"""PPTX 엔진 — 컨설팅 스타일 디자인 시스템 + 레이아웃 슬롯 채우기.

디자인 원칙 (references/design-system.md 와 같은 내용)
- 제목 = 액션 타이틀(결론 문장, 1~2줄). 모든 장에서 같은 위치·같은 크기.
- 화면 구조: 트래커 → 제목 → 구분선 → 내용 영역(근거 하나) → 출처 → 쪽번호
- 색: 남색(primary) 중심 + 회색. 강조색(accent)은 '봐야 할 한 곳'에만.
- 12단 그리드, 왼쪽 정렬, 표의 숫자는 오른쪽 정렬.
- 채운 뒤 balance()가 실제 분량을 재서 내용 묶음을 세로 '시각적 중심'에 놓는다.
- 슬라이드에는 레이아웃의 플레이스홀더만 쓴다. 새 텍스트박스를 덧그리지 않는다.
"""
from __future__ import annotations

import copy
import io
import re
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.parts.slide import SlideLayoutPart
from pptx.util import Pt, Emu

from ys_common import inch, split_bold, flatten_items

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
NSDECL = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

# =====================================================================
# 디자인 토큰
# =====================================================================
SLIDE_W, SLIDE_H = 13.333, 7.5          # 16:9 (인치)
M = 0.6                                  # 좌우 여백
CW = SLIDE_W - 2 * M                     # 콘텐츠 폭 12.133
GUTTER = 0.25                            # 단 사이 간격
COLS = 12
COL_W = (CW - GUTTER * (COLS - 1)) / COLS


def col(start, span):
    """12단 그리드: start(1~12)부터 span단 → (x, w)"""
    x = M + (start - 1) * (COL_W + GUTTER)
    return x, span * COL_W + (span - 1) * GUTTER


# 세로 리듬 (인치)
Y_TRACKER, H_TRACKER = 0.36, 0.26
Y_TITLE, H_TITLE = 0.62, 0.92            # 24pt 2줄까지 (구분선까지 여유 0.1in)
Y_RULE = 1.64                            # 제목 아래 구분선
Z_TOP, Z_BOT = 1.92, 6.45                # 내용 영역
Y_SOURCE, H_SOURCE = 6.58, 0.3
Y_FOOTER = 7.02

# 타입 스케일 (pt) — 이 크기만 쓴다
TYPE = {
    "cover": 36, "section": 32, "statement": 30, "number": 66, "closing": 30,
    "title": 24, "lead": 18, "head": 16, "body": 16, "sub": 14, "sub2": 12,
    "table": 12, "chart": 11, "tracker": 11, "source": 9, "footer": 8,
}
OPTICAL = 0.38          # 남는 세로 공간 중 위에 둘 비율 (0.5 = 정중앙)

# 한국어 줄바꿈: 어절 단위(latinLnBrk=0), 동아시아 금칙(eaLnBrk=1)
LNBRK = 'eaLnBrk="1" latinLnBrk="0" hangingPunct="1"'
SCH = lambda n: f'<a:solidFill><a:schemeClr val="{n}"/></a:solidFill>'  # noqa: E731
WHITE = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
RULE_GRAY = "C9CED6"
LINE_GRAY = "D9DDE3"


# =====================================================================
# XML 조각
# =====================================================================
def _fill(color):
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


def _rpr(sz=None, b=False, color=None, tag="a:defRPr"):
    attrs = 'lang="ko-KR" altLang="en-US"'
    if sz:
        attrs += f' sz="{int(sz * 100)}"'
    attrs += f' b="{1 if b else 0}"'
    fill = color if color is None or color.startswith("<") else _fill(color)
    return f"<{tag} {attrs}>{fill or ''}<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/></{tag}>"


def _lvl(n, sz, color=None, b=False, bullet=None, mar=0.0, ind=0.0, algn="l", spc_before=0, line=None, bu_color=None):
    bu = "<a:buNone/>"
    if bullet == "auto":
        bu = '<a:buClr><a:schemeClr val="accent2"/></a:buClr><a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>'
    elif bullet:
        buc = bu_color or '<a:schemeClr val="tx2"/>'
        bu = f'<a:buClr>{buc}</a:buClr><a:buSzPct val="80000"/><a:buFont typeface="Arial"/><a:buChar char="{bullet}"/>'
    ln = f'<a:lnSpc><a:spcPct val="{int(line * 1000)}"/></a:lnSpc>' if line else ""
    sb = f'<a:spcBef><a:spcPts val="{int(spc_before * 100)}"/></a:spcBef>'
    return (
        f'<a:lvl{n}pPr marL="{inch(mar)}" indent="{inch(ind)}" algn="{algn}" {LNBRK}>'
        f"{ln}{sb}{bu}{_rpr(sz, b, color)}</a:lvl{n}pPr>"
    )


def BODY_LEVELS(sz=TYPE["body"]):
    """본문 3단계 ■ / – / · . 1단계 크기 sz, 2단계 -2pt, 3단계 -4pt."""
    return (
        _lvl(1, sz, SCH("tx1"), bullet="■", mar=0.28, ind=-0.28, spc_before=10, line=115)
        + _lvl(2, sz - 2, SCH("tx1"), bullet="–", mar=0.58, ind=-0.25, spc_before=4, line=115)
        + _lvl(3, sz - 4, SCH("accent4"), bullet="·", mar=0.85, ind=-0.2, spc_before=2, line=115)
    )


def PLAIN(sz, color="tx1", b=False, algn="l", line=115, spc=0):
    c = color if color.startswith("<") else SCH(color)
    return _lvl(1, sz, c, b=b, algn=algn, line=line, spc_before=spc)


def ph(sid, name, ph_type, idx, x, y, w, h, levels, prompt, anchor="t", ins=(0, 0, 0, 0),
       fill=None, geom=None, vert=None, line=None):
    """레이아웃 플레이스홀더 <p:sp>."""
    type_attr = f' type="{ph_type}"' if ph_type else ""
    idx_attr = f' idx="{idx}"' if idx is not None else ""
    sz_attr = ' sz="quarter"' if ph_type in ("dt", "ftr", "sldNum") else ""
    sppr = ""
    if fill or geom or line:
        sppr += f'<a:prstGeom prst="{geom or "rect"}"><a:avLst/></a:prstGeom>'
        sppr += (fill if fill and fill.startswith("<") else (_fill(fill) if fill else "<a:noFill/>"))
        sppr += f'<a:ln w="{line[1]}">{_fill(line[0])}</a:ln>' if line else "<a:ln><a:noFill/></a:ln>"
    if ph_type in ("pic", "tbl", "chart"):
        txbody = ""
    elif ph_type == "sldNum":
        txbody = (
            f'<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle>{levels}</a:lstStyle>'
            '<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="ko-KR" altLang="en-US"/><a:t>‹#›</a:t></a:fld></a:p></p:txBody>'
        )
    else:
        v = f' vert="{vert}"' if vert else ' vert="horz"'
        txbody = (
            f'<p:txBody><a:bodyPr{v} wrap="square" lIns="{inch(ins[0])}" tIns="{inch(ins[1])}" rIns="{inch(ins[2])}" '
            f'bIns="{inch(ins[3])}" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
            f"<a:lstStyle>{levels}</a:lstStyle>"
            f'<a:p><a:r><a:rPr lang="ko-KR" altLang="en-US"/><a:t>{escape(prompt)}</a:t></a:r></a:p></p:txBody>'
        )
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{escape(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f"<p:nvPr><p:ph{type_attr}{sz_attr}{idx_attr}/></p:nvPr></p:nvSpPr>"
        f'<p:spPr><a:xfrm><a:off x="{inch(x)}" y="{inch(y)}"/><a:ext cx="{inch(w)}" cy="{inch(h)}"/></a:xfrm>{sppr}</p:spPr>'
        f"{txbody}</p:sp>"
    )


def rect(sid, name, x, y, w, h, fill, geom="rect"):
    fl = fill if fill.startswith("<") else _fill(fill)
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{inch(x)}" y="{inch(y)}"/><a:ext cx="{inch(w)}" cy="{inch(h)}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>{fl}<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ko-KR"/></a:p></p:txBody></p:sp>'
    )


def textbox(sid, name, x, y, w, h, text, sz, color, b=False, algn="l", anchor="ctr"):
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{inch(x)}" y="{inch(y)}"/><a:ext cx="{inch(w)}" cy="{inch(h)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}"/><a:lstStyle/>'
        f'<a:p><a:pPr algn="{algn}" {LNBRK}/><a:r>{_rpr(sz, b, color, tag="a:rPr")}<a:t>{escape(text)}</a:t></a:r></a:p></p:txBody></p:sp>'
    )


# =====================================================================
# 레이아웃 정의 (14종)
# =====================================================================
def _tracker(sid=40):
    return ph(sid, "트래커", "body", 13, M, Y_TRACKER, CW, H_TRACKER, PLAIN(TYPE["tracker"], "accent2", b=True),
              "섹션 이름 (예: 2. 현황 분석)", anchor="b")


def _title(sid=41):
    return ph(sid, "제목", "title", None, M, Y_TITLE, CW, H_TITLE, PLAIN(TYPE["title"], "tx2", b=True, line=108),
              "결론을 한 문장으로 (액션 타이틀, 2줄 이내)", anchor="t")


def _source(sid=42, idx=19):
    x, w = col(1, 10)
    return ph(sid, "출처", "body", idx, x, Y_SOURCE, w, H_SOURCE, PLAIN(TYPE["source"], "accent4", line=110),
              "출처: ", anchor="b")


def _sldnum(sid=43):
    return ph(sid, "슬라이드 번호", "sldNum", 12, SLIDE_W - M - 0.8, Y_FOOTER, 0.8, 0.26,
              PLAIN(TYPE["footer"], "accent4", algn="r"), "")


def _head(x, y, w, h, sid, name, idx, prompt, fill="tx2"):
    return ph(sid, name, "body", idx, x, y, w, h, PLAIN(TYPE["head"], WHITE, b=True, line=105), prompt,
              anchor="ctr", ins=(0.18, 0.04, 0.15, 0.04), fill=SCH(fill))


def content(name, shapes, source=True):
    s = [_tracker(), _title()] + shapes + ([_source()] if source else []) + [_sldnum()]
    return dict(name=name, master_shapes=True, bg=None, shapes=s)


def layout_defs(theme):
    C = theme["colors"]
    zh = Z_BOT - Z_TOP
    defs = []

    # 1 표지 — 흰 배경, 왼쪽 정렬, 제목 묶음은 세로 중앙, 아래 남색 띠
    defs.append(dict(name="표지", master_shapes=False, bg=None, shapes=[
        rect(20, "하단띠", 0, SLIDE_H - 0.28, SLIDE_W, 0.28, SCH("tx2")),
        ph(2, "제목", "ctrTitle", None, M, 2.3, col(1, 9)[1], 1.5, PLAIN(TYPE["cover"], "tx2", b=True, line=110), "문서 제목", anchor="t"),
        ph(3, "부제", "subTitle", 1, M, 3.95, col(1, 9)[1], 0.8, PLAIN(TYPE["lead"], "accent4", line=125), "부제: 이 문서가 답하는 질문 또는 결론 한 줄"),
        ph(4, "날짜·작성자", "body", 14, M, 6.35, col(1, 8)[1], 0.4, PLAIN(TYPE["sub2"], "accent4"), "2026. 11. 4. | 소속 · 작성자", anchor="b"),
        textbox(22, "회사명", col(9, 4)[0], 6.35, col(9, 4)[1], 0.4, theme["company"], TYPE["sub"], C["primary"], b=True, algn="r", anchor="b"),
    ]))
    # 2 목차 — 번호 목록, 세로 중앙
    defs.append(content("목차", [
        ph(4, "본문", "body", 1, col(2, 9)[0], Z_TOP, col(2, 9)[1], zh,
           _lvl(1, TYPE["lead"] + 2, SCH("tx1"), bullet="auto", mar=0.6, ind=-0.6, spc_before=14, line=110), "목차 항목"),
    ], source=False))
    # 3 섹션 — 흰 배경, 큰 번호 + 제목
    defs.append(dict(name="섹션", master_shapes=False, bg=None, shapes=[
        rect(20, "좌측선", M, 1.2, 0.06, 5.1, SCH("accent2")),
        ph(3, "번호", "body", 13, M + 0.4, 2.3, 4, 1.0, PLAIN(TYPE["number"] - 6, "accent2", b=True, line=100), "01", anchor="b"),
        ph(2, "제목", "title", None, M + 0.4, 3.4, col(1, 10)[1], 1.0, PLAIN(TYPE["section"], "tx2", b=True, line=108), "섹션 제목"),
        ph(4, "설명", "body", 14, M + 0.4, 4.5, col(1, 10)[1], 0.8, PLAIN(TYPE["lead"], "accent4", line=125), "이 섹션이 답하는 질문 한 줄"),
    ]))
    # 4 본문 — 단일 열 개조식 (10단 폭: 너무 긴 줄 방지)
    x, w = col(1, 11)
    defs.append(content("본문", [
        ph(4, "본문", "body", 1, x, Z_TOP, w, zh, BODY_LEVELS(TYPE["lead"]), "근거 (개조식, 명사형 종결)"),
    ]))
    # 5 2단비교 — 6+6단, 색 머리띠
    xl, wl = col(1, 6)
    xr, wr = col(7, 6)
    defs.append(content("2단비교", [
        _head(xl, Z_TOP, wl, 0.52, 4, "좌측제목", 15, "좌측 (예: 현재)", fill="accent4"),
        ph(5, "좌측본문", "body", 16, xl, Z_TOP + 0.7, wl, zh - 0.7, BODY_LEVELS(), "좌측 내용", ins=(0.1, 0, 0.05, 0)),
        _head(xr, Z_TOP, wr, 0.52, 6, "우측제목", 17, "우측 (예: 개선 후)", fill="tx2"),
        ph(7, "우측본문", "body", 18, xr, Z_TOP + 0.7, wr, zh - 0.7, BODY_LEVELS(), "우측 내용", ins=(0.1, 0, 0.05, 0)),
    ]))
    # 6 3단카드 — 4+4+4단. 요약(상황|발견|제언), 세 가지 축, KPI 셋
    shapes = []
    for i in range(3):
        x, w = col(1 + i * 4, 4)
        shapes.append(ph(4 + i * 2, f"카드{i+1}제목", "body", 20 + i * 2, x, Z_TOP, w, 0.95,
                         PLAIN(22, "tx2", b=True, line=100), f"카드 {i+1} 제목 또는 숫자",
                         anchor="b", ins=(0.22, 0.2, 0.2, 0.1), fill=C["light"]))
        shapes.append(ph(5 + i * 2, f"카드{i+1}본문", "body", 21 + i * 2, x, Z_TOP + 0.95, w, zh - 0.95,
                         BODY_LEVELS(TYPE["sub"] + 1), f"카드 {i+1} 설명", ins=(0.22, 0.1, 0.2, 0.28), fill=C["light"]))
    defs.append(content("3단카드", shapes))
    # 7 프로세스 — 4단계 화살표 + 설명 (3+3+3+3단). 머리 화살표는 고정 위치(움직이지 않음)
    shapes = []
    P_TOP = Z_TOP + 0.9
    for i in range(4):
        x, w = col(1 + i * 3, 3)
        geom = "homePlate" if i == 0 else "chevron"
        shapes.append(ph(4 + i * 2, f"단계{i+1}제목", "body", 30 + i * 2, x, P_TOP, w + 0.12, 0.72,
                         PLAIN(TYPE["head"], WHITE, b=True, line=100), f"{i+1}. 단계",
                         anchor="ctr", ins=(0.35 if i else 0.2, 0.04, 0.3, 0.04), fill=SCH("tx2" if i < 3 else "accent2"), geom=geom))
        shapes.append(ph(5 + i * 2, f"단계{i+1}본문", "body", 31 + i * 2, x, P_TOP + 0.97, w, Z_BOT - P_TOP - 0.97,
                         BODY_LEVELS(TYPE["sub"]), "활동·산출물", ins=(0.05, 0, 0.05, 0)))
    defs.append(content("프로세스", shapes))
    # 8 매트릭스 — 2x2 (7단) + 해설 (5단). 축 이름은 가로쓰기(세로쓰기는 프로그램마다 깨짐)
    mx, mw = col(1, 7)
    my = Z_TOP + 0.38
    qh = (Z_BOT - my - 0.42 - 0.08) / 2
    qw = (mw - 0.08) / 2
    shapes = [
        ph(4, "세로축", "body", 40, mx, Z_TOP, mw, 0.3, PLAIN(TYPE["sub2"], "accent4", b=True), "▲ 세로축 이름", anchor="b"),
        ph(5, "가로축", "body", 41, mx, my + qh * 2 + 0.16, mw, 0.3, PLAIN(TYPE["sub2"], "accent4", b=True, algn="r"), "가로축 이름 ▶", anchor="t"),
    ]
    quads = [("좌상", 0, 0, C["light"], "tx2"), ("우상", 1, 0, C["primary"], WHITE), ("좌하", 0, 1, C["light"], "accent4"), ("우하", 1, 1, C["light"], "tx2")]
    for k, (nm, cx, cy, fl, tc) in enumerate(quads):
        tcol = tc if tc.startswith("<") else SCH(tc)
        shapes.append(ph(6 + k, nm, "body", 42 + k, mx + cx * (qw + 0.08), my + cy * (qh + 0.08), qw, qh,
                         _lvl(1, TYPE["head"], tcol, b=True, spc_before=0, line=110)
                         + _lvl(2, TYPE["sub2"] + 1, tcol if tc == WHITE else SCH("tx1"), bullet="–", mar=0.25, ind=-0.2, spc_before=4, line=115,
                                bu_color='<a:srgbClr val="FFFFFF"/>' if tc == WHITE else None),
                         f"{nm}: 이름", anchor="t", ins=(0.22, 0.2, 0.18, 0.12), fill=fl))
    hx, hw = col(8, 5)
    shapes.append(rect(31, "해설선", hx, my, 0.05, qh * 2 + 0.08, SCH("accent2")))
    shapes.append(ph(10, "해설", "body", 14, hx + 0.3, my, hw - 0.3, qh * 2 + 0.08, BODY_LEVELS(TYPE["body"]), "매트릭스가 말하는 것", anchor="ctr"))
    defs.append(content("매트릭스", shapes))
    # 9 표
    defs.append(content("표", [ph(4, "표", "tbl", 1, M, Z_TOP, CW, zh, "", "표를 넣으세요")]))
    # 10 차트 — 차트 8단 + 시사점 상자 4단
    cx, cw = col(1, 8)
    sx, sw = col(9, 4)
    defs.append(content("차트", [
        ph(4, "차트", "chart", 1, cx, Z_TOP, cw, zh, "", "차트를 넣으세요"),
        rect(30, "시사점배경", sx, Z_TOP, sw, zh, C["light"]),
        rect(31, "시사점선", sx, Z_TOP, 0.06, zh, SCH("accent2")),
        textbox(32, "시사점라벨", sx + 0.3, Z_TOP + 0.25, sw - 0.5, 0.3, "시사점", TYPE["sub2"], C["accent"], b=True, anchor="t"),
        ph(5, "해설", "body", 14, sx + 0.3, Z_TOP + 0.7, sw - 0.5, zh - 0.95, BODY_LEVELS(TYPE["sub"] + 1), "차트에서 읽어야 할 점", anchor="t"),
    ]))
    # 11 빅넘버 — 큰 숫자 + 설명 (6단) | 근거 (6단)
    lx, lw = col(1, 6)
    rx, rw = col(7, 6)
    defs.append(content("빅넘버", [
        ph(4, "숫자", "body", 50, lx, Z_TOP, lw, 1.4, PLAIN(TYPE["number"], "accent2", b=True, line=100), "42%", anchor="b"),
        ph(5, "설명", "body", 51, lx, Z_TOP + 1.5, lw, 1.2, PLAIN(TYPE["lead"], "tx2", b=True, line=120), "숫자의 의미 한 줄"),
        rect(30, "구분선", rx - GUTTER / 2 - 0.01, Z_TOP + 0.6, 0.02, zh - 1.2, RULE_GRAY),
        ph(6, "본문", "body", 1, rx + 0.2, Z_TOP, rw - 0.2, zh, BODY_LEVELS(TYPE["body"]), "근거 2~4개", anchor="t"),
    ]))
    # 12 핵심메시지 — 한 문장 결론 (요약·전환)
    x, w = col(1, 10)
    defs.append(dict(name="핵심메시지", master_shapes=True, bg=None, shapes=[
        rect(20, "좌측선", M, 1.6, 0.06, 4.3, SCH("accent2")),
        ph(2, "제목", "title", None, x + 0.4, 1.9, w, 0.4, PLAIN(TYPE["sub2"] + 1, "accent2", b=True), "요약", anchor="b"),
        ph(3, "핵심메시지", "body", 13, x + 0.4, 2.45, w, 2.0, PLAIN(TYPE["statement"], "tx2", b=True, line=125), "가장 중요한 결론 한두 문장"),
        ph(4, "보조설명", "body", 14, x + 0.4, 4.7, w, 1.0, PLAIN(TYPE["lead"], "accent4", line=130), "근거 또는 요청 한 줄"),
        _sldnum(),
    ]))
    # 13 이미지 — 그림 7단 + 설명 5단
    ix, iw = col(1, 7)
    tx_, tw = col(8, 5)
    defs.append(content("이미지", [
        ph(4, "이미지", "pic", 1, ix, Z_TOP, iw, zh, "", "그림을 넣으세요"),
        ph(5, "캡션", "body", 14, tx_ + 0.1, Z_TOP, tw - 0.1, zh, BODY_LEVELS(TYPE["body"]), "그림에서 볼 곳", anchor="ctr"),
    ]))
    # 14 마무리 — 다음 단계 / 요청 / 문의
    defs.append(dict(name="마무리", master_shapes=False, bg=None, shapes=[
        rect(20, "하단띠", 0, SLIDE_H - 0.28, SLIDE_W, 0.28, SCH("tx2")),
        ph(2, "제목", "title", None, M, 2.45, col(1, 10)[1], 0.9, PLAIN(TYPE["closing"], "tx2", b=True, line=108), "다음 단계", anchor="t"),
        ph(3, "본문", "body", 1, M, 3.55, col(1, 10)[1], 2.3, BODY_LEVELS(TYPE["lead"]), "요청 사항 · 담당 · 기한 / 문의처"),
    ]))
    return defs


# =====================================================================
# 세로 균형
#   stacks: 서로 독립적으로 세로 배치되는 묶음들. rows = (슬롯들, 고정높이|None, 최소높이)
# =====================================================================
def _S(zone, rows, gaps=(), bias=OPTICAL):
    return dict(zone=zone, rows=rows, gaps=list(gaps), bias=bias)


BALANCE = {
    "본문": [_S((Z_TOP, Z_BOT), [(["본문"], None, 0.8)])],
    "목차": [_S((Z_TOP, Z_BOT), [(["본문"], None, 0.8)])],
    "2단비교": [_S((Z_TOP, Z_BOT), [(["좌측제목", "우측제목"], 0.52, 0), (["좌측본문", "우측본문"], None, 1.0)], [0.2])],
    "3단카드": [_S((Z_TOP, Z_BOT), [(["카드1제목", "카드2제목", "카드3제목"], 0.95, 0), (["카드1본문", "카드2본문", "카드3본문"], None, 1.4)], [0])],
    "프로세스": [_S((Z_TOP + 0.9, Z_BOT), [(["단계1제목", "단계2제목", "단계3제목", "단계4제목"], 0.72, 0),
                                   (["단계1본문", "단계2본문", "단계3본문", "단계4본문"], None, 1.0)], [0.25], bias=0)],
    "빅넘버": [_S((Z_TOP, Z_BOT), [(["숫자"], 1.3, 0), (["설명"], None, 0.5)], [0.15]),
              _S((Z_TOP, Z_BOT), [(["본문"], None, 0.8)])],
    "핵심메시지": [_S((0.7, 6.6), [(["제목"], 0.4, 0), (["핵심메시지"], None, 0.6), (["보조설명"], None, 0.4)], [0.2, 0.4])],
    "표지": [_S((1.2, 5.9), [(["제목"], None, 0.7), (["부제"], None, 0.4)], [0.35])],
    "섹션": [_S((1.2, 6.3), [(["번호"], 1.0, 0), (["제목"], None, 0.6), (["설명"], None, 0.4)], [0.1, 0.25], bias=0.45)],
    "마무리": [_S((1.4, 6.6), [(["제목"], None, 0.6), (["본문"], None, 0.4)], [0.3])],
}


SIBLINGS = {
    "3단카드": [["카드1제목", "카드2제목", "카드3제목", "카드1본문", "카드2본문", "카드3본문"]],
    "프로세스": [["단계1본문", "단계2본문", "단계3본문", "단계4본문"]],
    "2단비교": [["좌측본문", "우측본문"]],
}
# 비면 화면 균형이 깨지는 칸 (check.py 가 경고)
REQUIRED = {
    "3단카드": ["카드1제목", "카드2제목", "카드3제목", "카드1본문", "카드2본문", "카드3본문"],
    "프로세스": ["단계1제목", "단계2제목", "단계3제목", "단계4제목", "단계1본문", "단계2본문", "단계3본문", "단계4본문"],
    "2단비교": ["좌측제목", "우측제목", "좌측본문", "우측본문"],
    "매트릭스": ["좌상", "우상", "좌하", "우하"],
    "차트": ["해설"], "빅넘버": ["숫자", "본문"], "이미지": ["캡션"],
}


def is_yangsik(prs):
    return "yangsik" in (prs.core_properties.keywords or "")


def balance(slide, prs):
    stacks = BALANCE.get(slide.slide_layout.name)
    if not stacks:
        return
    import check as _chk
    ty = _chk.Checker().ty
    by_name = {s.name: s for s in slide.placeholders}
    for spec in stacks:
        rows, gaps = [], []
        for i, (names, fixed, min_h) in enumerate(spec["rows"]):
            present = [by_name[n] for n in names if n in by_name]
            if not present:
                continue
            h = fixed if fixed is not None else max(max(_chk.estimate_height(s, prs, ty) for s in present), min_h)
            if rows and i - 1 < len(spec["gaps"]):
                gaps.append(spec["gaps"][i - 1])
            rows.append((present, h))
        if not rows:
            continue
        z0, z1 = spec["zone"]
        total = sum(h for _, h in rows) + sum(gaps)
        if total >= (z1 - z0):
            rows[-1] = (rows[-1][0], max((z1 - z0) - (total - rows[-1][1]), 0.3))
            y = z0
        else:
            y = z0 + ((z1 - z0) - total) * spec["bias"]
        for k, (present, h) in enumerate(rows):
            for s in present:
                left, width = s.left, s.width          # 상속값을 먼저 읽어 둠 (xfrm 생성 시 0이 되지 않게)
                s.left, s.width = left, width
                s.top, s.height = inch(y), inch(h)
            y += h + (gaps[k] if k < len(gaps) else 0)


# =====================================================================
# 베이스 양식 빌드
# =====================================================================
def _set_theme(prs, theme):
    master = prs.slide_masters[0]
    tpart = master.part.part_related_by(RT.THEME)
    root = etree.fromstring(tpart.blob)
    C = theme["colors"]
    root.set("name", "양식대로")
    cs = root.find(".//a:clrScheme", NS)
    cs.set("name", "양식대로")
    mapping = {"dk1": C["text"], "lt1": "FFFFFF", "dk2": C["primary"], "lt2": C["light"],
               "accent1": C["primary"], "accent2": C["accent"], "accent3": C["highlight"],
               "accent4": C["muted"], "accent5": "9AA3B2", "accent6": "C9CED6",
               "hlink": C["accent"], "folHlink": C["muted"]}
    for tag, val in mapping.items():
        el = cs.find(f"a:{tag}", NS)
        for c in list(el):
            el.remove(c)
        etree.SubElement(el, f"{{{NS['a']}}}srgbClr", val=val)
    fs = root.find(".//a:fontScheme", NS)
    fs.set("name", "양식대로")
    for kind in ("majorFont", "minorFont"):
        f = fs.find(f"a:{kind}", NS)
        f.find("a:latin", NS).set("typeface", theme["fonts"]["latin"])
        f.find("a:ea", NS).set("typeface", theme["fonts"]["korean"])
        for sub in f.findall("a:font", NS):
            if sub.get("script") == "Hang":
                sub.set("typeface", theme["fonts"]["korean"])
    tpart._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _set_master(prs, theme):
    master = prs.slide_masters[0]
    el = master._element
    C = theme["colors"]
    tree = el.find(".//p:cSld/p:spTree", NS)
    for sp in list(tree):
        if sp.tag.endswith("}sp") or sp.tag.endswith("}pic"):
            tree.remove(sp)
    xmls = [
        ph(2, "제목", "title", None, M, Y_TITLE, CW, H_TITLE, "", "제목", anchor="t"),
        ph(3, "본문", "body", 1, M, Z_TOP, CW, Z_BOT - Z_TOP, "", "본문"),
        rect(10, "제목구분선", M, Y_RULE, CW, 0.012, RULE_GRAY),
        _sldnum(12),
    ]
    if not theme.get("logo"):
        xmls.append(textbox(11, "회사명", M, Y_FOOTER, 6, 0.26, theme["company"], TYPE["footer"], C["muted"], b=True))
    for x in xmls:
        tree.append(parse_xml(f'<p:wrap {NSDECL}>{x}</p:wrap>')[0])
    tx = el.find("p:txStyles", NS)
    tx.getparent().remove(tx)
    tx_xml = (
        f'<p:txStyles {NSDECL}>'
        f'<p:titleStyle>{PLAIN(TYPE["title"], "tx2", b=True, line=108)}</p:titleStyle>'
        f'<p:bodyStyle>{BODY_LEVELS()}'
        + "".join(_lvl(n, TYPE["sub2"], SCH("tx1"), mar=1.0 + n * 0.2) for n in range(4, 10))
        + '</p:bodyStyle>'
        f'<p:otherStyle>{PLAIN(TYPE["sub2"])}</p:otherStyle>'
        '</p:txStyles>'
    )
    idx = list(el).index(el.find("p:sldLayoutIdLst", NS)) + 1
    el.insert(idx, parse_xml(tx_xml))
    pres = prs.part._element
    dts = pres.find("p:defaultTextStyle", NS)
    if dts is not None:
        for d in dts.iter(f"{{{NS['a']}}}defRPr"):
            d.set("lang", "ko-KR")
            d.set("altLang", "en-US")


def _layout_xml(d):
    bg = (f'<p:bg><p:bgPr>{_fill(d["bg"])}<a:effectLst/></p:bgPr></p:bg>' if d["bg"] else "")
    show = "" if d["master_shapes"] else ' showMasterSp="0"'
    return (
        f'<p:sldLayout {NSDECL} preserve="1" userDrawn="1"{show}>'
        f'<p:cSld name="{escape(d["name"])}">{bg}<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + "".join(d["shapes"]) +
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
    ).encode("utf-8")


def _add_layout_part(prs, xml_bytes, n):
    master = prs.slide_masters[0]
    part = SlideLayoutPart.load(PackURI(f"/ppt/slideLayouts/slideLayout{n}.xml"), CT.PML_SLIDE_LAYOUT,
                                prs.part.package, xml_bytes)
    part.relate_to(master.part, RT.SLIDE_MASTER)
    rId = master.part.relate_to(part, RT.SLIDE_LAYOUT)
    lst = master._element.find("p:sldLayoutIdLst", NS)
    used = [int(x.get("id")) for x in lst] + [int(x.get("id")) for x in prs.part._element.iter(f"{{{NS['p']}}}sldMasterId")]
    lst.append(parse_xml(f'<p:sldLayoutId {NSDECL} id="{max(used) + 1}" r:id="{rId}"/>'))


def _pic_xml(sid, name, rId, x, y, w, h):
    return (
        f'<p:pic {NSDECL}><p:nvPicPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>'
        f'<p:blipFill><a:blip r:embed="{rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        f'<p:spPr><a:xfrm><a:off x="{inch(x)}" y="{inch(y)}"/><a:ext cx="{inch(w)}" cy="{inch(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )


def _add_logo(part, tree, path, x, y, h, max_w, right=False):
    from PIL import Image
    with Image.open(path) as im:
        ratio = im.width / im.height
    w = min(h * ratio, max_w)
    h = w / ratio
    if right:
        x = x - w
    _ip, rId = part.get_or_add_image_part(str(path))
    tree.append(parse_xml(_pic_xml(90, "로고", rId, x, y, w, h)))


def build_base(theme) -> bytes:
    prs = Presentation()
    prs.slide_width, prs.slide_height = inch(SLIDE_W), inch(SLIDE_H)
    _set_theme(prs, theme)
    _set_master(prs, theme)
    defs = layout_defs(theme)
    layouts = list(prs.slide_layouts)
    for layout, d in zip(layouts, defs):
        layout.part._element = parse_xml(_layout_xml(d))
    for extra in layouts[len(defs):]:
        prs.slide_layouts.remove(extra)
    for i, d in enumerate(defs[len(layouts):], len(layouts) + 1):
        _add_layout_part(prs, _layout_xml(d), i)
    logo = theme.get("logo")
    m = prs.slide_masters[0]
    mtree = m._element.find(".//p:cSld/p:spTree", NS)
    if theme.get("top_band"):      # 원본 양식의 상단 색띠를 살림
        mtree.append(parse_xml(f'<p:wrap {NSDECL}>{rect(15, "상단띠", 0, 0, SLIDE_W, 0.14, SCH("tx2"))}</p:wrap>')[0])
    if logo and Path(logo).exists():
        if theme.get("logo_position") == "top-right":
            _add_logo(m.part, mtree, logo, SLIDE_W - M, 0.26, 0.34, 2.2, right=True)
            mtree.append(parse_xml(f'<p:wrap {NSDECL}>{textbox(16, "회사명", M, Y_FOOTER, 6, 0.26, theme["company"], TYPE["footer"], theme["colors"]["muted"], b=True)}</p:wrap>')[0])
        else:
            _add_logo(m.part, mtree, logo, M, Y_FOOTER - 0.02, 0.26, 2.0)
        for l in prs.slide_layouts:
            if l.part._element.cSld.get("name") in ("표지", "마무리"):
                # l._element 는 교체 전 요소를 캐시하고 있을 수 있어 part 의 현재 요소를 쓴다
                _add_logo(l.part, l.part._element.find(".//p:cSld/p:spTree", NS), logo, M, 0.6, 0.5, 2.6)
    core = prs.core_properties
    core.title, core.author, core.language = "양식대로 표준 양식", theme["company"], "ko-KR"
    core.keywords = "yangsik"
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# =====================================================================
# 채우기
# =====================================================================
def layout_slots(layout):
    """레이아웃의 슬롯 이름 → (idx, type)."""
    out = {}
    for shp in layout.placeholders:
        pf = shp.placeholder_format
        t = str(pf.type).split(".")[-1].split(" ")[0]
        if "SLIDE_NUMBER" in t or "FOOTER" in t or "DATE" in t:
            continue
        out[shp.name] = (pf.idx, t)
    return out


def _set_runs(paragraph, text, level=0):
    paragraph.level = level
    for seg, bold in split_bold(text):
        r = paragraph.add_run()
        r.text = seg
        rpr = r._r.get_or_add_rPr()
        rpr.set("lang", "ko-KR")
        rpr.set("altLang", "en-US")
        if bold:
            r.font.bold = True


def set_text(shape, value):
    from pptx.text.text import _Paragraph
    tf = shape.text_frame
    items = flatten_items(value) if not isinstance(value, str) else [(0, value)]
    txBody = tf._txBody
    for p in list(txBody.findall(f"{{{NS['a']}}}p")):
        txBody.remove(p)
    for level, text in items:
        _set_runs(_Paragraph(txBody.add_p(), tf), text, level)
    if not items:
        txBody.add_p()


_NUM_TOKEN = re.compile(r"^[약~±+\-−–(]?[\d.,]*\d[\d.,]*(?:%p|%|배|x|×)?[가-힣]{0,2}\)?$")
_UNIT_WORD = {"원", "만", "억", "조", "명", "건", "개", "배", "주", "일", "년", "개월", "시간", "분", "달러", "감소", "증가",
              "예상", "이상", "이하", "내외", "수준", "p", "%p", "-", "→", "~"}


def is_number_like(v):
    """'1억 2,960만 원' '-2억 원' '2%p 감소 예상' '+20%' → True / '초기 8억 + 연 1.5억 원' → False"""
    toks = str(v).replace("\n", " ").split()
    if not toks or not any(ch.isdigit() for ch in v):
        return False
    return all(_NUM_TOKEN.match(t) or t in _UNIT_WORD for t in toks)


NUM_RE = re.compile(r"^$")  # 하위 호환 (사용 안 함)


def _tint(hexcolor, amount):
    """흰색과 섞기: amount=0.88 → 88% 흰색"""
    r, g, b = (int(hexcolor[i:i + 2], 16) for i in (0, 2, 4))
    mix = lambda c: round(c + (255 - c) * amount)  # noqa: E731
    return f"{mix(r):02X}{mix(g):02X}{mix(b):02X}"


def _cell_border(cell, side, w, color):
    tcPr = cell._tc.get_or_add_tcPr()
    tag = {"top": "a:lnT", "bottom": "a:lnB", "left": "a:lnL", "right": "a:lnR"}[side]
    for old in tcPr.findall(tag.replace("a:", f"{{{NS['a']}}}")):
        tcPr.remove(old)
    if w:
        xml = f'<{tag} {NSDECL} w="{w}" cap="flat" cmpd="sng" algn="ctr">{_fill(color)}<a:prstDash val="solid"/></{tag}>'
    else:
        xml = f'<{tag} {NSDECL} w="0"><a:noFill/></{tag}>'
    el = parse_xml(xml)
    # 순서: lnL lnR lnT lnB ... solidFill
    order = ["lnL", "lnR", "lnT", "lnB"]
    before = None
    for child in tcPr:
        nm = etree.QName(child).localname
        if nm not in order or order.index(nm) > order.index(tag[2:]):
            before = child
            break
    if before is not None:
        before.addprevious(el)
    else:
        tcPr.append(el)


def _cell_fill(cell, color):
    if color:
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(color)
    else:
        cell.fill.background()


def fill_table(ph_shape, spec, prs=None, yangsik=True):
    """컨설팅 스타일 표: 머리글 굵게 + 굵은 아래선, 행은 가는 가로선, 세로선 없음, 숫자 오른쪽 정렬."""
    ph_name = ph_shape.name
    header = spec.get("header") or spec.get("머리글") or []
    rows = spec.get("rows") or spec.get("행") or []
    hl = spec.get("highlight_col", spec.get("강조열"))
    n_cols = max([len(header)] + [len(r) for r in rows] + [1])
    all_rows = ([header] if header else []) + [list(r) + [""] * (n_cols - len(r)) for r in rows]
    avail = ph_shape.height            # 표 영역 높이 (insert 후에는 행 수 기준으로 바뀜)
    gf = ph_shape.insert_table(max(len(all_rows), 1), n_cols)
    table = gf.table
    numeric_col = []
    for c in range(n_cols):
        vals = [str(r[c]).strip() for r in rows if c < len(r) and str(r[c]).strip() not in ("", "-")]
        numeric_col.append(bool(vals) and sum(is_number_like(v) for v in vals) >= max(1, len(vals) * 0.6))
    widths = spec.get("col_widths") or spec.get("열너비")
    total = gf.width
    if widths:
        s = sum(widths[:n_cols])
        for i, w in enumerate(widths[:n_cols]):
            table.columns[i].width = Emu(int(total * w / s))
    if yangsik:
        tbl = gf._element.graphic.graphicData.tbl
        tblPr = tbl.tblPr
        sid = tblPr.find(f"{{{NS['a']}}}tableStyleId")
        if sid is None:
            sid = etree.SubElement(tblPr, f"{{{NS['a']}}}tableStyleId")
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid
        for k in ("bandRow", "firstRow"):
            tblPr.set(k, "0")
    C_TX2 = _scheme_rgb(prs, "dk2") if prs else "1F3A5F"
    C_ACC = _scheme_rgb(prs, "accent2") if prs else "2E75B6"
    C_HL = _tint(C_ACC, 0.87)          # 강조 열 배경: 강조색 13% (흰 배경과 확실히 구분)
    n_body = len(rows)
    fsz = 14 if (n_body <= 5 and n_cols <= 4) else (13 if n_body <= 8 else TYPE["table"])
    for r_i, row in enumerate(all_rows):
        is_head = bool(header) and r_i == 0
        for c_i in range(n_cols):
            cell = table.cell(r_i, c_i)
            cell.text_frame.text = ""
            p = cell.text_frame.paragraphs[0]
            for k, line in enumerate(str(row[c_i]).split("\n")):
                if k:
                    p = cell.text_frame.add_paragraph()
                _set_runs(p, line)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE if not is_head else MSO_ANCHOR.BOTTOM
            cell.margin_left = inch(0.12 if (hl is not None and c_i == hl) else (0.08 if c_i else 0.02))
            cell.margin_right = inch(0.08)
            cell.margin_top = cell.margin_bottom = inch(0.04)
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.RIGHT if numeric_col[c_i] else PP_ALIGN.LEFT
                for r in p.runs:
                    r.font.size = Pt(fsz)
                    if is_head or (c_i == 0 and spec.get("label_col", True)):
                        r.font.bold = True
                    if yangsik:
                        r.font.color.rgb = RGBColor.from_string(C_TX2 if is_head else "222222")
                        if hl is not None and c_i == hl:
                            r.font.color.rgb = RGBColor.from_string("FFFFFF" if is_head else C_TX2)
                            r.font.bold = True
            if yangsik:
                if hl is not None and c_i == hl:
                    _cell_fill(cell, C_ACC if is_head else C_HL)   # 머리글 강조색 + 흰 글자, 열은 연한 강조색
                else:
                    _cell_fill(cell, None)
                for side in ("left", "right", "top"):
                    _cell_border(cell, side, 0, None)
                last = r_i == len(all_rows) - 1
                if is_head:
                    _cell_border(cell, "bottom", 19050, C_TX2)
                elif last:
                    _cell_border(cell, "bottom", 12700, C_TX2)
                else:
                    _cell_border(cell, "bottom", 6350, LINE_GRAY)
    # 행 높이: 영역을 적당히 채우되(최대 0.62in) 넘치지 않게
    n = max(len(all_rows), 1)
    row_h = max(min(int(avail * 0.88 / n), inch(0.62)), inch(0.36))
    for r_i, row in enumerate(table.rows):
        row.height = Emu(inch(0.45) if (header and r_i == 0) else row_h)
    if yangsik:  # 표가 영역보다 많이 작으면 살짝 위쪽 중심에
        total = sum(r.height for r in table.rows)
        gf.height = Emu(total)
        if total < avail * 0.8:
            gf.top = Emu(gf.top + int((avail - total) * 0.25))
    gf.name = ph_name
    return gf


def _scheme_rgb(prs, name):
    try:
        theme = etree.fromstring(prs.slide_masters[0].part.part_related_by(RT.THEME).blob)
        el = theme.find(f".//a:clrScheme/a:{name}", NS)
        c = el[0]
        return c.get("val") or c.get("lastClr")
    except Exception:
        return {"dk2": "1F3A5F", "accent2": "2E75B6", "lt2": "F2F4F7"}.get(name, "333333")


CHART_TYPES = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "hbar": XL_CHART_TYPE.BAR_CLUSTERED, "stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "line": XL_CHART_TYPE.LINE_MARKERS, "pie": XL_CHART_TYPE.PIE, "doughnut": XL_CHART_TYPE.DOUGHNUT,
}
SERIES_GRAY = ["9AA3B2", "C9CED6", "6B7280"]


def fill_chart(ph_shape, spec, prs=None, yangsik=True):
    """컨설팅 스타일 차트: 눈금선·값축 없음, 데이터 레이블 직접 표시, 남색+회색, 강조 항목만 강조색."""
    ph_name = ph_shape.name
    cd = CategoryChartData()
    cd.categories = spec.get("categories") or spec.get("항목")
    series = spec.get("series") or spec.get("계열") or {}
    for name, vals in series.items():
        cd.add_series(name, vals)
    ctype = CHART_TYPES.get(spec.get("type", "bar"), XL_CHART_TYPE.COLUMN_CLUSTERED)
    gf = ph_shape.insert_chart(ctype, cd)
    ch = gf.chart
    ch.font.size = Pt(TYPE["chart"])
    pie = ctype in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT)
    ch.has_legend = len(series) > 1 or pie
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.TOP
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(TYPE["chart"])
    plot = ch.plots[0]
    if spec.get("labels", True):
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(TYPE["chart"])
        if spec.get("number_format"):
            dl.number_format = spec["number_format"]
            dl.number_format_is_linked = False
        if ctype in (XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED):
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
    if yangsik and prs is not None:
        primary = _scheme_rgb(prs, "dk2")
        accent = _scheme_rgb(prs, "accent2")
        hl = spec.get("highlight", spec.get("강조"))
        if not pie:
            try:
                plot.gap_width = 60
            except Exception:
                pass
            for i, s in enumerate(plot.series):
                color = primary if i == 0 else SERIES_GRAY[(i - 1) % len(SERIES_GRAY)]
                fmt = s.format
                if ctype == XL_CHART_TYPE.LINE_MARKERS:
                    fmt.line.color.rgb = RGBColor.from_string(color)
                    fmt.line.width = Pt(2.25)
                    s.marker.format.fill.solid()
                    s.marker.format.fill.fore_color.rgb = RGBColor.from_string(color)
                    s.smooth = False
                else:
                    fmt.fill.solid()
                    fmt.fill.fore_color.rgb = RGBColor.from_string(color)
                    if hl is not None and i == 0 and len(series) == 1:
                        for j in ([hl] if isinstance(hl, int) else hl):
                            pt = s.points[j]
                            pt.format.fill.solid()
                            pt.format.fill.fore_color.rgb = RGBColor.from_string(accent)
            va = ch.value_axis
            va.has_major_gridlines = False
            va.visible = not spec.get("labels", True)
            ca = ch.category_axis
            ca.format.line.color.rgb = RGBColor.from_string(RULE_GRAY)
            ca.tick_labels.font.size = Pt(TYPE["chart"])
            ca.tick_labels.font.color.rgb = RGBColor.from_string("4B5563")
            from pptx.enum.chart import XL_TICK_MARK
            ca.major_tick_mark = XL_TICK_MARK.NONE
        else:
            pal = [primary] + SERIES_GRAY + [accent]
            for j, pt in enumerate(plot.series[0].points):
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = RGBColor.from_string(accent if hl == j else pal[j % len(pal)])
    gf.name = ph_name
    return gf


def _add_slide_number(slide, layout):
    for shp in layout.placeholders:
        if "SLIDE_NUMBER" in str(shp.placeholder_format.type):
            slide.shapes._spTree.append(copy.deepcopy(shp._element))
            return


def delete_all_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        prs.part.drop_rel(sldId.get(f"{{{NS['r']}}}id"))
        sldIdLst.remove(sldId)


LAYOUT_ALIAS = {"요약": "3단카드", "세가지": "3단카드", "KPI": "3단카드", "비교": "2단비교", "단계": "프로세스",
                "2x2": "매트릭스", "숫자": "빅넘버", "결론": "핵심메시지", "그림": "이미지", "끝": "마무리"}


def get_layout(prs, name):
    for nm in (name, LAYOUT_ALIAS.get(name)):
        for l in prs.slide_layouts:
            if nm and l.name == nm:
                return l
    names = [l.name for l in prs.slide_layouts]
    raise KeyError(f"레이아웃 '{name}' 없음. 사용 가능: {names}")


SKIP_KEYS = {"layout", "레이아웃", "노트", "notes", "id"}


def _migrate(spec, slots):
    """옛 형식(제목=주제어 + 핵심메시지=결론) → 새 형식(트래커=주제어, 제목=결론)."""
    spec = dict(spec)
    if "핵심메시지" in spec and "핵심메시지" not in slots and "제목" in slots:
        if "트래커" in slots and "트래커" not in spec and spec.get("제목"):
            spec["트래커"] = spec["제목"]
        spec["제목"] = spec.pop("핵심메시지")
    return spec


def add_slide(prs, spec, base_dir=None, warn=print):
    layout = get_layout(prs, spec.get("layout") or spec.get("레이아웃"))
    slots = layout_slots(layout)
    spec = _migrate(spec, slots)
    yangsik = is_yangsik(prs)
    slide = prs.slides.add_slide(layout)
    by_idx = {shp.placeholder_format.idx: shp for shp in slide.placeholders}
    for name, (idx, _t) in slots.items():
        if idx in by_idx:
            by_idx[idx].name = name
    used = set()
    for key, value in spec.items():
        if key in SKIP_KEYS:
            continue
        if key not in slots:
            warn(f"  ! '{layout.name}' 레이아웃에 '{key}' 슬롯이 없습니다 (가능: {list(slots)})")
            continue
        idx, t = slots[key]
        shp = by_idx.get(idx)
        if shp is None:
            continue
        used.add(idx)
        if t == "TABLE":
            fill_table(shp, value, prs, yangsik)
        elif t == "CHART":
            fill_chart(shp, value, prs, yangsik)
        elif t == "PICTURE":
            p = Path(value)
            if base_dir and not p.is_absolute():
                p = Path(base_dir) / p
            if p.exists():
                shp.insert_picture(str(p))
            else:
                warn(f"  ! 이미지 파일 없음: {p} (빈 슬롯으로 둡니다)")
                used.discard(idx)
        else:
            set_text(shp, value)
    # 묶음 칸(카드·단계·2단)은 형제 칸이 하나라도 채워졌으면 빈 칸도 남겨서 높이·폭을 맞춘다
    keep = set()
    for group in SIBLINGS.get(layout.name, []):
        if any(slots.get(n, (None,))[0] in used for n in group):
            keep |= {slots[n][0] for n in group if n in slots}
    for idx, shp in by_idx.items():
        if idx in used or not shp.has_text_frame or shp.text_frame.text.strip():
            continue
        if idx in keep:
            set_text(shp, "\u200b")      # 보이지 않는 글자: 칸(배경)은 유지, 편집 화면 안내문은 숨김
            for p in shp.text_frame.paragraphs:
                p._p.get_or_add_pPr().insert(0, parse_xml(f'<a:buNone {NSDECL}/>'))
        else:
            shp._element.getparent().remove(shp._element)
    _add_slide_number(slide, layout)
    if yangsik:
        balance(slide, prs)
    notes = spec.get("노트") or spec.get("notes")
    if notes:
        slide.notes_slide.notes_text_frame.text = notes
    return slide


def set_company(prs, company):
    """마스터·레이아웃의 '회사명' 글상자를 실제 회사명으로 (빈 문자열이면 지움)."""
    n = 0
    for owner in [prs.slide_masters[0]] + list(prs.slide_layouts):
        for shp in owner.shapes:
            if shp.name == "회사명" and shp.has_text_frame:
                runs = [r for p in shp.text_frame.paragraphs for r in p.runs]
                if runs:
                    runs[0].text = company
                    for r in runs[1:]:
                        r.text = ""
                    n += 1
    return n


def fill(template_path, content, out_path, base_dir=None, warn=print, company=None):
    prs = Presentation(str(template_path))
    company = company if company is not None else content.get("meta", {}).get("company")
    if company is not None:
        set_company(prs, company)
    delete_all_slides(prs)
    for i, spec in enumerate(content.get("slides", []), 1):
        try:
            add_slide(prs, spec, base_dir, warn)
        except KeyError as e:
            warn(f"  ! {i}번 슬라이드: {e}")
            raise
    meta = content.get("meta", {})
    if meta.get("title"):
        prs.core_properties.title = meta["title"]
    prs.save(str(out_path))
    return out_path
