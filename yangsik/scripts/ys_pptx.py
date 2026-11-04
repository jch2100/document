"""PPTX 엔진: 공통 마스터/레이아웃 생성 + 레이아웃 슬롯 채우기.

원칙
- 슬라이드에는 레이아웃의 플레이스홀더만 쓴다. 새 텍스트박스를 덧그리지 않는다.
- 서식(글꼴·크기·색)은 레이아웃/마스터가 가진다. 채우기는 텍스트만 넣는다.
"""
from __future__ import annotations

import copy
import io
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
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

SLIDE_W, SLIDE_H = 13.333, 7.5
M = 0.6                      # 좌우 여백
CW = SLIDE_W - 2 * M         # 콘텐츠 폭
BODY_Y, BODY_H = 2.05, 4.55  # 본문 영역

# 한국어 줄바꿈: 어절 단위(latinLnBrk=0), 동아시아 금칙 적용(eaLnBrk=1)
LNBRK = 'eaLnBrk="1" latinLnBrk="0" hangingPunct="1"'


# ---------------------------------------------------------------------------
# XML 조각
# ---------------------------------------------------------------------------
def _fill(color):
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


def _scheme(name):
    return f'<a:solidFill><a:schemeClr val="{name}"/></a:solidFill>'


def _rpr(sz=None, b=False, color=None, tag="a:defRPr"):
    attrs = 'lang="ko-KR" altLang="en-US"'
    if sz:
        attrs += f' sz="{int(sz * 100)}"'
    attrs += f' b="{1 if b else 0}"'
    fill = color if color is None or color.startswith("<") else _fill(color)
    return f"<{tag} {attrs}>{fill or ''}<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/></{tag}>"


def _lvl(n, sz, color=None, b=False, bullet=None, mar=0.0, ind=0.0, algn="l", spc_before=0, line=None):
    bu = "<a:buNone/>"
    if bullet == "auto":
        bu = '<a:buClr><a:schemeClr val="accent2"/></a:buClr><a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>'
    elif bullet:
        bu = f'<a:buClr><a:schemeClr val="accent2"/></a:buClr><a:buSzPct val="80000"/><a:buFont typeface="Arial"/><a:buChar char="{bullet}"/>'
    ln = f'<a:lnSpc><a:spcPct val="{int(line * 1000)}"/></a:lnSpc>' if line else ""
    sb = f'<a:spcBef><a:spcPts val="{int(spc_before * 100)}"/></a:spcBef>' if spc_before else '<a:spcBef><a:spcPts val="0"/></a:spcBef>'
    return (
        f'<a:lvl{n}pPr marL="{inch(mar)}" indent="{inch(ind)}" algn="{algn}" {LNBRK}>'
        f"{ln}{sb}{bu}{_rpr(sz, b, color)}</a:lvl{n}pPr>"
    )


BODY_LEVELS = lambda sz=16: (  # noqa: E731  본문 3단계: ■ / – / ·
    _lvl(1, sz, "<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>", bullet="■", mar=0.28, ind=-0.28, spc_before=10, line=115)
    + _lvl(2, sz - 2, "<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>", bullet="–", mar=0.58, ind=-0.25, spc_before=4, line=115)
    + _lvl(3, sz - 3, "<a:solidFill><a:schemeClr val=\"accent4\"/></a:solidFill>", bullet="·", mar=0.85, ind=-0.22, spc_before=2, line=115)
)


def ph(sid, name, ph_type, idx, x, y, w, h, levels, prompt, anchor="t", extra_sppr=""):
    """레이아웃 플레이스홀더 <p:sp>."""
    type_attr = f' type="{ph_type}"' if ph_type else ""
    idx_attr = f' idx="{idx}"' if idx is not None else ""
    sz_attr = ' sz="quarter"' if ph_type in ("dt", "ftr", "sldNum") else ""
    autofit = "<a:normAutofit/>" if ph_type not in ("pic", "tbl", "chart", "sldNum", "ftr", "dt") else ""
    if ph_type in ("pic", "tbl", "chart"):
        txbody = ""
    else:
        txbody = (
            f'<p:txBody><a:bodyPr vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}">{autofit}</a:bodyPr>'
            f"<a:lstStyle>{levels}</a:lstStyle>"
            f'<a:p><a:r><a:rPr lang="ko-KR" altLang="en-US"/><a:t>{escape(prompt)}</a:t></a:r></a:p></p:txBody>'
        )
        if ph_type == "sldNum":
            txbody = (
                f'<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle>{levels}</a:lstStyle>'
                '<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="ko-KR" altLang="en-US"/><a:t>‹#›</a:t></a:fld></a:p></p:txBody>'
            )
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{escape(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f"<p:nvPr><p:ph{type_attr}{sz_attr}{idx_attr}/></p:nvPr></p:nvSpPr>"
        f'<p:spPr><a:xfrm><a:off x="{inch(x)}" y="{inch(y)}"/><a:ext cx="{inch(w)}" cy="{inch(h)}"/></a:xfrm>{extra_sppr}</p:spPr>'
        f"{txbody}</p:sp>"
    )


def rect(sid, name, x, y, w, h, fill, line=None):
    ln = f'<a:ln w="{line[1]}">{_fill(line[0])}</a:ln>' if line else "<a:ln><a:noFill/></a:ln>"
    fl = _fill(fill) if fill and not fill.startswith("<") else (fill or "<a:noFill/>")
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{inch(x)}" y="{inch(y)}"/><a:ext cx="{inch(w)}" cy="{inch(h)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fl}{ln}</p:spPr>'
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


# ---------------------------------------------------------------------------
# 레이아웃 정의 (11종) — 이름이 곧 API. 바꾸면 profile.yaml도 바뀐다.
# ---------------------------------------------------------------------------
def _title_ph(sid=2, color="<a:solidFill><a:schemeClr val=\"tx2\"/></a:solidFill>"):
    return ph(sid, "제목", "title", None, M, 0.42, CW, 0.72, _lvl(1, 26, color, b=True), "제목을 입력하세요 (한 줄)", anchor="b")


def _msg_ph(sid=3):
    return ph(
        sid, "핵심메시지", "body", 13, M, 1.24, CW, 0.62,
        _lvl(1, 16, "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>", b=True, line=110),
        "이 장의 결론 한 문장 (두괄식)",
    )


def _accent_bar(sid=20):
    return rect(sid, "강조선", M, 0.3, 0.55, 0.06, "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>")


def _src_ph(sid, idx=19):
    return ph(sid, "출처", "body", idx, M, 6.62, CW, 0.26,
              _lvl(1, 10, "<a:solidFill><a:schemeClr val=\"accent4\"/></a:solidFill>"), "출처: ")


def _sldnum(sid=30):
    return ph(sid, "슬라이드 번호", "sldNum", 12, SLIDE_W - M - 0.8, 7.02, 0.8, 0.28,
              _lvl(1, 9, "<a:solidFill><a:schemeClr val=\"accent4\"/></a:solidFill>", algn="r"), "")


def layout_defs(theme):
    C = theme["colors"]
    white = "<a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill>"
    tx2 = "<a:solidFill><a:schemeClr val=\"tx2\"/></a:solidFill>"
    muted = "<a:solidFill><a:schemeClr val=\"accent4\"/></a:solidFill>"
    defs = []

    # 1 표지
    defs.append(dict(name="표지", master_shapes=False, bg=None, shapes=[
        rect(20, "좌측띠", 0, 0, 0.32, SLIDE_H, "<a:solidFill><a:schemeClr val=\"tx2\"/></a:solidFill>"),
        rect(21, "강조선", 1.1, 2.35, 0.9, 0.08, "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>"),
        ph(2, "제목", "ctrTitle", None, 1.1, 2.6, 10.8, 1.5, _lvl(1, 38, tx2, b=True, line=105), "문서 제목", anchor="t"),
        ph(3, "부제", "subTitle", 1, 1.1, 4.2, 10.8, 0.9, _lvl(1, 18, muted, line=110), "부제 또는 한 줄 요약"),
        ph(4, "날짜·작성자", "body", 14, 1.1, 6.2, 8.0, 0.5, _lvl(1, 13, muted), "2026. 11. 4. | 소속 · 작성자"),
        textbox(22, "회사명", SLIDE_W - M - 4, 6.2, 4, 0.5, theme["company"], 13, C["primary"], b=True, algn="r"),
    ]))
    # 2 목차
    defs.append(dict(name="목차", master_shapes=True, bg=None, shapes=[
        _accent_bar(), _title_ph(),
        ph(4, "본문", "body", 1, M + 0.2, BODY_Y - 0.4, CW - 0.4, BODY_H + 0.4,
           _lvl(1, 20, "<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>", bullet="auto", mar=0.55, ind=-0.55, spc_before=14), "목차 항목"),
    ]))
    # 3 섹션
    defs.append(dict(name="섹션", master_shapes=False, bg=C["primary"], shapes=[
        ph(3, "번호", "body", 13, 1.1, 2.3, 4, 1.0, _lvl(1, 44, "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>", b=True), "01", anchor="b"),
        ph(2, "제목", "title", None, 1.1, 3.4, 11, 1.2, _lvl(1, 34, white, b=True), "섹션 제목"),
        ph(4, "설명", "body", 14, 1.1, 4.6, 11, 0.8, _lvl(1, 16, "<a:solidFill><a:srgbClr val=\"C9D3E0\"/></a:solidFill>"), "이 섹션에서 다룰 내용 한 줄"),
    ]))
    # 4 본문
    defs.append(dict(name="본문", master_shapes=True, bg=None, shapes=[
        _accent_bar(), _title_ph(), _msg_ph(),
        ph(4, "본문", "body", 1, M, BODY_Y, CW, BODY_H, BODY_LEVELS(17), "개조식 항목 (명사형 종결)"),
    ]))
    # 5 2단비교
    cw2 = (CW - 0.4) / 2
    defs.append(dict(name="2단비교", master_shapes=True, bg=None, shapes=[
        _accent_bar(), _title_ph(), _msg_ph(),
        rect(21, "좌측머리선", M, BODY_Y + 0.55, cw2, 0.03, "<a:solidFill><a:schemeClr val=\"tx2\"/></a:solidFill>"),
        rect(22, "우측머리선", M + cw2 + 0.4, BODY_Y + 0.55, cw2, 0.03, "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>"),
        ph(4, "좌측제목", "body", 15, M, BODY_Y, cw2, 0.5, _lvl(1, 18, tx2, b=True), "좌측 제목 (예: 현재)", anchor="b"),
        ph(5, "좌측본문", "body", 16, M, BODY_Y + 0.75, cw2, BODY_H - 0.75, BODY_LEVELS(15), "좌측 내용"),
        ph(6, "우측제목", "body", 17, M + cw2 + 0.4, BODY_Y, cw2, 0.5, _lvl(1, 18, "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>", b=True), "우측 제목 (예: 개선 후)", anchor="b"),
        ph(7, "우측본문", "body", 18, M + cw2 + 0.4, BODY_Y + 0.75, cw2, BODY_H - 0.75, BODY_LEVELS(15), "우측 내용"),
    ]))
    # 6 3단카드
    cw3 = (CW - 0.6) / 3
    shapes = [_accent_bar(), _title_ph(), _msg_ph()]
    for i in range(3):
        x = M + i * (cw3 + 0.3)
        shapes.append(rect(21 + i, f"카드{i+1}배경", x, BODY_Y, cw3, BODY_H, C["light"]))
        shapes.append(rect(24 + i, f"카드{i+1}선", x, BODY_Y, cw3, 0.06, "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>"))
        shapes.append(ph(4 + i * 2, f"카드{i+1}제목", "body", 20 + i * 2, x + 0.25, BODY_Y + 0.3, cw3 - 0.5, 0.85,
                         _lvl(1, 20, tx2, b=True, line=105), f"카드 {i+1} 제목 또는 숫자", anchor="t"))
        shapes.append(ph(5 + i * 2, f"카드{i+1}본문", "body", 21 + i * 2, x + 0.25, BODY_Y + 1.25, cw3 - 0.5, BODY_H - 1.5,
                         BODY_LEVELS(14), f"카드 {i+1} 설명"))
    defs.append(dict(name="3단카드", master_shapes=True, bg=None, shapes=shapes))
    # 7 표
    defs.append(dict(name="표", master_shapes=True, bg=None, shapes=[
        _accent_bar(), _title_ph(), _msg_ph(),
        ph(4, "표", "tbl", 1, M, BODY_Y, CW, BODY_H - 0.1, "", "표를 넣으세요"),
        _src_ph(5),
    ]))
    # 8 차트
    defs.append(dict(name="차트", master_shapes=True, bg=None, shapes=[
        _accent_bar(), _title_ph(), _msg_ph(),
        ph(4, "차트", "chart", 1, M, BODY_Y, CW * 0.66, BODY_H - 0.1, "", "차트를 넣으세요"),
        ph(5, "해설", "body", 14, M + CW * 0.66 + 0.35, BODY_Y + 0.2, CW * 0.34 - 0.35, BODY_H - 0.4, BODY_LEVELS(14), "차트에서 읽어야 할 점"),
        _src_ph(6),
    ]))
    # 9 핵심메시지
    defs.append(dict(name="핵심메시지", master_shapes=True, bg=C["light"], shapes=[
        rect(20, "강조선", M + 0.6, 2.35, 0.9, 0.08, "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>"),
        ph(2, "제목", "title", None, M + 0.6, 1.6, CW - 1.2, 0.6, _lvl(1, 16, muted, b=True), "요약", anchor="b"),
        ph(3, "핵심메시지", "body", 13, M + 0.6, 2.65, CW - 1.2, 2.2, _lvl(1, 32, tx2, b=True, line=120), "가장 중요한 결론 한두 문장"),
        ph(4, "보조설명", "body", 14, M + 0.6, 5.0, CW - 1.2, 1.2, _lvl(1, 16, muted, line=120), "근거 또는 요청 사항 한 줄"),
    ]))
    # 10 이미지
    defs.append(dict(name="이미지", master_shapes=True, bg=None, shapes=[
        _accent_bar(), _title_ph(), _msg_ph(),
        ph(4, "이미지", "pic", 1, M, BODY_Y, CW * 0.62, BODY_H, "", "그림을 넣으세요"),
        ph(5, "캡션", "body", 14, M + CW * 0.62 + 0.35, BODY_Y + 0.1, CW * 0.38 - 0.35, BODY_H - 0.2, BODY_LEVELS(14), "그림 설명"),
    ]))
    # 11 마무리
    defs.append(dict(name="마무리", master_shapes=False, bg=C["primary"], shapes=[
        ph(2, "제목", "title", None, 1.1, 2.4, 11, 1.2, _lvl(1, 36, white, b=True), "감사합니다", anchor="b"),
        rect(20, "강조선", 1.1, 3.8, 0.9, 0.08, "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>"),
        ph(3, "본문", "body", 1, 1.1, 4.1, 11, 2.2, _lvl(1, 16, "<a:solidFill><a:srgbClr val=\"DCE3EC\"/></a:solidFill>", line=130, spc_before=4), "문의: 이름 · 이메일 · 전화"),
    ]))
    for d in defs:
        if d["master_shapes"]:
            d["shapes"].append(_sldnum())
    return defs


# ---------------------------------------------------------------------------
# 베이스 양식 빌드
# ---------------------------------------------------------------------------
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
               "accent4": C["muted"], "accent5": "8FA9C4", "accent6": "C7D2DE",
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
        if sp.tag.endswith("}sp"):
            tree.remove(sp)
    xmls = [
        ph(2, "제목", "title", None, M, 0.42, CW, 0.72, "", "제목", anchor="b"),
        ph(3, "본문", "body", 1, M, BODY_Y, CW, BODY_H, "", "본문"),
        rect(10, "하단선", M, 6.95, CW, 0.012, C["light"] if C["light"] != "FFFFFF" else "DDDDDD"),
        textbox(11, "회사명", M, 7.02, 6, 0.28, theme["company"], 9, C["muted"], b=True),
        _sldnum(12),
    ]
    for x in xmls:
        tree.append(parse_xml(f'<p:wrap {NSDECL}>{x}</p:wrap>')[0])
    # 텍스트 스타일
    tx = el.find("p:txStyles", NS)
    tx.getparent().remove(tx)
    tx_xml = (
        f'<p:txStyles {NSDECL}>'
        f'<p:titleStyle>{_lvl(1, 26, "<a:solidFill><a:schemeClr val=\"tx2\"/></a:solidFill>", b=True)}</p:titleStyle>'
        f'<p:bodyStyle>{BODY_LEVELS(17)}'
        + "".join(_lvl(n, 12, "<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>", mar=1.0 + n * 0.2) for n in range(4, 10))
        + '</p:bodyStyle>'
        f'<p:otherStyle>{_lvl(1, 12, "<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>")}</p:otherStyle>'
        '</p:txStyles>'
    )
    # txStyles는 clrMap, sldLayoutIdLst 뒤
    idx = list(el).index(el.find("p:sldLayoutIdLst", NS)) + 1
    el.insert(idx, parse_xml(tx_xml))
    # 기본 텍스트(표·차트·새 텍스트박스)도 한국어
    pres = prs.part._element
    dts = pres.find("p:defaultTextStyle", NS)
    if dts is not None:
        for d in dts.iter(f"{{{NS['a']}}}defRPr"):
            d.set("lang", "ko-KR")
            d.set("altLang", "en-US")


def _write_layout(layout, d):
    bg = (f'<p:bg><p:bgPr>{_fill(d["bg"])}<a:effectLst/></p:bgPr></p:bg>' if d["bg"] else "")
    show = "" if d["master_shapes"] else ' showMasterSp="0"'
    xml = (
        f'<p:sldLayout {NSDECL} preserve="1" userDrawn="1"{show}>'
        f'<p:cSld name="{escape(d["name"])}">{bg}<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + "".join(d["shapes"]) +
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
    )
    layout.part._element = parse_xml(xml.encode("utf-8"))


def build_base(theme) -> bytes:
    prs = Presentation()
    prs.slide_width, prs.slide_height = inch(SLIDE_W), inch(SLIDE_H)
    _set_theme(prs, theme)
    _set_master(prs, theme)
    defs = layout_defs(theme)
    layouts = list(prs.slide_layouts)
    assert len(layouts) >= len(defs)
    for layout, d in zip(layouts, defs):
        _write_layout(layout, d)
    for extra in layouts[len(defs):]:
        prs.slide_layouts.remove(extra)
    core = prs.core_properties
    core.title, core.author, core.language = "양식대로 표준 양식", theme["company"], "ko-KR"
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 채우기
# ---------------------------------------------------------------------------
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
    tf = shape.text_frame
    items = flatten_items(value) if not isinstance(value, str) else [(0, value)]
    # 기존 문단 제거
    txBody = tf._txBody
    for p in list(txBody.findall(f"{{{NS['a']}}}p")):
        txBody.remove(p)
    for level, text in items:
        p = tf._txBody.add_p()
        from pptx.text.text import _Paragraph
        _set_runs(_Paragraph(p, tf), text, level)
    if not items:
        tf._txBody.add_p()


def _style_table(table, sizes=(13, 12)):
    for r_i, row in enumerate(table.rows):
        for cell in row.cells:
            cell.margin_left = cell.margin_right = inch(0.1)
            cell.margin_top = cell.margin_bottom = inch(0.05)
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(sizes[0] if r_i == 0 else sizes[1])


def fill_table(ph_shape, spec):
    ph_name = ph_shape.name
    header = spec.get("header") or spec.get("머리글") or []
    rows = spec.get("rows") or spec.get("행") or []
    n_cols = max([len(header)] + [len(r) for r in rows]) if (header or rows) else 1
    n_rows = len(rows) + (1 if header else 0)
    gf = ph_shape.insert_table(max(n_rows, 1), n_cols)
    table = gf.table
    table.first_row = bool(header)
    all_rows = ([header] if header else []) + rows
    for r_i, row in enumerate(all_rows):
        for c_i in range(n_cols):
            val = row[c_i] if c_i < len(row) else ""
            cell = table.cell(r_i, c_i)
            cell.text_frame.text = ""
            p = cell.text_frame.paragraphs[0]
            for k, line in enumerate(str(val).split("\n")):
                if k:
                    p = cell.text_frame.add_paragraph()
                _set_runs(p, line)
    widths = spec.get("col_widths") or spec.get("열너비")
    total = gf.width
    if widths:
        s = sum(widths)
        for i, w in enumerate(widths[:n_cols]):
            table.columns[i].width = Emu(int(total * w / s))
    # 행 높이: 가능한 한 컴팩트하게 (내용이 늘어나면 PowerPoint가 자동 확장)
    row_h = min(inch(0.55), int(gf.height / max(n_rows, 1)))
    for row in table.rows:
        row.height = Emu(row_h)
    _style_table(table)
    gf.name = ph_name
    return gf


CHART_TYPES = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "hbar": XL_CHART_TYPE.BAR_CLUSTERED, "stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "line": XL_CHART_TYPE.LINE_MARKERS, "pie": XL_CHART_TYPE.PIE, "doughnut": XL_CHART_TYPE.DOUGHNUT,
}


def fill_chart(ph_shape, spec):
    ph_name = ph_shape.name
    cd = CategoryChartData()
    cd.categories = spec.get("categories") or spec.get("항목")
    series = spec.get("series") or spec.get("계열") or {}
    for name, vals in series.items():
        cd.add_series(name, vals)
    ctype = CHART_TYPES.get(spec.get("type", "bar"), XL_CHART_TYPE.COLUMN_CLUSTERED)
    gf = ph_shape.insert_chart(ctype, cd)
    ch = gf.chart
    ch.font.size = Pt(12)
    ch.has_legend = len(series) > 1 or ctype in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT)
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
    if spec.get("labels", True):
        plot = ch.plots[0]
        plot.has_data_labels = True
        plot.data_labels.font.size = Pt(11)
        if spec.get("number_format"):
            plot.data_labels.number_format = spec["number_format"]
            plot.data_labels.number_format_is_linked = False
    if ctype not in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
        try:
            ch.value_axis.has_major_gridlines = True
            ch.value_axis.major_gridlines.format.line.color.rgb = __import__("pptx").dml.color.RGBColor(0xE5, 0xE7, 0xEB)
            ch.value_axis.format.line.fill.background()
        except Exception:
            pass
    gf.name = ph_name
    return gf


def _add_slide_number(slide, layout):
    for shp in layout.placeholders:
        if "SLIDE_NUMBER" in str(shp.placeholder_format.type):
            el = copy.deepcopy(shp._element)
            # 레이아웃 전용 속성 제거
            slide.shapes._spTree.append(el)
            return


def delete_all_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.get(f"{{{NS['r']}}}id")
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)


def get_layout(prs, name):
    for l in prs.slide_layouts:
        if l.name == name:
            return l
    names = [l.name for l in prs.slide_layouts]
    raise KeyError(f"레이아웃 '{name}' 없음. 사용 가능: {names}")


SKIP_KEYS = {"layout", "레이아웃", "노트", "notes", "id"}


def add_slide(prs, spec, base_dir=None, warn=print):
    layout = get_layout(prs, spec.get("layout") or spec.get("레이아웃"))
    slots = layout_slots(layout)
    slide = prs.slides.add_slide(layout)
    by_idx = {shp.placeholder_format.idx: shp for shp in slide.placeholders}
    # 슬라이드 플레이스홀더 이름을 레이아웃 슬롯 이름으로
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
            fill_table(shp, value)
        elif t == "CHART":
            fill_chart(shp, value)
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
    # 내용 없는 텍스트 슬롯은 지운다 (발표 시 '텍스트를 입력하세요'가 남지 않게)
    for idx, shp in by_idx.items():
        if idx not in used and shp.has_text_frame and not shp.text_frame.text.strip():
            shp._element.getparent().remove(shp._element)
    _add_slide_number(slide, layout)
    notes = spec.get("노트") or spec.get("notes")
    if notes:
        slide.notes_slide.notes_text_frame.text = notes
    return slide


def fill(template_path, content, out_path, base_dir=None, warn=print):
    prs = Presentation(str(template_path))
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
