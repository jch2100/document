"""DOCX 엔진: 표준 스타일 양식 생성 + 블록(JSON) 채우기.

원칙
- 모든 서식은 '스타일'로만 준다. 본문에 글꼴·크기를 직접 박지 않는다(→ Word에서 스타일 한 번에 수정 가능).
- 개조식 기호(□ ○ - ※)와 제목 번호(1. 가. 1))는 Word '번호 매기기'로 자동 생성된다.
  → 내용에 기호/번호를 직접 쓰지 않는다. 채우기 스크립트가 붙어 온 기호는 자동 제거한다.
"""
from __future__ import annotations

import copy
import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn, nsdecls
from docx.shared import Pt, Cm, RGBColor, Emu

from ys_common import split_bold, flatten_items

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 스타일 이름 = API. profile.yaml 과 check.py 가 이 이름을 쓴다.
S_TITLE, S_SUB = "Title", "Subtitle"
S_H = {1: "Heading 1", 2: "Heading 2", 3: "Heading 3"}
S_B = {0: "개조식 1", 1: "개조식 2", 2: "개조식 3"}
S_NOTE = "참고"
S_SUMMARY = "요약 박스"
S_TABLE_TEXT = "표 본문"
S_CAPTION = "Caption"
S_COVER_META = "표지 정보"
S_TABLE = "양식대로 표"
S_TOC_HEAD = "목차 제목"

BULLET_SYMBOLS = ["□", "○", "-"]
STRIP_RE = re.compile(r"^\s*(?:[□■○●◦\-–·•※▪▶]|\d{1,2}[.)]|[가나다라마바사아자차카타파하][.)]|\(\d+\)|\([가나다라마바사아자차카타파하]\)|[①-⑳]|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.)\s+")


def _rgb(hexstr):
    return RGBColor.from_string(hexstr)


# ---------------------------------------------------------------------------
# 번호 매기기 (numbering.xml)
# ---------------------------------------------------------------------------
def _lvl_xml(ilvl, fmt, text, left, hanging, font=None, sz=None, color=None, pstyle=None, start=1):
    rpr = ""
    if font or sz or color:
        rpr = "<w:rPr>"
        if font:
            rpr += f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:hint="eastAsia"/>'
        if color:
            rpr += f'<w:color w:val="{color}"/>'
        if sz:
            rpr += f'<w:sz w:val="{int(sz*2)}"/>'
        rpr += "</w:rPr>"
    ps = f'<w:pStyle w:val="{pstyle}"/>' if pstyle else ""
    return (
        f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/>{ps}'
        f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr>{rpr}</w:lvl>'
    )


def _add_numbering(doc, theme):
    numbering = doc.part.numbering_part.element
    font = theme["fonts"]["korean"]
    C = theme["colors"]
    # 900: 제목 번호 1. → 가. → 1)
    heading_lvls = (
        _lvl_xml(0, "decimal", "%1.", 400, 400, pstyle="Heading1")
        + _lvl_xml(1, "ganada", "%2.", 600, 400, pstyle="Heading2")
        + _lvl_xml(2, "decimal", "%3)", 800, 400, pstyle="Heading3")
    )
    # 901: 개조식 □ → ○ → -
    bullet_lvls = (
        _lvl_xml(0, "bullet", "□", 420, 320, font=font)
        + _lvl_xml(1, "bullet", "○", 760, 320, font=font)
        + _lvl_xml(2, "bullet", "-", 1080, 280, font=font)
    )
    # 902: 참고 ※
    note_lvls = _lvl_xml(0, "bullet", "※", 420, 320, font=font, color=C["muted"])
    for aid, lvls, kind in ((900, heading_lvls, "multilevel"), (901, bullet_lvls, "multilevel"), (902, note_lvls, "singleLevel")):
        xml = (f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{aid}">'
               f'<w:multiLevelType w:val="{kind}"/>{lvls}</w:abstractNum>')
        # abstractNum 은 num 보다 앞에 와야 한다
        first_num = numbering.find(qn("w:num"))
        el = parse_xml(xml)
        if first_num is not None:
            first_num.addprevious(el)
        else:
            numbering.append(el)
    for nid, aid in ((900, 900), (901, 901), (902, 902)):
        numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{nid}"><w:abstractNumId w:val="{aid}"/></w:num>'))


def _set_numpr(style, num_id, ilvl):
    pPr = style.element.get_or_add_pPr()
    for old in pPr.findall(qn("w:numPr")):
        pPr.remove(old)
    numPr = parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr>')
    pPr.insert(0, numPr)


# ---------------------------------------------------------------------------
# 스타일
# ---------------------------------------------------------------------------
def _fonts(rpr_owner_font, latin, korean):
    rpr = rpr_owner_font.element.get_or_add_rPr() if hasattr(rpr_owner_font, "element") else rpr_owner_font
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for k in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rf.get(qn(k)) is not None:
            del rf.attrib[qn(k)]
    rf.set(qn("w:ascii"), latin)
    rf.set(qn("w:hAnsi"), latin)
    rf.set(qn("w:eastAsia"), korean)
    rf.set(qn("w:hint"), "eastAsia")


def _pstyle(doc, name, base="Normal", size=None, bold=None, color=None, align=None,
            before=None, after=None, line=None, keep_next=False, left=None):
    styles = doc.styles
    try:
        st = styles[name]
    except KeyError:
        st = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = styles[base]
        st.quick_style = True
    f = st.font
    if size:
        f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    if color:
        f.color.rgb = _rgb(color)
    pf = st.paragraph_format
    if align:
        pf.alignment = align
    if before is not None:
        pf.space_before = Pt(before)
    if after is not None:
        pf.space_after = Pt(after)
    if line:
        pf.line_spacing = line
    if keep_next:
        pf.keep_with_next = True
    if left is not None:
        pf.left_indent = Cm(left)
    return st


def _border_xml(sides, color, sz=6, space=4):
    return "<w:pBdr>" + "".join(
        f'<w:{s} w:val="single" w:sz="{sz}" w:space="{space}" w:color="{color}"/>' for s in sides) + "</w:pBdr>"


def _apply_ppr_xml(style, xml):
    pPr = style.element.get_or_add_pPr()
    el = parse_xml(f'<w:x {nsdecls("w")}>{xml}</w:x>')
    for child in el:
        for old in pPr.findall(child.tag):
            pPr.remove(old)
        pPr.append(child)


def _build_styles(doc, theme):
    C, F = theme["colors"], theme["fonts"]
    lat, kor = F["latin"], F["korean"]
    # 문서 기본값: 한국어 + 동아시아 글꼴
    rpr_default = doc.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    _fonts(rpr_default, lat, kor)
    lang = rpr_default.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr_default.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "ko-KR")

    normal = doc.styles["Normal"]
    _fonts(normal, lat, kor)
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = _rgb(C["text"])
    pf = normal.paragraph_format
    pf.space_after = Pt(3)
    pf.line_spacing = 1.3
    # 한국어 단어 단위 줄바꿈 (Word: '한글 입력 체계에서 단어 잘림 허용' 해제와 같은 효과)
    _apply_ppr_xml(normal, '<w:wordWrap w:val="1"/><w:autoSpaceDE w:val="1"/><w:autoSpaceDN w:val="1"/>')

    t = _pstyle(doc, S_TITLE, size=20, bold=True, color=C["primary"], align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=4, line=1.15)
    _apply_ppr_xml(t, "<w:pBdr/>")
    _pstyle(doc, S_SUB, size=11.5, bold=False, color=C["muted"], align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=14)
    doc.styles[S_SUB].font.italic = False
    for st_name in (S_TITLE, S_SUB):
        _fonts(doc.styles[st_name], lat, kor)
        doc.styles[st_name].font.italic = False

    hdefs = {1: (14, C["primary"], 16, 6), 2: (12, C["text"], 10, 4), 3: (11, C["text"], 6, 2)}
    for lvl, (sz, col, b4, af) in hdefs.items():
        h = _pstyle(doc, S_H[lvl], size=sz, bold=True, color=col, before=b4, after=af, line=1.2, keep_next=True)
        _fonts(h, lat, kor)
        h.font.italic = False
        _set_numpr(h, 900, lvl - 1)
        if lvl == 1:
            _apply_ppr_xml(h, _border_xml(["bottom"], C["primary"], sz=8, space=3))

    for lvl, name in S_B.items():
        b = _pstyle(doc, name, size=10.5 if lvl == 0 else 10.5, before=4 if lvl == 0 else 1, after=1, line=1.3)
        _set_numpr(b, 901, lvl)
    note = _pstyle(doc, S_NOTE, size=9.5, color=C["muted"], before=2, after=2)
    _set_numpr(note, 902, 0)

    s = _pstyle(doc, S_SUMMARY, size=10.5, bold=False, before=6, after=10, line=1.35)
    _apply_ppr_xml(s, _border_xml(["top", "left", "bottom", "right"], C["accent"], sz=8, space=6)
                   + f'<w:shd w:val="clear" w:color="auto" w:fill="{C["light"]}"/>'
                   + '<w:ind w:left="140" w:right="140"/>')

    _pstyle(doc, S_TABLE_TEXT, size=9.5, before=0, after=0, line=1.15)
    cap = _pstyle(doc, S_CAPTION, size=9, bold=False, color=C["muted"], before=2, after=8)
    cap.font.italic = False
    th = _pstyle(doc, S_TOC_HEAD, size=14, bold=True, color=C["primary"], before=6, after=10, keep_next=True)
    _apply_ppr_xml(th, _border_xml(["bottom"], C["primary"], sz=8, space=3))
    _pstyle(doc, S_COVER_META, size=12, color=C["text"], align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=4)

    # 표 스타일 (Word에서 '표 디자인'으로 보임)
    try:
        ts = doc.styles[S_TABLE]
    except KeyError:
        ts = doc.styles.add_style(S_TABLE, WD_STYLE_TYPE.TABLE)
    ts.element.append(parse_xml(
        f'<w:tblPr {nsdecls("w")}><w:tblBorders>'
        f'<w:top w:val="single" w:sz="8" w:color="{C["primary"]}"/>'
        f'<w:bottom w:val="single" w:sz="8" w:color="{C["primary"]}"/>'
        f'<w:insideH w:val="single" w:sz="4" w:color="D0D5DD"/>'
        f'<w:insideV w:val="single" w:sz="4" w:color="D0D5DD"/>'
        f'</w:tblBorders><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>'))
    ts.element.append(parse_xml(
        f'<w:tblStylePr {nsdecls("w")} w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>'
        f'<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{C["primary"]}"/></w:tcPr></w:tblStylePr>'))


def _header_footer(doc, theme, doc_label):
    C = theme["colors"]
    sec = doc.sections[0]
    sec.header_distance = Cm(1.0)
    sec.footer_distance = Cm(1.0)
    hp = sec.header.paragraphs[0]
    for r in list(hp.runs):
        r._r.getparent().remove(r._r)
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = hp.add_run(f"{theme['company']}  |  {doc_label}")
    r.font.size = Pt(8.5)
    r.font.color.rgb = _rgb(C["muted"])
    hp._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="4" w:color="D0D5DD"/></w:pBdr>'))
    fp = sec.footer.paragraphs[0]
    for r in list(fp.runs):
        r._r.getparent().remove(r._r)
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for part in ("- ", None, " -"):
        if part is None:
            fp._p.append(parse_xml(
                f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"><w:r><w:rPr><w:sz w:val="17"/><w:color w:val="{C["muted"]}"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>'))
        else:
            rr = fp.add_run(part)
            rr.font.size = Pt(8.5)
            rr.font.color.rgb = _rgb(C["muted"])


def build_base(theme, doc_label="문서") -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)  # A4
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.top_margin, sec.bottom_margin = Cm(2.0), Cm(1.8)
    sec.left_margin = sec.right_margin = Cm(2.0)
    _build_styles(doc, theme)
    _add_numbering(doc, theme)
    _header_footer(doc, theme, doc_label)
    cp = doc.core_properties
    cp.title, cp.author, cp.language = doc_label, theme["company"], "ko-KR"
    # 빈 첫 문단 제거
    body = doc.element.body
    for p in body.findall(qn("w:p")):
        body.remove(p)
    return doc


# ---------------------------------------------------------------------------
# 채우기
# ---------------------------------------------------------------------------
def clean(text):
    return STRIP_RE.sub("", str(text), count=1)


_MISSING = set()


def _style_or_normal(doc, style):
    try:
        doc.styles[style]
        return style
    except KeyError:
        if style not in _MISSING:
            _MISSING.add(style)
            print(f"  ! 양식에 '{style}' 스타일이 없어 '본문(Normal)'으로 넣습니다 → normalize.py 로 스타일 추가 권장")
        return "Normal"


def add_par(doc, text, style, strip=True):
    p = doc.add_paragraph(style=_style_or_normal(doc, style))
    for seg, bold in split_bold(clean(text) if strip else text):
        r = p.add_run(seg)
        if bold:
            r.bold = True
    return p


def _cell_text(cell, value, header=False, align=None):
    cell.text = ""
    lines = str(value).split("\n")
    p = cell.paragraphs[0]
    for i, line in enumerate(lines):
        if i:
            p = cell.add_paragraph()
        try:
            p.style = S_TABLE_TEXT
        except (KeyError, ValueError):
            pass
        if align:
            p.alignment = align
        for seg, bold in split_bold(line):
            r = p.add_run(seg)
            if bold or header:
                r.bold = True


def _shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def _set_grid(table, widths_cm):
    """열 너비를 Word·LibreOffice·Google 모두 따르도록 tblGrid/tcW/tblW 를 함께 설정."""
    tbl = table._tbl
    grid = tbl.tblGrid
    cols = grid.findall(qn("w:gridCol"))
    for gc, w in zip(cols, widths_cm):
        gc.set(qn("w:w"), str(int(w * 567)))
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths_cm):
                cell.width = Cm(widths_cm[i])
    tblPr = tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(int(sum(widths_cm) * 567)))
    tblW.set(qn("w:type"), "dxa")
    lay = tblPr.find(qn("w:tblLayout"))
    if lay is None:
        lay = OxmlElement("w:tblLayout")
        tblPr.append(lay)
    lay.set(qn("w:type"), "fixed")


def _col_widths(table, widths, total_cm):
    s = sum(widths)
    _set_grid(table, [total_cm * w / s for w in widths])


def _table_style(doc, t):
    for name in (S_TABLE, "Table Grid", "표 구분선"):
        try:
            t.style = doc.styles[name]
            return
        except KeyError:
            continue


def _content_width_cm(doc):
    sec = doc.sections[-1]
    try:
        return (sec.page_width - sec.left_margin - sec.right_margin) / 360000
    except Exception:
        return 17.0


def add_table(doc, spec, theme, total_cm=None):
    total_cm = total_cm or _content_width_cm(doc)
    C = theme["colors"]
    header = spec.get("header") or spec.get("머리글") or []
    rows = spec.get("rows") or spec.get("행") or []
    ncol = max([len(header)] + [len(r) for r in rows] + [1])
    t = doc.add_table(rows=0, cols=ncol)
    _table_style(doc, t)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    if header:
        cells = t.add_row().cells
        for i in range(ncol):
            _cell_text(cells[i], header[i] if i < len(header) else "", header=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            _shade(cells[i], C["primary"])
            for r in cells[i].paragraphs[0].runs:
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        # 머리글 행 반복
        trPr = t.rows[0]._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    first_col_label = spec.get("label_col", False)
    for row in rows:
        cells = t.add_row().cells
        for i in range(ncol):
            _cell_text(cells[i], row[i] if i < len(row) else "")
            if first_col_label and i == 0:
                _shade(cells[i], C["light"])
                for p in cells[i].paragraphs:
                    for r in p.runs:
                        r.bold = True
    widths = spec.get("col_widths") or spec.get("열너비") or [1] * ncol
    _col_widths(t, widths[:ncol], total_cm)
    if spec.get("caption") or spec.get("캡션"):
        add_par(doc, spec.get("caption") or spec.get("캡션"), S_CAPTION, strip=False)
    else:
        sp = doc.add_paragraph(style="Normal")
        sp.paragraph_format.space_after = Pt(4)
    return t


def add_kv_table(doc, pairs, theme, cols=2):
    """[['보고일','2026. 11. 4.'], ['보고자','홍길동'], ...] → 라벨|값|라벨|값 표."""
    C = theme["colors"]
    per_row = cols
    t = doc.add_table(rows=0, cols=per_row * 2)
    _table_style(doc, t)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i in range(0, len(pairs), per_row):
        cells = t.add_row().cells
        for j in range(per_row):
            k, v = (pairs[i + j] if i + j < len(pairs) else ("", ""))
            _cell_text(cells[j * 2], k, header=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            _shade(cells[j * 2], C["light"])
            _cell_text(cells[j * 2 + 1], v)
    # 한 쌍이 한 행을 다 쓰는 경우(값이 김) 병합
    widths = [2.2, 6.3] * per_row
    _col_widths(t, widths, 17.0)
    # 첫 행 스타일 조건(머리글색) 끄기
    tblPr = t._tbl.tblPr
    look = tblPr.find(qn("w:tblLook"))
    if look is None:
        look = OxmlElement("w:tblLook")
        tblPr.append(look)
    look.set(qn("w:firstRow"), "0")
    look.set(qn("w:val"), "0000")
    sp = doc.add_paragraph(style="Normal")
    sp.paragraph_format.space_after = Pt(2)
    return t


def add_approval(doc, roles, theme):
    """결재란: 우측 정렬 작은 표 (직위 / 서명칸)."""
    C = theme["colors"]
    n = len(roles)
    t = doc.add_table(rows=2, cols=n + 1)
    _table_style(doc, t)
    t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    tblPr = t._tbl.tblPr
    look = OxmlElement("w:tblLook")
    look.set(qn("w:firstRow"), "0")
    look.set(qn("w:val"), "0000")
    tblPr.append(look)
    a = t.cell(0, 0).merge(t.cell(1, 0))
    _cell_text(a, "결\n재", header=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _shade(a, C["light"])
    for i, role in enumerate(roles, 1):
        _cell_text(t.cell(0, i), role, header=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _shade(t.cell(0, i), C["light"])
        _cell_text(t.cell(1, i), "")
    t.rows[1].height = Cm(1.5)
    _set_grid(t, [0.9] + [2.0] * n)
    sp = doc.add_paragraph(style="Normal")
    sp.paragraph_format.space_after = Pt(6)
    return t


def add_toc(doc, heading="목차"):
    if isinstance(heading, str) and heading:
        add_par(doc, heading, S_TOC_HEAD, strip=False)
    p = doc.add_paragraph(style="Normal")
    r = p.add_run()
    r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'))
    r2 = p.add_run()
    r2._r.append(parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve">TOC \\o "1-2" \\h \\z \\u</w:instrText>'))
    r3 = p.add_run()
    r3._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'))
    p.add_run("(목차: Word에서 이 자리를 마우스 오른쪽 → 필드 업데이트)")
    r4 = p.add_run()
    r4._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'))


def add_cover(doc, spec, theme):
    """표지: 빈 줄 대신 '문단 앞 간격'으로 배치 (빈 문단 반복은 check.py가 경고함)."""
    tp = add_par(doc, spec.get("title", ""), S_TITLE, strip=False)
    tp.paragraph_format.space_before = Pt(200)
    for r in tp.runs:
        r.font.size = Pt(28)
    if spec.get("subtitle"):
        add_par(doc, spec["subtitle"], S_SUB, strip=False)
    first = True
    for key in ("date", "company", "author"):
        if spec.get(key):
            mp = add_par(doc, spec[key], S_COVER_META, strip=False)
            if first:
                mp.paragraph_format.space_before = Pt(220)
                first = False
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def _block_type(block):
    if "type" in block:
        return block["type"], block.get("text", block.get("value"))
    for k, v in block.items():
        if k not in ("caption", "캡션", "width_cm"):
            return k, v
    return None, None


ALIASES = {"제목": "title", "부제": "subtitle", "대제목": "h1", "중제목": "h2", "소제목": "h3",
           "문단": "p", "개조식": "b", "참고": "note", "요약": "summary", "표": "table",
           "정보표": "meta", "결재": "approval", "쪽나눔": "pagebreak", "목차": "toc", "표지": "cover",
           "그림": "image"}


def add_block(doc, block, theme, base_dir=None, warn=print):
    kind, val = _block_type(block)
    kind = ALIASES.get(kind, kind)
    if kind == "title":
        add_par(doc, val, S_TITLE, strip=False)
    elif kind == "subtitle":
        add_par(doc, val, S_SUB, strip=False)
    elif kind in ("h1", "h2", "h3"):
        add_par(doc, val, S_H[int(kind[1])])
    elif kind == "p":
        for line in ([val] if isinstance(val, str) else val):
            add_par(doc, line, "Normal", strip=False)
    elif kind == "b":
        for level, text in flatten_items(val):
            add_par(doc, text, S_B[min(level, 2)])
    elif kind == "note":
        for line in ([val] if isinstance(val, str) else val):
            add_par(doc, line, S_NOTE)
    elif kind == "summary":
        items = [val] if isinstance(val, str) else val
        for level, text in flatten_items(items):
            add_par(doc, ("· " if level else "") + clean(text), S_SUMMARY, strip=False)
    elif kind == "table":
        spec = dict(val)
        if block.get("caption"):
            spec["caption"] = block["caption"]
        add_table(doc, spec, theme)
    elif kind == "meta":
        add_kv_table(doc, val, theme, cols=block.get("cols", 2))
    elif kind == "approval":
        add_approval(doc, val, theme)
    elif kind == "pagebreak":
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    elif kind == "toc":
        add_toc(doc, val if isinstance(val, str) else "목차")
    elif kind == "cover":
        add_cover(doc, val, theme)
    elif kind == "image":
        p = Path(val)
        if base_dir and not p.is_absolute():
            p = Path(base_dir) / p
        if p.exists():
            doc.add_picture(str(p), width=Cm(block.get("width_cm", 15)))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if block.get("caption"):
                add_par(doc, block["caption"], S_CAPTION, strip=False).alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            warn(f"  ! 그림 파일 없음: {p}")
    else:
        warn(f"  ! 알 수 없는 블록: {kind} (가능: title subtitle h1 h2 h3 p b note summary table meta approval pagebreak toc cover image)")


def clear_body(doc):
    body = doc.element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)


def theme_from_doc(doc):
    """채우기 때 표/결재란 색을 양식에서 다시 읽는다 (양식의 스타일이 기준)."""
    def color_of(style_name, fallback):
        try:
            c = doc.styles[style_name].font.color.rgb
            return str(c) if c else fallback
        except KeyError:
            return fallback
    primary = color_of(S_H[1], "1F3A5F")
    light = "F2F4F7"
    try:
        shd = doc.styles[S_SUMMARY].element.find(".//" + qn("w:shd"))
        if shd is not None:
            light = shd.get(qn("w:fill"))
    except KeyError:
        pass
    return {"colors": {"primary": primary, "light": light, "muted": color_of(S_CAPTION, "6B7280"),
                       "accent": primary, "text": color_of("Normal", "222222")},
            "fonts": {}, "company": ""}


def apply_style_map(styles):
    """profile.yaml 의 styles: {h1: '제목 1', bullet1: '본문 개조식', ...} 로 회사 스타일 이름을 쓴다."""
    g = globals()
    m = {"title": "S_TITLE", "subtitle": "S_SUB", "note": "S_NOTE", "summary": "S_SUMMARY", "table": "S_TABLE",
         "caption": "S_CAPTION", "table_text": "S_TABLE_TEXT"}
    for k, v in (styles or {}).items():
        if k in m:
            g[m[k]] = v
        elif k in ("h1", "h2", "h3"):
            S_H[int(k[1])] = v
        elif k in ("bullet1", "bullet2", "bullet3"):
            S_B[int(k[-1]) - 1] = v


def fill(template_path, content, out_path, base_dir=None, warn=print, styles=None):
    apply_style_map(styles)
    doc = Document(str(template_path))
    theme = theme_from_doc(doc)
    clear_body(doc)
    header_label = content.get("meta", {}).get("header")
    if header_label:
        hp = doc.sections[0].header.paragraphs[0]
        runs = [r for r in hp.runs if r.text.strip()]
        if runs:
            company = hp.text.split("|")[0].strip()
            runs[0].text = f"{company}  |  {header_label}" if company else header_label
            for r in hp.runs:
                if r is not runs[0] and r._r is not runs[0]._r:
                    r.text = ""
    for block in content.get("blocks", []):
        add_block(doc, block, theme, base_dir, warn)
    if content.get("meta", {}).get("title"):
        doc.core_properties.title = content["meta"]["title"]
    doc.save(str(out_path))
    return out_path
