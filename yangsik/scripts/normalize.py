#!/usr/bin/env python3
"""회사 양식 표준화 (원본은 그대로 두고 결과 파일을 새로 씀).

PPTX
  --auto                     플레이스홀더 이름 자동(제목/부제/본문/본문2/이미지/표/차트),
                             한글 글꼴·어절 단위 줄바꿈·ko-KR 설정
  --korean-font "맑은 고딕"  테마 한글 글꼴 지정 (기본: 비어 있을 때만 맑은 고딕)
  --rename-layouts "Title Slide=표지,Title and Content=본문"
  --rename-slots "본문:Text Placeholder 2=핵심메시지"      (레이아웃:기존=새이름, 쉼표로 여러 개)
  --promote 3 --as-layout 요약 [--slot-names 제목,핵심메시지,본문]
                             3번 슬라이드의 텍스트박스들을 그 레이아웃의 '슬롯'으로 승격(반자동)

DOCX
  --korean-font "맑은 고딕"  문서 기본 한글 글꼴
  --add-styles               채우기에 필요한 스타일(개조식 1~3, 참고, 요약 박스, 양식대로 표…) 추가
  --fix-symbols              □ ○ - ※ 로 시작하는 문단 → 개조식/참고 스타일 + 기호 삭제
  --headings-from-numbers    '1. ' '가. ' '1) '로 시작하는 짧은 문단 → 제목 1/2/3 스타일 + 번호 삭제
  --clear-direct             글꼴·크기 직접 서식 제거 (굵게·기울임·색은 유지)
  --remove-empty             연속된 빈 문단 정리
  --all                      위 DOCX 작업 전부

사용: python scripts/normalize.py 입력.pptx --auto --out 표준화.pptx
"""
from __future__ import annotations

import argparse
import copy
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
A = "{%s}" % A_NS
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
log = []


def note(msg):
    log.append(msg)
    print("  ·", msg)


# =====================================================================
# PPTX
# =====================================================================
TYPE_NAME = {"title": "제목", "center_title": "제목", "subtitle": "부제", "picture": "이미지", "table": "표",
             "chart": "차트", "body": "본문", "object": "본문"}


def _typ(pf):
    return str(pf.type).split(".")[-1].split(" ")[0].lower()


def pptx_auto_names(prs):
    for layout in prs.slide_layouts:
        counts = {}
        phs = [s for s in layout.placeholders if _typ(s.placeholder_format) not in ("date", "footer", "slide_number")]
        # 본문류는 면적 큰 순으로 본문, 본문2…
        phs.sort(key=lambda s: (0 if "title" in _typ(s.placeholder_format) else 1, -((s.width or 0) * (s.height or 0))))
        import check as chk
        for s in phs:
            base = TYPE_NAME.get(_typ(s.placeholder_format), "본문")
            if not re.match(r"^(Title|Subtitle|Text|Content|Picture|Table|Chart|제목|부제목|텍스트|내용|그림|표|차트|개체|Placeholder)", s.name or "", re.I) and chk.HANGUL.search(s.name or ""):
                counts[s.name] = 1
                continue  # 이미 한글 이름이면 유지
            counts[base] = counts.get(base, 0) + 1
            new = base if counts[base] == 1 else f"{base}{counts[base]}"
            if s.name != new:
                note(f"[{layout.name}] '{s.name}' → '{new}'")
                s.name = new


def pptx_korean(prs, font=None):
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from lxml import etree
    import check as chk
    master = prs.slide_masters[0]
    tpart = master.part.part_related_by(RT.THEME)
    root = etree.fromstring(tpart.blob)
    changed = False
    for kind in ("majorFont", "minorFont"):
        f = root.find(f".//{A}{kind}")
        ea = f.find(f"{A}ea")
        cur = ea.get("typeface") if ea is not None else ""
        hang = next((x for x in f.findall(f"{A}font") if x.get("script") == "Hang"), None)
        ck = chk.Checker()
        need = font or (None if ck.is_korean_font(cur) or (hang is not None and ck.is_korean_font(hang.get("typeface"))) else "맑은 고딕")
        if need:
            if ea is None:
                ea = etree.SubElement(f, f"{A}ea")
            ea.set("typeface", need)
            if hang is not None:
                hang.set("typeface", need)
            else:
                etree.SubElement(f, f"{A}font", script="Hang", typeface=need)
            note(f"테마 {kind} 한글 글꼴 → {need}")
            changed = True
    if changed:
        tpart._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    # 어절 단위 줄바꿈 + ko-KR
    n = 0
    els = [master._element] + [l._element for l in prs.slide_layouts]
    for el in els:
        for tag in [f"{A}lvl{i}pPr" for i in range(1, 10)] + [f"{A}pPr"]:
            for ppr in el.iter(tag):
                if ppr.get("latinLnBrk") != "0":
                    ppr.set("latinLnBrk", "0")
                    ppr.set("eaLnBrk", "1")
                    n += 1
        for d in el.iter(f"{A}defRPr"):
            d.set("lang", "ko-KR")
    if n:
        note(f"어절 단위 줄바꿈 설정 {n}곳 (한글 단어 중간 끊김 방지)")


def pptx_rename_layouts(prs, spec):
    pairs = dict(x.split("=", 1) for x in spec.split(",") if "=" in x)
    for l in prs.slide_layouts:
        if l.name in pairs:
            note(f"레이아웃 '{l.name}' → '{pairs[l.name]}'")
            l._element.cSld.set("name", pairs[l.name])


def pptx_rename_slots(prs, spec):
    for item in spec.split(","):
        m = re.match(r"(.+?):(.+?)=(.+)", item.strip())
        if not m:
            continue
        lname, old, new = m.groups()
        for l in prs.slide_layouts:
            if l.name == lname:
                for s in l.placeholders:
                    if s.name == old:
                        s.name = new
                        note(f"[{lname}] 슬롯 '{old}' → '{new}'")


def pptx_promote(prs, slide_no, layout_name=None, slot_names=None):
    """슬라이드의 일반 텍스트박스 → 레이아웃 플레이스홀더로 승격."""
    from pptx.oxml import parse_xml
    slide = prs.slides[slide_no - 1]
    layout = slide.slide_layout
    if layout_name:
        others = [s for s in prs.slides if s.slide_layout == layout and s != slide]
        if others:
            note(f"주의: 레이아웃 '{layout.name}'을 다른 슬라이드 {len(others)}장도 씀 — 이름만 바꾸고 슬롯을 추가함(기존 슬라이드는 영향 없음)")
        layout._element.cSld.set("name", layout_name)
    used = {s.placeholder_format.idx for s in layout.placeholders}
    nxt = max([30] + list(used)) + 1
    boxes = [s for s in slide.shapes if not s.is_placeholder and s.has_text_frame and s.text_frame.text.strip()]
    boxes.sort(key=lambda s: (s.top, s.left))
    names = (slot_names.split(",") if slot_names else [])
    for i, box in enumerate(boxes):
        name = names[i] if i < len(names) else f"슬롯{i+1}"
        idx = nxt + i
        sp = copy.deepcopy(box._element)
        # nvSpPr → 플레이스홀더
        nv = sp.find(f"{P}nvSpPr")
        nv.find(f"{P}cNvPr").set("name", name)
        csp = nv.find(f"{P}cNvSpPr")
        for k in list(csp.attrib):
            del csp.attrib[k]
        for c in list(csp):
            csp.remove(c)
        csp.append(parse_xml(f'<a:spLocks xmlns:a="{A_NS}" noGrp="1"/>'))
        nvpr = nv.find(f"{P}nvPr")
        for c in list(nvpr):
            nvpr.remove(c)
        nvpr.append(parse_xml(f'<p:ph xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="body" idx="{idx}"/>'))
        # 첫 run 서식 → lstStyle (레이아웃이 서식을 갖게)
        tx = sp.find(f"{P}txBody")
        first_rpr = tx.find(f".//{A}r/{A}rPr")
        first_ppr = tx.find(f".//{A}p/{A}pPr")
        lst = tx.find(f"{A}lstStyle")
        new_lst = parse_xml(f'<a:lstStyle xmlns:a="{A_NS}"/>')
        lvl = parse_xml(f'<a:lvl1pPr xmlns:a="{A_NS}" eaLnBrk="1" latinLnBrk="0"><a:buNone/></a:lvl1pPr>')
        if first_ppr is not None and first_ppr.get("algn"):
            lvl.set("algn", first_ppr.get("algn"))
        if first_rpr is not None:
            d = copy.deepcopy(first_rpr)
            d.tag = f"{A}defRPr"
            d.set("lang", "ko-KR")
            lvl.append(d)
        new_lst.append(lvl)
        if lst is not None:
            tx.replace(lst, new_lst)
        else:
            tx.insert(1, new_lst)
        # 레이아웃에는 안내문만
        for p in tx.findall(f"{A}p")[1:]:
            tx.remove(p)
        p0 = tx.find(f"{A}p")
        for r in p0.findall(f"{A}r")[1:]:
            p0.remove(r)
        t = p0.find(f".//{A}t")
        if t is not None:
            t.text = f"{name} 입력"
        layout.shapes._spTree.append(sp)
        # 슬라이드 쪽: 텍스트박스를 플레이스홀더 참조로 교체
        sl = box._element
        snv = sl.find(f"{P}nvSpPr")
        scsp = snv.find(f"{P}cNvSpPr")
        for k in list(scsp.attrib):
            del scsp.attrib[k]
        for c in list(scsp):
            scsp.remove(c)
        scsp.append(parse_xml(f'<a:spLocks xmlns:a="{A_NS}" noGrp="1"/>'))
        snvpr = snv.find(f"{P}nvPr")
        for c in list(snvpr):
            snvpr.remove(c)
        snvpr.append(parse_xml(f'<p:ph xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="body" idx="{idx}"/>'))
        snv.find(f"{P}cNvPr").set("name", name)
        note(f"슬라이드 {slide_no}: '{_clip(box.text_frame.text)}' → 레이아웃 '{layout.name}' 슬롯 '{name}'")
    if not boxes:
        note(f"슬라이드 {slide_no}에 승격할 텍스트박스가 없음")


def _clip(t, n=24):
    t = t.replace("\n", " ").strip()
    return t if len(t) <= n else t[: n - 1] + "…"


def run_pptx(a):
    from pptx import Presentation
    prs = Presentation(a.file)
    if a.auto:
        pptx_auto_names(prs)
        pptx_korean(prs, a.korean_font)
    elif a.korean_font:
        pptx_korean(prs, a.korean_font)
    if a.rename_layouts:
        pptx_rename_layouts(prs, a.rename_layouts)
    if a.rename_slots:
        pptx_rename_slots(prs, a.rename_slots)
    if a.promote:
        pptx_promote(prs, a.promote, a.as_layout, a.slot_names)
    prs.save(a.out)


# =====================================================================
# DOCX
# =====================================================================
def docx_korean(doc, font):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    rpr = doc.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    for owner in (rpr, doc.styles["Normal"].element.get_or_add_rPr()):
        rf = owner.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            owner.insert(0, rf)
        if rf.get(qn("w:eastAsiaTheme")) is not None:
            del rf.attrib[qn("w:eastAsiaTheme")]
        rf.set(qn("w:eastAsia"), font)
        rf.set(qn("w:hint"), "eastAsia")
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:eastAsia"), "ko-KR")
    note(f"기본 한글 글꼴 → {font}, 언어 ko-KR")


def docx_add_styles(doc):
    import ys_docx
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn
    theme = ys_docx.theme_from_doc(doc)
    C = theme["colors"]
    font = None
    rf = doc.styles["Normal"].element.find(f".//{W}rFonts")
    if rf is not None:
        font = rf.get(qn("w:eastAsia"))
    theme["fonts"] = {"korean": font or "맑은 고딕", "latin": font or "맑은 고딕"}
    numbering = doc.part.numbering_part.element
    have_ids = {n.get(qn("w:abstractNumId")) for n in numbering.findall(qn("w:abstractNum"))}
    if "901" not in have_ids:
        lv = (ys_docx._lvl_xml(0, "bullet", "□", 420, 320, font=theme["fonts"]["korean"])
              + ys_docx._lvl_xml(1, "bullet", "○", 760, 320, font=theme["fonts"]["korean"])
              + ys_docx._lvl_xml(2, "bullet", "-", 1080, 280, font=theme["fonts"]["korean"]))
        nl = ys_docx._lvl_xml(0, "bullet", "※", 420, 320, font=theme["fonts"]["korean"], color=C["muted"])
        for aid, lvls, kind in ((901, lv, "multilevel"), (902, nl, "singleLevel")):
            el = parse_xml(f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{aid}"><w:multiLevelType w:val="{kind}"/>{lvls}</w:abstractNum>')
            first = numbering.find(qn("w:num"))
            first.addprevious(el) if first is not None else numbering.append(el)
        for nid in (901, 902):
            numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{nid}"><w:abstractNumId w:val="{nid}"/></w:num>'))
    names = {s.name for s in doc.styles}
    from docx.shared import Pt
    for lvl, name in ys_docx.S_B.items():
        if name not in names:
            st = ys_docx._pstyle(doc, name, before=4 if lvl == 0 else 1, after=1)
            ys_docx._set_numpr(st, 901, lvl)
            note(f"스타일 추가: {name}")
    if ys_docx.S_NOTE not in names:
        st = ys_docx._pstyle(doc, ys_docx.S_NOTE, size=9.5, color=C["muted"], before=2, after=2)
        ys_docx._set_numpr(st, 902, 0)
        note(f"스타일 추가: {ys_docx.S_NOTE}")
    if ys_docx.S_SUMMARY not in names:
        st = ys_docx._pstyle(doc, ys_docx.S_SUMMARY, before=6, after=10)
        ys_docx._apply_ppr_xml(st, ys_docx._border_xml(["top", "left", "bottom", "right"], C["primary"], sz=8, space=6)
                               + f'<w:shd w:val="clear" w:color="auto" w:fill="{C["light"]}"/><w:ind w:left="140" w:right="140"/>')
        note(f"스타일 추가: {ys_docx.S_SUMMARY}")
    for name, kw in ((ys_docx.S_TABLE_TEXT, dict(size=9.5, before=0, after=0, line=1.15)),
                     (ys_docx.S_TOC_HEAD, dict(size=14, bold=True, color=C["primary"], before=6, after=10))):
        if name not in names:
            ys_docx._pstyle(doc, name, **kw)
            note(f"스타일 추가: {name}")
    if ys_docx.S_TABLE not in names:
        ts = doc.styles.add_style(ys_docx.S_TABLE, WD_STYLE_TYPE.TABLE)
        ts.element.append(parse_xml(
            f'<w:tblPr {nsdecls("w")}><w:tblBorders><w:top w:val="single" w:sz="8" w:color="{C["primary"]}"/>'
            f'<w:bottom w:val="single" w:sz="8" w:color="{C["primary"]}"/><w:insideH w:val="single" w:sz="4" w:color="D0D5DD"/>'
            f'<w:insideV w:val="single" w:sz="4" w:color="D0D5DD"/></w:tblBorders></w:tblPr>'))
        note(f"스타일 추가: {ys_docx.S_TABLE}")


SYM_LEVEL = [(re.compile(r"^\s*[□■]\s*"), "b0"), (re.compile(r"^\s*[○●◦◎]\s*"), "b1"),
             (re.compile(r"^\s*[-–·•▪]\s+"), "b2"), (re.compile(r"^\s*※\s*"), "note")]
HEAD_LEVEL = [(re.compile(r"^\s*(?:[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.|\d{1,2}\.)\s+(?=\S)"), 1), (re.compile(r"^\s*[가나다라마바사아자차카타파하]\.\s+(?=\S)"), 2),
              (re.compile(r"^\s*(?:\d{1,2}\)|[가나다라마바사아자차카타파하]\))\s+(?=\S)"), 3)]


def _strip_prefix(p, rx):
    first = next((r for r in p.runs if r.text), None)
    if first is None:
        return
    full = p.text
    m = rx.match(full)
    if not m:
        return
    cut = m.end()
    for r in p.runs:
        if cut <= 0:
            break
        n = len(r.text)
        r.text = r.text[cut:] if cut < n else ""
        cut -= n


def docx_fix_symbols(doc):
    import ys_docx
    n = 0
    for p in doc.paragraphs:
        if p._p.pPr is not None and p._p.pPr.numPr is not None:
            continue
        for rx, kind in SYM_LEVEL:
            if rx.match(p.text):
                style = ys_docx.S_NOTE if kind == "note" else ys_docx.S_B[int(kind[1])]
                try:
                    p.style = doc.styles[style]
                except KeyError:
                    continue
                _strip_prefix(p, rx)
                n += 1
                break
    note(f"기호 직접 입력 → 스타일 {n}곳")


def docx_headings(doc, max_len=40):
    n = 0
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t or len(t) > max_len or (p.style is not None and p.style.name.startswith("Heading")):
            continue
        for rx, lvl in HEAD_LEVEL:
            if rx.match(t):
                try:
                    p.style = doc.styles[f"Heading {lvl}"]
                except KeyError:
                    continue
                _strip_prefix(p, rx)
                n += 1
                break
    note(f"번호로 시작하는 짧은 문단 → 제목 스타일 {n}곳 (자동 번호가 대신 붙음)")


def docx_clear_direct(doc):
    from docx.oxml.ns import qn
    n = 0
    for p in doc.paragraphs:
        for r in p.runs:
            rpr = r._r.rPr
            if rpr is None:
                continue
            for tag in ("w:rFonts", "w:sz", "w:szCs"):
                el = rpr.find(qn(tag))
                if el is not None:
                    rpr.remove(el)
                    n += 1
    note(f"직접 서식(글꼴·크기) {n}개 제거")


def docx_remove_empty(doc):
    body = doc.element.body
    prev_empty = False
    n = 0
    for p in list(body.findall(f"{W}p")):
        txt = "".join(t.text or "" for t in p.iter(f"{W}t"))
        special = p.find(f".//{W}br") is not None or p.find(f".//{W}sectPr") is not None or p.find(f".//{W}drawing") is not None
        if not txt.strip() and not special:
            if prev_empty:
                body.remove(p)
                n += 1
            prev_empty = True
        else:
            prev_empty = False
    note(f"연속 빈 문단 {n}개 삭제")


def run_docx(a):
    from docx import Document
    doc = Document(a.file)
    if a.all:
        a.add_styles = a.fix_symbols = a.headings_from_numbers = a.clear_direct = a.remove_empty = True
        a.korean_font = a.korean_font or "맑은 고딕"
    if a.korean_font:
        docx_korean(doc, a.korean_font)
    if a.add_styles or a.fix_symbols:
        docx_add_styles(doc)
    if a.headings_from_numbers:
        docx_headings(doc)
    if a.fix_symbols:
        docx_fix_symbols(doc)
    if a.clear_direct:
        docx_clear_direct(doc)
    if a.remove_empty:
        docx_remove_empty(doc)
    doc.save(a.out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--out", help="결과 파일 (기본: 원본이름_표준화.확장자)")
    ap.add_argument("--auto", action="store_true")
    ap.add_argument("--korean-font")
    ap.add_argument("--rename-layouts")
    ap.add_argument("--rename-slots")
    ap.add_argument("--promote", type=int)
    ap.add_argument("--as-layout")
    ap.add_argument("--slot-names")
    ap.add_argument("--add-styles", action="store_true")
    ap.add_argument("--fix-symbols", action="store_true")
    ap.add_argument("--headings-from-numbers", action="store_true")
    ap.add_argument("--clear-direct", action="store_true")
    ap.add_argument("--remove-empty", action="store_true")
    ap.add_argument("--all", action="store_true")
    a = ap.parse_args()
    src = Path(a.file)
    a.out = a.out or str(src.with_name(src.stem + "_표준화" + src.suffix))
    print(f"[표준화] {src.name} → {a.out}")
    if src.suffix.lower() == ".pptx":
        run_pptx(a)
    elif src.suffix.lower() == ".docx":
        run_docx(a)
    else:
        sys.exit("pptx/docx만 지원")
    if not log:
        print("  · 변경 없음 (옵션을 지정하세요: --help)")


if __name__ == "__main__":
    main()
