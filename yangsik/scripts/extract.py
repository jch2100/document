#!/usr/bin/env python3
"""PPTX/DOCX → 내용 JSON (fill.py 의 역방향).

쓰임
  - 사람이 고친 파일에서 내용을 다시 뽑아 이어서 작업
  - 회사의 '잘 쓴 예시 문서'에서 문체·분량을 학습할 때 (profiles/회사/samples/)
  - 다른 양식으로 옮겨 담기: extract → fill(새 양식)

사용: python scripts/extract.py 보고서.pptx [--out 내용.json]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))


def _runs_md(runs):
    """직접 굵게 지정된 글자 조각 → **굵게** (양식이 원래 굵은 칸은 그대로)"""
    out = ""
    for r in runs:
        t = r.text
        if t.strip() and r.font.bold:
            out += f"**{t}**"
        else:
            out += t
    return out.replace("****", "")


def _nest(paragraphs):
    """[(level, text)] → ['a', ['a-1'], 'b']"""
    root = []
    stack = [(0, root)]
    for level, text in paragraphs:
        while len(stack) > 1 and stack[-1][0] > level:
            stack.pop()
        while stack[-1][0] < level:
            new = []
            stack[-1][1].append(new)
            stack.append((stack[-1][0] + 1, new))
        stack[-1][1].append(text)
    return root


def extract_pptx(path):
    from pptx import Presentation
    prs = Presentation(str(path))
    slides = []
    for s in prs.slides:
        spec = {"layout": s.slide_layout.name}
        for shp in s.shapes:
            if not shp.is_placeholder:
                if shp.has_text_frame and shp.text_frame.text.strip():
                    spec.setdefault("_양식밖_텍스트", []).append(shp.text_frame.text)
                continue
            t = str(shp.placeholder_format.type)
            if "SLIDE_NUMBER" in t or "FOOTER" in t or "DATE" in t:
                continue
            if getattr(shp, "has_table", False) and shp.has_table:
                tbl = shp.table
                rows = [[c.text for c in r.cells] for r in tbl.rows]
                t = {"header": rows[0], "rows": rows[1:]} if rows else {}
                tot = sum(cl.width for cl in tbl.columns) or 1
                t["col_widths"] = [round(cl.width / tot * 10, 1) for cl in tbl.columns]
                # 강조 열: 본문 행 셀이 모두 채우기 색을 가진 열
                for ci in range(len(tbl.columns)):
                    cells = [tbl.cell(ri, ci) for ri in range(1, len(tbl.rows))]
                    if cells and all(c._tc.tcPr is not None and c._tc.tcPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill") is not None for c in cells):
                        t["highlight_col"] = ci
                        break
                spec[shp.name] = t
            elif getattr(shp, "has_chart", False) and shp.has_chart:
                ch = shp.chart
                plot = ch.plots[0]
                spec[shp.name] = {"type": "bar", "categories": [str(c) for c in plot.categories],
                                  "series": {se.name: list(se.values) for se in plot.series}}
            elif shp.has_text_frame:
                paras = [(p.level, _runs_md(p.runs).replace("\u200b", "")) for p in shp.text_frame.paragraphs]
                paras = [x for x in paras if x[1].strip()]
                if not paras:
                    continue
                spec[shp.name] = paras[0][1] if len(paras) == 1 and paras[0][0] == 0 and "본문" not in shp.name else _nest(paras)
        if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip():
            spec["노트"] = s.notes_slide.notes_text_frame.text
        slides.append(spec)
    return {"meta": {"source": Path(path).name}, "slides": slides}


def extract_docx(path):
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(str(path))
    rev = {"Title": "title", "Subtitle": "subtitle", "Heading 1": "h1", "Heading 2": "h2", "Heading 3": "h3",
           "참고": "note", "요약 박스": "summary", "Caption": "p"}
    bullets = {"개조식 1": 0, "개조식 2": 1, "개조식 3": 2, "List Bullet": 0, "List Bullet 2": 1, "List Bullet 3": 2}
    blocks, pending = [], []

    def flush():
        if pending:
            blocks.append({"b": _nest(pending)})
            pending.clear()
    tables = {t._tbl: t for t in doc.tables}
    for el in doc.element.body:
        if el.tag == qn("w:tbl"):
            flush()
            t = tables.get(el)
            if t is None:
                continue
            rows = []
            for r in t.rows:
                row, seen = [], set()
                for c in r.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    row.append(c.text)
                rows.append(row)
            look = el.find(qn("w:tblPr") + "/" + qn("w:tblLook"))
            if rows and rows[0] and rows[0][0].replace("\n", "") == "결재":
                blocks.append({"approval": rows[0][1:]})
            elif look is not None and look.get(qn("w:firstRow")) == "0" and all(len(r) % 2 == 0 for r in rows):
                pairs = [[r[i], r[i + 1]] for r in rows for i in range(0, len(r), 2) if r[i]]
                blocks.append({"meta": pairs})
            else:
                blocks.append({"table": {"header": rows[0], "rows": rows[1:]}} if rows else {"table": {}})
            continue
        if el.tag != qn("w:p"):
            continue
        from docx.text.paragraph import Paragraph
        p = Paragraph(el, doc)
        text = p.text
        if not text.strip():
            continue
        st = p.style.name if p.style is not None else "Normal"
        if st in bullets:
            pending.append((bullets[st], text))
            continue
        flush()
        if st in ("Caption", "캡션") and blocks and "table" in blocks[-1]:
            blocks[-1]["caption"] = text
            continue
        kind = rev.get(st, "p")
        if kind == "summary" and blocks and "summary" in blocks[-1]:
            blocks[-1]["summary"].append(text)
        elif kind == "summary":
            blocks.append({"summary": [text]})
        else:
            blocks.append({kind: text})
    flush()
    return {"meta": {"source": Path(path).name}, "blocks": blocks}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--out")
    a = ap.parse_args()
    src = Path(a.file)
    data = extract_pptx(src) if src.suffix.lower() == ".pptx" else extract_docx(src)
    out = Path(a.out) if a.out else src.with_name(src.stem + "_내용.json")
    out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    n = len(data.get("slides", data.get("blocks", [])))
    print(f"[추출] {src.name} → {out} ({n}개)")


if __name__ == "__main__":
    main()
