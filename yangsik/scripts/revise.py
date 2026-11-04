#!/usr/bin/env python3
"""결재 피드백 부분 수정 — 고칠 곳만 바꾸고 나머지(사람이 손댄 서식 포함)는 그대로 둡니다.

사용: python scripts/revise.py 보고서_v1.pptx --ops 수정.json --out 보고서_v2.pptx

수정.json 예 (PPTX)
  {"ops": [
    {"slide": 3, "slot": "핵심메시지", "value": "이용률 하락의 원인은 접근성"},
    {"slide": 3, "slot": "본문", "value": ["항목1", ["하위"]]},
    {"slide": 5, "delete": true},
    {"insert_after": 2, "spec": {"layout": "본문", "제목": "추가 근거", "본문": ["..."]}},
    {"move": 6, "to": 2}
  ]}
수정.json 예 (DOCX) — 문단을 '포함된 글자'로 찾음 (스타일 유지)
  {"ops": [
    {"find": "월 활성 이용률", "replace": "월 활성 이용률 31% (전년 대비 7%p 하락)"},
    {"find": "붙임", "delete": true},
    {"after": "기대 효과", "blocks": [{"b": ["추가 효과"]}]}
  ]}
슬라이드 번호는 '수정 전' 기준이 아니라 각 작업 시점 기준입니다 (위에서부터 순서대로 적용).
변경 내역은 화면과 *_변경내역.md 로 남깁니다.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ys_common import load_json  # noqa: E402

changes = []


def _clip(t, n=36):
    t = str(t).replace("\n", " ")
    return t if len(t) <= n else t[: n - 1] + "…"


def revise_pptx(path, ops, out, base_dir):
    from pptx import Presentation
    import ys_pptx
    prs = Presentation(str(path))
    sldIdLst = prs.slides._sldIdLst
    for op in ops:
        if "slide" in op and "slot" in op:
            slide = prs.slides[op["slide"] - 1]
            target = next((s for s in slide.placeholders if s.name == op["slot"]), None)
            if target is None:
                names = [s.name for s in slide.placeholders]
                print(f"  ! 슬라이드 {op['slide']}에 '{op['slot']}' 없음 (있는 것: {names})")
                continue
            before = target.text_frame.text if target.has_text_frame else "(표/차트)"
            if target.has_text_frame:
                ys_pptx.set_text(target, op["value"])
                changes.append(f"슬라이드 {op['slide']} · {op['slot']}: \"{_clip(before)}\" → \"{_clip(target.text_frame.text)}\"")
        elif "slide" in op and op.get("delete"):
            sld = list(sldIdLst)[op["slide"] - 1]
            prs.part.drop_rel(sld.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"))
            sldIdLst.remove(sld)
            changes.append(f"슬라이드 {op['slide']} 삭제")
        elif "insert_after" in op:
            ys_pptx.add_slide(prs, op["spec"], base_dir)
            new = list(sldIdLst)[-1]
            sldIdLst.remove(new)
            sldIdLst.insert(op["insert_after"], new)
            changes.append(f"슬라이드 {op['insert_after']} 뒤에 '{op['spec'].get('제목', op['spec'].get('layout'))}' 추가")
        elif "move" in op:
            el = list(sldIdLst)[op["move"] - 1]
            sldIdLst.remove(el)
            sldIdLst.insert(op["to"] - 1, el)
            changes.append(f"슬라이드 {op['move']} → {op['to']}번 위치로 이동")
        else:
            print("  ! 알 수 없는 작업:", op)
    prs.save(str(out))


def revise_docx(path, ops, out, base_dir):
    from docx import Document
    import ys_docx
    doc = Document(str(path))
    theme = ys_docx.theme_from_doc(doc)
    for op in ops:
        key = op.get("find") or op.get("after")
        hits = [p for p in doc.paragraphs if key and key in p.text]
        if not hits:
            print(f"  ! '{key}' 를 포함한 문단 없음")
            continue
        p = hits[0]
        if len(hits) > 1:
            print(f"  · '{key}' 가 {len(hits)}곳 — 첫 번째만 수정 (더 길게 지정하면 정확해짐)")
        if "replace" in op:
            before = p.text
            runs = p.runs
            for r in runs[1:]:
                r._r.getparent().remove(r._r)
            if runs:
                runs[0].text = ys_docx.clean(op["replace"]) if p.style.name in ys_docx.S_B.values() else op["replace"]
            else:
                p.add_run(op["replace"])
            changes.append(f"문단 \"{_clip(before)}\" → \"{_clip(p.text)}\"")
        elif op.get("delete"):
            changes.append(f"문단 삭제: \"{_clip(p.text)}\"")
            p._p.getparent().remove(p._p)
        elif "blocks" in op:
            # 끝에 추가 후 해당 위치로 옮김
            body = doc.element.body
            before_len = len(body)
            for b in op["blocks"]:
                ys_docx.add_block(doc, b, theme, base_dir)
            new_els = list(body)[before_len - 1:-1] if body[-1].tag.endswith("sectPr") else list(body)[before_len:]
            anchor = p._p
            for el in new_els:
                if el.tag.endswith("sectPr"):
                    continue
                anchor.addnext(el)
                anchor = el
            changes.append(f"\"{_clip(p.text)}\" 뒤에 블록 {len(op['blocks'])}개 추가")
    doc.save(str(out))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--ops", required=True)
    ap.add_argument("--out")
    ap.add_argument("--no-check", action="store_true")
    a = ap.parse_args()
    src = Path(a.file)
    out = Path(a.out) if a.out else src.with_name(src.stem + "_수정" + src.suffix)
    data = load_json(a.ops)
    ops = data["ops"] if isinstance(data, dict) else data
    print(f"[수정] {src.name} → {out.name} (작업 {len(ops)}개)")
    if src.suffix.lower() == ".pptx":
        revise_pptx(src, ops, out, Path(a.ops).resolve().parent)
    else:
        revise_docx(src, ops, out, Path(a.ops).resolve().parent)
    for c in changes:
        print("  ✓", c)
    log = out.with_name(out.stem + "_변경내역.md")
    log.write_text(f"# 변경 내역: {src.name} → {out.name}\n\n" + "\n".join(f"- {c}" for c in changes) + "\n", encoding="utf-8")
    if not a.no_check:
        import check
        prof = src.parent / "profile.yaml"
        check.run(out, prof if prof.exists() else None)


if __name__ == "__main__":
    main()
