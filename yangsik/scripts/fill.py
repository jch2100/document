#!/usr/bin/env python3
"""양식에 내용 채우기 → 검증까지 한 번에.

사용:
  python scripts/fill.py --template pptx/01 --content 내용.json --out 결과.pptx
  python scripts/fill.py --template docx/01-1page-report --content 내용.json --out 결과.docx
  python scripts/fill.py --template profiles/우리회사/template.pptx --content 내용.json --out 결과.pptx

--template : 'pptx/번호' 'docx/이름' 같은 기본 양식, 또는 파일 경로(등록한 회사 양식)
--content  : JSON (PPTX는 slides 목록, DOCX는 blocks 목록 — references/content-format.md)
옆에 profile.yaml이 있으면 검증 기준으로 자동 사용합니다.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ys_common import SKILL_DIR, find_template, load_json, load_yaml  # noqa: E402


def resolve_template(spec: str) -> Path:
    p = Path(spec)
    if p.is_dir():
        found = next(iter(sorted(p.glob("template.*"))), None)
        if found:
            return found
    if p.exists():
        return p
    # '회사/<이름>' → ./profiles/<이름>/ 또는 SKILL_DIR/profiles/<이름>/
    if spec.startswith(("회사/", "profile/", "profiles/")):
        name = spec.split("/", 1)[1]
        for base in (Path("profiles"), SKILL_DIR / "profiles"):
            d = base / name
            found = next(iter(sorted(d.glob("template.*"))), None) if d.is_dir() else None
            if found:
                return found
    if "/" in spec:
        kind, name = spec.split("/", 1)
        return find_template(kind, name)
    sys.exit(f"양식 경로를 찾을 수 없음: {spec}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--template", required=True)
    ap.add_argument("--content", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--profile", help="검증 기준 profile.yaml (기본: 양식 옆 profile.yaml)")
    ap.add_argument("--no-check", action="store_true")
    ap.add_argument("--preview", action="store_true", help="LibreOffice로 썸네일 시트 생성")
    a = ap.parse_args()

    tpl = resolve_template(a.template)
    content = load_json(a.content)
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    base_dir = Path(a.content).resolve().parent
    kind = tpl.suffix.lower()
    if kind != out.suffix.lower():
        sys.exit(f"양식({kind})과 출력({out.suffix}) 형식이 다릅니다")

    profile_path = Path(a.profile) if a.profile else tpl.parent / "profile.yaml"
    profile = load_yaml(profile_path) if profile_path.exists() else {}

    print(f"[채우기] 양식: {tpl}")
    if kind == ".pptx":
        import ys_pptx
        if "slides" not in content:
            sys.exit("PPTX 내용 JSON에는 'slides' 목록이 필요합니다")
        ys_pptx.fill(tpl, content, out, base_dir)
        print(f"  슬라이드 {len(content['slides'])}장 → {out}")
    elif kind == ".docx":
        import ys_docx
        if "blocks" not in content:
            sys.exit("DOCX 내용 JSON에는 'blocks' 목록이 필요합니다")
        ys_docx.fill(tpl, content, out, base_dir, styles=profile.get("styles"))
        print(f"  블록 {len(content['blocks'])}개 → {out}")
    else:
        sys.exit("pptx/docx 양식만 지원합니다")

    verdict = "SKIP"
    if not a.no_check:
        import check
        verdict, _ = check.run(out, profile_path if profile_path.exists() else None)
    if a.preview:
        import render
        r = render.render(out, sheet=True, keep_pages=False)
        if r and r.get("sheet"):
            print(f"[미리보기] {r['sheet']}")
    sys.exit(1 if verdict == "FAIL" else 0)


if __name__ == "__main__":
    main()
