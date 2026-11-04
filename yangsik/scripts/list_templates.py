#!/usr/bin/env python3
"""기본 제공 양식 목록 + 등록한 회사 양식(profiles/) 목록.

사용: python scripts/list_templates.py [--profiles 폴더]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ys_common import SKILL_DIR, TEMPLATES_DIR, load_yaml  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--profiles", default="profiles")
    a = ap.parse_args()
    print("■ 기본 제공 양식 (내 양식이 없을 때)\n")
    print(f"  {'지정 이름':<26}{'문서':<16}{'언제 쓰나':<40}분량")
    for kind in ("pptx", "docx"):
        for d in sorted((TEMPLATES_DIR / kind).iterdir()):
            pf = d / "profile.yaml"
            if not pf.exists():
                continue
            p = load_yaml(pf)
            print(f"  {kind + '/' + d.name:<26}{p['title']:<16}{p['purpose'][:38]:<40}{p['length']}")
    shown = False
    for base in dict.fromkeys([Path(a.profiles), SKILL_DIR / "profiles"]):
        if not base.exists():
            continue
        for pf in sorted(base.glob("**/profile.yaml")):
            if not shown:
                print("\n■ 등록한 회사 양식\n")
                shown = True
            p = load_yaml(pf)
            rel = pf.parent.relative_to(base)
            short = "/".join(rel.parts[:2] + (rel.parts[2].split("-")[0],)) if len(rel.parts) == 3 else str(rel)
            print(f"  {short:<26}{p.get('title', '')[:30]:<32}{pf.parent}")
    print("\n사용 예: python scripts/fill.py --template pptx/01 --content 내용.json --out 결과.pptx --company \"회사명\"")
    print("        등록 양식: --template 그린텍/pptx/01")


if __name__ == "__main__":
    main()
