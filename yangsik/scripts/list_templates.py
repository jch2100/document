#!/usr/bin/env python3
"""기본 제공 양식 목록 + 등록한 회사 양식(profiles/) 목록.

사용: python scripts/list_templates.py [--profiles 폴더]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ys_common import TEMPLATES_DIR, load_yaml  # noqa: E402


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
    prof_dir = Path(a.profiles)
    if prof_dir.exists():
        rows = [d for d in sorted(prof_dir.iterdir()) if (d / "profile.yaml").exists()]
        if rows:
            print("\n■ 등록한 회사 양식\n")
            for d in rows:
                p = load_yaml(d / "profile.yaml")
                tpl = next(d.glob("template.*"), None)
                print(f"  {str(tpl or d):<40}{p.get('title', ''):<20}등급 {p.get('grade', '?')}")
    print("\n사용 예: python scripts/fill.py --template pptx/01 --content 내용.json --out 결과.pptx")


if __name__ == "__main__":
    main()
