#!/usr/bin/env python3
"""claude.ai 업로드용 dist/yangsik.zip 생성 (yangsik/ 폴더를 최상위로)."""
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "yangsik"
OUT = ROOT / "dist" / "yangsik.zip"
SKIP_DIRS = {"__pycache__"}
SKIP_SUFFIX = {".pyc"}

OUT.parent.mkdir(exist_ok=True)
n = 0
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for f in sorted(SKILL.rglob("*")):
        rel = f.relative_to(SKILL)
        if not f.is_file() or SKIP_DIRS & set(rel.parts) or f.suffix in SKIP_SUFFIX:
            continue
        if rel.parts[0] == "profiles" and f.name != "README.md":
            continue  # 회사 양식은 기본 배포에서 제외 (넣으려면 이 줄 삭제)
        z.write(f, Path("yangsik") / rel)
        n += 1
print(f"{OUT}  ({n}개 파일, {OUT.stat().st_size/1024:.0f} KB)")
