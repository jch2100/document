#!/usr/bin/env python3
"""환경 점검: 이 컴퓨터(또는 claude.ai 작업 환경)에서 양식대로의 어떤 기능이 되는지 알려 줌.

사용: python scripts/doctor.py
"""
import importlib
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ok = True
print(f"Python {sys.version.split()[0]}  ({sys.executable})")
need = [("pptx", "python-pptx", "PPT 만들기·검사 (필수)"), ("docx", "python-docx", "Word 만들기·검사 (필수)"),
        ("yaml", "PyYAML", "양식 계약서·규칙 읽기 (필수)"), ("lxml", "lxml", "파일 내부 편집 (필수)"),
        ("PIL", "Pillow", "로고·미리보기 시트 (선택)")]
for mod, pkg, use in need:
    try:
        m = importlib.import_module(mod)
        print(f"  ✓ {pkg:<12} {getattr(m, '__version__', '')}  — {use}")
    except ImportError:
        print(f"  ✗ {pkg:<12} 없음 — {use}  →  pip install {pkg}")
        if "필수" in use:
            ok = False
soffice = shutil.which("soffice") or shutil.which("libreoffice")
for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice", r"C:\Program Files\LibreOffice\program\soffice.exe"):
    if not soffice and Path(p).exists():
        soffice = p
print(f"  {'✓' if soffice else '·'} LibreOffice  {soffice or '없음'} — 미리보기 이미지 (없어도 만들기·검사는 됨)")
print(f"  {'✓' if shutil.which('pdftoppm') else '·'} pdftoppm     — 미리보기 PDF→PNG (없으면 PyMuPDF 시도)")
if shutil.which("fc-list"):
    try:
        out = subprocess.run(["fc-list", ":lang=ko", "family"], capture_output=True, text=True, timeout=10).stdout
        fams = sorted({f.split(",")[0].strip() for f in out.splitlines() if f.strip()})
        print(f"  · 한글 글꼴(미리보기용) {len(fams)}개: {', '.join(fams[:5]) or '없음 — 미리보기 한글이 깨질 수 있음 (파일 자체는 정상)'}")
    except Exception:
        pass
skill = HERE.parent
n_tpl = len(list((skill / "templates").glob("*/*/template.*")))
print(f"  {'✓' if n_tpl == 10 else '✗'} 기본 양식 {n_tpl}/10종  ({skill / 'templates'})")
if ok:
    try:
        sys.path.insert(0, str(HERE))
        import ys_pptx
        from ys_common import find_template
        with tempfile.TemporaryDirectory() as t:
            out = Path(t) / "t.pptx"
            ys_pptx.fill(find_template("pptx", "01"), {"slides": [{"layout": "본문", "제목": "환경 점검", "본문": ["항목"]}]}, out)
            print(f"  ✓ 시험 생성 성공 ({out.stat().st_size // 1024} KB)")
    except Exception as e:
        ok = False
        print(f"  ✗ 시험 생성 실패: {type(e).__name__}: {e}")
print("\n결과:", "사용 가능" if ok else "필수 패키지를 설치한 뒤 다시 실행하세요 (pip install -r requirements.txt)")
sys.exit(0 if ok else 1)
