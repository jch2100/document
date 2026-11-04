#!/usr/bin/env python3
"""PPTX/DOCX → PDF → PNG 미리보기 (LibreOffice 필요, 없으면 건너뜀).

사용: python scripts/render.py 결과.pptx [--out 폴더] [--sheet]
  --sheet : 전체 페이지를 한 장의 썸네일 시트(contact sheet)로 합침

주의: LibreOffice는 '맑은 고딕'이 없으면 다른 글꼴로 대체합니다.
      미리보기는 '배치 확인용'이며 실제 PowerPoint/Word와 줄바꿈이 조금 다를 수 있습니다.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def soffice_bin():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe"):
        if Path(p).exists():
            return p
    return None


def to_pdf(src: Path, outdir: Path) -> Path | None:
    sb = soffice_bin()
    if not sb:
        print("· LibreOffice가 없어 미리보기를 건너뜁니다 (구조 검사는 check.py로 가능)")
        return None
    with tempfile.TemporaryDirectory() as prof:
        cmd = [sb, f"-env:UserInstallation=file://{prof}", "--headless", "--convert-to", "pdf",
               "--outdir", str(outdir), str(src)]
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    pdf = outdir / (src.stem + ".pdf")
    return pdf if pdf.exists() else None


def to_pngs(pdf: Path, outdir: Path, dpi=60) -> list[Path]:
    if not shutil.which("pdftoppm"):
        try:
            import fitz  # PyMuPDF
        except ImportError:
            print("· pdftoppm/PyMuPDF가 없어 PDF까지만 만듭니다:", pdf)
            return []
        doc = fitz.open(pdf)
        out = []
        for i, page in enumerate(doc, 1):
            p = outdir / f"{pdf.stem}-{i:02d}.png"
            page.get_pixmap(dpi=dpi).save(p)
            out.append(p)
        return out
    prefix = outdir / pdf.stem
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(prefix)], check=True)
    return sorted(p for p in outdir.glob(f"{pdf.stem}-*.png") if p.stem.rsplit("-", 1)[-1].isdigit())


def contact_sheet(pngs: list[Path], out: Path, cols=3, pad=12):
    from PIL import Image, ImageDraw
    ims = [Image.open(p).convert("RGB") for p in pngs]
    if not ims:
        return None
    w, h = ims[0].size
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * (w + pad) + pad, rows * (h + pad + 18) + pad), "#E5E7EB")
    d = ImageDraw.Draw(sheet)
    for i, im in enumerate(ims):
        x = pad + (i % cols) * (w + pad)
        y = pad + (i // cols) * (h + pad + 18)
        sheet.paste(im.resize((w, h)), (x, y + 18))
        d.text((x, y + 2), str(i + 1), fill="#374151")
    sheet.save(out)
    return out


def render(src, out_dir=None, sheet=False, dpi=60, keep_pages=True):
    src = Path(src)
    out_dir = Path(out_dir or src.parent / (src.stem + "_" + src.suffix.lstrip(".") + "_preview"))
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf = to_pdf(src, out_dir)
    if not pdf:
        return None
    pngs = to_pngs(pdf, out_dir, dpi)
    result = {"pdf": pdf, "pages": pngs}
    if sheet and pngs:
        result["sheet"] = contact_sheet(pngs, out_dir / f"{src.stem}-sheet.png")
        if not keep_pages:
            for p in pngs:
                p.unlink()
            result["pages"] = []
    return result


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--out")
    ap.add_argument("--sheet", action="store_true")
    ap.add_argument("--dpi", type=int, default=60)
    a = ap.parse_args()
    r = render(a.file, a.out, a.sheet, a.dpi)
    if not r:
        sys.exit(0)
    print("PDF:", r["pdf"])
    if r.get("sheet"):
        print("썸네일 시트:", r["sheet"])
    print(f"페이지 이미지 {len(r['pages'])}장:", r["pages"][0].parent if r["pages"] else "-")


if __name__ == "__main__":
    main()
