#!/usr/bin/env python3
"""회사 양식 → 브랜드(색·글꼴·로고) 추출 → 표준 양식 10종을 그 브랜드로 재구성.

'칸 없이 그려진 양식(C등급)'을 고쳐 쓰는 대신, 브랜드만 살리고 구조는 표준 레이아웃으로 바꾼다.
원본 양식은 건드리지 않는다.

사용:
  python scripts/standardize.py 회사양식.pptx --name 우리회사 --company "우리회사"
  python scripts/standardize.py 회사보고서.docx 회사양식.pptx --name 우리회사     # 여러 파일에서 함께 추출
옵션:
  --logo 로고.png      로고를 직접 지정 (자동 추출이 틀렸을 때)
  --primary 0B2545     대표색 직접 지정 / --accent 1F6FEB 강조색 직접 지정
  --font "나눔고딕"     글꼴 직접 지정
  --out profiles       결과 상위 폴더

결과: profiles/<이름>/
  theme.yaml          추출한 브랜드 (사람이 고쳐서 다시 돌릴 수 있음)
  brand.md            무엇을 어디서 뽑았는지 + 판단 근거
  brand.png           색 견본
  <이름>.zip           위 폴더 전체 (사용자에게 전달용)
  logo.png            추출한 로고 (있을 때)
  pptx/01-…/ docx/01-…/  브랜드 적용 표준 양식 10종 (template + profile.yaml)
"""
from __future__ import annotations

import argparse
import colorsys
import io
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ys_common import dump_yaml, load_theme, load_yaml, SKILL_DIR  # noqa: E402

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
notes: list[str] = []


# ---------------------------------------------------------------------------
# 색 계산
# ---------------------------------------------------------------------------
def hex2rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def rgb2hex(r, g, b):
    return "".join(f"{max(0, min(255, round(v * 255))):02X}" for v in (r, g, b))


def lum(h):
    def ch(c):
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = hex2rgb(h)
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast(a, b):
    la, lb = sorted((lum(a), lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def hls(h):
    return colorsys.rgb_to_hls(*hex2rgb(h))


def sat(h):
    return hls(h)[2]


def with_light(h, l):
    hh, _l, s = hls(h)
    return rgb2hex(*colorsys.hls_to_rgb(hh, l, s))


def is_neutral(h):
    _h, l, s = hls(h)
    return s < 0.12 or l > 0.93 or l < 0.06


def darken_until(h, target=7.0):
    """흰 글자 대비가 target 이상이 될 때까지 어둡게."""
    _h, l, _s = hls(h)
    while contrast(h, "FFFFFF") < target and l > 0.05:
        l -= 0.03
        h = with_light(h, l)
    return h


# ---------------------------------------------------------------------------
# 추출
# ---------------------------------------------------------------------------
def theme_from_xml(xml_bytes):
    from lxml import etree
    root = etree.fromstring(xml_bytes)
    colors = {}
    cs = root.find(f".//{A}clrScheme")
    if cs is not None:
        for el in cs:
            tag = etree.QName(el).localname
            c = el[0] if len(el) else None
            if c is not None:
                colors[tag] = (c.get("val") or c.get("lastClr") or "").upper()
    fonts = {}
    for kind in ("majorFont", "minorFont"):
        f = root.find(f".//{A}{kind}")
        if f is None:
            continue
        lat = f.find(f"{A}latin")
        ea = f.find(f"{A}ea")
        hang = next((x.get("typeface") for x in f.findall(f"{A}font") if x.get("script") == "Hang"), "")
        fonts[kind] = {"latin": lat.get("typeface") if lat is not None else "", "ea": (ea.get("typeface") if ea is not None else "") or hang}
    return colors, fonts


def extract_pptx(path, out_dir):
    from pptx import Presentation
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    prs = Presentation(str(path))
    master = prs.slide_masters[0]
    colors, fonts = theme_from_xml(master.part.part_related_by(RT.THEME).blob)
    # 실제로 많이 쓰인 색 (도형 채우기·글자) — 테마가 기본값인 양식 대비
    used = Counter()
    for el_root in [master._element] + [l._element for l in prs.slide_layouts] + [s._element for s in list(prs.slides)[:15]]:
        for c in el_root.iter(f"{A}srgbClr"):
            v = (c.get("val") or "").upper()
            if v and not is_neutral(v):
                used[v] += 1
    # 로고: 마스터·레이아웃·앞 슬라이드의 그림 중 '여러 번 나오거나 모서리에 작게 있는 것'
    pics = Counter()
    blobs = {}
    pic_pos = {}
    sw, sh = prs.slide_width, prs.slide_height

    def scan(shapes, weight):
        for s in shapes:
            if s.shape_type is not None and "PICTURE" in str(s.shape_type):
                try:
                    blob = s.image.blob
                except Exception:
                    continue
                area = (s.width * s.height) / (sw * sh)
                corner = (s.top < sh * 0.2 or s.top + s.height > sh * 0.8)
                if area < 0.08 and corner:
                    key = hash(blob)
                    pics[key] += weight
                    blobs[key] = (blob, s.image.ext)
                    if s.left > sw * 0.55 and s.top < sh * 0.2:
                        pic_pos[key] = "top-right"
    scan(master.shapes, 5)
    for l in prs.slide_layouts:
        scan(l.shapes, 2)
    for s in list(prs.slides)[:10]:
        scan(s.shapes, 1)
    # 상단 색띠: 화면 폭 90% 이상, 위쪽 25% 안, 무채색이 아닌 채우기
    band = False
    for owner in [master] + list(prs.slide_layouts) + list(prs.slides)[:5]:
        for s in owner.shapes:
            try:
                if s.width and s.width > sw * 0.9 and s.top is not None and s.top < sh * 0.25 and s.height < sh * 0.2:
                    if s.fill.type is not None and str(s.fill.fore_color.rgb) and not is_neutral(str(s.fill.fore_color.rgb)):
                        band = True
            except Exception:
                pass
    if band:
        notes.append("원본에 상단 색띠가 있음 → 표준 양식 맨 위에 대표색 띠를 넣음")
    logo_pos = None
    logo = None
    if pics:
        key, _n = pics.most_common(1)[0]
        blob, ext = blobs[key]
        logo_pos = pic_pos.get(key)
        logo = out_dir / f"logo.{ext}"
        logo.write_bytes(blob)
        notes.append(f"PPTX 로고: 마스터/레이아웃/슬라이드 모서리의 작은 그림 중 가장 자주 나온 것 ({ext})")
    size = f"{sw/914400:.2f}x{sh/914400:.2f}in"
    if abs(sw / sh - 16 / 9) > 0.05:
        notes.append(f"원본 슬라이드 크기 {size} (16:9 아님) → 표준 양식은 16:9 로 만듦")
    company = prs.core_properties.author or ""
    if logo_pos:
        notes.append("원본 로고가 오른쪽 위 → 표준 양식도 내용 장 오른쪽 위에 로고")
    return dict(colors=colors, fonts=fonts, used=used, logo=logo, company=company, source=Path(path).name,
                band=band, logo_pos=logo_pos, aspect=f"{sw/914400:.2f}x{sh/914400:.2f}")


def extract_docx(path, out_dir):
    from docx import Document
    doc = Document(str(path))
    colors, fonts = {}, {}
    with zipfile.ZipFile(path) as z:
        th = [n for n in z.namelist() if n.startswith("word/theme/")]
        if th:
            colors, fonts = theme_from_xml(z.read(th[0]))
    used = Counter()
    for name in ("Title", "Heading 1", "Heading 2", "제목 1", "제목 2"):
        try:
            c = doc.styles[name].font.color.rgb
            if c is not None and not is_neutral(str(c)):
                used[str(c)] += 5
        except (KeyError, AttributeError):
            pass
    for shd in doc.element.body.iter(f"{W}shd"):
        v = (shd.get(f"{W}fill") or "").upper()
        if re.fullmatch(r"[0-9A-F]{6}", v) and not is_neutral(v):
            used[v] += 2
    for c in doc.element.body.iter(f"{W}color"):
        v = (c.get(f"{W}val") or "").upper()
        if re.fullmatch(r"[0-9A-F]{6}", v) and not is_neutral(v):
            used[v] += 1
    ea = None
    rf = doc.styles["Normal"].element.find(f".//{W}rFonts")
    if rf is not None:
        ea = rf.get(f"{W}eastAsia")
    if ea:
        fonts.setdefault("minorFont", {})["ea"] = ea
    logo = None
    for sec in doc.sections:
        for part in (sec.header, sec.first_page_header):
            try:
                rels = part.part.rels
            except Exception:
                continue
            for rel in rels.values():
                if "image" in rel.reltype:
                    ext = Path(rel.target_part.partname).suffix.lstrip(".")
                    logo = out_dir / f"logo.{ext}"
                    logo.write_bytes(rel.target_part.blob)
                    notes.append(f"DOCX 로고: 머리말 그림 ({ext})")
                    break
            if logo:
                break
        if logo:
            break
    company = doc.core_properties.author or ""
    return dict(colors=colors, fonts=fonts, used=used, logo=logo, company=company, source=Path(path).name,
                band=False, logo_pos=None, aspect=None)


OFFICE_DEFAULTS = {"44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "1F497D", "4F81BD", "C0504D",
                   "9BBB59", "8064A2", "4BACC6", "F79646", "0E2841", "156082", "E97132", "196B24", "0F9ED5", "A02B93",
                   # Word 기본 스타일 제목색 (회사색이 아님)
                   "17365D", "365F91", "1F3864", "2F5496", "1F4E79", "2E74B5", "0F4761", "243F60", "1F3763", "4472C4"}


def decide(found, overrides):
    """여러 파일의 추출 결과 → 테마 결정 + 근거."""
    used = Counter()
    theme_cands = []
    for f in found:
        used.update({k: v for k, v in f["used"].items() if k not in OFFICE_DEFAULTS})
        c = f["colors"]
        for k in ("dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
            v = c.get(k)
            if v and not is_neutral(v):
                theme_cands.append((k, v))
    default_theme = all(v in OFFICE_DEFAULTS for _k, v in theme_cands) if theme_cands else True
    if default_theme:
        notes.append("테마 색이 Office 기본값 → 테마 대신 실제로 많이 쓴 색을 기준으로 판단")
    # 대표색: 실제 사용 빈도 높은 진한 색 > 테마 dk2/accent1
    ranked = [c for c, _n in used.most_common() if not is_neutral(c)]
    if not default_theme:
        ranked = [v for k, v in theme_cands if k in ("dk2", "accent1")] + ranked
    primary = overrides.get("primary")
    if not primary:
        dark = [c for c in ranked if hls(c)[1] < 0.45]
        primary = dark[0] if dark else (ranked[0] if ranked else "0B2545")
        why = "실사용·테마 색 중 가장 우선인 진한 색" if dark else ("가장 많이 쓴 색을 어둡게 조정" if ranked else "추출 실패 → 기본 남색")
        notes.append(f"대표색 {primary}: {why}")
    fixed = darken_until(primary, 7.0)
    if fixed != primary:
        notes.append(f"대표색 {primary} → {fixed}: 흰 글자 대비 7:1 이상이 되도록 어둡게 (머리띠·표지에 흰 글자를 쓰기 때문)")
        primary = fixed
    accent = overrides.get("accent")
    if not accent:
        ph = hls(primary)[0]
        cands = [c for c in ranked if c != primary and 0.25 <= hls(c)[1] <= 0.65 and sat(c) > 0.35]
        # 대표색과 색상이 비슷하면 '같은 계열의 밝은 색', 다르면 그 색
        accent = cands[0] if cands else with_light(primary, 0.48)
    # 강조색은 원본 그대로(채우기·막대·머리띠). 작은 글씨에는 대비를 맞춘 어두운 버전(accent_text)을 따로 쓴다
    from ys_common import readable_on_white
    accent_text = readable_on_white(accent, 4.5)
    if accent_text != accent:
        notes.append(f"강조색 {accent} 은 그대로 쓰고, 작은 글씨(트래커 등)에만 대비 4.5:1 버전 {accent_text} 사용")
    # 대표색과 너무 비슷하면 강조가 안 보임 → 다른 후보 또는 밝게
    if contrast(accent, primary) < 1.8:
        alt = [c for c in ranked if c not in (primary, accent) and contrast(c, primary) >= 1.8 and sat(c) > 0.3]
        if alt and not overrides.get("accent"):
            notes.append(f"강조색 {accent} 이 대표색과 비슷해 {alt[0]} 로 바꿈 (강조가 보이도록)")
            accent = alt[0]
            accent_text = readable_on_white(accent, 4.5)
        else:
            lighter = with_light(accent, min(hls(primary)[1] + 0.35, 0.6))
            notes.append(f"강조색 {accent} 이 대표색과 비슷해 밝게 조정 → {lighter}")
            accent = lighter
            accent_text = readable_on_white(accent, 4.5)
    light = with_light(primary, 0.965)
    hh, _l, s = hls(light)
    light = rgb2hex(*colorsys.hls_to_rgb(hh, 0.965, min(s, 0.3)))
    font = overrides.get("font")
    if not font:
        for f in found:
            for kind in ("minorFont", "majorFont"):
                ea = f["fonts"].get(kind, {}).get("ea")
                if ea:
                    font = ea
                    break
            if font:
                break
        if font:
            notes.append(f"한글 글꼴 '{font}': 원본 테마/기본 스타일에서")
        else:
            font = "맑은 고딕"
            notes.append("한글 글꼴을 찾지 못함 → 맑은 고딕")
    import check
    if not check.Checker().is_korean_font(font):
        notes.append(f"'{font}'는 한글 글꼴 목록에 없음 → 맑은 고딕으로 대체 (rules/typography.yaml 에 추가하면 그대로 씀)")
        font = "맑은 고딕"
    logo = overrides.get("logo") or next((f["logo"] for f in found if f["logo"]), None)
    logo = str(Path(logo).resolve()) if logo else None
    company = overrides.get("company") or next((f["company"] for f in found if f["company"]), "") or "회사명"
    return {"company": company, "fonts": {"korean": font, "latin": font},
            "colors": {"primary": primary, "accent": accent, "accent_text": accent_text, "highlight": "E07A1F", "text": "1F2937",
                       "muted": "6B7280", "light": light},
            "logo": logo,
            "top_band": any(f.get("band") for f in found),
            "logo_position": next((f["logo_pos"] for f in found if f.get("logo_pos")), "footer")}


def swatch(theme, out):
    from PIL import Image, ImageDraw, ImageFont
    C = theme["colors"]
    keys = ["primary", "accent", "accent_text", "light", "text", "muted"]
    im = Image.new("RGB", (len(keys) * 170 + 20, 150), "white")
    d = ImageDraw.Draw(im)
    try:
        f = ImageFont.truetype(str(Path.home() / ".fonts/Pretendard-Regular.otf"), 16)
    except Exception:
        f = ImageFont.load_default()
    for i, k in enumerate(keys):
        x = 10 + i * 170
        d.rectangle([x, 10, x + 160, 100], fill="#" + C[k], outline="#D0D5DD")
        d.text((x, 108), f"{k}\n#{C[k]}", fill="#374151", font=f)
    im.save(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--name", required=True)
    ap.add_argument("--company")
    ap.add_argument("--logo")
    ap.add_argument("--primary")
    ap.add_argument("--accent")
    ap.add_argument("--font")
    ap.add_argument("--out", default="profiles")
    ap.add_argument("--no-preview", action="store_true")
    a = ap.parse_args()
    out_dir = Path(a.out) / a.name
    out_dir.mkdir(parents=True, exist_ok=True)
    found = []
    for f in a.files:
        p = Path(f)
        if p.suffix.lower() == ".pptx":
            found.append(extract_pptx(p, out_dir))
        elif p.suffix.lower() == ".docx":
            found.append(extract_docx(p, out_dir))
        else:
            sys.exit(f"pptx/docx 만 지원: {p.name}")
    overrides = {k: getattr(a, k) for k in ("company", "logo", "primary", "accent", "font") if getattr(a, k)}
    for k in ("primary", "accent"):
        if k in overrides:
            overrides[k] = overrides[k].lstrip("#").upper()
    theme = decide(found, overrides)
    dump_yaml(theme, out_dir / "theme.yaml")
    swatch(theme, out_dir / "brand.png")
    # 표준 양식 10종 생성
    import build_templates
    build_templates.build(out_dir / "theme.yaml", out_dir, preview=not a.no_preview)
    # 프로필 이름을 회사 이름으로
    for pf in out_dir.glob("*/*/profile.yaml"):
        prof = load_yaml(pf)
        prof["title"] = f"{theme['company']} {prof['title']}"
        prof["brand"] = {"source": [f["source"] for f in found], "theme": "../../theme.yaml"}
        dump_yaml(prof, pf)
    L = [f"# 브랜드 표준화: {a.name}", "", f"원본: {', '.join(f['source'] for f in found)}", "",
         "![색 견본](brand.png)", "", "| 항목 | 값 |", "|---|---|",
         f"| 회사명 | {theme['company']} |", f"| 한글 글꼴 | {theme['fonts']['korean']} |"]
    L += [f"| {k} | `#{v}` |" for k, v in theme["colors"].items()]
    L += [f"| 로고 | {Path(theme['logo']).name if theme['logo'] else '없음 (회사명 글자로 대체)'} ({theme['logo_position']}) |",
          f"| 상단 색띠 | {'있음' if theme['top_band'] else '없음'} |", "", "## 판단 근거", ""]
    L += [f"- {n}" for n in notes]
    L += ["- highlight(주황)는 원본 색이 아니라 경고·예외 표시용 기본값 — 거의 쓰지 않음", "",
          "## 원본과 달라지는 점 (사용자에게 미리 알릴 것)", "",
          "- 구조: 손으로 그린 상자 → 표준 레이아웃 14종의 '칸' (AI가 채우고 고칠 수 있게)",
          "- 제목: 가운데 정렬이었다면 왼쪽 정렬로 (컨설팅 관례, 줄바꿈 안정)",
          "- 글자 크기·위치: 표준 타입 스케일·12단 그리드로 통일 (references/design-system.md)",
          f"- 화면 비율: {next((f['aspect'] for f in found if f.get('aspect')), '-')} → 16:9" if any(f.get("aspect") and f["aspect"] != "13.33x7.50" for f in found) else "- 화면 비율: 16:9 유지",
          "- 원본 모양을 한 픽셀도 바꾸면 안 되면 → workflows/1-onboard.md 의 'B. 원본 그대로 등록'"]
    L += ["", "## 다음 단계", "",
          "1. `pptx/*/preview.png`, `docx/*/preview.png` 로 결과 확인",
          f"2. 색·글꼴이 틀리면 `theme.yaml` 을 고치고: `python SKILL_DIR/scripts/build_templates.py --theme {out_dir}/theme.yaml --out {out_dir}`",
          f"3. 채우기: `python SKILL_DIR/scripts/fill.py --template {out_dir}/pptx/01-exec-report/template.pptx --content 내용.json --out 결과.pptx --company \"{theme['company']}\"`",
          f"4. **다음 대화에서도 쓰려면** `{a.name}.zip` 을 사용자에게 전달 → 스킬 폴더 `profiles/` 에 넣어 다시 업로드 (또는 make_company_skill.py 로 전용 스킬)"]
    (out_dir / "brand.md").write_text("\n".join(L), encoding="utf-8")
    print(f"[표준화] {a.name}: 대표색 #{theme['colors']['primary']} · 강조색 #{theme['colors']['accent']} · 글꼴 {theme['fonts']['korean']} · 로고 {'있음' if theme['logo'] else '없음'}")
    for n in notes:
        print("  -", n)
    import shutil
    zbase = Path(a.out) / a.name
    zpath = shutil.make_archive(str(zbase), "zip", root_dir=Path(a.out), base_dir=a.name)
    print(f"  → {out_dir}/ (양식 10종 + brand.md)")
    print(f"  → {zpath}  ← 사용자에게 전달 (claude.ai 는 대화가 끝나면 폴더가 사라짐)")


if __name__ == "__main__":
    main()
