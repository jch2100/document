#!/usr/bin/env python3
"""회사 양식 진단 → 등급(A/B/C) + 할 일 + profile.yaml 초안.

사용:
  python scripts/inspect_template.py 우리회사양식.pptx --name 우리회사 [--out profiles/]
  python scripts/inspect_template.py 우리회사보고서.docx --name 우리회사

결과 (profiles/우리회사/ 에 생성):
  template.pptx|docx   원본 복사본 (원본은 건드리지 않음)
  profile.yaml         양식 계약서 초안 — 사람이 검토해서 고침
  diagnosis.md         진단 리포트

등급
  A  바로 채우기 가능 (레이아웃·플레이스홀더/스타일이 제대로 있음)
  B  normalize.py --auto 한 번이면 가능 (이름·글꼴 정리 수준)
  C  사람 손이 필요 (텍스트박스로 그린 양식, 스타일 없이 손으로 서식)
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ys_common import SKILL_DIR, dump_yaml, load_yaml  # noqa: E402

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
GENERIC = re.compile(r"^(Title|Subtitle|Text|Content|Picture|Table|Chart|Date|Footer|Slide Number|제목|부제목|텍스트|내용|그림|표|차트|날짜|바닥글|슬라이드 번호)( Placeholder)?\s*\d+$|개체 틀 \d+$|Placeholder \d+$", re.I)


def _typ(pf):
    return str(pf.type).split(".")[-1].split(" ")[0].lower()


def inspect_pptx(path):
    from pptx import Presentation
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from lxml import etree
    import check as chk
    ty = load_yaml(SKILL_DIR / "rules" / "typography.yaml")
    korean = lambda n: bool(n) and (any(k.lower() in n.lower() for k in ty["korean_fonts"]) or bool(chk.HANGUL.search(n)))  # noqa: E731

    prs = Presentation(str(path))
    info = {"slide_size": f"{prs.slide_width/914400:.2f} x {prs.slide_height/914400:.2f} in", "layouts": [], "issues": [], "todo": []}
    generic_names = 0
    body_layouts = 0
    used_layouts = Counter(s.slide_layout.name for s in prs.slides)
    for layout in prs.slide_layouts:
        phs = []
        for shp in layout.placeholders:
            t = _typ(shp.placeholder_format)
            if t in ("date", "footer", "slide_number"):
                continue
            phs.append({"name": shp.name, "type": t, "idx": shp.placeholder_format.idx,
                        "size_in": f"{(shp.width or 0)/914400:.1f}x{(shp.height or 0)/914400:.1f}"})
            if GENERIC.match(shp.name or ""):
                generic_names += 1
        if any(p["type"] in ("body", "object") for p in phs):
            body_layouts += 1
        info["layouts"].append({"name": layout.name, "used_by_slides": used_layouts.get(layout.name, 0), "placeholders": phs})

    # 슬라이드 실사용 분석
    tb_chars = ph_chars = 0
    tb_slides = []
    fonts = Counter()
    for i, s in enumerate(prs.slides, 1):
        n_tb = 0
        for shp in s.shapes:
            if not shp.has_text_frame:
                continue
            txt = shp.text_frame.text.strip()
            for p in shp.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.name:
                        fonts[r.font.name] += 1
            if shp.is_placeholder:
                ph_chars += len(txt)
            elif txt:
                tb_chars += len(txt)
                n_tb += 1
        if n_tb:
            tb_slides.append((i, n_tb))
    theme = etree.fromstring(prs.slide_masters[0].part.part_related_by(RT.THEME).blob)
    minor = theme.find(f".//{A}minorFont")
    ea = minor.find(f"{A}ea").get("typeface") if minor.find(f"{A}ea") is not None else ""
    latin = minor.find(f"{A}latin").get("typeface")
    hang = next((x.get("typeface") for x in minor.findall(f"{A}font") if x.get("script") == "Hang"), "")
    info["fonts"] = {"theme_latin": latin, "theme_korean": ea or hang or "(없음)", "direct_fonts": dict(fonts.most_common(5))}
    total = tb_chars + ph_chars
    tb_ratio = tb_chars / total if total else 0
    info["textbox_ratio"] = round(tb_ratio, 2)

    # 등급
    grade = "A"
    if body_layouts < 2:
        grade = "C"
        info["issues"].append(f"본문 플레이스홀더가 있는 레이아웃이 {body_layouts}개 — AI가 채울 '칸'이 없음")
        info["todo"].append("PowerPoint 보기 → 슬라이드 마스터에서 레이아웃마다 '개체 틀 삽입 → 텍스트' 추가, 또는 normalize.py --promote")
    if tb_ratio > 0.5 and len(prs.slides):
        grade = "C"
        info["issues"].append(f"슬라이드 글의 {tb_ratio:.0%}가 일반 텍스트박스 (예: {', '.join(f'{i}번({n}개)' for i, n in tb_slides[:5])})")
        info["todo"].append("대표 슬라이드마다 normalize.py --promote <슬라이드번호> --as-layout <이름> 으로 텍스트박스를 슬롯으로 승격")
    elif tb_ratio > 0.2:
        grade = max(grade, "B")
        info["issues"].append(f"슬라이드 글의 {tb_ratio:.0%}가 텍스트박스")
        info["todo"].append("자주 쓰는 텍스트박스 위치만 --promote 로 슬롯화")
    if generic_names:
        grade = max(grade, "B")
        info["issues"].append(f"플레이스홀더 이름이 기본값('Text Placeholder 2' 등) {generic_names}개 — AI가 어느 칸인지 구분 못함")
        info["todo"].append("normalize.py --auto (제목/본문/부제… 로 자동 이름 붙임)")
    if not korean(ea) and not korean(hang) and not korean(latin):
        grade = max(grade, "B")
        info["issues"].append("테마에 한글 글꼴 미지정 → PC마다 다른 글꼴로 보일 수 있음")
        info["todo"].append("normalize.py --auto --korean-font '맑은 고딕'")
    default_names = sum(1 for l in prs.slide_layouts if re.match(r"^(Title|Two|Comparison|Blank|Section|Picture|Content|Vertical|Title and|제목|빈 화면|비교|구역|콘텐츠|세로)", l.name))
    if default_names >= 3:
        info["issues"].append(f"레이아웃 이름이 기본값 {default_names}개 (예: 'Title and Content') — 용도가 드러나는 이름 권장")
        info["todo"].append("normalize.py --rename-layouts \"Title and Content=본문,Title Slide=표지\"")
    info["grade"] = grade
    return info, prs


def draft_profile_pptx(name, info, prs):
    import ys_pptx
    layouts = {}
    for l in prs.slide_layouts:
        slots = ys_pptx.layout_slots(l)
        if not slots:
            continue
        layouts[l.name] = {"use": "[용도를 적으세요]", "slots": {k: {"type": t.lower()} for k, (i, t) in slots.items()}}
    storyline = [s.slide_layout.name for s in prs.slides]
    return {
        "name": name, "kind": "pptx", "title": f"{name} 표준 PPT", "purpose": "[이 양식을 언제 쓰는지]",
        "audience": "[주 독자]", "file": "template.pptx", "grade": info["grade"],
        "storyline": storyline or ["[대표 슬라이드 순서]"], "layouts": layouts,
        "writing": {"style": "gaejosik", "rules": ["두괄식", "개조식: '~함/~임'", "숫자는 단위와 함께"], "banned": []},
        "check": {"fail_on": ["off_template_shape", "missing_korean_font", "overflow_high"],
                  "warn_on": ["overflow_maybe", "ai_tone", "ending_style", "empty_slot"]},
    }


def inspect_docx(path):
    from docx import Document
    import check as chk
    doc = Document(str(path))
    info = {"issues": [], "todo": []}
    style_count = Counter()
    manual = re.compile(load_yaml(SKILL_DIR / "rules" / "gaejosik.yaml")["manual_symbols"])
    manual_n = empty_streaks = runs = direct = 0
    streak = 0
    for p in doc.paragraphs:
        t = p.text
        st = p.style.name if p.style is not None else "Normal"
        if not t.strip():
            streak += 1
            if streak == 2:
                empty_streaks += 1
            continue
        streak = 0
        style_count[st] += 1
        has_num = p._p.pPr is not None and p._p.pPr.numPr is not None
        if manual.match(t) and not has_num:
            manual_n += 1
        for r in p.runs:
            runs += 1
            rpr = r._r.rPr
            if rpr is not None and (rpr.find(f"{W}sz") is not None or rpr.find(f"{W}rFonts") is not None):
                direct += 1
    n_par = sum(style_count.values()) or 1
    heading_used = sum(c for s, c in style_count.items() if re.search(r"heading|제목 ?\d|^제목\d", s, re.I))
    info["styles_used"] = dict(style_count.most_common(12))
    info["manual_symbols"] = manual_n
    info["direct_format_ratio"] = round(direct / runs, 2) if runs else 0
    info["tables"] = len(doc.tables)
    rf = doc.styles.element.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}rFonts")
    info["korean_font"] = (rf.get(f"{W}eastAsia") or rf.get(f"{W}eastAsiaTheme")) if rf is not None else None
    sec = doc.sections[0]
    info["page"] = f"{sec.page_width/360000:.1f} x {sec.page_height/360000:.1f} cm"
    have = {s.name for s in doc.styles}
    missing = [s for s in ("개조식 1", "개조식 2", "개조식 3", "참고", "요약 박스", "양식대로 표") if s not in have]
    info["missing_yangsik_styles"] = missing

    grade = "A"
    if heading_used == 0:
        grade = "C" if manual_n > 3 else "B"
        info["issues"].append("제목 스타일(Heading/제목 1~3)을 쓰지 않음 — 목차·탐색창·자동번호 불가")
        info["todo"].append("normalize.py --headings-from-numbers ('1. ', '가. '로 시작하는 짧은 문단을 제목 스타일로)")
    if manual_n:
        grade = "C" if manual_n / n_par > 0.3 else max(grade, "B")
        info["issues"].append(f"기호·번호를 손으로 입력한 문단 {manual_n}개 (□ ○ - 1. 가. 등)")
        info["todo"].append("normalize.py --fix-symbols (기호 삭제 + 개조식 스타일 적용)")
    if info["direct_format_ratio"] > 0.3:
        grade = max(grade, "B")
        info["issues"].append(f"글자 조각의 {info['direct_format_ratio']:.0%}에 글꼴·크기 직접 지정 — 스타일 수정이 반영 안 됨")
        info["todo"].append("normalize.py --clear-direct (직접 서식 제거, 굵게·기울임은 유지)")
    if empty_streaks:
        info["issues"].append(f"빈 줄로 간격 띄운 곳 {empty_streaks}곳")
        info["todo"].append("normalize.py --remove-empty")
    if missing:
        grade = max(grade, "B")
        info["issues"].append(f"채우기에 필요한 스타일 없음: {', '.join(missing)}")
        info["todo"].append("normalize.py --add-styles (회사 글꼴·색을 따라 생성)")
    kf = info["korean_font"]
    if not kf or (not kf.startswith("minor") and not kf.startswith("major") and not chk.HANGUL.search(kf)
                  and not any(k.lower() in kf.lower() for k in load_yaml(SKILL_DIR / "rules" / "typography.yaml")["korean_fonts"])):
        grade = max(grade, "B")
        info["issues"].append(f"기본 한글 글꼴: {kf or '(없음)'}")
        info["todo"].append("normalize.py --korean-font '맑은 고딕'")
    info["grade"] = grade
    return info, doc


def draft_profile_docx(name, info, doc):
    have = {s.name for s in doc.styles}
    pick = lambda *cands: next((c for c in cands if c in have), cands[-1])  # noqa: E731
    return {
        "name": name, "kind": "docx", "title": f"{name} 표준 문서", "purpose": "[이 양식을 언제 쓰는지]",
        "audience": "[주 독자]", "file": "template.docx", "grade": info["grade"],
        "outline": ["[대제목 1]", "[대제목 2]"],
        "styles": {"title": pick("Title", "제목"), "h1": pick("Heading 1", "제목 1"), "h2": pick("Heading 2", "제목 2"),
                   "h3": pick("Heading 3", "제목 3"), "bullet1": "개조식 1", "bullet2": "개조식 2", "bullet3": "개조식 3",
                   "note": "참고", "summary": "요약 박스", "table": "양식대로 표"},
        "writing": {"style": "gaejosik", "rules": ["두괄식", "개조식: '~함/~임'"], "banned": []},
        "check": {"fail_on": ["missing_korean_font"], "warn_on": ["ai_tone", "ending_style", "direct_symbol", "direct_format", "empty_paragraphs"]},
    }


def write_diag(path, info, kind, out):
    g = info["grade"]
    desc = {"A": "바로 채우기 가능", "B": "normalize.py 한 번이면 가능", "C": "사람 손이 필요 (반자동)"}[g]
    L = [f"# 양식 진단: {Path(path).name}", "", f"## 등급: {g} — {desc}", ""]
    if info["issues"]:
        L += ["## 발견한 문제", ""] + [f"- {x}" for x in info["issues"]] + [""]
        L += ["## 할 일 (순서대로)", ""] + [f"{i}. {x}" for i, x in enumerate(dict.fromkeys(info["todo"]), 1)] + [""]
    else:
        L += ["문제 없음. profile.yaml 의 '[ ]' 부분(용도·독자)만 채우세요.", ""]
    if kind == "pptx":
        L += [f"## 레이아웃 ({len(info['layouts'])}개, 슬라이드 크기 {info['slide_size']})", "",
              "| 레이아웃 | 사용 슬라이드 | 슬롯(플레이스홀더) |", "|---|---|---|"]
        for l in info["layouts"]:
            slots = ", ".join(f"{p['name']}({p['type']})" for p in l["placeholders"]) or "없음"
            L.append(f"| {l['name']} | {l['used_by_slides']} | {slots} |")
        L += ["", f"- 글꼴: 테마 영문 `{info['fonts']['theme_latin']}` / 한글 `{info['fonts']['theme_korean']}`",
              f"- 직접 지정된 글꼴: {info['fonts']['direct_fonts'] or '없음'}",
              f"- 텍스트박스 비율: {info['textbox_ratio']:.0%}"]
    else:
        L += ["## 사용 중인 스타일 (문단 수)", ""] + [f"- {k}: {v}" for k, v in info["styles_used"].items()]
        L += ["", f"- 용지: {info['page']}", f"- 기본 한글 글꼴: {info['korean_font']}", f"- 표: {info['tables']}개",
              f"- 직접 서식 비율: {info['direct_format_ratio']:.0%}", f"- 손으로 입력한 기호/번호: {info['manual_symbols']}개"]
    L += ["", "---", "다음 단계: 할 일을 처리한 뒤 `python scripts/inspect_template.py` 를 다시 돌려 등급을 확인하세요."]
    out.write_text("\n".join(L), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--name", required=True, help="프로필 이름 (예: 우리회사-보고서)")
    ap.add_argument("--out", default="profiles", help="프로필 상위 폴더 (기본 ./profiles)")
    ap.add_argument("--no-copy", action="store_true", help="진단만 (복사·프로필 생성 안 함)")
    a = ap.parse_args()
    src = Path(a.file)
    kind = src.suffix.lower().lstrip(".")
    if kind not in ("pptx", "docx"):
        sys.exit("pptx/docx만 진단합니다. (.ppt/.doc 는 PowerPoint/Word에서 '다른 이름으로 저장' → pptx/docx)")
    if kind == "pptx":
        info, obj = inspect_pptx(src)
        prof = draft_profile_pptx(a.name, info, obj)
    else:
        info, obj = inspect_docx(src)
        prof = draft_profile_docx(a.name, info, obj)
    out_dir = Path(a.out) / a.name
    out_dir.mkdir(parents=True, exist_ok=True)
    diag = out_dir / "diagnosis.md"
    write_diag(src, info, kind, diag)
    if not a.no_copy:
        dst = out_dir / f"template.{kind}"
        if src.resolve() != dst.resolve():
            shutil.copy(src, dst)
        pf = out_dir / "profile.yaml"
        if pf.exists():
            print(f"· 기존 profile.yaml 유지 (새 초안: profile.draft.yaml)")
            pf = out_dir / "profile.draft.yaml"
        dump_yaml(prof, pf)
    print(f"[진단] {src.name}: 등급 {info['grade']}")
    for x in info["issues"]:
        print("  -", x)
    print(f"  → {diag}")


if __name__ == "__main__":
    main()
