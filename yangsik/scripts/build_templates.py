#!/usr/bin/env python3
"""표준 양식 10종(PPTX 5 + DOCX 5)을 테마로부터 생성합니다.

사용:
  python scripts/build_templates.py                         # 기본 테마로 templates/ 재생성
  python scripts/build_templates.py --theme my.yaml --out my-templates/   # 우리 회사 색·글꼴로

각 양식 폴더에 생기는 것: template.(pptx|docx), profile.yaml, preview.png(LibreOffice 있을 때)
"""
from __future__ import annotations

import argparse
import io
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx import Presentation  # noqa: E402

import ys_docx  # noqa: E402
import ys_pptx  # noqa: E402
from ys_common import TEMPLATES_DIR, dump_yaml, load_json, load_theme  # noqa: E402

LAYOUT_USE = {
    "표지": "문서 첫 장", "목차": "목차·흐름", "섹션": "큰 단락 구분 (10장 이상일 때)",
    "본문": "개조식 근거 (가장 많이 씀)", "2단비교": "전/후, 문제/원인, 나쁜 예/좋은 예",
    "3단카드": "요약(상황|발견|제언), 세 가지 축, KPI 셋", "프로세스": "4단계 추진 방식·절차",
    "매트릭스": "2x2 우선순위 (영향×가능성 등)", "표": "비교표·일정·R&R (숫자는 자동 오른쪽 정렬)",
    "차트": "추이·비교 수치 + 시사점", "빅넘버": "숫자 하나가 결론일 때", "핵심메시지": "요약·전환 한 문장",
    "이미지": "화면·구성도·사진 + 설명", "마무리": "요청·다음 단계·문의",
}
SLOT_LIMITS = {  # 슬롯별 권장 분량 (check.py 넘침 추정과 함께 씀)
    "트래커": "섹션 이름, 15자 이내", "제목": "결론 문장(액션 타이틀), 2줄·45자 이내",
    "본문": "항목 3~6개, 하위 포함 10줄 이내", "좌측본문": "항목 2~5개", "우측본문": "항목 2~5개",
    "카드1제목": "10자 이내 또는 숫자", "카드1본문": "항목 1~3개", "단계1제목": "8자 이내", "단계1본문": "항목 1~3개",
    "숫자": "숫자+단위 하나", "설명": "한 줄", "해설": "2~3개", "부제": "한 줄", "보조설명": "한 줄",
    "좌상": "이름 + 항목 1~3개",
}

WRITING = {
    "style": "gaejosik",
    "rules": [
        "액션 타이틀: 장 제목 = 그 장의 결론 문장 (제목만 이어 읽어도 이야기가 됨)",
        "한 장 = 한 메시지. 제목에 '그리고'가 들어가면 장을 나눔",
        "개조식: 명사형·'~함/~임'으로 끝냄, '~습니다' 금지(마무리 장 제외)",
        "한 항목 = 한 주장, 한 줄 40자 안팎",
        "숫자는 단위와 함께 (예: 1,200만 원, 15%p)",
        "날짜는 '2026. 11. 4.' 형식",
        "기호·번호는 직접 쓰지 않음 (양식이 자동으로 붙임)",
    ],
    "banned": ["혁신적인", "획기적인", "다양한", "~하는 것이 중요합니다", "~라고 할 수 있습니다", "—", "이모지"],
}


def pptx_profile(name, skeleton, prs):
    layouts = {}
    for l in prs.slide_layouts:
        slots = ys_pptx.layout_slots(l)
        layouts[l.name] = {
            "use": LAYOUT_USE.get(l.name, ""),
            "slots": {k: {"type": t.lower(), **({"limit": SLOT_LIMITS[k]} if k in SLOT_LIMITS else {})}
                      for k, (idx, t) in slots.items()},
        }
    return {
        "name": name,
        "kind": "pptx",
        "title": skeleton["meta"]["title"],
        "purpose": skeleton["meta"]["purpose"],
        "audience": skeleton["meta"]["audience"],
        "length": skeleton["meta"]["length"],
        "file": "template.pptx",
        "storyline": [f'{s["layout"]}: {s.get("제목", "")}' for s in skeleton["slides"]],
        "layouts": layouts,
        "writing": WRITING,
        "check": {"fail_on": ["off_template_shape", "missing_korean_font", "overflow_high"],
                  "warn_on": ["overflow_maybe", "ai_tone", "ending_style", "empty_slot", "topic_title", "missing_source"]},
    }


def docx_profile(name, skeleton):
    return {
        "name": name,
        "kind": "docx",
        "title": skeleton["meta"]["title"],
        "purpose": skeleton["meta"]["purpose"],
        "audience": skeleton["meta"]["audience"],
        "length": skeleton["meta"]["length"],
        "file": "template.docx",
        "outline": [next(iter(b.values())) for b in skeleton["blocks"] if next(iter(b)) in ("h1", "h2")],
        "blocks": {
            "title": "문서 제목", "subtitle": "부제", "meta": "정보표 [[라벨, 값], ...]",
            "approval": "결재란 [직위, ...]", "summary": "요약 박스 (문자열 또는 목록)",
            "h1": "대제목 → 자동 '1.'", "h2": "중제목 → 자동 '가.'", "h3": "소제목 → 자동 '1)'",
            "b": "개조식 목록 → 자동 □ ○ - (중첩 목록 = 하위 단계)", "p": "서술형 문단",
            "note": "※ 참고", "table": "{header, rows, col_widths, label_col}", "cover": "표지 {title, subtitle, date, company}",
            "toc": "목차 필드", "pagebreak": "쪽 나눔", "image": "그림 경로",
        },
        "styles": {"title": ys_docx.S_TITLE, "h1": ys_docx.S_H[1], "h2": ys_docx.S_H[2], "h3": ys_docx.S_H[3],
                   "bullet1": ys_docx.S_B[0], "bullet2": ys_docx.S_B[1], "bullet3": ys_docx.S_B[2],
                   "note": ys_docx.S_NOTE, "summary": ys_docx.S_SUMMARY, "table": ys_docx.S_TABLE},
        "writing": {**WRITING, "rules": WRITING["rules"] + ["서술형 문단(p)은 제안서·품의 문구에만"]},
        "check": {"fail_on": ["missing_korean_font", "direct_symbol"],
                  "warn_on": ["ai_tone", "ending_style", "direct_format", "empty_paragraphs"]},
    }


def build(theme_path=None, out_dir=None, preview=True):
    theme = load_theme(theme_path)
    src = TEMPLATES_DIR
    out = Path(out_dir) if out_dir else TEMPLATES_DIR
    base_pptx = ys_pptx.build_base(theme)
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        base_path = tmp / "base.pptx"
        base_path.write_bytes(base_pptx)
        for d in sorted((src / "pptx").iterdir()):
            if not (d / "skeleton.json").exists():
                continue
            target = out / "pptx" / d.name
            target.mkdir(parents=True, exist_ok=True)
            sk = load_json(d / "skeleton.json")
            if target != d:
                shutil.copy(d / "skeleton.json", target / "skeleton.json")
            ys_pptx.fill(base_path, sk, target / "template.pptx")
            prs = Presentation(str(target / "template.pptx"))
            dump_yaml(pptx_profile(d.name, sk, prs), target / "profile.yaml")
            print("PPTX", d.name)
        for d in sorted((src / "docx").iterdir()):
            if not (d / "skeleton.json").exists():
                continue
            target = out / "docx" / d.name
            target.mkdir(parents=True, exist_ok=True)
            sk = load_json(d / "skeleton.json")
            if target != d:
                shutil.copy(d / "skeleton.json", target / "skeleton.json")
            base = ys_docx.build_base(theme, sk["meta"].get("label", sk["meta"]["title"]))
            bpath = tmp / f"{d.name}.docx"
            base.save(str(bpath))
            ys_docx.fill(bpath, sk, target / "template.docx")
            dump_yaml(docx_profile(d.name, sk), target / "profile.yaml")
            print("DOCX", d.name)
        if preview:
            from render import render
            for f in sorted(out.glob("*/*/template.*")):
                r = render(f, tmp / f.parent.name, sheet=True, dpi=45, keep_pages=False)
                if r and r.get("sheet"):
                    shutil.copy(r["sheet"], f.parent / "preview.png")
                    print("  preview", f.parent.name)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--theme", help="테마 YAML (기본: theme/default.yaml)")
    ap.add_argument("--out", help="출력 폴더 (기본: templates/ 덮어쓰기)")
    ap.add_argument("--no-preview", action="store_true")
    a = ap.parse_args()
    build(a.theme, a.out, not a.no_preview)
