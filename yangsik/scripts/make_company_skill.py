#!/usr/bin/env python3
"""등록한 회사 양식 1개 → 그 양식만 아는 '전용 스킬' 폴더 + zip.

사용:
  python scripts/make_company_skill.py profiles/우리회사-보고 --skill-name acme-monthly-report \
      --triggers "월간 보고 만들어줘,경영회의 자료,월간 실적 PPT" --out dist/

결과: dist/acme-monthly-report/ (SKILL.md, template, profile.yaml, scripts/, rules/, references/) + .zip
→ claude.ai 스킬 업로드 / 동료에게 전달
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ys_common import SKILL_DIR, load_yaml  # noqa: E402

SCRIPTS = ["ys_common.py", "ys_pptx.py", "ys_docx.py", "fill.py", "check.py", "revise.py", "extract.py", "render.py"]
REFS = ["content-format.md", "writing-guide.md", "troubleshooting.md"]

TEMPLATE = """---
name: {name}
description: "{company_doc} 작성 전용 스킬. {purpose} 트리거: {triggers}"
---

# {title}

이 폴더 = `SKILL_DIR`. 양식: `SKILL_DIR/{file}` · 계약서: `SKILL_DIR/profile.yaml`
처음 한 번: `pip install -r SKILL_DIR/requirements.txt` (이미 있으면 생략)

## 순서
1. `profile.yaml` 을 읽는다 — 흐름, 칸(슬롯/스타일) 이름, 말투 규칙
2. `references/writing-guide.md` 를 읽는다. profile 의 writing 규칙이 우선
3. 사용자 자료에서 **결론 한 문장**을 먼저 정한다 (두괄식)
4. 내용 JSON 작성 — 형식 `references/content-format.md`, 키는 profile 의 이름 그대로
5. 실행:
   `python SKILL_DIR/scripts/fill.py --template SKILL_DIR/{file} --content 내용.json --out 결과.{ext} --profile SKILL_DIR/profile.yaml`
6. 검증 FAIL/WARN → 내용 JSON 고쳐 다시 실행 (최대 3회)
7. 수정 요청이면 `scripts/revise.py` (전체 재생성 금지)
8. 보고: 결과 파일 · 검증 판정 · 남은 경고 · 사람이 확인할 숫자

## 권장 흐름
{flow}

## 이 양식의 말투 규칙
{rules}

## 하지 않는 것
- 텍스트박스 새로 그리기, 글꼴·색 직접 지정, 기호·번호 직접 입력
- 원자료에 없는 숫자 만들기 (없으면 `[확인 필요]`)
- 원본 양식 덮어쓰기
"""


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("profile_dir")
    ap.add_argument("--skill-name", required=True, help="영문 소문자·숫자·하이픈 (예: acme-monthly-report)")
    ap.add_argument("--triggers", default="", help="사람들이 실제로 하는 말, 쉼표로 구분")
    ap.add_argument("--out", default="dist")
    a = ap.parse_args()
    if not re.fullmatch(r"[a-z0-9][a-z0-9-]{1,62}", a.skill_name):
        sys.exit("스킬 이름은 영문 소문자·숫자·하이픈만 (claude.ai 규칙). 예: acme-monthly-report")
    pdir = Path(a.profile_dir)
    prof = load_yaml(pdir / "profile.yaml")
    tpl = pdir / prof.get("file", "template.pptx")
    if not tpl.exists():
        sys.exit(f"양식 파일 없음: {tpl}")
    out = Path(a.out) / a.skill_name
    if out.exists():
        shutil.rmtree(out)
    (out / "scripts").mkdir(parents=True)
    (out / "references").mkdir()
    shutil.copytree(SKILL_DIR / "rules", out / "rules")
    for s in SCRIPTS:
        shutil.copy(SKILL_DIR / "scripts" / s, out / "scripts" / s)
    for r in REFS:
        shutil.copy(SKILL_DIR / "references" / r, out / "references" / r)
    shutil.copy(SKILL_DIR / "requirements.txt", out / "requirements.txt")
    shutil.copy(tpl, out / tpl.name)
    shutil.copy(pdir / "profile.yaml", out / "profile.yaml")
    samples = pdir / "samples"
    if samples.is_dir():
        shutil.copytree(samples, out / "samples")
    # check.py 가 rules 를 SKILL_DIR/rules 에서 찾도록 ys_common 의 SKILL_DIR 은 scripts/.. = out 이 됨 (그대로 동작)
    flow = prof.get("storyline") or prof.get("outline") or []
    rules = (prof.get("writing", {}).get("rules") or []) + [f"금지어: {', '.join(prof['writing']['banned'])}"] if prof.get("writing", {}).get("banned") else (prof.get("writing", {}).get("rules") or [])
    trig = ", ".join(f"'{t.strip()}'" for t in a.triggers.split(",") if t.strip()) or "'" + prof.get("title", "") + " 만들어줘'"
    md = TEMPLATE.format(
        name=a.skill_name, company_doc=prof.get("title", a.skill_name), purpose=prof.get("purpose", "").rstrip(".") + ".",
        triggers=trig, title=prof.get("title", a.skill_name), file=tpl.name, ext=tpl.suffix.lstrip("."),
        flow="\n".join(f"{i}. {x}" for i, x in enumerate(flow, 1)) or "- profile.yaml 참고",
        rules="\n".join(f"- {r}" for r in rules) or "- 개조식, 두괄식",
    )
    (out / "SKILL.md").write_text(md, encoding="utf-8")
    zpath = Path(a.out) / f"{a.skill_name}.zip"
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(out.rglob("*")):
            if f.is_file() and "__pycache__" not in f.parts:
                z.write(f, Path(a.skill_name) / f.relative_to(out))
    print(f"[전용 스킬] {out}/  →  {zpath}")
    print("  claude.ai: 설정 → 기능 → 스킬 → 업로드 /  Claude Code: ~/.claude/skills/ 에 폴더 복사")


if __name__ == "__main__":
    main()
