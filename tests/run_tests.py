#!/usr/bin/env python3
"""회귀 테스트: 양식 생성 → 10종 채우기 → 검증 → 수정 → 추출 → 등록(진단·표준화) → 테마.

실행: python tests/run_tests.py   (LibreOffice 불필요)
"""
import json
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SK = ROOT / "yangsik"
PY = sys.executable
fails = []


def run(*args, expect=0, cwd=None):
    r = subprocess.run([PY, *map(str, args)], capture_output=True, text=True, cwd=cwd)
    ok = (r.returncode == expect) if expect is not None else True
    name = " ".join(Path(str(a)).name for a in args[:2])
    print(("  ok  " if ok else "  FAIL"), name, f"(exit {r.returncode})")
    if not ok:
        fails.append(name)
        print(r.stdout[-1500:], r.stderr[-1500:])
    return r


with tempfile.TemporaryDirectory() as t:
    t = Path(t)
    print("1) 테마 → 양식 10종 생성")
    run(SK / "scripts/build_templates.py", "--out", t / "tpl", "--no-preview")
    custom = t / "theme.yaml"
    custom.write_text('company: "테스트주식회사"\ncolors:\n  primary: "7A1F1F"\n  accent: "C0392B"\n', encoding="utf-8")
    run(SK / "scripts/build_templates.py", "--theme", custom, "--out", t / "tpl2", "--no-preview")
    assert len(list((t / "tpl").glob("*/*/template.*"))) == 10, "양식 10종이 아님"

    print("2) 10종 뼈대로 채우기 + 검증 (FAIL 없어야 함)")
    for sk in sorted((SK / "templates").glob("*/*/skeleton.json")):
        kind = sk.parent.parent.name
        out = t / f"{sk.parent.name}.{kind}"
        run(SK / "scripts/fill.py", "--template", sk.parent / f"template.{kind}", "--content", sk, "--out", out)

    print("3) 예시 내용 → PASS")
    for ex, tpl, ext in (("01-exec-report", "pptx/01", "pptx"), ("01-1page-report", "docx/01", "docx")):
        out = t / f"ex.{ext}"
        r = run(SK / "scripts/fill.py", "--template", tpl, "--content", ROOT / "examples" / ex / "content.json", "--out", out)
        if "PASS" not in r.stdout:
            fails.append(f"{ex} not PASS")
        run(SK / "scripts/revise.py", out, "--ops", ROOT / "examples" / ex / "revise-ops.json", "--out", t / f"ex_v2.{ext}")
        run(SK / "scripts/extract.py", t / f"ex_v2.{ext}", "--out", t / f"ex_v2_{ext}.json")
        data = json.loads((t / f"ex_v2_{ext}.json").read_text(encoding="utf-8"))
        # 왕복: 추출한 내용으로 다시 채우기
        run(SK / "scripts/fill.py", "--template", tpl, "--content", t / f"ex_v2_{ext}.json", "--out", t / f"rt.{ext}")

    print("4) 망가진 샘플 → PPTX는 FAIL 이어야 함")
    run(SK / "scripts/check.py", ROOT / "workshop/bad-samples/ai-made-bad.pptx", expect=1)
    r = run(SK / "scripts/check.py", ROOT / "workshop/bad-samples/ai-made-bad.docx", expect=None)
    if "direct_symbol" not in (t.name and "") and "기호" not in r.stdout:
        fails.append("bad docx: 기호 직접 입력 미검출")

    print("5) 등록: 진단 → 표준화 → 재진단 → 채우기")
    run(SK / "scripts/inspect_template.py", ROOT / "workshop/bad-samples/ai-made-bad.pptx", "--name", "t1", "--out", t / "profiles")
    p = t / "profiles/t1/template.pptx"
    run(SK / "scripts/normalize.py", p, "--auto", "--promote", "1", "--as-layout", "요약", "--slot-names", "제목,본문", "--out", p)
    r = run(SK / "scripts/inspect_template.py", p, "--name", "t1", "--out", t / "profiles", "--no-copy")
    if "등급 C" in r.stdout:
        fails.append("표준화 후에도 C")
    c = t / "c.json"
    c.write_text(json.dumps({"slides": [{"layout": "요약", "제목": "테스트", "본문": ["항목"]}]}, ensure_ascii=False), encoding="utf-8")
    run(SK / "scripts/fill.py", "--template", p, "--content", c, "--out", t / "u.pptx")
    run(SK / "scripts/inspect_template.py", ROOT / "workshop/bad-samples/ai-made-bad.docx", "--name", "t2", "--out", t / "profiles")
    d = t / "profiles/t2/template.docx"
    run(SK / "scripts/normalize.py", d, "--all", "--out", d)
    run(SK / "scripts/fill.py", "--template", d, "--content", ROOT / "examples/01-1page-report/content.json", "--out", t / "u.docx")
    run(SK / "scripts/make_company_skill.py", t / "profiles/t1", "--skill-name", "t-one", "--out", t / "dist")

print()
if fails:
    print(f"실패 {len(fails)}건:", fails)
    sys.exit(1)
print("모든 테스트 통과")
