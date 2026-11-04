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
        r = run(SK / "scripts/fill.py", "--template", tpl, "--content", ROOT / "examples" / ex / "content.json", "--out", out, "--company", "테스트상사")
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

    print("6) 브랜드 표준화 → 10종 + 채우기")
    bs = ROOT / "workshop/brand-sample"
    r = run(SK / "scripts/standardize.py", bs / "greentec-template.pptx", bs / "greentec-report.docx", "--name", "g", "--company", "그린텍",
            "--out", t / "profiles", "--no-preview")
    if "#0061" not in r.stdout and "#0070" not in r.stdout:
        fails.append("브랜드 대표색(녹색) 추출 실패")
    if len(list((t / "profiles/g").glob("*/*/template.*"))) != 10:
        fails.append("표준화 결과가 10종이 아님")
    r = run(SK / "scripts/fill.py", "--template", t / "profiles/g/pptx/01-exec-report/template.pptx",
            "--content", ROOT / "examples/01-exec-report/content.json", "--out", t / "g.pptx")
    if "PASS" not in r.stdout:
        fails.append("브랜드 양식 채우기 PASS 아님")

    print("7) 컨설팅 원칙 검사: 주제어 제목·출처 누락")
    c2 = t / "c2.json"
    c2.write_text(json.dumps({"slides": [{"layout": "표", "제목": "현황", "표": {"header": ["a", "b"], "rows": [["x", "12"]]}}]},
                             ensure_ascii=False), encoding="utf-8")
    r = run(SK / "scripts/fill.py", "--template", "pptx/01", "--content", c2, "--out", t / "c2.pptx")
    if "주제어" not in r.stdout or "출처" not in r.stdout:
        fails.append("topic_title / missing_source 미검출")

    print("8) QA 회귀: 표 교체 수정 · 회사명 · 2줄 제목 · DOCX 강조/정보표")
    ops = t / "ops.json"
    ops.write_text(json.dumps({"ops": [
        {"slide": 6, "slot": "표", "value": {"header": ["구분", "A", "B"], "rows": [["비용", "1억 원", "2억 원"]], "highlight_col": 2}},
        {"slide": 2, "slot": "없는칸", "value": "x"}]}, ensure_ascii=False), encoding="utf-8")
    base = t / "ex.pptx"
    run(SK / "scripts/fill.py", "--template", "pptx/01", "--content", ROOT / "examples/01-exec-report/content.json", "--out", base)
    r = run(SK / "scripts/revise.py", base, "--ops", ops, "--out", t / "rv.pptx", expect=2)
    if "새 내용으로 교체" not in r.stdout or "✗" not in r.stdout:
        fails.append("revise: 표 교체 또는 실패 보고 안 됨")
    from pptx import Presentation as _P
    tb = next(s_ for s_ in _P(str(t / "rv.pptx")).slides[5].shapes if s_.has_table)
    if tb.table.cell(1, 2).text != "2억 원":
        fails.append("revise: 표 내용이 바뀌지 않음")
    r = run(SK / "scripts/check.py", base)
    if "회사명" not in r.stdout:
        fails.append("company_placeholder 미검출")
    r = run(SK / "scripts/fill.py", "--template", "pptx/01", "--content", ROOT / "examples/01-exec-report/content.json", "--out", t / "co.pptx", "--company", "테스트상사")
    if "회사명" in r.stdout or "상자를 꽉 채움" in r.stdout:
        fails.append("--company 미적용 또는 2줄 제목 오탐")
    dj = t / "d.json"
    dj.write_text(json.dumps({"blocks": [{"meta": [["a", "1"], ["b", "2"], ["c", "3"]]}, {"table": {"header": ["x", "y"], "rows": [["1", "2억 원"]], "highlight_col": 1}}, {"note": "끝."}]}, ensure_ascii=False), encoding="utf-8")
    r = run(SK / "scripts/fill.py", "--template", "docx/01", "--content", dj, "--out", t / "d.docx", "--company", "테스트상사")
    from docx import Document as _D
    dd = _D(str(t / "d.docx"))
    if any("※" in p_.text for p_ in dd.paragraphs if "끝" in p_.text):
        pass
    if "회사명" in dd.sections[0].header.paragraphs[0].text:
        fails.append("docx --company 미적용")

print()
if fails:
    print(f"실패 {len(fails)}건:", fails)
    sys.exit(1)
print("모든 테스트 통과")
