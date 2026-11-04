# 예시: 입력 → 결과 → 검증 → 수정

| 폴더 | 양식 | 입력 | 결과 | 수정 |
|---|---|---|---|---|
| 01-exec-report | pptx/01 경영진 보고 | content.json | result.pptx (PASS) | revise-ops.json → result_v2.pptx |
| 01-1page-report | docx/01 1페이지 보고서 | content.json | result.docx (PASS) | revise-ops.json → result_v2.docx |

다시 만들기:
```bash
python yangsik/scripts/fill.py --template pptx/01 --content examples/01-exec-report/content.json --out examples/01-exec-report/result.pptx
```
```bash
python yangsik/scripts/revise.py examples/01-exec-report/result.pptx --ops examples/01-exec-report/revise-ops.json --out examples/01-exec-report/result_v2.pptx
```
