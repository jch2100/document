# 브랜드 표준화 예시: 가상의 회사 '그린텍'

- `greentec-template.pptx` : 4:3, Arial, 텍스트박스로 그린 전형적인 C등급 양식 (녹색 브랜드 + 모서리 로고)
- `greentec-report.docx` : 머리말 로고 + 녹색 제목의 보고서

```bash
python yangsik/scripts/standardize.py workshop/brand-sample/greentec-template.pptx workshop/brand-sample/greentec-report.docx --name 그린텍 --company 그린텍
```
→ `profiles/그린텍/` 에 녹색 브랜드의 표준 양식 10종 + `brand.md`(판단 근거)
