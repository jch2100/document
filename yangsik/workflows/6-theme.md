# 6. theme — 기본 양식 10종을 우리 회사 색·글꼴로

회사 양식이 없지만 "우리 회사 느낌"이 필요할 때.

1. 사용자에게 받을 것: 회사명, 대표색 1~2개 (HEX 또는 로고 이미지에서 추출), 글꼴(모르면 맑은 고딕)
2. `SKILL_DIR/theme/default.yaml` 을 복사해 `my-theme.yaml` 로 저장 후 값 수정
   - primary: 제목·표 머리글 (짙은 색) / accent: 강조선·핵심메시지 / light: 박스 배경 (아주 연한 색)
   - 대비 확인: primary 위 흰 글자가 읽혀야 함 (밝은 노랑·연두 금지)
3. 실행: `python SKILL_DIR/scripts/build_templates.py --theme my-theme.yaml --out my-templates`
4. 이후 fill 에서 `--template my-templates/pptx/01-exec-report/template.pptx`
5. 미리보기(`my-templates/*/*/preview.png`)를 보여 주고 확정
