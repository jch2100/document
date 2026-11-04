# profile.yaml 스키마

양식 1개 = 폴더 1개 = `template.(pptx|docx)` + `profile.yaml`. profile 은 AI와 사람이 같이 읽는 **양식 계약서**다.

```yaml
name: 우리회사-보고         # 폴더 이름과 같게
kind: pptx                  # pptx | docx
title: 우리회사 경영 보고
purpose: 월간 경영회의 보고  # 언제 쓰나 (AI가 양식 고를 때 봄)
audience: 대표·임원
length: 7~10장
file: template.pptx
grade: A                    # inspect 결과

# PPTX: 레이아웃 = 장 종류, slots = 채울 칸 이름 (fill 의 키)
layouts:
  본문:
    use: 개조식 설명
    slots:
      제목: {type: title, limit: "한 줄, 25자"}
      핵심메시지: {type: body, limit: "한 문장"}
      본문: {type: body, limit: "항목 3~6개"}
storyline:                  # 권장 흐름 (레이아웃: 제목)
  - "표지: …"
  - "핵심메시지: 요약"

# DOCX: 스타일 이름 매핑 (회사 스타일 이름이 다르면 여기서 연결)
styles:
  title: Title
  h1: Heading 1             # 예: "제목 1", "1수준"
  h2: Heading 2
  h3: Heading 3
  bullet1: 개조식 1          # 예: 회사의 "본문 □"
  bullet2: 개조식 2
  bullet3: 개조식 3
  note: 참고
  summary: 요약 박스
  table: 양식대로 표
outline: [추진 배경, 현황 및 문제점, 개선 방안]

writing:
  style: gaejosik           # gaejosik | prose
  rules: ["두괄식", "~함/~임 종결"]
  banned: [혁신적인, 다양한]   # check 가 경고

check:
  fail_on: [off_template_shape, missing_korean_font, overflow_high]
  warn_on: [ai_tone, ending_style]
  ignore: [gongmun]
```
