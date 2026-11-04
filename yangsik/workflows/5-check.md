# 5. check — 문서 검증

```
python SKILL_DIR/scripts/check.py 파일 [--profile profile.yaml]
```
결과: 화면 요약 + `파일_pptx_check.md`. FAIL 이면 종료코드 1.

## 결과 설명 방법

- 판정 → 실패 항목 → 경고 상위 3개 → "고치는 법" 순서로 짧게.
- 남이 만든 파일이면 **고쳐 줄지** 제안: 양식 밖 텍스트박스가 많으면 extract → 기본 양식으로 fill 하는 '옮겨 담기'를 권한다.
- 넘침은 글자 폭 **추정치**임을 밝힌다.

## 항목 사전

| 코드 | 뜻 | 기본 수준 |
|---|---|---|
| off_template_shape | 레이아웃 슬롯이 아닌 텍스트박스 | FAIL |
| missing_korean_font | 한글 글꼴 미지정 (PC마다 다르게 보임) | FAIL |
| overflow_high / overflow_maybe | 글이 상자를 넘침 (확실/가능) | FAIL / WARN |
| table_tall | 표가 아래 여백 침범 | WARN |
| line_break | 어절 단위 줄바꿈 꺼짐 | WARN |
| empty_slot | 빈 슬롯 | WARN |
| guide_left | `[안내문]` 이 남음 | WARN |
| ai_tone | AI 티 나는 표현 (rules/ai-tone.yaml) | WARN |
| ending_style | 개조식인데 '~습니다' | WARN |
| direct_symbol | Word에서 기호·번호를 손으로 입력 | WARN |
| direct_format | 직접 서식 과다 | WARN |
| empty_paragraphs | 빈 줄로 간격 | WARN |
| heading_number_dup | 자동 번호 + 손 번호 | WARN |
| gongmun / item_long / korean_lang / compat_toc / table_no_style | 참고 | INFO |

수준은 profile.yaml 의 `check.fail_on / warn_on / ignore` 로 바꾼다. 규칙 문구는 `rules/*.yaml` 에서 회사가 직접 고친다.
