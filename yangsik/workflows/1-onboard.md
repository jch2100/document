# 1. onboard — 회사 양식 등록·표준화

목표: 회사 양식 1개 → `profiles/<이름>/` (template + profile.yaml + diagnosis.md). 이후 fill 은 이 폴더를 쓴다.

## 순서

1. **확장자 확인**: .ppt/.doc/.hwp 는 처리 불가. → "PowerPoint/Word에서 '다른 이름으로 저장 → pptx/docx' 후 올려 주세요". HWP는 v2 예정이라고 안내.
2. **보안 한 줄 안내** (처음 한 번): "실제 고객명·금액이 든 파일이면 내용을 지운 빈 양식으로 주시는 게 안전합니다."
3. **진단**
   ```
   python SKILL_DIR/scripts/inspect_template.py <파일> --name <회사-문서종류> --out profiles
   ```
   `profiles/<이름>/diagnosis.md` 를 읽고 등급을 사용자에게 설명한다.

   | 등급 | 뜻 | 할 일 |
   |---|---|---|
   | A | 바로 채우기 가능 | 4단계로 |
   | B | 이름·글꼴만 정리 | `normalize.py --auto` (PPTX) / `--all` (DOCX) |
   | C | 텍스트박스로 그린 양식 | 반자동 승격 (아래) + 사용자 확인 |

4. **표준화** (B·C) — 결과를 같은 폴더의 template 으로 저장
   - PPTX: `python SKILL_DIR/scripts/normalize.py profiles/<이름>/template.pptx --auto --out profiles/<이름>/template.pptx`
   - PPTX C등급: 대표 슬라이드마다
     `normalize.py ... --promote <슬라이드번호> --as-layout <용도이름> --slot-names 제목,핵심메시지,본문`
     슬롯 이름은 위→아래, 왼→오 순서. 사용자에게 "3번 슬라이드의 상자 3개를 '제목/핵심메시지/본문' 칸으로 만들었습니다"처럼 보고하고 맞는지 묻는다.
   - 레이아웃 이름이 영어 기본값이면 용도 이름으로: `--rename-layouts "Title Slide=표지,Title and Content=본문"`
   - DOCX: `normalize.py profiles/<이름>/template.docx --all --out profiles/<이름>/template.docx`
5. **재진단** → 등급 A/B 확인. C가 남으면 무엇이 남았는지 정직하게 알린다 (자동으로 다 된다고 하지 않는다).
6. **profile.yaml 다듬기** — 초안의 `[ ]` 를 채운다
   - `purpose`, `audience`: 사용자에게 한 번에 묻거나 양식 내용에서 추론(추론이면 "가정"이라고 표시)
   - PPTX `layouts.*.use`: 레이아웃마다 용도 한 줄
   - `storyline` / `outline`: 이 양식의 대표 흐름
   - `writing`: 회사 문체 (개조식 기호 순서, 금지어). 사용자가 잘 쓴 예시 문서를 주면 `extract.py` 로 뽑아 문체·분량을 읽고 규칙으로 옮긴다
   - 스키마: `references/profile-schema.md`
7. **시험 채우기**: 짧은 예시 내용으로 fill → 검증 PASS + 미리보기 확인 → 사용자에게 보여 주고 등록 완료.

## 자주 생기는 경우

- 슬라이드 마스터가 여러 개: 첫 마스터 기준. 다른 마스터가 필요하면 사용자에게 알림 (확인 필요 기능).
- 로고가 그림으로 마스터에 있음: 그대로 유지된다 (건드리지 않음).
- 워드 양식이 표로 칸을 나눈 서식(신청서형): 표 안 칸 채우기는 현재 미지원 → `revise.py` 의 find/replace 로 칸별 교체를 안내.
