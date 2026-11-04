---
name: yangsik
description: "양식대로 — 회사 양식(또는 기본 제공 표준 양식 10종) 그대로, 편집 가능한 Word(.docx)·PowerPoint(.pptx)를 만들고 한국어 문서 기준(개조식·한글 글꼴·넘침·AI 문체)으로 검증하는 스킬. 트리거: 보고서/기획서/품의서/결과보고서/제안서/회의록/경영진 보고/발표자료/교육자료를 워드나 PPT로 만들어 달라, 우리 회사 양식에 맞춰 달라, 양식 등록·표준화, 팀장 피드백 반영해서 고쳐 달라, 문서 검사·검증. Use for Korean business DOCX/PPTX generation from company templates, template onboarding/normalization, partial revision, and Korean document QA."
---

# 양식대로 (yangsik)

**원칙: 새로 그리지 않는다. 양식을 복제해 '칸'에 내용만 채우고, 매번 검증한다.**

이 파일이 있는 폴더를 `SKILL_DIR` 이라 부른다. 모든 스크립트는 `python SKILL_DIR/scripts/<이름>.py` 로 실행한다.
처음 한 번: `pip install -r SKILL_DIR/requirements.txt` (이미 있으면 생략. 실패해도 python-pptx·python-docx·pyyaml이 있으면 동작)

## 1. 요청 → 모드 판별

| 사용자가 말하면 | 모드 | 읽을 파일 |
|---|---|---|
| "우리 회사 양식 등록/표준화해줘", 양식 파일만 올림 | **onboard** | `workflows/1-onboard.md` |
| "~를 PPT로/발표자료로", "슬라이드 만들어줘" | **fill-pptx** | `workflows/2-fill-pptx.md` |
| "~를 워드로/보고서로/품의서로/회의록으로" | **fill-docx** | `workflows/3-fill-docx.md` |
| "3번 장만 고쳐줘", "팀장님 코멘트 반영", 기존 파일 + 수정 요청 | **revise** | `workflows/4-revise.md` |
| "이 문서 검사해줘", 남이 만든 파일 점검 | **check** | `workflows/5-check.md` |
| "우리 회사 색으로 양식 바꿔줘" | **theme** | `workflows/6-theme.md` |
| "우리 회사 전용 스킬로 묶어줘", 동료에게 줄 zip | **export-skill** | `workflows/7-export-skill.md` |

판별이 애매하면: 파일이 첨부됐고 내용 요청이 없으면 onboard, 내용만 있으면 fill.
**해당 workflow 파일을 반드시 읽고 따른다.** 내용을 쓰기 전에 `references/writing-guide.md` 도 읽는다.

## 2. 어떤 양식을 쓸까

1. 사용자가 회사 양식을 줬거나 `profiles/<회사>/` 가 있으면 → 그것 (onboard 먼저)
2. 없으면 기본 양식 10종에서 용도로 고른다 (`python SKILL_DIR/scripts/list_templates.py`)

| PPTX | 언제 | DOCX | 언제 |
|---|---|---|---|
| `pptx/01` 경영진 보고 | 결정 받기 | `docx/01` 1페이지 보고서 | 한 장 보고 |
| `pptx/02` 제안·영업 | 고객 설득 | `docx/02` 기획서·품의서 | 예산·실행 승인 |
| `pptx/03` 프로젝트 계획 | 킥오프·합의 | `docx/03` 결과보고서 | 끝난 일 보고 |
| `pptx/04` 성과 발표 | 실적 보고 | `docx/04` 제안서 | 서술형 제출 문서 |
| `pptx/05` 교육·워크숍 | 강의 | `docx/05` 회의록 | 결정·할 일 기록 |

고른 양식 폴더의 `profile.yaml` 을 읽는다: 레이아웃·슬롯 이름, 권장 흐름(storyline/outline), 문체 규칙이 들어 있다.

## 3. 공통 작업 순서 (fill)

1. profile.yaml 읽기 → 흐름에 맞춰 **내용 JSON** 작성 (형식: `references/content-format.md`)
2. `python SKILL_DIR/scripts/fill.py --template <양식> --content 내용.json --out <결과>` (검증 자동 실행)
3. 검증 결과가 FAIL/WARN 이면 **내용 JSON을 고쳐 다시 실행** (최대 3회). 남은 경고는 사용자에게 그대로 알린다.
4. 가능하면 `--preview` 로 썸네일을 만들어 직접 눈으로 확인 (LibreOffice 없으면 생략)
5. 결과 파일 + 검증 요약(판정·남은 경고·사람이 확인할 점)을 전달

## 4. 절대 하지 않는 것

- python-pptx 로 텍스트박스를 새로 그리거나 좌표를 직접 잡지 않는다 → 레이아웃 슬롯만 쓴다
- 글꼴·크기·색을 내용에 박지 않는다 → 양식이 정한다 (강조는 `**굵게**` 만)
- 개조식 기호(□ ○ -)나 제목 번호(1. 가.)를 글에 직접 쓰지 않는다 → 양식이 자동으로 붙인다
- 검증 FAIL 인 파일을 "완성"이라고 넘기지 않는다
- 사용자의 원본 양식 파일을 덮어쓰지 않는다 (항상 새 파일)
- 회사 양식·내용을 공개 저장소에 올리라고 권하지 않는다

## 5. 결과 보고 형식 (사용자에게)

```
✅ 결과.pptx (7장, 양식: 경영진 보고)
검증: PASS — 실패 0 · 경고 0
확인해 주세요: 5장 표의 금액(원 자료 기준), 2장 결론 문장이 의도와 맞는지
```
경고가 남았으면 "왜 남겼는지"를 한 줄로.
