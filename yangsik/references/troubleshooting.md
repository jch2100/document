# 문제 해결

| 증상 | 원인 | 해결 |
|---|---|---|
| `ModuleNotFoundError: pptx` | 패키지 없음 | `pip install -r requirements.txt` |
| 레이아웃 '○○' 없음 | 슬롯/레이아웃 이름 불일치 | 오류에 나온 '사용 가능' 목록의 이름으로 |
| 미리보기가 안 생김 | LibreOffice 없음 | 정상. 검증(check)은 그대로 동작. 설치: libreoffice.org |
| 미리보기 줄바꿈이 PowerPoint와 다름 | LibreOffice가 맑은 고딕 대신 다른 글꼴 사용 | PowerPoint에서 최종 확인. 배치 확인용으로만 사용 |
| Word에서 □ 대신 네모 깨짐 | 기호 글꼴 없음 | 테마 한글 글꼴을 맑은 고딕으로 (normalize --korean-font) |
| 목차가 비어 있음 | 목차는 필드 | Word에서 목차 오른쪽 클릭 → 필드 업데이트 (Google 문서는 미지원) |
| 회사 양식에 채웠더니 글이 엉뚱한 칸에 | 플레이스홀더 이름이 기본값 | `normalize.py --auto` 후 profile.yaml slots 확인 |
| 텍스트박스 양식이라 등급 C | 양식이 '칸' 없이 그려짐 | `--promote` 로 대표 슬라이드 승격, 또는 PowerPoint 슬라이드 마스터에서 개체 틀 추가 |
| Windows에서 한글 경로 오류 | 콘솔 인코딩 | PowerShell에서 `chcp 65001`, 또는 영문 폴더 사용 |
| claude.ai에서 등록한 양식이 다음 대화에 없음 | 대화마다 작업 폴더가 새로 만들어짐 | 등록 결과 폴더를 zip으로 받아 두고, 스킬 폴더 `profiles/` 에 넣어 스킬을 다시 업로드 |
| .ppt / .doc / .hwp | 구형·한글 형식 | PowerPoint/Word에서 pptx/docx 로 저장. HWP는 v2 예정 |
