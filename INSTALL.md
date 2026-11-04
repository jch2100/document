# 설치 안내

## 준비물

| 항목 | 필요 | 비고 |
|---|---|---|
| Claude 유료 플랜 | claude.ai 경로 | Pro · Max · Team · Enterprise |
| 코드 실행 켜기 | claude.ai 경로 | 설정 → 기능 → 코드 실행 및 파일 생성 |
| Python 3.9+ | Claude Code 경로 | `python --version` |
| LibreOffice | 선택 | 미리보기 썸네일용. 없어도 생성·검증은 됨 |

## A. claude.ai / 데스크톱 앱

1. Releases에서 `yangsik.zip` 받기 (압축 풀지 않음)
   - 직접 만들려면: 저장소 루트에서 `python tests/make_dist.py` → `dist/yangsik.zip`
2. 설정 → 기능 → **스킬** → 스킬 업로드 → zip 선택
3. 목록에서 `yangsik` 이 켜져 있는지 확인
4. 새 대화: *"양식대로 스킬 목록 보여줘"* → 기본 양식 10종이 나오면 성공

> 메뉴 위치는 앱 업데이트로 바뀔 수 있습니다. "Skills" 또는 "스킬"을 설정에서 찾으세요.
> 조직 계정(Team·Enterprise)은 관리자가 스킬 기능을 켜야 할 수 있습니다.
> 스킬은 **사람마다** 올립니다. 팀 전체 배포 기능은 플랜에 따라 다릅니다 (확인 필요).

## B. Claude Code

```bash
git clone https://github.com/jch2100/document.git
```
개인용(모든 프로젝트):
```bash
mkdir -p ~/.claude/skills && cp -r document/yangsik ~/.claude/skills/yangsik
```
프로젝트용(이 폴더에서만):
```bash
mkdir -p .claude/skills && cp -r document/yangsik .claude/skills/yangsik
```
패키지:
```bash
pip install -r ~/.claude/skills/yangsik/requirements.txt
```
확인:
```bash
python ~/.claude/skills/yangsik/scripts/list_templates.py
```

## Windows 메모

- PowerShell에서 한글이 깨지면: `chcp 65001`
- `python` 대신 `py` 일 수 있음
- LibreOffice 설치 경로 `C:\Program Files\LibreOffice\program\soffice.exe` 는 자동 인식

## 업데이트

- claude.ai: 새 zip을 같은 이름으로 다시 업로드 (등록한 `profiles/` 는 zip에 넣어 함께 올림)
- Claude Code: `git pull` 후 다시 복사
