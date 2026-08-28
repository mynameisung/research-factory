---
name: 업그레이드
description: report-factory 를 최신판으로 올린다 — git pull·의존성 갱신·스모크 테스트·변경 요약 보고. 비개발자 전제라서 충돌도 Claude 가 정리한다. 사용자가 "업그레이드", "최신판 받아줘", "새 규칙 받아졌어?"라고 하거나 /업그레이드 를 호출할 때 사용.
---

# 업그레이드 — 최신판 내려받기

작업 폴더(리포 루트)에서 실행한다. 사용자는 git 을 모른다는 전제로, 모든 git 조작과 충돌 정리를 대신한다.

## 절차

1. **현재 상태 파악** — `git status --porcelain` 과 `git log --oneline -1`.
   추적 파일에 로컬 수정이 있으면 무엇이 바뀌어 있는지 한국어로 요약해 둔다(아직 묻지 않는다).

2. **받기**
   - 수정이 없으면: `git pull --ff-only`
   - 수정이 있으면: `git stash` → `git pull --ff-only` → `git stash pop`
   - `stash pop` 충돌 시: 충돌 파일을 열어 판단한다. **로컬 수정과 원격 갱신이 같은 내용이면
     원격판을 남기고 stash 를 버린다.** 다른 내용이면 두 쪽을 다 살리는 병합을 시도하고,
     판단이 서지 않는 부분만 사용자에게 두 안을 보여 주고 고르게 한다.

3. **의존성·검증** — `pip install -q -r requirements.txt` 후
   `PYTHONUTF8=1 python report-factory/engine/tools/verify.py --help` 로 기동 확인.

4. **보고** — 당겨온 커밋들(`git log --oneline <이전>..HEAD`)과 CHANGELOG 최신 항목을 읽어
   **무엇이 달라졌는지 연구원 언어로** 요약한다. 규칙이 바뀌었으면 규칙 번호와 한 줄 요지,
   도구가 바뀌었으면 체감 변화만. 커밋 해시 나열은 하지 않는다.

## 주의

- `git reset --hard`·`checkout --` 같은 파괴적 조작은 쓰지 않는다. 사용자의 로컬 수정은
  그 사람이 다듬어 온 개선분일 수 있다 — 지울지 말지는 항상 사용자가 정한다.
- `report-factory/projects/` 는 git 밖이다. 업그레이드가 조사 데이터를 건드릴 일은 없고,
  사용자가 불안해하면 그렇게 안심시켜 준다.
