---
name: 업그레이드
description: research-factory 를 최신판으로 올린다 — git pull·의존성 갱신·스모크 테스트·변경 요약 보고. 비개발자 전제라서 충돌도 Claude 가 정리한다. 사용자가 "업그레이드", "최신판 받아줘", "새 규칙 받아졌어?"라고 하거나 /업그레이드 를 호출할 때 사용.
---

# 업그레이드 — 최신판 내려받기

작업 폴더(리포 루트)에서 실행한다. 사용자는 git 을 모른다는 전제로, 모든 git 조작과 충돌 정리를 대신한다.

## 절차

1. **현재 상태 파악** — `git status --porcelain` 과 `git log --oneline -1`.
   `OLD=$(cat VERSION)` 으로 지금 판을 먼저 적어 둔다.
   추적 파일에 로컬 수정이 있으면 무엇이 바뀌어 있는지 한국어로 요약해 둔다(아직 묻지 않는다).

2. **받기**
   - 수정이 없으면: `git pull --ff-only`
   - 수정이 있으면: `git stash` → `git pull --ff-only` → `git stash pop`
   - 받은 뒤 `python research-factory/engine/tools/update_check.py mark --from "$OLD"` 로 표식을
     남긴다. 다음에 스킬을 부르면 「<새판> 으로 올라왔습니다」 한 줄이 뜬다.
   - 이어서 `python research-factory/engine/tools/migrate_folder.py` — 0.5.0 에서 안쪽 폴더 이름이
     `report-factory/` → `research-factory/` 로 바뀌었다. git 은 추적 파일만 옮기므로, git 밖에 있던 조사 데이터·코퍼스·
     인수인계 문서를 이 도구가 새 폴더로 옮기고 옛 자리에 연결(정션)을 건다. 멱등이라 매번 불러도 된다.
     「열린 파일 때문에 못 옮김」이면 엑셀 등을 닫게 한 뒤 다시 부른다. 「옛 쪽을 남김」이면 두 파일을 비교해 보여 주고
     사용자가 고르게 한다. 옛 폴더 `report-factory/` 에 남는 대리 파일 6개와 연결은 지우지 않는다(업그레이드 전에 열린
     세션이 옛 경로로 부른다).
   - `stash pop` 충돌 시: 충돌 파일을 열어 판단한다. **로컬 수정과 원격 갱신이 같은 내용이면
     원격판을 남기고 stash 를 버린다.** 다른 내용이면 두 쪽을 다 살리는 병합을 시도하고,
     판단이 서지 않는 부분만 사용자에게 두 안을 보여 주고 고르게 한다.
     **`report-factory/engine/tools/` 의 대리 파일에서 충돌하면 병합하지 않는다**(0.5.0 이전 로컬 수정이 옛 경로에 남은 것).
     `git checkout HEAD -- report-factory/engine/tools/<파일>` 로 대리 파일을 되돌리고, 로컬 수정은
     `research-factory/engine/tools/<같은 파일>` 에 다시 적용한다.

3. **의존성·검증** — `pip install -q -r requirements.txt` 후
   `PYTHONUTF8=1 python research-factory/engine/tools/verify.py --help` 로 기동 확인.

4. **보고** — `NEW=$(cat VERSION)` 을 읽고, 당겨온 커밋들(`git log --oneline <이전>..HEAD`)과 CHANGELOG 최신 항목을 읽어
   **무엇이 달라졌는지 연구원 언어로** 요약한다. 규칙이 바뀌었으면 규칙 번호와 한 줄 요지,
   도구가 바뀌었으면 체감 변화만. 커밋 해시 나열은 하지 않는다.
   이어서 `python research-factory/engine/tools/release_notes.py url "$NEW" --from "$OLD"` 로 링크를 받아
   보고 끝에 **「업데이트 내역: [research-factory <NEW>](<RELEASE_NOTES: 링크>)」** 를 마크다운 링크로 붙인다.
   누르면 달라진 점 · 새로 써볼 수 있는 것 · 기능 설명이 나온다고 한 줄 덧붙인다.
   `RELEASE_NOTES_SKIPPED:` 가 있으면(여러 판을 한 번에 올린 경우) 「건너뛴 판의 내역」으로 그 링크들도 붙인다.
   링크가 `NONE` 이면 `release_notes.py show "$NEW"` 본문의 「새로 써볼 수 있는 것」 절을 대신 보여 준다.
   `OLD` 가 0.5.0 보다 낮았으면 끝에 「Claude 세션을 한 번 닫고 새로 열어 주세요. 훅과 스킬이 새 폴더 경로로 붙습니다」를 덧붙인다.

## 주의

- `git reset --hard`·`checkout --` 같은 파괴적 조작은 쓰지 않는다(위의 대리 파일 되돌리기만 예외). `git stash -a`·`git clean -x` 도
  쓰지 않는다 — 옛 폴더의 연결(정션)을 따라가 조사 데이터를 건드린다. 사용자의 로컬 수정은
  그 사람이 다듬어 온 개선분일 수 있다 — 지울지 말지는 항상 사용자가 정한다.
- `research-factory/projects/` 는 git 밖이다. 업그레이드가 조사 데이터의 내용을 바꾸는 일은 없다(0.5.0 폴더 이름 변경 때
  자리만 옮긴다). 사용자가 불안해하면 그렇게 안심시켜 준다.
