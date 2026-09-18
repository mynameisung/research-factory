<!-- 이 블록의 원본은 engine/templates/SKILL_PREAMBLE.md 다. gen_preamble.py 가 모든 작업 스킬에 같은 내용을 채운다. 여기서 직접 고치지 않는다 -->
## PREAMBLE — 본 절차보다 먼저 실행한다

```bash
python "$(git rev-parse --show-toplevel)/research-factory/engine/tools/preamble.py" --skill {{SKILL}} <조사 폴더 인자가 있으면 그대로>
```

한 번의 호출이 새 판 확인·활성 조사 지정·조사 메모리 위치·실행 기록을 처리하고 `KEY: 값` 줄로 답한다. 줄마다 이렇게 한다.

- **`UPDATE: NONE`** 또는 **`DISABLED`** — 아무 말 없이 본 절차로 간다.
- **`UPDATE: UPGRADE_AVAILABLE <내판> <새판>`** — 먼저 `update_check.py whats-new <새판>` 으로 바뀐 점을 읽는다.
  `update_check.py config get auto_upgrade` 가 `true` 면 묻지 않고 `/업그레이드` 를 실행한 뒤 본 절차로 간다.
  아니면 AskUserQuestion 으로 묻는다 — 「research-factory **<새판>** 이 나왔습니다 (지금 <내판>). <바뀐 점 한 줄>. 올릴까요?」
  1. **지금 올린다** → `/업그레이드` → 본 절차
  2. **앞으로는 자동으로** → `update_check.py config set auto_upgrade true` → `/업그레이드` → 본 절차
  3. **지금 말고 나중에** → `update_check.py snooze <새판>`. 출력의 next 값대로 「다음 알림은 24시간 뒤입니다」라고 알린다(두 번째는 48시간, 세 번째부터 1주일) → 본 절차
  4. **다시 묻지 않는다** → `update_check.py config set update_check false`. 「다시 켜려면 `config set update_check true`」 → 본 절차
- **`UPDATE: JUST_UPGRADED <이전판> <내판>`** — `whats-new <내판>` 을 읽어 「<내판> 으로 올라왔습니다: <바뀐 점>」 한 줄만 말하고 본 절차로 간다.
- **`MIGRATE: …`** (있을 때만) — 0.5.0 폴더 이름 변경(`report-factory/` → `research-factory/`)의 뒷정리 결과다. 「옮김」뿐이면 「조사 폴더를 새 폴더로 옮겼습니다(옛 경로도 그대로 쓸 수 있습니다)」 한 줄만 말하고 본 절차로 간다. 「열린 파일 때문에 못 옮김」이나 「아직 옛 폴더에 있다」면 엑셀 등 열린 파일을 닫게 한 뒤 `python "$(git rev-parse --show-toplevel)/research-factory/engine/tools/migrate_folder.py"` 를 다시 부르고, 옮겨질 때까지 본 절차로 가지 않는다. 「옛 쪽을 남김」이면 두 파일을 비교해 보여 주고 사용자가 고르게 한다. 「권한 문제」·「대리 파일에 로컬 수정」은 줄에 적힌 대로 처리한다.
- **`UPDATE: CHECK_FAILED <이유>`** — 「판을 확인하지 못했습니다(<이유>). 그대로 진행합니다」 한 줄만 말하고 본 절차로 간다. 최신으로 읽지 않는다.
- **`MEMORY: <경로>`** — 그 `MEMORY.md` 를 읽는다. 「결정된 것」은 다시 묻지 않는다. 스킬이 끝날 때 세 절(결정된 것 · 지시의 요지 · 다음 세션이 알아야 할 것)을 갱신한다. `MEMORY_NEW: yes` 면 읽을 것이 없으니 끝날 때 채우기만 한다.
- **`MEMORY: NONE`** — 조사 폴더 인자가 없었다. 조사가 정해지는 즉시 조사 폴더를 인자로 이 명령을 다시 실행한다.
- **`PATHS: {…}`** — 역할 → 경로 JSON. 설문지·통계표·정본 JSON·제작본·최종보고서 같은 파일과 단계 폴더를 **여기서 찾는다.** 폴더 이름을 추측하거나 `input/` 을 뒤지지 않는다. 역할 하나만 필요하면 `paths.py <조사> --role <역할>`. `LAYOUT: v1` 이면 옛 배치이고 같은 역할 이름으로 풀려 있으니 그대로 쓴다(`paths.py migrate` 는 사람이 정한다). `LAYOUT: none` 이면 `paths.py new <프로젝트>/<조사>` 로 뼈대를 만든다. `files_all` 에 설문지나 통계표가 여러 벌 있으면 어느 것을 쓸지 사용자에게 확인한다.
- **`TRACKING: …`** 이 있으면 트래킹 조사다. `PREV_PATHS:` 의 이전 회차 통계표·최종보고서·제작가이드를 전차 비교의 정본으로 쓴다(전년 수치를 기억이나 추정으로 쓰지 않는다). 이전 회차가 `NONE` 이면 첫 회차이므로 비교 없이 간다.

사용자 입력과 손 편집은 훅이 같은 `_memory/` 폴더에 자동으로 남긴다(`knowledge/02_프로젝트-메모리.md`). 이 블록은 스킬 본문의 어떤 규칙보다 먼저다.
