<!-- 이 블록의 원본은 engine/templates/SKILL_PREAMBLE.md 다. gen_preamble.py 가 모든 작업 스킬에 같은 내용을 채운다. 여기서 직접 고치지 않는다 -->
## PREAMBLE — 본 절차보다 먼저 실행한다

```bash
python "$(git rev-parse --show-toplevel)/report-factory/engine/tools/preamble.py" --skill {{SKILL}} <조사 폴더 인자가 있으면 그대로>
```

한 번의 호출이 새 판 확인·활성 조사 지정·조사 메모리 위치·실행 기록을 처리하고 `KEY: 값` 줄로 답한다. 줄마다 이렇게 한다.

- **`UPDATE: NONE`** 또는 **`DISABLED`** — 아무 말 없이 본 절차로 간다.
- **`UPDATE: UPGRADE_AVAILABLE <내판> <새판>`** — 먼저 `update_check.py whats-new <새판>` 으로 바뀐 점을 읽는다.
  `update_check.py config get auto_upgrade` 가 `true` 면 묻지 않고 `/업그레이드` 를 실행한 뒤 본 절차로 간다.
  아니면 AskUserQuestion 으로 묻는다 — 「report-factory **<새판>** 이 나왔습니다 (지금 <내판>). <바뀐 점 한 줄>. 올릴까요?」
  1. **지금 올린다** → `/업그레이드` → 본 절차
  2. **앞으로는 자동으로** → `update_check.py config set auto_upgrade true` → `/업그레이드` → 본 절차
  3. **지금 말고 나중에** → `update_check.py snooze <새판>`. 출력의 next 값대로 「다음 알림은 24시간 뒤입니다」라고 알린다(두 번째는 48시간, 세 번째부터 1주일) → 본 절차
  4. **다시 묻지 않는다** → `update_check.py config set update_check false`. 「다시 켜려면 `config set update_check true`」 → 본 절차
- **`UPDATE: JUST_UPGRADED <이전판> <내판>`** — `whats-new <내판>` 을 읽어 「<내판> 으로 올라왔습니다: <바뀐 점>」 한 줄만 말하고 본 절차로 간다.
- **`UPDATE: CHECK_FAILED <이유>`** — 「판을 확인하지 못했습니다(<이유>). 그대로 진행합니다」 한 줄만 말하고 본 절차로 간다. 최신으로 읽지 않는다.
- **`MEMORY: <경로>`** — 그 `MEMORY.md` 를 읽는다. 「결정된 것」은 다시 묻지 않는다. 스킬이 끝날 때 세 절(결정된 것 · 지시의 요지 · 다음 세션이 알아야 할 것)을 갱신한다. `MEMORY_NEW: yes` 면 읽을 것이 없으니 끝날 때 채우기만 한다.
- **`MEMORY: NONE`** — 조사 폴더 인자가 없었다. 조사가 정해지는 즉시 조사 폴더를 인자로 이 명령을 다시 실행한다.

사용자 입력과 손 편집은 훅이 같은 `_memory/` 폴더에 자동으로 남긴다(`knowledge/02_프로젝트-메모리.md`). 이 블록은 스킬 본문의 어떤 규칙보다 먼저다.
