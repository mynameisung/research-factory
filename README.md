# report-factory &nbsp;<sub><sup>powered by seo-yeon</sup></sub>

> 통계표가 나온 순간부터 납품 PPT 까지, 리서치 보고서 제작을 Claude Code 팀으로 바꾼다.

통계표·설문지·조사개요를 넣으면 제작가이드(.xlsm)와 보고서 PPT 가 나온다. 연구원은 git 을 몰라도 된다.

## 1. 업무 흐름

```
설문 → 웹 구축 → 웹테스트 *1 → 실사 → 통계표 → 제작가이드 *2 → 제작본 PPT → 보고서 마감 *3 → 납품 → 회고 *4
```

`*` 가 붙은 네 곳을 스킬이 한다. 나머지는 실사팀·DP·제작팀 몫이다.

- **\*1 `/웹테스트`** — 실사 전에 웹 화면이 설문지대로 됐는지 브라우저가 대신 걸어 검수하고, 수정요청서(.xlsx)를 낸다.
- **\*2 `/제작가이드`** — 통계표·설문지로 제작팀에 보낼 지시서(.xlsm)를 만든다. 멘트·요약·결론 문장도 여기서 쓴다.
- **\*3 `/보고서작성`** — 제작팀 회신본에 값을 채우고 전수 검수·렌더 QA 로 마감한다.
- **\*4 `/회고`** — 이번 조사에서 배운 것을 규칙으로 올린다.

새 판 받기는 `/업그레이드`, 내가 고친 것을 나누는 것은 `/개선공유`.

## 2. 어떻게 돌아가는가

네 층이다. 위에서 아래로 시키고, 아래에서 위로 검사 결과가 올라온다.

```
사용자  ──/제작가이드 <조사 폴더>──▶ ┌─ 스킬 (오케스트레이터) ────────────────────────────┐
                                    │ 판 확인 → 조사 메모리 읽기 → 오리엔테이션 → 파트별 배정   │
                                    └───────┬──────────────────────────┬───────────────┘
                                            │ 배정 (명시적 ID 목록)          │ 수용은 기계 게이트가
                                    ┌───────▼──────────────┐   ┌───────▼──────────────┐
                                    │ 에이전트 12명           │   │ 검수 게이트            │
                                    │ 목차·Part Ⅰ·멘트·요약·   │   │ verify (V·D·K 검사)   │
                                    │ 결론·심사·반증·렌더 QA   │   │ audit · lint · 렌더    │
                                    └───────┬──────────────┘   └───────▲──────────────┘
                                            │ 팩(자기 규칙 슬라이스)만 읽는다   │
                                    ┌───────▼──────────────────────────┴───────────────┐
                                    │ 엔진 도구 40여 개 (파서 · 빌더 · 검수기 · 덱 주입기)      │
                                    └───────┬─────────────────────────────────────────┘
                                            │ 읽고 쓴다
                                    ┌───────▼─────────────────────────────────────────┐
                                    │ 정본 파일  tables.json · questionnaire.json · 목차.md     │
                                    │           slide_facts.json (숫자·문장의 정본) · 제작가이드  │
                                    └─────────────────────────────────────────────────┘
```

- **스킬**은 절차와 배정만 안다. 규칙 대장을 통째로 읽지 않고, 담당 에이전트에게 파트를 맡긴다.
- **에이전트**는 자기 파트의 규칙 슬라이스(팩)만 읽고, 팩트를 하우스 문형으로 옮길 뿐 숫자를 만들지 않는다.
- **엔진 도구**가 숫자를 만든다. 통계표를 파싱해 `tables.json`, 설문지를 파싱해 `questionnaire.json`, 목차와 멘트를 합쳐
  `slide_facts.json`(정본)을 만들고, 거기서 제작가이드와 PPT 를 찍어 낸다. 사람이 손으로 숫자를 옮기는 자리는 없다.
- **검수 게이트**가 수용을 정한다. 제작가이드는 `verify.py`(V-검사, 치명 0·경고 0), 덱은 `verify --mode deck`(D-검사)와
  `audit_deck_values`(표·차트를 장마다 정본과 대조), 렌더는 PNG 육안 검수 에이전트가 본다.
- **규칙**은 사람용 대장(`RULES.md`)에서 기계용 부분집합(`rules.yaml`)과 담당별 팩(`packs/*.md`)으로 흘러간다.
  `/회고` 가 조사의 흔적(사용자 지시 원문·손 편집 이력·반송 대장·검수 findings)에서 규칙 후보를 뽑고, 사람이 승인한 것만
  대장에 오른다.

## 3. 스킬 6개

| 스킬 | 하는 일 |
|---|---|
| `/제작가이드` | 통계표·설문지·조사개요 → 목차 설계 → 본문 블록(차트타입·베이스·QA·멘트) → 요약·결론 블록 → 검증 → 제작팀 송부본(.xlsm) |
| `/보고서작성` | 제작팀 회신본 마감(경로 A) · 예시 덱에서 직접 조립(경로 B) · 통계표 갱신 뒤 값만 재반영(경로 C) → 수치 전수 검수 → 렌더 QA → 반증 감사 |
| `/웹테스트` | 설문지 정본 → 시나리오 행렬(기본·이탈·숨김) → Playwright 가 한 문항씩 순회·덤프 → 정본 대조(문항 누락·분기·이탈·파이핑·보기·문구) → 수정요청서(.xlsx). 로그인은 사람이 한다 |
| `/회고` | 조사가 남긴 증거에서 규칙 후보를 채굴하고, 승인분만 `RULES.md`·`rules.yaml`·knowledge 에 등재 |
| `/업그레이드` | 최신판 받기(stash → pull → pop) · 의존성 · 스모크 · 변경 요약. 충돌은 Claude 가 정리한다 |
| `/개선공유` | 내가 다듬은 규칙·지식·도구를 임시 워크트리에서 PR 로 올린다. 작업 폴더는 건드리지 않는다 |

네 작업 스킬은 시작할 때 두 가지를 먼저 한다. **판 확인**(원격 `VERSION` 만 읽어 새 판이면 「올릴까요?」)과
**조사 메모리 읽기**(`projects/<조사>/_memory/MEMORY.md`).

## 4. 에이전트 12명

| 팀 | 에이전트 |
|---|---|
| 가이드 설계 | guide-toc(목차·배너) · guide-front(Part Ⅰ·Cover) · guide-body-ment(본문 멘트) |
| 요약·결론 | guide-summary(총괄·결론) · guide-summary-slide(꼭지별 문장) · guide-client-review(의뢰인 자리 심사) |
| 심사·검증 | guide-judge(설계 심사) · guide-summary-reviewer(품질) · guide-adversary(반증 감사) |
| PPT·회고 | deck-adversary(덱 감사) · deck-render-qa(렌더 검수) · retro-miner(회고 채굴) |

## 5. 정보 구조 (IA) — 무엇이 어디에 있나

```
embrain/                              ← 리포 루트 = Claude Code 작업 폴더
├─ README.md · CHANGELOG.md · VERSION   판은 세 자리. 병합할 때 VERSION 과 CHANGELOG 를 함께 올린다
├─ setup.ps1 · setup · requirements.txt 설치 (의존성 · PYTHONUTF8 · 스모크)
├─ .claude/
│   ├─ skills/                        스킬 6개 (SKILL.md 하나씩)
│   ├─ agents/                        에이전트 12명 정의
│   └─ settings.json                  훅 2개 — 사용자 입력 포착 · 손 편집 포착 (조사 폴더에 기록)
└─ report-factory/
    ├─ engine/                        ── 재사용 자산. 조사와 무관. git 이 추적 ──
    │   ├─ tools/                     파서(parse_table_* · parse_questionnaire) · 정본 빌더(build_facts) ·
    │   │                             가이드 빌더(build_guide) · 덱 주입(fill_deck · refill_deck · deckbuild · deckops) ·
    │   │                             검수(verify · verify_deck · audit_deck_values · deck_lint) · 팩 생성(make_packs) ·
    │   │                             판 확인(update_check) · 훅(capture_prompt · capture_edit) · 통계표 사전 검사(table_lint) ·
    │   │                             웹테스트(webtest_scenarios · webtest_walk · webtest_compare · webtest_selftest + webtest_fixture/ 모의 설문)
    │   ├─ knowledge/                 업무 지식. 00 업무개요 · 01 통계표 읽는 법 · 02 프로젝트 메모리 · 40 웹테스트(업무 C)
    │   │   ├─ guide/                 【업무 A】 10 목차 · 11 블록포맷 · 12 차트타입 결정표 · 13 멘트 · 14 요약·결론(규칙) · 14A 코퍼스 관찰 ·
    │   │   │                         14B 헤드라인 문형 사전 · 15 지시어 사전 · 15 한국어 문장지침 · 16 구조 대조 · 17 팀구조 · 18 오케스트레이션 · 19 요약 그래프
    │   │   └─ report/                【업무 B】 20 반복장표 · 21 멘트 보강 · 22 검수 체크리스트 · 23 팀구조 · 24 그래프 하우스표준 · 25 차트 카탈로그
    │   ├─ learning/                  규칙과 학습 루프
    │   │   ├─ RULES.md               규칙 대장 (사람용 · 카드마다 분류·근거·신뢰도·적용·이력)
    │   │   ├─ rules.yaml             기계 판독 부분집합 → verify 가 검사 항목으로 쓴다
    │   │   ├─ packs/                 담당별 슬라이스 6종 + _색인.md (make_packs 가 RULES 의 「요지」 줄만 뽑아 생성 — 손대지 않는다)
    │   │   ├─ OPEN_QUESTIONS.md      아직 규칙이 못 된 것
    │   │   ├─ retro/ · eval/ · golden_loop/   회고 감사 · 골든셋 역검증 기록 (산출 JSON 은 git 밖)
    │   │   └─ (observations.jsonl)   옛 전역 편집 로그. git 밖. 지금은 조사별 _memory/ 에 쌓인다
    │   ├─ templates/                 TEMPLATE_제작가이드.xlsm · TEMPLATE_프로젝트/ (새 조사 폴더의 뼈대)
    │   ├─ macros/통계표전처리/        DP 통계표 전처리 엑셀 추가 기능 (.xlam 과 소스)
    │   └─ refs/                      ── git 밖. 코퍼스 원본 5벌 · 하우스 그래프 가이드 · 분석 문서 (사내 드라이브로 전달) ──
    └─ projects/                      ── git 밖. 고객 데이터. 조사 하나 = 폴더 하나 ──
        ├─ .active                    훅이 기록할 조사 (스킬이 적는다)
        └─ <연도_클라이언트_프로젝트>/
            ├─ README.md · ORIENTATION.md · 00_공통/ (조사개요 · RFP · 회의록 · To Know List)
            └─ <NN_조사명>/
                ├─ input/             설문지 · 통계표 · 원자료
                ├─ 1_제작가이드/       work/ (정본 JSON · 목차.md · 멘트_확정.json) · out/ (.xlsm · 검수리포트 · .baseline/)
                ├─ 2_보고서작성/       in/ (제작본) · work/ (빌드 스크립트 · 렌더 · 검수) · out/ (최종 .pptx · 검수리포트)
                ├─ 3_웹테스트/         in/ (설문지) · work/ (시나리오 · 덤프 · 캡처 · 로그인 세션) · out/ (findings · 수정요청서 .xlsx)
                ├─ _반송대장/          가이드 팀 ↔ PPT 팀 반송 기록
                └─ _memory/           MEMORY.md (정리본) · 사용자입력.jsonl · 편집기록.jsonl (훅이 쌓는 원문)
```

**git 이 추적하는 것은 `engine/`·`.claude/`·루트 문서뿐이다.** 고객 조사(`projects/`)와 코퍼스 원본(`refs/`)은
어떤 경우에도 리포에 올라가지 않는다(.gitignore 강제). 개인 로그와 재생성 가능한 산출도 밖에 둔다.

## 6. 정본과 검수

| 정본 | 만드는 도구 | 검사 |
|---|---|---|
| `tables.json` (통계표) | `parse_table_xlsx` / `parse_table_biff` + `apply_table_fixes`(라벨 정정 선언) | `table_lint` (입고 직후 라벨·중복 검사) |
| `questionnaire.json` | `parse_questionnaire` | — |
| `목차.md` · `slide_facts.json` | guide-toc → `build_facts` (멘트·베이스·각주·파이핑 치환 포함) | `verify --mode toc` · `--mode ments` · `--mode summary` |
| 제작가이드 `.xlsm` | `build_guide` (+ 송부본) | `verify` (V-검사, 치명 0·경고 0) · guide-adversary |
| 보고서 `.pptx` | `fill_deck` / `refill_deck` / `deckbuild` (`deckops.apply_chart` — 차트 데이터는 PPT 안에) | `verify --mode deck` (D-검사) · `audit_deck_values` · `deck_lint` · deck-render-qa · deck-adversary |

숫자의 정본은 `tables.json`, 문장의 정본은 `slide_facts.json` 이다. PPT 팀은 문장을 짓지 않고, 가이드 팀은 덱을 만지지 않는다.
어긋나면 `_반송대장/` 에 적어 상대 팀이 정본을 고친다.

## 7. 규칙 체계

```
RULES.md (사람용 대장, 카드 R-nnn)
  ├─▶ rules.yaml            기계가 검사할 수 있는 것만 → verify 의 V-검사
  ├─▶ packs/<담당>.md        make_packs 가 분류 태그로 잘라 담당 에이전트에게 배달
  └─▶ knowledge/            카드가 설명하는 「왜」와 절차
```

카드에는 분류·규칙·근거·신뢰도(`가설` → `검증중` → `확정`, 사용자 지시는 `지시`)·적용·이력이 있다. `가설` 카드는
적용 전에 사용자에게 확인한다. 규칙을 고치면 `make_packs.py` 로 팩을 다시 만든다.

## 8. 조사 메모리 — 조사마다 따로 기억한다

연구원마다 맡는 조사가 다르므로 기억은 리포가 아니라 **조사 폴더**에 쌓인다. 훅 두 개가 자동으로 남긴다.

- `사용자입력.jsonl` — 이 폴더에서 입력한 지시 전부 (`capture_prompt.py`, UserPromptSubmit 훅)
- `편집기록.jsonl` — 사람이 산출물을 손으로 고친 이력 (`capture_edit.py`, PostToolUse 훅)
- `MEMORY.md` — 스킬이 시작할 때 읽고 끝날 때 갱신하는 정리본

어느 조사에 기록할지는 `projects/.active` 가 정하고, 스킬이 조사 폴더 인자를 받으면 적어 둔다. 자세한 것은
`knowledge/02_프로젝트-메모리.md`. `/회고` 는 이 두 원문을 ★★★ 관찰 소스로 읽는다.

## 9. 설치 — 30초

**요구 사항:** [Claude Code](https://docs.anthropic.com/en/docs/claude-code) · [Git](https://git-scm.com/) · Python 3.11+ · (개선 공유용) [GitHub CLI](https://cli.github.com/)

Claude Code 를 열고 아래 상자를 그대로 복사해 붙여넣으면 Claude 가 알아서 한다 (상자 오른쪽 위 복사 단추).

```text
report-factory 설치: git clone https://github.com/FineHourTeam/report-factory.git 로 원하는 위치에 받고, 그 폴더에서 powershell -File setup.ps1 (macOS/Linux 는 ./setup) 을 실행한 다음, 그 폴더를 Claude Code 작업 폴더로 열어 줘.
```

직접 하려면 터미널에서:

```bash
git clone https://github.com/FineHourTeam/report-factory.git
cd report-factory
powershell -File setup.ps1
```

macOS/Linux 는 마지막 줄 대신 `./setup`. 끝나면 그 폴더를 Claude Code 작업 폴더로 연다.

클론한 폴더가 곧 작업 폴더다. 스킬·에이전트·엔진이 상대 경로로 서로를 찾으므로 **폴더 안 구조는 바꾸지 않는다.**
설치되면 이 폴더에서 입력하는 지시는 조사별 `_memory/` 에 자동으로 저장된다.

## 10. 업그레이드와 협업

```
연구원 1·2·3                       관리자(리드)                 GitHub FineHourTeam/report-factory (비공개)
   │ 스킬 호출 → 판 확인 ◀────────────────────────────────────  main 의 VERSION
   │ 「올릴까요?」 예 → /업그레이드                                   │
   │ 조사 작업 (스킬이 알아서) · _memory 에 기억 쌓임                    │
   │ /회고 → 내 로컬 규칙 다듬어짐                                    │
   │ /개선공유 ─────────▶ PR 검토 · 공용만 병합 · VERSION 올림 ───▶  main 갱신
```

- **판 확인**: 스킬을 부를 때마다 `update_check.py` 가 `git fetch` 로 원격 `VERSION` 한 줄만 읽는다. 작업 폴더는 건드리지
  않는다. 새 판이면 「지금 올린다 / 항상 자동 / 나중에(24시간 → 48시간 → 1주일) / 다시 묻지 않기」를 묻는다.
  상태는 `~/.report-factory/` 에 있고, 설정은 그 안의 `config.json`(`auto_upgrade` · `update_check`)이다.
- **받기**: `/업그레이드`. 로컬 수정과 충돌하면 Claude 가 정리한다.
- **주기**: 관리자가 「PR 올리세요」 하면 각자 `/개선공유`. 변경분을 요약해 보여 주고 확인 후 PR 까지 대신한다.
- **고르기**: 관리자가 공용으로 쓸 PR 만 병합하고, 그때 `VERSION` 과 `CHANGELOG.md` 를 함께 올린다. CHANGELOG 항목이
  곧 연구원 화면에 뜨는 「바뀐 점」 문구다.

## 11. 새 조사 시작하기

1. `engine/templates/TEMPLATE_프로젝트/` 를 복사해 `projects/<연도>_<클라이언트>_<프로젝트>/` 로 이름을 바꾼다.
2. 입력물을 넣는다. 이름은 `<구분자>_<내용>_<일자>.<ext>` — `공통_`(프로젝트 전체) · `<조사약칭>_`(조사별, 예 `TA_`·`비상주_`).
3. `/제작가이드 projects/<프로젝트>/<조사>` 를 부른다. 오리엔테이션(조사 목적·의뢰인이 알고 싶은 것·요약·결론 필요 여부)부터 시작한다.
4. `engine/` 은 건드리지 않는다. 새로 배운 규칙은 `/회고` 로 올리고, 도구 개선은 `/개선공유` 로 보낸다.

## 12. 여기 없는 것

- `report-factory/projects/` — 고객 데이터. 어떤 경우에도 리포에 올리지 않는다.
- `report-factory/engine/refs/` — 코퍼스 원본 5벌과 하우스 그래프 가이드(157MB). 규칙의 근거이며 사내 드라이브로 전달한다.
  그 내용에서 「언제 어떤 그림을 쓰나」만 `knowledge/report/25_차트-카탈로그.md` 에 적어 두었다.
- API 키 — 전부 환경변수(`CLOVA_STUDIO_API_KEY` 등). 파일로 두지 않는다.

이 리포는 **비공개**다. 규칙 문서에 실제 납품 사례의 실측 수치가 들어 있으므로 공개 전환하지 않는다.
