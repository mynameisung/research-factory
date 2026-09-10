# report-factory

> 통계표가 나온 순간부터 제작팀 송부본까지, 리서치 보고서 제작을 Claude Code 팀으로 바꾼다.

**report-factory 는 엠브레인 리서치 보고서 제작 파이프라인이다.** 통계표(.xls)·설문지(.docx)·조사개요만 넣으면 — 목차 설계자, 본문 멘트 담당, 요약 설계자, 결론 담당, 적대 감사관, 렌더 QA 까지 **전문 에이전트 12명과 스킬 3개**가 제작가이드(.xlsm)와 보고서 PPT 를 만들고, 기계 검수(치명 0·경고 0 게이트)·칸 대 칸 대조·반증 감사로 스스로 검증한다.

규칙 186개 이상이 실제 납품 5개 사례(코퍼스)에서 채굴되어 `RULES.md` 에 근거와 함께 등재되어 있고, 조사가 끝날 때마다 `/회고` 가 새로 드러난 암묵지를 규칙으로 승격시킨다. **쓰면 쓸수록 똑똑해지는 공장이다.**

**누구를 위한 것인가**
- **연구원** — git 을 몰라도 된다. 클론 한 번이 설치의 전부이고, 업그레이드와 개선 공유도 스킬이 대신한다
- **보고서 제작 리드** — 두 팀 체제(가이드 정본 팀 / PPT 팀)와 반송 대장으로 세션 간 작업이 어긋나지 않는다

## Quick start

1. 설치 (아래 30초 설치)
2. 조사 폴더에 통계표·설문지를 넣고 `/제작가이드` — 목차 설계부터 제작팀 송부본까지
3. 제작본이 회신되면 `/보고서작성` — 반복 장표 채움·수치 전수 검수·렌더 QA
4. 납품 후 `/회고` — 이번에 배운 것을 규칙으로 등재
5. 가끔 관리자가 「PR 올리세요」 하면 `/개선공유` — 끝.

## 설치 — 30초

**요구 사항:** [Claude Code](https://docs.anthropic.com/en/docs/claude-code) · [Git](https://git-scm.com/) · Python 3.11+ · (개선 공유용) [GitHub CLI](https://cli.github.com/)

Claude Code 를 열고 이걸 붙여넣으면 Claude 가 알아서 한다.

> report-factory 설치: **`git clone https://github.com/FineHourTeam/report-factory.git`** 로 원하는 위치에 받고, 그 폴더에서 **`powershell -File setup.ps1`** (macOS/Linux 는 `./setup`) 을 실행한 다음, 그 폴더를 Claude Code 작업 폴더로 열어 줘.

클론한 폴더가 곧 작업 폴더다. 스킬·에이전트·엔진이 상대 경로로 서로를 찾으므로 **폴더 안 구조는 바꾸지 않는다.** 조사 자료는 `report-factory/projects/<조사명>/` 에 만들며, 이 폴더는 git 이 추적하지 않는다(고객 데이터는 리포에 올라가지 않는다).

## 스킬

| 스킬 | 하는 일 |
|---|---|
| `/제작가이드` | 통계표·설문지·조사개요 → 목차 설계 → 본문·요약·결론 블록 → 검증 → 제작팀 송부본(.xlsm) |
| `/보고서작성` | 제작팀 회신본의 반복 장표 채움 또는 예시 덱 직접 조립 → 수치 전수 검수 → 렌더 QA → 적대 감사 |
| `/회고` | 프로젝트가 남긴 증거에서 규칙 후보를 채굴, 승인분만 RULES.md·rules.yaml·knowledge 에 등재 |
| `/업그레이드` | 최신판 내려받기(git pull)·의존성 갱신·스모크 테스트 — 비개발자용 |
| `/개선공유` | 내가 다듬은 규칙·지식·도구 개선분을 PR 로 올린다 — git 몰라도 된다 |

## 에이전트 12명

| 팀 | 에이전트 |
|---|---|
| 가이드 설계 | guide-toc(목차) · guide-front(Part Ⅰ) · guide-body-ment(본문 멘트) |
| 요약·결론 | guide-summary(총괄·결론) · guide-summary-slide(꼭지별 문장) · guide-client-review(의뢰인 자리 심사) |
| 심사·검증 | guide-judge(설계 심사) · guide-summary-reviewer(품질) · guide-adversary(반증 감사) |
| PPT·회고 | deck-adversary(덱 감사) · deck-render-qa(렌더 검수) · retro-miner(회고 채굴) |

## 운영 모델 — 연구원은 git 을 모른다는 전제로 설계했다

```
연구원 1·2·3           관리자(리드)              GitHub (비공개)
   │ /업그레이드  ◀───────────────────────────  main (정본)
   │ 조사 작업 (스킬이 알아서)
   │ /회고 → 내 로컬 규칙 다듬어짐
   │ /개선공유  ─────▶  PR 검토·선별 병합  ───▶  main 갱신
```

- **알림**: `/제작가이드`·`/보고서작성`·`/회고` 를 부르면 원격의 판 번호를 먼저 읽는다(작업 폴더는 건드리지 않는다). 새 판이 있으면 「올릴까요?」를 묻고, 「나중에」는 24시간 → 48시간 → 1주일씩 미룬다.
- **받기**: `/업그레이드` 한 마디면 끝. 로컬 수정과 충돌하면 Claude 가 정리해 준다.
- **주기**: 관리자가 「PR 올리세요」 하면 각자 `/개선공유`. 스킬이 변경분을 요약해 보여 주고, 확인 후 브랜치·커밋·PR 까지 대신 한다. 작업 폴더는 건드리지 않는다.
- **고르기**: 관리자가 PR 중 공용으로 쓸 것만 병합한다. 취향이 갈리는 것은 프로젝트 폴더의 설정(하우스 프로파일)으로 내려보낸다.

## 구조

```
.claude/skills/      스킬 5개 (제작가이드 · 보고서작성 · 회고 · 업그레이드 · 개선공유)
.claude/agents/      에이전트 12명
report-factory/
  engine/tools/      파서·빌더·검수기·crosscheck 등 30여 개
  engine/knowledge/  블록 포맷·차트 결정표·멘트 문형·한국어 문장지침 …
  engine/learning/   RULES.md(규칙 대장) · rules.yaml · 회고 루프
  engine/templates/  제작가이드 xlsm 템플릿
  projects/          ← 조사 데이터. git 이 추적하지 않는다
```

## 여기 없는 것

- `report-factory/projects/` — **고객 데이터. 어떤 경우에도 리포에 올리지 않는다** (.gitignore 강제)
- `report-factory/engine/refs/` — 코퍼스 원자료 157MB. 규칙의 근거이며 사내 드라이브로 별도 전달
- API 키 — 전부 환경변수(`CLOVA_STUDIO_API_KEY` 등). 파일로 두지 않는다

이 리포는 **비공개**다. 규칙 문서에 실제 납품 사례의 실측 수치가 들어 있으므로 공개 전환하지 않는다.

## 업그레이드

스킬을 부를 때마다 `engine/tools/update_check.py` 가 원격 `VERSION` 을 읽어 새 판이 있으면 묻는다. gstack 의 판 확인을 옮겨 온 것이고, 상태(캐시·미룸·설정)는 `~/.report-factory/` 에 둔다.

- 「지금 올린다」 → `/업그레이드` 가 stash → pull → pop, 의존성 갱신, 스모크 테스트, 변경 요약까지 한다
- 「지금 말고 나중에」 → 24시간 뒤 다시 묻는다 (두 번째 48시간, 세 번째부터 1주일)
- 「앞으로는 자동으로」 / 「다시 묻지 않는다」 → `~/.report-factory/config.json` 의 `auto_upgrade` / `update_check`

직접 하려면 `/업그레이드` 라고 치거나, `git pull` 후 `pip install -r requirements.txt`.
