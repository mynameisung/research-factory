# `findings.json` 스키마와 엑셀 시트 구성

`scripts/make_report.py`의 입력. 완성 예시는 `assets/findings_example.json`.

시트는 **9개 고정**이다. 내용이 없어도 시트를 빼지 않는다 —
"검수했는데 없음"과 "검수 안 함"을 구분해야 하기 때문이다.

---

## 최상위 구조

```json
{
  "project":   { ... },
  "summary":   { "checks": [...], "overall": "..." },
  "questions": [ ... ],
  "issues":    [ ... ],
  "quota":     [ ... ],
  "base_logic":[ ... ],
  "structure": [ ... ],
  "codes":     [ ... ],
  "zero_opts": [ ... ],
  "false_positives": { "count": 0, "note": "..." },
  "scope":     { "verified": [...], "not_verified": [...] }
}
```

판정 값은 다섯 가지뿐이다: `적합` `이상` `확인필요` `참고` `미검증`.
다른 문자열을 쓰면 색이 안 입혀진다.

---

## `project` → `00_검수요약`

| 키 | 내용 |
|---|---|
| `name` | 프로젝트명 |
| `client` | 고객사 |
| `total_n` | 전체 표본수 — **테이블가이드 개요에서 읽은 값** |
| `questionnaire_file` / `guide_file` / `table_file` | 수령 파일명 |
| `checked_by` | 검수자 (기본: 사용자) |
| `checked_at` | 검수일 `YYYY-MM-DD` |
| `weighted` | 가중치 적용 여부와 방식 — **테이블가이드 개요에서 읽은 값** |

`total_n`과 `weighted`는 가이드가 출처다. 읽은 위치(시트·셀)를 `guide.json`의
`meta.source`에 남기고, 가이드에 없어 사용자에게 물은 경우에는
`"1000 (가이드 미기재, 사용자 확인)"`처럼 출처를 값에 붙여 적는다.

## `summary.checks` → `00_검수요약` 체크리스트 표

`key`는 고정 6종: `quota` `base` `logic` `code` `numbering` `structure`.

```json
{ "key": "base", "name": "2. Base 확인 (모든 문항)", "result": "이상",
  "count": 2, "note": "Q7, Q12 base 사례수 불일치." }
```

`count`를 생략하면 `issues`에서 `check`가 같은 건수를 자동으로 센다.
`summary.overall`은 종합 의견 — DP실에 회신할 문장 톤으로 쓴다.

---

## `questions` → `01_문항별검수`

**전 문항이 여기 있어야 한다.** 적합한 문항도 한 줄 남긴다.
이 시트가 "무엇을 봤는지"의 근거다.

| 키 | 내용 |
|---|---|
| `code` `title` `type` | 문번 / 문항 문구 / 응답유형 |
| `base_def` | 설문지의 base 정의 (문장 그대로) |
| `base_label` | 테이블에 표기된 base 문구 |
| `n_expected` | 기대 사례수. 산출 근거는 `note`에 |
| `n_actual` | 테이블 사례수. 표가 없으면 `null` |
| `opts_q` / `opts_tbl` | 보기 수 (설문지 / 테이블). 넷팅·평균 행은 빼고 센다 |
| `verdict` | 판정 |
| `note` | 판정 근거. 이상이면 반드시 채운다 |

## `issues` → `02_이슈목록`

심각도 순(`치명`→`중요`→`경미`→`확인필요`)으로 자동 정렬된다.

| 키 | 내용 |
|---|---|
| `severity` | `치명` / `중요` / `경미` / `확인필요` — 기준은 `checklist.md` |
| `check` | 6종 체크 키. 요약 시트의 건수 집계에 쓰인다 |
| `code` | 문번 |
| `category` | 이슈 유형 (`Base 사례수 불일치`, `문번 누락`, `역코딩` …) |
| `symptom` | 무엇이 어떻게 잘못됐는가. 숫자를 넣는다 |
| `evidence` | 근거 위치 — 설문지 페이지 / 가이드 행 / 테이블 블록 번호 |
| `action` | **DP실에 그대로 보낼 수 있는 문장으로** 쓴다 |
| `status` | `요청` / `확인중` / `수정완료` / `보류` |

## `quota` → `03_쿼터확인`

`group`(축) · `cell` · `target` · `actual` · `diff` · `verdict` · `note`.
각 축의 합계 행과 전체 합계 행도 넣는다.

## `base_logic` → `04_Base로직`

| 키 | 내용 |
|---|---|
| `code` | 문번 |
| `kind` | `Base` / `Prog` / `스킵` / `이탈` / `배타` / `표시조건` |
| `logic` | **설문지 문장 그대로.** 요약 금지 |
| `source` | 설문지 위치 |
| `expected` | 기대 base — 숫자와 산출 근거를 함께 (`"742 (Q6 1~3번 합)"`) |
| `actual` | 테이블 N |
| `verdict` `note` | |

## `structure` → `05_행열구조`

행 판정과 열 판정을 분리한다 (`row_verdict`, `col_verdict`).
가이드가 없으면 `guide_rows`/`guide_cols`에 `"가이드 미수령 — 설문지 기준"`이라 적는다.

## `codes` → `06_보기코드대조`

문항×보기 한 줄씩. 가장 큰 시트다.
전 보기를 다 넣기 어려우면 **최소한 이상·확인필요 보기와 각 문항의 첫·끝 보기**는 넣고,
`08_검증범위`에 "전수 대조 / 표본 대조" 중 무엇을 했는지 적는다.

척도 문항은 `note`에 넷팅 산술 검증 결과를 적는다 — "TOP2=412, 보기1+2=412 일치".

## `zero_opts` → `07_미출력보기`

**이슈가 아니다.** 응답 0건이라 DP가 빼고 낸 보기의 기록.

| 키 | 내용 |
|---|---|
| `code` `opt_code` `label` | 문번 / 보기 코드 / 보기 문구 |
| `value` | 확인한 셀 값 (0) |
| `verdict` | `참고` |
| `note` | "응답 0건으로 미출력. 이슈 아님." |

값이 0이 아니거나 확인 불가면 여기가 아니라 `issues`로 간다.

## `false_positives`

`count`와 `note`. 원시 경보 몇 건 중 몇 건을 왜 뺐는지 쓴다.
투명성이 리포트 신뢰의 근거다.

## `scope` → `08_검증범위`

```json
"scope": {
  "verified":     [ { "item": "전 문항 base 정의·사례수", "how": "설문지 base 지시문 → 앞 문항 사례수로 기대값 산출 후 대조" } ],
  "not_verified": [ { "item": "원자료 대비 집계 정확성", "reason": "raw data 미수령" } ]
}
```

문자열만 넣어도 되지만, `how`/`reason`을 채우는 편이 회신할 때 쓸모 있다.

**항상 `not_verified`에 들어가는 것들** — 원자료 대비 집계 정확성(raw data 없이는 불가),
가중치 산출 로직, 오픈형 코딩 품질, 유의차 검정값. 해당 없으면 그 사유를 적는다.

---

## 실행

```bash
python scripts/make_report.py findings.json -o "테이블검수_<프로젝트>_<YYYYMMDD>.xlsx"
```

`-o`를 생략하면 입력 파일명에 `.xlsx`를 붙인다.
Excel에서 파일이 열려 있으면 저장에 실패하니 닫고 실행한다.
