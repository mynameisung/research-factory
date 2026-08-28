# 제작가이드 스킬 — 다른 PC 에서 시작하기

2026-08-28 기준으로 묶었다. 이 압축본 하나로 `/제작가이드` · `/보고서작성` · `/회고` 세 스킬이 돈다.

---

## 1. 어디에 푸는가 — 여기가 제일 중요하다

**압축을 풀면 나오는 `embrain` 폴더를 통째로 작업 폴더로 삼는다.**
스킬과 도구가 서로를 **상대 경로**로 찾기 때문에 폴더 구조를 그대로 유지해야 한다.

```
어디든\embrain\              ← Claude Code 를 이 폴더에서 연다
├── .claude\
│   ├── skills\              제작가이드 · 보고서작성 · 회고
│   ├── agents\              담당 에이전트 12개
│   └── settings.json
└── report-factory\
    ├── engine\
    │   ├── tools\           파이썬 도구 40여 개
    │   ├── knowledge\       업무 지식 (목차 설계·차트타입·멘트 규칙 등)
    │   ├── learning\        RULES.md 규칙 대장 · rules.yaml · 담당 팩
    │   └── templates\       제작가이드 .xlsm 템플릿  ← 이게 없으면 산출물을 못 만든다
    ├── projects\            (비어 있음 — 아래 3번)
    └── README.md · INSTRUCTIONS.md · HANDOVER*.md
```

원래 PC 의 경로는 `C:\Users\flzkf\workspace\embrain` 이었다. 새 PC 에서 경로가 달라도 되지만
**`embrain` 아래의 구조는 바꾸지 않는다.**

풀고 나서 그 폴더에서 Claude Code 를 열면 `/제작가이드` 가 목록에 뜬다.

---

## 2. 파이썬 준비

Python 3.11 이상이면 된다 (원래 PC 는 3.13).

```bash
pip install openpyxl python-pptx lxml Pillow PyYAML xlrd
```

| 패키지 | 쓰이는 곳 | 원 PC 버전 |
|---|---|---|
| `openpyxl` | .xlsx/.xlsm 읽기·쓰기 — **필수** | 3.1.5 |
| `PyYAML` | `chart_rules.yaml` · `rules.yaml` — **필수** | 6.0.3 |
| `xlrd` | 구형 `.xls`(BIFF8) 통계표 파싱 | 2.0.2 |
| `python-pptx` · `lxml` · `Pillow` | 보고서작성(업무 B)의 PPT 조립·검수 | 1.0.2 · 6.1.1 · 12.1.1 |

**모든 파이썬 실행은 `PYTHONUTF8=1` 을 붙인다.** 안 붙이면 cp949 콘솔에서 한글이 깨지고
`--help` 조차 실패한다.

```bash
PYTHONUTF8=1 python engine/tools/build_facts.py ...
```

Windows PowerShell 이면 `$env:PYTHONUTF8=1` 을 한 번 걸어 두면 된다.

---

## 3. 이 압축본에 **없는** 것

용량 때문에 뺐다. 없어도 스킬은 돌지만, 아래 상황에서는 원래 PC 에서 가져와야 한다.

| 빠진 것 | 원래 크기 | 없으면 무엇이 안 되나 |
|---|---:|---|
| `engine/refs/` | 157 MB | 예시 제작가이드·제작본(LG·SKEE)과 참고자료 코퍼스 5사례. **규칙의 근거**이자 벤치마크 재료다. 새 규칙을 세우거나 실측으로 검증할 때 필요하다. 제작가이드를 **만드는 일** 자체에는 쓰이지 않는다 |
| `projects/` | 400 MB | 진행 중인 조사 데이터(통계표·설문지·산출물). 새 조사를 시작하는 데는 필요 없다 |

**진행 중인 에스원 작업을 이어가려면** `projects\2026_에스원_건물현황실사` 폴더를 따로 복사해
`report-factory\projects\` 아래에 넣는다. 같은 이름의 별도 압축본이 있으면 그것을 쓴다.

---

## 4. 먼저 읽을 것

새 세션에서 이어 할 때는 이 순서로 읽는다.

1. `report-factory/HANDOVER_260824.md` — 가장 최근 인수인계
2. `report-factory/engine/learning/RULES.md` — 확정 규칙 대장. **작업 전에 반드시 읽는다**
3. `report-factory/engine/knowledge/00_업무개요.md` · `01_통계표-읽는법.md`
4. `report-factory/engine/knowledge/guide/` — 목차 설계 · 블록 포맷 · 차트타입 · 멘트 · 팀구조 · 오케스트레이션

스킬을 부르면 `SKILL.md` 가 이 순서를 다시 안내한다.

---

## 5. 규칙 이름공간 — 헷갈리기 쉬운 자리

한 파일에 두 업무의 규칙이 섞여 있어 ID 를 갈라 두었다.

- `R-###` = **제작가이드**(업무 A)
- `R-PPT-###` = **보고서 제작**(업무 B)

새 규칙은 **자기 이름공간의 최대 번호 + 1** 이다. 등재 전에 중복을 확인한다.

```bash
grep -c "^#\{2,3\} R-<번호>" report-factory/engine/learning/RULES.md
```

---

## 6. 동작 확인

풀고 나서 이 두 줄이 돌면 준비가 끝난 것이다.

```bash
PYTHONUTF8=1 python report-factory/engine/tools/rules.py
```
(오류 없이 끝나면 규칙 엔진과 PyYAML 이 정상이다)

```bash
PYTHONUTF8=1 python -c "import openpyxl, yaml, xlrd; print('ok')"
```

템플릿이 살아 있는지도 본다 — **15 시트**가 나와야 한다
(끝의 `_치트시트`·`_스니펫` 둘은 제작팀 송부본을 만들 때 빠져 13 시트가 된다).

```bash
PYTHONUTF8=1 python -c "import openpyxl; wb=openpyxl.load_workbook(r'report-factory/engine/templates/TEMPLATE_제작가이드.xlsm', keep_vba=True); print(len(wb.sheetnames), wb.sheetnames)"
```

---

## 7. 이 압축본에 담긴 엔진 개선 (2026-08-25~28)

에스원 TA Pool 작업에서 잡은 것들이다. 이전 판으로 되돌리면 같은 결함이 재발한다.

| 도구 | 무엇 |
|---|---|
| `parse_table_xlsx.py` · `parse_table_biff.py` | `--in-groups` — 배너 절단선을 프로젝트별로 준다. 없으면 비상주 배너군이 하드코딩돼 IN 열이 「전체」 1열만 남는다 |
| `parse_table_common.py` | `평균 (개월)` 처럼 **대괄호 없는 평균 행**을 인식한다. 못 읽으면 `그래프 상단 Text` 수식어가 통째로 누락된다 |
| `rules.py` | 평균 단위를 표에서 읽는다(`점` 하드코딩 제거) · 자기참조 배너 **유형 ③ 분할 100%** 추가 · `sort_option_rows(resort=)` |
| `build_facts.py` | `--notes` 각주 · `--freq-banners` 빈도 표기 · `--mean-units` · `--chart-type` 오버라이드 · `--resort` · 셀 서식 기반 자릿수 · 목차 문항코드 폴백 · 동률을 `line1` 밖까지 검출 |
| `build_guide.py` | `--profile-groups` · `--profile-extra-first` · `--banner-note` · 장별 각주 기입 · 구조 블록 `Chart type` · 값 자릿수에 맞춘 셀 서식 |
| `apply_table_fixes.py` | **표 제목 정정**(`제목정정`) 지원 |
| `verify.py` | 패널 제목 정규식 확장(`A2- b.` 꼴) · 빈도 열 역산 검산 · 목차 제외 문항 반영 |
| `chart_rules.yaml` | `CT-890` — 기본형이라도 평균 행이 있으면 `그래프 상단 Text` 를 붙인다 |

⚠ **비상주(02_비상주)를 다시 빌드하면 산출물이 달라진다.** 평균 단위 8표가 바로잡히고
(`점` → 층·개·만원) 평균 행이 있는 표에 수식어가 새로 붙는다. r11 이 이미 납품된 상태라
재빌드 전에 연구원 확인이 필요하다.
