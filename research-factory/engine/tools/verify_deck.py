#!/usr/bin/env python3
"""제작본 `.pptx` 검수 — 납품 직전 마지막 관문 (업무 B).

`22_검수-체크리스트.md` 의 A(치명)·B(기계)·C(반자동) 항목 중 **기계로 판정 가능한 것**을
D0~D26 검사로 돌린다. `verify.py` 의 `Finding`/`Report`/`render` 를 그대로 공유하므로
가이드 모드와 리포트 형식·실패 태그가 같다.

설계 원칙 세 가지
-----------------
1. **기대값은 통계표·설문지·rules.py 에서만 나온다.** 예시 제작본은 회귀 스냅샷일 뿐
   기대값 소스가 아니다 (R-104). LG s6 의 `[Base : 전체 응답자(n=1,000), Unit : %]` 에서
   Base 형식을 배우면 오답이 규격이 된다.
2. **파싱을 다시 만들지 않는다.** 슬라이드 팩트는 `extract_ground_truth.extract_deck()`,
   수치 규칙은 `rules.py`, 기대 그리드는 `build_facts.build_grid()` 가 이미 낸다.
3. **없는 입력은 조용히 통과시키지 않는다.** 통계표가 없으면 값 대조를 건너뛰되
   `stats` 에 `미대조` 로 남긴다. 0건 대조가 만점으로 위장하는 것이 R-088 공허 일치다 —
   LG 본문 79장이 `top_ment` 전부 공백이라 멘트 검사가 0건으로 '통과' 하던 그 함정이다.

검사 항목
--------
  D0   패널 게이트        표·차트가 없는 본문 장표는 데이터 검사에서 뺀다
  D1   조인 · 문항코드     통계표 매칭 3단 폴백 · 중복 문항코드(장분리/오타)
  D2   미기입 마커        `???` · `내용을 입력하세요` · `Data 입력` · 표 셀 `?`      ★치명
  D3   미충전 반복 장표    표 값영역 채움도 (메모가 0건인 덱에서도 판정)             ★치명
  D4   제작팀 메모 잔존    채운 뒤 지웠어야 할 `Data 입력` 쪽지                      ★치명
  D5   Base n 정합        `[Base: …(n=NN)…]` ↔ 통계표 base_n                       ★치명
  D6   Base/Unit 형식     콜론 앞 공백 · `/` 대신 `,`
  D7   n<30 각주          별표 붙은 배너가 있으면 각주가 있어야 한다
  D8   표 셀 ↔ 기대 그리드                                                        ★치명
  D9   차트 ↔ 표          카테고리축과 표 데이터 행 라벨 정합
  D10  소수점 자릿수      % 1자리 · 평균 2자리 (뒤 0 유지)
  D11  합계 정합          TOP+SoSo+BOT = 100 · `계` = 100
  D12  0 표기             `-` / 공란 혼용                                    (가설·참고)
  D13  센티널 표기        【TOP2】 ↔ [7점 평균] 괄호 혼용
  D14  그래프 상단 Text   표 상단 행 / 독립 TextBox 매체별 표기
  D15  정렬               `rules.table_row_axis` 와 대조
  D16  차트타입           지시 ↔ 실제                                    (참고·규칙 후보)
  D17  Qu 텍스트          설문지 원문 대조
  D18  멘트 수치 환각     멘트의 모든 숫자가 그 슬라이드에 실제로 찍혀 있는가        ★치명
  D19  멘트 보기명        따옴표 안 보기명 ↔ 데이터 행 라벨
  D20  멘트 문체          개조식 · 마침표 없음
  D21  수준어 사다리      인용 값과 형용사가 맞는가                                 ★치명
  D22  따옴표 짝
  D23  배너 선별 타당성   2-proportion z 로 재계산                                 (참고)
  D24  목차 · 간지 정합
  D25  겹침막대 라벨 충돌                                                    (가설·참고)
  D26  제목 변경 이력     제작팀이 제목을 바꾼다 (R-093) — 오류가 아니라 이력이다

사용법
------
    python verify.py <제작본.pptx> [tables.json] --mode deck [--facts …]
        [--questionnaire …] [--guide-blocks …] [--deck-facts ground_truth.json]
        [--known-defects LG] [--out 검수리포트.md] [--json findings.json]

회귀 자산
--------
두 예시 덱에는 통계표(tables.json)가 없어 값 대조 경로(D5·D8·D10·D15·D18)를 실덱으로는
한 번도 태울 수 없다. 그래서 결함을 **일부러 심은** 최소 픽스처를 두었다:

    python verify.py engine/refs/fixture_deck_min/deck.pptx \\
        engine/refs/fixture_deck_min/tables.json --mode deck \\
        --deck-facts engine/refs/fixture_deck_min/deck_facts.json

기대 결과는 치명 3(D18 환각 · D5 Base n · D8 셀값) · 경고 5(D10 · D6×3 · D15) · 참고 1(D16).
`deck.pptx` 는 실제로 없다 — `--deck-facts` 가 있으면 파일을 열지 않는다.
"""

from __future__ import annotations

import io
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Iterator

sys.path.insert(0, str(Path(__file__).resolve().parent))

import build_facts as BF               # noqa: E402
import extract_ground_truth as EGT     # noqa: E402
import rules as R                      # noqa: E402
from verify import (                   # noqa: E402
    Report, check_base_format, check_ment_style, check_sentinels,
    find_check, load_rules, numbers_in,
)

# ─────────────────────────────────────────────────────────────
# 셀 패턴 — 표 행의 역할은 **인덱스가 아니라 라벨/값 모양**으로 판정한다
# ─────────────────────────────────────────────────────────────

#: `(100)` · `(8*)` — Base 행을 찾는 유일한 신호. 첫 셀이 `(Base)` 라는 가정은 못 쓴다.
#: LG 척도 장표는 헤더가 **표 하단**에 있고 base 행 첫 셀이 공란이라
#: `extract_table` 의 `base_row_index` 가 null 을 준다 (실측 s38).
BASE_CELL_RE = re.compile(r"^\(\d[\d,]*\*?\)$")

#: 표 셀에 남은 미기입 흔적. **정확 일치**여야 한다 — 부분 일치로 `?` 를 잡으면
#: 정상 라벨('얼마나 만족하십니까?')을 전건 오탐한다. LG 실측 660셀 / 34장.
SENTINEL_CELL_RE = re.compile(r"^[?？]+$")

#: 값처럼 보이는 셀. 헤더 행(라벨만 있는 행)과 데이터 행을 가르는 데 쓴다.
NUMERIC_CELL_RE = re.compile(r"^-?\d[\d,]*\.?\d*%?$")

#: 요약 행 라벨. 매체가 달라도(`【TOP2】` vs `TOP2(%)`) 같은 역할이다.
TOP_ROW_RE = re.compile(r"^\s*[\[【]?\s*(TOP|BOT|SoSo)\s*(\d*)")
MEAN_ROW_RE = re.compile(r"(\d+)\s*점\s*평균")
NPS_ROW_RE = re.compile(r"NPS")
TOTAL_ROW_RE = re.compile(r"^\s*계\s*$")

#: 범례 셀(`■ 7 매우 그렇다 …`). 값 칸이 통째로 비어 있는 게 정상이라
#: 채움도 계산에서 빼지 않으면 정상 표가 PARTIAL 로 오판된다.
LEGEND_MARK = "■"

#: 보기가 **카테고리축**인 차트. 이 유형에서만 차트 카테고리 = 표 데이터 행 라벨이다.
#: 척도(±세로막대)는 카테고리가 배너이고, 누적 100% 는 카테고리가 `['전체']` 하나뿐이라
#: 같은 잣대를 대면 전건 오탐한다 (실측 LG 41장 · SKEE 42장).
CATEGORY_IS_OPTION = ("가로막대", "가로겹침막대")

#: 값이 없다는 표기. `-` 는 **채워진 0** 이지 공란이 아니다 (R-110). SKEE 에 190개 있다.
ZERO_MARKS = ("-", "0.0")

#: PARTIAL 판정 임계. 두 덱에서 **한 번도 발화하지 않은 미검증 상태**이고,
#: 정상 표의 최대 blank 비율이 0.464(SKEE s82) 라 여유가 0.14 뿐이다.
#: 그래서 치명이 아니라 경고로 낸다 — 사람이 확인할 큐에 올리는 용도다.
PARTIAL_BLANK_RATIO = 0.6


def norm_cell(s: Any) -> str:
    """전각공백(U+3000)·NBSP·개행을 지운 비교용 문자열.

    이 정규화가 없으면 실측 셀들이 전부 어긋난다 —
    0 표기 `'-　'` 는 `strip()` 으로 안 지워지고, 배너 라벨은 `'1~2인\\n가구'` 처럼
    개행이 끼며, base 값은 `'(8*)'` 처럼 n<30 별표를 달고 온다.
    """
    if s is None:
        return ""
    t = str(s).replace("　", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", t).strip()


def norm_label(s: Any) -> str:
    """라벨 비교용 — 보기번호 접두(`3) `)와 공백까지 지운다."""
    t = R.chart_label(norm_cell(s))
    return re.sub(r"\s+", "", t)


# ─────────────────────────────────────────────────────────────
# 덱 로드
# ─────────────────────────────────────────────────────────────

def load_deck(deck: Path, deck_facts: Path | None = None) -> dict:
    """덱 팩트를 얻는다. 파싱을 재구현하지 않는다.

    `--deck-facts` 로 이미 뽑아 둔 `ground_truth.json` 을 주면 그걸 쓴다.
    99장 파싱이 1초 남짓이라 없어도 되지만, 골든셋 하네스처럼 같은 덱을 여러 번
    돌리는 자리에서는 매번 여는 게 낭비다.
    """
    if deck_facts and Path(deck_facts).exists():
        with io.open(deck_facts, encoding="utf-8") as fh:
            doc = json.load(fh)
        doc["_facts_source"] = str(deck_facts)
        return doc
    doc = EGT.extract_deck(deck)
    doc["_facts_source"] = "extract_deck()"
    return doc


def body_slides(doc: dict) -> list[dict]:
    return [s for s in doc.get("slides") or [] if s.get("classification") == "본문"]


def slide_panel(slide: dict) -> dict | None:
    """그 슬라이드가 **데이터 패널**(표 또는 차트)을 가졌는가.

    None 이면 D5~D16 을 통째로 건너뛴다. SKEE 본문 84장 중 17장이
    `table_count=0 · chart_count=0` 인 정성 인용 장표이고 15장은 base_unit·qu_text 까지
    공백이다 — 게이트 없이 데이터 검사를 돌리면 BASE 치명 15건이 전부 오탐이다.
    """
    tables = slide.get("tables") or []
    charts = slide.get("charts") or []
    if not tables and not charts:
        return None
    return {"table": tables[0] if tables else None,
            "chart": charts[0] if charts else None,
            "n_tables": len(tables), "n_charts": len(charts)}


# ─────────────────────────────────────────────────────────────
# 제작본 표 — 헤더 블록 복원과 행 역할 판정
# ─────────────────────────────────────────────────────────────

def table_base_row(table: dict) -> int | None:
    """(Base) 행을 **값 패턴**으로 찾는다.

    `(NN)` 꼴 셀이 `(n_cols-1)//2` 개 이상인 행이 base 행이다. 첫 셀이 `(Base)` 인지로
    찾으면 안 된다 — ±세로막대 표는 헤더가 하단에 있고 base 행 첫 셀이 공란이다.
    """
    rows = table.get("rows") or []
    need = max(1, (int(table.get("n_cols") or 0) - 1) // 2)
    for i, row in enumerate(rows):
        if sum(1 for c in row if BASE_CELL_RE.fullmatch(norm_cell(c))) >= need:
            return i
    return None


def table_value_columns(table: dict, base_row: int) -> list[int]:
    """(Base) 행에서 `(NN)` 인 열만 값 열로 본다. 라벨 열과 공란 열은 제외."""
    row = (table.get("rows") or [])[base_row]
    return [j for j, c in enumerate(row) if BASE_CELL_RE.fullmatch(norm_cell(c))]


def table_header(table: dict) -> dict | None:
    """배너 헤더 블록(그룹행 · 라벨행 · base행)을 복원한다.

    헤더는 표 위에 있을 수도(일반 배너표) 아래에 있을 수도(척도 장표) 있으므로
    base 행에서 **양쪽으로 인접한 라벨 행**을 훑어 붙인다. 그룹 행은 병합셀이라
    비어 있는 칸이 많다 — 그래서 '비지 않은 칸이 더 적은 쪽'이 그룹 행이다.
    """
    bi = table_base_row(table)
    if bi is None:
        return None
    cols = table_value_columns(table, bi)
    if not cols:
        return None
    rows = [[norm_cell(c) for c in r] for r in (table.get("rows") or [])]

    def is_label_row(r: list[str]) -> bool:
        vals = [r[j] for j in cols if j < len(r)]
        filled = [v for v in vals if v]
        # 값이 하나라도 숫자꼴이면 데이터 행이다. 라벨이 2개 미만이면 헤더로 볼 근거가 없다.
        return len(filled) >= 2 and not any(NUMERIC_CELL_RE.fullmatch(v) for v in filled)

    header = {bi}
    label_rows: list[int] = []
    for step in (-1, 1):
        for k in (1, 2):
            i = bi + step * k
            if 0 <= i < len(rows) and (i - step) in header and is_label_row(rows[i]):
                header.add(i)
                label_rows.append(i)

    group_row = label_row = None
    if len(label_rows) >= 2:
        # 병합셀이라 그룹 행은 빈 칸이 많다 — 비지 않은 칸이 **적은 쪽**이 그룹 행이다.
        # 실측 LG s25 (그룹 5 / 라벨 15) · s38 (그룹 4 / 라벨 11). 동수면 base 행에서
        # 먼 쪽을 그룹으로 본다 (헤더 블록은 언제나 [그룹][라벨][base] 순이다).
        counts = {i: sum(1 for j in cols if j < len(rows[i]) and rows[i][j]) for i in label_rows}
        order = sorted(label_rows, key=lambda i: (counts[i], -abs(i - bi)))
        group_row, label_row = order[0], order[-1]
    elif len(label_rows) == 1:
        label_row = label_rows[0]

    return {"base_row": bi, "value_cols": cols, "header_rows": sorted(header),
            "group_row": group_row, "label_row": label_row, "rows": rows}


def banner_axis(table: dict) -> list[dict] | None:
    """제작본 표 한 장의 배너 구조 — 열마다 (group, label, n, small).

    그룹은 병합셀이라 forward-fill 하고, 라벨이 비면 그룹명이 곧 라벨이다(`전체`).
    `*` 는 n<30 표식이므로 떼고 int 로 읽되 `small=True` 로 남긴다.
    열 수를 상수로 가정하지 않는다 — 실측 8/9/11/12/15/16/17/24열이다 (R-095).
    """
    h = table_header(table)
    if not h:
        return None
    rows = h["rows"]
    out: list[dict] = []
    group = ""
    for j in h["value_cols"]:
        g = rows[h["group_row"]][j] if h["group_row"] is not None and j < len(rows[h["group_row"]]) else ""
        if g:
            group = g
        lab = rows[h["label_row"]][j] if h["label_row"] is not None and j < len(rows[h["label_row"]]) else ""
        raw = rows[h["base_row"]][j]
        m = re.search(r"(\d[\d,]*)", raw)
        out.append({
            "col": j,
            "group": group,
            "label": lab or group,
            "n": int(m.group(1).replace(",", "")) if m else None,
            "small": "*" in raw,
        })
    return out


def iter_table_rows(deck_table: dict) -> Iterator[dict]:
    """표를 한 줄기로 훑으며 각 행의 **역할**을 라벨 패턴으로 판정한다.

    `extract_table` 의 `header_rows`/`data_rows` 분할을 믿으면 안 된다 —
    LG s38 은 header_rows 에 TOP3·평균이 들어가고 data_rows 는
    [범례][배너 라벨][base][배너 그룹] 순으로 역할이 뒤섞여 있다.
    """
    h = table_header(deck_table)
    rows = [[norm_cell(c) for c in r] for r in (deck_table.get("rows") or [])]
    cols = h["value_cols"] if h else list(range(1, len(rows[0]) if rows else 0))
    header = set(h["header_rows"]) if h else set()

    for i, row in enumerate(rows):
        label = row[0] if row else ""
        cells = [row[j] if j < len(row) else "" for j in cols]
        if i in header:
            if h and i == h["base_row"]:
                role = "base"
            elif h and i == h["group_row"]:
                role = "banner_group"
            else:
                role = "banner_label"
        elif LEGEND_MARK in label and not any(cells):
            role = "legend"
        elif TOP_ROW_RE.match(label):
            kind = TOP_ROW_RE.match(label).group(1)
            role = {"TOP": "top", "BOT": "bot", "SoSo": "soso"}[kind]
        elif MEAN_ROW_RE.search(label) or NPS_ROW_RE.search(label):
            role = "mean"
        elif TOTAL_ROW_RE.match(label):
            role = "total"
        else:
            role = "option"
        yield {"index": i, "role": role, "label": label, "cells": cells, "cols": cols}


def table_fill_state(deck_table: dict) -> tuple[str, int, int, int]:
    """(state, filled, blank, sentinel) — 표 값영역 채움도.

    ★ 이것이 미충전을 잡는 **유일하게 100% 맞는 신호**다. 실측 대조:
      LG `Data 입력` 메모 64장 → EMPTY 30 + SENTINEL 34, 비-메모 15장 → 전부 FILLED,
      불일치 0건. SKEE(메모 0건) 표 있는 본문 67장 전부 FILLED, 오탐 0.

    차트로는 판정이 원리적으로 불가능하다 — 제작팀의 '빈 차트' 는 None 이 아니라
    합이 100이 되도록 고른 더미 상수(50.0 / 25.0 / ±10.0)이고, 더 나쁘게는
    대표 슬라이드를 **값째로 복사**해 온다(LG s93 이 s90 의 값을 그대로 들고 있다).
    """
    h = table_header(deck_table)
    if not h:
        return ("NOBASE", 0, 0, 0)
    filled = blank = sentinel = 0
    for row in iter_table_rows(deck_table):
        if row["role"] in ("base", "banner_group", "banner_label", "legend"):
            continue
        for c in row["cells"]:
            if SENTINEL_CELL_RE.fullmatch(c):
                sentinel += 1
            elif not c:
                blank += 1
            else:
                filled += 1
    if sentinel:
        return ("SENTINEL", filled, blank, sentinel)
    if filled == 0:
        return ("EMPTY", filled, blank, sentinel)
    if blank and blank / (filled + blank) >= PARTIAL_BLANK_RATIO:
        return ("PARTIAL", filled, blank, sentinel)
    return ("FILLED", filled, blank, sentinel)


def rendered_values(slide: dict) -> set[str]:
    """그 슬라이드에 **실제로 찍힌** 모든 표시값. D18 멘트 환각 검출의 대조 집합.

    표 셀 · 차트 데이터 · 상단 TextBox 를 전부 긁는다. 차트 값은 저장된 무한소수라
    1·2자리 반올림 표기까지 함께 넣는다. ±세로막대는 음의 반쪽에 −1 을 곱해 두므로
    절댓값도 넣는다 — 멘트는 언제나 양수로 인용한다.
    """
    out: set[str] = set()

    def add_number_tokens(text: str) -> None:
        for m in re.finditer(r"-?\d[\d,]*\.?\d*", text or ""):
            out.add(m.group().replace(",", ""))
            out.add(m.group())

    for t in slide.get("tables") or []:
        for row in t.get("rows") or []:
            for c in row:
                add_number_tokens(norm_cell(c))
    for ch in slide.get("charts") or []:
        for ser in ch.get("series") or []:
            for v in ser.get("values") or []:
                if v is None:
                    continue
                a = abs(v)
                out.add(f"{R.round_half_up(a, 1):.1f}")
                out.add(f"{R.round_half_up(a, 2):.2f}")
                out.add(f"{R.round_half_up(a, 0):.0f}")
    for sh in slide.get("text_shapes") or []:
        add_number_tokens(sh.get("text") or "")
    return out


# ─────────────────────────────────────────────────────────────
# 기대값 — 전부 통계표에서 나온다 (R-104)
# ─────────────────────────────────────────────────────────────

def expected_grid(table: dict, rank_first: dict | None = None) -> dict:
    """기대 그리드. `build_facts.build_grid()` 를 그대로 부른다 — 재구현 금지."""
    return BF.build_grid(table, rank_first)


def expected_top_text(table: dict, decision: R.ChartDecision) -> list[dict]:
    """`그래프 상단 Text` 기대값. 매체별 표기는 rules.py 가 정한다.

    척도는 표 상단 행에 `TOP3(%)` · `7점 평균(점)` 로, 사다리형(누적 100%)은
    독립 TextBox 에 `【TOP2】` 로 들어간다 — 표 데이터 행만 뒤지면 전건 오탐이다.
    """
    return R.graph_top_text(table, decision.top_text_kinds, decision.top_text_medium)


# ─────────────────────────────────────────────────────────────
# 조인 — 문항코드가 있다는 가정을 하지 않는다 (V-091)
# ─────────────────────────────────────────────────────────────

def _option_labels(deck_table: dict) -> set[str]:
    return {norm_label(r["label"]) for r in iter_table_rows(deck_table)
            if r["role"] == "option" and r["label"]}


def _table_option_labels(table: dict) -> set[str]:
    return {norm_label(r.get("label")) for r in (table.get("rows") or [])
            if r.get("kind") == "option"}


def _jaccard(a: set[str], b: set[str]) -> float:
    a = {x for x in a if x}
    b = {x for x in b if x}
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def join_slides(deck_doc: dict, tables: list[dict],
                facts_doc: dict | None) -> tuple[dict[int, dict], list[tuple[int, str]]]:
    """덱 slide_no ↔ 통계표 index 조인. 3단 폴백이고 실패는 전부 MAP 으로 남긴다.

    ① question_code (LG 본문 79/79 가 이 경로)
    ② slide_facts 의 source_tables (문항코드가 같은 표가 여러 장일 때 갈라 준다)
    ③ 표 보기 라벨 집합 자카드 (SKEE 84/84 가 문항코드 'unknown' 이라 이 경로가 필수)

    ★ 슬라이드 번호로 seq 를 맞추려 들면 안 된다. 덱은 1-based slide_no 이고 본문이
      LG 기준 s16 부터 시작하는데 slide_facts.seq 는 본문만 세는 번호다.
    """
    by_code: dict[str, list[dict]] = defaultdict(list)
    for t in tables:
        code = (t.get("question_code") or "").strip().upper()
        if code:
            by_code[code].append(t)
    by_index = {t["index"]: t for t in tables}

    # 같은 문항코드가 여러 장인 문항(장분리·매트릭스 소항목)은 팩트도 여러 개다 —
    # 첫 장만 기억하면 뒷장들이 전부 첫 장의 표로 오조인된다 (실측 2026-08-28:
    # B7-1 웹 장이 모바일 표39 로, C2 7장이 전부 첫 표로 붙어 위양성 치명 32건).
    facts_by_code: dict[str, list[dict]] = defaultdict(list)
    for f in (facts_doc or {}).get("slides") or []:
        code = (f.get("question_code") or "").strip().upper()
        if code:
            facts_by_code[code].append(f)

    #: 같은 문항코드를 쓰는 덱 슬라이드의 **등장 순번**. 라벨로 갈리지 않을 때 쓴다.
    code_seq: dict[int, int] = {}
    _seen: dict[str, int] = {}
    for _s in body_slides(deck_doc):
        _c = (_s.get("question_code") or "").strip().upper()
        code_seq[_s["slide_no"]] = _seen.get(_c, 0)
        _seen[_c] = _seen.get(_c, 0) + 1

    def pick_fact(code: str, deck_slide: dict, deck_labels: set) -> dict | None:
        """코드가 같은 팩트가 여럿이면 덱의 Base n(장분리 판별자)으로, 없으면 라벨로 가른다."""
        flist = facts_by_code.get(code) or []
        if len(flist) <= 1:
            return flist[0] if flist else None
        bn = next((b.get("base_n") for b in deck_slide.get("base_items") or []
                   if b.get("base_n") is not None), None)
        hit = [f for f in flist if bn is not None and f.get("base_n") == bn]
        if len(hit) == 1:
            return hit[0]
        if deck_labels:
            def fl(f):
                return {str(r.get("label_clean") or r.get("label") or "").strip()
                        for r in (f.get("grid") or {}).get("rows") or []}
            scored = sorted(((_jaccard(deck_labels, fl(f)), f) for f in flist),
                            key=lambda kv: -kv[0])
            # 자카드가 **동률**이면 라벨로는 갈리지 않는다. 구간을 접어 두 표의 보기가
            # 똑같아지는 자리가 실제로 있다 — 비상주 A1 은 표7(지상 층수)을 표10(총 층수)과
            # 같은 구간으로 접어 네 보기가 한 글자도 다르지 않다(R-PPT-008 해소, 2026-09-01).
            # 그때는 **등장 순서**로 짝짓는다. 본문 슬라이드 순서와 팩트 page 순서는 같다.
            if len(scored) > 1 and abs(scored[0][0] - scored[1][0]) < 1e-9:
                idx = code_seq.get(deck_slide.get("slide_no"), 0)
                ordered = sorted(flist, key=lambda f: f.get("page") or f.get("seq") or 0)
                if idx < len(ordered):
                    return ordered[idx]
            return scored[0][1]
        return flist[0]

    joined: dict[int, dict] = {}
    unjoined: list[tuple[int, str]] = []

    for s in body_slides(deck_doc):
        no = s["slide_no"]
        panel = slide_panel(s)
        deck_tbl = (panel or {}).get("table")
        deck_labels = _option_labels(deck_tbl) if deck_tbl else set()
        code = (s.get("question_code") or "").strip().upper()

        cands = list(by_code.get(code) or [])
        how = "question_code"
        fact = pick_fact(code, s, deck_labels)
        if len(cands) > 1:
            picked = [by_index[i] for i in (fact or {}).get("source_tables") or []
                      if i in by_index]
            if picked:
                cands, how = picked, "question_code+facts"
            elif deck_labels:
                cands = sorted(cands,
                               key=lambda t: -_jaccard(deck_labels, _table_option_labels(t)))
                cands, how = cands[:1], "question_code+라벨 자카드"
        if not cands and deck_labels:
            scored = sorted(((_jaccard(deck_labels, _table_option_labels(t)), t)
                             for t in tables), key=lambda kv: -kv[0])
            if scored and scored[0][0] >= 0.6:
                cands, how = [scored[0][1]], f"라벨 자카드 {scored[0][0]:.2f}"

        if cands:
            joined[no] = {"tables": cands, "main": cands[0], "how": how,
                          "facts": fact}
        else:
            why = "문항코드 없음 + 표 라벨 자카드 미달" if not code else f"문항코드 {code} 가 통계표에 없음"
            if not deck_tbl and not (panel or {}).get("chart"):
                why = "데이터 패널 없음 (정성 인용 장표)"
            unjoined.append((no, why))
    return joined, unjoined


# ─────────────────────────────────────────────────────────────
# 검사
# ─────────────────────────────────────────────────────────────

def _where(slide: dict, extra: str = "") -> str:
    """`s26 (A10-2) 표 14` 형식. 문항코드가 없는 덱도 있으므로 있을 때만 붙인다."""
    head = f"s{slide['slide_no']}"
    code = (slide.get("question_code") or "").strip()
    if code:
        head += f" ({code})"
    return f"{head} {extra}".strip()


def _relabel(rep: Report, mark: int, mapping: dict[str, str], tag: str) -> None:
    """재사용한 가이드 모드 검사가 붙인 검사 ID(V9/V10/V11)를 덱 ID(D…)로 바꾼다.

    검사 로직은 그대로 두고 라벨만 갈아 끼운다 — 리포트에서 두 모드의 ID 체계가
    섞이면 골든셋 집계가 어긋난다.
    """
    for f in rep.findings[mark:]:
        f.check = mapping.get(f.severity, mapping.get("*", f.check))
        f.tag = tag


def check_unfilled(rep: Report, rules: dict, slide: dict, markers: list[str]) -> int:
    """D2/D3/D4 — 미기입 마커 · 미충전 반복 장표 · 제작팀 메모 잔존.

    마커 목록의 **정본은 rules.yaml V-070** 이다. `extract_ground_truth.UNFILLED_MARKERS`
    는 `TBD`·`XXX` 를 넣고 `텍스트를 입력`·`Data 입력` 을 빼서 목록이 다르다 —
    슬라이드 팩트의 `unfilled_markers` 만 믿으면 `Data 입력` 메모를 통째로 놓친다.
    """
    n = 0
    pat = re.compile("|".join(re.escape(m) for m in markers))
    # 제작팀 메모는 D4 가 전담한다. 여기서 또 잡으면 한 쪽지가 두 건으로 세어진다.
    memo_shapes = {m.get("shape_name") for m in slide.get("production_memos") or []}
    for sh in slide.get("text_shapes") or []:
        if sh.get("name") in memo_shapes:
            continue
        text = sh.get("text") or ""
        hit = pat.search(text)
        if hit:
            rep.add("D2", "치명", _where(slide, f"[{sh.get('name')}]"),
                    f"미기입 마커 `{hit.group()}` 가 남아 있습니다: {norm_cell(text)[:70]!r}",
                    tag="UNFILLED")
            n += 1

    # 표 셀은 **정확 일치**로 본다. 부분 일치로 `?` 를 잡으면 정상 문장을 오탐한다.
    for t in slide.get("tables") or []:
        hits = [(i, j) for i, row in enumerate(t.get("rows") or [])
                for j, c in enumerate(row) if SENTINEL_CELL_RE.fullmatch(norm_cell(c))]
        if hits:
            rep.add("D2", "치명", _where(slide, str(t.get("shape_name"))),
                    f"미기입 셀 `?` 가 {len(hits)}칸 남아 있습니다 "
                    f"(첫 자리 {hits[0][0]}행 {hits[0][1]}열)", tag="UNFILLED")
            n += 1

    for m in slide.get("production_memos") or []:
        rep.add("D4", "치명", _where(slide, f"[{m.get('shape_name')}]"),
                f"제작팀 메모가 지워지지 않았습니다: {norm_cell(m.get('text'))[:70]!r}\n"
                f"      → 채운 뒤 반드시 삭제합니다 (R-082/R-098)", tag="UNFILLED")
        n += 1
    return n


def check_fill_state(rep: Report, slide: dict, deck_table: dict) -> str:
    """D3 — 표 값영역이 통째로 비었거나 `?` 로 남아 있는가."""
    state, filled, blank, sentinel = table_fill_state(deck_table)
    where = _where(slide, str(deck_table.get("shape_name")))
    if state == "NOBASE":
        rep.add("D3", "참고", where,
                "(Base) 행을 찾지 못해 채움도를 판정하지 못했습니다 — 사람이 확인하세요",
                tag="UNFILLED")
    elif state == "EMPTY":
        rep.add("D3", "치명", where,
                f"데이터 칸 {blank}개가 전부 비어 있습니다 — 미충전 반복 장표입니다",
                tag="UNFILLED")
    elif state == "SENTINEL":
        rep.add("D3", "치명", where,
                f"데이터 칸에 `?` 가 {sentinel}개 남아 있습니다 (채움 {filled} · 공란 {blank})",
                tag="UNFILLED")
    elif state == "PARTIAL":
        rep.add("D3", "경고", where,
                f"데이터 칸의 {blank}/{filled + blank} 가 비어 있습니다 "
                f"(임계 {PARTIAL_BLANK_RATIO:.0%}) — 정상 표의 최대 공란 비율이 46% 라 "
                f"여유가 크지 않습니다. 사람이 확인하세요", tag="UNFILLED")
    return state


def check_base(rep: Report, rules: dict, slide: dict,
               deck_table: dict | None, table: dict | None) -> None:
    """D5/D6/D7 — Base n 정합 · Base/Unit 형식 · n<30 각주.

    형식 검사는 `verify.check_base_format()` 을 그대로 재사용한다 (LG s6 의 실제 오류를
    잡는 그 검사다). n 대조만 여기서 덧붙인다.
    """
    bu = norm_cell(slide.get("base_unit"))
    where = _where(slide)
    if not bu:
        if deck_table is not None or (slide.get("charts") or []):
            rep.add("D5", "경고", where, "Base/Unit 표기가 없습니다", tag="BASE")
        return

    mark = len(rep.findings)
    check_base_format(rep, rules, {"values": {"Base/Unit": bu}}, where)
    _relabel(rep, mark, {"*": "D6"}, "BASE")

    m = EGT.BASE_N_RE.search(bu)
    if m and table is not None:
        n = int(m.group(1).replace(",", ""))
        want = table.get("base_n")
        if want and n != want:
            rep.add("D5", "치명", where,
                    f"Base n={n:,} 인데 통계표 표{table['index']} 는 {want:,} 입니다\n"
                    f"      → 기대값은 통계표에서만 나옵니다. 제작본 숫자를 정답으로 삼지 않습니다",
                    tag="BASE")

    # D7 — n<30 배너가 있으면 각주가 있어야 한다. 문자열은 프로젝트마다 다르므로
    #      존재 여부만 본다 (rules.yaml V-032).
    axis = banner_axis(deck_table) if deck_table else None
    if axis and any(b["small"] for b in axis):
        if not (slide.get("base_unit_footnotes") or []):
            small = [b["label"] for b in axis if b["small"]][:4]
            rep.add("D7", "경고", where,
                    f"n<30 배너가 {sum(1 for b in axis if b['small'])}개인데 각주가 없습니다 "
                    f"(예: {', '.join(small)})", tag="BASE")


def _bind_banners(axis: list[dict], table: dict) -> dict[int, int]:
    """제작본 표 열 ↔ 통계표 banners[pos] 매칭. 위치가 아니라 (group,label) 문자열로 한다.

    열 수를 상수로 가정하지 않는다 (R-095). 못 붙은 열은 그냥 빠지고 BANNER 로 집계된다.
    """
    want: dict[tuple[str, str], int] = {}
    for b in R.report_banners(table):
        want[(norm_label(b.get("group")), norm_label(b.get("label")))] = b["pos"]
    out: dict[int, int] = {}
    for col in axis:
        key = (norm_label(col["group"]), norm_label(col["label"]))
        if key in want:
            out[col["col"]] = want[key]
    return out


def check_data(rep: Report, slide: dict, deck_table: dict,
               table: dict, grid: dict, axis: list[dict]) -> int:
    """D8 — 표 셀 ↔ 기대 그리드.

    행 단위로 묶어 내고 첫 불일치 열만 상세에 적는다. 값 단위로 내면 배너 열 하나가
    밀렸을 때 한 행에서 수십 건이 터져 리포트를 못 읽는다.
    """
    bind = _bind_banners(axis, table)
    if not bind:
        rep.add("D8", "경고", _where(slide, str(deck_table.get("shape_name"))),
                f"제작본 배너 열 {len(axis)}개 중 통계표 배너에 붙은 열이 없습니다 "
                f"— 값 대조를 건너뜁니다", tag="BANNER")
        return 0

    expected: dict[str, dict] = {}
    for row in grid.get("rows") or []:
        expected.setdefault(norm_label(row.get("label_clean") or row.get("label")), row)

    col_index = {c["col"]: k for k, c in enumerate(axis)}
    compared = 0
    for row in iter_table_rows(deck_table):
        if row["role"] in ("base", "banner_group", "banner_label", "legend"):
            continue
        key = norm_label(row["label"])
        want = expected.get(key)
        if want is None:
            continue
        bad: list[str] = []
        for col, pos in bind.items():
            k = col_index.get(col)
            if k is None or k >= len(row["cells"]):
                continue
            got = row["cells"][k]
            exp = (want.get("cells") or [None] * (pos + 1))[pos] if pos < len(want.get("cells") or []) else None
            if exp is None:
                continue
            compared += 1
            if got in ("", *ZERO_MARKS) and exp in ("-", "0.0"):
                continue        # 0 표기는 D12 가 따로 본다
            if not got:
                continue        # 전체 열 공란처럼 설계상 비는 칸 — D3 가 본다
            if got == exp:
                continue
            # 값은 같은데 표기만 다른 것(`25` vs `25.0`)은 소수점 문제이지 데이터 오류가
            # 아니다. D10 이 이미 잡으므로 여기서 또 치명으로 올리면 한 결함이 두 건이 된다.
            try:
                gf, ef = float(got.replace(",", "")), float(exp.replace(",", ""))
                if abs(gf - ef) < 1e-9:
                    continue
                # r14: 기대 그리드는 원값(R-191), 덱 표는 표시값이다 — 원값을 그 자리의
                # 표시 자릿수로 round-half-up 한 값과 같으면 데이터는 같은 것이다.
                from decimal import Decimal, ROUND_HALF_UP
                dg = len(got.split(".")[-1]) if "." in got else 0
                if str(Decimal(exp.replace(",", "")).quantize(
                        Decimal("1." + "0" * dg) if dg else Decimal("1"),
                        rounding=ROUND_HALF_UP)) == got:
                    continue
            except (ValueError, ArithmeticError):
                pass
            bad.append(f"{axis[k]['group']}/{axis[k]['label']} {got}≠{exp}")
        if bad:
            rep.add("D8", "치명", _where(slide, f"표 행 {row['index']}"),
                    f"`{row['label'][:28]}` 행이 통계표 표{table['index']} 와 다릅니다 "
                    f"({len(bad)}칸) — 첫 자리: {bad[0]}", tag="DATA")
    return compared


def check_chart_vs_table(rep: Report, slide: dict, deck_table: dict,
                         chart: dict, chart_type: str) -> None:
    """D9 — 차트 카테고리축 ↔ 표 데이터 행 라벨.

    보기가 카테고리축인 유형(가로막대·가로겹침막대)에서만 성립한다. 척도는 카테고리가
    **배너**이고 누적 100% 는 `['전체']` 하나뿐이라 같은 잣대를 대면 전건 오탐이다.
    제작팀이 차트 라벨에만 괄호 설명을 붙이는 일이 흔하므로 접두 일치까지 허용한다.
    """
    if not any(k in (chart_type or "") for k in CATEGORY_IS_OPTION):
        return
    cats = [norm_cell(c) for c in chart.get("categories") or []]
    labs = [row["label"] for row in iter_table_rows(deck_table)
            if row["role"] == "option" and row["label"]]
    if not cats or not labs or len(cats) != len(labs):
        return
    for i, (c, l) in enumerate(zip(cats, labs)):
        cn, ln = norm_label(c), norm_label(l)
        if cn == ln or cn.startswith(ln) or ln.startswith(cn):
            continue
        rep.add("D9", "참고", _where(slide, f"차트 카테고리 {i + 1}"),
                f"차트 라벨과 표 행 라벨이 다릅니다\n"
                f"      차트: {c[:60]}\n      표  : {l[:60]}", tag="DATA")
        return


def check_decimals(rep: Report, slide: dict, deck_table: dict) -> int:
    """D10 — 소수점 자릿수. % 는 1자리, 평균·NPS 는 2자리이고 **뒤 0을 살린다**.

    `33` 이 아니라 `33.0`, `4.2` 가 아니라 `4.20` 이다 (rules.yaml V-003).
    """
    spec = {"mean": 2}
    checked = 0
    for row in iter_table_rows(deck_table):
        if row["role"] in ("base", "banner_group", "banner_label", "legend"):
            continue
        want = spec.get(row["role"], 1)
        bad = []
        for c in row["cells"]:
            if not c or c in ZERO_MARKS or SENTINEL_CELL_RE.fullmatch(c):
                continue
            if not re.fullmatch(r"-?\d[\d,]*(\.\d+)?", c):
                continue
            checked += 1
            got = len(c.split(".")[1]) if "." in c else 0
            if got != want:
                bad.append(c)
        if bad:
            rep.add("D10", "경고", _where(slide, f"표 행 {row['index']}"),
                    f"`{row['label'][:24]}` 행의 소수점이 {want}자리가 아닙니다 "
                    f"({len(bad)}칸, 예: {', '.join(bad[:4])})", tag="DATA")
    return checked


def check_sums(rep: Report, slide: dict, deck_table: dict, table: dict | None) -> None:
    """D11 — 합계 정합. **일괄 적용하면 안 된다.**

    중복응답·1~N순위는 합 100 초과가 정상이고 `계` 행이 아예 없어야 한다 (V-004).
    통계표가 있으면 hints 로 가르고, 없으면 자기정합(TOP+SoSo+BOT=100)만 본다.
    """
    hints = (table or {}).get("hints") or {}
    by_role: dict[str, list[list[str]]] = defaultdict(list)
    for row in iter_table_rows(deck_table):
        by_role[row["role"]].append(row["cells"])

    def as_floats(cells: list[str]) -> list[float | None]:
        out = []
        for c in cells:
            try:
                out.append(float(c.replace(",", "")))
            except ValueError:
                out.append(None)
        return out

    # 척도 자기정합 — TOP + SoSo + BOT = 100
    if by_role.get("top") and by_role.get("bot"):
        top = as_floats(by_role["top"][0])
        bot = as_floats(by_role["bot"][0])
        soso = as_floats(by_role["soso"][0]) if by_role.get("soso") else [0.0] * len(top)
        for k in range(min(len(top), len(bot), len(soso))):
            vals = (top[k], bot[k], soso[k])
            if any(v is None for v in vals):
                continue
            total = sum(vals)
            if abs(total - 100.0) > 0.1:
                rep.add("D11", "경고", _where(slide, f"배너열 {k + 1}"),
                        f"TOP+SoSo+BOT = {total:.1f} 입니다 (100.0 ±0.1 이어야 합니다)",
                        tag="DATA")
                break

    if by_role.get("total"):
        if hints.get("is_multi") or (hints.get("is_ranking") and not hints.get("rank_is_first_only")):
            rep.add("D11", "경고", _where(slide),
                    "중복응답 · 1~N순위 표에는 `계` 행이 없어야 합니다", tag="DATA")
        else:
            vals = as_floats(by_role["total"][0])
            for k, v in enumerate(vals):
                if v is not None and abs(v - 100.0) > 0.1:
                    rep.add("D11", "경고", _where(slide, f"배너열 {k + 1}"),
                            f"`계` 가 {v:.1f} 입니다 (100.0 ±0.1 이어야 합니다)", tag="DATA")
                    break


def check_zero_notation(rep: Report, slide: dict, deck_table: dict) -> None:
    """D12 — 0 표기가 한 표 안에서 섞이는가.

    ⚠ R-110(단일선택 `-` / 중복응답 공란)은 실측이 반증했다. 0 이 실재하는 LG 본문
    10장 중 9장이 공란이고 1장(s20)만 `-` 이며 같은 `가로막대` 안에서도 갈린다 —
    문항 유형이 아니라 담당자 습관이다. 그래서 **참고로만** 낸다.

    ★ `전체` 열은 세지 않는다. 가로막대 장표는 표의 전체 열 데이터 칸이 **원래 비어 있고**
      (전체값은 차트가 들고 있다) 그걸 0 표기로 읽으면 SKEE 25장이 전건 오탐이다.
    """
    axis = banner_axis(deck_table) or []
    skip = {k for k, b in enumerate(axis) if norm_label(b["label"]) == "전체"}
    dash = blank = 0
    for row in iter_table_rows(deck_table):
        if row["role"] in ("base", "banner_group", "banner_label", "legend"):
            continue
        for k, c in enumerate(row["cells"]):
            if k in skip:
                continue
            if c == "-":
                dash += 1
            elif not c:
                blank += 1
    if dash and blank:
        rep.add("D12", "참고", _where(slide, str(deck_table.get("shape_name"))),
                f"0 표기가 섞여 있습니다 — `-` {dash}칸 · 공란 {blank}칸 "
                f"(R-110 은 가설이라 판정하지 않고 알려만 드립니다)", tag="DATA")


def check_sentinel_labels(rep: Report, rules: dict, slide: dict, deck_table: dict) -> None:
    """D13 — 센티널 라벨 표기. `verify.check_sentinels()` 를 그대로 재사용한다.

    괄호 종류가 섞이면(【】↔[]) 제작팀 RPA 가 그 행을 못 알아본다.
    다만 제작본 표는 **매체 표기**(`TOP3(%)` · `7점 평균(점)`)를 쓰므로 가이드 표기와
    다른 게 정상이다 — 독립 TextBox 표기(`【TOP2】`)만 이 검사에 건다.
    """
    rows = [{"row": r["index"], "label": r["label"], "modifier": None, "cells": r["cells"]}
            for r in iter_table_rows(deck_table) if r["label"].startswith(("【", "["))]
    if not rows:
        return
    mark = len(rep.findings)
    check_sentinels(rep, rules, {"rows": rows})
    for f in rep.findings[mark:]:
        f.check = "D13"
        f.tag = "LABEL"
        f.where = _where(slide, f.where)


def check_toptext(rep: Report, slide: dict, deck_table: dict | None,
                  expected: list[dict]) -> int:
    """D14 — `그래프 상단 Text`. 표 상단 행과 독립 TextBox **양쪽**을 뒤진다.

    어디에 찍히는지가 차트타입마다 다르다 — 척도는 표 header 행에 `TOP3(%)`,
    사다리형은 독립 TextBox 에 `【TOP2】\\n60.0 %` 다. 표 데이터 행만 뒤지면 전건 오탐이다.
    """
    if not expected:
        return 0
    haystack = " ".join(norm_cell(sh.get("text")) for sh in slide.get("text_shapes") or [])
    if deck_table:
        haystack += " " + " ".join(r["label"] for r in iter_table_rows(deck_table))
    checked = 0
    for item in expected:
        checked += 1
        label = item["label"]
        key = re.sub(r"[\s【】\[\]()%점]", "", label)
        flat = re.sub(r"[\s【】\[\]()%점]", "", haystack)
        if key and key not in flat:
            rep.add("D14", "경고", _where(slide),
                    f"`그래프 상단 Text` 가 없습니다 — {label} ({item['medium']})\n"
                    f"      안 적으면 Top·평균이 그래프 위에 표시되지 않습니다", tag="TOPTEXT")
    return checked


def check_order(rep: Report, slide: dict, deck_table: dict,
                table: dict, decision: R.ChartDecision) -> None:
    """D15 — 표 데이터 행 순서. `rules.table_row_axis` 가 기대 순서를 낸다 (SORT-045).

    갈림길은 하나 — 보기가 카테고리축인가 계열인가. 척도·누적 100% 는 표가 통계표
    순서를 유지하고, 가로막대·겹침막대·순위는 값 정렬이 표에도 똑같이 걸린다.
    """
    want, _rows, why = R.table_row_axis(table, decision.option_axis)
    got = [row["label"] for row in iter_table_rows(deck_table)
           if row["role"] == "option" and row["label"]]
    if not got or not want or len(got) != len(want):
        return
    if [norm_label(x) for x in got] == [norm_label(x) for x in want]:
        return
    first = next((i for i, (a, b) in enumerate(zip(got, want))
                  if norm_label(a) != norm_label(b)), 0)
    rep.add("D15", "경고", _where(slide, f"표 행 {first + 1}"),
            f"행 순서가 기대와 다릅니다 — 기대 근거: {why}\n"
            f"      제작본: {' / '.join(got[:4])}\n"
            f"      기대  : {' / '.join(want[:4])}", tag="SORT")


def check_chart_type(rep: Report, slide: dict, decision: R.ChartDecision | None,
                     guide_type: str | None) -> None:
    """D16 — 차트타입 대조. **게이트가 아니라 수집이다.**

    C-6 대로 불일치가 곧 오류가 아니다 — LG 79장 중 7장 불일치에서 3건은 연구원 쪽이
    틀렸다(s25~29·s17·s93). `derived_chart_type_confidence` 가 certain 이 아니면
    더더욱 판정 근거가 못 된다(SKEE heuristic 17장). 전부 참고로만 낸다.
    """
    got = slide.get("derived_chart_type")
    conf = slide.get("derived_chart_type_confidence")
    for want, src in ((decision.chart_type if decision else None, "통계표 결정표"),
                      (guide_type, "제작가이드 지시")):
        if not want or not got or want == got:
            continue
        rep.add("D16", "참고", _where(slide),
                f"차트타입이 {src}와 다릅니다 — 지시 `{want}` / 실제 `{got}` "
                f"(판독 신뢰도 {conf})\n"
                f"      → 오류가 아닐 수 있습니다. TYPE 은 규칙 후보 수집용입니다",
                tag="TYPE")


def check_qu(rep: Report, slide: dict, table: dict | None,
             questions: dict[str, dict]) -> None:
    """D17 — Qu 원문 대조. `Qu: ` 접두 + 설문지 원문 그대로."""
    qu = norm_cell(slide.get("qu_text"))
    if not qu or not questions or table is None:
        return
    q = questions.get((table.get("question_code") or "").strip())
    core = norm_cell((q or {}).get("text"))
    if not core:
        return
    if re.sub(r"\s+", "", core) not in re.sub(r"\s+", "", qu):
        rep.add("D17", "경고", _where(slide),
                f"Qu 가 설문지 원문과 다릅니다\n      제작본: {qu[:70]}\n      설문지: {core[:70]}",
                tag="QA")


def check_question_code_hygiene(rep: Report, deck_doc: dict) -> None:
    """D1 — 중복 문항코드. 연속이면 장분리, 떨어져 있으면 오타다.

    LG 실측: C2-1/C2-2/C2-3[81,82] · C3[83,84] · C4[87,88] 은 장분리(R-089)이고
    BB1-5[49,63] 은 오타다(63은 BD1-5 자리). 장분리 2장째의 차트가 1장째와 같은 건
    **정상**이다 — 복제 흔적으로 오해하면 안 된다.
    """
    seen: dict[str, list[int]] = defaultdict(list)
    for s in body_slides(deck_doc):
        code = (s.get("question_code") or "").strip()
        if code:
            seen[code].append(s["slide_no"])
    for code, nos in sorted(seen.items()):
        if len(nos) < 2:
            continue
        if all(b - a == 1 for a, b in zip(nos, nos[1:])):
            rep.add("D1", "참고", f"s{nos[0]}-{nos[-1]} ({code})",
                    f"같은 문항코드가 연속 {len(nos)}장입니다 — 배너가 쪼개진 장분리로 봅니다 (R-089)",
                    tag="MAP")
        else:
            rep.add("D1", "경고", f"s{'·s'.join(str(n) for n in nos)} ({code})",
                    f"같은 문항코드가 떨어진 자리에 있습니다 — 소제목 오타를 의심하세요",
                    tag="MAP")


def check_ment(rep: Report, rules: dict, slide: dict,
               values: set[str], deck_table: dict | None) -> tuple[int, int]:
    """D18~D22 — 멘트 검사. **D18 이 이 파일의 존재 이유다.**

    멘트에 나오는 모든 숫자가 그 슬라이드에 **실제로 찍힌 값**인지 본다. 통계표가 없어도
    돌아가는 자기정합 검사이고, 독자가 보는 화면 기준이라 오히려 더 엄격하다.
    실측: SKEE 본문 84장 멘트 97개 수치 전수 대조에서 불일치 0건.

    ⚠ LG 본문 79장은 top_ment 가 전부 공백이다. 0건 대조는 통과가 아니라 **미측정**이고
      그건 `stats.ment_numbers_checked` 로 리포트 상단에 찍힌다 (R-088 공허 일치).
    """
    ment = slide.get("top_ment") or ""
    lines = [l.strip() for l in ment.splitlines() if l.strip()]
    if not lines:
        return (0, 0)
    where = _where(slide)

    n_numbers = 0
    for line in lines:
        for tok in numbers_in(line):
            n_numbers += 1
            if tok not in values:
                rep.add("D18", "치명", where,
                        f"멘트의 수치 {tok} 가 이 슬라이드에 찍혀 있지 않습니다 ★환각 의심\n"
                        f"      멘트: {line[:80]}", tag="MENT-NUM")

    mark = len(rep.findings)
    check_ment_style(rep, rules, {"ment": lines}, where)
    _relabel(rep, mark, {"치명": "D21", "경고": "D20", "참고": "D20"}, "MENT-STYLE")

    # D22 따옴표 짝 — 실측 사고 SKEE C762 `소통 원활'` (여는 따옴표 누락)
    for line in lines:
        if line.count("'") % 2 or line.count('"') % 2:
            rep.add("D22", "경고", where,
                    f"따옴표 짝이 맞지 않습니다: {line[:80]}", tag="MENT-STYLE")

    # D19 멘트 보기명 ↔ 데이터 행 라벨 — 실측 사고 SKEE C744 `타건설사` vs `타 건설사`
    if deck_table is not None:
        labels = {norm_label(r["label"]) for r in iter_table_rows(deck_table)
                  if r["role"] == "option" and r["label"]}
        seen: set[str] = set()
        for line in lines:
            for quoted in re.findall(r"'([^']{2,40})'", line):
                key = norm_label(quoted)
                if not key or key in labels or key in seen:
                    continue
                seen.add(key)
                near = [l for l in labels if key in l or l in key]
                if near:
                    rep.add("D19", "경고", where,
                            f"멘트 보기명이 데이터 행 라벨과 다릅니다 — "
                            f"멘트 '{quoted}' / 행 라벨 '{sorted(near)[0]}'", tag="MENT-LABEL")
    return (len(lines), n_numbers)


def check_ment_pick(rep: Report, slide: dict, table: dict | None) -> None:
    """D23 — 배너 선별 타당성. 2-proportion z 로 다시 계산해 본다 (V-064).

    멘트 2행이 인용한 배너가 통계적으로 뽑을 만한 자리였는지 확인한다.
    임계는 실무 관행에서 온 것이라 **참고로만** 낸다.
    """
    ment = slide.get("top_ment") or ""
    if not ment or table is None:
        return
    facts = BF.build_ment_facts(table)
    picked = [d.get("banner_label") or d.get("label") or "" for d in facts.get("line2") or []]
    if not picked:
        return
    line2 = ment.splitlines()[-1] if len(ment.splitlines()) > 1 else ""
    if not line2 or any(norm_label(p) and norm_label(p) in norm_label(line2) for p in picked):
        return
    rep.add("D23", "참고", _where(slide),
            f"멘트 2행이 인용한 배너가 유의차 계산 결과와 다릅니다 — "
            f"계산상 후보: {', '.join(p for p in picked if p)}\n      멘트: {line2[:70]}",
            tag="MENT-PICK")


def check_label_clash(rep: Report, slide: dict, chart: dict, chart_type: str) -> None:
    """D25 — 겹침막대 두 계열 값 격차가 작으면 dLblPos=outEnd 라벨이 겹친다 (R-111).

    자동 해결이 불가능하므로 경고만 띄우고 사람이 옮긴다. ⚠ 임계 12%p 는 랜덤 데이터
    3건에서 뽑은 값이라 근거가 약하다 — 그래서 참고 고정이다.
    """
    if "겹침막대" not in (chart_type or ""):
        return
    sers = [s for s in chart.get("series") or [] if s.get("is_real")]
    if len(sers) < 2:
        return
    a, b = sers[0].get("values") or [], sers[1].get("values") or []
    clashes = [i for i, (x, y) in enumerate(zip(a, b))
               if x is not None and y is not None and abs(x - y) < 12.0]
    if clashes:
        rep.add("D25", "참고", _where(slide),
                f"겹침막대 라벨이 겹칠 수 있습니다 — 두 계열 격차 12%p 미만이 "
                f"{len(clashes)}자리 (카테고리 {clashes[0] + 1}번 등). 사람이 위치를 조정하세요",
                tag="LABEL")


def _toc_key(s: str) -> str:
    """목차 항목 비교용 — 앞머리 번호(`1.` `2)` `Part Ⅲ.`)와 공백을 지운다."""
    t = norm_cell(s)
    t = re.sub(r"^(Part\s*[ⅠⅡⅢⅣⅤIVX]+\s*[.·]?\s*)", "", t, flags=re.IGNORECASE)
    t = re.sub(r"^[A-Z0-9ⅠⅡⅢⅣⅤ]+\s*[.)]\s*", "", t)
    return re.sub(r"\s+", "", t)


def check_deck_toc(rep: Report, deck_doc: dict) -> None:
    """D24 — 목차 ↔ Part 간지 ↔ 챕터 간지 ↔ 본문 제목 정합.

    ⚠ 제목 불일치는 오류가 아니라 **변경 이력**일 수 있다 (C-7/R-093, 실측
    `건설사 대비` → `경쟁사 대비`). 목차와 간지가 **같이** 바뀌었는지만 본다.
    """
    slides = deck_doc.get("slides") or []
    toc_text: set[str] = set()
    for s in slides:
        if s.get("classification") not in ("목차", "간지"):
            continue
        for sh in s.get("text_shapes") or []:
            for line in (sh.get("text") or "").splitlines():
                key = _toc_key(line)
                if key:
                    toc_text.add(key)
    if not toc_text:
        return
    missing: dict[str, list[int]] = defaultdict(list)
    for s in body_slides(deck_doc):
        title = norm_cell(s.get("slide_title"))
        if not title:
            continue
        if _toc_key(title) not in toc_text:
            missing[title].append(s["slide_no"])
    for title, nos in sorted(missing.items(), key=lambda kv: kv[1][0]):
        rep.add("D24", "참고", f"s{nos[0]}",
                f"본문 제목이 목차·간지에 없습니다 ({len(nos)}장): {title!r}\n"
                f"      → 제작팀이 제목을 바꿨다면 목차와 간지도 같이 바꿔야 합니다 (R-093)",
                tag="TOC")


def check_title_history(rep: Report, deck_doc: dict, guide_blocks: list[dict]) -> None:
    """D26 — 제작가이드 제목 ↔ 제작본 제목. 오류가 아니라 이력으로 낸다."""
    if not guide_blocks:
        return
    guide_titles = {_toc_key(b.get("slide_title")) for b in guide_blocks if b.get("slide_title")}
    guide_titles.discard("")
    if not guide_titles:
        return
    seen: dict[str, int] = {}
    for s in body_slides(deck_doc):
        title = norm_cell(s.get("slide_title"))
        if title and _toc_key(title) not in guide_titles:
            seen.setdefault(title, s["slide_no"])
    for title, no in sorted(seen.items(), key=lambda kv: kv[1]):
        rep.add("D26", "참고", f"s{no}",
                f"제작가이드에 없는 제목입니다: {title!r} — 제작팀 변경 이력으로 봅니다 (R-093)",
                tag="TOC")


# ─────────────────────────────────────────────────────────────
# R-104 — 알려진 예시 결함 억제
# ─────────────────────────────────────────────────────────────

def _parse_defect_ref(ref: str) -> tuple[str | None, list[int]]:
    """`deck:LG#s26-29` · `#s82,s84,s86,s88` · `#본문 79장` 세 문법을 가른다.

    앞 둘은 슬라이드 주소이므로 그 슬라이드의 Finding 을 격하하고,
    `#본문 79장`·`#orphan` 처럼 주소가 안 되는 것은 **검사 자체를 끄는** scope 억제다.
    """
    if "#" not in ref:
        return (None, [])
    head, tail = ref.split("#", 1)
    deck_id = head.split(":", 1)[1] if ":" in head else None
    nos: list[int] = []
    for part in tail.split(","):
        m = re.fullmatch(r"\s*s(\d+)(?:-s?(\d+))?\s*", part)
        if not m:
            continue
        a = int(m.group(1))
        b = int(m.group(2) or a)
        nos.extend(range(a, b + 1))
    return (deck_id, nos)


def suppress_known(rep: Report, deck_id: str | None, chart_rules: dict) -> int:
    """알려진 예시 결함을 '예시오류' 로 격하한다. **--known-defects 를 명시했을 때만.**

    억제는 '이 오류를 이미 안다'는 표시일 뿐이고, 예시 덱의 오답을 기대값으로
    되먹이면 안 된다 (R-104). LG s6 과 남의 프로젝트 s6 은 아무 관계가 없으므로
    실제 프로젝트 덱에서는 절대 켜지면 안 된다.
    """
    if not deck_id:
        return 0
    want = deck_id.strip().lower()
    slide_tags: dict[int, set[str] | None] = {}
    scope_tags: set[str] = set()
    for entry in chart_rules.get("known_example_defects") or []:
        ref = str(entry.get("ref") or "")
        if not ref.lower().startswith("deck:"):
            continue
        did, nos = _parse_defect_ref(ref)
        if (did or "").strip().lower() != want:
            continue
        tag = entry.get("tag")
        if nos:
            for n in nos:
                if tag is None:
                    slide_tags[n] = None
                elif slide_tags.get(n, set()) is not None:
                    slide_tags.setdefault(n, set()).add(tag)
        elif tag:
            scope_tags.add(tag)

    n = 0
    for f in rep.findings:
        m = re.match(r"s(\d+)", f.where)
        no = int(m.group(1)) if m else None
        tag = f.tag or ""
        hit = False
        if no in slide_tags:
            allowed = slide_tags[no]
            hit = allowed is None or any(tag.startswith(t) for t in allowed)
        if not hit and any(tag.startswith(t) for t in scope_tags):
            hit = True
        if hit and f.severity != "예시오류":
            f.severity = "예시오류"
            n += 1
    return n


# ─────────────────────────────────────────────────────────────
# 진입점
# ─────────────────────────────────────────────────────────────

def verify_deck(deck: Path, tables: Path | None = None,
                facts: Path | None = None, questionnaire: Path | None = None,
                guide_blocks: Path | None = None, deck_facts: Path | None = None,
                known_defects: str | None = None) -> Report:
    """덱 모드 진입점. `verify.verify()` 와 반환 타입이 같아 `render()` 를 그대로 탄다."""
    rep = Report()
    rules = load_rules()
    if not rules:
        rep.add("D0", "경고", "engine/learning/rules.yaml",
                "규칙 파일을 읽지 못해 기본 검사만 수행했습니다 (PyYAML 설치 확인)")
    chart_rules = R.load_chart_rules()

    deck_doc = load_deck(deck, deck_facts)
    slides = deck_doc.get("slides") or []
    body = body_slides(deck_doc)

    tlist: list[dict] = []
    if tables and Path(tables).exists():
        with io.open(tables, encoding="utf-8") as fh:
            tdoc = json.load(fh)
        tlist = tdoc["tables"] if isinstance(tdoc, dict) else tdoc

    fdoc = None
    if facts and Path(facts).exists():
        with io.open(facts, encoding="utf-8") as fh:
            fdoc = json.load(fh)

    questions: dict[str, dict] = {}
    if questionnaire and Path(questionnaire).exists():
        with io.open(questionnaire, encoding="utf-8") as fh:
            questions = {q["code"]: q for q in json.load(fh).get("questions", [])
                         if q.get("code")}

    gblocks: list[dict] = []
    if guide_blocks and Path(guide_blocks).exists():
        with io.open(guide_blocks, encoding="utf-8") as fh:
            gdoc = json.load(fh)
        gblocks = gdoc.get("blocks") if isinstance(gdoc, dict) else (gdoc or [])
    guide_by_code = {(b.get("question_code") or "").strip(): b
                     for b in gblocks if b.get("question_code")}

    markers = (find_check(rules, "unfilled_markers", "V-070").get("search_strings")
               or ["???", "내용을 입력하세요", "텍스트를 입력", "Data 입력"])

    joined, unjoined = join_slides(deck_doc, tlist, fdoc) if tlist else ({}, [])
    for no, why in unjoined:
        rep.add("D1", "참고", f"s{no}", f"통계표 매칭 실패 — {why}", tag="MAP")
    check_question_code_hygiene(rep, deck_doc)

    # ── 덱 전체 검사 (본문 밖 슬라이드의 마커도 잡아야 한다: LG s2 목차 · s14 결론)
    for s in slides:
        if s.get("classification") != "본문":
            check_unfilled(rep, rules, s, markers)
    check_deck_toc(rep, deck_doc)
    check_title_history(rep, deck_doc, gblocks)

    panels = 0
    cells = 0
    decimals = 0
    toptext = 0
    ment_lines = 0
    ment_numbers = 0
    fill_states: Counter = Counter()

    for s in body:
        check_unfilled(rep, rules, s, markers)
        ml, mn = check_ment(rep, rules, s, rendered_values(s),
                            (slide_panel(s) or {}).get("table"))
        ment_lines += ml
        ment_numbers += mn

        panel = slide_panel(s)
        if panel is None:
            # D0 게이트 — 정성 인용 장표는 데이터 검사에서 통째로 뺀다
            continue
        panels += 1
        deck_table = panel["table"]
        chart = panel["chart"]
        chart_type = s.get("derived_chart_type") or ""

        if deck_table is not None:
            fill_states[check_fill_state(rep, s, deck_table)] += 1
            decimals += check_decimals(rep, s, deck_table)
            check_zero_notation(rep, s, deck_table)
            check_sentinel_labels(rep, rules, s, deck_table)
            if chart is not None:
                check_chart_vs_table(rep, s, deck_table, chart, chart_type)
        if chart is not None:
            check_label_clash(rep, s, chart, chart_type)

        hit = joined.get(s["slide_no"])
        table = (hit or {}).get("main")
        decision = None
        if table is not None:
            q = questions.get((table.get("question_code") or "").strip())
            decision = R.decide_chart_type(table, q, chart_rules)
        check_base(rep, rules, s, deck_table, table)
        check_qu(rep, s, table, questions)
        gb = guide_by_code.get((s.get("question_code") or "").strip())
        check_chart_type(rep, s, decision, (gb or {}).get("chart_type"))

        if table is None or deck_table is None:
            continue
        check_sums(rep, s, deck_table, table)
        rank_first = None
        if len(hit["tables"]) > 1:
            pairs = R.pair_ranking_tables(hit["tables"])
            rank_first = pairs[0][0] if pairs else None
        grid = expected_grid(table, rank_first)
        axis = banner_axis(deck_table)
        if axis:
            cells += check_data(rep, s, deck_table, table, grid, axis)
        check_order(rep, s, deck_table, table, decision)
        if decision is not None:
            toptext += check_toptext(rep, s, deck_table, expected_top_text(table, decision))
        check_ment_pick(rep, s, table)

    suppressed = suppress_known(rep, known_defects, chart_rules)

    rep.stats.update({
        "mode": "deck",
        "deck": Path(deck).name,
        "facts_source": deck_doc.get("_facts_source"),
        "slides": len(slides),
        "body_slides": len(body),
        "panels_checked": panels,
        "joined": len(joined),
        "unjoined": len(unjoined),
        "tables": len(tlist),
        "cells_compared": cells,
        "decimals_checked": decimals,
        "toptext_checked": toptext,
        "ment_lines": ment_lines,
        "ment_numbers_checked": ment_numbers,
        "fill_states": dict(fill_states),
        "suppressed_known_defects": suppressed,
        "rules_loaded": bool(rules),
        "blocks": len(body),
    })
    return rep


def deck_header_lines(rep: Report) -> list[str]:
    """리포트 상단에 붙일 덱 모드 요약. **미대조를 미대조라고 쓴다.**"""
    st = rep.stats
    lines = [
        f"- 슬라이드 {st.get('slides', 0)}장 · 본문 {st.get('body_slides', 0)}장 · "
        f"데이터 패널 {st.get('panels_checked', 0)}장 "
        f"(패널 없는 본문 {st.get('body_slides', 0) - st.get('panels_checked', 0)}장은 데이터 검사 제외)",
        f"- 통계표 조인 {st.get('joined', 0)} / 실패 {st.get('unjoined', 0)} · "
        f"표 셀 {st.get('cells_compared', 0)}칸 대조 · 소수점 {st.get('decimals_checked', 0)}칸 확인",
        f"- 멘트 {st.get('ment_lines', 0)}줄 · 수치 {st.get('ment_numbers_checked', 0)}개 대조",
    ]
    if not st.get("ment_numbers_checked"):
        lines.append("- ⚠ **멘트 미대조** — 대조한 수치가 0개입니다. 통과가 아니라 측정하지 못한 것입니다 (R-088)")
    if not st.get("tables"):
        lines.append("- ⚠ **값 미대조** — 통계표(tables.json)를 주지 않아 D5·D8·D15·D23 을 건너뛰었습니다")
    if st.get("fill_states"):
        lines.append("- 표 채움도 " + " · ".join(f"{k} {v}" for k, v in sorted(st["fill_states"].items())))
    if st.get("suppressed_known_defects"):
        lines.append(f"- 알려진 예시 결함 {st['suppressed_known_defects']}건을 '예시오류' 로 격하했습니다 (R-104)")
    return lines


def findings_json(rep: Report) -> str:
    """골든셋 하네스가 태그별로 집계할 수 있게 Finding 을 직렬화한다.

    `render()` 의 마크다운을 다시 파싱하게 두면 형식이 바뀔 때마다 하네스가 깨진다.
    """
    return json.dumps({
        "stats": rep.stats,
        "findings": [{"check": f.check, "severity": f.severity, "tag": f.tag,
                      "where": f.where, "detail": f.detail} for f in rep.findings],
    }, ensure_ascii=False, indent=1)


__all__ = [
    "verify_deck", "load_deck", "body_slides", "slide_panel",
    "norm_cell", "norm_label", "table_base_row", "table_value_columns",
    "table_header", "banner_axis", "iter_table_rows", "table_fill_state",
    "rendered_values", "expected_grid", "expected_top_text", "join_slides",
    "check_unfilled", "check_fill_state", "check_base", "check_data",
    "check_chart_vs_table", "check_decimals", "check_sums", "check_zero_notation",
    "check_sentinel_labels", "check_toptext", "check_order", "check_chart_type",
    "check_qu", "check_question_code_hygiene", "check_ment", "check_ment_pick",
    "check_label_clash", "check_deck_toc", "check_title_history",
    "suppress_known", "deck_header_lines", "findings_json",
]
