# -*- coding: utf-8 -*-
"""fill_deck.py -- 업무 B. 제작본 .pptx 를 훑고(--scan) 통계표 값으로 채운다(--fill).

제작팀(DP)은 양식이 같은 반복 장표를 **대표 1장만 만들고 나머지는 비워서** 보낸다.
그 빈칸을 채우는 게 업무 B 의 몸통이고, 이 파일이 그 두 단계를 맡는다.

    --scan   읽기 전용. "무엇을 채워야 하는가" 를 확정한다. 숫자를 읽지 않는다.
    --fill   통계표 값을 표·차트·Base/Unit·그래프 상단 Text 에 주입해 out/ 에 새로 쓴다.

★ 이 파일은 **숫자를 판단하지 않는다.** 반올림·정렬·수준어·Base 문자열·유의차는 전부
  rules.py 가 끝내고, XML 조작은 deckops.py 가, 줄바꿈은 textfit/kobreak 가 한다.
  fill_deck 은 "어느 값이 어느 칸으로 가는가" 라는 배선만 담당한다.

★ 원본을 덮어쓰지 않는다. deckops.Deck.save() 가 원본 경로를 거부한다.
★ 통계표에 없는 값은 **지어내지 않는다.** 비우고 리포트에 남긴다.
★ 제작팀 메모 도형은 지우지 않는다. 검수 후 사람이 지운다 (기본값; R-098 은 --fill 이
  아니라 납품 직전 단계의 일이라고 보고 이번 스펙에서 뺐다).

실측 근거 (두 예시 덱 전수 계측)
--------------------------------
R-105 표 y/h/bottom 고정, 데이터 행 높이 = 데이터영역 / n      -> deckops.set_row_count
R-106 차트 graphicFrame 은 보기 개수와 무관하게 고정            -> 아예 만지지 않는다
R-107 행 서식 [첫행][중간행][마지막행] 3종                      -> deckops.set_row_count
R-108 라벨 줄바꿈은 수식어구를 쪼개지 않는 자리                 -> textfit.wrap_ko
R-109 가로겹침막대는 order=0 에 긴 값(1+N순위)                  -> 계열 이름으로 매칭해 보존
R-110 0 표기(`-` / 공란)는 규칙이 아니라 담당자 습관            -> 클론 대표에서 상속
R-090 placeholder idx 로 sub_title/ment 를 가르면 안 된다       -> top_cm 으로 판정
R-094 계열의 개수·순서·이름은 제작본이 진리                     -> 이름으로 매칭, 값만 교체
R-095 표 열 수는 상수가 아니다 (실측 8~24열)                    -> (group,label) 문자열 매칭

--scan 이 표를 신호로 쓰는 이유 (실측)
--------------------------------------
· `Data 입력` 메모는 덱마다 있을 수도 없을 수도 있다. LG 64장 / SKEE 0장.
· 차트는 못 믿는다. 제작팀의 '빈 차트' 는 None 이 아니라 합이 100 이 되는 더미 상수다
  (50.0×2 · 25.0×4 · ±10.0×7). 심지어 대표 장표를 **값째로** 복사해 오기도 한다(LG s93).
· **표 값영역 채움도만 100% 맞는다.** LG: FILL 메모 64장 -> EMPTY 30 + SENTINEL 34,
  비-FILL 15장 -> 전부 FILLED, 불일치 0건. SKEE: 표 있는 본문 67장 전부 FILLED(오탐 0).

facts.json 어댑터 (이번 프로젝트는 통계표가 아직 없다)
------------------------------------------------------
--fill 의 값 출처는 원래 tables.json(parse_table_*.py 산출물)이다. 그런데 에스원 건은
통계표 입고 전이라 지금 당장 돌려 볼 수 없다. 그래서 **제작가이드 파싱 결과
(guide_blocks.json)도 같은 자리에 꽂을 수 있게** 어댑터를 뒀다 -- 가이드 블록의
`Base for %` 행 + 그 위 두 행(배너 그룹·라벨) + 아래 데이터 행이 통계표 표 한 장과
같은 모양이기 때문이다. 어느 쪽으로 읽었는지는 리포트 첫 줄에 찍힌다.

  ※ 값의 정본은 어디까지나 tables.json 이다 (21_멘트-보강 §4). guide_blocks 는 이미
    반올림·가공된 문자열이고 실측 오타가 있다(SKEE C163 `761.4%`). 통계표가 들어오면
    같은 CLI 에 tables.json 을 넘기면 되고, 그때 이 어댑터는 자동으로 비켜난다.

미결 (리포트에도 같이 찍는다)
-----------------------------
· 차트에 넣을 값이 표시값인가 원값인가. LG 실측이 갈린다 -- s38(척도)은 원값,
  s25(누적100%)는 표시값. 기본은 표시값(--chart-values display)으로 두고 플래그로 연다.
· PARTIAL 임계 0.6 은 두 덱에서 한 번도 발화하지 않은 미검증 값이다. 자동 채우기
  대상으로 올리지 않고 사람 확인 큐에 넣는다.
· 리터럴 `?` 셀(LG 660개/34장)이 제작팀 표준 표기인지 담당자 습관인지 미확정.

의존: lxml · openpyxl · Pillow(textfit) · PyYAML(chart_rules) -- 전부 이미 쓰던 것들
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:                     # 스크립트로 직접 실행될 때
    sys.path.insert(0, str(_HERE))

import deckops                                     # noqa: E402
import rules as R                                  # noqa: E402
import textfit                                     # noqa: E402
from extract_ground_truth import extract_deck      # noqa: E402
from parse_table_common import (                   # noqa: E402
    classify_row_label, col_letter, parse_title,
)

ENGINE = _HERE.parent
DEFAULT_OBSERVATIONS = ENGINE / "learning" / "observations.jsonl"

# ═══════════════════════════════════════════════════════════════════════════
# 0. 문자 정규화와 패턴
#
#    셀 문자열에 U+3000(전각 공백) 패딩이 붙는다. 실측 `-　`, `　`, `1~2인\n가구`.
#    strip() 으로는 안 지워지므로 비교·매칭 전에 반드시 norm_cell 을 통과시킨다.
# ═══════════════════════════════════════════════════════════════════════════

def norm_cell(s: Any) -> str:
    """전각공백·NBSP·개행을 공백으로 펴고 연속 공백을 하나로. 비교의 유일한 입구다."""
    if s is None:
        return ""
    t = (str(s).replace("　", " ").replace("\xa0", " ")
         .replace("\n", " ").replace("\r", " "))
    return " ".join(t.split()).strip()


def mask_digits(s: Any) -> str:
    """숫자를 `#` 로 가린다."""
    return re.sub(r"\d", "#", norm_cell(s))


def mask_cell(s: Any) -> str:
    """클론 패밀리 키에 넣을 셀 표기. **데이터가 키에 섞이는 것**을 막는다.

    값처럼 보이는 셀(숫자·`-`·`?`)은 통째로 `#` 하나로 접고, 나머지 글자는 숫자만 가린다.

    ★ 숫자 마스킹만으로는 부족하다. LG 척도 장표의 배너 라벨행은 extract 가
      TOP3 행을 잡아 오는데(헤더가 표 하단이라), 대표 s38 은 `TOP3(%) 58.0 75.0 …`
      이고 미기입 클론 s39 는 `TOP3(%) ? ? …` 다. 숫자만 가리면 `##.#` 과 `?` 가
      끝까지 달라 **대표가 자기 패밀리에서 떨어져 나가고 대표 없는 패밀리가 생긴다**
      (실측 F04/F05 · F07/F08 분리). 값 셀을 한 글자로 접으면 둘이 붙는다.
    """
    t = norm_cell(s)
    if not t:
        return ""
    if SENTINEL_CELL_RE.match(t) or _is_value_text(t):
        return "#"
    return mask_digits(t)


#: `(100)` `(8*)` 꼴 -- (Base) 행을 **값 패턴**으로 찾는 근거.
#: 첫 셀이 `(Base)` 인지로 찾으면 안 된다. LG ±세로막대 표는 헤더가 표 하단에 있고
#: base 행 첫 셀이 공란이라 extract_table 의 base_row_index 가 None 을 준다.
BASE_CELL_RE = re.compile(r"^\(\s*\d[\d,]*\s*\*?\s*\)$")

#: 표 셀의 미기입 센티널. **정확 일치**로만 본다 -- 정상 라벨에도 물음표가 섞인다
#: (`얼마나 만족하십니까?`). LG 실측 660셀 / 34장.
SENTINEL_CELL_RE = re.compile(r"^[?？]+$")

#: 텍스트 도형의 미기입 마커. 이쪽은 부분 일치로 본다.
#: rules.yaml V-070(4종)과 extract_ground_truth.UNFILLED_MARKERS(5종)의 합집합이다 --
#: 두 목록이 서로 다르므로 어느 한쪽만 믿으면 `Data 입력` 이나 `TBD` 를 놓친다.
TEXT_MARKERS = ("???", "내용을 입력하세요", "내용을입력하세요",
                "텍스트를 입력", "TBD", "XXX")

#: 그래프 상단 Text 행의 라벨. 매체마다 표기가 달라 전부 받아 준다.
#: 표 상단 행  : `TOP3(%)` `BOT2(%)` `Top2%` `7점 평균(점)` `5점 평균`
#: 독립 TextBox: `【TOP2】` `【BOT2】`
_TT_RE = re.compile(r"^(?:【\s*)?(TOP|BOT|SOSO)\s*(\d*)\s*(?:】)?\s*(?:\(\s*%?\s*\))?\s*%?$",
                    re.IGNORECASE)
_MEAN_RE = re.compile(r"^[\[【]?\s*(\d+)\s*점\s*평균\s*[\]】]?\s*(?:\(\s*점\s*\))?$")
_NPS_RE = re.compile(r"^[\[【]?\s*NPS\s*[\]】]?\s*(?:\(\s*점\s*\))?$", re.IGNORECASE)

#: 순수 숫자 셀(값)인지. 0 표기 `-` 도 값으로 친다 (R-110).
_NUMERIC_RE = re.compile(r"^-?\d[\d,]*(?:\.\d+)?%?$")


def toptext_kind(label: Any) -> str | None:
    """표 행 라벨 -> rules 의 row kind (`top`/`bot`/`soso`/`mean`). 아니면 None."""
    s = norm_cell(label)
    if not s:
        return None
    m = _TT_RE.match(s)
    if m:
        return {"TOP": "top", "BOT": "bot", "SOSO": "soso"}[m.group(1).upper()]
    if _MEAN_RE.match(s) or _NPS_RE.match(s):
        return "mean"
    return None


def _is_value_text(s: str) -> bool:
    """채워진 값처럼 보이는가. 0 표기 `-` 는 **채워진 값**이다 (SKEE 190개)."""
    return bool(s) and (s == "-" or bool(_NUMERIC_RE.match(s)))


# ═══════════════════════════════════════════════════════════════════════════
# 1. 제작팀 메모
# ═══════════════════════════════════════════════════════════════════════════

#: 메모 전문 -> 태그. 부분 문자열 판정이며 다중 태그를 허용한다.
_MEMO_TAGS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("FILL",      ("Data 입력",)),
    ("SPLIT",     ("장분리", "장 분리")),
    ("KEEPBLANK", ("범례만큼", "빈칸으로 유지")),
    ("SORT",      ("정렬", "순서")),
    ("REV",       ("송부",)),
)


def tag_memo(text: str) -> list[str]:
    """메모 전문 -> 태그 목록. 어디에도 안 걸리면 ['OTHER'].

    NOPT(보기 개수 지시)와 ASK(문의)는 조합 조건이라 따로 본다.
    학습 데이터이므로 OTHER 여도 버리지 않는다.
    """
    s = norm_cell(text)
    tags: list[str] = []
    for tag, needles in _MEMO_TAGS:
        if any(n in s for n in needles):
            tags.append(tag)
    if "보기" in s and ("개수" in s or "갯수" in s):
        tags.append("NOPT")
    if s.endswith("해주세요") or s.endswith("해 주세요"):
        tags.append("ASK")
    return tags or ["OTHER"]


def sweep_memos(slide_fact: dict, top_cm: float = 0.5) -> list[dict]:
    """슬라이드 팩트의 text_shapes 에서 제작팀 메모를 수집한다.

    조건은 (a) PLACEHOLDER 가 아님 (b) top <= top_cm (c) 텍스트 비어 있지 않음.

    extract_deck 의 production_memos 는 허용치 0.05cm · AUTO_SHAPE 한정이라 좁다 --
    LG s19 의 `선 색 #BFBFBF, 굵기 0.5` 가 top=0.41cm 라 통째로 빠진다(72 -> 73).
    0.5cm 로 넓혀도 두 덱 모두 오탐 0 이다. 그 구간의 유일한 다른 도형은 간지의
    `Slide Title` PLACEHOLDER(top=0.0)뿐이라 PLACEHOLDER 만 빼면 된다 (R-102 -- 남의
    파일을 고치지 않고 text_shapes 를 재훑는다).
    """
    out: list[dict] = []
    seen: set[tuple] = set()
    for sh in slide_fact.get("text_shapes") or []:
        top = sh.get("top_cm")
        if top is None or top > top_cm:
            continue
        if "PLACEHOLDER" in str(sh.get("shape_type") or "").upper():
            continue
        if sh.get("is_placeholder"):
            continue
        text = str(sh.get("text") or "")
        if not norm_cell(text):
            continue
        key = (norm_cell(text), round(float(sh.get("left_cm") or 0), 2))
        if key in seen:
            continue
        seen.add(key)
        out.append({
            "text": text,
            "lines": text.split("\n"),
            "tags": tag_memo(text),
            # extract_deck 의 text_shapes 는 이름 키가 `name` 이고 production_memos 는
            # `shape_name` 이다. 둘 다 받아야 학습 로그에 이름이 남는다.
            "shape_name": sh.get("shape_name") or sh.get("name"),
            "top_cm": top,
            "left_cm": sh.get("left_cm"),
        })
    # extract_deck 이 잡았는데 text_shapes 에 안 실린 것도 합친다 (그룹 안 도형 등)
    for memo in slide_fact.get("production_memos") or []:
        text = str(memo.get("text") or "")
        key = (norm_cell(text), round(float(memo.get("left_cm") or 0), 2))
        if not norm_cell(text) or key in seen:
            continue
        seen.add(key)
        out.append({
            "text": text,
            "lines": memo.get("lines") or text.split("\n"),
            "tags": tag_memo(text),
            "shape_name": memo.get("shape_name"),
            "top_cm": memo.get("top_cm"),
            "left_cm": memo.get("left_cm"),
        })
    out.sort(key=lambda d: (d.get("left_cm") or 0))
    return out


# ═══════════════════════════════════════════════════════════════════════════
# 2. 표 채움도 -- 미기입 판정의 유일한 신뢰 신호
# ═══════════════════════════════════════════════════════════════════════════

def table_base_row(table: dict) -> int | None:
    """(Base) 행 인덱스를 **값 패턴**으로 찾는다.

    `(NN)`/`(NN*)` 꼴 셀이 (n_cols-1)//2 개 이상인 행. extract_table 의
    base_row_index 는 첫 셀이 `(Base)` 일 때만 잡으므로 ±세로막대 표(첫 셀 공란·
    배너 헤더가 표 하단)에서 None 을 준다 -- 그 구멍을 메운다.
    """
    rows = table.get("rows") or []
    if not rows:
        return None
    need = max(2, (int(table.get("n_cols") or len(rows[0])) - 1) // 2)
    best, best_hits = None, 0
    for i, row in enumerate(rows):
        hits = sum(1 for c in row if BASE_CELL_RE.match(norm_cell(c)))
        if hits >= need and hits > best_hits:
            best, best_hits = i, hits
    return best


def table_value_columns(table: dict, base_row: int) -> list[int]:
    """(Base) 행에서 `(NN)` 패턴인 열만 값 열로 본다.

    라벨 열(0열 · 2단 라벨의 1열)과 공란 열은 제외된다.
    """
    rows = table.get("rows") or []
    if not (0 <= base_row < len(rows)):
        return []
    return [i for i, c in enumerate(rows[base_row]) if BASE_CELL_RE.match(norm_cell(c))]


@dataclass
class TableShape:
    """제작본 표 한 장의 구조 해석 결과. 인덱스가 아니라 **패턴**으로 잡은 것들이다."""
    base_row: int
    label_row: int
    group_row: int | None
    value_cols: list[int]
    data_rows: list[int]
    header_at_bottom: bool
    legend_rows: list[int] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return {
            "base_row": self.base_row, "label_row": self.label_row,
            "group_row": self.group_row, "value_cols": self.value_cols,
            "data_rows": self.data_rows, "header_at_bottom": self.header_at_bottom,
            "legend_rows": self.legend_rows,
        }


def table_shape(table: dict) -> TableShape | None:
    """표의 헤더 블록·값 열·데이터 행을 한 번에 판정한다. base 행을 못 찾으면 None.

    헤더가 위에 있는 표와 **아래에 있는 표**(척도 장표)를 같은 코드로 다룬다.
        표준 배너표   [그룹][라벨][(Base)][데이터…]
        척도 장표     [TOP3][평균][범례][라벨][(Base)][그룹]   <- 헤더가 하단, 역순
    라벨 행은 base 행에 붙은 두 후보 중 **글자(비숫자) 셀이 많은 쪽**으로 고른다.
    그룹 행은 남은 후보가 글자를 가지고 있으면 그쪽(척도형), 아니면 라벨 행 바깥쪽이다.
    """
    rows = table.get("rows") or []
    br = table_base_row(table)
    if br is None:
        return None
    vcols = table_value_columns(table, br)
    if not vcols:
        return None

    def labelness(i: int) -> int:
        if not (0 <= i < len(rows)):
            return -1
        row = rows[i]
        n = 0
        for c in vcols:
            s = norm_cell(row[c]) if c < len(row) else ""
            if s and not _is_value_text(s) and not SENTINEL_CELL_RE.match(s):
                n += 1
        return n

    cands = [i for i in (br - 1, br + 1) if 0 <= i < len(rows)]
    if not cands:
        return None
    label_row = max(cands, key=labelness)
    rest = [i for i in cands if i != label_row]
    if rest and labelness(rest[0]) > 0:
        group_row = rest[0]                 # 척도형: 그룹 행이 base 반대편에 있다
    else:
        step = label_row - br               # 라벨 행이 base 위면 -1, 아래면 +1
        gr = label_row + step
        group_row = gr if 0 <= gr < len(rows) else None

    header = {br, label_row}
    if group_row is not None:
        header.add(group_row)

    legend, data = [], []
    for i in range(len(rows)):
        if i in header:
            continue
        row = rows[i]
        empty_values = all(not norm_cell(row[c]) for c in vcols if c < len(row))
        if empty_values and norm_cell(row[0]).startswith("■"):
            legend.append(i)                # 척도 장표의 범례 셀(가로 병합)
            continue
        data.append(i)

    at_bottom = bool(header) and max(header) >= len(rows) - 1 and min(header) > 0
    return TableShape(br, label_row, group_row, vcols, data, at_bottom, legend)


def table_fill_state(table: dict, partial_blank_ratio: float = 0.6
                     ) -> tuple[str, int, int, int]:
    """(state, filled, blank, sentinel). state ∈ NOBASE|SENTINEL|EMPTY|PARTIAL|FILLED.

    · `?` 계열이 하나라도 있으면 SENTINEL  (LG 660셀 / 34장)
    · 숫자·`-` 가 0개면 EMPTY
    · blank 비율이 임계 이상이면 PARTIAL   (두 덱에서 발화 0건 -- 미검증)

    ★ 가로막대 장표는 표의 `전체` 열 데이터 칸이 **원래 비어 있다**(차트가 그 자리를
      덮는다). 그래서 열 하나가 통째로 비었다고 미기입으로 읽으면 안 된다 -- 비율로 본다.
    """
    shape = table_shape(table)
    if shape is None:
        return "NOBASE", 0, 0, 0
    rows = table.get("rows") or []
    filled = blank = sentinel = 0
    for i in shape.data_rows:
        row = rows[i]
        for c in shape.value_cols:
            s = norm_cell(row[c]) if c < len(row) else ""
            if SENTINEL_CELL_RE.match(s):
                sentinel += 1
            elif _is_value_text(s):
                filled += 1
            elif not s:
                blank += 1
            else:
                filled += 1               # 글자 값(예: `해당없음`)도 채워진 것으로 본다
    if sentinel:
        return "SENTINEL", filled, blank, sentinel
    total = filled + blank
    if total == 0:
        return "EMPTY", filled, blank, sentinel
    if filled == 0:
        return "EMPTY", filled, blank, sentinel
    if blank / total >= partial_blank_ratio:
        return "PARTIAL", filled, blank, sentinel
    return "FILLED", filled, blank, sentinel


def chart_fill_flags(chart: dict) -> set[str]:
    """차트에서 읽히는 미기입 흔적. **단독 판정 금지** -- 표를 봐야 한다.

    BLANKCAT  카테고리 라벨의 50% 이상이 공백        (오탐 0)
    ALLNONE   series_count>0 인데 real_series_count==0 (오탐 0)
    FLAT      비-None 값의 distinct <= 2             (SKEE 9건 오탐)

    제작팀의 '빈 차트' 는 None 이 아니라 합이 100 이 되는 더미 상수라(50.0/25.0/±10.0)
    차트만 보면 채워진 것처럼 보인다. 게다가 대표 장표를 값째로 복사해 오는 경우가
    있어(LG s93 = s90 의 값) 차트 기반 판정은 원리적으로 불가능하다.
    """
    flags: set[str] = set()
    cats = chart.get("categories") or []
    if cats and sum(1 for c in cats if not norm_cell(c)) / len(cats) >= 0.5:
        flags.add("BLANKCAT")
    if (chart.get("series_count") or 0) > 0 and not (chart.get("real_series_count") or 0):
        flags.add("ALLNONE")
    vals = [v for s in (chart.get("series") or []) for v in (s.get("values") or [])
            if v is not None]
    if vals and len({round(abs(float(v)), 6) for v in vals}) <= 2:
        flags.add("FLAT")
    return flags


_STATE_RANK = {"SENTINEL": 0, "EMPTY": 1, "PARTIAL": 2, "NOBASE": 3, "FILLED": 4}


def slide_fill_state(slide_fact: dict, *, partial_blank_ratio: float = 0.6
                     ) -> tuple[str, dict]:
    """슬라이드 1장의 최종 상태와 근거.

    표가 있으면 **표가 우선**(SENTINEL > EMPTY > PARTIAL > NOBASE > FILLED).
    표가 없고 차트만 있으면 ALLNONE/BLANKCAT 로 EMPTY 판정. 둘 다 없으면 NODATA --
    차트도 표도 없는 슬라이드에서 두 조건이 공허참이 되는 것을 막는 전제다.
    """
    tables = slide_fact.get("tables") or []
    charts = slide_fact.get("charts") or []
    ev: dict[str, Any] = {"tables": [], "chart_flags": []}
    if tables:
        best = "FILLED"
        for t in tables:
            st, f, b, s = table_fill_state(t, partial_blank_ratio)
            ev["tables"].append({"state": st, "filled": f, "blank": b, "sentinel": s})
            if _STATE_RANK[st] < _STATE_RANK[best]:
                best = st
        ev["by"] = "표 값영역 채움도"
        return best, ev
    if charts:
        flags: set[str] = set()
        for c in charts:
            fl = chart_fill_flags(c)
            ev["chart_flags"].append(sorted(fl))
            flags |= fl
        ev["by"] = "차트 플래그(표 없음)"
        if flags & {"ALLNONE", "BLANKCAT"}:
            return "EMPTY", ev
        return "FILLED", ev
    ev["by"] = "차트도 표도 없음"
    return "NODATA", ev


# ═══════════════════════════════════════════════════════════════════════════
# 3. 클론 패밀리 -- 메모가 0건인 덱에서도 형태로 반복 장표를 묶는다
# ═══════════════════════════════════════════════════════════════════════════

def clone_family_key(slide_fact: dict) -> tuple:
    """같은 양식의 반복 장표를 묶는 키.

    ★ 뺀 것과 그 근거 (전부 실측):
      (1) real_series_count -- 미기입 클론이면 0 이 된다(SKEE s6). 선언 계열수를 쓴다
      (2) data_label_num_fmts -- 대표 1장에만 계열별 override 가 걸려 있어 **대표가
          자기 패밀리에서 떨어져 나간다**(LG s38·s66, 대표없는 패밀리 2개 발생)
      (3) category_count / n_rows -- 보기 개수는 패밀리 안에서 달라지는 값이다
          (R-105·R-107 로 조정하는 대상). 넣으면 LG 18 -> 30 그룹 과분할
      (4) (Base) 행 값 -- 필터 문항은 base n 이 달라 같은 배너가 갈린다(SKEE 9 -> 34)
      (5) shape 이름 집합 -- `TextBox 15` 같은 자동 이름이 섞여 LG 18 -> 52 과분할
    배너 라벨 행은 숫자를 마스킹해서 넣는다. 안 하면 데이터가 키에 섞인다.
    """
    charts = tuple(
        (c.get("chart_type_id"), c.get("series_count"), c.get("overlap"),
         c.get("val_ax_min"), c.get("val_ax_max"))
        for c in (slide_fact.get("charts") or [])
    )
    tables = tuple(
        (t.get("n_cols"), tuple(mask_cell(x) for x in (t.get("banner_detail_row") or [])))
        for t in (slide_fact.get("tables") or [])
    )
    return (slide_fact.get("layout_name"), charts, tables)


def group_clone_families(body_slides: list[dict], *, partial_blank_ratio: float = 0.6
                         ) -> list[dict]:
    """clone_family_key 로 묶는다.

    차트·표가 모두 없는 슬라이드는 kind='nodata' 로 따로 뺀다 -- 빈 키가 서로 다
    같아져 한 덩어리로 뭉치는 것을 막는다(SKEE 텍스트 장표 17장).
    """
    buckets: dict[tuple, list[dict]] = {}
    nodata: list[dict] = []
    for s in body_slides:
        if not (s.get("charts") or s.get("tables")):
            nodata.append(s)
            continue
        buckets.setdefault(clone_family_key(s), []).append(s)

    fams: list[dict] = []
    for i, (key, members) in enumerate(sorted(buckets.items(),
                                              key=lambda kv: kv[1][0]["slide_no"]), 1):
        chart = members[0].get("charts") or []
        table = members[0].get("tables") or []
        fam = {
            "family_id": f"F{i:02d}",
            "kind": "chart+table" if (chart and table) else ("chart" if chart else "table"),
            "key": json.dumps(key, ensure_ascii=False, default=str),
            "layout": members[0].get("layout_name"),
            "chart_type": members[0].get("derived_chart_type"),
            "n_cols": table[0].get("n_cols") if table else None,
            "members": [m["slide_no"] for m in members],
            "states": {m["slide_no"]: slide_fill_state(m, partial_blank_ratio=partial_blank_ratio)[0]
                       for m in members},
        }
        fams.append(fam)
    if nodata:
        fams.append({
            "family_id": "F-ND", "kind": "nodata", "key": "", "layout": None,
            "chart_type": None, "n_cols": None,
            "members": [m["slide_no"] for m in nodata],
            "states": {m["slide_no"]: "NODATA" for m in nodata},
        })
    return fams


def pick_family_template(family: dict, *, memo_deck: bool,
                         memo_tags: dict[int, list[str]]) -> int | None:
    """패밀리 대표(=제작팀이 실제로 만든 1장) 슬라이드 번호. 후보 0개면 None.

    ① 메모 있는 덱: FILL 태그가 **없는** 멤버   ② 메모 없는 덱: 상태가 FILLED 인 멤버
    둘 다 slide_no 최소로 tie-break 한다. 후보가 2개 이상 나오는 패밀리가 실재하므로
    (LG s80·s81 -- 각각 `보기항목에 코드번호…`·`장분리` 메모라 FILL 태그가 없다)
    리포트에 후보 전량을 남긴다. 후보 0개면 자동 진행하지 않고 사람에게 넘긴다.
    """
    if family.get("kind") == "nodata":
        return None
    cands: list[int] = []
    for no in family["members"]:
        if memo_deck:
            if "FILL" not in memo_tags.get(no, []):
                cands.append(no)
        elif family["states"].get(no) == "FILLED":
            cands.append(no)
    if not cands:
        # 메모 덱이라도 메모가 틀릴 수 있다 -- 상태로 한 번 더 본다
        cands = [no for no in family["members"] if family["states"].get(no) == "FILLED"]
    return min(cands) if cands else None


# ═══════════════════════════════════════════════════════════════════════════
# 4. 미기입 잔존 · 중복 조인 키
# ═══════════════════════════════════════════════════════════════════════════

def find_unfilled_markers(slide_fact: dict) -> list[dict]:
    """미기입 잔존 전수.

    텍스트 도형은 **부분 문자열**, 표 셀은 **정확 일치** `^[?？]+$`.
    표 셀만 정확 일치인 이유는 정상 라벨에도 물음표가 섞이기 때문이고,
    텍스트 도형에 `?` 를 부분 일치로 넣으면 `얼마나 만족하십니까?` 가 전건 오탐이다.
    """
    out: list[dict] = []
    for sh in slide_fact.get("text_shapes") or []:
        text = str(sh.get("text") or "")
        hit = [m for m in TEXT_MARKERS if m in text]
        if hit:
            out.append({"where": "text", "shape_name": sh.get("shape_name"),
                        "markers": hit, "text": text[:80]})
    for ti, t in enumerate(slide_fact.get("tables") or []):
        for ri, row in enumerate(t.get("rows") or []):
            for ci, cell in enumerate(row):
                if SENTINEL_CELL_RE.match(norm_cell(cell)):
                    out.append({"where": f"표{ti + 1} r{ri}c{ci}", "markers": ["?"],
                                "text": norm_cell(cell)})
    return out


def duplicate_join_keys(body_slides: list[dict]) -> list[dict]:
    """같은 join_key 를 쓰는 슬라이드 묶음.

    슬라이드 번호가 **연속이면 장분리**(R-089), 떨어져 있으면 코드 중복(오타 의심).
    LG 실측: C2-1/C2-2/C2-3[81,82]·C3[83,84]·C3-1/…[85,86]·C4[87,88] = 장분리 /
    BB1-5[49,63] = 오타(63은 BD1-5 자리).
    ★ 장분리 2장째의 차트는 1장째와 **원래 같다** -- 복제 흔적이 아니다.
    """
    by: dict[str, list[int]] = {}
    for s in body_slides:
        key = s.get("join_key")
        if key:
            by.setdefault(str(key), []).append(s["slide_no"])
    out = []
    for key, nos in sorted(by.items(), key=lambda kv: kv[1][0]):
        if len(nos) < 2:
            continue
        nos = sorted(nos)
        consecutive = all(b - a == 1 for a, b in zip(nos, nos[1:]))
        out.append({"join_key": key, "slides": nos,
                    "verdict": "장분리(R-089)" if consecutive else "코드중복(오타 의심)"})
    return out


# ═══════════════════════════════════════════════════════════════════════════
# 5. --scan
# ═══════════════════════════════════════════════════════════════════════════

def scan_deck(pptx: Path, *, memo_top_cm: float = 0.5,
              partial_blank_ratio: float = 0.6) -> dict:
    """덱 1벌 스캔. 파일을 쓰지 않는다(순수함수).

    extract_deck() 결과 위에 memos/families/fill_state/targets/duplicates/unfilled 를
    얹은 dict 를 돌려준다. 통계표·제작가이드는 **읽지 않는다** -- 매핑과 값 주입은
    --fill 의 일이다.
    """
    doc = extract_deck(pptx)
    slides = doc["slides"]
    body = [s for s in slides if s.get("classification") == "본문"]

    memos: dict[int, list[dict]] = {}
    memo_tags: dict[int, list[str]] = {}
    for s in slides:
        found = sweep_memos(s, memo_top_cm)
        if found:
            memos[s["slide_no"]] = found
            tags: list[str] = []
            for m in found:
                tags.extend(m["tags"])
            memo_tags[s["slide_no"]] = sorted(set(tags))
    memo_deck = any("FILL" in t for t in memo_tags.values())

    states: dict[int, dict] = {}
    for s in body:
        st, ev = slide_fill_state(s, partial_blank_ratio=partial_blank_ratio)
        states[s["slide_no"]] = {"state": st, "evidence": ev}

    families = group_clone_families(body, partial_blank_ratio=partial_blank_ratio)
    for fam in families:
        fam["template"] = pick_family_template(fam, memo_deck=memo_deck, memo_tags=memo_tags)
        fam["targets"] = [n for n in fam["members"]
                          if fam["states"].get(n) in ("EMPTY", "SENTINEL")]
        fam["variants"] = sorted({fam["states"].get(n) for n in fam["members"]})

    by_no = {s["slide_no"]: s for s in body}
    targets = []
    for s in body:
        no = s["slide_no"]
        st = states[no]["state"]
        memo_fill = "FILL" in memo_tags.get(no, [])
        if st not in ("EMPTY", "SENTINEL") and not memo_fill:
            continue
        fam = next((f for f in families if no in f["members"]), None)
        targets.append({
            "slide_no": no,
            "join_key": s.get("join_key"),
            "join_key_kind": s.get("join_key_kind"),
            "question_code": s.get("question_code"),
            "state": st,
            "memo_fill": memo_fill,
            "agree": (st in ("EMPTY", "SENTINEL")) == memo_fill if memo_deck else None,
            "family": fam["family_id"] if fam else None,
            "template": fam.get("template") if fam else None,
            "chart_type": s.get("derived_chart_type"),
            "reason": states[no]["evidence"].get("by"),
        })

    unfilled = []
    for s in slides:
        found = find_unfilled_markers(s)
        if found:
            unfilled.append({"slide_no": s["slide_no"], "count": len(found),
                             "items": found[:6]})

    dups = duplicate_join_keys(body)

    ask = [{"slide_no": no, "tags": memo_tags[no],
            "text": memos[no][0]["text"]} for no in sorted(memo_tags)
           if "ASK" in memo_tags[no]]
    no_template = [f["family_id"] for f in families
                   if f["kind"] != "nodata" and f["template"] is None]
    nobase = [s["slide_no"] for s in body
              if any(table_fill_state(t)[0] == "NOBASE" for t in (s.get("tables") or []))]
    partial = [no for no, v in states.items() if v["state"] == "PARTIAL"]

    ment_empty = [s["slide_no"] for s in body if not norm_cell(s.get("top_ment"))]

    return {
        "source_file": doc.get("source_file"),
        "source_path": str(pptx),
        "slide_count": doc.get("slide_count"),
        "slide_width_cm": doc.get("slide_width_cm"),
        "slide_height_cm": doc.get("slide_height_cm"),
        "by_classification": (doc.get("summary") or {}).get("by_classification", {}),
        "by_layout": (doc.get("summary") or {}).get("by_layout", {}),
        "body_slide_count": len(body),
        "join_key_kinds": _count([s.get("join_key_kind") for s in body]),
        "memo_deck": memo_deck,
        "memo_slide_count": len(memos),
        "memo_total": sum(len(v) for v in memos.values()),
        "memos": {str(k): v for k, v in sorted(memos.items())},
        "memo_tags": {str(k): v for k, v in sorted(memo_tags.items())},
        "fill_state": {str(no): v for no, v in sorted(states.items())},
        "fill_state_counts": _count([v["state"] for v in states.values()]),
        "families": families,
        "targets": targets,
        "duplicates": dups,
        "unfilled": unfilled,
        "ment_empty": ment_empty,
        "ment_empty_count": len(ment_empty),
        "needs_human": {
            "families_without_template": no_template,
            "partial_slides": sorted(partial),
            "nobase_slides": sorted(nobase),
            "ask_memos": ask,
            "duplicate_codes": [d for d in dups if "오타" in d["verdict"]],
            "disagreements": [t["slide_no"] for t in targets if t.get("agree") is False],
        },
        "_body_slides": [s["slide_no"] for s in body],
        "_doc": doc,          # --fill 이 재사용한다. JSON 직렬화 때는 뺀다
    }


def _count(items: Iterable) -> dict[str, int]:
    out: dict[str, int] = {}
    for i in items:
        k = str(i)
        out[k] = out.get(k, 0) + 1
    return dict(sorted(out.items(), key=lambda kv: (-kv[1], kv[0])))


def render_scan_report(scan: dict) -> str:
    """사람이 읽을 마크다운. 콘솔에도 같은 내용을 찍는다."""
    L: list[str] = []
    a = L.append
    a(f"# 제작본 스캔 — {scan['source_file']}")
    a("")
    a("## ① 덱 요약")
    a(f"- 총 {scan['slide_count']}장 / 본문 {scan['body_slide_count']}장 "
      f"/ 규격 {scan['slide_width_cm']}×{scan['slide_height_cm']}cm")
    a(f"- 분류별: " + ", ".join(f"{k} {v}" for k, v in scan["by_classification"].items()))
    a(f"- 레이아웃별: " + ", ".join(f"{k} {v}" for k, v in scan["by_layout"].items()))
    a(f"- 조인 키: " + ", ".join(f"{k} {v}" for k, v in scan["join_key_kinds"].items()))
    a(f"- 채움 상태: " + ", ".join(f"{k} {v}" for k, v in scan["fill_state_counts"].items()))
    a(f"- 멘트 빈 본문 {scan['ment_empty_count']}/{scan['body_slide_count']}장")
    a("")
    a("## ② 제작팀 메모")
    if not scan["memos"]:
        a("- 메모 0건 (DP-003). 클론 패밀리와 표 채움도만으로 판정했다.")
    else:
        a(f"- 총 {scan['memo_total']}건 / {scan['memo_slide_count']}장 "
          f"(덱 유형: {'메모 있음' if scan['memo_deck'] else '메모 없음'})")
        for no, items in scan["memos"].items():
            for m in items:
                head = " / ".join(m["tags"])
                body = " ⏎ ".join(x for x in m["lines"] if x.strip())
                a(f"  - s{no} [{head}] {body}")
    a("")
    a("## ③ 클론 패밀리")
    a("| ID | 멤버 | 대표 | 차트타입 | 표 열수 | 상태 |")
    a("|---|---|---|---|---|---|")
    for f in scan["families"]:
        mem = ",".join(f"s{n}" for n in f["members"][:12])
        if len(f["members"]) > 12:
            mem += f" …(+{len(f['members']) - 12})"
        tmpl = f"s{f['template']}" if f.get("template") else "**없음**"
        a(f"| {f['family_id']} | {mem} | {tmpl} | {f.get('chart_type') or '-'} "
          f"| {f.get('n_cols') or '-'} | {'/'.join(str(v) for v in f['variants'])} |")
    a("")
    a("## ④ 채우기 대상")
    states = _count([t["state"] for t in scan["targets"]])
    a(f"- {len(scan['targets'])}장 (" + ", ".join(f"{k} {v}" for k, v in states.items()) + ")")
    done = [t["slide_no"] for t in scan["targets"] if t["state"] == "FILLED"]
    if done:
        a(f"- ⚠ 이 중 {len(done)}장은 표가 **이미 채워져 있는데** `Data 입력` 메모가 남아 있다 "
          f"— 값 주입은 끝났고 메모만 지우면 된다(검수 후 사람이 지운다): "
          f"{', '.join('s%d' % n for n in done[:20])}"
          + (f" …(+{len(done) - 20})" if len(done) > 20 else ""))
    a("| 슬라이드 | 조인키 | 상태 | 메모 | 패밀리(대표) | 근거 |")
    a("|---|---|---|---|---|---|")
    for t in scan["targets"]:
        a(f"| s{t['slide_no']} | {t.get('join_key') or '-'} | {t['state']} "
          f"| {'FILL' if t['memo_fill'] else '-'} "
          f"| {t.get('family')}({'s%s' % t['template'] if t.get('template') else '없음'}) "
          f"| {t.get('reason')} |")
    a("")
    a("## ⑤ 중복 join_key")
    if not scan["duplicates"]:
        a("- 없음")
    for d in scan["duplicates"]:
        a(f"- `{d['join_key']}` → {', '.join('s%d' % n for n in d['slides'])} — {d['verdict']}")
    a("")
    a("## ⑥ 미기입 잔존")
    if not scan["unfilled"]:
        a("- 없음")
    for u in scan["unfilled"]:
        sample = "; ".join(f"{i['where']}:{'/'.join(i['markers'])}" for i in u["items"])
        a(f"- s{u['slide_no']} {u['count']}건 — {sample}")
    a("")
    a("## ⑦ 사람 판단 필요")
    nh = scan["needs_human"]
    a(f"- 대표 미확정 패밀리: {nh['families_without_template'] or '없음'}")
    a(f"- PARTIAL(미검증 임계): {nh['partial_slides'] or '없음'}")
    a(f"- (Base) 행 미검출 표: {nh['nobase_slides'] or '없음'}")
    a(f"- 코드중복(오타 의심): "
      f"{[d['join_key'] for d in nh['duplicate_codes']] or '없음'}")
    a(f"- 메모/표 판정 불일치: {nh['disagreements'] or '없음'}")
    if nh["ask_memos"]:
        for m in nh["ask_memos"]:
            a(f"- s{m['slide_no']} 문의 메모: {m['text'][:70]}")
    return "\n".join(L)


def write_observations(scan: dict, path: Path) -> int:
    """메모 전량을 observations.jsonl 에 append. 기존 스키마(ts/kind)를 따른다.

    학습 데이터이므로 태그가 OTHER 여도 버리지 않는다.
    """
    if not scan["memos"]:
        return 0
    ts = datetime.now(timezone.utc).isoformat(timespec="seconds")
    path.parent.mkdir(parents=True, exist_ok=True)
    n = 0
    with path.open("a", encoding="utf-8") as fh:
        for no, items in scan["memos"].items():
            for m in items:
                fh.write(json.dumps({
                    "ts": ts, "tool": "fill_deck", "kind": "dp_memo",
                    "source": scan["source_file"], "slide_no": int(no),
                    "text": m["text"], "lines": m["lines"], "tags": m["tags"],
                    "shape_name": m.get("shape_name"), "top_cm": m.get("top_cm"),
                }, ensure_ascii=False) + "\n")
                n += 1
    return n


def scan_json(scan: dict) -> dict:
    """--json 으로 내보낼 형태. 원본 extract_deck 문서는 뺀다(수 MB)."""
    return {k: v for k, v in scan.items() if not k.startswith("_")}


# ═══════════════════════════════════════════════════════════════════════════
# 6. facts.json 어댑터 -- tables.json 도 guide_blocks.json 도 같은 자리에 꽂는다
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class FactBook:
    """값의 출처 한 벌. 어느 파일에서 왔는지를 끝까지 들고 다닌다."""
    source: str                       # 'tables.json' | 'guide_blocks.json'
    path: Path
    tables: list[dict]
    by_code: dict[str, list[dict]] = field(default_factory=dict)
    by_join: dict[str, list[dict]] = field(default_factory=dict)
    by_title: dict[str, list[dict]] = field(default_factory=dict)
    by_sub: dict[str, list[dict]] = field(default_factory=dict)

    def index(self) -> "FactBook":
        for t in self.tables:
            code = norm_cell(t.get("question_code"))
            if code:
                self.by_code.setdefault(code, []).append(t)
            jk = norm_cell(t.get("join_key"))
            if jk:
                self.by_join.setdefault(jk, []).append(t)
            title = _join_title(t.get("slide_title"), t.get("sub_title"))
            if title:
                self.by_title.setdefault(title, []).append(t)
            sub = norm_cell(t.get("sub_title"))
            if sub:
                self.by_sub.setdefault(sub, []).append(t)
        for bucket in (self.by_code, self.by_join, self.by_title, self.by_sub):
            for k in bucket:
                bucket[k].sort(key=_table_rank)
        return self


def _table_rank(t: dict) -> tuple:
    """같은 키를 두 표가 주장할 때의 우선순위.

    요약 장표의 곁다리 패널이 본문 전용 표를 밀어내는 것을 막는다 -- 실측 A12 는
    요약 블록7의 패널(배너 2열)과 본문 블록31의 표(배너 16열)가 같은 코드를 갖는데,
    등록 순서대로면 요약 쪽이 이겨서 s32 의 배너가 통째로 안 붙는다.
    **패널이 하나뿐인 블록 > 배너가 많은 표** 순으로 본다.
    """
    return (1 if (t.get("_panel_count") or 1) > 1 else 0,
            -len(t.get("banners") or []),
            str(t.get("index")))


def _join_title(title: Any, sub: Any) -> str:
    t, s = norm_cell(title), norm_cell(sub)
    if not (t or s):
        return ""
    return f"{t.replace(' ', '')}|{s.replace(' ', '')}"


def load_facts(path: Path) -> FactBook:
    """facts.json 을 읽는다. tables.json / guide_blocks.json 을 자동 판별한다."""
    doc = json.loads(Path(path).read_text(encoding="utf-8"))
    if isinstance(doc, dict) and isinstance(doc.get("tables"), list):
        tables = doc["tables"]
        for t in tables:
            t.setdefault("_facts_source", "tables.json")
        return FactBook("tables.json", Path(path), tables).index()
    if isinstance(doc, dict) and isinstance(doc.get("blocks"), list):
        return FactBook("guide_blocks.json", Path(path),
                        adapt_guide_blocks(doc)).index()
    raise ValueError(f"facts.json 형태를 모르겠다 (tables/blocks 키가 없다): {path}")


def adapt_guide_blocks(doc: dict) -> list[dict]:
    """제작가이드 파싱 결과 -> tables.json 과 같은 모양의 표 목록.

    ★ 이건 **임시 다리**다. 값의 정본은 tables.json 이고(21_멘트-보강 §4), 가이드 값은
      이미 반올림·가공된 문자열이라 실측 오타가 섞인다(SKEE C163 `761.4%`).
      통계표가 들어오면 같은 CLI 에 tables.json 을 넘기면 이 함수는 비켜난다.

    가이드 블록의 데이터 그리드는 통계표 표와 같은 골격이다:
        [배너 그룹행][배너 라벨행][Base for % 행][보기 행…][계][【TOP】][…]
    `Base for %` 행을 `(NN)` 패턴으로 찾고 그 위 두 행을 헤더로 삼는다 -- 라벨이
    `Base for %` 가 아닌 프로젝트가 있어도 패턴 쪽이 먼저 걸린다 (V-090).
    """
    out: list[dict] = []
    for b in doc.get("blocks") or []:
        out.extend(_tables_from_block(b))
    for b in doc.get("orphan_tables") or []:
        out.extend(_tables_from_block(b))
    return out


def _tables_from_block(block: dict) -> list[dict]:
    """가이드 블록 1개 -> 표 0~N장.

    패널은 **세로로도 가로로도** 늘어난다. 요약 장표는 한 행에 `Base for %` 블록이
    두셋 나란히 서고(LG 블록7 실측 B|C · I|J,K · N|O), 본문 장표는 세로로 하나뿐이다.
    그래서 base 행을 세로로 훑고, 각 base 행 안에서 **라벨 셀을 경계로 가로로 쪼갠다.**
    이렇게 안 하면 요약 블록의 서로 다른 문항 세 개가 한 표로 뭉쳐 엉뚱한 배너를 갖게 되고,
    그 표가 문항코드 색인을 선점해 본문 매핑을 통째로 오염시킨다(실측 A12·C1).
    """
    body = block.get("body_rows") or []
    if not body:
        return []
    base_idx = []
    for i, r in enumerate(body):
        hits = sum(1 for v in (r.get("cells") or {}).values()
                   if BASE_CELL_RE.match(norm_cell(v)))
        if hits >= 1:
            base_idx.append(i)
    if not base_idx:
        return []

    out: list[dict] = []
    seq = 0
    for p, bi in enumerate(base_idx):
        end = base_idx[p + 1] - 3 if p + 1 < len(base_idx) else len(body)
        for label_col, vcols in _split_panels(body[bi].get("cells") or {}):
            t = _table_from_panel(block, body, bi, max(bi + 1, end),
                                  label_col, vcols, seq)
            if t is not None:
                out.append(t)
            seq += 1
    return out


def _split_panels(base_cells: dict) -> list[tuple[str, list[str]]]:
    """base 행을 (라벨 열, 값 열들) 묶음으로 쪼갠다. 라벨 셀이 경계다."""
    cols = sorted(base_cells, key=deckops.col_to_index)
    out: list[tuple[str, list[str]]] = []
    label: str | None = None
    vals: list[str] = []
    for c in cols:
        v = norm_cell(base_cells[c])
        if BASE_CELL_RE.match(v):
            vals.append(c)
        elif v:
            if vals:
                out.append((label or col_letter(max(0, deckops.col_to_index(vals[0]) - 1)),
                            vals))
            label, vals = c, []
    if vals:
        out.append((label or col_letter(max(0, deckops.col_to_index(vals[0]) - 1)), vals))
    return out


def _table_from_panel(block: dict, body: list[dict], bi: int, end: int,
                      label_col: str, vcols: list[str], panel: int) -> dict | None:
    base_cells = body[bi].get("cells") or {}
    if not vcols:
        return None

    lbl_row = (body[bi - 1].get("cells") or {}) if bi >= 1 else {}
    grp_row = (body[bi - 2].get("cells") or {}) if bi >= 2 else {}

    banners: list[dict] = []
    cur_group = ""
    for pos, c in enumerate(vcols):
        g = norm_cell(grp_row.get(c))
        if g:
            cur_group = g
        label = norm_cell(lbl_row.get(c)) or cur_group
        m = re.search(r"\d[\d,]*", norm_cell(base_cells.get(c)))
        n = int(m.group().replace(",", "")) if m else 0
        banners.append({"pos": pos, "col": deckops.col_to_index(c), "col_letter": c,
                        "group": cur_group or label, "label": label, "n": n,
                        "in_report": True})

    rows: list[dict] = []
    pos = 0
    for r in body[bi + 1:end]:
        cells = r.get("cells") or {}
        raw_label = cells.get(label_col)
        has_values = any(c in cells and cells[c] not in (None, "") for c in vcols)
        if not has_values and not norm_cell(raw_label):
            continue
        info = classify_row_label(raw_label)
        values: list[dict | None] = []
        for b in banners:
            v = cells.get(b["col_letter"])
            if v is None or v == "":
                values.append(None)
            elif isinstance(v, str):
                values.append({"raw": None, "display": v.strip(), "text": v.strip()})
            else:
                fv = float(v)
                digits = 2 if info.kind == "mean" else 1
                values.append({"raw": fv, "value": R.round_half_up(fv, digits),
                               "display": f"{R.round_half_up(fv, digits):.{digits}f}",
                               "fmt": None})
        rows.append({
            "pos": pos, "row": r.get("row"),
            "label_raw": raw_label if isinstance(raw_label, str) else None,
            "label": info.text, "label_group": None, "indent": info.indent,
            "kind": info.kind, "option_no": info.option_no,
            "option_text": info.option_text, "scale_points": info.scale_points,
            "mean_unit": info.mean_unit, "values": values,
        })
        pos += 1
    if not rows:
        return None
    _rescue_option_rows(rows)

    title, title_text = _panel_title(block, label_col, body[bi].get("row"))
    ti = parse_title(title_text or "")
    kinds = [r["kind"] for r in rows]
    scale_pts = next((r["scale_points"] for r in rows if r["scale_points"]), None)
    has_scale = any(k in ("bot", "soso", "top") for k in kinds)

    code = (title or {}).get("code") or block.get("question_code") or ti.question_code

    return {
        "index": f"{block.get('block_index')}.{panel}",
        "_facts_source": "guide_blocks.json",
        "question_code": code or ti.question_code,
        "title_full": title_text,
        "title_base": ti.title_base,
        "suffix": ti.suffix,
        "suffix_parts": ti.suffix_parts,
        "matrix_item_no": ti.matrix_item_no,
        "base_n": block.get("base_n") or (banners[0]["n"] if banners else None),
        "banners": banners,
        "rows": rows,
        "hints": {
            "has_scale_rows": has_scale,
            "scale_points": scale_pts or (5 if has_scale else None),
            "is_ranking": ti.is_ranking, "rank_scope": ti.rank_scope,
            "rank_is_first_only": ti.rank_is_first_only,
            "is_multi": ti.is_multi, "is_summary": ti.summary_kind is not None,
            "summary_kind": ti.summary_kind,
            "is_matrix_item": ti.matrix_item_no is not None,
            "has_total_row": "total" in kinds, "has_mean_row": "mean" in kinds,
            "n_option_rows": sum(1 for k in kinds if k == "option"),
            "n_rows": len(rows), "n_banners": len(banners),
            "n_banners_in_report": len(banners),
        },
        # 가이드에서만 오는 부가 정보 -- 조인 폴백과 멘트 재사용에만 쓴다(값 출처 아님)
        "join_key": block.get("join_key"),
        "join_key_kind": block.get("join_key_kind"),
        "slide_title": block.get("slide_title"),
        "sub_title": block.get("sub_title"),
        "guide_chart_type": block.get("chart_type"),
        "guide_base_unit": block.get("base_unit"),
        "guide_ments": block.get("ments") or [],
        # 같은 문항코드를 두 표가 주장할 때의 우선순위 재료 (요약 패널 < 본문 전용 표)
        "_panel_count": block.get("panel_count") or 1,
        "_block_index": block.get("block_index"),
    }


def _panel_title(block: dict, label_col: str, base_row: Any) -> tuple[dict | None, str]:
    """패널 제목을 **열로** 찾는다. 패널 순번으로 집으면 요약 블록에서 어긋난다.

    LG 블록7 실측: panel_titles 가 [B,I,N](139행) + [B,I](159행) 5개인데 base 행은
    두 개다. 순번으로 집으면 두 번째 base 행의 B 패널이 A14 제목을 물어 온다.
    """
    titles = [t for t in (block.get("panel_titles") or [])
              if norm_cell(t.get("col")) == norm_cell(label_col)]
    if titles and base_row is not None:
        above = [t for t in titles if (t.get("row") or 0) <= base_row]
        if above:
            best = max(above, key=lambda t: t.get("row") or 0)
            return best, norm_cell(best.get("text"))
    if titles:
        return titles[0], norm_cell(titles[0].get("text"))
    return None, norm_cell(block.get("question_reference"))


def _rescue_option_rows(rows: list[dict]) -> None:
    """`1. 라벨` 꼴 보기를 보기 행으로 되살린다.

    classify_row_label 은 `N) ` 접두만 보기번호로 읽는다. 그런데 가이드에는 마침표형
    (`1. 기능형 어르신 물건 이동 보조`)으로 적힌 문항군이 있고(LG C3·C4 실측),
    그러면 option 행이 0개가 되어 **정렬도 주입도 통째로 빈손이 된다.**
    보기 행이 하나도 없을 때만, 값을 가진 일반 행을 보기로 승격한다.
    """
    if any(r["kind"] == "option" for r in rows):
        return
    for r in rows:
        if r["kind"] != "other":
            continue
        label = norm_cell(r.get("label"))
        if not label or toptext_kind(label) or label.lower().startswith("base"):
            continue
        if not any(v is not None for v in (r.get("values") or [])):
            continue
        r["kind"] = "option"
        m = re.match(r"^\s*(\d+)\s*[.)]", label)
        if m:
            r["option_no"] = int(m.group(1))


# ═══════════════════════════════════════════════════════════════════════════
# 7. 배너축 -- 제작본 표 열 ↔ 통계표 배너
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class BannerAxis:
    """제작본 표 한 장의 배너 구조 전부. 프로젝트 상수(BANNER_COLS/BASE_N/…)를 대체한다."""
    cols: list[int]                       # 표 열 인덱스
    group: list[str]
    label: list[str]
    n: list[int]
    small: list[bool]                     # `*` = n<30 표식
    total_col: int | None
    shape: TableShape

    def key(self, i: int) -> tuple[str, str]:
        return (self.group[i], self.label[i])


def banner_axis(deck_table: dict) -> BannerAxis | None:
    """제작본 표에서 배너 헤더 블록을 복원한다. 헤더가 표 하단인 척도 장표도 같이 잡힌다.

    그룹은 병합셀이라 forward-fill 하고, 라벨이 비면 그룹으로 떨어뜨린다
    (parse_table_common.build_banners 와 같은 규칙). `*` 는 n<30 표식이라 떼고 읽는다.
    """
    shape = table_shape(deck_table)
    if shape is None:
        return None
    rows = deck_table.get("rows") or []
    lbl = rows[shape.label_row]
    grp = rows[shape.group_row] if shape.group_row is not None else []

    cols, groups, labels, ns, smalls = [], [], [], [], []
    cur = ""
    ncols = int(deck_table.get("n_cols") or len(lbl))
    for c in range(ncols):
        g = norm_cell(grp[c]) if c < len(grp) else ""
        if g:
            cur = g
        if c not in shape.value_cols:
            continue
        label = (norm_cell(lbl[c]) if c < len(lbl) else "") or cur
        raw = norm_cell(rows[shape.base_row][c])
        m = re.search(r"\d[\d,]*", raw)
        cols.append(c)
        groups.append(cur or label)
        labels.append(label)
        ns.append(int(m.group().replace(",", "")) if m else 0)
        smalls.append("*" in raw)

    total_col = None
    for i, c in enumerate(cols):
        if groups[i] == "전체" or labels[i] == "전체":
            total_col = c
            break
    return BannerAxis(cols, groups, labels, ns, smalls, total_col, shape)


def bind_banners(axis: BannerAxis, table: dict) -> list[int | None]:
    """제작본 표 열 i ↔ 통계표 banners[pos]. 못 붙은 열은 None.

    키는 (group,label) 을 정규화한 값이다. 열 수를 상수로 가정하지 않는다 (R-095 --
    실측 8/9/11/12/15/16/17/24열). 위치로 맞추면 표마다 다른 데를 자르게 된다.
    """
    banners = table.get("banners") or []
    exact: dict[tuple[str, str], int] = {}
    by_label: dict[str, list[int]] = {}
    for b in banners:
        g, l = norm_cell(b.get("group")), norm_cell(b.get("label"))
        exact.setdefault((g, l), b["pos"])
        by_label.setdefault(l, []).append(b["pos"])

    out: list[int | None] = []
    for i in range(len(axis.cols)):
        g, l = axis.key(i)
        pos = exact.get((g, l))
        if pos is None:
            cand = by_label.get(l) or []
            pos = cand[0] if len(cand) == 1 else None
        out.append(pos)
    return out


# ═══════════════════════════════════════════════════════════════════════════
# 8. 매핑 -- 조인 키 사다리
# ═══════════════════════════════════════════════════════════════════════════

def map_slide(slide: dict, book: FactBook) -> tuple[dict | None, str, str]:
    """슬라이드 -> 통계표 표. (표, 방법, 설명).

    사다리: ①question_code ②slide_title+sub_title 정규화 ③sub_title 단독이 유일할 때
    ④Qu 텍스트. SKEE 실측으로 검증 -- 본문 84장 중 exact 57 + sub 단독 유일 11 =
    데이터를 가진 68장 100%. exact 가 깨지는 이유는 제작팀의 제목 변경이다 (R-093).
    실패는 MAP 으로 집계하고 **그 슬라이드는 건드리지 않는다.**
    """
    code = norm_cell(slide.get("question_code"))
    if code and code in book.by_code:
        cands = book.by_code[code]
        return cands[0], "question_code", f"문항코드 `{code}` ({len(cands)}장 후보)"

    title = _join_title(slide.get("slide_title"), slide.get("sub_title"))
    if title and title in book.by_title:
        cands = book.by_title[title]
        return cands[0], "title+subtitle", f"제목+서브제목 ({len(cands)}장 후보)"

    jk = norm_cell(slide.get("join_key"))
    if jk and jk in book.by_join:
        cands = book.by_join[jk]
        return cands[0], "join_key", f"조인키 `{jk}` ({len(cands)}장 후보)"

    sub = norm_cell(slide.get("sub_title"))
    if sub and len(book.by_sub.get(sub, [])) == 1:
        return book.by_sub[sub][0], "sub_title", f"서브제목 단독 유일 `{sub}`"

    qu = norm_cell(slide.get("qu_text")).replace("Qu:", "").strip()
    if qu:
        for t in book.tables:
            if norm_cell(t.get("title_base")) and norm_cell(t.get("title_base")) in qu:
                return t, "qu_text", "Qu 텍스트 부분일치"
    return None, "none", "사다리 4단 전부 실패"


# ═══════════════════════════════════════════════════════════════════════════
# 9. 패널 계획 -- "어느 값이 어느 칸으로 가는가"
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class RowPlan:
    label: str
    lines: list[str] | None                     # 수동 줄바꿈 (R-108). None 이면 한 줄
    cells: dict[int, str]                       # 표 열 인덱스 -> 텍스트
    kind: str = "option"


@dataclass
class SlidePlan:
    kind: str                                   # bar1 | rank2 | stack100 | scale | table
    data_rows: list[RowPlan] = field(default_factory=list)
    chart_series: list[list[float | None]] | None = None
    chart_categories: list[str] | None = None
    base_unit: str | None = None
    footnote: str | None = None
    top_text: list[tuple[str, str]] = field(default_factory=list)   # (【TOP2】, '60.0 %')
    ments: list[str] = field(default_factory=list)
    tags: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def panel_kind(slide: dict, axis: BannerAxis | None) -> str:
    """제작본이 실제로 어떤 장표인가. **덱이 진리다** (R-094) -- 가이드 문자열이 아니라.

    표 구조를 먼저 본다. 헤더가 하단이고 데이터 행이 전부 TOP/평균 라벨이면 척도다 --
    LG=COLUMN_STACKED(52) / SKEE=COLUMN_STACKED_100(53) 로 차트 종류가 덱마다 달라
    차트 id 로 가르면 SKEE 에서 누적100%와 뒤섞인다.
    """
    tables = slide.get("tables") or []
    charts = slide.get("charts") or []
    if axis is not None and tables:
        rows = tables[0].get("rows") or []
        idx = axis.shape.data_rows
        if axis.shape.header_at_bottom and idx and \
                all(toptext_kind(rows[i][0]) for i in idx):
            return "scale"
    if not charts:
        return "table" if tables else "other"
    ch = charts[0]
    cid = ch.get("chart_type_id")
    ncat = ch.get("category_count") or 0
    if cid == 59 or (ch.get("has_legend") and ncat == 1):
        return "stack100"
    if cid in (52, 53):
        return "scale"
    if (ch.get("series_count") or 0) >= 2:
        return "rank2"
    if (ch.get("series_count") or 0) == 1:
        return "bar1"
    return "other"


def _deck_top_text_kinds(slide: dict) -> tuple[str, ...]:
    """슬라이드에 실제로 놓인 `【TOP2】` 류 텍스트 상자에서 지표 종류를 읽는다.

    브래킷 표기(`【】`)가 붙은 줄만 본다 -- 표 상단 행 표기(`TOP2(%)`)는 표 쪽에서
    따로 채우므로 여기서 또 잡으면 이중으로 쓴다.
    """
    kinds: list[str] = []
    for sh in slide.get("text_shapes") or []:
        for line in str(sh.get("text") or "").split("\n"):
            if "【" not in line:
                continue
            k = toptext_kind(line)
            if k:
                kinds.append(k)
    order = {"bot": 0, "soso": 1, "top": 2, "mean": 3}
    return tuple(sorted(dict.fromkeys(kinds), key=lambda k: order.get(k, 9)))


def family_style(template_slide: dict | None) -> dict:
    """클론 패밀리 대표에서 **표기 습관을 상속한다.**

    · fill_cols  대표 장표에서 실제로 값이 들어가는 표 열. 가로막대 장표는 `전체` 열이
                 원래 비어 있고(차트가 그 자리를 덮는다) 누적100% 장표는 채워진다.
                 이걸 규칙으로 박으면 한쪽이 전량 오탐이라 대표에서 읽는다.
    · zero       0 표기. R-110(단일선택 `-`/중복응답 공란)은 실측이 반증했다 --
                 0 이 실재하는 LG 본문 10장 중 9장 공란·1장(s20)만 dash 이고 같은
                 `가로막대` 안에서도 갈린다. 문항 유형이 아니라 담당자 습관이다.
    """
    style = {"fill_cols": None, "zero": "", "from": None}
    if not template_slide:
        return style
    tables = template_slide.get("tables") or []
    if not tables:
        return style
    shape = table_shape(tables[0])
    if shape is None:
        return style
    rows = tables[0].get("rows") or []
    fill_cols, dash = set(), 0
    for i in shape.data_rows:
        for c in shape.value_cols:
            s = norm_cell(rows[i][c]) if c < len(rows[i]) else ""
            if s == "-":
                dash += 1
                fill_cols.add(c)
            elif _is_value_text(s):
                fill_cols.add(c)
    style["fill_cols"] = sorted(fill_cols)
    style["zero"] = "-" if dash else ""
    style["from"] = template_slide["slide_no"]
    return style


def _fmt_value(row: dict, pos: int | None, zero: str) -> str:
    """통계표 셀 -> 제작본 표기.

    · 없는 값은 **지어내지 않는다.** 통계표에서 빠진 칸은 0% 라는 뜻이고(DP가 0 을
      생략한다) 제작본에서는 0 표기로 나간다.
    · 반올림은 rules 가 한다. % 1자리, 평균·NPS 2자리, ROUND_HALF_UP.
    """
    if pos is None:
        return ""
    vals = row.get("values") or []
    v = R.cell_value(vals[pos]) if pos < len(vals) else None
    if v is None or float(v) == 0.0:
        return zero
    return R.fmt_mean(v) if row.get("kind") == "mean" else R.fmt_pct(v)


def _label_lines(label: str, tbl_metrics: tuple[int, float] | None,
                 font: str) -> list[str] | None:
    """보기 라벨의 수동 줄바꿈 (R-108). 실측 폭을 못 얻으면 줄바꿈하지 않는다."""
    if not tbl_metrics or not font:
        return None
    usable, pt = tbl_metrics
    if usable <= 0:
        return None
    try:
        lines = textfit.wrap_ko(label, usable, font, pt)
    except Exception:
        return None
    return lines if len(lines) > 1 else None


def _series_for_names(names: Sequence[str | None], table: dict,
                      axis: BannerAxis, bound: Sequence[int | None],
                      opt_rows: Sequence[dict | None], opt_labels: Sequence[str],
                      kind: str, zero_as_none: bool = True
                      ) -> tuple[list[list[float | None]] | None, list[str]]:
    """계열 이름으로 값을 찾는다. **계열의 개수·순서·이름은 제작본이 진리다** (R-094).

    · 이름이 통계표 배너 라벨과 맞으면 그 열을 세로로 읽는다 (겹침막대의 1순위/1+N순위)
    · 이름이 보기 라벨과 맞으면 그 행을 가로로 읽는다 (누적100%·척도)
    · 어느 쪽에도 안 붙는 계열(LG 실측 `do not know`)은 **건드리지 않는다** -> None 반환
    """
    notes: list[str] = []
    banners = table.get("banners") or []
    by_banner = {norm_cell(b.get("label")): b["pos"] for b in banners}
    by_opt: dict[str, dict] = {}
    for r, lab in zip(opt_rows, opt_labels):
        if r is not None:
            by_opt[norm_cell(lab)] = r
            by_opt[norm_cell(R.chart_label(r.get("label")))] = r

    total = R.total_banner(table)
    tpos = total["pos"] if total else 0

    out: list[list[float | None]] = []
    for name in names:
        nm = norm_cell(name)
        if kind == "rank2" and nm in by_banner:
            pos = by_banner[nm]
            out.append([_num(r, pos) for r in opt_rows])
            continue
        if kind == "stack100" and nm in by_opt:
            out.append([_num(by_opt[nm], tpos)])
            continue
        if kind == "scale":
            row = by_opt.get(nm)
            want = None
            m = re.match(r"^(\d+)\s*점?$", nm)
            if m:
                want = int(m.group(1))
            if row is None and want is not None:
                row = next((r for r in opt_rows
                            if r is not None and r.get("option_no") == want), None)
            pts = (table.get("hints") or {}).get("scale_points") or 0
            if row is None and want is not None and 1 <= want <= int(pts or 0):
                # SORT-011 -- 값이 0% 라 통계표에서 아예 빠진 척도점이다. 0 으로 복원한다.
                notes.append(f"척도점 `{name}` 이 통계표에 없다 — 0% 로 복원 (SORT-011)")
                out.append([0.0] * len(axis.cols))
                continue
            if row is None:
                notes.append(f"계열 `{name}` 을 통계표에서 못 찾았다")
                return None, notes
            sign = _scale_sign(nm, table)
            vals = []
            for i in range(len(axis.cols)):
                pos = bound[i]
                v = _num(row, pos)
                vals.append(None if v is None else (v * sign if v else 0.0))
            out.append(vals)
            continue
        if kind == "bar1":
            pos = by_banner.get(nm, tpos)
            out.append([_num(r, pos) for r in opt_rows])
            continue
        notes.append(f"계열 `{name}` 은 통계표 배너·보기 어느 쪽과도 안 붙는다")
        return None, notes
    return out, notes


def _num(row: dict | None, pos: int | None) -> float | None:
    """차트에 넣을 숫자. 통계표에서 빠진 칸은 0% 다 (DP가 0 을 생략한다)."""
    if row is None or pos is None:
        return None
    vals = row.get("values") or []
    v = R.cell_value(vals[pos]) if pos < len(vals) else None
    if v is None:
        return 0.0
    digits = 2 if row.get("kind") == "mean" else 1
    return R.round_half_up(float(v), digits)


def _scale_sign(name: str, table: dict) -> float:
    """척도 발산형(±세로막대)에서 음의 반쪽에 -1 을 곱한다.

    중립점 = (척도점+1)//2. 7점이면 4 이하가 음의 반쪽이다 (LG s38 실측 부호와 일치).
    0 은 **0.0 그대로** 넣는다 -- -0.0 을 만들면 PowerPoint 가 그 값을 계산에 쓴다.
    """
    m = re.match(r"^(\d+)", norm_cell(name))
    if not m:
        return 1.0
    pts = (table.get("hints") or {}).get("scale_points") or 0
    if not pts:
        return 1.0
    neutral = (int(pts) + 1) // 2
    return -1.0 if int(m.group(1)) <= neutral else 1.0


def plan_slide(slide: dict, table: dict, axis: BannerAxis, kind: str,
               style: dict, *, font: str = "", metrics: tuple[int, float] | None = None,
               chart_values: str = "display", with_ment: bool = True) -> SlidePlan:
    """슬라이드 1장의 주입 계획. **여기서 숫자를 만들지 않는다** -- rules 가 낸 값을 배선만."""
    plan = SlidePlan(kind=kind)
    bound = bind_banners(axis, table)
    unbound = [i for i, p in enumerate(bound) if p is None]
    fill_cols = style.get("fill_cols")
    zero = style.get("zero", "")

    decision = _decide(table, slide)
    mode = {"scale": "scale_diverging", "stack100": "grade_stack_reverse",
            "rank2": "rank_sum_desc"}.get(kind, decision.option_axis or "table_order")

    if unbound:
        named = [f"{axis.group[i]}/{axis.label[i]}" for i in unbound]
        # `전체` 열은 통계표 배너에 없을 수 있다(순위 표는 배너가 1순위/1+N순위다)
        hard = [n for i, n in zip(unbound, named)
                if fill_cols is None or axis.cols[i] in (fill_cols or [])]
        if hard:
            plan.tags.append("BANNER")
            plan.notes.append(f"통계표 배너에 못 붙은 열 {len(hard)}개: {', '.join(hard[:6])}")

    rows = slide["tables"][0].get("rows") or []
    shape = axis.shape

    # ── ① 그래프 상단 Text 행 (표 안) ──────────────────────────────────────
    tt_rows = {}
    for i in shape.data_rows:
        k = toptext_kind(rows[i][0])
        if k:
            tt_rows[i] = k

    # ── ② 데이터 행 계획 ───────────────────────────────────────────────────
    if kind == "scale":
        # 척도 장표의 표는 TOP/평균 두 행뿐이다. 행 개수를 바꾸지 않는다.
        for i in shape.data_rows:
            k = tt_rows.get(i)
            src = R.first_row_of_kind(table, k) if k else None
            if src is None:
                plan.tags.append("TOPTEXT")
                plan.notes.append(f"표 r{i} `{norm_cell(rows[i][0])}` 에 맞는 통계표 행이 없다")
                continue
            cells = {axis.cols[j]: _fmt_value(src, bound[j], zero)
                     for j in range(len(axis.cols))
                     if fill_cols is None or axis.cols[j] in fill_cols}
            plan.data_rows.append(RowPlan(norm_cell(rows[i][0]), None, cells, k))
    else:
        labels, opt_rows, why = R.table_row_axis(table, mode)
        plan.notes.append(f"행 순서: {why}")
        for lab, r in zip(labels, opt_rows):
            cells = {}
            for j in range(len(axis.cols)):
                col = axis.cols[j]
                if fill_cols is not None and col not in fill_cols:
                    continue
                cells[col] = _fmt_value(r, bound[j], zero)
            plan.data_rows.append(
                RowPlan(lab, _label_lines(lab, metrics, font), cells, "option"))
        # 누적100% 장표는 표 하단에 TOP2/BOT2 요약 행이 따라온다
        for i in sorted(tt_rows):
            k = tt_rows[i]
            src = R.first_row_of_kind(table, k)
            if src is None:
                plan.tags.append("TOPTEXT")
                plan.notes.append(f"표 r{i} `{norm_cell(rows[i][0])}` 에 맞는 통계표 행이 없다")
                continue
            cells = {axis.cols[j]: _fmt_value(src, bound[j], zero)
                     for j in range(len(axis.cols))
                     if fill_cols is None or axis.cols[j] in fill_cols}
            plan.data_rows.append(RowPlan(norm_cell(rows[i][0]), None, cells, k))

    # ── ③ 차트 계열 ────────────────────────────────────────────────────────
    charts = slide.get("charts") or []
    if charts and kind != "table":
        ch = charts[0]
        names = ch.get("series_names") or []
        if kind == "scale":
            opt_labels, opt_rows, _ = R.chart_option_axis(table, "scale_diverging")
        else:
            opt_labels, opt_rows, _ = R.chart_option_axis(table, mode)
        series, notes = _series_for_names(names, table, axis, bound,
                                          opt_rows, opt_labels, kind)
        plan.notes.extend(notes)
        if series is None:
            plan.tags.append("SERIES")
        else:
            if chart_values == "raw":
                series = _raw_series(names, table, axis, bound, opt_rows, opt_labels, kind)
            plan.chart_series = series
            if kind == "scale":
                plan.chart_categories = None          # 배너 라벨은 덱이 진리 (중복 접미 포함)
            elif kind == "stack100":
                plan.chart_categories = None          # 카테고리 1개 -- 손대지 않는다
            else:
                plan.chart_categories = [R.chart_label(l) for l in opt_labels]

    # ── ④ Base/Unit ────────────────────────────────────────────────────────
    #
    # 제작본의 숫자를 정답으로 삼지 않는다 -- LG s6 은 실제 n=100 인데 `n=1,000` 이라고
    # 적혀 있고 포맷도 다르다. rules 가 낸 문자열로 **덮어쓴다**. 다만 기존 문자열과
    # 달라지면 근거에 남긴다 (`Unit: %` vs `Unit: %, 점` 처럼 관례가 갈리는 자리가 있다).
    plan.base_unit = R.base_unit_string(table)
    plan.footnote = R.small_base_warning(table)
    old_base = norm_cell(slide.get("base_unit"))
    if old_base and old_base != norm_cell(plan.base_unit):
        plan.notes.append(f"Base/Unit 교체: `{old_base}` → `{plan.base_unit}`")

    # ── ⑤ 독립 TextBox 형 그래프 상단 Text ─────────────────────────────────
    #
    # 어떤 지표를 싣는지는 **덱이 정한다** -- 사다리형(누적100%)은 TOP+BOT 쌍인데
    # decide_chart_type 의 기본 kinds 는 (top, mean) 이라 그대로 쓰면 【BOT2】 가
    # 통째로 누락된다(LG s26~29 실측). 슬라이드에 실제로 놓인 상자를 읽어 맞춘다.
    box_kinds = _deck_top_text_kinds(slide)
    if box_kinds:
        for item in R.graph_top_text(table, box_kinds, "독립 TextBox"):
            if item.get("value") is not None:
                plan.top_text.append((item["label"], f"{item['value']} {item['unit']}"))
            else:
                plan.tags.append("TOPTEXT")
                plan.notes.append(f"`{item['label']}` 에 넣을 값이 통계표에 없다 — 비워 둔다")

    # ── ⑥ 멘트 (가이드가 이미 쓴 것만 옮긴다 -- 새로 짓지 않는다) ──────────
    if with_ment:
        plan.ments = [m for m in (table.get("guide_ments") or []) if norm_cell(m)]

    plan.notes.append(f"차트타입 판정: {decision.chart_type} "
                      f"({decision.rule_id or '규칙없음'}/{decision.confidence})")
    return plan


def _raw_series(names, table, axis, bound, opt_rows, opt_labels, kind):
    """--chart-values raw. 표시값 대신 통계표 원값을 차트에 넣는다.

    ★ 미결: 제작본 실측이 갈린다 -- LG s38(척도)은 원값(17.647058823529413)을,
      s25(누적100%)는 표시값(15/45/37/3=100)을 담고 있다. 기본은 표시값이다.
    """
    banners = table.get("banners") or []
    by_banner = {norm_cell(b.get("label")): b["pos"] for b in banners}
    by_opt = {norm_cell(R.chart_label(r.get("label"))): r for r in opt_rows if r}
    total = R.total_banner(table)
    tpos = total["pos"] if total else 0

    def raw(row, pos):
        if row is None or pos is None:
            return None
        vals = row.get("values") or []
        cell = vals[pos] if pos < len(vals) else None
        if not isinstance(cell, dict):
            return 0.0
        v = cell.get("raw")
        return 0.0 if v is None else float(v)

    out = []
    for name in names:
        nm = norm_cell(name)
        if kind == "rank2" and nm in by_banner:
            out.append([raw(r, by_banner[nm]) for r in opt_rows])
        elif kind == "stack100" and nm in by_opt:
            out.append([raw(by_opt[nm], tpos)])
        elif kind == "scale":
            row = by_opt.get(nm)
            if row is None:
                m = re.match(r"^(\d+)", nm)
                row = next((r for r in opt_rows if r is not None and m
                            and r.get("option_no") == int(m.group(1))), None)
            if row is None:
                out.append([0.0] * len(axis.cols))       # SORT-011 복원
                continue
            sign = _scale_sign(nm, table)
            out.append([None if (v := raw(row, bound[i])) is None else (v * sign if v else 0.0)
                        for i in range(len(axis.cols))])
        else:
            out.append([raw(r, by_banner.get(nm, tpos)) for r in opt_rows])
    return out


_CHART_RULES: dict | None = None


def _decide(table: dict, slide: dict) -> R.ChartDecision:
    """차트타입 결정. 규칙은 chart_rules.yaml 에 있고 여기서는 부르기만 한다."""
    global _CHART_RULES
    if _CHART_RULES is None:
        _CHART_RULES = R.load_chart_rules()
    try:
        return R.decide_chart_type(table, None, _CHART_RULES)
    except Exception:
        return R.ChartDecision(R.FALLBACK_BASIC, None, "결정 실패", confidence="가설",
                               unresolved=True)


# ═══════════════════════════════════════════════════════════════════════════
# 10. 주입
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class SlideResult:
    slide_no: int
    ok: bool
    kind: str = ""
    join: str = ""
    table_index: Any = None
    rows_written: int = 0
    cells_written: int = 0
    series_written: int = 0
    tags: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def _shape_lines(sp, lines: Sequence[str]) -> bool:
    """문단 구조를 지키며 도형 텍스트를 갈아 끼운다.

    deckops.set_shape_text 는 도형 **이름**이 있어야 하고 전체를 첫 런에 몰아 넣는다.
    `【TOP2】 \\n60.0 %` 처럼 문단이 두 개이고 문단마다 서식이 다른 상자는 그러면
    줄이 붙어 버리므로, 여기서는 문단 i 의 첫 런에 lines[i] 를 넣고 나머지 런을 비운다.
    """
    body = sp.find(deckops.q(deckops.P, "txBody"))
    if body is None:
        return False
    paras = body.findall(deckops.q(deckops.A, "p"))
    if not paras:
        return False
    for i, para in enumerate(paras):
        if i >= len(lines):
            break
        ts = para.findall(f".//{deckops.q(deckops.A, 't')}")
        if not ts:
            continue
        ts[0].text = lines[i]
        for extra in ts[1:]:
            extra.text = ""
    return True


_FORMULA_RE = re.compile(r"^=\s*\$?([A-Za-z]+)\$?(\d+)\s*(?:\*\s*-?\s*1)?\s*$")


def sync_formula_sources(deck: deckops.Deck, chart_part: str,
                         series: Sequence[Sequence[float | None]]) -> int:
    """차트 범위가 **수식**이면 그 수식이 가리키는 입력 블록에 값을 쓴다.

    ★ 이걸 안 하면 척도 장표는 사용자가 차트를 더블클릭하는 순간 옛 더미 값으로 되돌아간다.
      LG chart78 실측 -- 차트가 보는 r2~r8 은 전부 수식이고(`=B15`, 중립점 아래는
      `=B14*-1`) 사람이 값을 넣는 곳은 r11~r17 의 `▼데이터 입력칸` 이다.
      update_embedded_workbook(preserve_formulas=True) 는 그 수식을 보존만 하므로,
      **진짜 값의 주입 지점**을 여기서 정한다(그게 fill_deck 의 결정 사항이다).

    `*-1` 은 발산형(±)의 부호 반전 장치다. 입력 블록에는 언제나 **양수 %** 를 쓴다 --
    부호는 수식이 붙인다. 여기서 음수를 넣으면 부호가 두 번 뒤집힌다.
    반환: 실제로 쓴 셀 개수 (수식이 없으면 0).
    """
    grid = deckops.workbook_grid(deck, chart_part)
    if not grid:
        return 0
    lay = deckops.chart_layout(deck.xml(chart_part))
    cells: dict[tuple[int, int], Any] = {}

    def at(r: int, c: int) -> Any:
        if 1 <= r <= len(grid) and 1 <= c <= len(grid[r - 1]):
            return grid[r - 1][c - 1]
        return None

    for s, values in enumerate(series):
        for i, v in enumerate(values):
            if lay.orientation == "rows":
                row = lay.val_rows[s] if s < len(lay.val_rows) and lay.val_rows[s] else None
                col = deckops.col_to_index(lay.first_data_col) + i
                if row is None:
                    continue
            else:
                col_letter_ = lay.val_cols[s] if s < len(lay.val_cols) else None
                if not col_letter_:
                    continue
                row = lay.first_data_row + i
                col = deckops.col_to_index(col_letter_)
            m = _FORMULA_RE.match(str(at(row, col) or ""))
            if not m:
                continue
            cells[(int(m.group(2)), deckops.col_to_index(m.group(1)))] = \
                None if v is None else abs(float(v))
    if cells:
        deckops.write_workbook_cells(deck, chart_part, cells)
    return len(cells)


def _find_shape_by_text(tree, predicate: Callable[[str], bool]):
    for sp in deckops.iter_shapes(tree):
        if predicate(deckops.shape_text(sp)):
            return sp
    return None


def ment_slot(doc: dict) -> str | None:
    """멘트 플레이스홀더 idx 를 **덱 전체에 한 번만** 판정한다 (R-090 / DP-001).

    idx 로 sub_title/ment 를 가르면 안 된다 -- LG=(12=sub, 13=ment),
    SKEE=(13=sub, 12=ment) 로 역할이 반대다. 두 덱 모두 멘트가 아래쪽(top_cm 큰 쪽)이다.
    """
    tally: dict[str, int] = {}
    for s in doc["slides"]:
        if s.get("classification") != "본문":
            continue
        phs = [(v.get("top_cm"), k) for k, v in (s.get("placeholders_by_idx") or {}).items()
               if v.get("top_cm") is not None and str(v.get("type", "")).startswith("BODY")]
        if len(phs) < 2:
            continue
        phs.sort()
        idx = phs[-1][1]
        tally[idx] = tally.get(idx, 0) + 1
    return max(tally, key=tally.get) if tally else None


def apply_slide(deck: deckops.Deck, part: str, slide: dict, plan: SlidePlan,
                *, ment_idx: str | None, keep_ment: bool = True,
                on_warn: Callable[[str], None] | None = None) -> SlideResult:
    """한 슬라이드에 표·차트·Base/Unit·상단 Text·멘트를 주입한다.

    실패는 예외 대신 SlideResult.tags 로 돌려준다 -- 한 장이 깨져도 나머지는 채운다.
    ★ 제작팀 메모 도형은 지우지 않는다. 검수 후 사람이 지운다.
    """
    warn = on_warn or (lambda _m: None)
    res = SlideResult(slide["slide_no"], True, plan.kind, tags=list(plan.tags),
                      notes=list(plan.notes))
    tree = deck.xml(part)
    dirty = False

    # ── 표 ────────────────────────────────────────────────────────────────
    try:
        tbl = deckops.first_table(tree)
    except LookupError:
        tbl = None
    if tbl is not None and plan.data_rows:
        fact_tbl = (slide.get("tables") or [{}])[0]
        shape = table_shape(fact_tbl)
        xml_rows = deckops.table_rows(tbl)
        want = len(plan.data_rows)
        if shape is None:
            res.tags.append("BASE")
            res.notes.append("표에서 (Base) 행을 못 찾아 표를 건드리지 않았다")
        else:
            idx = list(shape.data_rows)
            if len(idx) != want and not shape.header_at_bottom and \
                    idx and idx[0] == shape.base_row + 1 and idx[-1] == len(xml_rows) - 1:
                # 상단 헤더형만 행 개수를 조정한다 (R-105/R-107 은 deckops 가 안다)
                try:
                    deckops.set_row_count(tbl, want, header_rows=shape.base_row + 1,
                                          mid_style="derive", on_warn=warn)
                    idx = list(range(shape.base_row + 1, shape.base_row + 1 + want))
                    res.notes.append(f"데이터 행 {len(shape.data_rows)} → {want} 조정")
                except ValueError as e:
                    res.tags.append("ROWCOUNT")
                    res.notes.append(f"행 개수 조정 실패: {e}")
            elif len(idx) != want:
                res.tags.append("ROWCOUNT")
                res.notes.append(
                    f"데이터 행 {len(idx)}개인데 계획은 {want}개 — 겹치는 만큼만 채웠다")
            xml_rows = deckops.table_rows(tbl)
            for rp, ri in zip(plan.data_rows, idx):
                if ri >= len(xml_rows):
                    break
                cells = deckops.row_cells(xml_rows[ri])
                if rp.kind == "option" and cells:
                    deckops.put_cell(cells[0], rp.label, rp.lines)
                for col, text in rp.cells.items():
                    if col < len(cells):
                        deckops.put_cell(cells[col], text)
                        res.cells_written += 1
                res.rows_written += 1
            dirty = True

    # ── Base/Unit (위치가 아니라 `[Base:` 텍스트 패턴으로 찾는다 · R-092) ──
    if plan.base_unit:
        sp = _find_shape_by_text(tree, lambda t: t.lstrip().startswith("[Base"))
        if sp is not None:
            old = deckops.shape_text(sp).split("\n")
            lines = [plan.base_unit] + old[1:]
            if plan.footnote and plan.footnote not in "\n".join(lines):
                lines.append(plan.footnote)
            if _shape_lines(sp, lines):
                dirty = True
        else:
            res.tags.append("BASE")
            res.notes.append("`[Base:` 도형을 못 찾았다")

    # ── 독립 TextBox 형 그래프 상단 Text (【TOP2】 / 【BOT2】) ─────────────
    for label, value in plan.top_text:
        head = label.strip("【】 ")
        sp = _find_shape_by_text(tree, lambda t, h=head: h in t and "【" in t)
        if sp is None:
            res.tags.append("TOPTEXT")
            res.notes.append(f"`{label}` 텍스트 상자를 못 찾았다")
            continue
        old = deckops.shape_text(sp).split("\n")
        lines = [old[0] if old else label, value]
        if _shape_lines(sp, lines):
            dirty = True

    # ── 멘트 ──────────────────────────────────────────────────────────────
    if plan.ments and ment_idx is not None:
        cur = norm_cell((slide.get("placeholders_by_idx") or {})
                        .get(ment_idx, {}).get("text"))
        if cur and keep_ment:
            res.tags.append("MENT-KEEP")
            res.notes.append("제작본에 이미 멘트가 있어 보존했다")
        elif deckops.set_ment(tree, ment_idx, plan.ments):
            dirty = True
            res.notes.append(f"멘트 {len(plan.ments)}줄 주입(가이드 원문 재사용)")

    if dirty:
        deck.put(part, tree)

    # ── 차트 (표와 별개 파트다) ───────────────────────────────────────────
    if plan.chart_series is not None:
        parts = deck.chart_parts(part)
        if not parts:
            res.tags.append("SERIES")
            res.notes.append("슬라이드에 차트 파트가 없다")
        else:
            try:
                got = deckops.apply_chart(deck, parts[0], plan.chart_series,
                                          plan.chart_categories, strict=True,
                                          on_warn=warn)
                res.series_written = len(plan.chart_series)
                res.notes.extend(got.get("warnings") or [])
                n = sync_formula_sources(deck, parts[0], plan.chart_series)
                if n:
                    res.notes.append(
                        f"차트 범위가 수식이라 입력 블록 {n}칸에 값을 썼다 "
                        "(척도 차트의 ▼데이터 입력칸 — 안 쓰면 편집 시 옛 값으로 되돌아간다)")
            except (ValueError, KeyError) as e:
                res.tags.append("SERIES")
                res.notes.append(f"차트 주입 실패(그대로 둔다): {e}")

    res.ok = not any(t in ("MAP", "BASE", "SERIES", "ROWCOUNT") for t in res.tags)
    return res


# ═══════════════════════════════════════════════════════════════════════════
# 11. --fill 오케스트레이터
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class RunReport:
    deck: str
    facts: str
    facts_source: str
    out: str
    targets: int = 0
    filled: int = 0
    skipped: int = 0
    results: list[SlideResult] = field(default_factory=list)
    unmapped: list[dict] = field(default_factory=list)
    tag_counts: dict[str, int] = field(default_factory=dict)
    scan: dict | None = None
    notes: list[str] = field(default_factory=list)


def fill_deck(pptx: Path, facts: Path, out: Path, *,
              only: set[int] | None = None,
              chart_values: str = "display",
              with_ment: bool = True,
              keep_ment: bool = True,
              partial_blank_ratio: float = 0.6,
              memo_top_cm: float = 0.5,
              dry_run: bool = False) -> RunReport:
    """①스캔 ②대상 선별 ③매핑 ④계획 ⑤주입 ⑥저장 ⑦리포트."""
    scan = scan_deck(pptx, memo_top_cm=memo_top_cm,
                     partial_blank_ratio=partial_blank_ratio)
    doc = scan["_doc"]
    book = load_facts(facts)
    rep = RunReport(str(pptx), str(facts), book.source, str(out), scan=scan_json(scan))
    rep.notes.append(f"값 출처: {book.source} — 표 {len(book.tables)}장")
    if book.source == "guide_blocks.json":
        rep.notes.append("⚠ 제작가이드 파싱본으로 돌렸다. 값의 정본은 tables.json 이다 "
                         "(가이드 값은 이미 반올림·가공된 문자열이라 오타가 섞인다).")
    rep.notes.append(f"차트 값 표기: {chart_values} "
                     f"(미결 — LG 실측이 척도=원값 / 누적100%=표시값으로 갈린다)")

    deck = deckops.Deck(pptx)
    by_no = {s["slide_no"]: s for s in doc["slides"]}
    fams = {f["family_id"]: f for f in scan["families"]}
    styles: dict[str, dict] = {}
    idx_ment = ment_slot(doc)
    if idx_ment:
        rep.notes.append(f"멘트 플레이스홀더 idx={idx_ment} (top_cm 위치 판정, R-090)")

    targets = [t for t in scan["targets"] if not only or t["slide_no"] in only]
    rep.targets = len(targets)

    typo_keys = {d["join_key"] for d in scan["duplicates"] if "오타" in d["verdict"]}
    split_keys = {d["join_key"] for d in scan["duplicates"] if "장분리" in d["verdict"]}
    used_tables: dict[Any, list[int]] = {}

    for t in targets:
        no = t["slide_no"]
        slide = by_no[no]
        fam = fams.get(t.get("family") or "")
        table, how, why = map_slide(slide, book)
        if table is None:
            rep.unmapped.append({"slide_no": no, "join_key": slide.get("join_key"),
                                 "reason": why})
            rep.results.append(SlideResult(no, False, tags=["MAP"], notes=[why]))
            rep.skipped += 1
            continue

        tables = slide.get("tables") or []
        axis = banner_axis(tables[0]) if tables else None
        if axis is None:
            rep.results.append(SlideResult(no, False, tags=["BASE"],
                                           notes=["(Base) 행을 못 찾아 건드리지 않았다"]))
            rep.skipped += 1
            continue

        kind = panel_kind(slide, axis)
        if kind in ("other",):
            rep.results.append(SlideResult(no, False, kind=kind, tags=["TYPE"],
                                           notes=["지원 밖 장표 — 건드리지 않았다"]))
            rep.skipped += 1
            continue

        fid = t.get("family") or ""
        if fid not in styles:
            tmpl = by_no.get(fam.get("template")) if fam else None
            styles[fid] = family_style(tmpl)
        style = styles[fid]

        part = deck.slide_at(no)
        font = deck.theme_font(part)
        try:
            tbl_xml = deckops.first_table(deck.xml(part))
            metrics = deckops.label_metrics(tbl_xml, header_rows=axis.shape.base_row + 1,
                                            row=axis.shape.data_rows[0]
                                            if axis.shape.data_rows else None)
        except (LookupError, IndexError):
            metrics = None

        plan = plan_slide(slide, table, axis, kind, style, font=font, metrics=metrics,
                          chart_values=chart_values, with_ment=with_ment)
        plan.notes.insert(0, f"매핑: {how} — {why} → 표 {table.get('index')} "
                             f"`{table.get('title_full') or ''}`"[:160])
        if style.get("from"):
            plan.notes.append(f"표기 상속: 대표 s{style['from']} "
                              f"(0 표기 {'`-`' if style['zero'] else '공란'}, "
                              f"채우는 열 {len(style['fill_cols'] or [])}개)")
        else:
            plan.tags.append("NOTEMPLATE")
            plan.notes.append("클론 대표가 없어 표기 습관을 상속하지 못했다 — 사람 확인 필요")

        jk = str(slide.get("join_key") or "")
        if jk in typo_keys:
            plan.tags.append("DUPKEY")
            plan.notes.append(
                f"조인키 `{jk}` 가 떨어진 두 장에서 겹친다 — 제작본 문항코드 오타일 수 있다. "
                "매핑한 표가 맞는지 사람이 확인해야 한다")
        if jk in split_keys:
            plan.notes.append(
                f"조인키 `{jk}` 는 장분리(R-089)다 — 배너군으로 장을 갈라야 하는데 "
                "지금은 같은 표를 두 장에 쓴다. 통계표(tables.json)로 돌릴 때 "
                "rules.split_banner_chapters 로 해소한다")
        used_tables.setdefault(table.get("index"), []).append(no)

        if dry_run:
            res = SlideResult(no, True, kind, how, table.get("index"),
                              len(plan.data_rows),
                              sum(len(r.cells) for r in plan.data_rows),
                              len(plan.chart_series or []), list(plan.tags),
                              list(plan.notes))
        else:
            res = apply_slide(deck, part, slide, plan, ment_idx=idx_ment,
                              keep_ment=keep_ment,
                              on_warn=lambda m: plan.notes.append(m))
            res.join, res.table_index = how, table.get("index")
        rep.results.append(res)
        if res.ok:
            rep.filled += 1
        else:
            rep.skipped += 1

    shared = {k: v for k, v in used_tables.items() if len(v) > 1}
    if shared:
        rep.notes.append(
            "같은 통계표를 두 장 이상에 쓴 곳: " +
            ", ".join(f"표{k}→{'·'.join('s%d' % n for n in v)}" for k, v in shared.items()) +
            " (장분리 미해소 — 사람 확인 필요)")

    for r in rep.results:
        for tag in r.tags:
            rep.tag_counts[tag] = rep.tag_counts.get(tag, 0) + 1

    if not dry_run and rep.filled:
        out.parent.mkdir(parents=True, exist_ok=True)
        deck.save(out)
        rep.notes.append(f"고친 파트 {len(deck.touched())}개 — 나머지는 바이트 그대로")
    elif not dry_run:
        rep.notes.append("채운 슬라이드가 0장이라 저장하지 않았다")
    return rep


def render_fill_report(rep: RunReport) -> str:
    """검수리포트.md. 슬라이드마다 무엇을 넣었는지 남긴다 -- baseline diff 학습용이다."""
    L: list[str] = []
    a = L.append
    a(f"# 채우기 검수 리포트 — {Path(rep.deck).name}")
    a("")
    a(f"- 제작본: `{rep.deck}`")
    a(f"- 값 출처: `{rep.facts}` ({rep.facts_source})")
    a(f"- 산출물: `{rep.out}`")
    a(f"- 대상 {rep.targets}장 / 채움 {rep.filled}장 / 건너뜀 {rep.skipped}장")
    if rep.tag_counts:
        a(f"- 태그: " + ", ".join(f"{k} {v}" for k, v in sorted(rep.tag_counts.items())))
    a("")
    for n in rep.notes:
        a(f"> {n}")
    a("")
    a("## 매핑 실패 (건드리지 않았다)")
    if not rep.unmapped:
        a("- 없음")
    for u in rep.unmapped:
        a(f"- s{u['slide_no']} `{u.get('join_key')}` — {u['reason']}")
    a("")
    a("## 슬라이드별 기록")
    a("| 슬라이드 | 종류 | 조인 | 표 | 행 | 셀 | 계열 | 태그 |")
    a("|---|---|---|---|---|---|---|---|")
    for r in rep.results:
        a(f"| s{r.slide_no} | {r.kind or '-'} | {r.join or '-'} | {r.table_index or '-'} "
          f"| {r.rows_written} | {r.cells_written} | {r.series_written} "
          f"| {'/'.join(r.tags) or '-'} |")
    a("")
    a("## 근거")
    for r in rep.results:
        if not r.notes:
            continue
        a(f"- **s{r.slide_no}**")
        for n in r.notes:
            a(f"  - {n}")
    return "\n".join(L)


# ═══════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════

def _parse_only(s: str | None) -> set[int] | None:
    if not s:
        return None
    out: set[int] = set()
    for part in s.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            lo, hi = part.split("-", 1)
            out.update(range(int(lo), int(hi) + 1))
        else:
            out.add(int(part))
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        prog="fill_deck.py",
        description="제작본 .pptx 를 훑고(--scan) 통계표 값으로 채운다(--fill).")
    ap.add_argument("--scan", metavar="DECK",
                    help="제작본 .pptx 를 읽기 전용으로 훑는다")
    ap.add_argument("--fill", nargs="+", metavar="ARG",
                    help="DECK FACTS OUT — 값을 주입해 OUT 에 새로 쓴다")
    ap.add_argument("--out", metavar="PPTX", help="--fill 의 산출 경로(위치인자 대신)")
    ap.add_argument("--json", metavar="JSON", help="--scan 결과를 기계용 JSON 으로")
    ap.add_argument("--md", metavar="MD", help="--scan 리포트를 마크다운 파일로")
    ap.add_argument("--report", metavar="MD", help="--fill 검수리포트 경로")
    ap.add_argument("--observations", metavar="JSONL", default=str(DEFAULT_OBSERVATIONS),
                    help="제작팀 메모를 적재할 관찰 로그 (기본 engine/learning/observations.jsonl)")
    ap.add_argument("--no-observations", action="store_true", help="관찰 로그 적재를 끈다")
    ap.add_argument("--memo-top-cm", type=float, default=0.5,
                    help="메모 판정 상단 허용치 cm (기본 0.5)")
    ap.add_argument("--partial-blank-ratio", type=float, default=0.6,
                    help="PARTIAL 판정 임계 (기본 0.6 — 미검증)")
    ap.add_argument("--only", metavar="18,23,25-29", help="슬라이드 번호 화이트리스트")
    ap.add_argument("--chart-values", choices=("display", "raw"), default="display",
                    help="차트에 넣을 값 (기본 display=표시값)")
    ap.add_argument("--no-ment", action="store_true", help="멘트를 건드리지 않는다")
    ap.add_argument("--overwrite-ment", action="store_true",
                    help="제작본에 이미 있는 멘트도 덮어쓴다 (기본은 보존)")
    ap.add_argument("--dry-run", action="store_true", help="쓰지 않고 리포트만 낸다")
    ap.add_argument("--quiet", action="store_true", help="콘솔 리포트를 줄인다")
    args = ap.parse_args(argv)

    if bool(args.scan) == bool(args.fill):
        ap.error("--scan 또는 --fill 중 하나를 정확히 하나 골라야 한다")

    # ── --scan ────────────────────────────────────────────────────────────
    if args.scan:
        deck_path = Path(args.scan)
        if not deck_path.exists():
            print(f"[오류] 파일이 없다: {deck_path}", file=sys.stderr)
            return 3
        scan = scan_deck(deck_path, memo_top_cm=args.memo_top_cm,
                         partial_blank_ratio=args.partial_blank_ratio)
        report = render_scan_report(scan)
        if not args.quiet:
            print(report)
        if args.json:
            p = Path(args.json)
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(json.dumps(scan_json(scan), ensure_ascii=False, indent=1),
                         encoding="utf-8")
            print(f"\n[JSON] {p}")
        if args.md:
            p = Path(args.md)
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(report, encoding="utf-8")
            print(f"[MD] {p}")
        if not args.no_observations:
            n = write_observations(scan, Path(args.observations))
            print(f"[관찰로그] dp_memo {n}건 → {args.observations}")
        nh = scan["needs_human"]
        need = bool(nh["families_without_template"] or nh["nobase_slides"])
        todo = [t for t in scan["targets"] if t["state"] in ("EMPTY", "SENTINEL")]
        print(f"\n요약: 본문 {scan['body_slide_count']}장 / 대상 {len(scan['targets'])}장"
              f"(실제 미기입 {len(todo)}장) "
              f"/ 패밀리 {len([f for f in scan['families'] if f['kind'] != 'nodata'])}개 "
              f"/ 사람 판단 {'필요' if need else '불필요'}")
        return 1 if need else 0

    # ── --fill ────────────────────────────────────────────────────────────
    rest = list(args.fill)
    if len(rest) == 3:
        deck_path, facts_path, out_path = (Path(rest[0]), Path(rest[1]), Path(rest[2]))
    elif len(rest) == 2 and args.out:
        deck_path, facts_path, out_path = Path(rest[0]), Path(rest[1]), Path(args.out)
    else:
        ap.error("--fill 은 DECK FACTS OUT 세 개(또는 DECK FACTS 와 --out)를 받는다")
        return 3
    for p in (deck_path, facts_path):
        if not p.exists():
            print(f"[오류] 파일이 없다: {p}", file=sys.stderr)
            return 3
    if out_path.resolve() == deck_path.resolve():
        print("[오류] 원본을 덮어쓸 수 없다", file=sys.stderr)
        return 3

    rep = fill_deck(deck_path, facts_path, out_path,
                    only=_parse_only(args.only),
                    chart_values=args.chart_values,
                    with_ment=not args.no_ment,
                    keep_ment=not args.overwrite_ment,
                    partial_blank_ratio=args.partial_blank_ratio,
                    memo_top_cm=args.memo_top_cm,
                    dry_run=args.dry_run)

    md = render_fill_report(rep)
    report_path = Path(args.report) if args.report else \
        out_path.with_name(out_path.stem + "_검수리포트.md")
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(md, encoding="utf-8")

    if not args.quiet:
        head = md.split("## 슬라이드별 기록")[0]
        print(head)
    print(f"[리포트] {report_path}")
    print(f"대상 {rep.targets}장 / 채움 {rep.filled}장 / 매핑실패 {len(rep.unmapped)}장 "
          f"/ 경고 {sum(rep.tag_counts.values())}건")
    if not args.dry_run and rep.filled:
        print(f"[산출물] {out_path}")

    if any(t in rep.tag_counts for t in ("BASE", "MENT-NUM")):
        return 1
    if rep.unmapped:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
