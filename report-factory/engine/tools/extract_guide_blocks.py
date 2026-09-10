# -*- coding: utf-8 -*-
"""
extract_guide_blocks.py -- 제작가이드(.xlsm) 블록 추출기

연구원이 작성한 제작가이드에서 **저작 블록(authored block)** 을 통째로 뽑아
ground_truth_*.json (제작본 실측 정답지) 과 대조 가능한 JSON으로 만든다.

블록 = 라벨 열에 아래 6개 영문 라벨이 **세로로 연속**한 구간.

    Slide title
    Sub title
    Chart type
    Base/Unit
    Question/Reference
    Data

그 아래로:
    멘트 (Data 라벨 행 바로 아래 값 열, 라벨 없음)
    -> 패널 제목 행 -> 배너 헤더 행 -> `Base for %` 행 -> 데이터 행
    -> `      계`(공백 6) -> BOT2/BOT3 -> SoSo
    -> `그래프 상단 Text` + TOP2/TOP3 -> `그래프 상단 Text` + [5점 평균]/[7점 평균]
    (빈 행 2줄) -> 다음 블록

시트/열 배치 (실측):
    LG   : sheet `본문_제작용`  라벨=A열, 값=B열
    SKEE : sheet `본문`         A열=섹션 인덱스, 라벨=B열, 값=C열

다단(多段) 슬라이드는 `Chart type` / `Base/Unit` / 패널제목 / 데이터를 **가로로 반복**한다.
열 위치는 고정이 아니다 (LG 실측 B/C/I/N/Q, SKEE 실측 C/D/G/H) -> 행 전체를 스캔해
비어 있지 않은 셀을 전부 패널로 잡는다.

블록 아래에 **헤더 없이 표만 붙은 구간**(LG 시나리오 7~15)은 `orphan_tables` 로 별도 수집한다.

사용법:
    python extract_guide_blocks.py                    # refs/ 의 예시 2벌
    python extract_guide_blocks.py <in.xlsm> <sheet> <label_col> <out.json>
"""

from __future__ import annotations

import json
import re
import sys
import warnings
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

REFS_DIR = Path(__file__).resolve().parents[1] / "refs"

GUIDES = [
    # (파일명, 시트명, 라벨 열 index(1-based), 출력 파일명)
    ("예시_제작가이드_LG홈로봇_0317.xlsm", "본문_제작용", 1, "guide_blocks_LG홈로봇.json"),
    ("1_SKEE만족도/1-4. skee 만족도 제작가이드.xlsm", "본문", 2, "guide_blocks_SKEE.json"),
]

BLOCK_LABELS = (
    "Slide title",
    "Sub title",
    "Chart type",
    "Base/Unit",
    "Question/Reference",
    "Data",
)

# 차트타입 문자열 후보 판별
CHART_TYPE_RE = re.compile(r"^\s*(전체|배너)\s*[:：]")
STRUCT_CHART_TYPES = {"간지", "표지", "목차"}

# Base/Unit 문자열
BASE_UNIT_RE = re.compile(r"^\s*\[\s*Base", re.IGNORECASE)
BASE_N_RE = re.compile(r"\(\s*n\s*=\s*([0-9,]+)\s*\)", re.IGNORECASE)

# Question/Reference
QU_PREFIX_RE = re.compile(r"^\s*Qu\s*[:：]", re.IGNORECASE)

# 문항코드 (제작본 소제목과 동일 형태) - 하이픈류 전부 허용
_DASH = r"[\-‐‑‒–—―−]"
_CODE_ATOM = r"(?:SQ|DQ|[A-Z]{1,3})\d+(?:\s*" + _DASH + r"\s*\d+)*"
QUESTION_CODE_RE = re.compile(
    r"^\s*(" + _CODE_ATOM + r"(?:\s*/\s*" + _CODE_ATOM + r")*)\s*[.)]"
)

# 데이터 행 센티널 (통계표 요약행). 전각 렌티큘러 괄호에 유의.
SENTINELS = {
    "base": ("Base for %", "(Base)"),
    "total": ("계",),                      # 실제 표기는 앞 공백 6 + 계
    "bot": ("【BOT2】", "【BOT3】"),
    "soso": ("【SoSo】",),
    "top": ("【TOP2】", "【TOP3】"),
    "mean": ("[5점 평균]", "[7점 평균]", "[10점 평균]", "[100점 평균]"),
}

# 행 단위 수식어 (데이터 행 라벨 열에 적는 제작팀 지시어)
ROW_MODIFIERS = (
    "그래프 상단 Text",
    "100%세로막대",
    "100%가로막대",
    "가로막대",
    "세로막대",
    "그래프 색",
    "각 그래프마다 base 표기",
    "text",
)

# 레이아웃 지시 (패널 옆 빈 셀에 적음)
LAYOUT_DIRECTIVE_RE = re.compile(
    r"공간만 만들어주세요|참고 이미지|TEXT 박스|부탁드립니다|장분리|장 분리"
)


# --------------------------------------------------------------------------- #
# 저수준
# --------------------------------------------------------------------------- #


def cell_val(ws, r: int, c: int):
    v = ws.cell(row=r, column=c).value
    if isinstance(v, str):
        v = v.replace("\x0b", "\n")
    return v


def is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def as_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    return str(v)


def row_cells(ws, r: int, from_col: int, to_col: int) -> dict[str, Any]:
    """비어 있지 않은 셀만 {열문자: 값}."""
    out = {}
    for c in range(from_col, to_col + 1):
        v = cell_val(ws, r, c)
        if not is_blank(v):
            out[get_column_letter(c)] = v
    return out


def row_cells_all(ws, r: int, from_col: int, to_col: int) -> list:
    """원문 보존용 - 끝의 None 은 잘라낸다."""
    vals = [cell_val(ws, r, c) for c in range(from_col, to_col + 1)]
    while vals and is_blank(vals[-1]):
        vals.pop()
    return vals


def norm_key(s) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", "", as_text(s).replace("：", ":")).lower()


# --------------------------------------------------------------------------- #
# 블록 탐지
# --------------------------------------------------------------------------- #


def find_block_starts(ws, label_col: int) -> list[int]:
    """라벨 6줄이 정확히 연속하는 시작 행 번호들."""
    starts = []
    for r in range(1, ws.max_row - 4):
        ok = True
        for i, lab in enumerate(BLOCK_LABELS):
            v = cell_val(ws, r + i, label_col)
            if not isinstance(v, str) or v.strip() != lab:
                ok = False
                break
        if ok:
            starts.append(r)
    return starts


def classify_sentinel(label: str) -> str | None:
    s = as_text(label).strip()
    if not s:
        return None
    for kind, keys in SENTINELS.items():
        for k in keys:
            if s == k.strip():
                return kind
    return None


def is_dash_placeholder(v) -> bool:
    """`lovelySeoyeon` 매크로가 '눈에 빈 칸'을 채워 넣은 문자열 `-`."""
    return isinstance(v, str) and v.strip() == "-"


def collect_panels_from_row(ws, r: int, from_col: int, max_col: int, pred) -> dict[str, Any]:
    """행 r 에서 pred(값) 을 만족하는 셀을 {열문자: 값} 으로. `-` 자리표시자는 제외."""
    out = {}
    for c in range(from_col, max_col + 1):
        v = cell_val(ws, r, c)
        if is_blank(v) or is_dash_placeholder(v):
            continue
        if pred(v):
            out[get_column_letter(c)] = as_text(v).strip()
    return out


def is_chart_type_value(v) -> bool:
    s = as_text(v).strip()
    if not s:
        return False
    if s in STRUCT_CHART_TYPES:
        return True
    return bool(CHART_TYPE_RE.match(s))


def is_base_unit_value(v) -> bool:
    return bool(BASE_UNIT_RE.match(as_text(v)))


# --------------------------------------------------------------------------- #
# 블록 1개 파싱
# --------------------------------------------------------------------------- #


def find_orphan_cut(ws, body_start: int, end: int, max_col: int) -> int | None:
    """
    블록 본문 안에서 **더 이상 이 블록이 선언한 것이 아닌 표**가 시작되는 행을 찾는다.

    LG 가이드는 시나리오 7~15 통계표를 헤더 블록 없이 직전 블록 뒤에 그대로 붙여 놨다.
    빈 행 개수로는 구분되지 않으므로(구분자가 1줄뿐인 곳도 있다) 아래 3조건을 모두 쓴다.

      (a) 직전 표와 이번 표 사이에 인라인 `Chart type` 선언 행이 **없다**
      (b) 이번 표의 `Base for %` 가 직전 표와 **같은 열**에 있다 (= 가로 패널이 아니다)
      (c) 둘 사이에 **문항코드가 붙은 패널 제목 행**이 새로 등장한다

    세 조건이 동시에 성립하면 그 패널 제목 행부터가 헤더 없는 표다.
    반환값은 잘라낼 행 번호(그 행부터 orphan). 없으면 None.
    """
    tables: list[tuple[int, str]] = []      # (row, 첫 Base for % 열)
    inline_ct_rows: list[int] = []
    code_rows: list[int] = []

    for rr in range(body_start, end + 1):
        cells = row_cells(ws, rr, 1, max_col)
        if not cells:
            continue
        base_col = None
        for c, v in cells.items():
            s = as_text(v).strip()
            if s in ("Base for %", "(Base)") and base_col is None:
                base_col = c
            if is_chart_type_value(v):
                inline_ct_rows.append(rr)
            if not QU_PREFIX_RE.match(s) and QUESTION_CODE_RE.match(s):
                code_rows.append(rr)
        if base_col:
            tables.append((rr, base_col))

    for i in range(1, len(tables)):
        prev_row, prev_col = tables[i - 1]
        cur_row, cur_col = tables[i]
        if any(prev_row < r < cur_row for r in inline_ct_rows):
            continue                                    # (a) 위배 - 새 패널이 선언됨
        if cur_col != prev_col:
            continue                                    # (b) 위배 - 가로 패널
        between_codes = [r for r in code_rows if prev_row < r < cur_row]
        if not between_codes:
            continue                                    # (c) 위배 - 같은 표의 반복 그룹
        return min(between_codes)
    return None


def parse_block(ws, start: int, end: int, label_col: int, max_col: int) -> dict:
    """
    start: 'Slide title' 행 번호
    end:   이 블록이 차지하는 마지막 행(포함) - 보통 다음 블록 시작 - 1
    """
    value_col = label_col + 1
    vc = get_column_letter(value_col)

    hdr = {}
    for i, lab in enumerate(BLOCK_LABELS):
        r = start + i
        hdr[lab] = {
            "row": r,
            "value": cell_val(ws, r, value_col),
            "row_cells": row_cells(ws, r, value_col, max_col),
        }

    slide_title = as_text(hdr["Slide title"]["value"]).strip()
    sub_title = as_text(hdr["Sub title"]["value"]).strip()
    question_ref = as_text(hdr["Question/Reference"]["value"]).strip()

    # ---- 헤더행 패널 -------------------------------------------------------- #
    chart_row = hdr["Chart type"]["row"]
    base_row = hdr["Base/Unit"]["row"]
    chart_panels = collect_panels_from_row(ws, chart_row, value_col, max_col, lambda v: True)
    base_panels = collect_panels_from_row(ws, base_row, value_col, max_col, lambda v: True)

    # 실측 결함: 간지 블록에서 '간지'를 Chart type 행이 아니라 Base/Unit 행에 적은 경우가 있다.
    misplaced_chart_type = {
        col: v
        for col, v in base_panels.items()
        if as_text(v).strip() in STRUCT_CHART_TYPES
    }

    # ---- 멘트 -------------------------------------------------------------- #
    data_row = hdr["Data"]["row"]
    ments: list[str] = []
    ment_rows: list[int] = []
    if as_text(hdr["Data"]["value"]).strip():
        ments.append(as_text(hdr["Data"]["value"]).strip())
        ment_rows.append(data_row)
    r = data_row + 1
    while r <= end:
        v = cell_val(ws, r, value_col)
        if is_blank(v):
            break
        s = as_text(v).strip()
        # 이미 데이터 영역(패널제목/센티널/차트타입)이 시작됐으면 중단
        if classify_sentinel(s) or is_chart_type_value(s) or is_base_unit_value(s):
            break
        ments.append(s)
        ment_rows.append(r)
        r += 1

    body_start = r

    # 헤더 없이 붙은 표가 이어지면 그 앞에서 블록을 끊는다 (나머지는 orphan_tables 로).
    orphan_cut = find_orphan_cut(ws, body_start, end, max_col)
    body_end = (orphan_cut - 1) if orphan_cut else end

    # ---- 본문(데이터) 영역 스캔 --------------------------------------------- #
    body_rows: list[dict] = []
    inline_chart_panels: list[dict] = []
    panel_titles: list[dict] = []
    sentinel_rows: list[dict] = []
    row_modifier_rows: list[dict] = []
    layout_directives: list[dict] = []
    base_for_pct_rows: list[int] = []
    codes_in_body: list[dict] = []
    dash_placeholder_cells = 0

    for rr in range(body_start, body_end + 1):
        cells = row_cells(ws, rr, 1, max_col)
        if not cells:
            body_rows.append({"row": rr, "cells": {}})
            continue

        raw = row_cells_all(ws, rr, 1, max_col)
        dash_placeholder_cells += sum(
            1 for v in raw if isinstance(v, str) and v.strip() == "-"
        )

        # 인라인 Chart type 행? (아래 행에 Base/Unit 이 오면 확정)
        ct_here = {c: as_text(v).strip() for c, v in cells.items() if is_chart_type_value(v)}
        if ct_here:
            bu_next = collect_panels_from_row(ws, rr + 1, 1, max_col, is_base_unit_value)
            inline_chart_panels.append(
                {
                    "row": rr,
                    "chart_types": ct_here,
                    "base_units": bu_next,
                    "confirmed_by_base_row": bool(bu_next),
                }
            )

        for c, v in cells.items():
            s = as_text(v).strip()

            # 레이아웃 지시
            if LAYOUT_DIRECTIVE_RE.search(s) and not is_chart_type_value(s):
                layout_directives.append({"row": rr, "col": c, "text": s})

            # 행 수식어
            if s in ROW_MODIFIERS:
                row_modifier_rows.append({"row": rr, "col": c, "modifier": s})

            # 센티널
            kind = classify_sentinel(s)
            if kind:
                sentinel_rows.append(
                    {
                        "row": rr,
                        "col": c,
                        "kind": kind,
                        "label": as_text(v),
                        "values": row_cells(
                            ws, rr, column_index_from_string(c) + 1, max_col
                        ),
                    }
                )
                if kind == "base":
                    base_for_pct_rows.append(rr)

            # 문항코드
            if not QU_PREFIX_RE.match(s):
                m = QUESTION_CODE_RE.match(s)
                if m:
                    code = re.sub(r"\s+", "", m.group(1))
                    codes_in_body.append({"row": rr, "col": c, "code": code, "text": s})
                    panel_titles.append({"row": rr, "col": c, "text": s, "code": code})

        body_rows.append({"row": rr, "cells": cells})

    while body_rows and not body_rows[-1]["cells"]:
        body_rows.pop()

    # ---- 차트타입 / Base 정리 ------------------------------------------------ #
    chart_type_values = [as_text(v).strip() for v in chart_panels.values()]
    chart_type_values += [
        v for p in inline_chart_panels for v in p["chart_types"].values()
    ]
    if not chart_type_values and misplaced_chart_type:
        chart_type_values = [as_text(v).strip() for v in misplaced_chart_type.values()]

    base_unit_values = [
        as_text(v).strip() for v in base_panels.values() if is_base_unit_value(v)
    ]
    base_unit_values += [
        as_text(v).strip() for p in inline_chart_panels for v in p["base_units"].values()
    ]

    primary_code = codes_in_body[0]["code"] if codes_in_body else None

    if primary_code:
        join_key, join_kind = primary_code, "question_code"
    else:
        jk = norm_key(slide_title) + "|" + norm_key(sub_title)
        join_key, join_kind = (jk if jk.strip("|") else None), "title+subtitle"

    n_match = BASE_N_RE.search(base_unit_values[0]) if base_unit_values else None

    # 같은 행에 Base for % 가 여러 열에 있으면 가로 데이터 패널이다.
    horizontal_data_panels = 0
    for rr in base_for_pct_rows:
        n = sum(
            1
            for v in row_cells(ws, rr, 1, max_col).values()
            if as_text(v).strip() in ("Base for %", "(Base)")
        )
        horizontal_data_panels = max(horizontal_data_panels, n)

    return {
        "block_index": None,
        "start_row": start,
        "end_row": body_end,
        "orphan_cut_row": orphan_cut,
        "horizontal_data_panels": horizontal_data_panels,
        "label_col": get_column_letter(label_col),
        "value_col": vc,
        "section_index": cell_val(ws, start, 1) if label_col > 1 else None,
        "slide_title": slide_title,
        "sub_title": sub_title,
        "question_reference": question_ref,
        "question_reference_is_qu_prefixed": bool(QU_PREFIX_RE.match(question_ref)),
        "chart_type": chart_type_values[0] if chart_type_values else None,
        "chart_types_all": chart_type_values,
        "chart_type_panels": chart_panels,
        "chart_type_misplaced_in_base_row": misplaced_chart_type or None,
        "base_unit": base_unit_values[0] if base_unit_values else None,
        "base_units_all": base_unit_values,
        "base_unit_panels": base_panels,
        "base_n": int(n_match.group(1).replace(",", "")) if n_match else None,
        "panel_count": len(chart_panels)
        + sum(len(p["chart_types"]) for p in inline_chart_panels),
        "inline_panels": inline_chart_panels,
        "ments": ments,
        "ment_rows": ment_rows,
        "ment_count": len(ments),
        "panel_titles": panel_titles,
        "question_codes_in_body": codes_in_body,
        "question_code": primary_code,
        "join_key": join_key,
        "join_key_kind": join_kind,
        "sentinel_rows": sentinel_rows,
        "sentinel_kinds": sorted({s["kind"] for s in sentinel_rows}),
        "is_scale_question": any(
            s["kind"] in ("top", "bot", "soso", "mean") for s in sentinel_rows
        ),
        "base_for_pct_rows": base_for_pct_rows,
        "table_count_in_block": len(base_for_pct_rows),
        "row_modifiers": row_modifier_rows,
        "graph_top_text_count": sum(
            1 for m in row_modifier_rows if m["modifier"] == "그래프 상단 Text"
        ),
        "layout_directives": layout_directives,
        "dash_placeholder_cells": dash_placeholder_cells,
        "body_rows": body_rows,
    }


# --------------------------------------------------------------------------- #
# 헤더 없는 표 (orphan) 수집
# --------------------------------------------------------------------------- #


def collect_orphan_tables(ws, covered: set[int], max_col: int) -> list[dict]:
    """
    블록에 속하지 않는 행 구간에서 `Base for %` 행을 찾아 표 단위로 묶는다.
    LG 가이드가 시나리오 7~15 통계표를 헤더 블록 없이 붙여 놓은 구간이 여기 해당.
    """
    orphan_rows = sorted(set(range(1, ws.max_row + 1)) - covered)
    if not orphan_rows:
        return []

    regions: list[list[int]] = []
    cur = [orphan_rows[0]]
    for r in orphan_rows[1:]:
        if r == cur[-1] + 1:
            cur.append(r)
        else:
            regions.append(cur)
            cur = [r]
    regions.append(cur)

    tables: list[dict] = []
    for reg in regions:
        if not any(row_cells(ws, r, 1, max_col) for r in reg):
            continue
        base_rows: list[tuple[int, str]] = []
        for r in reg:
            for c, v in row_cells(ws, r, 1, max_col).items():
                if as_text(v).strip() in ("Base for %", "(Base)"):
                    base_rows.append((r, c))
                    break
        if not base_rows:
            continue
        for i, (br, bcol) in enumerate(base_rows):
            tbl_start = reg[0] if i == 0 else base_rows[i - 1][0] + 1
            code = title = None
            code_row = None
            for rr in range(max(tbl_start, br - 6), br):
                for c, v in row_cells(ws, rr, 1, max_col).items():
                    s = as_text(v).strip()
                    if QU_PREFIX_RE.match(s):
                        continue
                    m = QUESTION_CODE_RE.match(s)
                    if m:
                        code = re.sub(r"\s+", "", m.group(1))
                        title = s
                        code_row = rr
                        break
                if code:
                    break
            tbl_end = (base_rows[i + 1][0] - 1) if i + 1 < len(base_rows) else reg[-1]
            while tbl_end > br and not row_cells(ws, tbl_end, 1, max_col):
                tbl_end -= 1
            rows_out = []
            for rr in range(code_row or br, tbl_end + 1):
                rows_out.append({"row": rr, "cells": row_cells(ws, rr, 1, max_col)})
            sent = []
            for rr in range(br, tbl_end + 1):
                for c, v in row_cells(ws, rr, 1, max_col).items():
                    k = classify_sentinel(as_text(v).strip())
                    if k:
                        sent.append({"row": rr, "col": c, "kind": k, "label": as_text(v)})
            tables.append(
                {
                    "region_start": reg[0],
                    "region_end": reg[-1],
                    "base_row": br,
                    "base_col": bcol,
                    "start_row": code_row or br,
                    "end_row": tbl_end,
                    "question_code": code,
                    "panel_title": title,
                    "join_key": code,
                    "join_key_kind": "question_code" if code else None,
                    "has_header_block": False,
                    "sentinel_kinds": sorted({s["kind"] for s in sent}),
                    "is_scale_question": any(
                        s["kind"] in ("top", "bot", "soso", "mean") for s in sent
                    ),
                    "rows": rows_out,
                }
            )
    return tables


# --------------------------------------------------------------------------- #
# 가이드 1벌
# --------------------------------------------------------------------------- #


def extract_guide(xlsm_path: str | Path, sheet: str, label_col: int) -> dict:
    xlsm_path = Path(xlsm_path)
    wb = openpyxl.load_workbook(str(xlsm_path), data_only=True)
    ws = wb[sheet]
    max_col = ws.max_column
    max_row = ws.max_row

    starts = find_block_starts(ws, label_col)
    blocks: list[dict] = []
    covered: set[int] = set()

    for i, st in enumerate(starts):
        end = (starts[i + 1] - 1) if i + 1 < len(starts) else max_row
        blk = parse_block(ws, st, end, label_col, max_col)
        blk["block_index"] = i
        blocks.append(blk)

    for blk in blocks:
        last = blk["body_rows"][-1]["row"] if blk["body_rows"] else blk["start_row"] + 5
        covered |= set(range(blk["start_row"], last + 1))
        blk["end_row"] = last

    orphans = collect_orphan_tables(ws, covered, max_col)

    ct_counts: dict[str, int] = {}
    for b in blocks:
        for ct in b["chart_types_all"]:
            ct_counts[ct] = ct_counts.get(ct, 0) + 1

    coded = [b for b in blocks if b["question_code"]]
    multi = [b for b in blocks if b["panel_count"] > 1]

    return {
        "source_file": xlsm_path.name,
        "source_path": str(xlsm_path),
        "sheet": sheet,
        "label_col": get_column_letter(label_col),
        "value_col": get_column_letter(label_col + 1),
        "max_row": max_row,
        "max_col": max_col,
        "summary": {
            "block_count": len(blocks),
            "blocks_with_question_code": len(coded),
            "blocks_with_chart_type": sum(1 for b in blocks if b["chart_type"]),
            "blocks_with_ments": sum(1 for b in blocks if b["ments"]),
            "ment_line_total": sum(b["ment_count"] for b in blocks),
            "multi_panel_blocks": len(multi),
            "scale_blocks": sum(1 for b in blocks if b["is_scale_question"]),
            "graph_top_text_total": sum(b["graph_top_text_count"] for b in blocks),
            "chart_type_counts": dict(sorted(ct_counts.items(), key=lambda kv: -kv[1])),
            "orphan_table_count": len(orphans),
            "orphan_tables_with_code": sum(1 for t in orphans if t["question_code"]),
            "orphan_row_span": (
                [orphans[0]["region_start"], orphans[-1]["region_end"]] if orphans else None
            ),
            "orphan_codes": sorted({t["question_code"] for t in orphans if t["question_code"]}),
            "blocks_with_chart_type_in_base_row": sum(
                1 for b in blocks if b["chart_type_misplaced_in_base_row"]
            ),
        },
        "blocks": blocks,
        "orphan_tables": orphans,
    }


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #


def _write(doc: dict, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(doc, ensure_ascii=False, indent=1, default=str), encoding="utf-8"
    )


def main(argv: list[str]) -> int:
    if len(argv) == 5:
        doc = extract_guide(argv[1], argv[2], int(argv[3]))
        _write(doc, Path(argv[4]))
        print("%s -> %s (%d blocks)" % (argv[1], argv[4], doc["summary"]["block_count"]))
        return 0
    if len(argv) != 1:
        print(__doc__)
        return 2

    for src, sheet, lab, dst in GUIDES:
        p = REFS_DIR / src
        if not p.exists():
            print("SKIP (없음): %s" % p)
            continue
        doc = extract_guide(p, sheet, lab)
        _write(doc, REFS_DIR / dst)
        s = doc["summary"]
        print("=" * 78)
        print("%s [%s] -> %s" % (src, sheet, dst))
        print("  블록:", s["block_count"], " 문항코드 보유:", s["blocks_with_question_code"])
        print("  멘트 보유 블록:", s["blocks_with_ments"], " 멘트 줄 총계:", s["ment_line_total"])
        print("  다단 블록:", s["multi_panel_blocks"], " 척도 블록:", s["scale_blocks"])
        print("  그래프 상단 Text:", s["graph_top_text_total"])
        print("  헤더없는 표(orphan):", s["orphan_table_count"], " 행범위:", s["orphan_row_span"])
        print("  차트타입 문자열:")
        for k, v in list(s["chart_type_counts"].items())[:20]:
            print("      %3d  %s" % (v, k))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main(sys.argv))
