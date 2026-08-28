"""통계표(.xls/.xlsx) 공통 파싱 로직.

`parse_table_biff.py`(BIFF8 OLE2 .xls)와 `parse_table_xlsx.py`(.xlsx)가 이 모듈을
공유한다. 리더(reader)가 시트를 `Grid`로 읽어오면, 이 모듈이 표 경계 탐지 →
배너 해석 → 행 분류 → 서식 적용까지 수행하고 동일한 `tables.json` 스키마를 만든다.

핵심 규약 (engine 지식 §3.3 / §3.6):

    row T     표 제목            (A열, 병합)
    row T+1   배너 그룹 헤더      (B 또는 C열 = '전체')
    row T+2   배너 세부 헤더
    row T+3   'Base for %' + '(500)'
    row T+4~  데이터 행

⚠ 표1(응답자 특성)만 행 라벨이 2열(A·B)이라 배너가 C열부터 시작한다.

셀에 저장된 값은 반올림되지 않은 float(`13.986013986`)이고 화면 표기만 서식
(`###0.0`)으로 결정된다. 보고서 수치는 **서식 적용 후 값**을 써야 하므로
모든 셀에 `raw`(원값)와 `display`(서식 적용 문자열)를 함께 담는다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Callable, Iterable, Sequence

SCHEMA_VERSION = "1.0"

# ── 배너 절단선 (R-055 · R-056) ─────────────────────────────────────────────
# 절단선은 프로젝트마다 다르다. 파서 CLI 의 `--in-groups` 로 그룹명을 주고,
# 안 주면 아래 기본값(비상주)을 쓴다. 위치가 아니라 그룹명으로 자른다.
#   비상주  : 전체 · 용도 · 권역 · 관리 운영 방식        (17열, B~R)
#   TA Pool : 전체 · 건물 소유 형태 · 지역 · 연면적       (최대 18열)
BANNER_GROUPS_IN_REPORT: tuple[str, ...] = ("전체", "용도", "권역", "관리 운영 방식")

BASE_ROW_SENTINEL = "Base for %"
TOTAL_ROW_LABEL = "계"  # 원문은 공백 6칸 들여쓰기: '      계'

# 전각 렌티큘러 괄호 요약행 vs 반각 대괄호 평균행 — 괄호 종류가 다르다.
KIND_BOT = "bot"
KIND_SOSO = "soso"
KIND_TOP = "top"
KIND_MEAN = "mean"
KIND_OPTION = "option"
KIND_TOTAL = "total"
KIND_OTHER = "other"

_SENTINEL_KINDS: dict[str, str] = {
    "【BOT2】": KIND_BOT,
    "【BOT3】": KIND_BOT,
    "【SoSo】": KIND_SOSO,
    "【TOP2】": KIND_TOP,
    "【TOP3】": KIND_TOP,
}

# '[5점 평균]' / '[7점 평균]' / '[평균 : 만원]' / '[평균 : 개]' / '[평균 : 층]'
_RE_SCALE_MEAN = re.compile(r"^\[\s*(\d+)\s*점\s*평균\s*\]$")
_RE_UNIT_MEAN = re.compile(r"^\[\s*평균\s*[:：]\s*(.*?)\s*\]$")
# 대괄호를 쓰지 않는 DP 도 있다 — '평균 (개월)' · '평균'.
# 에스원 TA Pool 실측: 34표 중 20표가 이 형태라 대괄호형만 보면 `has_mean_row` 가 전량 false 가 되고,
# 그러면 `그래프 상단 Text` 수식어(R-047)가 20장에서 통째로 누락된다.
_RE_PLAIN_MEAN = re.compile(r"^평균(?:\s*[(（]\s*(.*?)\s*[)）])?$")

_RE_OPTION = re.compile(r"^(\s*)(\d+)\)\s*(.*)$")
_RE_QCODE = re.compile(r"^([A-Z]{1,3}\d+(?:-\d+)*)\.\s*(.*)$", re.DOTALL)
_RE_BASE_N = re.compile(r"^\(\s*([\d,]+)\s*\)$")
_RE_RANK_SCOPE = re.compile(r"(\d+(?:\s*~\s*\d+)?)\s*순위")
_RE_MATRIX_ITEM = re.compile(r"^(\d+)\)\s*(.*)$", re.DOTALL)


class TableParseError(RuntimeError):
    """구조가 예상과 다를 때 즉시 던진다. 부분 출력보다 실패가 낫다."""


# ── Grid ───────────────────────────────────────────────────────────────────
@dataclass
class Grid:
    """리더 독립적인 시트 표현.

    values[r][c] : str | float | int | None   (빈 셀은 None)
    formats[r][c]: str                        (Excel 표시 형식 문자열, 'General' 등)
    """

    name: str
    values: list[list[Any]]
    formats: list[list[str]]

    @property
    def nrows(self) -> int:
        return len(self.values)

    @property
    def ncols(self) -> int:
        return len(self.values[0]) if self.values else 0

    def value(self, r: int, c: int) -> Any:
        if 0 <= r < self.nrows and 0 <= c < len(self.values[r]):
            return self.values[r][c]
        return None

    def text(self, r: int, c: int) -> str:
        v = self.value(r, c)
        return v if isinstance(v, str) else ""

    def fmt(self, r: int, c: int) -> str:
        if 0 <= r < len(self.formats) and 0 <= c < len(self.formats[r]):
            return self.formats[r][c] or "General"
        return "General"


# ── 숫자 서식 ───────────────────────────────────────────────────────────────
def format_decimals(fmt: str | None) -> int | None:
    """Excel 표시 형식이 강제하는 소수 자릿수. General/텍스트면 None."""
    if not fmt:
        return None
    section = fmt.split(";")[0]
    section = re.sub(r'"[^"]*"', "", section)
    section = re.sub(r"\\.", "", section)
    section = re.sub(r"\[[^\]]*\]", "", section)
    section = section.strip()
    if not section or section.lower() == "general" or section == "@":
        return None
    if "." not in section:
        # 자릿수 지정자가 하나도 없으면 서식으로 보지 않는다.
        return 0 if re.search(r"[0#?]", section) else None
    frac = section.split(".", 1)[1]
    m = re.match(r"[0#?]*", frac)
    return len(m.group(0)) if m else 0


def _is_percent_format(fmt: str | None) -> bool:
    if not fmt:
        return False
    return "%" in re.sub(r'"[^"]*"', "", fmt)


def general_display(v: float | int) -> str:
    if isinstance(v, bool):
        return str(v)
    if float(v).is_integer():
        return str(int(v))
    return repr(float(v))


def apply_number_format(value: float | int, fmt: str | None) -> tuple[float, str]:
    """(서식 적용 후 값, 표시 문자열)을 돌려준다.

    Excel 화면 반올림은 half-up(0에서 멀어지는 방향)이라 Python `round()`의
    banker's rounding을 쓰면 안 된다. Decimal(repr(v)) + ROUND_HALF_UP 사용.
    """
    v = float(value)
    if _is_percent_format(fmt):
        v *= 100.0
    dec = format_decimals(fmt)
    if dec is None:
        return v, general_display(v)
    q = Decimal(repr(v)).quantize(Decimal(1).scaleb(-dec), rounding=ROUND_HALF_UP)
    return float(q), f"{q:f}"


# ── 라벨 분류 ───────────────────────────────────────────────────────────────
@dataclass
class RowLabelInfo:
    kind: str
    indent: int
    text: str            # 앞뒤 공백 제거한 라벨
    option_no: int | None = None
    option_text: str | None = None
    scale_points: int | None = None
    mean_unit: str | None = None


def classify_row_label(raw: Any, *, two_label_cols: bool = False) -> RowLabelInfo:
    """행 라벨 문자열 → kind ∈ {option,total,bot,soso,top,mean,other}."""
    if raw is None or raw == "":
        return RowLabelInfo(KIND_OTHER, 0, "")
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        # A1(층수)처럼 보기 라벨이 숫자로 저장된 경우 — 보기 행이다.
        return RowLabelInfo(KIND_OPTION, 0, general_display(raw))

    s = str(raw)
    stripped = s.strip()
    indent = len(s) - len(s.lstrip(" "))

    if stripped in _SENTINEL_KINDS:
        return RowLabelInfo(_SENTINEL_KINDS[stripped], indent, stripped)

    m = _RE_SCALE_MEAN.match(stripped)
    if m:
        return RowLabelInfo(KIND_MEAN, indent, stripped,
                            scale_points=int(m.group(1)), mean_unit="점")
    m = _RE_UNIT_MEAN.match(stripped)
    if m:
        return RowLabelInfo(KIND_MEAN, indent, stripped, mean_unit=m.group(1))
    m = _RE_PLAIN_MEAN.match(stripped)
    if m:
        return RowLabelInfo(KIND_MEAN, indent, stripped, mean_unit=m.group(1) or None)

    if stripped == TOTAL_ROW_LABEL:
        return RowLabelInfo(KIND_TOTAL, indent, stripped)

    m = _RE_OPTION.match(s)
    if m:
        return RowLabelInfo(KIND_OPTION, len(m.group(1)), stripped,
                            option_no=int(m.group(2)), option_text=m.group(3).strip())

    if two_label_cols:
        # 표1(응답자 특성)은 보기 번호 없이 그룹/항목 2열 라벨을 쓴다.
        return RowLabelInfo(KIND_OPTION, indent, stripped)

    return RowLabelInfo(KIND_OTHER, indent, stripped)


# ── 제목 파싱 ───────────────────────────────────────────────────────────────
@dataclass
class TitleInfo:
    full: str
    question_code: str | None
    title_base: str
    suffix: str
    suffix_parts: list[str]
    is_multi: bool
    is_ranking: bool
    rank_scope: str | None
    rank_is_first_only: bool
    summary_kind: str | None      # 'top2' | 'mean' | None
    matrix_item_no: int | None
    matrix_item_label: str | None


def parse_title(title: str) -> TitleInfo:
    full = title.strip()
    m = _RE_QCODE.match(full)
    if m:
        code, rest = m.group(1), m.group(2).strip()
    else:
        code, rest = None, full

    if "_" in rest:
        head, tail = rest.split("_", 1)
        title_base = head.strip()
        suffix = "_" + tail
        suffix_parts = [p.strip() for p in tail.split("_")]
    else:
        title_base = rest
        suffix = ""
        suffix_parts = []

    is_multi = "(중복응답)" in full
    rank_m = _RE_RANK_SCOPE.search(suffix)
    rank_scope = rank_m.group(1).replace(" ", "") if rank_m else None
    is_ranking = rank_scope is not None
    rank_is_first_only = bool(rank_scope) and "~" not in rank_scope

    summary_kind = None
    matrix_item_no = None
    matrix_item_label = None
    for part in suffix_parts:
        low = part.lower()
        if low.endswith("top2 summary") or low.endswith("top3 summary"):
            summary_kind = "top2"
        elif low.endswith("mean summary"):
            summary_kind = "mean"
        mi = _RE_MATRIX_ITEM.match(part)
        if mi and matrix_item_no is None:
            matrix_item_no = int(mi.group(1))
            matrix_item_label = mi.group(2).strip()

    return TitleInfo(
        full=full,
        question_code=code,
        title_base=title_base,
        suffix=suffix,
        suffix_parts=suffix_parts,
        is_multi=is_multi,
        is_ranking=is_ranking,
        rank_scope=rank_scope,
        rank_is_first_only=rank_is_first_only,
        summary_kind=summary_kind,
        matrix_item_no=matrix_item_no,
        matrix_item_label=matrix_item_label,
    )


# ── 열 문자 ─────────────────────────────────────────────────────────────────
def col_letter(idx: int) -> str:
    """0-based 열 인덱스 → Excel 열 문자(A, B, ..., AA)."""
    if idx < 0:
        raise ValueError(f"음수 열 인덱스: {idx}")
    s = ""
    n = idx + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


# ── 표 경계 탐지 ────────────────────────────────────────────────────────────
@dataclass
class TableSpan:
    index: int
    title_row: int
    banner_col_start: int   # 1 = B열(일반표), 2 = C열(표1 응답자 특성)
    end_row: int            # exclusive


def find_table_spans(grid: Grid) -> list[TableSpan]:
    """제목행 = A열이 문자열이고 바로 다음 행 B 또는 C열이 '전체'인 행."""
    spans: list[tuple[int, int]] = []
    for r in range(grid.nrows - 3):
        t = grid.value(r, 0)
        if not isinstance(t, str) or not t.strip():
            continue
        if grid.text(r + 1, 1) == "전체":
            spans.append((r, 1))
        elif grid.text(r + 1, 2) == "전체":
            spans.append((r, 2))
    out: list[TableSpan] = []
    for i, (r, off) in enumerate(spans):
        end = spans[i + 1][0] if i + 1 < len(spans) else grid.nrows
        out.append(TableSpan(index=i + 1, title_row=r, banner_col_start=off, end_row=end))
    return out


# ── 배너 ───────────────────────────────────────────────────────────────────
def _parse_base_n(raw: Any, *, where: str) -> int:
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return int(raw)
    m = _RE_BASE_N.match(str(raw).strip())
    if not m:
        raise TableParseError(f"{where}: Base 셀을 '(NNN)' 형태로 해석할 수 없음 → {raw!r}")
    return int(m.group(1).replace(",", ""))


def find_base_row(grid: Grid, span: TableSpan) -> int:
    """'Base for %' 행 위치. 기본 판형은 제목+3이지만, 배너 헤더와 Base 사이에 빈 행이 끼는
    판형이 있다(분기 통합 Convergence 표 실측: 제목+4 / 백산수 단독 Convergence: 제목+5).
    제목+6까지 내려가며 센티널을 찾는다."""
    limit = min(span.title_row + 7, span.end_row)
    for r in range(span.title_row + 3, limit):
        v = grid.value(r, 0)
        if isinstance(v, str) and v.strip() == BASE_ROW_SENTINEL:
            return r
    v = grid.value(span.title_row + 3, 0)
    raise TableParseError(
        f"표{span.index}(행{span.title_row}): {span.title_row + 3}행 A열이 '{BASE_ROW_SENTINEL}'이 아님 → {v!r}"
    )


def build_banners(grid: Grid, span: TableSpan,
                  in_report_groups: Sequence[str]) -> list[dict[str, Any]]:
    gr = span.title_row + 1     # 배너 그룹 헤더
    lr = span.title_row + 2     # 배너 세부 헤더
    br = find_base_row(grid, span)      # Base for % (기본 제목+3, 변형 판형은 아래로 탐색)
    off = span.banner_col_start

    cols = [c for c in range(off, grid.ncols) if grid.value(br, c) not in (None, "")]
    if not cols:
        raise TableParseError(f"표{span.index}(행{span.title_row}): 배너 열을 찾지 못함")
    if cols != list(range(cols[0], cols[0] + len(cols))):
        raise TableParseError(
            f"표{span.index}(행{span.title_row}): 배너 열이 연속적이지 않음 → {cols}"
        )

    banners: list[dict[str, Any]] = []
    current_group = ""
    for pos, c in enumerate(cols):
        g = grid.text(gr, c).strip()
        if g:
            current_group = g
        label = grid.text(lr, c).strip() or current_group
        n = _parse_base_n(grid.value(br, c), where=f"표{span.index} {col_letter(c)}{br + 1}")
        banners.append({
            "pos": pos,
            "col": c,
            "col_letter": col_letter(c),
            "group": current_group,
            "label": label,
            "n": n,
            "in_report": current_group in in_report_groups,
        })
    return banners


# ── 데이터 행 ───────────────────────────────────────────────────────────────
def build_rows(grid: Grid, span: TableSpan,
               banners: Sequence[dict[str, Any]]) -> list[dict[str, Any]]:
    off = span.banner_col_start
    two_label_cols = off == 2
    first = find_base_row(grid, span) + 1
    rows: list[dict[str, Any]] = []
    group_ff = ""     # 표1 A열 그룹 라벨 forward-fill
    pos = 0
    for r in range(first, span.end_row):
        label_cells = [grid.value(r, c) for c in range(off)]
        has_values = any(grid.value(r, c) not in (None, "") for c in range(off, grid.ncols))
        if not has_values and all(v in (None, "") for v in label_cells):
            continue  # 표 사이 빈 행
        if two_label_cols:
            g = grid.text(r, 0).strip()
            if g:
                group_ff = g
            raw_label = grid.value(r, 1)
        else:
            raw_label = grid.value(r, 0)

        info = classify_row_label(raw_label, two_label_cols=two_label_cols)
        if raw_label in (None, "") and has_values:
            # 표1의 무라벨 100% 행처럼 값만 있는 행
            info = RowLabelInfo(KIND_OTHER, 0, "")

        values: list[dict[str, Any] | None] = []
        for b in banners:
            v = grid.value(r, b["col"])
            if v is None or v == "":
                values.append(None)
                continue
            if isinstance(v, str):
                values.append({"raw": None, "display": v.strip(), "text": v.strip()})
                continue
            fmt = grid.fmt(r, b["col"])
            rounded, disp = apply_number_format(v, fmt)
            values.append({"raw": float(v), "display": disp, "value": rounded, "fmt": fmt})

        if raw_label is None or (isinstance(raw_label, str) and not raw_label.strip()):
            label_raw = None
        elif isinstance(raw_label, str):
            label_raw = raw_label
        else:
            label_raw = general_display(raw_label)

        rows.append({
            "pos": pos,
            "row": r,
            "label_raw": label_raw,
            "label": info.text,
            "label_group": (group_ff or None) if two_label_cols else None,
            "indent": info.indent,
            "kind": info.kind,
            "option_no": info.option_no,
            "option_text": info.option_text,
            "scale_points": info.scale_points,
            "mean_unit": info.mean_unit,
            "values": values,
        })
        pos += 1

    # ── 보기 행 보완 판정 (2026-08-24)
    #    통계표가 보기 번호와 들여쓰기를 모두 뺀 판형이면 위 분류로는 보기가 하나도 안 잡힌다.
    #    그때만 등장 순서로 보기를 되살린다 — 센티널(계·BOT2·SoSo·TOP2·평균) **앞**에 놓이고
    #    값을 가진 라벨 행이 보기다. 보기가 이미 있으면 건드리지 않는다(오탐 방지).
    if rows and not any(r["kind"] == KIND_OPTION for r in rows):
        first_sentinel = next((k for k, r in enumerate(rows)
                               if r["kind"] in (KIND_TOTAL, KIND_BOT, KIND_SOSO, KIND_TOP, KIND_MEAN)),
                              len(rows))
        no = 0
        for r in rows[:first_sentinel]:
            if r["kind"] != KIND_OTHER or not r.get("label_raw"):
                continue
            if all(v is None for v in (r.get("values") or [])):
                continue
            no += 1
            r["kind"] = KIND_OPTION
            r["option_no"] = no
            r["option_text"] = r["label"]
            r["option_inferred"] = True      # 번호·들여쓰기 없이 순서로 되살린 보기

    if not rows:
        raise TableParseError(f"표{span.index}(행{span.title_row}): 데이터 행이 하나도 없음")
    return rows


# ── 힌트 ───────────────────────────────────────────────────────────────────
def build_hints(title: TitleInfo, rows: Sequence[dict[str, Any]],
                banners: Sequence[dict[str, Any]]) -> dict[str, Any]:
    kinds = [r["kind"] for r in rows]
    scale_pts = next((r["scale_points"] for r in rows if r["scale_points"]), None)
    mean_units = [r["mean_unit"] for r in rows if r["kind"] == KIND_MEAN and r["mean_unit"]]
    has_scale_rows = any(k in (KIND_BOT, KIND_SOSO, KIND_TOP) for k in kinds)
    n_options = sum(1 for k in kinds if k == KIND_OPTION)
    return {
        "has_scale_rows": has_scale_rows,
        "scale_points": scale_pts or (5 if has_scale_rows else None),
        "is_ranking": title.is_ranking,
        "rank_scope": title.rank_scope,
        "rank_is_first_only": title.rank_is_first_only,
        "is_multi": title.is_multi,
        "is_summary": title.summary_kind is not None,
        "summary_kind": title.summary_kind,
        "is_matrix_item": title.matrix_item_no is not None,
        "has_total_row": KIND_TOTAL in kinds,
        "has_mean_row": KIND_MEAN in kinds,
        "mean_units": sorted(set(mean_units)),
        "n_option_rows": n_options,
        "n_rows": len(rows),
        "n_banners": len(banners),
        "n_banners_in_report": sum(1 for b in banners if b["in_report"]),
        "row_kind_counts": {k: kinds.count(k) for k in sorted(set(kinds))},
    }


# ── 표 하나 ────────────────────────────────────────────────────────────────
def build_table(grid: Grid, span: TableSpan,
                in_report_groups: Sequence[str]) -> dict[str, Any]:
    raw_title = grid.value(span.title_row, 0)
    if not isinstance(raw_title, str):
        raise TableParseError(f"표{span.index}: 제목 셀이 문자열이 아님 → {raw_title!r}")
    title = parse_title(raw_title)
    banners = build_banners(grid, span, in_report_groups)
    rows = build_rows(grid, span, banners)
    base_n = banners[0]["n"]

    first_row = rows[0]["row"]
    last_row = rows[-1]["row"]
    last_col = banners[-1]["col"]
    return {
        "index": span.index,
        "sheet": grid.name,
        "title_full": title.full,
        "question_code": title.question_code,
        "title_base": title.title_base,
        "suffix": title.suffix,
        "suffix_parts": title.suffix_parts,
        "matrix_item_no": title.matrix_item_no,
        "matrix_item_label": title.matrix_item_label,
        "base_n": base_n,
        "label_cols": list(range(span.banner_col_start)),
        "banner_col_start": span.banner_col_start,
        "source_rows": {
            "basis": "0-based sheet row index",
            "title": span.title_row,
            "banner_group": span.title_row + 1,
            "banner_label": span.title_row + 2,
            "base": first_row - 1,
            "data_first": first_row,
            "data_last": last_row,
            "block_end_exclusive": span.end_row,
        },
        "excel_range": f"A{span.title_row + 1}:{col_letter(last_col)}{last_row + 1}",
        "banners": banners,
        "rows": rows,
        "hints": build_hints(title, rows, banners),
    }


# ── 목차 ───────────────────────────────────────────────────────────────────
def build_index(index_grid: Grid | None,
                tables: Sequence[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[str]]:
    """`목 차` 시트를 읽어 Table 시트에서 찾은 표와 대조한다."""
    warnings: list[str] = []
    entries: list[dict[str, Any]] = []
    if index_grid is None:
        warnings.append("목차 시트를 찾지 못해 교차검증을 건너뜀")
        return entries, warnings

    for r in range(index_grid.nrows):
        no = index_grid.value(r, 0)
        title = index_grid.value(r, 1)
        if not isinstance(title, str) or not title.strip():
            continue
        if not isinstance(no, (int, float)) or isinstance(no, bool):
            continue  # '목  차' 헤더행
        entries.append({"no": int(no), "title": title.strip(), "row": r})

    by_no = {t["index"]: t for t in tables}
    if len(entries) != len(tables):
        warnings.append(f"목차 항목 {len(entries)}개 vs Table 시트 표 {len(tables)}개 — 개수 불일치")
    for e in entries:
        t = by_no.get(e["no"])
        if t is None:
            e["match"] = "missing_in_table"
            warnings.append(f"목차 #{e['no']} '{e['title']}' 에 대응하는 표가 Table 시트에 없음")
        elif t["title_full"] != e["title"]:
            e["match"] = "title_mismatch"
            e["table_title"] = t["title_full"]
            warnings.append(
                f"목차 #{e['no']} 제목 불일치: 목차={e['title']!r} / Table={t['title_full']!r}"
            )
        else:
            e["match"] = "ok"
    matched = {e["no"] for e in entries}
    for t in tables:
        if t["index"] not in matched:
            warnings.append(f"Table 시트 표{t['index']} '{t['title_full']}' 이 목차에 없음")
    return entries, warnings


# ── 최상위 빌더 ────────────────────────────────────────────────────────────
def build_document(table_grid: Grid,
                   index_grid: Grid | None,
                   *,
                   source: str,
                   source_format: str,
                   parser: str,
                   in_report_groups: Sequence[str] = BANNER_GROUPS_IN_REPORT,
                   extra_meta: dict[str, Any] | None = None,
                   skip_bad_tables: bool = False) -> dict[str, Any]:
    spans = find_table_spans(table_grid)
    if not spans:
        raise TableParseError(
            f"'{table_grid.name}' 시트에서 표를 하나도 찾지 못했습니다. "
            "제목행 다음 행의 B 또는 C열이 '전체'인 구조인지 확인하세요."
        )
    tables: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    for s in spans:
        try:
            tables.append(build_table(table_grid, s, in_report_groups))
        except TableParseError as e:
            if not skip_bad_tables:
                raise
            skipped.append({
                "index": s.index,
                "title_row": s.title_row,
                "title": str(table_grid.value(s.title_row, 0) or "").strip(),
                "error": str(e),
            })
    if skip_bad_tables and not tables:
        raise TableParseError("--skip-bad: 전 표가 스킵되어 남은 표가 없습니다.")
    entries, warnings = build_index(index_grid, tables)
    warnings += [f"표{k['index']} 스킵({k['title'][:40]}): {k['error']}" for k in skipped]

    meta: dict[str, Any] = {
        "schema_version": SCHEMA_VERSION,
        "parser": parser,
        "source": source,
        "source_format": source_format,
        "sheet_table": table_grid.name,
        "sheet_index": index_grid.name if index_grid else None,
        "n_tables": len(tables),
        "banner_groups_in_report": list(in_report_groups),
        "banner_cut_rule": "그룹명 화이트리스트 (R-056) — IN: "
                           + " · ".join(in_report_groups)
                           + " / 그 밖의 배너군은 전부 OUT",
        "value_note": "raw=저장된 무한소수, display=셀 서식(###0.0 등) 적용 후 문자열, value=서식 적용 후 float",
        "warnings": warnings,
        "skipped_tables": skipped,
    }
    if extra_meta:
        meta.update(extra_meta)
    return {"meta": meta, "index": entries, "tables": tables}


# ── 요약 출력 ──────────────────────────────────────────────────────────────
def summarize(doc: dict[str, Any]) -> str:
    tables = doc["tables"]
    lines = [
        f"표 {len(tables)}장 파싱 완료 (목차 항목 {len(doc['index'])}개)",
        f"  척도표(BOT2/SoSo/TOP2 보유) : {sum(1 for t in tables if t['hints']['has_scale_rows'])}장",
        f"  순위표                      : {sum(1 for t in tables if t['hints']['is_ranking'])}장",
        f"  중복응답표                  : {sum(1 for t in tables if t['hints']['is_multi'])}장",
        f"  요약표(TOP2/Mean Summary)   : {sum(1 for t in tables if t['hints']['is_summary'])}장",
    ]
    ns = sorted({t["hints"]["n_banners"] for t in tables})
    lines.append(f"  배너 열 수 분포             : {ns}")
    small = [t for t in tables if t["base_n"] < 30]
    if small:
        lines.append(f"  ⚠ n<30 표 {len(small)}장: " +
                     ", ".join(f"표{t['index']}(n={t['base_n']})" for t in small))
    if doc["meta"]["warnings"]:
        lines.append(f"  ⚠ 경고 {len(doc['meta']['warnings'])}건:")
        lines += [f"     - {w}" for w in doc["meta"]["warnings"]]
    return "\n".join(lines)
