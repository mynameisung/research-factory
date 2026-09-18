"""통계표 .xlsx/.xlsm → tables.json.

    python parse_table_xlsx.py <input.xlsx> <output.json> [--strict] [--quiet]

`parse_table_biff.py`와 **출력 스키마가 완전히 동일**하다. 차이는 리더뿐이며
표 구조 해석·배너 절단·행 분류·서식 적용은 전부 `parse_table_common`이 한다.

TA Pool 통계표가 .xlsx로 들어올 경우를 대비한 대칭 도구다.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import sys
from pathlib import Path
from typing import Any

try:
    from parse_table_common import (
        BANNER_GROUPS_IN_REPORT,
        Grid,
        TableParseError,
        build_document,
        summarize,
    )
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from parse_table_common import (  # type: ignore
        BANNER_GROUPS_IN_REPORT,
        Grid,
        TableParseError,
        build_document,
        summarize,
    )

ZIP_MAGIC = b"PK\x03\x04"
INDEX_SHEET_HINTS = ("목", "차", "index", "toc", "contents")
TABLE_SHEET_HINTS = ("table", "표")


def read_grids(path: Path) -> tuple[dict[str, Grid], str]:
    """워크북 전체를 시트명 → Grid 로 읽는다."""
    try:
        import openpyxl  # type: ignore
    except ImportError as e:  # pragma: no cover
        raise TableParseError(
            f"openpyxl 이 필요합니다 (`pip install openpyxl`). 원인: {e}"
        ) from e

    with path.open("rb") as fh:
        head = fh.read(4)
    if head != ZIP_MAGIC:
        raise TableParseError(
            f"{path.name}: OOXML(zip) 파일이 아닙니다 (선두 {head.hex().upper()}). "
            "BIFF8 .xls 라면 parse_table_biff.py 를 쓰세요."
        )

    wb = openpyxl.load_workbook(str(path), data_only=True, rich_text=False)
    grids: dict[str, Grid] = {}
    n_formula_holes = 0
    for ws in wb.worksheets:
        values: list[list[Any]] = []
        formats: list[list[str]] = []
        for row in ws.iter_rows():
            vrow: list[Any] = []
            frow: list[str] = []
            for cell in row:
                v = cell.value
                if isinstance(v, str):
                    v = v if v != "" else None
                elif v is not None and not isinstance(v, (int, float, bool)):
                    v = str(v)          # datetime 등은 문자열로 떨어뜨린다
                if v is None and cell.data_type == "f":
                    n_formula_holes += 1
                vrow.append(v)
                frow.append(cell.number_format or "General")
            values.append(vrow)
            formats.append(frow)
        grids[ws.title] = Grid(name=ws.title, values=values, formats=formats)
    wb.close()

    if n_formula_holes:
        raise TableParseError(
            f"{path.name}: 수식 셀 {n_formula_holes}개의 캐시된 결과값이 없습니다. "
            "Excel에서 한 번 열어 저장한 뒤 다시 시도하세요 (openpyxl은 수식을 계산하지 않음)."
        )
    return grids, f"openpyxl {openpyxl.__version__}"


def pick_sheets(grids: dict[str, Grid]) -> tuple[Grid, Grid | None]:
    """(본표 시트, 목차 시트)를 고른다."""
    table = next((g for g in grids.values()
                  if any(h in g.name.lower() for h in TABLE_SHEET_HINTS)), None)
    if table is None:
        table = max(grids.values(), key=lambda g: g.nrows)
    index = next((g for g in grids.values()
                  if g is not table
                  and any(h in g.name.lower() for h in INDEX_SHEET_HINTS)), None)
    return table, index


def parse(path: str | Path, *,
          in_report_groups: tuple[str, ...] = BANNER_GROUPS_IN_REPORT) -> dict[str, Any]:
    p = Path(path)
    if not p.is_file():
        raise TableParseError(f"입력 파일이 없습니다: {p}")
    grids, reader = read_grids(p)
    table_grid, index_grid = pick_sheets(grids)
    return build_document(
        table_grid, index_grid,
        source=str(p.resolve()),
        source_format="ooxml-xlsx",
        parser=Path(__file__).name,
        in_report_groups=in_report_groups,
        extra_meta={
            "reader": reader,
            "sheets": list(grids),
            "generated_at": _dt.datetime.now().isoformat(timespec="seconds"),
        },
    )


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="통계표 .xlsx/.xlsm → tables.json")
    ap.add_argument("input", help="통계표 .xlsx 경로")
    ap.add_argument("output", help="출력 tables.json 경로")
    ap.add_argument("--strict", action="store_true",
                    help="목차 교차검증 경고도 실패로 취급")
    ap.add_argument("--in-groups", default="",
                    help="보고서에 싣는 배너군 이름을 세미콜론으로 구분해 준다 (R-056 그룹명 화이트리스트). "
                         "예: \"전체;건물 소유 형태;지역;연면적\". 생략하면 비상주 기본값을 쓴다")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    in_groups = tuple(g.strip() for g in args.in_groups.split(";") if g.strip()) \
        or BANNER_GROUPS_IN_REPORT

    try:
        doc = parse(args.input, in_report_groups=in_groups)
    except TableParseError as e:
        print(f"[parse_table_xlsx] 파싱 실패: {e}", file=sys.stderr)
        return 2

    warnings = doc["meta"]["warnings"]
    if args.strict and warnings:
        print("[parse_table_xlsx] --strict: 경고가 있어 출력하지 않습니다.", file=sys.stderr)
        for w in warnings:
            print(f"  - {w}", file=sys.stderr)
        return 3

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    if not args.quiet:
        print(summarize(doc))
        print(f"→ {out}  ({out.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
