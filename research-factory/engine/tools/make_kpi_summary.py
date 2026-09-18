#!/usr/bin/env python3
"""누락된 척도 요약표(TOP2 Summary / Mean Summary)를 통계표에서 직접 만든다.

왜 필요한가
-----------
연구원은 테이블 가이드 단계에서 **척도 문항이 연속 5개 이상이면** DP에 요약표를 함께
요청해 둔다(녹취 51:26 "척도 문항이 한 5개 이상 있으면 위에 탑2 민을 넣어요"). 요약·결론
장표는 그 요약표 한 장으로 만들기 때문이다.

그런데 요청이 빠지는 경우가 있다. 인수인계에서도 짚었다 —
    "만약에 이걸 서연이가 누락했었으면 이거 가지고 따로 만들었어야겠어"
    "이거 가지고 탑두값 하나씩 긁어 붙여 놓으셔야 되죠"   (녹취 50:44~50:53)

실제로 비상주 통계표가 그 상태다. 매트릭스 척도(A8·D6)에는 요약표 4장이 있지만,
**단일 5점 척도 문항들**(컨셉 이해도→유용성→신뢰도→문제해결→매력도→이용의향→
가격 후 이용의향→추천의향→차별성)을 묶는 요약표는 없다. 이 라인이 곧 신상품 수용도
KPI이고 요약·결론의 핵심 소재라서, 없으면 결론을 쓸 재료가 없다.

이 스크립트가 각 표의 【TOP2】·[N점 평균] 행을 긁어 그 요약표를 만든다. 손으로 하면
9문항 × 17배너 = 306칸을 옮겨 적어야 하는 일이다.

출력
----
1. `<out>.json` — tables.json 과 **동일한 스키마**의 표 2장. build_facts.py 가 그대로 먹는다.
2. `<out>.xlsx` — 사람이 눈으로 확인하고 제작가이드에 붙여넣을 수 있는 형태.

사용법
------
    python make_kpi_summary.py <tables.json> <out_prefix>
    python make_kpi_summary.py <tables.json> <out_prefix> --min-group 5
"""

from __future__ import annotations

import argparse
import io
import json
import sys
from pathlib import Path
from typing import Any

# 이 kind 를 가진 행에서 값을 긁는다.
TOP_KIND = "top"
MEAN_KIND = "mean"


def load_tables(path: Path) -> dict[str, Any]:
    with io.open(path, encoding="utf-8") as fh:
        return json.load(fh)


def single_scale_tables(tables: list[dict]) -> list[dict]:
    """요약표를 만들 대상 = 요약표도 매트릭스 항목도 아닌 순수 척도 문항."""
    out = []
    for t in tables:
        h = t.get("hints") or {}
        if not h.get("has_scale_rows"):
            continue
        if h.get("is_summary") or h.get("is_matrix_item"):
            continue
        out.append(t)
    return out


def already_summarized(tables: list[dict]) -> set[str]:
    """이미 `_TOP2 Summary` 가 있는 문항코드. 중복 생성을 막는다."""
    codes = set()
    for t in tables:
        if (t.get("hints") or {}).get("is_summary"):
            code = t.get("question_code")
            if code:
                codes.add(code)
    return codes


def row_of_kind(table: dict, kind: str) -> dict | None:
    for r in table.get("rows") or []:
        if r.get("kind") == kind:
            return r
    return None


def banner_key(table: dict, in_report_only: bool = True) -> list[tuple[str, str]]:
    return [
        (b["group"], b["label"])
        for b in table["banners"]
        if b.get("in_report") or not in_report_only
    ]


def build_summary(
    sources: list[dict],
    kind: str,
    title: str,
    index: int,
    in_report_only: bool = True,
) -> dict:
    """소스 표들에서 한 종류의 행(top 또는 mean)을 긁어 요약표 하나를 만든다."""
    ref = sources[0]
    ref_key = banner_key(ref, in_report_only)

    banners = [b for b in ref["banners"] if b.get("in_report") or not in_report_only]
    banners = [dict(b, pos=i) for i, b in enumerate(banners)]

    rows = []
    for n, src in enumerate(sources, start=1):
        # 배너 구성이 표마다 다를 수 있다(비상주 통계표는 15~45열로 가변).
        # 어긋난 채로 값을 옮기면 숫자가 다른 배너 칸에 들어간다 — 조용히 넘어가면 안 된다.
        if banner_key(src, in_report_only) != ref_key:
            raise SystemExit(
                f"[중단] 배너 구성 불일치: 표{src['index']} {src['question_code']} 가 "
                f"표{ref['index']} {ref['question_code']} 와 배너가 다릅니다.\n"
                f"  기준: {[l for _, l in ref_key]}\n"
                f"  실제: {[l for _, l in banner_key(src, in_report_only)]}\n"
                f"  → 요약표를 만들 수 없습니다. 배너가 같은 문항끼리만 묶으세요."
            )

        src_row = row_of_kind(src, kind)
        if src_row is None:
            raise SystemExit(
                f"[중단] 표{src['index']} {src['question_code']} 에 '{kind}' 행이 없습니다. "
                f"척도 문항으로 잡혔지만 요약 행이 빠졌습니다 — 통계표를 확인하세요."
            )

        values = [
            v
            for v, b in zip(src_row["values"], src["banners"])
            if b.get("in_report") or not in_report_only
        ]
        rows.append(
            {
                "label_raw": f"{n}) {src['title_base']}",
                "label": f"{n}) {src['title_base']}",
                "label_group": None,
                "indent": 0,
                "kind": "option",
                "option_no": n,
                "values": values,
                # 이 행이 어느 표에서 왔는지 — 검증과 회고에서 되짚을 수 있어야 한다.
                "source_table_index": src["index"],
                "source_question_code": src["question_code"],
            }
        )

    scale_points = (ref.get("hints") or {}).get("scale_points")
    return {
        "index": index,
        "sheet": "(생성)",
        "title_full": title,
        "question_code": None,
        "title_base": title,
        "suffix": "TOP2 Summary" if kind == TOP_KIND else "Mean Summary",
        "suffix_parts": ["TOP2 Summary" if kind == TOP_KIND else "Mean Summary"],
        "matrix_item_no": None,
        "matrix_item_label": None,
        "base_n": ref.get("base_n"),
        "label_cols": [0],
        "banner_col_start": 1,
        "source_rows": None,
        "excel_range": None,
        "generated_by": "make_kpi_summary.py",
        "generated_from": [s["index"] for s in sources],
        "banners": banners,
        "rows": rows,
        "hints": {
            "has_scale_rows": False,
            "scale_points": scale_points,
            "is_ranking": False,
            "is_multi": False,
            "is_summary": True,
            # R-049: summary_kind 도메인은 top2 | top3 | mean ('top' 아님). 7점 척도면 top3
            "summary_kind": (("top3" if scale_points == 7 else "top2") if kind == TOP_KIND else "mean"),
            "is_matrix_item": False,
            "has_total_row": False,
            "has_mean_row": kind == MEAN_KIND,
            "n_option_rows": len(rows),
            "n_rows": len(rows),
            "n_banners": len(banners),
            "n_banners_in_report": sum(1 for b in banners if b.get("in_report")),
        },
    }


def write_xlsx(summaries: list[dict], path: Path) -> None:
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
    except ImportError:
        print("openpyxl 이 없어 xlsx 를 건너뜁니다 (json 은 정상 생성).", file=sys.stderr)
        return

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    head = Font(bold=True)
    fill = PatternFill("solid", fgColor="EEEEEE")

    for s in summaries:
        ws = wb.create_sheet(s["suffix"][:31])
        ws.cell(1, 1, s["title_full"]).font = Font(bold=True, size=12)

        # 배너 2단 헤더 — 통계표 원본과 같은 모양으로 둔다(제작팀이 그대로 읽는다).
        r_group, r_label, r_base = 3, 4, 5
        prev, start = None, None
        for i, b in enumerate(s["banners"]):
            c = 2 + i
            if b["group"] != prev:
                if prev is not None and start is not None and c - 1 > start:
                    ws.merge_cells(start_row=r_group, start_column=start, end_row=r_group, end_column=c - 1)
                ws.cell(r_group, c, b["group"]).font = head
                prev, start = b["group"], c
            ws.cell(r_label, c, b["label"]).font = head
            ws.cell(r_base, c, f"({b['n']})")
            for r in (r_group, r_label, r_base):
                ws.cell(r, c).alignment = Alignment(horizontal="center", wrap_text=True)
        if start is not None and 1 + len(s["banners"]) > start:
            ws.merge_cells(start_row=r_group, start_column=start, end_row=r_group,
                           end_column=1 + len(s["banners"]))
        ws.cell(r_base, 1, "Base for %").font = head
        for r in (r_group, r_label, r_base):
            ws.cell(r, 1).fill = fill

        fmt = "0.00" if s["hints"]["summary_kind"] == "mean" else "0.0"
        for j, row in enumerate(s["rows"]):
            r = r_base + 1 + j
            ws.cell(r, 1, row["label"])
            for i, v in enumerate(row["values"]):
                cell = ws.cell(r, 2 + i, v.get("value") if isinstance(v, dict) else v)
                cell.number_format = fmt
                cell.alignment = Alignment(horizontal="center")

        ws.column_dimensions["A"].width = 46
        for i in range(len(s["banners"])):
            ws.column_dimensions[ws.cell(1, 2 + i).column_letter].width = 11
        ws.freeze_panes = ws.cell(r_base + 1, 2)

    wb.save(path)


def main() -> int:
    ap = argparse.ArgumentParser(description="통계표에서 누락된 척도 요약표를 생성한다")
    ap.add_argument("tables_json", type=Path)
    ap.add_argument("out_prefix", type=Path, help="확장자 없이. .json 과 .xlsx 두 개가 생긴다")
    ap.add_argument("--min-group", type=int, default=5,
                    help="이 개수 미만이면 요약표를 만들지 않는다 (기본 5, 녹취 51:34)")
    ap.add_argument("--title", default="주요 척도 문항")
    ap.add_argument("--all-banners", action="store_true",
                    help="보고서 제외 배너까지 포함 (기본은 in_report 만)")
    args = ap.parse_args()

    doc = load_tables(args.tables_json)
    tables = doc["tables"] if isinstance(doc, dict) else doc

    singles = single_scale_tables(tables)
    done = already_summarized(tables)
    targets = [t for t in singles if t.get("question_code") not in done]

    print(f"척도 문항 {len(singles)}개 중 요약표 없는 것 {len(targets)}개")
    for t in targets:
        print(f"  표{t['index']:>3}  {t['question_code']:<8} {t['title_base']}")

    if len(targets) < args.min_group:
        print(f"\n{args.min_group}개 미만이라 요약표를 만들지 않습니다 "
              f"(척도 문항이 몇 개 없으면 요약표가 필요 없다 — 녹취 51:26).")
        return 0

    in_report_only = not args.all_banners
    summaries = [
        build_summary(targets, TOP_KIND, f"{args.title}_TOP2 Summary",
                      index=9001, in_report_only=in_report_only),
        build_summary(targets, MEAN_KIND, f"{args.title}_Mean Summary",
                      index=9002, in_report_only=in_report_only),
    ]

    out_json = args.out_prefix.with_suffix(".json")
    out_xlsx = args.out_prefix.with_suffix(".xlsx")
    out_json.parent.mkdir(parents=True, exist_ok=True)

    with io.open(out_json, "w", encoding="utf-8") as fh:
        json.dump(
            {
                "meta": {
                    "generated_by": "make_kpi_summary.py",
                    "source": str(args.tables_json),
                    "note": "통계표에 없던 요약표. tables.json 과 동일 스키마이므로 병합해 쓸 수 있다.",
                },
                "tables": summaries,
            },
            fh, ensure_ascii=False, indent=1,
        )
    write_xlsx(summaries, out_xlsx)

    n_b = summaries[0]["hints"]["n_banners"]
    print(f"\n생성: {len(targets)}문항 × {n_b}배너 = {len(targets)*n_b}칸 × 2종")
    print(f"  {out_json}")
    print(f"  {out_xlsx}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
