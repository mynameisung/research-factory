# -*- coding: utf-8 -*-
"""
render_guide_blocks.py — extract_guide_blocks 산출 JSON → 사람이 읽는 마크다운 격자

  python engine/tools/render_guide_blocks.py <gb.json> <out.md> [--from N] [--to M] [--title 제목]

요약·결론 블록처럼 **다단 패널이 가로로 늘어선 구간**을 원래 셀 배치 그대로 보여 준다.
엑셀을 열지 않고 「어느 열에 무엇이 있었나」를 읽게 하는 것이 목적이다(코퍼스 분석 2026-08-24).
값은 원본 그대로 옮기되 부동소수 잔재(59.285714…)는 소수 1~2자리로 줄여 읽기 쉽게 한다.
"""
from __future__ import annotations
import argparse, json
from pathlib import Path


def fmt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        s = f"{v:.2f}".rstrip("0").rstrip(".")
        return s
    return str(v).replace("\n", "⏎").replace("\t", " ")


def render_block(b: dict) -> str:
    out = [f"### 블록 {b['block_index']} — {b.get('slide_title') or '(제목 없음)'}"
           + (f" | {b['sub_title']}" if b.get("sub_title") else "")]
    meta = []
    if b.get("chart_type"):
        meta.append(f"Chart type: `{fmt(b['chart_type'])}`")
    if b.get("chart_types_all") and len(b["chart_types_all"]) > 1:
        meta.append("Chart type(전체): " + " · ".join(f"`{fmt(c)}`" for c in b["chart_types_all"]))
    if b.get("base_units_all"):
        meta.append("Base/Unit: " + " · ".join(f"`{fmt(x)}`" for x in b["base_units_all"]))
    if b.get("question_reference"):
        meta.append(f"Question/Reference: `{fmt(b['question_reference'])}`")
    meta.append(f"행 {b['start_row']}~{b['end_row']} · 가로 패널 {b.get('horizontal_data_panels')} · 표 {b.get('table_count_in_block')}")
    out.append("- " + "\n- ".join(meta))
    if b.get("ments"):
        out.append("\n**멘트**")
        for m in b["ments"]:
            out.append(f"  - {fmt(m)}")
    if b.get("layout_directives"):
        out.append("\n**레이아웃 지시**: " + " · ".join(fmt(x) for x in b["layout_directives"]))
    rows = b.get("body_rows") or []
    if rows:
        out.append("\n```")
        for r in rows:
            cells = r.get("cells") or {}
            if not cells:
                continue
            line = "  ".join(f"{col}={fmt(v)}" for col, v in cells.items())
            out.append(f"r{r['row']:>5}  {line}")
        out.append("```")
    return "\n".join(out)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("gb", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument("--from", dest="a", type=int, default=0)
    ap.add_argument("--to", dest="b", type=int, default=10**9)
    ap.add_argument("--title", default="")
    a = ap.parse_args()
    doc = json.loads(a.gb.read_text(encoding="utf-8"))
    blocks = [x for x in doc["blocks"] if a.a <= x["block_index"] <= a.b]
    head = [f"# {a.title or a.gb.stem}", "",
            f"원본: `{doc.get('source_file')}` 시트 `{doc.get('sheet')}` (라벨 열 {doc.get('label_col')})",
            f"블록 {a.a}~{a.b} 중 {len(blocks)}개", ""]
    a.out.write_text("\n".join(head + [render_block(b) for b in blocks]) + "\n",
                     encoding="utf-8", newline="\n")
    print(f"→ {a.out}  ({len(blocks)} blocks)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
