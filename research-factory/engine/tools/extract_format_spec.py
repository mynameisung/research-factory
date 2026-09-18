#!/usr/bin/env python3
"""예시 제작가이드 → `format_spec.yaml`.

**왜 추출하나**
서식 기준을 코드에 박으면 템플릿이 바뀌는 순간 거짓말이 된다. LG 의 `s="178"` 은
그 파일의 스타일 인덱스일 뿐이라 다른 파일에서는 아무 의미가 없다.
그래서 기준을 **예시에서 실측해 뽑아 두고**, 빌더와 검수기가 그 파일을 계약으로 공유한다.
(골든셋에서 정답을 제작본에서 역산한 것과 같은 구조다.)

**무엇을 뽑나**
- 시트 스켈레톤 — 목록·순서·상태·탭색·본문 시트명
- 블록 골격 — 라벨 6종·라벨 열·값 열·블록 간격
- **행 서식** — `Slide title` 행과 나머지 라벨 행의 행 스타일(`s` + `customFormat` + 높이)
  ★ 이번 형광색 사고의 정체가 이것이다. 노랑은 셀이 아니라 **행 전체**에 걸려 있다
- 블록 종류 — `Chart type` 값으로 판별되는 표지/목차/간지/본문
- 표 서식 — 패널제목·배너헤더·Base 행·데이터 행의 테두리·정렬·병합
- 구조 요구 — 반드시 있어야 할 블록 종류, 목차 시트와 안내문

사용법
------
    python extract_format_spec.py <예시.xlsm> <format_spec.yaml> [--sheet 본문_제작용]
"""

from __future__ import annotations

import argparse
import io
import json
import re
import sys
import warnings
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

warnings.filterwarnings("ignore")
import openpyxl  # noqa: E402

LABELS = ["Slide title", "Sub title", "Chart type",
          "Base/Unit", "Question/Reference", "Data"]

#: 구조 블록은 `Chart type` 에 이 값이 들어간다 (자유 서술인 본문과 구분된다)
STRUCTURE_KINDS = ("표지", "목차", "간지")


def sheet_part(zf: zipfile.ZipFile, name: str) -> str | None:
    """시트 이름 → `xl/worksheets/sheetN.xml` 경로.

    `workbook.xml` 의 `<sheet>` 는 속성값에 `http://` 가 들어갈 수 있어
    `[^/>]*` 류 정규식이 조용히 빗나간다. 속성만 따로 파싱한다.
    """
    wb = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rmap = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
    for m in re.finditer(r"<sheet\b(.*?)/>", wb, re.S):
        attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
        if attrs.get("name") == name:
            tgt = rmap.get(attrs.get("r:id", ""))
            if not tgt:
                return None
            return ("xl/" + tgt.lstrip("/")) if not tgt.startswith("xl/") else tgt
    return None


def row_attrs(xml: str, row: int) -> dict[str, str]:
    m = re.search(r'<row r="%d"([^>]*)>' % row, xml)
    return dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1))) if m else {}


def fill_of(cell) -> str | None:
    f = cell.fill
    try:
        if f is not None and getattr(f, "patternType", None) and f.fgColor is not None:
            rgb = f.fgColor.rgb
            return rgb if isinstance(rgb, str) else None
    except AttributeError:
        pass  # GradientFill 등은 patternType 이 없다
    return None


def border_of(cell) -> str:
    b = cell.border
    return "".join(k[0] for k in ("left", "right", "top", "bottom")
                   if getattr(b, k) and getattr(b, k).style)


def cell_style(cell) -> dict[str, Any]:
    fo = cell.font
    return {
        "fill": fill_of(cell),
        "bold": bool(fo.bold),
        "size": float(fo.sz) if fo.sz else None,
        "border": border_of(cell) or None,
        "align": cell.alignment.horizontal,
    }


def find_blocks(ws, label_col: int) -> list[int]:
    return [r for r in range(1, ws.max_row + 1)
            if ws.cell(r, label_col).value == LABELS[0]]


def detect_columns(ws) -> tuple[int, int]:
    """라벨 열 / 값 열을 찾는다. LG 는 A/B, SKEE 는 B/C 다."""
    for lab in (1, 2, 3):
        hits = sum(1 for r in range(1, min(ws.max_row, 400) + 1)
                   if ws.cell(r, lab).value == LABELS[0])
        if hits:
            return lab, lab + 1
    raise SystemExit("[중단] `Slide title` 라벨 열을 찾지 못했습니다.")


def main() -> int:
    ap = argparse.ArgumentParser(description="예시 제작가이드 → format_spec.yaml")
    ap.add_argument("example", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument("--sheet", default=None, help="본문 시트명 (생략 시 자동 탐색)")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.example, keep_vba=True)
    body_name = args.sheet or next(
        (n for n in wb.sheetnames if n.startswith("본문")), None)
    if not body_name:
        raise SystemExit("[중단] 본문 시트를 찾지 못했습니다.")
    ws = wb[body_name]
    label_col, value_col = detect_columns(ws)
    blocks = find_blocks(ws, label_col)
    if not blocks:
        raise SystemExit("[중단] 블록을 찾지 못했습니다.")

    zf = zipfile.ZipFile(args.example)
    part = sheet_part(zf, body_name)
    xml = zf.read(part).decode("utf-8") if part else ""

    # ── 행 서식: Slide title 행 vs 나머지 라벨 행
    #    이게 노랑 띠의 정체다. 셀이 아니라 행에 걸린다.
    title_rows = [row_attrs(xml, r) for r in blocks[:20]]
    other_rows = [row_attrs(xml, r + i) for r in blocks[:20] for i in range(1, 6)]

    def dominant(rows: list[dict], key: str) -> str | None:
        vals = [r.get(key) for r in rows if r.get(key) is not None]
        return Counter(vals).most_common(1)[0][0] if vals else None

    row_fmt = {
        "slide_title": {
            "s": dominant(title_rows, "s"),
            "customFormat": dominant(title_rows, "customFormat"),
            "ht": dominant(title_rows, "ht"),
            "cell": cell_style(ws.cell(blocks[0], label_col)),
        },
        "other_label": {
            "s": dominant(other_rows, "s"),
            "customFormat": dominant(other_rows, "customFormat"),
            "ht": dominant(other_rows, "ht"),
            "cell": cell_style(ws.cell(blocks[0] + 1, label_col)),
        },
    }

    # ── 블록 간격 · 종류 분포
    gaps = Counter(b - a for a, b in zip(blocks, blocks[1:]))
    kinds = Counter()
    kind_examples: dict[str, int] = {}
    for r in blocks:
        ct = str(ws.cell(r + 2, value_col).value or "").strip()
        k = ct if ct in STRUCTURE_KINDS else ("본문" if ct and ct != "-" else "미지정")
        kinds[k] += 1
        kind_examples.setdefault(k, r)

    # ── 표 서식: 본문 블록 하나를 골라 행 종류별 셀 서식을 뜬다
    body_row = kind_examples.get("본문")
    table_fmt: dict[str, Any] = {}
    if body_row:
        rr = body_row + 6
        end = next((b for b in blocks if b > body_row), ws.max_row)
        seen: dict[str, int] = {}
        while rr < end:
            lab = str(ws.cell(rr, value_col + 1).value or ws.cell(rr, value_col).value or "").strip()
            if lab == "Base for %":
                seen.setdefault("base_row", rr)
            elif lab.startswith("【") or lab.startswith("["):
                seen.setdefault("summary_row", rr)
            elif lab.strip() == "계":
                seen.setdefault("total_row", rr)
            elif re.match(r"^[A-Z]+[\d-]*\.", lab):
                seen.setdefault("panel_title", rr)
            elif lab and "panel_title" in seen and "data_row" not in seen \
                    and "base_row" in seen and rr > seen["base_row"]:
                seen.setdefault("data_row", rr)
            rr += 1
        for k, r in seen.items():
            table_fmt[k] = {
                "row_offset_from_block": r - body_row,
                "label_cell": cell_style(ws.cell(r, value_col + 1)),
                "first_data_cell": cell_style(ws.cell(r, value_col + 2)),
            }
        # 배너 헤더는 Base 행 바로 위 1~2줄
        if "base_row" in seen:
            b = seen["base_row"]
            table_fmt["banner_header"] = {
                "rows_above_base": 2,
                "cell": cell_style(ws.cell(b - 1, value_col + 2)),
                "merged_ranges_sample": [str(m) for m in ws.merged_cells.ranges
                                         if m.min_row in (b - 2, b - 1)][:6],
            }

    spec = {
        "_설명": "예시 제작가이드에서 실측한 서식 기준. 빌더와 검수기가 이 파일을 계약으로 공유한다. "
                 "템플릿이나 클라이언트가 바뀌면 다시 뽑아라 — 스타일 인덱스는 파일마다 다르다.",
        "source": args.example.name,
        "body_sheet": body_name,
        "sheet_skeleton": [
            {"name": s.title, "state": s.sheet_state,
             "tab_color": (s.sheet_properties.tabColor.rgb
                           if s.sheet_properties.tabColor
                           and isinstance(s.sheet_properties.tabColor.rgb, str) else None)}
            for s in wb.worksheets
        ],
        "block": {
            "labels": LABELS,
            "label_col": label_col,
            "value_col": value_col,
            "count": len(blocks),
            "first_row": blocks[0],
            "common_gap": gaps.most_common(3),
            "row_format": row_fmt,
        },
        "block_kinds": {
            "counts": dict(kinds),
            "structure_kinds": list(STRUCTURE_KINDS),
            "required": [k for k in STRUCTURE_KINDS if kinds.get(k)],
        },
        "table_format": table_fmt,
        "toc_sheet": {
            "name": next((n for n in wb.sheetnames if n.startswith("목차(")), None),
            "note_cell": "C1",
            "note": None,
        },
    }
    tn = spec["toc_sheet"]["name"]
    if tn:
        spec["toc_sheet"]["note"] = wb[tn]["C1"].value

    args.out.parent.mkdir(parents=True, exist_ok=True)
    try:
        import yaml  # type: ignore
        text = yaml.safe_dump(spec, allow_unicode=True, sort_keys=False, width=100)
    except ImportError:
        text = json.dumps(spec, ensure_ascii=False, indent=2)
    with io.open(args.out, "w", encoding="utf-8") as fh:
        fh.write(text)

    print(f"예시 {args.example.name} · 시트 {body_name}")
    print(f"  블록 {len(blocks)}개 · 라벨 열 {label_col} / 값 열 {value_col}")
    print(f"  행 서식  Slide title: s={row_fmt['slide_title']['s']} "
          f"ht={row_fmt['slide_title']['ht']} fill={row_fmt['slide_title']['cell']['fill']}")
    print(f"           나머지 라벨: s={row_fmt['other_label']['s']}")
    print(f"  블록 종류: {dict(kinds)}")
    print(f"  표 서식 행: {list(table_fmt)}")
    print(f"\n→ {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
