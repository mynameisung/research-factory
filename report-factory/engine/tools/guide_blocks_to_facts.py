# -*- coding: utf-8 -*-
"""
guide_blocks_to_facts.py — extract_guide_blocks 산출 JSON → `slide_facts.json` 스키마 (벤치마크 어댑터)

  python engine/tools/guide_blocks_to_facts.py <gb.json> <out_facts.json> [--from N] [--to M]

왜: 코퍼스의 남의 제작가이드(.xlsm/.xlsx)를 우리 파이프라인(본문 빌더·렌더러)에 그대로 먹여
「같은 입력으로 우리가 만들면 무엇이 달라지는가」를 재는 것이 벤치마크다. 그러려면 남의 가이드를
우리 정본 포맷으로 옮기는 다리가 필요하다 (코퍼스 작업 2026-08-24 · PPT 팀과 스키마 합의).

내는 것 (PPT 팀 렌더러가 읽는 필드를 기준으로 맞췄다):
  필수  slides[] · order(1부터 연속) · question_code · grid
        grid.rows[].label_clean · kind · cells[] · grid.banner_group_row · banner_label_row · base_row
  선택  page · chart_type · qu · ment · slide_title · sub_title · title_full · base_unit · base_note
        grid.rank_columns · rows[].label · rows[].rank_first_total

블록 하나의 실측 구조 (코퍼스 4벌 공통):
    (패널 제목 행)              ← 행 라벨 열에 문항 제목·문항코드
    배너 그룹 행                ← 행 라벨 열은 비고, 데이터 열에 군 이름(군의 첫 열에만)
    배너 라벨 행                ← 집단 라벨
    `Base for %` 행             ← 행 라벨 열에 `Base for %`, 데이터 열에 `(n)`
    데이터 행…                  ← 행 라벨 열에 보기, 데이터 열에 값

**다단 패널 블록(요약·결론)은 grid 하나로 옮길 수 없다** — 건너뛰고 `skipped` 에 사유와 함께 남긴다.
요약은 `summary_map.json` 이 담당하는 자리이기 때문이다.
"""
from __future__ import annotations
import argparse, json, re
from pathlib import Path
from openpyxl.utils import column_index_from_string, get_column_letter

RE_TOTAL = re.compile(r"^\s*(계|합계|소계)\s*$")
RE_TOP = re.compile(r"【\s*TOP")
RE_BOT = re.compile(r"【\s*BOT")
RE_SOSO = re.compile(r"【\s*SoSo|【\s*보통")
RE_MEAN = re.compile(r"\[\s*\d*\s*점?\s*평균|평균\s*\[|^\s*평균|\[100점 환산")
RE_NPS = re.compile(r"\[\s*NPS")
#: 순위 문항의 두 전체 열 — `1순위` / `1+2+3순위` · `1+N순위` · `1~3순위`
RE_RANK_COL = re.compile(r"^\s*\d+\s*순위\s*$|^\s*1\s*[+~][\d+N~]*\s*순위\s*$")


def row_kind(label: str) -> str:
    s = str(label or "").strip()
    if RE_TOTAL.match(s):
        return "total"
    if RE_TOP.search(s):
        return "top"
    if RE_BOT.search(s):
        return "bot"
    if RE_SOSO.search(s):
        return "soso"
    if RE_NPS.search(s):
        return "nps"
    if RE_MEAN.search(s):
        return "mean"
    return "option"


def cell(v) -> str:
    """값 칸 하나 — 표기 문자열로 옮긴다. 평균은 소수 2자리, 비율은 1자리."""
    if v is None:
        return "-"
    if isinstance(v, (int, float)):
        return f"{float(v):.1f}"
    s = str(v).strip()
    return s or "-"


def mean_cell(v) -> str:
    if v is None:
        return "-"
    if isinstance(v, (int, float)):
        return f"{float(v):.2f}"
    return str(v).strip() or "-"


def find_row_label_col(blocks: list[dict]) -> str | None:
    """행 라벨 열을 **`Base for %` 가 실제로 놓인 열**로 정한다.

    하우스마다 오프셋이 다르다 — 라벨 열 기준으로 +1 인 곳(대전한옥 A/B)도 있고 +2 인 곳(LG A/B/C)도 있으며
    값 열과 같은 곳(SH B/C)도 있다. 고정 오프셋으로 가정하면 한 벌이 통째로 0건이 된다(실측 2026-08-24: LG 79블록 전부 실패).
    """
    tally: dict[str, int] = {}
    for b in blocks:
        for r in b.get("body_rows") or []:
            for col, v in (r.get("cells") or {}).items():
                if str(v or "").strip() == "Base for %":
                    tally[col] = tally.get(col, 0) + 1
    return max(tally, key=tally.get) if tally else None


def block_to_slide(b: dict, lab: str) -> tuple[dict | None, str]:
    """블록 하나 → 슬라이드 팩트. 못 옮기면 (None, 사유). `lab` 은 행 라벨 열 문자."""
    rows = b.get("body_rows") or []
    if not rows:
        return None, "본문 행 없음(간지·표지·텍스트 블록)"
    base_idx = [i for i, r in enumerate(rows)
                if str((r.get("cells") or {}).get(lab) or "").strip() == "Base for %"]
    if not base_idx:
        return None, "`Base for %` 행 없음"
    if len(base_idx) > 1 or (b.get("horizontal_data_panels") or 0) > 1:
        return None, f"다단 패널({len(base_idx)}표 · 가로 {b.get('horizontal_data_panels')}) — summary_map 소관"
    bi = base_idx[0]
    base_cells = rows[bi].get("cells") or {}
    data_cols = [c for c in base_cells if c != lab]
    data_cols.sort(key=column_index_from_string)
    if not data_cols:
        return None, "Base 행에 데이터 열 없음"

    # 배너 헤더 — Base 행 위쪽에서 행 라벨이 비고 데이터 열에 값이 있는 줄 (최대 2줄)
    hdr = []
    for i in range(bi - 1, max(-1, bi - 4), -1):
        c = rows[i].get("cells") or {}
        if str(c.get(lab) or "").strip():
            break
        if any(str(c.get(col) or "").strip() for col in data_cols):
            hdr.append(c)
        if len(hdr) == 2:
            break
    hdr.reverse()
    group_src = hdr[0] if hdr else {}
    label_src = hdr[1] if len(hdr) > 1 else (hdr[0] if hdr else {})

    banner_group_row = ["Base for %"] + [str(group_src.get(c) or "").strip() for c in data_cols]
    banner_label_row = ["Base for %"] + [str(label_src.get(c) or "").strip() for c in data_cols]
    base_row = ["Base for %"] + [str(base_cells.get(c) or "").strip() for c in data_cols]

    # ── 「전체 성격」 선두 열이 몇 개인가 — 배너 집단 라벨이 없는 앞쪽 열이 그것이다.
    #    하우스마다 1개가 아니다: 순위 문항은 `1순위`/`1+N순위` 두 열, 시계열 조사는 `2024년 전체`/`2025년 전체`
    #    두 열이다. 렌더러가 2 로 박아 두면 배너 이름과 값이 한 칸 어긋난다 (PPT 팀 실측 2026-08-24).
    #    추정하게 두지 말고 **개수와 이름을 데이터로 준다.**
    n_total = 0
    for i in range(1, len(banner_label_row)):
        if banner_label_row[i]:
            break
        n_total += 1
    total_columns = banner_group_row[1:1 + n_total]

    # 순위 겹침 — 두 전체 열이 순위 이름이면 우리 정본 포맷으로 옮긴다:
    #   cells[0] = 1+N순위(대표값) · rank_first_total = 1순위 · 배너는 cells[1] 부터
    rank = (n_total == 2 and all(RE_RANK_COL.match(x or "") for x in total_columns))
    # 두 열의 **순서는 하우스마다 다르다** — LG 는 `1순위 · 1+2+3순위`, 대전한옥은 `1+2+3순위 · 1순위` 다.
    # 위치로 정하면 대표값과 1순위가 뒤바뀐다(실측 2026-08-24). 이름으로 정한다.
    i_cum = i_first = None
    if rank:
        for i, nm in enumerate(total_columns):
            (i_cum, i_first) = (i, i_first) if ("+" in nm or "~" in nm) else (i_cum, i)
        if i_cum is None or i_first is None:            # 둘 다 단일 순위면 겹침이 아니다
            rank = False
    grid_rows = []
    for r in rows[bi + 1:]:
        c = r.get("cells") or {}
        label = str(c.get(lab) or "").strip()
        if not label:
            continue
        kind = row_kind(label)
        f = mean_cell if kind == "mean" else cell
        cells = [f(c.get(col)) for col in data_cols]
        entry = {
            "label": label, "label_clean": label.strip(), "kind": kind,
            "cells": cells,
            "modifier": next((str(v).strip() for k, v in c.items()
                              if k not in data_cols and k != lab and "그래프" in str(v)), None),
        }
        if rank:
            entry["rank_first_total"] = cells[i_first]  # 1순위
            entry["cells"] = [cells[i_cum]] + cells[2:]  # cells[0] = 1+N순위(대표값), 그 뒤 배너
        grid_rows.append(entry)
    if not grid_rows:
        return None, "데이터 행 없음"
    if rank:
        def collapse(seq, keep):
            return [seq[0], seq[1 + keep]] + seq[3:]
        banner_group_row = collapse(banner_group_row, i_cum)
        banner_label_row = collapse(banner_label_row, i_cum)
        base_row = collapse(base_row, i_cum)

    title = b.get("slide_title") or ""
    sub = b.get("sub_title") or ""
    return {
        "page": None, "order": None,                       # 호출부에서 채운다
        "question_code": b.get("question_code"),
        "slide_title": title, "sub_title": sub,
        "title_full": (f"{title} | {sub}" if sub else title),
        "chart_type": b.get("chart_type") or "",
        "base_unit": b.get("base_unit") or "",
        "base_note": None,
        "qu": b.get("question_reference") or "",
        "ment": list(b.get("ments") or []),
        "base_n": b.get("base_n"),
        "grid": {
            "banner_group_row": banner_group_row,
            "banner_label_row": banner_label_row,
            "base_row": base_row,
            "rows": grid_rows,
            "rank_overlay": rank,
            "rank_columns": ([total_columns[i_first], total_columns[i_cum]] if rank else None),
            "n_total_columns": 1 if rank else n_total,
            "total_columns": ["1+N순위(대표)"] if rank else list(total_columns),
        },
        "_source_block": b.get("block_index"), "_source_rows": [b.get("start_row"), b.get("end_row")],
    }, ""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("gb", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument("--from", dest="a", type=int, default=0)
    ap.add_argument("--to", dest="b", type=int, default=10 ** 9)
    a = ap.parse_args()
    doc = json.loads(a.gb.read_text(encoding="utf-8"))
    lab = find_row_label_col(doc["blocks"])
    if lab is None:
        print("[adapter] `Base for %` 행을 어느 열에서도 찾지 못했습니다 — 형식이 다릅니다")
        return 2
    print(f"[adapter] 행 라벨 열 = {lab} (라벨 열 {doc.get('label_col')})")
    slides, skipped = [], []
    for blk in doc["blocks"]:
        if not (a.a <= blk["block_index"] <= a.b):
            continue
        s, why = block_to_slide(blk, lab)
        if s is None:
            skipped.append({"block": blk["block_index"], "title": blk.get("slide_title"), "why": why})
            continue
        s["order"] = len(slides) + 1
        s["page"] = len(slides) + 1
        slides.append(s)
    out = {"meta": {"source": doc.get("source_file"), "sheet": doc.get("sheet"),
                    "adapter": "guide_blocks_to_facts.py", "n_slides": len(slides),
                    "n_skipped": len(skipped)},
           "slides": slides, "skipped": skipped}
    a.out.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"→ {a.out}  슬라이드 {len(slides)} · 건너뜀 {len(skipped)}")
    for s in skipped[:8]:
        print(f"    skip 블록 {s['block']} {str(s['title'])[:34]!r} — {s['why']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
