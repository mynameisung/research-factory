#!/usr/bin/env python3
"""골든셋 예측 — 가이드 블록의 **데이터만** 보고 슬라이드 사양을 예측한다.

역검증 하네스의 ① 단계다. 예시 제작가이드의 본문 블록에서 통계표 부분만 읽어
파이프라인이 무엇을 만들었을지 예측하고, 검수관이 그것을 제작본(정답)과 대조한다.

★ 반드시 지킬 것
----------------
가이드에 이미 적혀 있는 `Chart type` / `Sub title` 을 **읽지 않는다.** 그건 연구원의
답이고 우리가 맞혀야 할 대상이다. 읽는 순간 채점이 무의미해진다.
(`base_unit` 과 `question_reference` 는 예측 대상이 아니라 전사 대상이라 그대로 쓴다 —
 정답인 제작본 쪽 표기와 일치하는지가 채점 항목이다.)

정답은 가이드가 아니라 **제작본**이다. 연구원이 적은 것과 제작팀이 만든 것이 다를 수
있고, 그 불일치 자체가 규칙 후보다.

사용법
------
    python predict_golden.py <guide_blocks.json> <predict.json>
"""

from __future__ import annotations

import argparse
import io
import json
import re
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))
import rules as R  # noqa: E402

COL_RE = re.compile(r"^[A-Z]+$")


def col_index(letter: str) -> int:
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - 64)
    return n


def to_float(v: Any) -> float | None:
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        if s in ("", "-", "–"):
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def reconstruct_table(block: dict) -> dict | None:
    """가이드 블록의 붙여넣은 통계표를 `tables.json` 스키마 비슷하게 되살린다.

    룰엔진이 통계표를 먹도록 만들어져 있으므로, 가이드 쪽 입력도 같은 모양으로 맞춘다.
    """
    body = block.get("body_rows") or []
    sentinels = {s["row"]: s for s in (block.get("sentinel_rows") or [])}
    base_rows = block.get("base_for_pct_rows") or []
    if not base_rows:
        return None
    base_row_no = base_rows[0]

    # ── 배너 헤더: Base for % 바로 위 1~2줄.
    #    패널 제목 행이 바로 위에 붙어 있어서 그대로 잡으면 헤더 대신 제목을 집는다.
    #    순위 표는 헤더 행에 `1순위`/`1+2+3순위` 가 들어 있으므로, 이걸 놓치면
    #    순위 문항을 전부 기본형으로 오판한다.
    # 이 문항이 차지하는 열 범위 (다단 블록 대응). 아래 배너 헤더 필터에 쓴다.
    _titles = sorted((t for t in (block.get("panel_titles") or []) if t.get("col")),
                     key=lambda t: col_index(t["col"]))
    panel_span = None
    if len(_titles) > 1:
        _k = block.get("join_key")
        _i = next((i for i, t in enumerate(_titles) if t.get("code") == _k), None)
        if _i is not None:
            panel_span = (col_index(_titles[_i]["col"]),
                          col_index(_titles[_i + 1]["col"]) if _i + 1 < len(_titles) else 10**6)

    title_rows = {t["row"] for t in (block.get("panel_titles") or [])}
    header_rows = [r for r in body
                   if base_row_no - 3 <= r["row"] < base_row_no
                   and r["cells"] and r["row"] not in title_rows][-2:]
    def in_panel(k: str) -> bool:
        return COL_RE.match(k) is not None and (not panel_span or
                                                panel_span[0] <= col_index(k) < panel_span[1])

    group_cells: dict[str, str] = {}
    label_cells: dict[str, str] = {}
    if header_rows:
        group_cells = {k: str(v) for k, v in header_rows[0]["cells"].items() if in_panel(k)}
    if len(header_rows) > 1:
        label_cells = {k: str(v) for k, v in header_rows[-1]["cells"].items() if in_panel(k)}
    # 2단 배너 헤더(그룹행 + 라벨행)가 실재하는가. 장분리(MAP-050)의 전제 조건이다 —
    # 1단짜리 표(SKEE 지표비교표 등)에서는 "배너군"이라는 개념 자체가 없다.
    two_tier = len(header_rows) > 1

    base_cells = (sentinels.get(base_row_no) or {}).get("values") or {}
    cols = sorted((c for c in base_cells if COL_RE.match(c)), key=col_index)
    if not cols:
        return None

    # ── 다단 블록: 패널이 가로로 나란히 놓인다 (요약 슬라이드가 대부분).
    #    합쳐서 읽으면 옆 패널의 `1순위` 헤더를 이 문항 것으로 오독한다 —
    #    실제로 A1(기본형)이 옆 패널 A12(순위) 때문에 순위로 잘못 판정됐다.
    #    이 블록의 join_key 에 해당하는 패널의 열 범위만 남긴다.
    titles = sorted((t for t in (block.get("panel_titles") or []) if t.get("col")),
                    key=lambda t: col_index(t["col"]))
    if len(titles) > 1:
        key = block.get("join_key")
        mine = next((i for i, t in enumerate(titles) if t.get("code") == key), None)
        if mine is not None:
            lo = col_index(titles[mine]["col"])
            hi = col_index(titles[mine + 1]["col"]) if mine + 1 < len(titles) else 10**6
            cols = [c for c in cols if lo <= col_index(c) < hi]
            if not cols:
                return None

    banners = []
    cur_group = ""
    for pos, c in enumerate(cols):
        if c in group_cells:
            cur_group = group_cells[c]
        label = label_cells.get(c) or group_cells.get(c) or cur_group
        m = re.search(r"\(?(\d[\d,]*)\)?", str(base_cells[c]))
        banners.append({
            "pos": pos, "col_letter": c,
            "group": cur_group or "전체",
            "label": label,
            "n": int(m.group(1).replace(",", "")) if m else 0,
            # 골든셋 예측에서는 배너 절단을 적용하지 않는다 — LG 배너군이 다르다
            "in_report": True,
        })

    # ── 데이터 행: Base for % 아래, 셀에 값이 있는 행
    rows = []
    for r in body:
        if r["row"] <= base_row_no or not r["cells"]:
            continue
        cells = r["cells"]
        # 센티널 행은 추출기가 이미 라벨을 확정해 뒀다 — 그걸 우선한다.
        sn = sentinels.get(r["row"])
        label = sn.get("label") if sn else None
        if label is None:
            # 라벨은 데이터 시작 열 **바로 왼쪽**이다. 더 왼쪽 열에는 행 수식어
            # (`그래프 상단 Text`)가 들어가므로, 첫 열을 집으면 그걸 라벨로 오독한다.
            left = [c for c in cells if COL_RE.match(c) and col_index(c) < col_index(cols[0])]
            if left:
                label = str(cells[max(left, key=col_index)])
        if label is None:
            continue

        kind = "option"
        s = label.strip()
        if s == "계":
            kind = "total"
        elif re.match(r"^【BOT", s):
            kind = "bot"
        elif re.match(r"^【SoSo", s):
            kind = "soso"
        elif re.match(r"^【TOP", s):
            kind = "top"
        elif re.match(r"^\[.*(평균|NPS)", s):
            kind = "mean"

        values = []
        for c in cols:
            v = to_float(cells.get(c))
            values.append({"value": v, "display": None if v is None else f"{v}"})
        rows.append({"label_raw": label, "label": s, "kind": kind, "values": values,
                     "option_no": int(m.group(1)) if (m := re.match(r"^\s*(\d+)\)", label)) else None})

    opts = [r for r in rows if r["kind"] == "option"]
    tops = [r for r in rows if r["kind"] == "top"]
    scale_points = None
    if tops:
        m = re.search(r"TOP(\d)", tops[0]["label"])
        scale_points = 7 if (m and m.group(1) == "3") else 5
    means = [r for r in rows if r["kind"] == "mean"]
    if means and scale_points is None:
        m = re.search(r"\[(\d+)점", means[0]["label"])
        scale_points = int(m.group(1)) if m else None

    # ── 순위 판별: 배너 **라벨 완전일치** 로만 본다.
    #    헤더 문자열 부분일치로 보면 교차표 머리글(`C4 1+2+3순위` × `C2 1순위`)이
    #    순위로 잡힌다 — 라운드1 요약 s10 이 실제로 그렇게 오탐됐다 (RC-2).
    rank_first, rank_wide = R.rank_banners({"banners": banners})
    is_ranking = rank_first is not None and rank_wide is not None
    base_unit = block.get("base_unit") or ""

    return {
        "index": block.get("block_index"),
        "title_full": (block.get("panel_titles") or [{}])[0].get("text"),
        "question_code": block.get("question_code"),
        "base_n": block.get("base_n"),
        "banners": banners,
        "rows": rows,
        "suffix": "",
        "hints": {
            "has_scale_rows": bool(tops or [r for r in rows if r["kind"] in ("bot", "soso")]),
            "scale_points": scale_points,
            "is_ranking": is_ranking,
            "rank_is_first_only": False,
            "is_multi": "중복" in base_unit,
            "is_summary": False,
            # 매트릭스 항목 = 한 문항이 항목별로 쪼개진 표. LG 는 `A10-1`, `BA1-3` 처럼
            # 문항코드에 하이픈 인덱스가 붙는다.
            "is_matrix_item": bool(re.match(r"^[A-Z]+\d+-\d+$", block.get("question_code") or "")),
            "two_tier_banner_header": two_tier,
            "has_total_row": any(r["kind"] == "total" for r in rows),
            "has_mean_row": bool(means),
            "n_option_rows": len(opts),
            "n_banners": len(banners),
            "n_banners_in_report": len(banners),
        },
    }


def body_start_index(blocks: list[dict], rules: dict) -> int | None:
    """본문 구간이 시작하는 블록 인덱스 (MAP-020).

    `조사 결과` Part 간지 뒤부터가 본문이다. 그 앞(조사 개요 · 요약 및 결론)에도
    통계표를 품은 블록이 있어서 "표가 있으면 본문"으로 처리하면 비본문이 섞여 나온다 —
    라운드1 에서 LG 4건(응답자 특성 · 요약 3장)이 그렇게 새어 나왔다.

    마커를 못 찾으면 `None` 을 돌려 게이트를 걸지 않는다. 구조를 모르는 가이드에서
    조용히 절반을 버리는 것보다 현행 동작을 유지하는 쪽이 안전하다.
    """
    gate = (rules or {}).get("body_slide_gate") or {}
    det = gate.get("detect_part_marker") or {}
    pat = re.compile(det.get("pattern") or r"^\s*Part\s")
    prefer = [str(x) for x in (det.get("prefer_title_contains") or [])]

    parts = [i for i, b in enumerate(blocks) if pat.match(str(b.get("slide_title") or ""))]
    if not parts:
        return None
    named = [i for i in parts
             if any(w.strip() in str(blocks[i].get("slide_title") or "") for w in prefer)]
    return (named[-1] if named else parts[-1]) + 1


def main() -> int:
    ap = argparse.ArgumentParser(description="가이드 블록 → 예측 사양")
    ap.add_argument("guide_blocks", type=Path)
    ap.add_argument("out", type=Path)
    args = ap.parse_args()

    with io.open(args.guide_blocks, encoding="utf-8") as fh:
        doc = json.load(fh)
    blocks = doc["blocks"] if isinstance(doc, dict) else doc

    chart_rules = R.load_chart_rules()
    if not chart_rules:
        print("[중단] chart_rules.yaml 을 읽지 못했습니다.", file=sys.stderr)
        return 1

    start = body_start_index(blocks, chart_rules)
    if start is None:
        print("[주의] `조사 결과` Part 간지를 못 찾았다 — 본문 게이트(MAP-020)를 걸지 않는다.",
              file=sys.stderr)

    preds, skipped = [], []
    for i, b in enumerate(blocks):
        key = b.get("join_key")
        if not key:
            skipped.append({"block_index": b.get("block_index"), "reason": "join_key 없음"})
            continue
        if start is not None and i < start:
            skipped.append({"block_index": b.get("block_index"), "join_key": key,
                            "reason": "본문 구간 이전 (조사개요·요약·간지) — MAP-020"})
            continue
        t = reconstruct_table(b)
        if t is None:
            skipped.append({"block_index": b.get("block_index"), "join_key": key,
                            "reason": "표를 되살릴 수 없음 (Base for % 행 없음)"})
            continue

        # 차트타입은 문항 단위로 한 번만 정한다. 장을 나눠도 같은 차트가 두 장에 들어간다.
        d = R.decide_chart_type(t, None, chart_rules)
        # 한 문항이 슬라이드 두 장으로 전개되는가 (MAP-050). 배너가 쪼개질 뿐 문항은 하나다.
        views = R.split_banner_chapters(t, chart_rules)

        for view in views:
            # 가이드 그리드 행 순서(통계표 유지) · 차트 옵션축 · 제작본 표 행 순서는 **다른 산출물**이다.
            grid_rows, _ = R.sort_option_rows(view)
            axis, axis_rows, sort_why = R.chart_option_axis(view, d.option_axis)
            tbl_rows, _, tbl_why = R.table_row_axis(view, d.option_axis)
            tb = R.representative_banner(view)
            pos = tb["pos"] if tb else 0

            preds.append({
                "join_key": key,
                "block_index": b.get("block_index"),
                "question_code": b.get("question_code"),
                "slide_title": b.get("slide_title"),
                "chapter": view.get("chapter", 1),
                "chapter_count": view.get("chapter_count", 1),
                "chapter_banner_groups": view.get("chapter_banner_groups"),
                "chapter_basis": view.get("chapter_reason"),
                "predicted_chart_type": d.chart_type,
                "chart_rule": {"id": d.rule_id, "confidence": d.confidence, "reason": d.reason},
                "base_unit": b.get("base_unit"),
                "base_n": b.get("base_n"),
                "qu": b.get("question_reference"),
                "option_order": axis,
                "option_values": [
                    (R.cell_value((r.get("values") or [None] * (pos + 1))[pos]) if r else 0.0)
                    for r in axis_rows
                ],
                "table_row_order": tbl_rows,
                "table_row_basis": tbl_why,
                "guide_grid_row_order": [r["label_raw"] for r in grid_rows],
                "sort_basis": sort_why,
                "representative_banner": (tb or {}).get("label"),
                "graph_top_text_rows": [x["label"] for x in
                                        R.graph_top_text(view, d.top_text_kinds, d.top_text_medium)],
                "graph_top_text": R.graph_top_text(view, d.top_text_kinds, d.top_text_medium),
                # ★ `[]` 로 비우면 "멘트 없음"과 "멘트 생성기 미구현"이 구별되지 않는다.
                #    정답도 비어 있는 골든셋(LG)에서는 빈칸-빈칸이 만점으로 위장된다 (RC-MENT-002).
                "ment": None,
                "ment_status": "미구현 — 이 예측기는 멘트를 생성하지 않는다 (채점 제외)",
                "signals": view["hints"],
            })

    args.out.parent.mkdir(parents=True, exist_ok=True)
    with io.open(args.out, "w", encoding="utf-8") as fh:
        json.dump({
            "meta": {
                "generated_by": "predict_golden.py",
                "source": str(args.guide_blocks),
                "chart_rules_version": chart_rules.get("schema_version"),
                "note": "가이드의 Chart type / Sub title 은 읽지 않았다. 데이터에서만 예측했다.",
            },
            "predictions": preds,
            "skipped": skipped,
        }, fh, ensure_ascii=False, indent=1)

    from collections import Counter
    ct = Counter(p["predicted_chart_type"] for p in preds)
    print(f"블록 {len(blocks)}개 → 예측 {len(preds)}개 · 건너뜀 {len(skipped)}개")
    for k, v in ct.most_common():
        print(f"  {v:>3}  {k}")
    print(f"\n→ {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
