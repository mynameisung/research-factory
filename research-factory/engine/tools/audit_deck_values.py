#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""audit_deck_values.py -- 덱의 **표와 차트를 한 장씩** 정본과 대조한다.

왜 따로 만드는가
----------------
`verify.py --mode deck` 은 표 셀을 통계표와 전수 대조한다. 그런데 **차트 값**은
「차트 ↔ 표 라벨」만 본다(D9). 표는 맞는데 차트만 낡은 상태가 이 업무에서 가장 자주
나오는 결함인데(주입이 표와 차트를 따로 하므로), 그 자리가 비어 있다.

이 도구는 **덱 실물에서 값을 다시 읽어** 정본 그리드와 1:1로 맞춰 본다.
주입 코드를 부르지 않고 기대값을 정본에서 직접 세우므로, 주입이 일부만 됐거나
엉뚱한 차트에 들어갔을 때 걸린다.

무엇을 대조하는가 (본문 장표 한 장마다)
---------------------------------------
    표    값 셀 · 행 라벨 · (Base) 행 · 배너 라벨 행
    차트  카테고리 · 계열 값

읽는 규칙은 주입 규칙과 같다 — 「전체」 열은 평균 행만 채우고, 차트는 형태마다
값의 출처가 다르다(가로막대=전체 열 · 겹침막대=누적/1순위 · ±세로막대=배너 열).

사용:
    python engine/tools/audit_deck_values.py <덱.pptx> <slide_facts.json>         [--md <리포트>] [--json <리포트>] [--tol 0.05]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent))

import deckops as D                                       # noqa: E402
import refill_deck as RF                                  # noqa: E402

q, A, P, C = D.q, D.A, D.P, D.C


# --------------------------------------------------------------------------- #
# 기대값 -- 정본에서 직접 세운다 (주입 코드를 부르지 않는다)
# --------------------------------------------------------------------------- #

def expect_table(fact: dict[str, Any]) -> dict[str, Any]:
    """이 장 표의 기대 모습. {행: [(열, 값)], 라벨, base}"""
    grid = fact["grid"]
    rows = RF.body_rows(grid)
    out_rows = []
    for r in rows:
        disp = r.get("_display") or []
        is_mean = r.get("kind") == "mean"
        cells: list[str | None] = []
        for i, v in enumerate(disp):
            if i == 0 and not is_mean:
                cells.append(None)          # 「전체」 열은 차트가 덮는다 -- 비어 있어야 맞다
            else:
                cells.append("-" if v is None else v)
        out_rows.append({"label": RF.row_label(r), "kind": r.get("kind"),
                         "cells": cells})
    return {"rows": out_rows,
            "base_row": grid.get("base_row") or [],
            "banner_label_row": grid.get("banner_label_row") or []}


def expect_chart(fact: dict[str, Any]) -> dict[str, Any] | None:
    """이 장 차트의 기대 모습. 형태마다 값의 출처가 다르다."""
    grid = fact["grid"]
    rows = RF.body_rows(grid)
    kind = RF.chart_kind(fact.get("chart_type") or "")

    if kind == "scale":
        # 계열 = 척도 단계, 범주 = 배너 열. 부호는 덱이 정하므로 절댓값으로 본다.
        opts = {RF._norm(r.get("label_clean") or ""): r
                for r in rows if r.get("kind") == "option"}
        return {"kind": kind, "by_series": {
            k: [RF._num(x) for x in (r.get("cells") or [])] for k, r in opts.items()}}

    axis = [r if r.get("kind") == "option" else None for r in rows]
    cats = [(r.get("label_clean") if r else "") for r in axis]
    if kind == "rank":
        # ★ 순위 겹침막대의 두 계열은 **배너 열이 아니다.** 「전체」 열 하나에
        #   두 지표를 겹쳐 그린다: 누적(1+N순위)=cells[0] · 1순위=rank_first_total.
        #   `rank_columns` 의 인덱스를 cells 인덱스로 쓰면 「전체 열」과 「첫 배너 열」을
        #   두 계열로 읽어, 멀쩡한 덱을 전부 불일치로 잡는다(실측: 69장 중 10장).
        cum, first = RF.rank_names(grid)
        series = [[(RF._num(r["cells"][0]) if r else None) for r in axis],
                  [(RF._num(r.get("rank_first_total")) if r else None) for r in axis]]
        names = [cum, first]
    else:
        series = [[(RF._num(r["cells"][0]) if r else None) for r in axis]]
        names = ["전체"]
    return {"kind": kind, "cats": cats, "series": series, "names": names}


# --------------------------------------------------------------------------- #
# 실물 -- 덱에서 다시 읽는다
# --------------------------------------------------------------------------- #

def actual_table(deck: D.Deck, part: str) -> dict[str, Any] | None:
    tree = deck.xml(part)
    tbl = next(iter(tree.iter(q(A, "tbl"))), None)
    if tbl is None:
        return None
    trs = D.data_rows(tbl, D.HEADER_ROWS)
    rows = []
    for tr in trs:
        tcs = D.row_cells(tr)
        rows.append({"label": D.cell_text(tcs[0]).strip(),
                     "cells": [D.cell_text(tc).strip() or None for tc in tcs[1:]]})
    head = D.table_rows(tbl)[:D.HEADER_ROWS]
    base = [D.cell_text(tc).strip() for tc in D.row_cells(head[2])] if len(head) > 2 else []
    labels = [D.cell_text(tc).strip() for tc in D.row_cells(head[1])] if len(head) > 1 else []
    return {"rows": rows, "base_row": base, "banner_label_row": labels}


def actual_chart(deck: D.Deck, part: str) -> dict[str, Any] | None:
    cps = deck.chart_parts(part)
    if not cps:
        return None
    t = deck.xml(cps[0])
    return {"part": cps[0], "cats": D.chart_categories(t),
            "names": D.series_names(t), "series": D.series_values(t)}


# --------------------------------------------------------------------------- #
# 대조
# --------------------------------------------------------------------------- #

def _num_eq(a: Any, b: Any, tol: float) -> bool:
    if a is None and b is None:
        return True
    fa, fb = RF._num(a), RF._num(b)
    if fa is None or fb is None:
        # 둘 다 숫자가 아니면 문자열로 본다. `-` 와 빈칸은 같은 뜻이다.
        sa = "" if a in (None, "-", "·") else str(a).strip()
        sb = "" if b in (None, "-", "·") else str(b).strip()
        return sa == sb
    return abs(fa - fb) <= tol


def audit_scale_table(deck: D.Deck, part: str, fact: dict[str, Any],
                      rep: dict[str, Any], *, tol: float) -> dict[str, Any]:
    """±세로막대 장표의 표 -- 지표 행(TOP2·평균)만 라벨의 뜻으로 찾아 대조한다."""
    tree = deck.xml(part)
    tbl = next(iter(tree.iter(q(A, "tbl"))), None)
    if tbl is None:
        rep["issues"].append({"where": "표", "what": "표가 없다"})
        return {}
    by_kind = {r.get("kind"): r for r in (fact["grid"].get("rows") or [])}
    n_cell = n_bad = 0
    seen: list[str] = []
    for tr in D.table_rows(tbl):
        tcs = D.row_cells(tr)
        if not tcs:
            continue
        label = D.cell_text(tcs[0]).strip()
        if not label:
            continue
        kind = next((k for k, marks in RF.SCALE_ROW_KINDS
                     if any(m in label for m in marks)), None)
        row = by_kind.get(kind) if kind else None
        if row is None:
            continue
        seen.append(f"{label}<-{kind}")
        disp = row.get("_display") or []
        for i, tc in enumerate(tcs[1:]):
            if i >= len(disp) or disp[i] is None:
                continue
            n_cell += 1
            got = D.cell_text(tc).strip()
            if not _num_eq(got, disp[i], tol):
                n_bad += 1
                if n_bad <= 6:
                    rep["issues"].append({"where": f"표 {label} / 열{i + 1}",
                                          "what": "값 다름",
                                          "deck": got, "facts": disp[i]})
    if not seen:
        rep["issues"].append({"where": "표", "what": "지표 행(TOP2·평균)을 못 찾았다"})
    return {"cells": n_cell, "bad": n_bad, "rows": len(seen)}


def audit_slide(deck: D.Deck, part: str, fact: dict[str, Any], slide_no: int,
                *, tol: float) -> dict[str, Any]:
    rep: dict[str, Any] = {"slide": slide_no, "code": fact.get("question_code"),
                           "sub": fact.get("sub_title"),
                           "chart_type": RF.chart_kind(fact.get("chart_type") or ""),
                           "table": {}, "chart": {}, "issues": []}

    # ── 표 ────────────────────────────────────────────────────────────
    if rep["chart_type"] == "scale":
        # ★ ±세로막대 장표의 표는 **구조가 뒤집혀 있다** -- 지표 행(TOP2·평균)이 위,
        #   배너 헤더가 아래다. 헤더 3행을 떼고 읽는 보통 방식으로 보면 배너 라벨을
        #   데이터 행으로 읽어 전 칸이 불일치로 잡힌다. 라벨의 뜻으로 지표 행만 본다.
        rep["table"] = audit_scale_table(deck, part, fact, rep, tol=tol)
        want = None
        got = None
    else:
        want = expect_table(fact)
        got = actual_table(deck, part)
    if want is None:
        pass
    elif got is None:
        rep["issues"].append({"where": "표", "what": "표가 없다"})
    else:
        n_cell = n_bad = 0
        if len(got["rows"]) != len(want["rows"]):
            rep["issues"].append({"where": "표", "what": "행 수 불일치",
                                  "deck": len(got["rows"]), "facts": len(want["rows"])})
        for i, (g, w) in enumerate(zip(got["rows"], want["rows"])):
            if RF._norm(g["label"]) != RF._norm(w["label"]):
                rep["issues"].append({"where": f"표 행{i + 1} 라벨", "what": "다름",
                                      "deck": g["label"], "facts": w["label"]})
            for j, (gc, wc) in enumerate(zip(g["cells"], w["cells"])):
                n_cell += 1
                if not _num_eq(gc, wc, tol):
                    n_bad += 1
                    if n_bad <= 6:
                        col = (want["banner_label_row"][j + 1]
                               if j + 1 < len(want["banner_label_row"]) else f"열{j + 1}")
                        rep["issues"].append({
                            "where": f"표 {w['label']} / {col or '전체'}",
                            "what": "값 다름", "deck": gc, "facts": wc})
        if got["base_row"] and want["base_row"]:
            for j, (gb, wb) in enumerate(zip(got["base_row"], want["base_row"])):
                # 제작본은 저베이스에 별표를 단다 -- `(13*)` 와 `(13)` 은 같은 뜻이다
                if RF._norm(gb).replace("*", "") != RF._norm(wb).replace("*", "") \
                        and j > 0:
                    rep["issues"].append({"where": f"표 (Base) 열{j}", "what": "다름",
                                          "deck": gb, "facts": wb})
        rep["table"] = {"cells": n_cell, "bad": n_bad, "rows": len(got["rows"])}

    # ── 차트 ──────────────────────────────────────────────────────────
    w = expect_chart(fact)
    g = actual_chart(deck, part)
    if g is None:
        rep["issues"].append({"where": "차트", "what": "차트가 없다"})
        return rep
    if w is None:
        return rep

    if w["kind"] == "scale":
        n_val = n_bad = 0
        for nm, vals in zip(g["names"], g["series"]):
            exp = w["by_series"].get(RF._norm(nm or ""))
            if exp is None:
                # 응답 0인 척도 단계는 통계표에 행이 없다 -- 덱은 빈칸이어야 맞다
                if any(v is not None for v in vals):
                    rep["issues"].append({"where": f"차트 계열 {nm}",
                                          "what": "정본에 없는 단계인데 값이 있다",
                                          "deck": [round(v, 1) for v in vals if v is not None][:4]})
                continue
            for j, gv in enumerate(vals):
                n_val += 1
                ev = exp[j] if j < len(exp) else None
                # 부호는 덱의 수식이 만든다 -- 크기만 본다
                if not _num_eq(None if gv is None else abs(gv),
                               None if ev is None else abs(ev), tol):
                    n_bad += 1
                    if n_bad <= 6:
                        rep["issues"].append({
                            "where": f"차트 {nm} / 열{j + 1}", "what": "값 다름",
                            "deck": None if gv is None else round(abs(gv), 2),
                            "facts": None if ev is None else round(abs(ev), 2)})
        rep["chart"] = {"values": n_val, "bad": n_bad, "series": len(g["series"])}
        return rep

    # 가로막대 · 겹침막대
    if [RF._norm(x) for x in g["cats"]] != [RF._norm(x) for x in w["cats"]]:
        rep["issues"].append({"where": "차트 카테고리", "what": "다름",
                              "deck": g["cats"][:4], "facts": w["cats"][:4]})
    n_val = n_bad = 0
    for k, exp in enumerate(w["series"]):
        if k >= len(g["series"]):
            rep["issues"].append({"where": f"차트 계열{k + 1}", "what": "계열이 없다"})
            continue
        for j, ev in enumerate(exp):
            n_val += 1
            gv = g["series"][k][j] if j < len(g["series"][k]) else None
            if not _num_eq(gv, ev, tol):
                n_bad += 1
                if n_bad <= 6:
                    cat = w["cats"][j] if j < len(w["cats"]) else f"행{j + 1}"
                    rep["issues"].append({
                        "where": f"차트 {w['names'][k]} / {cat[:20]}",
                        "what": "값 다름",
                        "deck": None if gv is None else round(gv, 2),
                        "facts": None if ev is None else round(ev, 2)})
    # 덱에 남는 계열(범례 「없음」 자리)은 비어 있어야 한다
    for k in range(len(w["series"]), len(g["series"])):
        if any(v is not None for v in g["series"][k]):
            rep["issues"].append({"where": f"차트 여분 계열 {g['names'][k]}",
                                  "what": "비어 있어야 하는데 값이 있다"})
    rep["chart"] = {"values": n_val, "bad": n_bad, "series": len(g["series"])}
    return rep


def audit(deck_path: Path, facts_path: Path, *, tol: float = 0.05) -> dict[str, Any]:
    deck = D.Deck(deck_path)
    facts = RF.load_facts(facts_path)
    pairs = RF.pair_up(deck, facts)
    slides = [audit_slide(deck, p.part, p.fact, p.slide_no, tol=tol) for p in pairs]
    return {"deck": str(deck_path), "facts": str(facts_path),
            "slides": slides,
            "cells": sum((s["table"] or {}).get("cells", 0) for s in slides),
            "cells_bad": sum((s["table"] or {}).get("bad", 0) for s in slides),
            "chart_values": sum((s["chart"] or {}).get("values", 0) for s in slides),
            "chart_bad": sum((s["chart"] or {}).get("bad", 0) for s in slides),
            "slides_bad": sum(1 for s in slides if s["issues"])}


def render(rep: dict[str, Any]) -> str:
    L = ["# 덱 ↔ 정본 한 장씩 대조", "",
         f"- 덱: `{rep['deck']}`",
         f"- 정본: `{rep['facts']}`",
         f"- 본문 {len(rep['slides'])}장 · 표 {rep['cells']}칸(불일치 {rep['cells_bad']})"
         f" · 차트 {rep['chart_values']}값(불일치 {rep['chart_bad']})",
         f"- **지적이 있는 장 {rep['slides_bad']}장**", "",
         "## 장별", "",
         "| 슬라이드 | 문항 | 형태 | 표 칸(불일치) | 차트 값(불일치) | 지적 |",
         "|---|---|---|---|---|---|"]
    for s in rep["slides"]:
        t, c = s.get("table") or {}, s.get("chart") or {}
        mark = "**" + str(len(s["issues"])) + "건**" if s["issues"] else "-"
        L.append(f"| s{s['slide']} | {s['code']} | {s['chart_type']} |"
                 f" {t.get('cells', 0)}({t.get('bad', 0)}) |"
                 f" {c.get('values', 0)}({c.get('bad', 0)}) | {mark} |")
    bad = [s for s in rep["slides"] if s["issues"]]
    if bad:
        L += ["", "## 지적 상세", ""]
        for s in bad:
            L.append(f"### s{s['slide']} {s['code']} — {s['sub']}")
            for i in s["issues"]:
                d, f = i.get("deck"), i.get("facts")
                L.append(f"- **{i['where']}** {i['what']}"
                         + (f" — 덱 `{d}` / 정본 `{f}`" if d is not None or f is not None else ""))
            L.append("")
    else:
        L += ["", "## 지적 상세", "", "없다 — 표와 차트가 정본과 전부 일치한다.", ""]
    return "\n".join(L)


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="덱의 표·차트를 한 장씩 정본과 대조한다")
    ap.add_argument("deck")
    ap.add_argument("facts")
    ap.add_argument("--tol", type=float, default=0.05,
                    help="숫자 허용 오차 (기본 0.05 — 표시 반올림 폭)")
    ap.add_argument("--md")
    ap.add_argument("--json")
    args = ap.parse_args(argv)

    rep = audit(Path(args.deck), Path(args.facts), tol=args.tol)
    text = render(rep)
    print(text)
    if args.md:
        Path(args.md).write_text(text, encoding="utf-8")
        print(f"\n[리포트] {args.md}")
    if args.json:
        Path(args.json).write_text(json.dumps(rep, ensure_ascii=False, indent=1),
                                   encoding="utf-8")
    return 1 if rep["slides_bad"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
