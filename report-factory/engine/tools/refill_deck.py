#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""refill_deck.py -- **이미 만들어진 덱**의 표·차트를 정본으로 다시 채운다.

들어오는 자리
-------------
    통계표가 갱신됐다 / DP 가 값을 고쳐 보냈다 / 제작팀 회신본에 값이 안 들어왔다
        -> tables.json 재파싱 -> build_facts.py -> slide_facts.json
        -> **refill_deck.py** -> 덱의 표·차트가 정본과 다시 일치한다

fill_deck.py 와 무엇이 다른가
------------------------------
`fill_deck --fill` 은 **빈 제작본을 처음 채우는** 도구다. 문항코드로 표를 찾으므로
장분리(같은 코드가 여러 장)에서 엉뚱한 표를 붙이고, Base 라벨을 제 판단으로 바꾼다.
이 모듈은 **장표 순서가 정본과 1:1 이라는 사실**을 조인 키로 쓴다. 순서는 목차가
정하는 것이라 사람이 확인할 수 있고, 장분리에도 흔들리지 않는다.

원칙 -- 구조는 덱이 진리, 값은 정본이 진리
-------------------------------------------
행 수·열 수·계열 수는 **덱에 이미 내려진 결정**이다(꼬리 접기·배너 절단·범례 구성).
값만 갈아 끼우는 것이 이 도구의 기본 동작이고, 구조가 어긋나면 **멈춰서 알린다**.
`--fit-rows` 를 주면 행 수까지 정본에 맞춘다(빈 껍데기를 채울 때 쓴다).

제작본을 존중한다
-----------------
제작팀이 손본 서식은 값 주입으로 지워지지 않는다:
  · (Base) 행        -- 건드리지 않는다. 저베이스 별표 `(13*)` 가 거기 있다
  · 「전체」 열       -- 평균 행만 채운다(나머지 행은 차트가 그 자리를 덮는다)
  · 셀 서식·테두리    -- put_cell 이 런 서식을 물려 쓴다
  · 차트 graphicFrame -- 크기·위치를 손대지 않는다 (R-106)

의존: deckops · deckbuild (engine/tools)
"""

from __future__ import annotations

import argparse
import copy
import json
import posixpath
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lxml import etree                                    # noqa: E402
import deckops as D                                       # noqa: E402
import deckbuild as B                                     # noqa: E402

A, P, C, R = D.A, D.P, D.C, D.R
q = D.q

#: 본문 표·차트에 싣지 않는 행 종류. 「계」는 100% 라 정보가 없다.
BODY_SKIP_KINDS = {"total"}

#: 순위 누적 열임을 알리는 표식 -- `1+2+3순위` · `1~3순위 (중복응답)`
CUMULATIVE_RANK = ("+", "~", "중복")


# =========================================================================== #
# 1. 장표 식별 -- 이름이 아니라 구조로 (R-090 · R-092)
# =========================================================================== #

def is_body_slide(deck: D.Deck, part: str) -> bool:
    """본문 장표인가. `Qu:` 로 시작하는 텍스트가 있으면 본문이다.

    레이아웃 이름으로 가르지 않는다 -- 덱마다 다르고(LG `3_master-1`,
    에스원 제작본 `1_master-2`) 제작팀이 바꾸면 조용히 어긋난다.
    문항이 실린 장만 본문이라는 정의는 어느 하우스에서나 같다.
    """
    tree = deck.xml(part)
    for sp in D.iter_shapes(tree):
        if D.shape_text(sp).lstrip().startswith("Qu:"):
            return True
    return False


#: 부록 간지를 알아보는 표식. 이 간지 **뒤로는 본문이 아니다.**
APPENDIX_RE = re.compile(r"^\s*(Appendix|APPENDIX|부\s*록)\s*$")


def is_appendix_divider(deck: D.Deck, part: str) -> bool:
    """부록 간지인가. 짧은 텍스트 몇 개뿐이고 그 중 하나가 「Appendix」다."""
    tree = deck.xml(part)
    texts = [D.shape_text(sp).strip() for sp in D.iter_shapes(tree)]
    return any(APPENDIX_RE.match(t) for t in texts if t)


def body_slides(deck: D.Deck) -> list[str]:
    """표시 순서대로 본문 장표 파트 이름.

    ★ 부록 장표에도 `Qu:` 가 붙는다 -- 본문에 싣지 않은 배너를 부록으로 빼는 관행이
      있어서다(에스원 실측 2장). 부록 간지를 만나면 거기서 끊는다. 안 끊으면 본문
      장수가 정본보다 많아지고, 순서 매핑이 통째로 한 칸씩 밀린다.
    """
    out: list[str] = []
    for p in deck.slide_order():
        if is_appendix_divider(deck, p):
            break
        if is_body_slide(deck, p):
            out.append(p)
    return out


@dataclass
class Slots:
    """장표의 텍스트 슬롯 지도. 이름이 아니라 위치로 찾은 결과다."""
    part: str
    title: etree._Element | None = None        # 「1. 건물 기본 현황」
    sub: etree._Element | None = None          # 「1) 보유규모」
    ment: etree._Element | None = None         # 멘트 3줄
    heading: etree._Element | None = None      # 「DQ1. 소유 또는 관리 건물 수」
    qu: etree._Element | None = None
    base: etree._Element | None = None
    table: etree._Element | None = None        # a:tbl
    chart_frames: list[etree._Element] = field(default_factory=list)


def find_slots(deck: D.Deck, part: str) -> Slots:
    """장표의 슬롯을 위치·텍스트로 찾는다.

    제작팀은 도형을 자기 편의대로 다시 만든다 -- 소제목이 `AUTO_SHAPE '소제목'` 이던
    것이 회신본에서는 `GROUP > Freeform 39` 로 들어온다(에스원 실측). 이름을 믿으면
    거기서 죽는다. 그래서 순서를 이렇게 둔다:
        1) 이름이 맞으면 그것         (LG 계열 덱에서 가장 빠르고 정확하다)
        2) `Qu:` 접두 · Base 대괄호   (텍스트 자체가 정체를 말하는 슬롯)
        3) 표 위 밴드의 텍스트 도형   (heading -- 표 상단에서 가장 가까운 것)
    """
    tree = deck.xml(part)
    s = Slots(part=part)

    tbl = next(iter(tree.iter(q(A, "tbl"))), None)
    s.table = tbl
    for frame in tree.iter(q(P, "graphicFrame")):
        data = frame.find(f"{q(A, 'graphic')}/{q(A, 'graphicData')}")
        if data is not None and data.find(q(C, "chart")) is not None:
            s.chart_frames.append(frame)

    # 1) 이름
    s.title = D.find_shape(tree, "Slide Title")
    s.qu = D.find_shape(tree, "Qu")
    s.base = D.find_shape(tree, "Base/Unit")
    s.heading = D.find_shape(tree, "소제목")

    # 2) 텍스트가 스스로 밝히는 슬롯
    if s.qu is None or s.base is None:
        for sp in D.iter_shapes(tree):
            t = D.shape_text(sp).strip()
            if s.qu is None and t.startswith("Qu:"):
                s.qu = sp
            elif s.base is None and t.startswith("[Base"):
                s.base = sp

    # 3) 위치 -- 표 상단 위 2.2cm 밴드에서 표에 가장 가까운 텍스트 도형이 heading 이다
    known = {id(x) for x in (s.title, s.qu, s.base) if x is not None}
    tbl_top = _table_top(tbl) if tbl is not None else None
    if s.heading is None and tbl_top is not None:
        best, best_gap = None, None
        for sp in D.iter_shapes(tree):
            if id(sp) in known or not D.shape_text(sp).strip():
                continue
            top = _shape_top(sp)
            if top is None or top >= tbl_top:
                continue
            gap = tbl_top - top
            if gap > int(2.2 * D.EMU_PER_CM):
                continue
            if best_gap is None or gap < best_gap:
                best, best_gap = sp, gap
        s.heading = best

    # 소제목·멘트는 플레이스홀더다. top 순서로 가른다 (R-090)
    phs = [(int(x["top"] or 0), x) for x in D.placeholder_slots(tree)
           if x.get("idx") is not None and x.get("top") is not None]
    phs.sort()
    body_phs = [x for _, x in phs if x.get("name") != "Slide Title"]
    if len(body_phs) >= 2:
        s.sub = D.find_placeholder(tree, body_phs[0]["idx"])
        s.ment = D.find_placeholder(tree, body_phs[1]["idx"])
    if s.ment is None:
        # ★ 제작팀이 소제목을 그룹 안 자유 도형으로, 멘트를 플레이스홀더 아닌 텍스트 상자로
        #   만드는 경우가 있다(에스원 2차 회신본 실측: 멘트가 「텍스트 개체 틀 3」인데 idx 가
        #   없어 위 분기에 안 잡혔다). 그때 keep_ment=False 를 줘도 아무것도 안 바뀌고, 낡은
        #   수치가 든 멘트가 그대로 남았다(가이드 검수가 잡음, 덱 검수는 못 잡음).
        #   표 위에서 **가장 긴 텍스트 도형**을 멘트로 본다 — 멘트는 두세 줄이라 헤딩·소제목보다
        #   늘 길다. Qu·Base·제목·헤딩은 뺀다.
        skip = {id(x) for x in (s.title, s.qu, s.base, s.heading, s.sub) if x is not None}
        best, best_len = None, 0
        for sp in D.iter_shapes(tree):
            if id(sp) in skip:
                continue
            t = D.shape_text(sp).strip()
            if len(t) < 15 or t.startswith(("Qu:", "[Base")):
                continue
            top = _shape_top(sp)
            # 플레이스홀더는 xfrm 이 레이아웃에 있어 top 이 None 이다 — 위치를 모르면 배제하지 않는다.
            if top is not None and tbl_top is not None and top >= tbl_top:
                continue
            if len(t) > best_len:
                best, best_len = sp, len(t)
        s.ment = best
    return s


def _shape_top(sp: etree._Element) -> int | None:
    off = sp.find(f".//{q(A, 'off')}")
    return D._int(off.get("y")) if off is not None else None


def _table_top(tbl: etree._Element) -> int | None:
    return D.frame_box(D.table_frame(tbl)).get("y")


# =========================================================================== #
# 2. 정본 -- slide_facts.json
# =========================================================================== #

def load_facts(path: Path) -> list[dict[str, Any]]:
    """정본 본문 장표를 order 순서로. 표시값 반올림을 미리 계산해 둔다."""
    doc = json.loads(Path(path).read_text(encoding="utf-8"))
    slides = sorted(doc["slides"], key=lambda s: s.get("order") or s.get("seq") or 0)
    for s in slides:
        if s.get("grid"):
            _display_cells(s["grid"])
    return slides


def _display_cells(grid: dict[str, Any]) -> None:
    """원값 문자열 -> 표시값. 평균은 2자리, 나머지는 1자리 (R-013)."""
    for row in grid.get("rows") or []:
        digits = 2 if row.get("kind") == "mean" else 1
        disp: list[str | None] = []
        for raw in row.get("cells") or []:
            v = _num(raw)
            if v is None:
                disp.append(None if raw in (None, "") else str(raw).strip())
            else:
                disp.append(f"{_round_half_up(v, digits):.{digits}f}")
        row["_display"] = disp


def _num(v: Any) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(",", "")
    if t in ("-", "·", "*"):
        return None
    try:
        return float(t)
    except ValueError:
        return None


def _round_half_up(v: float, digits: int) -> float:
    from decimal import Decimal, ROUND_HALF_UP
    quant = Decimal(1).scaleb(-digits)
    return float(Decimal(repr(v)).quantize(quant, rounding=ROUND_HALF_UP))


def body_rows(grid: dict[str, Any]) -> list[dict[str, Any]]:
    """표·차트에 실을 행. 「계」를 뺀다."""
    return [r for r in (grid.get("rows") or [])
            if r.get("kind") not in BODY_SKIP_KINDS]


def row_label(row: dict[str, Any]) -> str:
    """표에 인쇄할 라벨. 통계표 표식 대괄호를 벗긴다 -- `[평균 : 층]` -> `평균(층)`."""
    t = (row.get("label_clean") or row.get("label") or "").strip()
    m = re.fullmatch(r"\[\s*(.+?)\s*[:：]\s*(.+?)\s*\]", t)
    if m:
        return f"{m.group(1)}({m.group(2)})"
    m = re.fullmatch(r"[\[【]\s*(.*?)\s*[\]】]", t)
    return (m.group(1) if m else t).strip()


# =========================================================================== #
# 3. 매핑 -- 순서가 조인 키다
# =========================================================================== #

@dataclass
class Pair:
    slide_no: int
    part: str
    fact: dict[str, Any]
    checks: dict[str, bool] = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return all(self.checks.values())


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", (s or "")).replace("’", "'").replace("‘", "'")


def pair_up(deck: D.Deck, facts: Sequence[dict[str, Any]],
            *, parts: Sequence[str] | None = None) -> list[Pair]:
    """본문 장표와 정본을 순서대로 짝짓고 **교차 검증**한다.

    순서만 믿지 않는다. Qu 텍스트·소제목·배너 열 수를 대조해 어긋난 자리를 짚는다.
    한 장이라도 어긋나면 호출자가 멈춰야 한다 -- 잘못 채운 장표는 검수에서 잡기 어렵다.
    """
    body = list(parts) if parts is not None else body_slides(deck)
    order = {p: i + 1 for i, p in enumerate(deck.slide_order())}
    if len(body) != len(facts):
        raise RuntimeError(
            f"본문 {len(body)}장인데 정본은 {len(facts)}장이다 -- 목차가 어긋났는가? "
            f"덱 본문 위치: {[order[p] for p in body][:8]}...")
    out: list[Pair] = []
    for part, fact in zip(body, facts):
        s = find_slots(deck, part)
        checks: dict[str, bool] = {}
        if s.qu is not None:
            checks["qu"] = _norm(D.shape_text(s.qu))[:40] == _norm(fact.get("qu") or "")[:40]
        if s.sub is not None and fact.get("sub_title"):
            checks["sub"] = _norm(D.shape_text(s.sub)).startswith(
                _norm(fact["sub_title"])[:8])
        if s.table is not None:
            want = len(fact["grid"]["banner_label_row"])
            checks["cols"] = D.n_cols(s.table) == want
        out.append(Pair(order[part], part, fact, checks))
    return out


# =========================================================================== #
# 4. 주입 -- 표
# =========================================================================== #

#: ±세로막대 장표의 표에서 지표 행을 알아보는 라벨 표식.
#: 「5점 평균(점)」·「TOP2(%)」 처럼 하우스마다 표기가 조금씩 다르므로 조각으로 본다.
SCALE_ROW_KINDS = (("mean", ("평균",)), ("top", ("TOP", "탑")))


def refill_table_scale(deck: D.Deck, part: str, fact: dict[str, Any], *,
                       on_warn: Callable[[str], None] | None = None) -> dict[str, Any]:
    """±세로막대 장표의 표. **행 순서가 아니라 라벨의 뜻으로** 자리를 찾는다.

    이 표는 구조가 뒤집혀 있다 -- 값 행(TOP2·평균)이 위, 배너 헤더가 아래다.
    게다가 제작팀이 두 행의 **순서를 바꿔 보낸다**(에스원 실측: 우리는 TOP2 가 먼저,
    회신본은 평균이 먼저). 그래서 행 번호로 짚으면 두 지표가 통째로 뒤바뀐다.
    라벨에 「평균」이 있으면 평균 행, 「TOP」이 있으면 TOP2 행이다.
    """
    warn = on_warn or (lambda _m: None)
    tree = deck.xml(part)
    tbl = next(iter(tree.iter(q(A, "tbl"))), None)
    if tbl is None:
        return {"skipped": "표 없음"}

    grid = fact["grid"]
    by_kind = {r.get("kind"): r for r in (grid.get("rows") or [])}
    n_cols = D.n_cols(tbl)
    want_cols = len(grid["banner_label_row"])
    if n_cols != want_cols:
        warn(f"열 수 {n_cols} != 정본 {want_cols} (R-095)")
        return {"skipped": f"열 수 불일치 {n_cols}/{want_cols}"}

    filled: list[str] = []
    for tr in D.table_rows(tbl):
        tcs = D.row_cells(tr)
        if not tcs:
            continue
        label = D.cell_text(tcs[0]).strip()
        if not label:
            continue
        kind = next((k for k, marks in SCALE_ROW_KINDS
                     if any(m in label for m in marks)), None)
        row = by_kind.get(kind) if kind else None
        if row is None:
            continue
        disp = row.get("_display") or []
        for i, tc in enumerate(tcs[1:]):
            if i < len(disp) and disp[i] is not None:
                D.put_cell(tc, disp[i])
        filled.append(f"{label}<-{kind}")
    if not filled:
        return {"skipped": "지표 행(평균·TOP2)을 못 찾았다"}
    deck.put(part, tree)
    return {"rows": filled, "cols": n_cols}


def refill_table(deck: D.Deck, part: str, fact: dict[str, Any], *,
                 fit_rows: bool = False,
                 on_warn: Callable[[str], None] | None = None) -> dict[str, Any]:
    """표의 값 셀을 정본으로 다시 채운다. (Base) 행과 배너 헤더는 손대지 않는다.

    「전체」 열은 **평균 행만** 채운다 -- 나머지 행에서는 차트가 그 자리를 덮는다.
    제작본 23장 전수 관찰에서 예외가 없었다(에스원 실측 2026-09-03).
    """
    warn = on_warn or (lambda _m: None)
    if chart_kind(fact.get("chart_type") or "") == "scale":
        return refill_table_scale(deck, part, fact, on_warn=warn)

    tree = deck.xml(part)
    tbl = next(iter(tree.iter(q(A, "tbl"))), None)
    if tbl is None:
        return {"skipped": "표 없음"}

    grid = fact["grid"]
    rows = body_rows(grid)
    n_cols = D.n_cols(tbl)
    want_cols = len(grid["banner_label_row"])
    if n_cols != want_cols:
        warn(f"열 수 {n_cols} != 정본 {want_cols} -- 배너 절단이 다르다 (R-095)")
        return {"skipped": f"열 수 불일치 {n_cols}/{want_cols}"}

    # ── (Base) 행 ─────────────────────────────────────────────────────
    # ★ 여기를 「제작팀이 손본 것」으로 보고 보존하면 안 된다. 도형을 대표 장에서
    #   떠 온 장표는 **대표의 표본 수가 그대로 따라온다** -- 실측: 문항 응답자
    #   286·266 명짜리 장표에 대표의 (500) 이 실렸다. 「500명 중 …」 으로 읽히는
    #   오독을 만들고, 저베이스 경고도 사라진다.
    #   정본 base_row 는 저베이스 별표 `(21*)` 까지 담고 있으므로 그대로 쓰면 된다.
    base_fixed = 0
    want_base = grid.get("base_row") or []
    head = D.table_rows(tbl)[:D.HEADER_ROWS]
    if len(head) >= D.HEADER_ROWS and want_base:
        for j, tc in enumerate(D.row_cells(head[D.HEADER_ROWS - 1])):
            if j == 0 or j >= len(want_base):
                continue                       # 0열은 「(Base)」 라벨 자리다
            if D.cell_text(tc).strip() != str(want_base[j]):
                D.put_cell(tc, str(want_base[j]))
                base_fixed += 1

    trs = D.data_rows(tbl, D.HEADER_ROWS)
    mode = "order"
    if len(trs) != len(rows):
        # ★ 순서가 중요하다. 라벨 매칭을 먼저 하면 **정본에만 있는 보기가 조용히
        #   빠진다** -- 실측 s48: 새 통계표가 「상황에 따라 다름」을 새 보기로 주었는데
        #   덱의 7행이 정본 8행 중 7개와 라벨로 붙어 버려 그 보기가 안 실렸다.
        #   그래서 `--fit-rows` 를 준 자리에서는 **행 수를 맞추는 쪽이 먼저**다.
        #   라벨 매칭은 행 수를 못 바꾸는 자리(제작팀이 의도적으로 줄인 장표)의
        #   안전장치로 남는다.
        paired = None if fit_rows else _pair_rows_by_label(trs, rows)
        if paired is not None:
            trs, rows, mode = [x for x, _ in paired], [y for _, y in paired], "label"
        elif fit_rows:
            trs = D.set_row_count(tbl, len(rows), header_rows=D.HEADER_ROWS,
                                  mid_style="derive", on_warn=warn)
            mode = "fit"
        else:
            warn(f"데이터 행 {len(trs)} != 정본 {len(rows)} -- 라벨로도 못 붙였다. "
                 f"꼬리 접기 판단이 얽혀 있어 값만 갈지 않는다. --fit-rows 로 맞출 수 있다")
            return {"skipped": f"행 수 불일치 {len(trs)}/{len(rows)}"}

    n_filled = 0
    for tr, row in zip(trs, rows):
        tcs = D.row_cells(tr)
        if mode != "label":
            D.put_cell(tcs[0], row_label(row))   # 라벨 매칭이면 덱 라벨을 존중한다
        disp = row.get("_display") or []
        is_mean = row.get("kind") == "mean"
        for i, tc in enumerate(tcs[1:]):
            if i >= len(disp):
                D.put_cell(tc, "")
                continue
            if i == 0 and not is_mean:
                continue                       # 「전체」 열은 차트가 덮는다
            v = disp[i]
            D.put_cell(tc, "-" if v is None else v)
            n_filled += 1
    deck.put(part, tree)
    return {"rows": len(rows), "cols": n_cols, "cells": n_filled, "match": mode,
            "base_fixed": base_fixed}


def _pair_rows_by_label(trs: Sequence[etree._Element],
                        rows: Sequence[dict[str, Any]]
                        ) -> list[tuple[etree._Element, dict[str, Any]]] | None:
    """덱의 데이터 행을 라벨로 정본 행에 붙인다. 하나라도 못 붙으면 None.

    덱 라벨은 줄바꿈·자동 말줄임이 섞이므로 공백을 없앤 접두 비교로 본다.
    한 행이라도 정본에서 짝을 못 찾으면 **전부 포기한다** -- 절반만 채운 표가
    가장 위험하다(옛 값과 새 값이 한 표에 섞인다).
    """
    index: dict[str, dict[str, Any]] = {}
    for r in rows:
        key = _norm(row_label(r))
        if key:
            index.setdefault(key, r)
    out: list[tuple[etree._Element, dict[str, Any]]] = []
    for tr in trs:
        tcs = D.row_cells(tr)
        if not tcs:
            return None
        key = _norm(D.cell_text(tcs[0]))
        if not key:
            return None
        hit = index.get(key)
        if hit is None:
            hit = next((r for k, r in index.items()
                        if k.startswith(key[:10]) or key.startswith(k[:10])), None)
        if hit is None:
            return None
        out.append((tr, hit))
    return out


# =========================================================================== #
# 5. 주입 -- 차트
# =========================================================================== #

def chart_kind(spec: str) -> str:
    """차트타입 문자열 -> 주입 방식. 정본 표기를 그대로 읽는다."""
    t = (spec or "").replace(" ", "")
    if "세로막대" in t or "세로누적" in t:
        return "scale"
    if "겹침막대" in t:
        return "rank"
    return "bar"


def rank_names(grid: dict[str, Any]) -> tuple[str, str]:
    """(누적 순위 이름, 1순위 이름). 위치가 아니라 이름의 뜻으로 고른다."""
    names = list(grid.get("rank_columns") or [])
    cum = next((n for n in names if any(k in n for k in CUMULATIVE_RANK)), None)
    first = next((n for n in names if n != cum), None)
    return cum or (names[0] if names else ""), first or ""


def refill_chart(deck: D.Deck, part: str, fact: dict[str, Any], *,
                 on_warn: Callable[[str], None] | None = None) -> dict[str, Any]:
    """차트 값을 정본으로 다시 채운다. graphicFrame 은 열지 않는다 (R-106).

    ★ 계열 수는 **덱이 진리**다(R-094·R-096). 제작본은 정본보다 계열이 하나 많은
      경우가 흔하다 -- 범례에 「없음」 자리를 미리 만들어 두기 때문이다
      (제작팀 메모 「범례만큼 데이터가 없을때는 빈칸으로 유지해주세요」).
      남는 계열은 **0 이 아니라 빈칸(None)** 으로 채운다.
    """
    warn = on_warn or (lambda _m: None)
    cps = deck.chart_parts(part)
    if not cps:
        return {"skipped": "차트 없음"}
    cp = cps[0]
    grid = fact["grid"]
    rows = body_rows(grid)
    kind = chart_kind(fact.get("chart_type") or "")

    if kind == "scale":
        return refill_chart_scale(deck, part, fact, on_warn=warn)

    axis = [r if r.get("kind") == "option" else None for r in rows]
    cats = [(r.get("label_clean") if r else "") for r in axis]

    if kind == "rank":
        # ★ 두 계열의 값은 **배너 열이 아니다.** 순위 겹침막대는 「전체」 열 하나에
        #   두 지표를 겹쳐 그린다:
        #       누적(1+N순위) = cells[0]            (전체 열)
        #       1순위        = rank_first_total    (별도 필드)
        #   `rank_columns` 의 인덱스를 cells 인덱스로 쓰면 「전체 열」과 「첫 배너 열」
        #   두 개를 두 계열로 넣게 되어, 1순위가 누적보다 큰 그림이 나온다
        #   (실측 s50: 공동주택 50.0 이 누적 자리에, 전체 54.5 가 1순위 자리에 들어갔다).
        #   파일은 멀쩡히 열리고 표 값도 맞아서 수치 검사가 통과한다 -- 그래서 위험하다.
        cum, first = rank_names(grid)
        # 순위 겹침막대: 긴 값(누적)이 order=0, 1순위가 그 위에 겹친다 (R-109)
        want = [[(_num(r["cells"][0]) if r else None) for r in axis],
                [(_num(r.get("rank_first_total")) if r else None) for r in axis]]
        headers: list[str | None] = [cum, first]
    else:
        want = [[(_num(r["cells"][0]) if r else None) for r in axis]]
        headers = ["전체"]

    tree = deck.xml(cp)
    have = len(D.chart_series(tree))
    if have > len(want):
        want = list(want) + [[None] * len(cats) for _ in range(have - len(want))]
        headers = list(headers) + [None] * (have - len(headers))
    elif have < len(want):
        warn(f"차트 계열 {have}개인데 정본은 {len(want)}세트다 -- 계열 수를 바꾸면 "
             f"색·라벨 서식이 어긋난다(R-096). 사람이 봐야 한다")
        return {"skipped": f"계열 수 부족 {have}/{len(want)}"}

    res = D.apply_chart(deck, cp, want, cats, series_headers=headers, on_warn=warn)
    set_series_names(deck, cp, headers)
    return {"chart": cp, "n": len(cats), "series": have, "warnings": res["warnings"]}


def set_series_names(deck: D.Deck, chart_part: str,
                     names: Sequence[str | None]) -> int:
    """차트의 계열명 캐시(`c:tx/strCache`)를 갱신한다.

    안 하면 워크시트에는 「1~2순위」인데 차트 캐시에는 대표 장에서 딸려온 옛 이름이
    남아, PowerPoint 「데이터 편집」에서 둘이 어긋나 보인다. `apply_chart` 의
    `series_headers` 는 워크시트만 고친다.
    """
    tree = deck.xml(chart_part)
    n = 0
    for ser, name in zip(D.chart_series(tree), names):
        if name is None:
            continue
        for v in ser.findall(f"{q(C, 'tx')}//{q(C, 'v')}"):
            v.text = name
            n += 1
    deck.put(chart_part, tree)
    return n


def _pick(row: dict[str, Any] | None, idx: int) -> float | None:
    if row is None:
        return None
    cells = row.get("cells") or []
    return _num(cells[idx]) if idx < len(cells) else None


# --------------------------------------------------------------------------- #
# 5-b. ±세로막대 -- 계열은 척도 단계, 범주는 배너 열. 워크북이 수식으로 굴러간다
# --------------------------------------------------------------------------- #

#: 척도 워크북의 입력 블록 머리. 이 아래 행들이 원값(양수)을 담는다.
INPUT_MARK = "▼데이터 입력"

#: 음수 계열을 알아보는 수식 꼴. 단항 마이너스(`=-B11`)와 곱셈(`=B14*-1`)이 둘 다 쓰인다.
#: ★ `"-1" in f` 로 보면 `=-B11` 을 못 잡아 부호가 사라진다(실제로 물렸다).
NEG_FORMULA = re.compile(r"=\s*-|\*\s*-\s*1")


def scale_series_plan(deck: D.Deck, chart_part: str
                      ) -> tuple[list[int], list[bool], list[int]]:
    """(계열별 워크시트 입력 행, 계열별 음수 여부, 입력 블록 전체 행).

    ★ 척도 차트의 워크북은 두 층이다 --
        계열 행    차트가 읽는 행. 전부 수식(`=B15` · `=B14*-1`)
        입력 블록  사람이 채우는 원값(양수) 자리
      ± 부호는 그 수식의 `*-1` 이 만든다. **값을 계열 행에 직접 쓰면 부호 장치가
      깨진다.** 그래서 어느 계열이 음수 자리인지도 수식에서 읽는다 -- 현재 값으로
      판정하면 값이 0인 계열(응답 0단계)에서 갈리지 않는다.
      행 번호를 상수로 박지 않는 이유: 원본이 바뀌면 조용히 어긋나느니 못 찾는 게 낫다.
    """
    import io
    import openpyxl
    part = deck.workbook_part(chart_part)
    if part is None or part not in deck.parts:
        return [], [], []
    ws = openpyxl.load_workbook(io.BytesIO(deck.parts[part])).worksheets[0]
    lay = D.chart_layout(deck.xml(chart_part))
    rows: list[int] = []
    negs: list[bool] = []
    for r in lay.val_rows:
        if not r:
            rows.append(0)
            negs.append(False)
            continue
        f = str(ws.cell(r, 2).value or "")
        m = re.search(r"B(\d+)", f)
        rows.append(int(m.group(1)) if m else 0)
        negs.append(bool(NEG_FORMULA.search(f)))
    block: list[int] = []
    for r in range(1, (ws.max_row or 0) + 1):
        if INPUT_MARK in (str(ws.cell(r, 1).value or "") + str(ws.cell(r, 2).value or "")):
            block = list(range(r + 1, (ws.max_row or 0) + 1))
            break
    return rows, negs, block


def refill_chart_scale(deck: D.Deck, part: str, fact: dict[str, Any], *,
                       on_warn: Callable[[str], None] | None = None) -> dict[str, Any]:
    """±세로막대 차트 값을 정본으로 다시 채운다.

    **계열 구조는 덱이 진리다.** 응답 0인 척도 단계는 이미 솎여 있고(실측 s87 은
    5단계 문항인데 계열이 4개다), 다시 솎으려 들면 색·범례가 어긋난다(R-096).
    그래서 계열을 **이름으로** 정본 보기 행에 붙이고 값만 넣는다.
    """
    warn = on_warn or (lambda _m: None)
    cps = deck.chart_parts(part)
    if not cps:
        return {"skipped": "차트 없음"}
    cp = cps[0]
    tree = deck.xml(cp)
    names = D.series_names(tree)
    grid = fact["grid"]
    opts = {_norm(r.get("label_clean") or ""): r
            for r in (grid.get("rows") or []) if r.get("kind") == "option"}

    rows, negs, block = scale_series_plan(deck, cp)
    if not rows or not block:
        warn("척도 워크북의 입력 블록을 못 찾았다 -- 구조가 바뀌었다")
        return {"skipped": "입력 블록 없음"}

    # 범주는 보기가 아니라 **배너 열**이다(표와 축이 뒤집혀 있다).
    cats = ["전체"] + [x for x in grid["banner_label_row"][2:]]
    series: list[list[float | None]] = []
    unmatched: list[str] = []
    for nm, neg in zip(names, negs):
        row = opts.get(_norm(nm or ""))
        if row is None:
            unmatched.append(nm or "")
            series.append([None] * len(cats))
            continue
        vals = [_num(x) for x in (row.get("cells") or [])]
        vals = (vals + [None] * len(cats))[:len(cats)]
        series.append([(-v if (neg and v is not None) else v) for v in vals])
    if unmatched:
        warn(f"계열 {unmatched} 를 정본 보기 행에서 못 찾았다 -- 빈칸으로 두었다")

    res = D.apply_chart(deck, cp, series, cats, series_headers=list(names),
                        preserve_formulas=True, on_warn=warn)
    set_series_names(deck, cp, list(names))

    # ★ 입력 블록을 따로 채운다. apply_chart 는 수식 셀을 보존하느라 못 건드린다.
    #   안 하면 화면은 맞는데 「데이터 편집」을 열면 옛 값이 그대로 보인다.
    cells: dict[tuple[int, int], Any] = {}
    for j, (r, nm) in enumerate(zip(rows, names)):
        if not r or j >= len(series):
            continue
        cells[(r, 1)] = nm
        for c, v in enumerate(series[j]):
            cells[(r, 2 + c)] = None if v is None else abs(v)   # 부호는 수식이 붙인다
    for r in block:
        if r in rows:
            continue
        for c in range(len(cats)):
            cells[(r, 2 + c)] = None                            # 옛 값 잔재 제거
    D.write_workbook_cells(deck, cp, cells)
    return {"chart": cp, "n": len(cats), "series": len(series),
            "workbook_cells": len(cells),
            "warnings": [w for w in res["warnings"] if "수식 셀" not in w]}


# =========================================================================== #
# 6. 주입 -- 텍스트
# =========================================================================== #

def refill_text(deck: D.Deck, part: str, fact: dict[str, Any], *,
                keep_ment: bool = True,
                on_warn: Callable[[str], None] | None = None) -> dict[str, Any]:
    """제목·소제목·표제목·Qu·Base/Unit. 멘트는 기본으로 보존한다.

    keep_ment
      제작본에 이미 들어온 멘트는 사람이 손봤을 수 있다. 값이 바뀌어 멘트도 다시
      써야 하면 정본(업무 A)에서 고쳐 내려오는 것이 정상 경로다(헌법 1).
    """
    warn = on_warn or (lambda _m: None)
    s = find_slots(deck, part)
    tree = deck.xml(part)
    done: dict[str, Any] = {}

    if s.title is not None and fact.get("slide_title"):
        done["title"] = _set_text(s.title, fact["slide_title"])
    if s.sub is not None and fact.get("sub_title"):
        done["sub"] = _set_text(s.sub, fact["sub_title"], split="|")
    if s.heading is not None and fact.get("title_full"):
        done["heading"] = _set_text(s.heading, fact["title_full"])
    if s.qu is not None and fact.get("qu"):
        done["qu"] = _set_text(s.qu, fact["qu"])
    if s.base is not None and fact.get("base_unit"):
        # ★ 각주 문단을 **지우지 않는다.** 정본의 base_note 가 비어 있어도 제작본에는
        #   「*Too small base(n<30)」 가 들어와 있는 경우가 있다 — 정본의 결손을
        #   제작팀이 메꿔 준 것이다(에스원 실측 69장 전수). 정본대로 문단을 다시
        #   쓰면 그 각주가 통째로 사라지고, 저베이스 열이 주석 없이 실린다.
        # 별표만 다르면 덱을 존중한다. `(n=20*)` 의 별표는 「이 Base 는 30 미만」
        # 이라는 표식이고, 정본 base_unit 에는 없는 경우가 있다(에스원 실측 6장).
        # 정본으로 덮어쓰면 저베이스 표식이 조용히 사라진다.
        have = D.shape_text(s.base).split("\n")[0]
        want = fact["base_unit"]
        if _norm(have.replace("*", "")) != _norm(want.replace("*", "")):
            done["base"] = _set_paragraph(s.base, 0, want)
        note = fact.get("base_note") or ""
        if note:
            done["base_note"] = _set_paragraph(s.base, 1, note)
    if not keep_ment and s.ment is not None:
        done["ment"] = _set_lines(s.ment, fact.get("ment") or [])
    for k, v in done.items():
        if v is False:
            warn(f"{k} 주입 실패")
    deck.put(part, tree)
    return done


def _set_text(sp: etree._Element, text: str, *, split: str | None = None) -> bool:
    body = sp.find(q(P, "txBody"))
    if body is None:
        return False
    ts = body.findall(f".//{q(A, 't')}")
    if not ts:
        return False
    D._write_runs(ts, text, split=split)
    return True


def _set_paragraph(sp: etree._Element, index: int, text: str) -> bool:
    """도형의 `index` 번째 **줄**만 갈아 끼우고 나머지 줄은 그대로 남긴다.

    ★ 줄이 어떻게 나뉘어 있는지를 가정하면 안 된다. 세 가지가 다 쓰인다:
        문단(`a:p`) 경계 · `<a:br/>` · **런 텍스트 안의 `\\n` 문자**
      제작팀 회신본은 세 번째였다(에스원 실측 69장 -- 문단 1개·br 0개인데
      마지막 런이 `]\\n*Too small base(n<30)`). 앞 둘만 보면 「한 줄 상자」로
      읽히고, 그 줄을 갈아 끼울 때 각주가 통째로 사라진다.
      그래서 **이어붙인 전체 텍스트에서 줄을 세고, 고친 전체를 다시 쓴다.**
    """
    body = sp.find(q(P, "txBody"))
    if body is None:
        return False
    ts = body.findall(f".//{q(A, 't')}")
    if not ts:
        return False
    cur = D.shape_text(sp)
    lines = cur.split("\n")
    while len(lines) <= index:
        lines.append("")
    lines[index] = text
    D._write_runs(ts, "\n".join(lines))
    return True


def legend_lines(deck: D.Deck, part: str, *, mark: str = "■") -> list[str]:
    """장표에서 범례 줄을 읽는다. 표 셀에 있든 별도 상자에 있든 찾는다.

    하우스마다 자리가 다르다 -- 우리 조립본은 표의 라벨 셀에 넣고, 제작팀 회신본은
    표 밖 「범례」 텍스트 상자에 넣는다. 척도 라벨은 **문항마다 다르므로**
    (만족도 「매우 만족」 · 추천의향 「매우 추천 의향이 있다」) 남의 장에서 복제한
    상자를 그대로 두면 문항과 어긋난 범례가 실린다.
    """
    tree = deck.xml(part)
    for sp in D.iter_shapes(tree):
        t = D.shape_text(sp)
        if mark in t:
            return [x for x in t.split("\n") if x.strip()]
    tbl = next(iter(tree.iter(q(A, "tbl"))), None)
    if tbl is not None:
        for tr in D.table_rows(tbl):
            tcs = D.row_cells(tr)
            if tcs and mark in D.cell_text(tcs[0]):
                return [x for x in D.cell_text(tcs[0]).split("\n") if x.strip()]
    return []


def set_legend(deck: D.Deck, part: str, lines: Sequence[str],
               *, mark: str = "■") -> bool:
    """장표의 범례 상자 텍스트를 갈아 끼운다. 상자의 서식·위치는 그대로 둔다."""
    if not lines:
        return False
    tree = deck.xml(part)
    for sp in D.iter_shapes(tree):
        if mark in D.shape_text(sp):
            ok = _set_lines(sp, list(lines))
            if ok:
                deck.put(part, tree)
            return ok
    return False


def _set_lines(sp: etree._Element, lines: Sequence[str]) -> bool:
    """여러 문단. 첫 문단 서식을 나머지가 물려받는다. <a:br/> 로 잇지 않는다."""
    body = sp.find(q(P, "txBody"))
    if body is None:
        return False
    paras = body.findall(q(A, "p"))
    if not paras:
        return False
    proto = [copy.deepcopy(x) for x in paras]
    for old in paras:
        body.remove(old)
    for i, line in enumerate(lines):
        para = copy.deepcopy(proto[min(i, len(proto) - 1)])
        ts = para.findall(f".//{q(A, 't')}")
        if not ts:
            continue
        ts[0].text = line
        for extra in ts[1:]:
            extra.text = ""
        body.append(para)
    return True


# =========================================================================== #
# 7. 이식 -- 도형이 없는 장에 표·차트를 옮겨 심는다
# =========================================================================== #

def graft_table(deck: D.Deck, dst_part: str, src_part: str, *,
                src_deck: D.Deck | None = None) -> bool:
    """표 graphicFrame 을 통째로 옮겨 심는다. 표는 관계가 없어 복사로 끝난다.

    제작팀이 「18p 동일 포맷」 이라고만 적고 도형을 비워 보내는 관행이 있다
    (에스원 실측 36장). 그때 대표 장의 표를 이 함수로 심고 값을 채운다.
    """
    src = (src_deck or deck).xml(src_part)
    frame = None
    for tbl in src.iter(q(A, "tbl")):
        frame = D.table_frame(tbl)
        break
    if frame is None:
        return False
    dst = deck.xml(dst_part)
    sp_tree = dst.find(f"{q(P, 'cSld')}/{q(P, 'spTree')}")
    if sp_tree is None:
        return False
    for old in list(dst.iter(q(A, "tbl"))):
        old_frame = D.table_frame(old)
        if old_frame is not None:
            old_frame.getparent().remove(old_frame)
    sp_tree.append(copy.deepcopy(frame))
    deck.put(dst_part, dst)
    return True


def graft_shape(deck: D.Deck, dst_part: str, src_part: str, *,
                name: str | None = None,
                text_mark: str | None = None,
                src_deck: D.Deck | None = None,
                skip_if_present: bool = True) -> str | None:
    """대표 장에 있고 대상 장에 없는 **보조 도형**을 옮겨 심는다.

    범례 상자처럼 표·차트에 딸리지 않은 도형이 그 대상이다. 제작팀은 이런 도형을
    대표 장에만 만들어 두고 클론 장에는 넣지 않는다(에스원 실측: ±세로막대 범례가
    s35·s36 에만 있었고 나머지 19장에는 없었다).

    name        도형 이름으로 찾는다
    text_mark   이름을 못 믿을 때 쓰는 텍스트 표식(예 「■」)
    skip_if_present  대상 장에 같은 도형이 이미 있으면 넣지 않는다 -- **중복 방지가
                핵심이다.** 표 셀에 범례를 또 써 넣었더니 s36 에 범례가 두 벌
                겹쳐 보였다(실측).
    """
    sdeck = src_deck or deck
    src = sdeck.xml(src_part)

    def pick(tree):
        for sp in D.iter_shapes(tree):
            if name and D.shape_name(sp) == name:
                return sp
            if text_mark and text_mark in D.shape_text(sp):
                return sp
        return None

    node = pick(src)
    if node is None:
        return None
    dst = deck.xml(dst_part)
    if skip_if_present and pick(dst) is not None:
        return None
    sp_tree = dst.find(f"{q(P, 'cSld')}/{q(P, 'spTree')}")
    if sp_tree is None:
        return None
    sp_tree.append(copy.deepcopy(node))
    deck.put(dst_part, dst)
    return D.shape_name(node) or name or text_mark


def graft_chart(deck: D.Deck, dst_part: str, src_part: str, *,
                src_deck: D.Deck | None = None,
                frame_from: str | None = None) -> str | None:
    """차트를 옮겨 심는다. 차트 파트·서식·워크북까지 **새로 떠서** 붙인다.

    ★ 파트를 공유하면 한 장을 고칠 때 다른 장이 함께 바뀐다. 반드시 복제다.
    frame_from
      차트 **배치**를 이 장표에서 빌린다 -- graphicFrame 의 위치·크기와 함께
      `plotArea/layout/manualLayout` 까지 옮긴다.

      ★ 프레임만 빌리면 막대가 어긋난다. 막대 영역은 프레임 안의 **비율**로
        정해지므로, 폭이 다른 프레임에 남의 비율을 얹으면 시작점이 밀린다.
        실측: 제작팀 프레임(2016000·x 9.1%·w 59.6%)에 우리 비율(x 5.6%·w 75.3%)을
        얹었더니 막대가 0.2cm 왼쪽으로 밀려 **표의 라벨 열을 덮었다**
        (「0개」가 「0기」로 잘려 읽혔다).
    """
    sdeck = src_deck or deck
    src = sdeck.xml(src_part)
    frame = None
    for f in src.iter(q(P, "graphicFrame")):
        data = f.find(f"{q(A, 'graphic')}/{q(A, 'graphicData')}")
        if data is not None and data.find(q(C, "chart")) is not None:
            frame = f
            break
    if frame is None:
        return None

    src_rels = {r["Id"]: r["Part"] for r in sdeck.rels(src_part) if r["Part"]}
    node = copy.deepcopy(frame)
    chart_el = node.find(f"{q(A, 'graphic')}/{q(A, 'graphicData')}/{q(C, 'chart')}")
    old_rid = chart_el.get(q(R, "id")) if chart_el is not None else None
    old_part = src_rels.get(old_rid or "")
    if not old_part:
        return None

    namer = B._Namer(deck)
    new_part = _copy_private(deck, sdeck, old_part, namer, {})
    new_rid = _add_rel(
        deck, dst_part, new_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")
    chart_el.set(q(R, "id"), new_rid)

    dst = deck.xml(dst_part)
    if frame_from:
        borrow = deck.xml(frame_from)
        for f in borrow.iter(q(P, "graphicFrame")):
            data = f.find(f"{q(A, 'graphic')}/{q(A, 'graphicData')}")
            if data is None or data.find(q(C, "chart")) is None:
                continue
            _copy_frame_box(f, node)
            lend = data.find(q(C, "chart"))
            lend_rid = lend.get(q(R, "id")) if lend is not None else None
            lend_part = {r["Id"]: r["Part"]
                         for r in deck.rels(frame_from) if r["Part"]}.get(lend_rid or "")
            if lend_part:
                _copy_plot_layout(deck.xml(lend_part), deck.xml(new_part))
                deck.put(new_part, deck.xml(new_part))
            break
    sp_tree = dst.find(f"{q(P, 'cSld')}/{q(P, 'spTree')}")
    dropped_rids: list[str] = []
    for old in list(dst.iter(q(P, "graphicFrame"))):
        data = old.find(f"{q(A, 'graphic')}/{q(A, 'graphicData')}")
        if data is None or data.find(q(C, "chart")) is None:
            continue
        ch = data.find(q(C, "chart"))
        rid = ch.get(q(R, "id")) if ch is not None else None
        if rid and rid != new_rid:
            dropped_rids.append(rid)
        old.getparent().remove(old)
    sp_tree.append(node)
    deck.put(dst_part, dst)
    # ★ 도형만 지우고 관계를 남기면 **옛 차트 파트가 계속 먼저 잡힌다.**
    #   `chart_parts()` 는 rels 순서를 따르므로 `cps[0]` 이 지운 더미를 가리키고,
    #   값 주입도 감사도 그 유령을 상대한다(실측 s36 -- 새 차트를 붙였는데도
    #   ±20.0 더미가 계속 보고됐다). 관계까지 끊어야 이식이 끝난다.
    if dropped_rids:
        _drop_rels(deck, dst_part, dropped_rids)
    B.sync_content_types(deck)
    return new_part


def _drop_rels(deck: D.Deck, part: str, rids: Sequence[str]) -> int:
    """파트의 .rels 에서 주어진 rId 항목을 지운다. 남은 파트는 gc 가 걷는다."""
    name = D.Deck.rels_name(part)
    if name not in deck.parts:
        return 0
    tree = etree.fromstring(deck.parts[name])
    n = 0
    for rel in list(tree):
        if rel.get("Id") in set(rids):
            tree.remove(rel)
            n += 1
    if n:
        deck.put_bytes(name, etree.tostring(
            tree, xml_declaration=True, encoding="UTF-8", standalone=True))
    return n


def _copy_plot_layout(src_chart: etree._Element, dst_chart: etree._Element) -> bool:
    """`plotArea/layout` 을 통째로 옮긴다. 막대가 프레임 안에서 앉는 자리다."""
    s = src_chart.find(f".//{q(C, 'plotArea')}/{q(C, 'layout')}")
    if s is None:
        return False
    area = dst_chart.find(f".//{q(C, 'plotArea')}")
    if area is None:
        return False
    old = area.find(q(C, "layout"))
    node = copy.deepcopy(s)
    if old is not None:
        area.replace(old, node)
    else:
        area.insert(0, node)
    return True


def _copy_frame_box(src_frame: etree._Element, dst_frame: etree._Element) -> None:
    """graphicFrame 의 p:xfrm(off·ext)을 그대로 옮긴다."""
    s = src_frame.find(q(P, "xfrm"))
    d = dst_frame.find(q(P, "xfrm"))
    if s is None or d is None:
        return
    for tag in ("off", "ext"):
        se, de = s.find(q(A, tag)), d.find(q(A, tag))
        if se is not None and de is not None:
            for k, v in se.attrib.items():
                de.set(k, v)


def _copy_private(deck: D.Deck, src_deck: D.Deck, part: str,
                  namer: Any, seen: dict[str, str]) -> str:
    """차트/워크북/서식 파트를 다른 덱에서 새 이름으로 떠 온다. rels 까지 재귀."""
    if part in seen:
        return seen[part]
    new = namer.take_like(part)
    seen[part] = new
    deck.put_bytes(new, src_deck.parts[part])
    rels_name = D.Deck.rels_name(part)
    if rels_name not in src_deck.parts:
        return new
    tree = etree.fromstring(src_deck.parts[rels_name])
    src_dir, dst_dir = posixpath.dirname(part), posixpath.dirname(new)
    for rel in tree:
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target") or ""
        tgt = posixpath.normpath(posixpath.join(src_dir, target))
        if B.PRIVATE_ROOTS.match(tgt) and tgt in src_deck.parts:
            mapped = _copy_private(deck, src_deck, tgt, namer, seen)
        else:
            mapped = tgt
        rel.set("Target", posixpath.relpath(mapped, dst_dir) if dst_dir else mapped)
    deck.put_bytes(D.Deck.rels_name(new), etree.tostring(
        tree, xml_declaration=True, encoding="UTF-8", standalone=True))
    return new


def _add_rel(deck: D.Deck, part: str, target: str, rel_type: str) -> str:
    """파트의 .rels 에 관계를 추가하고 새 rId 를 준다."""
    name = D.Deck.rels_name(part)
    NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    if name in deck.parts:
        tree = etree.fromstring(deck.parts[name])
    else:
        tree = etree.Element(f"{{{NS}}}Relationships")
    used = set()
    for rel in tree:
        m = re.fullmatch(r"rId(\d+)", rel.get("Id") or "")
        if m:
            used.add(int(m.group(1)))
    rid = f"rId{max(used) + 1 if used else 1}"
    rel = etree.SubElement(tree, f"{{{NS}}}Relationship")
    rel.set("Id", rid)
    rel.set("Type", rel_type)
    rel.set("Target", posixpath.relpath(target, posixpath.dirname(part)))
    deck.put_bytes(name, etree.tostring(
        tree, xml_declaration=True, encoding="UTF-8", standalone=True))
    return rid


# =========================================================================== #
# 8. 잔재 제거 -- 제작팀 쪽지
# =========================================================================== #

#: 제작팀 쪽지로 판정하는 상단 밴드. 슬라이드 맨 위에 걸친 채움 도형이다.
MEMO_MAX_Y = int(1.5 * D.EMU_PER_CM)

#: 쪽지 텍스트 패턴 -- 「데이터 입력」·「18p 동일 포맷」·「항목 10개만 작업」 …
MEMO_RE = re.compile(r"(데이터\s*입력|동일\s*포맷|Data\s*입력|확인해주세요|"
                     r"각주\s*추가|맨\s*아래로|유지해주세요|항목\s*\d+개만)")


def drop_memos(deck: D.Deck, part: str, *,
               drop_big_pictures: bool = False) -> list[str]:
    """제작팀 쪽지 도형을 걷어낸다. 이름이 아니라 **위치와 텍스트**로 잡는다.

    drop_big_pictures
      제작팀이 작업 지시를 **그림으로** 붙여 보내는 경우가 있다(에스원 s23 실측:
      「보기 개수 -1 로 셀 분할」 안내가 전면 캡처로 들어왔다). 슬라이드를 거의
      덮는 그림은 배경이 아니라 지시서다 -- 납품물에 실리면 사고다.
    """
    tree = deck.xml(part)
    dropped: list[str] = []
    for sp in list(D.iter_shapes(tree)):
        top = _shape_top(sp)
        text = D.shape_text(sp).strip()
        if top is not None and top < MEMO_MAX_Y and text and MEMO_RE.search(text):
            parent = sp.getparent()
            if parent is not None:
                parent.remove(sp)
                dropped.append(text[:30])
    if drop_big_pictures:
        for pic in list(tree.iter(q(P, "pic"))):
            # ★ `.//a:ext` 로 찾으면 안 된다 -- 같은 태그 이름이 두 뜻으로 쓰인다.
            #   `<a:ext uri="{FF2B5EF4-…}">` 는 확장 목록이고 크기가 아니다.
            #   실측: 그 자리를 잡으면 cx 가 없어 전면 그림이 조용히 안 걸린다.
            ext = pic.find(f"{q(P, 'spPr')}/{q(A, 'xfrm')}/{q(A, 'ext')}")
            if ext is None:
                continue
            cx, cy = D._int(ext.get("cx")), D._int(ext.get("cy"))
            if cx > int(25 * D.EMU_PER_CM) and cy > int(12 * D.EMU_PER_CM):
                parent = pic.getparent()
                if parent is not None:
                    parent.remove(pic)
                    dropped.append(f"[전면 그림 {cx / D.EMU_PER_CM:.0f}x"
                                   f"{cy / D.EMU_PER_CM:.0f}cm]")
    if dropped:
        deck.put(part, tree)
    return dropped


# =========================================================================== #
# 8-b. 차트 전수 감사 -- 표는 verify.py 가 보고, 차트는 여기서 본다
# =========================================================================== #

#: 두 계열의 값이 이만큼(축 최대의 비율) 안쪽으로 붙으면 데이터 라벨이 겹친다.
#: 실측 근거: s50 「27.3/45.5」(차이 18.2, 축 54.5 -> 33%)는 붙어 읽혔고
#: 「27.3/54.5」(차이 27.2 -> 50%)는 떨어졌다. 경계를 40% 로 잡는다.
LABEL_CLASH_RATIO = 0.40


def audit_charts(deck: D.Deck, facts: Sequence[dict[str, Any]]
                 ) -> list[dict[str, Any]]:
    """본문 차트를 정본과 전수 대조한다. **표가 맞아도 차트는 틀릴 수 있다.**

    실제로 물렸던 자리 -- 순위 겹침막대의 두 계열에 「전체 열」과 「첫 배너 열」을
    넣었더니 표 값은 전부 맞고 파일도 멀쩡히 열렸다. 1순위가 누적보다 큰 그림만
    남았고, 수치 검사(verify.py)는 표만 보므로 통과했다. 그래서 차트는 따로 본다.

    검사 여덟 -- 값·카테고리·논리·캐시·계열명·부호·빈계열·라벨겹침
    """
    out: list[dict[str, Any]] = []
    pairs = pair_up(deck, facts)
    for p in pairs:
        cps = deck.chart_parts(p.part)
        rec: dict[str, Any] = {"slide": p.slide_no,
                               "code": p.fact.get("question_code"),
                               "sub": p.fact.get("sub_title"),
                               "kind": chart_kind(p.fact.get("chart_type") or ""),
                               "findings": []}
        add = rec["findings"].append
        if not cps:
            add(("C0", "high", "차트가 없다"))
            out.append(rec)
            continue
        cp = cps[0]
        tree = deck.xml(cp)
        sers = D.chart_series(tree)
        vals = D.series_values(tree)
        names = D.series_names(tree)
        cats = D.chart_categories(tree)
        grid = p.fact["grid"]
        rows = body_rows(grid)
        rec["chart"] = cp
        rec["series"] = len(sers)

        # --- C4 차트 캐시 <-> 임베디드 워크시트 ---
        book = D.workbook_grid(deck, cp)
        if not book:
            add(("C4", "mid", "임베디드 워크시트를 못 읽었다 -- 더블클릭하면 옛 값으로 돌아간다"))

        # --- C9 더미 값: 제작팀은 껍데기를 「그럴듯한 숫자」로 채워 보낸다 ---
        #     실측: s23 차트가 전부 50.0, s36 이 다섯 계열 전부 ±20.0 이었다.
        #     빈칸이 아니라 값이라서 채움도 검사(FILLED)를 통과하고, 표 검사도
        #     표만 보므로 통과한다. **같은 절댓값만 나오면 데이터가 아니다.**
        flat = [abs(x) for v in vals for x in v if x is not None]
        if len(flat) >= 4 and len(set(round(x, 3) for x in flat)) == 1:
            add(("C9", "high", f"모든 값이 {flat[0]} 하나뿐이다 -- 제작팀 더미가 "
                               f"그대로 남았다"))

        if rec["kind"] == "scale":
            # --- C7 부호: 부정 계열은 음수여야 ±막대가 아래로 그려진다 ---
            neg = [n for n, v in zip(names, vals)
                   if any(x is not None and x < 0 for x in v)]
            pos = [n for n, v in zip(names, vals)
                   if any(x is not None and x > 0 for x in v)]
            if not neg:
                add(("C7", "high", "음수 계열이 없다 -- ±세로막대가 한쪽으로만 그려진다"))
            rec["neg_series"], rec["pos_series"] = neg, pos
            # --- C1 값: 척도는 배너가 카테고리다. 표 지표행(TOP2·평균)으로 검산 ---
            top = next((r for r in (grid.get("rows") or [])
                        if r.get("kind") == "top"), None)
            if top is not None and pos:
                want = _num(top["cells"][0])
                got = sum(v[0] for n, v in zip(names, vals)
                          if n in pos and v and v[0] is not None)
                if want is not None and abs(got - want) > 0.6:
                    add(("C1", "high", f"전체 열 TOP2 합이 안 맞는다: 차트 {got:.1f} vs 정본 {want:.1f}"))
            out.append(rec)
            continue

        axis = [r if r.get("kind") == "option" else None for r in rows]
        want_cats = [(r.get("label_clean") if r else "") for r in axis]

        # --- C2 카테고리 ---
        if len(cats) != len(want_cats):
            add(("C2", "high", f"카테고리 {len(cats)}개 != 정본 {len(want_cats)}개"))
        else:
            bad = [(i, a, b) for i, (a, b) in enumerate(zip(cats, want_cats))
                   if _norm(a) != _norm(b)]
            if bad:
                i, a, b = bad[0]
                add(("C2", "mid", f"카테고리 {len(bad)}개가 표 행 라벨과 다르다 "
                                  f"(예 {i}: 차트 {a!r} vs 정본 {b!r})"))

        # --- C1 값 ---
        if rec["kind"] == "rank":
            want = [[(_num(r["cells"][0]) if r else None) for r in axis],
                    [(_num(r.get("rank_first_total")) if r else None) for r in axis]]
            cum_name, first_name = rank_names(grid)
            want_names = [cum_name, first_name]
        else:
            want = [[(_num(r["cells"][0]) if r else None) for r in axis]]
            want_names = ["전체"]
        for si, wv in enumerate(want):
            if si >= len(vals):
                add(("C1", "high", f"계열 {si} 이 차트에 없다"))
                continue
            gv = vals[si]
            n = min(len(gv), len(wv))
            diffs = [(i, gv[i], wv[i]) for i in range(n)
                     if not _close(gv[i], wv[i])]
            if diffs:
                i, g, w = diffs[0]
                add(("C1", "high", f"계열 {si}({names[si]!r}) 값 {len(diffs)}곳이 정본과 다르다 "
                                   f"(예 {i}: 차트 {g} vs 정본 {w})"))
            if len(gv) != len(wv):
                add(("C1", "mid", f"계열 {si} 길이 {len(gv)} != 정본 {len(wv)}"))

        # --- C5 계열 이름 ---
        for si, wn in enumerate(want_names):
            if si < len(names) and names[si] and _norm(names[si]) != _norm(wn):
                add(("C5", "mid", f"계열 {si} 이름이 정본과 다르다: "
                                  f"{names[si]!r} vs {wn!r}"))

        # --- C8 남는 계열은 빈칸이어야 한다 ---
        for si in range(len(want), len(vals)):
            filled = [x for x in vals[si] if x is not None]
            if filled:
                add(("C8", "high", f"여분 계열 {si}({names[si]!r})에 값이 {len(filled)}개 "
                                   f"남아 있다 -- 범례에 없는 막대가 그려진다"))

        # --- C3 순위 논리: 누적 >= 1순위 ---
        if rec["kind"] == "rank" and len(vals) >= 2:
            bad = [(i, a, b) for i, (a, b) in enumerate(zip(vals[0], vals[1]))
                   if a is not None and b is not None and b > a + 0.05]
            if bad:
                i, a, b = bad[0]
                add(("C3", "high", f"1순위가 누적보다 큰 자리 {len(bad)}곳 "
                                   f"(예 {i}: 누적 {a} < 1순위 {b}) -- 계열이 뒤바뀌었다"))

        # --- C6 데이터 라벨 겹침 위험 ---
        #     ★ 이미 손으로 옮긴 자리는 세지 않는다. 제작팀은 겹치는 라벨만
        #       개별 dLbl 로 옮겨 두므로(s34 1개·s47 5개·s82 7개), 그것까지 세면
        #       「고쳐야 할 곳」과 「이미 고친 곳」이 뒤섞여 목록이 쓸모없어진다.
        if rec["kind"] == "rank" and len(vals) >= 2:
            span = max([x for x in vals[0] if x is not None] or [0]) or 1
            moved = _moved_label_idx(sers[1]) | _moved_label_idx(sers[0])
            clash = [i for i, (a, b) in enumerate(zip(vals[0], vals[1]))
                     if a is not None and b is not None
                     and (a - b) / span < LABEL_CLASH_RATIO
                     and i not in moved]
            rec["labels_moved"] = sorted(moved)
            if clash:
                add(("C6", "mid", f"두 계열 값이 붙어 라벨이 겹칠 자리 {len(clash)}곳 "
                                  f"(행 {clash[:6]}) -- 개별 데이터 라벨 이동이 필요하다"
                                  f"{f' / 이미 옮긴 자리 {len(moved)}곳' if moved else ''}"))
        out.append(rec)
    return out


def _moved_label_idx(ser: etree._Element) -> set[int]:
    """계열에서 **개별 데이터 라벨**이 붙은 데이터 포인트 인덱스.

    사람이 손으로 옮겨 놓은 자리다. `c:dLbls/c:dLbl/c:idx` 로 찾는다.
    """
    out: set[int] = set()
    dl = ser.find(q(C, "dLbls"))
    if dl is None:
        return out
    for one in dl.findall(q(C, "dLbl")):
        idx = one.find(q(C, "idx"))
        if idx is not None and idx.get("val"):
            out.add(int(idx.get("val")))
    return out


def _close(a: float | None, b: float | None, tol: float = 0.06) -> bool:
    """표시값끼리 견주므로 반올림 차이만 허용한다."""
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(float(a) - float(b)) <= tol


def render_chart_audit(recs: Sequence[dict[str, Any]]) -> str:
    lines = ["# 차트 전수 감사", ""]
    n_bad = sum(1 for r in recs if r["findings"])
    sev = {}
    for r in recs:
        for code, level, _ in r["findings"]:
            sev[f"{code}/{level}"] = sev.get(f"{code}/{level}", 0) + 1
    lines.append(f"- 차트 {len(recs)}개 · 지적 있는 장 {n_bad}개")
    lines.append(f"- 유형별: {sev or '없음'}")
    lines.append("")
    for r in recs:
        if not r["findings"]:
            continue
        lines.append(f"## s{r['slide']} {r['code']} — {r['sub']} [{r['kind']}]")
        for code, level, msg in r["findings"]:
            lines.append(f"- **{code}** ({level}) {msg}")
        lines.append("")
    return "\n".join(lines)


# =========================================================================== #
# 9. CLI
# =========================================================================== #

def refill(deck_path: Path, facts_path: Path, out_path: Path, *,
           fit_rows: bool = False, keep_ment: bool = True,
           keep_rows: set[int] | None = None,
           only: set[int] | None = None,
           text: bool = True, table: bool = True, chart: bool = True,
           drop_memos_too: bool = False,
           dry_run: bool = False) -> dict[str, Any]:
    deck = D.Deck(deck_path)
    facts = load_facts(facts_path)
    pairs = pair_up(deck, facts)

    rep: dict[str, Any] = {"deck": str(deck_path), "facts": str(facts_path),
                           "body": len(pairs), "mismatch": [], "slides": [],
                           "warnings": []}
    for p in pairs:
        if not p.ok:
            rep["mismatch"].append({
                "slide": p.slide_no, "code": p.fact.get("question_code"),
                "failed": [k for k, v in p.checks.items() if not v]})
    if rep["mismatch"]:
        return rep

    for p in pairs:
        if only and p.slide_no not in only:
            continue
        warns: list[str] = []
        row: dict[str, Any] = {"slide": p.slide_no,
                               "code": p.fact.get("question_code"),
                               "sub": p.fact.get("sub_title")}
        if text:
            row["text"] = refill_text(deck, p.part, p.fact,
                                      keep_ment=keep_ment, on_warn=warns.append)
        if table:
            # ★ 제작팀이 **의도적으로 보기를 줄인** 장표는 행 수를 건드리지 않는다
            #   (「항목 10개만 작업 · 각주 추가」). 정본 39행을 한 장에 넣을 수 없다.
            row["table"] = refill_table(
                deck, p.part, p.fact,
                fit_rows=fit_rows and p.slide_no not in (keep_rows or set()),
                on_warn=warns.append)
        if chart:
            row["chart"] = refill_chart(deck, p.part, p.fact, on_warn=warns.append)
        if drop_memos_too:
            row["memos"] = drop_memos(deck, p.part)
        if warns:
            row["warnings"] = warns
            rep["warnings"].extend(f"s{p.slide_no}: {w}" for w in warns)
        rep["slides"].append(row)

    if not dry_run:
        saved = B.save_checked(deck, out_path, com_verify=False)
        rep["out"] = str(saved)
    return rep


def _parse_only(text: str | None) -> set[int] | None:
    if not text:
        return None
    out: set[int] = set()
    for chunk in text.split(","):
        chunk = chunk.strip()
        if "-" in chunk:
            a, b = chunk.split("-", 1)
            out.update(range(int(a), int(b) + 1))
        elif chunk:
            out.add(int(chunk))
    return out


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="이미 만들어진 덱의 표·차트를 정본(slide_facts.json)으로 다시 채운다")
    ap.add_argument("deck", help="대상 .pptx")
    ap.add_argument("facts", help="정본 slide_facts.json")
    ap.add_argument("--out", help="산출 .pptx (없으면 대조만 한다)")
    ap.add_argument("--fit-rows", action="store_true",
                    help="표 행 수를 정본에 맞춘다 (빈 껍데기를 채울 때)")
    ap.add_argument("--keep-rows", default="",
                    help="행 수를 건드리지 않을 슬라이드 번호. 제작팀이 보기를 "
                         "의도적으로 줄인 장표에 쓴다. 예 92")
    ap.add_argument("--overwrite-ment", action="store_true",
                    help="멘트도 정본으로 덮어쓴다 (기본은 보존)")
    ap.add_argument("--only", help="슬라이드 번호 화이트리스트 (예 17,20,34-40)")
    ap.add_argument("--no-text", action="store_true")
    ap.add_argument("--no-table", action="store_true")
    ap.add_argument("--no-chart", action="store_true")
    ap.add_argument("--drop-memos", action="store_true",
                    help="제작팀 쪽지 도형을 함께 걷어낸다")
    ap.add_argument("--audit", action="store_true", help="쓰지 않고 대조만")
    ap.add_argument("--audit-charts", action="store_true",
                    help="차트를 정본과 전수 대조한다 (값·카테고리·논리·계열명·부호)")
    ap.add_argument("--json", help="리포트를 JSON 으로")
    ap.add_argument("--md", help="차트 감사 리포트를 마크다운으로")
    args = ap.parse_args(argv)

    if args.audit_charts:
        deck = D.Deck(Path(args.deck))
        facts = load_facts(Path(args.facts))
        recs = audit_charts(deck, facts)
        report = render_chart_audit(recs)
        print(report)
        if args.md:
            Path(args.md).write_text(report, encoding="utf-8")
            print(f"[리포트] {args.md}")
        if args.json:
            Path(args.json).write_text(
                json.dumps(recs, ensure_ascii=False, indent=2, default=str),
                encoding="utf-8")
        return 1 if any(l == "high" for r in recs for _, l, _ in r["findings"]) else 0

    out = Path(args.out) if args.out else Path(args.deck)
    rep = refill(Path(args.deck), Path(args.facts), out,
                 fit_rows=args.fit_rows,
                 keep_ment=not args.overwrite_ment,
                 keep_rows=_parse_only(args.keep_rows),
                 only=_parse_only(args.only),
                 text=not args.no_text, table=not args.no_table,
                 chart=not args.no_chart,
                 drop_memos_too=args.drop_memos,
                 dry_run=args.audit or not args.out)

    print(f"본문 {rep['body']}장 / 처리 {len(rep['slides'])}장 / "
          f"경고 {len(rep['warnings'])}건")
    if rep["mismatch"]:
        print(f"\n★ 정본과 어긋난 장 {len(rep['mismatch'])}개 -- 아무것도 쓰지 않았다")
        for m in rep["mismatch"][:20]:
            print(f"   s{m['slide']} {m['code']} -- 불일치: {', '.join(m['failed'])}")
        return 2
    skipped = [(r["slide"], (r.get("table") or {}).get("skipped"),
                (r.get("chart") or {}).get("skipped"))
               for r in rep["slides"]
               if (r.get("table") or {}).get("skipped")
               or (r.get("chart") or {}).get("skipped")]
    if skipped:
        print(f"\n건너뛴 자리 {len(skipped)}건")
        for no, t, c in skipped[:40]:
            print(f"   s{no}: 표={t or '-'} / 차트={c or '-'}")
    for w in rep["warnings"][:30]:
        print(f"   ! {w}")
    if rep.get("out"):
        print(f"\n산출: {rep['out']}")
    if args.json:
        Path(args.json).write_text(json.dumps(rep, ensure_ascii=False, indent=2),
                                   encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
