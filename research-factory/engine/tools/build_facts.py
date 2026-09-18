#!/usr/bin/env python3
"""통계표 + 설문지 → `slide_facts.json`.

파이프라인의 합류점이다. 여기서 **숫자에 대한 판단이 전부 끝난다.**
AI는 이 파일이 준 팩트를 문장으로 옮기기만 하고, 새 수치를 만들거나 계산하지 않는다.
그래서 `verify.py` 가 생성된 멘트의 모든 숫자를 통계표 원본과 1:1 대조할 수 있다.

하는 일
------
1. 통계표 90장 → 슬라이드 목록으로 접는다
   · SQ 문항 제외 · 응답자 특성은 조사개요로 · 순위 표 2장을 1장으로 병합
   · 요약표(TOP2/Mean Summary)는 본문에서 빼고 요약·결론용으로 분류
2. 슬라이드마다 결정 (전부 `engine/knowledge/guide/` 규칙 기반)
   · 차트타입      chart_rules.yaml decision_table
   · Base/Unit
   · Qu 텍스트     설문지 원문 그대로
   · 데이터 그리드  제작가이드 Data 영역에 그대로 붙일 형태
   · 멘트 팩트     1행(전체값 상위 N) · 2행(배너 튀는 값) · 수준어

사용법
------
    python build_facts.py <questionnaire.json> <tables.json> <out/slide_facts.json>
        [--kpi KPI요약표_직접생성.json] [--toc 목차.md]
"""

from __future__ import annotations

import re
import argparse
import io
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))
import rules as R  # noqa: E402


# ─────────────────────────────────────────────────────────────
# 1행에서 읽을 보기 개수  (knowledge/guide/13 §1-②)
# ─────────────────────────────────────────────────────────────

def n_options_to_read(n_options: int) -> int:
    """보기 수 → 멘트 1행에서 읽어줄 개수.

    "보기 개수가 한 10개가 되면 한 3개 읽고 보기 개수가 몇 개 없으면
     한두 개만 읽어도 돼요" (녹취 14:01)
    """
    if n_options <= 2:
        return 1
    if n_options <= 4:
        return 2
    return 3


def use_pattern_d(table: dict, opts: list[dict]) -> bool:
    """패턴 D(퍼센트 없이 '~가 있음')를 써야 하나.

    n이 1~2명이면 100%·50% 같은 수치가 오해를 부른다 (knowledge/guide/13 §2-D).
    """
    n = table.get("base_n") or 0
    if 0 < n <= 2:
        return True
    tb = R.total_banner(table)
    if tb is None or not opts:
        return False
    # 보기별 응답이 1건씩뿐 = 값이 전부 (1/n) 근처
    unit = 100.0 / n if n else 0
    vals = [R.cell_value((r.get("values") or [None] * (tb["pos"] + 1))[tb["pos"]]) for r in opts]
    vals = [v for v in vals if v]
    return bool(vals) and unit > 0 and all(abs(v - unit) < 0.05 for v in vals)


# ─────────────────────────────────────────────────────────────
# 슬라이드 접기
# ─────────────────────────────────────────────────────────────

def is_sq(table: dict) -> bool:
    return (table.get("question_code") or "").upper().startswith("SQ")


def is_profile(table: dict) -> bool:
    return table.get("index") == 1 or "응답자 특성" in (table.get("title_full") or "")


def fold_slides(tables: list[dict],
                exclude: dict[int, str] | None = None) -> tuple[list[dict], dict[str, list[dict]]]:
    """통계표 목록 → (본문 슬라이드 단위, 본문 밖으로 뺀 것들).

    슬라이드 단위 하나는 표 1장 또는 순위 표 2장이다.
    `exclude` = {표 인덱스: 사유} — R-023 예외④(인지 여부 등 무의미한 표). **사람이 목차에서
    확정한 것만** 들어온다. 룰엔진이 스스로 판단하지 않는다.
    """
    aside: dict[str, list[dict]] = {"sq": [], "profile": [], "summary": [], "excluded": []}
    body: list[dict] = []
    exclude = exclude or {}

    # 순위 짝을 먼저 잡아 두 표를 한 단위로 묶는다.
    pairs = R.pair_ranking_tables(tables)
    paired_ids = {id(t) for pair in pairs for t in pair if t is not None}
    pair_by_combined = {id(combined): first for first, combined in pairs}

    for t in tables:
        if t.get("index") in exclude:
            aside["excluded"].append({**t, "_reason": exclude[t["index"]]})
            continue
        if is_sq(t):
            aside["sq"].append(t)
            continue
        if is_profile(t):
            aside["profile"].append(t)
            continue
        if (t.get("hints") or {}).get("is_summary"):
            # 요약표는 본문에 넣지 않는다 — 요약·결론에서만 쓴다 (녹취 50:24)
            aside["summary"].append(t)
            continue
        if id(t) in paired_ids:
            if id(t) in pair_by_combined:
                first = pair_by_combined[id(t)]
                body.append({"main": t, "rank_first": first})
            # 1순위 표 단독은 건너뛴다 — 합산표 쪽에 병합됐다
            continue
        body.append({"main": t, "rank_first": None})

    return body, aside


# ─────────────────────────────────────────────────────────────
# 데이터 그리드 (제작가이드 Data 영역에 그대로 붙일 형태)
# ─────────────────────────────────────────────────────────────

#: 셀 서식에서 소수 자릿수를 읽는다 — `###0.0` → 1 · `###0.00` → 2 · `####.00` → 2.
_FMT_DECIMALS = re.compile(r"\.(0+)")


def _cell_digits(cell: Any, default: int) -> int:
    """통계표 셀 서식이 정한 소수 자릿수. 서식이 없으면 `default`."""
    if isinstance(cell, dict):
        m = _FMT_DECIMALS.search(cell.get("fmt") or "")
        if m:
            return len(m.group(1))
    return default


def build_grid(table: dict, rank_first: dict | None = None, extra_groups: "set[str] | None" = None,
               resort: bool = False, freq_banners: "set[str] | None" = None,
               small_base_mark: bool = True,
               show_digits: int | None = None) -> dict:
    """배너 헤더 + Base for % + 데이터 행 + 요약 행을 통째로 만든다.

    `그래프 상단 Text` 는 **행 라벨 열의 왼쪽 칸**에 놓는다. 이 문자열을 적어야
    제작팀이 Top·평균 값을 그래프 위에 표시한다 — 안 적으면 안 나온다 (녹취 33:52).
    `extra_groups`: 본문 절단선 밖(in_report=False)이어도 살릴 배너군 — 요약 패널이 `배너 표 (건물 연면적)` 처럼
    원표 교차를 지정할 때 쓴다(R-070, 실측 2026-08-23: 요약 1 연면적 4열이 빠져 나감).
    """
    banners = R.report_banners(table)
    if extra_groups:
        have = {id(b) for b in banners}
        extra = [b for b in (table.get("banners") or []) if b.get("group") in extra_groups and id(b) not in have]
        banners = sorted(banners + extra, key=lambda b: b["pos"])
    tb = R.total_banner(table)

    group_row: list[str] = [""]
    label_row: list[str] = [""]
    base_row: list[str] = ["Base for %"]
    prev_group = None
    for b in banners:
        group_row.append(b["group"] if b["group"] != prev_group else "")
        prev_group = b["group"]
        label_row.append(b["label"] if b["label"] != b["group"] else "")
        # n<30 인 배너 셀에는 별표를 단다 — `(13*)`. **제작가이드에 적는다** (사용자 지시 2026-08-26).
        #
        # 하우스 제작가이드 4벌(LG·SKEE·SH청안주·한옥) base 셀 2,886개에는 별표가 0건이고
        # 통계표 원본도 `(9)` 로 준다. 그러나 그것은 **적지 않아도 된다는 근거가 아니라
        # 적지 않아서 뒤로 밀린 일**이다 — SKEE 는 제작가이드 89개 블록이 전부 n<30 인데
        # 표기가 없었고, 최종본에 가서야 `Too small base(n<30)` 가 붙었다.
        # 저베이스 경고는 보고서에 반드시 실려야 하므로, 뒤에서 누가 챙기기를 기대하지 않고
        # 가이드에 적는다. 별표를 달면 build_guide 가 Base/Unit 뒤에 범례도 함께 붙인다.
        n_b = b.get("n") or 0
        if b.get("n_label"):
            # 열 하나에 표본이 하나가 아닌 파생 열 — 「해당 업무 위탁 중」처럼 행(업무)마다 n 이
            # 다르다. 한 숫자로 적으면 거짓이 되므로 표가 준 표기(「(117~273)」)를 그대로 싣는다
            # (에스원 비상주 표9007, 2026-09-06). 별표 판정은 최소 n 으로 한다.
            base_row.append(str(b["n_label"]))
        else:
            base_row.append(f"({n_b}*)" if (small_base_mark and 0 < n_b < R.SMALL_BASE) else f"({n_b})")

    mean_table = (table.get("hints") or {}).get("summary_kind") == "mean"   # Mean Summary 표는 행 kind 가 option 이어도 평균값 (V17 실측 2026-08-24: 4.08→4.1 로 깎였음)

    #: 행마다 셀 표시 서식을 함께 모은다. build_guide 가 이 값으로 `###0.0` 을 건다.
    row_formats: dict[int, list[str | None]] = {}

    force_digits = show_digits

    def row_cells(row: dict, _fmt_sink: list | None = None) -> list[str]:
        vals = row.get("values") or []
        default_digits = 2 if (row.get("kind") == "mean" or mean_table) else 1
        out: list[str] = []
        fmts: list[str | None] = _fmt_sink if _fmt_sink is not None else []
        for b in banners:
            cell = vals[b["pos"]] if b["pos"] < len(vals) else None
            v = R.cell_value(cell)
            # 사람이 지정한 열은 **백분율 대신 빈도**를 적는다. n 이 아주 작은 열에서
            # 백분율은 0.0/100.0 으로만 나오거나(응답 1건) 100 을 넘어(중복 슬롯) 읽을 수 없고,
            # 그 값 자체가 개별 응답의 노출이다 — 통계법 조항에 걸린다.
            # 평균 행은 백분율이 아니므로 그대로 둔다.
            if freq_banners and b["label"] in freq_banners and row.get("kind") != "mean":
                n_b = b.get("n") or 0
                out.append("-" if v is None else str(int(R.round_half_up(v * n_b / 100.0, 0))))
                fmts.append(None if v is None else "###0")      # 빈도는 정수다
                continue
            # 자릿수의 정본은 **통계표 셀 서식**이다. `% 는 1자리` 는 통례이지
            # 불변식이 아니다 — 에스원 TA Pool 표31(공간 구성 비율)은 보기 행인데
            # 서식이 `###0.00` 이라 통계표 표시값이 `86.13` 이다. 1자리로 깎으면
            # 검증기가 「통계표에 없는 값」으로 잡는다(V2 치명 36건).
            # 값은 **통계표 원값 그대로** 싣고 자릿수는 표시 서식이 지킨다
            # (사용자 지시 2026-08-26). 통계표가 그렇게 되어 있다 — 셀에는
            # 25.7142857142857 이 들어 있고 서식이 25.7 로 보여 준다.
            # 반올림해서 담으면 제작팀이 그 값을 다시 반올림할 때 자릿수가 한 번 더 깎인다.
            digits = _cell_digits(cell, default_digits)
            # 파서는 `raw`(원값)와 `value`(표시값)를 둘 다 담는다. 제작가이드에는 **원값**을 싣는다 —
            # 통계표 셀이 그렇게 되어 있고(25.7142857142857 + 서식 `###0.0`),
            # 표시값을 담으면 자릿수가 한 번 더 깎인 채 굳는다.
            rawv = cell.get("raw") if isinstance(cell, dict) else None
            out.append("-" if v is None else repr(float(rawv if rawv is not None else v)))
            # 화면 자릿수. `--show-digits` 는 **백분율 행에만** 건다 —
            # 평균(개월·층·개)은 통계표가 2자리로 주고 그 자릿수가 뜻을 가진다.
            is_mean = row.get("kind") == "mean" or mean_table
            shown = digits if (force_digits is None or is_mean) else force_digits
            fmts.append(None if v is None else ("###0." + "0" * shown if shown else "###0"))
        return out

    sorted_opts, sort_why = R.sort_option_rows(table, resort=resort)

    data_rows: list[dict] = []

    # 순위 슬라이드는 1순위 표의 **전체 열만** 가져와 앞에 붙인다 (녹취 29:04).
    rank_lookup: dict[str, float | None] = {}
    if rank_first is not None:
        rtb = R.total_banner(rank_first)
        for r in rank_first.get("rows") or []:
            if r.get("kind") == "option" and rtb is not None:
                vals = r.get("values") or []
                rank_lookup[(r.get("label") or "").strip()] = (
                    R.cell_value(vals[rtb["pos"]]) if rtb["pos"] < len(vals) else None)

    for r in sorted_opts:
        fmt_sink: list[str | None] = []
        cells = row_cells(r, fmt_sink)
        # 그리드에는 `label_raw`(원문 간격 보존)를 쓴다. 보기 행은 앞 공백 2칸,
        # `계` 행은 6칸이 규격이고(rules.yaml V-010/V-011), 정규화된 `label` 을 쓰면
        # 그 간격이 사라져 제작팀 RPA 가 행을 못 알아본다.
        entry: dict[str, Any] = {
            "label": r.get("label_raw") or r.get("label"),
            "label_clean": r.get("label"),
            "kind": "option", "cells": cells, "cell_formats": fmt_sink,
        }
        if rank_first is not None:
            v = rank_lookup.get((r.get("label") or "").strip())  # 조인은 정규화 라벨로
            entry["rank_first_total"] = R.fmt_pct(v) if v is not None else "-"
        data_rows.append(entry)

    for kind in ("total", "bot", "soso", "top", "mean"):
        for r in R.rows_of_kind(table, kind):
            fs: list[str | None] = []
            data_rows.append({
                "label": r.get("label_raw") or r.get("label"),
                "label_clean": r.get("label"),
                "kind": kind,
                "cells": row_cells(r, fs), "cell_formats": fs,
                # Top·평균 행에만 제작팀 신호어를 붙인다
                "modifier": "그래프 상단 Text" if kind in ("top", "mean") else None,
            })

    return {
        "banner_group_row": group_row,
        "banner_label_row": label_row,
        "base_row": base_row,
        "rows": data_rows,
        "sort": sort_why,
        # 백분율 대신 빈도를 적은 열 — 검증기가 이 열을 통계표 표시값과 곧이곧대로
        # 대조하면 전량 「통계표에 없는 값」으로 잡힌다. 무엇을 어떻게 바꿨는지 남긴다.
        "freq_banners": sorted(freq_banners & {b["label"] for b in banners}) if freq_banners else [],
        "n_banners": len(banners),
        "rank_overlay": rank_first is not None,
        "rank_columns": ["1순위", table.get("suffix", "").lstrip("_") or "1+2+3순위"]
        if rank_first is not None else None,
    }


# ─────────────────────────────────────────────────────────────
# 멘트 팩트
# ─────────────────────────────────────────────────────────────

def build_ment_facts(table: dict, resort: bool = False) -> dict:
    """멘트를 쓰는 데 필요한 숫자를 전부 확정해 둔다. AI는 여기 있는 값만 쓴다."""
    tb = R.total_banner(table)
    opts, sort_why = R.sort_option_rows(table, resort=resort)
    h = table.get("hints") or {}

    facts: dict[str, Any] = {
        "sort": sort_why,
        "pattern": None,
        "line1": [],
        "line2": [],
        "level": None,
        "n_options": len(opts),
    }

    # ── 척도: 【TOP2】 값으로 읽는다 (보기값이 아니라)
    if h.get("has_scale_rows"):
        top = R.first_row_of_kind(table, "top")
        mean = R.first_row_of_kind(table, "mean")
        if top is not None and tb is not None:
            tv = R.cell_value((top.get("values") or [])[tb["pos"]])
            word, phrase = R.level_word(tv)
            facts["pattern"] = "A"
            facts["level"] = {
                "top_label": top.get("label"),
                "top_pct": R.round_half_up(tv, 1) if tv is not None else None,
                "top_display": R.fmt_pct(tv),
                "word": word,
                "phrase": phrase,
            }
            if mean is not None:
                mv = R.cell_value((mean.get("values") or [])[tb["pos"]])
                facts["level"]["mean_label"] = mean.get("label")
                facts["level"]["mean_display"] = R.fmt_mean(mv)
            devs = R.row_deviations(table, top)
            facts["line2"] = [vars(d) for d in R.pick_deviations(devs)]
            facts["line2_candidates"] = len(devs)
        return facts

    # ── 그 외: 전체값 상위 N 개를 읽는다
    if use_pattern_d(table, opts):
        facts["pattern"] = "D"
        take = min(2, len(opts))
    else:
        facts["pattern"] = "C"
        take = min(n_options_to_read(len(opts)), len(opts))

    if tb is not None:
        for r in R.top_options_by_value(table, take):
            cell = (r.get("values") or [None] * (tb["pos"] + 1))[tb["pos"]]
            v = R.cell_value(cell)
            # 표시 자릿수는 통계표 셀 서식을 따른다 — 그리드와 같은 규칙이어야
            # 멘트 수치가 데이터 행과 대조된다(V3). `_cell_digits` 주석 참조.
            digits = _cell_digits(cell, 1)
            facts["line1"].append({
                "label": r.get("label"),
                "pct": R.round_half_up(v, digits) if v is not None else None,
                "display": R.fmt_mean(v) if digits >= 2 else R.fmt_pct(v),
            })
        # 동률 묶기 — 같은 값이면 한 괄호에 몰아 쓴다 (knowledge/guide/13 §2-C).
        # ★ 동률은 **전체 보기**에서 찾는다. line1 안에서만 보면 상위 N 절단선 밖에 있는
        #   같은 값을 놓쳐, 멘트가 동률 중 한쪽만 인용해 나머지를 가린다 —
        #   에스원 TA Pool 실측 4곳(표10·12 의 1년/2년 단위 22.9 · 표27 의 15층 미만/25층 이상 25.6 ·
        #   표29 의 1개 층/3개 층 이상 29.1). 표27 은 가려진 쪽이 「25층 이상」이라
        #   고층 비중이 문장에서 통째로 사라졌다 (guide-adversary F-10).
        line1_displays = {item["display"] for item in facts["line1"]}
        ties: dict[str, list[str]] = {}
        for r in opts:
            cell = (r.get("values") or [None] * (tb["pos"] + 1))[tb["pos"]]
            v = R.cell_value(cell)
            if v is None:
                continue
            digits = _cell_digits(cell, 1)
            disp = R.fmt_mean(v) if digits >= 2 else R.fmt_pct(v)
            if disp in line1_displays:
                ties.setdefault(disp, []).append(r.get("label") or "")
        facts["ties"] = {k: v for k, v in ties.items() if len(v) > 1}

        if opts:
            devs = R.row_deviations(table, opts[0])
            facts["line2"] = [vars(d) for d in R.pick_deviations(devs)]
            facts["line2_candidates"] = len(devs)

    # ── R-139 수량형 보기 단위 (사용자 지시 2026-08-23: 「'5' 보다 5층」) — 보기가 전부 순수 숫자이고
    #    평균 행 라벨이 「[평균 : 층]」 꼴이면 그 단위를 멘트용으로 확정한다. 담당은 '5' 대신 5층 이라 쓴다(따옴표 없이).
    labels = [str(r.get("label") or "").strip() for r in opts]
    if labels and all(re.fullmatch(r"\d+", l) for l in labels):
        mean = R.first_row_of_kind(table, "mean")
        m = re.search(r"\[\s*평균\s*[:：]\s*([^\]\s]+)\s*\]", str((mean or {}).get("label") or ""))
        if m:
            facts["unit"] = m.group(1)
            for item in facts["line1"]:
                if item.get("label") and re.fullmatch(r"\d+", str(item["label"]).strip()):
                    item["label_unit"] = f"{str(item['label']).strip()}{m.group(1)}"
            for d in facts["line2"]:
                if d.get("row_label") and re.fullmatch(r"\d+", str(d["row_label"]).strip()):
                    d["row_label_unit"] = f"{str(d['row_label']).strip()}{m.group(1)}"

    return facts


# ─────────────────────────────────────────────────────────────
# 조립
# ─────────────────────────────────────────────────────────────

#: 문항문 말미의 응답 형식 지시. 설문지에는 프로그래머용으로 붙어 있지만
#: 보고서 `Qu:` 는 **응답자가 읽은 질문**만 담는다 (예시 제작본 LG 79장 전부 마커 없음).
RESPONSE_MARKER = re.compile(
    r"\s*\[(?:\d+개\s*선택|모두\s*선택|직접\s*입력|\d+행\s*필수\s*입력"
    r"|행별\s*\d+개\s*선택|항목별\s*\d+개\s*선택"
    r"|최소\s*\d+\s*개\s*,?\s*최대\s*\d+\s*개\s*선택|최대\s*\d+\s*개\s*선택"
    r"|\d+순위[^\]]*|1순위[^\]]*)\]\s*")

#: 접미 중 순위·요약표를 가리키는 조각. 나머지는 매트릭스 항목으로 본다.
RANK_PART = re.compile(r"^(?:\d+(?:[~+]\d+)*순위.*|TOP\d+\s*Summary|Mean\s*Summary)$", re.I)

#: 매트릭스 항목 접미 앞에 붙는 보기번호 (`1) 지상` → `지상`)
ITEM_NO = re.compile(r"^\s*\d+\)\s*")


PIPING = re.compile(r"\[#[^\]]*?파이핑[^\]]*?#\]")

#: 받침 유무에 따라 갈리는 조사 쌍 (받침 있음, 받침 없음)
PARTICLES = [("을", "를"), ("은", "는"), ("이", "가"), ("과", "와"), ("으로", "로")]
PARTICLE_RE = re.compile(r"([가-힣])(을|를|은|는|이|가|과|와|으로|로)(?=\s|$|[,.\]])")


def has_final(ch: str) -> bool:
    """한글 음절에 받침이 있나."""
    code = ord(ch) - 0xAC00
    return 0 <= code <= 11171 and code % 28 != 0


def fix_particles(text: str) -> str:
    """치환 뒤 조사를 받침에 맞춘다.

    설문지는 파이핑 자리에 어떤 값이 올지 모르니 조사를 임의로 붙여 뒀다.
    치환어를 넣으면 `업무을` 처럼 어긋나므로 바로잡는다. (`로`는 ㄹ 받침 예외)
    """
    def fix(m: re.Match) -> str:
        ch, part = m.group(1), m.group(2)
        final = has_final(ch)
        for a, b in PARTICLES:
            if part in (a, b):
                if (a, b) == ("으로", "로"):
                    # ㄹ 받침 뒤에는 `로` 를 쓴다
                    want = a if (final and (ord(ch) - 0xAC00) % 28 != 8) else b
                else:
                    want = a if final else b
                return ch + want
        return m.group()
    return PARTICLE_RE.sub(fix, text)


def resolve_piping(text: str, piping_map: dict) -> tuple[str, bool]:
    """파이핑 자리표시자를 독자가 읽을 말로 바꾼다.

    응답자는 실제 값을 봤으므로 보고서 `Qu:` 에 자리표시자가 남으면 안 된다 —
    예시 제작본 150장에 파이핑 마커는 **단 한 건도 없다**.
    치환어는 프로젝트별 `파이핑치환.json` 에서 온다 (참조 문항의 주제에서 뽑는다).
    """
    if not PIPING.search(text):
        return text, False
    table = (piping_map or {}).get("치환") or {}
    unresolved = False

    def sub(m: re.Match) -> str:
        nonlocal unresolved
        entry = table.get(m.group())
        if entry:
            return entry["치환어"] if isinstance(entry, dict) else str(entry)
        unresolved = True
        return m.group()

    out = PIPING.sub(sub, text)
    # 설문지는 자리표시자 뒤에 조사를 띄어 쓴 곳이 있다(`[# LIST 파이핑 #] 이 맞습니까?`).
    # 값이 들어가면 붙여 써야 자연스럽다.
    out = re.sub(r"([가-힣])\s+(을|를|은|는|이|가|과|와|으로|로)(?=\s|$|[,.])",
                 lambda m: m.group(1) + m.group(2), out)
    return re.sub(r"\s{2,}", " ", fix_particles(out)).strip(), unresolved


def compose_qu(q: dict, table: dict, piping_map: dict | None = None, *, suffix_on: bool = True) -> str:
    """`Qu:` 본문을 만든다.

    두 가지를 지킨다 (예시 제작본 실측):
    - 응답 형식 마커(`[직접 입력]` 등)는 **뺀다.** 응답자가 읽은 문장만 남긴다.
    - 접미 표기가 종류마다 다르다 —
      순위/요약표는 밑줄(`…_1+2+3순위 (중복응답)`),
      매트릭스 항목은 파이프(`… | 편의/시간 지연`)이고 보기번호는 떼어낸다.
    """
    text = RESPONSE_MARKER.sub(" ", (q.get("text") or "")).strip()
    text = re.sub(r"\s{2,}", " ", text)
    # `※ …` 부기는 응답자에게 준 기입 안내다. 제작팀은 그것으로 그릴 것이 없으므로 뗀다
    # (사용자 지시 2026-08-26 — 「딱 질문만」).
    text = re.split(r"\s*※", text)[0].strip()
    text, _ = resolve_piping(text, piping_map or {})

    suffix = (table.get("suffix") or "").strip().lstrip("_")
    if not suffix or not suffix_on:
        return text

    items, ranks = [], []
    for part in (p for p in suffix.split("_") if p.strip()):
        (ranks if RANK_PART.match(part.strip()) else items).append(ITEM_NO.sub("", part).strip())

    # 중복 방지는 **조립된 형태 그대로** 찾는다. 맨 문자열로 찾으면
    # `지상` 이 "지상 층수…" 안에 이미 있다는 이유로 접미가 통째로 사라져
    # 서로 다른 슬라이드의 Qu 가 같아진다 (A1 의 지상/지하/엘리베이터가 그랬다).
    out = text
    if items:
        joined = "_".join(items)
        if f"| {joined}" not in out:
            out = f"{out} | {joined}" if out else joined
    for r in ranks:
        if f"_{r}" not in out:
            out = f"{out}_{r}" if out else r
    return out


#: 마크다운 표의 행을 셀로 쪼갠다. `\|` (이스케이프된 파이프)는 열 구분자가 아니다.
_MD_PIPE = re.compile(r"(?<!\\)\|")


def _split_md_row(line: str) -> list[str]:
    body = line.strip()
    body = _MD_PIPE.sub("\x00", body).strip("\x00")
    return [c.replace("\\|", "|").strip() for c in body.split("\x00")]


def _toc_groups(path: Path) -> dict[str, str]:
    """첫 펜스 블록 Part Ⅲ 의 대분류 줄(`A. 건물 관리 현황`) 아래 챕터들을 그 대분류에 매핑한다 (R-140)."""
    text = io.open(path, encoding="utf-8").read()
    m = re.search(r"```\n(.*?)```", text, re.S)
    if not m:
        return {}
    out: dict[str, str] = {}
    cur_part = None
    grp = None
    for ln in m.group(1).splitlines():
        mp = re.match(r"^Part\s+([ⅠⅡⅢⅣ])\.", ln.strip())
        if mp:
            cur_part = mp.group(1); grp = None; continue
        if cur_part != "Ⅲ" or not ln.strip():
            continue
        mg = re.match(r"^([A-Z])\.\s+(\S.*?)\s*$", ln.strip())
        if mg:
            grp = f"{mg.group(1)}. {mg.group(2).strip()}"; continue
        if grp:
            for tok in re.split(r"\s{2,}", ln.strip()):
                mi = re.match(r"^(\d+)\.\s*(.+?)\s*$", tok.strip())
                if mi:
                    out[f"{mi.group(1)}. {mi.group(2)}"] = grp
    return out


def parse_toc(path: Path) -> tuple[dict[int, dict], list[str]]:
    r"""`목차.md` 의 문항 배정표를 읽어 seq → (슬라이드 타이틀, 서브 타이틀) 로 만든다.

    목차는 사람이 확정한다. 여기서는 그 결정을 **읽기만** 하고 지어내지 않는다 —
    빠진 seq 가 있으면 채우지 말고 경고로 올린다(§2-3① 전수성은 사람이 책임진다).

    기대 형식:
        | P | 꼭지 | seq | 문항 | 서브 타이틀 | 표 | 구성 및 BASE |
        | 1 | 1. 현황 | 6 | A3 | 1) 소유 \| 형태 | 12 | 전체 500 |

    서브 타이틀은 3단 구조를 `중분류 \| 세부` 로 한 칸에 담는다 (LG 예시 실측 표기).
    마크다운 표에서 `|` 는 열 구분자이므로 `\|` 로 이스케이프되어 있다 — 여기서 되돌린다.
    """
    rows: dict[int, dict] = {}
    warns: list[str] = []
    header_cols: list[str] | None = None
    group_of = _toc_groups(path)                      # R-140 대분류: 챕터 → `A. 건물 관리 현황`

    for line in io.open(path, encoding="utf-8"):
        if not line.lstrip().startswith("|"):
            continue
        cells = _split_md_row(line)
        if header_cols is None:
            if "seq" in cells and "꼭지" in cells and "서브 타이틀" in cells:
                header_cols = cells
            continue
        if set("".join(cells)) <= set("-: "):      # 구분선
            continue
        if len(cells) != len(header_cols):
            continue
        rec = dict(zip(header_cols, cells))
        raw = rec.get("seq", "")
        if not raw.isdigit():
            continue
        seq = int(raw)
        if seq in rows:
            warns.append(f"seq {seq} 가 배정표에 두 번 나옵니다 (MECE 위반)")
            continue
        rows[seq] = {
            "page": int(rec["P"]) if rec.get("P", "").isdigit() else None,
            "slide_title": rec.get("꼭지") or None,
            "group": group_of.get(re.sub(r"\s+", " ", (rec.get("꼭지") or "").strip())),
            "sub_title": rec.get("서브 타이틀") or None,
            "base_note_toc": rec.get("구성 및 BASE") or None,
            "question_code": rec.get("문항") or None,
        }

    if header_cols is None:
        warns.append("배정표를 찾지 못했습니다 (헤더에 `seq`·`꼭지`·`서브 타이틀` 필요)")
    return rows, warns


#: `구성 및 BASE` 칸의 선두 조건 토큰 — `A3=1` · `B2=3,4` · `C1≥1` · `D7=1~3` · `C4=1~7`
#: `≠` 와 `A2-a` 꼴(하위 항목이 알파벳)도 받는다 — 에스원 TA Pool 의 실제 조건이
#: `A2-a ≠ 1)` · `A4 -a ≠ 1)` · `B3≠0` 이고, 좁게 잡으면 조건식이 통째로 군 이름이 된다.
_BASE_COND = re.compile(r"^\s*[A-Z]{1,2}\d*(?:\s*-\s*(?:\d+|[A-Za-z]))?\s*[=≥≤<>≠]+\s*[\d,~\.]+\)?\s*")
#: 끝 사례수 — `300` 도 `(n=85)` 도 벗긴다. 비상주 목차가 `A3=1 기업소유 (n=85)` 로 적어
#: 라벨이 `기업소유 (n=85)` 로 오염됐다(base_unit 이 `(n=85)(n=85)` 가 된다).
_TRAIL_N = re.compile(r"\s*(?:[\(（]\s*n\s*=\s*[\d,]+\s*[\)）]|[\d,]+)\s*$", re.I)


def derive_base_label(base_note_toc: str | None, q: dict | None, base_n: int | None,
                      total_n: int | None) -> str:
    """서브베이스 장의 `base_label` (군 이름). base == 전체면 "전체".

    ① 목차 `구성 및 BASE` 에서 선두 조건 토큰과 끝 사례수를 떼면 남는 것이 군 이름이다
       (`B2=3,4 외부 위탁 300` → `외부 위탁`). 사람이 확정한 문구라 1순위.
    ② 없으면 설문지 base_condition 의 `|` 뒤 주해 (`B2=3,4 | 외부 위탁` → `외부 위탁`).
    ③ 둘 다 없으면 조건식 자체라도 남긴다 — `전체` 로 위장하지 않는다.
    """
    if base_n is not None and total_n is not None and base_n >= total_n:
        return "전체"
    txt = (base_note_toc or "").strip()
    if txt:
        label = _TRAIL_N.sub("", _BASE_COND.sub("", txt)).strip(" ·-—")
        if label and label != "전체":
            return label
    cond = (q or {}).get("base_condition") or ""
    if "|" in cond:
        note = cond.split("|", 1)[1].strip()
        if note:
            return note
    return cond.strip() or "서브베이스"


def build_slide(unit: dict, questions: dict[str, dict], chart_rules: dict,
                seq: int, piping_map: dict | None = None,
                resort_tables: "set[int] | None" = None,
                freq_banners: "set[str] | None" = None,
                small_base_mark: bool = True,
                qu_suffix: bool = True,
                show_digits: int | None = None) -> dict:
    table = unit["main"]
    # DP 정렬이 깨진 표만 사람이 지정해 재정렬한다 (R.sort_option_rows 의 resort 예외).
    resort = bool(resort_tables) and table.get("index") in (resort_tables or ())
    rank_first = unit.get("rank_first")
    code = table.get("question_code")
    q = questions.get(code) or {}

    decision = R.decide_chart_type(table, q, chart_rules)
    # 소표본 여부는 **언제나** 판정한다 — 사람이 보는 `flags` 신호이기 때문이다.
    # `Base/Unit` 칸에 실을 `base_note` 는 기본으로 만든다 (`--no-small-base-mark` 로만 끈다).
    small_base = R.small_base_warning(table)
    base_note = small_base if small_base_mark else None

    flags: list[str] = []
    if decision.unresolved:
        flags.append("차트타입_확인필요")
    if small_base:
        flags.append("소표본")
    # 배너 열 상한은 고정값이 아니다. 예시 제작본 실측에서 본문 표가 8~24열로
    # 제각각이었고(표 전체 폭 32cm를 유지한 채 열 너비를 슬라이드마다 다시 잡는다),
    # 플랜의 "1.5cm × 14열"은 상한이 아니라 흔한 한 경우였다.
    # 24열을 넘어가면 실측 범위 밖이므로 그때만 사람이 본다.
    if (table.get("hints") or {}).get("n_banners_in_report", 0) > 24:
        flags.append("배너_실측범위초과")

    qu_text = compose_qu(q, table, piping_map, suffix_on=qu_suffix)

    slide = {
        "seq": seq,
        "question_code": code,
        "source_tables": [t["index"] for t in (table, rank_first) if t],
        "title_full": table.get("title_full"),
        # 목차 단계에서 사람이 확정한다. 비워 두는 게 맞다 — 지어내지 않는다.
        "slide_title": None,
        "sub_title": None,
        "chart_type": decision.chart_type,
        "chart_rule": {
            "id": decision.rule_id,
            "confidence": decision.confidence,
            "reason": decision.reason,
        },
        "modifiers": decision.modifiers,
        "base_unit": R.base_unit_string(table),
        # 이 표를 **누가 봤는가** — 앞단 문항의 어느 보기를 고른 사람에게만 물은 것인지.
        # 설문 파서가 뽑아 둔 `base_condition`(`D7=4~5 | 매력도 높음`)을 그대로 옮긴다.
        # 모수를 모르면 해석이 통째로 어긋난다 (사용자 지시 2026-08-24).
        "base_condition": (q or {}).get("base_condition") or None,
        "base_note": base_note,
        "base_n": table.get("base_n"),
        "qu": f"Qu: {qu_text}" if qu_text else None,
        "grid": build_grid(table, rank_first, resort=resort, freq_banners=freq_banners,
                           small_base_mark=small_base_mark, show_digits=show_digits),
        "ment_facts": build_ment_facts(table, resort=resort),
        "flags": flags,
        "_table": table,          # 목차 적용 단계에서만 쓰는 임시 참조 — 직렬화 전에 뺀다
    }
    apply_label_unit(slide)
    return slide


def apply_label_unit(slide: dict) -> None:
    """R-139 (가): 수량형 보기의 단위를 **표 행 라벨에도** 붙인다 — 멘트(5층)와 표(5)가 한 장에서 어긋나지 않게
    (PPT 팀 반송 R-PPT-007, 2026-08-23). `label` 은 원문 그대로, `label_clean`·`label_unit` 이 단위 병기형."""
    unit = (slide.get("ment_facts") or {}).get("unit")
    if not unit:
        return
    for r in (slide.get("grid") or {}).get("rows") or []:
        lab = str(r.get("label") or "").strip()
        if r.get("kind") == "option" and re.fullmatch(r"\d+", lab):
            r["label_unit"] = f"{lab}{unit}"
            r["label_clean"] = f"{lab}{unit}"


def main() -> int:
    ap = argparse.ArgumentParser(description="통계표+설문지 → slide_facts.json")
    ap.add_argument("questionnaire", type=Path)
    ap.add_argument("tables", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument("--kpi", type=Path, help="make_kpi_summary.py 가 만든 요약표 json")
    ap.add_argument("--piping", type=Path,
                    help="파이핑 자리표시자 치환 맵 (프로젝트별 파이핑치환.json)")
    ap.add_argument("--toc", type=Path,
                    help="사람이 확정한 목차.md — 슬라이드/서브 타이틀을 여기서 읽는다")
    ap.add_argument("--ments", type=Path,
                    help="확정된 멘트 json ({\"ments\":[{\"seq\":1,\"lines\":[...]}]})")
    ap.add_argument("--freq-banners", default="",
                    help="백분율 대신 **빈도**를 적을 배너 열 라벨 (세미콜론 구분). 예: \"서울시 강서구\". "
                         "n 이 아주 작아 백분율이 0.0/100.0 으로만 나오거나 100 을 넘는 열에 쓴다 — "
                         "그 값이 곧 개별 응답이라 통계법 조항에 걸린다. 각주는 반드시 함께 단다")
    ap.add_argument("--show-digits", type=int, default=None, metavar="N",
                    help="표 데이터 칸의 **표시** 소수 자릿수를 N 으로 통일한다. "
                         "값 자체는 통계표 원값 그대로 들어가고 서식만 바뀐다. "
                         "안 주면 통계표 셀 서식을 그대로 따른다(백분율 1자리 · 평균 2자리)")
    ap.add_argument("--no-qu-suffix", action="store_true",
                    help="`Qu:` 뒤에 매트릭스 소항목 접미(`… | 계약 기간`)를 붙이지 않는다. "
                         "`Question/Reference` 칸에는 응답자가 읽은 질문만 넣는다는 규율 (사용자 지시 2026-08-26). "
                         "LG 예시에는 파이프 접미가 5건 있으나 SKEE·SH청안주에는 0건이다. "
                         "소항목 구분은 `Sub title` 이 이미 하고 있다")
    ap.add_argument("--no-small-base-mark", action="store_true",
                    help="n<30 배너 셀의 별표 `(13*)` 와 `*Too small base(n<30)` 범례를 **끈다**. "
                         "기본은 켬이다 — 저베이스 경고는 보고서에 반드시 실려야 하고, "
                         "가이드에 적지 않으면 뒤 단계에서 누락되기 쉽다 (사용자 지시 2026-08-26). "
                         "통계표가 이미 별표를 달아 오는 경우에만 끈다")
    ap.add_argument("--notes", type=Path,
                    help="장별 각주 json — {\"notes\":[{\"seq\":5,\"lines\":[\"* …\"]}]}. "
                         "목차의 각주 배정표를 옮긴 것이다. `Question/Reference` 칸에 Qu 다음 줄로 붙는다")
    ap.add_argument("--chart-type", action="append", default=[], metavar="표=문자열",
                    help="차트타입을 사람이 확정해 덮어쓴다 — \"31=전체 : 가로 누적 100%% 막대 그래프\" (반복 가능). "
                         "판별식 밖의 판단일 때만 쓴다. 룰엔진이 낸 값과 다르면 사유를 목차에 적는다")
    ap.add_argument("--mean-units", default="",
                    help="평균 행에 단위가 안 적힌 표의 단위를 준다 — \"27=층;32=개\" (표인덱스=단위, 세미콜론 구분). "
                         "DP 가 `평균 (개월)` 처럼 적어 주면 파서가 읽지만, `평균` 만 적힌 표는 알 수 없다. "
                         "안 주면 Unit 에 평균 단위를 붙이지 않는다 — 틀린 단위보다 낫다")
    ap.add_argument("--resort", default="",
                    help="통계표 정렬이 깨져 전체 값 내림차순으로 다시 세울 표 인덱스 (세미콜론 구분). "
                         "예: \"8;34\". DP 가 이미 정렬해 주는 것이 기본 전제라 자동 판정하지 않는다 — "
                         "구간형 표에 켜면 구간 순서가 망가진다")
    ap.add_argument("--exclude-tables", default="",
                    help="본문에서 빼는 표 인덱스 (R-023 예외④, 사람 확정). 예: \"60=D1-1 n=6 단일값 100%%\" "
                         "— `인덱스=사유` 를 세미콜론으로 구분")
    args = ap.parse_args()

    with io.open(args.questionnaire, encoding="utf-8") as fh:
        qdoc = json.load(fh)
    questions = {q["code"]: q for q in qdoc.get("questions", []) if q.get("code")}

    with io.open(args.tables, encoding="utf-8") as fh:
        tdoc = json.load(fh)
    tables = tdoc["tables"] if isinstance(tdoc, dict) else tdoc

    generated: list[dict] = []
    if args.kpi and args.kpi.exists():
        with io.open(args.kpi, encoding="utf-8") as fh:
            generated = json.load(fh).get("tables", [])

    chart_rules = R.load_chart_rules()
    if not chart_rules:
        print("[중단] chart_rules.yaml 을 읽지 못했습니다. "
              "PyYAML 설치와 engine/knowledge/guide/chart_rules.yaml 을 확인하세요.",
              file=sys.stderr)
        return 1

    piping_map: dict = {}
    if getattr(args, "piping", None) and args.piping.exists():
        with io.open(args.piping, encoding="utf-8") as fh:
            piping_map = json.load(fh)

    exclude: dict[int, str] = {}
    for tok in (args.exclude_tables or "").split(";"):
        tok = tok.strip()
        if not tok:
            continue
        idx, _, why = tok.partition("=")
        if idx.strip().isdigit():
            exclude[int(idx)] = why.strip() or "사람 확정 제외 (R-023 예외④)"
    body, aside = fold_slides(tables, exclude)
    toc, toc_warns = ({}, [])
    if getattr(args, "toc", None) and args.toc.exists():
        toc, toc_warns = parse_toc(args.toc)

    notes_map: dict[int, list[str]] = {}
    if getattr(args, "notes", None) and args.notes.exists():
        with io.open(args.notes, encoding="utf-8") as fh:
            for n in (json.load(fh).get("notes") or []):
                if n.get("seq") is not None and n.get("lines"):
                    notes_map[int(n["seq"])] = [str(x) for x in n["lines"]]

    chart_override: dict[int, str] = {}
    for tok in args.chart_type or []:
        idx, _, ct = tok.partition("=")
        if idx.strip().isdigit() and ct.strip():
            chart_override[int(idx)] = ct.strip()
    resort_tables = {int(x) for x in re.findall(r"\d+", args.resort or "")}
    # 평균 단위 보완 — 파서가 라벨에서 못 읽은 표에만 사람이 준다 (`평균` 만 적힌 표).
    for tok in (args.mean_units or "").split(";"):
        idx, _, unit = tok.strip().partition("=")
        if not (idx.strip().isdigit() and unit.strip()):
            continue
        t = next((x for x in tables if x.get("index") == int(idx)), None)
        if t and (t.get("hints") or {}).get("has_mean_row") and not (t["hints"].get("mean_units") or []):
            t["hints"]["mean_units"] = [unit.strip()]
    freq_banners = {x.strip() for x in (args.freq_banners or "").split(";") if x.strip()}
    slides = [build_slide(u, questions, chart_rules, i, piping_map, resort_tables, freq_banners,
                          small_base_mark=not args.no_small_base_mark,
                          qu_suffix=not args.no_qu_suffix,
                          show_digits=args.show_digits)
              for i, u in enumerate(body, start=1)]
    for sl in slides:
        if sl["seq"] in notes_map:
            sl["note"] = notes_map[sl["seq"]]
        for tidx in sl.get("source_tables") or []:
            if tidx in chart_override:
                sl["chart_type"] = chart_override[tidx]
                sl["chart_rule"] = "사람 확정 (판별식 밖)"
                if "차트타입_확인필요" in sl.get("flags") or []:
                    sl["flags"].remove("차트타입_확인필요")
                break

    # ── 목차 적용: 사람이 확정한 슬라이드/서브 타이틀을 얹는다.
    #    없는 seq 는 채우지 않고 경고로 올린다 — 목차 전수성은 사람이 책임진다.
    if toc:
        missing, mismatched = [], []
        for sl in slides:
            row = toc.get(sl["seq"])
            if not row:
                missing.append(sl["seq"])
                continue
            # 통계표에 문항 코드가 없는 표(매트릭스 항목별로 쪼개져 온 표)는 `Qu` 가 빈 채로 온다.
            # 목차 `문항` 칸이 사람이 확정한 코드이므로 그것으로 설문지를 찾아 채운다.
            # 에스원 TA Pool 실측: 표6·7·8·15~18 일곱 장이 `question_code=null` 이었다.
            if not sl.get("qu") and row.get("question_code"):
                fq = questions.get(row["question_code"].strip())
                if fq:
                    fqu = compose_qu(fq, sl.get("_table") or {}, piping_map)
                    if fqu:
                        sl["qu"] = f"Qu: {fqu}"
                        sl["question_code"] = sl.get("question_code") or row["question_code"].strip()
            sl["slide_title"] = row["slide_title"]
            if row.get("group"):
                sl["group"] = row["group"]              # R-140 대분류 (PPT 팀이 목차·간지에 쓴다)
            sl["sub_title"] = row["sub_title"]
            sl["page"] = row.get("page")
            sl["base_note_toc"] = row.get("base_note_toc")
            want, got = row.get("question_code"), sl.get("question_code")
            if want and got and want != got:
                mismatched.append(f"seq {sl['seq']}: 목차 {want} vs 통계표 {got}")
        if missing:
            toc_warns.append(f"목차에 없는 슬라이드 {len(missing)}장: "
                             + ", ".join(str(x) for x in missing[:15]))
        if mismatched:
            toc_warns.append("문항코드 불일치 " + str(len(mismatched)) + "건: "
                             + " / ".join(mismatched[:5]))
    # ── 사람 확정 제외 표(R-023 예외④)의 흔적을 원 문항 장의 `Question/Reference` 주석으로 남긴다.
    #    조용히 빠뜨리지 않는다 — 보기 1개 100% 같은 단일값은 한 줄이면 전달된다. 수치는 통계표 표시값.
    for ex in aside["excluded"]:
        code = ex.get("question_code") or ""
        parent_code = code.rsplit("-", 1)[0] if "-" in code else None
        parent = next((sl for sl in slides if parent_code and sl.get("question_code") == parent_code), None)
        if parent is None:
            continue
        opts = [r for r in (ex.get("rows") or []) if r.get("kind") == "option"]
        vals = "; ".join(f"'{r.get('label', '').strip()}' {R.cell_value((r.get('values') or [None])[0])}%"
                         for r in opts[:3]) if opts else ""
        # ★ 이 메모는 **독자 대면 자리(Qu)에 쓰지 않는다.** 종전에는 Qu 에 붙여 요약 장표에
        #   그대로 실렸는데 ① 「본문 제외」는 제작 내부 용어이고 ② % 를 인용할 수 없어서 뺀
        #   표(n<30)의 % 를 각주가 그대로 노출해 자기모순이었다 (적대 감사 F16 · 사용자 지적
        #   2026-08-29). 정본에는 남겨 두되(조용히 빠뜨리지 않는다) 렌더 대상에서는 뺀다 —
        #   제외 사실은 연구원 확인 목록으로 전달한다.
        parent.setdefault("exclusion_notes", []).append({
            "question_code": code,
            "title_full": ex.get("title_full"),
            "base_n": ex.get("base_n"),
            "top_options": vals,
        })

    # ── 라벨 정정 표기 — 통계표 후처리가 보기 라벨 내용을 바꿨으면 그 장 Qu 에 한 줄 (ADV-11:
    #    제작팀·연구부가 설문지 원문과 대조할 때 보기명 불일치로 반송하지 않도록)
    for corr in ((tdoc.get("meta") or {}).get("label_corrections") or []) if isinstance(tdoc, dict) else []:
        for sl in slides:
            if corr.get("table") in (sl.get("source_tables") or []):
                note = (f"* 보기 {corr.get('option_no')} 라벨 '{corr.get('new')}' 는 설문지 원문 "
                        f"'{corr.get('old')}'(보기 2와 중복 오타)를 정정 표기한 것 — 연구원 확인")
                sl["qu"] = ((sl.get("qu") or "") + "\n" + note).strip()

    # ── base_label — 서브베이스 장은 군 이름 필수 (guide/17 소비 계약 · `[Base: 전체(n=3)]` 금지)
    total_n = max((t.get("base_n") or 0) for t in tables) if tables else None
    # ── base_is_sub — 전체 표본보다 좁은 모수인가. 서브베이스면 문장에 모수를 반드시 밝혀야 한다
    #    (사용자 지시 2026-08-24: 「해당 테이블이 앞단 설문 어떤 문항을 선택한 사람들한테만 보인 건지」).
    for sl in slides:
        bn = sl.get("base_n")
        sl["base_is_sub"] = bool(bn and total_n and bn < total_n)
    for sl in slides:
        main = next((u["main"] for u in body if u["main"].get("index") in sl["source_tables"]), None)
        q = questions.get(sl.get("question_code")) or {}
        sl["base_label"] = derive_base_label(sl.get("base_note_toc"), q, sl.get("base_n"), total_n)
        if main is not None:
            sl["base_unit"] = R.base_unit_string(main, base_label=sl["base_label"])
    if toc:
        # 제작가이드는 **보고서 페이지 순**으로 나가야 제작팀이 순서대로 만든다.
        # 통계표 순(seq)과 보고서 순(P)은 다르다 — 목차가 흐름을 다시 짰기 때문이다.
        if all(sl.get("page") for sl in slides):
            slides.sort(key=lambda x: x["page"])
            for i, sl in enumerate(slides, start=1):
                sl["order"] = i
        else:
            toc_warns.append("일부 슬라이드에 페이지 번호가 없어 통계표 순서를 유지합니다")

        extra = sorted(set(toc) - {sl["seq"] for sl in slides})
        if extra:
            toc_warns.append(f"목차에만 있고 본문에 없는 seq {len(extra)}개: "
                             + ", ".join(str(x) for x in extra[:15]))

    doc = {
        "meta": {
            "generated_by": "build_facts.py",
            "questionnaire": str(args.questionnaire),
            "tables": str(args.tables),
            "chart_rules_version": chart_rules.get("schema_version"),
            "n_tables": len(tables),
            "n_body_slides": len(slides),
        },
        "excluded": {
            "sq": [{"index": t["index"], "code": t.get("question_code")} for t in aside["sq"]],
            "profile": [{"index": t["index"], "title": t.get("title_full")} for t in aside["profile"]],
            "summary_for_conclusion": [
                {"index": t["index"], "code": t.get("question_code"), "suffix": t.get("suffix")}
                for t in aside["summary"]
            ],
            "human_excluded": [
                {"index": t["index"], "code": t.get("question_code"), "title": t.get("title_full"),
                 "reason": t.get("_reason")}
                for t in aside["excluded"]
            ],
        },
        "generated_summaries": [
            {"index": t["index"], "title": t.get("title_full"), "rows": len(t.get("rows") or [])}
            for t in generated
        ],
        "slides": slides,
    }
    for sl in slides:
        sl.pop("_table", None)      # 목차 적용용 임시 참조를 직렬화 전에 뺀다

    # ── 멘트 주입: AI가 쓴 문장을 얹는다. 숫자는 이미 룰엔진이 확정했으므로
    #    여기서 새로 계산하지 않는다. verify.py 가 멘트의 모든 수치를 그리드와 대조한다.
    if getattr(args, "ments", None) and args.ments.exists():
        with io.open(args.ments, encoding="utf-8") as fh:
            mdoc = json.load(fh)
        mmap = {m["seq"]: m.get("lines") or [] for m in (mdoc.get("ments") or [])
                if isinstance(m.get("seq"), int)}
        filled = 0
        for sl in slides:
            lines = mmap.get(sl["seq"])
            if lines:
                sl["ment"] = [str(x).strip() for x in lines if str(x).strip()]
                filled += 1
        toc_warns.append(f"멘트 주입 {filled}/{len(slides)}장")
        gap = [sl["seq"] for sl in slides if not sl.get("ment")]
        if gap:
            toc_warns.append(f"멘트 없는 슬라이드 {len(gap)}장: "
                             + ", ".join(str(x) for x in gap[:15]))

    args.out.parent.mkdir(parents=True, exist_ok=True)
    with io.open(args.out, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)

    # ── 요약 보고
    print(f"통계표 {len(tables)}장 → 본문 슬라이드 {len(slides)}장")
    print(f"  제외: SQ {len(aside['sq'])} · 응답자특성 {len(aside['profile'])} · 사람확정 {len(aside['excluded'])} "
          f"· 요약표 {len(aside['summary'])}(요약·결론용)")
    if generated:
        print(f"  추가 생성 요약표 {len(generated)}장 (요약·결론용)")

    from collections import Counter
    ct = Counter(s["chart_type"] for s in slides)
    print("\n차트타입:")
    for k, v in ct.most_common():
        print(f"  {v:>3}  {k}")

    flags = Counter(f for s in slides for f in s["flags"])
    if flags:
        print("\n확인 필요:")
        for k, v in flags.most_common():
            print(f"  {v:>3}  {k}")
    for w in toc_warns:
        print(f"[목차] {w}")

    print(f"\n→ {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
