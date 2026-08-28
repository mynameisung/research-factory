# -*- coding: utf-8 -*-
"""
extract_ground_truth.py -- 골든셋 정답지(ground truth) 추출기

제작본 PPTX 1개를 읽어 슬라이드별 사실(fact)을 JSON으로 떨군다.

  ★ 설계 원칙 ★
  정답은 **제작본**이다. 제작가이드의 `Chart type` 문자열이 아니라,
  제작팀이 실제로 만든 차트/표를 역산해 차트타입을 판정한다.
  (가이드 문자열과의 불일치는 extract_guide_blocks.py 결과와 대조해 별도 리포트)

  추출은 **명시적**으로 한다. 휴리스틱이 불가피한 곳에는 confidence 필드를 남긴다.
  - confidence "certain"  : 도형 이름/레이아웃 이름/좌표가 전 슬라이드에서 일관되게 확인됨
  - confidence "high"     : 규칙이 단일 후보만 남김
  - confidence "heuristic": 위치·문맥 추론이 개입함
  - confidence "unknown"  : 판정 불가

사용법:
    python extract_ground_truth.py                     # refs/ 의 예시 2벌 모두 처리
    python extract_ground_truth.py <in.pptx> <out.json>

import 해서 쓰려면:
    from extract_ground_truth import extract_deck
    doc = extract_deck(pptx_path)
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

# --------------------------------------------------------------------------- #
# 상수
# --------------------------------------------------------------------------- #

EMU_PER_CM = 360000
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS = {"c": C_NS}

REFS_DIR = Path(__file__).resolve().parents[1] / "refs"

DECKS = [
    ("예시_제작본_LG홈로봇_최종.pptx", "ground_truth_LG홈로봇.json"),
    ("예시_제작본_SKEE만족도_0127.pptx", "ground_truth_SKEE.json"),
]

# 본문 레이아웃. 두 덱 모두 `3_master-1`.
BODY_LAYOUT = "3_master-1"          # 예시 2벌의 본문 레이아웃. 코퍼스 대조 결과 **사례마다 다르다**

#: 레이아웃 이름 → 장 역할. 이름은 **사례마다 대소문자·꼬리표가 다르다.**
#:   실측(코퍼스 5사례): `cover`(한옥·물리보안) vs `Cover-1`(SH) ·
#:                       `section`(한옥·물리보안) vs `Section`(SH) ·
#:                       `contents`(전 사례 동일) · `contact us`·`Ending-1`(마무리)
#:   예전에는 소문자 정확일치라 SH 의 표지 1장·간지 3장이 **본문으로 떨어졌다.**
LAYOUT_ROLE = (
    ("cover", "표지"),
    ("contents", "목차"),
    ("toc", "목차"),
    ("section", "간지"),
    ("contact", "부록"),
    ("ending", "부록"),
)


def layout_role(layout_name: str) -> str | None:
    """레이아웃 이름에서 장 역할을 읽는다. 대소문자·꼬리표(`-1`)를 무시한다."""
    key = re.sub(r"[\s_-]+", "", (layout_name or "").lower())
    for needle, role in LAYOUT_ROLE:
        if key.startswith(needle):
            return role
    return None


#: 본문 레이아웃 후보 — `master-1`·`3_master-1` 처럼 master 계열 이름
_MASTER_RE = re.compile(r"^\d*_?master", re.IGNORECASE)


def dominant_body_layouts(layout_names: list[str]) -> set[str]:
    """덱에서 **본문을 지고 있는** master 레이아웃을 고른다.

    ★ 본문 레이아웃 이름은 고객사 템플릿마다 다르다 —
      `3_master-1`(SKEE 84장·SH 97장·LG 138장) · `4_master-1`(물리보안 220장) ·
      한옥은 1~5_master-1 에 고루 흩어져 있다(각 4~7장).
      그래서 상수로 박으면 안 되고, **그 덱에서 가장 많이 쓰인 것**을 본문으로 본다.
      흩어져 있으면 어느 것도 압도적이지 않으므로 빈 집합을 주고 Part 문맥에 맡긴다.
    """
    from collections import Counter
    c = Counter(n for n in layout_names if n and _MASTER_RE.match(n))
    if not c:
        return set()
    top, n = c.most_common(1)[0]
    return {top} if n >= max(3, len(layout_names) * 0.25) else set()

# 분류 버킷
BUCKETS = ("표지", "목차", "간지", "조사개요", "요약", "결론", "본문", "부록", "여백")

# Part 간지 제목 → 이후 슬라이드가 속하는 Part
PART_TITLE_TO_PART = [
    ("조사 개요", "조사개요"),
    ("조사개요", "조사개요"),
    ("요약 및 결론", "요약결론"),
    ("요약및결론", "요약결론"),
    ("조사 결과", "본문"),
    ("조사결과", "본문"),
    ("Appendix", "부록"),
    ("APPENDIX", "부록"),
    ("부록", "부록"),
]

# 결론 슬라이드 판별 (요약결론 Part 안에서)
CONCLUSION_TITLE_RE = re.compile(r"결론|제언|SWOT|Implication", re.IGNORECASE)

# 문항코드: 소제목 선두 토큰.
#   A1.  A10-1)  BA1-1)  BB1-5)  C2.  C2-1/C2-2/C2-3.  C2-1-1/C2-2-1/C2-3-1.
#   SQ7-1)  DQ2.
_CODE_ATOM = r"(?:SQ|DQ|[A-Z]{1,3})\d+(?:\s*[-‐-―]\s*\d+)*"
QUESTION_CODE_RE = re.compile(
    r"^\s*(" + _CODE_ATOM + r"(?:\s*/\s*" + _CODE_ATOM + r")*)\s*[.)]"
)

# Base/Unit 텍스트박스: 대괄호 Base 로 시작
BASE_PREFIX_RE = re.compile(r"^\s*\[\s*Base", re.IGNORECASE)
# Base/Unit 2번째 줄 각주
FOOTNOTE_RE = re.compile(r"^\s*\*")
# Base n 추출
BASE_N_RE = re.compile(r"\(\s*n\s*=\s*([0-9,]+)\s*\)", re.IGNORECASE)

QU_PREFIX_RE = re.compile(r"^\s*Qu\s*[:：]", re.IGNORECASE)

# 미기입 잔존 마커 — 예시 제작본의 알려진 결함 탐지용
UNFILLED_MARKERS = ("???", "내용을 입력하세요", "내용을입력하세요", "TBD", "XXX")

# 표 센티널 라벨
BASE_ROW_LABELS = ("(Base)", "Base for %", "(base)")

# 순위 계열 이름:  `1순위` `1+2순위` `1+2+3순위`
RANK_SERIES_RE = re.compile(r"^\s*\d+(\s*\+\s*\d+)*\s*순위\s*$")

# 제작팀 메모 판정: top ≈ 0 인 AutoShape
MEMO_TOP_TOLERANCE_CM = 0.05


# --------------------------------------------------------------------------- #
# 저수준 유틸
# --------------------------------------------------------------------------- #


def cm(emu: int | None) -> float | None:
    """EMU → cm (소수 2자리). None 은 그대로 통과."""
    if emu is None:
        return None
    return round(int(emu) / EMU_PER_CM, 2)


def _text(shape) -> str:
    """도형 텍스트. vertical-tab(\\x0b)은 줄바꿈으로 정규화."""
    if not shape.has_text_frame:
        return ""
    return shape.text_frame.text.replace("\x0b", "\n")


def _lines(s: str) -> list[str]:
    return [ln for ln in s.split("\n")]


def norm_ws(s: str | None) -> str:
    """공백 정규화 (매핑 키 비교용)."""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", s.replace("　", " ")).strip()


def norm_key(s: str | None) -> str:
    """조인 키 정규화 — 공백 전부 제거 + 콜론 통일 + 소문자화."""
    if s is None:
        return ""
    s = s.replace("：", ":")
    return re.sub(r"\s+", "", s).lower()


def flatten_shapes(shapes, depth: int = 0, path: tuple = ()) -> list[dict]:
    """그룹을 재귀적으로 펼쳐 (depth, path, shape) 목록을 만든다."""
    out: list[dict] = []
    for sh in shapes:
        entry = {"depth": depth, "path": path + (sh.name,), "shape": sh}
        out.append(entry)
        if sh.shape_type == 6:  # GROUP
            try:
                out.extend(flatten_shapes(sh.shapes, depth + 1, path + (sh.name,)))
            except Exception:  # pragma: no cover - 방어
                pass
    return out


def geom(shape) -> dict:
    return {
        "left_cm": cm(shape.left),
        "top_cm": cm(shape.top),
        "width_cm": cm(shape.width),
        "height_cm": cm(shape.height),
    }


def _para_alignment(shape) -> str | None:
    """첫 문단 정렬. 명시가 없으면 None (상속)."""
    if not shape.has_text_frame:
        return None
    for p in shape.text_frame.paragraphs:
        if p.alignment is not None:
            return str(p.alignment)
    return None


def _fill_desc(shape) -> str | None:
    try:
        f = shape.fill
        if f.type is None:
            return None
        if f.type == 1:  # solid
            fc = f.fore_color
            if fc.type == 1:  # RGB
                return "solid:#%s" % str(fc.rgb)
            return "solid:%s" % str(fc.theme_color)
        return str(f.type)
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# 차트 추출
# --------------------------------------------------------------------------- #


def _axis_scaling(chart_space, ax_tag: str) -> dict:
    """valAx / catAx 의 scaling min/max 를 뽑는다."""
    out: dict[str, Any] = {"min": None, "max": None, "num_fmt": None}
    for ax in chart_space.iter("{%s}%s" % (C_NS, ax_tag)):
        sc = ax.find("c:scaling", NS)
        if sc is not None:
            mn = sc.find("c:min", NS)
            mx = sc.find("c:max", NS)
            if mn is not None and out["min"] is None:
                out["min"] = float(mn.get("val"))
            if mx is not None and out["max"] is None:
                out["max"] = float(mx.get("val"))
        nf = ax.find("c:numFmt", NS)
        if nf is not None and out["num_fmt"] is None:
            out["num_fmt"] = nf.get("formatCode")
    return out


def _data_label_num_fmts(chart_space) -> list[str]:
    """dLbls 아래 numFmt formatCode 를 등장 순서대로."""
    out = []
    for dl in chart_space.iter("{%s}dLbls" % C_NS):
        nf = dl.find("c:numFmt", NS)
        if nf is not None:
            out.append(nf.get("formatCode"))
    return out


def extract_chart(shape) -> dict:
    """차트 1개 → dict. python-pptx 가 못 읽는 값은 None 으로 남긴다."""
    chart = shape.chart
    cs = chart._chartSpace

    plots = list(chart.plots)
    plot = plots[0] if plots else None

    series = []
    for s in chart.series:
        vals = [None if v is None else float(v) for v in s.values]
        real = any(v is not None for v in vals)
        series.append(
            {
                "name": s.name,
                "values": vals,
                "is_real": real,  # None 으로만 채워진 더미 계열 구분
                "has_negative": any(v is not None and v < 0 for v in vals),
            }
        )

    try:
        categories = [None if c is None else str(c) for c in plot.categories] if plot else []
    except Exception:
        categories = []

    val_ax = _axis_scaling(cs, "valAx")
    cat_ax = _axis_scaling(cs, "catAx")

    # 순위 계열 (`1순위` / `1+2+3순위`) — 가로겹침막대의 결정적 표지
    rank_names = [s["name"] for s in series if s["name"] and RANK_SERIES_RE.match(s["name"])]

    # ★ 겹침막대 판별의 핵심: **계열이 같은 카테고리를 공유하는가**
    #
    #   진짜 가로겹침막대(순위)  : 모든 카테고리에 두 계열이 동시에 값을 가진다
    #   색 분리용 분할 계열       : 카테고리를 서로 **나눠** 가진다 (겹치는 칸 0개)
    #        실측 - LG s20 `do not know`(잘 모르겠다 칸만), LG s96/s97 `열2`(그룹 헤더 칸만),
    #               SKEE s46/s63 `열2`(없다/특별한 이유 없음 칸만)
    real_series = [s for s in series if s["is_real"]]
    n_cat = max((len(s["values"]) for s in real_series), default=0)
    overlapping_categories = 0
    cat_sums: list[float] = []
    for i in range(n_cat):
        filled = 0
        tot = 0.0
        for s in real_series:
            v = s["values"][i] if i < len(s["values"]) else None
            if v is not None:
                filled += 1
                tot += v
        if filled >= 2:
            overlapping_categories += 1
        if filled:
            cat_sums.append(tot)
    complements_to_100 = bool(cat_sums) and all(abs(x - 100.0) < 0.6 for x in cat_sums)

    return {
        "shape_name": shape.name,
        "chart_type": str(chart.chart_type),
        "chart_type_id": int(chart.chart_type) if chart.chart_type is not None else None,
        "plot_class": type(plot).__name__ if plot is not None else None,
        "plot_count": len(plots),
        "series_count": len(series),
        "real_series_count": len(real_series),
        "rank_series_names": rank_names,
        "overlapping_categories": overlapping_categories,
        "complements_to_100": complements_to_100,
        "series_names": [s["name"] for s in series],
        "series": series,
        "any_negative_series": any(s["has_negative"] for s in series),
        "categories": categories,
        "category_count": len(categories),
        "overlap": getattr(plot, "overlap", None) if plot is not None else None,
        "gap_width": getattr(plot, "gap_width", None) if plot is not None else None,
        "val_ax_min": val_ax["min"],
        "val_ax_max": val_ax["max"],
        "val_ax_num_fmt": val_ax["num_fmt"],
        "cat_ax_num_fmt": cat_ax["num_fmt"],
        "data_label_num_fmts": _data_label_num_fmts(cs),
        "has_legend": chart.has_legend,
        **geom(shape),
    }


# --------------------------------------------------------------------------- #
# 표 추출
# --------------------------------------------------------------------------- #


def extract_table(shape) -> dict:
    """
    표 1개 → dict.

    통계표 표준 구조:
        r0  배너 그룹 행   ['　','전체','나이','','','','가구 구성','', ...]
        r1  배너 상세 행   ['','','20대','30대', ...]
        r2  (Base) 행      ['(Base)','(100)','(8*)', ...]
        r3+ 데이터 행
    다만 슬라이드에 따라 헤더가 1행이거나 (Base) 행 위치가 다를 수 있으므로
    `base_row_index` 는 실제로 찾아서 기록한다 (없으면 None).
    """
    tbl = shape.table
    rows: list[list[str]] = []
    for r in tbl.rows:
        rows.append([c.text.replace("\x0b", "\n") for c in r.cells])

    n_rows = len(rows)
    n_cols = len(tbl.columns) if n_rows else 0

    # (Base) / Base for % 행 찾기 — 첫 열 기준
    base_row_index = None
    for i, row in enumerate(rows):
        first = norm_ws(row[0]) if row else ""
        if first in ("(Base)", "Base for %", "(base)") or first.replace(" ", "") in (
            "(base)",
            "basefor%",
        ):
            base_row_index = i
            break

    header_rows = rows[:base_row_index] if base_row_index is not None else rows[:2]
    data_rows = rows[base_row_index + 1 :] if base_row_index is not None else rows[2:]

    return {
        "shape_name": shape.name,
        "n_rows": n_rows,
        "n_cols": n_cols,
        "col_widths_cm": [cm(c.width) for c in tbl.columns],
        "row_heights_cm": [cm(r.height) for r in tbl.rows],
        "first_row_header": tbl.first_row,
        "banner_group_row": header_rows[0] if len(header_rows) >= 1 else None,
        "banner_detail_row": header_rows[1] if len(header_rows) >= 2 else None,
        "header_rows": header_rows,
        "base_row_index": base_row_index,
        "base_row_label": rows[base_row_index][0] if base_row_index is not None else None,
        "base_row_values": rows[base_row_index][1:] if base_row_index is not None else None,
        "data_rows": data_rows,
        "data_row_labels": [r[0] for r in data_rows],
        "rows": rows,  # 전량 원문 보존 (정답지이므로)
        **geom(shape),
    }


# --------------------------------------------------------------------------- #
# 슬라이드 분류
# --------------------------------------------------------------------------- #


def _part_of_section_title(title: str) -> str | None:
    for needle, part in PART_TITLE_TO_PART:
        if needle in title:
            return part
    return None


def classify(layout_name: str, current_part: str | None, slide_title: str, has_any_text: bool,
             body_layouts: set[str] | None = None):
    """
    (classification, confidence, reason) 반환.
    레이아웃 이름이 1차 근거이고, 그것으로 안 갈리는 경우에만 직전 Part 간지 문맥을 쓴다.
    """
    role = layout_role(layout_name)
    if role:
        return role, "certain", "layout=%s (역할=%s)" % (layout_name, role)
    if layout_name == "빈 화면":
        if not has_any_text:
            return "여백", "certain", "layout=빈 화면 & 텍스트 없음"
        return "간지", "high", "layout=빈 화면 & 타이틀 있음(장 구분 간지)"
    if layout_name in (body_layouts if body_layouts is not None else {BODY_LAYOUT}):
        return "본문", "certain", "layout=%s (덱의 본문 레이아웃)" % layout_name

    # master-1 / 1_master-1 / 2_master-1 / 44_master-1 → 직전 Part 간지 문맥으로 판정
    if current_part == "조사개요":
        return "조사개요", "high", "직전 Part 간지='조사 개요'"
    if current_part == "요약결론":
        if CONCLUSION_TITLE_RE.search(slide_title or ""):
            return "결론", "high", "직전 Part 간지='요약 및 결론' & 타이틀에 결론/제언"
        return "요약", "high", "직전 Part 간지='요약 및 결론'"
    if current_part == "부록":
        return "부록", "high", "직전 Part 간지='Appendix'"
    if current_part == "본문":
        return "본문", "heuristic", "직전 Part 간지='조사 결과' (비표준 레이아웃)"
    return "본문", "unknown", "Part 문맥 없음 — 기본값"


# --------------------------------------------------------------------------- #
# 차트타입 역산
# --------------------------------------------------------------------------- #


def derive_chart_type(
    charts: list[dict], tables: list[dict], classification: str, has_body_text: bool = False
):
    """
    제작본 실측 → 가이드 차트타입 문자열 역산.

    §9.2 역산표:
      전체 : 가로막대, 배너 : 표          BAR_CLUSTERED, 실계열 1 (+None 더미), 표 있음
      전체 : 가로겹침막대, 배너 : 표        BAR_CLUSTERED, 실계열 2, overlap=100, 긴 값 계열 먼저
      전체 : +-세로막대 그래프             COLUMN_STACKED(±100, 7점) / COLUMN_STACKED_100(±1, 5점)
                                          음수 계열 존재, numFmt "0.0;0.0;\\ "
      전체 : 원그래프                     DOUGHNUT
      전체 : 가로 누적 100% 막대 그래프     BAR_STACKED_100
      전체 : 표                          차트 없음, 표만
      간지 / 표지 / 목차                  차트 없음, 레이아웃이 section/cover/contents

    반환: (derived, confidence, evidence:list[str])
    다단 슬라이드는 패널별 문자열을 ' + ' 로 연결한다.
    """
    ev: list[str] = []

    if classification == "표지":
        return "표지", "certain", ["classification=표지"], ["표지"]
    if classification == "목차":
        return "목차", "certain", ["classification=목차"], ["목차"]
    if classification == "간지":
        return "간지", "certain", ["classification=간지"], ["간지"]
    if classification == "여백":
        return "여백", "certain", ["classification=여백"], ["여백"]

    if not charts:
        if tables:
            return "전체 : 표", "certain", ["차트 0개, 표 %d개" % len(tables)], ["전체 : 표"]
        if has_body_text:
            # 차트도 표도 없는 본문 = 텍스트만 있는 장표
            # (SKEE 의 IDI 정성 인용 장표, 개선사항 텍스트 박스 장표)
            return ("전체 : 텍스트 박스", "heuristic", ["차트 0개, 표 0개, 본문 텍스트만 존재"], ["전체 : 텍스트 박스"])
        return None, "unknown", ["차트 0개, 표 0개"], []

    parts: list[str] = []
    confs: list[str] = []

    for ci, ch in enumerate(charts):
        ct = ch["chart_type"]
        real = ch["real_series_count"]
        total = ch["series_count"]
        overlap = ch["overlap"]
        neg = ch["any_negative_series"]
        vmin, vmax = ch["val_ax_min"], ch["val_ax_max"]

        if ct.startswith("DOUGHNUT") or ct.startswith("PIE"):
            parts.append("전체 : 원그래프")
            confs.append("certain")
            ev.append("chart[%d] %s → 원그래프" % (ci, ct))

        elif ct.startswith("BAR_STACKED_100"):
            parts.append("전체 : 가로 누적 100% 막대 그래프")
            confs.append("certain")
            ev.append("chart[%d] %s → 가로 누적 100%%" % (ci, ct))

        elif ct.startswith("COLUMN_STACKED_100"):
            # 5점척도 +- 세로막대 (valAx ±1)
            if neg:
                parts.append("전체 : +-세로막대 그래프")
                confs.append("certain")
                ev.append(
                    "chart[%d] COLUMN_STACKED_100 valAx[%s,%s] 음수계열有 → +-세로막대(5점)"
                    % (ci, vmin, vmax)
                )
            else:
                parts.append("전체 : 세로 누적 100% 막대 그래프")
                confs.append("heuristic")
                ev.append(
                    "chart[%d] COLUMN_STACKED_100 음수계열無 → 누적100%%(부호 미사용)" % ci
                )

        elif ct.startswith("COLUMN_STACKED"):
            # 7점척도 +- 세로막대 (valAx ±100)
            if neg:
                parts.append("전체 : +-세로막대 그래프")
                confs.append("certain")
                ev.append(
                    "chart[%d] COLUMN_STACKED valAx[%s,%s] 음수계열有 → +-세로막대(7점)"
                    % (ci, vmin, vmax)
                )
            else:
                parts.append("전체 : 세로 누적 막대 그래프")
                confs.append("heuristic")
                ev.append("chart[%d] COLUMN_STACKED 음수계열無" % ci)

        elif ct.startswith("BAR_CLUSTERED"):
            ranks = ch.get("rank_series_names") or []
            real_series = [s for s in ch["series"] if s["is_real"]]
            sums = [
                sum(v for v in (s["values"] or []) if v is not None) for s in real_series
            ]

            if len(ranks) >= 2:
                # 순위 겹침막대 — 계열 이름이 `1+2+3순위` / `1순위`
                parts.append("전체 : 가로겹침막대, 배너 : 표")
                confs.append("certain" if overlap == 100 else "heuristic")
                ev.append(
                    "chart[%d] BAR_CLUSTERED 순위계열%s overlap=%s 계열합%s 긴값먼저=%s"
                    % (
                        ci,
                        ranks,
                        overlap,
                        [round(x, 1) for x in sums],
                        bool(sums) and sums[0] >= max(sums),
                    )
                )
            elif real >= 2 and ch.get("overlapping_categories") == 0:
                # 계열이 카테고리를 나눠 가졌다 = 특정 막대만 색을 달리하려고 쪼갠 것.
                # 겹침막대가 아니라 실질 1계열 가로막대다.
                parts.append("전체 : 가로막대, 배너 : 표")
                confs.append("high")
                ev.append(
                    "chart[%d] BAR_CLUSTERED 실계열%d 이나 겹치는 카테고리 0개 "
                    "(색 분리용 분할 계열 %s) -> 실질 1계열"
                    % (ci, real, ch["series_names"])
                )
            elif real >= 2:
                parts.append("전체 : 가로겹침막대, 배너 : 표")
                confs.append("heuristic")
                ev.append(
                    "chart[%d] BAR_CLUSTERED 실계열%d overlap=%s 겹칩카테고리%d 계열명%s "
                    "(순위계열명 아님)"
                    % (ci, real, overlap, ch.get("overlapping_categories"), ch["series_names"])
                )
            else:
                parts.append("전체 : 가로막대, 배너 : 표")
                confs.append("certain")
                ev.append(
                    "chart[%d] BAR_CLUSTERED 실계열1 (총계열%d, 더미%d) 표%d개"
                    % (ci, total, total - real, len(tables))
                )

        elif ct.startswith("COLUMN_CLUSTERED"):
            parts.append("전체 : 세로막대")
            confs.append("heuristic")
            ev.append("chart[%d] COLUMN_CLUSTERED (역산표 미등재)" % ci)

        elif ct.startswith("XY_SCATTER"):
            parts.append("전체 : 분산형(IPA)")
            confs.append("heuristic")
            ev.append("chart[%d] XY_SCATTER (역산표 미등재)" % ci)

        else:
            parts.append("UNKNOWN(%s)" % ct)
            confs.append("unknown")
            ev.append("chart[%d] %s 역산표 미등재" % (ci, ct))

    # 다단 슬라이드는 같은 타입이 15번씩 반복되기도 한다 -> 순서를 지키며 중복만 제거.
    # 원본 패널 목록은 derived_chart_type_panels 로 따로 남긴다.
    derived = " + ".join(dict.fromkeys(parts))
    order = {"unknown": 0, "heuristic": 1, "high": 2, "certain": 3}
    conf = min(confs, key=lambda c: order[c]) if confs else "unknown"
    if tables:
        ev.append("표 %d개 (열수 %s)" % (len(tables), [t["n_cols"] for t in tables]))
    return derived, conf, ev, parts


# --------------------------------------------------------------------------- #
# 슬라이드 1장 추출
# --------------------------------------------------------------------------- #


def extract_slide(slide, slide_no: int, current_part: str | None,
                  body_layouts: set[str] | None = None) -> dict:
    layout_name = slide.slide_layout.name
    flat = flatten_shapes(slide.shapes)

    # ---- 플레이스홀더 수집 -------------------------------------------------- #
    placeholders: dict[int, dict] = {}
    for e in flat:
        sh = e["shape"]
        try:
            if not sh.is_placeholder:
                continue
        except Exception:
            continue
        idx = sh.placeholder_format.idx
        placeholders[idx] = {
            "idx": idx,
            "type": str(sh.placeholder_format.type),
            "name": sh.name,
            "text": _text(sh),
            **geom(sh),
        }

    slide_title = ""
    if 0 in placeholders:
        slide_title = placeholders[0]["text"].strip()

    # ---- sub_title / top_ment ---------------------------------------------- #
    #
    # ⚠ 두 덱에서 placeholder idx 의 역할이 서로 뒤바뀐다 (실측):
    #     LG   : idx=12 → sub_title (top≈2.62, h≈0.60) / idx=13 → top_ment (top≈3.42, h≈2.00)
    #     SKEE : idx=13 → sub_title (top≈2.68, h≈0.75) / idx=12 → top_ment (top≈3.60, h≈2.50)
    # 따라서 idx 가 아니라 **기하(제목 아래 첫 줄 = sub_title, 그 아래 = top_ment)** 로
    # 판정하고, idx 별 원문도 함께 보존한다.
    body_phs = [
        p
        for i, p in placeholders.items()
        if i in (12, 13) and p["top_cm"] is not None
    ]
    body_phs.sort(key=lambda p: p["top_cm"])
    sub_title = body_phs[0]["text"].strip() if len(body_phs) >= 1 else ""
    top_ment = body_phs[1]["text"].strip() if len(body_phs) >= 2 else ""
    ph_role_conf = "high" if len(body_phs) == 2 else ("heuristic" if body_phs else "unknown")

    # ---- 소제목 (문항코드 보유 도형) ---------------------------------------- #
    #
    # LG   : AutoShape name='소제목', L≈0.93 T≈5.54, 0.5×0.5cm, word_wrap=False
    # SKEE : 소제목 도형 자체가 없다(실측). → 문항코드 없음.
    subtitle_shapes = []
    for e in flat:
        sh = e["shape"]
        t = _text(sh).strip()
        if not t:
            continue
        if sh.name == "소제목":
            subtitle_shapes.append(
                {
                    "source": "shape_name=소제목",
                    "text": t,
                    "depth": e["depth"],
                    "word_wrap": sh.text_frame.word_wrap if sh.has_text_frame else None,
                    **geom(sh),
                }
            )
    # 이름이 소제목이 아니어도 문항코드를 달고 있는 도형이 있으면 후보로 잡는다.
    if not subtitle_shapes:
        for e in flat:
            sh = e["shape"]
            t = _text(sh).strip()
            if not t or QU_PREFIX_RE.match(t):
                continue
            if QUESTION_CODE_RE.match(t) and (sh.top or 0) / EMU_PER_CM < 7.0:
                subtitle_shapes.append(
                    {
                        "source": "code_pattern",
                        "text": t,
                        "depth": e["depth"],
                        "word_wrap": sh.text_frame.word_wrap if sh.has_text_frame else None,
                        **geom(sh),
                    }
                )

    question_code = None
    question_code_conf = "unknown"
    if subtitle_shapes:
        m = QUESTION_CODE_RE.match(subtitle_shapes[0]["text"])
        if m:
            question_code = re.sub(r"\s+", "", m.group(1))
            question_code_conf = (
                "certain" if subtitle_shapes[0]["source"] == "shape_name=소제목" else "high"
            )
        else:
            question_code_conf = "unknown"

    # ---- Qu 텍스트박스 ------------------------------------------------------ #
    qu_items = []
    for e in flat:
        sh = e["shape"]
        t = _text(sh).strip()
        if not t:
            continue
        is_qu_named = sh.name == "Qu"
        is_qu_prefixed = bool(QU_PREFIX_RE.match(t))
        if is_qu_named or is_qu_prefixed:
            qu_items.append(
                {
                    "text": t,
                    "source": "shape_name=Qu" if is_qu_named else "prefix=Qu:",
                    "alignment": _para_alignment(sh),
                    **geom(sh),
                }
            )
    qu_items.sort(key=lambda q: (q["top_cm"] if q["top_cm"] is not None else 0,
                                 q["left_cm"] if q["left_cm"] is not None else 0))
    qu_text = qu_items[0]["text"] if qu_items else None

    # ---- Base/Unit 텍스트박스 ----------------------------------------------- #
    base_items = []
    for e in flat:
        sh = e["shape"]
        t = _text(sh)
        if not BASE_PREFIX_RE.match(t):
            continue
        lines = [ln for ln in _lines(t)]
        first = lines[0].strip()
        rest = [ln.strip() for ln in lines[1:] if ln.strip()]
        n_match = BASE_N_RE.search(first)
        base_items.append(
            {
                "text": t.strip(),
                "base_unit_line": first,
                "footnotes": [ln for ln in rest if FOOTNOTE_RE.match(ln)],
                "other_lines": [ln for ln in rest if not FOOTNOTE_RE.match(ln)],
                "base_n": int(n_match.group(1).replace(",", "")) if n_match else None,
                "shape_name": sh.name,
                "alignment": _para_alignment(sh),
                **geom(sh),
            }
        )
    base_items.sort(key=lambda b: (b["top_cm"] if b["top_cm"] is not None else 0,
                                   b["left_cm"] if b["left_cm"] is not None else 0))
    base_unit = base_items[0]["base_unit_line"] if base_items else None
    base_unit_footnotes = base_items[0]["footnotes"] if base_items else []

    # ---- 차트 / 표 ----------------------------------------------------------- #
    charts, tables = [], []
    for e in flat:
        sh = e["shape"]
        try:
            if sh.has_chart:
                charts.append(extract_chart(sh))
        except Exception as exc:  # pragma: no cover
            charts.append({"shape_name": sh.name, "error": repr(exc)})
        try:
            if sh.has_table:
                tables.append(extract_table(sh))
        except Exception as exc:  # pragma: no cover
            tables.append({"shape_name": sh.name, "error": repr(exc)})
    charts.sort(key=lambda c: (c.get("top_cm") or 0, c.get("left_cm") or 0))
    tables.sort(key=lambda t: (t.get("top_cm") or 0, t.get("left_cm") or 0))

    # ---- 제작팀 메모 (top ≈ 0 에 핀된 직사각형) ------------------------------ #
    memos = []
    for e in flat:
        sh = e["shape"]
        if sh.shape_type != 1:  # AUTO_SHAPE 만
            continue
        if sh.top is None or abs(sh.top) > MEMO_TOP_TOLERANCE_CM * EMU_PER_CM:
            continue
        t = _text(sh).strip()
        if not t:
            continue
        memos.append(
            {
                "text": t,
                "lines": [ln.strip() for ln in _lines(t) if ln.strip()],
                "shape_name": sh.name,
                "depth": e["depth"],
                "fill": _fill_desc(sh),
                **geom(sh),
            }
        )
    memos.sort(key=lambda m: (m["depth"], m["left_cm"] or 0))

    # ---- 텍스트 도형 전량 인벤토리 -------------------------------------------- #
    # 정답지는 자기충족적이어야 한다. Qu/Base/소제목/플레이스홀더로 분류되지 않은
    # 자유 텍스트박스(범례, 【TOP2】 ???% 주석, `내용을 입력하세요` 템플릿 잔존 등)도
    # 전부 보존해 채점기가 나중에 참조할 수 있게 한다.
    text_shapes = []
    unfilled_markers = []
    for e in flat:
        sh = e["shape"]
        t = _text(sh)
        if not t.strip():
            continue
        rec = {
            "name": sh.name,
            "shape_type": str(sh.shape_type),
            "depth": e["depth"],
            "text": t.strip(),
            **geom(sh),
        }
        text_shapes.append(rec)
        for mk in UNFILLED_MARKERS:
            if mk in t:
                unfilled_markers.append({"marker": mk, "shape": sh.name, "text": t.strip()})
    for tb in tables:
        for row in tb.get("rows", []):
            for cellv in row:
                for mk in UNFILLED_MARKERS:
                    if mk in cellv:
                        unfilled_markers.append(
                            {"marker": mk, "shape": tb.get("shape_name"), "text": cellv.strip()}
                        )

    # ---- 분류 ---------------------------------------------------------------- #
    has_any_text = bool(text_shapes)
    classification, cls_conf, cls_reason = classify(
        layout_name, current_part, slide_title, has_any_text, body_layouts
    )

    # 제목/서브타이틀/Qu/Base 외에 본문 영역(T>4cm)에 실제 내용 텍스트가 있는가
    has_body_text = any(
        (ts["top_cm"] or 0) > 4.0
        and ts["name"] not in ("Qu", "Base/Unit", "소제목")
        and not BASE_PREFIX_RE.match(ts["text"])
        for ts in text_shapes
    )
    derived, derived_conf, derived_ev, derived_panels = derive_chart_type(
        charts, tables, classification, has_body_text
    )

    # ---- 조인 키 -------------------------------------------------------------- #
    #  1순위: 문항코드 (LG)
    #  2순위: slide_title ‖ sub_title 정규화 (SKEE — 덱에 문항코드가 없다)
    if question_code:
        join_key, join_kind = question_code, "question_code"
    else:
        jk = norm_key(slide_title) + "|" + norm_key(sub_title)
        join_key, join_kind = (jk if jk.strip("|") else None), "title+subtitle"

    return {
        "slide_no": slide_no,
        "layout_name": layout_name,
        "classification": classification,
        "classification_confidence": cls_conf,
        "classification_reason": cls_reason,
        "part_context": current_part,
        "question_code": question_code,
        "question_code_confidence": question_code_conf,
        "subtitle_shape_text": subtitle_shapes[0]["text"] if subtitle_shapes else None,
        "subtitle_shapes": subtitle_shapes,
        "join_key": join_key,
        "join_key_kind": join_kind,
        "slide_title": slide_title,
        "sub_title": sub_title,
        "top_ment": top_ment,
        "placeholder_role_confidence": ph_role_conf,
        "placeholders_by_idx": {str(k): v for k, v in sorted(placeholders.items())},
        "qu_text": qu_text,
        "qu_items": qu_items,
        "base_unit": base_unit,
        "base_unit_footnotes": base_unit_footnotes,
        "base_items": base_items,
        "charts": charts,
        "chart_count": len(charts),
        "tables": tables,
        "table_count": len(tables),
        "production_memos": memos,
        "production_memo_count": len(memos),
        "text_shapes": text_shapes,
        "unfilled_markers": unfilled_markers,
        "derived_chart_type": derived,
        "derived_chart_type_panels": derived_panels,
        "derived_panel_count": len(derived_panels),
        "derived_chart_type_confidence": derived_conf,
        "derived_chart_type_evidence": derived_ev,
    }


# --------------------------------------------------------------------------- #
# 덱 1벌 추출
# --------------------------------------------------------------------------- #


def extract_deck(pptx_path: str | Path) -> dict:
    pptx_path = Path(pptx_path)
    prs = Presentation(str(pptx_path))

    slides: list[dict] = []
    current_part: str | None = None

    # ★ 본문 레이아웃 이름은 고객사 템플릿마다 다르다 — 덱을 한 번 훑어 정한다.
    body_layouts = dominant_body_layouts([sl.slide_layout.name for sl in prs.slides])

    for i, slide in enumerate(prs.slides, start=1):
        rec = extract_slide(slide, i, current_part, body_layouts)
        slides.append(rec)
        # 간지를 지나면 Part 문맥이 바뀐다.
        # ★ 예전에는 `== "section"` 이라 SH 의 `Section` 을 못 알아봤고, 그 바람에
        #   Part 문맥이 끝까지 안 잡혀 **112장 전부가 본문**으로 떨어졌다(실측).
        if layout_role(slide.slide_layout.name) == "간지":
            part = _part_of_section_title(rec["slide_title"])
            if part:
                current_part = part

    # 집계
    by_class: dict[str, int] = {}
    for s in slides:
        by_class[s["classification"]] = by_class.get(s["classification"], 0) + 1
    by_layout: dict[str, int] = {}
    for s in slides:
        by_layout[s["layout_name"]] = by_layout.get(s["layout_name"], 0) + 1
    memo_total = sum(s["production_memo_count"] for s in slides)
    memo_texts: dict[str, int] = {}
    for s in slides:
        for m in s["production_memos"]:
            memo_texts[m["text"]] = memo_texts.get(m["text"], 0) + 1
    chart_types: dict[str, int] = {}
    for s in slides:
        for c in s["charts"]:
            ct = c.get("chart_type", "ERROR")
            chart_types[ct] = chart_types.get(ct, 0) + 1

    body = [s for s in slides if s["classification"] == "본문"]
    codes = [s["question_code"] for s in body if s["question_code"]]
    dup = sorted({c for c in codes if codes.count(c) > 1})

    return {
        "source_file": pptx_path.name,
        "source_path": str(pptx_path),
        "slide_count": len(slides),
        "slide_width_cm": cm(prs.slide_width),
        "slide_height_cm": cm(prs.slide_height),
        "summary": {
            "by_classification": dict(sorted(by_class.items())),
            "by_layout": dict(sorted(by_layout.items())),
            "body_slide_count": len(body),
            "body_with_question_code": len(codes),
            "duplicate_question_codes": dup,
            "production_memo_total": memo_total,
            "production_memo_slides": sum(
                1 for s in slides if s["production_memo_count"] > 0
            ),
            "production_memo_texts": dict(
                sorted(memo_texts.items(), key=lambda kv: -kv[1])
            ),
            "chart_type_counts": dict(sorted(chart_types.items(), key=lambda kv: -kv[1])),
            "slides_with_unfilled_markers": [
                {"slide_no": s["slide_no"], "markers": sorted({m["marker"] for m in s["unfilled_markers"]})}
                for s in slides
                if s["unfilled_markers"]
            ],
            "derived_chart_type_counts": _count(
                [s["derived_chart_type"] for s in slides if s["derived_chart_type"]]
            ),
        },
        "slides": slides,
    }


def _count(items) -> dict:
    out: dict[str, int] = {}
    for it in items:
        out[it] = out.get(it, 0) + 1
    return dict(sorted(out.items(), key=lambda kv: -kv[1]))


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #


def _write(doc: dict, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8"
    )


def main(argv: list[str]) -> int:
    if len(argv) == 3:
        doc = extract_deck(argv[1])
        _write(doc, Path(argv[2]))
        print("%s → %s  (%d slides)" % (argv[1], argv[2], doc["slide_count"]))
        return 0

    if len(argv) != 1:
        print(__doc__)
        return 2

    for src, dst in DECKS:
        src_p = REFS_DIR / src
        if not src_p.exists():
            print("SKIP (없음): %s" % src_p)
            continue
        doc = extract_deck(src_p)
        _write(doc, REFS_DIR / dst)
        s = doc["summary"]
        print("=" * 78)
        print(
            "%s  →  %s\n  slides=%d  %.2f x %.2f cm"
            % (src, dst, doc["slide_count"], doc["slide_width_cm"], doc["slide_height_cm"])
        )
        print("  분류:", s["by_classification"])
        print("  본문:", s["body_slide_count"], " / 문항코드 보유:", s["body_with_question_code"])
        print("  중복 문항코드:", s["duplicate_question_codes"])
        print("  제작팀 메모:", s["production_memo_total"], "개 /", s["production_memo_slides"], "장")
        print("  실측 차트타입:", s["chart_type_counts"])
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main(sys.argv))
