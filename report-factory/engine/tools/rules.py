#!/usr/bin/env python3
"""룰 엔진 — 통계표에서 슬라이드 팩트를 결정론적으로 뽑는다.

설계 원칙: **AI는 숫자를 읽지 않는다.**
여기서 모든 수치 판단을 끝내고, AI는 그 결과를 문장으로 옮기기만 한다. 그래야 환각이
구조적으로 불가능해지고, 생성된 멘트의 모든 숫자를 통계표 원본과 1:1 대조할 수 있다.

담당:
  - 차트타입 결정        `chart_rules.yaml` 기반
  - 정렬                전체값 내림차순 vs 원래 순서 유지
  - Base/Unit 문자열
  - 배너 유의차          2-proportion z-test + 그룹별 최대편차 선별
  - 수준어 판정          Top2 → 미흡/보통/약간 긍정/긍정/우수
  - 소수점              % 1자리, 평균·NPS 2자리 (ROUND_HALF_UP)

근거는 전부 `engine/learning/RULES.md`. 규칙을 바꾸려면 거기부터 고친다.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any, Iterable

ENGINE = Path(__file__).resolve().parents[1]
CHART_RULES = ENGINE / "knowledge" / "guide" / "chart_rules.yaml"


# ─────────────────────────────────────────────────────────────
# 소수점
# ─────────────────────────────────────────────────────────────

def round_half_up(value: float, digits: int) -> float:
    """엑셀 표시 서식과 같은 반올림.

    파이썬 기본 round() 는 banker's rounding(짝수로 맞춤)이라 0.5 를 내림하는 경우가 있다.
    통계표는 ROUND_HALF_UP 이므로 그대로 쓰면 표시값이 어긋난다.
    """
    q = Decimal(1).scaleb(-digits)
    return float(Decimal(repr(value)).quantize(q, rounding=ROUND_HALF_UP))


def fmt_pct(value: float | None) -> str:
    """% 는 소수 1자리. `.0` 도 유지한다 (40 → '40.0')."""
    if value is None:
        return "-"
    return f"{round_half_up(value, 1):.1f}"


def fmt_mean(value: float | None) -> str:
    """평균·NPS 는 소수 2자리."""
    if value is None:
        return "-"
    return f"{round_half_up(value, 2):.2f}"


# ─────────────────────────────────────────────────────────────
# 수준어 사다리
# ─────────────────────────────────────────────────────────────

#: 보고서 `조사 개요 > 만족도 산출 방법` 슬라이드에 문서로 박혀 있는 해석 기준.
#: "50.0% 미만은 미흡, 50.0%대는 보통, 60%대는 약간 긍정, 70%대는 긍정, 80이상은 우수"
LEVEL_LADDER: tuple[tuple[float, str, str], ...] = (
    (50.0, "미흡", "미흡한 수준임"),
    (60.0, "보통", "보통 수준임"),
    (70.0, "약간 긍정", "약간 긍정적인 수준임"),
    (80.0, "긍정", "긍정적인 수준임"),
    (math.inf, "우수", "우수한 수준임"),
)


def level_word(top_pct: float | None) -> tuple[str | None, str | None]:
    """Top2(또는 Top3) 값 → (형용사, 멘트 어미)."""
    if top_pct is None:
        return None, None
    v = round_half_up(top_pct, 1)
    for upper, word, phrase in LEVEL_LADDER:
        if v < upper:
            return word, phrase
    return LEVEL_LADDER[-1][1], LEVEL_LADDER[-1][2]


# ─────────────────────────────────────────────────────────────
# 행/값 접근
# ─────────────────────────────────────────────────────────────

def cell_value(cell: Any) -> float | None:
    """파서가 낸 셀에서 **서식 적용 후** 값을 꺼낸다.

    원값은 무한소수(13.986013986…)이고 표시 서식이 1자리로 반올림한다.
    보고서에 쓸 숫자는 표시값이어야 합계가 맞는다.
    """
    if cell is None:
        return None
    if isinstance(cell, dict):
        v = cell.get("value")
        if v is None:
            v = cell.get("raw")
        return v
    if isinstance(cell, (int, float)):
        return float(cell)
    return None


def rows_of_kind(table: dict, *kinds: str) -> list[dict]:
    want = set(kinds)
    return [r for r in (table.get("rows") or []) if r.get("kind") in want]


def first_row_of_kind(table: dict, *kinds: str) -> dict | None:
    rows = rows_of_kind(table, *kinds)
    return rows[0] if rows else None


def self_referential_banner_groups(table: dict) -> dict[str, str]:
    """표 안에서 **정보량이 0인 배너군** — {그룹명: 사유}.

    세 유형이 실측된다 (2026 에스원 비상주 90표 · TA Pool 34표 전수 적용):
      ① 대각 100% — 배너 변수가 이 표의 문항 그 자체다 (비상주 표31 B2 × 관리 운영 방식).
         보기 행마다 그 군에 값이 정확히 한 칸이고 그 값이 100.0 이다.
      ② 단일 범주 n = base_n — 서브베이스 조건이 그 군을 한 범주로 접어 전체 열의
         완전한 복제가 됐다 (비상주 표32 B3: 관리 운영 방식 = 「일부 외부 전문업체 위탁」 198 = base).
      ③ 분할 — ①의 일반형이다. **행이 배너군의 하위 분류**여서 열마다 값이 있는 행이
         여럿이지만 그 합이 100.0 이고 열끼리 행 집합이 서로소다.
         TA Pool 표4(A1 소유 형태)가 그렇다 — 행은 사후 오픈코딩 8보기인데 배너는 원본 5보기라
         `기타` 열에 `기관간접투자`·`리츠간접소유`·`법인간접소유`·`법인` 이 14.3/14.3/14.3/57.1 로 실린다.
         ①만 보면 못 잡고, 그 열은 5분류 리코드의 내부를 그대로 노출한다.
    `전체` 군은 원래 1열이므로 대상이 아니다. 판정은 그룹명·값으로만 한다 — 열 위치가 아니다(R-056).
    """
    out: dict[str, str] = {}
    base_n = table.get("base_n")
    groups: dict[str, list[dict]] = {}
    for b in table.get("banners") or []:
        if b.get("in_report") and b.get("group") not in (None, "", "전체"):
            groups.setdefault(b["group"], []).append(b)
    opts = [r for r in (table.get("rows") or []) if r.get("kind") == "option"]
    for g, bs in groups.items():
        if len(bs) == 1 and base_n and (bs[0].get("n") or 0) == base_n:
            out[g] = f"단일 범주 n={base_n} = base — 전체 열 복제"
            continue
        if not opts or len(bs) < 2:
            continue
        # 열별로 값이 있는 행 집합과 그 합을 모은다.
        col_rows: list[set[int]] = []
        col_sums: list[float] = []
        for b in bs:
            rows_here: set[int] = set()
            total = 0.0
            for i, r in enumerate(opts):
                vals = r.get("values") or []
                v = cell_value(vals[b["pos"]]) if b["pos"] < len(vals) else None
                if v is not None:
                    rows_here.add(i)
                    total += v
            col_rows.append(rows_here)
            col_sums.append(total)
        if any(not s for s in col_rows):
            continue                                   # 값이 통째로 빈 열이 있으면 판정하지 않는다
        disjoint = all(col_rows[i].isdisjoint(col_rows[j])
                       for i in range(len(col_rows)) for j in range(i + 1, len(col_rows)))
        if not (disjoint and all(abs(t - 100.0) <= 0.15 for t in col_sums)):
            continue
        if all(len(s) == 1 for s in col_rows):
            out[g] = "대각 100% — 배너 변수가 이 문항 자체"
        else:
            out[g] = "분할 100% — 행이 이 배너군의 하위 분류다 (열마다 합 100, 열끼리 행 서로소)"
    return out


def report_banners(table: dict) -> list[dict]:
    """보고서에 쓰는 배너만. §3.6 절단선은 위치가 아니라 **그룹명**으로 판정된다.

    통계표 배너 열 수는 표마다 다르다(비상주 15~45열) — DP가 base 0 인 범주를 통째로
    빼기 때문이다. 위치로 자르면 표마다 다른 데를 자르게 된다.
    자기참조 군(`self_referential_banner_groups`)은 정보량이 0 이라 함께 뺀다.
    """
    drop = self_referential_banner_groups(table)
    return [b for b in (table.get("banners") or [])
            if b.get("in_report") and b.get("group") not in drop]


def banner_indices(table: dict, banners: Iterable[dict]) -> list[int]:
    return [b["pos"] for b in banners]


def total_banner(table: dict) -> dict | None:
    for b in table.get("banners") or []:
        if b.get("group") == "전체" or b.get("label") == "전체":
            return b
    bs = table.get("banners") or []
    return bs[0] if bs else None


#: 순위 표의 열 머리글. `1순위` 열과 `1+2+3순위`(= 넓은 스코프) 열이 나란히 온다.
RANK_FIRST_RE = re.compile(r"^1\s*순위$")
RANK_SUM_RE = re.compile(r"^1\s*(?:\+\s*\d+\s*)*\+\s*\d+\s*순위$|^1\s*~\s*\d+\s*순위$")


def rank_banners(table: dict) -> tuple[dict | None, dict | None]:
    """(1순위 열, 넓은 순위 스코프 열). 없으면 (None, None)."""
    first = wide = None
    for b in table.get("banners") or []:
        lbl = str(b.get("label") or "").strip()
        if first is None and RANK_FIRST_RE.match(lbl):
            first = b
        elif wide is None and RANK_SUM_RE.match(lbl):
            wide = b
    return first, wide


def banner_groups(table: dict) -> list[tuple[str, list[dict]]]:
    """배너를 헤더 그룹 단위로 묶는다 — `[(그룹명, [배너,…]), …]`, 원순서 유지."""
    out: list[tuple[str, list[dict]]] = []
    for b in table.get("banners") or []:
        g = str(b.get("group") or "")
        if out and out[-1][0] == g:
            out[-1][1].append(b)
        else:
            out.append((g, [b]))
    return out


def _banner_view(table: dict, positions: list[int], chapter: int,
                 chapters: int, groups: list[str], why: str) -> dict:
    """배너 일부만 남긴 표 뷰. 행 값도 같은 위치로 잘라 낸다."""
    src = table.get("banners") or []
    banners = []
    for new_pos, p in enumerate(positions):
        b = dict(src[p])
        b["pos"] = new_pos
        banners.append(b)
    rows = []
    for r in table.get("rows") or []:
        vals = r.get("values") or []
        rows.append({**r, "values": [vals[p] if p < len(vals) else {"value": None, "display": None}
                                     for p in positions]})
    hints = dict(table.get("hints") or {})
    hints["n_banners"] = hints["n_banners_in_report"] = len(banners)
    return {**table, "banners": banners, "rows": rows, "hints": hints,
            "chapter": chapter, "chapter_count": chapters,
            "chapter_banner_groups": groups, "chapter_reason": why}


def split_banner_chapters(table: dict, rules: dict | None = None) -> list[dict]:
    """MAP-050 — 배너가 한 장에 안 들어가면 **배너군 경계**에서 장(章)을 나눈다.

    한 문항이 슬라이드 두 장으로 전개되는 케이스다. 문항이 쪼개지는 게 아니라
    **배너가 쪼개진다** — 차트(전체)는 두 장에 똑같이 들어가고 배너표만 갈린다.
    그래서 두 장의 차트가 동일한 것은 복제 흔적이 아니라 정상이다.

    분리는 표 폭이 물리적으로 안 되기 때문이므로 열 수 임계로 판정한다(R-057).
    다만 다음을 모두 만족할 때만 나눈다 — 없는 신호로 장을 복제하면 중복 장표를 양산한다:
      ① 2단 배너 헤더(그룹행 + 라벨행)가 실재한다
      ② 다열 배너군이 2개 이상이다
      ③ 떼어 낼 마지막 배너군이 충분히 크다
    선두의 1열짜리 군(`전체` · `1순위` · `1+N순위`)은 **양쪽 장에 모두** 싣는다.
    """
    cfg = ((rules or {}).get("chapter_split") or {})
    if cfg.get("enabled") is False:
        return [table]
    limit = int(cfg.get("max_banner_columns") or 22)
    min_tail = int(cfg.get("min_tail_group_columns") or 3)

    banners = table.get("banners") or []
    if len(banners) <= limit:
        return [table]
    if not (table.get("hints") or {}).get("two_tier_banner_header"):
        return [table]

    groups = banner_groups(table)
    if len(groups) < 3 or len(groups[-1][1]) < min_tail:
        return [table]
    if sum(1 for _, bs in groups if len(bs) > 1) < 2:
        return [table]

    # 선두의 1열짜리 군 = 양쪽에 다 싣는 대표 열 (`전체` / `1순위` / `1+N순위`)
    carry_upto = 0
    while carry_upto < len(groups) and len(groups[carry_upto][1]) == 1:
        carry_upto += 1
    if carry_upto == 0 or carry_upto >= len(groups) - 1:
        return [table]

    carry = [b["pos"] for _, bs in groups[:carry_upto] for b in bs]
    head = [b["pos"] for _, bs in groups[carry_upto:-1] for b in bs]
    tail = [b["pos"] for b in groups[-1][1]]
    if not head or not tail:
        return [table]

    names = [g for g, _ in groups]
    why = (f"배너 {len(banners)}열 > 상한 {limit}열 — 마지막 배너군 "
           f"`{names[-1]}`({len(tail)}열)을 떼어 2장으로 전개 (MAP-050)")
    return [
        _banner_view(table, carry + head, 1, 2, names[:carry_upto] + names[carry_upto:-1], why),
        _banner_view(table, carry + tail, 2, 2, names[:carry_upto] + [names[-1]], why),
    ]


def representative_banner(table: dict) -> dict | None:
    """정렬·대표값에 쓰는 '전체' 열.

    ★ 순위 문항에서는 이게 `1순위` 열이 아니라 **넓은 순위 스코프**(1+2+3순위) 열이다
    (chart_rules#option_axis.modes.rank_sum_desc, transcript:28:46).
    `total_banner()` 는 첫 열을 집기 때문에 순위 표에서 1순위를 집어 버린다 —
    라운드1 C2 에서 실제로 그렇게 틀렸다.
    """
    _, wide = rank_banners(table)
    if wide is not None:
        return wide
    return total_banner(table)


# ─────────────────────────────────────────────────────────────
# 정렬
# ─────────────────────────────────────────────────────────────

def sort_option_rows(table: dict, resort: bool = False) -> tuple[list[dict], str]:
    """제작가이드에 실을 보기 행 순서. **통계표가 준 순서를 그대로 쓴다.**

    이게 규칙이다 (knowledge/guide/13 §1-①):
        "실무상 DP가 이미 정렬해서 준다 — 중복응답·순위 표는 전체 값 내림차순,
         단일선택·척도 표는 보기번호 오름차순. 그러니 통계표 순서를 그대로 따르면 된다."

    직접 내림차순으로 재정렬하면 오히려 망가진다. 예: 표7 `A1 지상 층수` 는 층수 분포라
    DP가 1,2,3,4… 오름차순으로 주는데, 값 기준으로 다시 정렬하면 5,4,6,3,7… 로 흐트러진다.
    보기번호가 없는(숫자 리코드) 표라서 "고정 순서" 판별로도 걸러낼 수 없다.

    멘트 1행에서 "가장 높은 것 N개"를 읽을 때만 값 기준 정렬이 필요하고,
    그건 `top_options_by_value()` 가 따로 한다.

    ★ 이건 **가이드 그리드**의 행 순서다. 차트가 보여주는 옵션축은 이것과 다르다 —
      척도는 계열이 5,6,7 → 4,3,2,1 로 가고 누적 100% 는 역순으로 쌓인다.
      차트 축은 `chart_option_axis()` 가 따로 낸다 (RULES R-054).
      둘을 한 함수로 겸하면 척도·누적·순위가 전부 틀린다 — 라운드1 SORT 0/53 의 근인이다.

    ★ `resort=True` 는 **DP 정렬이 실제로 깨진 표에만** 사람이 켜는 예외다.
      기본 전제("DP 가 이미 정렬해서 준다")가 그 표에서 성립하지 않을 때만 쓴다.
      실측 사례 — 에스원 TA Pool 표8 은 `외주관리`(31.4)가 `자체관리`(38.4)보다 위,
      표34 는 `없음`(11.7)이 `대형 피트니스/실내골프장`(14.3)보다 위다.
      구간형 표(`1년 이하` · `15층 미만`)에 켜면 구간 순서가 망가지므로 **자동 판정하지 않는다.**
      정렬은 전체 값 내림차순이고 기타·모름 계열은 최하단으로 내린다 (SORT-030 · SORT-040).
    """
    opts = [r for r in (table.get("rows") or []) if r.get("kind") == "option"]
    if not resort:
        return opts, "통계표 순서 유지 (DP 정렬 결과를 따름)"

    def _key(r: dict) -> tuple[int, float]:
        label = (r.get("label") or "").strip()
        is_etc = any(p.search(label) for p in ETC_OPTION_RES)
        vals = r.get("values") or []
        v = vals[0].get("value") if vals and isinstance(vals[0], dict) else None
        return (1 if is_etc else 0, -(v if isinstance(v, (int, float)) else -1.0))

    return (sorted(opts, key=_key),
            "전체 값 내림차순 · 기타·모름 최하단 (SORT-030·040) — 통계표 정렬이 깨져 사람이 재정렬 지정")


#: 보기번호 접두. 차트 라벨에서만 떼고, 가이드 그리드(label_raw)에서는 떼지 않는다.
#: 라벨에 내재된 번호(`8. 공간형 …`, `A. 소유방식`)는 마침표형이라 여기 걸리지 않는다.
OPTION_NO_PREFIX_RE = re.compile(r"^\s*\d+\)\s*")

#: 값 정렬에서 빼고 최하단에 고정하는 보기 (chart_rules#option_axis.etc_last, SORT-040)
#: `없음`으로 끝난다고 다 기타가 아니다 — `안전사고 발생 이력 없음` 은 실질 보기다.
#: 옮기지 않는 쪽이 안전한 기본값이므로(대개 DP 순서가 맞다) 좁게 잡는다.
ETC_OPTION_RES = tuple(re.compile(p) for p in (
    r"^기타",
    r"모르겠다",
    r"해당\s*없[음다]",
    r"^없[음다]$",
    r"(?:특별|딱히|별다른)\S*.*없[음다]",
    r"상황에\s*따라",
))


def chart_label(label: str | None) -> str:
    """차트에 찍히는 보기 라벨. 보기번호 접두 `N) ` 만 제거한다 (LABEL-010)."""
    return OPTION_NO_PREFIX_RE.sub("", str(label or "")).strip()


def is_etc_option(label: str | None) -> bool:
    """`기타`·`잘 모르겠다`·`해당 없음` 류인가 (SORT-040)."""
    s = chart_label(label)
    return any(r.search(s) for r in ETC_OPTION_RES)


def _banner_value(row: dict, pos: int) -> float | None:
    vals = row.get("values") or []
    return cell_value(vals[pos]) if 0 <= pos < len(vals) else None


def _split_etc(opts: list[dict]) -> tuple[list[dict], list[dict]]:
    main = [r for r in opts if not is_etc_option(r.get("label"))]
    etc = [r for r in opts if is_etc_option(r.get("label"))]
    return main, etc


def _scale_axis(table: dict, opts: list[dict]) -> tuple[list[str], list[dict | None], str] | None:
    """척도 발산형(+-세로막대)의 계열 순서 (SORT-010 / SORT-011).

    카테고리축은 배너이고 **보기가 계열**이다. 계열은 중앙에서 바깥으로 두 덩어리:
      양의 반쪽 = 중립점+1 → 최고점,  음의 반쪽 = 중립점 → 1점
    (7점 → 5,6,7 / 4,3,2,1). 값이 0%라 통계표에서 빠진 점수도 복원한다.
    """
    pts = (table.get("hints") or {}).get("scale_points")
    if not pts or pts < 3:
        return None

    by_no: dict[int, dict] = {}
    if all(r.get("option_no") for r in opts) and opts:
        by_no = {int(r["option_no"]): r for r in opts}
    elif len(opts) == pts:
        by_no = {i + 1: r for i, r in enumerate(opts)}
    else:
        for i, r in enumerate(opts):
            n = r.get("option_no")
            if n:
                by_no[int(n)] = r
    if not by_no:
        return None

    neutral = (pts + 1) // 2
    order = list(range(neutral + 1, pts + 1)) + list(range(neutral, 0, -1))

    labels: list[str] = []
    rows: list[dict | None] = []
    for n in order:
        r = by_no.get(n)
        rows.append(r)
        if r is None:
            labels.append(f"{n}점")          # 0% 라 통계표에서 빠진 척도점 복원
            continue
        txt = chart_label(r.get("label"))
        # 수치형 척도는 라벨이 점수 숫자로 시작한다 ("5", "7매우 그렇다") → "N점"
        labels.append(f"{n}점" if txt.startswith(str(n)) else txt)

    missing = pts - len(by_no)
    why = (f"척도 발산형 — 보기는 계열, 중앙에서 바깥으로 "
           f"{'·'.join(str(n) for n in order[:pts - neutral])} → "
           f"{'·'.join(str(n) for n in order[pts - neutral:])} (SORT-010)")
    if missing > 0:
        why += f" / 통계표에서 빠진 척도점 {missing}개 복원 (SORT-011)"
    return labels, rows, why


def chart_option_axis(table: dict,
                      mode: str | None = None) -> tuple[list[str], list[dict | None], str]:
    """차트가 보여주는 **옵션축** 순서와 그 근거.

    ★ 가이드 그리드의 행 순서(`sort_option_rows`)와는 다른 산출물이다.
      그리드는 통계표 순서를 유지하고, 차트는 타입별로 축을 다시 잡는다.
      권위는 `chart_rules.yaml#option_axis`, 모드는 `decision_table.then.option_axis`.

    반환: (축 라벨, 축 라벨에 대응하는 원본 행 | None, 근거 문자열).
    `None` 은 통계표에 없어 복원한 척도점이다 — 값이 0%다.
    """
    opts = [r for r in (table.get("rows") or []) if r.get("kind") == "option"]
    mode = mode or "table_order"

    if mode == "scale_diverging":
        got = _scale_axis(table, opts)
        if got:
            return got
        mode = "table_order"      # 척도점을 못 세면 조용히 틀리느니 통계표 순서로 둔다

    if mode == "grade_stack_reverse":
        rev = list(reversed(opts))
        return ([chart_label(r.get("label")) for r in rev], rev,
                "등급 사다리 누적 100% — 보기는 계열, 통계표 역순으로 쌓음 (SORT-020)")

    if mode == "rank_sum_desc":
        _, wide = rank_banners(table)
        if wide is not None:
            main, etc = _split_etc(opts)
            ordered = sorted(
                main,
                key=lambda r: -(v if (v := _banner_value(r, wide["pos"])) is not None else -math.inf),
            ) + etc
            why = f"순위 문항 — `{wide.get('label')}` 값 내림차순, 동점은 통계표 원순서 (SORT-030)"
            if etc:
                why += " / 기타·모름 보기 최하단 고정 (SORT-040)"
            return [chart_label(r.get("label")) for r in ordered], ordered, why
        mode = "table_order"

    main, etc = _split_etc(opts)
    why = "통계표 순서 유지 (DP 정렬 결과를 따름)"
    if etc:
        why += " / 기타·모름 보기 최하단 고정 (SORT-040)"
    ordered = main + etc
    return [chart_label(r.get("label")) for r in ordered], ordered, why


#: 보기가 **계열**이 되는 축 모드. 이 두 유형만 표와 그래프의 순서가 갈린다 (SORT-045).
SERIES_AXIS_MODES = ("scale_diverging", "grade_stack_reverse")


def table_row_axis(table: dict,
                   mode: str | None = None) -> tuple[list[str], list[dict], str]:
    """제작본 **배너표**의 데이터 행 순서와 그 근거 (SORT-045).

    ★ 정렬 규칙이 그래프에만 걸리고 표에는 안 걸리던 자리다 (라운드2 SORT 표 13/15).
      갈림길은 딱 하나 — **보기가 카테고리축인가, 계열인가.**

      - 보기 = 카테고리축 (가로막대 · 가로겹침막대 · 순위) → 표 행 순서 = 차트 옵션축 순서.
        값 내림차순(SORT-030)도 기타·모름 최하단(SORT-040)도 표에 **똑같이** 적용된다.
      - 보기 = 계열 (척도 발산 · 등급 사다리 누적 100%) → 표는 통계표 순서를 유지한다.
        그래프는 계열을 쌓는 순서라 표와 같을 이유가 없다.

    라벨 위생(`N) ` 접두 제거, LABEL-010)은 두 경우 모두 적용한다 — 제작본 표에도 접두가 없다.
    가이드 그리드(제작가이드 .xlsm)의 행 순서는 이것과 또 다른 산출물이고,
    거기서는 원문 라벨을 그대로 쓴다 (`sort_option_rows`, R-066).
    """
    opts = [r for r in (table.get("rows") or []) if r.get("kind") == "option"]
    if mode in SERIES_AXIS_MODES:
        return ([chart_label(r.get("label")) for r in opts], opts,
                "보기가 계열인 유형 — 표는 통계표 순서를 유지한다 (SORT-045)")
    _, rows, why = chart_option_axis(table, mode)
    rows = [r for r in rows if r is not None]
    return ([chart_label(r.get("label")) for r in rows], rows,
            why + " / 그래프와 표에 동시 적용 (SORT-045)")


def top_options_by_value(table: dict, limit: int) -> list[dict]:
    """전체 열 값이 큰 순으로 보기 행 N개. **멘트 1행 전용.**

    그리드 순서와는 무관하다 — 표는 통계표 순서를 유지하고, 멘트만 "뭐가 제일 높고"
    순으로 읽는다 (녹취 14:01).
    """
    opts = [r for r in (table.get("rows") or []) if r.get("kind") == "option"]
    tb = total_banner(table)
    pos = tb["pos"] if tb else 0

    def key(r: dict) -> float:
        vals = r.get("values") or []
        v = cell_value(vals[pos]) if pos < len(vals) else None
        return -(v if v is not None else -math.inf)

    return sorted(opts, key=key)[:limit]


# ─────────────────────────────────────────────────────────────
# 배너 유의차
# ─────────────────────────────────────────────────────────────

@dataclass
class Deviation:
    """배너 셀 하나가 전체 대비 얼마나 튀는가."""
    group: str
    label: str
    n: int
    row_label: str
    cell_pct: float
    total_pct: float
    diff: float          # 셀 − 전체 (%p)
    z: float             # 2-proportion z (셀 vs 나머지)
    direction: str       # "높음" | "낮음"

    @property
    def abs_diff(self) -> float:
        return abs(self.diff)


def two_proportion_z(p1: float, n1: int, p2: float, n2: int) -> float:
    """비율 두 개의 차이 검정. p 는 0~100 퍼센트로 받는다."""
    if n1 <= 0 or n2 <= 0:
        return 0.0
    a, b = p1 / 100.0, p2 / 100.0
    pool = (a * n1 + b * n2) / (n1 + n2)
    denom = pool * (1 - pool) * (1 / n1 + 1 / n2)
    if denom <= 0:
        return 0.0
    return (a - b) / math.sqrt(denom)


def row_deviations(
    table: dict,
    row: dict,
    *,
    min_z: float = 2.0,
    min_diff: float = 4.4,
    min_cell_n: int = 30,
) -> list[Deviation]:
    """한 행에서 전체 대비 튀는 배너 셀을 전부 찾아 편차 크기순으로 준다.

    기본 임계는 2σ / n=500 기준 4.4%p (녹취 16:34). 단 서연 말대로 이 기준을 전부
    적용하면 읽을 게 너무 많아진다 — 선별은 `pick_deviations()` 가 한다.

    `min_cell_n` 미만인 배너는 아예 제외한다. n<30 셀의 %는 흔들려서 "튀는 값"이 아니라
    표본 잡음이다.
    """
    tb = total_banner(table)
    if tb is None:
        return []
    values = row.get("values") or []
    total_pct = cell_value(values[tb["pos"]] if tb["pos"] < len(values) else None)
    total_n = tb.get("n") or table.get("base_n") or 0
    if total_pct is None or not total_n:
        return []

    out: list[Deviation] = []
    for b in report_banners(table):
        if b["pos"] == tb["pos"]:
            continue
        n1 = b.get("n") or 0
        if n1 < min_cell_n:
            continue
        p1 = cell_value(values[b["pos"]] if b["pos"] < len(values) else None)
        if p1 is None:
            continue

        n2 = total_n - n1
        if n2 <= 0:
            continue
        # 나머지(전체 − 이 셀)의 비율. 전체는 부분들의 가중평균이므로 역산할 수 있다.
        p2 = (total_pct * total_n - p1 * n1) / n2

        diff = p1 - total_pct
        z = two_proportion_z(p1, n1, p2, n2)
        # R-065: 유의(|z|≥2) **이면서** 실무 하한(4.4%p) 이상인 것만 후보다 — 둘 중 하나로 풀면
        # n 이 큰 배너에서 1%p 차이가 z 만으로 올라온다 (실측: A1 지상 '1' 0.4→1.5, z=2.33).
        if abs(z) < min_z or abs(diff) < min_diff:
            continue

        out.append(Deviation(
            group=b.get("group") or "",
            label=b.get("label") or "",
            n=n1,
            row_label=row.get("label") or "",
            cell_pct=round_half_up(p1, 1),
            total_pct=round_half_up(total_pct, 1),
            diff=round_half_up(diff, 1),
            z=round(z, 2),
            direction="높음" if diff > 0 else "낮음",
        ))

    out.sort(key=lambda d: d.abs_diff, reverse=True)
    return out


def pick_deviations(devs: list[Deviation], *, max_groups: int = 2) -> list[Deviation]:
    """멘트 2행에 실제로 읽어줄 것만 고른다.

    서연의 선별 방식을 그대로 옮긴 것이다 —
        "얘 튀긴 튀죠. 근데 얘가 훨씬 튀잖아요. 그럼 도심 것만 읽으면 돼. 얘는 생략해도 돼"
        (녹취 16:57~17:15)

    즉 **배너 그룹 안에서는 가장 튀는 것 하나만** 읽고, 그다음 **그룹 사이에서도 상위 몇 개만**
    읽는다. 기준을 만족하는 것을 다 읽으면 문장이 감당이 안 된다.
    """
    best_per_group: dict[str, Deviation] = {}
    for d in devs:  # 이미 편차 내림차순
        best_per_group.setdefault(d.group, d)
    ranked = sorted(best_per_group.values(), key=lambda d: d.abs_diff, reverse=True)
    return ranked[:max_groups]


# ─────────────────────────────────────────────────────────────
# Base/Unit
# ─────────────────────────────────────────────────────────────

def base_unit_string(
    table: dict,
    *,
    base_label: str = "전체",
    unit: str = "%",
    multi: bool = False,
    with_mean: bool = False,
) -> str:
    """`[Base: 전체(n=500)/ Unit: %]` 형태. 실측 표기를 그대로 따른다.

    평균 단위는 **표에서 읽는다.** 척도 조사에서만 `점` 이고, 이번 조사처럼 계약기간이면 `개월` 이다
    (에스원 TA Pool 실측: `평균 (개월)` 14장). `점` 을 하드코딩하면 그 20장이 전부 틀린다.
    단위를 알 수 없는 평균 행(`평균` 만 있는 표)은 **아무것도 붙이지 않는다** — 틀린 단위보다 낫다.
    """
    n = table.get("base_n")
    hints = table.get("hints") or {}
    extras = []
    if multi or hints.get("is_multi"):
        extras.append("중복")
    if with_mean or hints.get("has_mean_row"):
        units = [u for u in (hints.get("mean_units") or []) if u]
        if units:
            extras.extend(units)
        elif hints.get("scale_points"):
            extras.append("점")
    unit_part = unit + ((", " + ", ".join(extras)) if extras else "")
    return f"[Base: {base_label}(n={n})/ Unit: {unit_part}]"


#: 이 미만이면 % 를 단독 인용하지 않는다. 표기 규약: 배너 셀 `(13*)` + 장표 범례.
SMALL_BASE = 30


def small_base_warning(table: dict, threshold: int = SMALL_BASE) -> str | None:
    """n<30 이면 %를 그대로 인용하면 안 된다. 표기를 남겨 판단하게 한다.

    **전체 base 만 보면 안 된다.** 배너 셀 하나만 작아도 그 열의 %는 못 쓴다 —
    비상주는 전체가 500 이어도 배너 「분당구」가 n=13 이다.
    LG 실측이 그렇다: 제작본 s9 는 `[Base: 전체(n=100)/ Unit: %] *Too small base(n<30)`
    이고 표의 (Base) 행이 `(23*) (22*) (30) (19*) (6*)` 다 —
    **전체는 100 인데 배너 셀 때문에 범례가 붙었다.**
    """
    ns = [table.get("base_n") or 0]
    ns += [b.get("n") or 0 for b in report_banners(table)]
    if any(0 < n < threshold for n in ns):
        return f"*Too small base(n<{threshold})"
    return None


# ─────────────────────────────────────────────────────────────
# 차트타입
# ─────────────────────────────────────────────────────────────

@dataclass
class ChartDecision:
    chart_type: str
    rule_id: str | None = None
    reason: str = ""
    confidence: str = "확정"       # 확정 | 검증중 | 가설
    modifiers: list[str] = field(default_factory=list)   # 예: 그래프 상단 Text
    unresolved: bool = False
    #: 차트 옵션축 모드 — chart_rules.yaml#option_axis.modes
    option_axis: str = "table_order"
    #: `그래프 상단 Text` 를 붙일 행 종류. 사다리형은 top·bot 쌍이다 (RC-1)
    top_text_kinds: tuple[str, ...] = ("top", "mean")
    #: 그 행을 **어떤 매체**에 싣는가 — `표 상단 행` | `독립 TextBox` (R-A)
    top_text_medium: str = "표 상단 행"


def load_chart_rules(path: Path | None = None) -> dict:
    """`chart_rules.yaml` 을 읽는다. 없으면 내장 폴백을 쓴다."""
    p = path or CHART_RULES
    if not p.exists():
        return {}
    try:
        import yaml  # type: ignore
    except ImportError:
        return {}
    with p.open(encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


#: chart_rules.yaml 이 아직 없거나 어떤 규칙에도 안 걸릴 때의 최소 폴백.
#: 실측 빈도 기준 절대 다수가 기본형이므로 그것으로 떨어뜨리되, `unresolved` 로 표시해
#: 사람이 반드시 확인하게 한다. 조용히 기본형으로 넘기면 틀린 걸 못 잡는다.
FALLBACK_BASIC = "전체 : 가로막대, 배너 : 표"


def _get_path(ctx: dict, path: str) -> Any:
    """`table.has_scale_rows` 같은 점 경로를 컨텍스트에서 꺼낸다."""
    cur: Any = ctx
    for part in path.split("."):
        if isinstance(cur, dict):
            cur = cur.get(part)
        else:
            return None
        if cur is None:
            return None
    return cur


def _match_condition(actual: Any, expected: Any) -> bool:
    """decision_table 의 `when` 항목 하나를 판정한다."""
    if isinstance(expected, dict):
        if "min" in expected and (actual is None or actual < expected["min"]):
            return False
        if "max" in expected and (actual is None or actual > expected["max"]):
            return False
        return True
    if isinstance(expected, list):
        return actual in expected
    return actual == expected


def build_context(table: dict, question: dict | None = None,
                  slide_role: str | None = None, **extra: Any) -> dict:
    """decision_table 이 참조하는 필드를 한 곳에 모은다."""
    h = dict(table.get("hints") or {})
    h.setdefault("n_option_rows",
                 sum(1 for r in (table.get("rows") or []) if r.get("kind") == "option"))
    ctx = {
        "table": {**h, "index": table.get("index"),
                  "title_full": table.get("title_full"),
                  "question_code": table.get("question_code")},
        "question": dict(question or {}),
        "slide_role": slide_role,
    }
    ctx.update(extra)
    return ctx


def decide_chart_type(table: dict, question: dict | None = None,
                      rules: dict | None = None, *,
                      slide_role: str | None = None, **extra: Any) -> ChartDecision:
    """통계표 한 장 → Chart type 행에 적을 문자열.

    판단의 권위는 `engine/knowledge/guide/chart_rules.yaml` 의 `decision_table` 에 있다.
    이 함수는 그 표를 위에서부터 평가하는 실행기일 뿐이다 — 규칙을 바꾸려면 YAML 을 고친다.
    """
    rules = rules if rules is not None else load_chart_rules()
    table_rules = (rules or {}).get("decision_table") or []

    if not table_rules:
        # YAML 을 못 읽으면 판단하지 않는다. 조용히 기본형으로 떨어뜨리면 틀린 걸 못 잡는다.
        return ChartDecision(
            FALLBACK_BASIC, None,
            "chart_rules.yaml 을 읽지 못했다 — 규칙 없이 기본형으로 둠",
            confidence="가설", unresolved=True,
        )

    ctx = build_context(table, question, slide_role, **extra)

    for rule in table_rules:
        when = rule.get("when") or {}   # `when: {}` 은 무조건 매치되는 폴백(CT-900)
        if all(_match_condition(_get_path(ctx, k), v) for k, v in when.items()):
            then = rule.get("then") or {}
            chart = then.get("chart_type") or FALLBACK_BASIC
            if "{slide_role}" in chart and slide_role:
                chart = chart.replace("{slide_role}", slide_role)

            conf = rule.get("confidence") or "확정"
            pts = ctx["table"].get("scale_points")
            k = 3 if pts and pts >= 7 else 2
            tmpl = then.get("graph_top_text_rows") or []
            mods = [str(t).replace("{k}", str(k)).replace("{scale_points}", str(pts or ""))
                    for t in tmpl]
            # 상단 Text 대상 행은 규칙 카드가 정한다 — 코드에 'TOP만' 을 박아 두면
            # 사다리형(CT-050)의 【BOT】 이 통째로 누락된다 (라운드1 TOPTEXT 실패 5건).
            kinds = tuple(kind for token, kind in (("TOP", "top"), ("BOT", "bot"),
                                                   ("평균", "mean"), ("NPS", "mean"))
                          if any(token in str(t) for t in tmpl))
            return ChartDecision(
                chart_type=chart, rule_id=rule.get("id"),
                reason=rule.get("name") or "", confidence=conf,
                modifiers=(["그래프 상단 Text"] if mods else []),
                unresolved=(conf != "확정"),
                option_axis=then.get("option_axis") or "table_order",
                top_text_kinds=kinds or ("top", "mean"),
                top_text_medium=then.get("graph_top_text_medium") or "표 상단 행",
            )

    return ChartDecision(FALLBACK_BASIC, None, "어떤 규칙에도 걸리지 않음",
                         confidence="가설", unresolved=True)


def pair_ranking_tables(tables: list[dict]) -> list[tuple[dict | None, dict]]:
    """순위 표를 (1순위표, 1~N순위표) 짝으로 묶는다.

    순위 문항은 통계표가 두 장이지만 **슬라이드는 한 장**이다. 가로겹침막대의 흐린 막대가
    1~N순위, 진한 막대가 1순위이고, 정렬은 1~N순위 기준이다. 1순위 표에서 가져오는 것은
    전체 열뿐이다 (녹취 28:45~29:04).

    B7-1 처럼 한 문항이 앱/웹으로 갈라져 네 장이 되는 경우가 있으므로,
    문항코드가 아니라 **문항코드 + 매트릭스/분기 라벨** 로 묶어야 섞이지 않는다.
    """
    def group_key(t: dict) -> tuple:
        h = t.get("hints") or {}
        parts = [p for p in (t.get("suffix_parts") or [])
                 if not str(p).startswith("1")]  # 순위 표기만 제거
        return (t.get("question_code"), t.get("matrix_item_no"),
                t.get("matrix_item_label"), tuple(parts))

    groups: dict[tuple, list[dict]] = {}
    for t in tables:
        if (t.get("hints") or {}).get("is_ranking"):
            groups.setdefault(group_key(t), []).append(t)

    pairs: list[tuple[dict | None, dict]] = []
    for key, members in groups.items():
        first = next((m for m in members if (m.get("hints") or {}).get("rank_is_first_only")), None)
        combined = next((m for m in members if not (m.get("hints") or {}).get("rank_is_first_only")), None)
        if combined is None:
            # 1순위만 있고 합산표가 없다 — 짝이 깨진 것이므로 그대로 한 장으로 둔다
            for m in members:
                pairs.append((None, m))
            continue
        pairs.append((first, combined))
    return pairs


# ─────────────────────────────────────────────────────────────
# 한글 조사
# ─────────────────────────────────────────────────────────────

_PARTICLES = {"이": "가", "은": "는", "을": "를", "과": "와", "으로": "로"}


def has_final_consonant(word: str) -> bool | None:
    """마지막 글자에 받침이 있나. 한글이 아니면 None."""
    for ch in reversed(word.strip()):
        if ch.isspace() or ch in "()[]{}'\"”’":
            continue
        code = ord(ch)
        if 0xAC00 <= code <= 0xD7A3:
            return (code - 0xAC00) % 28 != 0
        if ch.isdigit():
            # 숫자는 읽는 소리 기준: 0·1·3·6·7·8 은 받침이 있다
            return ch in "0136789" and ch not in "29"
        return None
    return None


def with_particle(word: str, particle: str) -> str:
    """'서북권' + '이' → '서북권이',  '도심권' + '이' → '도심권이',  '분당구' + '이' → '분당구가'.

    받침 여부를 몰라도(영문·기호로 끝나는 라벨) 문장이 깨지지 않도록,
    판단 불가일 때는 관용 표기 '(이)가' 대신 받침 없는 쪽을 쓴다.
    """
    if particle in _PARTICLES:
        with_final, without_final = particle, _PARTICLES[particle]
    elif particle in _PARTICLES.values():
        without_final = particle
        with_final = next(k for k, v in _PARTICLES.items() if v == particle)
    else:
        return word + particle

    fin = has_final_consonant(word)
    if fin is None:
        fin = False
    return word + (with_final if fin else without_final)


def graph_top_text_rows(table: dict,
                        kinds: Iterable[str] = ("top", "mean")) -> list[dict]:
    """`그래프 상단 Text` 를 붙여야 할 행.

    Top·평균 행 라벨 열에 이 문자열을 적어야 제작팀이 그래프 위에 표시한다.
    **안 적으면 안 나온다** (녹취 33:52 "안 나와 이게 안 나와요").

    어떤 행 종류가 대상인지는 차트 규칙이 정한다 (`ChartDecision.top_text_kinds`).
    척도는 TOP+평균, 등급 사다리(누적 100%)는 TOP+BOT 쌍이다 —
    평균 행이 없는 사다리형에서 BOT 을 빼면 그래프 위 요약이 반쪽이 된다 (RC-1).
    """
    order = {"bot": 0, "soso": 1, "top": 2, "mean": 3}
    rows = rows_of_kind(table, *kinds)
    return sorted(rows, key=lambda r: order.get(r.get("kind"), 9))


#: `그래프 상단 Text` 를 싣는 **매체**. 표기법이 매체마다 다르다 (R-A).
TOP_TEXT_MEDIA = {
    "표 상단 행":   "배너표의 맨 위 두 행. 배너별 값과 함께 싣는다 (그래프가 표 아래)",
    "독립 TextBox": "막대 좌우 끝에 겹쳐 붙이는 텍스트 상자 (BOT 왼쪽 · TOP 오른쪽)",
}

_TOP_KIND_RE = re.compile(r"(TOP|BOT|SoSo)\s*(\d*)")
_MEAN_PTS_RE = re.compile(r"(\d+)\s*점")


def top_text_label(row: dict, medium: str = "표 상단 행") -> str:
    """`그래프 상단 Text` 한 줄의 **표기**. 같은 지표라도 매체가 다르면 표기가 다르다.

    통계표(가이드) 쪽 센티널 라벨은 `【TOP3】` / `[7점 평균]` 인데, 이건 **통계표 표기**다.
    제작본에 실릴 때는 매체를 따른다 (라운드2 TOPTEXT 라벨 5/41 의 근인):

      - 표 상단 행  → `TOP3(%)` · `BOT2(%)` · `7점 평균(점)`  — 단위를 괄호로 붙인다
      - 독립 TextBox → `【TOP2】` · `【BOT2】`                  — 브래킷 표기를 유지하고 수치는 다음 줄

    한 예측 안에서 브래킷이 `【】` 와 `[]` 로 섞이던 것도 여기서 통일된다.
    """
    kind = row.get("kind")
    s = str(row.get("label") or "").strip()
    m = _TOP_KIND_RE.search(s)

    if medium == "독립 TextBox":
        if s.startswith("【"):
            return s
        return f"【{m.group(1)}{m.group(2)}】" if m else s

    if kind in ("top", "bot", "soso") and m:
        return f"{m.group(1)}{m.group(2)}(%)"
    if kind == "mean":
        pts = _MEAN_PTS_RE.search(s)
        if pts:
            return f"{pts.group(1)}점 평균(점)"
        if "NPS" in s:
            return "NPS(점)"
    return s


def top_text_unit(row: dict) -> str:
    return "점" if row.get("kind") == "mean" else "%"


def graph_top_text(table: dict,
                   kinds: Iterable[str] = ("top", "mean"),
                   medium: str = "표 상단 행") -> list[dict]:
    """`그래프 상단 Text` 산출물 — 라벨만이 아니라 **매체·배치·단위·수치**까지 싣는다.

    라벨 배열만 내면 제작팀도 채점기도 "어디에 어떻게 쓰라는 건지" 를 알 수 없고,
    멘트 렌즈는 인용할 수치가 없어 채점 자체가 불가능해진다 (R-B · R-MENT-010).
    """
    tb = representative_banner(table)
    pos = tb["pos"] if tb else 0
    out = []
    for r in graph_top_text_rows(table, kinds):
        v = _banner_value(r, pos)
        unit = top_text_unit(r)
        out.append({
            "label": top_text_label(r, medium),
            "source_label": str(r.get("label") or "").strip(),
            "kind": r.get("kind"),
            "medium": medium,
            "placement": TOP_TEXT_MEDIA.get(medium, ""),
            "unit": unit,
            "banner": (tb or {}).get("label"),
            "value": None if v is None else (fmt_mean(v) if unit == "점" else fmt_pct(v)),
        })
    return out


__all__ = [
    "round_half_up", "fmt_pct", "fmt_mean",
    "LEVEL_LADDER", "level_word",
    "cell_value", "rows_of_kind", "first_row_of_kind",
    "report_banners", "total_banner",
    "sort_option_rows", "top_options_by_value",
    "chart_label", "is_etc_option", "chart_option_axis",
    "SERIES_AXIS_MODES", "table_row_axis",
    "banner_groups", "split_banner_chapters",
    "rank_banners", "representative_banner",
    "Deviation", "two_proportion_z", "row_deviations", "pick_deviations",
    "pair_ranking_tables", "has_final_consonant", "with_particle",
    "base_unit_string", "small_base_warning", "SMALL_BASE",
    "ChartDecision", "load_chart_rules", "build_context", "decide_chart_type",
    "graph_top_text_rows", "graph_top_text", "top_text_label", "top_text_unit",
    "TOP_TEXT_MEDIA",
]
