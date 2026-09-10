#!/usr/bin/env python3
"""banner_insight.py — 배너(용도·권역·…)에서 **요약에 쓸 만한 인사이트**를 찾는다.

## 무엇을 푸는가

요약 꼭지는 전체값만으로 쓰면 「무난한 사실」만 남는다. 값어치는 **전체와 다르게 움직이는
집단**에 있다. 그런데 배너 표는 한 장이 17열 × 10행이라 눈으로 훑으면 사람마다 다른 것을
집고, 크게 보이는 값을 과대 해석하기 쉽다.

    사용자 지시 (2026-09-01)
    「전체적으로 [2단계] 시설 관리 + 종합 자산관리(PM) 및 임대차까지 도입을 희망하지만
     용도별로 보니 종교는 1단계가 우세하잖아. 이것도 하나의 인사이트고 요약에서 한 줄 써도 돼.
     용도와 권역(배너들)을 기준으로 전체값과 차이가 나는 값을 찾아서 인사이트를 도출해 봐.
     **표 해석을 오버하지 않고, 중요성 기준을 실무적으로 잡아서.**」

이 모듈은 그 「중요성 기준」을 코드로 못박는다. 사람이 표를 보는 순서를 그대로 따른다.

## 세 가지를 찾는다

| 갈래 | 무엇 | 왜 값어치가 있는가 |
|---|---|---|
| **순위 역전** (`flip`) | 그 집단의 1위 보기가 전체의 1위와 **다르다** | 정도의 차이가 아니라 **선택이 다르다**. 한 줄로 쓸 수 있는 가장 강한 발견이다 |
| **큰 편차** (`gap`) | 순위는 같은데 어느 보기의 값이 전체에서 크게 벌어진다 | 같은 방향인데 **세기가 다르다**. 세그먼트 우선순위의 근거가 된다 |
| **단조** (`monotone`) | 순서가 있는 배너(연면적·입주사 수·준공연도)에서 값이 한 방향으로만 간다 | 집단 하나가 아니라 **축 전체가 말을 한다**. 가장 안정적인 신호다 |

## 오버 해석을 막는 다섯 빗장

실무에서 틀리는 자리를 하나씩 막았다. 임계는 전부 이 파일 맨 위 상수에 있다.

1. **모수 하한** — `n < 30` 집단은 아예 보지 않는다. % 가 흔들려 「튀는 값」이 아니라
   표본 잡음이다(R-128 · R-156). `n = 30~49` 는 찾되 **「방향만」** 등급으로 낮춰
   수치 인용을 막는다.
2. **작은 보기 무시** — 전체값이 `MIN_TOTAL_PCT`(5%) 미만인 보기는 건너뛴다.
   1.0 → 3.0 은 3배지만 2%p 이고, 이런 자리에서 나오는 「3배」가 오해의 대표 사례다.
3. **두 관문을 동시에** — 편차는 `MIN_DIFF`(%p)와 `MIN_Z`(2-proportion z)를 **함께**
   넘어야 한다. 큰 모수에서는 작은 차이도 유의해지고, 작은 모수에서는 큰 차이도 우연이다.
4. **역전은 격차를 본다** — 1위가 바뀌었어도 그 집단의 1·2위 차이가 `MIN_FLIP_GAP`(10%p)
   미만이면 **역전이라 부르지 않는다**. 사실상 동률인데 「우세하다」고 쓰면 과장이다.
   (비상주 표86 교육연구시설 37.8 ↔ 34.1 = 3.7%p 가 이 빗장에 걸린다.)
5. **군마다 하나** — 한 배너군에서 가장 센 것 하나만 올린다. 기준을 넘는 것을 다 쓰면
   요약 한 줄에 담기지 않고, 다중비교로 우연히 걸린 것까지 실린다(`pick` 의 규율).

## 쓰는 법

    from banner_insight import find_insights
    ins = find_insights(table, groups=["용도", "권역"])
    for i in ins:
        print(i.sentence())

    # CLI — 표 하나를 훑는다
    python banner_insight.py tables.json --table 86 --groups 용도 권역
    # 요약 정본이 쓰는 표 전부를 훑는다
    python banner_insight.py tables.json --summary-map work/summary_map.json

**이 모듈은 문장을 확정하지 않는다.** 후보와 그 근거 좌표를 낼 뿐이고, 요약에 올릴지는
설계자가 정한다 — 값어치는 통계가 아니라 의뢰인의 물음이 정하기 때문이다.
"""
from __future__ import annotations

import argparse
import io
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import rules as R  # noqa: E402

# ── 중요성 기준 (실무 임계) ──────────────────────────────────────────────
#: 이 아래 모수는 보지 않는다. % 인용 금지선과 같다 (R-128).
MIN_CELL_N = 30
#: 30~49 는 찾되 「방향만」으로 낮춘다 — 수치를 문장에 쓰지 않는다.
WEAK_CELL_N = 50
#: 전체값이 이보다 작은 보기는 건너뛴다. 작은 값의 배수 차이는 발견이 아니다.
MIN_TOTAL_PCT = 5.0
#: 편차의 두 관문. 둘을 **함께** 넘어야 한다.
MIN_DIFF = 10.0          # %p — 실무에서 「눈에 띈다」고 부르는 폭
MIN_Z = 2.0              # 2-proportion z (p<0.05)
STRONG_Z = 2.58          # p<0.01 — 「확실」 등급
#: 순위가 바뀌었어도 이 격차 미만이면 역전이라 부르지 않는다.
MIN_FLIP_GAP = 10.0      # %p
#: 순서가 있는 배너군 — 단조 판정 대상. 이름으로 알아본다.
ORDERED_GROUP_HINTS = ("연면적", "층수", "입주사", "준공", "규모", "면적", "가격", "금액", "연령")


@dataclass
class Insight:
    """요약에 올릴 만한 후보 하나. **문장이 아니라 근거다.**"""
    kind: str                # flip | gap | monotone
    group: str               # 배너군 이름 (용도 · 권역 …)
    label: str               # 집단 이름 (종교시설 …)
    n: int
    row_label: str           # 그 보기(행)
    cell: float
    total: float
    diff: float              # 집단 − 전체 (%p)
    z: float
    strength: str            # 확실 | 참고 | 방향만
    detail: str = ""
    evidence: list[dict] = field(default_factory=list)

    @property
    def score(self) -> float:
        """정렬용. 역전을 가장 앞에 두고, 그다음은 편차 크기다."""
        base = {"flip": 1000.0, "monotone": 500.0, "gap": 0.0}[self.kind]
        return base + abs(self.diff)

    def sentence(self) -> str:
        """설계자가 읽을 한 줄. **요약 문장이 아니다** — 후보의 요지다."""
        num = "" if self.strength == "방향만" else f" ({self.cell:.1f}% ↔ 전체 {self.total:.1f}%)"
        return (f"[{self.kind}·{self.strength}] {self.group} 「{self.label}」(n={self.n}) "
                f"— {self.detail}{num}")


def _pct(table: dict, row: dict, pos: int) -> float | None:
    vs = row.get("values") or []
    return R.cell_value(vs[pos]) if pos < len(vs) else None


def _options(table: dict) -> list[dict]:
    return [r for r in (table.get("rows") or []) if r.get("kind") == "option"]


def _strength(n: int, z: float) -> str:
    if n < WEAK_CELL_N:
        return "방향만"
    return "확실" if abs(z) >= STRONG_Z else "참고"


def find_insights(table: dict, groups: list[str] | None = None,
                  *, min_diff: float = MIN_DIFF, min_z: float = MIN_Z) -> list[Insight]:
    """표 한 장에서 배너 인사이트 후보를 찾는다. 군마다 가장 센 것 하나씩만 돌려준다."""
    tb = R.total_banner(table)
    opts = _options(table)
    if tb is None or not opts:
        return []
    total_n = tb.get("n") or table.get("base_n") or 0
    if not total_n:
        return []

    # 전체 열의 보기 순위 — 역전 판정의 기준선이다.
    tot = [(r, _pct(table, r, tb["pos"])) for r in opts]
    tot = [(r, v) for r, v in tot if v is not None]
    if not tot:
        return []
    tot_sorted = sorted(tot, key=lambda x: -x[1])
    tot_top_label = str(tot_sorted[0][0].get("label") or "").strip()
    tot_by_label = {str(r.get("label") or "").strip(): v for r, v in tot}

    banners = [b for b in R.report_banners(table)
               if b["pos"] != tb["pos"] and (b.get("n") or 0) >= MIN_CELL_N
               and (groups is None or b.get("group") in groups)]

    found: list[Insight] = []
    for b in banners:
        n1 = b.get("n") or 0
        n2 = total_n - n1
        if n2 <= 0:
            continue
        cells = [(str(r.get("label") or "").strip(), _pct(table, r, b["pos"])) for r in opts]
        cells = [(lab, v) for lab, v in cells if v is not None]
        if not cells:
            continue
        ranked = sorted(cells, key=lambda x: -x[1])
        top_label, top_v = ranked[0]
        second_v = ranked[1][1] if len(ranked) > 1 else 0.0

        def dev(label: str, v: float) -> tuple[float, float]:
            """(전체 대비 %p, z). 전체는 부분들의 가중평균이므로 나머지를 역산해 검정한다."""
            tv = tot_by_label.get(label)
            if tv is None:
                return 0.0, 0.0
            p2 = (tv * total_n - v * n1) / n2
            return v - tv, R.two_proportion_z(v, n1, p2, n2)

        # ① 순위 역전 — 1위가 다르고, 그 집단 안에서 1·2위가 충분히 벌어져 있을 때만.
        if top_label != tot_top_label and (top_v - second_v) >= MIN_FLIP_GAP:
            d, z = dev(top_label, top_v)
            if abs(z) >= min_z:
                found.append(Insight(
                    "flip", b.get("group") or "", b["label"], n1, top_label, top_v,
                    tot_by_label.get(top_label, 0.0), d, z, _strength(n1, z),
                    detail=(f"전체 1위는 「{tot_top_label}」인데 이 집단은 「{top_label}」이 "
                            f"{top_v - second_v:.1f}%p 앞선다"),
                    evidence=[{"table": table.get("index"), "row_label": top_label,
                               "col_label": b["label"], "value": f"{top_v:.1f}"},
                              {"table": table.get("index"), "row_label": tot_top_label,
                               "col_label": "전체", "value": f"{tot_by_label.get(tot_top_label, 0):.1f}"}]))
                continue                      # 역전을 찾았으면 그 집단은 그것으로 족하다

        # ② 큰 편차 — 두 관문을 함께 넘고, 전체값이 너무 작은 보기는 보지 않는다.
        best: Insight | None = None
        for label, v in cells:
            tv = tot_by_label.get(label)
            if tv is None or tv < MIN_TOTAL_PCT:
                continue
            d, z = dev(label, v)
            if abs(d) < min_diff or abs(z) < min_z:
                continue
            cand = Insight("gap", b.get("group") or "", b["label"], n1, label, v, tv, d, z,
                           _strength(n1, z),
                           detail=f"「{label}」이 전체보다 {abs(d):.1f}%p {'높다' if d > 0 else '낮다'}",
                           evidence=[{"table": table.get("index"), "row_label": label,
                                      "col_label": b["label"], "value": f"{v:.1f}"},
                                     {"table": table.get("index"), "row_label": label,
                                      "col_label": "전체", "value": f"{tv:.1f}"}])
            if best is None or abs(cand.diff) > abs(best.diff):
                best = cand
        if best is not None:
            found.append(best)

    # ③ 단조 — 순서 있는 배너군은 축 전체가 한 방향인지 본다.
    found += _monotone(table, tb, opts, tot_by_label, groups)

    # 군마다 가장 센 것 하나만 — 다만 **역전은 전부 남긴다.**
    #
    # `gap` 은 정도의 차이라서 한 군에서 여럿 실으면 같은 이야기를 되풀이하게 되고,
    # 다중비교로 우연히 걸린 것까지 올라온다(`pick_deviations` 의 규율).
    # `flip` 은 다르다 — 드물고, **여러 집단이 저마다 다른 보기를 1위로 꼽는다는 사실
    # 자체가 발견**이다. 비상주 표86 권역이 그렇다: 서남권은 1단계, 동남권은 3단계,
    # 나머지는 2단계가 1위다. 군당 하나로 줄이면 그 그림이 통째로 사라진다.
    out = [i for i in found if i.kind == "flip"]
    best_per_group: dict[str, Insight] = {}
    for i in sorted((x for x in found if x.kind != "flip"), key=lambda x: -x.score):
        best_per_group.setdefault(f"{i.kind}:{i.group}", i)
    return sorted(out + list(best_per_group.values()), key=lambda x: -x.score)


def _monotone(table: dict, tb: dict, opts: list[dict],
              tot_by_label: dict[str, float], groups: list[str] | None) -> list[Insight]:
    """순서 있는 배너군에서 값이 한 방향으로만 가는 보기를 찾는다.

    집단 하나가 튀는 것보다 **축 전체가 기우는 것**이 안정적이다 — 표본 잡음은 단조를
    만들지 못한다. 양 끝 차이가 `MIN_DIFF` 를 넘고 집단이 셋 이상일 때만 본다.
    """
    out: list[Insight] = []
    total_n = tb.get("n") or table.get("base_n") or 0
    for gname, bs in R.banner_groups(table):
        if gname in ("", "전체") or (groups is not None and gname not in groups):
            continue
        if not any(h in gname for h in ORDERED_GROUP_HINTS):
            continue
        cols = [b for b in bs if b.get("in_report") and (b.get("n") or 0) >= MIN_CELL_N]
        if len(cols) < 3:
            continue
        for r in opts:
            label = str(r.get("label") or "").strip()
            tv = tot_by_label.get(label)
            if tv is None or tv < MIN_TOTAL_PCT:
                continue
            vs = [_pct(table, r, b["pos"]) for b in cols]
            if any(v is None for v in vs):
                continue
            up = all(vs[i] <= vs[i + 1] for i in range(len(vs) - 1))
            down = all(vs[i] >= vs[i + 1] for i in range(len(vs) - 1))
            span = vs[-1] - vs[0]
            if not (up or down) or abs(span) < MIN_DIFF:
                continue
            n1 = cols[-1].get("n") or 0
            n2 = max(total_n - n1, 1)
            p2 = (tv * total_n - vs[-1] * n1) / n2
            z = R.two_proportion_z(vs[-1], n1, p2, n2)
            out.append(Insight(
                "monotone", gname, f"{cols[0]['label']} → {cols[-1]['label']}",
                min(c.get("n") or 0 for c in cols), label, vs[-1], tv, span, z,
                _strength(min(c.get("n") or 0 for c in cols), z),
                detail=(f"「{label}」이 {cols[0]['label']} {vs[0]:.1f}% 에서 "
                        f"{cols[-1]['label']} {vs[-1]:.1f}% 로 {'올라간다' if up else '내려간다'}"),
                evidence=[{"table": table.get("index"), "row_label": label,
                           "col_label": c["label"], "value": f"{v:.1f}"}
                          for c, v in zip(cols, vs)]))
    return out


# ── CLI ───────────────────────────────────────────────────────────────────

def _load_tables(path: Path) -> dict[int, dict]:
    with io.open(path, encoding="utf-8") as fh:
        doc = json.load(fh)
    out = {t["index"]: t for t in (doc["tables"] if isinstance(doc, dict) else doc)}
    kpi = path.parent / "KPI요약표_직접생성.json"
    if kpi.exists():
        with io.open(kpi, encoding="utf-8") as fh:
            for t in json.load(fh).get("tables", []):
                out[t["index"]] = t
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="배너에서 요약에 쓸 인사이트 후보를 찾는다")
    ap.add_argument("tables", type=Path, help="tables.json")
    ap.add_argument("--table", type=int, action="append", help="표 인덱스 (여러 번 가능)")
    ap.add_argument("--summary-map", type=Path, help="summary_map.json — 요약이 쓰는 표를 전부 훑는다")
    ap.add_argument("--groups", nargs="*", default=None, help="볼 배너군 (기본: 전부)")
    ap.add_argument("--min-diff", type=float, default=MIN_DIFF)
    ap.add_argument("--min-z", type=float, default=MIN_Z)
    ap.add_argument("--json-out", type=Path)
    args = ap.parse_args()

    by_index = _load_tables(args.tables)
    want: list[int] = list(args.table or [])
    if args.summary_map and args.summary_map.exists():
        with io.open(args.summary_map, encoding="utf-8") as fh:
            smap = json.load(fh)
        for s in smap.get("summaries", []):
            for p in s.get("panels") or []:
                ref = str(p.get("ref") or "")
                if ref.startswith("표") and ref[1:].isdigit():
                    want.append(int(ref[1:]))
            for c in s.get("claims") or []:
                for e in c.get("evidence") or []:
                    if e.get("table"):
                        want.append(int(e["table"]))
    if not want:
        want = sorted(by_index)
    seen, order = set(), []
    for i in want:
        if i not in seen:
            seen.add(i)
            order.append(i)

    all_out = []
    for idx in order:
        t = by_index.get(idx)
        if t is None:
            print(f"⚠ 표{idx} 없음")
            continue
        ins = find_insights(t, args.groups, min_diff=args.min_diff, min_z=args.min_z)
        if not ins:
            continue
        print(f"\n=== 표{idx} {t.get('title_full')} (base {t.get('base_n')})")
        for i in ins:
            print("   " + i.sentence())
            all_out.append({**i.__dict__, "table": idx})
    print(f"\n후보 {len(all_out)}건 — 요약에 올릴지는 설계자가 정한다")
    if args.json_out:
        with io.open(args.json_out, "w", encoding="utf-8") as fh:
            json.dump(all_out, fh, ensure_ascii=False, indent=1)
        print(f"[OK] {args.json_out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
