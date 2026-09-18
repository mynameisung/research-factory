# -*- coding: utf-8 -*-
"""
table_lint.py — 통계표를 받자마자 돌리는 사전 검사 (2026-08-24 신설)

  python engine/tools/table_lint.py <tables.json> [--out 리포트.md] [--quiet]

왜: 통계표는 사람이 만든 산출물이라 **라벨이 틀려도 파서는 경고를 내지 않는다.** 실제로 D4(표63)에서
「신뢰가 가지 않는다」가 0.2 행과 71.6 행에 두 번 찍혀 있었고, 파서 경고는 0건이었다. 산출물이 다 나온
뒤에 사람 눈으로 잡으면 이미 늦다.

이 검사기는 **라벨과 수치를 서로 검산**한다. 척도표는 【BOT2】·【TOP2】가 어느 보기 행의 합인지를 값으로
역산할 수 있으므로, 그렇게 찾아낸 행의 라벨이 실제로 그 방향(부정/긍정)의 말인지 대조한다. 라벨이
뒤바뀌면 여기서 걸린다.

검사 항목
  T-01 보기 라벨 중복        같은 표에 같은 보기 문구가 두 번 (치명)
  T-02 척도 극성 불일치      TOP2 를 이루는 행에 부정 어휘 / BOT2 를 이루는 행에 긍정 어휘 (치명)
  T-03 척도 요약 검산 불일치  BOT2+SoSo+TOP2 ≠ 100, 또는 요약행이 보기 합으로 재현되지 않음 (치명)
  T-04 계 ≠ 100             단일 응답표의 합계 이탈 (경고)
  T-05 보기 수 부족          척도 점수보다 보기 행이 적다 — 응답 0인 보기가 생략된 자리 (참고)
  T-06 평균 검산 불일치      [N점 평균] 이 보기 분포로 재현되지 않음 (경고)
  T-07 저베이스              base < 30 (참고) · 배너 집단 n < 30 개수 (참고)
  T-08 빈 표 / 값 없음       (경고)
"""
from __future__ import annotations
import argparse, io, json, re, sys
from pathlib import Path

POS = re.compile(r"매우|그렇다|간다|있다|만족|높|충분|긍정|우수|필요하다|의향이 있|알고 있|이용하고 있|보유|예$|맞|일치")
NEG = re.compile(r"전혀|않|못|없|불만|낮|부정|아니|모르|미보유|비허용|의향이 없")


def num(v):
    try:
        return float(str(v).replace(",", "").replace("*", "").strip())
    except (TypeError, ValueError):
        return None


def total_col(t: dict) -> int:
    for b in t.get("banners") or []:
        if b.get("group") == "전체":
            return b["pos"]
    return 0


def cell(row: dict, pos: int):
    vals = row.get("values") or []
    if pos >= len(vals) or vals[pos] is None:
        return None
    return num(vals[pos].get("display"))


class Report:
    def __init__(self):
        self.items: list[tuple] = []

    def add(self, code, sev, where, msg):
        self.items.append((code, sev, where, msg))

    def count(self, sev):
        return sum(1 for x in self.items if x[1] == sev)


def lint_table(t: dict, rep: Report) -> None:
    idx = t.get("index")
    code = t.get("question_code") or ""
    where = f"표{idx} {code}".strip()
    rows = t.get("rows") or []
    tp = total_col(t)
    opts = [r for r in rows if r.get("kind") == "option"]
    hints = t.get("hints") or {}

    if not rows or not opts:
        rep.add("T-08", "경고", where, "보기 행이 없습니다")
        return

    # ── T-01 보기 라벨 중복
    seen: dict[str, int] = {}
    for r in opts:
        lab = str(r.get("label") or "").strip()
        if not lab:
            continue
        if lab in seen:
            rep.add("T-01", "치명", where,
                    f"보기 라벨이 두 번 나옵니다: {lab!r} (행 {seen[lab] + 1} · {r['row'] + 1}) "
                    f"— 한쪽 문구가 잘못 찍힌 것입니다")
        seen[lab] = r["row"]

    # ── 척도표 검사
    def sentinel(kind):
        return next((r for r in rows if r.get("kind") == kind), None)
    bot, soso, top = sentinel("bot"), sentinel("soso"), sentinel("top")
    if bot is not None and top is not None:
        b, s_, tv = cell(bot, tp), cell(soso, tp) if soso is not None else None, cell(top, tp)
        vals = [(r, cell(r, tp)) for r in opts]
        vals = [(r, v) for r, v in vals if v is not None]

        # T-03 세 요약이 100 인가
        if None not in (b, tv):
            tot = b + tv + (s_ or 0)
            if abs(tot - 100.0) > 0.6:
                rep.add("T-03", "치명", where,
                        f"BOT2+SoSo+TOP2 = {tot:.1f} 로 100 이 아닙니다 "
                        f"(BOT2 {b} · SoSo {s_} · TOP2 {tv})")

        # 어느 보기 행이 TOP2·BOT2 를 이루는지 값으로 역산한다 (뒤에서부터 / 앞에서부터 누적)
        def prefix_match(seq, target):
            acc = 0.0
            for k, (r, v) in enumerate(seq):
                acc += v
                if target is not None and abs(acc - target) <= 0.15:
                    return [x[0] for x in seq[:k + 1]]
            return None
        bot_rows = prefix_match(vals, b)
        top_rows = prefix_match(list(reversed(vals)), tv)

        # T-02 극성 대조 — 라벨이 그 방향의 말인가
        for rs, want_pos, name in ((top_rows, True, "TOP2"), (bot_rows, False, "BOT2")):
            if not rs:
                continue
            for r in rs:
                lab = str(r.get("label") or "")
                is_neg, is_pos = bool(NEG.search(lab)), bool(POS.search(lab))
                if want_pos and is_neg and not (is_pos and not is_neg):
                    rep.add("T-02", "치명", where,
                            f"{name} 를 이루는 행인데 라벨이 부정형입니다 — 행 {r['row'] + 1} {lab!r} "
                            f"(값 {cell(r, tp)}). 라벨이 뒤바뀐 것으로 보입니다")
                if (not want_pos) and is_pos and not is_neg:
                    rep.add("T-02", "치명", where,
                            f"{name} 를 이루는 행인데 라벨이 긍정형입니다 — 행 {r['row'] + 1} {lab!r} "
                            f"(값 {cell(r, tp)})")
        if b is not None and bot_rows is None:
            rep.add("T-03", "경고", where, f"BOT2 {b} 를 보기 행 합으로 재현하지 못했습니다")
        if tv is not None and top_rows is None:
            rep.add("T-03", "경고", where, f"TOP2 {tv} 를 보기 행 합으로 재현하지 못했습니다")

        # T-05 보기 수
        pts = hints.get("scale_points")
        if pts and len(opts) < pts:
            rep.add("T-05", "참고", where,
                    f"{pts}점 척도인데 보기 행이 {len(opts)}개입니다 — 응답 0인 보기가 생략된 것으로 보입니다")

        # T-06 평균 검산 — 보기 순서를 1..N 점으로 두고 가중평균
        mean_row = sentinel("mean")
        if mean_row is not None and pts and len(opts) == pts:
            m = cell(mean_row, tp)
            tot_pct = sum(v for _, v in vals)
            if m is not None and tot_pct and abs(tot_pct - 100) < 1.5:
                calc = sum((k + 1) * v for k, (_, v) in enumerate(vals)) / tot_pct
                if abs(calc - m) > 0.06:
                    rep.add("T-06", "경고", where,
                            f"[{pts}점 평균] {m} 이 보기 분포로 재현되지 않습니다 (계산 {calc:.2f}) "
                            f"— 보기 순서나 라벨을 확인해야 합니다")

    # ── T-04 계
    tot_row = next((r for r in rows if r.get("kind") == "total"), None)
    if tot_row is not None and not hints.get("is_multi") and not hints.get("is_ranking"):
        v = cell(tot_row, tp)
        if v is not None and abs(v - 100.0) > 0.6:
            rep.add("T-04", "경고", where, f"계가 {v} 입니다 (단일 응답표)")

    # ── T-07 저베이스
    bn = t.get("base_n")
    if bn is not None and bn < 30:
        rep.add("T-07", "참고", where, f"base n={bn} (30 미만)")
    small = [b for b in (t.get("banners") or []) if 0 < (b.get("n") or 0) < 30]
    if small:
        tiny = [b for b in small if (b.get("n") or 0) <= 2]
        rep.add("T-07", "참고", where,
                f"배너 집단 {len(small)}개가 n<30" + (f" (그중 n≤2 가 {len(tiny)}개)" if tiny else ""))


def main() -> int:
    ap = argparse.ArgumentParser(description="통계표 사전 검사")
    ap.add_argument("tables", type=Path)
    ap.add_argument("--out", type=Path)
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args()
    with io.open(a.tables, encoding="utf-8") as fh:
        doc = json.load(fh)
    tables = doc["tables"] if isinstance(doc, dict) else doc
    rep = Report()
    for t in tables:
        lint_table(t, rep)

    order = {"치명": 0, "경고": 1, "참고": 2}
    rep.items.sort(key=lambda x: (order.get(x[1], 9), x[0]))
    head = (f"# 통계표 사전 검사 — {a.tables.name}\n\n"
            f"표 {len(tables)}장 · 치명 {rep.count('치명')} · 경고 {rep.count('경고')} · 참고 {rep.count('참고')}\n")
    lines = [head]
    for c, s, w, m in rep.items:
        lines.append(f"- `{c}` **[{s}]** {w} — {m}")
    text = "\n".join(lines) + "\n"
    if a.out:
        a.out.write_text(text, encoding="utf-8", newline="\n")
        print(f"→ {a.out}")
    if not a.quiet:
        print(head.strip())
        for c, s, w, m in rep.items:
            if s != "참고":
                print(f"  [{s}] {c} {w} — {m}")
    return 1 if rep.count("치명") else 0


if __name__ == "__main__":
    raise SystemExit(main())
