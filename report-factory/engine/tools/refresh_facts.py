#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""refresh_facts.py -- 통계표가 다시 왔을 때 `slide_facts.json` 의 **값만** 갱신한다.

왜 build_facts 를 다시 돌리지 않는가
------------------------------------
`slide_facts.json` 은 `build_facts.py` 산출물 그대로가 아니다. 그 뒤에 사람과
여러 도구가 손을 얹는다 -- 라벨 정정(`표라벨정정.json`: 「0」→「0층」) · 파이핑 치환 ·
선지 번호 제거 · 멘트 확정 · 저베이스 표기 · 목차 페이지 · 요약 설계 · 파생 구간
접기(R-PPT-008: A1 지상 층수 19행을 총 층수 구간 6행으로). 옵션 기록이 남아 있지
않으므로 **재실행은 그 손질을 전부 잃는다.**

그래서 이 도구는 반대로 간다: 기존 정본의 **구조를 진리로 두고** 새 통계표에서
값만 떠서 갈아 끼운다. `build_facts.build_grid` 를 그대로 불러 새 그리드를 만들고,
배너 열과 행 종류(kind) 배열이 기존과 **일치하는 장에만** 값을 옮긴다.

구조가 달라진 장은 손대지 않고 보고한다
----------------------------------------
보기가 추가·삭제되거나 배너 절단이 바뀐 장은 장표 행 수가 달라지는 일이고,
꼬리 접기·표 높이 판단이 얽혀 있어 기계가 대신 정하면 안 된다.
`--allow-structure-change <표번호,...>` 로 장을 하나씩 지정해 통째 교체한다.

사용:
    python engine/tools/refresh_facts.py <slide_facts.json> <새 tables.json>       [--out <새 slide_facts.json>] [--allow-structure-change 37] [--md <리포트>]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent))

import build_facts as BF                                  # noqa: E402


def build_new_grid(slide: dict[str, Any], tables: dict[int, dict]) -> dict | None:
    """이 장의 원표로 새 그리드를 만든다. 기존 그리드가 쓴 옵션을 그대로 준다."""
    src = slide.get("source_tables") or []
    if not src or src[0] not in tables:
        return None
    main = tables[src[0]]
    rank = tables.get(src[1]) if len(src) > 1 else None
    freq = set((slide.get("grid") or {}).get("freq_banners") or []) or None
    return BF.build_grid(main, rank_first=rank, freq_banners=freq)


def structure_diff(old: dict, new: dict) -> list[str]:
    """두 그리드의 **구조** 차이. 값 차이는 보지 않는다."""
    out: list[str] = []
    if old.get("banner_label_row") != new.get("banner_label_row"):
        out.append(f"배너 열 {len(old.get('banner_label_row') or [])}"
                   f" -> {len(new.get('banner_label_row') or [])}")
    ok = [r.get("kind") for r in old.get("rows") or []]
    nk = [r.get("kind") for r in new.get("rows") or []]
    if ok != nk:
        out.append(f"행 구성 {len(ok)} -> {len(nk)}")
    return out


#: 값으로 취급해 갈아 끼우는 필드. 라벨은 **건드리지 않는다**(정정이 얹혀 있다).
VALUE_KEYS = ("cells", "cell_formats", "rank_first_total")


def _norm(s: Any) -> str:
    return re.sub(r"\s+", "", str(s or ""))


def graft_values(old: dict, new: dict) -> dict[str, Any]:
    """값을 옮긴다. **라벨로 붙이고, 새 순서를 따른다.**

    ★ 위치로 붙이면 안 된다. DP 가 보기 순서를 바꿔 보내는 일이 있고(실측: A6
      청구방식과 D10-1 망설임요인에서 두 보기의 자리가 서로 바뀌었다), 그때
      위치로 이식하면 **라벨은 옛 순서·값은 새 순서**가 되어 두 보기의 수치가
      맞바뀐다. 파일은 멀쩡하고 표 채움도도 정상이라 눈으로는 안 보인다
      (수치 검사 D8 이 「행이 통계표와 다릅니다」로 잡았다).

    라벨이 전부 붙으면 새 순서대로 다시 세운다 -- 「통계표 순서를 따른다」가
    정렬 규칙이므로 순서 자체도 정본이다. 라벨이 하나라도 안 붙으면 위치로
    돌아간다: 기존 라벨에 정정이 얹혀 있어 글자가 다른 표가 있다(「0층」 vs 「0」).
    그 경우 순서 변경은 감지할 수 없으므로 `by="pos"` 를 리포트에 남긴다.
    """
    o_rows = old.get("rows") or []
    n_rows_ = new.get("rows") or []
    idx: dict[str, dict] = {}
    for r in o_rows:
        k = _norm(r.get("label_clean") or r.get("label"))
        if k and k not in idx:
            idx[k] = r
    paired: list[tuple[dict, dict]] = []
    for n in n_rows_:
        k = _norm(n.get("label_clean") or n.get("label"))
        hit = idx.get(k)
        if hit is None:
            paired = []
            break
        paired.append((hit, n))

    by = "label"
    if paired and len(paired) == len(o_rows):
        order_changed = [o for o, _ in paired] != o_rows
        rows_out = [o for o, _ in paired]
    else:
        by, order_changed = "pos", False
        paired = list(zip(o_rows, n_rows_))
        rows_out = o_rows

    n_cells = 0
    for o, n in paired:
        for k in VALUE_KEYS:
            if k in n:
                if k == "cells" and o.get("cells") != n.get("cells"):
                    n_cells += sum(1 for a, b in zip(o.get("cells") or [],
                                                     n.get("cells") or []) if a != b)
                o[k] = n[k]
    old["rows"] = rows_out
    old["base_row"] = new.get("base_row", old.get("base_row"))
    for k in ("sort", "n_banners", "rank_overlay", "rank_columns", "freq_banners"):
        if k in new:
            old[k] = new[k]
    return {"rows": len(paired), "cells": n_cells, "by": by,
            "order_changed": order_changed}


BASE_N_RE = re.compile(r"\(n=([\d,]+)\)")


def refresh_base(slide: dict[str, Any], table: dict) -> str | None:
    """Base n 이 바뀌었으면 `base_n` 과 `base_unit` 문자열을 함께 고친다.

    ★ 둘 중 하나만 고치면 표의 (Base) 행과 장표 하단 Base 문구가 어긋난다.
    """
    want = table.get("base_n")
    have = slide.get("base_n")
    if want is None or want == have:
        return None
    slide["base_n"] = want
    bu = slide.get("base_unit") or ""
    if BASE_N_RE.search(bu):
        slide["base_unit"] = BASE_N_RE.sub(f"(n={want})", bu, count=1)
    return f"{have} -> {want}"


def _cell_by_label(table: dict, row_label: str, col_label: str) -> float | None:
    """(행 라벨, 열 라벨)로 표에서 값을 찾는다. 요약 근거 좌표가 쓰는 방식이다."""
    pos = None
    for b in table.get("banners") or []:
        if _norm(b.get("label")) == _norm(col_label):
            pos = b["pos"]
            break
    if pos is None and _norm(col_label) == _norm("전체"):
        pos = 0
    if pos is None:
        return None
    for r in table.get("rows") or []:
        if _norm(r.get("label")) == _norm(row_label) or \
           _norm(r.get("label_raw")) == _norm(row_label):
            vals = r.get("values") or []
            cell = vals[pos] if pos < len(vals) else None
            raw = (cell or {}).get("raw")
            return None if raw is None else float(raw)
    return None


#: 요약 패널의 원표 참조. 「표78」(통계표 번호) · 「P65」(본문 장 순번) 두 꼴이 쓰인다.
REF_TABLE = re.compile(r"^\s*표\s*(\d+)")
REF_SLIDE = re.compile(r"^\s*P\s*(\d+)", re.I)


def _ref_base_n(ref: str, tables: dict[int, dict],
                slide_base: dict[int, int] | None) -> int | None:
    """패널 참조에서 표본 수를 찾는다.

    ★ 참조가 두 꼴이라 하나만 보면 절반을 놓친다 -- 「표78」은 통계표 번호이고
      「P65」는 **본문 장 순번**이다(요약이 본문 장표를 그대로 인용하는 자리).
    """
    m = REF_TABLE.match(ref)
    if m:
        t = tables.get(int(m.group(1)))
        return t.get("base_n") if t else None
    m = REF_SLIDE.match(ref)
    if m and slide_base:
        return slide_base.get(int(m.group(1)))
    return None


def refresh_summary_evidence(summary_path: Path, tables: dict[int, dict],
                             out_path: Path, *,
                             slide_base: dict[int, int] | None = None) -> dict[str, Any]:
    """`summary_map.json` 의 근거값(`claims[].evidence[].value`)을 새 통계표로 맞춘다.

    ★ 요약의 근거값은 **검증의 기준**이다(`verify.py --mode summary` 가 이걸로 대조한다).
      통계표가 바뀌었는데 여기를 안 고치면, 덱을 올바르게 갱신해도 검증이 계속
      「근거 값 불일치」로 잡는다 -- 어느 쪽이 낡았는지 사람이 헷갈린다.
      **문장은 고치지 않는다.** 수치가 든 문장은 뜻이 달라질 수 있어 사람이 본다.
    """
    doc = json.loads(Path(summary_path).read_text(encoding="utf-8"))
    rep: dict[str, Any] = {"changed": [], "missing": [], "sentences": [],
                           "base": []}
    # ── 패널 base_n 과 문장의 「N곳」 표기 ────────────────────────────────
    #    ★ 표본 수는 값이 아니라 **문항의 조건**이다. 근거값만 고치고 여기를 두면
    #      「266곳의 장벽」 같은 문장이 남아 표와 어긋난다(V14 가 치명으로 잡는다).
    #      숫자만 갈아 끼우는 것이라 뜻이 바뀌지 않으므로 기계가 해도 된다.
    base_moves: dict[int, int] = {}
    for sm in doc.get("summaries") or []:
        for p in sm.get("panels") or []:
            if p.get("base_n") is None:
                continue
            want = _ref_base_n(str(p.get("ref") or ""), tables, slide_base)
            if want is not None and p["base_n"] != want:
                base_moves[p["base_n"]] = want
                rep["base"].append({"summary": sm.get("no"),
                                    "table": p.get("ref"),
                                    "was": p["base_n"], "now": want})
                p["base_n"] = want
    if base_moves:
        for sm in doc.get("summaries") or []:
            for c in sm.get("claims") or []:
                t0 = str(c.get("text") or "")
                for was, now in base_moves.items():
                    t0 = re.sub(rf"(?<!\d){was}(?=\s*곳)", str(now), t0)
                if t0 != c.get("text"):
                    rep["base"].append({"summary": sm.get("no"), "table": "문장",
                                        "was": c.get("text"), "now": t0})
                    c["text"] = t0

    for sm in doc.get("summaries") or []:
        for c in sm.get("claims") or []:
            for e in c.get("evidence") or []:
                t = tables.get(e.get("table"))
                if t is None:
                    rep["missing"].append({"summary": sm.get("no"),
                                           "table": e.get("table")})
                    continue
                got = _cell_by_label(t, e.get("row_label") or "",
                                     e.get("col_label") or "")
                if got is None:
                    rep["missing"].append({"summary": sm.get("no"),
                                           "table": e.get("table"),
                                           "row": e.get("row_label"),
                                           "col": e.get("col_label")})
                    continue
                # 자릿수는 **표가 정한다.** 평균은 2자리인데, 라벨만 보면
                # Mean Summary 표(행 라벨이 보기명이고 값이 평균)를 놓친다 --
                # 실측: 표20 의 4.02 가 4.0 으로 깎였다.
                mean_table = (t.get("hints") or {}).get("summary_kind") == "mean"
                digits = 2 if (mean_table or "평균" in str(e.get("row_label") or "")) else 1
                want = f"{BF.R.round_half_up(got, digits):.{digits}f}"
                if str(e.get("value")) != want:
                    rep["changed"].append({"summary": sm.get("no"),
                                           "table": e.get("table"),
                                           "row": e.get("row_label"),
                                           "col": e.get("col_label"),
                                           "was": e.get("value"), "now": want})
                    e["value"] = want
        # 수치가 든 문장은 사람이 본다 -- 목록만 남긴다
        for key in ("headline", "sub_title", "subtitle"):
            if re.search(r"\d", str(sm.get(key) or "")):
                rep["sentences"].append({"summary": sm.get("no"), "where": key,
                                         "text": sm.get(key)})
        for b in sm.get("body") or []:
            if re.search(r"\d", str(b)):
                rep["sentences"].append({"summary": sm.get("no"), "where": "body",
                                         "text": b})
        for c in sm.get("claims") or []:
            if re.search(r"\d", str(c.get("text") or "")):
                rep["sentences"].append({"summary": sm.get("no"), "where": "claim",
                                         "text": c.get("text")})
    Path(out_path).write_text(json.dumps(doc, ensure_ascii=False, indent=1),
                              encoding="utf-8")
    rep["out"] = str(out_path)
    return rep


def refresh(facts_path: Path, tables_path: Path, out_path: Path, *,
            allow: set[int] | None = None) -> dict[str, Any]:
    facts = json.loads(Path(facts_path).read_text(encoding="utf-8"))
    tdoc = json.loads(Path(tables_path).read_text(encoding="utf-8"))
    tables = {t["index"]: t for t in (tdoc["tables"] if isinstance(tdoc, dict) else tdoc)}
    allow = allow or set()

    rep: dict[str, Any] = {"facts": str(facts_path), "tables": str(tables_path),
                           "changed": [], "structure": [], "missing": [],
                           "base": [], "replaced": []}
    for slide in sorted(facts["slides"], key=lambda s: s.get("order") or s.get("seq") or 0):
        new = build_new_grid(slide, tables)
        if new is None:
            rep["missing"].append({"order": slide.get("order"),
                                   "code": slide.get("question_code"),
                                   "src": slide.get("source_tables")})
            continue
        old = slide["grid"]
        diff = structure_diff(old, new)
        src = (slide.get("source_tables") or [None])[0]
        if diff:
            row = {"order": slide.get("order"), "code": slide.get("question_code"),
                   "table": src, "sub": slide.get("sub_title"), "why": diff}
            if src in allow:
                slide["grid"] = new
                rep["replaced"].append(row)
            else:
                rep["structure"].append(row)
                continue
        else:
            got = graft_values(old, new)
            if got["cells"] or got["order_changed"]:
                rep["changed"].append({"order": slide.get("order"),
                                       "code": slide.get("question_code"),
                                       "table": src, "sub": slide.get("sub_title"),
                                       "cells": got["cells"], "by": got["by"],
                                       "order_changed": got["order_changed"]})
            if got["by"] == "pos":
                rep.setdefault("by_position", []).append(
                    {"order": slide.get("order"), "code": slide.get("question_code"),
                     "table": src})
        moved = refresh_base(slide, tables[src])
        if moved:
            rep["base"].append({"order": slide.get("order"),
                                "code": slide.get("question_code"), "n": moved})

    meta = facts.setdefault("meta", {})
    hist = meta.setdefault("refreshed", [])
    hist.append({"tables": str(tables_path),
                 "changed_slides": len(rep["changed"]),
                 "replaced_slides": len(rep["replaced"])})
    Path(out_path).write_text(json.dumps(facts, ensure_ascii=False, indent=1),
                              encoding="utf-8")
    rep["out"] = str(out_path)
    return rep


def render(rep: dict[str, Any]) -> str:
    L = ["# 정본 값 갱신 리포트", "",
         f"- 정본: `{rep['facts']}`",
         f"- 새 통계표: `{rep['tables']}`",
         f"- 값이 바뀐 장 {len(rep['changed'])} · 구조가 바뀐 장 {len(rep['structure'])}"
         f" · 통째 교체 {len(rep['replaced'])} · Base 변경 {len(rep['base'])}", ""]
    if rep["changed"]:
        L += ["## 값이 바뀐 장", "",
              "| 본문 | 문항 | 표 | 바뀐 칸 | 붙인 방식 | 순서 변경 | 꼭지 |",
              "|---|---|---|---|---|---|---|"]
        L += [f"| {c['order']} | {c['code']} | {c['table']} | {c['cells']} |"
              f" {c.get('by','?')} | {'예' if c.get('order_changed') else '-'} | {c['sub']} |"
              for c in rep["changed"]]
        L.append("")
    if rep.get("by_position"):
        L += ["## 라벨이 안 붙어 **위치로** 이식한 장 — 보기 순서가 바뀌었으면 못 잡는다", ""]
        L += [f"- 본문 {b['order']} {b['code']} (표 {b['table']})"
              for b in rep["by_position"]] + [""]
    if rep["base"]:
        L += ["## Base 가 바뀐 장", ""] + [f"- 본문 {b['order']} {b['code']}: {b['n']}"
                                          for b in rep["base"]] + [""]
    if rep["replaced"]:
        L += ["## 그리드를 통째로 교체한 장 (구조 변경 승인)", ""]
        L += [f"- 본문 {r['order']} {r['code']} (표 {r['table']}): {' · '.join(r['why'])}"
              for r in rep["replaced"]] + [""]
    if rep["structure"]:
        L += ["## 구조가 달라 **손대지 않은** 장 — 사람 판단이 필요하다", ""]
        L += [f"- 본문 {r['order']} {r['code']} (표 {r['table']}): {' · '.join(r['why'])}"
              f" — {r['sub']}" for r in rep["structure"]] + [""]
    if rep["missing"]:
        L += ["## 원표를 못 찾은 장", ""] + [f"- 본문 {m['order']} {m['code']}: {m['src']}"
                                            for m in rep["missing"]] + [""]
    return "\n".join(L)


def render_summary(rep: dict[str, Any]) -> str:
    L = ["## 요약 근거값 갱신", "",
         f"- 고친 근거 {len(rep['changed'])}건 · 표본 수 표기 {len(rep.get('base') or [])}건"
         f" · 못 찾은 좌표 {len(rep['missing'])}건"
         f" · 수치가 든 문장 {len(rep['sentences'])}개(사람이 본다)", ""]
    if rep.get("base"):
        L += ["### 표본 수 표기", ""] + [
            f"- 요약 {b['summary']} ({b['table']}): {b['was']} -> {b['now']}"
            for b in rep["base"]] + [""]
    if rep["changed"]:
        L += ["| 요약 | 표 | 행 | 열 | 였던 값 | 새 값 |", "|---|---|---|---|---|---|"]
        L += [f"| {c['summary']} | {c['table']} | {c['row']} | {c['col']} |"
              f" {c['was']} | {c['now']} |" for c in rep["changed"]] + [""]
    if rep["missing"]:
        L += ["### 좌표를 못 찾았다 — 표가 바뀌었는지 사람이 본다", ""]
        L += [f"- 요약 {m['summary']} 표{m.get('table')} "
              f"{m.get('row','')!r}/{m.get('col','')!r}" for m in rep["missing"]] + [""]
    if rep["sentences"]:
        L += ["### 수치가 든 문장 — 값이 바뀌었으면 뜻도 바뀔 수 있다", ""]
        L += [f"- 요약 {s['summary']} [{s['where']}] {s['text']}"
              for s in rep["sentences"]] + [""]
    return "\n".join(L)


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="통계표가 다시 왔을 때 slide_facts.json 의 값만 갱신한다")
    ap.add_argument("facts", help="기존 slide_facts.json")
    ap.add_argument("tables", help="새 tables.json (선지 번호를 먼저 제거한 것)")
    ap.add_argument("--out", help="산출 slide_facts.json (없으면 대조만)")
    ap.add_argument("--allow-structure-change", default="",
                    help="그리드를 통째로 교체할 표 번호. 예 37,44")
    ap.add_argument("--summary", help="summary_map.json — 요약 근거값도 함께 맞춘다")
    ap.add_argument("--md", help="리포트를 마크다운으로")
    args = ap.parse_args(argv)

    allow = {int(x) for x in re.findall(r"\d+", args.allow_structure_change)}
    out = Path(args.out) if args.out else Path(args.facts).with_suffix(".dryrun.json")
    rep = refresh(Path(args.facts), Path(args.tables), out, allow=allow)
    report = render(rep)
    if args.summary:
        tdoc = json.loads(Path(args.tables).read_text(encoding="utf-8"))
        tables = {t["index"]: t for t in (tdoc["tables"] if isinstance(tdoc, dict) else tdoc)}
        sout = Path(args.summary) if args.out else Path(args.summary).with_suffix(".dryrun.json")
        fdoc = json.loads(Path(out if args.out else args.facts).read_text(encoding="utf-8"))
        slide_base = {s0["order"]: s0.get("base_n") for s0 in fdoc.get("slides") or []
                      if s0.get("order") is not None}
        srep = refresh_summary_evidence(Path(args.summary), tables, sout,
                                        slide_base=slide_base)
        report += "\n" + render_summary(srep)
        rep["summary"] = srep
        if not args.out:
            sout.unlink(missing_ok=True)
    print(report)
    if args.md:
        Path(args.md).write_text(report, encoding="utf-8")
        print(f"[리포트] {args.md}")
    if not args.out:
        out.unlink(missing_ok=True)
        print("(--out 이 없어 쓰지 않았다)")
    return 2 if rep["structure"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
