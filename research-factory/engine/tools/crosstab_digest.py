# -*- coding: utf-8 -*-
"""
crosstab_digest.py — tables.json → 요약 설계용 교차표 요약 패킷 (R-145, 2026-08-24)

  python engine/tools/crosstab_digest.py <work_dir> [--out work/variants/_inputs/crosstab]

guide-summary 가 tables.json(수 MB)을 통째로 읽지 않고 **표 하나씩** 군별 display 값을 보게 한다.
  _index.json   표 번호 · 본문 page · 문항코드 · 제목 · base_n · hints · 군별 집단 수
  표{N}.json     rows[]{label, kind, 전체, <군>:{집단 라벨: display}} · group_n{군:{집단: n}} · group_in_report{군: bool}
값은 통계표 display 문자열 그대로다 — 계산·반올림하지 않는다. `계` 행은 뺀다.
"""
from __future__ import annotations
import argparse, json, os
from pathlib import Path


def digest(tables: list[dict], fact_slides: list[dict], questions: dict | None = None) -> list[dict]:
    page_by = {(s.get("question_code"), (s.get("title_full") or "").strip()): s["page"]
               for s in fact_slides if s.get("page")}
    total_n = max((s.get("base_n") or 0) for s in fact_slides) if fact_slides else 0
    out = []
    for t in tables:
        groups: dict[str, list[dict]] = {}
        for b in t.get("banners") or []:
            groups.setdefault(b["group"], []).append(b)
        rows = []
        for r in t.get("rows") or []:
            if r.get("kind") == "total":
                continue
            vals = r.get("values") or []

            def v(pos: int):
                return vals[pos]["display"] if pos < len(vals) and isinstance(vals[pos], dict) else None
            e = {"label": r.get("label"), "kind": r.get("kind"), "전체": v(0)}
            for g, bs in groups.items():
                if g != "전체":
                    e[g] = {b["label"]: v(b["pos"]) for b in bs}
            rows.append(e)
        out.append({
            "표": t["index"],
            "page": page_by.get((t.get("question_code"), (t.get("title_full") or "").strip())),
            "code": t.get("question_code"), "title": t.get("title_full"), "base_n": t.get("base_n"),
            "hints": {k: (t.get("hints") or {}).get(k) for k in ("is_multi", "has_scale_rows", "is_ranking", "is_summary", "summary_kind")},
            # 이 표를 누가 봤는가 — 앞단 문항의 어느 보기를 고른 사람에게만 물은 것인지 (사용자 지시 2026-08-24)
            "base_condition": ((questions or {}).get(t.get("question_code")) or {}).get("base_condition"),
            "base_is_sub": bool(total_n and (t.get("base_n") or 0) and t["base_n"] < total_n),
            "group_n": {g: {b["label"]: b["n"] for b in bs} for g, bs in groups.items() if g != "전체"},
            "group_in_report": {g: all(b.get("in_report") for b in bs) for g, bs in groups.items() if g != "전체"},
            "rows": rows,
        })
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("work", type=Path)
    ap.add_argument("--out", type=Path)
    a = ap.parse_args()
    W = a.work
    T = json.loads((W / "tables.json").read_text(encoding="utf-8"))
    tables = T["tables"] if isinstance(T, dict) else T
    kpi = W / "KPI요약표_직접생성.json"
    if kpi.exists():
        tables = tables + (json.loads(kpi.read_text(encoding="utf-8")).get("tables") or [])
    facts = W / "slide_facts.json"
    slides = json.loads(facts.read_text(encoding="utf-8")).get("slides", []) if facts.exists() else []
    out_dir = a.out or (W / "variants/_inputs/crosstab")
    os.makedirs(out_dir, exist_ok=True)
    qdoc = json.loads((W / "questionnaire.json").read_text(encoding="utf-8")) if (W / "questionnaire.json").exists() else {}
    qlist = qdoc.get("questions") if isinstance(qdoc, dict) else qdoc
    questions = {q["code"]: q for q in (qlist or []) if q.get("code")}
    d = digest(tables, slides, questions)
    idx = []
    for x in d:
        (out_dir / f"표{x['표']}.json").write_text(json.dumps(x, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        idx.append({k: x[k] for k in ("표", "page", "code", "title", "base_n", "base_condition", "base_is_sub", "hints")}
                   | {"groups": {g: len(v) for g, v in x["group_n"].items()}, "rows": len(x["rows"])})
    (out_dir / "_index.json").write_text(json.dumps(idx, ensure_ascii=False, indent=0), encoding="utf-8", newline="\n")
    print(f"→ {out_dir}  표 {len(idx)}개 (page 매핑 {sum(1 for i in idx if i['page'])}개)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
