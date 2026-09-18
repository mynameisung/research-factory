#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""webtest_scenarios.py — questionnaire.json → 웹테스트 시나리오 행렬.

연구원이 손으로 짜던 「최소 테스트 시나리오 매트릭스」(행 = 문항, 열 = 시나리오 1~N)를
설문지 정본(`parse_questionnaire.py` 산출)에서 기계적으로 뽑는다. 요구사항 원천은
`docs/requirements-from-call.md`(2026-07-23 연구원 통화)다.

무엇을 시나리오로 만드나
    S0  기본 경로     — 원천 문항마다 가장 많은 딸린 문항을 보이게 하는 답을 골라 끝까지 간다 (파이핑 기대값도 여기서 계산)
    B*  분기 경로     — S0 에서 숨은 문항(서로 배타인 분기의 다른 쪽)을 한 번은 밟는다
    T*  이탈 경로     — 조사 중단(terminate) 보기를 하나씩 골라 그 자리에서 끝나는지 본다
    H*  숨김 경로     — 노출 조건(BASE)을 어기는 답을 골라 딸린 문항이 안 보이는지 본다
                        (상한 --max-branch · 원천 문항마다 하나씩 먼저 · 상한에 걸려 뺀 수는 meta.h_skipped 와 경고로 알린다)
    표 보기의 표시 조건(show_if)이 앞 답으로 안 맞는 보기는 고르지 않는다 — B·T·H 에서 원천 답을 바꾸면 그런 보기를 고른 문항도 다시 고른다.

시나리오 하나 = {id, purpose, answers{문항코드: {choose:[보기번호…], text:"…"}}, expect{visible, hidden, terminate_at, piping}}.
답을 정하지 않은 문항은 순회기(webtest_walk.py)가 첫 보기로 답한다.

사용: python webtest_scenarios.py <questionnaire.json> <out_dir> [--max-branch 12]
산출: <out_dir>/시나리오.json · <out_dir>/시나리오_매트릭스.md
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

RE_OR_GROUP = re.compile(r"\)\s*(?:or|또는|OR)\s*\(")
RE_COND = re.compile(
    r"(?P<code>[A-Z]{1,3}\d+(?:\s*-\s*[A-Za-z0-9]+)?)"      # SQ1 · A2-a · A4 -a
    r"(?:\s+(?P<item>\d+)\))?"                                 # (B1-1 1)≠0 의 항목 번호
    r"\s*(?P<op>≥|≤|≠|!=|<>|>=|<=|=|>|<)\s*"
    r"(?P<vals>\d[\d\)\s,~]*(?:(?:or|또는)\s*\d[\d\)\s,~]*)*)")
RE_COUNT = re.compile(r"(?P<code>[A-Z]{1,3}\d+(?:\s*-\s*[A-Za-z0-9]+)?)\s*(?:≥|>=)\s*(?P<n>\d+)\s*개")
RE_PIPE = re.compile(r"\[#\s*([A-Z]{1,3}\d+(?:-\d+)?)[^#\]]*#\]")
RE_RANGE_NUM = re.compile(r"(\d+)")
RE_SPAN = re.compile(r"(\d+)\s*\)?\s*~\s*(\d+)")


def _vals(s: str) -> set[str]:
    """「3 or 5」「3,4」「1) ~ 7)」「4~5」 → 값 집합."""
    out: set[str] = set()
    for a, b in RE_SPAN.findall(s):
        out.update(str(i) for i in range(int(a), int(b) + 1))
    rest = RE_SPAN.sub(" ", s)
    out.update(RE_RANGE_NUM.findall(rest))
    return out


def parse_condition(s: str | None) -> list[list[tuple[str, str, set[str]]]]:
    """노출 조건 문자열 → OR 묶음 목록(각 묶음은 AND 절 목록). 못 읽으면 빈 목록.

    읽는 표기: SQ1=1 · A2-a ≠ 1) · SQ3=1,2 · B1 = 3 or 5 · D7=4~5 · C4 = 1) ~ 7) · C1 ≥ 1개 이상 ·
              (B1-1 1)≠0) or (B1-1 3)≠0)   ← 복수 문항의 항목 1 이 선택됨 = (B1-1 = 1)
    「|」 뒤의 라벨은 버린다."""
    if not s:
        return []
    s = str(s).split("|")[0]
    groups = RE_OR_GROUP.split(s) if RE_OR_GROUP.search(s) else [s]
    out = []
    for g in groups:
        clauses: list[tuple[str, str, set[str]]] = []
        mc = RE_COUNT.search(g)
        if mc:
            clauses.append((re.sub(r"\s+", "", mc.group("code")), "count>=", {mc.group("n")}))
            g = g[:mc.start()] + g[mc.end():]
        for m in RE_COND.finditer(g):
            code = re.sub(r"\s+", "", m.group("code"))
            op = {"!=": "≠", "<>": "≠", "≥": ">=", "≤": "<="}.get(m.group("op"), m.group("op"))
            vals = _vals(m.group("vals"))
            item = m.group("item")
            if item:
                if op == "≠" and vals == {"0"}:
                    clauses.append((code, "=", {item}))
                elif op == "=" and vals == {"0"}:
                    clauses.append((code, "≠", {item}))
                else:
                    clauses.append((code, op, vals))
            else:
                clauses.append((code, op, vals))
        if clauses:
            out.append(clauses)
    return out


def _chosen(a: dict) -> set[str]:
    chosen = {str(x) for x in (a.get("choose") or [])}
    if a.get("text") is not None and not chosen:
        chosen = {str(a["text"])}
    return chosen


def clause_holds(clause, answers: dict, by_code: dict) -> bool | None:
    code, op, vals = clause
    a = answers.get(code)
    if a is None:
        return None
    chosen = _chosen(a)
    if op == "=":
        return bool(chosen & vals)
    if op == "≠":
        return not (chosen & vals)
    if op == "count>=":
        none_nos = {str(o.get("no")) for o in (by_code.get(code) or {}).get("options") or [] if o.get("is_none")}
        return len(chosen - none_nos) >= int(next(iter(vals)))
    try:  # 수치 비교는 입력값으로만 본다
        x = float(next(iter(chosen))); y = float(next(iter(vals)))
    except Exception:
        return None
    return {">": x > y, "<": x < y, ">=": x >= y, "<=": x <= y}[op]


def cond_holds(groups, answers: dict, by_code: dict | None = None) -> bool | None:
    """OR 묶음 중 하나라도 성립하면 True. 답이 없는 문항이 끼어 판단 못 하면 None."""
    by_code = by_code or {}
    if not groups:
        return True
    unknown = False
    for clauses in groups:
        rs = [clause_holds(c, answers, by_code) for c in clauses]
        if all(r is True for r in rs):
            return True
        if any(r is None for r in rs):
            unknown = True
    return None if unknown else False


def show_if_holds(o: dict, answers: dict | None, by_code: dict | None = None) -> bool | None:
    """보기의 표시 조건(show_if — 표 보기 행의 `[PROG: …]`)이 답으로 성립하나. 조건이 없으면 True, 답이 없거나 조건을 못 읽으면 None."""
    s = o.get("show_if")
    if not s:
        return True
    groups = parse_condition(" · ".join(s) if isinstance(s, list) else str(s)) if answers is not None else []
    return cond_holds(groups, answers, by_code) if groups else None


def option_nos(q: dict) -> list[int]:
    return [int(o.get("no") or i + 1) for i, o in enumerate(q.get("options") or [])]


def default_answer(q: dict, avoid_terminate=True, need=None, forbid=None, prefer=None, answers=None, by_code=None) -> dict:
    """문항 하나의 기본 답. need/forbid 는 노출 조건이 요구하는 보기 번호 집합, prefer 는 {보기번호: 가중치}(높을수록 먼저).
    answers(앞 문항 답)를 주면 표시 조건(show_if)이 안 맞아 웹이 숨길 보기는 맨 뒤로 미룬다."""
    rt = q.get("response_type") or ""
    opts = q.get("options") or []
    if rt in ("숫자입력",):
        rng = " ".join(q.get("directives", {}).get("range", []) or [])
        nums = [int(n) for n in RE_RANGE_NUM.findall(rng)]
        return {"text": str(nums[0] if nums else 30)}
    if rt in ("개방형",) or not opts:
        return {"text": "테스트 응답"}
    cand = []
    for i, o in enumerate(opts):
        no = int(o.get("no") or 0)
        if avoid_terminate and o.get("terminate"):
            continue
        if forbid and str(no) in forbid:
            continue
        if need and str(no) not in need:
            continue
        if o.get("is_none") and not need and len(opts) > 1:
            continue
        cand.append((show_if_holds(o, answers, by_code) is False, -(prefer or {}).get(str(no), 0), i, no))
    if not cand:
        cand = [(show_if_holds(o, answers, by_code) is False, 0, i, int(o.get("no") or 1)) for i, o in enumerate(opts)]
    cand = [no for *_, no in sorted(cand)]
    if rt == "복수":           # 두 개를 고르되 `[최소 N개, 최대 M개 선택]` 이 있으면 그 범위 안에서
        lim = q.get("multi_limit") or {}
        n = max(int(lim.get("min") or 1), min(2, int(lim.get("max") or 2)))
        return {"choose": sorted(cand[:n])}
    if rt == "순위":
        return {"choose": cand[:3]}
    return {"choose": cand[:1]}


def build(qs: list[dict], max_branch: int) -> dict:
    by_code = {q["code"]: q for q in qs}
    conds_of = {q["code"]: parse_condition(q.get("base_condition")) for q in qs}
    unparsed = [q["code"] for q in qs if q.get("base_condition") and not conds_of[q["code"]]]
    # 보기가 없는 문항(숫자 입력·개방형)을 보기 번호로 비교하는 조건은 기계가 판단할 수 없다 — 사람 몫으로 넘긴다
    for q in qs:
        bad = any(c[1] in ("=", "≠") and not (by_code.get(c[0]) or {}).get("options") and c[0] in by_code
                  for grp in conds_of[q["code"]] for c in grp)
        if bad:
            conds_of[q["code"]] = []
            unparsed.append(q["code"])

    def first_group(code: str) -> list:
        return conds_of[code][0] if conds_of[code] else []

    # ── S0 기본 경로: 원천 문항마다 「가장 많은 딸린 문항을 보이게 하는 답」을 고른다
    votes: dict[str, dict[str, int]] = {}
    forbid: dict[str, set] = {}
    for q in qs:
        for code, op, vals in first_group(q["code"]):
            if op == "=":
                for v in vals:
                    votes.setdefault(code, {})[v] = votes.setdefault(code, {}).get(v, 0) + 1
            elif op == "≠":
                forbid.setdefault(code, set()).update(vals)
    s0: dict[str, dict] = {}
    for q in qs:
        s0[q["code"]] = default_answer(q, forbid=forbid.get(q["code"]), prefer=votes.get(q["code"]), answers=s0, by_code=by_code)

    def fit(a: dict, keep: set) -> dict:
        """원천 답을 바꾼 뒤, 표시 조건(show_if)이 안 맞아 웹이 숨길 보기를 고른 문항은 다시 고른다 (keep 문항은 그대로)."""
        for q in qs:
            c = q["code"]
            byno = {int(o.get("no") or 0): o for o in q.get("options") or []}
            if c not in keep and any(show_if_holds(byno.get(int(n)) or {}, a, by_code) is False
                                     for n in (a.get(c) or {}).get("choose") or []):
                a[c] = default_answer(q, forbid=forbid.get(c), prefer=votes.get(c), answers=a, by_code=by_code)
        return a

    def expect_for(answers: dict, terminate_at: str | None) -> dict:
        visible, hidden = [], []
        for q in qs:
            h = cond_holds(conds_of[q["code"]], answers, by_code)
            (visible if h in (True, None) else hidden).append(q["code"])
            if terminate_at and q["code"] == terminate_at:
                break
        return {"visible": visible, "hidden": hidden, "terminate_at": terminate_at, "piping": piping_for(answers, visible)}

    def piping_for(answers: dict, visible: list[str]) -> list[dict]:
        out = []
        for q in qs:
            if q["code"] not in visible:
                continue
            for src in RE_PIPE.findall(q.get("text_full") or q.get("text") or ""):
                a = answers.get(src); sq = by_code.get(src)
                if not a or not sq:
                    continue
                labels = [o.get("text") for o in sq.get("options") or [] if int(o.get("no") or 0) in (a.get("choose") or [])]
                exp = labels[0] if labels else a.get("text")
                if exp:
                    out.append({"code": q["code"], "source": src, "expects_text": exp})
        return out

    def answer_for_clause(code: str, op: str, vals: set, satisfy: bool) -> dict | None:
        """조건 절을 만족(또는 위반)하는 원천 문항의 답. 못 만들면 None."""
        srcq = by_code.get(code)
        if not srcq:
            return None
        if op == "count>=":
            if satisfy:
                return default_answer(srcq)
            none_nos = [int(o.get("no")) for o in srcq.get("options") or [] if o.get("is_none")]
            return {"choose": none_nos[:1]} if none_nos else None
        if op == "=":
            return default_answer(srcq, need=vals) if satisfy else default_answer(srcq, forbid=vals)
        if op == "≠":
            return default_answer(srcq, forbid=vals) if satisfy else default_answer(srcq, need=vals)
        return None

    def fmt(a: dict) -> str:
        return ",".join(str(x) for x in a.get("choose", [])) or str(a.get("text", ""))

    s0_exp = expect_for(s0, None)
    scenarios = [{"id": "S0", "purpose": "기본 경로 — 가장 많은 문항이 보이는 답으로 끝까지 간다", "answers": s0, "expect": s0_exp}]

    # ── B* 분기 경로: S0 에서 숨은 문항을 한 번은 밟는다
    covered = set(s0_exp["visible"]); n_b = 0
    for q in qs:
        if q["code"] in covered or not first_group(q["code"]):
            continue
        a = {k: dict(v) for k, v in s0.items()}
        changed = []
        for code, op, vals in first_group(q["code"]):
            alt = answer_for_clause(code, op, vals, satisfy=True)
            if alt is None:
                break
            a[code] = alt; changed.append(f"{code}={fmt(alt)}")
        else:
            fit(a, {c for c, _, _ in first_group(q["code"])})
            exp = expect_for(a, None)
            if q["code"] in exp["visible"]:
                n_b += 1
                newly = [c for c in exp["visible"] if c not in s0_exp["visible"]]
                covered.update(newly)
                scenarios.append({"id": f"B{n_b}_{q['code']}", "purpose": f"분기 — {' · '.join(changed)} 로 답해 {'·'.join(newly)} 문항이 노출되어야 한다",
                                  "answers": a, "expect": exp})

    # ── T* 이탈 경로
    for q in qs:
        for o in q.get("options") or []:
            if o.get("terminate"):
                a = {k: dict(v) for k, v in s0.items()}
                a[q["code"]] = {"choose": [int(o.get("no") or 1)]}
                fit(a, {q["code"]})
                scenarios.append({"id": f"T_{q['code']}_{o.get('no')}", "purpose": f"이탈 — {q['code']} 보기 {o.get('no')} 「{o.get('text')}」에서 조사 중단",
                                  "answers": a, "expect": expect_for(a, q["code"])})

    # ── H* 숨김 경로: S0 에서 보이는 문항의 노출 조건을 하나씩 깨뜨린다 (상한)
    #    후보를 다 모은 뒤 원천 문항마다 하나씩 먼저 고르고, 남는 자리를 나머지로 채운다. 상한에 걸려 뺀 수는 meta.h_skipped
    seen = set(); cands = []
    for q in qs:
        if q["code"] not in s0_exp["visible"]:
            continue
        for code, op, vals in first_group(q["code"]):
            alt = answer_for_clause(code, op, vals, satisfy=False)
            if alt is None or alt == s0.get(code):
                continue
            key = (code, fmt(alt))              # 같은 답으로 깨뜨리는 조건은 하나로 본다
            if key in seen:
                continue
            seen.add(key)
            a = {k: dict(v) for k, v in s0.items()}
            a[code] = alt
            fit(a, {code})
            exp = expect_for(a, None)
            newly_hidden = [c for c in exp["hidden"] if c not in s0_exp["hidden"]]
            if newly_hidden:
                cands.append((code, op, vals, alt, a, exp, newly_hidden))
    pick: list[int] = []; picked_src: set[str] = set()
    for i, c in enumerate(cands):
        if len(pick) < max_branch and c[0] not in picked_src:
            pick.append(i); picked_src.add(c[0])
    for i in range(len(cands)):
        if len(pick) >= max_branch:
            break
        if i not in pick:
            pick.append(i)
    for n_h, i in enumerate(sorted(pick), 1):
        code, op, vals, alt, a, exp, newly_hidden = cands[i]
        scenarios.append({"id": f"H{n_h}_{code}", "purpose": f"숨김 — {code}={fmt(alt)} 로 답해 조건 「{code}{op if op != 'count>=' else '≥'}{'/'.join(sorted(vals))}」 을 깨뜨린다. {'·'.join(newly_hidden)} 문항이 노출되지 않아야 한다",
                          "answers": a, "expect": exp})

    never = [q["code"] for q in qs if q["code"] not in covered and q["code"] not in unparsed]
    return {"meta": {"n_questions": len(qs), "n_scenarios": len(scenarios), "unparsed_conditions": unparsed, "never_visited": never,
                     "h_candidates": len(cands), "h_skipped": len(cands) - len(pick),
                     "note": "답을 정하지 않은 문항은 순회기가 첫 보기로 답한다. unparsed_conditions 는 사람이 시나리오를 보태야 한다. never_visited 는 어느 시나리오도 밟지 못한 문항이다. "
                             "h_skipped 는 숨김 경로 상한(--max-branch)에 걸려 만들지 않은 수다"},
            "scenarios": scenarios}


def matrix_md(qs: list[dict], doc: dict) -> str:
    sc = doc["scenarios"]
    head = "| 문항 | " + " | ".join(s["id"] for s in sc) + " |"
    lines = ["# 웹테스트 시나리오 매트릭스", "", f"문항 {len(qs)} · 시나리오 {len(sc)}. 칸 = 고를 보기 번호(또는 입력값), — = 노출되면 안 됨, ■ = 여기서 조사 중단.", "",
             "| ID | 목적 |", "|---|---|"] + [f"| {s['id']} | {s['purpose']} |" for s in sc] + ["", head, "|" + "---|" * (len(sc) + 1)]
    for q in qs:
        row = [q["code"]]
        for s in sc:
            e = s["expect"]
            if q["code"] in e["hidden"]:
                row.append("—")
            else:
                a = s["answers"].get(q["code"], {})
                cell = ",".join(str(x) for x in a.get("choose", [])) or a.get("text", "") or "1"
                if cell == "테스트 응답":
                    cell = "텍스트"
                if e.get("terminate_at") == q["code"]:
                    cell = f"■{cell}"
                row.append(cell)
        lines.append("| " + " | ".join(row) + " |")
    if doc["meta"]["unparsed_conditions"]:
        lines += ["", "⚠ 노출 조건을 기계가 못 읽은 문항 (사람이 시나리오를 보탠다): " + " · ".join(doc["meta"]["unparsed_conditions"])]
    if doc["meta"].get("never_visited"):
        lines += ["", "⚠ 어느 시나리오도 밟지 못한 문항 (사람이 시나리오를 보탠다): " + " · ".join(doc["meta"]["never_visited"])]
    if doc["meta"].get("h_skipped"):
        lines += ["", f"⚠ 숨김 경로 후보 {doc['meta']['h_candidates']}개 중 {doc['meta']['h_skipped']}개를 상한(--max-branch)으로 뺐다 — 모두 보려면 --max-branch {doc['meta']['h_candidates']}"]
    return "\n".join(lines) + "\n"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("questionnaire"); ap.add_argument("out_dir")
    ap.add_argument("--max-branch", type=int, default=12, help="숨김 경로 시나리오 상한")
    a = ap.parse_args(argv)
    q = json.loads(Path(a.questionnaire).read_text(encoding="utf-8"))
    qs = q.get("questions") or []
    doc = build(qs, a.max_branch)
    out = Path(a.out_dir); out.mkdir(parents=True, exist_ok=True)
    (out / "시나리오.json").write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    (out / "시나리오_매트릭스.md").write_text(matrix_md(qs, doc), encoding="utf-8")
    print(f"시나리오 {doc['meta']['n_scenarios']}개 (문항 {len(qs)}) → {out}")
    if doc["meta"]["unparsed_conditions"]:
        print("⚠ 못 읽은 노출 조건:", ", ".join(doc["meta"]["unparsed_conditions"]))
    if doc["meta"].get("never_visited"):
        print("⚠ 어느 시나리오도 밟지 못한 문항:", ", ".join(doc["meta"]["never_visited"]))
    if doc["meta"].get("h_skipped"):
        print(f"⚠ 숨김 경로 후보 {doc['meta']['h_candidates']}개 중 {doc['meta']['h_skipped']}개를 상한 --max-branch {a.max_branch} 로 뺐다"
              f" — 모두 만들려면 --max-branch {doc['meta']['h_candidates']}")
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    raise SystemExit(main())
