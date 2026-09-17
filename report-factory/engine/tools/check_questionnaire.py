#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""설문 명세 검수 — 구조화 대비 누락 · 로직 참조 · 하우스 규칙 · 파서 왕복.

    python check_questionnaire.py <설문명세.json> [--structure structure.json] [--docx 설문지.docx]
                                  [--out 검수.json] [--md 검수.md]

검사 (코드 = findings 의 check)
    Q-01 치명  문항번호 형식·중복·파트 접두 불일치
    Q-02 경고  파트 안 번호 순서 역행
    Q-03 치명  BASE·파이핑이 없는 문항 / 뒤에 나오는 문항을 가리킴   (PROG 는 없는 문항만)
    Q-04 치명  BASE 의 보기 번호가 그 문항 보기 수를 넘음           (`A3=1~3`, `B2≠4`, `SQ5=2,3`)
    Q-05 치명  응답유형별 필수 요소 누락 (선택형 보기 2개 미만 · 척도 라벨 · 행렬 행 · 순위 수 > 보기 수)
    Q-06 경고  기타 보기가 끝이 아님(없음·모름 앞은 허용) · 기타 기입란 없음
    Q-07 경고  복수 문항의 없음/모름/경험 없음 보기에 배타(exclusive) 표시 없음
    Q-08 경고  척도 방향 — 1점 라벨이 긍정으로 보임 (하우스 규칙: 1=부정 → 5=긍정)
    Q-09 경고  숫자 입력인데 RANGE 없음
    Q-10 경고  보기 5개 이상 명목형인데 로테이션 결정이 없음 (rotation: true/false 를 명시)
    Q-11 경고  구조화 행이 어떤 문항의 src 에도 없음 (구조화 누락)          --structure
    Q-12 정보  구조화에 없는 문항 (src 없음) · 확인 필요(review) 문항 목록
    Q-13 경고  SQ 파트에 조사 중단 장치가 없음 (보기 terminate 또는 PROG 조사 종료)
    Q-14 치명  파서 왕복 — docx 를 parse_questionnaire 로 읽은 문항번호·응답유형·보기 수가 명세와 다름   --docx
    Q-15 경고  파서 경고
    Q-16 경고  PROG·BASE·RANGE 안에 대괄호가 겹침 (파서가 첫 `]` 에서 자름)
판정: 치명 0 이면 통과(exit 0). 치명이 있으면 exit 1.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from build_questionnaire import RE_NONE, _as_list, load_spec, norm_options, scale_labels  # noqa: E402

RE_CODE = re.compile(r"^(SQ|DQ|[A-Z]{1,2})(\d+)((?:-\d+)*)$")
RE_REF = re.compile(r"(?<![A-Za-z#])((?:SQ|DQ|[A-Z]{1,2})\d+(?:-\d+)*)(?![\d])")
RE_COND = re.compile(r"((?:SQ|DQ|[A-Z]{1,2})\d+(?:-\d+)*)\s*(=|≠|!=)\s*([\d~,\s]+(?:or\s*[\d~,\s]+)*)")
EXPECT_TYPE = {"single": "단일", "multi": "복수", "rank": "순위", "scale": "척도", "matrix": "행렬",
               "number": "숫자입력", "open": "개방형", "yesno": "예-아니오"}
POSITIVE_FIRST = re.compile(r"^(매우\s*(만족|그렇|중요|좋)|아주)")


def _key(code: str) -> tuple:
    m = RE_CODE.match(code)
    if not m:
        return (code,)
    return (m.group(1), int(m.group(2))) + tuple(int(x) for x in m.group(3).split("-") if x)


def _nums(expr: str) -> list[int]:
    out: list[int] = []
    for part in re.split(r"[,\s]|or", expr):
        part = part.strip()
        if not part:
            continue
        if "~" in part:
            a, _, b = part.partition("~")
            if a.strip().isdigit() and b.strip().isdigit():
                out += list(range(int(a), int(b) + 1))
        elif part.isdigit():
            out.append(int(part))
    return out


def n_answers(it: dict[str, Any]) -> int | None:
    kind = it.get("kind")
    if kind in ("single", "multi", "rank"):
        if it.get("options_table"):
            return sum(len(g["options"]) for g in it["options_table"])
        return len(norm_options(it.get("options") or []))
    if kind == "yesno":
        return len(it.get("options") or ["예", "아니오"])
    if kind == "list":
        return len(it["options"]) if it.get("options") else None
    if kind == "scale":
        return len(scale_labels(it)) + len(_as_list((it.get("scale") or {}).get("extra")))
    return None


def check(spec: dict[str, Any], structure: dict[str, Any] | None, docx_path: Path | None) -> list[dict[str, Any]]:
    F: list[dict[str, Any]] = []

    def add(check_id, sev, where, msg):
        F.append({"check": check_id, "severity": sev, "where": where, "message": msg})

    items: list[tuple[str, dict[str, Any]]] = []
    for pt in spec["parts"]:
        for it in pt.get("items", []):
            if it.get("code"):
                items.append((pt["code"], it))
    order = {it["code"]: i for i, (_, it) in enumerate(items)}
    by_code = {it["code"]: it for _, it in items}

    # Q-01 / Q-02
    seen: set[str] = set()
    last: dict[str, tuple] = {}
    for part, it in items:
        code = it["code"]
        m = RE_CODE.match(code)
        if not m:
            add("Q-01", "치명", code, "문항번호 형식이 `A1` · `A1-1` · `SQ3` 꼴이 아님 (parse_questionnaire 가 못 읽음)")
            continue
        if code in seen:
            add("Q-01", "치명", code, "문항번호 중복")
        seen.add(code)
        if m.group(1) != part:
            add("Q-01", "치명", code, f"파트 {part} 안에 있는데 접두가 {m.group(1)} — 파서가 문항으로 못 읽음")
        k = _key(code)
        if part in last and k < last[part]:
            add("Q-02", "경고", code, "앞 문항보다 번호가 작음 — 순서 확인")
        last[part] = k

    # Q-03 / Q-04
    for part, it in items:
        code = it["code"]
        refs_strict = []
        for field in ("base",):
            refs_strict += [(field, r) for r in RE_REF.findall(str(it.get(field) or ""))]
        for pip in re.findall(r"\[#([^\]]*)\]", str(it.get("text", "")) + str(it.get("note", ""))):
            refs_strict += [("파이핑", r) for r in RE_REF.findall(pip)]
        for field, r in refs_strict:
            if r not in by_code:
                add("Q-03", "치명", code, f"{field} 가 없는 문항 {r} 를 가리킴")
            elif order[r] >= order[code]:
                add("Q-03", "치명", code, f"{field} 가 뒤에 나오는 문항 {r} 를 가리킴")
        for pg in _as_list(it.get("prog")) + _as_list(it.get("prog_before")):
            for r in RE_REF.findall(str(pg)):
                if r not in by_code and not r.startswith(("SQ0",)):
                    add("Q-03", "경고", code, f"PROG 가 없는 문항 {r} 를 가리킴: {pg}")
        for q, op, expr in RE_COND.findall(str(it.get("base") or "")):
            tgt = by_code.get(q)
            if not tgt:
                continue
            n = n_answers(tgt)
            if n is None:
                continue
            bad = [x for x in _nums(expr) if x < 1 or x > n]
            if bad:
                add("Q-04", "치명", code, f"BASE `{q}{op}{expr.strip()}` — {q} 의 응답 코드는 1~{n} 인데 {bad} 를 씀")

    # Q-05 ~ Q-10
    for part, it in items:
        code, kind = it["code"], it.get("kind", "single")
        if kind not in EXPECT_TYPE and kind != "list":
            add("Q-05", "치명", code, f"알 수 없는 kind `{kind}`")
            continue
        opts = norm_options(it.get("options") or []) if kind in ("single", "multi", "rank") else []
        if kind in ("single", "multi", "rank") and not it.get("options_table") and len(opts) < 2:
            add("Q-05", "치명", code, "선택형인데 보기가 2개 미만")
        if kind == "scale":
            try:
                labs = scale_labels(it)
            except Exception as e:
                add("Q-05", "치명", code, str(e)); labs = []
            if labs and POSITIVE_FIRST.match(labs[0]):
                add("Q-08", "경고", code, f"1점 라벨이 긍정으로 보임: `{labs[0]}` — 1=부정 → 5=긍정이 하우스 규칙")
        if kind == "matrix":
            if not it.get("rows"):
                add("Q-05", "치명", code, "행렬인데 rows 가 없음")
            if not it.get("scale") and not it.get("columns"):
                add("Q-05", "치명", code, "행렬인데 scale 도 columns 도 없음")
            if it.get("scale"):
                try:
                    labs = scale_labels(it)
                    if POSITIVE_FIRST.match(labs[0]):
                        add("Q-08", "경고", code, f"1점 라벨이 긍정으로 보임: `{labs[0]}`")
                except Exception as e:
                    add("Q-05", "치명", code, str(e))
        if kind == "rank":
            rk = it.get("rank") or {}
            mx = int(rk.get("optional") or rk.get("required") or 1)
            n = n_answers(it) or 0
            if mx > n:
                add("Q-05", "치명", code, f"순위 {mx}개를 받는데 보기가 {n}개")
        if kind == "multi" and it.get("max") and int(it["max"]) > len(opts) and not it.get("options_table"):
            add("Q-05", "치명", code, f"최대 {it['max']}개 선택인데 보기가 {len(opts)}개")
        if opts:
            others = [i for i, o in enumerate(opts) if o.get("other")]
            for i in others:
                tail = opts[i + 1:]
                if any(not (o.get("exclusive") or RE_NONE.search(o["text"])) for o in tail):
                    add("Q-06", "경고", code, f"기타 보기 {opts[i]['no']}) 가 끝이 아님")
                if "(" not in opts[i]["text"]:
                    add("Q-06", "경고", code, "기타 보기에 기입란 `(      )` 없음")
            if kind == "multi":
                for o in opts:
                    if RE_NONE.search(o["text"]) and not o.get("exclusive") and not o.get("other"):
                        add("Q-07", "경고", code, f"`{o['no']}) {o['text']}` — 복수 선택에서 배타 보기로 보이나 exclusive 표시 없음")
            nominal = kind in ("single", "multi", "rank") and not it.get("ordinal")
            if nominal and len(opts) >= 5 and "rotation" not in it:
                add("Q-10", "경고", code, f"보기 {len(opts)}개 명목형 — rotation 을 true/false 로 명시 (구간·척도형이면 ordinal: true)")
        if kind in ("open", "number") and any(re.match(r"^\s*\d+\)", str(ln)) for ln in _as_list(it.get("lines"))):
            add("Q-05", "경고", code, "입력형 기입 줄이 `1) …` 로 시작 — 파서가 보기로 읽는다. `a. …` 나 fields 로 쓴다")
        if kind == "number" and not it.get("range") and not it.get("columns"):
            add("Q-09", "경고", code, "숫자 입력인데 RANGE 없음")
        if kind == "number" and it.get("columns") and not it.get("range"):
            add("Q-09", "경고", code, "연도별 숫자 입력인데 RANGE 없음")
        if not it.get("src"):
            add("Q-12", "정보", code, "구조화에 없는 문항(src 없음) — 추가한 이유를 메모에 남긴다")
        if it.get("review"):
            add("Q-12", "정보", code, f"확인 필요: {it['review']}")

    # Q-16 — 지시문 안에 대괄호가 겹치면 파서가 첫 `]` 에서 잘라 읽는다 (`[PROG: 이후 [#연도별]은 …]`)
    for pt in spec["parts"]:
        pieces = [(pt["code"], "파트 PROG", x) for x in _as_list(pt.get("prog"))]
        for it in pt.get("items", []):
            where = it.get("code") or f"{pt['code']} {it.get('kind')}"
            for f in ("prog", "prog_before", "base", "range"):
                pieces += [(where, f, x) for x in _as_list(it.get(f))]
            if it.get("kind") == "text":
                for m in re.finditer(r"\[(?:BASE|PROG|RANGE)\s*:([^\]]*)", str(it.get("text", ""))):
                    if "[" in m.group(1):
                        add("Q-16", "경고", where, f"지시문 안에 대괄호가 겹침 — 파서가 잘라 읽는다: {it['text'][:50]}")
        for where, f, x in pieces:
            body = re.sub(r"^\s*\[(?:BASE|PROG|RANGE)\s*:", "", str(x)).rstrip("]").strip()
            if "[" in body or "]" in body:
                add("Q-16", "경고", where, f"{f} 안에 대괄호 `{x}` — 파서가 첫 `]` 에서 잘라 읽는다. `#연도별` 처럼 괄호 없이 쓴다")

    # Q-13
    sq = [it for p, it in items if p == "SQ"]
    if sq:
        has_term = any(o.get("terminate") for it in sq for o in norm_options(it.get("options") or []))
        has_prog = any(re.search(r"조사\s*(종료|중단)", " ".join(_as_list(it.get("prog")))) for it in sq)
        if not (has_term or has_prog):
            add("Q-13", "경고", "SQ", "스크리닝 파트에 조사 중단 장치가 없음")

    # Q-11
    if structure:
        cited: set[str] = set()
        for _, it in items:
            cited.update(_as_list(it.get("src")))
        for pt in spec["parts"]:
            for it in pt.get("items", []):
                if not it.get("code"):
                    cited.update(_as_list(it.get("src")))
        skipped = set(_as_list(spec.get("meta", {}).get("skipped_src")))
        for sh in structure["sheets"]:
            if sh["kind"] != "structure":
                continue
            for rec in sh["rows"]:
                f = rec["fields"]
                if not (f.get("question") or f.get("code") or f.get("attribute") or f.get("section")):
                    continue
                if rec.get("banner") and not f.get("question"):
                    continue
                if rec["ref"] in cited or rec["ref"] in skipped:
                    continue
                label = f.get("code") or f.get("attribute") or f.get("section") or ""
                q = (f.get("question") or "")[:40].replace("\n", " ")
                add("Q-11", "경고", rec["ref"], f"구조화 행이 설문에 반영되지 않음: {label} {q}".strip()
                    + " — 반영하거나 meta.skipped_src 에 사유와 함께 넣는다")

    # Q-14 / Q-15
    if docx_path:
        import parse_questionnaire as pq
        try:
            doc = pq.parse(docx_path)
        except Exception as e:
            add("Q-14", "치명", str(docx_path), f"파서가 읽지 못함: {e}")
            return F
        got = {q["code"]: q for q in doc["questions"] if q["code"] != "ED"}
        want = [it["code"] for _, it in items]
        missing = [c for c in want if c not in got]
        extra = [c for c in got if c not in by_code]
        if missing:
            add("Q-14", "치명", "왕복", f"파서가 못 찾은 문항: {missing}")
        if extra:
            add("Q-14", "치명", "왕복", f"명세에 없는 문항을 파서가 찾음: {extra}")
        for code, q in got.items():
            it = by_code.get(code)
            if not it:
                continue
            exp = EXPECT_TYPE.get(it.get("kind", "single"))
            if it.get("kind") == "number" and q["response_type"] == "개방형" and not it.get("unit"):
                pass
            elif exp and q["response_type"] != exp and not (exp == "단일" and q["response_type"] == "예-아니오"):
                add("Q-14", "치명", code, f"응답유형 왕복 불일치: 명세 {exp} ↔ 파서 {q['response_type']}")
            n = n_answers(it)
            if n is not None and it.get("kind") in ("single", "multi", "rank", "yesno") and q["n_options"] != n:
                add("Q-14", "치명", code, f"보기 수 왕복 불일치: 명세 {n} ↔ 파서 {q['n_options']}")
            if it.get("base") and not q.get("base_condition"):
                add("Q-14", "치명", code, "BASE 가 파서에서 문항에 붙지 않음")
        list_codes = {c for c, it in by_code.items() if it.get("kind") == "list"}
        for w in doc["meta"]["warnings"]:
            code = w.split(":", 1)[0]
            if code in list_codes and "보기가 없음" in w:
                add("Q-15", "정보", code, f"리스트 문항 — 파서는 보기 없는 개방형으로 읽음 (정상): {w}")
            else:
                add("Q-15", "경고", "파서", w)
    return F


def to_md(F: list[dict[str, Any]], spec: dict[str, Any]) -> str:
    n = {s: sum(1 for f in F if f["severity"] == s) for s in ("치명", "경고", "정보")}
    title = spec.get("meta", {}).get("title", "")
    out = [f"# 설문지 검수 — {title}", "",
           f"치명 {n['치명']} · 경고 {n['경고']} · 정보 {n['정보']} → **{'통과' if n['치명'] == 0 else '보류'}**", ""]
    for sev in ("치명", "경고", "정보"):
        rows = [f for f in F if f["severity"] == sev]
        if not rows:
            continue
        out += [f"## {sev} {len(rows)}건", "", "| 검사 | 위치 | 내용 |", "|---|---|---|"]
        out += [f"| {f['check']} | {f['where']} | {f['message'].replace('|', '/')} |" for f in rows]
        out.append("")
    return "\n".join(out)


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="설문 명세 검수")
    ap.add_argument("spec")
    ap.add_argument("--structure")
    ap.add_argument("--docx")
    ap.add_argument("--out")
    ap.add_argument("--md")
    a = ap.parse_args(argv)
    spec = load_spec(Path(a.spec))
    structure = json.loads(Path(a.structure).read_text(encoding="utf-8")) if a.structure else None
    F = check(spec, structure, Path(a.docx) if a.docx else None)
    if a.out:
        Path(a.out).write_text(json.dumps(F, ensure_ascii=False, indent=1), encoding="utf-8")
    md = to_md(F, spec)
    if a.md:
        Path(a.md).write_text(md, encoding="utf-8")
    fatal = sum(1 for f in F if f["severity"] == "치명")
    warn = sum(1 for f in F if f["severity"] == "경고")
    info = len(F) - fatal - warn
    print(f"[check_questionnaire] 치명 {fatal} · 경고 {warn} · 정보 {info} → {'통과' if not fatal else '보류'}")
    for f in F:
        if f["severity"] != "정보":
            print(f"  {f['severity']} {f['check']} {f['where']}: {f['message']}")
    return 1 if fatal else 0


if __name__ == "__main__":
    raise SystemExit(main())
