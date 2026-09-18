#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""webtest_compare.py — 순회 덤프를 설문지 정본과 대조해 findings 와 수정요청서(.xlsx)를 만든다.

무엇을 잡나 (심각도)
    치명  QUESTION_MISSING   보여야 할 문항이 안 나옴(미구현)      · QUESTION_UNEXPECTED  안 보여야 할 문항이 나옴(분기 오류)
          TERMINATE_MISMATCH 이탈해야 할 자리에서 계속 진행(또는 엉뚱한 자리에서 종료)
          PIPING_MISSING     앞 문항 답이 문장에 안 들어감
    경고  OPTION_MISSING · OPTION_EXTRA  보기 누락·추가      · TEXT_MISMATCH  문항 문장이 정본과 다름(오탈자·수정 미반영)
    참고  OPTION_ORDER       보기 순서가 다름(로테이션이면 정상)  · UNMATCHED_PAGE  정본의 어느 문항과도 안 맞는 페이지
          OPTION_MISSING     표시 조건([PROG: …], 보기의 show_if)이 붙은 표 보기가 조건째 없음(숨김이면 정상)

사용: python webtest_compare.py <questionnaire.json> <시나리오.json> <out_dir> [덤프_*.json …]
      덤프를 안 주면 <out_dir>/덤프_*.json 전부를 읽는다.
산출: <out_dir>/findings.json · <out_dir>/수정요청서_<날짜>.xlsx · <out_dir>/대조요약.md
"""
from __future__ import annotations

import argparse
import difflib
import json
import re
import sys
import time
from pathlib import Path

SIM_MATCH = 0.60     # 이 이상이면 같은 문항으로 본다
SIM_EXACT = 0.97     # 이 미만이면 문장이 다르다고 본다
RE_PIPE = re.compile(r"\[#[^#\]]*#\]")


def norm(s: str) -> str:
    s = re.sub(r"\[[^\]]*\]", " ", s or "")          # [1개 선택] 같은 표지 제거
    s = re.sub(r"^\s*(SQ\d+|[A-K]{1,2}\d+(?:-\d+)?)\s*[\.\)]\s*", "", s)
    return re.sub(r"[\s\W_]+", "", s).lower()


def sim(a: str, b: str) -> float:
    a, b = norm(a), norm(b)
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


RE_CODE_POS = re.compile(r"(?<![A-Za-z가-힣])(SQ\d+|[A-K]{1,2}\d+(?:-\d+)?)\s*[\.\)]")
BUTTON_LINES = {"다음", "이전", "next", "prev", "previous", "확인", "제출", "다음▶", "▶다음", "▶", "◀이전"}


def page_question_text(page: dict) -> str:
    """페이지 본문에서 문항 문장 후보.
    문번호가 있으면 거기서 시작(머리글·진행률 제거), 첫 보기 라벨 전에서 끊고, 버튼 문구와 두 글자 이하 조각(단위 「세」 등)은 버린다."""
    t = page.get("text") or ""
    m = RE_CODE_POS.search(t[:600])
    if m:
        t = t[m.start():]
    labels = [o.get("label") for o in page.get("options") or [] if o.get("label")]
    cut = len(t)
    for l in labels[:3]:
        key = RE_LABEL_NO.sub("", l, count=1)[:20] if l else ""
        i = t.find(key) if key else -1
        if i > 0:
            cut = min(cut, i)
    lines = [ln.strip() for ln in t[:cut].splitlines()]
    lines = [ln for ln in lines if ln and ln.lower().replace(" ", "") not in BUTTON_LINES and len(ln) > 2]
    return "\n".join(lines).strip()[:600]


def match_pages(expected: list[dict], pages: list[dict]) -> list[tuple[dict | None, dict | None, float]]:
    """정본 문항(노출 순)과 덤프 페이지를 순서대로 짝짓는다. 코드가 있으면 코드로, 없으면 문장 유사도로."""
    pairs = []
    ei = 0
    by_code = {q["code"]: i for i, q in enumerate(expected)}
    for pg in pages:
        if pg.get("end"):
            continue
        code = pg.get("code_guess")
        ptxt = page_question_text(pg)
        if code and code in by_code and by_code[code] >= ei:
            j = by_code[code]
        else:
            j, best = None, 0.0
            for k in range(ei, min(ei + 4, len(expected))):
                s = sim(expected[k].get("text") or "", ptxt)
                if s > best:
                    best, j = s, k
            if best < SIM_MATCH:
                pairs.append((None, pg, best)); continue
        for k in range(ei, j):            # 건너뛴 정본 문항 = 미구현
            pairs.append((expected[k], None, 0.0))
        pairs.append((expected[j], pg, sim(expected[j].get("text") or "", ptxt)))
        ei = j + 1
    for k in range(ei, len(expected)):
        pairs.append((expected[k], None, 0.0))
    return pairs


RE_LABEL_NO = re.compile(r"^\s*[\(（]?\s*(\d{1,2}|[①-⑳]|[가-힣])\s*[\)）\.．:]\s*")
OPT_SIM = 0.70       # 보기 문구가 이 이상 닮았으면 「같은 보기인데 문구가 다름」, 아래면 「없음/추가」


def norm_label(s: str) -> str:
    """웹 보기 라벨의 앞 번호(1) · ① · (1) · 가.)를 떼고 정규화한다."""
    return norm(RE_LABEL_NO.sub("", s or "", count=1))


def compare_options(q: dict, pg: dict) -> list[dict]:
    """정본 보기 ↔ 웹 보기를 탐욕적으로 짝짓는다. 정확히 맞는 것부터, 다음은 닮은 것, 남은 것은 누락/추가.
    표시 조건(show_if — 표 보기 행의 `[PROG: …]`)이 붙은 보기가 조건째 통째로 없으면 숨김일 수 있어 참고로 낮춘다."""
    opts = q.get("options") or []
    exp = [o.get("text") or "" for o in opts]
    got = [o.get("label") or "" for o in pg.get("options") or [] if o.get("kind") in ("radio", "checkbox", "select")]
    out = []
    if not exp or not got:
        return out
    en = [norm(e) for e in exp]
    gn = [norm_label(g) for g in got]
    pair: dict[int, int] = {}          # exp idx → got idx
    used: set[int] = set()
    for i, e in enumerate(en):         # 1차: 같거나 앞부분이 같은 것
        for k, g in enumerate(gn):
            if k in used:
                continue
            if g == e or (e and g.startswith(e)) or (g and e.startswith(g) and len(g) >= 4):
                pair[i] = k; used.add(k); break
    for i, e in enumerate(en):         # 2차: 닮은 것
        if i in pair:
            continue
        best, bk = 0.0, None
        for k, g in enumerate(gn):
            if k in used:
                continue
            r = difflib.SequenceMatcher(None, e, g).ratio()
            if r > best:
                best, bk = r, k
        if bk is not None and best >= OPT_SIM:
            pair[i] = bk; used.add(bk)
            out.append({"kind": "TEXT_MISMATCH", "sev": "경고", "detail": f"보기 {i+1} 문구 다름 — 정본 「{exp[i]}」 / 웹 「{got[bk]}」"})
    cond_of = {i: " · ".join(o["show_if"]) if isinstance(o["show_if"], list) else str(o["show_if"])
               for i, o in enumerate(opts) if o.get("show_if")}
    shown_conds = {cond_of[i] for i in pair if i in cond_of}     # 같은 조건 보기가 하나라도 나왔으면 그 묶음은 보이는 중
    hidden_groups: dict[str, list[int]] = {}
    for i, e in enumerate(exp):
        if i not in pair:
            if i in cond_of and cond_of[i] not in shown_conds:
                hidden_groups.setdefault(cond_of[i], []).append(i + 1)
                continue
            out.append({"kind": "OPTION_MISSING", "sev": "경고", "detail": f"보기 {i+1} 「{e}」 가 웹에 없음"})
    for cond, nos in hidden_groups.items():
        span = f"{nos[0]}~{nos[-1]}" if len(nos) > 1 and nos == list(range(nos[0], nos[-1] + 1)) else ",".join(map(str, nos))
        out.append({"kind": "OPTION_MISSING", "sev": "참고",
                    "detail": f"표시 조건 [PROG: {cond}] 이 붙은 보기 {span} ({len(nos)}개) 가 웹에 없음 — 이 시나리오에서 조건에 안 맞아 숨긴 것이면 정상"})
    for k, g in enumerate(got):
        if k not in used:
            out.append({"kind": "OPTION_EXTRA", "sev": "경고", "detail": f"웹 보기 {k+1} 「{g}」 가 정본에 없음"})
    order_got = [pair[i] for i in sorted(pair)]
    if order_got != sorted(order_got):
        out.append({"kind": "OPTION_ORDER", "sev": "참고", "detail": "보기 순서가 정본과 다름 (로테이션 지시가 있으면 정상)"})
    return out


def compare(questionnaire: dict, scen: dict, dump: dict) -> list[dict]:
    qs = {q["code"]: q for q in questionnaire.get("questions") or []}
    order = [q["code"] for q in questionnaire.get("questions") or []]
    e = scen["expect"]
    expected = [qs[c] for c in e["visible"] if c in qs]
    hidden = set(e.get("hidden") or [])
    pages = dump.get("pages") or []
    f: list[dict] = []
    sid = scen["id"]

    def add(kind, sev, code, detail, page=None):
        f.append({"scenario": sid, "kind": kind, "sev": sev, "code": code, "detail": detail, "page": (page or {}).get("n")})

    for q, pg, s in match_pages(expected, pages):
        if q and not pg:
            add("QUESTION_MISSING", "치명", q["code"], f"보여야 할 문항 {q['code']} 이(가) 순회에서 나오지 않음 — 미구현이거나 분기 오류")
        elif pg and not q:
            code = pg.get("code_guess")
            if code in hidden:
                add("QUESTION_UNEXPECTED", "치명", code, f"안 보여야 할 문항 {code} 이(가) 나옴 (노출 조건 미적용)", pg)
            else:
                add("UNMATCHED_PAGE", "참고", code, f"정본의 어느 문항과도 안 맞는 페이지 (유사도 {s:.2f}): 「{page_question_text(pg)[:80]}」", pg)
        else:
            piped = bool(RE_PIPE.search(q.get("text") or ""))
            if s < (SIM_EXACT - 0.12 if piped else SIM_EXACT):   # 파이핑 문항은 치환값만큼 달라지므로 느슨하게
                add("TEXT_MISMATCH", "경고", q["code"], f"문항 문장이 정본과 다름 (유사도 {s:.2f}) — 정본 「{(q.get('text') or '')[:80]}」 / 웹 「{page_question_text(pg)[:80]}」", pg)
            for x in compare_options(q, pg):
                add(x["kind"], x["sev"], q["code"], x["detail"], pg)
    # 이탈
    term = e.get("terminate_at")
    last_code = next((p.get("code_guess") for p in reversed(pages) if p.get("code_guess")), None)
    if term:
        seen_after = [p for p in pages if p.get("code_guess") and order.index(p["code_guess"]) > order.index(term)] if term in order else []
        if seen_after:
            add("TERMINATE_MISMATCH", "치명", term, f"{term} 에서 조사 중단이어야 하는데 {seen_after[0]['code_guess']} 로 계속 진행됨", seen_after[0])
        elif not dump.get("ended_by") or "종료 문구" not in str(dump.get("ended_by")):
            add("TERMINATE_MISMATCH", "경고", term, f"{term} 뒤에 종료 문구를 확인하지 못함 (끝: {dump.get('ended_by')})")
    # 파이핑
    for pp in e.get("piping") or []:
        pg = next((p for p in pages if p.get("code_guess") == pp["code"]), None)
        if pg and pp["expects_text"] and norm(pp["expects_text"]) not in norm(pg.get("text") or ""):
            add("PIPING_MISSING", "치명", pp["code"], f"{pp['source']} 의 답 「{pp['expects_text']}」 이(가) {pp['code']} 문장에 없음", pg)
    return f


def dedupe(findings: list[dict]) -> list[dict]:
    """같은 결함이 시나리오마다 잡히면 한 줄로 합치고 시나리오를 나열한다 — 수정요청서는 결함 하나에 한 줄."""
    merged: dict[tuple, dict] = {}
    for x in findings:
        d = re.sub(r"\(유사도 [\d.]+\)", "", x["detail"]).strip()
        if x["kind"] == "PIPING_MISSING":          # 시나리오마다 기대값이 달라도 결함은 하나
            d = re.sub(r"「[^」]*」", "「…」", d)
        key = (x["kind"], x["code"], d)
        if key in merged:
            m = merged[key]
            if x["scenario"] not in m["scenario"].split("·"):
                m["scenario"] += "·" + x["scenario"]
        else:
            merged[key] = dict(x)
    return list(merged.values())


def write_xlsx(findings: list[dict], path: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook(); ws = wb.active; ws.title = "수정요청"
    head = ["No", "시나리오", "문항", "구분", "심각도", "현상", "요청", "페이지"]
    ws.append(head)
    for c in ws[1]:
        c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="DDE6F0")
    ask = {"QUESTION_MISSING": "문항 구현 및 노출 조건 확인 요청", "QUESTION_UNEXPECTED": "노출 조건(BASE) 적용 요청", "TERMINATE_MISMATCH": "조사 중단(이탈) 로직 적용 요청",
           "PIPING_MISSING": "파이핑 적용 요청", "OPTION_MISSING": "보기 추가 요청", "OPTION_EXTRA": "보기 삭제 또는 정본 확인 요청", "TEXT_MISMATCH": "문구 정정 요청",
           "OPTION_ORDER": "로테이션 여부 확인", "UNMATCHED_PAGE": "정본 대조 확인"}
    ask_ref = {"OPTION_MISSING": "표시 조건(PROG) 적용 확인 — 조건상 숨김이면 조치 불필요"}   # 참고로 낮춘 것
    sev_rank = {"치명": 0, "경고": 1, "참고": 2}
    for i, x in enumerate(sorted(findings, key=lambda r: (sev_rank.get(r["sev"], 9), r["code"] or "")), 1):
        req = (ask_ref.get(x["kind"]) if x["sev"] == "참고" else None) or ask.get(x["kind"], "")
        ws.append([i, x["scenario"], x["code"] or "", x["kind"], x["sev"], x["detail"], req, x.get("page") or ""])
    for col, w in zip("ABCDEFGH", (5, 12, 10, 20, 8, 70, 30, 8)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(str(path))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("questionnaire"); ap.add_argument("scenarios"); ap.add_argument("out_dir"); ap.add_argument("dumps", nargs="*")
    a = ap.parse_args(argv)
    q = json.loads(Path(a.questionnaire).read_text(encoding="utf-8"))
    sc = {s["id"]: s for s in json.loads(Path(a.scenarios).read_text(encoding="utf-8"))["scenarios"]}
    out = Path(a.out_dir); out.mkdir(parents=True, exist_ok=True)
    dumps = [Path(d) for d in a.dumps] or sorted(out.glob("덤프_*.json"))
    all_f: list[dict] = []
    for d in dumps:
        dump = json.loads(d.read_text(encoding="utf-8"))
        s = sc.get(dump.get("scenario"))
        if not s:
            print(f"시나리오 {dump.get('scenario')} 를 시나리오.json 에서 찾지 못함 — {d.name} 건너뜀"); continue
        if not s.get("expect"):           # 사람이 답만 적어 보탠 시나리오 — 기대 노출은 S0 의 것을 빌린다
            s = dict(s, expect=sc["S0"]["expect"]) if "S0" in sc else None
            if s is None:
                print(f"시나리오 {dump.get('scenario')} 에 expect 가 없고 S0 도 없음 — {d.name} 건너뜀"); continue
        all_f += compare(q, s, dump)
    (out / "findings.json").write_text(json.dumps(all_f, ensure_ascii=False, indent=1), encoding="utf-8")
    stamp = time.strftime("%y%m%d")
    write_xlsx(dedupe(all_f), out / f"수정요청서_{stamp}.xlsx")
    from collections import Counter
    c = Counter(x["sev"] for x in all_f); k = Counter(x["kind"] for x in all_f)
    uniq = dedupe(all_f)
    lines = [f"# 웹테스트 대조 요약 ({time.strftime('%Y-%m-%d %H:%M')})", "",
             f"덤프 {len(dumps)}개 · 지적 {len(all_f)}건(결함 {len(uniq)}건) — 치명 {c.get('치명',0)} · 경고 {c.get('경고',0)} · 참고 {c.get('참고',0)}", "",
             "| 구분 | 건수 |", "|---|---:|"] + [f"| {kk} | {v} |" for kk, v in k.most_common()] + ["", "| 시나리오 | 문항 | 심각도 | 현상 |", "|---|---|---|---|"]
    for x in sorted(uniq, key=lambda r: ({"치명": 0, "경고": 1, "참고": 2}.get(r["sev"], 9), r["code"] or "", r["scenario"])):
        lines.append(f"| {x['scenario']} | {x['code'] or ''} | {x['sev']} | {x['detail'][:140].replace(chr(10), ' ')} |")
    (out / "대조요약.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"지적 {len(all_f)}건(결함 {len(uniq)}건) — 치명 {c.get('치명',0)} · 경고 {c.get('경고',0)} · 참고 {c.get('참고',0)} → {out}/findings.json · 수정요청서_{stamp}.xlsx · 대조요약.md")
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    raise SystemExit(main())
