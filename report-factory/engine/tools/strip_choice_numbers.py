# -*- coding: utf-8 -*-
"""선지 번호 접두 제거 — 통계표가 「번호 있는 버전」으로 입고됐을 때 파이프라인 산출물을 정규화한다.

보통 통계표는 선지에 번호 없는 버전으로 오지만, 번호 있는 버전으로 받은 경우
행 라벨이 `1) 직접 알고 있음` 형태로 파싱된다. 이 도구는 그 접두(`n) `)를
tables.json · slide_facts.json · KPI요약표 산출물에서 일괄 제거한다.

건드리지 않는 것 (선지 번호가 아니므로):
  - slide_title (`1. 현황` — 챕터 번호) · sub_title (`1) 소유 | 형태` — 중분류 번호)
  - base_row (`(50)` — 베이스 n) · qu/QA (설문 원문의 자체 번호)
  - questionnaire.json 전체 (설문지는 원문 보존)

안전장치:
  - 제거 후 같은 표 안에서 라벨이 중복되면 **중단**한다 (라벨 매칭 검사가 깨진다)
  - 유일한 기지 충돌 = D4 보기 4 오타(설문지 단계부터 `신뢰가 가지 않는다`가 2)와 중복).
    5점 척도 구조상(BOT2=1+2, TOP2=4+5) 4)는 `신뢰가 간다`가 맞으므로 함께 정정한다
    (프로젝트 이슈 #7 「라벨 정정」). --no-d4-fix 로 끌 수 있다.
  - 기본 dry-run. --apply 시에만 쓰며, 원본은 work/_backup_선지번호제거/ 에 복사한다
  - tables.json meta 에 postprocess 기록을 남긴다

사용:
  python engine/tools/strip_choice_numbers.py <work_dir> [--apply] [--no-d4-fix]
"""
import argparse
import datetime as _dt
import json
import re
import shutil
import sys
from collections import Counter
from pathlib import Path

# 1~3자리: 보통 선지는 1~2자리, 특수코드(998=없음 등)가 3자리로 온다
PREFIX = re.compile(r"^(\s*)\d{1,3}\)\s*")
D4_BAD = re.compile(r"^(\s*)4\)\s*신뢰가 가지 않는다\s*$")
D4_GOOD = "신뢰가 간다"

stats = Counter()


def fix(s, where, d4=True):
    if not isinstance(s, str):
        return s
    if d4:
        m = D4_BAD.match(s)
        if m:
            stats[f"{where}|D4정정"] += 1
            return m.group(1) + D4_GOOD
    m = PREFIX.match(s)
    if m:
        stats[f"{where}|제거"] += 1
        return PREFIX.sub(r"\1", s, count=1)
    return s


def fix_row(row, where, d4):
    for k in ("label", "label_clean", "label_raw"):
        if k in row:
            row[k] = fix(row[k], where, d4)


def collisions(groups, what):
    """groups: iterable of (이름, [라벨...]). 중복 라벨을 찾아 보고한다."""
    bad = []
    for name, labels in groups:
        c = Counter(l.strip() for l in labels if isinstance(l, str) and l.strip()
                    and not l.strip().startswith("【") and l.strip() not in ("계", "[5점 평균]"))
        dup = {k: v for k, v in c.items() if v > 1}
        if dup:
            bad.append((name, dup))
    if bad:
        print(f"\n🔴 {what} 라벨 충돌 — 중단한다:")
        for name, dup in bad:
            print(f"   {name}: {dup}")
    return bad


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("work_dir")
    ap.add_argument("--apply", action="store_true", help="실제로 쓴다 (기본 dry-run)")
    ap.add_argument("--no-d4-fix", action="store_true", help="D4 보기4 오타 정정을 하지 않는다")
    args = ap.parse_args()
    work = Path(args.work_dir)
    d4 = not args.no_d4_fix

    targets = {}  # path → (수정된 객체, dumps kwargs)

    # ── tables.json ──
    p_tab = work / "tables.json"
    tj = json.loads(p_tab.read_text(encoding="utf-8"))
    for t in tj["tables"]:
        for r in t.get("rows", []):
            fix_row(r, "tables.json rows", d4)
        if isinstance(t.get("suffix_parts"), list):
            t["suffix_parts"] = [fix(x, "tables.json suffix_parts", d4) for x in t["suffix_parts"]]
    tj.setdefault("meta", {}).setdefault("postprocess", []).append({
        "tool": "strip_choice_numbers.py",
        "date": _dt.date.today().isoformat(),
        "what": "선지 번호 접두(n) ) 제거" + (" + D4 보기4 라벨 정정(이슈 #7)" if d4 else ""),
    })
    # 라벨 **내용**을 바꾼 정정은 표·보기 단위로 남긴다 — 빌더가 해당 장 Qu 에 정정 표기를 단다 (ADV-11)
    if d4:
        corr = tj["meta"].setdefault("label_corrections", [])
        for t in tj["tables"]:
            for r in t.get("rows", []):
                if r.get("option_no") == 4 and (r.get("label") or "").strip() == D4_GOOD and (t.get("question_code") == "D4"):
                    corr.append({"table": t.get("index"), "question_code": "D4", "option_no": 4,
                                 "old": "신뢰가 가지 않는다", "new": D4_GOOD,
                                 "why": "설문지 보기 2·4 중복 오타 — 5점 척도 구조(TOP2=4+5)상 긍정 라벨"})
    targets[p_tab] = tj

    # ── slide_facts.json (Step 1 에서는 아직 없다 — 있을 때만 정규화) ──
    p_sf = work / "slide_facts.json"
    sf = json.loads(p_sf.read_text(encoding="utf-8")) if p_sf.exists() else {"slides": []}
    for s in sf.get("slides", []):
        g = s.get("grid") or {}
        for r in g.get("rows", []):
            fix_row(r, "slide_facts grid", d4)
        mf = s.get("ment_facts") or {}
        for key in ("line1", "line2"):
            for e in mf.get(key) or []:
                for k in ("label", "row_label"):
                    if k in e:
                        e[k] = fix(e[k], f"slide_facts ment_facts.{key}", d4)
    def fix_ties(node):
        """ties = ment_facts.ties: {"표.열": [동점 라벨...]} — 리스트 원소만 고친다."""
        if isinstance(node, dict):
            for k, v in node.items():
                if isinstance(v, list):
                    node[k] = [fix(x, "slide_facts ties", d4) for x in v]
                else:
                    fix_ties(v)

    fix_ties(sf.get("ties"))
    for s in sf.get("slides", []):
        fix_ties(s.get("ties"))
        fix_ties((s.get("ment_facts") or {}).get("ties"))
    if p_sf.exists():
        targets[p_sf] = sf

    # ── KPI요약표_직접생성.json ──
    p_kpi = work / "KPI요약표_직접생성.json"
    if p_kpi.exists():
        kj = json.loads(p_kpi.read_text(encoding="utf-8"))
        for t in kj.get("tables", []):
            for r in t.get("rows", []):
                fix_row(r, "KPI json rows", d4)
        targets[p_kpi] = kj

    # ── 충돌 검사 (적용 전) ──
    bad = collisions(
        [(f"표{t.get('index')}({t.get('question_code')})",
          [r.get("label", "") for r in t.get("rows", [])]) for t in tj["tables"]],
        "tables.json",
    )
    bad += collisions(
        [(f"slide{s.get('seq')}({s.get('question_code')})",
          [r.get("label", "") for r in (s.get("grid") or {}).get("rows", [])])
         for s in sf.get("slides", [])],
        "slide_facts.json",
    )
    if bad:
        sys.exit(1)

    # ── 보고 ──
    print("── 변경 집계 ──")
    for k in sorted(stats):
        print(f"  {stats[k]:>5}  {k}")
    print(f"  합계 {sum(stats.values())}")

    if not args.apply:
        print("\n(dry-run — 아무것도 쓰지 않았다. 적용하려면 --apply)")
        return

    # ── 백업 후 쓰기 ──
    bak = work / "_backup_선지번호제거"
    bak.mkdir(exist_ok=True)
    for p in targets:
        dst = bak / p.name
        if not dst.exists():
            shutil.copy2(p, dst)
    # KPI xlsx 도 백업·정규화
    p_x = work / "KPI요약표_직접생성.xlsx"
    if p_x.exists():
        dst = bak / p_x.name
        if not dst.exists():
            shutil.copy2(p_x, dst)
        import openpyxl
        wb = openpyxl.load_workbook(p_x)
        n = 0
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str):
                        nv = fix(c.value, "KPI xlsx", d4)
                        if nv != c.value:
                            c.value = nv
                            n += 1
        wb.save(p_x)
        print(f"  KPI xlsx 셀 {n}개 정규화")

    for p, obj in targets.items():
        p.write_text(json.dumps(obj, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        print(f"  썼다: {p.name}")
    print(f"  백업: {bak}")


if __name__ == "__main__":
    main()
