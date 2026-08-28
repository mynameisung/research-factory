# -*- coding: utf-8 -*-
"""
apply_table_fixes.py — 통계표 라벨 정정을 파싱 직후 적용한다 (2026-08-24 신설)

  python engine/tools/apply_table_fixes.py <tables.json> <work/표라벨정정.json>

통계표는 사람이 만든 산출물이라 라벨 오타가 있고 파서는 그것을 경고하지 않는다. 손으로 고치면
재파싱 때마다 되돌아가므로, 정정을 **파일로 선언**하고 파이프라인이 매번 적용하게 한다.

**수치는 고치지 않는다.** 라벨만 고치고, 고친 자리는 `label_fixed` 로 표시해 나중에 추적할 수 있게 남긴다.
지정한 before 문구가 그 행에 실제로 있지 않으면 **적용하지 않고 실패로 알린다** — 표가 바뀌었는데
정정이 남아 엉뚱한 행을 고치는 사고를 막는다.
"""
from __future__ import annotations
import argparse, io, json, sys
from pathlib import Path


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("tables", type=Path)
    ap.add_argument("fixes", type=Path)
    a = ap.parse_args()
    if not a.fixes.exists():
        print("[apply_table_fixes] 정정 파일이 없습니다 — 건너뜁니다")
        return 0
    with io.open(a.tables, encoding="utf-8") as fh:
        doc = json.load(fh)
    tables = doc["tables"] if isinstance(doc, dict) else doc
    by_index = {t["index"]: t for t in tables}
    fixes = json.loads(a.fixes.read_text(encoding="utf-8")).get("정정") or []

    ok = fail = 0
    for f in fixes:
        t = by_index.get(int(f["표"]))
        if t is None:
            print(f"[실패] 표{f['표']} 를 찾지 못했습니다"); fail += 1; continue
        row = next((r for r in (t.get("rows") or []) if r.get("row") == int(f["행"]) - 1), None)
        if row is None:
            print(f"[실패] 표{f['표']} 행{f['행']} 을 찾지 못했습니다"); fail += 1; continue
        cur = str(row.get("label") or "").strip()
        if cur != str(f["before"]).strip():
            print(f"[실패] 표{f['표']} 행{f['행']} 의 현재 라벨이 {cur!r} 라 정정 대상({f['before']!r})과 다릅니다 "
                  f"— 통계표가 바뀌었는지 확인하십시오")
            fail += 1
            continue
        row["label"] = f["after"]
        if row.get("label_raw") is not None:
            row["label_raw"] = str(row["label_raw"]).replace(f["before"], f["after"])
        if row.get("option_text"):
            row["option_text"] = f["after"]
        row["label_fixed"] = {"before": f["before"], "근거": f.get("근거"), "검출": f.get("검출")}
        ok += 1
        print(f"[적용] 표{f['표']} 행{f['행']} {f['before']!r} → {f['after']!r}")

    # ── 표 제목 정정 — 통계표 제목이 문항을 잘못 적은 경우 (행 라벨과 채널이 다르다).
    #    에스원 TA Pool 표15~18 이 `A4- b. 자산관리 항목` 인데 A4 는 시설관리(FM) 문항이다.
    #    그대로 옮기면 한 슬라이드 안에서 제목(시설관리)과 표 참조(자산관리)가 반대말을 한다.
    titles = json.loads(a.fixes.read_text(encoding="utf-8")).get("제목정정") or []
    for f in titles:
        t = by_index.get(int(f["표"]))
        if t is None:
            print(f"[실패] 표{f['표']} 를 찾지 못했습니다"); fail += 1; continue
        cur = str(t.get("title_full") or "")
        if str(f["before"]) not in cur:
            print(f"[실패] 표{f['표']} 제목에 {f['before']!r} 가 없습니다 (현재 {cur[:50]!r})")
            fail += 1
            continue
        for key in ("title_full", "title_base"):
            if t.get(key):
                t[key] = str(t[key]).replace(str(f["before"]), str(f["after"]))
        t["title_fixed"] = {"before": f["before"], "근거": f.get("근거"), "검출": f.get("검출")}
        ok += 1
        print(f"[적용] 표{f['표']} 제목 {f['before']!r} → {f['after']!r}")

    doc.setdefault("meta", {})["label_fixes"] = {"applied": ok, "failed": fail, "source": str(a.fixes)}
    a.tables.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"→ {a.tables}  (적용 {ok} · 실패 {fail})")
    return 1 if fail else 0


if __name__ == "__main__":
    raise SystemExit(main())
