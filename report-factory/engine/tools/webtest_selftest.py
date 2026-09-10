#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""webtest_selftest.py — 웹테스트 도구 셋의 자체 시험. 모의 설문(webtest_fixture/설문.html)을 실제 Chromium 으로 돈다.

기대: 픽스처에 심어 둔 결함 3종을 대조기가 잡는다.
    A2 보기 2 문구 오타 → TEXT_MISMATCH(보기) · A2 보기 4 누락 → OPTION_MISSING · A3 파이핑 미적용 → PIPING_MISSING
    그리고 이탈 시나리오(T_SQ1_3)는 종료 문구로 끝나고, 숨김 시나리오(H1_SQ1)에서는 A1 이 안 나와 QUESTION_MISSING 이 0건이어야 한다.

사용: python webtest_selftest.py [--keep]
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
FIX = HERE / "webtest_fixture"


def run(*args: str) -> str:
    r = subprocess.run([sys.executable, *args], capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        print(r.stdout); print(r.stderr); raise SystemExit(f"실패: {' '.join(args[:2])}")
    return r.stdout


def main() -> int:
    keep = "--keep" in sys.argv
    out = Path(tempfile.mkdtemp(prefix="webtest_selftest_"))
    try:
        run(str(HERE / "webtest_scenarios.py"), str(FIX / "questionnaire.json"), str(out))
        sc = json.loads((out / "시나리오.json").read_text(encoding="utf-8"))
        ids = [s["id"] for s in sc["scenarios"]]
        assert "S0" in ids and "T_SQ1_3" in ids and any(i.startswith("H1_SQ1") for i in ids), ids
        url = FIX.joinpath("설문.html").as_uri()
        run(str(HERE / "webtest_walk.py"), url, str(out / "시나리오.json"), str(out), "--scenario", "all", "--max-pages", "20")
        run(str(HERE / "webtest_compare.py"), str(FIX / "questionnaire.json"), str(out / "시나리오.json"), str(out))
        f = json.loads((out / "findings.json").read_text(encoding="utf-8"))
        kinds = {(x["scenario"], x["kind"], x["code"]) for x in f}
        checks = {
            "A2 보기 4 누락 → OPTION_MISSING": ("S0", "OPTION_MISSING", "A2") in kinds,
            "A2 보기 2 오타 → TEXT_MISMATCH": ("S0", "TEXT_MISMATCH", "A2") in kinds,
            "A3 파이핑 미적용 → PIPING_MISSING": ("S0", "PIPING_MISSING", "A3") in kinds,
            "S0 미구현 문항 0건": not any(k[0] == "S0" and k[1] == "QUESTION_MISSING" for k in kinds),
            "이탈 시나리오 종료 확인": not any(k[0] == "T_SQ1_3" and k[1] == "TERMINATE_MISMATCH" for k in kinds),
            "숨김 시나리오에서 A1 미노출": not any(k[0].startswith("H1_SQ1") and k[1] in ("QUESTION_UNEXPECTED",) for k in kinds),
        }
        ok = True
        for name, passed in checks.items():
            print(("  OK   " if passed else "  FAIL ") + name); ok &= passed
        print(f"\n지적 {len(f)}건 · 산출 {out}" + ("" if keep else " (지움)"))
        if not ok:
            for x in f: print("   ", x["scenario"], x["kind"], x["code"], x["detail"][:100])
            return 1
        print("=== 전부 통과 ===")
        return 0
    finally:
        if not keep:
            shutil.rmtree(out, ignore_errors=True)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    raise SystemExit(main())
