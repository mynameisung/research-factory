#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""questionnaire_selftest.py — 설문지작성 도구 셋의 자체 시험.

기대
    · 모의 명세(questionnaire_fixture/sample_spec.json) → docx → 검수: 치명 0, 파서 왕복 일치
    · 결함을 심은 명세는 검수기가 잡는다
        번호 중복 → Q-01 · 없는 문항 BASE → Q-03 · BASE 코드 범위 초과 → Q-04 · 뒤 문항 파이핑 → Q-03
        보기 1개 단일 → Q-05 · 1점 긍정 척도 → Q-08 · 기타가 중간 → Q-06 · 구조화 행 누락 → Q-11

사용: python questionnaire_selftest.py [--keep]
"""
from __future__ import annotations

import copy
import json
import shutil
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import build_questionnaire as bq  # noqa: E402
import check_questionnaire as cq  # noqa: E402

FIX = HERE / "questionnaire_fixture" / "sample_spec.json"


def _item(spec, code):
    for pt in spec["parts"]:
        for it in pt["items"]:
            if it.get("code") == code:
                return it
    raise KeyError(code)


def main(argv: list[str]) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    tmp = Path(tempfile.mkdtemp(prefix="questionnaire_selftest_"))
    ok = True

    def expect(cond: bool, label: str) -> None:
        nonlocal ok
        print(f"  {'OK  ' if cond else 'FAIL'} {label}")
        ok &= cond

    spec = bq.load_spec(FIX)
    out = tmp / "sample.docx"
    bq.Builder(copy.deepcopy(spec)).build(out)
    F = cq.check(spec, None, out)
    fatal = [f for f in F if f["severity"] == "치명"]
    expect(out.is_file() and not fatal, f"모의 명세 빌드·왕복 치명 0 ({len(fatal)})")
    if fatal:
        for f in fatal:
            print("       ", f)

    bad = copy.deepcopy(spec)
    _item(bad, "SQ2")["code"] = "SQ1"                                  # 중복
    _item(bad, "A2")["base"] = "A9=1 | 없는 문항"                         # 없는 문항
    _item(bad, "B2")["base"] = "B1=4~7 | 범위 초과"                       # B1 은 5점
    _item(bad, "A1")["text"] = "[#A2 응답값#] 을 기준으로 답해 주십시오."     # 뒤 문항 파이핑
    _item(bad, "DQ1")["options"] = ["미혼"]                              # 보기 1개
    _item(bad, "A2")["scale"] = {"labels": ["매우 만족한다", "만족한다", "보통이다", "불만족한다", "매우 불만족한다"]}
    _item(bad, "B2")["options"] = ["원두", "기타", "캡슐"]                # 기타 중간
    structure = {"sheets": [{"kind": "structure", "rows": [
        {"ref": "구조화!3", "fields": {"code": "SQ1", "question": "성별"}},
        {"ref": "구조화!99", "fields": {"code": "Z1", "question": "빠뜨린 문항"}},
    ]}]}
    F = cq.check(bad, structure, None)
    got = {(f["check"], f["where"]) for f in F}
    has = lambda chk, where: any(c == chk and w.startswith(where) for c, w in got)  # noqa: E731
    expect(has("Q-01", "SQ1"), "번호 중복 → Q-01")
    expect(has("Q-03", "A2"), "없는 문항 BASE → Q-03")
    expect(has("Q-04", "B2"), "BASE 코드 범위 초과 → Q-04")
    expect(has("Q-03", "A1"), "뒤 문항 파이핑 → Q-03")
    expect(has("Q-05", "DQ1"), "보기 1개 → Q-05")
    expect(has("Q-08", "A2"), "1점 긍정 척도 → Q-08")
    expect(has("Q-06", "B2"), "기타 중간 → Q-06")
    expect(has("Q-11", "구조화!99"), "구조화 행 누락 → Q-11")
    expect(not has("Q-11", "구조화!3"), "반영된 구조화 행은 누락 아님")

    if "--keep" in argv:
        print(f"산출 {tmp} (남김)")
    else:
        shutil.rmtree(tmp, ignore_errors=True)
    print("=== 전부 통과 ===" if ok else "=== 실패 있음 ===")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
