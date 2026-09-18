#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""questionnaire_selftest.py — 설문지작성 도구 셋의 자체 시험. setup 이 설치 끝에 돌린다.

기대
    · 모의 명세(questionnaire_fixture/sample_spec.json, 모든 kind) → docx → 검수: 치명·경고 0, 파서 왕복 일치, 문서 위생 깨끗
    · 결함을 심은 명세는 검수기가 그 문항에서 잡는다 (위치는 정확히 일치 — `A2` 기대가 `A2-1` 결과로 통과하지 않는다)
        번호 중복 → Q-01 · 번호 역행 → Q-02 · 없는 문항 BASE · 뒤 문항 파이핑 → Q-03 · BASE 코드 범위 초과 → Q-04
        보기 1개 → Q-05 · 기타 중간 → Q-06 · 없음 보기 배타 누락 → Q-07 · 1점 긍정 척도 → Q-08 · RANGE 없음 → Q-09
        로테이션 미정 → Q-10 · 구조화 행 누락 → Q-11 · 중단 장치 없음 → Q-13 · 지시문 속 대괄호 → Q-16
    · 고친 결함이 되살아나지 않는다 — 검수기(거짓 경보·BASE/PROG 왕복·skipped_src) · 빌더(기입란·넓은 행렬·동의 박스) ·
      파서(표 보기·숫자 표) · 구조화 읽기(빈 행 뒤 본표·같은 머리말·병합 머리·색) · 연계(웹테스트·제작가이드·경로)
입력 xlsx·docx 는 임시 폴더에 그때그때 만든다 (바이너리 고정물을 두지 않는다).

사용: python questionnaire_selftest.py [--keep]
"""
from __future__ import annotations

import argparse
import contextlib
import copy
import io
import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import build_questionnaire as bq  # noqa: E402
import check_questionnaire as cq  # noqa: E402
import parse_questionnaire as pq  # noqa: E402

FIX = HERE / "questionnaire_fixture" / "sample_spec.json"
KINDS = {"single", "multi", "rank", "scale", "matrix", "number", "open", "yesno", "list", "info", "text", "table"}


def _item(spec, code):
    for pt in spec["parts"]:
        for it in pt["items"]:
            if it.get("code") == code:
                return it
    raise KeyError(code)


def _insert_before(spec, code, new):
    for pt in spec["parts"]:
        for i, it in enumerate(pt["items"]):
            if it.get("code") == code:
                pt["items"].insert(i, new)
                return
    raise KeyError(code)


# 검수 결과 조회 — 위치(where)는 정확히 일치해야 한다
def _has(F, chk, where, sev=None) -> bool:
    return any(f["check"] == chk and f["where"] == where and (sev is None or f["severity"] == sev) for f in F)


def _msg(F, chk, where) -> str:
    return " / ".join(f["message"] for f in F if f["check"] == chk and f["where"] == where)


def _where(F, chk) -> set[str]:
    return {f["where"] for f in F if f["check"] == chk}


def _cli(fn, args: list[str]) -> tuple[int | str, str]:
    """도구 main() 을 이 프로세스 안에서 — (종료 코드, stderr). 화면 출력은 삼킨다."""
    err = io.StringIO()
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(err):
        try:
            rc = fn(args)
        except SystemExit as e:
            rc = e.code
        except Exception as e:                      # 한 줄 오류 대신 트레이스백으로 죽음
            rc = f"예외 {type(e).__name__}"
    return rc, err.getvalue()


def main(argv: list[str]) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="설문지작성 도구 자체 시험 — 모의 명세 왕복 + 고친 결함 회귀")
    ap.add_argument("--keep", action="store_true", help="임시 산출(docx·xlsx)을 지우지 않고 남긴다")
    a = ap.parse_args(argv)
    tmp = Path(tempfile.mkdtemp(prefix="questionnaire_selftest_"))
    ok = True

    def expect(cond, label: str, detail=None) -> None:
        nonlocal ok
        print(f"  {'OK  ' if cond else 'FAIL'} {label}")
        if not cond and detail:
            for d in detail if isinstance(detail, list) else [detail]:
                print("       ", str(d).replace("\n", " ⏎ "))
        ok &= bool(cond)

    @contextlib.contextmanager
    def section(name: str):
        print(f"[{name}]")
        try:
            yield
        except Exception as e:                      # 한 묶음이 터져도 나머지 묶음은 돌린다
            expect(False, f"{name} — 예외 {type(e).__name__}: {e}")

    spec = bq.load_spec(FIX)
    out = tmp / "sample.docx"

    # ── 모의 명세 ─────────────────────────────────────────────────────────
    with section("모의 명세"):
        bq.Builder(copy.deepcopy(spec)).build(out)
        F = cq.check(spec, None, out)
        fatal = [f for f in F if f["severity"] == "치명"]
        warn = [f for f in F if f["severity"] == "경고"]
        expect(out.is_file() and not fatal, f"모의 명세 빌드·왕복 치명 0 ({len(fatal)})", fatal)
        expect(not warn, f"모의 명세 경고 0 — BASE·PROG 왕복 잡음 없음 ({len(warn)})", warn)
        kinds = {it.get("kind", "single") for pt in spec["parts"] for it in pt["items"]}
        expect(kinds >= KINDS, "모의 명세에 모든 kind (text 포함)", f"빠짐 {sorted(KINDS - kinds)}")
        import doc_hygiene
        hy = doc_hygiene.scan_file(out, "엠브레인")
        expect(not hy["props"] and not hy["invisible"], "문서 위생 깨끗 — 작성자 조직명 · lastPrinted 없음",
               hy["props"] + list(hy["invisible"]))
        with zipfile.ZipFile(out) as z:
            names = set(z.namelist())
            app = z.read("docProps/app.xml").decode("utf-8") if "docProps/app.xml" in names else ""
        expect("docProps/custom.xml" not in names and "<Pages>" not in app, "양식 잔재 없음 — custom.xml · 원본 페이지 수")

    # ── 검수기: 심은 결함 ─────────────────────────────────────────────────
    with section("검수기 — 심은 결함"):
        bad = copy.deepcopy(spec)
        _item(bad, "SQ2")["code"] = "SQ1"                                  # 중복
        _item(bad, "A2")["base"] = "A9=1 | 없는 문항"                         # 없는 문항
        _item(bad, "B2")["base"] = "B1=4~7 | 범위 초과"                       # B1 은 5점
        _item(bad, "A1")["text"] = "[#A2 응답값#] 을 기준으로 답해 주십시오."     # 뒤 문항 파이핑
        _item(bad, "DQ1")["options"] = ["미혼"]                              # 보기 1개
        _item(bad, "A2")["scale"] = {"labels": ["매우 만족한다", "만족한다", "보통이다", "불만족한다", "매우 불만족한다"]}
        _item(bad, "B2")["options"] = ["원두", "기타", "캡슐"]                # 기타 중간
        _item(bad, "A2-1")["options"] = _item(bad, "A2-1")["options"] + ["해당 없음"]   # 배타 표시 없음
        _item(bad, "DQ3")["code"] = "DQ1-1"                                 # DQ2 뒤로 번호 역행
        del _item(bad, "DQ2")["range"]                                      # RANGE 없음
        del _item(bad, "A1-1")["rotation"]                                  # 보기 6개 명목형
        _item(bad, "A4")["prog"] = ["이후 [#연도별]은 반복"]                    # 지시문 속 대괄호
        structure = {"sheets": [{"kind": "structure", "rows": [
            {"ref": "구조화!3", "fields": {"code": "SQ1", "question": "성별"}},
            {"ref": "구조화!99", "fields": {"code": "Z1", "question": "빠뜨린 문항"}},
        ]}]}
        F = cq.check(bad, structure, None)
        for chk, where, label in (
                ("Q-01", "SQ1", "번호 중복"), ("Q-02", "DQ1-1", "번호 역행"), ("Q-03", "A2", "없는 문항 BASE"),
                ("Q-03", "A1", "뒤 문항 파이핑"), ("Q-04", "B2", "BASE 코드 범위 초과"), ("Q-05", "DQ1", "보기 1개"),
                ("Q-06", "B2", "기타 중간"), ("Q-07", "A2-1", "없음 보기 배타 누락"), ("Q-08", "A2", "1점 긍정 척도"),
                ("Q-09", "DQ2", "RANGE 없음"), ("Q-10", "A1-1", "로테이션 미정"), ("Q-16", "A4", "지시문 속 대괄호"),
                ("Q-11", "구조화!99", "구조화 행 누락")):
            expect(_has(F, chk, where), f"{label} → {chk} {where}")
        expect(not _has(F, "Q-11", "구조화!3"), "반영된 구조화 행은 누락 아님")
        nosq = copy.deepcopy(spec)
        nosq["parts"][0]["items"] = [copy.deepcopy(_item(spec, "SQ1"))]
        expect(_has(cq.check(nosq, None, None), "Q-13", "SQ"), "스크리닝에 중단 장치 없음 → Q-13 SQ")

    # ── 검수기: 거짓 경보 · 새 검사 ────────────────────────────────────────
    with section("검수기 — 거짓 경보·새 검사"):
        s = copy.deepcopy(spec)
        _item(s, "A1-1")["base"] = "A1=1~2 | B2B 거래 기업 담당자, 2025년 Q1 이용자, PM2.5 민감, K5·S24"
        _item(s, "DQ1")["options"] = ["미혼", "기혼", {"text": "모름", "no": 9}]
        _insert_before(s, "DQ2", {"code": "DQ1-1", "kind": "single", "base": "DQ1=9 | 모름 응답자",
                                  "text": "그렇게 답하신 이유는 무엇입니까?", "options": ["가", "나"], "src": ["구조화!19"]})
        _item(s, "DQ2")["base"] = "DQ1=3 | 없는 코드"
        _item(s, "B2")["base"] = "A7=1~10 | 산지 응답자"                   # 묶음 표 이어 번호 1~10
        _item(s, "B3")["base"] = "A7=11 | 없는 산지"
        _item(s, "B4")["lines"] = ["1) 첫째 이유", "2) 둘째 이유"]
        F = cq.check(s, None, None)
        expect(not (_has(F, "Q-03", "A1-1") or _has(F, "Q-04", "A1-1")),
               "BASE `|` 뒤 설명의 B2B·Q1·K5 는 문항이 아님 → Q-03·Q-04 없음", _msg(F, "Q-03", "A1-1"))
        expect(not _has(F, "Q-04", "DQ1-1"), "명시 번호 보기(no: 9)를 BASE 가 쓰면 통과", _msg(F, "Q-04", "DQ1-1"))
        expect(_has(F, "Q-04", "DQ2") and "1~2, 9" in _msg(F, "Q-04", "DQ2"), "없는 코드 3 → Q-04 (응답 코드 1~2, 9)",
               _msg(F, "Q-04", "DQ2"))
        expect(not _has(F, "Q-04", "B2") and _has(F, "Q-04", "B3"), "묶음 표 이어 번호 1~10 은 통과, 11 은 Q-04")
        expect(_has(F, "Q-05", "B4", "치명"), "입력형 기입 줄이 `1)` 로 시작 → Q-05 치명")

        s = copy.deepcopy(spec)
        _item(s, "A2")["scale"] = {"preset": "만족도"}                     # 없는 preset — A2-1 BASE 가 A2 를 가리킨다
        try:
            ok_preset = _has(cq.check(s, None, None), "Q-05", "A2", "치명")
        except Exception:
            ok_preset = False
        expect(ok_preset, "없는 척도 preset → 멈추지 않고 Q-05 치명")
        badj = tmp / "bad.json"
        badj.write_text("{ 'parts': ", encoding="utf-8")
        rc, err = _cli(cq.main, [str(badj)])
        expect(rc == 2 and "명세 오류" in err and "Traceback" not in err, f"명세 JSON 을 못 읽으면 한 줄 오류 · exit 2 (rc {rc})")

        sc = {"parts": [{"code": "B", "items": []}]}
        its, want = sc["parts"][0]["items"], set()
        for name, labs in bq.SCALE_PRESETS.items():                      # 프리셋마다 역순
            its.append({"code": f"B{len(its) + 1}", "kind": "scale", "text": name, "scale": {"labels": labs[::-1]}})
            want.add(its[-1]["code"])
        its.append({"code": f"B{len(its) + 1}", "kind": "matrix", "text": "역순 행렬", "rows": ["가", "나"],
                    "scale": {"labels": bq.SCALE_PRESETS["추천"][::-1]}})
        want.add(its[-1]["code"])
        for labs in (bq.SCALE_PRESETS["추천"], ["매우 그렇지 않다", "그렇지 않다", "보통이다", "그렇다", "매우 그렇다"],
                     ["매우 적다", "적다", "보통", "많다", "매우 많다"], ["매우 어렵다", "어렵다", "보통", "쉽다", "매우 쉽다"]):
            its.append({"code": f"B{len(its) + 1}", "kind": "scale", "text": "정방향", "scale": {"labels": labs}})
        expect(cq._reversed_scale(["매우 잘 안다", "대체로 안다", "보통", "별로 모르겠다", "전혀 모르겠다"])
               and not cq._reversed_scale(["1점", "2점", "3점", "4점", "잘 모르겠다"]),
               "Q-08 — `전혀 모르겠다` 는 인지 척도의 끝점, 라벨 전체가 `잘 모르겠다` 인 것만 모름으로 뺌")
        got = _where(cq.check(sc, None, None), "Q-08")
        expect(got == want, f"척도 방향 — 프리셋 {len(bq.SCALE_PRESETS)}종 역순·역순 행렬은 Q-08, "
                            "`매우 그렇지 않다` 시작·부정어 없는 척도는 통과",
               f"빠짐 {sorted(want - got)} · 과잉 {sorted(got - want)}")

        s = copy.deepcopy(spec)
        _insert_before(s, "B1", {"kind": "info", "base": "A1=1 | 현재 이용자", "text": "다음은 현재 이용자에게만 여쭙니다."})
        _insert_before(s, "A4", {"kind": "info", "base": "SQ1=2 | 여성", "text": "다음은 여성에게만 여쭙니다."})
        _item(s, "A5")["prog_before"] = ["A5~A6 은 연도별로 제시"]
        p = tmp / "roundtrip.docx"
        bq.Builder(copy.deepcopy(s)).build(p)
        F = cq.check(s, None, p)
        expect(_has(F, "Q-14", "B1", "치명") and "BASE 가 없는데" in _msg(F, "Q-14", "B1"),
               "안내 박스 BASE 가 BASE 없는 다음 문항에 붙음 → Q-14 치명", _msg(F, "Q-14", "B1"))
        expect(_has(F, "Q-14", "A4", "치명") and "왕복 불일치" in _msg(F, "Q-14", "A4"),
               "안내 박스 BASE 가 다음 문항 BASE 보다 먼저 붙음 → Q-14 치명", _msg(F, "Q-14", "A4"))
        expect(_has(F, "Q-14", "A5", "경고") and "A4 에 붙음" in _msg(F, "Q-14", "A5"),
               "BASE 없는 문항의 prog_before 가 앞 문항으로 → Q-14 경고", _msg(F, "Q-14", "A5"))

        s = copy.deepcopy(spec)
        s["meta"]["skipped_src"] = ["구조화!99 — 빨간 행(삭제 표시)", "구조화 ver3!*", "구조화!25"]
        structure = {"sheets": [
            {"name": "구조화", "kind": "structure", "rows": [
                {"ref": "구조화!3", "fields": {"code": "SQ1", "question": "성별"}},
                {"ref": "구조화!99", "fields": {"code": "Z1", "question": "빨간 행"}},
                {"ref": "구조화!2", "fields": {"code": "Z2", "question": "빠뜨린 행"}}]},
            {"name": "구조화 ver3", "kind": "structure", "rows": [
                {"ref": f"구조화 ver3!{i}", "fields": {"code": f"A{i}", "question": "이전 버전"}} for i in range(2, 6)]}]}
        got = _where(cq.check(s, structure, None), "Q-11")
        expect(got == {"구조화!2"}, "skipped_src — `ref — 사유` · `시트!*` 는 빼고, `구조화!25` 가 `구조화!2` 를 가리지 않음",
               sorted(got))

    # ── 빌더 ──────────────────────────────────────────────────────────────
    with section("빌더"):
        import docx
        edge = {
            "meta": {"title": "엣지", "cover": {"target": "만 20~59세 남녀", "extra": "■ 조사 기간 : 9월"},
                     "closing": {"gift": "1,000원 금액권", "items": ["이름", "휴대전화번호"]}},
            "parts": [
                {"code": "SQ", "items": [
                    {"code": "SQ1", "kind": "single", "text": "성별", "options": ["남성", "여성"], "cols": 2,
                     "prog": ["만 20~59세 외 조사 종료"]}]},
                {"code": "A", "title": "엣지", "items": [
                    {"code": "A1", "kind": "open", "text": "근무 기간과 사업자번호를 적어 주십시오.",
                     "lines": ["a. 근무 기간 : ___년 ___개월", "b. 사업자번호 : ___-__-_____", "________________"]},
                    {"code": "A2", "kind": "open", "text": "__밑줄__ 은 살아야 한다", "box": False},
                    {"code": "A3", "kind": "multi", "min": 2, "text": "최소 2개", "options": ["가", "나", "다", "라"]},
                    {"code": "A4", "kind": "single", "text": "파이핑 PROG", "options": ["가", "나"],
                     "prog": ["[#SQ1 응답값#] 보기별로 반복 질문"]}]}]}
        ep = tmp / "edge.docx"
        review = bq.Builder(copy.deepcopy(edge)).build(ep)
        pars = docx.Document(str(ep)).paragraphs
        texts = [x.text for x in pars]
        fatal = [f for f in cq.check(edge, None, ep) if f["severity"] == "치명"]
        expect(not fatal, f"빌더 엣지 명세 왕복 치명 0 ({len(fatal)})", fatal)
        expect("- 만 20~59세 남녀" in texts and "- 만" not in texts and "■ 조사 기간 : 9월" in texts,
               "표지 target · extra 를 문자열 하나로 줘도 한 줄")
        blanks = [x for x in pars if "___" in x.text]
        expect([x.text for x in blanks] == ["a. 근무 기간 : ___년 ___개월", "b. 사업자번호 : ___-__-_____", "________________"]
               and not any(r.underline for x in blanks for r in x.runs),
               "기입란 `___년 ___개월` · `________` 은 밑줄 표기로 삼키지 않음", [x.text for x in blanks])
        expect(any(r.text == "밑줄" and r.underline for x in pars for r in x.runs), "`__밑줄__` 은 밑줄")
        expect("[PROG: [#SQ1 응답값#] 보기별로 반복 질문]" in texts, "`[#…#]` 로 시작하는 PROG 도 [PROG: …] 로 감쌈")
        q = {x["code"]: x for x in pq.parse(ep)["questions"]}
        expect(q["A3"]["response_type"] == "복수" and "최소 2개 이상 선택" in q["A3"]["directives"]["prog"],
               "복수 min 만 → [모두 선택] + PROG 최소 2개 이상 선택", q["A3"]["directives"]["prog"])
        ed = q.get("ED") or {}
        expect(ed.get("response_type") == "예-아니오" and ed.get("n_options") == 2,
               "동의 표가 ED 보기(예·아니오)에 섞이지 않음", [o["text"] for o in ed.get("options", [])])
        box = next((t.get("text") or "" for t in ed.get("tables", []) if t["kind"] == "box"), "")
        grid = next((t for t in ed.get("tables", []) if t["kind"] == "input_grid"), {})
        labels = [x["label"] for x in grid.get("items", [])]
        progs = " ".join(ed.get("directives", {}).get("prog", []))
        expect("3. 개인정보 항목 : 이름, 휴대전화번호" in box and "거부할 권리" in box
               and labels == ["동의 여부", "이름", "휴대전화번호"] and "이름, 휴대전화번호 입력" in progs
               and any(r.startswith("ED:") for r in review),
               "동의 박스 — 고지 3항·입력 칸·PROG 가 같은 수집 항목, 거부권 한 줄, 확인 목록에 오름",
               [labels, progs, box[:80]])

    with section("빌더 — 넓은 행렬"):
        mx = {"meta": {"title": "넓은 행렬", "closing": False}, "parts": [{"code": "A", "title": "행렬", "items": [
            {"code": "A1", "kind": "matrix", "text": "11점 + 모름", "rows": ["가", "나"],
             "scale": {"labels": [f"{i}점" for i in range(11)], "extra": ["모름"]}},
            {"code": "A2", "kind": "matrix", "text": "선택 12열", "rows": ["가", "나"],
             "columns": [f"c{i}" for i in range(1, 13)]}]}]}
        mp = tmp / "matrix.docx"
        bq.Builder(copy.deepcopy(mx)).build(mp)
        q = {x["code"]: x for x in pq.parse(mp)["questions"]}
        expect(q["A1"]["response_type"] == "행렬" and (q["A1"]["scale"] or {}).get("points") == 12
               and q["A2"]["response_type"] == "행렬" and len((q["A2"]["matrix"] or {}).get("column_labels") or []) == 12,
               "11점+모름 · 12열 선택 행렬을 그리고 파서가 행렬로 읽음")
        _item(mx, "A2")["columns"] = [f"c{i}" for i in range(1, 17)]
        try:
            bq.Builder(copy.deepcopy(mx)).build(tmp / "wide.docx")
            raised = False
        except bq.SpecError:
            raised = True
        wj = tmp / "wide.json"
        wj.write_text(json.dumps(mx, ensure_ascii=False), encoding="utf-8")
        rc, err = _cli(bq.main, [str(wj), str(tmp / "wide.docx")])
        expect(raised and rc == 2 and "명세 오류" in err and "Traceback" not in err,
               f"16열 행렬 → 명세 오류 한 줄 · exit 2 (rc {rc})", err.strip())

    # ── 파서 ──────────────────────────────────────────────────────────────
    with section("파서"):
        import docx
        d = docx.Document()

        def tbl(rows):
            t = d.add_table(rows=len(rows), cols=max(len(r) for r in rows))
            for i, r in enumerate(rows):
                for j, c in enumerate(r):
                    if isinstance(c, list):                 # 칸 안 여러 단락 (실제 설문지의 흔한 꼴)
                        t.cell(i, j).text = c[0]
                        for more in c[1:]:
                            t.cell(i, j).add_paragraph(more)
                    else:
                        t.cell(i, j).text = c

        tbl([["Part A. 파서"]])
        d.add_paragraph("A1. 귀하의 성별은 무엇입니까? [1개 선택]")
        tbl([["1) 남성", "2) 여성", "3) 응답 거절"], ["", "", ""]])
        d.add_paragraph("A2. 아래 표를 보고 가장 최근 구매 시기를 선택해 주십시오. [1개 선택]")
        tbl([["최초 출시", "리뉴얼", "현재 모델"], ["2019", "2022", "2025"]])
        for x in ("1) 1년 이내", "2) 1~3년", "3) 3년 이상"):
            d.add_paragraph(x)
        d.add_paragraph("A3. 이용하시는 서비스를 모두 선택해 주십시오. [모두 선택]")
        tbl([["구분", "보기"], ["온라인", "1) 앱\t2) 웹"], ["오프라인", "1) 매장\t2) 전화"]])
        d.add_paragraph("A4. 전반적으로 만족하십니까? [1개 선택]")
        tbl([["매우 불만족", "불만족", "보통", "만족", "매우 만족"], ["1", "2", "3", "4", "5"]])
        d.add_paragraph("A5. 이용하시는 서비스를 모두 선택해 주십시오. [모두 선택]")
        tbl([["구분", "보기"], ["온라인", ["1) 앱", "2) 웹"]], ["오프라인", ["3) 매장", "4) 전화"]]])
        d.add_paragraph("A6. 추천 의향은? [1개 선택]")
        tbl([["1) 전혀 그렇지 않다", "2) 그렇지 않다", "3) 보통", "4) 그렇다", "5) 매우 그렇다"], ["", "", "", "", ""]])
        d.add_paragraph("A7. 업종을 선택해 주십시오. [1개 선택]")
        tbl([["대분류", "보기"], ["판매업[PROG: A1=1]", "1) 소매업"], ["판매업[PROG: A1=1]", "2) 도매업"],
             ["서비스업[PROG: A1=2]", "3) 숙박업"], ["기타", "4) 기타"]])
        d.add_paragraph("A8. 중요한 순서대로 최대 3개까지 선택해 주십시오. [최소 1개, 최대 3개 선택]")
        for i, x in enumerate(("가격", "품질", "디자인", "브랜드"), start=1):
            d.add_paragraph(f"{i}) {x}")
        d.add_paragraph("A9. 이용 서비스를 모두 선택해 주십시오. [모두 선택]")
        tbl([["구분", "보기"], ["온라인", "1) 앱 2) 웹"], ["오프라인", "3) 매장"]])
        d.add_paragraph("A10. 동의 정도는? [1개 선택]")
        tbl([["매우 그렇다", "그렇다", "보통", "아니다", "전혀 아니다"], ["5", "4", "3", "2", "1"]])
        d.add_paragraph("A11. 사용 중인 기종을 선택해 주십시오. [1개 선택]")
        tbl([["구분", "보기"], ["휴대폰", ["1) 갤럭시(S24) 시리즈", "2) 갤럭시 Z 폴드"]],
             ["가전", "3) 공기청정기(PM2.5)\t4) 가습기"], ["기타", "5) 모델(가24) 한정판"]])
        d.add_paragraph("A12. 아래 표를 보고 가장 최근 구매 시기를 선택해 주십시오. [1개 선택]")
        tbl([["재작년", "작년", "올해"], ["2024", "2025", "2026"]])
        for x in ("1) 재작년", "2) 작년", "3) 올해"):
            d.add_paragraph(x)
        d.add_paragraph("A13. 이용 경험이 있으십니까? [1개 선택]")
        tbl([["구분", "보기"], ["응답", ["1) 있음", "2) 없음", "9) 모름"]], ["기타", ["97) 기타", "99) 무응답"]]])
        d.add_paragraph("A14. 사용 중인 기종을 선택해 주십시오. [1개 선택]")
        tbl([["구분", "보기"], ["휴대폰", ["1) 갤럭시 S24", "2) 갤럭시 Z"]]])
        sp = tmp / "synth.docx"
        d.save(str(sp))
        q = {x["code"]: x for x in pq.parse(sp)["questions"]}

        def opts(c):
            return [o["text"] for o in q[c]["options"]]

        def anom(c):
            return " ".join(q[c]["anomalies"])

        expect(q["A1"]["response_type"] == "단일" and opts("A1") == ["남성", "여성", "응답 거절"],
               "번호 머리 표(척도 낱말 없음)는 척도가 아니라 표 보기", [q["A1"]["response_type"], opts("A1")])
        expect(q["A2"]["response_type"] == "단일" and q["A2"]["scale"] is None, "연도가 적힌 참고 표는 척도가 아님",
               q["A2"]["response_type"])
        expect(q["A12"]["response_type"] == "단일" and q["A12"]["scale"] is None,
               "이어진 연도(2024 2025 2026)도 0·1 에서 시작하지 않으면 척도가 아님", q["A12"]["response_type"])
        expect(all(q[c]["response_type"] == "척도" for c in ("A4", "A6", "A10")),
               "연속 번호(1~5 · 5~1)·척도 낱말 번호 머리 표는 그대로 척도", [q[c]["response_type"] for c in ("A4", "A6", "A10")])
        expect(opts("A5") == ["앱", "웹", "매장", "전화"] and "분리 의심" not in anom("A5"),
               "칸 안 단락이 붙은 `1) 앱2) 웹` → 보기 넷", opts("A5"))
        expect("분리 의심" in anom("A9"), "한 칸에 `1) 앱 2) 웹` (한 칸 띄움) → 표 보기 분리 의심 경고", opts("A9"))
        expect(opts("A11") == ["갤럭시(S24) 시리즈", "갤럭시 Z 폴드", "공기청정기(PM2.5)", "가습기", "모델(가24) 한정판"]
               and "분리 의심" not in anom("A11"),
               "보기 글자 속 `(S24)`·`(PM2.5)`·`(가24)` 는 끊지 않고, 붙은 `…시리즈2) 갤럭시` 는 끊음", [opts("A11"), anom("A11")])
        expect([(o["no"], o["text"]) for o in q["A13"]["options"]] == [(1, "있음"), (2, "없음"), (9, "모름"), (97, "기타"), (99, "무응답")]
               and "분리 의심" not in anom("A13"),
               "붙은 번호가 건너뛴 코드(`9) 모름`·`97) 기타`)여도 끊음", [[(o["no"], o["text"]) for o in q["A13"]["options"]], anom("A13")])
        expect("분리 의심" in anom("A14"), "숫자로 끝난 보기에 붙은 번호(`갤럭시 S242) 갤럭시 Z`)는 못 끊어도 분리 의심 경고",
               [opts("A14"), anom("A14")])
        g = copy.deepcopy(spec)
        grp = _item(g, "A7")["options_table"][1]["options"]
        grp[0], grp[1] = "브라질(S24) 산토스", "콜롬비아(PM2.5)"
        gp = tmp / "glue.docx"
        bq.Builder(copy.deepcopy(g)).build(gp)
        fatal = [f for f in cq.check(g, None, gp) if f["severity"] == "치명"]
        expect(not fatal, "표 보기에 `(S24)`·`(PM2.5)` 가 있어도 빌더 → 파서 왕복 치명 0", fatal)
        expect(opts("A3") == ["앱", "웹", "매장", "전화"] and [o["no"] for o in q["A3"]["options"]] == [1, 2, 3, 4]
               and "번호 중복" in anom("A3"), "묶음마다 1) 부터 다시 매긴 표 → 모두 두고 새 번호 + 경고", opts("A3"))
        expect([o.get("show_if") for o in q["A7"]["options"]] == [["A1=1"], ["A1=1"], ["A1=2"], None],
               "표 보기 행의 [PROG: …] → 보기 show_if", [o.get("show_if") for o in q["A7"]["options"]])
        lim = q["A8"].get("multi_limit") or {}
        expect(q["A8"]["response_type"] == "복수" and (lim.get("min"), lim.get("max")) == (1, 3) and "순서대로" in anom("A8"),
               "[최소 1개, 최대 3개 선택] → multi_limit, `순서대로` 면 순위 의심 경고", [lim, anom("A8")])

    # ── 구조화 읽기 ───────────────────────────────────────────────────────
    with section("구조화 읽기"):
        import openpyxl
        from openpyxl.styles import Font, PatternFill
        from openpyxl.styles.colors import Color

        import read_structure as rs
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        def sheet(name, rows, merges=()):
            ws = wb.create_sheet(name)
            for r, row in enumerate(rows, start=1):
                for c, v in enumerate(row, start=1):
                    if v is not None:
                        ws.cell(r, c, v)
            for m in merges:
                ws.merge_cells(m)
            return ws

        sheet("무앵커", [["중분류", "소분류", "문항", "보기", "응답대상"],
                       ["인지", "브랜드 인지", "아는 브랜드", "목록", "all"],
                       ["이용", "주 이용", "주 이용 브랜드", "목록", "all"],
                       [None] * 5,
                       ["평가", "전반 만족", "전반 만족도", "5점", "all"],
                       ["불만", "불만족 이유", "주관식", None, "불만족 시"]])
        sheet("병합초안", [["대분류", "문번", "문항", "보기", "응답대상"],
                        ["PART1", None, "성별", "남/여", "all"],
                        ["PART1", None, "연령", "주관식", "all"],
                        [None] * 5,
                        ["PART2", None, "전반 만족도", "5점", "all"],
                        [None, None, "추천 의향", "5점", "all"],
                        [None, None, "불만 이유", "주관식", "불만족 시"]], merges=("A5:A7",))
        sheet("행렬", [["파트", "중분류", "속성", "문번", "문항", "보기", "응답대상"],
                     ["PART1", "인지", None, "A1", "아는 브랜드", "목록", "all"],
                     [None] * 7,
                     ["PART2", "속성 만족", "맛", "B1", "속성별 만족도", "5점", "이용자"],
                     [None, None, "가격", None, None, None, None],
                     [None, None, "매장", None, None, None, None],
                     [None] * 7,
                     [None, None, None, None, "(대안) 속성별로 얼마나 만족하십니까?", None, None],
                     [None, None, None, None, "(대안) 다음 항목에 대한 만족도", None, None]])
        sheet("합친머리", [["No", "중분류", None, "문항", "보기", "응답대상"],
                        [1, "서비스", "속도", "서비스 속도 만족", "5점", "all"]], merges=("B1:C1",))
        sheet("빈머리", [["No", "중분류", None, "문항", "보기", "응답대상"],
                       [1, "서비스", "속도", "서비스 속도 만족", "5점", "all"]])
        sheet("같은머리", [["문번", "문항", "보기", "보기", "보기", "응답대상"],
                        ["A1", "주 이용 브랜드", "스타벅스", "투썸", "이디야", "all"]])
        sheet("곁표", [["No", "중분류", "문항", "응답대상", None, "구분", "보기"],
                     [1, "국가", "방문 국가", "all", None, "아시아", "일본, 중국"]])
        ws = sheet("색", [["문번", "문항", "보기", "응답대상"]] + [[f"A{i}", f"행 {i}", "5점", "all"] for i in range(1, 10)])
        ws["B2"].fill = PatternFill("solid", fgColor="FFFFFF00")                # RGB 노랑
        ws["B3"].fill = PatternFill("solid", fgColor=Color(theme=5, tint=0.4))  # 테마색 + 명암
        ws["B4"].fill = PatternFill("solid", fgColor=Color(indexed=13))         # 인덱스색 (.xls 변환본)
        ws["B5"].font = Font(color="FFFF0000")                                  # 빨간 글자
        ws["B6"].fill = PatternFill("solid", fgColor=Color(indexed=64))         # 시스템 기본 — 뺀다
        ws["B7"].fill = PatternFill("solid", fgColor=Color(theme=0))            # 흰 테마색 — 뺀다
        ws["B8"].font = Font(color="FF333333")                                  # 회색 글자 — 뺀다
        ws["B9"].font = Font(color=Color(theme=1, tint=0.35))                   # 기본 글자색 옅게 — 뺀다
        ws["B10"].font = Font(color=Color(theme=4))                             # 강조 테마 글자 — 남긴다
        xp = tmp / "structure.xlsx"
        wb.save(str(xp))
        doc = rs.read(xp)
        sh = {x["name"]: x for x in doc["sheets"]}

        def refs(name):
            return [r["ref"] for r in sh[name]["rows"]]

        def below(name):
            return [x["ref"] for x in sh[name].get("below_lines", [])]

        def row(ref):
            return next(r for x in doc["sheets"] for r in x.get("rows", []) if r["ref"] == ref)

        expect(refs("무앵커") == ["무앵커!2", "무앵커!3", "무앵커!5", "무앵커!6"] and not below("무앵커"),
               "번호·파트 열 없는 시트 — 빈 행 뒤 행도 본표", [refs("무앵커"), below("무앵커")])
        expect({"병합초안!6", "병합초안!7"} <= set(refs("병합초안")) and row("병합초안!7")["fields"].get("part") == "PART2",
               "빈 행 뒤 세로 병합 대분류(문번 빈 초안) — 이어지는 행도 본표", refs("병합초안"))
        expect({"행렬!4", "행렬!5", "행렬!6"} <= set(refs("행렬")) and row("행렬!5")["fields"].get("attribute") == "가격"
               and below("행렬") == ["행렬!8", "행렬!9"],
               "빈 행 뒤 행렬 이어 행(속성만)은 본표, 한 칸짜리 대안 문구는 본표 아래 메모", [refs("행렬"), below("행렬")])
        f = row("합친머리!2")["fields"]
        expect(f.get("section") == "서비스 | 속도" and f.get("question") and f.get("options") and f.get("base")
               and not sh["합친머리"].get("side_lines"), "가로 병합 머리(중분류 B:C) — 같은 역할로 잇고 문항·보기·대상 유지", f)
        r = row("빈머리!2")
        expect(r["fields"].get("question") and r["fields"].get("options") and r["fields"].get("base")
               and r["extra"].get("col3") == "속도", "빈 머리 칸이 있어도 문항 열까지는 본표", r)
        expect(row("같은머리!2")["fields"].get("options") == "스타벅스 | 투썸 | 이디야",
               "같은 머리말 열(보기|보기|보기) → 값을 ` | ` 로 이음", row("같은머리!2")["fields"])
        side = " ".join(x["text"] for x in sh["곁표"].get("side_lines", []))
        expect("아시아" in side and "options" not in row("곁표!2")["fields"], "빈 열 뒤 곁표(구분|보기)는 곁표로 남음")
        fills = {r["ref"]: r["fills"].get("문항") for r in sh["색"]["rows"]}
        want = {"색!2": "FFFF00", "색!3": "theme5+0.4", "색!4": "FFFF00", "색!5": "글자FF0000", "색!6": None,
                "색!7": None, "색!8": None, "색!9": None, "색!10": "글자theme4"}
        expect(fills == want, "셀 색 — RGB·테마·인덱스 배경과 강조 글자색은 남기고 흰 배경·검정/회색 글자는 뺌", fills)
        mini = {"parts": [{"code": "A", "items": [
            {"code": f"A{i}", "kind": "single", "text": t, "options": ["가", "나"], "src": [ref]}
            for i, (t, ref) in enumerate((("성별", "병합초안!2"), ("연령", "병합초안!3"), ("만족", "병합초안!5")), start=1)]}]}
        got = {w for w in _where(cq.check(mini, doc, None), "Q-11") if w.startswith("병합초안!")}
        expect(got == {"병합초안!6", "병합초안!7"}, "읽은 구조화 → 검수: 빈 행 뒤 행을 빠뜨리면 Q-11", sorted(got))

    # ── 연계: 파서 결과를 쓰는 도구 · 작업 폴더 ─────────────────────────────
    with section("연계 — 웹테스트"):
        import webtest_compare as wc
        import webtest_scenarios as wsn
        tq = {"options": [{"text": "소매업", "show_if": ["A1=1"]}, {"text": "도매업", "show_if": ["A1=1"]},
                          {"text": "숙박업", "show_if": ["A1=2"]}, {"text": "기타"}]}

        def page(*labels):
            return {"options": [{"kind": "radio", "label": x} for x in labels]}

        def missing(res, sev):
            return [x for x in res if x["kind"] == "OPTION_MISSING" and x["sev"] == sev]

        r1 = wc.compare_options(tq, page("1) 소매업", "2) 도매업", "4) 기타"))
        r2 = wc.compare_options(tq, page("1) 소매업", "4) 기타"))
        expect(not missing(r1, "경고") and len(missing(r1, "참고")) == 1 and "판정하지 못함" in missing(r1, "참고")[0]["detail"],
               "답 없이(판정 못 함) 표시 조건째 숨은 표 보기 → 조건별 참고 한 줄 (경고 아님)", r1)
        expect(len(missing(r2, "경고")) == 1 and "도매업" in missing(r2, "경고")[0]["detail"],
               "같은 조건 보기 중 하나만 빠지면 → 경고", r2)
        a1 = {"A1": {"choose": [1]}}
        r3 = wc.compare_options(tq, page("1) 소매업", "2) 도매업", "4) 기타"), a1)
        r4 = wc.compare_options(tq, page("3) 숙박업", "4) 기타"), a1)
        shown = [x for x in r4 if x["kind"] == "OPTION_EXTRA" and x["sev"] == "경고" and "[PROG: A1=2]" in x["detail"]]
        expect(not r3, "답(A1=1)에 맞는 묶음만 보이고 안 맞는 묶음은 숨김 → 지적 없음", r3)
        expect(len(missing(r4, "경고")) == 2 and len(shown) == 1,
               "답(A1=1)에 맞는 묶음이 없으면 누락 경고, 안 맞는 묶음(A1=2)이 보이면 경고", r4)
        tq_dup = {"options": [{"text": "소매업", "show_if": ["A1=1"]}, {"text": "기타", "show_if": ["A1=1"]},
                              {"text": "숙박업", "show_if": ["A1=2"]}, {"text": "기타", "show_if": ["A1=2"]}]}
        r5 = wc.compare_options(tq_dup, page("숙박업", "기타"), {"A1": {"choose": [2]}})
        expect(not r5, "묶음마다 같은 「기타」가 있어도 조건에 맞는 쪽과 짝지음 → 맞는 페이지는 지적 없음", r5)
        tq_no = {"response_type": "단일", "options": [dict(o, no=i) for i, o in enumerate(tq["options"], start=1)]}
        expect(wsn.default_answer(tq_no, answers={"A1": {"choose": [2]}}) == {"choose": [3]}
               and wsn.default_answer(tq_no) == {"choose": [1]},
               "기본 답이 앞 답(A1=2)으로 숨는 표 보기(show_if A1=1)를 피함")
        mini = [{"code": "A1", "response_type": "단일", "options": [{"no": 1, "text": "판매"}, {"no": 2, "text": "서비스"}]},
                dict(tq_no, code="A2"),
                {"code": "A3", "response_type": "단일", "base_condition": "A1=1", "options": [{"no": 1, "text": "가"}, {"no": 2, "text": "나"}]}]
        by_id = {x["id"]: x["answers"] for x in wsn.build(mini, 12)["scenarios"]}
        expect(by_id["S0"]["A2"] == {"choose": [1]} and by_id.get("H1_A1", {}).get("A2") == {"choose": [3]},
               "숨김 경로에서 원천 답(A1)을 바꾸면 딸린 표 보기 답(A2)도 보이는 보기로 다시 고름", by_id)
        six = [{"no": i, "text": t} for i, t in enumerate("가나다라마바", start=1)]

        def n_pick(lim):
            return len(wsn.default_answer({"response_type": "복수", "options": six, "multi_limit": lim})["choose"])

        expect((n_pick({"min": 3, "max": 5}), n_pick({"min": 1, "max": 1}), n_pick(None)) == (3, 1, 2),
               "복수 기본 답이 multi_limit 범위 안 (최소 3 → 3개 · 최대 1 → 1개 · 제한 없음 → 2개)")
        qs = pq.parse(out)["questions"]                 # 모의 명세 docx

        def h_src(doc):
            return [sc["id"].split("_", 1)[1] for sc in doc["scenarios"] if sc["id"].startswith("H")]

        full = wsn.build(qs, 100)
        srcs = set(h_src(full))
        cap = wsn.build(qs, len(srcs))
        one = wsn.build(qs, 1)["meta"]
        expect(full["meta"]["h_candidates"] > len(srcs) and full["meta"]["h_skipped"] == 0
               and set(h_src(cap)) == srcs and cap["meta"]["h_skipped"] == full["meta"]["h_candidates"] - len(srcs)
               and one["h_skipped"] == one["h_candidates"] - 1,
               "숨김 경로 상한 — 원천 문항마다 하나씩 먼저 고르고, 뺀 수를 meta.h_skipped 로 알림",
               [full["meta"], h_src(full), h_src(cap)])

    with section("연계 — 제작가이드·작업 폴더"):
        import build_facts as bf
        expect(all(bf.RESPONSE_MARKER.sub(" ", f"중요한 것을 고르십시오. {m}").strip() == "중요한 것을 고르십시오."
                   for m in ("[최소 1개, 최대 3개 선택]", "[최대 3개 선택]", "[최소1개,최대3개 선택]")),
               "제작가이드 Qu: 에서 개수 제한 표기를 뗌")
        import capture_edit as ce
        import paths as pth
        kinds_ = [ce._classify(Path(x)) for x in ("projects/P/01_A/03_설문지/work/설문명세.json",
                                                  "projects/P/01_A/input/work/설문명세.json",
                                                  "projects/P/01_A/03_설문지/A_구조화_260918.xlsx")]
        expect(kinds_ == ["questionnaire", "questionnaire", "input"],
               "편집 기록 — 03_설문지/work(v1 input/work)는 questionnaire, 받은 구조화 파일은 input", kinds_)

        def v2_role(name):
            return next((rl for rx, rl in pth.RE_INPUT_SPLIT if rx.search(name)), None)

        roles_ = [v2_role(x) for x in ("work", "A_구조화_260918.xlsx", "원자료.xlsx", "workbook.xlsx")]
        expect(roles_ == ["설문지_dir", "설문지_dir", "실사", None], "v1→v2 옮김 규칙 — input/work · 구조화 파일은 03_설문지 로", roles_)
        sv = tmp / "P" / "01_A"                         # v1 조사 — /설문지작성 은 설문지 폴더(input)/work 에 쓴다
        (sv / "input" / "work" / "검수").mkdir(parents=True)
        (sv / "1_제작가이드").mkdir()
        (sv / "input" / "work" / "설문명세.json").write_text("{}", encoding="utf-8")
        (sv / "input" / "A_구조화_260918.xlsx").write_bytes(b"")
        before = pth.resolve(sv)
        with contextlib.redirect_stdout(io.StringIO()):
            pth.cmd_migrate(str(sv), True)
        after = pth.resolve(sv)
        expect(before["layout"] == "v1" and before["dirs"]["설문지_dir"].endswith("/input")
               and (before["files"]["설문명세"] or "").endswith("/input/work/설문명세.json")
               and after["layout"] == "v2" and (after["files"]["설문명세"] or "").endswith("/03_설문지/work/설문명세.json")
               and (after["files"]["구조화"] or "").endswith("/03_설문지/A_구조화_260918.xlsx"),
               "v1 조사 — input/work 를 써도 v1 그대로, migrate 하면 input/work · 구조화가 03_설문지 로",
               [before["layout"], before["files"]["설문명세"], after["layout"], after["files"]["설문명세"], after["files"]["구조화"]])

    if a.keep:
        print(f"산출 {tmp} (남김)")
    else:
        shutil.rmtree(tmp, ignore_errors=True)
    print("=== 전부 통과 ===" if ok else "=== 실패 있음 ===")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
