#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""테이블 검수 결과(findings.json)를 엑셀 리포트로 변환한다.

시트 구성은 고정이다 (references/report-schema.md 참조).
입력 키가 없으면 해당 시트는 "해당 없음" 한 줄로 남긴다 — 시트를 빼지 않는다.
검수하지 않은 것과 검수했는데 없는 것을 구분해야 하기 때문이다.
"""

import argparse
import json
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.stderr.write("오류: openpyxl이 없습니다. `pip install openpyxl` 후 다시 실행하세요.\n")
    sys.exit(1)

HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(color="FFFFFF", bold=True, size=10)
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
LABEL_FILL = PatternFill("solid", fgColor="D9E2F3")
BODY_FONT = Font(size=10)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

VERDICT_FILL = {
    "적합": PatternFill("solid", fgColor="E2EFDA"),
    "이상": PatternFill("solid", fgColor="FCE4E4"),
    "확인필요": PatternFill("solid", fgColor="FFF2CC"),
    "참고": PatternFill("solid", fgColor="F2F2F2"),
    "미검증": PatternFill("solid", fgColor="F2F2F2"),
}
VERDICT_FONT = {
    "적합": Font(size=10, color="375623"),
    "이상": Font(size=10, color="C00000", bold=True),
    "확인필요": Font(size=10, color="7F6000", bold=True),
    "참고": Font(size=10, color="595959"),
    "미검증": Font(size=10, color="595959"),
}
SEVERITY_ORDER = {"치명": 0, "중요": 1, "경미": 2, "확인필요": 3}

# (시트명, findings 키, [(헤더, 필드, 너비), ...])
SHEETS = [
    ("01_문항별검수", "questions", [
        ("No", "_no", 6), ("문번", "code", 12), ("문항 문구", "title", 46),
        ("응답유형", "type", 12), ("설문지 Base 정의", "base_def", 30),
        ("테이블 Base 표기", "base_label", 30), ("기대 N", "n_expected", 10),
        ("테이블 N", "n_actual", 10), ("보기수(설문지)", "opts_q", 13),
        ("보기수(테이블)", "opts_tbl", 13), ("판정", "verdict", 11), ("비고", "note", 44),
    ]),
    ("02_이슈목록", "issues", [
        ("No", "_no", 6), ("심각도", "severity", 10), ("문번", "code", 12),
        ("유형", "category", 16), ("현상", "symptom", 50),
        ("근거(설문지/가이드)", "evidence", 40), ("조치 요청", "action", 44),
        ("상태", "status", 12),
    ]),
    ("03_쿼터확인", "quota", [
        ("구분", "group", 16), ("셀", "cell", 22), ("목표", "target", 10),
        ("실제", "actual", 10), ("차이", "diff", 10), ("판정", "verdict", 11),
        ("비고", "note", 40),
    ]),
    ("04_Base로직", "base_logic", [
        ("문번", "code", 12), ("로직 구분", "kind", 12),
        ("설문지 로직 문장", "logic", 48), ("근거 위치", "source", 20),
        ("기대 Base", "expected", 14), ("테이블 N", "actual", 12),
        ("판정", "verdict", 11), ("비고", "note", 40),
    ]),
    ("05_행열구조", "structure", [
        ("문번", "code", 12), ("가이드 행 구성", "guide_rows", 34),
        ("테이블 행 구성", "tbl_rows", 34), ("행 판정", "row_verdict", 11),
        ("가이드 열(배너)", "guide_cols", 30), ("테이블 열(배너)", "tbl_cols", 30),
        ("열 판정", "col_verdict", 11), ("비고", "note", 36),
    ]),
    ("06_보기코드대조", "codes", [
        ("문번", "code", 12), ("보기코드", "opt_code", 10),
        ("설문지 보기", "q_label", 40), ("테이블 행 라벨", "tbl_label", 40),
        ("판정", "verdict", 11), ("비고", "note", 36),
    ]),
    ("07_미출력보기", "zero_opts", [
        ("문번", "code", 12), ("보기코드", "opt_code", 10), ("보기 문구", "label", 40),
        ("해당 셀 값", "value", 12), ("판정", "verdict", 11), ("비고", "note", 44),
    ]),
]

CHECKLIST = [
    ("quota", "1. 쿼터 일치 확인 (응답자 특성)"),
    ("base", "2. Base 확인 (모든 문항)"),
    ("logic", "3. 로직 Check (base/prog 반영)"),
    ("code", "4. Code 일치 확인"),
    ("numbering", "5. 문번 확인 (넘버링 누락)"),
    ("structure", "6. 행·열 구조 / 보기 중복·누락"),
]


def die(msg):
    sys.stderr.write("오류: " + msg + "\n")
    sys.exit(1)


def as_text(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "Y" if v else "N"
    if isinstance(v, (list, tuple)):
        return ", ".join(as_text(x) for x in v)
    if isinstance(v, dict):
        return json.dumps(v, ensure_ascii=False)
    return v


def style_head(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def write_grid(ws, cols, rows, start_row):
    """헤더 + 데이터. 판정 계열 컬럼에는 색을 입힌다."""
    for i, (head, _, width) in enumerate(cols, start=1):
        ws.cell(row=start_row, column=i, value=head)
        ws.column_dimensions[get_column_letter(i)].width = width
    style_head(ws, start_row, len(cols))

    if not rows:
        ws.cell(row=start_row + 1, column=1, value="해당 없음 (검수 결과 항목 없음)")
        ws.cell(row=start_row + 1, column=1).font = Font(size=10, italic=True, color="808080")
        return start_row + 1

    r = start_row
    for n, item in enumerate(rows, start=1):
        r += 1
        for i, (_, field, _w) in enumerate(cols, start=1):
            value = n if field == "_no" else as_text(item.get(field))
            cell = ws.cell(row=r, column=i, value=value)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if field.endswith("verdict") or field == "severity":
                key = str(value).strip()
                if field == "severity":
                    key = {"치명": "이상", "중요": "이상", "경미": "확인필요"}.get(key, key)
                if key in VERDICT_FILL:
                    cell.fill = VERDICT_FILL[key]
                    cell.font = VERDICT_FONT[key]
                cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    ws.auto_filter.ref = "A%d:%s%d" % (start_row, get_column_letter(len(cols)), r)
    return r


def count_by(rows, field, value):
    return sum(1 for r in rows if str(r.get(field, "")).strip() == value)


def sheet_summary(ws, data):
    project = data.get("project", {}) or {}
    issues = data.get("issues", []) or []
    questions = data.get("questions", []) or []
    summary = data.get("summary", {}) or {}

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 62
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 62

    ws["A1"] = "통계표(Table) 검수 결과"
    ws["A1"].font = TITLE_FONT
    ws.row_dimensions[1].height = 26

    info = [
        ("프로젝트", project.get("name")),
        ("고객사", project.get("client")),
        ("전체 표본수", project.get("total_n")),
        ("설문지 파일", project.get("questionnaire_file")),
        ("테이블가이드 파일", project.get("guide_file")),
        ("테이블 파일", project.get("table_file")),
        ("검수자", project.get("checked_by")),
        ("검수일", project.get("checked_at")),
        ("가중치 적용", project.get("weighted")),
    ]
    r = 3
    for label, value in info:
        ws.cell(row=r, column=1, value=label).fill = LABEL_FILL
        ws.cell(row=r, column=1).font = Font(size=10, bold=True)
        ws.cell(row=r, column=1).border = BORDER
        cell = ws.cell(row=r, column=2, value=as_text(value) or "-")
        cell.font = BODY_FONT
        cell.border = BORDER
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="검수 체크리스트")
    ws.cell(row=r, column=1).font = Font(size=12, bold=True, color="1F3864")
    r += 1
    head_row = r
    for i, h in enumerate(["항목", "결과", "이슈 건수", "비고"], start=1):
        ws.cell(row=head_row, column=i, value=h)
    style_head(ws, head_row, 4)
    checks = {c.get("key"): c for c in (summary.get("checks") or [])}
    for key, name in CHECKLIST:
        r += 1
        c = checks.get(key, {})
        ws.cell(row=r, column=1, value=c.get("name") or name).font = BODY_FONT
        verdict = as_text(c.get("result")) or "미검증"
        vc = ws.cell(row=r, column=2, value=verdict)
        if verdict in VERDICT_FILL:
            vc.fill = VERDICT_FILL[verdict]
            vc.font = VERDICT_FONT[verdict]
        else:
            vc.font = BODY_FONT
        vc.alignment = Alignment(horizontal="center", vertical="center")
        cnt = c.get("count")
        if cnt is None:
            cnt = count_by(issues, "check", key)
        ws.cell(row=r, column=3, value=cnt).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=3).font = BODY_FONT
        nc = ws.cell(row=r, column=4, value=as_text(c.get("note")))
        nc.font = BODY_FONT
        nc.alignment = Alignment(vertical="top", wrap_text=True)
        for i in range(1, 5):
            ws.cell(row=r, column=i).border = BORDER

    r += 2
    ws.cell(row=r, column=1, value="집계")
    ws.cell(row=r, column=1).font = Font(size=12, bold=True, color="1F3864")
    r += 1
    stats = [
        ("검수 문항 수", len(questions)),
        ("적합", count_by(questions, "verdict", "적합")),
        ("이상", count_by(questions, "verdict", "이상")),
        ("확인필요", count_by(questions, "verdict", "확인필요")),
        ("이슈 - 치명", count_by(issues, "severity", "치명")),
        ("이슈 - 중요", count_by(issues, "severity", "중요")),
        ("이슈 - 경미", count_by(issues, "severity", "경미")),
        ("이슈 합계", len(issues)),
        ("0건 미출력 보기 (참고)", len(data.get("zero_opts", []) or [])),
        ("오탐 제외 건수", (data.get("false_positives") or {}).get("count", 0)),
    ]
    for label, value in stats:
        ws.cell(row=r, column=1, value=label).fill = LABEL_FILL
        ws.cell(row=r, column=1).font = Font(size=10, bold=True)
        ws.cell(row=r, column=1).border = BORDER
        ws.cell(row=r, column=2, value=value).font = BODY_FONT
        ws.cell(row=r, column=2).border = BORDER
        r += 1

    note = summary.get("overall") or (data.get("false_positives") or {}).get("note")
    if note:
        r += 1
        ws.cell(row=r, column=1, value="종합 의견").fill = LABEL_FILL
        ws.cell(row=r, column=1).font = Font(size=10, bold=True)
        ws.cell(row=r, column=1).border = BORDER
        cell = ws.cell(row=r, column=2, value=as_text(note))
        cell.font = BODY_FONT
        cell.border = BORDER
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = 60


def sheet_scope(ws, data):
    scope = data.get("scope", {}) or {}
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 70

    ws["A1"] = "검증 범위"
    ws["A1"].font = TITLE_FONT
    r = 3
    for title, key, tail in (("검증한 항목", "verified", "확인 방법"),
                             ("검증하지 못한 항목", "not_verified", "사유")):
        ws.cell(row=r, column=1, value=title)
        ws.cell(row=r, column=1).font = Font(size=12, bold=True, color="1F3864")
        r += 1
        for i, h in enumerate(["No", "항목", tail], start=1):
            ws.cell(row=r, column=i, value=h)
        style_head(ws, r, 3)
        items = scope.get(key) or []
        if not items:
            r += 1
            ws.cell(row=r, column=1, value="-")
            ws.cell(row=r, column=2, value="기재 없음").font = Font(size=10, italic=True, color="808080")
        for n, item in enumerate(items, start=1):
            r += 1
            if isinstance(item, str):
                item = {"item": item}
            ws.cell(row=r, column=1, value=n).font = BODY_FONT
            ws.cell(row=r, column=2, value=as_text(item.get("item"))).font = BODY_FONT
            ws.cell(row=r, column=3, value=as_text(item.get("how") or item.get("reason"))).font = BODY_FONT
            for i in range(1, 4):
                c = ws.cell(row=r, column=i)
                c.border = BORDER
                c.alignment = Alignment(vertical="top", wrap_text=True)
        r += 3


def build(data, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "00_검수요약"
    sheet_summary(ws, data)

    for name, key, cols in SHEETS:
        sheet = wb.create_sheet(name)
        rows = data.get(key) or []
        if key == "issues":
            rows = sorted(rows, key=lambda x: SEVERITY_ORDER.get(str(x.get("severity", "")).strip(), 9))
        write_grid(sheet, cols, rows, 1)

    sheet_scope(wb.create_sheet("08_검증범위"), data)

    try:
        wb.save(out_path)
    except PermissionError:
        die("파일이 열려 있어 저장하지 못했습니다. Excel에서 닫고 다시 실행하세요: " + out_path)
    except Exception as exc:
        die("저장하지 못했습니다 (%s): %s" % (type(exc).__name__, exc))


def main():
    p = argparse.ArgumentParser(
        prog="make_report.py",
        description="테이블 검수 결과(findings.json)를 엑셀 리포트로 변환한다.",
        epilog=("예시:\n"
                "  python make_report.py findings.json -o 테이블검수_프로젝트_20260826.xlsx\n\n"
                "입력 스키마는 references/report-schema.md, 예시는 assets/findings_example.json."),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("findings", help="검수 결과 JSON 파일")
    p.add_argument("-o", "--out", help="출력 .xlsx 경로 (기본: 입력파일명.xlsx)")
    args = p.parse_args()

    src = os.path.abspath(args.findings)
    if not os.path.exists(src):
        die("파일을 찾을 수 없습니다: " + src)
    try:
        with open(src, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except json.JSONDecodeError as exc:
        die("JSON 형식이 아닙니다 (%s행): %s" % (exc.lineno, exc.msg))
    if not isinstance(data, dict):
        die("최상위가 객체(JSON object)여야 합니다.")

    out = os.path.abspath(args.out) if args.out else os.path.splitext(src)[0] + ".xlsx"
    if not out.lower().endswith((".xlsx", ".xlsm")):
        out += ".xlsx"
    parent = os.path.dirname(out)
    if parent and not os.path.isdir(parent):
        die("출력 폴더가 없습니다: " + parent)

    missing = [k for k in ("project", "questions") if not data.get(k)]
    if missing:
        sys.stderr.write("경고: 비어 있는 키가 있습니다 — %s (해당 시트는 '해당 없음'으로 남습니다)\n"
                         % ", ".join(missing))

    build(data, out)
    print("저장: %s" % out)
    print("문항 %d개 / 이슈 %d건" % (len(data.get("questions") or []), len(data.get("issues") or [])))


if __name__ == "__main__":
    main()
