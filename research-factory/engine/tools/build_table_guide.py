# -*- coding: utf-8 -*-
"""테이블가이드(.xlsx) 빌더 — DP 발주서를 통계표·설문지에서 역산해 만든다.

집을 짓고 나서 설계도를 다시 그리는 일에 해당한다. 통계표가 이미 입고된 조사에서
DP 가 실제로 무엇을 산출했는지(표 목록·베이스·통계값·배너 구성·쿼터)를 한 장의
발주서 형식으로 복원하고, 복원 근거와 추가 산출 요청을 같은 파일에 함께 남긴다.

사용
    python build_table_guide.py <tables.json> <questionnaire.json> <out.xlsx>
        --config <tableguide_config.json> --template <TEMPLATE.xlsx>

서식은 템플릿(같은 연구원·DP 가 쓴 기존 테이블가이드)을 복제해 그대로 물려받는다.
글꼴·채움·테두리·열 너비를 손으로 다시 만들지 않는다.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 시트/좌표 상수 (템플릿 구조)
SH_BANNER, SH_QUOTA, SH_BASIC = "Banner", "Quota", "Basic Table"
BANNER_FIRST_ROW = 14          # 배너 정의 첫 행
BANNER_NOTE_LABEL = 26         # '유의사항' 라벨 행 (A26:A32 병합)
BANNER_NOTE_LAST = 32
BANNER_ADD_LABEL = 33          # '추가사항' 라벨 행 (A33:A39 병합)
BANNER_ADD_LAST = 39
BASIC_FIRST_ROW = 3            # 표1 -> 3행. 표N -> (N+2)행
SH_ADD, SH_TRACE = "Add Table", "복원 근거"


def stat_spec(table):
    """표의 실제 행 구성에서 통계값 문구를 도출한다."""
    hints = table.get("hints") or {}
    kinds = set(r.get("kind") for r in table.get("rows", []))
    if hints.get("is_summary"):
        return "top2(%)" if hints.get("summary_kind") == "top2" else "Mean"
    parts = ["%"]
    if kinds & {"bot", "soso", "top"}:
        parts.append("BOT2/SoSo/TOP2")
    if "mean" in kinds:
        parts.append("Mean")
    return ", ".join(parts)


def mean_unit(table):
    for r in table.get("rows", []):
        if r.get("kind") == "mean" and r.get("mean_unit"):
            return r["mean_unit"]
    return None


def tidy_base(raw):
    """questionnaire 의 'A3=1 | 기업소유' 를 'A3=1 (기업소유)' 로 다듬는다."""
    if not raw:
        return ""
    if "|" in raw:
        cond, note = raw.split("|", 1)
        return "%s (%s)" % (cond.strip(), note.strip())
    return raw.strip()


def put(ws, coord, value, style_from=None):
    """병합 영역이면 좌상단에만 쓴다. style_from 을 주면 그 셀의 서식을 복사한다."""
    cell = ws[coord]
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            cell = ws.cell(row=rng.min_row, column=rng.min_col)
            break
    if style_from is not None:
        src = ws[style_from]
        cell.font = copy(src.font)
        cell.fill = copy(src.fill)
        cell.border = copy(src.border)
        cell.alignment = copy(src.alignment)
        cell.number_format = src.number_format
    cell.value = value
    return cell


def clear_range(ws, min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                            min_col=min_col, max_col=max_col):
        for c in row:
            if not isinstance(c, MergedCell):
                c.value = None


def write_banner(ws, cfg, banner_inventory):
    put(ws, "B5", cfg.get("kms_no", ""))
    put(ws, "D5", cfg.get("project_name", ""))
    put(ws, "B6", cfg.get("researcher", ""))
    put(ws, "D6", cfg.get("dp", ""))
    for key, coord in (("result_table", "B8"), ("result_banner", "D8"),
                       ("result_data", "B9"), ("result_extra", "D9"),
                       ("result_qname", "B10"), ("result_tracking", "D10")):
        if cfg.get(key) is not None:
            put(ws, coord, cfg[key])

    rows = cfg["banner"]
    last_template_row = BANNER_NOTE_LABEL - 3      # A24:D24 앞까지가 배너 블록
    if BANNER_FIRST_ROW + len(rows) - 1 > last_template_row:
        raise SystemExit(
            "배너 정의가 %d건인데 템플릿 배너 블록은 %d행뿐입니다."
            % (len(rows), last_template_row - BANNER_FIRST_ROW + 1))

    clear_range(ws, BANNER_FIRST_ROW, last_template_row, 1, 4)
    for i, spec in enumerate(rows):
        r = BANNER_FIRST_ROW + i
        labels = banner_inventory.get(spec["group"])
        if labels is None:
            raise SystemExit("배너군 '%s' 이 통계표에 없습니다." % spec["group"])
        value_label = "\n".join(
            "%d) %s" % (n, lab) for n, lab in enumerate(labels, 1))
        if spec.get("recode"):
            value_label += "\nRECODE: " + spec["recode"]
        put(ws, "A%d" % r, spec["var"], style_from="A%d" % BANNER_FIRST_ROW)
        put(ws, "B%d" % r, spec["label"], style_from="B%d" % BANNER_FIRST_ROW)
        put(ws, "C%d" % r, value_label, style_from="C%d" % BANNER_FIRST_ROW)
        ws.row_dimensions[r].height = max(15.0, 12.6 * (value_label.count("\n") + 1))

    clear_range(ws, BANNER_NOTE_LABEL, BANNER_ADD_LAST, 2, 4)
    wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    small = Font(name="돋움", size=9)

    def _block(first, last, lines):
        for i, line in enumerate(lines[:last - first + 1]):
            cell = put(ws, "B%d" % (first + i), line)
            cell.alignment = wrap
            cell.font = small
            ws.row_dimensions[first + i].height = max(
                21.0, 13.5 * (len(line) // 52 + 1))

    _block(BANNER_NOTE_LABEL, BANNER_NOTE_LAST, cfg.get("notes", []))
    _block(BANNER_ADD_LABEL, BANNER_ADD_LAST, cfg.get("additions", []))


def write_quota(ws, quota):
    clear_range(ws, 1, max(ws.max_row, 24), 1, max(ws.max_column, 14))
    thin = Side(style="thin", color="FF808080")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="FFDCE6F1")
    base = Font(name="돋움", size=10)
    bold = Font(name="돋움", size=10, bold=True)

    ws["A1"] = quota.get("title", "쿼터")
    ws["A1"].font = Font(name="돋움", size=11, bold=True)
    ws["A2"] = quota.get("note", "")
    ws["A2"].font = base

    r0 = 4
    cols = quota["cols"]
    ncol = len(cols)
    ws.cell(row=r0, column=1,
            value="%s \\ %s" % (quota["row_name"], quota["col_name"]))
    for j, c in enumerate(cols):
        ws.cell(row=r0, column=2 + j, value=c)
    ws.cell(row=r0, column=2 + ncol, value="합계")
    for j in range(1, ncol + 3):
        c = ws.cell(row=r0, column=j)
        c.font, c.fill, c.border = bold, head_fill, bd
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for i, pair in enumerate(quota["rows"]):
        rname, vals = pair[0], pair[1]
        r = r0 + 1 + i
        ws.cell(row=r, column=1, value=rname)
        for j, v in enumerate(vals):
            ws.cell(row=r, column=2 + j, value=v)
        ws.cell(row=r, column=2 + ncol, value=sum(vals))
        for j in range(1, ncol + 3):
            c = ws.cell(row=r, column=j)
            c.font = bold if j in (1, 2 + ncol) else base
            c.border = bd
            c.alignment = Alignment(horizontal="center", vertical="center")
            if j in (1, 2 + ncol):
                c.fill = head_fill

    rt = r0 + 1 + len(quota["rows"])
    ws.cell(row=rt, column=1, value="합계")
    for j in range(ncol):
        ws.cell(row=rt, column=2 + j,
                value=sum(p[1][j] for p in quota["rows"]))
    ws.cell(row=rt, column=2 + ncol,
            value=sum(sum(p[1]) for p in quota["rows"]))
    for j in range(1, ncol + 3):
        c = ws.cell(row=rt, column=j)
        c.font, c.fill, c.border = bold, head_fill, bd
        c.alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions["A"].width = 18.4
    for j in range(ncol + 1):
        ws.column_dimensions[get_column_letter(2 + j)].width = 15.5
    ws.row_dimensions[r0].height = 30.0

    for i, line in enumerate(quota.get("footnotes", [])):
        c = ws.cell(row=rt + 2 + i, column=1, value=line)
        c.font = base


def write_basic(ws, tables, qbank, cfg):
    over = dict((int(k), v) for k, v in cfg.get("table_overrides", {}).items())
    last = BASIC_FIRST_ROW + len(tables) - 1
    clear_range(ws, BASIC_FIRST_ROW, max(ws.max_row, last), 1, 6)
    total_n = cfg.get("total_n")

    for t in tables:
        r = BASIC_FIRST_ROW + t["index"] - 1
        ov = over.get(t["index"], {})
        q = qbank.get(t.get("question_code") or "")

        a = ov.get("question_text")
        if a is None:
            a = " ".join(str((q or {}).get("text_full", "")).split())
        b = t["title_full"]
        c = ov.get("extra_banner", "")
        if not c:
            for grp in dict.fromkeys(x["group"] for x in t["banners"]):
                if grp in cfg.get("extra_banner_source", {}):
                    c = cfg["extra_banner_source"][grp]
        d = ov.get("base")
        if d is None:
            d = "ALL" if t["base_n"] == total_n else tidy_base((q or {}).get("base_condition"))
        e = ov.get("stat") or stat_spec(t)
        u = mean_unit(t)
        if u and "Mean" in e and "단위" not in e:
            e = "%s (단위: %s)" % (e, u)

        notes = []
        if t["base_n"] != total_n:
            notes.append("실산출 base n=%d" % t["base_n"])
        if ov.get("note"):
            notes.append(ov["note"])
        f = " · ".join(notes)

        for col, val, align in (("A", a, None), ("B", b, None), ("C", c, "center"),
                                ("D", d, "center"), ("E", e, "center"), ("F", f, None)):
            cell = put(ws, "%s%d" % (col, r), val or None,
                       style_from="%s%d" % (col, BASIC_FIRST_ROW))
            cell.font = Font(name="맑은 고딕", size=10)
            cell.alignment = Alignment(horizontal=align, vertical="center")
            if col == "F" and val:
                cell.fill = PatternFill("solid", fgColor="FFFFFF00")
    return last


def write_sheet_table(wb, name, title, headers, rows, widths, intro=None):
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    thin = Side(style="thin", color="FF808080")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    base = Font(name="맑은 고딕", size=10)

    ws["A1"] = title
    ws["A1"].font = Font(name="맑은 고딕", size=14, bold=True)
    r = 2
    for line in (intro or []):
        r += 1
        ws.cell(row=r, column=1, value=line).font = base
    head_row = r + 2
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=head_row, column=j, value=h)
        c.font = Font(name="맑은 고딕", size=10, bold=True)
        c.fill = PatternFill("solid", fgColor="FFDCE6F1")
        c.border = bd
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row, 1):
            c = ws.cell(row=head_row + i, column=j, value=v)
            c.font = base
            c.border = bd
            c.alignment = Alignment(vertical="center", wrap_text=True,
                                    horizontal="center" if j in (1, 2) else None)
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[head_row].height = 28.0
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)
    return ws


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("tables")
    ap.add_argument("questionnaire")
    ap.add_argument("out")
    ap.add_argument("--config", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--creator", default="마크로밀엠브레인",
                    help="문서 속성의 작성자·최종 수정자에 넣을 조직명")
    args = ap.parse_args()

    tj = json.loads(Path(args.tables).read_text(encoding="utf-8"))
    qj = json.loads(Path(args.questionnaire).read_text(encoding="utf-8"))
    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    tables = tj["tables"]
    qbank = dict((q["code"], q) for q in qj["questions"])

    inventory = {}
    for t in tables:
        for b in t["banners"]:
            inventory.setdefault(b["group"], [])
            if b["label"] not in inventory[b["group"]]:
                inventory[b["group"]].append(b["label"])

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(args.template, str(out))
    wb = openpyxl.load_workbook(str(out))

    write_banner(wb[SH_BANNER], cfg, inventory)
    write_quota(wb[SH_QUOTA], cfg["quota"])
    last = write_basic(wb[SH_BASIC], tables, qbank, cfg)

    if cfg.get("add_table"):
        write_sheet_table(
            wb, SH_ADD, cfg["add_table"]["title"],
            ["#", "구분", "요청 내용", "대상 표 / 문항", "근거", "상태"],
            [[i + 1, x["kind"], x["request"], x["target"], x["reason"], x["status"]]
             for i, x in enumerate(cfg["add_table"]["items"])],
            [5, 14, 62, 26, 62, 16],
            cfg["add_table"].get("intro"))

    total_n = cfg.get("total_n")
    trace_rows = []
    for t in tables:
        q = qbank.get(t.get("question_code") or "")
        trace_rows.append([
            t["index"], BASIC_FIRST_ROW + t["index"] - 1,
            t.get("question_code") or "-", t["title_full"],
            "ALL" if t["base_n"] == total_n else tidy_base((q or {}).get("base_condition")),
            t["base_n"], t["hints"]["n_banners"], t["hints"]["n_banners_in_report"],
            stat_spec(t), t["excel_range"],
        ])
    write_sheet_table(
        wb, SH_TRACE, cfg["trace"]["title"],
        ["표", "가이드 행", "문항", "표 제목", "설문지 BASE 정의",
         "실산출 n", "배너 열", "보고서 IN 열", "통계값", "통계표 셀 범위"],
        trace_rows, [6, 9, 10, 52, 34, 9, 8, 11, 26, 14],
        cfg["trace"].get("intro"))

    wb.save(str(out))

    # 문서 위생 (R-137): 템플릿에서 따라온 개인 이름을 조직명으로 바꾸고 보이지 않는 문자를 제거한다.
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import doc_hygiene
    tmp = out.with_suffix(out.suffix + ".hyg")
    hr = doc_hygiene.clean_file(out, tmp, args.creator)
    shutil.move(str(tmp), str(out))
    if hr["props_changed"] or hr["invisible"]:
        print("     문서 위생 — 속성 %s · 제거 문자 %d개"
              % ("정리" if hr["props_changed"] else "그대로", sum(hr["invisible"].values())))

    print("[OK] %s" % out)
    print("     Basic Table %d~%d행 = 표1~표%d · 배너 %d군"
          % (BASIC_FIRST_ROW, last, len(tables), len(cfg["banner"])))
    return 0


if __name__ == "__main__":
    sys.exit(main())
