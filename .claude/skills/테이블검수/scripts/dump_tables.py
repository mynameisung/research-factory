#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""통계표(.xlsx/.csv)를 문항 블록 단위로 잘라 구조화해 덤프한다.

블록 분할은 휴리스틱이다. 반드시 --list 로 훑고 --block 으로 2~3개를 눈으로
확인한 뒤 대조에 쓴다. 분할이 어긋나면 --sep-rows / --title-regex 를 조정한다.
"""

import argparse
import csv
import json
import os
import re
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

# 문번 패턴: Q1, SQ2-1, DQ3, A1_2, 문5, [Q7], 1., BQ10 ...
DEFAULT_TITLE_RE = r"(?:^|[\[\(\s])((?:S|D|B|F|A|Q|문)?Q?\s*\d+(?:\s*[-_.]\s*\d+)*)\s*[.)\]]?"

# 사례수 행으로 볼 라벨
N_LABELS = ("사례수", "사례 수", "응답자수", "응답자 수", "base", "Base", "BASE",
            "전체", "계", "합계", "표본수", "n수", "N수", "(명)")

# base 표기로 볼 라벨
BASE_HINTS = ("base", "Base", "BASE", "기준", "대상", "응답대상", "분석대상")


def die(msg):
    sys.stderr.write("오류: " + msg + "\n")
    sys.exit(1)


def cell_str(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if v == int(v):
            return str(int(v))
        return "%g" % v
    return str(v).replace("\n", " ").strip()


def fill_merged_cells(ws, grid):
    """병합 셀의 값을 병합 범위 전체에 복사한다 (라벨이 빈칸으로 읽히는 것 방지)."""
    for rng in list(ws.merged_cells.ranges):
        r0, c0 = rng.min_row - 1, rng.min_col - 1
        r1, c1 = rng.max_row - 1, rng.max_col - 1
        if r0 >= len(grid) or c0 >= len(grid[r0]):
            continue
        val = grid[r0][c0]
        if not val:
            continue
        for r in range(r0, min(r1 + 1, len(grid))):
            for c in range(c0, c1 + 1):
                while len(grid[r]) <= c:
                    grid[r].append("")
                if r == r0 and c == c0:
                    continue
                if not grid[r][c]:
                    grid[r][c] = val


def load_xlsx(path, only_sheets, fill_merged):
    try:
        from openpyxl import load_workbook
    except ImportError:
        die("openpyxl이 없습니다. `pip install openpyxl` 후 다시 실행하세요.")
    try:
        wb = load_workbook(path, data_only=True, read_only=not fill_merged)
    except Exception as exc:
        die("파일을 열지 못했습니다 (%s): %s" % (type(exc).__name__, exc))
    sheets = []
    for ws in wb.worksheets:
        if only_sheets and ws.title not in only_sheets:
            continue
        grid = [[cell_str(v) for v in row] for row in ws.iter_rows(values_only=True)]
        if fill_merged:
            fill_merged_cells(ws, grid)
        sheets.append((ws.title, grid))
    try:
        wb.close()
    except Exception:
        pass
    if only_sheets and not sheets:
        die("지정한 시트를 찾을 수 없습니다: " + ", ".join(sorted(only_sheets)))
    return sheets


def load_csv(path):
    for enc in ("utf-8-sig", "cp949", "utf-8"):
        try:
            with open(path, "r", encoding=enc, newline="") as fh:
                sample = fh.read(8192)
                fh.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",\t|;")
                except csv.Error:
                    dialect = csv.excel_tab if "\t" in sample else csv.excel
                grid = [[cell_str(c) for c in row] for row in csv.reader(fh, dialect)]
            return [(os.path.basename(path), grid)]
        except UnicodeDecodeError:
            continue
    die("인코딩을 판별하지 못했습니다 (utf-8/cp949 모두 실패): " + path)


def load_xls(path, only_sheets, fill_merged):
    """구형 .xls(BIFF8)를 xlrd로 읽는다. DP실이 .xls로 내는 경우가 흔하다."""
    try:
        import xlrd
    except ImportError:
        die(".xls를 읽으려면 xlrd가 필요합니다. `pip install xlrd` 후 다시 실행하세요.")
    try:
        wb = xlrd.open_workbook(path, formatting_info=fill_merged)
    except Exception as exc:
        die("파일을 열지 못했습니다 (%s): %s\n"
            "     신형 확장자만 .xls인 파일(HTML/xlsx)일 수 있습니다. Excel에서 .xlsx로 저장해 보세요."
            % (type(exc).__name__, exc))
    sheets = []
    for sh in wb.sheets():
        if only_sheets and sh.name not in only_sheets:
            continue
        grid = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                v = sh.cell_value(r, c)
                row.append("" if v == "" else cell_str(v))
            grid.append(row)
        if fill_merged:
            # xlrd의 merged_cells는 (rlo, rhi, clo, chi) — 끝값 배타
            for r0, r1, c0, c1 in sh.merged_cells:
                if r0 >= len(grid) or c0 >= len(grid[r0]):
                    continue
                val = grid[r0][c0]
                if not val:
                    continue
                for r in range(r0, min(r1, len(grid))):
                    for c in range(c0, c1):
                        while len(grid[r]) <= c:
                            grid[r].append("")
                        if (r, c) != (r0, c0) and not grid[r][c]:
                            grid[r][c] = val
        sheets.append((sh.name, grid))
    if only_sheets and not sheets:
        die("지정한 시트를 찾을 수 없습니다: " + ", ".join(sorted(only_sheets)))
    return sheets


def row_empty(row):
    return all(not c for c in row)


def trim(rows):
    """블록의 오른쪽·아래쪽 빈 영역을 잘라낸다."""
    while rows and row_empty(rows[-1]):
        rows.pop()
    width = 0
    for row in rows:
        for i in range(len(row) - 1, -1, -1):
            if row[i]:
                width = max(width, i + 1)
                break
    out = []
    for row in rows:
        cut_row = row[:width]
        out.append(cut_row + [""] * (width - len(cut_row)))
    return out


def split_blocks(grid, sep_rows):
    """연속된 빈 행 sep_rows개 이상을 구분자로 보고 블록을 나눈다."""
    blocks, cur, blank, start = [], [], 0, 0
    for idx, row in enumerate(grid):
        if row_empty(row):
            blank += 1
            if cur and blank >= sep_rows:
                blocks.append((start, trim(cur)))
                cur = []
            continue
        if not cur:
            start = idx
        blank = 0
        cur.append(list(row))
    if cur:
        blocks.append((start, trim(cur)))
    return [(s, b) for s, b in blocks if b]


def first_text(row):
    for c in row:
        if c:
            return c
    return ""


def join_row(row):
    return " ".join(c for c in row if c).strip()


def is_num(s):
    if not s:
        return False
    return bool(re.fullmatch(r"-?[\d,]+(?:\.\d+)?%?", s.strip()))


def to_num(s):
    try:
        return float(s.strip().rstrip("%").replace(",", ""))
    except ValueError:
        return None


def analyze(rows, title_re):
    title = ""
    for row in rows[:3]:
        t = join_row(row)
        if t:
            title = t
            break
    m = re.search(title_re, title)
    qcode = re.sub(r"\s+", "", m.group(1)) if m else ""

    base_label = ""
    for row in rows[:6]:
        t = join_row(row)
        if any(h in t for h in BASE_HINTS):
            base_label = t
            break

    n_row_idx, n_values = None, []
    for i, row in enumerate(rows):
        lab = first_text(row)
        if not lab:
            continue
        if any(lab == x or lab.startswith(x) for x in N_LABELS):
            vals = [to_num(c) for c in row if is_num(c)]
            vals = [v for v in vals if v is not None]
            if vals:
                n_row_idx, n_values = i, vals
                break

    labels = []
    start = (n_row_idx + 1) if n_row_idx is not None else 1
    for i in range(start, len(rows)):
        lab = first_text(rows[i])
        if lab and not is_num(lab):
            labels.append({"row": i, "label": lab})

    return {
        "title": title,
        "qcode": qcode,
        "base_label": base_label,
        "n_row": n_row_idx,
        "total_n": n_values[0] if n_values else None,
        "n_values": n_values,
        "n_rows": len(rows),
        "n_cols": max((len(r) for r in rows), default=0),
        "row_labels": labels,
    }


def build(path, only_sheets, sep_rows, title_re, fill_merged):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".tsv", ".txt"):
        sheets = load_csv(path)
    elif ext in (".xlsx", ".xlsm"):
        sheets = load_xlsx(path, only_sheets, fill_merged)
    elif ext == ".xls":
        sheets = load_xls(path, only_sheets, fill_merged)
    else:
        die("지원하지 않는 확장자입니다(.xlsx/.xlsm/.xls/.csv/.tsv/.txt): " + path)

    out, idx = [], 0
    for name, grid in sheets:
        for srow, block in split_blocks(grid, sep_rows):
            info = analyze(block, title_re)
            info.update({"index": idx, "sheet": name, "sheet_row": srow + 1, "rows": block})
            out.append(info)
            idx += 1
    return {"file": os.path.abspath(path), "block_count": len(out), "blocks": out}


def cut(s, n):
    s = s.replace("\t", " ")
    return s if len(s) <= n else s[: max(n - 1, 1)] + "~"


def print_list(data, grep):
    print("파일: %s   블록 %d개" % (data["file"], data["block_count"]))
    print("%-5s %-14s %-6s %-10s %-8s %-8s %s"
          % ("IDX", "SHEET", "ROW", "문번", "N", "행x열", "제목"))
    print("-" * 112)
    shown = 0
    for b in data["blocks"]:
        if grep and grep not in json.dumps(b, ensure_ascii=False):
            continue
        shown += 1
        n = "-" if b["total_n"] is None else "%g" % b["total_n"]
        print("%-5d %-14s %-6d %-10s %-8s %-8s %s"
              % (b["index"], cut(b["sheet"], 14), b["sheet_row"], cut(b["qcode"] or "-", 10),
                 n, "%dx%d" % (b["n_rows"], b["n_cols"]), cut(b["title"], 46)))
    if grep:
        print("-" * 112)
        print("'%s' 포함 블록 %d개" % (grep, shown))


def print_block(data, i):
    blocks = data["blocks"]
    if i < 0 or i >= len(blocks):
        die("블록 번호는 0 ~ %d 범위입니다: %d" % (len(blocks) - 1, i))
    b = blocks[i]
    print("[블록 %d] 시트=%s 시작행=%d  문번=%s  N=%s  %d행 x %d열"
          % (i, b["sheet"], b["sheet_row"], b["qcode"] or "-",
             "-" if b["total_n"] is None else "%g" % b["total_n"],
             b["n_rows"], b["n_cols"]))
    if b["base_label"]:
        print("Base 표기: " + b["base_label"])
    print("")
    widths = []
    for c in range(b["n_cols"]):
        lens = [len(r[c]) for r in b["rows"] if c < len(r)]
        widths.append(min(max(lens + [3]), 22))
    print("     " + " | ".join(("c%d" % c).ljust(widths[c])[: widths[c]] for c in range(b["n_cols"])))
    for ri, row in enumerate(b["rows"]):
        cells = [cut(row[c] if c < len(row) else "", widths[c]).ljust(widths[c])
                 for c in range(b["n_cols"])]
        mark = "N>" if b["n_row"] == ri else "  "
        print("%s%3d %s" % (mark, ri, " | ".join(cells)))


def main():
    p = argparse.ArgumentParser(
        prog="dump_tables.py",
        description="통계표를 문항 블록 단위로 잘라 구조화해 덤프한다.",
        epilog=(
            "예시:\n"
            "  python dump_tables.py 테이블.xlsx --list\n"
            "  python dump_tables.py 테이블.xlsx --list --grep Q7\n"
            "  python dump_tables.py 테이블.xlsx --block 12\n"
            "  python dump_tables.py 테이블.xlsx --json -o tables.json\n\n"
            "블록 분할이 어긋나면 --sep-rows 2 로 늘리거나 --title-regex 를 바꾼다."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("path", help="테이블 파일 (.xlsx/.xlsm/.xls/.csv/.tsv/.txt)")
    p.add_argument("--list", action="store_true", help="블록 목록을 표로 출력한다 (기본)")
    p.add_argument("--block", type=int, metavar="N", help="N번 블록의 원본 격자를 그대로 출력한다")
    p.add_argument("--json", action="store_true", dest="as_json", help="전체를 JSON으로 출력한다")
    p.add_argument("--grep", metavar="TEXT", help="--list에서 해당 문자열을 포함한 블록만 본다")
    p.add_argument("--sheets", metavar="A,B", help="쉼표로 구분한 시트명만 읽는다")
    p.add_argument("--sep-rows", type=int, default=1, metavar="N",
                   help="블록 구분자로 볼 연속 빈 행 수 (기본 1)")
    p.add_argument("--title-regex", default=DEFAULT_TITLE_RE, metavar="RE",
                   help="제목에서 문번을 뽑는 정규식 (그룹 1이 문번)")
    p.add_argument("--no-fill-merged", action="store_true",
                   help="병합 셀 값 채우기를 끈다 (기본은 켜짐)")
    p.add_argument("-o", "--out", metavar="FILE", help="출력 파일 (기본 표준출력)")
    args = p.parse_args()

    path = os.path.abspath(args.path)
    if not os.path.exists(path):
        die("파일을 찾을 수 없습니다: " + path)
    if os.path.isdir(path):
        die("디렉터리입니다. 파일 경로를 지정하세요: " + path)
    try:
        re.compile(args.title_regex)
    except re.error as exc:
        die("--title-regex가 올바른 정규식이 아닙니다: %s" % exc)
    if args.sep_rows < 1:
        die("--sep-rows는 1 이상이어야 합니다.")

    only = set(s.strip() for s in args.sheets.split(",") if s.strip()) if args.sheets else None
    data = build(path, only, args.sep_rows, args.title_regex, not args.no_fill_merged)

    stream = open(args.out, "w", encoding="utf-8") if args.out else sys.stdout
    old = sys.stdout
    sys.stdout = stream
    try:
        if args.as_json:
            print(json.dumps(data, ensure_ascii=False, indent=2))
        elif args.block is not None:
            print_block(data, args.block)
        else:
            print_list(data, args.grep)
    finally:
        sys.stdout = old
        if args.out:
            stream.close()
            sys.stderr.write("저장: %s\n" % os.path.abspath(args.out))


if __name__ == "__main__":
    main()
