#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""설문 원안(.docx)에서 문단/표를 문서 순서대로, run 서식과 함께 추출한다."""

import argparse
import json
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

try:
    from docx import Document
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run
except ImportError:
    sys.stderr.write("오류: python-docx가 설치되어 있지 않습니다. `pip install python-docx` 후 다시 실행하세요.\n")
    sys.exit(1)


def die(msg):
    sys.stderr.write("오류: " + msg + "\n")
    sys.exit(1)


def iter_block_items(parent):
    """문단과 표를 문서에 나타난 실제 순서대로 내보낸다."""
    body = parent._tc if hasattr(parent, "_tc") else parent.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def paragraph_runs(paragraph):
    """하이퍼링크 안의 run까지 포함해 run 객체를 순서대로 돌려준다."""
    nodes = paragraph._p.xpath("./w:r | ./w:hyperlink/w:r")
    return [Run(r, paragraph) for r in nodes]


def highlight_name(value):
    name = getattr(value, "name", None)
    if name:
        return str(name)
    text = str(value)
    if "." in text:
        text = text.rsplit(".", 1)[1]
    return text.split(" ")[0]


def run_format(run):
    fmt = {}
    font = run.font
    try:
        rgb = font.color.rgb if font.color is not None else None
    except (AttributeError, ValueError, TypeError):
        rgb = None
    if rgb is not None:
        fmt["color"] = str(rgb)
    if font.strike:
        fmt["strike"] = True
    if run.bold:
        fmt["bold"] = True
    if run.underline:
        fmt["underline"] = True
    if font.highlight_color is not None:
        fmt["highlight"] = highlight_name(font.highlight_color)
    return fmt


def merged_runs(paragraph):
    """서식이 같은 연속 run을 하나로 합친다 (Word가 쪼개 놓은 run 정리)."""
    out = []
    for run in paragraph_runs(paragraph):
        text = run.text
        if not text:
            continue
        fmt = run_format(run)
        if out and out[-1]["fmt"] == fmt:
            out[-1]["text"] += text
        else:
            out.append({"text": text, "fmt": fmt})
    return out


def format_tag(fmt):
    parts = []
    if "color" in fmt:
        parts.append("color=" + fmt["color"])
    if fmt.get("strike"):
        parts.append("strike")
    if fmt.get("bold"):
        parts.append("bold")
    if fmt.get("underline"):
        parts.append("underline")
    if "highlight" in fmt:
        parts.append("highlight=" + fmt["highlight"])
    return ",".join(parts)


def tagged_text(paragraph):
    """서식이 있는 run만 [color=FF0000,strike]...[/] 로 감싸고 평문은 그대로 둔다."""
    chunks = []
    for item in merged_runs(paragraph):
        tag = format_tag(item["fmt"])
        if tag:
            chunks.append("[" + tag + "]" + item["text"] + "[/]")
        else:
            chunks.append(item["text"])
    return "".join(chunks)


def cell_text(cell):
    """셀 안의 문단과 중첩 표를 한 문자열로 만든다."""
    pieces = []
    for block in iter_block_items(cell):
        if isinstance(block, Paragraph):
            text = tagged_text(block).strip()
            if text:
                pieces.append(text)
        else:
            for row in block.rows:
                pieces.append(" | ".join(cell_text(c) for c in row.cells))
    return " / ".join(pieces)


def table_rows(table):
    rows = []
    for row in table.rows:
        rows.append([cell_text(c) for c in row.cells])
    return rows


def collect(path):
    document = Document(path)
    blocks = []
    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            runs = []
            for item in merged_runs(block):
                entry = {"text": item["text"]}
                entry.update(item["fmt"])
                runs.append(entry)
            try:
                style = block.style.name if block.style is not None else None
            except (AttributeError, KeyError):
                style = None
            blocks.append({
                "type": "paragraph",
                "style": style,
                "text": tagged_text(block),
                "plain": block.text,
                "runs": runs,
            })
        else:
            blocks.append({"type": "table", "rows": table_rows(block)})
    return blocks


def print_text(blocks):
    table_no = 0
    for block in blocks:
        if block["type"] == "paragraph":
            print(block["text"])
        else:
            table_no += 1
            rows = block["rows"]
            cols = max((len(r) for r in rows), default=0)
            print("")
            print("[표 %d] %d행 x %d열" % (table_no, len(rows), cols))
            for row in rows:
                print("| " + " | ".join(row) + " |")
            print("")


def main():
    parser = argparse.ArgumentParser(
        prog="read_docx.py",
        description=(
            "설문 원안 Word(.docx) 파일에서 문단과 표를 문서에 적힌 실제 순서대로 추출한다.\n"
            "글자색/취소선/굵게/밑줄/형광펜 서식을 run 단위로 함께 표기하므로,\n"
            "색이나 취소선으로 표기된 이탈·로테이션 조건을 놓치지 않는다."
        ),
        epilog=(
            "예시:\n"
            "  python read_docx.py 설문원안.docx\n"
            "  python read_docx.py 설문원안.docx --json > 원안.json\n\n"
            "서식 표기: 서식이 있는 run만 [color=FF0000,strike]본문[/] 형태로 감싸고,\n"
            "평문 run은 그대로 출력한다."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("path", help="읽을 .docx 파일 경로")
    parser.add_argument(
        "--json",
        action="store_true",
        dest="as_json",
        help="JSON으로 출력한다 (문단은 run 배열, 표는 행/열 중첩 배열)",
    )
    args = parser.parse_args()

    path = os.path.abspath(args.path)
    if not os.path.exists(path):
        die("파일을 찾을 수 없습니다: " + path)
    if os.path.isdir(path):
        die("디렉터리입니다. .docx 파일 경로를 지정하세요: " + path)
    if not path.lower().endswith(".docx"):
        die(".docx 파일이 아닙니다(구형 .doc는 지원하지 않습니다. Word에서 .docx로 저장하세요): " + path)

    try:
        blocks = collect(path)
    except Exception as exc:
        die("문서를 읽지 못했습니다 (%s): %s" % (type(exc).__name__, exc))

    if args.as_json:
        print(json.dumps({"file": path, "blocks": blocks}, ensure_ascii=False, indent=2))
    else:
        print_text(blocks)


if __name__ == "__main__":
    main()
