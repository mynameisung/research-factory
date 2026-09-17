"""설문지 .docx → questionnaire.json.

    python parse_questionnaire.py <input.docx> <output.json> [--strict] [--quiet]

추출 대상
    · 파트 코드/제목            (`SQ. Screening Question`, `Part A. …` 1셀 표)
    · 문항번호 · 문항 텍스트    — **원문 그대로**. 제작가이드 `Qu:` 줄에 그대로 들어간다.
    · 응답유형                  단일/복수/순위/척도/숫자입력/개방형/예-아니오/행렬
    · 척도 점수 · 척도 라벨     5점/7점, 가로형 표 및 행렬 표 모두
    · 순위 사양                 `[3순위 필수]`, `[1순위 필수, 3순위 선택]`
    · 보기 목록                 번호 · 텍스트 · 기타/오픈/파이핑/조사중단 여부
    · 프로그래머 지시           `[BASE: …]` `[RANGE: …]` `[PROG: …]` `[# … 파이핑 #]`

⚠ 원문의 화살표는 Wingdings 심볼(`<w:sym w:font="Wingdings" w:char="F0E0"/>`)이다.
python-docx의 `paragraph.text`는 이걸 통째로 버리기 때문에 `→ 조사 중단`이
`  조사 중단`으로 보인다. 이 모듈은 런을 직접 훑어 심볼을 복원한다.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterator

SCHEMA_VERSION = "1.0"

# ── Wingdings/Symbol 사설영역 매핑 ──────────────────────────────────────────
SYMBOL_MAP: dict[tuple[str, int], str] = {
    ("wingdings", 0xF0E0): "→",
    ("wingdings", 0xF0DF): "←",
    ("wingdings", 0xF0DE): "↑",
    ("wingdings", 0xF0E2): "↓",
    ("wingdings", 0xF0FC): "✓",
    ("wingdings", 0xF0FE): "☑",
    ("wingdings", 0xF0A7): "▪",
    ("wingdings", 0xF0B7): "●",
    ("wingdings", 0xF06C): "●",
    ("symbol", 0xF0AE): "→",
    ("symbol", 0xF0B7): "•",
}

# ── 응답유형 마커 ──────────────────────────────────────────────────────────
MARKER_SINGLE = "1개 선택"
MARKER_MULTI = "모두 선택"
MARKER_OPEN = "직접 입력"
MARKER_ROW_REQUIRED = "1행 필수 입력"
MARKER_MATRIX_ROW = "행별 1개 선택"
MARKER_MATRIX_ITEM = "항목별 1개 선택"
PLAIN_MARKERS = (
    MARKER_SINGLE, MARKER_MULTI, MARKER_OPEN,
    MARKER_ROW_REQUIRED, MARKER_MATRIX_ROW, MARKER_MATRIX_ITEM,
)

TYPE_SINGLE = "단일"
TYPE_MULTI = "복수"
TYPE_RANK = "순위"
TYPE_SCALE = "척도"
TYPE_NUMBER = "숫자입력"
TYPE_OPEN = "개방형"
TYPE_YESNO = "예-아니오"
TYPE_MATRIX = "행렬"
TYPE_UNKNOWN = "미상"

# ── 정규식 ─────────────────────────────────────────────────────────────────
RE_QCODE_HEAD = re.compile(r"^\s*((?:[A-Z]{1,3}\d+(?:-\d+)*)|ED)\.\s*(.*)$", re.DOTALL)
RE_PART_HEAD = re.compile(r"^\s*(?:PART|Part|part)?\s*([A-Z]{1,2})\.\s*(.+)$", re.DOTALL)
RE_OPTION = re.compile(r"^\s*(\d+)\)\s*(.*)$", re.DOTALL)
RE_OPTION_SPLIT = re.compile(r"(?:\t+|\s{3,})(?=\d+\)\s)")
RE_RANK_SPEC = re.compile(r"\[\s*(\d+)\s*순위\s*필수(?:\s*,\s*(\d+)\s*순위\s*선택)?\s*\]")
# 개수 제한 복수 선택: `[최소 1개, 최대 3개 선택]` · `[최대 3개 선택]`
RE_MULTI_LIMIT = re.compile(r"\[\s*(?:최소\s*(\d+)\s*개\s*,?\s*)?최대\s*(\d+)\s*개\s*선택\s*\]")
RE_MARKER = re.compile(r"\[([^\[\]]{1,40})\]")
RE_BASE = re.compile(r"\[BASE\s*:\s*(.*?)\]", re.DOTALL)
RE_RANGE = re.compile(r"\[RANGE\s*:\s*(.*?)\]", re.DOTALL)
RE_PROG = re.compile(r"\[PROG\s*:\s*(.*?)\]", re.DOTALL)
RE_PIPING = re.compile(r"\[#\s*(.*?)\s*#\]", re.DOTALL)
RE_DIRECTIVE_LINE = re.compile(r"^\s*\[(?:BASE|RANGE|PROG)\s*:")
RE_CIRCLED = re.compile(r"[①-⑳㉑-㉟]")   # ①~⑳ 등
RE_SCALE_NUM_CELL = re.compile(r"^\s*\d+\)?\s*$")   # `1)` 또는 `1`
RE_OPEN_BLANK = re.compile(r"\(\s{2,}\)|\(\s*\)")
YESNO_TOKENS = ({"예", "아니오"}, {"예", "아니요"}, {"네", "아니오"})
# 기입란 `(      )` 바로 뒤에 단위가 붙으면 숫자 입력. 문항 본문 단어에는 걸리지
# 않도록 반드시 닫는 괄호 직후만 본다 ('개선사항'·'명칭' 오탐 방지).
RE_NUMBER_ANSWER = re.compile(
    r"\)\s*(?:만\s*원|개월|층|개|대|명|년|월|일|평|원|%|㎡|시간|건|회|점)"
)


class QuestionnaireParseError(RuntimeError):
    """구조가 예상과 다를 때 즉시 던진다."""


# ═══════════════════════════════════════════════════════════════════════════
# docx 저수준 읽기
# ═══════════════════════════════════════════════════════════════════════════
def _tag(el: Any) -> str:
    return el.tag.split("}")[-1] if isinstance(el.tag, str) else ""


def _sym_char(node: Any, unmapped: set[str]) -> str:
    """`<w:sym>` 을 실제 문자로 되살린다."""
    font = ""
    char = ""
    for k, v in node.attrib.items():
        name = k.split("}")[-1]
        if name == "font":
            font = v
        elif name == "char":
            char = v
    try:
        code = int(char, 16)
    except ValueError:
        unmapped.add(f"{font}:{char}")
        return ""
    hit = SYMBOL_MAP.get((font.lower(), code))
    if hit:
        return hit
    # Wingdings 계열은 F020~F0FF 가 ASCII 0x20~0xFF 에 대응한다. 최후의 폴백.
    if 0xF020 <= code <= 0xF0FF:
        unmapped.add(f"{font}:{char}")
        return chr(code - 0xF000)
    unmapped.add(f"{font}:{char}")
    return ""


def _element_text(el: Any, unmapped: set[str]) -> str:
    """단락/셀 XML에서 텍스트를 뽑는다. `w:sym`·`w:tab`·`w:br` 을 살린다."""
    out: list[str] = []
    for node in el.iter():
        t = _tag(node)
        if t == "t":
            out.append(node.text or "")
        elif t == "tab":
            out.append("\t")
        elif t in ("br", "cr"):
            out.append("\n")
        elif t == "sym":
            out.append(_sym_char(node, unmapped))
    return "".join(out)


@dataclass
class Block:
    kind: str                      # 'p' | 'tbl'
    index: int
    text: str = ""                 # 단락 텍스트 (kind='p')
    cells: list[list[str]] = field(default_factory=list)   # 표 (kind='tbl')

    @property
    def stripped(self) -> str:
        return self.text.strip()

    @property
    def n_rows(self) -> int:
        return len(self.cells)

    @property
    def n_cols(self) -> int:
        return max((len(r) for r in self.cells), default=0)

    def cell(self, r: int, c: int) -> str:
        if 0 <= r < len(self.cells) and 0 <= c < len(self.cells[r]):
            return self.cells[r][c].strip()
        return ""

    def flat_text(self) -> str:
        return "\n".join(" | ".join(row) for row in self.cells)


def read_blocks(path: Path) -> tuple[list[Block], set[str]]:
    """본문을 문서 순서대로 단락/표 블록 리스트로 읽는다."""
    try:
        import docx  # type: ignore
        from docx.table import Table
    except ImportError as e:  # pragma: no cover
        raise QuestionnaireParseError(
            f"python-docx 가 필요합니다 (`pip install python-docx`). 원인: {e}"
        ) from e

    doc = docx.Document(str(path))
    unmapped: set[str] = set()
    blocks: list[Block] = []
    for i, child in enumerate(doc.element.body.iterchildren()):
        t = _tag(child)
        if t == "p":
            blocks.append(Block("p", i, text=_element_text(child, unmapped)))
        elif t == "tbl":
            tbl = Table(child, doc)
            cells: list[list[str]] = []
            for row in tbl.rows:
                try:
                    row_cells = row.cells
                except IndexError:  # 병합이 깨진 표
                    continue
                cells.append([_element_text(c._tc, unmapped).strip() for c in row_cells])
            blocks.append(Block("tbl", i, cells=cells))
    if not blocks:
        raise QuestionnaireParseError(f"{path.name}: 본문 블록이 하나도 없습니다.")
    return blocks, unmapped


# ═══════════════════════════════════════════════════════════════════════════
# 파싱 도우미
# ═══════════════════════════════════════════════════════════════════════════
def norm_space(s: str) -> str:
    return re.sub(r"[ \t]+", " ", s.replace(" ", " ")).strip()


def extract_directives(text: str) -> dict[str, list[str]]:
    return {
        "base": [norm_space(m) for m in RE_BASE.findall(text)],
        "range": [norm_space(m) for m in RE_RANGE.findall(text)],
        "prog": [norm_space(m) for m in RE_PROG.findall(text)],
        "piping": [norm_space(m) for m in RE_PIPING.findall(text)],
    }


def merge_directives(dst: dict[str, list[str]], src: dict[str, list[str]]) -> None:
    for k, v in src.items():
        for item in v:
            if item not in dst[k]:
                dst[k].append(item)


def find_response_markers(text: str) -> list[str]:
    """문항 텍스트 안의 `[…]` 중 응답유형 마커만 원문 그대로 골라낸다."""
    out: list[str] = []
    for m in RE_MARKER.finditer(text):
        inner = norm_space(m.group(1))
        if (inner in PLAIN_MARKERS or RE_RANK_SPEC.fullmatch(f"[{inner}]")
                or RE_MULTI_LIMIT.fullmatch(f"[{inner}]")):
            out.append(f"[{inner}]")
    return out


def parse_rank_spec(markers: list[str]) -> dict[str, Any] | None:
    for mk in markers:
        m = RE_RANK_SPEC.fullmatch(mk)
        if m:
            required = int(m.group(1))
            optional = int(m.group(2)) if m.group(2) else None
            return {
                "spec": mk,
                "required_ranks": required,
                "optional_ranks": optional,
                "max_ranks": optional or required,
            }
    return None


def split_option_line(text: str) -> list[str]:
    """`1) 예\\t\\t2) 아니오` 처럼 한 줄에 여러 보기가 붙은 경우를 쪼갠다."""
    parts = [p for p in RE_OPTION_SPLIT.split(text) if p.strip()]
    return parts or [text]


def parse_option(raw: str) -> dict[str, Any] | None:
    m = RE_OPTION.match(raw)
    if not m:
        return None
    no = int(m.group(1))
    body = m.group(2)
    directives = extract_directives(body)
    clean = RE_PIPING.sub("", body)
    clean = RE_BASE.sub("", RE_RANGE.sub("", RE_PROG.sub("", clean)))
    clean = norm_space(clean)
    terminate = "조사 중단" in body or "조사중단" in body
    text_wo_terminate = norm_space(re.sub(r"→?\s*조사\s*중단\s*$", "", clean))
    return {
        "no": no,
        "raw": raw.strip(),          # 공백·화살표까지 원문 그대로
        "text": text_wo_terminate or clean,
        "is_other": text_wo_terminate.startswith("기타"),
        "is_none": bool(re.match(r"^(현재로서는\s*)?없음|^정보를 얻기 위해", text_wo_terminate)),
        "has_open_blank": bool(RE_OPEN_BLANK.search(body)),
        "terminate": terminate,
        "piping": directives["piping"] or None,
    }


# ── 표 분류 ────────────────────────────────────────────────────────────────
def classify_table(blk: Block) -> dict[str, Any]:
    """설문지 안의 표를 용도별로 분류하고 척도 정보를 뽑는다."""
    rows, cols = blk.n_rows, blk.n_cols
    info: dict[str, Any] = {
        "block": blk.index, "kind": "other", "n_rows": rows, "n_cols": cols,
        "cells": blk.cells,
    }
    if rows == 0 or cols == 0:
        return info

    # 가로형 척도: 2행 × N열, 2행이 전부 `N)`
    if rows == 2 and cols >= 3 and all(
            RE_SCALE_NUM_CELL.match(blk.cell(1, c)) for c in range(cols)):
        labels = [blk.cell(0, c) for c in range(cols)]
        info.update(kind="scale_horizontal", scale_points=cols, scale_labels=labels)
        return info

    # 가로형 척도(번호가 머리 칸에 있는 꼴): 1행이 전부 `N) 라벨`, 2행은 빈 칸
    if rows == 2 and cols >= 3 and all(
            RE_OPTION.match(blk.cell(0, c)) and not blk.cell(1, c) for c in range(cols)):
        labels = [norm_space(RE_OPTION.match(blk.cell(0, c)).group(2)) for c in range(cols)]
        info.update(kind="scale_horizontal", scale_points=cols, scale_labels=labels)
        return info

    # 행렬: 1행 헤더 + 항목 행. 본문 셀이 ①~⑤ 면 척도 행렬.
    body_has_circled = any(
        RE_CIRCLED.search(blk.cell(r, c))
        for r in range(1, rows) for c in range(1, cols)
    )
    header_has_circled = any(RE_CIRCLED.search(blk.cell(0, c)) for c in range(1, cols))
    if cols >= 3 and rows >= 2 and (body_has_circled or header_has_circled):
        items = []
        for r in range(1, rows):
            label = blk.cell(r, 0)
            if not label:
                continue
            om = RE_OPTION.match(label) or re.match(r"^\s*(\d+)\.\s*(.*)$", label, re.DOTALL)
            items.append({
                "no": int(om.group(1)) if om else None,
                "text": norm_space(om.group(2)) if om else norm_space(label),
                "raw": norm_space(label),
                "piping": [norm_space(x) for x in RE_PIPING.findall(label)] or None,
            })
        labels = [blk.cell(0, c) for c in range(1, cols)]
        info.update(
            kind="matrix_scale" if body_has_circled else "matrix_choice",
            scale_points=cols - 1 if body_has_circled else None,
            scale_labels=labels,
            choice_labels=None if body_has_circled else labels,
            items=items,
        )
        return info

    # 입력 그리드: 2열 표 (항목/응답, 구분/개수 …)
    if cols == 2 and rows >= 2:
        info.update(kind="input_grid", items=[
            {"label": blk.cell(r, 0), "spec": blk.cell(r, 1)} for r in range(rows)
        ])
        return info

    if rows == 1 and cols == 1:
        info.update(kind="box", text=blk.cell(0, 0))
        return info
    return info


def part_header(blk: Block) -> tuple[str, str] | None:
    """1셀 표가 파트 헤더면 (파트코드, 파트제목) 반환."""
    if blk.kind != "tbl" or blk.n_rows != 1 or blk.n_cols != 1:
        return None
    text = norm_space(blk.cell(0, 0))
    if not text or "\n" in blk.cell(0, 0):
        return None
    m = RE_PART_HEAD.match(text)
    if not m:
        return None
    code, title = m.group(1), norm_space(m.group(2))
    if not title or len(title) > 60:
        return None
    return code, text


# ═══════════════════════════════════════════════════════════════════════════
# 응답유형 판정
# ═══════════════════════════════════════════════════════════════════════════
def decide_response_type(markers: list[str], options: list[dict[str, Any]],
                         tables: list[dict[str, Any]], answer_context: str,
                         directives: dict[str, list[str]]) -> tuple[str, list[str]]:
    """(응답유형, 이상징후 메모) 반환."""
    anomalies: list[str] = []
    plain = [m.strip("[]") for m in markers]
    scale_tbl = next((t for t in tables
                      if t["kind"] in ("scale_horizontal", "matrix_scale")), None)
    matrix_tbl = next((t for t in tables
                       if t["kind"] in ("matrix_scale", "matrix_choice")), None)

    if parse_rank_spec(markers):
        return TYPE_RANK, anomalies
    if MARKER_MATRIX_ROW in plain or MARKER_MATRIX_ITEM in plain or MARKER_ROW_REQUIRED in plain:
        if matrix_tbl is None and MARKER_ROW_REQUIRED not in plain:
            anomalies.append("행렬 마커이나 행렬 표를 찾지 못함")
        return TYPE_MATRIX, anomalies
    if MARKER_MULTI in plain or any(RE_MULTI_LIMIT.fullmatch(m) for m in markers):
        return TYPE_MULTI, anomalies
    if MARKER_OPEN in plain:
        numeric = bool(directives["range"]) or bool(RE_NUMBER_ANSWER.search(answer_context))
        return (TYPE_NUMBER if numeric else TYPE_OPEN), anomalies
    if MARKER_SINGLE in plain:
        if scale_tbl is not None:
            return TYPE_SCALE, anomalies
        texts = {o["text"] for o in options}
        if len(options) == 2 and any(texts == y for y in YESNO_TOKENS):
            return TYPE_YESNO, anomalies
        if not options:
            anomalies.append(f"'{MARKER_SINGLE}' 표기이나 보기가 없음 — 개방 기입으로 보임")
            return TYPE_OPEN, anomalies
        return TYPE_SINGLE, anomalies

    # 마커 없음 — 구조로 추론
    if not markers:
        anomalies.append("응답유형 마커가 없어 구조로 추론함")
    if scale_tbl is not None:
        return TYPE_SCALE, anomalies
    texts = {o["text"] for o in options}
    if len(options) == 2 and any(texts == y for y in YESNO_TOKENS):
        return TYPE_YESNO, anomalies
    if options:
        return TYPE_SINGLE, anomalies
    if matrix_tbl is not None:
        return TYPE_MATRIX, anomalies
    return TYPE_UNKNOWN, anomalies


# ═══════════════════════════════════════════════════════════════════════════
# 본 파서
# ═══════════════════════════════════════════════════════════════════════════
def _options_from_tables(tables: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """표 칸의 `N) 보기` 를 번호 순으로 모은다. 번호만 있는 칸(`1)`)은 건너뛰고, 같은 번호는 처음 것만."""
    found: dict[int, dict[str, Any]] = {}
    for t in tables:
        if t.get("kind") in ("scale_horizontal", "matrix_scale", "box"):
            continue
        for row in t.get("cells") or []:
            for cell in row:
                for line in str(cell).split("\n"):
                    for piece in split_option_line(line.strip()):
                        opt = parse_option(piece)
                        if opt and opt["text"] and opt["no"] not in found:
                            found[opt["no"]] = opt
    return [found[k] for k in sorted(found)]


def _is_question_head(text: str, known_parts: set[str]) -> tuple[str, str] | None:
    m = RE_QCODE_HEAD.match(text)
    if not m:
        return None
    code, rest = m.group(1), m.group(2)
    prefix = re.match(r"^([A-Z]+)", code).group(1)
    if prefix not in known_parts and code != "ED":
        return None
    if not rest.strip():
        return None
    return code, rest


def _collect_parts(blocks: list[Block]) -> list[dict[str, Any]]:
    parts: list[dict[str, Any]] = []
    for blk in blocks:
        hit = part_header(blk)
        if hit:
            code, title = hit
            parts.append({"code": code, "title": title, "block": blk.index,
                          "question_codes": []})
    return parts


def parse(path: str | Path) -> dict[str, Any]:
    p = Path(path)
    if not p.is_file():
        raise QuestionnaireParseError(f"입력 파일이 없습니다: {p}")
    blocks, unmapped_syms = read_blocks(p)

    parts = _collect_parts(blocks)
    known_parts = {pt["code"] for pt in parts} | {"ED", "SQ", "DQ"}
    if not parts:
        raise QuestionnaireParseError(
            f"{p.name}: 파트 헤더(1셀 표 'SQ. …' / 'Part A. …')를 하나도 찾지 못했습니다."
        )

    warnings: list[str] = []
    questions: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    current_part = parts[0]["code"]
    part_by_block = {pt["block"]: pt for pt in parts}

    # 문항 앞에 붙는 [BASE:] 블록을 모아 두는 버퍼
    pending: dict[str, list[str]] = {"base": [], "range": [], "prog": [], "piping": []}
    pending_open = False

    def flush_pending(q: dict[str, Any]) -> None:
        nonlocal pending, pending_open
        merge_directives(q["directives"], pending)
        if pending["base"]:
            q["base_condition"] = pending["base"][0]
        pending = {"base": [], "range": [], "prog": [], "piping": []}
        pending_open = False

    for blk in blocks:
        if blk.index in part_by_block:
            current_part = part_by_block[blk.index]["code"]
            current = None
            pending = {"base": [], "range": [], "prog": [], "piping": []}
            pending_open = False
            continue

        if blk.kind == "p":
            raw = blk.text
            text = raw.strip()
            if not text:
                continue

            head = _is_question_head(text, known_parts)
            if head:
                code, rest = head
                markers = find_response_markers(text)
                body_directives = extract_directives(text)
                current = {
                    "index": len(questions) + 1,
                    "code": code,
                    # ED(종료 문항)는 전용 파트 헤더 표가 없다. DQ에 섞이지 않게 분리.
                    "part": "ED" if code == "ED" else current_part,
                    "block": blk.index,
                    "text_raw": text,                 # 공백까지 원문 그대로
                    "text_full": norm_space(text),    # 연속 공백만 정리
                    "text": norm_space(rest),
                    "response_markers": markers,
                    "response_type": TYPE_UNKNOWN,
                    "rank": parse_rank_spec(markers),
                    "scale": None,
                    "options": [],
                    "matrix": None,
                    "tables": [],
                    "directives": {"base": [], "range": [], "prog": [], "piping": []},
                    "base_condition": None,
                    "terminate_options": [],
                    "anomalies": [],
                }
                merge_directives(current["directives"], body_directives)
                flush_pending(current)
                questions.append(current)
                continue

            if RE_DIRECTIVE_LINE.match(text):
                d = extract_directives(text)
                forward = bool(d["base"]) or pending_open
                if forward:
                    merge_directives(pending, d)
                    pending_open = True
                elif current is not None:
                    merge_directives(current["directives"], d)
                else:
                    merge_directives(pending, d)
                    pending_open = True
                continue

            if current is not None and RE_OPTION.match(text):
                for piece in split_option_line(text):
                    opt = parse_option(piece)
                    if opt:
                        current["options"].append(opt)
                continue

            # 보기도 지시문도 아닌 본문 (오픈 기입란, 안내 문구 등)
            if current is not None:
                merge_directives(current["directives"], extract_directives(text))
                current.setdefault("extra_lines", []).append(norm_space(text))
            continue

        # 표
        info = classify_table(blk)
        if current is not None:
            current["tables"].append(info)
            merge_directives(current["directives"], extract_directives(blk.flat_text()))

    if not questions:
        raise QuestionnaireParseError(
            f"{p.name}: 문항을 하나도 찾지 못했습니다. 문항 단락이 'A1. …' 형태인지 확인하세요."
        )

    # ── 후처리: 응답유형 · 척도 · 행렬 ────────────────────────────────────
    for q in questions:
        # 응답 기입란은 문항 본문이 아니라 아래 딸린 줄·표에 있다.
        answer_context = " ".join(q.get("extra_lines", []))
        for t in q["tables"]:
            answer_context += " " + " ".join(
                " ".join(row) for row in t.get("cells", []))
        # 보기가 표 안에만 있는 선택형 (대분류|보기 묶음 표 · 보기×연도 체크 표). 단락 보기가 없을 때만 줍는다.
        # 유형 판정보다 먼저 해야 `[1개 선택]` + 보기 없음 → 개방형 오판을 피한다.
        choice_marked = any(
            m.strip("[]") in (MARKER_SINGLE, MARKER_MULTI)
            or RE_RANK_SPEC.fullmatch(m) or RE_MULTI_LIMIT.fullmatch(m)
            for m in q["response_markers"])
        if not q["options"] and choice_marked:
            q["options"] = _options_from_tables(q["tables"])
            if q["options"]:
                q["options_in_table"] = True
        rtype, anomalies = decide_response_type(
            q["response_markers"], q["options"], q["tables"],
            answer_context, q["directives"])
        q["response_type"] = rtype
        q["anomalies"].extend(anomalies)

        scale_tbl = next((t for t in q["tables"]
                          if t["kind"] in ("scale_horizontal", "matrix_scale")), None)
        if scale_tbl:
            q["scale"] = {
                "points": scale_tbl.get("scale_points"),
                "labels": scale_tbl.get("scale_labels"),
                "layout": "matrix" if scale_tbl["kind"] == "matrix_scale" else "horizontal",
            }
            if scale_tbl["kind"] == "scale_horizontal" and not q["options"]:
                q["options"] = [
                    {"no": i + 1, "raw": lab, "text": norm_space(lab), "is_other": False,
                     "is_none": False, "has_open_blank": False, "terminate": False,
                     "piping": None}
                    for i, lab in enumerate(scale_tbl.get("scale_labels") or [])
                ]

        matrix_tbl = next((t for t in q["tables"]
                           if t["kind"] in ("matrix_scale", "matrix_choice")), None)
        if matrix_tbl:
            q["matrix"] = {
                "items": matrix_tbl.get("items") or [],
                "column_labels": matrix_tbl.get("scale_labels")
                or matrix_tbl.get("choice_labels"),
                "is_scale": matrix_tbl["kind"] == "matrix_scale",
            }

        q["terminate_options"] = [o["no"] for o in q["options"] if o["terminate"]]
        q["n_options"] = len(q["options"])
        if q["response_type"] == TYPE_RANK and not q["options"]:
            q["anomalies"].append("순위 문항인데 보기 목록이 비어 있음")
        if q["response_type"] in (TYPE_SINGLE, TYPE_MULTI) and not q["options"]:
            q["anomalies"].append("선택형인데 보기 목록이 비어 있음")

    # 파트별 문항 코드 (ED는 헤더 표가 없으므로 가상 파트를 만든다)
    part_index = {pt["code"]: pt for pt in parts}
    for q in questions:
        pt = part_index.get(q["part"])
        if pt is None:
            pt = {"code": q["part"], "title": f"{q['part']}. (파트 헤더 없음)",
                  "block": q["block"], "question_codes": [], "synthetic": True}
            parts.append(pt)
            part_index[q["part"]] = pt
        pt["question_codes"].append(q["code"])

    # 중복 코드 검사
    seen: dict[str, int] = {}
    for q in questions:
        if q["code"] in seen:
            warnings.append(f"문항코드 중복: {q['code']} (블록 {seen[q['code']]} / {q['block']})")
        seen[q["code"]] = q["block"]

    for q in questions:
        for a in q["anomalies"]:
            warnings.append(f"{q['code']}: {a}")
    if unmapped_syms:
        warnings.append(
            "매핑 테이블에 없는 심볼 폰트 문자 (ASCII 폴백 적용): "
            + ", ".join(sorted(unmapped_syms))
        )

    counts = {
        "n_questions": len(questions),
        "by_part": {pt["code"]: len(pt["question_codes"]) for pt in parts},
        "by_response_type": {},
        "n_scale": sum(1 for q in questions if q["scale"]),
        "n_rank": sum(1 for q in questions if q["rank"]),
        "n_multi": sum(1 for q in questions if q["response_type"] == TYPE_MULTI),
        "scale_codes": [q["code"] for q in questions if q["scale"]],
        "rank_codes": [q["code"] for q in questions if q["rank"]],
        "multi_codes": [q["code"] for q in questions
                        if q["response_type"] == TYPE_MULTI],
    }
    for q in questions:
        counts["by_response_type"][q["response_type"]] = \
            counts["by_response_type"].get(q["response_type"], 0) + 1

    return {
        "meta": {
            "schema_version": SCHEMA_VERSION,
            "parser": Path(__file__).name,
            "source": str(p.resolve()),
            "survey_title": _survey_title(blocks),
            "cover": _cover_info(blocks),
            "generated_at": _dt.datetime.now().isoformat(timespec="seconds"),
            "counts": counts,
            "warnings": warnings,
        },
        "parts": parts,
        "questions": questions,
    }


def _survey_title(blocks: list[Block]) -> str | None:
    for blk in blocks[:3]:
        if blk.kind == "tbl" and blk.n_rows == 1 and blk.n_cols == 2:
            t = norm_space(blk.cell(0, 1))
            if t:
                return t
    return None


def _cover_info(blocks: list[Block]) -> dict[str, Any]:
    info: dict[str, Any] = {}
    lines: list[str] = []
    for blk in blocks[:12]:
        if blk.kind == "p" and blk.stripped:
            lines.append(norm_space(blk.text))
        elif blk.kind == "tbl" and blk.n_cols >= 3:
            info["quota_table"] = blk.cells
    info["lines"] = lines
    for ln in lines:
        if ln.startswith("■ 샘플 수"):
            info["sample"] = ln.split(":", 1)[-1].strip()
    return info


# ═══════════════════════════════════════════════════════════════════════════
def summarize(doc: dict[str, Any]) -> str:
    c = doc["meta"]["counts"]
    lines = [
        f"문항 {c['n_questions']}개 파싱 완료 — 파트별 {c['by_part']}",
        f"  응답유형 분포 : {c['by_response_type']}",
        f"  척도 문항 {c['n_scale']}개 : {', '.join(c['scale_codes']) or '-'}",
        f"  순위 문항 {c['n_rank']}개 : {', '.join(c['rank_codes']) or '-'}",
        f"  복수 문항 {c['n_multi']}개 : {', '.join(c['multi_codes']) or '-'}",
    ]
    if doc["meta"]["warnings"]:
        lines.append(f"  ⚠ 경고 {len(doc['meta']['warnings'])}건:")
        lines += [f"     - {w}" for w in doc["meta"]["warnings"]]
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="설문지 .docx → questionnaire.json")
    ap.add_argument("input", help="설문지 .docx 경로")
    ap.add_argument("output", help="출력 questionnaire.json 경로")
    ap.add_argument("--strict", action="store_true", help="경고가 있으면 출력하지 않음")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    try:
        doc = parse(args.input)
    except QuestionnaireParseError as e:
        print(f"[parse_questionnaire] 파싱 실패: {e}", file=sys.stderr)
        return 2

    if args.strict and doc["meta"]["warnings"]:
        print("[parse_questionnaire] --strict: 경고가 있어 출력하지 않습니다.", file=sys.stderr)
        for w in doc["meta"]["warnings"]:
            print(f"  - {w}", file=sys.stderr)
        return 3

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    if not args.quiet:
        print(summarize(doc))
        print(f"→ {out}  ({out.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
