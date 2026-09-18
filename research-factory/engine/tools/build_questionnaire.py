#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""설문 명세(JSON) → 엠브레인 양식 설문지 .docx.

    python build_questionnaire.py <설문명세.json> <출력.docx>

업무 흐름 ③ 설문지 구성. 구조화 엑셀을 사람이(또는 Claude 가) 읽고 문항 문장·보기·로직을 정한 결과가
`설문명세.json` 이고, 이 도구는 그것을 **서식만** 입힌다. 문장을 짓거나 보기를 보태지 않는다.
명세 형식은 `knowledge/50_설문지-작성.md` 4절. 양식의 틀(로고·여백·글꼴)은 `templates/TEMPLATE_설문지.docx`.
저장 뒤 문서 위생(R-137, `doc_hygiene`)을 건다 — 작성자·최종 수정자를 조직명으로, lastPrinted 제거.

하우스 서식 (로직 구성 예시 · 2025~2026 납품 설문지에서 뽑음)
    · 맑은 고딕 9pt, A4, 파트 머리는 검정 바탕 흰 굵은 글씨 1셀 표 (`SQ. Screening Question` · `PART A. …`)
    · 문항은 굵은 번호 `A1.` + 문장 + 응답 표기 `[1개 선택]` · `[모두 선택]` · `[1순위 필수, 3순위 선택]` · `[직접 입력]` · `[행별 1개 선택]`
    · 보기는 `1) …` 들여쓰기 단락, 조사 중단 보기는 빨간 `→ 조사 중단`
    · `[BASE: …]` 는 문항 **위**, `[PROG: …]` · `[RANGE: …]` 는 보기 **아래**, 파이핑 `[#A2 응답값#]` — 모두 파란 글씨
    · 척도는 2행 표(1행 라벨 회색 바탕, 2행 `1)`~`5)`), 행렬은 `구분` 머리 + `1. 항목` 행 + `①`~`⑤` 칸
산출물은 `parse_questionnaire.py` 가 그대로 읽는다 — `check_questionnaire.py --docx` 가 왕복을 검사한다.
    · 파서는 문항 사이 지시문을 이웃 문항에 붙여 읽는다. `kind: info` 의 `base` 는 **다음 문항**의 BASE 가 되고
      (그 문항에 BASE 가 따로 있어도 info 쪽이 이긴다), info 의 PROG 는 앞 문항(파트 첫머리면 다음 문항)에 붙는다.
      `prog_before` 는 그 문항에 `base` 가 있을 때만 그 문항에 붙고, 없으면 **앞 문항** PROG 로 읽힌다.
      노출 조건은 문항마다 `base` 로 단다 — /웹테스트 가 파싱 결과의 BASE 로 시나리오를 짠다.

문장 안 인라인 표기
    [#…#] · [#…]  [BASE: …]  [PROG: …]  [RANGE: …]  → 파란 글씨
    [쿼터]  → 조사 중단                          → 빨간 글씨 (화살표는 Wingdings 심볼)
    **굵게**  __밑줄__  ==노란 형광 (사람이 채울 자리)==  줄바꿈은 \\n
    기입란 `___년 ___개월` 처럼 밑줄 3개 이상 이어진 것은 밑줄 표기로 읽지 않는다.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE.parent / "templates" / "TEMPLATE_설문지.docx"

FONT = "맑은 고딕"
SIZE_PT = 9
BLUE, RED, WHITE = "0000FF", "FF0000", "FFFFFF"
FILL_PART, FILL_HEAD, FILL_QUOTA, FILL_SUM = "000000", "D9D9D9", "595959", "F2F2F2"
TABLE_W = 10471                     # twips — 예시 설문지의 표 폭
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮"     # 맑은 고딕 글리프가 ⑮ 까지 (KS X 1001) — 행렬 열 상한

# 척도 라벨 사전 — 1점이 부정, 끝점이 긍정 (하우스 규칙: 오름차순)
SCALE_PRESETS: dict[str, list[str]] = {
    "만족": ["전혀 만족하지 않는다", "만족하지 않는다", "보통이다", "만족한다", "매우 만족한다"],
    "동의": ["전혀 그렇지 않다", "그렇지 않다", "보통이다", "그렇다", "매우 그렇다"],
    "중요": ["전혀 중요하지 않다", "중요하지 않다", "보통이다", "중요하다", "매우 중요하다"],
    "추천": ["추천할 의향이 전혀 없다", "추천할 의향이 없다", "보통이다", "추천할 의향이 있다", "추천할 의향이 매우 있다"],
    "이용의향": ["전혀 이용 의향이 없다", "이용 의향이 없다", "보통이다", "이용 의향이 있다", "매우 이용 의향이 있다"],
    "구매의향": ["전혀 구매 의향이 없다", "구매 의향이 없다", "보통이다", "구매 의향이 있다", "매우 구매 의향이 있다"],
    "지속이용": ["지속적으로 이용할 의향이 전혀 없다", "지속적으로 이용할 의향이 없다", "보통이다",
             "지속적으로 이용할 의향이 있다", "지속적으로 이용할 의향이 매우 있다"],
    "기여": ["전혀 기여하지 못했다", "기여하지 못했다", "보통이다", "기여했다", "매우 기여했다"],
    "도움": ["전혀 도움이 되지 않았다", "도움이 되지 않았다", "보통이다", "도움이 되었다", "매우 도움이 되었다"],
    "이해": ["전혀 이해하지 못했다", "이해하지 못했다", "보통이다", "이해했다", "매우 이해했다"],
    "신뢰": ["전혀 신뢰가 가지 않는다", "신뢰가 가지 않는다", "보통이다", "신뢰가 간다", "매우 신뢰가 간다"],
    "매력": ["전혀 매력적이지 않다", "매력적이지 않다", "보통이다", "매력적이다", "매우 매력적이다"],
    "필요": ["전혀 필요하지 않다", "필요하지 않다", "보통이다", "필요하다", "매우 필요하다"],
    "인지": ["전혀 모른다", "잘 모른다", "보통이다", "어느 정도 안다", "매우 잘 안다"],
}

# 밑줄 표기는 앞뒤가 밑줄·공백이 아닌 `__…__` 만 — 기입란 `___년 ___개월` 을 표기로 삼키지 않는다
RE_INLINE = re.compile(
    r"(\[#[^\]]*\]|\[(?:BASE|PROG|RANGE)\s*:.*?\]|\[쿼터\]|\*\*.+?\*\*|(?<!_)__(?![_\s]).+?(?<![_\s])__(?!_)|==.+?==|→\s*조사\s*중단|\n)",
    re.DOTALL,
)


class SpecError(ValueError):
    pass


# ═══════════════════════════════════════════════════════════════════════════
# 저수준 서식
# ═══════════════════════════════════════════════════════════════════════════
def _oxml():
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    return OxmlElement, qn


def _font(run, bold=False, color=None, underline=False, size=SIZE_PT, highlight=False):
    from docx.shared import Pt, RGBColor
    OxmlElement, qn = _oxml()
    run.font.name = FONT
    run.font.size = Pt(size)
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(k), FONT)
    if bold:
        run.bold = True
    if underline:
        run.underline = True
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if highlight:
        from docx.enum.text import WD_COLOR_INDEX
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW


def _arrow(par, color=RED):
    """Wingdings → (원문 설문지와 같은 심볼). parse_questionnaire 가 '→' 로 되살린다."""
    OxmlElement, qn = _oxml()
    run = par.add_run()
    _font(run, color=color)
    sym = OxmlElement("w:sym")
    sym.set(qn("w:font"), "Wingdings")
    sym.set(qn("w:char"), "F0E0")
    run._r.append(sym)


def rich(par, text: str, bold=False, color=None, underline=False, size=SIZE_PT, highlight=False):
    """인라인 표기를 해석해 런으로 붙인다."""
    if text is None:
        return
    for i, tok in enumerate(RE_INLINE.split(str(text))):
        if not tok:
            continue
        mark = i % 2 == 1        # split 은 [평문, 표기, 평문, …] — 평문의 `________`(기입 줄)은 표기로 풀지 않는다
        if tok == "\n":
            par.add_run().add_break()
            continue
        if tok.startswith("[#") or re.match(r"\[(BASE|PROG|RANGE)\s*:", tok):
            r = par.add_run(tok)
            _font(r, bold=bold, color=BLUE, underline=underline, size=size)
        elif tok == "[쿼터]":
            r = par.add_run(tok)
            _font(r, bold=bold, color=RED, size=size)
        elif mark and tok.startswith("**") and tok.endswith("**"):
            rich(par, tok[2:-2], bold=True, color=color, underline=underline, size=size, highlight=highlight)
        elif mark and tok.startswith("__") and tok.endswith("__"):
            rich(par, tok[2:-2], bold=bold, color=color, underline=True, size=size, highlight=highlight)
        elif mark and tok.startswith("==") and tok.endswith("=="):
            rich(par, tok[2:-2], bold=bold, color=color, underline=underline, size=size, highlight=True)
        elif re.fullmatch(r"→\s*조사\s*중단", tok):
            _arrow(par)
            r = par.add_run(" 조사 중단")
            _font(r, color=RED, size=size)
        else:
            r = par.add_run(tok)
            _font(r, bold=bold, color=color, underline=underline, size=size, highlight=highlight)


def _spacing(par, before=0, after=0, line=None, left=None):
    from docx.shared import Pt, Twips
    pf = par.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing = line
    if left is not None:
        pf.left_indent = Twips(left)


def _shade(cell, fill):
    OxmlElement, qn = _oxml()
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def _borders(table):
    OxmlElement, qn = _oxml()
    tblpr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        b.append(e)
    tblpr.append(b)


def _widths(table, widths: list[int]):
    from docx.shared import Twips
    table.autofit = False
    for row in table.rows:
        for i, c in enumerate(row.cells):
            if i < len(widths):
                c.width = Twips(widths[i])


def _center(cell):
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _word_break(par):
    """한글을 글자가 아니라 어절 단위로 줄바꿈. 원본 설문지처럼 값 없는 `<w:wordWrap/>` 이어야 Word 가 따른다
    (`w:val="0"` 은 글자 단위로 자른다 — 실측)."""
    OxmlElement, qn = _oxml()
    ppr = par._p.get_or_add_pPr()
    if ppr.find(qn("w:wordWrap")) is None:
        ppr.insert(0, OxmlElement("w:wordWrap"))


def cell_text(cell, text, bold=False, color=None, fill=None, center=False, size=SIZE_PT):
    # 칸 안 줄바꿈은 단락이 아니라 <w:br/> 로 — parse_questionnaire 는 칸의 단락을 이어 붙여 읽는다
    # (단락으로 나누면 「4) 베트남5) 인도네시아」가 된다).
    par = cell.paragraphs[0]
    _spacing(par, 0, 0)
    _word_break(par)
    rich(par, str(text), bold=bold, color=color, size=size)
    if fill:
        _shade(cell, fill)
    if center:
        _center(cell)


# ═══════════════════════════════════════════════════════════════════════════
# 블록 렌더러
# ═══════════════════════════════════════════════════════════════════════════
class Builder:
    def __init__(self, spec: dict[str, Any]):
        import docx
        self.spec = spec
        self.doc = docx.Document(str(TEMPLATE))
        self.review: list[str] = []
        self._fill_title(spec.get("meta", {}).get("title") or "설문지")

    # ── 공통 ───────────────────────────────────────────────────────────────
    def _fill_title(self, title):
        cell = self.doc.tables[0].rows[0].cells[1]
        par = cell.paragraphs[0]
        for r in par.runs:
            if "{{TITLE}}" in r.text:
                r.text = r.text.replace("{{TITLE}}", title)
                _font(r, bold=True, size=14)

    def para(self, text="", bold=False, color=None, before=0, after=0, left=None, line=None, highlight=False):
        p = self.doc.add_paragraph()
        _word_break(p)
        _spacing(p, before, after, line=line, left=left)
        rich(p, text, bold=bold, color=color, highlight=highlight)
        return p

    def blank(self):
        return self.para("")

    def directive(self, kind: str, body: str):
        body = str(body).strip()
        if re.match(r"\[(BASE|PROG|RANGE)\s*:", body):     # 이미 감싼 것만 그대로 — `[#SQ1 응답값#] …` 은 감싼다
            text = body
        else:
            text = f"[{kind}: {body}]"
        return self.para(text, color=BLUE, before=5, after=2, line=1.15)

    def table(self, rows: int, cols: int, widths: list[int] | None = None):
        t = self.doc.add_table(rows=rows, cols=cols)
        _borders(t)
        if widths is None:
            widths = [TABLE_W // cols] * cols
        _widths(t, widths)
        return t

    def box(self, text: str, fill=None, bold=False):
        t = self.table(1, 1, [TABLE_W])
        cell_text(t.cell(0, 0), text, bold=bold, fill=fill)
        return t

    # ── 표지 ───────────────────────────────────────────────────────────────
    def cover(self):
        meta = self.spec.get("meta", {})
        cov = meta.get("cover") or {}
        greeting = meta.get("greeting")
        if cov.get("target"):
            self.para("■ **조사 대상자**", before=6)
            for ln in _as_list(cov["target"]):          # 문자열 하나로 줘도 한 줄
                self.para(f"- {ln}")
        for key, label in (("method", "조사 방법"), ("area", "조사 지역"), ("sample", "샘플 수")):
            if cov.get(key):
                self.para(f"■ **{label}** : {cov[key]}", before=4)
        for ln in _as_list(cov.get("extra")):
            self.para(ln, before=4)
        qt = cov.get("quota_table")
        if qt:
            self.para("■ **조사 쿼터**", before=4)
            t = self.table(len(qt), len(qt[0]))
            for r, row in enumerate(qt):
                for c, v in enumerate(row):
                    if r == 0:
                        cell_text(t.cell(r, c), v, bold=True, color=WHITE, fill=FILL_QUOTA, center=True)
                    else:
                        cell_text(t.cell(r, c), v, center=c > 0)
        if greeting:
            self.blank()
            self.box(greeting)
        self.blank()

    # ── 파트 ───────────────────────────────────────────────────────────────
    def part(self, pt: dict[str, Any]):
        code = pt["code"]
        title = pt.get("title") or {"SQ": "Screening Question", "DQ": "Demographic Question"}.get(code, "")
        head = f"{code}. {title}" if code in ("SQ", "DQ") else f"PART {code}. {title}"
        self.box(head, fill=FILL_PART, bold=True)
        # 흰 글씨
        for r in self.doc.tables[-1].rows[0].cells[0].paragraphs[0].runs:
            _font(r, bold=True, color=WHITE)
        if pt.get("intro"):
            self.para(pt["intro"], before=4)
        for pg in _as_list(pt.get("prog")):
            self.directive("PROG", pg)
        for it in pt.get("items", []):
            self.item(it)
        self.blank()

    # ── 문항 ───────────────────────────────────────────────────────────────
    def item(self, it: dict[str, Any]):
        kind = it.get("kind", "single")
        if kind == "text":
            self.para(it["text"], before=4, bold=it.get("bold", False))
            for pg in _as_list(it.get("prog")):
                self.directive("PROG", pg)
            return
        if kind == "info":
            if it.get("base"):
                self.directive("BASE", it["base"])
            self.blank()
            self.box(it["text"], bold=it.get("bold", False))
            for pg in _as_list(it.get("prog")) or ["안내문 별도 페이지 제시"]:
                self.directive("PROG", pg)
            return
        if kind == "closing":
            self.closing(it)
            return
        if kind == "table":                     # 참고 표 (응답자 분류 · 파트 진입 조건 등). 첫 행이 머리
            rows = it["rows"]
            if it.get("title"):
                self.para(it["title"], bold=True, before=4)
            ncol = max(len(r) for r in rows)
            t = self.table(len(rows), ncol)
            for r, row in enumerate(rows):
                for c in range(ncol):
                    v = row[c] if c < len(row) else ""
                    cell_text(t.cell(r, c), v, bold=r == 0, fill=FILL_HEAD if r == 0 else None, center=r == 0)
            return

        code = it["code"]
        if it.get("base"):
            self.directive("BASE", it["base"])
        for pg in _as_list(it.get("prog_before")):
            self.directive("PROG", pg)

        marker = it.get("marker") or default_marker(it)
        p = self.doc.add_paragraph()
        _word_break(p)
        _spacing(p, before=5, after=2, line=1.15)
        review = bool(it.get("review"))
        r = p.add_run(f"{code}. ")
        _font(r, bold=True, highlight=review)
        if it.get("quota"):
            rich(p, "[쿼터] ")
        rich(p, it["text"])
        if it.get("note"):
            rich(p, "\n" + it["note"])
        if marker:
            rich(p, f" {marker}")
        if review:
            self.review.append(f"{code}: {it['review']}")

        for bx in _as_list(it.get("boxes")):
            self.box(bx)

        getattr(self, f"body_{kind}", self.body_single)(it)

        for pg in auto_prog(it) + _as_list(it.get("prog")):
            self.directive("PROG", pg)
        if it.get("range"):
            self.directive("RANGE", it["range"])

    # 보기 목록 (단일·복수·순위·예아니오)
    def options(self, it: dict[str, Any]):
        opts = norm_options(it.get("options") or [])
        cols = int(it.get("cols") or 1)
        if cols <= 1:
            for o in opts:
                self.para(option_text(o), left=567)
            return
        for i in range(0, len(opts), cols):
            chunk = opts[i:i + cols]
            self.para("\t\t".join(option_text(o) for o in chunk), left=567)

    def body_single(self, it):
        if it.get("options_table"):
            self.options_table(it)
        elif it.get("columns"):
            self.options_grid(it)
        else:
            self.options(it)

    def options_grid(self, it):
        """보기 × 열(연도 등) 체크 표. 칸마다 보기 번호를 둔다 (연도별 복수 선택)."""
        opts = norm_options(it.get("options") or [])
        cols = list(it["columns"])
        first = 4200
        rest = (TABLE_W - first) // len(cols)
        t = self.table(len(opts) + 1, len(cols) + 1, [first] + [rest] * len(cols))
        cell_text(t.cell(0, 0), it.get("row_head") or "구분", bold=True, fill=FILL_HEAD, center=True)
        for c, h in enumerate(cols, start=1):
            cell_text(t.cell(0, c), h, bold=True, fill=FILL_HEAD, center=True)
        for r, o in enumerate(opts, start=1):
            cell_text(t.cell(r, 0), option_text(o))
            for c in range(1, len(cols) + 1):
                cell_text(t.cell(r, c), f"{o['no']})", center=True)

    body_multi = body_single
    body_rank = body_single

    def body_list(self, it):
        """리스트·드롭다운·지도 선택. 보기(시도 목록 등)를 주면 싣고, 없으면 PROG 로만 남긴다."""
        if it.get("options"):
            self.options(it)

    def body_yesno(self, it):
        it = dict(it)
        it.setdefault("options", ["예", "아니오"])
        it.setdefault("cols", 2)
        self.options(it)

    def options_table(self, it):
        """보기를 묶음 표로 (대분류 | 보기들). 예: 대륙별 국가."""
        groups = it["options_table"]
        t = self.table(len(groups) + 1, 2, [2200, TABLE_W - 2200])
        heads = it.get("options_table_head") or ["구분", "보기"]
        for c, h in enumerate(heads):
            cell_text(t.cell(0, c), h, bold=True, fill=FILL_HEAD, center=True)
        n = 0
        for r, g in enumerate(groups, start=1):
            cell_text(t.cell(r, 0), g["group"], center=True)
            lines = []
            for raw, o in zip(g["options"], norm_options(g["options"])):
                explicit = isinstance(raw, dict) and "no" in raw
                n = o["no"] if explicit else n + 1      # 묶음을 넘어 번호를 잇는다
                o["no"] = n
                lines.append(option_text(o))
            cols = int(g.get("cols") or 4)
            rows = ["\t".join(lines[i:i + cols]) for i in range(0, len(lines), cols)]
            cell_text(t.cell(r, 1), "\n".join(rows))

    def body_scale(self, it):
        labels = scale_labels(it)
        extra = _as_list((it.get("scale") or {}).get("extra"))
        allc = labels + extra
        t = self.table(2, len(allc))
        for c, lab in enumerate(allc):
            cell_text(t.cell(0, c), lab, bold=True, fill=FILL_HEAD, center=True)
            cell_text(t.cell(1, c), f"{c + 1})", center=True)

    def body_matrix(self, it):
        labels = scale_labels(it) if it.get("scale") else list(it["columns"])
        extra = _as_list((it.get("scale") or {}).get("extra"))
        allc = labels + extra
        if len(allc) > len(CIRCLED):
            raise SpecError(f"{it.get('code')}: 행렬 열이 {len(allc)}개 — {len(CIRCLED)}개(①~⑮)까지만 그린다. 문항을 나눈다")
        rows = it["rows"]
        first = 3600
        rest = (TABLE_W - first) // len(allc)
        t = self.table(len(rows) + 1, len(allc) + 1, [first] + [rest] * len(allc))
        cell_text(t.cell(0, 0), it.get("row_head") or "구분", bold=True, fill=FILL_HEAD, center=True)
        for c, lab in enumerate(allc, start=1):
            head = f"{CIRCLED[c - 1]} {lab}" if not it.get("scale") else lab
            cell_text(t.cell(0, c), head, bold=True, fill=FILL_HEAD, center=True)
        for r, row in enumerate(rows, start=1):
            label = row if isinstance(row, str) else row["text"]
            cell_text(t.cell(r, 0), f"{r}. {label}")
            for c in range(1, len(allc) + 1):
                mark = CIRCLED[c - 1] if it.get("scale") else f"{c})"
                cell_text(t.cell(r, c), mark, center=True)

    def body_number(self, it):
        unit = it.get("unit", "")
        cols = it.get("columns")
        if cols:
            t = self.table(2, len(cols))
            for c, h in enumerate(cols):
                cell_text(t.cell(0, c), h, fill=FILL_HEAD, center=True)
                cell_text(t.cell(1, c), f"(          ) {unit}".rstrip(), center=True)
            return
        if it.get("fields"):
            self.body_grid(it)
            return
        self.para(f"(                    ) {unit}".rstrip(), left=567)

    def body_open(self, it):
        if it.get("fields"):
            self.body_grid(it)
            return
        if it.get("lines"):
            for ln in _as_list(it["lines"]):
                self.para(ln, left=567)
            return
        if it.get("box", True):
            t = self.table(1, 1, [TABLE_W])
            t.rows[0].height = __import__("docx").shared.Twips(900)
            cell_text(t.cell(0, 0), "")
        else:
            self.para("(                                        )", left=567)

    def body_grid(self, it):
        fields = it["fields"]
        head = it.get("fields_head")
        t = self.table(len(fields) + (1 if head else 0), 2, [3000, TABLE_W - 3000])
        r0 = 0
        if head:
            for c, h in enumerate(head):
                cell_text(t.cell(0, c), h, bold=True, fill=FILL_HEAD, center=True)
            r0 = 1
        for r, f in enumerate(fields, start=r0):
            if isinstance(f, str):
                f = {"label": f}
            cell_text(t.cell(r, 0), f["label"], bold=True, fill=FILL_SUM if f.get("sum") else None)
            ans = f.get("answer") or f"(          ) {f.get('unit', '')}".rstrip()
            cell_text(t.cell(r, 1), ans, fill=FILL_SUM if f.get("sum") else None)

    # ── 끝맺음 ──────────────────────────────────────────────────────────────
    def closing(self, it: dict[str, Any] | None = None):
        cl = it or self.spec.get("meta", {}).get("closing") or {}
        self.blank()
        gift = cl.get("gift")
        if not gift:
            self.para(cl.get("text") or "– 설문에 응답해 주셔서 감사합니다. –", bold=True, before=6)
            return
        self.para("긴 시간 동안 설문에 참여해 주셔서 감사합니다.", before=6)
        self.para(f"**설문에 참여해주신 분들께는 __{gift}__을 발송할 예정입니다.**")
        p = self.doc.add_paragraph()
        _spacing(p, before=5, after=2, line=1.15)
        r = p.add_run("ED. ")
        _font(r, bold=True)
        rich(p, "기프티콘 수령을 원하시는 경우 “예”에 체크를 해주십시오.(원하지 않을 경우 “아니오” 체크) [1개 선택]")
        self.para("1) 예\t\t2) 아니오", left=567)
        company = cl.get("company", "(주)마크로밀엠브레인")
        keep = cl.get("retention", "조사완료 후 3개월간")
        # 수집 항목은 한 값 — 고지(3.)·입력 칸·PROG 가 같은 목록을 쓴다 (하우스 원문은 고지 없이 이름까지 받았다)
        fields = [str(x) for x in _as_list(cl.get("items"))] or ["휴대전화번호"]
        what = ", ".join(fields)
        # 법적 고지 문안 — 박스 전체를 노란 형광으로 두고 확인 목록에 올린다 (개인정보 담당 확인 뒤 납품)
        self.box(
            "==**[개인정보 수집 및 이용 동의]**\n"
            "조사 사례인 모바일 상품권 발송을 위해 아래와 같이 귀하의 개인정보를 수집합니다.\n"
            "제공하신 개인정보는 목적 외에는 사용하지 않으며, 이용 기간 후 즉시 파기할 예정입니다.\n"
            f"**1. 제공받는 자 :** {company}\n"
            "**2. 이용 목적 :** 모바일 기프티콘 발송\n"
            f"**3. 개인정보 항목 :** {what}\n"
            f"**4. 보유 및 이용기간 :** {keep}\n"
            "※ 귀하는 개인정보 수집·이용 동의를 거부할 권리가 있습니다.\n"
            "※ 기프티콘 수령을 원하셨으나 개인정보 수집 관련 동의 여부에 \"비동의\"를 체크하신 경우 기프티콘 수령을 받으실 수 없습니다.\n"
            "※ 미기재 또는 잘못된 정보일 경우 기프티콘 발송에서 자동 제외될 수 있으며, 이럴 경우 재발송은 하지 않습니다.=="
        )
        # 동의 항목 (하우스 원본의 「동의 여부 · 입력 칸」 줄). 표로 둬야 파서가 ED 보기(예·아니오)에 섞지 않는다
        t = self.table(1 + len(fields), 2, [2200, TABLE_W - 2200])
        cell_text(t.cell(0, 0), "동의 여부", bold=True, fill=FILL_HEAD, center=True)
        cell_text(t.cell(0, 1), "1) 동의\t\t2) 비동의")
        for r, f in enumerate(fields, start=1):
            cell_text(t.cell(r, 0), f, bold=True, fill=FILL_HEAD, center=True)
            cell_text(t.cell(r, 1), "(           ) – (           ) – (           )" if re.search(r"전화|연락처", f) else "")
        self.directive("PROG", f"ED=1 응답자에게만 동의 여부 제시, 동의 여부 1) 동의하신 경우에만 {what} 입력하도록 구현")
        self.review.append(f"ED: 개인정보 수집·이용 동의 박스 — 고지 문안·수집 항목({what}) 확인 (노란 형광)")

    # ── 전체 ───────────────────────────────────────────────────────────────
    def build(self, out: Path):
        self.cover()
        for pt in self.spec["parts"]:
            self.part(pt)
        if self.spec.get("meta", {}).get("closing") is not False:
            self.closing()
        out.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(out))
        # 문서 위생 (R-137): 고객 송부본이다 — 작성자·최종 수정자를 조직명으로, lastPrinted 제거, 보이지 않는 문자 제거.
        # deckops.save 와 같은 꼴. 실패해도 저장은 살린다.
        try:
            import doc_hygiene
            tmp = out.with_suffix(out.suffix + ".hyg")
            doc_hygiene.clean_file(out, tmp, "엠브레인")
            tmp.replace(out)
        except Exception as e:                       # noqa: BLE001
            print(f"[build_questionnaire] 문서 위생 건너뜀: {e}", file=sys.stderr)
        return self.review


# ═══════════════════════════════════════════════════════════════════════════
# 명세 해석 도우미 (check_questionnaire 도 쓴다)
# ═══════════════════════════════════════════════════════════════════════════
def _as_list(v) -> list:
    if v is None or v is False:
        return []
    return list(v) if isinstance(v, (list, tuple)) else [v]


RE_OTHER = re.compile(r"^기타\s*(\(.*\))?\s*$|^기타\s*[:：(]")
RE_NONE = re.compile(r"(없음|없다|모름|모르겠|해당\s*없|경험\s*없)")


def norm_options(opts: list) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    n = 0
    for o in opts:
        d = {"text": o} if isinstance(o, str) else dict(o)
        if "no" in d:
            n = int(d["no"])
        else:
            n += 1
            d["no"] = n
        t = d["text"].strip()
        if d.get("other") is None:
            d["other"] = bool(RE_OTHER.match(t))
        if d["other"] and "(" not in t:
            d["text"] = f"{t}(            )"
        out.append(d)
    return out


def option_text(o: dict[str, Any]) -> str:
    s = f"{o['no']}) {o['text']}"
    if o.get("note"):
        s += f" {o['note']}"
    if o.get("terminate"):
        s += " → 조사 중단"
    return s


def scale_labels(it: dict[str, Any]) -> list[str]:
    sc = it.get("scale") or {}
    labels = sc.get("labels")
    if labels:
        return list(labels)
    preset = sc.get("preset") or "만족"
    if preset not in SCALE_PRESETS:
        raise SpecError(f"{it.get('code')}: 척도 preset '{preset}' 이 없습니다 — labels 를 직접 주세요 ({', '.join(SCALE_PRESETS)})")
    base = SCALE_PRESETS[preset]
    pts = int(sc.get("points") or 5)
    if pts != 5:
        raise SpecError(f"{it.get('code')}: 5점 외 척도는 labels 를 직접 주세요")
    return list(base)


def default_marker(it: dict[str, Any]) -> str:
    kind = it.get("kind", "single")
    if kind in ("single", "scale", "yesno", "list"):
        return "[1개 선택]"
    if kind == "multi":
        mx, mn = it.get("max"), it.get("min")
        if mx:
            return f"[최소 {mn or 1}개, 최대 {mx}개 선택]"
        return "[모두 선택]"            # min 만 있으면 PROG 로 남긴다 (auto_prog) — 파서는 `[최소 N개 …]` 단독 표기를 못 읽는다
    if kind == "rank":
        rk = it.get("rank") or {}
        req, opt = int(rk.get("required", 1)), rk.get("optional")
        if opt and int(opt) > req:
            return f"[{req}순위 필수, {int(opt)}순위 선택]"
        return f"[{req}순위 필수]"
    if kind == "matrix":
        return "[항목별 1개 선택]" if it.get("per_item") else "[행별 1개 선택]"
    if kind in ("number", "open"):
        return "[직접 입력]"
    return ""


def auto_prog(it: dict[str, Any]) -> list[str]:
    """보기 플래그에서 나오는 PROG 줄. 명세에 같은 뜻의 PROG 가 이미 있으면 만들지 않는다."""
    out: list[str] = []
    given = " ".join(_as_list(it.get("prog")))
    if it.get("kind") in ("multi",):
        for o in norm_options(it.get("options") or []):
            if o.get("exclusive") and "중복" not in given and "복수 선택 불가" not in given:
                out.append(f"{o['no']}) 복수 선택 불가")
        mn = str(it.get("min") or "").strip()        # max 없이 min 만 → 표기 대신 PROG (check 도 이 함수를 부른다)
        if mn.isdigit() and int(mn) > 1 and not it.get("max") and "최소" not in given and "MIN" not in given.upper():
            out.append(f"최소 {int(mn)}개 이상 선택")
    rot = it.get("rotation")
    if rot and "otation" not in given and "로테이션" not in given:
        opts = norm_options(it.get("options") or [])
        fixed = [o for o in opts if o.get("other") or o.get("exclusive") or o.get("fixed")]
        if rot is True and fixed:
            names = "·".join(("기타" if o.get("other") else re.sub(r"\(.*", "", o["text"]).strip()) for o in fixed)
            out.append(f"{names} 제외 보기 Rotation")
        elif rot is True:
            out.append("보기 Rotation")
        else:
            out.append(str(rot))
    return out


def load_spec(path: Path) -> dict[str, Any]:
    spec = json.loads(Path(path).read_text(encoding="utf-8"))
    if "parts" not in spec or not spec["parts"]:
        raise SpecError("명세에 parts 가 없습니다")
    return spec


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="설문 명세 JSON → 엠브레인 양식 설문지 .docx")
    ap.add_argument("spec")
    ap.add_argument("out")
    a = ap.parse_args(argv)
    try:
        spec = load_spec(Path(a.spec))
        review = Builder(spec).build(Path(a.out))
    except (SpecError, KeyError, IndexError, TypeError, ValueError) as e:
        # 명세 모양이 틀리면 트레이스백 대신 한 줄로 (KeyError 는 빠진 필드 이름)
        why = e if isinstance(e, SpecError) else f"{type(e).__name__}: {e}"
        print(f"[build_questionnaire] 명세 오류: {why}", file=sys.stderr)
        return 2
    n = sum(1 for pt in spec["parts"] for it in pt.get("items", []) if it.get("code"))
    print(f"설문지 생성: 파트 {len(spec['parts'])}개 · 문항 {n}개 → {a.out}")
    if review:
        print(f"  ⚑ 확인 필요(노란 번호) {len(review)}건:")
        for r in review:
            print(f"     - {r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
