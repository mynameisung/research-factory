#!/usr/bin/env python3
"""`slide_facts.json` → 제작가이드 `.xlsm`.

템플릿(LG 예시 기반)을 복제해 `본문_제작용` 시트를 채운다. 숨은 검증 시트 9종과
VBA(`lovelySeoyeon`, `검수 실행` 버튼)는 손대지 않는다 — `Verification_*` 시트에
"수정금지, 시트삭제 금지"라고 적혀 있고, 사내 RPA 애드인이 그 시트들을 읽는다.

본문은 **조사결과 블록만이 아니다.** `--toc 목차.md` 를 주면 보고서 순서 그대로
구조 블록을 함께 깐다 (guide/11 §5 · R-035/R-036, LG `본문_제작용` 실측 배치):

    표지 → 목차
    간지 Part Ⅰ. 조사 개요        Data = 세부 목차
      1. 조사목적 / 2. 조사설계 / 3. 응답자특성
    간지 Part Ⅱ. 요약 및 결론     Data = 세부 목차
      요약 꼭지들 / 결론
    간지 Part Ⅲ. 조사 결과        Data = 챕터 목록
      간지 1. 소유                Data = 그 챕터의 서브 타이틀 목록   ← 챕터마다 반복
        …본문 블록…

열 배치 (LG `본문_제작용` A1:A6 실측 — 라벨이 곧 A열이다. seq 열은 두지 않는다)
------------------------------------------------------------------------------
    A  블록 라벨          B  값 · 멘트 · 행수식어
    C  행 라벨            D  전체            E~ 배너

    순위 슬라이드는 전체 열이 둘로 갈린다:
    C  행 라벨            D  1순위          E  1~N순위          F~ 배너

빈 칸은 **빈 칸으로 둔다.** `-` 로 메우지 않는다 — LG 송부본의 `본문_제작용` 에는 `-`
가 한 개도 없다. `lovelySeoyeon` 매크로(R-037)가 빈칸을 `-` 로 바꾸는 것은 맞지만
VBA 원문이 `For Each c In Selection.Cells` 다: 연구원이 **고른 영역에만** 수동으로
거는 도구이지 시트 전체에 자동으로 도는 것이 아니다.

산출물 3벌
----------
    out/{이름}.xlsm              작업용 (치트시트·스니펫 포함)
    out/.baseline/{이름}.xlsm    생성 직후 스냅샷 — 사람이 고친 diff가 학습 데이터가 된다
    out/{이름}_제작팀송부본.xlsm   `_` 접두 시트 제거본

사용법
------
    python build_guide.py <slide_facts.json> <out/제작가이드.xlsm> [--toc work/목차.md]
        [--title "조사명"] [--date 2026-09-01]

`--toc` 를 주면 같은 폴더에서 부재료를 자동으로 찾는다 (없으면 그 블록만 건너뛴다):

    <facts 폴더>/tables.json     표1 응답자 특성 — 전체 열 (guide/10 §1)
    <toc 폴더>/조사개요.json      조사 목적 · 조사 설계 (조사개요·RFP 에서 뽑아 둔 것)
"""

from __future__ import annotations

import argparse
import io
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.styles.cell_style import StyleArray
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))   # build_facts 를 옆에서 가져온다

ENGINE = Path(__file__).resolve().parents[1]
TEMPLATE = ENGINE / "templates" / "TEMPLATE_제작가이드.xlsm"

#: 본문 시트 이름은 파일마다 다르다 (LG `본문_제작용` / SKEE `본문`). 먼저 잡히는 것을 쓴다.
BODY_SHEETS = ("본문_제작용", "본문")
TOC_SHEET = "목차(보고서용)"
RESEARCH_TOC_SHEET = "part 3 목차(연구부 확인용)"

LABELS = ["Slide title", "Sub title", "Chart type",
          "Base/Unit", "Question/Reference", "Data"]

COL_LABEL, COL_VALUE, COL_ROWLABEL = 1, 2, 3
COL_FIRST_DATA = 4           # 전체 (순위면 1순위)


def _num(v):
    """수치 칸 하나. `-` 는 값이 아니라 **값 없음** 표시이므로 빈 칸으로 돌린다.

    `build_facts.row_cells` 가 통계표의 빈 셀을 `-` 로 적어 두는데, LG `본문_제작용`
    실측(r202 등)은 그런 칸을 **빈 칸**으로 둔다 — 파일 전체에 `-` 가 0개다.
    """
    return None if isinstance(v, str) and v.strip() in ("", "-") else v


def _style(ws, row: int, col: int, *, bold=False, fill: str | None = None,
           align: str | None = None, wrap=False) -> None:
    c = ws.cell(row, col)
    if bold:
        c.font = Font(bold=True)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if align or wrap:
        c.alignment = Alignment(horizontal=align, wrap_text=wrap, vertical="center")


# ══ LG 예시 서식 재현 ════════════════════════════════════════════════════════
#
#  실측 대상 — engine/refs/예시_제작가이드_LG홈로봇_0317.xlsm · `본문_제작용`
#
#  ① 서식은 **셀이 아니라 행**에 걸린다.
#       Slide title 행  <row s=178|169 customFormat="1" ht="16.5" customHeight="1">
#       나머지 라벨 행  <row s=179|161 customFormat="1" ht="16.5" customHeight="1">
#       표 구간         행 서식 없음. ht 와 thickTop/thickBot 만 붙는다
#     행 서식을 안 걸고 템플릿이 물려준 것을 그대로 두면 **LG 의 원래 행 위치**에
#     노랑 띠가 남아 우리 블록과 어긋난다. 그래서 매 빌드마다 행 서식을 전부 지우고
#     새로 건다(`ws.row_dimensions.clear()`).
#
#  ② 섹션마다 서식이 **두 벌**이다. 색·테두리 체계는 같고 글꼴만 다르다.
#       구조(표지·목차·간지·조사개요) + 요약·결론 → 11pt (LG 전용 서체)  s=178/179, 196/197
#       본문(Part Ⅲ 조사 결과)                    →  9pt 맑은 고딕        s=169/161
#     스타일 인덱스를 박아 두지 않는다 — 템플릿의 노랑 행에서 **실측해** 가져온다.
#
#  ③ 6줄 라벨 칸은 밝은 띠 / 어두운 띠가 번갈아 든다 (LG A1~A6 실측).
#       Slide title  노랑 + 볼드          Sub title   밝은 띠 + 아래 실선
#       Chart type   어두운 띠 + 흰 글씨   Base/Unit   〃
#       Question     밝은 띠 + 위아래 실선  Data        어두운 띠 + 흰 글씨 + 위 실선
#
#  ④ 표는 **바깥 굵은 선 · 안쪽 가는 선**의 상자다 (LG r357~366 실측).
#       문항 제목 행   볼드 · 행 라벨 열부터 표 오른쪽 끝까지 병합
#       배너 그룹 행   그룹명 가로 병합 + 가운데 정렬 · 위 굵은 선
#       배너 세부 행   가운데 정렬 · 아래 굵은 선
#       모서리 칸      행 라벨 열이 두 행 세로 병합 (상자의 왼쪽 위 빈칸)
#       Base for %    위 굵은 선
#       데이터 행      행 라벨 왼쪽 정렬 · 수치 오른쪽 정렬
#       마지막 행      아래 굵은 선 (`계`·`【TOP2】`·`[N점 평균】` 무엇이든 위치로 정한다)

YELLOW = "FFFFFF00"
#: 라벨 띠 색. LG 는 테마 8(강조 4)의 tint 로 밝은 띠/어두운 띠를 만든다.
BAND_LIGHT = 0.7999816888943144
BAND_DARK = -0.249977111117893

_S_THIN = Side(style="thin", color=Color(theme=0, tint=0.0))
_T_THIN = Side(style="thin")
_T_THICK = Side(style="thick")

#: 표 안의 행 높이 (LG 실측)
HT_LABEL = 16.5
HT_PANEL = 16.2
HT_HEADER = 18.0
HT_DATA = 17.4


def _band(dark: bool) -> PatternFill:
    return PatternFill("solid",
                       fgColor=Color(theme=8, tint=BAND_DARK if dark else BAND_LIGHT))


class Skin:
    """한 계열의 서식 한 벌. 글꼴만 다르고 색·테두리는 LG 실측값 그대로다."""

    def __init__(self, name: str, size: float,
                 row_title: StyleArray | None, row_plain: StyleArray | None):
        self.row_title = row_title
        self.row_plain = row_plain
        self.font = Font(name=name, size=size)
        self.font_b = Font(name=name, size=size, bold=True)
        self.font_w = Font(name=name, size=size, color=Color(theme=0))
        self.fill_yellow = PatternFill("solid", fgColor=YELLOW)
        # 6줄 라벨 (LG A1~A6)
        self.labels = [
            (self.font_b, self.fill_yellow, None),
            (self.font, _band(False), Border(bottom=_S_THIN)),
            (self.font_w, _band(True), Border(top=_S_THIN, bottom=_S_THIN)),
            (self.font_w, _band(True), Border(top=_S_THIN, bottom=_S_THIN)),
            (self.font, _band(False), Border(top=_S_THIN, bottom=_S_THIN)),
            (self.font_w, _band(True), Border(top=_S_THIN)),
        ]

    # ── 라벨 6줄 ------------------------------------------------------------
    def label_cell(self, ws, row: int, idx: int) -> None:
        font, fill, border = self.labels[idx]
        c = ws.cell(row, COL_LABEL)
        c.font, c.fill = font, fill
        if border is not None:
            c.border = border
        c.alignment = Alignment(vertical="center")

    def value_cell(self, ws, row: int, idx: int) -> None:
        """값 열(B). Slide title 줄만 노랑 볼드고 나머지는 맨 글꼴이다 (LG B17/B348)."""
        c = ws.cell(row, COL_VALUE)
        if idx == 0:
            c.font, c.fill = self.font_b, self.fill_yellow
        else:
            c.font = self.font
        c.alignment = Alignment(vertical="center", wrap_text=False)

    # ── 행 서식 -------------------------------------------------------------
    def row_label(self, ws, row: int, title: bool = False) -> None:
        rd = ws.row_dimensions[row]
        src = self.row_title if title else self.row_plain
        if src is not None:
            # `customFormat` · `customHeight` 는 openpyxl 이 파생시키는 읽기 전용
            # 속성이다 — 행 서식과 높이만 주면 저장할 때 `customFormat="1"` 이 붙는다.
            rd._style = StyleArray(src)
        rd.height = HT_LABEL


def measure_skins(ws) -> tuple[Skin, Skin]:
    """템플릿(LG 복제본)의 **행 서식**에서 두 계열을 실측한다.

    찾는 것은 `노랑 채움 + 볼드` 인 행 — 그것이 `Slide title` 행이고, 바로 다음 행이
    그 계열의 맨 행이다. 글꼴 크기가 큰 쪽이 구조·요약 계열, 작은 쪽이 본문 계열이다.
    인덱스를 박지 않았으므로 템플릿이 바뀌어도 따라간다.
    """
    seen: dict[float, tuple] = {}
    for r in sorted(ws.row_dimensions):
        rd = ws.row_dimensions[r]
        if rd._style is None:
            continue
        fill, font = rd.fill, rd.font
        rgb = getattr(fill.fgColor, "rgb", None)
        if fill.patternType != "solid" or not isinstance(rgb, str):
            continue
        if not rgb.upper().endswith("FFFF00") or not font.b:
            continue
        nxt = ws.row_dimensions.get(r + 1)
        size = float(font.sz or 11)
        if size in seen or nxt is None or nxt._style is None:
            continue
        seen[size] = (font.name or "맑은 고딕", size,
                      StyleArray(rd._style), StyleArray(nxt._style))

    if not seen:                       # 템플릿에 단서가 없으면 색만이라도 맞춘다
        return (Skin("맑은 고딕", 11, None, None), Skin("맑은 고딕", 9, None, None))
    big = seen[max(seen)]
    small = seen[min(seen)]
    return Skin(*big), Skin(*small)


# ── 표 서식 (LG 본문 블록 r357~366 실측 재현) ─────────────────────────────────

def _tbl(font, *, l=None, r=None, t=None, b=None, halign=None,
         valign="top", numfmt=None):
    return {"font": font,
            "border": Border(left=l, right=r, top=t, bottom=b),
            "align": Alignment(horizontal=halign, vertical=valign),
            "numfmt": numfmt}


def _put(ws, row: int, col: int, spec: dict) -> None:
    c = ws.cell(row, col)
    c.font = spec["font"]
    c.border = spec["border"]
    c.alignment = spec["align"]
    if spec.get("numfmt"):
        c.number_format = spec["numfmt"]


def style_table(ws, skin: Skin, *, panel_row: int, hdr_row: int, base_row: int,
                last_row: int, label_col: int, first_col: int, last_col: int,
                spans: list[tuple[int, int]], vmerge: list[int],
                label_span: int = 1, mod_col: int | None = COL_VALUE,
                tail: bool = True) -> None:
    """블록 하나의 표에 LG 서식을 입힌다.

    `hdr_row` 는 배너 **그룹** 행이고 그 다음 행이 배너 세부 행이다.
    `spans` 는 그룹 헤더의 가로 병합 구간 [(시작열, 끝열), …],
    `vmerge` 는 두 행 세로 병합할 열(행 라벨 모서리 · 전체 · 순위 열)이다.

    요약 장표의 다단 패널을 위해 세 가지를 열어 뒀다. 기본값을 그대로 두면
    **본문 블록의 출력은 한 칸도 달라지지 않는다.**

        label_span  행 라벨 칸이 차지할 열 수. 2 이상이면 안쪽 세로선을 빼서 한 칸처럼
                    보이게 한다 — 오른쪽 칸을 비워 두므로 긴 라벨이 그대로 흘러 나온다
                    (LG 요약 패널도 병합 없이 이렇게 뒀다: I137·N137).
        mod_col     `그래프 상단 Text` 가 들어갈 행수식어 열. 패널마다 다르다
                    (첫 패널은 값 열 B, 그 다음 패널은 앞의 빈 열).
        tail        표 아래 여백 한 줄에 thickTop 을 남길지. 다단에서는 짧은 패널의
                    다음 줄이 옆 패널의 데이터 줄이라 켜면 남의 행 높이를 건드린다.
    """
    f, fb = skin.font, skin.font_b
    lbl_row = hdr_row + 1
    label_end = label_col + label_span - 1

    def edge(col):
        return (_T_THICK if col == first_col else _T_THIN,
                _T_THICK if col == last_col else _T_THIN)

    def label_band(row: int, *, t=None, b=None) -> None:
        """행 라벨 칸 한 줄. 바깥 세로선만 굵게 두고 안쪽 세로선은 지운다."""
        for col in range(label_col, label_end + 1):
            _put(ws, row, col,
                 _tbl(f, l=_T_THICK if col == label_col else None,
                      r=_T_THICK if col == label_end else None,
                      t=t, b=b, halign="left"))

    # ⓪ 병합을 **먼저** 한다. openpyxl 의 `merge_cells` 는 병합할 때 범위 안 셀들의
    #    테두리를 왼쪽 위 칸 기준으로 다시 쓴다 — 서식을 먼저 입히면 그때 덮인다.
    #    (배너 세부 행의 아래 굵은 선이 사라지는 증상이 이것이었다.)
    if last_col > COL_VALUE:
        ws.merge_cells(start_row=panel_row, start_column=label_col,
                       end_row=panel_row, end_column=last_col)
    for r0, r1 in spans:
        if r1 > r0:
            ws.merge_cells(start_row=hdr_row, start_column=r0,
                           end_row=hdr_row, end_column=r1)
    ws.merge_cells(start_row=hdr_row, start_column=label_col,
                   end_row=lbl_row, end_column=label_end)
    for col in vmerge:
        ws.merge_cells(start_row=hdr_row, start_column=col,
                       end_row=lbl_row, end_column=col)

    # ① 문항 제목 — 행 라벨 열부터 표 오른쪽 끝까지 병합, 볼드
    for c in range(label_col, last_col + 1):
        _put(ws, panel_row, c, _tbl(fb, halign="left", valign="center"))
    ws.row_dimensions[panel_row].height = HT_PANEL
    ws.row_dimensions[panel_row].thickBot = True

    # ② 배너 헤더 2줄 — 바깥 굵은 선 상자의 윗변
    label_band(hdr_row, t=_T_THICK)
    label_band(lbl_row, b=_T_THICK)
    for col in range(first_col, last_col + 1):
        l, r = edge(col)
        # 세로 병합된 열(전체·순위)은 두 행이 한 칸이므로 위아래 굵은 선을 함께 준다.
        merged = col in vmerge
        _put(ws, hdr_row, col,
             _tbl(f, l=l, r=r, t=_T_THICK, b=_T_THICK if merged else _T_THIN,
                  halign="center", valign="center"))
        _put(ws, lbl_row, col,
             _tbl(f, l=l, r=r, t=_T_THICK if merged else _T_THIN, b=_T_THICK,
                  halign="center", valign="center"))
    for rr, ht in ((hdr_row, HT_HEADER), (lbl_row, HT_HEADER)):
        ws.row_dimensions[rr].height = ht
    ws.row_dimensions[hdr_row].thickTop = True
    ws.row_dimensions[lbl_row].thickBot = True

    # ③ Base for % · 데이터 행 · 마지막 행
    for rr in range(base_row, last_row + 1):
        top = _T_THICK if rr == base_row else None
        bot = _T_THICK if rr == last_row else None
        if mod_col:
            _put(ws, rr, mod_col, _tbl(f, halign="left"))   # 행수식어 열 (테두리 없음)
        label_band(rr, t=top, b=bot)
        # LG 는 수치 칸에 표시 형식을 걸어 둔다 — 평균 행만 소수점 두 자리다.
        label = str(ws.cell(rr, label_col).value or "")
        numfmt = None if rr == base_row else ("###0.00" if "평균" in label else "###0.0")
        for col in range(first_col, last_col + 1):
            l, r = edge(col)
            halign = "left" if rr == base_row else "right"
            # 표시 자릿수의 정본은 팩트의 `cell_formats`(=통계표 셀 서식)다. 값 쓸 때
            # 이미 걸어 두었으므로 여기서는 **그대로 존중**한다.
            #
            # 값의 소수부 길이로 서식을 추론하지 않는다 — R-191 로 셀에 담는 값이
            # 통계표 **원값**(86.36363636363636)이 되면서 그 추론이 무너졌다.
            # 그대로 두면 `###0.00000000000000` 이 붙어 화면에 원값이 그대로 보인다.
            # 값은 원값을 지키고 화면만 1자리로 보이는 것이 이 규칙의 요구다.
            cell_fmt = numfmt
            if numfmt:
                cur = str(ws.cell(rr, col).number_format or "")
                if cur.startswith("###0"):
                    cell_fmt = cur              # 값 쓸 때 통계표 서식을 이미 걸었다
                elif isinstance(ws.cell(rr, col).value, str) and \
                        re.fullmatch(r"-?\d+", ws.cell(rr, col).value.strip()):
                    cell_fmt = "###0"           # 빈도 열은 정수라 소수점을 붙이지 않는다
            _put(ws, rr, col, _tbl(f, l=l, r=r, t=top, b=bot, halign=halign,
                                   numfmt=cell_fmt))
        ws.row_dimensions[rr].height = HT_HEADER if rr in (base_row, last_row) else HT_DATA
    ws.row_dimensions[base_row].thickTop = True
    ws.row_dimensions[last_row].thickBot = True

    # ④ 표 아래 여백 한 줄 — LG 는 여기에 thickTop 을 남긴다
    if tail:
        rd = ws.row_dimensions[last_row + 1]
        rd.height = HT_PANEL
        rd.thickTop = True


def dash_empty_cells(ws, mark: str = "-") -> int:
    """표 데이터 칸의 빈 자리를 `-` 로 메운다. 돌려주는 값은 채운 칸 수다.

    하우스가 갈리는 자리다 — SKEE 본문에는 `-` 가 수백 개이고 LG 송부본에는 0개다.
    빈 칸으로 두면 제작팀이 「값이 없음」인지 「적는 걸 잊었음」인지 구분하지 못하므로,
    사내 매크로 `Sub lovelySeoyeon()` 이 선택 영역의 빈 칸을 `-` 로 바꾼다.
    이 함수는 그 매크로를 빌드 단계에서 미리 돌려 두는 것이다.

    대상은 **표시 서식이 `###0` 으로 시작하는 칸**뿐이다 — 그 서식은 데이터 칸에만 붙으므로
    라벨·`Base for %`·머리글은 건드리지 않는다.
    """
    n = 0
    for row in ws.iter_rows():
        for cell in row:
            if not str(cell.number_format or "").startswith("###0"):
                continue
            if cell.value is None:
                cell.value = mark          # 서식은 그대로 둔다 — 오른쪽 정렬이라 `-` 도 값과 같은 자리에 선다
                n += 1
    return n


def numify_data_cells(ws) -> int:
    """데이터 칸에 문자열로 적힌 숫자를 실제 숫자로 되돌린다. 돌려주는 값은 바꾼 칸 수다.

    자릿수는 **셀 표시 서식**(`###0.0` · `###0.00` · `###0`)이 지키므로 값을 숫자로 두어도
    화면 글자는 그대로다. 문자열로 두면 두 가지가 깨진다 —
    ① 엑셀이 칸마다 「텍스트로 저장된 숫자」 경고(초록 삼각형)를 띄운다
    ② 제작팀이 표를 복사해 차트로 옮길 때 값으로 잡히지 않고 합계 검산도 막힌다.
    LG 예시 실측: 데이터 칸은 float 이다(`62.5`), 문자열은 `Base for %` 행의 `(24)` 뿐이다.

    대상은 **표시 서식이 `###0` 으로 시작하는 칸**뿐이다. 그 서식은 데이터 칸에만 붙으므로
    `(86)` · `-` · 라벨은 건드리지 않는다.
    """
    changed = 0
    for row in ws.iter_rows():
        for cell in row:
            if not str(cell.number_format or "").startswith("###0"):
                continue
            v = cell.value
            if not isinstance(v, str):
                continue
            txt = v.strip()
            if re.fullmatch(r"-?\d+\.\d+", txt):
                cell.value = float(txt)
                changed += 1
            elif re.fullmatch(r"-?\d+", txt):
                cell.value = int(txt)
                changed += 1
    return changed


def guide_rows(grid: dict) -> list[dict]:
    """제작가이드 데이터 표에 싣는 행 — 사용자 지시(2026-08-23):
    ① `계`(kind=total) 행은 싣지 않는다(표마다 있다 없다 해서 일관성이 없었다).
    ② 보기 행의 들여쓰기를 통일한다 — 통계표는 보기마다 선행 공백이 0~2칸으로 제각각(마지막 추가 보기가 1칸)이라
       그대로 옮기면 마지막 보기만 튀어 보인다. 보기 행은 공백 2칸, 그 외(【TOP2】·[5점 평균] 등)는 원문 유지.
    grid 자체는 바꾸지 않는다(slide_facts 는 정본) — 쓰는 순간에만 정리한다."""
    out = []
    for r in grid["rows"]:
        if r.get("kind") == "total" or str(r.get("label") or "").strip() in ("계", "합계"):
            continue
        if r.get("kind") == "option":
            r = {**r, "label": "  " + str(r.get("label_unit") or r["label"]).strip()}   # R-139: 단위 병기형이 있으면 그것
        out.append(r)
    return out


def _group_spans(groups: list[str], start_col: int) -> list[tuple[int, int]]:
    """배너 그룹 헤더의 가로 병합 구간. 그룹명은 첫 열에만 있고 나머지는 빈칸이다."""
    spans: list[tuple[int, int]] = []
    cur: int | None = None
    for j, g in enumerate(groups):
        col = start_col + j
        if g:
            if cur is not None:
                spans.append((cur, col - 1))
            cur = col
        elif cur is None:
            cur = col
    if cur is not None:
        spans.append((cur, start_col + len(groups) - 1))
    return spans


def write_block(ws, row: int, slide: dict, skin: Skin) -> int:
    """슬라이드 하나를 쓴다. 돌려주는 값은 다음 블록의 시작 행이다.

    빈 칸은 **비워 둔다** — `-` 로 메우지 않는다 (LG 실측). 검수는 이 블록을 팩트와
    **등장 순서**로 조인하므로 seq 열이 없어도 짝을 잃지 않는다.
    """
    grid = slide["grid"]
    overlay = grid.get("rank_overlay")

    values = [
        slide.get("slide_title") or "",
        slide.get("sub_title") or "",
        slide["chart_type"],
        slide["base_unit"] + (f"  {slide['base_note']}" if slide.get("base_note") else ""),
        # `Question/Reference` 는 문항 원문과 **장별 각주**를 함께 담는 칸이다.
        # 각주는 목차의 각주 배정표에서 오고 `build_facts --notes` 가 실어 준다.
        # 붙이지 않으면 「판단은 문서에만 적히고 셀에는 안 들어가는」 자리가 생긴다.
        "\n".join([slide.get("qu") or ""] + list(slide.get("note") or [])).strip(),
        "",
    ]
    for i, (label, value) in enumerate(zip(LABELS, values)):
        ws.cell(row + i, COL_LABEL, label)
        ws.cell(row + i, COL_VALUE, value)
        skin.label_cell(ws, row + i, i)
        skin.value_cell(ws, row + i, i)
        skin.row_label(ws, row + i, title=(i == 0))
    row += len(LABELS)

    # ── 멘트: Data 라벨 아래부터 라벨 없이 값 열에 이어 쓴다
    for line in slide.get("ment") or []:
        ws.cell(row, COL_VALUE, line).font = skin.font
        skin.row_label(ws, row)
        row += 1
    skin.row_label(ws, row)          # 멘트와 표 사이 빈 줄
    row += 1

    # ── 문항 제목 (패널 제목)
    # LG 는 이 제목을 **행 라벨 열(C)** 에 둔다(실측 C열 59건 / B열 0건).
    # B(값 열)는 `그래프 상단 Text` 신호어 자리이므로(LG 77건) 제목을 거기 넣으면
    # 제작팀이 신호어로 오독한다.
    ws.cell(row, COL_ROWLABEL, slide.get("title_full") or "")
    panel_row = row
    row += 1

    # ── 배너 헤더 2줄
    first = COL_FIRST_DATA
    if overlay:
        cols = grid.get("rank_columns") or ["1순위", "1+2+3순위"]
        ws.cell(row, first, cols[0])
        ws.cell(row, first + 1, cols[1])
        banner_start = first + 2
    else:
        banner_start = first + 1

    group_row = grid["banner_group_row"][1:]     # [0] 은 라벨 열 자리
    label_row = grid["banner_label_row"][1:]
    base_row = grid["base_row"][1:]

    # 전체 열은 그룹 헤더에서 그대로 '전체'
    hdr_row = row
    if not overlay:
        ws.cell(row, first, group_row[0])
    for j, g in enumerate(group_row[1:], start=0):
        if g:
            ws.cell(row, banner_start + j, g)
    row += 1

    # 첫 열의 세부 라벨 — 보통은 '전체'라 그룹 행과 같아 비워 두지만(LG 관례), 첫 열이
    # 전체가 아닌 그리드(브랜드 열부터 시작하는 발췌 등)에서는 이 칸을 비우면 열 이름이 사라진다.
    if not overlay and label_row and label_row[0] and label_row[0] != group_row[0]:
        ws.cell(row, first, label_row[0])
    for j, l in enumerate(label_row[1:], start=0):
        ws.cell(row, banner_start + j, l)
    row += 1

    data_r0 = row
    # ── Base for %
    ws.cell(row, COL_ROWLABEL, "Base for %")
    if overlay:
        ws.cell(row, first, _num(base_row[0]))
        ws.cell(row, first + 1, _num(base_row[0]))
    else:
        ws.cell(row, first, _num(base_row[0]))
    for j, b in enumerate(base_row[1:], start=0):
        ws.cell(row, banner_start + j, _num(b))
    row += 1

    # ── 데이터 행
    for r in guide_rows(grid):
        if r.get("modifier"):
            ws.cell(row, COL_VALUE, r["modifier"])          # '그래프 상단 Text'
        ws.cell(row, COL_ROWLABEL, r["label"])
        cells = r["cells"]
        # 값은 통계표 **원값**이고 자릿수는 표시 서식이 지킨다 (`cell_formats`).
        # 서식을 여기서 걸어 두면 뒤따르는 `_style_body_table` 이 그것을 존중한다.
        fmts = r.get("cell_formats") or []

        def put(col: int, v, i: int):
            c = ws.cell(row, col, _num(v))
            f = fmts[i] if i < len(fmts) else None
            if f:
                c.number_format = f
            return c

        if overlay and r["kind"] == "option":
            ws.cell(row, first, _num(r.get("rank_first_total")))
            put(first + 1, cells[0], 0)
        else:
            put(first, cells[0], 0)
            if overlay:
                put(first + 1, cells[0], 0)
        for j, v in enumerate(cells[1:], start=0):
            put(banner_start + j, v, j + 1)
        row += 1

    last_col = banner_start + max((len(r["cells"]) - 1) for r in grid["rows"]) - 1
    # 첫 열의 그룹·세부 헤더 세로 병합은 그 열이 '전체'처럼 이름 한 개일 때만 한다.
    # 이름이 둘(그룹 + 세부 라벨)이면 병합이 세부 라벨을 가린다.
    head_merge = (not label_row) or (not label_row[0]) or label_row[0] == group_row[0]
    style_table(ws, skin,
                panel_row=panel_row, hdr_row=hdr_row, base_row=data_r0,
                last_row=row - 1, label_col=COL_ROWLABEL,
                first_col=first, last_col=last_col,
                spans=_group_spans(group_row[1:], banner_start),
                vmerge=([first, first + 1] if overlay else ([first] if head_merge else [])))
    return row + 2                                       # 블록 사이 빈 행 2줄


#: 요약 장표의 차트 지시. 우리 요약은 미니표 없는 **순수 텍스트형**이라
#: SKEE 가 같은 경우에 쓴 `전체 : 텍스트 박스` 를 따른다.
#: (LG 요약은 미니차트를 얹어 `전체 : 원그래프` 등을 쓴다 — 성격이 다르다)
SUMMARY_CHART = "전체 : 텍스트 박스"
SUMMARY_BASE = "[Base: 전체(n=500)/ Unit: %]"


def conclusion_layout(n_items: int) -> str:
    """결론 장표의 레이아웃 지시.

    LG 는 이 칸에 산문으로 지시한다 — `제목 - 내용 형태의 박스 3개 -> 총결론 박스 하나
    구성 부탁드립니다.` (B308). 비워 두면 제작팀이 만들 수가 없다.
    항목 수에 맞춰 단 구성을 제안한다.
    """
    if n_items <= 3:
        return f"제목 - 내용 형태의 박스 {n_items}개 구성 부탁드립니다."
    cols = 3 if n_items % 3 == 0 else 2
    rows = -(-n_items // cols)
    return (f"제목 - 내용 형태의 박스 {n_items}개를 {cols}단 × {rows}행으로 "
            f"구성 부탁드립니다.")


def write_struct_block(ws, row: int, skin: Skin, *, slide_title: str = "",
                       sub_title: str = "", chart_type: str = "",
                       base_unit: str = "", qu: str = "",
                       rows: list[tuple] | None = None,
                       table: bool = False) -> int:
    """구조 블록(표지·목차·간지·조사개요·요약·결론) 하나를 쓴다.

    라벨 6줄은 본문 블록과 **똑같이** 두고 `Chart type` 칸에 구분어만 적는다 (R-035).
    나머지 칸은 LG 예시 그대로 **비워 둔다** — `-` 로 메우지 않는다.
    검수 V7(미기입 잔존)은 구조 블록을 제외 대상으로 안다.

    `rows` 는 `(값 열 B, 행 라벨 열 C, 전체 열 D …)` 튜플 목록이다.
    행 라벨 열이 비어도 검수 파서는 멘트로 오독하지 않는다 — 멘트는 `Data` 라벨
    **바로 아래 연속 구간**이고, 구조 블록은 그 자리에 빈 행을 한 줄 두기 때문이다.
    """
    values = [slide_title, sub_title, chart_type, base_unit, qu, ""]
    for i, (label, value) in enumerate(zip(LABELS, values)):
        ws.cell(row + i, COL_LABEL, label)
        if value not in (None, ""):
            ws.cell(row + i, COL_VALUE, value)
        skin.label_cell(ws, row + i, i)
        skin.value_cell(ws, row + i, i)
        skin.row_label(ws, row + i, title=(i == 0))
    row += len(LABELS)
    skin.row_label(ws, row)
    row += 1                                   # Data 라벨 다음 한 줄 비움 (LG 실측)

    body_r0 = row
    width = 2
    for cells in rows or []:
        cells = list(cells)
        width = max(width, len(cells))
        for j, v in enumerate(cells):
            c = ws.cell(row, COL_VALUE + j)
            if v not in (None, ""):
                c.value = v
            c.font = skin.font
            c.alignment = Alignment(vertical="center", wrap_text=(j == 0))
        skin.row_label(ws, row)
        row += 1

    # 응답자 특성처럼 실제 통계표가 들어가는 구조 블록은 본문과 같은 상자를 두른다.
    if table and rows and width >= 3:
        _style_struct_table(ws, skin, body_r0, row - 1, COL_VALUE + width - 1)
    return row + 2


def _style_struct_table(ws, skin: Skin, r0: int, r1: int, last_col: int) -> None:
    """구조 블록 안의 표 (LG `3. 응답자 특성` r76~r112 실측).

    본문 표와 같은 상자다 — 첫 줄이 `Base for %` 이므로 헤더 두 줄 없이
    바깥 굵은 선만 두르고, 구분명 열(B)·항목 열(C)·수치 열(D~)로 나눈다.
    LG 실측(r76~r112, s=213/214/215) 그대로: 구분 열 왼쪽 굵은 선, 항목 열 오른쪽
    굵은 선으로 라벨 구역을 닫고, 수치 열은 왼쪽 굵은 선으로 다시 연다.
    """
    f = skin.font
    for rr in range(r0, r1 + 1):
        top = _T_THICK if rr == r0 else None
        bot = _T_THICK if rr == r1 else None
        _put(ws, rr, COL_VALUE, _tbl(f, l=_T_THICK, t=top, b=bot, halign="center"))
        _put(ws, rr, COL_ROWLABEL, _tbl(f, r=_T_THICK, t=top, b=bot, halign="left"))
        for col in range(COL_FIRST_DATA, last_col + 1):
            # 수치 칸에는 본문 표와 같은 표시 서식을 건다 — LG 실측 r79~r112 가 `###0.0` 이다.
            # 서식이 없으면 뒤따르는 `numify_data_cells` 가 이 칸을 데이터로 알아보지 못해
            # 값이 문자열로 남고, 엑셀이 「텍스트로 저장된 숫자」로 잡는다.
            cv = ws.cell(rr, col).value
            numfmt = "###0.0" if isinstance(cv, str) and re.fullmatch(r"-?\d+(\.\d+)?", cv.strip()) else None
            _put(ws, rr, col,
                 _tbl(f, l=_T_THICK if col == COL_FIRST_DATA else _T_THIN,
                      r=_T_THICK if col == last_col else _T_THIN,
                      t=top, b=bot, halign="right", numfmt=numfmt))
        ws.row_dimensions[rr].height = HT_LABEL
    ws.row_dimensions[r0].thickTop = True
    ws.row_dimensions[r1].thickBot = True


# ── 요약 장표의 다단 패널 (LG `본문_제작용` r128~170 실측 재현) ──────────────────
#
#  요약 장표는 순수 텍스트가 아니다. 본문에서 표를 2~4개 뽑아 한 페이지에 늘어놓고
#  **그 표들을 요약한 멘트**를 다는 것이다 (녹취 41:13). LG 요약 「1. 홈로봇 기존 인식」의
#  실측 배치가 그대로 규격이다:
#
#      A130 Chart type   B130=전체 : 원그래프   I130=전체 : 가로 겹침 막대   N130=전체 : 가로 막대
#      A131 Base/Unit    B131=[Base: 전체(n=100)/ Unit: %]   I131=(같은 형식)   N131=(같은 형식)
#      A133 Data         B134·B135=멘트
#      B137=패널1 제목   I137=패널2 제목   N137=패널3 제목
#      B139=패널1 문항   I139=패널2 문항   N139=패널3 문항
#      그 아래로 각 패널의 표가 같은 열 오프셋에 나란히
#
#  즉 `Chart type` · `Base/Unit` · 제목 · 문항 · 표가 **가로로 반복**되고 패널마다
#  열 오프셋이 다르다. LG 는 한 장표 안에서 단을 **세로로도 한 번 더** 쌓았다
#  (r153~170: 라벨 열 없이 Chart type · Base/Unit 부터 다시 시작한다).
#
#  열은 우리 규약을 따른다 — 첫 패널의 행 라벨은 C(폭 44), 데이터는 D~.
#  두 번째부터는 본문 배너 열(폭 11) 위에 놓이므로 라벨 칸을 **4열 폭**으로 잡아
#  44 를 맞춘다. 병합하지 않고 오른쪽 세 칸을 비워 두면 긴 라벨이 그대로 흘러 나온다
#  (LG 도 I137·N137 을 병합 없이 그렇게 뒀다).

#: 두 번째 이후 패널의 행 라벨 폭. 11 × 4 = 44 로 C열과 같아진다.
SUMMARY_LABEL_SPAN = 4
#: 패널 사이 빈 열 한 칸. 이 열이 다음 패널의 행수식어(`그래프 상단 Text`) 열이 된다.
SUMMARY_GAP = 1
#: 가로로 밀어 넣을 수 있는 오른쪽 한계. 본문 블록이 쓰는 폭과 같다.
SUMMARY_MAX_COL = 22


def panel_chart_type(chart: str) -> str:
    """목차.md 표기 → 가이드 표기. `전체 가로막대` → `전체 : 가로막대`.

    배너 **차트** 패널(`배너 가로 누적 100% 막대 (관리 운영 방식)`)은 표를 전치해 싣는다(행 = 집단) —
    제작팀이 행 방향으로 막대를 쌓아도 각 행이 100% 가 되도록 표기에 못 박는다 (R-145, 홈로봇 실측 r153).
    """
    scope, _, rest = chart.partition(" ")
    if scope == "배너" and rest and not rest.startswith("표"):
        m = re.search(r"\(([^)]*)\)", rest)
        groups = m.group(1).strip() if m else "배너"
        body = rest[:m.start()].strip() if m else rest.strip()
        if not body.endswith("그래프"):
            body += " 그래프"
        # 합이 100 이 되는 것은 **누적 100%** 뿐이다. 묶음 막대(중복응답·TOP2 요약·행 선별)는 합이 100 이 아니므로
        # 「각 행 Base 100%」라고 적으면 제작팀이 정규화해 정본과 어긋난다 (반증관 실측 2026-08-24:
        # 표74 행 합 171.5 · 표81 선별 3행 합 124.8 · 표9001 TOP2 3지표 합 198.2).
        tail = ("각 행 Base 100%" if "누적" in body
                else "각 행 Base = 그 집단 n · 행 합 100% 아님")
        return f"배너 : {body} (행 = {groups} 집단, {tail})"
    return f"{scope} : {rest}".strip() if rest else scope


def is_transposed(chart: str) -> bool:
    """배너 패널인데 `표` 가 아니면(막대 등) 전치 교차표 — 행 = 전체 + 배너 집단, 열 = 보기."""
    scope, _, rest = chart.partition(" ")
    return scope == "배너" and bool(rest) and not rest.lstrip().startswith("표")


#: 패널 배치 지시의 **행 선별** 꼬리 — `전체 가로막대 · 상위 3` / `… · 행 'A'|'B'`.
#: 요약 장표는 상위·격차 포인트만 싣는다(사용자 지시 2026-08-23). 구간형·척도는 꼬리 없이 전 행.
RE_PANEL_SELECT = re.compile(r"\s*·\s*(?:상위\s*(?P<top>\d+)|행\s*(?P<rows>'[^']*'(?:\s*\|\s*'[^']*')*))\s*$")


def split_panel_chart(chart: str) -> tuple[str, dict]:
    """`전체 가로막대 · 상위 3` → (`전체 가로막대`, {"top": 3}) / `… · 행 'A'|'B'` → (…, {"rows": [A, B]})."""
    sel: dict = {}
    m = RE_PANEL_SELECT.search(chart)
    while m:
        if m.group("top"):
            sel["top"] = int(m.group("top"))
        else:
            sel["rows"] = [x.strip().strip("'") for x in m.group("rows").split("|")]
        chart = chart[:m.start()].rstrip()
        m = RE_PANEL_SELECT.search(chart)
    return chart, sel



def select_panel_columns(grid: dict, groups: list[str]) -> dict:
    """배너 표 패널의 **열 선별** — `배너 표 (관리 운영 방식)` 처럼 괄호 안에 적은 배너군만 남긴다.

    cells[0](전체)은 항상 남긴다. 그리드의 banner_group_row 는 군 이름을 첫 열에만 적고 나머지는
    빈 문자열이므로 앞으로 채워(forward-fill) 군을 판정한다. 값은 옮기기만 한다.
    R-070(요약은 필요한 부분만 잘라 온다) · R-129(절단 사유는 적지 않는다).
    """
    want = {g.strip() for g in groups if g.strip()}
    grow = grid.get("banner_group_row") or []
    if not want or len(grow) < 2:
        return grid
    # banner_group_row[0] 은 라벨 열 자리, [1] 부터 데이터 열. cells 는 [0]=전체, [1..]=배너
    filled, cur = [], ""
    for g in grow[1:]:
        cur = g if g else cur
        filled.append(cur)
    keep = [0] + [i for i, g in enumerate(filled) if g in want and i > 0]
    if len(keep) == len(filled):
        return grid

    def pick(seq):
        return [seq[i] for i in keep if i < len(seq)]
    out = {**grid, "_columns": keep}
    out["banner_group_row"] = [grow[0]] + pick(grow[1:])
    # 군 이름은 그 군의 첫 열에만 — 선별 후 첫 열이 빈 문자열이면 군 이름을 되살린다
    out["banner_group_row"] = [grow[0]] + [
        filled[i] if (j == 0 or filled[i] != filled[keep[j - 1]]) else "" for j, i in enumerate(keep)]
    out["banner_group_row"][1] = grow[1]     # 전체 열 머리글은 원본 그대로 (`전체`)
    lrow = grid.get("banner_label_row") or []
    out["banner_label_row"] = [lrow[0]] + pick(lrow[1:]) if lrow else lrow
    brow = grid.get("base_row") or []
    out["base_row"] = [brow[0]] + pick(brow[1:]) if brow else brow
    out["rows"] = [{**r, "cells": pick(r.get("cells") or [])} for r in grid.get("rows") or []]
    return out

def select_panel_rows(grid: dict, sel: dict, warn) -> dict:
    """그리드의 보기 행을 선별한다. 값은 옮기기만 하고 새로 계산하지 않는다.

    `top` — 전체 값(순위면 1~N순위 합 = cells[0]) 내림차순 상위 N 보기 행. 보기 외 행(계·【TOP】·평균)은
    요약 패널의 목적(상위 포인트)과 무관하므로 함께 떨군다 — 단 척도 요약행은 차트의 근거라 유지한다.
    `rows` — 라벨 원문 일치 행만 (격차 포인트: 최고 vs 최저).
    """
    if not sel:
        return grid
    rows = grid.get("rows") or []
    opts = [r for r in rows if r.get("kind") == "option"]
    others = [r for r in rows if r.get("kind") != "option"]

    def total(r):
        try:
            return float(str((r.get("cells") or ["-"])[0]).replace("*", ""))
        except ValueError:
            return float("-inf")

    if "rows" in sel:
        want = [w.strip() for w in sel["rows"]]
        # 지정한 **순서**가 곧 패널의 행 순서다 — 퍼널(85→66→64→47)처럼 순서 자체가 논지다 (ADV-04)
        # 명시 라벨은 보기 행만이 아니라 요약 행(【TOP2】 등)도 고를 수 있다 — 용도·권역
        # 의향 패널처럼 TOP2 한 행만 싣는 지시가 있다 (r14, 반증 F-16 반영).
        picked = []
        for w in want:
            hit = next((r for r in rows if (r.get("label") or "").strip() == w or (r.get("label_clean") or "").strip() == w), None)
            if hit is not None:
                picked.append(hit)
        missing = [w for w in want if not any((r.get("label") or "").strip() == w or (r.get("label_clean") or "").strip() == w for r in rows)]
        if missing:
            warn(f"요약 패널 행 선별 — 라벨을 찾지 못했습니다: {missing}")
    else:
        picked = sorted(opts, key=total, reverse=True)[: sel["top"]]
    # 명시 목록에 요약 행이 섞였으면 「보기 + 자동 요약행」 규약을 끈다 — 지정한 것만 싣는다
    if any(r.get("kind") != "option" for r in picked):
        return {**grid, "rows": picked, "_selected": sel}
    keep_others = [r for r in others if re.search(r"【TOP|【BOT|평균", r.get("label") or "")]
    # 짝 Mean 표에서 붙인 `[5점 평균] 항목` 행은 선별된 보기와 짝이 맞는 것만 남긴다 (V17 실측 2026-08-24: 선별 밖 3항목 평균이 따라붙음)
    picked_labels = {str(r.get("label_clean") or r.get("label") or "").strip() for r in picked}
    keep_others = [r for r in keep_others
                   if not str(r.get("label") or "").startswith("[5점 평균] ")
                   or str(r.get("label") or "")[len("[5점 평균] "):].strip() in picked_labels]
    # 짝 평균 행은 **막대 행과 같은 순서**로 세운다 (반증 F1, 2026-08-28) — Mean 표 원본
    # 순서를 그대로 두면 선별 순서와 어긋나 평균이 다른 막대 위에 붙는다.
    order = {lab: i for i, lab in enumerate(str(r.get("label_clean") or r.get("label") or "").strip()
                                            for r in picked)}
    keep_others.sort(key=lambda r: order.get(
        str(r.get("label") or "")[len("[5점 평균] "):].strip()
        if str(r.get("label") or "").startswith("[5점 평균] ") else "", 10 ** 6))
    return {**grid, "rows": picked + keep_others, "_selected": sel}


def _mean_unit_of(mean_rows: list[dict]) -> str:
    """평균 행의 단위 — `[평균 : 만원]` 라벨에서 뽑는다. `[5점 평균] …` 짝 행이면 「점」."""
    for r in mean_rows:
        m = re.search(r"\[평균\s*[:：]\s*([^\]]+)\]", str(r.get("label") or ""))
        if m:
            return m.group(1).strip()
        if r.get("mean_unit"):
            return str(r["mean_unit"]).strip()
    return "점"


def panel_source(spec: dict, by_page: dict[int, dict],
                 by_index: dict[int, dict], warn) -> dict | None:
    """배치 지시 한 줄 → 그릴 수 있는 패널. **수치는 옮기기만 한다.**

    `P{n}` 은 본문 슬라이드의 페이지 번호이므로 팩트의 `grid` 를 그대로 쓰고,
    `표{n}` 은 본문에 없는 Summary 표라 본문 빌더와 **같은 함수**(`build_facts.build_grid`)
    로 그리드를 만든다. 둘 다 통계표에서 온 값 그대로다.
    """
    ref = spec["ref"]
    chart_txt, sel = split_panel_chart(spec["chart"])
    spec = {**spec, "chart": chart_txt}
    banner = spec["chart"].startswith("배너")
    if ref.startswith("P"):
        s = by_page.get(int(ref[1:]))
        if s is None:
            warn(f"요약 패널 {ref} — `page={ref[1:]}` 인 본문 슬라이드가 없습니다")
            return None
        grid = s["grid"]
        # 요약 패널의 문항 줄은 **`{문항코드}. {원문}`** 이다 — LG 실측 r139
        # (`A1. 집에서 로봇은 사람과 어떤 관계여야 한다고 생각하십니까?`).
        # 본문 블록의 `Qu: 원문` 과 표기가 다르다: 요약에서는 띠(주제명)가 코드를 안 달기
        # 때문에 코드가 여기 아니면 장표 어디에도 남지 않는다.
        question = s.get("qu") or s.get("title_full") or ""
        code = (s.get("question_code") or "").strip()
        if code and question.startswith("Qu: "):
            question = f"{code}. {question[4:]}"
        base_unit = s["base_unit"]
        base_note = s.get("base_note")
        base_n = s.get("base_n")
        extra = None
    else:
        t = by_index.get(int(ref[1:]))
        if t is None:
            warn(f"요약 패널 {ref} — 통계표 인덱스 {ref[1:]} 를 찾지 못했습니다")
            return None
        import build_facts as BF                      # 본문 그리드와 같은 조립기
        m_extra = re.search(r"^배너\s+[^(]*\(([^)]*)\)", spec.get("chart") or "")
        extra = {g.strip() for g in m_extra.group(1).split("·")} if m_extra else None
        grid = BF.build_grid(t, extra_groups=extra, small_base_mark=False)   # 절단선 밖 군도 패널이 지정하면 원표에서. 별표는 R-188 로 뺀다
        question = t.get("title_full") or ""
        base_n = t.get("base_n")
        base_unit = f"[Base: 전체(n={base_n})/ Unit: %]"
        base_note = "*Too small base(n<30)"
    if base_n is not None and base_n != spec["base_n"]:
        warn(f"요약 패널 {ref} — 목차.md 는 Base {spec['base_n']} 인데 "
             f"원자료는 {base_n} 입니다 (원자료를 씁니다)")

    if "5점평균" in (spec.get("chart") or "").replace(" ", "") and not any(r.get("kind") == "mean" for r in grid["rows"]):
        # 차트가 「top2 세로막대 + 5점평균 상단 표」인데 TOP2 Summary 표에는 평균 행이 없다 — 짝 Mean Summary 표에서 붙인다
        # (비판 평가 2026-08-23: 요약 4 중·우 패널에 [5점 평균] 행 부재)
        idx = int(ref[1:]) if ref.startswith("표") else None
        mate = by_index.get(idx + 1) if idx is not None else None
        if mate is not None and re.search(r"Mean|평균", str(mate.get("title_full") or "")):
            import build_facts as BF
            mg = BF.build_grid(mate, extra_groups=extra if ref.startswith("표") else None, small_base_mark=False)
            for r in mg["rows"]:
                if r.get("kind") == "option":
                    grid["rows"].append({**r, "label": "[5점 평균] " + str(r.get("label_clean") or r.get("label")).strip(),
                                         "label_clean": "[5점 평균] " + str(r.get("label_clean") or r.get("label")).strip(),
                                         "kind": "mean", "modifier": "그래프 상단 Text"})
        else:
            warn(f"요약 패널 {ref} — 차트가 5점평균을 요구하지만 짝 Mean 표를 찾지 못했습니다")
    if sel:
        grid = select_panel_rows(grid, sel, warn)
    if banner:
        m_groups = re.search(r"\(([^)]*)\)", spec["chart"])
        if m_groups:
            grid = select_panel_columns(grid, re.split(r"[·,/]", m_groups.group(1)))
    # 범례는 **이 패널에 실리는 열** 기준으로 — 선별로 저베이스 열이 빠졌으면 범례도 뺀다 (ADV-06).
    # 비배너 패널은 전체 열 하나만 실리므로 전체 n 만 본다 (비판 평가 2026-08-23: n=500 패널에 별표 주석이 붙고
    # 정작 n<30 열이 있는 배너 패널에는 없던 반전을 고침). 슬라이드가 물려준 base_unit 의 주석도 떼고 다시 판정한다.
    base_unit = re.sub(r"\s*\*Too small base\(n<30\)\s*$", "", str(base_unit))
    shown = (grid.get("base_row") or [])[1:]
    if not banner:
        shown = shown[:1]
    low = any("*" in str(c) for c in shown if c is not None)
    if low:
        base_unit = f"{base_unit}  {base_note or '*Too small base(n<30)'}"
    overlay = bool(grid.get("rank_overlay"))
    transpose = banner and is_transposed(spec["chart"])
    mean_rows = [r for r in grid.get("rows") or [] if r.get("kind") == "mean"]
    if mean_rows and "Unit: %]" in base_unit:
        base_unit = base_unit.replace("Unit: %]", f"Unit: %, {_mean_unit_of(mean_rows)}]")
    n_total = 2 if overlay else 1                     # 전체 열 (순위면 두 칸)
    if transpose:
        n_data = 1 + len(guide_rows(grid))            # Base for % 열 + 보기 열
    else:
        n_data = n_total + (max(len(r["cells"]) for r in grid["rows"]) - 1 if banner else 0)
    ct = panel_chart_type(spec["chart"])
    if transpose and mean_rows and "누적" in ct:
        # 누적 100% 는 구간 열끼리의 이야기다 — 평균 열이 합산에 끼면 행 합이 100 을 넘는다 (반증 F4)
        ct += " · [평균] 열은 막대에서 제외하고 우측 텍스트로"
    return {**spec, "grid": grid, "banner": banner, "overlay": overlay, "transpose": transpose,
            "n_data": n_data, "question": question, "base_unit": base_unit,
            "chart_type": ct}


def write_panel(ws, skin: Skin, *, top: int, label_col: int, label_span: int,
                panel: dict) -> tuple[int, int]:
    """패널 하나. 제목 → 빈 줄 → 문항 → 배너 헤더 2줄 → Base for % → 데이터.

    `top` 이 제목 행이다. 돌려주는 값은 (마지막 데이터 행, 마지막 열).
    서식은 본문 표와 **같은 함수**(`style_table`)를 쓴다.
    """
    if panel.get("transpose"):
        return write_panel_transposed(ws, skin, top=top, label_col=label_col,
                                      label_span=label_span, panel=panel)
    grid, banner, overlay = panel["grid"], panel["banner"], panel["overlay"]
    label_end = label_col + label_span - 1
    first = label_end + 1
    mod_col = label_col - 1

    ws.cell(top, label_col, panel["title"]).font = skin.font_b
    skin.row_label(ws, top)
    skin.row_label(ws, top + 1)                       # 제목과 문항 사이 빈 줄 (LG 138)

    panel_row = top + 2
    ws.cell(panel_row, label_col, panel["question"])

    group_row = grid["banner_group_row"][1:]          # [0] 은 라벨 열 자리
    label_row = grid["banner_label_row"][1:]
    base_vals = grid["base_row"][1:]

    hdr = panel_row + 1
    if overlay:
        cols = grid.get("rank_columns") or ["1순위", "1+2+3순위"]
        ws.cell(hdr, first, cols[0])
        ws.cell(hdr, first + 1, cols[1])
        banner_start, vmerge = first + 2, [first, first + 1]
    else:
        ws.cell(hdr, first, group_row[0])
        banner_start, vmerge = first + 1, [first]

    spans: list[tuple[int, int]] = []
    if banner:
        for j, g in enumerate(group_row[1:]):
            if g:
                ws.cell(hdr, banner_start + j, g)
        for j, l in enumerate(label_row[1:]):
            ws.cell(hdr + 1, banner_start + j, l)
        spans = _group_spans(group_row[1:], banner_start)

    base_r = hdr + 2
    ws.cell(base_r, label_col, "Base for %")
    ws.cell(base_r, first, _num(base_vals[0]))
    if overlay:
        ws.cell(base_r, first + 1, _num(base_vals[0]))
    if banner:
        for j, b in enumerate(base_vals[1:]):
            ws.cell(base_r, banner_start + j, _num(b))

    row = base_r + 1
    for r in guide_rows(grid):
        if r.get("modifier"):
            ws.cell(row, mod_col, r["modifier"])      # '그래프 상단 Text'
        ws.cell(row, label_col, r["label"])
        cells = r["cells"]
        if overlay and r["kind"] == "option":
            ws.cell(row, first, _num(r.get("rank_first_total")))
            ws.cell(row, first + 1, _num(cells[0]))
        else:
            ws.cell(row, first, _num(cells[0]))
            if overlay:
                ws.cell(row, first + 1, _num(cells[0]))
        if banner:
            for j, v in enumerate(cells[1:]):
                ws.cell(row, banner_start + j, _num(v))
        row += 1

    last_col = (banner_start + max(len(r["cells"]) for r in grid["rows"]) - 2
                if banner else first + (1 if overlay else 0))
    style_table(ws, skin, panel_row=panel_row, hdr_row=hdr, base_row=base_r,
                last_row=row - 1, label_col=label_col, first_col=first,
                last_col=last_col, spans=spans, vmerge=vmerge,
                label_span=label_span, mod_col=mod_col, tail=False)
    return row - 1, last_col


def write_panel_transposed(ws, skin: Skin, *, top: int, label_col: int, label_span: int,
                           panel: dict) -> tuple[int, int]:
    """전치 교차표 패널 (R-145, 홈로봇 실측 r157~ 「실수 유형별 허용 가능 빈도」 형식).

        (라벨)        Base for %   보기1   보기2 …      ← 헤더 1줄
        전체          (500)        31.8    21.2 …
        [배너군]                                       ← 군 이름 행 (값 없음)
        집단 A        (148)        …
    각 행이 한 집단이므로 가로 누적 100% 막대를 행 방향으로 쌓아도 합이 100 이다.
    값은 그리드를 **전치해 옮기기만** 한다 — 열 = guide_rows(grid) 순서.
    """
    f, fb = skin.font, skin.font_b
    grid = panel["grid"]
    label_end = label_col + label_span - 1
    first = label_end + 1                            # Base for % 열
    mod_col = label_col - 1
    rows = guide_rows(grid)
    cols = [str(r["label"]).strip() for r in rows]
    group_row = grid["banner_group_row"][1:]
    label_row = grid["banner_label_row"][1:]
    base_vals = grid["base_row"][1:]
    last_col = first + len(cols)

    ws.cell(top, label_col, panel["title"]).font = fb
    skin.row_label(ws, top)
    skin.row_label(ws, top + 1)
    panel_row = top + 2
    ws.cell(panel_row, label_col, panel["question"])
    hdr = panel_row + 1
    ws.cell(hdr, first, "Base for %")
    for j, c in enumerate(cols):
        ws.cell(hdr, first + 1 + j, c)

    # 데이터 행: 전체 → 군 이름 행 → 집단 행들
    row = hdr + 1
    lines = []                                       # (라벨, base, 값들, 군이름행?)
    lines.append(("전체", base_vals[0], [r["cells"][0] for r in rows], False))
    cur = ""
    for i in range(1, len(label_row)):
        g = group_row[i] if i < len(group_row) else ""
        if g and g != cur:
            cur = g
            lines.append((f"[{g}]", None, [], True))
        lines.append((str(label_row[i]).strip(), base_vals[i] if i < len(base_vals) else None,
                      [(r["cells"][i] if i < len(r["cells"]) else None) for r in rows], False))
    first_data = row
    for lab, base, vals, is_group in lines:
        ws.cell(row, label_col, lab)
        if not is_group:
            ws.cell(row, first, _num(base))
            for j, v in enumerate(vals):
                ws.cell(row, first + 1 + j, _num(v))
        row += 1
    last_row = row - 1

    # ── 서식 (style_table 은 2줄 배너 헤더 + Base 행 전제라 여기서 직접 입힌다)
    def edge(col):
        return (_T_THICK if col == first else _T_THIN, _T_THICK if col == last_col else _T_THIN)
    if last_col > COL_VALUE:
        ws.merge_cells(start_row=panel_row, start_column=label_col, end_row=panel_row, end_column=last_col)
    for c in range(label_col, last_col + 1):
        _put(ws, panel_row, c, _tbl(fb, halign="left", valign="center"))
    ws.row_dimensions[panel_row].height = HT_PANEL
    ws.row_dimensions[panel_row].thickBot = True
    for c in range(label_col, label_end + 1):
        _put(ws, hdr, c, _tbl(f, l=_T_THICK if c == label_col else None, r=_T_THICK if c == label_end else None,
                              t=_T_THICK, b=_T_THICK, halign="left"))
    for c in range(first, last_col + 1):
        l, r = edge(c)
        _put(ws, hdr, c, _tbl(f, l=l, r=r, t=_T_THICK, b=_T_THICK, halign="center", valign="center"))
    ws.row_dimensions[hdr].height = HT_HEADER
    ws.row_dimensions[hdr].thickTop = True
    ws.row_dimensions[hdr].thickBot = True
    for rr in range(first_data, last_row + 1):
        bot = _T_THICK if rr == last_row else None
        is_group = str(ws.cell(rr, label_col).value or "").startswith("[")
        if mod_col:
            _put(ws, rr, mod_col, _tbl(f, halign="left"))
        for c in range(label_col, label_end + 1):
            _put(ws, rr, c, _tbl(fb if is_group else f, l=_T_THICK if c == label_col else None,
                                 r=_T_THICK if c == label_end else None, b=bot, halign="left"))
        for c in range(first, last_col + 1):
            l, r = edge(c)
            if c == first:
                numfmt = None
            else:
                numfmt = "###0.00" if "평균" in cols[c - first - 1] else "###0.0"
            _put(ws, rr, c, _tbl(f, l=l, r=r, b=bot, halign="left" if c == first else "right", numfmt=numfmt))
        ws.row_dimensions[rr].height = HT_HEADER if rr == last_row else HT_DATA
    ws.row_dimensions[last_row].thickBot = True
    return last_row, last_col


def plan_summary(panels: list[dict]) -> tuple[list[tuple], list[tuple], list[tuple], bool]:
    """패널 목록 → (전폭 상단, 가로 배치, 세로로 쌓을 것, 세로쌓기 여부).

    가로 배치는 `[(패널, 라벨 열, 라벨 폭)]` 이다. `우하` 는 `우` 아래에 붙인다.

    `상` 은 **전폭 상단**이다 — 멘트가 첫 줄부터 짚는 표가 배너 교차라 폭이 넓을 때
    쓴다(R-197). 자리 이름만 `좌` 로 바꾸면 18열짜리 표가 1/3칸에 눌리므로, 층을
    하나 더 두어 넓은 표를 위로 올린다.

    ⚠ 배너 표 패널은 폭이 18열이라 둘을 나란히 놓을 수 없다. 억지로 밀어 넣지 않고
    **세로로 쌓는다** — 그때는 모든 패널이 전체 폭(C열 기준)을 쓴다.
    """
    full_above = [p for p in panels if p["position"] == "상"]      # 전폭 상단 — 좌 열에 앵커
    across = [p for p in panels if p["position"] in ("좌", "중", "우")]
    below = [p for p in panels if p["position"] == "우하"]
    full_below = [p for p in panels if p["position"] == "하"]      # 전폭 하단 — 좌 열에 앵커

    def lay(seq: list[dict]) -> list[tuple]:
        out, col = [], COL_ROWLABEL
        for p in seq:
            span = 1 if col == COL_ROWLABEL else SUMMARY_LABEL_SPAN
            out.append((p, col, span))
            col += span + p["n_data"] + SUMMARY_GAP
        return out

    above = [(p, COL_ROWLABEL, 1) for p in full_above]
    placed = lay(across)
    if placed and max(c + s + p["n_data"] - 1 for p, c, s in placed) > SUMMARY_MAX_COL:
        return above, lay(across[:1]), [(p, COL_ROWLABEL, 1) for p in across[1:] + below], True
    anchor = placed[-1] if placed else (None, COL_ROWLABEL, 1)
    return above, placed, ([(p, anchor[1], anchor[2]) for p in below]
                           + [(p, COL_ROWLABEL, 1) for p in full_below]), False


def summary_subtitles(toc_path: Path | None) -> dict[int, str]:
    """summary_map.json 의 subtitle(꼭지가 아우르는 문항 묶음 — guide/14 §6) — 요약 블록 Sub title 칸 (비판 평가 2026-08-23)."""
    try:
        sm = json.loads((Path(toc_path).parent / "summary_map.json").read_text(encoding="utf-8"))
        return {int(x["no"]): str(x.get("subtitle") or "") for x in sm.get("summaries", [])}
    except Exception:                                # noqa: BLE001
        return {}


def write_summary_block(ws, row: int, skin: Skin, *, slide_title: str, qu: str,
                        ment: list[str], panels: list[dict], sub_title: str = "") -> int:
    """요약 꼭지 하나 — 라벨 6줄 + 멘트 + 다단 패널. 돌려주는 값은 다음 블록 시작 행."""
    above, across, stacked, _ = plan_summary(panels)

    def head_col(col: int) -> int:
        """단 머리(Chart type · Base/Unit)를 적을 열 = 패널의 라벨 열 그대로.

        이전엔 첫 패널만 값 열 B로 보냈는데(LG 실측 B130 모방), LG 는 첫 패널 표
        자체가 B열이라 정렬됐던 것이다. 우리 규약은 첫 패널 표가 C열이므로 B에
        적으면 머리만 한 칸 왼쪽으로 어긋난다 (PPT 팀 반송 R-PPT: 요약 1~4 재현).
        """
        return col

    # ── 라벨 6줄. Chart type · Base/Unit 은 패널마다 제 열에 반복한다 (LG I130/N130)
    values = [slide_title, sub_title, "", "", qu, ""]
    for i, (label, value) in enumerate(zip(LABELS, values)):
        ws.cell(row + i, COL_LABEL, label)
        if value:
            ws.cell(row + i, COL_VALUE, value)
        skin.label_cell(ws, row + i, i)
        skin.value_cell(ws, row + i, i)
        skin.row_label(ws, row + i, title=(i == 0))
    for p, col, _span in across:
        c = head_col(col)
        ws.cell(row + 2, c, p["chart_type"]).font = skin.font
        ws.cell(row + 3, c, p["base_unit"]).font = skin.font
    row += len(LABELS)

    # ── 멘트: Data 라벨 아래부터 라벨 없이 값 열에 이어 쓴다 (본문 블록과 같다)
    for line in ment:
        ws.cell(row, COL_VALUE, line).font = skin.font
        skin.row_label(ws, row)
        row += 1
    skin.row_label(ws, row)
    row += 1                                          # 멘트와 표 사이 빈 줄

    # ── 전폭 상단(상) — 멘트가 첫 줄부터 짚는 넓은 표를 가로 단보다 위에 둔다 (R-197).
    #    쓰는 방식은 세로 단과 같다: 라벨 열 없이 Chart type · Base/Unit 부터 시작한다.
    for p, col, span in above:
        c = head_col(col)
        ws.cell(row + 1, c, p["chart_type"]).font = skin.font
        ws.cell(row + 2, c, p["base_unit"]).font = skin.font
        top = row + 4                                 # chart · base · 빈 줄
        for rr in range(row, top):
            skin.row_label(ws, rr)
        p_end, _ = write_panel(ws, skin, top=top, label_col=col, label_span=span, panel=p)
        row = p_end + 2                               # 패널 사이 빈 줄
        skin.row_label(ws, row - 1)

    # ── 가로 단
    end = row
    right_end = row
    for p, col, span in across:
        p_end, _ = write_panel(ws, skin, top=row, label_col=col, label_span=span,
                               panel=p)
        end = max(end, p_end)
        right_end = p_end                             # 마지막(=우) 패널의 끝

    # ── 세로로 쌓는 단. 라벨 열 없이 Chart type · Base/Unit 부터 다시 시작한다
    #    (LG r153~157). `우하` 는 `우` 바로 아래, 세로쌓기면 앞 단 아래에 붙는다.
    anchor = right_end
    for p, col, span in stacked:
        c = head_col(col)
        if col == COL_ROWLABEL:                       # 전폭(하) 패널은 가로 단 **전체**의 끝 아래에 — 좌 패널이 더 길면 우 끝 기준으로는 겹친다
            anchor = max(anchor, end)                 # (비판 평가 2026-08-23: 요약 3 좌 5행 라벨이 하 패널 머리에 덮임)
        top = anchor + 5                              # 빈 줄 · chart · base · 빈 줄
        ws.cell(anchor + 2, c, p["chart_type"]).font = skin.font
        ws.cell(anchor + 3, c, p["base_unit"]).font = skin.font
        for rr in range(anchor + 1, top):
            skin.row_label(ws, rr)
        p_end, _ = write_panel(ws, skin, top=top, label_col=col, label_span=span,
                               panel=p)
        anchor = p_end
        end = max(end, p_end)

    # ── 자기검사: 모든 패널의 행 라벨이 제자리에 남아 있는가 (겹쳐 쓰기 검출)
    for p, col, span in list(across) + list(stacked):
        for r in guide_rows(p["grid"]):
            lab = str(r["label"]).strip()
            if p.get("transpose"):                    # 전치 패널은 보기가 열 머리로 남는다
                found = any(str(ws.cell(rr, cc).value or "").strip() == lab
                            for rr in range(row, end + 1) for cc in range(col, col + span + p["n_data"] + 1))
            else:
                found = any(str(ws.cell(rr, col).value or "").strip() == lab for rr in range(row, end + 1))
            if not found:
                print(f"⚠ 요약 패널 겹침 — '{slide_title}' {p.get('ref')} 행 라벨 '{lab[:30]}' 이 열 {col} 에 남아 있지 않습니다")
                break
    # ── 표 아래 여백 한 줄 — LG 는 여기에 thickTop 을 남긴다
    tail = ws.row_dimensions[end + 1]
    tail.height = HT_PANEL
    tail.thickTop = True
    return end + 3                                    # 블록 사이 빈 행 2줄


# ── 목차.md 파싱 ───────────────────────────────────────────────────────────────
#    확정 목차(코드블록) · Part Ⅲ 배정표 · Part Ⅱ 요약/결론 설계를 읽는다.
#    **수치를 새로 만들지 않는다** — 목차.md 에 적힌 문장을 그대로 옮긴다.

#: 확정 목차 코드블록의 한 줄은 여러 항목을 **두 칸 이상**으로 벌려 적는다.
#    `   1. 시장      2. 공백      8. 결론 및 제언`
#    항목명 자체에 공백이 들어갈 수 있으므로(`결론 및 제언` — 고정 섹션명) 먼저
#    두 칸 이상으로 쪼갠 뒤 각 토큰을 통째로 읽는다. 꼬리 괄호(`(12장)`)는 버린다.
RE_TOC_ITEM = re.compile(r"^(\d+)\.\s*(.+?)\s*(?:\([^)]*\))?$")
RE_PART_HEAD = re.compile(r"^Part\s+([ⅠⅡⅢⅣ])\.\s*(.+?)\s*$")
RE_TOC_GROUP = re.compile(r"^([A-Z])\.\s+(\S.*?)\s*$")      #: 대분류 행 `A. 건물 관리 현황` (한 줄에 하나, 챕터 항목과 같은 줄 금지)
RE_PII_HEAD = re.compile(r"^\*\*(\d+)\.\s*(.+?)\*\*\s*←\s*(.*?)\s*$")

#: 요약 꼭지의 **패널 배치 지시**. 꼭지 제목 바로 아래 펜스 코드블록 안에 한 줄씩 적는다.
#:
#:     [좌]   P61 B5. 월 평균 지불 비용      전체 가로막대 · Base 300
#:     [중]   표19 A8. 관리 업무 만족도_TOP2 Summary   전체 가로막대 · Base 500
#:     [좌]   P65 D10. 가격 제시 후 이용 의향   배너 표 (용도·권역·관리 운영 방식) · Base 500
#:
#: `P{n}` 은 본문 슬라이드의 **페이지 번호**(slide_facts 의 `page`), `표{n}` 은 통계표
#: 인덱스다 (본문에 없는 Summary 표라 tables.json 에서 직접 가져온다).
#: 제목과 차트 지시는 **두 칸 이상**으로 갈린다 — 제목 안에는 두 칸이 없다.
RE_PANEL = re.compile(
    r"^\[(?P<pos>상|좌|중|우|우하|하)\]\s+(?P<ref>P\d+|표\d+)\s+(?P<title>.+?)\s{2,}"
    r"(?P<chart>.+?)\s*·\s*Base\s*(?P<base>[\d,]+)\s*$")

#: 패널 순서 = 읽는 순서. `우하` 는 `우` 아래에 놓으므로 가로 배치에서는 빠진다.
#: `하` 는 **전폭 하단 층** — 좌 열(C)에서 시작해 오른쪽으로 뻗는다. 배너 표처럼 넓은 패널의 자리다.
#: 가이드의 패널 셀 위치가 곧 PPT 배치다 (열 = 가로 위치, 층 = 세로 위치, 같은 열 두 층 = 상하 분할 —
#: LG s8 ↔ 본문_제작용 r130/r153 1:1 실측, 사용자 진술 2026-08-23).
PANEL_ORDER = ("상", "좌", "중", "우", "우하", "하")

#: 마크다운 표의 `\|` 는 열 구분자가 아니라 셀 안의 층위 구분자다
#    (`1) 소유 \| 형태` — LG 예시가 3단 구조에 쓰는 표기).
_MD_PIPE = re.compile(r"(?<!\\)\|")


def _split_md_row(line: str) -> list[str]:
    body = _MD_PIPE.sub("\x00", line.strip()).strip("\x00")
    return [c.replace("\\|", "|").strip() for c in body.split("\x00")]


def _toc_items(line: str) -> list[str]:
    out: list[str] = []
    for tok in re.split(r"\s{2,}", line.strip()):
        m = RE_TOC_ITEM.match(tok.strip())
        if m:
            out.append(f"{m.group(1)}. {m.group(2)}")
    return out


def _clean(s: str) -> str:
    return s.replace("**", "").replace("`", "").strip()


def _sentences(text: str) -> list[str]:
    """문단을 문장 단위로 쪼갠다. 소수점은 뒤에 공백이 없으므로 안 걸린다."""
    text = _clean(text).lstrip("- ").strip()
    return [s.strip() for s in re.split(r"(?<=\.)\s+", text) if s.strip()]


def _closes_ment(text: str) -> bool:
    """개조식 한 줄이 끝났는가 — 마지막 글자에 **ㅁ 받침**이 있으면 끝이다.

    `임 · 음 · 함 · 섬 · 름 · 감 · 답함` … 개조식 종결은 전부 ㅁ 받침이다
    (rules.yaml V-060 이 강제하는 그 어미). 목차.md 의 멘트는 마크다운 소프트랩으로
    두세 줄에 걸쳐 적혀 있어서, 물리 줄을 그대로 옮기면 한 문장이 두 줄로 쪼개진다.
    """
    ch = (text or "").rstrip().rstrip("*」』)]”'\"")[-1:]
    if not ch:
        return False
    o = ord(ch)
    return 0xAC00 <= o <= 0xD7A3 and (o - 0xAC00) % 28 == 16      # 종성 16 = ㅁ


def _ment_lines(body: list[str]) -> list[str]:
    """물리 줄 목록 → 멘트 줄 목록. 개조식이 끝날 때까지 이어 붙인다."""
    out: list[str] = []
    buf = ""
    for ln in body:
        t = _clean(ln).strip()
        if not t:
            continue
        buf = f"{buf} {t}" if buf else t
        if _closes_ment(buf):
            out.append(buf)
            buf = ""
    if buf:
        out.append(buf)
    return out


def parse_toc(path: Path, part_label: "dict[str, str] | None" = None) -> dict[str, Any]:
    text = io.open(path, encoding="utf-8").read()
    lines = text.splitlines()

    # ① 확정 목차 — 첫 펜스 코드블록
    fence, block = False, []
    for ln in lines:
        if ln.strip().startswith("```"):
            if fence:
                break
            fence = True
            continue
        if fence:
            block.append(ln)

    parts: dict[str, dict] = {}
    cur = None
    grp = None
    for ln in block:
        m = RE_PART_HEAD.match(ln.strip())
        if m:
            cur = m.group(1)
            grp = None
            # 표기 번호는 갈아 끼울 수 있다 — 요약·결론(Part Ⅱ)을 만들지 않는 조사에서
            # 「Ⅰ 다음에 Ⅲ」이 비어 보이면 「1, 2」로 이어 붙인다 (사용자 지시 2026-08-26).
            # 내부 키는 로마숫자 그대로 두어야 아래 배정 로직이 성립한다.
            shown = (part_label or {}).get(cur, cur)
            parts[cur] = {"title": f"Part {shown}. {m.group(2)}", "items": [], "groups": []}
            continue
        if not (cur and ln.strip()):
            continue
        # R-140 대분류 층 (사용자 지시 2026-08-23 · SKEE 실측 A~F): `A. 건물 관리 현황` 한 줄 — 그 아래 챕터들이 묶인다
        mg = RE_TOC_GROUP.match(ln.strip())
        if mg:
            grp = {"label": f"{mg.group(1)}. {mg.group(2).strip()}", "items": []}
            parts[cur]["groups"].append(grp)
            continue
        items = _toc_items(ln)
        parts[cur]["items"].extend(items)
        if grp is not None:
            grp["items"].extend(items)
    for part in parts.values():
        part["group_of"] = {it: g["label"] for g in part["groups"] for it in g["items"]}

    # ② Part Ⅲ 배정표 — 챕터별 서브 타이틀 (중복 제거 · 등장 순서 유지)
    chapter_subs: dict[str, list[str]] = {}
    assignments: list[dict] = []
    for ln in lines:
        if not ln.startswith("|"):
            continue
        col = _split_md_row(ln)
        if len(col) < 5 or not col[0].isdigit():
            continue
        chapter, sub = _clean(col[1]), _clean(col[4])
        subs = chapter_subs.setdefault(chapter, [])
        if sub and sub not in subs:
            subs.append(sub)
        # 연구부 확인용 시트는 이 표를 **한 행도 접지 않고** 그대로 편다.
        assignments.append({
            "page": int(col[0]), "chapter": chapter,
            "seq": _clean(col[2]) if len(col) > 2 else "",
            "question": _clean(col[3]) if len(col) > 3 else "",
            "sub": sub,
            "tables": _clean(col[5]) if len(col) > 5 else "",
            "base": _clean(col[6]) if len(col) > 6 else "",
        })

    # ③ Part Ⅱ 요약 · 결론 — `**N. 이름** ← 근거` + 본문
    def harvest(start_pat: str, stop_pat: str) -> list[dict]:
        """`**N. 이름** ← 근거` 를 머리로 하는 꼭지들을 긁어 온다.

        꼭지 제목 바로 아래의 **펜스 코드블록은 패널 배치 지시**이고 (`panels`),
        그 뒤 문단이 멘트다 (`body`). 결론 꼭지에는 코드블록이 없으므로 같은 함수가
        그대로 통한다.
        """
        out: list[dict] = []
        inside, cur_item, fence = False, None, False
        for ln in lines:
            if not fence and re.match(start_pat, ln):
                inside, cur_item, fence = True, None, False
                continue
            if inside and not fence and re.match(stop_pat, ln):
                break
            if not inside:
                continue
            if cur_item is not None and ln.strip().startswith("```"):
                fence = not fence
                continue
            if fence:
                m = RE_PANEL.match(ln.strip())
                if m:
                    cur_item["panels"].append({
                        "position": m.group("pos"),
                        "ref": m.group("ref"),
                        "title": _clean(m.group("title")),
                        "chart": _clean(m.group("chart")),
                        "base_n": int(m.group("base").replace(",", "")),
                    })
                continue
            m = RE_PII_HEAD.match(ln.strip())
            if m:
                cur_item = {"no": int(m.group(1)), "title": _clean(m.group(2)),
                            "source": _clean(m.group(3)), "body": [], "panels": []}
                out.append(cur_item)
            elif (cur_item is not None and ln.strip() and not ln.startswith("|")
                    and not re.fullmatch(r"[-=_*\s]+", ln)):
                cur_item["body"].append(ln.strip())
        return out

    summaries = harvest(r"^###\s+요약\s", r"^###\s+결론|^##\s")
    conclusions = harvest(r"^###\s+결론\s", r"^##\s")
    # 결론 레이아웃 지시 원문 — `레이아웃: …` 줄 (없으면 빌더가 박스 수로 제안한다)
    conclusion_layout_text = None
    inside = False
    for ln in lines:
        if re.match(r"^###\s+결론\s", ln):
            inside = True
            continue
        if inside and ln.startswith("레이아웃:"):
            conclusion_layout_text = _clean(ln.split(":", 1)[1])
            break
        if inside and re.match(r"^##\s", ln):
            break

    return {
        "parts": parts,
        "chapter_subs": chapter_subs,
        "assignments": assignments,
        "summaries": {s["no"]: s for s in summaries},
        "conclusions": conclusions,
        "conclusion_layout": conclusion_layout_text,
    }


def load_overview(path: Path | None) -> dict[str, Any]:
    """조사 목적 · 조사 설계. 조사개요·RFP 에서 사람이 뽑아 둔 JSON 을 읽기만 한다."""
    if not path or not path.exists():
        return {}
    with io.open(path, encoding="utf-8") as fh:
        return json.load(fh)


def load_tables_index(tables: Path | None) -> dict[int, dict]:
    """통계표 인덱스 → 표. 요약 패널의 `표{n}` 참조가 여기서 그리드를 가져온다."""
    if not tables or not tables.exists():
        return {}
    with io.open(tables, encoding="utf-8") as fh:
        doc = json.load(fh)
    tlist = doc["tables"] if isinstance(doc, dict) else doc
    return {t["index"]: t for t in tlist if t.get("index") is not None}


def load_profile(tables: Path | None, extra: list[tuple[int, str]] | None = None,
                 groups: list[str] | None = None, extra_first: bool = False) -> list[tuple]:
    """표1(응답자 특성)의 **전체 열**. 항상 첫 표다 (guide/10 §1).

    `extra` = [(표 인덱스, 그룹명)] — 표1 에 없는 응답자 속성을 사람이 확정해 덧붙인다
    (예: SQ2 응답자 역할 — B2B 조사에서 「누가 답했나」가 표1 에 없을 때. R-022 의 예외).

    `groups` = 표1 에서 **실을 `label_group` 화이트리스트**. 주면 그 군만 싣는다.
        DP 가 배너 변수를 전치해 표1 을 만들면 그 행들이 본문 장과 그대로 겹친다
        (에스원 TA Pool: 46행 중 33행이 본문 표4·6~8·15~18 과 동일). 전량 실으면
        같은 수치가 보고서에 두 번 나와 MECE 가 깨진다. 채택 규칙은
        **「본문에 장이 없는 변수만 싣는다」** 이고, 그 판정은 사람이 해서 이 인자로 준다.

    `extra_first` = `extra` 를 표1 행보다 **앞에** 놓는다. 표1 에 응답자 본인 속성이
        한 행도 없는 조사에서 「누가 답했나」를 맨 위에 두어 슬라이드 제목을 참으로 만든다.
    """
    if not tables or not tables.exists():
        return []
    with io.open(tables, encoding="utf-8") as fh:
        doc = json.load(fh)
    tlist = doc["tables"] if isinstance(doc, dict) else doc
    t1 = next((t for t in tlist if t.get("index") == 1), None)
    if not t1:
        return []
    # LG 는 `Base for %` 위에 배너 헤더 줄을 한 줄 둔다 (실측 `D76 = 전체`).
    # 이게 없으면 아래 `(500)` 이 무엇의 사례수인지 표에 안 적힌 채로 나간다.
    head_rows: list[tuple] = [
        ("", "", "전체"),
        ("", "Base for %", f"({t1.get('base_n')})"),
    ]
    t1_rows: list[tuple] = []
    seen_group = None
    for r in t1.get("rows") or []:
        if r.get("kind") != "option" or not r.get("label"):
            continue
        grp = r.get("label_group") or ""
        if groups is not None and grp not in groups:
            continue
        head = grp if grp != seen_group else ""
        seen_group = grp
        vals = r.get("values") or []
        display = vals[0].get("display") if vals else None
        t1_rows.append((head, r["label"], _num(display) or ""))
    extra_rows: list[tuple] = []
    for idx, group in extra or []:
        t = next((x for x in tlist if x.get("index") == idx), None)
        if not t:
            continue
        first = True
        for r in t.get("rows") or []:
            if r.get("kind") != "option" or not r.get("label"):
                continue
            vals = r.get("values") or []
            display = vals[0].get("display") if vals else None
            extra_rows.append((group if first else "", r["label"].strip(), _num(display) or ""))
            first = False
    return head_rows + (extra_rows + t1_rows if extra_first else t1_rows + extra_rows)


#: 결론 장표 스케치 앵커 — write_structure_head 가 채우고 main 이 소비한다
SKETCH_ANCHORS: list[dict] = []
SKETCH_ROWS = 6 + 2 + 6 + 3


def _xml_text(t: str) -> str:
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;"))


def draw_conclusion_sketch(out: Path, sheet_name: str, anchor_row: int, titles: list[str]) -> str | None:
    """결론 블록 아래에 LG 식 레이아웃 스케치를 **도형으로** 그려 넣는다.

    LG `본문_제작용` 결론 블록 실측: Chart type 산문 지시 옆에 둥근 박스 3개(제목 띠 포함) → 아래 화살표 →
    전폭 총결론 박스. 제작팀은 이 그림을 보고 PPT 를 짠다 — 글로만 쓰면 배치가 제작팀 재량이 된다.
    마지막 항목이 총결론이다. 박스 N-1 개를 한 줄에, 총결론 1개를 전폭으로.
    openpyxl 은 도형을 못 쓰므로 zip 수술로 drawing 파트를 추가한다 (R-039 계열).
    """
    import zipfile
    n = len(titles)
    if n < 2:
        return None
    top_titles, final_title = titles[:-1], titles[-1]
    # 열 배치: C(2) 부터 시작, 박스 하나당 3열 폭 + 1열 간격. 행: 앵커 + 1 부터 박스 6행, 화살표 2행, 총결론 6행
    r0 = anchor_row - 1                          # 0-based
    box_h, gap_h, final_h = 6, 2, 6
    col0, box_w, gap_w = 2, 3, 1

    def anchor(c1, r1, c2, r2, shape_id, name, prst, fill, line, text="", bold=False, size=1100, tcolor="FFFFFF"):
        body = ""
        if text:
            body = (f'<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
                    f'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="ko-KR" sz="{size}" b="{1 if bold else 0}">'
                    f'<a:solidFill><a:srgbClr val="{tcolor}"/></a:solidFill></a:rPr><a:t>{_xml_text(text)}</a:t></a:r></a:p></xdr:txBody>')
        ln = f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line else '<a:ln><a:noFill/></a:ln>'
        fl = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else '<a:noFill/>'
        return (f'<xdr:twoCellAnchor><xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                f'<xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{shape_id}" name="{_xml_text(name)}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
                f'<xdr:spPr><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fl}{ln}</xdr:spPr>{body}</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>')

    shapes, sid = [], 1001
    for i, t in enumerate(top_titles):
        c1 = col0 + i * (box_w + gap_w)
        shapes.append(anchor(c1, r0, c1 + box_w, r0 + box_h, sid, f"결론박스{i + 1}", "roundRect", None, "2E75B6")); sid += 1
        shapes.append(anchor(c1, r0, c1 + box_w, r0 + 1, sid, f"결론제목{i + 1}", "rect", "5B9BD5", None, text=f"{i + 1}. {t}", bold=True)); sid += 1
    width_cols = len(top_titles) * (box_w + gap_w) - gap_w
    mid = col0 + width_cols // 2
    shapes.append(anchor(mid, r0 + box_h, mid + 1, r0 + box_h + gap_h, sid, "화살표", "downArrow", "5B9BD5", None)); sid += 1
    shapes.append(anchor(col0, r0 + box_h + gap_h, col0 + width_cols, r0 + box_h + gap_h + final_h, sid, "총결론박스", "rect", "5B9BD5", None,
                         text=f"총결론 — {final_title}", bold=True, size=1200)); sid += 1
    drawing_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                   'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + "".join(shapes) + '</xdr:wsDr>')

    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        data = {n_: z.read(n_) for n_ in names}
    wb_xml = data["xl/workbook.xml"].decode("utf-8")
    m = re.search(rf'<sheet[^>]*name="{re.escape(sheet_name)}"[^>]*/>', wb_xml)
    if not m:
        return None
    rid = re.search(r'r:id="(rId\d+)"', m.group(0)).group(1)
    wrels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")
    tgt = re.search(rf'<Relationship[^>]*Id="{rid}"[^>]*Target="([^"]+)"', wrels) or re.search(rf'<Relationship[^>]*Target="([^"]+)"[^>]*Id="{rid}"', wrels)
    tgt = tgt.group(1)
    sheet_part = tgt.lstrip("/") if tgt.startswith("/") else ("xl/" + tgt if not tgt.startswith("xl/") else tgt)
    sx = data[sheet_part].decode("utf-8")
    if "<drawing " in sx:
        return None                                     # 이미 드로잉이 있는 시트는 건드리지 않는다
    k = 1
    while f"xl/drawings/drawing{k}.xml" in data:
        k += 1
    dname = f"xl/drawings/drawing{k}.xml"
    rels_name = f"xl/worksheets/_rels/{sheet_part.split('/')[-1]}.rels"
    rels_xml = data.get(rels_name, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>').decode("utf-8")
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)]
    drid = f"rId{max(ids) + 1 if ids else 1}"
    rels_xml = rels_xml.replace("</Relationships>",
        f'<Relationship Id="{drid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{k}.xml"/></Relationships>')
    # <drawing> 은 pageSetup/headerFooter 뒤, legacyDrawing/tableParts/extLst 앞
    tag = f'<drawing r:id="{drid}"/>'
    if "xmlns:r=" not in sx[:600]:
        sx = sx.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
    m2 = re.search(r"<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b", sx)
    sx = sx[:m2.start()] + tag + sx[m2.start():] if m2 else sx.replace("</worksheet>", tag + "</worksheet>")
    ct = data["[Content_Types].xml"].decode("utf-8")
    if f'PartName="/{dname}"' not in ct:
        ct = ct.replace("</Types>", f'<Override PartName="/{dname}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    data[dname] = drawing_xml.encode("utf-8")
    data[rels_name] = rels_xml.encode("utf-8")
    data[sheet_part] = sx.encode("utf-8")
    data["[Content_Types].xml"] = ct.encode("utf-8")
    tmp = out.with_suffix(out.suffix + ".tmp")
    with zipfile.ZipFile(out) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        done = set()
        for info in zin.infolist():
            zout.writestr(info, data[info.filename]); done.add(info.filename)
        for n_, b in data.items():
            if n_ not in done:
                zout.writestr(n_, b)
    tmp.replace(out)
    return dname


def write_structure_head(ws, row: int, toc: dict, overview: dict,
                         profile: list[tuple], skin: Skin,
                         by_page: dict[int, dict] | None = None,
                         by_index: dict[int, dict] | None = None,
                         warn=None) -> tuple[int, int, int]:
    """표지 → 목차 → Part Ⅰ → Part Ⅱ → Part Ⅲ 간지까지. (다음 행, 블록 수, 최대 열)

    구조 블록과 요약·결론 블록은 LG 에서 **같은 계열**(11pt)이다 — 본문(9pt 맑은 고딕)과
    다른 것은 글꼴뿐이고 노랑 띠·밝은 띠·어두운 띠 체계는 동일하다.
    """
    n = 0
    max_col = COL_FIRST_DATA
    by_page = by_page or {}
    by_index = by_index or {}
    warn = warn or (lambda msg: None)
    parts = toc["parts"]

    row = write_struct_block(ws, row, skin, chart_type="표지"); n += 1
    row = write_struct_block(ws, row, skin, chart_type="목차"); n += 1

    # ── Part Ⅰ. 조사 개요
    p1 = parts.get("Ⅰ")
    if p1:
        row = write_struct_block(ws, row, skin, slide_title=p1["title"], chart_type="간지",
                                 rows=[(i, "") for i in p1["items"]]); n += 1
        items = p1["items"]

        def kv_rows(groups) -> list[tuple]:
            """LG 실측 배치 — 값 열에 구분명, 행 라벨 열에 내용. 구분명은 첫 줄에만.
            `table` 이 있는 구분(표본 구성 — 사용자 지시 2026-08-23 「표로 제시」)은 헤더 줄 + 행을 열 단위 튜플로 내보낸다
            (C=구분 · D=항목 · E=사례수 · F=비율) — 제작팀이 PPT 표로 옮긴다."""
            out: list[tuple] = []
            for g in groups or []:
                first = True
                tbl = g.get("table")
                if tbl:
                    out.append((g.get("label"), *(tbl.get("columns") or tbl.get("header") or [])))
                    first = False
                    for r in tbl.get("rows", []):
                        out.append(("", *r))
                for v in g.get("values") or []:
                    out.append((g.get("label") if first else "", v))
                    first = False
            return out

        # 1. 조사목적  /  2. 조사설계
        if len(items) > 0:
            row = write_struct_block(ws, row, skin, slide_title=items[0],
                                     rows=kv_rows(overview.get("purpose"))); n += 1
        if len(items) > 1:
            row = write_struct_block(ws, row, skin, slide_title=items[1],
                                     rows=kv_rows(overview.get("design"))); n += 1
        # 3번째 이후 항목 — 기본은 응답자특성(통계표 표1, Chart type = 전체 : 표).
        # 트래킹 PT 처럼 Part Ⅰ 에 정의 장(Convergence 정의 등)이 더 붙는 하우스는
        # overview["extra"][항목명] 의 그룹(purpose/design 과 같은 스키마)으로 내용을 준다.
        extra_items = overview.get("extra") or {}
        for item in items[2:]:
            if "응답자" in item and profile:
                # profile[0] 은 배너 헤더 줄(`전체`), profile[1] 이 `Base for %` 줄이다 — 첫 줄을
                # 집으면 `[Base: 전체(n=전체)]` 가 나간다 (guide-front F-002 실측).
                base_row = next((r for r in profile if r[1] == "Base for %"), None)
                base_n = base_row[2].strip("()") if base_row else ""
                # 군 이름은 조사개요가 줄 수 있다. 표본 단위가 「전체」로 뭉뚱그려지면
                # 안 되는 조사가 있다 — 에스원 TA Pool 은 86 이 건물이 아니라 응답 건이다(R-186).
                base_label = overview.get("profile_base_label") or "전체"
                row = write_struct_block(ws, row, skin, slide_title=item,
                    chart_type="전체 : 표",
                    base_unit=f"[Base: {base_label}(n={base_n})/ Unit: %]",
                    qu="\n".join(overview.get("profile_notes") or []),
                    rows=profile, table=True); n += 1
            else:
                g = extra_items.get(item) or extra_items.get(item.split(".", 1)[-1].strip())
                if g is None and "응답자" in item:
                    g = []          # 표1 이 없는 프로젝트의 응답자특성 — 빈 블록으로라도 자리를 낸다
                # 정의 장에도 제작 지시가 필요하다. 비어 있으면 제작팀이 무엇을 그릴지 모른다.
                ct = (overview.get("extra_chart_types") or {}).get(item, "")
                row = write_struct_block(ws, row, skin, slide_title=item,
                    chart_type=ct, rows=kv_rows(g)); n += 1

    # ── Part Ⅱ. 요약 및 결론
    p2 = parts.get("Ⅱ")
    if p2:
        row = write_struct_block(ws, row, skin, slide_title=p2["title"], chart_type="간지",
                                 rows=[(i, "") for i in p2["items"]]); n += 1
        for item in p2["items"]:
            no = int(item.split(".")[0])
            s = toc["summaries"].get(no)
            if s:
                qu = f"* 근거: {s['source']}" if s.get("source") else ""
                # `Slide title` 은 **확정 목차**의 항목명을 쓴다 (검수가 이 이름으로
                # 요약 블록을 알아본다). 요약 꼭지 제목과 어긋나면 목차.md 안에서
                # 두 곳이 따로 논다는 뜻이므로 조용히 넘기지 않는다.
                if s["title"] and s["title"] != item.split(".", 1)[-1].strip():
                    warn(f"요약 {item} — 확정 목차는 '{item}' 인데 요약 꼭지 제목은 "
                         f"'{no}. {s['title']}' 입니다 (확정 목차를 씁니다)")
                panels = [p for p in (
                    panel_source(spec, by_page, by_index, warn)
                    for spec in sorted(s.get("panels") or [],
                                       key=lambda x: PANEL_ORDER.index(x["position"])))
                    if p]
                if panels:
                    above, across, stacked, _ = plan_summary(panels)
                    max_col = max([max_col] + [c + sp + p["n_data"] - 1
                                               for p, c, sp in above + across + stacked])
                    row = write_summary_block(ws, row, skin, slide_title=item, qu=qu,
                        ment=_ment_lines(s["body"]), panels=panels,
                        sub_title=(toc.get("_subtitles") or {}).get(no, "")); n += 1
                    continue
                # 배치 지시가 없는 꼭지는 예전처럼 텍스트만 (LG 이전 판)
                warn(f"요약 {item} — 목차.md 에 패널 배치 지시가 없어 텍스트로만 씁니다")
                body: list[tuple] = []
                for ln in s["body"]:
                    body += [(t, "") for t in _sentences(ln)]
                row = write_struct_block(ws, row, skin, slide_title=item,
                    chart_type=SUMMARY_CHART, base_unit=SUMMARY_BASE,
                    qu=qu, rows=body); n += 1
            else:
                # 목차.md 의 요약 번호에 없는 항목 = 결론 꼭지
                body = []
                for c in toc["conclusions"]:
                    body.append((f"{c['no']}. {c['title']}", f"← {c['source']}"))
                    for ln in c["body"]:
                        body += [(t, "") for t in _sentences(ln)]
                n_items = len(toc["conclusions"])
                # 스케치 도형 앵커 — `Data` 라벨 아래 첫 행. 제목 띠에 박스 제목을 적는다 (LG 본문_제작용 결론 블록 실측)
                SKETCH_ANCHORS.append({"row": row + len(LABELS) + 1 + len(body) + 2,   # 텍스트 줄 아래
                                       "titles": [c["title"] for c in toc["conclusions"]]})
                row_before = row
                row = write_struct_block(ws, row, skin, slide_title=item,
                    chart_type=toc.get("conclusion_layout") or conclusion_layout(n_items),
                    # 결론은 표가 없는 텍스트 장표다. Base 를 하나로 못 박으면 거짓이 된다 —
                    # 박스마다 근거 요약이 다르고 그 요약들의 Base 가 87·198·266·300·500 으로 다르다.
                    # LG 도 이 칸을 비워 뒀다(r309). 근거는 `Question/Reference` 의 요약 추적으로 남는다.
                    base_unit="",
                    # 근거 표기는 결론 항목들이 실제로 가리키는 것을 모아 쓴다. 요약 꼭지에서
                    # 도출한 덱이면 「요약 꼭지」가, 본문 장에서 직접 끌어온 덱이면 그 장 번호가 나온다.
                    qu=("* 근거: " + " / ".join(
                        f"{c['no']}. {c['source']}" for c in toc["conclusions"] if c.get("source"))
                        if body and any(c.get("source") for c in toc["conclusions"])
                        else ("* 근거: Part Ⅱ 요약 꼭지" if body else "")),
                    rows=body); n += 1
                # 스케치 도형 자리 예약 — 박스 6행 + 화살표 2행 + 총결론 6행 + 여백 (다음 블록과 겹치지 않게)
                row += SKETCH_ROWS

    # ── Part Ⅲ. 조사 결과 (챕터 간지는 본문 사이에 들어간다)
    p3 = parts.get("Ⅲ")
    if p3:
        row = write_struct_block(ws, row, skin, slide_title=p3["title"], chart_type="간지",
                                 rows=[(i, "") for i in p3["items"]]); n += 1
    return row, n, max_col


def write_toc(ws, slides: list[dict], toc: dict | None = None) -> None:
    """`목차(보고서용)` — A열 Part 행, B열 세부 항목. C1 안내문은 건드리지 않는다 (R-036).

    목차.md 가 없으면 슬라이드 타이틀만으로 Part Ⅲ 만 적는다 — 지어내지 않는다.
    """
    if toc and toc.get("parts"):
        r = 2
        for key in ("Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ"):
            part = toc["parts"].get(key)
            if not part:
                continue
            ws.cell(r, 1, part["title"])
            _style(ws, r, 1, bold=True)
            r += 1
            # R-140 대분류는 목차 시트에 쓰지 않는다 (사용자 지시 2026-08-23 「이렇게 소목차 하지마」) —
            # 챕터만 B열에 평평하게. 대분류는 slide_facts 의 group 필드로만 남긴다.
            for item in part["items"]:
                ws.cell(r, 2, item)
                r += 1
            r += 1
        return

    titles: list[str] = []
    for s in slides:
        t = s.get("slide_title")
        if t and t not in titles:
            titles.append(t)
    if not titles:
        return
    ws.cell(3, 1, "Part Ⅲ. 조사 결과")
    for i, t in enumerate(titles):
        ws.cell(4 + i, 2, t)


#: 연구부 확인용 시트의 열. `표` 가 이 시트의 존재 이유다 —
#: **목차 한 줄마다 어느 통계표에서 나오는지를 못 박는 자리**(사용자 지시 I-030).
#: `Chart type`·`Data` 는 본문_제작용 블록의 같은 이름 라벨을 그대로 옮긴 것이다 (사용자 지시 2026-08-27).
#: 연구원이 이 시트만 훑어도 「무슨 그림으로 그리는가 · 무슨 멘트가 붙는가」를 볼 수 있게 한다 —
#: 본문 시트로 건너뛰지 않아도 되도록. 자리는 `구성 및 BASE` 바로 뒤다.
RESEARCH_TOC_COLS = ("챕터", "서브 타이틀", "문항코드", "질문", "표", "구성 및 BASE",
                     "Chart type", "Data", "P", "요약·결론", "링크")


def appendix_specs(toc: dict, by_page: dict, by_index: dict, warn) -> list[dict]:
    """요약·결론이 쓴 표 가운데 **본문에 실리지 않은 것**을 부록 대상으로 모은다 (사용자 지시 2026-08-24).

    「요약 및 결론에 나온 숫자는 다 조사 결과에서 찾을 수 있어야 한다」가 요구다. 교차 패널은 본문 절단선
    밖 배너군(건물 연면적·입주사 수 등)을 원표에서 끌어오므로, 그 수치는 본문 어디에도 없다.
    그래서 **그 표를 부록에 싣는다** — 부록에 실리면 독자가 찾을 수 있고, 검수 V18 도 통과한다.

    돌려주는 각 항목: {ref, table, groups, title, why}
      · `groups` 가 비면 표 전체가 본문에 없다는 뜻이고, 차 있으면 그 배너군만 본문 밖이라는 뜻이다.
    """
    def body_groups_of(page: int | None, tidx: int | None) -> set[str] | None:
        """본문에 실린 그 표의 배너군 집합. 본문에 없으면 None."""
        sl = by_page.get(page) if page is not None else None
        if sl is None and tidx is not None:
            # 슬라이드는 원천 표를 `source_tables` 로 들고 있다 (순위 슬라이드는 두 개)
            sl = next((x for x in by_page.values() if tidx in (x.get("source_tables") or [])), None)
        if sl is None:
            return None
        g = (sl.get("grid") or {}).get("banner_group_row") or []
        return {str(x).strip() for x in g[1:] if str(x).strip()}

    out: list[dict] = []
    seen: set[tuple] = set()
    for s in (toc.get("summaries") or {}).values():
        for spec in s.get("panels") or []:
            ref = str(spec.get("ref") or "")
            chart = str(spec.get("chart") or "")
            if not chart.startswith("배너"):
                continue                       # 전체값 패널은 본문에도 같은 값이 있다
            m = re.search(r"\(([^)]*)\)", chart)
            groups = [g.strip() for g in re.split(r"[·,/]", m.group(1))] if m else []
            page = int(ref[1:]) if ref.startswith("P") else None
            tidx = int(ref[1:]) if ref.startswith("표") else None
            bg = body_groups_of(page, tidx)
            if bg is None:
                why, missing = "본문에 없는 표", groups
            else:
                missing = [g for g in groups if g and g not in bg]
                if not missing:
                    continue                   # 본문에 이미 같은 배너군으로 실려 있다
                why = "본문에 싣지 않은 배너 기준"   # 독자 대면 표기 — 내부 은어(절단선·배너군) 금지 (초독자 감사 2026-08-28)
            key = (ref, tuple(sorted(missing)))
            if key in seen:
                continue
            seen.add(key)
            t = by_index.get(tidx) if tidx is not None else None
            title = (t or {}).get("title_full") if t else (by_page.get(page) or {}).get("title_full")
            out.append({"ref": ref, "table": tidx, "page": page, "groups": missing,
                        "title": str(title or spec.get("topic") or ref), "why": why,
                        "used_by": f"요약 {s.get('no')}"})
    return out


def write_appendix(ws, row: int, skin_body: Skin, skin_struct: Skin, specs: list[dict],
                   by_page: dict, by_index: dict, warn) -> tuple[int, int, int]:
    """Part Ⅳ 간지 + 부록 블록. 돌려주는 값은 (다음 행, 블록 수, 최대 열)."""
    if not specs:
        return row, 0, 0
    import build_facts as BF
    row = write_struct_block(ws, row, skin_struct, slide_title="Part Ⅳ. Appendix",
                             chart_type="간지",
                             rows=[(f"{i}. {sp['title']}", "") for i, sp in enumerate(specs, 1)])
    n, max_col = 1, 0
    for i, sp in enumerate(specs, 1):
        if sp["table"] is not None:
            t = by_index.get(sp["table"])
            if t is None:
                warn(f"부록 {sp['ref']} — 통계표를 찾지 못해 건너뜁니다")
                continue
            grid = BF.build_grid(t, extra_groups=set(sp["groups"]) or None, small_base_mark=False)
            qu = t.get("title_full") or ""
            base_n = t.get("base_n")
        else:
            sl = by_page.get(sp["page"])
            if sl is None:
                warn(f"부록 {sp['ref']} — 본문 슬라이드를 찾지 못해 건너뜁니다")
                continue
            src = next((x for x in by_index.values()
                        if x.get("question_code") == sl.get("question_code")
                        and str(x.get("title_full") or "").strip() == str(sl.get("title_full") or "").strip()), None)
            grid = BF.build_grid(src, extra_groups=set(sp["groups"]) or None, small_base_mark=False) if src else sl["grid"]
            qu = re.sub(r"^Qu:\s*", "", str(sl.get("qu") or "")).strip()
            base_n = sl.get("base_n")
        slide = {"slide_title": f"[부록 {i}] {sp['title']}",
                 "sub_title": f"{sp['used_by']} 근거 — {sp['why']}"
                              + (f" ({' · '.join(sp['groups'])})" if sp["groups"] else ""),
                 "chart_type": "전체 : 표, 배너 : 표",
                 "base_unit": (f"[Base: 전체(n={base_n})/ Unit: %"
                               + (f", {_mean_unit_of([r for r in grid['rows'] if r.get('kind') == 'mean'])}"
                                  if any(r.get('kind') == 'mean' for r in grid['rows']) else "")
                               + "]") if base_n else "",
                 "qu": f"Qu: {qu}" if qu else "", "ment": [], "grid": grid}
        row = write_block(ws, row, slide, skin_body)
        n += 1
        max_col = max(max_col, COL_FIRST_DATA + max(len(r["cells"]) for r in grid["rows"]) - 1)
    return row, n, max_col


def part2_usage(toc_path: Path | None) -> dict:
    """Part Ⅱ 에 쓰인 표·페이지 → 꼭지/결론 라벨. summary_map.json 의 panels(ref P13/표19)·claims 와
    conclusion_map.json 의 evidence(table) 를 읽는다 — 연구부 시트 「요약·결론」 열(사용자 지시 2026-08-23)."""
    out: dict = {}
    if not toc_path:
        return out
    W = Path(toc_path).parent
    def add(key, label):
        out.setdefault(key, [])
        if label not in out[key]:
            out[key].append(label)
    try:
        sm = json.loads((W / "summary_map.json").read_text(encoding="utf-8"))
        for sres in sm.get("summaries", []):
            # 「요약1」 이 아니라 실제 꼭지 이름을 적는다 — 연구원이 번호를 다시 찾아보지
            # 않아도 되게 (사용자 지시 2026-08-27). 제목이 없으면 번호로 물러선다.
            _t = str(sres.get("title") or "").strip()
            lab = f"{sres['no']}. {_t}" if _t else f"요약{sres['no']}"
            for p in sres.get("panels", []):
                ref = str(p.get("ref") or "")
                mk = re.match(r"표(9\d{3})$", ref)
                if mk:
                    try:
                        kj = json.loads((W / "KPI요약표_직접생성.json").read_text(encoding="utf-8"))
                        kt = next((x for x in kj.get("tables", []) if x.get("index") == int(mk.group(1))), None)
                        picked = set(re.findall(r"'([^']+)'", str(p.get("chart") or "")))
                        for kr in (kt or {}).get("rows", []):
                            if kr.get("source_table_index") and (not picked or str(kr.get("label") or "").strip() in picked):
                                add(("T", int(kr["source_table_index"])), lab)
                    except Exception:                     # noqa: BLE001
                        pass
                m = re.match(r"P(\d+)$", ref)
                if m:
                    add(("P", int(m.group(1))), lab)
                m = re.match(r"표(\d+)$", ref)
                if m:
                    add(("T", int(m.group(1))), lab)
            for c in sres.get("claims", []):
                for e in c.get("evidence", []):
                    if isinstance(e.get("table"), int):
                        add(("T", e["table"]), lab + "(근거)")
    except Exception as e:                           # noqa: BLE001
        print(f"[research_toc] summary_map 읽기 실패: {e}")
    try:
        cm = json.loads((W / "conclusion_map.json").read_text(encoding="utf-8"))
        for it in cm.get("items", []):
            _t = str(it.get("title") or "").strip()
            lab = f"결론 {it['no']}. {_t}" if _t else f"결론{it['no']}"
            for e in it.get("evidence", []):
                if isinstance(e.get("table"), int):
                    add(("T", e["table"]), lab)
    except Exception as e:                           # noqa: BLE001
        print(f"[research_toc] conclusion_map 읽기 실패: {e}")
    # 9001/9002 직접 생성표 → 구성 표 매핑은 없으므로 그대로 둔다(표 번호 열에 안 나타남)
    return out


def write_research_toc(ws, toc: dict | None, banner_note: str | None = None, part2_use: dict | None = None,
                       body_index: dict | None = None) -> None:
    """`part 3 목차(연구부 확인용)` — 본문 한 장 = 한 행.

    LG 는 이 시트에서 여러 문항을 한 행에 묶었다(`A1, A2, A3, A5, A6`). 우리는 접지
    않는다 — 이 시트의 쓰임이 **누락 검출**이기 때문이다(knowledge/guide/10 §2-2).
    묶어 놓으면 빠진 문항이 묶음 안에 숨고, 표 번호도 한 칸에 여러 개가 들어가
    「목차 ↔ 표」 대응이 1:1 로 읽히지 않는다. 70장이면 70행이다.

    서식은 LG 실측 그대로다 — 나눔고딕 11pt, 채움·테두리 없음. 헤더만 굵게 한다.
    """
    if not toc or not toc.get("assignments"):
        return
    rows = sorted(toc["assignments"], key=lambda a: a["page"])

    # 템플릿이 물려준 LG 내용을 먼저 비운다. 열 너비는 아래에서 다시 잡는다.
    if ws.max_row > 1:
        ws.delete_rows(1, ws.max_row)

    font = Font(name="나눔고딕", size=11)

    def put(r: int, c: int, v, *, bold=False):
        cell = ws.cell(r, c, v)
        cell.font = Font(name="나눔고딕", size=11, bold=bold)
        return cell

    put(2, 1, "Part Ⅲ. 조사 결과", bold=True)
    put(2, 3, banner_note or "배너 3군 공통 — 용도(6) · 권역(6) · 관리 운영 방식(4)")
    for i, name in enumerate(RESEARCH_TOC_COLS):
        put(4, 1 + i, name, bold=True)

    r = 5
    prev_chapter = None
    group_of = (toc.get("parts", {}).get("Ⅲ") or {}).get("group_of") or {}
    prev_group = None
    for a in rows:
        if prev_chapter is not None and a["chapter"] != prev_chapter:
            r += 1                                   # 챕터 사이 한 줄 비움 (LG 실측)
        prev_chapter = a["chapter"]                  # 대분류 행은 쓰지 않는다 (사용자 지시 2026-08-23)
        bi = (body_index or {}).get(a["page"]) or {}
        put(r, 1, a["chapter"])
        put(r, 2, a["sub"])
        put(r, 3, a["question"])
        put(r, 4, bi.get("qu") or "")
        put(r, 5, a["tables"])
        put(r, 6, a["base"])
        put(r, 7, bi.get("chart_type") or "")
        c_ment = put(r, 8, bi.get("ment") or "")
        c_ment.alignment = Alignment(vertical="top", wrap_text=True)
        put(r, 9, a["page"])
        if part2_use:
            tags = set(part2_use.get(("P", a["page"]), [])) | {t for n in re.findall(r"\d+", str(a.get("tables") or "")) for t in part2_use.get(("T", int(n)), [])}
            if tags:
                put(r, 10, " · ".join(sorted(tags, key=lambda x: (x.startswith("결론"), x))))
        if bi.get("row"):                            # 본문_제작용 의 Slide title 행으로 가는 하이퍼링크 (사용자 지시 2026-08-23)
            c = put(r, 11, f"P{a['page']} 표 보기")
            from openpyxl.worksheet.hyperlink import Hyperlink
            c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{bi.get('sheet', '본문_제작용')}'!A{bi['row']}", display=f"P{a['page']} 표 보기")
            c.font = Font(name="나눔고딕", size=11, underline="single", color="0563C1")
        r += 1

    for col, w in zip("ABCDEFGHIJK", (14.6, 26.0, 10.0, 60.0, 10.0, 40.0, 30.0, 64.0, 6.0, 30.0, 12.0)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"


def swap_logo(out: Path, logo: Path) -> tuple[int, int] | None:
    """Cover 시트의 클라이언트 로고를 갈아 끼운다.

    템플릿은 예시(LG 홈로봇)에서 복제했으므로 `xl/media/image1.png` 가 **LG전자 로고**다.
    그대로 두면 에스원 조사 가이드가 LG 로고를 달고 제작팀에 간다.

    **가로세로비를 지킨다.** 로고는 「글자의 꼴, 굵기, 비례, 자간의 임의 변경 금지」가
    CI 규정이라 늘려 끼울 수 없다(에스원 CI 안내 실측). 앵커 상자의 **높이를 유지하고
    너비만 새 이미지 비율로 다시 계산**한다 — 왼쪽 위 자리는 그대로 남는다.

    돌려주는 값은 (새 너비 EMU, 높이 EMU). 로고나 그림 파트가 없으면 None.
    """
    import zipfile
    from PIL import Image

    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
        if "xl/media/image1.png" not in names or "xl/drawings/drawing1.xml" not in names:
            return None
        draw = z.read("xl/drawings/drawing1.xml").decode("utf-8")
        parts = {n: z.read(n) for n in z.namelist()}
        infos = list(z.infolist())

    with Image.open(logo) as im:
        ratio = im.width / im.height

    m = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"\s*/>', draw)
    m_off = re.search(r"<from>.*?<colOff>(\d+)</colOff>.*?</from>", draw, re.S)
    if not m or not m_off:
        return None
    cy = int(m.group(2))
    cx = round(cy * ratio)
    draw = draw[:m.start()] + f'<a:ext cx="{cx}" cy="{cy}"/>' + draw[m.end():]
    # twoCellAnchor 의 `to` 도 같이 좁힌다 — 안 맞추면 Excel 이 상자를 다시 늘린다.
    left = int(m_off.group(1))
    draw = re.sub(r"(<to>.*?<colOff>)\d+(</colOff>)",
                  lambda mm: f"{mm.group(1)}{left + cx}{mm.group(2)}", draw, count=1, flags=re.S)

    parts["xl/drawings/drawing1.xml"] = draw.encode("utf-8")
    parts["xl/media/image1.png"] = logo.read_bytes()

    tmp = out.with_suffix(out.suffix + ".logo")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in infos:
            zout.writestr(item, parts[item.filename])
    shutil.move(str(tmp), str(out))
    return cx, cy


#: 어느 PC 에나 있는 서체. 이 목록에 없는 이름은 **예시사 전용 서체**로 본다.
#  템플릿을 LG 홈로봇에서 복제했기 때문에 `LG Smart UI Regular` 가 딸려 온다 —
#  제작팀 PC 에 그 서체가 없으면 Excel 이 임의로 대체하면서 행 높이가 어긋나고,
#  무엇보다 **에스원 조사 가이드에 LG 전용 서체가 박혀 나간다.**
SAFE_FONTS = {
    "맑은 고딕", "나눔고딕", "굴림", "굴림체", "돋움", "돋움체", "바탕", "바탕체",
    "Arial", "Calibri", "Consolas", "Times New Roman", "Wingdings", "Malgun Gothic",
}

RE_FONT_NAME = re.compile(r'<name val="([^"]+)"/>')


def normalize_fonts(out: Path, to: str = "맑은 고딕") -> dict[str, int]:
    """예시사 전용 서체를 표준 서체로 바꾼다. 돌려주는 값은 {바뀐 이름: 정의 수}.

    `styles.xml` 한 곳만 고치면 **행 서식과 셀 서식이 동시에** 따라온다 —
    둘 다 같은 글꼴 정의를 인덱스로 가리키기 때문이다. 셀을 훑어 고치면
    행 서식(`<row s="178">`)에 걸린 글꼴이 남는다.

    크기·굵기·색은 건드리지 않는다. **이름만** 바꾼다.
    """
    import zipfile

    with zipfile.ZipFile(out) as z:
        parts = {n: z.read(n) for n in z.namelist()}
        infos = list(z.infolist())
    if "xl/styles.xml" not in parts:
        return {}

    xml = parts["xl/styles.xml"].decode("utf-8")
    hit: dict[str, int] = {}

    def sub(m: re.Match) -> str:
        name = m.group(1)
        if name in SAFE_FONTS or name == to:
            return m.group(0)
        hit[name] = hit.get(name, 0) + 1
        return f'<name val="{to}"/>'

    xml = RE_FONT_NAME.sub(sub, xml)
    if not hit:
        return {}

    parts["xl/styles.xml"] = xml.encode("utf-8")
    tmp = out.with_suffix(out.suffix + ".font")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in infos:
            zout.writestr(item, parts[item.filename])
    shutil.move(str(tmp), str(out))
    return hit


def restore_ooxml_parts(template: Path, out: Path) -> dict[str, Any]:
    """openpyxl 이 흘린 OOXML 파트를 템플릿에서 되살린다.

    왜 필요한가: `keep_vba=True` 로 저장해도 openpyxl 은 워크시트 XML의 **꼬리**를 다시
    쓰면서 `<controls>` 블록(폼 컨트롤 정의)과 `printerSettings`·`ctrlProp` 관계를 버린다.
    `vbaProject.bin` 은 살아남지만 **버튼이 매크로에 연결된 고리가 끊긴다.**

    실측 손실: `xl/drawings/drawing2.xml`(ItemCheck_PPT `초기화`),
    `drawing3.xml`(ItemCheck_한글 `초기화`), `printerSettings*.bin` 9개,
    워크시트 `_rels` 7개, 그리고 Cover 의 `검수 실행` 버튼 `<controls>` 정의.

    Cover 시트 안내문이 "검수 실행은 … 반드시 실행해서 TRUE를 꼭 확인해주세요" 라고
    지시하는 그 버튼이므로, 끊긴 채로 제작팀에 보낼 수 없다.

    안전한 이유: 생성물은 sharedStrings 없이 인라인 문자열을 쓴다. 셀 데이터는 그대로 두고
    `<pageSetup>` 이후 꼬리만 템플릿 것으로 갈아끼우므로 인덱스가 어긋날 여지가 없다.
    """
    import os
    import re
    import shutil as _sh
    import tempfile

    needs_tail = re.compile(r"<controls>|<drawing\b|<oleObjects\b")

    # 꼬리 시작점: pageSetup 이 있으면 거기서, 없으면(ItemCheck_한글) drawing 에서.
    # 시트마다 어떤 요소로 끝나는지가 달라 하나의 정규식으로는 못 잡는다.
    _tail_res = (
        re.compile(r"<pageSetup\b.*?</worksheet>", re.S),
        re.compile(r"<drawing\b.*?</worksheet>", re.S),
        re.compile(r"<legacyDrawing\b.*?</worksheet>", re.S),
    )

    def find_tail(xml: str):
        for rx in _tail_res:
            m = rx.search(xml)
            if m:
                return m
        return None

    with zipfile.ZipFile(template) as zt, zipfile.ZipFile(out) as zg:
        t_names, g_names = set(zt.namelist()), set(zg.namelist())
        patched: dict[str, bytes] = {}
        restored: list[str] = []
        spliced: list[str] = []

        # ① 사라진 파트를 템플릿에서 그대로 되살린다.
        #    단 workbook.xml.rels 가 가리키지 않는 워크시트는 **일부러 지운 것**(송부본의 `_` 시트)이다 —
        #    되살리면 탭에는 안 보이는 고아 파트로 내용이 통째로 남는다 (guide-adversary ADV-02).
        g_rels = zg.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        referenced = {("xl/" + t.lstrip("/")) if not t.startswith("xl/") else t
                      for t in re.findall(r'Target="([^"]+)"', g_rels)}
        for name in sorted(t_names - g_names):
            if name == "xl/sharedStrings.xml":
                continue          # 생성물은 인라인 문자열을 쓴다 — 되살리면 오히려 충돌
            m_ws = re.match(r"xl/worksheets/(sheet\d+\.xml)$", name)
            if m_ws and name not in referenced:
                continue          # 삭제한 시트 — 고아 파트로 되살리지 않는다
            if re.match(r"xl/worksheets/_rels/(sheet\d+\.xml)\.rels$", name) and \
               name.replace("/_rels/", "/").removesuffix(".rels") not in referenced:
                continue
            patched[name] = zt.read(name)
            restored.append(name)

        # ①-b 버튼 도형(AlternateContent)이 있는 드로잉을 openpyxl 재직렬화로 잃은 경우.
        #     버튼만 있는 드로잉(초기화)은 파트째 사라져 ①이 되살리지만, 그림이 함께 있는
        #     드로잉(Cover: 로고+검수 실행)은 파트가 '남아서' ①에 안 걸린다 — 그림만 남고
        #     버튼 도형이 소리 없이 빠진다. 템플릿 것으로 통째로 되살린다.
        #     (로고 교체 swap_logo 는 이 복원 '뒤'에 돌므로 그림 편집은 유지된다)
        for name in sorted(n for n in t_names & g_names
                           if re.match(r"xl/drawings/drawing\d+\.xml$", n)):
            t_xml = zt.read(name).decode("utf-8")
            if "<mc:AlternateContent" not in t_xml:
                continue
            if "<mc:AlternateContent" in zg.read(name).decode("utf-8"):
                continue
            patched[name] = t_xml.encode("utf-8")
            restored.append(name)
            rel = name.replace("drawings/", "drawings/_rels/") + ".rels"
            if rel in t_names:
                patched[rel] = zt.read(rel)

        # ② 버튼·드로잉이 있는 시트는 꼬리를 템플릿 것으로 갈아끼운다
        for name in sorted(n for n in t_names if n.startswith("xl/worksheets/sheet")):
            t_xml = zt.read(name).decode("utf-8")
            if not needs_tail.search(t_xml) or name not in g_names:
                continue
            t_tail = find_tail(t_xml)
            if not t_tail:
                continue
            g_xml = zg.read(name).decode("utf-8")
            g_tail = find_tail(g_xml)
            if g_tail:
                g_xml = g_xml[:g_tail.start()] + t_tail.group(0)
            else:
                g_xml = g_xml.replace("</worksheet>", t_tail.group(0))
            # 템플릿 꼬리는 r:/xdr:/mc: 접두를 쓴다. openpyxl 이 쓴 루트에는 이 선언이
            # 없어서(inline 으로만 붙인다) 그대로 접합하면 unbound prefix 로 깨진다.
            ns = {
                "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
                # MCE: `<mc:Choice Requires="x14">` 는 x14 접두가 in-scope 여야 한다. 없으면 XML 은
                # well-formed 라 openpyxl 은 읽지만 **Excel 은 0x800A03EC 로 파일을 거부한다**
                # (guide-adversary ADV-01 · Excel COM 실측). LG 예시 루트에는 선언돼 있다.
                "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
            }
            missing = " ".join(f'{k}="{v}"' for k, v in ns.items() if f"{k}=" not in g_xml[:400])
            if missing:
                g_xml = g_xml.replace("<worksheet ", f"<worksheet {missing} ", 1)
            patched[name] = g_xml.encode("utf-8")
            spliced.append(name)
            # 그 시트의 관계 파일도 템플릿 것으로 (drawing·ctrlProp·printerSettings rId 포함)
            rel = name.replace("worksheets/", "worksheets/_rels/") + ".rels"
            if rel in t_names:
                patched[rel] = zt.read(rel)

        # ③ [Content_Types].xml — 되살린 파트 선언을 합친다
        ct = "[Content_Types].xml"
        t_ct = zt.read(ct).decode("utf-8")
        g_ct = zg.read(ct).decode("utf-8")
        final_names = (g_names - {"[Content_Types].xml"}) | set(patched)
        for frag in re.findall(r"<(?:Override|Default)[^>]*/>", t_ct):
            key = re.search(r'(?:PartName|Extension)="([^"]+)"', frag)
            if not key or f'"{key.group(1)}"' in g_ct:
                continue
            # 실제로 존재하는 파트만 선언한다. 없는 파트를 선언하면 Excel 이 복구를 요구하고
            # openpyxl 은 KeyError 로 죽는다 (되살리지 않은 sharedStrings.xml 이 그 경우다).
            name = key.group(1)
            if name.startswith("/") and name.lstrip("/") not in final_names:
                continue
            g_ct = g_ct.replace("</Types>", frag + "</Types>")
        patched[ct] = g_ct.encode("utf-8")

        # ④ 새 zip 으로 다시 쓴다
        fd, tmp_name = tempfile.mkstemp(suffix=".xlsm")
        os.close(fd)          # Windows 는 열린 핸들이 있으면 rename 이 막힌다
        tmp = Path(tmp_name)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
            for item in zg.infolist():
                zw.writestr(item, patched.pop(item.filename, zg.read(item.filename)))
            for name, data in patched.items():
                zw.writestr(name, data)

    _sh.move(str(tmp), str(out))
    return {"restored": restored, "spliced": spliced}


def main() -> int:
    ap = argparse.ArgumentParser(description="slide_facts.json → 제작가이드 .xlsm")
    ap.add_argument("facts", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument("--template", type=Path, default=TEMPLATE)
    ap.add_argument("--toc", type=Path, default=None,
                    help="확정 목차(work/목차.md) — 표지·목차·간지·요약 블록의 원본")
    ap.add_argument("--tables", type=Path, default=None,
                    help="통계표 JSON (기본: facts 폴더의 tables.json) — 표1 응답자 특성용")
    ap.add_argument("--overview", type=Path, default=None,
                    help="조사 목적·설계 JSON (기본: 목차.md 폴더의 조사개요.json)")
    ap.add_argument("--title", default=None, help="Cover!C4 보고서 제목")
    ap.add_argument("--date", default=None, help="Cover!C6 제출일")
    ap.add_argument("--font", default="맑은 고딕",
                    help="예시사 전용 서체를 이 서체로 정규화한다 (빈 값이면 정규화 안 함)")
    ap.add_argument("--logo", type=Path,
                    help="Cover 로고 PNG — 템플릿의 예시사(LG) 로고를 이걸로 갈아 끼운다")
    ap.add_argument("--no-send-copy", action="store_true", help="_제작팀송부본을 만들지 않음")
    ap.add_argument("--creator", default="엠브레인", help="문서 속성 작성자·최종 수정자에 넣을 조직명 (R-137)")
    ap.add_argument("--dash-empty", action="store_true",
                    help="표 데이터 칸의 빈 자리를 `-` 로 메운다. 하우스가 갈리는 자리다 — "
                         "SKEE 본문에는 `-` 가 수백 개이고 LG 송부본에는 0개다. "
                         "빈 칸으로 두면 제작팀이 「값이 없음」과 「적는 걸 잊었음」을 구분하지 못한다")
    ap.add_argument("--part-label", default="",
                    help="Part 번호의 **표기**를 갈아 끼운다 — \"Ⅰ=1;Ⅲ=2\" (세미콜론 구분). "
                         "요약·결론(Part Ⅱ)을 만들지 않아 「Ⅰ 다음에 Ⅲ」이 되는 조사에서 쓴다. "
                         "목차.md 와 내부 키는 로마숫자 그대로 두고 화면 글자만 바뀐다")
    ap.add_argument("--chapter-covers", action="store_true",
                    help="Part Ⅲ 챕터 간지 블록을 본문에 넣는다 (기본 꺼짐 — R-138, 사용자 지시 2026-08-23)")
    ap.add_argument("--profile-extra", default="",
                    help="응답자 특성에 덧붙일 표 — \"3=응답자 역할\" (인덱스=그룹명, 세미콜론 구분). R-022 예외, 사람 확정")
    ap.add_argument("--profile-groups", default="",
                    help="표1 에서 실을 행 그룹 화이트리스트 — \"지역;연면적\" (세미콜론 구분). "
                         "생략하면 표1 전 행을 싣는다. DP 가 배너 변수를 전치해 표1 을 만든 조사에서는 "
                         "본문 장과 겹치는 군을 여기서 떨군다 (채택 규칙: 본문에 장이 없는 변수만 싣는다)")
    ap.add_argument("--profile-extra-first", action="store_true",
                    help="--profile-extra 행을 표1 행보다 앞에 놓는다. 표1 에 응답자 본인 속성이 없는 조사에서 "
                         "「누가 답했나」를 맨 위로 올려 슬라이드 제목을 참으로 만든다")
    ap.add_argument("--banner-note", default=None,
                    help="연구부 확인용 시트 C2 배너 안내 문구 (기본: 배너 3군 공통 …)")
    ap.add_argument("--cover", action="append", default=[], metavar="CELL=값",
                    help="Cover 시트 셀 직접 지정 (반복 가능). 예: --cover \"C19=에스원 Blue(#0033A0)\" "
                         "— 템플릿은 범용이므로 프로젝트 값은 인자로 준다 (R-127)")
    args = ap.parse_args()

    if not args.template.exists():
        print(f"[중단] 템플릿이 없습니다: {args.template}", file=sys.stderr)
        return 1

    with io.open(args.facts, encoding="utf-8") as fh:
        doc = json.load(fh)
    slides = doc["slides"]

    part_label = {}
    for kv in (args.part_label or "").split(";"):
        if "=" in kv:
            k, v = kv.split("=", 1)
            part_label[k.strip()] = v.strip()
    toc = parse_toc(args.toc, part_label) if args.toc and args.toc.exists() else None
    tables_path = args.tables or (args.facts.parent / "tables.json")
    overview_path = args.overview or (args.toc.parent / "조사개요.json" if args.toc else None)
    extra_profile: list[tuple[int, str]] = []
    for tok in (args.profile_extra or "").split(";"):
        tok = tok.strip()
        if tok and "=" in tok and tok.split("=", 1)[0].strip().isdigit():
            i, g = tok.split("=", 1)
            extra_profile.append((int(i), g.strip()))
    profile_groups = [g.strip() for g in (args.profile_groups or "").split(";") if g.strip()] or None
    profile = load_profile(tables_path, extra_profile,
                           groups=profile_groups,
                           extra_first=args.profile_extra_first) if toc else []
    overview = load_overview(overview_path) if toc else {}

    args.out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(args.template, args.out)

    wb = openpyxl.load_workbook(args.out, keep_vba=True)
    body_name = next((n for n in BODY_SHEETS if n in wb.sheetnames), None)
    if body_name is None:
        print(f"[중단] 본문 시트가 없습니다 (찾은 이름: {', '.join(BODY_SHEETS)})",
              file=sys.stderr)
        return 1
    ws = wb[body_name]

    # ── ① 템플릿의 LG 원본 행에서 서식 두 벌을 실측한다 (인덱스를 박지 않는다)
    skin_struct, skin_body = measure_skins(ws)
    # ── ② 물려받은 행 서식을 **전부 지운다**. 안 지우면 LG 의 원래 행 위치에 노랑 띠가
    #      남아 우리 블록과 어긋난다 (사용자가 본 "이상한 형광색"의 정체).
    ws.row_dimensions.clear()
    ws.merged_cells.ranges = []

    if args.title:
        wb["Cover"]["C4"] = args.title
    if args.date:
        wb["Cover"]["C6"] = args.date
    for spec in args.cover:
        cell, _, value = spec.partition("=")
        cell = cell.strip().upper()
        if not re.fullmatch(r"[A-Z]{1,2}\d{1,3}", cell) or not value:
            print(f"[중단] --cover 형식 오류: {spec!r} (CELL=값)", file=sys.stderr)
            return 1
        wb["Cover"][cell] = value

    row = 3
    first_row = row
    max_col = COL_FIRST_DATA
    n_struct = 0

    notes: list[str] = []
    if toc:
        # 요약 패널의 원자료 — `P{n}` 은 본문 슬라이드의 페이지, `표{n}` 은 통계표 인덱스다.
        by_page = {s["page"]: s for s in slides if s.get("page") is not None}
        by_index = load_tables_index(tables_path)
        # 통계표에 없어 직접 만든 요약표(make_kpi_summary · 인덱스 9001~)도 `표{n}` 으로 참조된다.
        # tables.json 과 같은 스키마라 그대로 합친다 (R-019).
        kpi_path = args.facts.parent / "KPI요약표_직접생성.json"
        if kpi_path.exists():
            by_index.update(load_tables_index(kpi_path))
        toc["_subtitles"] = summary_subtitles(args.toc)
        row, n_struct, sum_col = write_structure_head(
            ws, row, toc, overview, profile, skin_struct,
            by_page=by_page, by_index=by_index, warn=notes.append)
        max_col = max(max_col, sum_col)

    body_index: dict = {}       # page → 본문 행·질문 (연구부 시트 하이퍼링크·질문 열용, 사용자 지시 2026-08-23)
    chapter = object()          # 첫 슬라이드에서 반드시 간지가 한 번 나가게 하는 센티널
    for slide in slides:
        if toc and args.chapter_covers:          # R-138: 조사 결과(Part Ⅲ) 간지는 기본적으로 내지 않는다 (사용자 지시 2026-08-23)
            title = slide.get("slide_title")
            if title != chapter:
                chapter = title
                subs = toc["chapter_subs"].get(title or "", [])
                row = write_struct_block(ws, row, skin_struct,
                                         slide_title=title or "",
                                         chart_type="간지",
                                         rows=[(s, "") for s in subs])
                n_struct += 1
        n_b = slide["grid"]["n_banners"]
        overlay = slide["grid"].get("rank_overlay")
        max_col = max(max_col, COL_FIRST_DATA + n_b + (1 if overlay else 0))
        # 키는 **본문 블록의 등장 순번**이다 — 연구부 확인용 시트의 `P` 열과 같은 축.
        # 팩트의 `page` 를 키로 쓰면, 한 장이 블록 여러 개로 쪼개지는 덱에서 질문 열과
        # 하이퍼링크가 통째로 다른 블록을 가리킨다 (adversary F-02·F-03 실측).
        _ment = slide.get("ment")
        body_index[len(body_index) + 1] = {
            "row": row, "qu": re.sub(r"^Qu:\s*", "", str(slide.get("qu") or "")).strip(),
            "sheet": ws.title,
            "chart_type": str(slide.get("chart_type") or "").strip(),
            # 멘트는 여러 줄이다 — 셀 안에서 줄바꿈으로 잇는다 (연구부 시트 `Data` 열)
            "ment": "\n".join(str(x).strip() for x in (_ment if isinstance(_ment, list) else [_ment])
                              if x and str(x).strip())}
        row = write_block(ws, row, slide, skin_body)

    # ── Part Ⅳ. Appendix — 요약이 쓴 표 중 본문에 없는 것 (사용자 지시 2026-08-24)
    apx = appendix_specs(toc, by_page, by_index, notes.append) if toc else []
    if apx:
        row, n_apx, apx_col = write_appendix(ws, row, skin_body, skin_struct, apx,
                                             by_page, by_index, notes.append)
        n_struct += n_apx
        max_col = max(max_col, apx_col)
        print("부록 " + str(len(apx)) + "표 = " + " · ".join(
            f"{a['ref']}({'·'.join(a['groups']) or '표 전체'})" for a in apx))
    last_row = row

    ws.column_dimensions["A"].width = 20     # 블록 라벨
    ws.column_dimensions["B"].width = 62     # 값 · 멘트 · 행수식어
    ws.column_dimensions["C"].width = 44     # 행 라벨
    for c in range(COL_FIRST_DATA, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 11

    if args.dash_empty:
        n_dash = dash_empty_cells(ws)
        print(f"빈 칸 `-` 채움: {n_dash}개 (매크로 lovelySeoyeon 과 같은 처리)")
    n_num = numify_data_cells(ws)
    if n_num:
        print(f"데이터 칸 숫자화: {n_num}개 (문자열로 두면 엑셀이 「텍스트로 저장된 숫자」로 잡는다)")

    write_toc(wb[TOC_SHEET], slides, toc)
    if RESEARCH_TOC_SHEET in wb.sheetnames:
        write_research_toc(wb[RESEARCH_TOC_SHEET], toc, args.banner_note, part2_use=part2_usage(args.toc), body_index=body_index)
    wb.save(args.out)

    # ── 제작팀 송부본: `_` 접두 시트 제거.
    #    파트 복원 '전' 상태에서 만든다 — 복원된 파일을 openpyxl 로 다시 열면
    #    되살린 mc:AlternateContent 를 파서가 감당하지 못한다.
    send_path, removed = None, []
    if not args.no_send_copy:
        send_path = args.out.with_name(args.out.stem + "_제작팀송부본" + args.out.suffix)
        shutil.copy2(args.out, send_path)
        wb2 = openpyxl.load_workbook(send_path, keep_vba=True)
        removed = [n for n in wb2.sheetnames if n.startswith("_")]
        for n in removed:
            del wb2[n]
        if "Basic_Set" in wb2.sheetnames:            # 템플릿 내부 시트(블록 라벨 원본·RPA 픽커) — 제작팀이 볼 것이 아니다 (사용자 지시 2026-08-23)
            wb2["Basic_Set"].sheet_state = "hidden"
        wb2.save(send_path)
        restore_ooxml_parts(args.template, send_path)

    # ── openpyxl 이 흘린 파트를 되살린다 (버튼·컨트롤·프린터설정)
    fixed = restore_ooxml_parts(args.template, args.out)

    # ── 클라이언트 로고 교체 (파트 복원 뒤 — 복원은 템플릿에서 파트를 되살리므로
    #    먼저 끼우면 LG 로고가 되돌아온다)
    logo_size = None
    if args.logo:
        logo_size = swap_logo(args.out, args.logo)
        if send_path:
            swap_logo(send_path, args.logo)

    # ── 예시사 전용 서체 정규화 (파트 복원 뒤 — 복원이 styles.xml 을 되돌릴 수 있다)
    fonts_fixed: dict[str, int] = {}
    if args.font:
        fonts_fixed = normalize_fonts(args.out, args.font)
        if send_path:
            normalize_fonts(send_path, args.font)

    # ── `dimension` 선언을 실제 범위에 맞춘다.
    #    openpyxl 로 쓴 뒤 zip 수술로 시트를 갈아 끼우면 선언이 갱신되지 않아
    #    Excel 이 "내용에 문제가 있습니다" 복구 안내를 띄운다. **매 빌드마다 다시
    #    어긋나므로** 별도 도구로 두지 않고 여기서 부른다 (잊으면 그대로 재발한다).
    from fix_dimension import fix as fix_dimension
    dim_fixed = fix_dimension(args.out)
    if send_path:
        fix_dimension(send_path)

    # ── 결론 장표 스케치 도형 (LG 식) — 드로잉 파트 주입은 모든 zip 수술 뒤에
    sketch_parts: list[str] = []
    for a in SKETCH_ANCHORS:
        for target in ([args.out] + ([send_path] if send_path else [])):
            d_ = draw_conclusion_sketch(target, body_name, a["row"], a["titles"])
            if d_:
                sketch_parts.append(f"{target.name}:{d_}")

    # ── 문서 위생 (R-137): 문서 속성의 개인 이름·템플릿 잔재를 조직명으로, 보이지 않는 문자 제거 — 두 파일 모두
    import doc_hygiene
    for target in ([args.out] + ([send_path] if send_path else [])):
        tmp = target.with_suffix(target.suffix + ".hyg")
        hr = doc_hygiene.clean_file(target, tmp, args.creator)
        shutil.move(tmp, target)
        if hr["props_changed"] or hr["invisible"]:
            print(f"문서 위생: {target.name} — 속성 {'정리' if hr['props_changed'] else '그대로'} · 제거 문자 {sum(hr['invisible'].values())}개")

    # ── baseline 스냅샷: 이후 사람이 고친 diff 가 학습 데이터가 된다
    baseline = args.out.parent / ".baseline" / args.out.name
    baseline.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(args.out, baseline)

    if args.logo:
        if logo_size:
            print(f"Cover 로고 교체: {args.logo.name} "
                  f"({logo_size[0] / 914400 * 2.54:.1f} × {logo_size[1] / 914400 * 2.54:.1f} cm, 비율 유지)")
        else:
            print(f"⚠ 로고를 끼우지 못했습니다 — 템플릿에 그림 파트가 없습니다")

    if fonts_fixed:
        print("전용 서체 정규화 → " + args.font + ": " + " · ".join(
            f"{k}({v}개 정의)" for k, v in fonts_fixed.items()))

    if dim_fixed:
        print("dimension 선언 정정: " + " · ".join(
            f"{n.rsplit('/', 1)[-1]} {decl}→{act}" for n, decl, act in dim_fixed))

    untitled = sum(1 for s in slides if not s.get("slide_title"))
    no_ment = sum(1 for s in slides if not s.get("ment"))

    print(f"블록 {n_struct + len(slides)}개 = 구조 {n_struct} + 본문 {len(slides)}"
          f" · {last_row - first_row}행 · 최대 {max_col}열  [{body_name}]")
    if toc:
        parts = " / ".join(f"{k}({len(v['items'])})" for k, v in toc["parts"].items())
        print(f"목차: Part {parts} · 챕터 간지 {len(toc['chapter_subs'])}개")
        n_panels = sum(len(s.get("panels") or []) for s in toc["summaries"].values())
        if n_panels:
            print(f"요약 패널 {n_panels}개 = " + " · ".join(
                f"{no}.{s['title']}({len(s.get('panels') or [])})"
                for no, s in sorted(toc["summaries"].items())))
        for w in notes:
            print(f"⚠ {w}")
        if not profile:
            print("⚠ 통계표를 못 찾아 `응답자특성` 블록이 비었습니다: " + str(tables_path))
        if not overview:
            print("⚠ 조사개요 JSON 이 없어 `조사목적`·`조사설계` 블록이 비었습니다: "
                  + str(overview_path))
    else:
        print("⚠ --toc 를 주지 않아 구조 블록(표지·목차·간지·요약)이 빠졌습니다")
    if not args.dash_empty:
        print("빈 칸은 그대로 둡니다 — `-` 로 메우지 않습니다 (LG 송부본 실측: `-` 0개)")
    print(f"OOXML 파트 복원: {len(fixed['restored'])}개 · 시트 꼬리 재접합 {len(fixed['spliced'])}개")
    if sketch_parts:
        print(f"결론 스케치 도형 주입: {', '.join(sketch_parts)}")
    print(f"\n  작업용     {args.out}")
    print(f"  스냅샷     {baseline}")
    if send_path:
        print(f"  송부본     {send_path}  (제거한 시트: {', '.join(removed) or '없음'})")
    if untitled:
        print(f"\n⚠ 슬라이드 타이틀 미확정 {untitled}장 — 목차를 확정한 뒤 다시 생성하세요.")
    if no_ment:
        print(f"⚠ 멘트 미작성 {no_ment}장 — 멘트를 미리 넣어야 제작팀이 멘트까지 만들어 줍니다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
