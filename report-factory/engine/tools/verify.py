#!/usr/bin/env python3
"""제작가이드 `.xlsm` 검수 — 생성물의 모든 수치를 통계표 원본과 1:1 대조한다.

이 검사가 파이프라인의 안전장치다. 룰엔진이 숫자를 확정하고 AI는 문장만 쓰게 설계했으므로,
**생성물에 통계표에 없는 숫자가 있으면 그건 환각이거나 버그**다. 둘 다 잡아야 한다.

검사 항목
--------
  V1  블록 무결성      6개 영문 라벨이 순서대로 있는가
  V2  그리드 수치      데이터 칸이 통계표 원본 값과 일치하는가
  V3  멘트 수치        멘트에 등장하는 모든 숫자가 그 슬라이드 그리드 안에 있는가  ★환각 검출
  V4  Base 표기        [Base: …(n=NN)…] 의 n 이 통계표 base 와 같은가
  V5  Qu 텍스트        설문지와 일치하는가 (응답형식 마커·파이핑은 정규화 후 대조)
  V6  신호어           【TOP】·[평균] 행에 `그래프 상단 Text` 가 붙어 있는가
  V7  미기입 잔존      ???, 내용을 입력하세요, 빈 필수 칸
  V8  템플릿 무결성    숨은 검증 시트 9종과 VBA 가 살아 있는가
  V13 목차 규칙        한 어절 · 중복 배정 없음 · 전수성  (RULES R-027~029)

사용법
------
    python verify.py <guide.xlsm> <tables.json> [--facts slide_facts.json]
        [--questionnaire questionnaire.json] [--out 검수리포트.md]

`목차.md` 를 주면 **목차 단독 모드**(`--mode toc`)로 V13 만 돈다 — 빌드 전에 목차 게이트를
거는 자리다 (guide/18 §5-4 · 같은 규칙의 판정을 두 벌 만들지 않으려고 verify 안에 둔다, R-118).

    python verify.py <work/목차.md> --mode toc --questionnaire questionnaire.json --facts slide_facts_pre.json

제작본 `.pptx` 를 주면 **덱 검수 모드**(`verify_deck.py`, D0~D26)로 갈아탄다.
확장자로 자동 분기하고 `--mode guide|deck` 으로 못 박을 수 있다. 가이드 모드의
인자·동작·종료코드는 그대로다 — 골든셋 회귀 기준을 보존해야 한다 (R-103).

    python verify.py <제작본.pptx> [tables.json] [--facts …] [--questionnaire …]
        [--guide-blocks …] [--deck-facts ground_truth.json] [--known-defects LG]
        [--out 검수리포트.md] [--json findings.json]
"""

from __future__ import annotations

import argparse
import io
import os
import json
import re
import sys
import warnings
import zipfile
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable

warnings.filterwarnings("ignore")
import openpyxl  # noqa: E402

LABELS = ["Slide title", "Sub title", "Chart type",
          "Base/Unit", "Question/Reference", "Data"]

#: 본문 시트 이름은 파일마다 다르다 (LG `본문_제작용` / SKEE `본문`). 먼저 잡히는 것을 쓴다.
#: 검사 내용은 그대로다 — 시트를 **찾는 방법**만 넓혔다.
BODY_SHEETS = ("본문_제작용", "본문")

#: 열 배치는 LG `본문_제작용` 실측 그대로다 — **라벨이 곧 A열**이고 seq 열은 없다.
#:     A 블록 라벨 · B 값/멘트/행수식어 · C 행 라벨 · D 전체(순위면 1순위) · E~ 배너
COL_LABEL, COL_VALUE, COL_ROWLABEL = 1, 2, 3
COL_FIRST_DATA = 4

#: `Chart type` 이 이 셋 중 하나면 구조 블록이다 (R-035).
STRUCT_CHART_TYPES = {"표지", "목차", "간지"}

NUMBER_PAT = re.compile(r"-?\d+\.\d+|-?\d+")

RULES_YAML = Path(__file__).resolve().parents[1] / "learning" / "rules.yaml"


def load_rules() -> dict:
    """`engine/learning/rules.yaml` 을 읽는다.

    검사 항목을 코드에 박지 않고 여기서 가져오는 이유: `/회고` 가 규칙을 추가하면
    **다음 검수부터 바로 적용**돼야 하기 때문이다. 규칙이 늘수록 검증이 촘촘해진다.
    """
    if not RULES_YAML.exists():
        return {}
    try:
        import yaml  # type: ignore
    except ImportError:
        return {}
    with RULES_YAML.open(encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def find_check(rules: dict, section: str, cid: str) -> dict:
    for item in rules.get(section) or []:
        if item.get("id") == cid:
            return item
    return {}


# ── rules.yaml 이 없을 때만 쓰는 최소 폴백
FALLBACK_PROTECTED = ["Cover", "PPT_data", "목차", "한글_data", "ItemCheck_PPT",
                      "ItemCheck_한글", "Verification_PPT", "Verification_한글", "Coding No"]
FALLBACK_MARKERS = ["???", "내용을 입력하세요", "텍스트를 입력"]


@dataclass
class Finding:
    check: str
    severity: str          # 치명 | 경고 | 참고 | 예시오류
    where: str
    detail: str
    #: 실패 태그 (22 체크리스트 §E · rules.yaml#failure_tags). 덱 모드가 채운다.
    #: 기본값이 있어야 가이드 모드의 기존 4인자 생성자 호출이 그대로 산다.
    tag: str | None = None


@dataclass
class Report:
    findings: list[Finding] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)

    def add(self, check: str, severity: str, where: str, detail: str,
            tag: str | None = None) -> None:
        self.findings.append(Finding(check, severity, where, detail, tag))

    @property
    def fatal(self) -> int:
        return sum(1 for f in self.findings if f.severity == "치명")


def _display(cell) -> object:
    """셀을 **화면에 보이는 글자**로 돌려준다.

    데이터 칸의 값은 숫자다(문자열로 두면 엑셀이 「텍스트로 저장된 숫자」로 잡는다).
    그런데 통계표 대조는 **표시값** 기준이므로, 숫자를 그대로 `str()` 하면
    `50.0` 이 `50` 이 되어 「통계표에 없는 값」으로 잡힌다. 자릿수의 정본은
    셀 표시 서식(`###0.0` · `###0.00` · `###0`)이므로 그것을 적용해 문자열로 만든다.
    """
    v = cell.value
    if v is None or isinstance(v, str) or isinstance(v, bool):
        return v
    fmt = str(cell.number_format or "")
    m = re.search(r"###0(?:\.(0+))?", fmt)
    if not m:
        return v
    digits = len(m.group(1) or "")
    # 엑셀은 0.5 를 **올린다**. 파이썬 f-string 은 짝수 반올림이라
    # 31.25 를 31.2 로 만들어 통계표(31.3)와 어긋난다.
    q = Decimal("1." + "0" * digits) if digits else Decimal("1")
    return f"{Decimal(str(float(v))).quantize(q, rounding=ROUND_HALF_UP):f}"


def _blank(v: Any) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def read_blocks(ws, part2_items: list[str] | None = None) -> list[dict]:
    """본문 시트를 블록 단위로 되읽는다.

    빈 칸은 빈 칸이다 — 빌더가 `-` 로 메우지 않는다(LG 실측). 그래서 **멘트와 데이터
    행을 가르는 기준을 `-` 의 유무에 두지 않는다.** 멘트는 `Data` 라벨 바로 아래의
    **연속 구간**이고 그 뒤 첫 빈 행에서 끝난다. 빌더는 본문 블록이면 멘트와 표 사이에,
    구조 블록이면 `Data` 라벨 바로 다음에 빈 행을 한 줄 두므로 이 규칙만으로 갈린다.
    """
    blocks: list[dict] = []
    r, last = 1, ws.max_row
    while r <= last:
        if ws.cell(r, COL_LABEL).value == LABELS[0]:
            block: dict[str, Any] = {
                "row": r, "index": len(blocks), "seq": None, "slide": None,
                "labels": [], "values": {}, "ment": [], "rows": [],
                "banner_row": None, "banner_groups": [], "banner_labels": [],
            }
            for i, label in enumerate(LABELS):
                got = ws.cell(r + i, COL_LABEL).value
                block["labels"].append(got)
                block["values"][label] = ws.cell(r + i, COL_VALUE).value
            rr = r + len(LABELS)
            # ── 멘트: `Data` 라벨 바로 아래의 연속 구간. 첫 빈 행에서 끝난다.
            #    그 뒤의 값 열 문자열은 패널 제목(문항 제목)이므로 멘트가 아니다 —
            #    제목을 멘트로 오독하면 제목 속 문항번호를 수치로 잡아 오탐이 난다.
            while rr <= last:
                v = ws.cell(rr, COL_VALUE).value
                if _blank(v):
                    break
                if ws.cell(rr, COL_LABEL).value is not None or \
                        not _blank(ws.cell(rr, COL_ROWLABEL).value):
                    break
                if isinstance(v, str) and v.strip() != "-":
                    block["ment"].append(v.strip())
                rr += 1
            # ── 데이터 행 · 배너 헤더
            while rr <= last and ws.cell(rr, COL_LABEL).value != LABELS[0]:
                lbl = ws.cell(rr, COL_ROWLABEL).value
                if _blank(lbl):
                    # 값 열·행 라벨 열이 비었는데 전체 열이 차 있으면 배너 헤더 줄이다
                    # (`전체` 또는 `1순위`). **본문 블록에만** 있는 모양이라 구조 블록과
                    # 본문 블록을 가르는 표지가 된다 — 구조 블록의 표(응답자 특성)는
                    # 모든 줄에 행 라벨이 있다.
                    if block["banner_row"] is None and _blank(ws.cell(rr, COL_VALUE).value) \
                            and not _blank(ws.cell(rr, COL_FIRST_DATA).value):
                        block["banner_row"] = rr
                        # 배너 헤더는 두 줄이다 — 위가 그룹, 아래가 라벨.
                        # 좌표 대조(V14)가 이 줄을 통계표 배너 순서와 맞대 본다.
                        _wide = range(COL_FIRST_DATA, ws.max_column + 1)
                        block["banner_groups"] = [ws.cell(rr, c).value for c in _wide]
                        block["banner_labels"] = [ws.cell(rr + 1, c).value for c in _wide]
                else:
                    # 오른쪽 끝까지 읽고 **뒤쪽 빈 칸만** 잘라낸다. 첫 빈 칸에서 멈추면
                    # 값이 없는 배너 칸(통계표에 사례가 없어 빈 칸인 곳)에서 행이
                    # 잘려 뒤쪽 수치를 대조하지 못한다 — `-` 로 메우지 않으니 더 그렇다.
                    cells = [_display(ws.cell(rr, c))
                             for c in range(COL_FIRST_DATA, ws.max_column + 1)]
                    while cells and _blank(cells[-1]):
                        cells.pop()
                    block["rows"].append({
                        "row": rr, "label": str(lbl),
                        "modifier": ws.cell(rr, COL_VALUE).value,
                        "cells": cells,
                    })
                rr += 1
            block["end"] = rr - 1
            # 문항 제목 = 행 라벨 열의 `A3. …` `BA1-1) …` 같은 문항코드 시작 줄
            block["has_panel_title"] = any(
                PANEL_TITLE.match(str(rw.get("label") or ""))
                for rw in block["rows"])
            block["is_struct"] = is_struct_block(block)
            blocks.append(block)
            r = rr
        else:
            r += 1
    mark_struct_blocks(blocks, part2_items)
    return blocks


#: Part 간지의 `Slide title`. 보고서가 어디서 조사 결과로 넘어가는지 알려 주는 표지다.
PART_TITLE = re.compile(r"^\s*Part\s+([ⅠⅡⅢⅣ])\.")


def mark_struct_blocks(blocks: list[dict], part2_items: list[str] | None = None) -> None:
    """구조 블록을 **보고서 안에서의 자리**로 가른다.

    본문(조사 결과)은 `Part Ⅲ` 간지 **뒤에서** 시작한다. 표지·목차·Part Ⅰ·요약·결론은
    전부 그 앞이다. 이 경계는 보고서 구조 자체이므로 표의 생김새와 무관하다.

    ⚠ 왜 블록별 판정만으로는 안 되나: 요약 장표는 본문에서 표를 2~4개 뽑아 붙이는
    다단 패널이라 **문항 제목(`A3. …`)을 그대로 갖는다.** `is_struct_block` 의 문항
    제목 규칙만 두면 요약 7블록이 전부 본문으로 새어 팩트 조인이 7칸 밀린다.
    (그 전에는 배너 헤더의 유무로 갈랐다가 응답자 특성 블록에서 같은 사고가 났다 —
    표의 생김새는 판정 근거가 못 된다.)

    `part2_items` 를 주면 요약·결론 항목명과 `Slide title` 이 맞는지로 한 번 더 확인한다.
    경계를 못 찾으면(= `--toc` 없이 만든 가이드라 간지가 아예 없다) 예전 규칙 그대로 둔다.
    """
    gate = None
    for i, b in enumerate(blocks):
        ct = b["values"].get("Chart type")
        ct = ct.strip() if isinstance(ct, str) else ""
        m = PART_TITLE.match(str(b["values"].get("Slide title") or ""))
        if ct == "간지" and m and m.group(1) == "Ⅲ":
            gate = i
            break
    names = {s.strip() for s in (part2_items or []) if s and s.strip()}
    for i, b in enumerate(blocks):
        title = str(b["values"].get("Slide title") or "").strip()
        if gate is not None:
            # 경계를 찾았으면 그것만 믿는다. 항목명 대조를 여기에 겹치면 Part Ⅱ 항목명이
            # Part Ⅲ 챕터명과 겹칠 때 본문 블록이 구조로 새어 나간다.
            b["is_struct"] = True if i <= gate else is_struct_block(b)
        elif names and title in names:
            b["is_struct"] = True
        else:
            b["is_struct"] = is_struct_block(b)


#: 본문 블록의 문항 제목. `A3. 소유 형태` / `BA1-1) …` / `D7-1. …`
#: 패널 제목 = 문항코드로 시작하는 제목 줄. `A3. 소유 형태` · `B1-1. 공간 구성` · `A2- b. 자산관리 항목`.
#: 하위 항목 표기는 숫자(`B1-1`)일 수도 알파벳(`A2-b`)일 수도 있고, DP 가 `A2- b.` 처럼
#: 띄어 쓰기도 한다 — 에스원 TA Pool 실측 7표. 좁게 잡으면 그 7장이 구조 블록으로 오판돼
#: 조인이 통째로 깨진다(V1 치명).
PANEL_TITLE = re.compile(r"^\s*[A-Z]{1,3}\d+(?:\s*-\s*(?:\d+|[A-Za-z]))?\s*[.)]\s")


def is_struct_block(block: dict) -> bool:
    """구조 블록(표지·목차·간지·조사개요·요약·결론)인가 — 팩트에 대응이 없는 블록.

    두 가지로 본다. `Chart type` 이 구분어(표지·목차·간지)이거나 비어 있으면 구조다.
    그 밖에는 **문항 제목의 유무**로 가른다 — 본문 블록은 데이터 영역에 반드시
    `A3. 소유 형태` 같은 **문항코드로 시작하는 제목 줄**을 갖는다. `3. 응답자특성` 처럼
    통계표를 품은 구조 블록은 `Chart type` 이 `전체 : 표` 라 문자열만으로는 안 갈리는데,
    문항 제목이 없다는 점은 확실히 다르다.

    ⚠ 이전 판은 **배너 헤더 줄의 유무**로 갈랐다. 응답자특성 블록에 `전체` 헤더를
    한 줄 추가하자 그 블록이 본문으로 새어 조인이 통째로 한 칸 밀렸다.
    표의 생김새는 바뀔 수 있으니 판정 근거로 삼으면 안 된다 — 문항 제목은 의미다.
    """
    ct = block["values"].get("Chart type")
    ct = ct.strip() if isinstance(ct, str) else ""
    if ct in STRUCT_CHART_TYPES or ct in ("", "-"):
        return True
    return not block.get("has_panel_title")


def block_where(b: dict) -> str:
    """지적 위치 문자열. seq 는 조인된 팩트에서 온다 (A열에 더 이상 없다)."""
    seq = b.get("seq")
    return f"본문!{b['row']}행" + (f" (seq={seq})" if seq is not None else "")


def join_facts(rep: Report, blocks: list[dict], slides: list[dict]) -> None:
    """본문 블록을 **등장 순서**로 `slide_facts.json` 의 슬라이드와 짝짓는다.

    A열 seq 를 폐지했으므로(LG `본문_제작용` 배치) 조인 키가 따로 없다. 빌더가 팩트
    순서 그대로 본문 블록을 쓰므로 등장 순번이 곧 키다. 구조 블록은 팩트에 대응이
    없으니 조인 대상에서 뺀다.

    개수가 어긋나면 **조용히 넘어가지 않는다.** 조인이 밀린 채로 검사를 계속하면
    V2·V3·V4·V5 가 엉뚱한 슬라이드와 대조해 전부 거짓 지적이 되기 때문이다.
    """
    body = [b for b in blocks if not b["is_struct"]]
    if len(body) != len(slides):
        rep.add("V1", "치명", "본문 시트",
                f"본문 블록 {len(body)}개 ↔ 팩트 슬라이드 {len(slides)}개 — 개수가 맞지 않아 "
                f"순번 조인을 할 수 없습니다 (전체 {len(blocks)}블록 중 구조 블록 "
                f"{len(blocks) - len(body)}개 제외). 수치 대조를 건너뜁니다")
        return
    for b, s in zip(body, slides):
        b["slide"] = s
        b["seq"] = s.get("seq")


#: 보기 라벨은 작은따옴표로 감싼다(멘트 규칙). 그 안의 숫자는 라벨의 일부이지 통계치가 아니다.
QUOTED = re.compile(r"'[^']*'")


#: 통계치가 아닌 숫자 표기 — 순위 스코프(`1~3순위`·`1순위`), 용도 등급(`제1종`), 척도 점수(`5점 척도`),
#: 단계 태그(`[2단계]`). 멘트 관례상 따옴표 없이 쓰이므로 따옴표 마스크로는 못 거른다.
NOT_STAT = re.compile(r"\d+\s*~\s*\d+\s*순위|\d+\s*순위|제\d+종|\d+\s*점\s*척도|\[\d+단계\]|TOP\d|BOT\d|\d+(?:층|개|대|명|동|호|회|건)(?![\d.%])|제시가[^\d]{0,8}\d+만\s?원|\d+만\s?원대")   # 뒤쪽: R-139 단위 병기 보기(5층·1대) · R-142 설문 제시가·가격 눈금(통계치 아님)


def numbers_in(text: str, labels: "Iterable[str] | None" = None) -> set[str]:
    """문장에서 **통계치**를 뽑는다. 라벨·번호·연도는 통계치가 아니므로 뺀다.

    `labels` 를 주면 그 문자열(배너 라벨 `제1종근린생활시설` 등)이 따옴표 없이 등장해도
    통째로 라벨로 본다 — 멘트 2행은 배너명을 따옴표 없이 쓰는 것이 관례다 (패턴 B).
    """
    out: set[str] = set()
    # 따옴표 구간은 통째로 라벨이다 — `'2개 이상 ~ 10개 미만'` 의 2·10 은 수치가 아니고,
    # 층수 문항의 `'5'` 도 마찬가지다. 구간 안의 숫자는 전부 건너뛴다.
    label_spans = [(m.start(), m.end()) for m in QUOTED.finditer(text)]
    label_spans += [(m.start(), m.end()) for m in NOT_STAT.finditer(text)]
    for lab in sorted({l.strip() for l in (labels or []) if isinstance(l, str) and l.strip()}, key=len, reverse=True):
        # 순수 숫자 라벨(`'4'` 층수)은 따옴표 마스크가 맡는다 — 여기서 가리면 `18.0` 안의 `1` 까지 가린다
        if any(ch.isdigit() for ch in lab) and not re.fullmatch(r"[\d.,~\s]+", lab):
            start = 0
            while (i := text.find(lab, start)) >= 0:
                j = i + len(lab)
                edge = (text[i - 1] if i > 0 else " ") + (text[j] if j < len(text) else " ")
                if not re.search(r"[\d.]", edge):       # 숫자·소수점에 붙어 있으면 라벨이 아니라 수치의 일부
                    label_spans.append((i, j))
                start = j

    def in_label(i: int) -> bool:
        return any(a <= i < b for a, b in label_spans)

    for m in NUMBER_PAT.finditer(text):
        if in_label(m.start()):
            continue
        tok = m.group()
        s, e = m.start(), m.end()
        before = text[max(0, s - 2):s]
        after = text[e:e + 1]
        if re.search(r"[A-Za-z]$", before):        # A7, D10 같은 문항코드
            continue
        if after == ")" and "." not in tok:        # 보기번호 `3)`
            continue
        if "." not in tok and len(tok) == 4 and tok.startswith(("19", "20")):
            continue                                # 연도
        out.add(tok)
    return out


def check_freq_cells(rep: Report, block: dict, srcs: list[dict],
                     freq_labels: set[str]) -> None:
    """빈도로 바꿔 적은 배너 열을 통계표에서 역산해 검산한다.

    사람이 `--freq-banners` 로 지정한 열은 `표시값 × n ÷ 100` 으로 적힌다.
    n 이 아주 작은 열에서 백분율이 0.0/100.0 으로만 나오거나 100 을 넘어(중복 슬롯)
    읽을 수 없고, 그 값 자체가 개별 응답의 노출이기 때문이다.
    **대조를 건너뛰는 것이 아니라 다른 식으로 대조한다** — 여기서 안 보면 그 열은 무검증이 된다.
    """
    if not srcs or not freq_labels:
        return
    # 통계표에서 그 열의 (행 라벨 → 빈도) 표를 만든다.
    want: dict[str, dict[str, str]] = {}
    for t in srcs:
        for bn in t.get("banners") or []:
            if str(bn.get("label") or "").strip() not in freq_labels:
                continue
            n = bn.get("n") or 0
            for r in t.get("rows") or []:
                if r.get("kind") == "mean":
                    continue                       # 평균은 백분율이 아니라 그대로 둔다
                vals = r.get("values") or []
                cell = vals[bn["pos"]] if bn["pos"] < len(vals) else None
                v = cell.get("value") if isinstance(cell, dict) else None
                if v is None:
                    continue
                key = str(r.get("label") or "").strip()
                want.setdefault(key, {})[str(bn["label"]).strip()] = str(int(round(v * n / 100.0)))

    labels = [str(x or "").strip() for x in (block.get("banner_labels") or [])]
    for row in block["rows"]:
        if row["label"] == "Base for %":
            continue
        key = str(row["label"]).strip()
        if key not in want:
            continue
        for ci, cell in enumerate(row["cells"]):
            if ci >= len(labels) or labels[ci] not in freq_labels:
                continue
            got = str(cell).strip() if cell is not None else ""
            if got in ("-", ""):
                continue
            exp = want[key].get(labels[ci])
            if exp is not None and got != exp:
                rep.add("V2", "치명", f"본문!{row['row']}행",
                        f"빈도 열 `{labels[ci]}` 의 값이 통계표 역산과 다릅니다: "
                        f"{got} (기대 {exp}, 행 {key[:30]!r})")


def check_sentinels(rep: Report, rules: dict, block: dict) -> None:
    """V-010/V-011 — 센티널 라벨 표기와 보기 라벨 들여쓰기.

    괄호 종류가 섞이면(【】↔[]) 제작팀 RPA 가 그 행을 못 알아본다.
    """
    spec = find_check(rules, "sentinels", "V-010").get("exact_strings") or {}
    if not spec:
        return
    literals = {v["literal"] for v in spec.values() if isinstance(v, dict) and "literal" in v}
    # 라벨이 센티널 '비슷한' 모양인데 정확히 일치하지 않으면 표기 오류다
    near = re.compile(r"^\s*[\[【](TOP|BOT|SoSo|\d+점 평균|NPS)")
    # 예외 — `[5점 평균] {항목명}` 합성 라벨. 짝 Mean 표를 top2 패널에 얹을 때 빌더가
    # 항목명을 붙여 만드는 정규 표기다(build_guide 짝 Mean 접합 · R-145, r11 실물 307행~).
    # 괄호가 【】 로 어긋난 합성은 예외에 안 걸리므로 여전히 잡힌다.
    paired_mean = re.compile(r"^\s*\[\d+점 평균\]\s+\S")
    for row in block["rows"]:
        label = row["label"]
        if paired_mean.match(label):
            continue
        if near.match(label) and label not in literals:
            rep.add("V9", "치명", f"본문!{row['row']}행",
                    f"센티널 라벨 표기가 규격과 다릅니다: {label!r} "
                    f"(허용: {', '.join(sorted(literals))})")
        if label.strip() == "계" and label != "      계":
            rep.add("V9", "경고", f"본문!{row['row']}행",
                    f"`계` 행은 앞 공백 6칸이어야 합니다: {label!r}")


def check_base_format(rep: Report, rules: dict, block: dict, where: str) -> None:
    """V-030 — Base/Unit 문자열 형식. 예시 제작본 s6 의 실제 오류를 잡는 검사다."""
    spec = find_check(rules, "base_unit", "V-030")
    if not spec:
        return
    bu = str(block["values"].get("Base/Unit") or "")
    if not bu or bu == "-":
        return
    for f in spec.get("forbid_patterns") or []:
        if re.search(f["pattern"], bu):
            rep.add("V10", "경고", where,
                    f"Base/Unit 표기 오류 ({f['reason']}): {bu!r}")


def check_ment_style(rep: Report, rules: dict, block: dict, where: str) -> None:
    """V-060/V-061 — 개조식 어미와 **수준어 사다리 일치**.

    수준어 검사가 핵심이다. 멘트가 인용한 TOP2 값과 형용사가 어긋나면
    (예: 46.8% 인데 '긍정적인 수준임') 독자가 숫자를 잘못 읽게 된다.
    SKEE 멘트 코퍼스 40여 줄 전수에서 예외 0건이라 강제해도 안전하다.
    """
    ladder = find_check(rules, "ment", "V-061").get("ladder") or []
    style = find_check(rules, "ment", "V-060")
    for line in block["ment"]:
        if style and line.rstrip().endswith("."):
            rep.add("V11", "경고", where, f"멘트는 마침표를 찍지 않습니다: {line[:60]}")
        if not ladder:
            continue
        m = re.search(r"(\d+\.\d)\s*%\s*\(TOP\d 기준\)", line)
        if not m:
            continue
        val = float(m.group(1))
        want = None
        for rung in ladder:
            lo = rung.get("min", float("-inf"))
            hi = rung.get("max_exclusive", float("inf"))
            if lo <= val < hi:
                want = rung
                break
        if want and want["phrase"] not in line and want["adjective"] not in line:
            rep.add("V11", "치명", where,
                    f"수준어가 값과 어긋납니다 — {val}% 는 '{want['phrase']}' 여야 합니다\n"
                    f"      멘트: {line[:80]}")


#: 목차 항목명에서 두 개념을 잇는 표시. 이게 보이면 MECE 가 깨진 것이므로
#: 문구를 다듬는 게 아니라 항목을 쪼개야 한다 (RULES R-028/R-029).
TOC_CONJUNCTIONS = (" 및 ", "및 ", " 및", "/", "·")


def normalize_qu(text: str, piping_map: dict | None = None) -> str:
    """설문지 원문 → 보고서 `Qu:` 표기로 정규화.

    빌더(`build_facts.compose_qu`)와 **같은 함수를 쓴다.** 정규화 규칙이 두 곳에
    따로 있으면 한쪽만 바뀌었을 때 검수가 조용히 무력해진다.
    """
    try:
        import build_facts as BF
        out = BF.RESPONSE_MARKER.sub(" ", text or "").strip()
        out = re.sub(r"\s{2,}", " ", out)
        out, _ = BF.resolve_piping(out, piping_map or {})
        return out
    except Exception:
        return (text or "").strip()


def check_toc_naming(rep: Report, rules: dict, blocks: list[dict]) -> None:
    """R-028/R-029 — 목차 항목명은 한 어절. `및`·`/` 로 두 개념을 잇지 않는다.

    `및` 을 쓰고 싶어졌다는 건 한 꼭지에 두 개념을 욱여넣었다는 신호다.
    """
    spec = find_check(rules, "toc", "V-101")
    marks = spec.get("forbid_tokens") if spec else None
    marks = marks or TOC_CONJUNCTIONS
    # `Part Ⅱ. 요약 및 결론` 은 고정 Part 명이라 한 어절 규칙 대상이 아니다 (guide/10 §1).
    exempt = tuple((spec or {}).get("exempt_prefixes") or ("Part ", "PART ", "파트 "))
    # 고정 섹션명(`8. 결론 및 제언`)도 같은 이유로 면제다 — 번호 접두를 떼고 대조한다.
    exempt_names = tuple((spec or {}).get("exempt_names") or ())
    #: 한 어절 규칙은 **Part Ⅲ 조사 결과 목차**의 작명 규칙이다 (guide/10 §2-3②).
    #  Part Ⅱ 요약 꼭지 제목은 대상이 아니다 — 예시 실측이 그렇다:
    #  LG `2. 시나리오 제시에 따른 인식 변화` · `3. 폼팩터별 시장 가능성 & 수용성 비교`.
    #  요약은 표 2~4개를 한 장에 합친 것이라 무엇을 합쳤는지가 제목에 드러나야 하고,
    #  한 어절로 줄이면 「시장? 무슨 시장」이 된다. 구조 블록 전체를 면제한다.
    skip_struct = True if spec is None else bool(spec.get("exempt_struct", True))
    for b in blocks:
        if skip_struct and b.get("is_struct"):
            continue
        for label in ("Slide title", "Sub title"):
            v = b["values"].get(label)
            if not isinstance(v, str) or not v.strip() or v.strip() == "-":
                continue
            if v.strip().startswith(exempt):
                continue
            if re.sub(r"^\s*\d+[.)]\s*", "", v.strip()) in exempt_names:
                continue
            hit = next((m for m in marks if m in v), None)
            if hit:
                rep.add("V13", "경고", block_where(b),
                        f"`{label}` 이 한 어절이 아닙니다 ({hit.strip()!r} 로 두 개념을 이었습니다): {v!r}\n"
                        f"      → 문구를 다듬지 말고 항목을 쪼개세요 (MECE 위반 신호)")
            # R-132 — 명사형 종결. 문항 원문(의문문)을 그대로 옮긴 항목명을 잡는다.
            #
            # 끝 **글자**만 보면 명사를 서술어미로 오인한다 — 「수요」·「개요」의 `요`,
            # 「의향」에는 없지만 「인지」의 `지` 같은 자리다. 그래서 마지막 **어절**이
            # 규칙이 인정한 명사면 통과시킨다 (`exempt_ending_words`, 사용자 지시 2026-08-27).
            ends = tuple((spec or {}).get("forbid_endings") or ())
            exempt_words = tuple((spec or {}).get("exempt_ending_words") or ())
            tail = re.sub(r"^\s*\d+[.)]\s*", "", v.strip()).rstrip(")").split()[-1:] or [""]
            if ends and v.strip().rstrip(")").endswith(ends) and tail[0] not in exempt_words:
                rep.add("V13", "경고", block_where(b),
                        f"`{label}` 이 명사형으로 끝나지 않습니다 (R-132): {v!r}")


def check_toc_coverage(rep: Report, rules: dict, blocks: list[dict],
                       questions: dict[str, dict],
                       exclusions: "list[str] | None" = None) -> None:
    """R-027/R-029 — 전 문항이 한 번씩 배정됐는가 (전수성 + 상호배타).

    문항코드는 `Qu:` 문자열이 아니라 팩트에서 가져온다 — 가이드의 `Qu:` 는
    설문지 문항 **원문**만 담고 코드는 넣지 않는다(예시 파일도 동일).
    팩트는 `join_facts` 가 순번으로 붙여 둔 `block["slide"]` 에서 읽는다.
    """
    if not questions or not any(b.get("slide") for b in blocks):
        return

    # ── 상호배타: **같은 표**가 두 블록에 들어가면 중복이다.
    #    문항코드로 세면 안 된다 — 매트릭스 척도(A8 7항목·D6 5항목·C2 7항목)나
    #    다항목 숫자입력(A1 지상/지하/엘리베이터/총층수)은 한 문항이 여러 표를 갖는 게 정상이다.
    by_table: dict[tuple, list[int]] = {}
    covered: set[str] = set()
    for b in blocks:
        slide = b.get("slide") or {}
        if slide.get("question_code"):
            covered.add(slide["question_code"])
        key = tuple(slide.get("source_tables") or [])
        if key:
            by_table.setdefault(key, []).append(b["row"])

    for key, rows in by_table.items():
        if len(rows) > 1:
            rep.add("V13", "경고", f"본문!{rows[0]}행",
                    f"통계표 {list(key)} 가 {len(rows)}개 블록에 중복 배정됐습니다 "
                    f"(행 {rows}) — 한 표는 한 꼭지에만 들어갑니다")
    seen = {c: [] for c in covered}

    # SQ 와 개방형은 보고서에서 빼는 게 정상이므로 누락으로 세지 않는다.
    # 목차에 **제외 사유를 적어 둔 문항**은 누락이 아니다 (R-027 — 검산식은
    # `설문지 문항 수 = 배정 + 명시적 제외` 이고, 조용히 빠뜨린 것만 잡는다).
    excluded = {c.upper() for c in (exclusions or [])}
    missing = [c for c in questions
               if not c.upper().startswith(("SQ", "DQ", "ED"))
               and c not in seen
               and c.upper() not in excluded
               and (questions[c].get("response_type") or "") != "개방형"]
    if missing:
        rep.add("V13", "참고", "목차 전수성",
                f"제작가이드에 블록이 없는 문항 {len(missing)}개: {', '.join(missing[:20])}"
                f"{' …' if len(missing) > 20 else ''}\n"
                f"      → 의도적으로 뺀 것이면 `part 3 목차` 에 제외 사유를 남기세요")


def check_send_copy(rep: Report, rules: dict, wb, guide: Path) -> None:
    """V-081 — 송부본에 `_` 접두 시트(우리 작업용)가 남아 있으면 안 된다."""
    if "송부본" not in guide.stem:
        return
    leftover = [n for n in wb.sheetnames if n.startswith("_")]
    if leftover:
        rep.add("V12", "치명", guide.name,
                f"송부본에 작업용 시트가 남아 있습니다: {', '.join(leftover)}")


def load_part2_items(toc: Path | None) -> list[str]:
    """`목차.md` 확정 목차의 Part Ⅱ 항목명 (`1. 시장` …). 파서는 빌더 것을 그대로 쓴다."""
    if not toc or not toc.exists():
        return []
    try:
        import build_guide as BG
        return (BG.parse_toc(toc).get("parts", {}).get("Ⅱ", {}) or {}).get("items", [])
    except Exception:
        return []


RE_PANEL_REF = re.compile(r"\b(P\d+|표\d+)\b")
RE_APPENDIX = re.compile(r"^\s*\[부록\s*\d+\]")


def check_summary_traceable(rep: "Report", ws, blocks: list[dict]) -> None:
    """V18 — 요약·결론 장표에 적힌 수치를 **조사 결과에서 찾을 수 있는가** (사용자 지시 2026-08-24).

    「요약 및 결론에 나온 숫자들은 다 조사 결과에서 찾을 수 있어야 한다」가 요구다. V17 은 요약 패널 수치를
    **근거 표**(통계표)에 대조하는데, 그것만으로는 독자가 보고서 안에서 그 숫자를 찾지 못한다 — 교차 패널이
    쓰는 절단선 밖 배너군은 본문 어디에도 안 실리기 때문이다. 그래서 대조 대상을 **본문 + 부록**으로 둔다.
    찾지 못한 수치는 ① 오기이거나 ② 그 표를 부록으로 빼지 않은 것이다. 둘 다 고쳐야 한다.

    멘트·헤드라인의 수치도 본다 — 장표에서 눈에 먼저 들어오는 것이 그 문장이다.
    """
    def numeric_cells(b: dict) -> set[str]:
        out: set[str] = set()
        for rr in range(b["row"], b["end"] + 1):
            for cc in range(COL_VALUE, ws.max_column + 1):
                v = ws.cell(rr, cc).value
                if isinstance(v, (int, float)):
                    out |= {f"{float(v):.1f}", f"{float(v):.2f}", str(v)}
                elif isinstance(v, str):
                    sv = v.strip()
                    if re.fullmatch(r"-?\d+(?:\.\d+)?\*?", sv):
                        out.add(sv.rstrip("*"))
                    else:
                        m = re.fullmatch(r"\((\d[\d,]*)\*?\)", sv)
                        if m:
                            out.add(m.group(1).replace(",", ""))
        return out

    body_pool: set[str] = set()
    part2: list[dict] = []
    n_apx = 0
    for b in blocks:
        title = str(b["values"].get("Slide title") or "")
        if RE_APPENDIX.match(title):
            body_pool |= numeric_cells(b)
            n_apx += 1
            continue
        if not b.get("is_struct"):                     # 본문(조사 결과)
            body_pool |= numeric_cells(b)
            continue
        ref = str(b["values"].get("Question/Reference") or "")
        if ref.startswith("* 근거"):                    # 요약 꼭지·결론 박스
            part2.append(b)
    if not part2 or not body_pool:
        return

    n_checked = 0
    for b in part2:
        where = f"본문!{b['row']}행 {str(b['values'].get('Slide title') or '')[:22]!r}"
        # ① 표 칸의 수치
        for rr in range(b["row"] + len(LABELS), b["end"] + 1):
            for cc in range(COL_VALUE, ws.max_column + 1):
                v = ws.cell(rr, cc).value
                cand: set[str] = set()
                if isinstance(v, (int, float)):
                    cand = {f"{float(v):.1f}", f"{float(v):.2f}", str(v)}
                elif isinstance(v, str):
                    sv = v.strip()
                    if re.fullmatch(r"-?\d+(?:\.\d+)?\*?", sv):
                        cand = {sv.rstrip("*")}
                    else:
                        m = re.fullmatch(r"\((\d[\d,]*)\*?\)", sv)
                        if m:
                            cand = {m.group(1).replace(",", "")}
                if not cand:
                    continue
                n_checked += 1
                if not (cand & body_pool):
                    rep.add("V18", "치명", f"{where} {rr}행 {cc}열",
                            f"요약·결론 수치 {sorted(cand)[0]} 를 조사 결과(본문·부록)에서 찾을 수 없습니다 "
                            f"— 오기이거나 그 표를 부록으로 빼지 않은 것입니다")
        # ② 멘트·헤드라인의 수치
        labels = [r["label"] for r in b["rows"]] + list(b.get("banner_labels") or []) + list(b.get("banner_groups") or [])
        for line in b["ment"]:
            for tok in numbers_in(line, labels):
                n_checked += 1
                if tok not in body_pool:
                    rep.add("V18", "치명", where,
                            f"요약·결론 문장의 수치 {tok} 를 조사 결과에서 찾을 수 없습니다\n      문장: {line[:80]}")
    rep.stats["traceable_numbers_checked"] = n_checked
    rep.stats["appendix_blocks"] = n_apx



def check_summary_panels(rep: "Report", ws, blocks: list[dict], by_index: dict[int, dict],
                         fact_slides: list[dict]) -> None:
    """V17 — 요약 블록(Part Ⅱ)의 패널 수치가 `* 근거: P27 · 표31` 에 적힌 표들의 값 집합 안에 있는가.

    요약 블록은 구조 블록이라 V2(본문 수치 ⊂ 통계표)가 건너뛴다. 전치 교차표(R-145)까지 포함해
    블록 구간의 **모든 수치 셀**을 근거 표(슬라이드 grid + 통계표 display + 짝 Mean 표)의 값 집합에 대조한다.
    괄호 Base `(300)` 은 Base 집합에 따로 대조한다. 문항코드·연도 같은 라벨 속 숫자는 문자열 셀이라 대상이 아니다.
    """
    by_page = {s["page"]: s for s in fact_slides if s.get("page")}

    def table_values(t: dict) -> tuple[set[str], set[str]]:
        vals, bases = set(), set()
        for r in t.get("rows") or []:
            for v in r.get("values") or []:
                if isinstance(v, dict) and v.get("display"):
                    vals.add(str(v["display"]))
        for b in t.get("banners") or []:
            if b.get("n") is not None:
                bases.add(str(b["n"]))
        if t.get("base_n") is not None:
            bases.add(str(t["base_n"]))
        return vals, bases

    def grid_values(g: dict) -> tuple[set[str], set[str]]:
        vals, bases = set(), set()
        for r in g.get("rows") or []:
            for c in r.get("cells") or []:
                if c is not None and str(c).strip() not in ("", "-"):
                    vals.add(str(c).replace("*", "").strip())
            if r.get("rank_first_total") is not None:
                vals.add(str(r["rank_first_total"]))
        for b in (g.get("base_row") or [])[1:]:
            if b is not None:
                bases.add(re.sub(r"[()*\s]", "", str(b)))
        return vals, bases

    n_checked = 0
    for b in blocks:
        if not b.get("is_struct"):
            continue
        ref_line = str(b["values"].get("Question/Reference") or "")
        if not ref_line.startswith("* 근거"):
            continue
        refs = RE_PANEL_REF.findall(ref_line)
        if not refs:                                  # 결론 블록(`* 근거: Part Ⅱ 요약 꼭지`)은 표가 아니라 요약을 근거로 한다
            continue
        allowed, bases = set(), set()
        for ref in refs:
            if ref.startswith("P"):
                sl = by_page.get(int(ref[1:]))
                if sl:
                    v, bs = grid_values(sl.get("grid") or {})
                    allowed |= v; bases |= bs
                    # 본문 슬라이드의 원표(절단선 밖 배너군까지) — 요약 패널이 extra_groups 로 가져올 수 있다
                    for t in by_index.values():
                        if t.get("question_code") == sl.get("question_code") and \
                                (t.get("title_full") or "").strip() == (sl.get("title_full") or "").strip():
                            v, bs = table_values(t); allowed |= v; bases |= bs
            else:
                idx = int(ref[1:])
                for k in (idx, idx + 1):              # 짝 Mean Summary 표([5점 평균] 행)
                    t = by_index.get(k)
                    if t and (k == idx or re.search(r"Mean|평균", str(t.get("title_full") or ""))):
                        v, bs = table_values(t); allowed |= v; bases |= bs
        if not allowed:
            rep.add("V17", "경고", block_where(b), f"요약 블록 근거 표를 찾지 못했습니다: {ref_line[:60]}")
            continue
        where = f"본문!{b['row']}행 {str(b['values'].get('Slide title') or '')[:20]!r}"
        start, end = b["row"] + len(LABELS), b["end"]
        for rr in range(start, end + 1):
            for cc in range(COL_VALUE, ws.max_column + 1):
                v = ws.cell(rr, cc).value
                if v is None:
                    continue
                if isinstance(v, (int, float)):
                    txt = f"{v:.2f}".rstrip("0").rstrip(".") if isinstance(v, float) else str(v)
                    cands = {txt, f"{v:.1f}", f"{v:.2f}"} if isinstance(v, (int, float)) else {txt}
                    n_checked += 1
                    if not (cands & allowed):
                        rep.add("V17", "치명", f"{where} {rr}행 {cc}열", f"요약 패널 수치 {txt} 가 근거 표({' · '.join(refs)})에 없습니다")
                    continue
                sv = str(v).strip()
                m = re.fullmatch(r"\((\d[\d,]*)\)\*?", sv)
                if m:
                    n_checked += 1
                    if m.group(1).replace(",", "") not in bases:
                        rep.add("V17", "치명", f"{where} {rr}행 {cc}열", f"요약 패널 Base {sv} 가 근거 표의 집단 n 에 없습니다")
                    continue
                if re.fullmatch(r"-?\d+(?:\.\d+)?\*?", sv):
                    n_checked += 1
                    if sv.rstrip("*") not in allowed:
                        rep.add("V17", "치명", f"{where} {rr}행 {cc}열", f"요약 패널 수치 {sv} 가 근거 표({' · '.join(refs)})에 없습니다")
    rep.stats["summary_panel_numbers_checked"] = n_checked


def _collect_values(obj, out: set[str]) -> None:
    """중첩 구조에서 수치로 읽히는 값을 전부 긁어 문자열 집합에 넣는다."""
    if isinstance(obj, dict):
        for v in obj.values():
            _collect_values(v, out)
    elif isinstance(obj, list):
        for v in obj:
            _collect_values(v, out)
    elif isinstance(obj, bool):
        return
    elif isinstance(obj, (int, float)):
        out.add(str(obj))
        out.add(f"{float(obj):.1f}")
        out.add(f"{float(obj):.2f}")
        out.add(str(round(float(obj))))
    elif isinstance(obj, str):
        t = obj.strip()
        if re.fullmatch(r"-?\d+(\.\d+)?", t):
            out.add(t)


def verify(guide: Path, tables: Path, facts: Path | None,
           questionnaire: Path | None, piping: Path | None = None,
           toc: Path | None = None, extra_values: Path | None = None) -> Report:
    rep = Report()
    #: 통계표 밖에서 온 값의 출처 (전년 보고서 차트 등). V2 는 이 집합도 정본으로 인정한다 —
    #: 값을 지어낸 것이 아니라 **다른 정본 파일**에서 옮겨 실었다는 뜻이다.
    external_values: set[str] = set()
    if extra_values and Path(extra_values).exists():
        with io.open(extra_values, encoding="utf-8") as fh:
            _collect_values(json.load(fh), external_values)
    rules = load_rules()
    if not rules:
        rep.add("V0", "경고", "engine/learning/rules.yaml",
                "규칙 파일을 읽지 못해 기본 검사만 수행했습니다 (PyYAML 설치 확인)")

    with io.open(tables, encoding="utf-8") as fh:
        tdoc = json.load(fh)
    tlist = tdoc["tables"] if isinstance(tdoc, dict) else tdoc
    by_index = {t["index"]: t for t in tlist}

    fdoc = None
    if facts and facts.exists():
        with io.open(facts, encoding="utf-8") as fh:
            fdoc = json.load(fh)
    fact_slides = (fdoc or {}).get("slides", [])

    # 빌더가 쓴 파이핑 치환 맵. Qu 대조에 **같은 규칙**을 적용해야 오탐이 안 난다.
    piping_map: dict = {}
    if piping and piping.exists():
        with io.open(piping, encoding="utf-8") as fh:
            piping_map = json.load(fh)

    questions: dict[str, dict] = {}
    if questionnaire and questionnaire.exists():
        with io.open(questionnaire, encoding="utf-8") as fh:
            questions = {q["code"]: q for q in json.load(fh).get("questions", [])
                         if q.get("code")}

    # 확정 목차의 Part Ⅱ 항목명. 요약·결론 블록을 **의미로** 알아보기 위한 것이다
    # (`mark_struct_blocks` 참조). 팩트 옆에 있는 목차.md 를 조용히 찾아 쓴다 —
    # 없으면 Part Ⅲ 간지 경계만으로 가른다.
    part2_items = load_part2_items(toc or (facts.parent / "목차.md" if facts else None))

    # ── V8 템플릿 무결성
    with zipfile.ZipFile(guide) as z:
        if "xl/vbaProject.bin" not in z.namelist():
            rep.add("V8", "치명", guide.name,
                    "vbaProject.bin 이 없습니다 — lovelySeoyeon 매크로와 검수 실행 버튼이 사라졌습니다")
        # ── MCE 접두 in-scope — `<mc:Choice Requires="x14">` 를 쓰는 시트 루트에 `xmlns:x14` 가 없으면
        #    XML 은 well-formed 라 openpyxl 은 읽지만 **Excel 은 파일을 거부한다**(0x800A03EC).
        #    guide-adversary ADV-01 실측 — 기계 검사가 「Excel 이 여는가」를 묻지 않아 두 번 놓친 결함.
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8") if "xl/_rels/workbook.xml.rels" in z.namelist() else ""
        for name in z.namelist():
            if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                continue
            xml = z.read(name).decode("utf-8", "replace")
            root = re.search(r"<worksheet[^>]*>", xml)
            for prefix in set(re.findall(r'Requires="([^"]+)"', xml)):
                for pfx in prefix.split():
                    if root and f"xmlns:{pfx}=" not in root.group(0):
                        rep.add("V8", "치명", name,
                                f"MCE `Requires=\"{pfx}\"` 를 쓰는데 루트에 xmlns:{pfx} 선언이 없습니다 — Excel 이 파일을 열지 못합니다")
            if name.split("xl/")[1] not in rels and name.split("/")[-1] not in rels:
                rep.add("V8", "경고", name, "workbook 이 참조하지 않는 고아 워크시트 파트입니다 (탭에는 안 보이나 내용이 회수됩니다)")
    wb = openpyxl.load_workbook(guide, keep_vba=True)
    protected = (find_check(rules, "pre_submit", "V-082").get("required_sheets")
                 or FALLBACK_PROTECTED)
    for name in protected:
        if name not in wb.sheetnames:
            rep.add("V8", "치명", name, "사내 RPA 가 참조하는 시트가 삭제되었습니다 (수정금지 시트)")
    body_name = next((n for n in BODY_SHEETS if n in wb.sheetnames), None)
    if body_name is None:
        rep.add("V8", "치명", guide.name, "본문 시트가 없습니다")
        return rep

    check_send_copy(rep, rules, wb, guide)

    markers = (find_check(rules, "unfilled_markers", "V-070").get("search_strings")
               or FALLBACK_MARKERS)
    marker_pat = re.compile("|".join(re.escape(m) for m in markers))

    blocks = read_blocks(wb[body_name], part2_items)
    rep.stats["blocks"] = len(blocks)
    rep.stats["rules_loaded"] = bool(rules)
    rep.stats["tables"] = len(tlist)

    # A열 seq 가 없으므로 **등장 순서**로 팩트를 붙인다. 개수가 어긋나면 여기서 치명.
    if fact_slides:
        join_facts(rep, blocks, fact_slides)
    rep.stats["struct_blocks"] = sum(1 for b in blocks if b["is_struct"])

    check_toc_naming(rep, rules, blocks)
    check_toc_coverage(rep, rules, blocks, questions,
                       _toc_exclusions(toc) if toc else None)

    # ── 좌표 단위 대조 (V14·V15·V16) — **아직 캘리브레이션 전이라 기본 꺼짐.**
    #    V2 의 집합 소속 검사가 못 잡는 자리를 덮는 것이 목적이다
    #    (같은 표 안 값 맞바꿈 · 배너 라벨 어긋남 · 차트타입 오기).
    #    2026-08-22 현재 정상본에도 1,003건이 뜬다 — 행 라벨 매칭이 매트릭스 표에서
    #    엉뚱한 행에 붙는 것으로 보인다(표18 처럼 한 문항이 여러 표로 갈리는 경우).
    #    **오탐을 0 으로 맞추기 전에는 켜지 않는다.** `--coord` 로만 켠다.
    if os.environ.get("VERIFY_COORD") == "1":
        from verify_coord import check_grid_coordinates, check_chart_type
        check_grid_coordinates(rep, blocks, tlist, block_where)
        check_chart_type(rep, blocks, block_where)

    # 한 장(page)에 실린 값 전체 — 한 장이 패널 여러 개로 쪼개지는 덱에서 V3 의 대조
    # 집합이 된다. 팩트에 `page` 가 없거나 한 장에 블록이 하나뿐이면 블록 집합과 같다.
    page_numbers: dict[Any, set[str]] = {}
    for b in blocks:
        pg = (b.get("slide") or {}).get("page")
        if pg is None:
            continue
        acc = page_numbers.setdefault(pg, set())
        for row in b["rows"]:
            for cell in row["cells"]:
                if cell is None:
                    continue
                s = str(cell).strip()
                if s and s not in ("-",):
                    acc.add(s)

    n_ment_numbers = 0
    for b in blocks:
        where = block_where(b)
        slide = b.get("slide")

        # ── V1 블록 무결성
        if b["labels"] != LABELS:
            rep.add("V1", "치명", where,
                    f"라벨이 규격과 다릅니다: {b['labels']}")

        # ── V7 미기입
        #    구조 블록(표지·목차·간지·조사개요·요약·결론)은 Base/Unit·Qu·Sub title 이
        #    **비어 있는 게 정상**이다 (LG `본문_제작용` 실측 — 그 칸들이 빈칸이다).
        #    빌더가 빈칸을 `-` 로 메우지 않으므로 여기서 빼지 않으면 전부 오탐이 된다.
        #    미기입 표시(`???` 등)는 구조 블록에서도 그대로 잡는다.
        for label in ("Chart type", "Base/Unit", "Question/Reference"):
            v = b["values"].get(label)
            if not v or (isinstance(v, str) and not v.strip()):
                if not b["is_struct"]:
                    rep.add("V7", "치명", where, f"`{label}` 이 비어 있습니다")
            elif isinstance(v, str) and marker_pat.search(v):
                rep.add("V7", "치명", where, f"`{label}` 에 미기입 표시가 남아 있습니다: {v!r}")
        for label in ("Slide title", "Sub title"):
            v = b["values"].get(label)
            if isinstance(v, str) and marker_pat.search(v):
                rep.add("V7", "치명", where, f"`{label}` 에 미기입 표시가 남아 있습니다: {v!r}")
            elif not b["is_struct"] and not (v or "").strip():
                rep.add("V7", "경고", where, f"`{label}` 이 비어 있습니다 (목차 확정 필요)")

        check_sentinels(rep, rules, b)
        check_base_format(rep, rules, b, where)
        check_ment_style(rep, rules, b, where)

        src_indices = (slide or {}).get("source_tables") or []
        srcs = [by_index[i] for i in src_indices if i in by_index]
        src = srcs[0] if srcs else None
        if not b["is_struct"] and not srcs:
            # 대조 기준이 없는 본문 블록. 여기서 V2·V3 을 그냥 돌리면 그리드가 비어
            # 있으니 모든 수치가 "통계표에 없는 값"으로 잡혀 진짜 원인이 파묻힌다.
            rep.add("V2", "경고", where,
                    "이 블록에 대응하는 통계표를 찾지 못해 수치·멘트 대조를 건너뜁니다"
                    + ("" if slide else " (팩트 조인 실패 — 위 V1 참조)"))

        # ── V4 Base 표기
        bu = str(b["values"].get("Base/Unit") or "")
        m = re.search(r"n=([\d,]+)", bu)
        if m and src:
            n = int(m.group(1).replace(",", ""))
            if n != src.get("base_n"):
                rep.add("V4", "치명", where,
                        f"Base n={n} 인데 통계표 표{src['index']} 는 {src.get('base_n')} 입니다")

        # ── V5 Qu 텍스트
        # 설문지 **원문 그대로**가 아니라 **보고서 표기로 정규화한 것**과 비교한다.
        # 응답 형식 마커(`[직접 입력]`)와 파이핑 자리표시자는 의도적으로 걷어내므로
        # 원문과 곧이곧대로 대조하면 정상 슬라이드가 전부 경고로 뜬다.
        qu = str(b["values"].get("Question/Reference") or "")
        if src and questions:
            q = questions.get(src.get("question_code"))
            if q and q.get("text"):
                core = normalize_qu(q["text"], piping_map)
                if core and core not in qu:
                    rep.add("V5", "경고", where,
                            f"Qu 가 설문지와 다릅니다\n      가이드: {qu[:70]}\n      설문지(정규화): {core[:70]}")

        # ── V2 그리드 수치 · V6 신호어
        grid_numbers: set[str] = set()
        if srcs:
            # 순위 슬라이드는 1순위 표(전체 열)와 1~N순위 표가 함께 들어간다.
            src_values: set[str] = set()
            for t in srcs:
                for r in t.get("rows") or []:
                    for v in r.get("values") or []:
                        if isinstance(v, dict) and v.get("display"):
                            src_values.add(str(v["display"]))
            # 백분율 대신 **빈도**를 적기로 한 열(사람 확정 `--freq-banners`)은 통계표
            # 표시값과 곧이곧대로 대조하면 전량 「없는 값」으로 잡힌다. 그 열의 셀은
            # `표시값 × n ÷ 100` 이므로 역산해 따로 검산한다 — 눈감아 주는 것이 아니다.
            freq_labels = set(((b.get("slide") or {}).get("grid") or {}).get("freq_banners") or [])
            freq_cols: set[int] = set()
            if freq_labels:
                freq_cols = {i for i, lab in enumerate(b.get("banner_labels") or [])
                             if str(lab or "").strip() in freq_labels}
                check_freq_cells(rep, b, srcs, freq_labels)
            for row in b["rows"]:
                if row["label"] == "Base for %":
                    continue
                for ci, cell in enumerate(row["cells"]):
                    if ci in freq_cols:
                        continue                      # 빈도 열 — 아래 _check_freq_cells 가 따로 검산한다
                    if cell is None:
                        continue                      # 값 없음 — 빈 칸으로 둔 자리
                    s = str(cell).strip()
                    if s in ("-", ""):
                        continue
                    grid_numbers.add(s)
                    if re.fullmatch(r"-?\d+(\.\d+)?", s) and s not in src_values                             and s not in external_values:
                        srcs_txt = "·".join(f"표{t['index']}" for t in srcs)
                        rep.add("V2", "치명", f"본문!{row['row']}행",
                                f"통계표 {srcs_txt} 에 없는 값입니다: {s} "
                                f"(행 {row['label'][:30]!r})")
                if re.search(r"【TOP\d?】|\[\d+점 평균\]|\[NPS\]", row["label"]):
                    if row["modifier"] != "그래프 상단 Text":
                        rep.add("V6", "치명", f"본문!{row['row']}행",
                                f"`{row['label']}` 행에 `그래프 상단 Text` 가 없습니다 "
                                f"— 이 값이 그래프에 표시되지 않습니다")

        # ── V3 멘트 수치  ★환각 검출
        #    그리드가 비어 있으면(대조 기준 없음) 돌리지 않는다 — 바로 위에서 경고했다.
        #    한 장이 패널 여러 개로 쪼개지는 덱(팩트에 `page` 가 있고 같은 page 를 여러
        #    슬라이드가 공유)에서는 멘트 상자가 장에 하나뿐이라 인용 수치가 다른 패널에
        #    있을 수 있다. 그때는 **같은 장의 값 전체**를 대조 집합으로 삼는다.
        cmp_numbers = grid_numbers | page_numbers.get((slide or {}).get("page"), set())
        for line in (b["ment"] if srcs else []):
            v3_labels = [r["label"] for r in b["rows"]] + list(b.get("banner_labels") or []) + list(b.get("banner_groups") or [])
            for tok in numbers_in(line, v3_labels):
                n_ment_numbers += 1
                if tok not in cmp_numbers:
                    rep.add("V3", "치명", where,
                            f"멘트의 수치 {tok} 가 이 슬라이드 데이터에 없습니다\n      멘트: {line[:80]}")

    rep.stats["ment_numbers_checked"] = n_ment_numbers
    rep.stats["ment_lines"] = sum(len(b["ment"]) for b in blocks)
    kpi = tables.parent / "KPI요약표_직접생성.json"                           # 표9001·9002 (직접 생성 Summary)
    by_index_all = dict(by_index)
    if kpi.exists():
        with io.open(kpi, encoding="utf-8") as fh:
            for t in (json.load(fh).get("tables") or []):
                by_index_all.setdefault(t["index"], t)
    check_summary_panels(rep, wb[body_name], blocks, by_index_all, fact_slides)   # V17 (R-145)
    check_summary_traceable(rep, wb[body_name], blocks)                          # V18 (사용자 지시 2026-08-24)
    return rep


# ─────────────────────────────────────────────────────────────
# 멘트 수용 게이트 — 담당(guide-body-ment) 샤드 반환분을 정본에 넣기 전에 돌린다
# ─────────────────────────────────────────────────────────────

MENT_MAX_CHARS = 60     #: guide/17 문장 정책 — PPT 멘트 상자 한 줄 실측 74자, 여유 14자
SUMMARY_LINE_WRAP = 74  #: 요약·결론 줄의 물리 한계 (R-198 — 60자는 강제하지 않는다)


def _has_final_mieum(text: str) -> bool:
    """마지막 글자가 ㅁ 받침인가 (임·함·음·됨·순임 …). 한글 음절 종성 인덱스 16 = ㅁ."""
    # 꼬리 괄호 주석(`…가장 높음(1~3순위 기준)`)은 종결 판정에서 뗀다 — 개조식 관례
    t = re.sub(r"\s*\([^()]*\)\s*$", "", text.rstrip()).rstrip()
    if not t:
        return False
    ch = t[-1]
    code = ord(ch) - 0xAC00
    return 0 <= code < 11172 and code % 28 == 16


def ment_gate(line: str, slide: dict, rules: dict) -> list[str]:
    """한 줄의 멘트를 한 장의 팩트 슬라이스에 대조한다. 반려 사유 목록을 돌려준다 (빈 목록 = 수용)."""
    reasons: list[str] = []
    text = line.strip()
    if not text:
        return ["빈 줄"]
    body = text[2:] if text.startswith("- ") else text
    # ① 길이
    if len(body) > MENT_MAX_CHARS:
        reasons.append(f"{len(body)}자 > {MENT_MAX_CHARS}자")
    # ② 종결
    if body.endswith("."):
        reasons.append("마침표")
    if not _has_final_mieum(body):
        reasons.append(f"ㅁ받침 종결 아님: …{body[-6:]!r}")
    # ③ 따옴표 짝
    if body.count("'") % 2:
        reasons.append("작은따옴표 짝 불일치")
    # ④ 보기명 원문 — 따옴표 인용은 행 라벨 / 배너 라벨 / 배너군 과 글자 단위로 같아야 한다
    grid = slide.get("grid") or {}
    labels: set[str] = set()
    for r in grid.get("rows") or []:
        for k in ("label", "label_clean"):
            v = r.get(k)
            if isinstance(v, str) and v.strip():
                labels.add(v.strip())
    for k in ("banner_label_row", "banner_group_row"):
        for v in grid.get(k) or []:
            if isinstance(v, str) and v.strip():
                labels.add(v.strip())
    for q in QUOTED.findall(body):
        inner = q.strip("'").strip()
        if inner and inner not in labels:
            reasons.append(f"보기명 원문 불일치: {inner!r}")
    # ④-b 배너 인용은 룰엔진 후보(ment_facts.line2) 안에서만 — R-065 선별을 담당이 우회하지 못하게
    mf = slide.get("ment_facts") or {}
    cand = {d.get("label") for d in (mf.get("line2") or [])}
    banner_labels = [v.strip() for v in (grid.get("banner_label_row") or []) if isinstance(v, str) and v.strip() and v.strip() != "전체"]
    for lab in sorted(set(banner_labels), key=len, reverse=True):
        if lab in body and lab not in cand and "별로" in body:
            reasons.append(f"배너 '{lab}' 인용 — 룰엔진 유의차 후보 밖 (R-065)")
    # ④-c R-139: 수량형 보기는 단위를 붙여 따옴표 없이 쓴다 — `'5'` 가 아니라 `5층`
    unit = mf.get("unit")
    if unit:
        for q in QUOTED.findall(body):
            if re.fullmatch(r"\d+", q.strip("'").strip()):
                reasons.append(f"R-139 수량형 보기 {q} → {q.strip(chr(39))}{unit} (단위 병기, 따옴표 없이)")
                break
        wrong = {u for u in re.findall(r"\d+(층|개|대|명|동|호|회|건)(?![\d.%])", body) if u != unit}
        if wrong:
            reasons.append(f"R-139 단위 불일치 — 표의 평균 행 단위는 '{unit}' 인데 {sorted(wrong)} 를 썼습니다")
    # ⑤ 수치 ⊆ 슬라이스 (V3 와 같은 numbers_in · 같은 그리드 집합)
    nums: set[str] = set()
    for r in grid.get("rows") or []:
        fmts = r.get("cell_formats") or []
        for i, c in enumerate(r.get("cells") or []):
            if c is None:
                continue
            cs = str(c).strip()
            if not cs or cs == "-":
                continue
            nums.add(cs)
            # 셀에는 통계표 **원값**이 들어 있고 화면에는 서식이 깎은 값이 보인다.
            # 멘트는 화면 숫자를 인용하므로 표시값도 후보에 넣는다.
            f = fmts[i] if i < len(fmts) else None
            m = re.search(r"###0(?:\.(0+))?", str(f or ""))
            if m:
                try:
                    d = len(m.group(1) or "")
                    q = Decimal("1." + "0" * d) if d else Decimal("1")
                    nums.add(f"{Decimal(cs).quantize(q, rounding=ROUND_HALF_UP):f}")
                except Exception:
                    pass
    for c in grid.get("base_row") or []:
        cs = str(c).strip("() ") if c is not None else ""
        if cs:
            nums.add(cs)
    for tok in numbers_in(body, labels):
        if tok not in nums:
            reasons.append(f"수치 {tok} 가 슬라이스에 없음")
    # ⑥ 수준어 사다리 (V-061) — check_ment_style 을 그대로 태운다
    tmp = Report()
    check_ment_style(tmp, rules, {"ment": [body]}, "ment")
    reasons.extend(f.detail.splitlines()[0] for f in tmp.findings if f.severity == "치명")
    # ⑦ 초소표본 (V-063): base_n ≤ 2 인데 % 를 썼다
    bn = slide.get("base_n")
    if isinstance(bn, int) and bn <= 2 and re.search(r"\d+\.\d%", body):
        reasons.append(f"n={bn} 초소표본인데 % 인용 (패턴 D 필요)")
    return reasons


MENT_MAX_LINES = 3      #: guide/15 §1-9 — 장당 3줄 이하 (A2 실측: 5줄 멘트는 과밀)


def slide_gate(lines: list[str], slide: dict) -> list[str]:
    """장 단위 검사 (guide/15 §1-9, A vs A2 실측 2026-08-23) — 경고 사유 목록.
    (a) 순위형 문항은 첫 줄에 「(1~N순위 기준)」 한정구  (b) 줄 수 ≤ 3
    (c) 동률(ment_facts.ties) 보기는 묶어 적는다 — 한쪽만 단독 2위처럼 쓰지 않는다  (d) 배너명에 따옴표 금지"""
    rs: list[str] = []
    body = [str(x).strip() for x in lines if str(x).strip()]
    title = str(slide.get("title_full") or "")
    mf = slide.get("ment_facts") or {}
    grid = slide.get("grid") or {}
    # (a) 순위 한정구
    m = re.search(r"(\d)\s*~\s*(\d)\s*순위", title)
    if m and body and not re.search(r"\(\s*\d\s*~\s*\d\s*순위 기준\)", body[0]):
        rs.append(f"순위형 문항인데 첫 줄에 '({m.group(1)}~{m.group(2)}순위 기준)' 한정구 없음 (guide/15 §1-9)")
    # (b) 줄 수
    if len(body) > MENT_MAX_LINES:
        rs.append(f"멘트 {len(body)}줄 > {MENT_MAX_LINES}줄 (guide/15 §1-9)")
    # (c) 동률 묶음 — ties: {값: [라벨…]} 또는 [[라벨…]] 형태 모두 허용
    ties = mf.get("ties") or {}
    groups = list(ties.values()) if isinstance(ties, dict) else list(ties)
    joined = "\n".join(body)
    for g in groups:
        labs = [str(x) for x in (g if isinstance(g, (list, tuple)) else [g])]
        if len(labs) < 2:
            continue
        hit = [l for l in labs if f"'{l}'" in joined]
        if hit and len(hit) < len(labs) and "동률" not in joined:
            rs.append(f"동률 보기 {labs} 중 {hit} 만 적음 — 묶어서 '동률임' 으로 (guide/15 §1-9)")
    # (d) 배너명 따옴표
    banners = {v.strip() for v in (grid.get("banner_label_row") or []) if isinstance(v, str) and v.strip() and v.strip() != "전체"}
    rows = {str(r.get("label") or "").strip() for r in (grid.get("rows") or [])}
    for q in QUOTED.findall(joined):
        inner = q.strip("'").strip()
        if inner in banners and inner not in rows:
            rs.append(f"배너명 '{inner}' 에 따옴표 — 배너명은 따옴표 없이 (guide/15 §1-9)")
    return rs


def check_doc_hygiene(rep: "Report", target: Path) -> None:
    """V16 문서 위생 (R-137) — 문서 속성의 개인 이름·템플릿 잔재, 본문의 보이지 않는 문자. xlsx/xlsm/pptx/docx."""
    if target.suffix.lower() not in (".xlsx", ".xlsm", ".pptx", ".docx"):
        return
    try:
        import doc_hygiene
        r = doc_hygiene.scan_file(target, "엠브레인")
    except Exception as e:                      # noqa: BLE001 — 검수가 위생 도구 때문에 죽으면 안 된다
        rep.add("V16", "참고", "문서 위생", f"스캔 실패: {e}")
        return
    for p in r["props"]:
        rep.add("V16", "경고", "문서 속성", p + "\n      → doc_hygiene.py clean --inplace (build_guide 는 자동 적용)")
    for k, v in r["invisible"].items():
        rep.add("V16", "경고", "보이지 않는 문자", f"{k} ×{v} — {', '.join(r['parts'][:3])}")


# ── 한국어 문체 예방 게이트 (guide/15 §5 · 외부 자료 im-not-ai · korean-skills 채굴 2026-08-23, K-xx) — 전부 경고
_STYLE_RULES = [
    ("K-01", re.compile(r"^(이어|뒤이어)\s|뒤를 이음$|응답도 \d+\.\d%로 (높|낮)음$"), "2·3위 줄 정형 이탈 — 「다음으로 'Y'가 NN.N%임」 / 「'Z'는 NN.N%로 그 다음임」"),
    ("K-02", re.compile(r"가장 많(음|고)"), "1위 서술어는 「가장 높음/높고」로 통일"),
    ("K-03", re.compile(r"가장 (높|많)고 '"), "「가장 높고,」 뒤 쉼표 누락"),
    ("K-06", re.compile(r"모바일앱|월평균|^외주 "), "문두 맥락어 표기 — 문항 제목 표기(모바일 앱 · 월 평균 · 외부 위탁)를 따른다"),
    ("K-11", re.compile(r"보면 상대적으로 .*(높|낮)은 편임$"), "배너 줄은 B 정형 하나 — 「…의 '보기' 응답이 상대적으로 높음」"),
    ("K-14", re.compile(r"[가-힣]+별로 (?!보면)"), "「~별로 보면」 정형 탈락"),
    ("K-12", re.compile(r"(에 대(한|해|하여)|에 있어서?|[을를] 통(해|하여)|(으)?로 인(해|한)|것으로 (나타|보|판단|예상|여겨)|되어지|지게 됨|에 의(해|한)|다양한|적극적|효과적|중요한 역할|혁신적|획기적|압도적|시사하는 바|주목할 만|결론적으로|요약하면|할 필요가 있|가지고 있)"), "번역체·AI 상투어"),
    ("K-09", re.compile(r"타겟|어플|컨셉"), "외래어 표기(타깃·앱·콘셉트) — 보기명 원문은 예외"),
    ("K-16", re.compile(r"평균\s*[\d][\d,.]*\s*[가-힣]{0,3}(?:이며|이고),?\s"), "평균은 독립 문장으로 — 한 문장에 통계량 하나 (R-193 개정 2026-08-28)"),
]


def _strip_quotes(text: str) -> str:
    """작은따옴표 안(보기명 원문)은 검사에서 뺀다."""
    return re.sub(r"'[^']*'", "'…'", text)


def style_gate(text: str, *, first_line: bool = True) -> list[str]:
    """한 줄에 대한 문체 경고 사유 목록 (반려 아님)."""
    body = _strip_quotes(text.strip())
    out = []
    for code, rx, msg in _STYLE_RULES:
        if code == "K-03" and not first_line:
            continue
        if rx.search(body):
            out.append(f"{code} {msg}")
    return out


def term_consistency(ments: list[str], texts: list[str]) -> list[str]:
    """K-05 — 요약·결론이 부르는 지표명이 본문 멘트의 측정 대상어와 다른 이름이면 경고 (예: 멘트 「이용 의향」 ↔ 요약 「도입 의향」).
    지표명 = 접미어(의향·만족도…) 바로 앞 한 단어. 조사로 끝나는 말·관형어·접속은 지표명이 아니다."""
    rx = re.compile(r"(?:^|[\s'(])([가-힣]+) (의향|만족도|이해도|신뢰도|유용성|매력도|차별성)")
    skip = {"시", "및", "의", "때", "다만", "이미", "전반적인", "관리", "컨셉", "순간", "후", "전"}
    def terms(t: str) -> set[str]:
        out = set()
        for m in rx.finditer(_strip_quotes(t)):
            w = m.group(1)
            if w in skip or w[-1] in "은는이가도의을를로와과며고나":
                continue
            out.add(f"{w} {m.group(2)}")
        return out
    ment_terms: set[str] = set()
    for t in ments:
        ment_terms |= terms(t)
    by_suffix: dict[str, set[str]] = {}
    for t in ment_terms:
        by_suffix.setdefault(t.split()[-1], set()).add(t)
    out = []
    for t in texts:
        for term in terms(t):
            suf = term.split()[-1]
            if term not in ment_terms and by_suffix.get(suf):
                out.append(f"K-05 지표명 「{term}」 — 본문 멘트는 {sorted(by_suffix[suf])} 로 부름")
    return sorted(set(out))


def unit_mix(texts: list[str]) -> list[str]:
    """K-13 — 한 문서에 「N만원」과 「N만 원」이 같이 있으면 경고."""
    joined = "\n".join(_strip_quotes(t) for t in texts)
    if re.search(r"\d만원", joined) and re.search(r"\d만 원", joined):
        return ["K-13 통화 단위 표기 혼용 — 「만원」과 「만 원」이 한 문서에"]
    return []


def verify_ments(ments: Path, facts: Path) -> Report:
    """샤드 반환 JSON(`{slides:[{id|seq, ment:[...]}]}` 또는 `{ments:[{seq, lines}]}`) 을 팩트에 대조한다.

    수용/반려를 `rep.stats['accepted']` / `['rejected']` 에 seq 별로 남긴다 —
    오케스트레이터는 accepted 만 정본에 병합하고 rejected 는 사유를 붙여 재발주한다.
    """
    rules = load_rules()
    rep = Report()
    with io.open(facts, encoding="utf-8") as fh:
        by_seq = {s["seq"]: s for s in json.load(fh).get("slides", [])}
    with io.open(ments, encoding="utf-8") as fh:
        doc = json.load(fh)
    items = doc.get("slides") or doc.get("ments") or []
    accepted: dict[int, list[str]] = {}
    rejected: dict[int, list[dict]] = {}
    for it in items:
        seq = it.get("seq", it.get("id"))
        try:
            seq = int(str(seq).lstrip("P").lstrip("seq").strip())
        except (TypeError, ValueError):
            rep.add("V11", "치명", str(it.get("id")), "seq 를 읽을 수 없습니다")
            continue
        lines = it.get("ment") or it.get("lines") or []
        slide = by_seq.get(seq)
        if slide is None:
            rep.add("V11", "치명", f"seq {seq}", "팩트에 없는 seq")
            continue
        if not lines:
            rep.add("V11", "경고", f"seq {seq}", "멘트가 비었습니다")
            rejected[seq] = [{"line": "", "reasons": ["빈 멘트"]}]
            continue
        bad = []
        for ln in lines:
            rs = ment_gate(str(ln), slide, rules)
            if rs:
                bad.append({"line": str(ln), "reasons": rs})
                rep.add("V11", "치명", f"seq {seq}", "; ".join(rs) + f"\n      멘트: {str(ln)[:80]}")
        for r in slide_gate([str(x) for x in lines], slide):     # 장 단위 — 경고(반려 아님), 담당 재발주 판단은 사람
            rep.add("V11", "경고", f"seq {seq}", r)
        for i, ln in enumerate(lines):
            for r in style_gate(str(ln), first_line=(i == 0)):
                rep.add("V11", "경고", f"seq {seq}", r + f"\n      멘트: {str(ln)[:70]}")
        if bad:
            rejected[seq] = bad
        else:
            accepted[seq] = [str(x) for x in lines]
    rep.stats.update({"ment_items": len(items), "accepted": accepted, "rejected": rejected,
                      "ment_lines": sum(len(v) for v in accepted.values())})
    return rep


# ─────────────────────────────────────────────────────────────
# 요약·결론 설계 게이트 — guide-summary 의 summary_map.json 을 정본(통계표·팩트)에 대조
# ─────────────────────────────────────────────────────────────

#: 요약 1줄(L1)의 길이 상한. 코퍼스 5사례 실측 평균 51.6~74.6자 · 최장 149자.
#: R-130(수치 없는 판단문 ≤40자)은 근거가 LG 가이드 1건뿐이고 그 문장조차 최종본에서
#: 수치 4줄로 교체됐다 — 폐기하고 물리 한계(한 줄 74자)를 상한으로 쓴다 (R-205).
HEADLINE_MAX = 74
#: 헤드라인 외 줄 상한. 코퍼스 실측 — SKEE 3줄 11장·4줄 8장 / LG 4줄 7장.
#: 4줄짜리는 마지막 줄이 **소결론(➔)** 이거나 정성층일 때 나온다. 근거 2줄 + 소결론 1줄을
#: 허용해야 요약이 「값 나열」에서 「판단이 얹힌 장」이 된다 (사용자 지적 2026-08-31).
SUMMARY_BODY_MAX = 3
#: R-126 의 「박스당 150자·장표 900자」는 코퍼스 실측과 어긋난다 — SH Key point 01 은 3줄 218자,
#: 02 는 9줄 약 650자로 2.5배 기울고 장표 합계가 900자를 넘는다. 실제 제약은 글자 수가 아니라
#: **박스 높이**(3.59cm vs 9.02cm)이고, 4사례 공통 골격은 「번호·라벨 + 제목 1줄 + 내용 3~9줄」이다(R-166).
#: 그래서 글자 수는 **참고치**로만 두고 경고에 근거를 병기하며, 실질 검사는 줄 수로 한다.
CONCLUSION_BOX_MAX = 150     #: 참고치 (R-126 원안 · 실측과 어긋남)
CONCLUSION_TOTAL_MAX = 900   #: 참고치 (R-126 원안 · 실측과 어긋남)
CONCLUSION_LINES = (3, 9)    #: R-166 확정 — 4사례 공통 골격


def _table_cell(t: dict, row_label: str, col_label: str):
    """표에서 (행 라벨, 열 라벨) 의 표시값. 행 라벨은 공백 정규화 후 대조, 열은 배너 라벨 또는 `전체`."""
    want = re.sub(r"\s+", " ", (row_label or "").strip())
    row = next((r for r in t.get("rows") or []
                if re.sub(r"\s+", " ", (r.get("label") or "").strip()) == want
                or re.sub(r"\s+", " ", (r.get("label_clean") or "").strip()) == want), None)
    if row is None:
        return None, "행 없음"
    col = (col_label or "전체").strip()
    pos = None
    for b in t.get("banners") or []:
        if (b.get("label") or "").strip() == col or (col == "전체" and b.get("group") == "전체"):
            pos = b["pos"]
            break
    if pos is None:
        return None, "열 없음"
    vals = row.get("values") or []
    if pos >= len(vals) or vals[pos] is None:
        return None, "빈 셀"
    v = vals[pos]
    return (v.get("display") if isinstance(v, dict) else str(v)), None


CONCLUSION_TITLE_MAX = 30    #: guide/15 §1-9 — 결론 제목 30자 이하
CONCLUSION_LINE_MAX = 60     #: K-10 — 결론 줄(→ 제언 포함)도 멘트처럼 60자


def _ments_from_facts(facts: Path | None) -> list[str]:
    """slide_facts.json 의 주입된 멘트(ment) 전부 — K-05 지표명 대조용."""
    if not facts or not Path(facts).exists():
        return []
    try:
        with io.open(facts, encoding="utf-8") as fh:
            return [str(l) for s in json.load(fh).get("slides", []) for l in (s.get("ment") or [])]
    except Exception:            # noqa: BLE001
        return []


def korean_style_checks(rep: "Report", check: str, where_prefix: str, texts: list[tuple[str, str]], ments: list[str],
                        *, title_max: int | None = None, line_max: int | None = None) -> None:
    """요약·결론 공통 문체 경고. texts = [(where, text)…]."""
    for where, t in texts:
        for r in style_gate(t, first_line=False):
            rep.add(check, "경고", f"{where_prefix}{where}", r + f"\n      {t[:60]}")
        if line_max and t.strip().startswith("→") and len(t.strip()) > line_max:
            rep.add(check, "경고", f"{where_prefix}{where}", f"K-10 제언 줄 {len(t.strip())}자 > {line_max}자 — 행동이 둘이면 → 줄을 나눈다")
        if title_max and where.endswith("제목") and len(t.strip()) > title_max:
            rep.add(check, "경고", f"{where_prefix}{where}", f"결론 제목 {len(t.strip())}자 > {title_max}자 (guide/15 §1-9)")
    for r in term_consistency(ments, [t for _, t in texts]):
        rep.add(check, "경고", f"{where_prefix}지표명", r)
    for r in unit_mix(ments + [t for _, t in texts]):
        rep.add(check, "경고", f"{where_prefix}단위", r)


def verify_summary(smap: Path, tables: Path, facts: Path | None) -> Report:
    rules = load_rules()
    rep = Report()
    with io.open(smap, encoding="utf-8") as fh:
        doc = json.load(fh)
    with io.open(tables, encoding="utf-8") as fh:
        tdoc = json.load(fh)
    tlist = tdoc["tables"] if isinstance(tdoc, dict) else tdoc
    by_index = {t["index"]: t for t in tlist}
    kpi = tables.parent / "KPI요약표_직접생성.json"
    if kpi.exists():
        with io.open(kpi, encoding="utf-8") as fh:
            for t in json.load(fh).get("tables", []):
                by_index[t["index"]] = t
    by_page: dict[int, dict] = {}
    if facts and facts.exists():
        with io.open(facts, encoding="utf-8") as fh:
            by_page = {x["page"]: x for x in json.load(fh).get("slides", []) if x.get("page")}

    # ── 모수 조건 지도 — 표마다 「누가 그 문항을 봤는가」 (R-185)
    total_n = max((t.get("base_n") or 0) for t in tlist) if tlist else 0
    sub_base: dict[int, dict] = {}
    for t in tlist:
        bn = t.get("base_n") or 0
        if bn and total_n and bn < total_n:
            sl = next((x for x in by_page.values() if t["index"] in (x.get("source_tables") or [])), None)
            sub_base[t["index"]] = {
                "n": bn,
                "label": (sl or {}).get("base_label") or "",
                "cond": (sl or {}).get("base_condition") or "",
            }

    def names_base(text: str, info: dict) -> bool:
        """문장이 그 표의 모수를 밝히고 있는가 — base n 을 적었거나 모수 이름의 실마리 어절을 썼는가."""
        t = str(text or "")
        if str(info["n"]) in t or f"{info['n']:,}" in t:
            return True
        lab = re.sub(r"\(.*?\)", " ", str(info.get("label") or ""))
        words = [w for w in re.split(r"[\s·,/]+", lab) if len(w) >= 2]
        return any(w in t for w in words)

    n_claims = n_ok = 0
    summaries = doc.get("summaries") or []
    for sm in summaries:
        no = sm.get("no")
        where = f"요약 {no}"
        title = (sm.get("title") or "").strip()
        if not title:
            rep.add("V14", "치명", where, "제목 없음")
        head = (sm.get("headline") or "").strip()
        hb = head[2:] if head.startswith("- ") else head
        if not hb:
            rep.add("V14", "치명", where, "헤드라인 없음 (R-130)")
        else:
            if len(hb) > HEADLINE_MAX:
                rep.add("V14", "경고", where, f"헤드라인 {len(hb)}자 > {HEADLINE_MAX}자 — 렌더 두 줄")
            # 1줄에 대표 수치가 오는 것이 코퍼스 정본이다 (5/5 사례 과반: 90·83·83·70·60%).
            # 수치가 없으면 오히려 알린다 — 판단만 있고 값이 없는 L1 은 소수파다 (R-205).
            if not re.search(r"\d", hb):
                rep.add("V14", "참고", where, f"헤드라인에 수치가 없습니다 — 코퍼스 정본은 L1 에 대표 수치를 둡니다: {hb[:50]}")
        # 패널
        panels = sm.get("panels") or []
        if not 1 <= len(panels) <= 6:
            rep.add("V14", "경고", where, f"패널 {len(panels)}개 (권장 2~5)")
        n_banner = 0
        for pn in panels:
            ref = str(pn.get("ref") or "")
            chart = str(pn.get("chart") or "")
            if chart.startswith("배너 표"):
                n_banner += 1          # 원형 배너 표(17열)만 폭 제한 대상 — 전치 교차 패널은 아니다 (R-145 ⑤)
            if ref.startswith("P"):
                sl = by_page.get(int(ref[1:])) if ref[1:].isdigit() else None
                if sl is None:
                    rep.add("V14", "치명", where, f"패널 {ref} — 본문 페이지 없음")
                    continue
                bn = sl.get("base_n")
            elif ref.startswith("표"):
                t = by_index.get(int(ref[1:])) if ref[1:].isdigit() else None
                if t is None:
                    rep.add("V14", "치명", where, f"패널 {ref} — 통계표 인덱스 없음")
                    continue
                bn = t.get("base_n")
            else:
                rep.add("V14", "치명", where, f"패널 ref 형식 오류: {ref!r}")
                continue
            if pn.get("base_n") is not None and bn is not None and int(pn["base_n"]) != int(bn):
                rep.add("V14", "치명", where, f"패널 {ref} base {pn['base_n']} ≠ 원자료 {bn}")
            topic = str(pn.get("topic") or "")
            if not topic or re.match(r"^[A-Z]{1,2}\d", topic):
                rep.add("V14", "경고", where, f"패널 {ref} 주제명이 없거나 문항코드로 시작 (R-131): {topic!r}")
        if n_banner > 1:
            # 폭 제한은 **원형 배너 표**(17열)에만 걸린다. 전치 교차 패널(R-145)은 폭이 「1 + 보기 수」라
            # 한 층에 둘을 나란히 놓을 수 있다 — 그것이 전치를 쓰는 이유다.
            rep.add("V14", "경고", where, f"배너 표 패널 {n_banner}개 — 한 층에 하나만 (세로 쌓기)")
        n_cross = sum(1 for pn in panels if str(pn.get("chart") or "").startswith("배너")
                      and not str(pn.get("chart") or "").startswith("배너 표"))
        if n_cross > 2:
            rep.add("V14", "경고", where, f"교차 패널 {n_cross}개 — 꼭지당 2개까지 (R-145 ⑤)")
        # 근거 줄 — 요약 멘트는 헤드라인 포함 ≤3줄 (R-032 · guide/14 §7: LG 2줄 · SKEE 3줄. 사용자 지시 2026-08-23)
        if len(sm.get("body") or []) > SUMMARY_BODY_MAX:
            rep.add("V14", "치명", where, f"요약 근거 줄 {len(sm['body'])}줄 > {SUMMARY_BODY_MAX}줄 — 헤드라인 + {SUMMARY_BODY_MAX}줄까지")
        for ln in sm.get("body") or []:
            t = str(ln).strip()
            # R-198 — 요약·결론 줄은 60자를 강제하지 않는다(가독 우선, 사용자 지시 2026-08-28).
            # 본문 멘트와 달리 요약은 여러 표를 한 줄로 잇느라 길어지고, 성분을 깎으면 뜻이 흐려진다.
            # 한 줄 물리 한계(맑은 고딕 12pt · 32cm 상자 실측 74자)를 넘을 때만 알린다 — 렌더 두 줄.
            if len(t) > SUMMARY_LINE_WRAP:
                rep.add("V14", "경고", where,
                        f"근거 줄 {len(t)}자 > {SUMMARY_LINE_WRAP}자 — 렌더에서 두 줄이 된다: {t[:50]}")
            if t.endswith("."):
                rep.add("V14", "경고", where, f"마침표: {t[:50]}")
            # 코퍼스 불변식 — 요약은 인과·역접·제언을 담지 않는다 (5/5 사례)
            if re.match(r"^\s*(따라서|그러므로|이에|그러나|하지만|또한|한편)", t):
                rep.add("V14", "경고", where,
                        f"요약 헤드라인에 문두 접속부사 — 인과·역접은 결론이 독점한다(코퍼스 5/5 0회): {t[:40]}")
            if re.search(r"(필요함|필요가 있음|요구됨|해야 함|제안함|권장)", t):
                rep.add("V14", "경고", where,
                        f"요약에 제언 — 제언 줄 비율 코퍼스 0~3%. 결론으로 옮긴다: {t[:40]}")
            if not _has_final_mieum(t):
                rep.add("V14", "치명", where, f"ㅁ받침 종결 아님: …{t[-8:]!r}")
        # claims — 근거 셀 좌표 대조
        claims = sm.get("claims") or []
        if not claims:
            rep.add("V14", "치명", where, "claims 없음 — 근거 없는 꼭지")
        for c in claims:
            ev = c.get("evidence") or []
            if not ev:
                # ★ 소결론(➔)은 근거 셀이 **없는 것이 정상**이다. 앞 줄들에서 끌어낸
                #   판단이지 표에서 옮긴 값이 아니다 — 코퍼스 L4 자리 (R-206).
                #   여기에 셀을 요구하면 요약이 다시 「표 옮겨 적기」로 되돌아간다.
                if str(c.get("text") or "").lstrip().startswith(("➔", "→")):
                    continue
                # ★ 판단문(수치 없는 줄)도 마찬가지다 (R-222 · 2026-09-01).
                #   「가장 중요한데 만족은 가장 낮음」은 어느 셀에도 없다 — 표들 사이에서
                #   나온 판단이기 때문이다. 여기에 좌표를 요구하면 판단문이 게이트를 못 넘고,
                #   넘기려면 문장이 수치를 끌어안아야 해서 결국 표 옮겨 적기로 되돌아간다.
                #   계약이 R-206 과 반대 방향을 가리키던 자리다.
                #   대신 **그 판단을 받칠 패널이 이 꼭지에 있는가**를 확인한다.
                _txt = str(c.get("text") or "")
                if not numbers_in(_txt):
                    if not (sm.get("panels") or c.get("supports")):
                        rep.add("V14", "치명", where,
                                f"판단문인데 받칠 패널이 이 꼭지에 없음: {_txt[:50]}")
                    continue
                rep.add("V14", "치명", where,
                        f"수치를 쓴 주장에 근거 셀이 없음: {_txt[:50]}")
                continue
            # V19 — 이 주장이 서브베이스 표에 기대는데 문장이 모수를 안 밝혔는가 (R-185)
            for e in ev:
                try:
                    _ti = int(str(e.get("table")).lstrip("표"))
                except (TypeError, ValueError):
                    continue
                info = sub_base.get(_ti)
                if info and not names_base(c.get("text"), info):
                    rep.add("V19", "경고", where,
                            f"표{_ti} 는 전체가 아니라 「{info['label'] or info['cond'] or 'n=' + str(info['n'])}"
                            f"」(n={info['n']}) 에게만 물은 문항인데 문장이 모수를 밝히지 않았습니다\n"
                            f"      주장: {str(c.get('text'))[:70]}")
                    break
            for e in ev:
                n_claims += 1
                ti = e.get("table")
                try:
                    ti = int(str(ti).lstrip("표"))
                except (TypeError, ValueError):
                    rep.add("V14", "치명", where, f"근거 표 식별 불가: {e}")
                    continue
                t = by_index.get(ti)
                if t is None:
                    rep.add("V14", "치명", where, f"근거 표{ti} 없음")
                    continue
                got, err = _table_cell(t, e.get("row_label") or e.get("row") or "", e.get("col_label") or e.get("col") or "전체")
                want = str(e.get("value") or "").strip().rstrip("%")
                if err:
                    rep.add("V14", "치명", where, f"근거 셀 표{ti} [{e.get('row_label') or e.get('row')!r} × {e.get('col_label') or e.get('col') or '전체'!r}] — {err}")
                elif want and got is not None and want != str(got).strip():
                    rep.add("V14", "치명", where, f"근거 값 불일치 표{ti} [{e.get('row_label') or e.get('row')!r}] 주장 {want} ≠ 원본 {got}")
                else:
                    n_ok += 1
    # 결론
    concl = doc.get("conclusions") or []
    total_chars = 0
    nos = {sm.get("no") for sm in summaries}
    for c in concl:
        where = f"결론 {c.get('no')}"
        body = " ".join(str(x) for x in (c.get("body") or []))
        total_chars += len(body)
        _lines = [x for x in (c.get("body") or []) if str(x).strip()]
        _lo, _hi = CONCLUSION_LINES
        if _lines and not (_lo <= len(_lines) <= _hi):
            rep.add("V14", "경고", where,
                    f"내용 {len(_lines)}줄 — 4사례 공통 골격은 {_lo}~{_hi}줄이다 (R-166)")
        if len(body) > CONCLUSION_BOX_MAX:
            rep.add("V14", "참고", where,
                    f"박스 {len(body)}자 > 참고치 {CONCLUSION_BOX_MAX}자 — R-126 원안이나 "
                    f"실측(SH 02 는 9줄 약 650자)과 어긋난다. 실제 제약은 박스 높이다(R-166)")
        src = c.get("from_summary") or []
        if not src:
            rep.add("V14", "치명", where, "근거 요약 꼭지 없음 (R-078)")
        for sno in src:
            try:
                sn = int(str(sno).lstrip("요약 ").strip())
            except ValueError:
                sn = None
            if sn not in nos:
                rep.add("V14", "치명", where, f"되짚은 요약 {sno!r} 가 없음 (R-078)")
    if total_chars > CONCLUSION_TOTAL_MAX:
        rep.add("V14", "참고", "결론",
                f"장표 합계 {total_chars}자 > 참고치 {CONCLUSION_TOTAL_MAX}자 — R-126 원안이나 "
                f"실측 장표는 이 값을 넘는다(R-166). 박스 높이가 실제 제약이다")
    # 한국어 문체 예방 게이트 (K-xx) — 헤드라인·해석 줄
    texts = []
    for it in summaries:
        no = it.get("no")
        texts.append((f"요약 {no} 헤드라인", str(it.get("headline") or "").lstrip("- ")))
        for i, b in enumerate(it.get("body") or [], 1):
            texts.append((f"요약 {no} 해석 {i}", str(b)))
    korean_style_checks(rep, "V14", "", texts, _ments_from_facts(facts))

    # ── V20 요약 패널의 그래프 심사 (SC-xxx · guide/summary_chart_rules.yaml)
    #    「평균을 독립 그래프로 그리지 않는다」·「최상위 단독 차트를 만들지 않는다」처럼
    #    **값이 아니라 그림의 설계**를 보는 자리다. 값 대조(V13·V18)가 통과해도 여기서 걸린다.
    #    근거는 납품 최종본 5벌의 요약 차트 358개 실측(guide/19)이고, 규칙은 YAML 이 갖는다.
    try:
        import rules as _R
        cmap_path = smap.parent / "conclusion_map.json"
        cdoc = None
        if cmap_path.exists():
            with io.open(cmap_path, encoding="utf-8") as fh:
                cdoc = json.load(fh)
        for f in _R.audit_summary_panels(summaries, by_index, conclusion=cdoc):
            rep.add("V20", "치명" if f.severity == "fatal" else "경고", f.where,
                    f.message + (f" → {f.remedy}" if f.remedy else ""))
    except Exception as e:                                # 규칙 파일이 없거나 PyYAML 부재
        rep.add("V20", "참고", "요약 그래프 심사", f"건너뜀 — {e}")

    rep.stats.update({"summaries": len(summaries), "conclusions": len(concl),
                      "claims_cells": n_claims, "claims_ok": n_ok, "conclusion_chars": total_chars})
    return rep


def verify_conclusion(cmap: Path, tables: Path, smap: Path | None) -> Report:
    """guide-conclusion 의 conclusion_map.json — 근거 셀 좌표 · from_summary 존재 · 길이(R-126)."""
    rep = Report()
    with io.open(cmap, encoding="utf-8") as fh:
        doc = json.load(fh)
    with io.open(tables, encoding="utf-8") as fh:
        tdoc = json.load(fh)
    tlist = tdoc["tables"] if isinstance(tdoc, dict) else tdoc
    by_index = {t["index"]: t for t in tlist}
    kpi = tables.parent / "KPI요약표_직접생성.json"
    if kpi.exists():
        with io.open(kpi, encoding="utf-8") as fh:
            for t in json.load(fh).get("tables", []):
                by_index[t["index"]] = t
    nos: set[int] = set()
    if smap and smap.exists():
        with io.open(smap, encoding="utf-8") as fh:
            nos = {int(x["no"]) for x in json.load(fh).get("summaries", [])}
    items = doc.get("items") or []
    total = 0
    n_cells = n_ok = 0
    for c in items:
        where = f"결론 {c.get('no')}"
        if not (c.get("title") or "").strip():
            rep.add("V15", "치명", where, "제목(판단 문장) 없음")
        body = " ".join(str(x) for x in (c.get("body") or []))
        total += len(body)
        if len(body) > CONCLUSION_BOX_MAX:
            rep.add("V15", "경고", where, f"박스 {len(body)}자 > {CONCLUSION_BOX_MAX}자 (R-126)")
        # 제언은 body 안의 `→` 줄이거나 **action 필드**다. 후자는 조립기가 굵게 뽑는다.
        if not (c.get("action") or any("→" in str(x) for x in (c.get("body") or []))):
            rep.add("V15", "경고", where, "제언 없음 (R-126: 판단 + → 제언)")
        src = c.get("from_summary") or []
        if not src:
            rep.add("V15", "치명", where, "근거 요약 꼭지 없음 (R-078)")
        for sno in src:
            try:
                sn = int(str(sno).replace("요약", "").strip())
            except ValueError:
                sn = None
            if nos and sn not in nos:
                rep.add("V15", "치명", where, f"되짚은 요약 {sno!r} 가 없음 (R-078)")
        cited = numbers_in(body)
        ev = c.get("evidence") or []
        ev_vals = set()
        for e in ev:
            n_cells += 1
            try:
                ti = int(str(e.get("table")).lstrip("표"))
            except (TypeError, ValueError):
                rep.add("V15", "치명", where, f"근거 표 식별 불가: {e}")
                continue
            t = by_index.get(ti)
            if t is None:
                rep.add("V15", "치명", where, f"근거 표{ti} 없음")
                continue
            got, err = _table_cell(t, e.get("row_label") or "", e.get("col_label") or "전체")
            want = str(e.get("value") or "").strip().rstrip("%")
            ev_vals.add(want)
            if err:
                rep.add("V15", "치명", where, f"근거 셀 표{ti} [{e.get('row_label')!r} × {e.get('col_label') or '전체'!r}] — {err}")
            elif want and got is not None and want != str(got).strip():
                rep.add("V15", "치명", where, f"근거 값 불일치 표{ti} [{e.get('row_label')!r}] 주장 {want} ≠ 원본 {got}")
            else:
                n_ok += 1
        # 본문에 인용한 수치는 전부 evidence 에 있어야 한다 (합산값은 evidence.value 에 검산값으로 등재)
        for tok in cited:
            if tok not in ev_vals:
                rep.add("V15", "경고", where, f"본문 수치 {tok} 가 evidence 에 없음 (합산이면 검산값으로 등재)")
    if total > CONCLUSION_TOTAL_MAX:
        rep.add("V15", "경고", "결론", f"장표 합계 {total}자 > {CONCLUSION_TOTAL_MAX}자 (R-126)")
    # 한국어 문체 예방 게이트 (K-xx) — 제목·본문·제언 줄
    texts = []
    for it in items:
        no = it.get("no")
        texts.append((f"결론 {no} 제목", str(it.get("title") or "")))
        for i, b in enumerate(it.get("body") or [], 1):
            texts.append((f"결론 {no} 줄 {i}", str(b)))
    korean_style_checks(rep, "V15", "", texts, _ments_from_facts(cmap.parent / "slide_facts.json"),
                        title_max=CONCLUSION_TITLE_MAX, line_max=CONCLUSION_LINE_MAX)
    rep.stats.update({"conclusions": len(items), "conclusion_chars": total, "claims_cells": n_cells, "claims_ok": n_ok})
    return rep


# ─────────────────────────────────────────────────────────────
# 목차 단독 모드 — 빌드 전 게이트 (V-100 전수성 · V-101 한 어절 · V-102 중복 · seq/P 정합)
# ─────────────────────────────────────────────────────────────

def _toc_pseudo_blocks(assign: dict[int, dict], seq_index: dict[int, dict]) -> list[dict]:
    """목차 배정표 한 행을 `read_blocks` 가 돌려주는 블록 모양으로 흉내 낸다.

    그래야 `check_toc_naming` · `check_toc_coverage` 를 **그대로** 재사용한다 —
    목차용 판정을 따로 만들면 같은 규칙이 두 벌이 된다(R-118).
    """
    blocks: list[dict] = []
    for seq, rec in sorted(assign.items()):
        fact = seq_index.get(seq) or {}
        blocks.append({
            "row": rec.get("page") or seq, "index": len(blocks), "seq": seq,
            "is_struct": False,
            "slide": {"question_code": fact.get("question_code") or rec.get("question_code"),
                      "source_tables": fact.get("source_tables") or []},
            "values": {"Slide title": rec.get("slide_title"), "Sub title": rec.get("sub_title")},
            "labels": LABELS[:], "ment": [], "rows": [],
        })
    return blocks


_QCODE = re.compile(r"[A-Z]{1,2}\d*(?:-\d+)?")


def _toc_part3_items(toc: Path) -> set[str]:
    """첫 펜스 블록의 Part Ⅲ 항목명 집합(번호 제거)."""
    text = io.open(toc, encoding="utf-8").read()
    m = re.search(r"```\n(.*?)```", text, re.S)
    if not m:
        return set()
    block = m.group(1)
    m3 = re.search(r"Part Ⅲ[^\n]*\n(.*)", block, re.S)
    if not m3:
        return set()
    # 줄 단위로 읽는다. `\d+\.\s*[^\s]+` 로 잡으면 「3. 시설관리(FM) 위탁 현황」이
    # 첫 어절까지만 잡혀 여러 어절짜리 챕터명이 전부 불일치로 떨어진다.
    out: set[str] = set()
    for ln in m3.group(1).splitlines():
        mm = re.match(r"^\s*(\d+)\.\s+(\S.*?)\s*$", ln)
        if mm:
            out.add(mm.group(2).strip())
    return out


def _toc_exclusion_rows(toc: Path) -> list[dict]:
    """제외 문항 표 — 첫 열이 `-` 인 행. 두 번째 칸 = 문항코드, **마지막 칸 = 제외 사유**."""
    rows: list[dict] = []
    for line in io.open(toc, encoding="utf-8"):
        if not line.lstrip().startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 2 or cells[0].isdigit():
            continue
        if cells[0] in ("-", "—", "×", "제외"):
            # 고정 포맷은 `| - | 문항 | 유형 | 표 | 제외 사유 |` 이지만, 문항코드가 다른 칸에 와도 읽는다 (초안 실측 2026-08-23)
            code = next((c for c in cells[1:] if _QCODE.fullmatch(c)), None)
            if code:
                reason = cells[-1] if len(cells) > 2 else ""
                if not reason or _QCODE.fullmatch(reason) or reason.replace(" ", "") in ("전체500",):
                    reason = next((c for c in cells[1:] if len(c) > 8 and not _QCODE.fullmatch(c)), reason)
                rows.append({"code": code, "reason": reason})
    return rows


def _toc_exclusions(toc: Path) -> list[str]:
    return [r["code"] for r in _toc_exclusion_rows(toc)]


def _toc_table_cells(toc: Path) -> dict[int, tuple[str, str]]:
    """배정표의 `표`·`구성 및 BASE` 칸 — build_facts.parse_toc 은 안 읽으므로 여기서 직접 읽는다."""
    out: dict[int, tuple[str, str]] = {}
    header: list[str] | None = None
    for line in io.open(toc, encoding="utf-8"):
        if not line.lstrip().startswith("|"):
            continue
        cells = [c.replace("\\|", "|").strip() for c in re.split(r"(?<!\\)\|", line.strip().strip("|"))]
        if header is None:
            if "seq" in cells and "표" in cells:
                header = cells
            continue
        if len(cells) != len(header) or not cells[header.index("seq")].isdigit():
            continue
        rec = dict(zip(header, cells))
        out[int(rec["seq"])] = (rec.get("표", ""), rec.get("구성 및 BASE", ""))
    return out


def verify_toc(toc: Path, questionnaire: Path | None, facts: Path | None) -> Report:
    """`목차.md` 만으로 V13 목차 규칙을 돈다. 빌드 산출물 없이 목차 초안을 게이트한다."""
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import build_facts as BF  # noqa: E402
    rules = load_rules()
    rep = Report()

    assign, warns = BF.parse_toc(toc)
    for w in warns:
        rep.add("V13", "치명" if "찾지 못했" in w else "경고", toc.name, w)
    rep.stats["toc_rows"] = len(assign)

    seq_index: dict[int, dict] = {}
    if facts and facts.exists():
        with io.open(facts, encoding="utf-8") as fh:
            seq_index = {s["seq"]: s for s in json.load(fh).get("slides", [])}
    questions: dict[str, dict] = {}
    if questionnaire and questionnaire.exists():
        with io.open(questionnaire, encoding="utf-8") as fh:
            qdoc = json.load(fh)
        questions = {q["code"]: q for q in qdoc.get("questions", [])}

    # ── seq 정합: 팩트의 seq 집합과 배정표의 seq 집합이 같아야 한다
    if seq_index:
        want, got = set(seq_index), set(assign)
        if want - got:
            rep.add("V13", "치명", "배정표", f"배정되지 않은 seq {len(want - got)}개: {sorted(want - got)}")
        if got - want:
            rep.add("V13", "치명", "배정표", f"팩트에 없는 seq {len(got - want)}개: {sorted(got - want)}")
        for seq, rec in assign.items():
            fc = (seq_index.get(seq) or {}).get("question_code")
            if fc and rec.get("question_code") and fc != rec["question_code"]:
                rep.add("V13", "치명", f"seq {seq}",
                        f"문항코드 불일치 — 배정표 {rec['question_code']!r} vs 팩트 {fc!r}")
    # ── G-1: 배정표의 꼭지는 첫 펜스 Part Ⅲ 항목 목록에 있어야 한다 (제외표 행이 배정으로 흡수되면 여기서 잡힌다)
    part3 = _toc_part3_items(toc)
    if part3:
        bad = sorted({str(r.get("slide_title") or "") for r in assign.values()
                      if re.sub(r"^\d+\.\s*", "", str(r.get("slide_title") or "")) not in part3})
        for t in bad:
            rep.add("V13", "치명", "배정표", f"꼭지 {t!r} 가 첫 펜스의 Part Ⅲ 항목 목록에 없습니다 — 제외 문항 표에 숫자 seq 가 남았거나 챕터명이 어긋났습니다")
    # ── G-2: 제외 문항 표는 고정 5열 `| - | 문항 | 유형 | 표 | 제외 사유 |` — seq 열을 두지 않는다
    for ln in io.open(toc, encoding="utf-8"):
        cells = [c.strip() for c in ln.strip().strip("|").split("|")] if ln.lstrip().startswith("|") else []
        if len(cells) >= 5 and cells[0] in ("-", "—", "×", "제외") and any(_QCODE.fullmatch(c) for c in cells[1:]):
            if len(cells) >= 7 and cells[2].isdigit():
                rep.add("V13", "치명", "제외 문항", f"제외 행에 숫자 seq 가 있습니다 ({cells[3] if len(cells) > 3 else cells[1]}) — parse_toc 가 본문 배정으로 읽습니다. 고정 5열로 쓰세요")
            elif len(cells) != 5:
                rep.add("V13", "경고", "제외 문항", f"제외 표가 고정 5열이 아닙니다 ({len(cells)}열) — `| - | 문항 | 유형 | 표 | 제외 사유 |`")
    # ── P 연속성: 1..N 빠짐·중복 없음
    pages = sorted(r.get("page") for r in assign.values() if r.get("page"))
    if pages and pages != list(range(1, len(assign) + 1)):
        rep.add("V13", "경고", "배정표", f"P 가 1..{len(assign)} 연속이 아닙니다 (실제 {pages[:5]}…{pages[-3:]})")
    # ── 빈 칸
    for seq, rec in assign.items():
        for k, lab in (("slide_title", "꼭지"), ("sub_title", "서브 타이틀"), ("base_note_toc", "구성 및 BASE")):
            if not rec.get(k):
                rep.add("V13", "경고", f"seq {seq}", f"`{lab}` 이 비었습니다")

    # ── 표 열 ↔ 팩트 source_tables · 구성 및 BASE 끝 숫자 ↔ base_n (심사 J-08 보강)
    if seq_index:
        cells = _toc_table_cells(toc)
        for seq, (tbl, base) in cells.items():
            fact = seq_index.get(seq) or {}
            want = set(int(x) for x in fact.get("source_tables") or [])
            got = set(int(x) for x in re.findall(r"\d+", tbl))
            if want and got != want:
                rep.add("V13", "치명", f"seq {seq}",
                        f"`표` 칸 {sorted(got)} 이 팩트 source_tables {sorted(want)} 와 다릅니다")
            m = re.search(r"(\d[\d,]*)\s*$", base)
            bn = fact.get("base_n")
            if m and bn is not None and int(m.group(1).replace(",", "")) != int(bn):
                rep.add("V13", "치명", f"seq {seq}",
                        f"`구성 및 BASE` 의 사례수 {m.group(1)} 이 통계표 base {bn} 와 다릅니다")
    for ex in _toc_exclusion_rows(toc):
        if len(ex["reason"]) < 8:
            rep.add("V13", "경고", f"제외 {ex['code']}", "제외 사유가 비었거나 너무 짧습니다 (R-027 — 뺐다고 명시)")

    blocks = _toc_pseudo_blocks(assign, seq_index)
    check_toc_naming(rep, rules, blocks)
    check_toc_coverage(rep, rules, blocks, questions,
                       _toc_exclusions(toc) if toc else None)

    # ── 전수성 검증식 (V-100): 설문지 문항 수 == 배정 + 명시 제외
    if questions:
        assigned = {b["slide"]["question_code"] for b in blocks if b["slide"].get("question_code")}
        excluded = set(_toc_exclusions(toc))
        both = assigned & excluded
        if both:
            rep.add("V13", "경고", "제외 문항", f"배정과 제외에 동시에 있는 문항: {sorted(both)}")
        neither = [c for c in questions if c not in assigned and c not in excluded]
        rep.stats.update({"questions": len(questions), "assigned": len(assigned & set(questions)),
                          "excluded": len(excluded & set(questions))})
        if neither:
            rep.add("V13", "치명", "목차 전수성",
                    f"배정도 제외도 되지 않은 문항 {len(neither)}개: {', '.join(neither)}\n"
                    f"      → 검증식 {len(questions)} = {len(assigned & set(questions))}(배정) + "
                    f"{len(excluded & set(questions))}(제외) 이 깨졌습니다 (R-027)")
        unknown = sorted((assigned | excluded) - set(questions))
        if unknown:
            rep.add("V13", "경고", "목차 전수성", f"설문지에 없는 문항코드: {unknown}")
    return rep


def render(rep: Report, guide: Path) -> str:
    lines = [f"# 검수 리포트 — {guide.name}", ""]
    if rep.stats.get("mode") == "deck":
        from verify_deck import deck_header_lines
        lines.extend(deck_header_lines(rep))
    else:
        lines.append(f"- 블록 {rep.stats.get('blocks', 0)}개 · 통계표 {rep.stats.get('tables', 0)}장")
        lines.append(f"- 멘트 {rep.stats.get('ment_lines', 0)}줄 · 수치 {rep.stats.get('ment_numbers_checked', 0)}개 대조")
    lines.append("")

    from collections import Counter
    sev = Counter(f.severity for f in rep.findings)
    if not rep.findings:
        lines.append("## ✅ 지적사항 없음")
        lines.append("")
        lines.append("모든 수치가 통계표 원본과 일치하고, 블록 규격·신호어·템플릿 무결성 모두 통과했습니다.")
        return "\n".join(lines)

    head = f"## 요약 — 치명 {sev.get('치명', 0)} · 경고 {sev.get('경고', 0)} · 참고 {sev.get('참고', 0)}"
    if sev.get("예시오류"):
        head += f" · 예시오류 {sev['예시오류']}"
    lines.append(head)
    lines.append("")
    by_check = Counter(f.check for f in rep.findings)
    names = {"V0": "규칙 로드", "V1": "블록 무결성", "V2": "그리드 수치",
             "V3": "멘트 수치(환각)", "V4": "Base n 정합", "V5": "Qu 텍스트",
             "V6": "신호어", "V7": "미기입 잔존", "V8": "템플릿 무결성",
             "V9": "센티널 표기", "V10": "Base/Unit 형식", "V11": "멘트 문체·수준어",
             "V12": "송부본", "V13": "목차 규칙",
             # ── 덱 모드 (verify_deck.py)
             "D0": "규칙 로드", "D1": "조인·문항코드", "D2": "미기입 마커",
             "D3": "미충전 반복 장표", "D4": "제작팀 메모 잔존", "D5": "Base n 정합",
             "D6": "Base/Unit 형식", "D7": "n<30 각주", "D8": "표 셀 ↔ 통계표",
             "D9": "차트 ↔ 표", "D10": "소수점 자릿수", "D11": "합계 정합",
             "D12": "0 표기", "D13": "센티널 표기", "D14": "그래프 상단 Text",
             "D15": "정렬", "D16": "차트타입", "D17": "Qu 텍스트",
             "D18": "멘트 수치(환각)", "D19": "멘트 보기명", "D20": "멘트 문체",
             "D21": "수준어 사다리", "D22": "따옴표 짝", "D23": "배너 선별",
             "D24": "목차·간지 정합", "D25": "겹침막대 라벨 충돌", "D26": "제목 변경 이력"}
    for k, v in sorted(by_check.items()):
        lines.append(f"- **{k} {names.get(k, '')}** — {v}건")
    lines.append("")

    for sv in ("치명", "경고", "참고", "예시오류"):
        group = [f for f in rep.findings if f.severity == sv]
        if not group:
            continue
        lines.append(f"## {sv} ({len(group)}건)")
        lines.append("")
        for f in group[:200]:
            lines.append(f"- `{f.check}` **{f.where}**  \n  {f.detail}")
        if len(group) > 200:
            lines.append(f"- … 외 {len(group) - 200}건")
        lines.append("")
    return "\n".join(lines)


def main() -> int:
    ap = argparse.ArgumentParser(description="제작가이드(.xlsm) · 제작본(.pptx) 검수")
    ap.add_argument("target", type=Path,
                    help=".xlsm/.xlsx → 가이드 검수, .pptx → 덱 검수")
    ap.add_argument("tables", type=Path, nargs="?",
                    help="통계표 JSON. 가이드 모드는 필수, 덱 모드는 없으면 값 대조를 건너뛴다")
    ap.add_argument("--mode", choices=["auto", "guide", "deck", "toc", "ments", "summary", "conclusion"], default="auto")
    ap.add_argument("--facts", type=Path)
    ap.add_argument("--questionnaire", type=Path)
    ap.add_argument("--piping", type=Path,
                    help="빌더가 쓴 파이핑 치환 맵 — Qu 대조에 같은 규칙을 적용한다")
    ap.add_argument("--toc", type=Path,
                    help="확정 목차(work/목차.md). 안 주면 facts 옆에서 찾는다 — "
                         "요약·결론 블록을 항목명으로 알아보는 데 쓴다")
    ap.add_argument("--guide-blocks", type=Path)
    ap.add_argument("--deck-facts", type=Path,
                    help="extract_deck() 재실행을 건너뛰고 쓸 ground_truth.json")
    ap.add_argument("--known-defects", metavar="DECK_ID",
                    help="예시 덱 ID(LG·SKEE). 명시했을 때만 R-104 억제가 켜진다")
    ap.add_argument("--extra-values", type=Path, dest="extra_values",
                    help="통계표 밖에서 온 값의 출처 파일 (예: 전년 보고서 차트에서 추출한 시계열). "
                         "{\"values\": [\"71.5\", …]} 또는 그 값들을 담은 JSON. V2 대조 집합에 더한다")
    ap.add_argument("--out", type=Path)
    ap.add_argument("--json", type=Path, dest="json_out")
    args = ap.parse_args()

    mode = args.mode
    if mode == "auto":
        ext = args.target.suffix.lower()
        if ext in (".xlsm", ".xlsx"):
            mode = "guide"
        elif ext == ".pptx":
            mode = "deck"
        elif ext == ".md":
            mode = "toc"
        elif ext == ".json":
            mode = "ments"
        else:
            ap.error(f"확장자로 모드를 정할 수 없습니다: {args.target.name} "
                     f"(--mode guide|deck 으로 지정하세요)")

    if mode == "toc":
        rep = verify_toc(args.target, args.questionnaire, args.facts)
    elif mode == "ments":
        if not args.facts:
            ap.error("멘트 모드는 --facts slide_facts.json 이 필요합니다")
        rep = verify_ments(args.target, args.facts)
    elif mode == "summary":
        if args.tables is None:
            ap.error("요약 모드는 tables.json 이 필요합니다")
        rep = verify_summary(args.target, args.tables, args.facts)
    elif mode == "conclusion":
        if args.tables is None:
            ap.error("결론 모드는 tables.json 이 필요합니다")
        rep = verify_conclusion(args.target, args.tables, args.target.parent / "summary_map.json")
    elif mode == "guide":
        if args.tables is None:
            ap.error("가이드 모드는 tables.json 이 필요합니다")
        rep = verify(args.target, args.tables, args.facts, args.questionnaire,
                     getattr(args, "piping", None), getattr(args, "toc", None),
                     getattr(args, "extra_values", None))
        check_doc_hygiene(rep, args.target)
    else:
        # 순환 임포트를 피해 여기서만 부른다 — verify_deck 이 이 모듈을 import 한다.
        from verify_deck import verify_deck
        rep = verify_deck(args.target, args.tables, args.facts, args.questionnaire,
                          args.guide_blocks, args.deck_facts, args.known_defects)
        check_doc_hygiene(rep, args.target)

    text = render(rep, args.target)

    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        with io.open(args.out, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"검수 리포트 → {args.out}")
    if args.json_out and mode != "ments":
        from verify_deck import findings_json
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        with io.open(args.json_out, "w", encoding="utf-8") as fh:
            fh.write(findings_json(rep))
        print(f"Finding JSON → {args.json_out}")

    from collections import Counter
    sev = Counter(f.severity for f in rep.findings)
    if mode == "deck":
        from verify_deck import deck_header_lines
        for line in deck_header_lines(rep):
            print(line.lstrip("- "))
    elif mode == "conclusion":
        print(f"결론 {rep.stats.get('conclusions', 0)}박스({rep.stats.get('conclusion_chars', 0)}자) · "
              f"근거 셀 {rep.stats.get('claims_ok', 0)}/{rep.stats.get('claims_cells', 0)} 일치")
    elif mode == "summary":
        print(f"요약 {rep.stats.get('summaries', 0)}꼭지 · 결론 {rep.stats.get('conclusions', 0)}박스({rep.stats.get('conclusion_chars', 0)}자) · "
              f"근거 셀 {rep.stats.get('claims_ok', 0)}/{rep.stats.get('claims_cells', 0)} 일치")
    elif mode == "ments":
        acc, rej = rep.stats.get("accepted", {}), rep.stats.get("rejected", {})
        print(f"멘트 {rep.stats.get('ment_items', 0)}장 — 수용 {len(acc)} · 반려 {len(rej)}"
              + (f" (반려 seq: {sorted(rej)})" if rej else ""))
        if args.json_out:
            with io.open(args.json_out, "w", encoding="utf-8") as fh:
                json.dump({"accepted": acc, "rejected": rej}, fh, ensure_ascii=False, indent=1)
    elif mode == "toc":
        print(f"배정 {rep.stats.get('toc_rows', 0)}행 · 문항 {rep.stats.get('questions', 0)} = "
              f"배정 {rep.stats.get('assigned', 0)} + 제외 {rep.stats.get('excluded', 0)}")
    else:
        print(f"블록 {rep.stats.get('blocks', 0)}개 · 멘트 수치 {rep.stats.get('ment_numbers_checked', 0)}개 대조")
    tail = f" · 예시오류 {sev['예시오류']}" if sev.get("예시오류") else ""
    print(f"치명 {sev.get('치명', 0)} · 경고 {sev.get('경고', 0)} · 참고 {sev.get('참고', 0)}{tail}")
    for f in rep.findings[:12]:
        print(f"  [{f.severity}] {f.check} {f.where} — {f.detail.splitlines()[0][:90]}")
    if len(rep.findings) > 12:
        print(f"  … 외 {len(rep.findings) - 12}건 (리포트 참조)")

    return 1 if rep.fatal else 0


if __name__ == "__main__":
    sys.exit(main())
