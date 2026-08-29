#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""deckbuild.py -- 덱 조립. 슬라이드 삭제 / 복제 / 재배열 / 부분집합 추출.

deckops.Deck 을 **받아서** 조작한다. Deck 을 상속하지도 재구현하지도 않는다.
이 모듈은 슬라이드 XML 본문을 **한 바이트도 해석하지 않는다** — 파트를 통째로 복사하고,
관계(rels)와 sldIdLst 와 [Content_Types].xml 만 손본다.
그래서 R-105(표 기하) · R-106(차트 프레임) · R-107(행 서식) · R-109(겹침막대)는
이 모듈을 통과해도 원본 그대로다. 기하를 바꾸는 코드는 여기에 없다.

  ① 조립(장수 확정)  ← 이 모듈
  ② LG 자산 제거(테마 색·폰트·로고)
  ③ 본문 채우기(fill_deck)

pptx 에 슬라이드를 하나 더하거나 빼려면 아래를 동시에 맞춰야 한다.
하나라도 어긋나면 PowerPoint 가 "복구가 필요합니다"(0x80070570) 로 파일을 거부한다.

  1. ppt/slides/slideN.xml                    슬라이드 본체
  2. ppt/slides/_rels/slideN.xml.rels         레이아웃·이미지·차트 연결
  3. ppt/charts/{chart,colors,style}N.xml     차트 일습 (+ ppt/embeddings/*.xlsx)
  4. ppt/_rels/presentation.xml.rels          rId -> 슬라이드 파트
  5. ppt/presentation.xml  <p:sldIdLst>       표시 순서 (여기가 순서의 유일한 원천)
  6. [Content_Types].xml                      Override 등록

실측으로 확인한 사실 (LG 99장 덱 · PowerPoint COM 대조)
-------------------------------------------------------
· 표시 순서는 sldIdLst 자식 순서가 100% 결정한다. slideN.xml 의 N 은 순서와 무관하다.
  그래서 이 모듈의 주소 지정은 **표시 위치(1-based)** 또는 **파트 이름** 두 가지뿐이다.
  파일번호 주소는 제공하지 않는다 — 조립 후 파일번호는 의미를 잃는다.
· 재배열은 <p:sldId> 요소를 통째로 옮기는 것으로 끝난다. r:id 만 바꿔치기하면
  id <-> rId 쌍이 어긋나 sldId/@id 의 의미가 뒤집힌다. 절대 하지 않는다.
· 슬라이드 번호(하단 페이지 번호)는 레이아웃의 <a:fld type="slidenum"/> 가 자동 계산한다.
  삭제·복제·재배열 어느 것도 번호를 손댈 필요가 없다. 그래서 API 대상에서 뺐다.
· PowerPoint 가 강제하는 것은 관계 그래프 무결성**만이 아니다.** 2026-08-22 재실측에서
  이 서술을 고쳤다 — 이전 판은 'Content_Types 위생은 검사하지 않는다'고 적혀 있었고
  그게 validate() 에 구멍을 냈다. 실제로는 **누락과 중복을 갈라서** 본다.
      Override 를 지운다        연다 22/22   (xml Default 로 폴백)
      Override 를 댕글링으로    연다 22/22   (없는 파트를 가리켜도 무시)
      Override PartName 중복    거부 0x80CB8001
      Default  Extension 중복   거부 0x80CB8000
      Default  Extension 삭제   거부 0x80CB8002  ← 그 확장자 파트가 전부 무타입이 된다
  즉 **빠진 것은 봐주고 겹치거나 타입이 아예 없는 것은 거부**한다. 누락 2종은 그래도
  규격 위반이고 다른 뷰어에서 검증하지 못했으므로 항상 동기화한다.
· 관계 그래프 쪽도 '대상이 있는가'만으로는 부족하다. **관계를 통째로 지우면** 댕글링이
  안 생겨서 대상 존재 검사를 전부 통과하는데, 슬라이드→레이아웃 · 레이아웃→마스터 ·
  마스터→테마 · presentation→마스터 4종은 없으면 PowerPoint 가 0x80070570 으로 거부한다.
· 반대로 **거부당하지 않는데 더 위험한 파손**이 하나 있다. sldIdLst 항목만 빠진
  '유령 슬라이드'는 파트·rels·Override 가 전부 살아 있어 어떤 참조 검사에도 안 걸리고,
  PowerPoint 는 22장 덱을 **경고 없이 21장으로** 연다. 그래서 이 모듈의 fatal 은
  'PowerPoint 가 거부한다'가 아니라 **'이대로 납품하면 안 된다'**로 정의한다.
· 고아 파트를 남기면 파일이 30% 넘게 불고, **지운 슬라이드의 원문·차트·엑셀 데이터가
  파일 안에 그대로 남는다.** 다른 고객사 자료가 납품물에 실려 나가는 정보 유출이다.
  그래서 gc 는 선택이 아니라 기본값이다.
· 복제 시 새로 뜨는 파트는 ppt/charts · ppt/embeddings · ppt/notesSlides 세 갈래뿐이다.
  media 와 slideLayout 은 Target 을 그대로 두어 공유한다.
  (챕터 간지 복제가 새 파트 1개로 끝나는 이유. 요약 장표는 차트 5개라 21개가 뜬다.)
· 임베디드 워크북(.xlsx)만은 반드시 새로 떠야 한다. 공유하면 PowerPoint 의
  '데이터 편집'이 두 장표를 동시에 바꾼다.

이 모듈이 피해 가는 실측 함정
-----------------------------
(가) 차트가 2개 이상인 슬라이드. 프로토타입 slideadd.py 는 rels 를 루프로 훑으며
     src_chart 를 덮어써 마지막 하나만 복제했다 — 요약 7장 전부가 깨진다.
     여기서는 rels 를 전수 순회하며 **각 차트마다** 일습을 뜬다.
(나) chart10~chart24 15개는 colors/style rels 가 아예 없다. '항상 있다'고 가정하면
     KeyError 거나 조용히 남의 파트를 복사한다. 여기서는 rels 에 있는 것만 복제한다.
(다) ppt/embeddings/Microsoft_Excel_Worksheet.xlsx (번호 없음) 1개. 채번 정규식을
     (\\d*) 로 쓰면 int('') 로 죽는다. (\\d+) 로 잡고 번호 없는 놈은 후보에서 뺀다.
(라) Deck.flush() 는 _dirty 를 돌며 _tree_cache 의 트리를 parts 에 다시 써 넣는다.
     parts 에서만 지우면 저장할 때 되살아난다. remove_part 는 세 곳을 동시에 비운다.
(마) 슬라이드가 직접 쥔 이미지를 레이아웃도 쥐고 있을 수 있다. 슬라이드만 보고 GC 하면
     레이아웃이 깨진다. protected_parts() 로 '슬라이드를 거치지 않고 도달하는 파트'를
     먼저 확보하고 그 밖에서만 지운다.

의존: lxml (필수) / deckops (필수) / PowerPoint COM (assert_openable 만, 선택)
자체 검증:  PYTHONIOENCODING=utf-8 python deckbuild.py --selftest
"""

from __future__ import annotations

import json
import posixpath
import re
import subprocess
import sys
import tempfile
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable, Literal, Sequence

from lxml import etree

from deckops import Deck, q, A, C, P, R, PR

__all__ = [
    "SHARED_ROOTS", "PRIVATE_ROOTS", "CT_OVERRIDE",
    "resolve", "positions",
    "owned_parts", "protected_parts", "refcount", "sole_owner_parts",
    "DeleteReport", "delete_slides", "keep_only",
    "CloneResult", "clone_slide", "next_free",
    "reorder", "move_slide", "apply_plan",
    "remove_part", "sync_content_types", "gc_orphans", "refresh_app_xml",
    "Problem", "validate", "assert_openable", "save_checked",
    "deck_map", "layout_index",
    "S1_KEEP", "build_s1_skeleton",
]

# --------------------------------------------------------------------------- #
# 경계선 두 개
# --------------------------------------------------------------------------- #

#: 의존 그래프 탐색을 **중단**하는 경계.
#: 이 경계가 없으면 슬라이드 -> 레이아웃 -> 마스터 -> 레이아웃 14개로 번져
#: 모든 레이아웃 미디어가 'x99 공유' 로 보인다(실측 확인).
SHARED_ROOTS = re.compile(r"^ppt/(slideLayouts|slideMasters|theme|notesMasters|tableStyles)")

#: 복제 시 **새 파트로 뜨는** 갈래. 나머지(media 등)는 Target 그대로 공유한다.
PRIVATE_ROOTS = re.compile(r"^ppt/(charts|embeddings|notesSlides)/")

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

#: media(.png/.jpeg/.emf/.svg) 와 워크북(.xlsx) 은 [Content_Types].xml 의 Default
#: 확장자가 이미 커버한다. Override 를 만들면 안 된다.
CT_OVERRIDE: dict[str, str] = {
    "slide": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    "chart": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    "colors": "application/vnd.ms-office.chartcolorstyle+xml",
    "style": "application/vnd.ms-office.chartstyle+xml",
    "notesSlide": "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
}

#: 파트 이름 규칙. 채번 정규식이 (\d+) 인 것이 핵심 —
#: 번호 없는 ppt/embeddings/Microsoft_Excel_Worksheet.xlsx 를 후보에서 뺀다.
_KIND_PATTERNS: dict[str, tuple[str, str]] = {
    "slide":      (r"^ppt/slides/slide(\d+)\.xml$",                              "ppt/slides/slide{n}.xml"),
    "chart":      (r"^ppt/charts/chart(\d+)\.xml$",                              "ppt/charts/chart{n}.xml"),
    "colors":     (r"^ppt/charts/colors(\d+)\.xml$",                             "ppt/charts/colors{n}.xml"),
    "style":      (r"^ppt/charts/style(\d+)\.xml$",                              "ppt/charts/style{n}.xml"),
    "xlsx":       (r"^ppt/embeddings/Microsoft_Excel_Worksheet(\d+)\.xlsx$",     "ppt/embeddings/Microsoft_Excel_Worksheet{n}.xlsx"),
    "notesSlide": (r"^ppt/notesSlides/notesSlide(\d+)\.xml$",                    "ppt/notesSlides/notesSlide{n}.xml"),
}

SLDID_MIN = 256
SLDID_MAX = 2147483647

#: int = 1-based 표시 위치(sldIdLst 순서) · str = 파트 이름. 파일번호 주소는 없다.
SlideRef = int | str

PRES = "ppt/presentation.xml"
PRES_RELS = "ppt/_rels/presentation.xml.rels"
ROOT_RELS = "_rels/.rels"
APP_XML = "docProps/app.xml"


# --------------------------------------------------------------------------- #
# 저수준 헬퍼
# --------------------------------------------------------------------------- #

def _xml_bytes(tree: etree._Element) -> bytes:
    """OOXML 파트 직렬화. 선언부를 원본과 같은 꼴로 유지한다."""
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def _norm_target(base_dir: str, target: str) -> str:
    """rels Target -> 패키지 절대 경로(선행 / 없음).

    Deck.rels() 는 절대 Target('/ppt/...')을 정규화하지 않는다. LG 덱에는 절대 Target 이
    0건이지만 제작팀 회신본에는 있을 수 있어 여기서 흡수한다.
    """
    if target.startswith("/"):
        return posixpath.normpath(target).lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")


def _rel_rows(deck: Deck, part: str) -> list[dict[str, str | None]]:
    """파트의 관계 목록. Deck.rels() 와 같은 모양이되 절대 Target 을 정규화한다.

    part='' 을 주면 패키지 루트(_rels/.rels)를 읽는다.
    """
    name = Deck.rels_name(part)
    if name not in deck.parts:
        return []
    base = posixpath.dirname(part)
    out: list[dict[str, str | None]] = []
    for rel in etree.fromstring(deck.parts[name]):
        target = rel.get("Target") or ""
        external = rel.get("TargetMode") == "External"
        out.append({
            "Id": rel.get("Id") or "",
            "Type": rel.get("Type") or "",
            "Target": target,
            "TargetMode": rel.get("TargetMode") or "Internal",
            "Part": None if external else _norm_target(base, target),
        })
    return out


def _reach(deck: Deck, roots: Iterable[str], *,
           enter: Callable[[str], bool] | None = None) -> set[str]:
    """roots 에서 rels 를 따라 도달하는 파트 집합(각 파트의 .rels 파트 포함).

    enter(target) 이 False 를 주면 그 가지로 **들어가지도 않고 집합에 넣지도 않는다.**
    """
    seen: set[str] = set()
    stack = [r for r in roots]
    while stack:
        part = stack.pop()
        if part in seen or part not in deck.parts:
            continue
        seen.add(part)
        rels_name = Deck.rels_name(part)
        if rels_name in deck.parts:
            seen.add(rels_name)
            for row in _rel_rows(deck, part):
                tgt = row["Part"]
                if not tgt:
                    continue
                if enter is not None and not enter(tgt):
                    continue
                stack.append(tgt)
    return seen


def _root_children(deck: Deck) -> list[str]:
    return [r["Part"] for r in _rel_rows(deck, "") if r["Part"]]


def _kind_of(part: str) -> str | None:
    for kind, (rx, _) in _KIND_PATTERNS.items():
        if re.match(rx, part):
            return kind
    return None


def _ct_kind(part: str) -> str | None:
    """[Content_Types].xml Override 가 필요한 파트인가. 필요하면 그 종류를 준다."""
    kind = _kind_of(part)
    return kind if kind in CT_OVERRIDE else None


# --------------------------------------------------------------------------- #
# 채번
# --------------------------------------------------------------------------- #

class _Namer:
    """충돌 없는 새 파트 이름을 배치로 뽑는다.

    69회 복제 × 파트 수백 개에서 매번 전체 파트 목록을 훑지 않으려고 한 번만 스캔하고
    이후에는 카운터를 올린다. 뽑은 이름은 예약 집합에 넣어 같은 배치 안에서도 안 겹친다.
    """

    def __init__(self, deck: Deck):
        self.deck = deck
        self._next: dict[str, int] = {}
        self._taken: set[str] = set(deck.parts)
        for kind, (rx, _) in _KIND_PATTERNS.items():
            hi = 0
            crx = re.compile(rx)
            for name in deck.parts:
                m = crx.match(name)
                if m:
                    hi = max(hi, int(m.group(1)))
            self._next[kind] = hi + 1

    def take(self, kind: str) -> str:
        template = _KIND_PATTERNS[kind][1]
        n = self._next[kind]
        while True:
            name = template.format(n=n)
            n += 1
            if name not in self._taken:
                self._next[kind] = n
                self._taken.add(name)
                return name

    def take_like(self, src: str) -> str:
        """알려진 규칙이 없는 파트(예: 번호 없는 Microsoft_Excel_Worksheet.xlsx)용 폴백.

        확장자를 유지한 채 '__c{n}' 을 붙인다. 규칙에 없는 이름을 억지로 규칙에 끼워
        맞추다 기존 파트를 덮어쓰는 사고를 막는다.
        """
        kind = _kind_of(src)
        if kind is not None:
            return self.take(kind)
        stem, ext = posixpath.splitext(src)
        i = 1
        while True:
            name = f"{stem}__c{i}{ext}"
            i += 1
            if name not in self._taken:
                self._taken.add(name)
                return name


def next_free(deck: Deck, kind: Literal["slide", "chart", "colors", "style", "xlsx", "notesSlide"]) -> str:
    """충돌 없는 다음 파트 이름 하나. (배치로 뽑을 땐 내부 _Namer 가 쓰인다)"""
    if kind not in _KIND_PATTERNS:
        raise ValueError(f"모르는 파트 종류: {kind}")
    return _Namer(deck).take(kind)


# --------------------------------------------------------------------------- #
# 주소 지정
# --------------------------------------------------------------------------- #

def resolve(deck: Deck, ref: SlideRef) -> str:
    """SlideRef -> 슬라이드 파트 이름.

    int  : 1-based **표시 위치**(sldIdLst 순서). 파일번호가 아니다.
    str  : 파트 이름 그대로. 덱에 없으면 KeyError.
    """
    if isinstance(ref, bool):                       # bool 은 int 의 하위형이라 먼저 막는다
        raise TypeError(f"슬라이드 주소로 bool 은 못 쓴다: {ref!r}")
    if isinstance(ref, int):
        order = deck.slide_order()
        if not 1 <= ref <= len(order):
            raise IndexError(f"슬라이드 위치 {ref} (총 {len(order)}장)")
        return order[ref - 1]
    if isinstance(ref, str):
        if ref not in deck.parts:
            raise KeyError(f"슬라이드 파트 없음: {ref}")
        return ref
    raise TypeError(f"SlideRef 는 int(표시 위치) 또는 str(파트 이름): {ref!r}")


def positions(deck: Deck) -> dict[str, int]:
    """파트 이름 -> 1-based 표시 위치 역인덱스."""
    return {part: i for i, part in enumerate(deck.slide_order(), start=1)}


def _resolve_all(deck: Deck, refs: Sequence[SlideRef]) -> list[str]:
    """순서 유지 + 중복 제거."""
    out: list[str] = []
    seen: set[str] = set()
    for ref in refs:
        part = resolve(deck, ref)
        if part not in seen:
            seen.add(part)
            out.append(part)
    return out


# --------------------------------------------------------------------------- #
# 의존 그래프
# --------------------------------------------------------------------------- #

def owned_parts(deck: Deck, slide_part: str) -> set[str]:
    """슬라이드에서 rels 로 도달하는 파트 중 SHARED_ROOTS 가지를 뺀 집합.

    반환값에는 슬라이드 자신과 각 파트의 .rels 파트가 포함된다.
    · 레이아웃/마스터/테마로는 들어가지 않는다 — 들어가면 모든 레이아웃 미디어가
      '전 슬라이드 공유' 로 보여 GC 판정이 통째로 무너진다(실측).
    · **다른 슬라이드로도 들어가지 않는다.** 슬라이드 A 가 슬라이드 B 로 점프하는
      hyperlink 를 갖고 있으면 B 와 B 의 차트까지 A 의 소유로 셈해 버린다.
      (notesSlide 가 자기 슬라이드를 되가리키는 관계도 여기서 걸린다.)
      LG 덱에는 슬라이드 간 관계가 0건이지만 제작팀 회신본에는 있을 수 있다.
    · External 관계(mailto:, https:)는 Part 가 None 이라 자연히 빠진다.
    """
    return _reach(deck, [slide_part],
                  enter=lambda t: not SHARED_ROOTS.match(t) and not t.startswith("ppt/slides/"))


def protected_parts(deck: Deck) -> set[str]:
    """슬라이드를 거치지 **않고** 패키지 루트에서 도달하는 파트 집합.

    레이아웃·마스터·테마·presProps·docProps 와 그들이 쥔 media 가 여기 들어온다.
    슬라이드가 직접 쥔 이미지를 레이아웃도 쥐고 있을 수 있으므로, GC 는 이 집합을
    절대 건드리면 안 된다.
    """
    return _reach(deck, _root_children(deck),
                  enter=lambda t: not t.startswith("ppt/slides/"))


def refcount(deck: Deck) -> dict[str, int]:
    """파트 -> 그 파트를 쥔 슬라이드 수. sldIdLst 에 있는 슬라이드만 센다.

    chart/embedding 이 이 덱에서 1:1 이라는 것은 실측 사실이지 규격 보장이 아니다.
    그래서 코드는 언제나 이 refcount 로 판정한다.
    """
    counts: dict[str, int] = {}
    for slide in deck.slide_order():
        for part in owned_parts(deck, slide):
            counts[part] = counts.get(part, 0) + 1
    return counts


def sole_owner_parts(deck: Deck, slide_parts: Iterable[str]) -> set[str]:
    """주어진 슬라이드들만 쓰는 파트 = 그 슬라이드와 함께 지워도 안전한 것."""
    doomed = list(slide_parts)
    doomed_set = set(doomed)
    survivors = [s for s in deck.slide_order() if s not in doomed_set]

    doomed_reach: set[str] = set()
    for s in doomed:
        doomed_reach |= owned_parts(deck, s)
    keep: set[str] = protected_parts(deck)
    for s in survivors:
        keep |= owned_parts(deck, s)
    return doomed_reach - keep


# --------------------------------------------------------------------------- #
# 패키지 위생
# --------------------------------------------------------------------------- #

def remove_part(deck: Deck, name: str) -> bool:
    """파트를 지운다. deck.parts / _tree_cache / _dirty 세 곳을 동시에 비운다.

    ★ parts 에서만 지우면 Deck.flush() 가 _dirty 에 남은 이름을 _tree_cache 에서
      되살려 저장 시 부활한다. 그러면 sldIdLst 에 없는 고아 슬라이드가 되어 파일만
      불어나고, 발견이 아주 늦다. 그 함정을 여기서 막는다.
    .rels 파트도 함께 지운다.
    """
    hit = False
    for target in (name, Deck.rels_name(name)):
        if target in deck.parts:
            deck.parts.pop(target, None)
            deck._tree_cache.pop(target, None)
            deck._dirty.discard(target)
            hit = True
    return hit


def sync_content_types(deck: Deck) -> dict[str, list[str]]:
    """[Content_Types].xml 을 현재 파트 목록에 맞춘다.

    · 파트가 사라진 Override 제거
    · Override 가 없는 slide/chart/colors/style/notesSlide 추가
    · media(.png/.jpeg/.emf/.svg) 와 워크북(.xlsx) 은 손대지 않는다 —
      Default 확장자가 이미 커버하므로 Override 를 만들면 오히려 규격 위반이다.

    반환 {'added': [...], 'removed': [...]} (둘 다 파트 이름, 선행 / 없음)
    """
    name = "[Content_Types].xml"
    tree = etree.fromstring(deck.parts[name])
    ovr_tag = q(CT_NS, "Override")

    have: dict[str, etree._Element] = {}
    removed: list[str] = []
    for el in list(tree):
        if el.tag != ovr_tag:
            continue
        part = (el.get("PartName") or "").lstrip("/")
        if part not in deck.parts:
            tree.remove(el)
            removed.append(part)
            continue
        have[part] = el

    added: list[str] = []
    for part in deck.parts:
        kind = _ct_kind(part)
        if kind is None or part in have:
            continue
        el = etree.SubElement(tree, ovr_tag)
        el.set("PartName", "/" + part)
        el.set("ContentType", CT_OVERRIDE[kind])
        added.append(part)

    if added or removed:
        deck.put_bytes(name, _xml_bytes(tree))
    return {"added": sorted(added), "removed": sorted(removed)}


def gc_orphans(deck: Deck, *, dry_run: bool = False) -> list[str]:
    """_rels/.rels 에서 도달하지 못하는 파트를 제거한다. 삭제 누락 뒤처리용 안전망.

    ★ PowerPoint 가 재저장하며 하는 GC 를 정답으로 삼지 않는다. PowerPoint 는 SVG 의
      PNG 래스터 폴백까지 버린다(실측: media 42 -> 30). 그건 PowerPoint 고유 최적화이지
      규격 요구가 아니다. 여기서는 **도달 가능한 것은 전부 유지**하는 보수적 GC 만 한다.
    """
    live = _reach(deck, _root_children(deck))
    live.add(ROOT_RELS)
    live.add("[Content_Types].xml")
    orphans = sorted(p for p in deck.parts if p not in live)
    if not dry_run:
        for p in orphans:
            remove_part(deck, p)
        if orphans:
            sync_content_types(deck)
    return orphans


# --------------------------------------------------------------------------- #
# presentation.xml / presentation.xml.rels 조작
# --------------------------------------------------------------------------- #

def _sld_id_lst(deck: Deck) -> etree._Element:
    tree = deck.xml(PRES)
    lst = tree.find(q(P, "sldIdLst"))
    if lst is None:
        raise RuntimeError("ppt/presentation.xml 에 <p:sldIdLst> 가 없다")
    return lst


def _pres_rid_map(deck: Deck) -> dict[str, str]:
    """rId -> 슬라이드 파트."""
    return {r["Id"]: r["Part"] for r in _rel_rows(deck, PRES)
            if r["Part"] and r["Type"].endswith("/slide")}


def _add_pres_rel(deck: Deck, target_part: str) -> str:
    """presentation.xml.rels 에 슬라이드 관계를 더하고 새 rId 를 준다."""
    tree = deck.xml(PRES_RELS)
    used = {el.get("Id") for el in tree}
    hi = 0
    for rid in used:
        m = re.match(r"^rId(\d+)$", rid or "")
        if m:
            hi = max(hi, int(m.group(1)))
    n = hi + 1
    while f"rId{n}" in used:
        n += 1
    rid = f"rId{n}"
    el = etree.SubElement(tree, q(PR, "Relationship"))
    el.set("Id", rid)
    el.set("Type", R + "/slide")
    el.set("Target", posixpath.relpath(target_part, posixpath.dirname(PRES)))
    deck.put(PRES_RELS, tree)
    return rid


def _drop_pres_rels(deck: Deck, parts: set[str]) -> list[str]:
    """presentation.xml.rels 에서 주어진 슬라이드 파트를 가리키는 관계를 지운다."""
    tree = deck.xml(PRES_RELS)
    base = posixpath.dirname(PRES)
    dropped: list[str] = []
    for el in list(tree):
        if el.get("TargetMode") == "External":
            continue
        tgt = _norm_target(base, el.get("Target") or "")
        if tgt in parts:
            tree.remove(el)
            dropped.append(el.get("Id") or "")
    if dropped:
        deck.put(PRES_RELS, tree)
    return dropped


def _new_sld_id(lst: etree._Element, reserved: set[int]) -> int:
    """규격 범위(256~2147483647) 안에서 안 겹치는 sldId/@id.

    실측 N6: id=100 을 쓰면 PowerPoint 가 파일을 거부한다. 하한 256 은 규격이다.
    """
    used = {int(s.get("id")) for s in lst if (s.get("id") or "").isdigit()} | reserved
    cand = max(used) + 1 if used else SLDID_MIN
    if cand < SLDID_MIN:
        cand = SLDID_MIN
    if cand > SLDID_MAX:                     # 상한을 넘겼으면 아래에서 빈 번호를 줍는다
        cand = SLDID_MIN
        while cand in used and cand <= SLDID_MAX:
            cand += 1
    if cand > SLDID_MAX or cand in used:
        raise RuntimeError("sldId/@id 를 더 뽑을 수 없다 (규격 상한 초과)")
    return cand


# --------------------------------------------------------------------------- #
# 삭제
# --------------------------------------------------------------------------- #

@dataclass
class DeleteReport:
    removed_slides: list[str] = field(default_factory=list)
    removed_parts: list[str] = field(default_factory=list)
    kept_shared: list[str] = field(default_factory=list)
    dropped_overrides: list[str] = field(default_factory=list)
    remaining: int = 0
    freed_bytes: int = 0

    def as_dict(self) -> dict[str, Any]:
        return {
            "removed_slides": self.removed_slides,
            "removed_parts": len(self.removed_parts),
            "kept_shared": self.kept_shared,
            "dropped_overrides": len(self.dropped_overrides),
            "remaining": self.remaining,
            "freed_bytes": self.freed_bytes,
        }


def delete_slides(deck: Deck, refs: Sequence[SlideRef], *,
                  gc: bool = True, dry_run: bool = False) -> DeleteReport:
    """슬라이드를 지운다. 공유되지 않는 종속 파트도 함께 정리한다.

    순서: sldIdLst 제거 -> presentation.xml.rels 제거 -> 슬라이드 파트+rels 제거
          -> gc 면 sole_owner 파트 제거 -> [Content_Types].xml 동기화

    gc=False 는 고아를 남긴다(디버그용). PowerPoint 는 열지만 파일이 30% 넘게 불고
    지운 장표의 원문·차트·엑셀이 파일 안에 남는다 — 납품물에는 절대 쓰지 마라.
    """
    doomed = _resolve_all(deck, refs)
    if not doomed:
        return DeleteReport(remaining=len(deck.slide_order()))

    order = deck.slide_order()
    doomed_set = set(doomed)
    remaining = [s for s in order if s not in doomed_set]
    if not remaining:
        raise ValueError("슬라이드를 전부 지울 수 없다 (0장 덱). keep_only 로 최소 1장은 남겨라")

    # 살아남는 슬라이드가 지울 슬라이드를 가리키고 있으면 REL_TARGET_MISSING 이 된다.
    # LG 덱에는 슬라이드 간 관계가 0건이지만, 지우고 나서 validate 로 발견하면 늦다.
    inbound: list[str] = []
    for s in remaining:
        for row in _rel_rows(deck, s):
            if row["Part"] in doomed_set:
                inbound.append(f"{s} -> {row['Part']} ({row['Id']})")
    if inbound:
        raise ValueError(
            "살아남는 슬라이드가 지울 슬라이드를 참조한다 (내부 점프 링크). "
            "먼저 그 관계를 끊어라:\n  " + "\n  ".join(inbound[:10]))

    # 무엇을 지울지 먼저 전부 계산한다 (dry_run 이 부작용 없이 같은 답을 주도록)
    doomed_reach: set[str] = set()
    for s in doomed:
        doomed_reach |= owned_parts(deck, s)
    keep: set[str] = protected_parts(deck)
    for s in remaining:
        keep |= owned_parts(deck, s)

    if gc:
        to_remove = doomed_reach - keep
    else:
        to_remove = set()
        for s in doomed:
            to_remove.add(s)
            rn = Deck.rels_name(s)
            if rn in deck.parts:
                to_remove.add(rn)

    rep = DeleteReport(
        removed_slides=list(doomed),
        removed_parts=sorted(to_remove),
        kept_shared=sorted(doomed_reach & keep),
        remaining=len(remaining),
        freed_bytes=sum(len(deck.parts[p]) for p in to_remove if p in deck.parts),
    )
    if dry_run:
        rep.dropped_overrides = sorted(p for p in to_remove if _ct_kind(p))
        return rep

    # 1) sldIdLst
    lst = _sld_id_lst(deck)
    rid_map = _pres_rid_map(deck)
    for sld in list(lst):
        if rid_map.get(sld.get(q(R, "id"))) in doomed_set:
            lst.remove(sld)
    deck.put(PRES, deck.xml(PRES))

    # 2) presentation.xml.rels
    _drop_pres_rels(deck, doomed_set)

    # 3) 파트 제거
    for p in sorted(to_remove):
        remove_part(deck, p)

    # 4) Content_Types
    ct = sync_content_types(deck)
    rep.dropped_overrides = ct["removed"]
    return rep


def keep_only(deck: Deck, refs: Sequence[SlideRef], *,
              gc: bool = True, dry_run: bool = False) -> DeleteReport:
    """refs 를 제외한 전부를 지운다. **표시 순서는 바꾸지 않는다**(원래 순서 유지).

    refs 가 비면 ValueError — 0장 덱은 PowerPoint 가 열긴 하지만 사고다.
    """
    if not refs:
        raise ValueError("keep_only 에 빈 목록을 줄 수 없다 (0장 덱)")
    keepers = set(_resolve_all(deck, refs))
    doomed = [s for s in deck.slide_order() if s not in keepers]
    if not doomed:
        return DeleteReport(remaining=len(deck.slide_order()))
    return delete_slides(deck, doomed, gc=gc, dry_run=dry_run)


# --------------------------------------------------------------------------- #
# 복제
# --------------------------------------------------------------------------- #

@dataclass
class CloneResult:
    slide_part: str
    position: int                                    # 1-based 표시 위치
    rid: str
    sld_id: int
    part_map: dict[str, str] = field(default_factory=dict)   # 옛 파트 -> 새 파트
    shared: list[str] = field(default_factory=list)           # 복제 않고 공유한 파트

    @property
    def charts(self) -> dict[str, str]:
        """옛 차트 파트 -> 새 차트 파트. fill_deck 가 차트를 찾아갈 때 쓴다."""
        return {k: v for k, v in self.part_map.items() if re.match(r"^ppt/charts/chart\d+\.xml$", k)}


def _clone_private(deck: Deck, src: str, namer: _Namer,
                   part_map: dict[str, str], shared: set[str]) -> str:
    """PRIVATE_ROOTS 파트 하나를 새 이름으로 뜬다. 그 파트의 rels 도 재귀 복제한다."""
    if src in part_map:
        return part_map[src]
    dst = namer.take_like(src)
    part_map[src] = dst
    deck.put_bytes(dst, deck.parts[src])
    _clone_rels(deck, src, dst, namer, part_map, shared)
    return dst


def _clone_rels(deck: Deck, src: str, dst: str, namer: _Namer,
                part_map: dict[str, str], shared: set[str]) -> None:
    """src 의 .rels 를 dst 용으로 복제한다.

    · part_map 에 이미 있는 대상 -> 그 새 파트로 다시 가리킨다
      (notesSlide 가 원본 슬라이드를 되가리키는 관계를 여기서 바로잡는다)
    · PRIVATE_ROOTS 대상        -> 새로 떠서 가리킨다 (차트/워크북/노트)
    · 그 외(media·slideLayout)  -> Target 그대로 공유
    · External(mailto:/https:)  -> 손대지 않는다
    """
    src_rels = Deck.rels_name(src)
    if src_rels not in deck.parts:
        return
    tree = etree.fromstring(deck.parts[src_rels])
    src_dir = posixpath.dirname(src)
    dst_dir = posixpath.dirname(dst)

    for rel in tree:
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target") or ""
        part = _norm_target(src_dir, target)
        mapped = part_map.get(part)
        if mapped is None and PRIVATE_ROOTS.match(part) and part in deck.parts:
            mapped = _clone_private(deck, part, namer, part_map, shared)
        if mapped is None:
            shared.add(part)
        final = mapped or part
        rel.set("Target", posixpath.relpath(final, dst_dir) if dst_dir else final)

    deck.put_bytes(Deck.rels_name(dst), _xml_bytes(tree))


def clone_slide(deck: Deck, ref: SlideRef, *,
                at: int | None = None, count: int = 1) -> list[CloneResult]:
    """슬라이드를 복제한다. 새 CloneResult 목록(각각 1-based 표시 위치)을 준다.

    at : **0-based** sldIdLst 삽입 인덱스. None 이면 맨 뒤.
         count>1 이면 at, at+1, at+2 ... 로 연달아 꽂힌다.
         (다른 API 의 위치 인자는 1-based 다 — 여기만 sldIdLst 인덱스라 0-based 다.
          move_slide(to=) 는 1-based 이므로 헷갈리지 마라.)

    차트가 N개면 chart/colors/style/워크북을 각각 N벌 뜬다. colors/style 이 없는
    차트(LG chart10~24)는 rels 에 없으니 애초에 복제 대상이 아니다.
    media 와 slideLayout 은 Target 그대로 공유한다.

    ★ 같은 슬라이드 파트를 sldIdLst 에 두 번 올리는 '복제 흉내'는 불가능하다.
      서로 다른 sldId 2개가 같은 rId 를 가리키면 PowerPoint 가 파일을 거부한다(실측 N7).
      그래서 복제는 반드시 실제 파트 복사다.
    """
    if count < 1:
        raise ValueError(f"count 는 1 이상: {count}")
    src = resolve(deck, ref)

    lst = _sld_id_lst(deck)
    n_now = len(lst)
    if at is None:
        insert_at = n_now
    else:
        insert_at = at if at >= 0 else n_now + 1 + at
        if not 0 <= insert_at <= n_now:
            raise IndexError(f"삽입 위치 {at} (현재 {n_now}장, 0-based 0..{n_now})")

    namer = _Namer(deck)
    reserved: set[int] = set()
    results: list[CloneResult] = []

    for i in range(count):
        part_map: dict[str, str] = {}
        shared: set[str] = set()

        new_slide = namer.take("slide")
        part_map[src] = new_slide                     # notesSlide 되가리킴 대비
        deck.put_bytes(new_slide, deck.parts[src])
        _clone_rels(deck, src, new_slide, namer, part_map, shared)

        rid = _add_pres_rel(deck, new_slide)
        lst = _sld_id_lst(deck)
        sld_id = _new_sld_id(lst, reserved)
        reserved.add(sld_id)
        el = etree.Element(q(P, "sldId"))
        el.set("id", str(sld_id))
        el.set(q(R, "id"), rid)
        lst.insert(insert_at + i, el)
        deck.put(PRES, deck.xml(PRES))

        results.append(CloneResult(
            slide_part=new_slide,
            position=insert_at + i + 1,
            rid=rid,
            sld_id=sld_id,
            part_map=dict(part_map),
            shared=sorted(shared),
        ))

    sync_content_types(deck)                          # 배치 끝에 한 번만
    # 앞에 끼워 넣으면 뒤 슬라이드 위치가 밀린다 -> 최종 위치로 다시 읽는다
    pos = positions(deck)
    for res in results:
        res.position = pos[res.slide_part]
    return results


# --------------------------------------------------------------------------- #
# 재배열
# --------------------------------------------------------------------------- #

def reorder(deck: Deck, order: Sequence[SlideRef]) -> list[str]:
    """<p:sldId> 요소를 통째로 재배치한다. order 는 현재 전체 슬라이드의 순열이어야 한다.

    ★ r:id 속성만 치환하는 짓은 하지 않는다. id <-> rId 쌍이 어긋나면 sldId/@id 의
      의미가 뒤집힌다.
    """
    want = _resolve_all(deck, order)
    current = deck.slide_order()
    if len(want) != len(current) or set(want) != set(current):
        missing = [p for p in current if p not in set(want)]
        extra = [p for p in want if p not in set(current)]
        raise ValueError(
            f"order 는 현재 슬라이드 {len(current)}장의 순열이어야 한다 "
            f"(빠짐 {len(missing)}: {missing[:3]} / 넘침 {len(extra)}: {extra[:3]})")

    lst = _sld_id_lst(deck)
    rid_map = _pres_rid_map(deck)
    by_part: dict[str, etree._Element] = {}
    for sld in list(lst):
        part = rid_map.get(sld.get(q(R, "id")))
        if part:
            by_part[part] = sld
        lst.remove(sld)
    for part in want:
        lst.append(by_part[part])
    deck.put(PRES, deck.xml(PRES))
    return want


def move_slide(deck: Deck, ref: SlideRef, to: int) -> int:
    """슬라이드 하나를 1-based 목표 위치로 옮긴다. 새 위치를 준다."""
    part = resolve(deck, ref)
    order = deck.slide_order()
    n = len(order)
    if not 1 <= to <= n:
        raise IndexError(f"목표 위치 {to} (총 {n}장)")
    rest = [p for p in order if p != part]
    rest.insert(to - 1, part)
    reorder(deck, rest)
    return to


def apply_plan(deck: Deck, plan: Sequence[SlideRef], *, gc: bool = True) -> list[str]:
    """plan 이 현재 덱의 부분집합이면 나머지를 지운 뒤 plan 순서로 배치한다.

    keep_only + reorder 를 한 번에. plan 은 **먼저 파트 이름으로 굳혀** 두므로
    삭제로 위치가 밀려도 안전하다.
    """
    want = _resolve_all(deck, plan)
    if not want:
        raise ValueError("apply_plan 에 빈 계획을 줄 수 없다")
    keep_only(deck, want, gc=gc)
    return reorder(deck, want)


# --------------------------------------------------------------------------- #
# docProps/app.xml
# --------------------------------------------------------------------------- #

_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


def _slide_title(deck: Deck, slide_part: str) -> str:
    """슬라이드 제목 개체틀 텍스트. 없으면 PowerPoint 기본 문자열."""
    try:
        tree = deck.xml(slide_part)
    except Exception:
        return "PowerPoint 프레젠테이션"
    for sp in tree.iter(q(P, "sp")):
        ph = sp.find(f".//{q(P, 'nvSpPr')}/{q(P, 'nvPr')}/{q(P, 'ph')}")
        if ph is None or ph.get("type") not in ("title", "ctrTitle"):
            continue
        txt = "".join(t.text or "" for t in sp.iter(q(A, "t"))).strip()
        if txt:
            return txt
    return "PowerPoint 프레젠테이션"


def refresh_app_xml(deck: Deck, *, drop_fonts: Sequence[str] = ()) -> None:
    """docProps/app.xml 을 현재 덱 기준으로 다시 쓴다.

    PowerPoint 는 stale 한 app.xml 을 무시하고 정상 오픈하지만, **파일 속성창과
    탐색기 미리보기에 틀린 장수와 남의 회사 글꼴 이름이 뜬다.** 납품물에 LG 흔적을
    남기지 않으려면 반드시 돌려야 한다.

    drop_fonts 로 'LG EI Text TTF Regular' 같은 잔재를 글꼴 목록에서 뺀다.
    """
    if APP_XML not in deck.parts:
        return
    tree = etree.fromstring(deck.parts[APP_XML])
    order = deck.slide_order()

    # 기존 HeadingPairs 에서 글꼴 수 / 테마 수를 읽어 TitlesOfParts 를 잘라낸다
    pairs = tree.find(q(_EP, "HeadingPairs"))
    titles = tree.find(q(_EP, "TitlesOfParts"))
    fonts: list[str] = []
    themes: list[str] = []
    if pairs is not None and titles is not None:
        vec = pairs.find(q(_VT, "vector"))
        tvec = titles.find(q(_VT, "vector"))
        if vec is not None and tvec is not None:
            names = [(e.text or "") for e in tvec.iter(q(_VT, "lpstr"))]
            variants = list(vec.iter(q(_VT, "variant")))
            cursor = 0
            for i in range(0, len(variants) - 1, 2):
                label_el = variants[i].find(q(_VT, "lpstr"))
                count_el = variants[i + 1].find(q(_VT, "i4"))
                label = (label_el.text or "") if label_el is not None else ""
                cnt = int(count_el.text or 0) if count_el is not None else 0
                chunk = names[cursor:cursor + cnt]
                cursor += cnt
                if "글꼴" in label or "font" in label.lower():
                    fonts = chunk
                elif "테마" in label or "theme" in label.lower():
                    themes = chunk
    drop = set(drop_fonts)
    fonts = [f for f in fonts if f not in drop]
    slide_titles = [_slide_title(deck, s) for s in order]

    # 스칼라 카운트
    def _set(tag: str, value: str) -> None:
        el = tree.find(q(_EP, tag))
        if el is None:
            el = etree.SubElement(tree, q(_EP, tag))
        el.text = value

    n_par = 0
    n_word = 0
    for s in order:
        try:
            st = deck.xml(s)
        except Exception:
            continue
        n_par += sum(1 for _ in st.iter(q(A, "p")))
        for t in st.iter(q(A, "t")):
            n_word += len((t.text or "").split())
    _set("Slides", str(len(order)))
    _set("Paragraphs", str(n_par))
    _set("Words", str(n_word))
    _set("Notes", str(len(deck.names(r"^ppt/notesSlides/notesSlide\d+\.xml$"))))

    groups: list[tuple[str, list[str]]] = []
    if fonts:
        groups.append(("사용한 글꼴", fonts))
    if themes:
        groups.append(("테마", themes))
    groups.append(("슬라이드 제목", slide_titles))

    if pairs is not None:
        for ch in list(pairs):
            pairs.remove(ch)
    else:
        pairs = etree.SubElement(tree, q(_EP, "HeadingPairs"))
    vec = etree.SubElement(pairs, q(_VT, "vector"))
    vec.set("size", str(len(groups) * 2))
    vec.set("baseType", "variant")
    for label, items in groups:
        v1 = etree.SubElement(vec, q(_VT, "variant"))
        etree.SubElement(v1, q(_VT, "lpstr")).text = label
        v2 = etree.SubElement(vec, q(_VT, "variant"))
        etree.SubElement(v2, q(_VT, "i4")).text = str(len(items))

    if titles is not None:
        for ch in list(titles):
            titles.remove(ch)
    else:
        titles = etree.SubElement(tree, q(_EP, "TitlesOfParts"))
    flat = [x for _, items in groups for x in items]
    tvec = etree.SubElement(titles, q(_VT, "vector"))
    tvec.set("size", str(len(flat)))
    tvec.set("baseType", "lpstr")
    for x in flat:
        etree.SubElement(tvec, q(_VT, "lpstr")).text = x

    deck.put_bytes(APP_XML, _xml_bytes(tree))


# --------------------------------------------------------------------------- #
# 검증
# --------------------------------------------------------------------------- #

#: 슬라이드/레이아웃/마스터가 **반드시** 가져야 하는 관계. 하나라도 빠지면 PowerPoint 가
#: 0x80070570 으로 파일을 거부한다(4종 전수 실측, 2026-08-22).
REQUIRED_RELS: tuple[tuple[str, str], ...] = (
    (r"^ppt/presentation\.xml$",               "/slideMaster"),
    (r"^ppt/slides/slide\d+\.xml$",            "/slideLayout"),
    (r"^ppt/slideLayouts/slideLayout\d+\.xml$", "/slideMaster"),
    (r"^ppt/slideMasters/slideMaster\d+\.xml$", "/theme"),
)

#: 관계 네임스페이스 접두. 이 네임스페이스의 **모든** 속성은 예외 없이 rels 의 Id 를 가리킨다
#: (id·embed·link·pict·href·dm·lo·qs·cs). 그래서 이름을 열거하지 않고 네임스페이스로 잡는다.
_R_BRACE = "{" + R + "}"

#: 여기 걸린 r:id 가 끊기면 PowerPoint 는 파일을 열되 그 자리를 오류 상자로 그린다.
#: 나머지(c:chart·c:externalData 등)는 파일 자체를 거부한다. 실측 기준.
_RENDER_ONLY_TAGS = {"blip", "svgBlip"}


def _read_tree(deck: Deck, name: str) -> etree._Element | None:
    """캐시에 트리가 있으면 그것을, 없으면 바이트를 파싱해서 준다. **캐시를 더럽히지 않는다.**

    deck.xml() 을 쓰면 검사만 했는데 덱의 모든 XML 파트가 트리로 캐시에 눌러앉는다
    (99장 덱에서 100MB+). 반대로 deck.parts 만 읽으면 put() 안 하고 트리를 제자리에서
    고친 코드의 수정분을 놓친다. 둘 다 피하려고 캐시 우선 · 없으면 임시 파싱으로 간다.
    """
    tree = deck._tree_cache.get(name)
    if tree is not None:
        return tree
    try:
        return etree.fromstring(deck.parts[name])
    except (etree.XMLSyntaxError, KeyError):
        return None


def _part_ext(part: str) -> str:
    """파트 이름의 확장자(소문자, 점 없음). OPC 규격 = **마지막 마침표 뒤 조각**.

    ★ posixpath.splitext 를 쓰면 안 된다. '_rels/.rels' 를 유닉스 숨김 파일로 보고
      확장자를 '' 로 준다. 그러면 패키지의 모든 .rels 가 '무타입' 으로 오탐된다
      (실제로 물렸다). PowerPoint 는 <Default Extension="rels"> 로 이들을 덮는다.
    """
    base = part.rsplit("/", 1)[-1]
    return base.rsplit(".", 1)[-1].lower() if "." in base else ""


@dataclass
class Problem:
    level: Literal["fatal", "warn"]
    code: str
    part: str | None
    detail: str
    #: PowerPoint 가 실제로 어떻게 반응하는가. 등급의 근거이지 등급 자체가 아니다.
    #:   reject        파일을 거부한다 (복구 대화상자 / HRESULT 오류)
    #:   silent-loss   열리는데 내용이 **경고 없이** 사라진다
    #:   broken-render 열리는데 그 자리가 오류 상자나 빈칸이 된다
    #:   spec          규격 위반이지만 PowerPoint 는 정상 처리한다
    effect: Literal["reject", "silent-loss", "broken-render", "spec"] = "reject"

    def __str__(self) -> str:
        return f"[{self.level}/{self.effect}] {self.code} {self.part or ''} :: {self.detail}"


def validate(deck: Deck) -> list[Problem]:
    """패키지 무결성 점검. **fatal = 저장을 막아야 하는 것.**

    fatal 의 뜻은 'PowerPoint 가 거부한다'가 아니라 **'이대로 납품하면 안 된다'**이다.
    거부당하는 것과 조용히 내용을 잃는 것은 똑같이 막아야 한다 — 오히려 후자가 더 늦게
    발각돼서 위험하다. 어느 쪽인지는 Problem.effect 에 남긴다.

    아래 등급은 전부 **파손본을 만들어 PowerPoint COM 으로 열어 본 실측**이다
    (2026-08-22, 세팅본 22장에 19종 주입 · _selftest [7] 이 16종을 회귀 테스트로 붙들고
    있고, [7b] 는 'PowerPoint 가 여는 것을 fatal 로 올리지 않았는가'를 반대로 지킨다).

      코드                     PowerPoint 실동작              effect
      ---------------------------------------------------------------------------
      SLDID_DEAD_RID           거부 0x80070570                reject
      SLDID_DUP_ID             거부 0x80070570                reject
      SLDID_ID_RANGE           거부 (id=100 실측)             reject
      SLDID_DUP_RID            거부 0x80070570                reject
      SLIDE_PART_MISSING       거부 0x80070570                reject
      REL_TARGET_MISSING       거부 0x80070570                reject
      CT_UNTYPED_PART          거부 0x80CB8002                reject        ★신규
      CT_DUP_OVERRIDE          거부 0x80CB8001                reject        ★신규
      CT_DUP_DEFAULT           거부 0x80CB8000                reject        ★신규
      REL_REQUIRED_MISSING     거부 0x80070570                reject        ★신규
      REL_ID_DANGLING          맥락별(아래 주석)              reject/broken-render ★신규
      SLIDE_NOT_IN_SLDIDLST    연다 — 그 장이 사라진다        silent-loss   ★신규
      ---------------------------------------------------------------------------
      CT_OVERRIDE_ORPHAN       연다 22/22 정상                spec     warn
      CT_OVERRIDE_MISSING      연다 22/22 정상 — 과잉 아님    spec     warn
      PART_ORPHAN              연다 정상 (파일만 분다)        spec     warn
      APP_XML_STALE            연다 정상                      spec     warn

    ★신규 5종은 전부 **이전 판이 fatal 도 warn 도 없이 통과시키던** 파손이다.
    특히 SLIDE_NOT_IN_SLDIDLST('유령 슬라이드')는 파트·rels·Override 가 전부 살아 있어
    어떤 참조 검사에도 걸리지 않는다. PowerPoint 는 22장짜리 덱을 21장으로 조용히 연다.

    REL_ID_DANGLING 의 맥락별 실측 — **일부러 넓게 fatal 로 잡는다**:
      <c:chart r:id>        거부          납품 불가
      <c:externalData r:id> 거부          납품 불가
      <a:blip r:embed>      연다 → 그 자리가 「그림을 표시할 수 없습니다」 상자. 표지 로고가 그랬다
      <a:svgBlip r:embed>   연다 → 렌더 동일 (래스터 폴백이 받아 준다). **여기만 과잉 등급**
    svgBlip 하나 때문에 등급을 낮추지 않는 이유: 댕글링 r:id 가 의도적으로 옳은 경우는
    없다. 실덱 2종(세팅본 22장 · LG 99장)에서 733건을 전수 검사해 댕글링 0건이므로
    과잉 등급의 비용은 0이고, 놓쳤을 때의 비용은 납품 거부다.
    """
    problems: list[Problem] = []

    # 루트에서 도달 가능한 파트. 고아 파트는 PowerPoint 가 아예 읽지 않으므로 아래
    # 구조 검사(필수 관계 · 본문 r:id)의 대상에서 뺀다. 고아라는 사실 자체는
    # PART_ORPHAN warn 이 따로 보고한다. 여기서 안 빼면 지워야 할 고아 슬라이드가
    # '레이아웃 관계 없음' fatal 로 둔갑해 원인을 가린다.
    live = _reach(deck, _root_children(deck)) | {ROOT_RELS, "[Content_Types].xml"}

    # ---- sldIdLst ----
    rid_map = {r["Id"]: r["Part"] for r in _rel_rows(deck, PRES)}
    slide_rids = {r["Id"] for r in _rel_rows(deck, PRES) if r["Type"].endswith("/slide")}
    try:
        lst = _sld_id_lst(deck)
    except RuntimeError as exc:
        problems.append(Problem("fatal", "SLDID_DEAD_RID", PRES, str(exc)))
        lst = []

    seen_ids: set[str] = set()
    seen_rids: set[str] = set()
    for sld in lst:
        sid = sld.get("id") or ""
        rid = sld.get(q(R, "id")) or ""
        if sid in seen_ids:
            problems.append(Problem("fatal", "SLDID_DUP_ID", PRES, f"sldId/@id 중복: {sid}"))
        seen_ids.add(sid)
        if not sid.isdigit() or not (SLDID_MIN <= int(sid) <= SLDID_MAX):
            problems.append(Problem("fatal", "SLDID_ID_RANGE", PRES,
                                    f"sldId/@id={sid} 는 {SLDID_MIN}~{SLDID_MAX} 밖"))
        if rid in seen_rids:
            problems.append(Problem("fatal", "SLDID_DUP_RID", PRES,
                                    f"두 sldId 가 같은 rId 를 가리킴: {rid}"))
        seen_rids.add(rid)
        if rid not in rid_map or rid not in slide_rids:
            problems.append(Problem("fatal", "SLDID_DEAD_RID", PRES,
                                    f"r:id={rid} 가 presentation.xml.rels 에 없다"))
        else:
            part = rid_map[rid]
            if part not in deck.parts:
                problems.append(Problem("fatal", "SLIDE_PART_MISSING", part,
                                        f"sldId r:id={rid} 가 없는 파트를 가리킨다"))

    # ---- 유령 슬라이드 ★ ----
    # presentation.xml.rels 에는 슬라이드 관계가 있는데 sldIdLst 가 안 가리키는 파트.
    # 참조는 전부 성립하므로 위의 어떤 검사에도 안 걸린다. PowerPoint 는 경고 없이 뺀다.
    for rid in sorted(slide_rids - seen_rids):
        part = rid_map.get(rid)
        if part and part in deck.parts:
            problems.append(Problem(
                "fatal", "SLIDE_NOT_IN_SLDIDLST", part,
                f"rels 에 {rid} 로 살아 있는데 sldIdLst 가 안 가리킨다 "
                f"— PowerPoint 는 경고 없이 이 장을 빼고 연다",
                effect="silent-loss"))

    # ---- 모든 관계의 대상 존재 ----
    rels_parts = [n for n in deck.parts if n.endswith(".rels")]
    for rels_name in rels_parts:
        owner_dir = posixpath.dirname(posixpath.dirname(rels_name))
        for rel in etree.fromstring(deck.parts[rels_name]):
            if rel.get("TargetMode") == "External":
                continue
            tgt = _norm_target(owner_dir, rel.get("Target") or "")
            if tgt not in deck.parts:
                code = "SLIDE_PART_MISSING" if tgt.startswith("ppt/slides/") else "REL_TARGET_MISSING"
                problems.append(Problem("fatal", code, rels_name,
                                        f"{rel.get('Id')} -> {tgt} (없는 파트)"))

    # ---- 필수 관계 ★ ----
    # 대상이 살아 있는지가 아니라 **관계 자체가 있는지**를 본다. 관계를 통째로 지우면
    # 댕글링이 안 생겨서 위의 검사가 전부 통과해 버린다.
    for rx, suffix in REQUIRED_RELS:
        crx = re.compile(rx)
        for part in sorted(p for p in deck.parts if p in live and crx.match(p)):
            if not any(r["Type"].endswith(suffix) for r in _rel_rows(deck, part)):
                problems.append(Problem("fatal", "REL_REQUIRED_MISSING", part,
                                        f"필수 관계 {suffix} 가 없다"))

    # ---- 본문이 가리키는 r:id 가 rels 에 있는가 ★ ----
    for name in sorted(live):
        if not name.endswith(".xml") or name.endswith(".rels") or name not in deck.parts:
            continue
        root = _read_tree(deck, name)
        if root is None:
            continue
        ids = {r["Id"] for r in _rel_rows(deck, name)}
        for el in root.iter():
            if not isinstance(el.tag, str):
                continue
            for key, val in el.attrib.items():
                # 빈 값은 정상이다. <a:blip r:embed=""> 는 '그림 없음'을 뜻한다.
                if not val or not key.startswith(_R_BRACE) or val in ids:
                    continue
                tag = etree.QName(el).localname
                problems.append(Problem(
                    "fatal", "REL_ID_DANGLING", name,
                    f"<{tag} {etree.QName(key).localname}=\"{val}\"> "
                    f"— {Deck.rels_name(name)} 에 그 Id 가 없다",
                    effect="broken-render" if tag in _RENDER_ONLY_TAGS else "reject"))

    # ---- OOXML 자식 순서 · 식별자 중복 ★ ----
    # PowerPoint 가 **조용히 무시하는** 두 종류다. 파일은 열리고, 관계 그래프도 멀쩡하고,
    # XML 만 들여다보면 완벽하다. COM 으로 표를 되읽어야 드러난다.
    #   (가) extLst / endParaRPr 뒤에 형제를 붙인 것 — 그 요소가 통째로 무시된다
    #   (나) 같은 tblGrid 안에 a16:colId 중복 — PowerPoint 가 한 열로 취급한다
    #   (다) 누적 막대에 dLblPos="outEnd" — 규격 위반이라 파일이 아예 안 열린다
    # (가)(나)는 '표를 넓혔는데 늘린 열이 화면에서 사라지는' 형태다(실측 2026-08-22).
    # (다)는 반대로 **시끄럽게** 죽지만 원인이 멀다 — 차트 종류만 바꿨는데
    #   원본에 있던 레이블 위치가 갑자기 불법이 된다(실측 2026-08-24).
    for name in sorted(live):
        if not name.startswith("ppt/") or not name.endswith(".xml") or name not in deck.parts:
            continue
        root = _read_tree(deck, name)
        if root is None:
            continue
        for el in root.iter():
            if not isinstance(el.tag, str):
                continue
            kids = [k for k in el if isinstance(k.tag, str)]
            for i, k in enumerate(kids[:-1]):
                local = etree.QName(k).localname
                if local in ("extLst", "endParaRPr"):
                    problems.append(Problem(
                        "fatal", "XML_CHILD_ORDER", name,
                        f"<{etree.QName(el).localname}> 안에서 <{local}> 뒤에 "
                        f"<{etree.QName(kids[i + 1]).localname}> 가 온다 "
                        f"— PowerPoint 는 뒤따르는 형제를 무시한다",
                        effect="silent-loss"))
                    break
        # (라) 문단이 하나도 없는 txBody — 불법 OOXML. lxml 정형성·패키지 검사는 다
        #      통과하는데 PowerPoint 는 **파일 전체를 거부**한다
        #      (실측 2026-08-28: 빈 항목 목록을 write_lines 로 주입해 간지가 이 꼴이 됐고,
        #       COM 이분 탐색으로야 찾았다. 여기서 잡으면 저장 단계에서 바로 죽는다).
        for tb in root.iter():
            if not (isinstance(tb.tag, str) and etree.QName(tb).localname == "txBody"):
                continue
            if not any(isinstance(k.tag, str) and etree.QName(k).localname == "p" for k in tb):
                problems.append(Problem(
                    "fatal", "TXBODY_EMPTY", name,
                    "txBody 에 <a:p> 가 하나도 없다 — PowerPoint 가 파일을 거부한다",
                    effect="reject"))
        for grid in root.iter(q(A, "tblGrid")):
            seen_ids: dict[str, int] = {}
            for e in grid.iter():
                if isinstance(e.tag, str) and etree.QName(e).localname == "colId":
                    v = e.get("val")
                    if v:
                        seen_ids[v] = seen_ids.get(v, 0) + 1
            dups = sorted(v for v, n in seen_ids.items() if n > 1)
            if dups:
                problems.append(Problem(
                    "fatal", "TBL_DUP_COLID", name,
                    f"tblGrid 안에 colId 가 겹친다: {dups[:4]} "
                    f"— PowerPoint 가 같은 열로 보고 내용을 버린다",
                    effect="silent-loss"))
        if name.startswith("ppt/charts/chart"):
            # (다) 누적 막대인데 데이터 레이블이 막대 **바깥끝**(outEnd)에 붙어 있다.
            #      막대가 이어 붙는 그림에 '바깥끝'이라는 자리가 없어서 규격 위반이고,
            #      PowerPoint 는 무시하지 않고 **파일 전체를 거부**한다
            #      (실측 2026-08-24: clustered 로만 바꾸면 같은 파일이 열린다).
            #      가로막대 슬롯을 누적으로 전용할 때 반드시 밟는 함정이다 —
            #      바깥끝은 clustered 에서는 정상값이라 원본에 그대로 들어 있다.
            for plot in root.iter(q(C, "barChart")):
                g = plot.find(q(C, "grouping"))
                if g is None or g.get("val") not in ("stacked", "percentStacked"):
                    continue
                bad = [e for e in plot.iter(q(C, "dLblPos")) if e.get("val") == "outEnd"]
                if bad:
                    problems.append(Problem(
                        "fatal", "CHART_DLBLPOS_STACKED", name,
                        f"누적 막대({g.get('val')})에 dLblPos=\"outEnd\" 가 {len(bad)}개 있다 "
                        f"— PowerPoint 가 파일을 열지 못한다. ctr·inEnd·inBase 중 하나여야 한다",
                        effect="reject"))

    # ---- Content_Types ----
    # 규격상 Extension 도 PartName 도 비교는 대소문자를 무시한다. 그래서 키는 소문자로
    # 잡고, 사람에게 보여 줄 이름만 원래 표기로 따로 들고 있는다.
    ct = etree.fromstring(deck.parts["[Content_Types].xml"])
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}          # 소문자 파트 이름 -> ContentType
    ovr_shown: dict[str, str] = {}          # 소문자 파트 이름 -> 원래 표기
    for el in ct:
        local = etree.QName(el).localname
        if local == "Default":
            ext = (el.get("Extension") or "").lower()
            if ext in defaults:
                problems.append(Problem("fatal", "CT_DUP_DEFAULT", "[Content_Types].xml",
                                        f"Default Extension=\"{ext}\" 가 2번 이상 나온다"))
            defaults[ext] = el.get("ContentType") or ""
        elif local == "Override":
            part = (el.get("PartName") or "").lstrip("/")
            if part.lower() in overrides:
                problems.append(Problem("fatal", "CT_DUP_OVERRIDE", part,
                                        f"Override PartName=\"/{part}\" 가 2번 이상 나온다"))
            overrides[part.lower()] = el.get("ContentType") or ""
            ovr_shown[part.lower()] = part

    # 모든 파트는 타입이 있어야 한다 — Override(정확 이름) 아니면 Default(확장자).
    # 어느 쪽도 없으면 PowerPoint 는 0x80CB8002 로 거부한다. Default 하나만 지워도
    # 그 확장자를 쓰는 파트가 전부 무타입이 되므로 media 수십 개가 한꺼번에 뜬다.
    untyped: list[str] = []
    for part in sorted(deck.parts):
        # 디렉터리 엔트리('ppt/media/')는 파트가 아니다. LG 덱·세팅본에는 0건이지만
        # 제작팀 회신본이 다른 도구로 저장돼 오면 들어 있을 수 있다.
        if part.endswith("/") or part == "[Content_Types].xml" or part.lower() in overrides:
            continue
        ext = _part_ext(part)
        if ext not in defaults:
            untyped.append(part)
    for part in untyped[:20]:
        ext = _part_ext(part)
        problems.append(Problem("fatal", "CT_UNTYPED_PART", part,
                                f"Override 도 Default Extension=\"{ext}\" 도 없다"))
    if len(untyped) > 20:
        problems.append(Problem("fatal", "CT_UNTYPED_PART", None,
                                f"…무타입 파트 {len(untyped)}개 중 20개만 위에 적었다"))

    live_lower = {p.lower() for p in deck.parts}
    for key in sorted(set(overrides) - live_lower):
        problems.append(Problem("warn", "CT_OVERRIDE_ORPHAN", ovr_shown[key],
                                "Override 만 있고 파트가 없다", effect="spec"))
    for part in sorted(deck.parts):
        if _ct_kind(part) and part.lower() not in overrides:
            problems.append(Problem("warn", "CT_OVERRIDE_MISSING", part,
                                    "파트에 Override 가 없다 (xml Default 로 폴백 — PowerPoint 는 연다)",
                                    effect="spec"))

    # ---- 고아 파트 ----
    for part in sorted(p for p in deck.parts if p not in live):
        problems.append(Problem("warn", "PART_ORPHAN", part, "루트에서 도달 불가", effect="spec"))

    # ---- app.xml ----
    if APP_XML in deck.parts:
        m = re.search(rb"<Slides>(\d+)</Slides>", deck.parts[APP_XML])
        n = len(deck.slide_order())
        if m and int(m.group(1)) != n:
            problems.append(Problem("warn", "APP_XML_STALE", APP_XML,
                                    f"<Slides>{m.group(1).decode()}</Slides> != 실제 {n}장",
                                    effect="spec"))
    return problems


_PS_VERIFY = r"""
$ErrorActionPreference = 'Stop'
$res = @{ ok = $false; slides = 0; layouts = @(); exported = @(); error = '' }
$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($PPTX, $true, $false, $false)
    $n = $pres.Slides.Count
    $res.slides = $n
    $names = New-Object System.Collections.ArrayList
    if ($n -gt 0) {
        for ($i = 1; $i -le $n; $i++) {
            [void]$names.Add($pres.Slides.Item($i).CustomLayout.Name)
        }
    }
    $res.layouts = $names.ToArray()
    $exp = New-Object System.Collections.ArrayList
    foreach ($pos in $EXPORTS) {
        if ($pos -ge 1 -and $pos -le $n) {
            $png = Join-Path $OUTDIR ("slide_{0}.png" -f $pos)
            $pres.Slides.Item([int]$pos).Export($png, "PNG", 1600, 900)
            [void]$exp.Add($png)
        }
    }
    $res.exported = $exp.ToArray()
    $res.ok = $true
} catch {
    $res.error = $_.Exception.Message
} finally {
    if ($pres -ne $null) { try { $pres.Close() } catch {} }
    if ($ppt  -ne $null) { try { $ppt.Quit()   } catch {} }
}
$res | ConvertTo-Json -Depth 5 | Out-File -FilePath $JSONOUT -Encoding utf8
"""


def assert_openable(path: str | Path, *, export_positions: Sequence[int] = (),
                    timeout: int = 300) -> dict[str, Any]:
    """PowerPoint COM 으로 실제 열어 본다. 복구 대화상자 없이 열려야 통과다.

    zipfile 로 열린다고 PowerPoint 가 연다는 보장이 없다 — 관계 그래프가 깨지면
    PowerPoint 는 0x80070570 'corrupted and unreadable' 로 거부한다. validate() 가
    놓친 파손은 여기서 잡힌다.

    반환 {'slides': int, 'layouts': [...], 'exported': [png...]}
    실패면 RuntimeError.
    """
    pptx = Path(path).resolve()
    if not pptx.exists():
        raise FileNotFoundError(pptx)
    tmp = Path(tempfile.gettempdir()) / f"deckbuild_{uuid.uuid4().hex}"
    tmp.mkdir(parents=True, exist_ok=True)
    js = tmp / "result.json"
    ps = tmp / "verify.ps1"
    header = (
        f"$PPTX = '{pptx}'\n"
        f"$JSONOUT = '{js}'\n"
        f"$OUTDIR = '{tmp}'\n"
        f"$EXPORTS = @({','.join(str(int(p)) for p in export_positions) or ''})\n"
    )
    ps.write_text(header + _PS_VERIFY, encoding="utf-8-sig")
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
         "-File", str(ps)],
        capture_output=True, timeout=timeout,
    )
    if not js.exists():
        raise RuntimeError(
            "PowerShell COM 검증이 결과를 남기지 않았다.\n"
            f"stdout={proc.stdout[:800]!r}\nstderr={proc.stderr[:800]!r}")
    data = json.loads(js.read_text(encoding="utf-8-sig"))
    if not data.get("ok"):
        raise RuntimeError(f"PowerPoint 가 파일을 열지 못했다: {pptx}\n  {data.get('error')}")
    layouts = data.get("layouts") or []
    if isinstance(layouts, str):
        layouts = [layouts]
    exported = data.get("exported") or []
    if isinstance(exported, str):
        exported = [exported]
    return {"slides": int(data.get("slides") or 0),
            "layouts": list(layouts),
            "exported": list(exported),
            "workdir": str(tmp)}


def save_checked(deck: Deck, path: str | Path, *,
                 com_verify: bool = True,
                 drop_fonts: Sequence[str] = (),
                 export_positions: Sequence[int] = ()) -> Path:
    """sync_content_types -> refresh_app_xml -> validate(fatal 이면 예외) -> save -> COM 오픈.

    fatal 이 하나라도 있으면 저장 자체를 막는다. fatal 11종 중 9종은 PowerPoint 가
    파일을 거부하고, 2종(유령 슬라이드 · blip 댕글링)은 열리되 내용이 사라지거나
    그림 자리가 오류 상자가 된다. 어느 쪽이든 납품할 수 없는 파일이다.
    등급 근거는 validate() 독스트링의 실측표에 있다.

    ★ 이 함수는 gc_orphans 를 부르지 않는다. 고아 파트는 warn 이라 저장을 막지 않지만
      지운 슬라이드의 원문이 파일에 남는 정보 유출이다. 삭제 작업 뒤에는 호출자가
      gc_orphans 를 직접 부르거나 delete_slides 의 기본 gc 를 켠 채로 둘 것.
    """
    sync_content_types(deck)
    refresh_app_xml(deck, drop_fonts=drop_fonts)
    problems = validate(deck)
    fatal = [p for p in problems if p.level == "fatal"]
    if fatal:
        raise RuntimeError("fatal 문제가 있어 저장을 막는다:\n  " +
                           "\n  ".join(str(p) for p in fatal[:20]))
    out = deck.save(path)
    if com_verify:
        assert_openable(out, export_positions=export_positions)
    return out


# --------------------------------------------------------------------------- #
# 메타 조회 (세팅 담당용)
# --------------------------------------------------------------------------- #

def _layout_name(deck: Deck, layout_part: str) -> str:
    try:
        tree = deck.xml(layout_part)
    except Exception:
        return ""
    csld = tree.find(q(P, "cSld"))
    return (csld.get("name") or "") if csld is not None else ""


def _has_slidenum(deck: Deck, part: str) -> bool:
    for fld in deck.xml(part).iter(q(A, "fld")):
        if fld.get("type") == "slidenum":
            return True
    return False


def deck_map(deck: Deck) -> list[dict[str, Any]]:
    """조립 계획을 짤 때 보는 표.

    [{pos, part, layout_part, layout_name, title, n_charts, n_tables, has_slidenum}]
    ★ 복제 전에 n_charts 를 반드시 확인해라. LG slide10/11 은 차트가 15개라
      한 번 복제하면 새 파트가 61개 뜬다.
    """
    rows: list[dict[str, Any]] = []
    for i, part in enumerate(deck.slide_order(), start=1):
        layouts = [r["Part"] for r in _rel_rows(deck, part)
                   if r["Part"] and r["Type"].endswith("/slideLayout")]
        layout = layouts[0] if layouts else ""
        tree = deck.xml(part)
        rows.append({
            "pos": i,
            "part": part,
            "layout_part": layout,
            "layout_name": _layout_name(deck, layout) if layout else "",
            "title": _slide_title(deck, part),
            "n_charts": len([r for r in _rel_rows(deck, part)
                             if r["Part"] and r["Type"].endswith("/chart")]),
            "n_tables": sum(1 for _ in tree.iter(q(A, "tbl"))),
            "has_slidenum": _has_slidenum(deck, layout) if layout else False,
        })
    return rows


def layout_index(deck: Deck) -> dict[str, dict[str, Any]]:
    """레이아웃 파트 -> {name, users(장수), has_slidenum}.

    has_slidenum=False 인 레이아웃(LG: cover/contents/section/contact us/빈 화면)에는
    페이지 번호가 안 찍힌다. 원본 덱의 설계이지 버그가 아니다.
    """
    users: dict[str, int] = {}
    for slide in deck.slide_order():
        for r in _rel_rows(deck, slide):
            if r["Part"] and r["Type"].endswith("/slideLayout"):
                users[r["Part"]] = users.get(r["Part"], 0) + 1
    out: dict[str, dict[str, Any]] = {}
    for part in sorted(deck.names(r"^ppt/slideLayouts/slideLayout\d+\.xml$"),
                       key=lambda n: int(re.findall(r"(\d+)", n)[-1])):
        out[part] = {"name": _layout_name(deck, part),
                     "users": users.get(part, 0),
                     "has_slidenum": _has_slidenum(deck, part)}
    return out


# --------------------------------------------------------------------------- #
# 이번 세팅 시나리오 (얇은 래퍼)
# --------------------------------------------------------------------------- #

#: LG 99장에서 남길 표시 위치.
#: 1 표지 / 2 목차 / 3 PART간지 / 4~6 목적·설계·응답자 / 7 PART간지 /
#: 8~14 요약 7장(원본 유지 — 7장이 차트 5/2/15/15/12/6/0개로 배치가 전부 다르다) /
#: 15 PART간지 / 16 챕터간지 원본 / 18 본문 견본 / 98 contact us / 99 마무리
S1_KEEP: list[int] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 18, 98, 99]

#: docProps/app.xml 에서 뺄 LG 전용 글꼴 이름
LG_FONTS = ("LG EI Text TTF Bold", "LG EI Text TTF Regular", "LG Smart UI Regular")


def build_s1_skeleton(src: str | Path, out: str | Path, *,
                      chapters: int = 6,
                      keep: Sequence[int] = tuple(S1_KEEP),
                      chapter_ref: SlideRef = 16,
                      com_verify: bool = True) -> dict[str, Any]:
    """LG 99장 -> 에스원 세팅 골격. 얇은 래퍼일 뿐 콘텐츠는 건드리지 않는다.

    keep_only(keep) -> 챕터 간지를 chapters-1 회 복제(원본 바로 뒤에 연달아) -> save_checked.
    챕터 간지 복제는 파트가 1개씩만 늘어 비용이 사실상 0이다(이미지 공유).

    ★ LG 자산 제거(테마 색·폰트·로고)는 이 함수가 하지 않는다. 조립과 별개 단계다.
      순서는 ① 조립 ② LG 자산 제거 ③ 본문 채우기.
    """
    deck = Deck(src)
    # ★ 위치 주소는 keep_only 로 밀린다. 삭제 전에 파트 이름으로 굳혀 둔다.
    base = resolve(deck, chapter_ref)
    rep = keep_only(deck, list(keep))
    if base not in deck.parts:
        raise ValueError(f"챕터 간지 {chapter_ref} 가 keep 목록에 없어 삭제됐다: {base}")
    clones: list[CloneResult] = []
    if chapters > 1:
        at = positions(deck)[base]      # 1-based 위치 == 0-based '바로 뒤' 삽입 인덱스
        clones = clone_slide(deck, base, at=at, count=chapters - 1)
    path = save_checked(deck, out, com_verify=com_verify, drop_fonts=LG_FONTS)
    return {"path": str(path),
            "slides": len(deck.slide_order()),
            "parts": len(deck.parts),
            "clones": clones,
            "report": rep,
            "problems": validate(deck)}


# --------------------------------------------------------------------------- #
# 자체 검증
# --------------------------------------------------------------------------- #

_DEFAULT_SRC = Path(__file__).resolve().parents[1] / "refs" / "예시_제작본_LG홈로봇_최종.pptx"


def _selftest(src: Path, workdir: Path, *, com: bool = True) -> int:
    import zipfile

    workdir.mkdir(parents=True, exist_ok=True)
    ok = True

    def say(tag: str, msg: str) -> None:
        print(f"  {tag:<5} {msg}")

    def check(cond: bool, msg: str) -> None:
        nonlocal ok
        print(f"  {'OK  ' if cond else 'FAIL'}  {msg}")
        if not cond:
            ok = False

    print(f"원본: {src}")
    deck = Deck(src)
    n0, p0 = len(deck.slide_order()), len(deck.parts)
    say("info", f"원본 {n0}장 / 파트 {p0}개")

    # ---------- 1. keep_only 16장 ----------
    print("\n[1] keep_only -> 16장")
    keep = [1, 2, 3, 4, 5, 6, 7, 8, 15, 16, 18, 19, 20, 37, 98, 99]
    rep = keep_only(deck, keep)
    check(len(deck.slide_order()) == 16, f"덱 내부 슬라이드 수 = {len(deck.slide_order())}")
    say("info", f"삭제 파트 {len(rep.removed_parts)}개 / 공유라 살린 것 {len(rep.kept_shared)}개 "
                f"/ 회수 {rep.freed_bytes:,} B")
    say("info", f"살린 공유 파트: {rep.kept_shared[:6]}")
    f1 = workdir / "t1_keep16.pptx"
    save_checked(deck, f1, com_verify=False, drop_fonts=LG_FONTS)
    with zipfile.ZipFile(f1) as z:
        bad = z.testzip()
        n_slide = len([n for n in z.namelist() if re.match(r"^ppt/slides/slide\d+\.xml$", n)])
    check(bad is None, "zipfile 무결성 (testzip)")
    check(n_slide == 16, f"zip 안 슬라이드 파트 = {n_slide}")
    d1 = Deck(f1)
    check(len(d1.slide_order()) == 16, f"다시 열어 sldIdLst = {len(d1.slide_order())}장")
    check(len(d1.parts) < p0, f"파트 {p0} -> {len(d1.parts)}")

    # ---------- 2. clone_slide 7장 ----------
    print("\n[2] clone_slide -> +7 = 23장")
    dmap = {r["pos"]: r for r in deck_map(d1)}
    say("info", "복제 전 n_charts: " + ", ".join(f"{p}:{dmap[p]['n_charts']}" for p in sorted(dmap)))
    # 간지(차트 0개) 3장 + 본문(차트 1개) 4장 + 요약(차트 다수) 로 섞어 복제한다
    plan = [(10, 3), (11, 3), (8, 1)]        # 위치 10=챕터간지, 11=본문, 8=요약(차트 5)
    made: list[CloneResult] = []
    for pos, cnt in plan:
        part = resolve(d1, pos)
        res = clone_slide(d1, part, count=cnt)
        made += res
        say("info", f"pos{pos} ({dmap[pos]['n_charts']}차트) x{cnt} -> 1장당 새 파트 "
                    f"{len(res[0].part_map)}개, 공유(복제 안 함) {len(res[0].shared)}개")
    check(len(made) == 7, f"복제 결과 {len(made)}장")
    check(len(d1.slide_order()) == 23, f"덱 내부 = {len(d1.slide_order())}장")
    # 요약 장표(차트 5개) 복제가 차트 일습을 전부 떴는가
    summary = [c for c in made if c.part_map and len(c.charts) >= 2]
    if summary:
        c = summary[0]
        check(len(c.charts) == dmap[8]["n_charts"],
              f"차트 {dmap[8]['n_charts']}개 슬라이드 -> 새 차트 {len(c.charts)}개 (다중 차트 함정)")
        xl = [v for k, v in c.part_map.items() if k.endswith(".xlsx")]
        check(len(xl) == dmap[8]["n_charts"], f"임베디드 워크북도 {len(xl)}벌 새로 떴다")
    f2 = workdir / "t2_clone23.pptx"
    save_checked(d1, f2, com_verify=False, drop_fonts=LG_FONTS)
    with zipfile.ZipFile(f2) as z:
        check(z.testzip() is None, "zipfile 무결성")
        n_slide = len([n for n in z.namelist() if re.match(r"^ppt/slides/slide\d+\.xml$", n)])
    check(n_slide == 23, f"zip 안 슬라이드 파트 = {n_slide}")
    d2 = Deck(f2)
    check(len(d2.slide_order()) == 23, f"다시 열어 sldIdLst = {len(d2.slide_order())}장")

    # ---------- 3. reorder 뒤집기 + PowerPoint COM ----------
    print("\n[3] reorder (완전 역순) -> PowerPoint COM 오픈")
    before = d2.slide_order()
    reorder(d2, list(reversed(range(1, len(before) + 1))))
    after = d2.slide_order()
    check(after == list(reversed(before)), "sldIdLst 가 정확히 역순")
    f3 = workdir / "t3_reversed.pptx"
    save_checked(d2, f3, com_verify=False, drop_fonts=LG_FONTS)
    probs = validate(d2)
    check(not [p for p in probs if p.level == "fatal"], f"validate fatal 0건 (전체 {len(probs)}건)")
    for p in probs[:8]:
        say("warn", str(p))
    if com:
        info = assert_openable(f3, export_positions=[1, 2, len(after)])
        check(info["slides"] == 23, f"PowerPoint 가 연 슬라이드 수 = {info['slides']} (복구 대화상자 없음)")
        say("info", f"레이아웃 이름 보존: {sorted(set(info['layouts']))}")
        say("info", f"PNG: {info['exported']}")
        # 원본 1번(표지)이 역순 후 마지막에 왔는가 -- 레이아웃 이름으로 대조
        d3 = Deck(f3)
        lay_first = _layout_name(d3, [r["Part"] for r in _rel_rows(d3, d3.slide_at(1))
                                      if r["Part"] and r["Type"].endswith("/slideLayout")][0])
        check(info["layouts"][0] == lay_first,
              f"COM 이 보는 1번 레이아웃 '{info['layouts'][0]}' == XML 판독 '{lay_first}'")
    else:
        say("skip", "COM 검증 건너뜀 (--no-com)")

    # ---------- 4. 고아 파트 0 ----------
    print("\n[4] 고아 파트 / 정합성")
    for tag, f in (("t1", f1), ("t2", f2), ("t3", f3)):
        d = Deck(f)
        orph = gc_orphans(d, dry_run=True)
        pr = validate(d)
        check(not orph, f"{tag} 고아 파트 {len(orph)}개 {orph[:5]}")
        check(not [x for x in pr if x.level == "fatal"], f"{tag} validate fatal 0건")
        ct_warn = [x for x in pr if x.code.startswith("CT_")]
        check(not ct_warn, f"{tag} Content_Types 정합 (누락/댕글링 {len(ct_warn)}건)")
        app_warn = [x for x in pr if x.code == "APP_XML_STALE"]
        check(not app_warn, f"{tag} app.xml <Slides> 최신")
        say("info", f"{tag}: {len(d.slide_order())}장 / 파트 {len(d.parts)} / "
                    f"{Path(f).stat().st_size:,} B")

    # ---------- 5. 복제본이 진짜 독립인가 ----------
    print("\n[5] 복제 독립성 — 슬라이드 간 비-media 파트 공유")
    d5 = Deck(f2)
    rc = refcount(d5)
    shared_nonmedia = sorted(p for p, c in rc.items()
                             if c > 1 and not p.startswith("ppt/media/") and not p.endswith(".rels"))
    shared_media = sorted(p for p, c in rc.items() if c > 1 and p.startswith("ppt/media/"))
    check(not shared_nonmedia, f"비-media 공유 파트 {len(shared_nonmedia)}개 {shared_nonmedia[:4]}")
    say("info", f"media 는 의도적으로 공유: {len(shared_media)}개 {shared_media[:4]}")
    xlsx = [p for p in d5.parts if p.startswith("ppt/embeddings/")]
    check(len(set(xlsx)) == len(xlsx) and all(rc.get(p, 0) <= 1 for p in xlsx),
          f"임베디드 워크북 {len(xlsx)}개가 전부 1장 전속 (데이터 편집이 옆 장표를 안 건드린다)")

    # ---------- 6. at= 삽입 위치 / 복제본 삭제 / apply_plan ----------
    print("\n[6] at= 삽입 · 복제본 삭제 · apply_plan")
    d6 = Deck(f1)                                   # 16장에서 다시 시작
    src_pos = 11                                    # 본문 견본 (차트 1개)
    src_part = resolve(d6, src_pos)
    before6 = d6.slide_order()
    res6 = clone_slide(d6, src_part, at=3, count=2)  # 0-based 3 -> 4·5번째 자리
    check([r.position for r in res6] == [4, 5], f"at=3, count=2 -> 위치 {[r.position for r in res6]}")
    after6 = d6.slide_order()
    check(after6[:3] == before6[:3] and after6[5:] == before6[3:],
          "삽입 지점 앞뒤의 기존 순서가 그대로다")

    priv6 = sorted(set(res6[0].part_map.values()) - {res6[0].slide_part})
    check(all(p in d6.parts for p in priv6), f"복제본 전속 파트 {len(priv6)}개 존재")
    src_priv = sorted(owned_parts(d6, src_part) - {src_part} - {Deck.rels_name(src_part)})
    rep6 = delete_slides(d6, [res6[0].slide_part])
    check(len(d6.slide_order()) == 17, f"복제본 1장 삭제 -> {len(d6.slide_order())}장")
    check(all(p not in d6.parts for p in priv6),
          f"복제본 전속 파트 {len(priv6)}개가 함께 사라졌다")
    check(all(p in d6.parts for p in src_priv if not p.startswith("ppt/media/")),
          "원본 슬라이드의 차트/워크북은 그대로 남았다")
    check(all(p in d6.parts for p in src_priv if p.startswith("ppt/media/")),
          "공유 media 는 살아남았다")
    say("info", f"삭제 회수 {rep6.freed_bytes:,} B / 공유라 살린 것 {len(rep6.kept_shared)}개")

    plan = [17, 1, 5, 9, 2]                          # 부분집합 + 새 순서
    want = [resolve(d6, p) for p in plan]
    apply_plan(d6, plan)
    check(d6.slide_order() == want, f"apply_plan -> {len(d6.slide_order())}장, 계획 순서 일치")
    f6 = workdir / "t6_plan5.pptx"
    save_checked(d6, f6, com_verify=False, drop_fonts=LG_FONTS)
    d6b = Deck(f6)
    check(len(d6b.slide_order()) == 5, f"저장 후 {len(d6b.slide_order())}장")
    check(not gc_orphans(d6b, dry_run=True), "고아 0")
    check(not [x for x in validate(d6b) if x.level == "fatal"], "validate fatal 0건")
    if com:
        info6 = assert_openable(f6)
        check(info6["slides"] == 5, f"PowerPoint 가 연 슬라이드 수 = {info6['slides']}")
    say("info", f"t6: 5장 / 파트 {len(d6b.parts)} / {f6.stat().st_size:,} B")

    # ---------- 7. 가드 — validate 등급이 PowerPoint 실동작과 맞는가 ----------
    # 여기 있는 등급은 전부 실측이다. 파손본을 만들어 PowerPoint COM 으로 열어 보고
    # (거부 / 연다+유실 / 연다+정상) 셋 중 어디인지를 확인해 정했다. 2026-08-22 · 18종.
    print("\n[7] 파손 재현 — validate 등급 ↔ PowerPoint 실동작")

    def _mut_pres(d: Deck, fn: Callable[[etree._Element], None]) -> None:
        lst = _sld_id_lst(d)
        fn(lst)
        d.put(PRES, d.xml(PRES))

    def _mut_ct(d: Deck, fn: Callable[[etree._Element], None]) -> None:
        t = etree.fromstring(d.parts["[Content_Types].xml"])
        fn(t)
        d.put_bytes("[Content_Types].xml", _xml_bytes(t))

    def _ct_el(t: etree._Element, local: str, attr: str, value: str) -> etree._Element:
        return next(e for e in t if etree.QName(e).localname == local
                    and (e.get(attr) or "").lower() == value.lower())

    def _drop_default(d: Deck, ext: str) -> None:
        _mut_ct(d, lambda t: t.remove(_ct_el(t, "Default", "Extension", ext)))

    def _dup_override(d: Deck) -> None:
        def go(t: etree._Element) -> None:
            src = next(e for e in t if etree.QName(e).localname == "Override")
            el = etree.SubElement(t, q(CT_NS, "Override"))
            el.set("PartName", src.get("PartName"))
            el.set("ContentType", src.get("ContentType"))
        _mut_ct(d, go)

    def _dup_default(d: Deck) -> None:
        def go(t: etree._Element) -> None:
            src = _ct_el(t, "Default", "Extension", "png")
            el = etree.SubElement(t, q(CT_NS, "Default"))
            el.set("Extension", src.get("Extension"))
            el.set("ContentType", src.get("ContentType"))
        _mut_ct(d, go)

    def _drop_rel(d: Deck, part: str, type_suffix: str) -> None:
        """파트의 rels 에서 관계를 통째로 지운다. 댕글링이 안 생겨 참조 검사를 다 통과한다."""
        name = Deck.rels_name(part)
        t = etree.fromstring(d.parts[name])
        for el in [e for e in t if (e.get("Type") or "").endswith(type_suffix)]:
            t.remove(el)
        d.put_bytes(name, _xml_bytes(t))

    def _dangle_body_rid(d: Deck, part: str, localname: str, attr: str) -> None:
        """본문의 r:* 속성만 없는 rId 로 바꾼다. rels 는 멀쩡해서 대상 존재 검사를 통과한다."""
        t = d.xml(part)
        el = next(e for e in t.iter() if isinstance(e.tag, str)
                  and etree.QName(e).localname == localname and e.get(q(R, attr)))
        el.set(q(R, attr), "rId9999")
        d.put(part, t)

    def _ghost_slide(d: Deck) -> None:
        """sldIdLst 항목만 삭제. 파트·rels·Override 는 전부 살려 둔다 — 어떤 참조도 안 끊긴다."""
        _mut_pres(d, lambda l: l.remove(l[5]))
        d.put_bytes(APP_XML, re.sub(rb"<Slides>\d+</Slides>",
                                    f"<Slides>{len(d.slide_order())}</Slides>".encode(),
                                    d.parts[APP_XML]))

    def _first_chart_slide(d: Deck) -> str:
        return next(s for s in d.slide_order() if d.chart_parts(s))

    #  (라벨, 기대 fatal 코드, 주입 함수)
    breakages: list[tuple[str, str, Callable[[Deck], None]]] = [
        # --- 기존 6종: PowerPoint 가 거부한다 ---
        ("N1  sldIdLst r:id 가 rels 에 없음      [거부 0x80070570]", "SLDID_DEAD_RID",
         lambda d: _mut_pres(d, lambda l: l[2].set(q(R, "id"), "rId9999"))),
        ("N4  sldId/@id 중복                     [거부 0x80070570]", "SLDID_DUP_ID",
         lambda d: _mut_pres(d, lambda l: l[3].set("id", l[2].get("id")))),
        ("N6  sldId/@id=100 (최소 256 위반)      [거부]", "SLDID_ID_RANGE",
         lambda d: _mut_pres(d, lambda l: l[1].set("id", "100"))),
        ("N7  두 sldId 가 같은 rId               [거부 0x80070570]", "SLDID_DUP_RID",
         lambda d: _mut_pres(d, lambda l: l[4].set(q(R, "id"), l[3].get(q(R, "id"))))),
        ("N2  슬라이드 파트만 증발               [거부 0x80070570]", "SLIDE_PART_MISSING",
         lambda d: d.parts.pop(d.slide_at(6))),
        ("N3  차트 파트만 증발                   [거부 0x80070570]", "REL_TARGET_MISSING",
         lambda d: d.parts.pop(next(c for s in d.slide_order() for c in d.chart_parts(s)))),
        # --- 신규 5종: 이전 판은 fatal 도 warn 도 없이 통과시켰다 ---
        ("N8  CT Default png 삭제                [거부 0x80CB8002]", "CT_UNTYPED_PART",
         lambda d: _drop_default(d, "png")),
        ("N8b CT Default rels 삭제               [거부 0x80CB8002]", "CT_UNTYPED_PART",
         lambda d: _drop_default(d, "rels")),
        ("N9  CT Override PartName 중복          [거부 0x80CB8001]", "CT_DUP_OVERRIDE",
         _dup_override),
        ("N10 CT Default Extension 중복          [거부 0x80CB8000]", "CT_DUP_DEFAULT",
         _dup_default),
        ("N11 차트 externalData rel 삭제         [거부]", "REL_ID_DANGLING",
         lambda d: _drop_rel(d, d.chart_parts(_first_chart_slide(d))[0], "/package")),
        ("N12 c:chart r:id 댕글링                [거부]", "REL_ID_DANGLING",
         lambda d: _dangle_body_rid(d, _first_chart_slide(d), "chart", "id")),
        ("N13 blip r:embed 댕글링                [연다 — 그림 자리가 오류 상자]", "REL_ID_DANGLING",
         lambda d: _dangle_body_rid(d, next(s for s in d.slide_order()
                                            if b"blip" in d.parts[s]), "blip", "embed")),
        ("N14 슬라이드→레이아웃 관계 삭제        [거부]", "REL_REQUIRED_MISSING",
         lambda d: _drop_rel(d, d.slide_at(2), "/slideLayout")),
        ("N15 레이아웃→마스터 관계 삭제          [거부 0x80070570]", "REL_REQUIRED_MISSING",
         lambda d: _drop_rel(d, d.related(d.slide_at(2), "/slideLayout")[0], "/slideMaster")),
        ("N16 유령 슬라이드 (sldId 항목만 삭제)  [연다 — 한 장이 조용히 사라진다]",
         "SLIDE_NOT_IN_SLDIDLST", _ghost_slide),
    ]
    for i, (label, code, mutate) in enumerate(breakages):
        d7 = Deck(f1)
        mutate(d7)
        codes = {p.code for p in validate(d7) if p.level == "fatal"}
        check(code in codes, f"{label} -> {sorted(codes)}")
        if i == 0:                    # COM 은 대표 1건만 (PowerPoint 기동이 느리다)
            f7 = workdir / "t7_broken.pptx"
            d7.flush()
            with zipfile.ZipFile(f7, "w", zipfile.ZIP_DEFLATED) as z:
                for nm, b in d7.parts.items():
                    z.writestr(nm, b)
            with zipfile.ZipFile(f7) as z:
                check(z.testzip() is None, "  파손본도 zipfile 로는 멀쩡히 열린다(그래서 COM 검증이 필요하다)")
            if com:
                try:
                    assert_openable(f7)
                    check(False, "  PowerPoint 가 파손본을 열어 버렸다")
                except RuntimeError as exc:
                    check("0x80070570" in str(exc) or "손상" in str(exc) or "corrupt" in str(exc).lower(),
                          f"  PowerPoint 가 파손본을 거부: {str(exc).splitlines()[-1].strip()[:60]}")
        try:
            save_checked(d7, workdir / "t7_should_not_exist.pptx", com_verify=False)
            check(False, f"  save_checked 가 {code} 를 막지 못했다")
        except RuntimeError:
            pass
    say("info", f"save_checked 는 fatal 이 있으면 저장 자체를 막는다 ({len(breakages)}/{len(breakages)} 차단)")

    # ---- 과잉 등급 가드 ----
    # PowerPoint 가 **정상으로 여는** 것을 fatal 로 올리면 멀쩡한 납품물의 저장이 막힌다.
    # 아래 2종은 규격 위반이지만 실측상 22/22 정상 오픈이므로 warn 에 머물러야 한다.
    print("\n[7b] 과잉 등급 가드 — PowerPoint 가 여는 것은 fatal 이 아니어야 한다")

    def _drop_override(d: Deck) -> None:
        part = "/" + d.slide_at(2)
        _mut_ct(d, lambda t: t.remove(_ct_el(t, "Override", "PartName", part)))

    tolerated: list[tuple[str, str, Callable[[Deck], None]]] = [
        ("슬라이드 Override 삭제 (xml Default 로 폴백) [연다 정상]",
         "CT_OVERRIDE_MISSING", _drop_override),
        ("고아 파트 1개 추가                            [연다 정상]",
         "PART_ORPHAN", lambda d: d.put_bytes("ppt/slides/slide999.xml", b"<x/>")),
    ]
    for label, code, mutate in tolerated:
        d7b = Deck(f1)
        mutate(d7b)
        probs = validate(d7b)
        fatal = sorted({p.code for p in probs if p.level == "fatal"})
        warns = sorted({p.code for p in probs if p.level == "warn"})
        check(not fatal and code in warns, f"{label} -> warn {warns} / fatal {fatal}")

    # 무손상 원본은 아무것도 안 나와야 한다 (오탐 0)
    check(not validate(Deck(f1)), f"무손상 16장 덱 -> 문제 0건")
    for tag, f in (("t2 복제 23장", f2), ("t3 역순 23장", f3), ("t6 재배열 5장", f6)):
        check(not validate(Deck(f)), f"무손상 {tag} -> 문제 0건")

    d8 = Deck(f1)
    for label, fn in (("keep_only([])", lambda: keep_only(d8, [])),
                      ("delete_slides(전부)", lambda: delete_slides(d8, list(range(1, 17)))),
                      ("reorder(부분집합)", lambda: reorder(d8, [1, 2, 3]))):
        try:
            fn()
            check(False, f"{label} 가 통과해 버렸다")
        except ValueError:
            check(True, f"{label} -> ValueError 로 막힘")

    print("\n" + ("=== 전부 통과 ===" if ok else "=== 실패 있음 ==="))
    return 0 if ok else 1


def main(argv: Sequence[str]) -> int:
    args = list(argv)
    com = "--no-com" not in args
    args = [a for a in args if a != "--no-com"]
    if args and args[0] == "--selftest":
        args = args[1:]
    src = Path(args[0]) if args else _DEFAULT_SRC
    workdir = Path(args[1]) if len(args) > 1 else Path(tempfile.gettempdir()) / "deckbuild_selftest"
    return _selftest(src, workdir, com=com)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
