# -*- coding: utf-8 -*-
"""deckops.py -- 제작본 .pptx 를 파트(zip 엔트리) 단위로 열고 고치는 저수준 도구.

python-pptx 로 열었다 다시 저장하면 서식이 미묘하게 어긋난다(차트 스타일 파트 소실,
테마 상속 rPr 평탄화 등). 그래서 여기서는 **zip 을 파트 dict 로 통째로 들고 있다가
건드린 파트만 XML 로 파싱해 고쳐 되넣고 나머지는 바이트 그대로 다시 쓴다.**
손대지 않은 파트는 바이트가 1비트도 안 바뀐다 -- 그게 서식 100% 보존의 근거다.

담당 범위 (전부 "XML 을 어떻게 만지는가" 뿐이다):
  · Deck            zip <-> 파트 dict, rels 해석, 파트 이름 계산
  · 표 행 개수       set_row_count  (R-105 / R-106 / R-107)
  · 셀 텍스트        clear_text / set_text  (R-096, 다중 런·다중 문단·수동 줄바꿈)
  · 도형 텍스트      set_shape_text / set_ph_text / drop_shape
  · 멘트 주입        set_ment  (R-090 -- idx 는 **호출자가 위치로 판정해** 넘긴다)
  · 차트 계열        sync_chart + 임베디드 워크시트 갱신
  · 기하 실측        measure_table_geometry / label_metrics  (R-108 줄바꿈 폭)

★ 이 모듈은 **숫자를 만들지 않는다.** 무슨 값을 어디에 넣을지는 rules.py 가 정하고
  fill_deck.py 가 결정한다. 여기는 그 결정을 XML 에 옮기는 손발일 뿐이다.

★ 이 모듈은 **프로젝트 상수를 갖지 않는다.** 배너 열 문자·base n·배너명·그룹 같은
  덱마다 다른 것은 전부 인자로 받는다. (프로토타입의 BANNER_COLS/BASE_N/NAMES/GROUPS 는
  LG 전용이라 승격 과정에서 전량 제거했다. 표 구조는 호출자가 BannerAxis 로 넘긴다.)

실측 근거
---------
R-105 (표 기하)   예시_제작본_LG홈로봇 26장 계측
      표 y=2,220,713 / h=3,960,000 고정, 헤더 3행 = 216000+468000+180000 = 864,000 고정
      데이터 행 높이 = (3,960,000 - 864,000) / n = 3,096,000 / n
      n=4 -> 774000 · n=5 -> 619200 · n=10 -> 309600  (전부 실측 일치)
      나머지 EMU 는 첫 행에 몰아 합을 정확히 맞춘다.
      ※ 다만 이 값들은 **기본값일 뿐 상수가 아니다.** 같은 LG 덱에서도 가로누적100%
        장표(s25)는 y=3,804,713 · 총높이 2,376,000 · 데이터 영역 1,512,000 이다.
        그래서 set_row_count 는 기본적으로 **현재 표에서 데이터 영역 높이를 실측해**
        그 합을 보존한다. DATA_AREA_H 상수는 실측이 불가능할 때의 폴백이다.
R-106 (차트 기하) 차트 graphicFrame 은 보기 개수와 무관하게 고정 -- 절대 건드리지 않는다.
      plotArea 가 n개 범주를 같은 영역에 등분하므로 표 행과 자동으로 정렬된다.
      이 모듈에는 차트 프레임을 **쓰는** 코드가 아예 없다(읽기만 한다). 그게 R-106 준수다.
R-107 (행 서식)   [첫행][중간행...][마지막행] 3종. 마지막 행만 lnB=solid (표 하단 실선).
      축소는 중간부터 지우고, 확대는 중간을 복제해 마지막 행 **앞에** 끼운다.
      데이터 행이 2개뿐인 표에는 중간 템플릿이 없다 -> mid_template 인자로 같은
      클론 패밀리의 다른 장표에서 가져온다. 안 주면 첫 행을 복제하고 on_warn 으로 알린다.
R-096 (셀 서식)   `cell.text = ...` 는 서식을 날린다. 반드시 런 단위로 교체한다.
      원본 셀은 런이 3개이거나 문단이 2개인 경우가 있다(LG s23 라벨열 10칸 중 3칸).
R-108 (줄바꿈)    자동 줄바꿈에 맡기지 않고 <a:br/> 를 박는다.
      가용 폭 = tblGrid[col].w - tcPr.marL - tcPr.marR (label_metrics 가 실측).
R-090 (멘트 슬롯) placeholder idx 로 sub_title/ment 를 가르면 안 된다.
      LG=(12=sub_title, 13=ment) / SKEE=(13=sub_title, 12=ment) 로 역할이 반대다.
      set_ment 는 idx 를 **인자로 받는다** -- 위치 판정은 호출자의 몫이다.

이번 승격에서 새로 드러난 것 (프로토타입이 LG s18/s23 만 밟아 못 본 지점. 규칙 후보)
---------------------------------------------------------------------------------
(가) **임베디드 워크시트 배치가 차트타입마다 다르다.** 계열이 열로 서기도 하고 행으로
     눕기도 한다. LG 실측:
         chart64 겹침막대     tx=$B$1  cat=$A$2:$A$11  val=$B$2:$B$11   계열 = 열
         chart78 ±세로막대    tx=$A$2  cat=$B$1:$L$1   val=$B$2:$L$2    계열 = 행
         chart66 가로누적100% tx=$A$2  cat=$B$1        val=$B$2         행 + 카테고리 1개
     프로토타입은 전부 `$COL$2:$COL$(n+1)` 로 고쳐 썼다 -- 척도 장표(LG 36장·SKEE 44장)의
     워크시트 연결이 통째로 끊긴다. series_orientation() 이 범위 모양으로 방향을 판정한다.
(나) **척도 차트의 워크시트는 수식으로 굴러간다.** chart78 실측:
         r2~r8   차트가 보는 범위인데 전부 수식 (`=B15`, 중립점 아래는 `=B14*-1`)
         r10     `▼데이터 입력칸`
         r11~r17 사람이 넣는 생 % 값
     발산형(±)의 부호 반전이 이 수식이다. 생값을 덮어쓰면 장치가 사라진다.
     update_embedded_workbook(preserve_formulas=True, 기본) 이 수식 셀을 건드리지 않고,
     입력 블록은 호출자가 write_workbook_cells() 로 채운다.
     ※ 그럼 값의 진짜 주입 지점이 어디인가(차트 범위냐 입력 블록이냐)는 fill_deck 의
       결정 사항이다. deckops 는 둘 다 할 수 있게 열어 두고 판단하지 않는다.
(다) **R-107 의 실제 신호는 lnB 의 존재가 아니라 prstDash 값이다.** lnB 엘리먼트는 모든
     행에 있다. LG s17/s18/s23 실측: 첫행(solid,dot) 중간(dot,dot) 마지막(dot,solid).
     row_border_dashes() 로 읽는다.

의존: lxml (필수) / openpyxl (임베디드 워크시트를 만질 때만 지연 임포트)
"""

from __future__ import annotations

import copy
import io
import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator, Sequence

from lxml import etree

# --------------------------------------------------------------------------- #
# 네임스페이스
# --------------------------------------------------------------------------- #

A = "http://schemas.openxmlformats.org/drawingml/2006/main"            # DrawingML
P = "http://schemas.openxmlformats.org/presentationml/2006/main"       # PresentationML
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"           # Chart
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"    # .rels 파일 안

NS = {"a": A, "p": P, "c": C, "r": R, "pr": PR}

EMU_PER_CM = 360000
EMU_PER_PT = 12700

# --------------------------------------------------------------------------- #
# R-105 실측 기본값
#
# 상수로 박아 쓰라고 두는 게 아니라 **실측이 불가능할 때의 폴백**이다.
# 정상 경로는 measure_table_geometry() / set_row_count(data_area_h=None) 의 실측이다.
# --------------------------------------------------------------------------- #

HEADER_ROWS = 3          # 배너표 헤더 행 수 (그룹행 / 라벨행 / (Base)행)
HEADER_H = 864000        # 216000 + 468000 + 180000
DATA_AREA_H = 3096000    # 3960000 - 864000
TABLE_Y = 2220713
TABLE_H = 3960000
TABLE_BOTTOM = TABLE_Y + TABLE_H     # 6180713
CHART_Y = 3037299        # R-106: 읽기 전용. 이 모듈은 차트 프레임을 쓰지 않는다
CHART_H = 3200013

DEFAULT_CELL_MARGIN = 91440          # ECMA-376 tcPr marL/marR 기본값 (0.1인치)

# 임베디드 워크시트에서 값이 시작하는 지점의 기본값. 두 덱 전수(2,685개 c:f)에서
# 1행/1열이 머리글이고 데이터는 2부터였다. 실제 값은 chart_layout() 이 c:f 에서 읽으므로
# 이 상수는 참조가 없을 때의 폴백일 뿐이다.
WORKBOOK_FIRST_DATA_ROW = 2
WORKBOOK_FIRST_DATA_COL = "B"


# --------------------------------------------------------------------------- #
# 태그 헬퍼
# --------------------------------------------------------------------------- #

def q(ns: str, tag: str) -> str:
    """Clark 표기 태그 이름. q(A, 'tr') -> '{...drawingml...}tr'."""
    return f"{{{ns}}}{tag}"


def _int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _tostring(tree: etree._Element) -> bytes:
    """OOXML 파트 직렬화. 선언부를 원본과 같은 꼴로 유지한다."""
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


# --------------------------------------------------------------------------- #
# Deck -- pptx 를 파트 dict 로 들고 있는 컨테이너
# --------------------------------------------------------------------------- #

class Deck:
    """pptx 를 열어 {파트이름: bytes} 로 들고 있다가 통째로 다시 쓴다.

    · 손대지 않은 파트는 바이트 그대로 나간다 (서식·차트스타일 100% 보존)
    · 원본 파일은 절대 건드리지 않는다. save() 는 언제나 새 경로에 쓴다
    · dict 는 삽입 순서를 지키므로 파트 순서도 원본 그대로 유지된다

    ★ 고친 트리는 반드시 put() 으로 되돌려야 저장된다.
      xml() 은 **읽기**로 취급한다 -- 읽기만 한 파트를 다시 직렬화하면 lxml 이 속성
      순서·자기닫는태그·공백을 원본과 다르게 뱉어 바이트가 바뀐다. 그러면 '손댄 파트'
      목록이 오염돼 baseline diff(R-097)에 노이즈가 섞인다. 그래서 put() 한 파트만
      더럽다고 표시하고 나머지는 원본 바이트 그대로 내보낸다.
    """

    def __init__(self, path: str | Path):
        self.source = Path(path)
        with zipfile.ZipFile(self.source) as z:
            self.parts: dict[str, bytes] = {n: z.read(n) for n in z.namelist()}
        self._tree_cache: dict[str, etree._Element] = {}
        self._dirty: set[str] = set()

    # ---------- 파트 입출력 ----------

    def __contains__(self, name: str) -> bool:
        return name in self.parts

    def names(self, pattern: str | None = None) -> list[str]:
        """파트 이름 목록. pattern 은 정규식(부분일치)."""
        if pattern is None:
            return list(self.parts)
        rx = re.compile(pattern)
        return [n for n in self.parts if rx.search(n)]

    def xml(self, name: str) -> etree._Element:
        """파트를 XML 트리로 파싱한다. 같은 파트를 두 번 열어도 같은 트리를 준다.

        캐시가 없으면 두 번째 xml() 호출이 첫 번째 수정분을 못 본 채로 새 트리를 만들어
        조용히 덮어쓴다 -- 프로토타입에서 실제로 물렸던 함정이라 캐시를 둔다.
        """
        tree = self._tree_cache.get(name)
        if tree is None:
            tree = etree.fromstring(self.parts[name])
            self._tree_cache[name] = tree
        return tree

    def put(self, name: str, tree: etree._Element) -> None:
        """고친 XML 트리를 파트로 되돌린다. 이걸 불러야 저장된다."""
        self.parts[name] = _tostring(tree)
        self._tree_cache[name] = tree
        self._dirty.add(name)

    def put_bytes(self, name: str, data: bytes) -> None:
        self.parts[name] = data
        self._tree_cache.pop(name, None)
        self._dirty.add(name)

    def touched(self) -> list[str]:
        """put() 으로 실제로 고친 파트 이름. 검수리포트·baseline diff 에 쓴다."""
        return sorted(self._dirty)

    def flush(self) -> None:
        """더럽다고 표시된 트리만 직렬화한다. save() 가 자동으로 부른다."""
        for name in self._dirty:
            tree = self._tree_cache.get(name)
            if tree is not None:
                self.parts[name] = _tostring(tree)

    def save(self, path: str | Path) -> Path:
        """새 경로에 쓴다. 원본 경로와 같으면 거부한다 (원본 덮어쓰기 금지)."""
        out = Path(path)
        if out.resolve() == self.source.resolve():
            raise ValueError(f"원본을 덮어쓸 수 없다: {out}")
        self.flush()
        out.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
            for name, data in self.parts.items():
                zo.writestr(name, data)
        # 문서 위생 (R-137): 문서 속성의 개인 이름·이전 작업자 잔재를 조직명으로, 보이지 않는 문자 제거.
        # 덱 파이프라인의 모든 저장이 이 메서드를 지나므로 여기 한 곳에서 건다. 실패해도 저장은 살린다.
        try:
            import doc_hygiene
            tmp = out.with_suffix(out.suffix + ".hyg")
            doc_hygiene.clean_file(out, tmp, "엠브레인")
            tmp.replace(out)
        except Exception as e:                       # noqa: BLE001
            print(f"[deckops] 문서 위생 건너뜀: {e}")
        return out

    # ---------- 관계(.rels) ----------

    @staticmethod
    def rels_name(part: str) -> str:
        """파트의 .rels 파트 이름. 'ppt/slides/slide1.xml' -> 'ppt/slides/_rels/slide1.xml.rels'."""
        d, b = posixpath.split(part)
        return posixpath.join(d, "_rels", b + ".rels") if d else f"_rels/{b}.rels"

    def rels(self, part: str) -> list[dict[str, str]]:
        """파트의 관계 목록. [{'Id','Type','Target','TargetMode','Part'}]

        'Part' 는 Target 을 패키지 절대 경로로 정규화한 값이다(외부 링크면 None).
        """
        name = self.rels_name(part)
        if name not in self.parts:
            return []
        base = posixpath.dirname(part)
        out: list[dict[str, str]] = []
        for rel in etree.fromstring(self.parts[name]):
            target = rel.get("Target") or ""
            external = rel.get("TargetMode") == "External"
            resolved = None if external else posixpath.normpath(posixpath.join(base, target))
            out.append({
                "Id": rel.get("Id") or "",
                "Type": rel.get("Type") or "",
                "Target": target,
                "TargetMode": rel.get("TargetMode") or "Internal",
                "Part": resolved,
            })
        return out

    def related(self, part: str, type_suffix: str) -> list[str]:
        """관계 Type 이 주어진 접미사로 끝나는 대상 파트 이름들."""
        return [r["Part"] for r in self.rels(part)
                if r["Part"] and r["Type"].endswith(type_suffix)]

    # ---------- 슬라이드 / 차트 ----------

    def slide_part(self, slide_no: int) -> str:
        """1-based 슬라이드 번호 -> 파트 이름.

        주의: 파일명 번호(slideN.xml)와 **표시 순서는 다를 수 있다.** 표시 순서가
        필요하면 slide_order() 를 쓴다. extract_deck 의 slide_no 는 표시 순서다.
        """
        name = f"ppt/slides/slide{slide_no}.xml"
        if name not in self.parts:
            raise KeyError(f"슬라이드 파트 없음: {name}")
        return name

    def slide_order(self) -> list[str]:
        """presentation.xml 의 sldIdLst 순서(= 화면 표시 순서)대로 슬라이드 파트 이름."""
        pres = "ppt/presentation.xml"
        by_rid = {r["Id"]: r["Part"] for r in self.rels(pres)}
        tree = self.xml(pres)
        lst = tree.find(q(P, "sldIdLst"))
        if lst is None:
            return sorted(self.names(r"^ppt/slides/slide\d+\.xml$"))
        out = []
        for sld in lst:
            part = by_rid.get(sld.get(q(R, "id")))
            if part:
                out.append(part)
        return out

    def slide_at(self, position: int) -> str:
        """표시 순서 기준 1-based 위치 -> 슬라이드 파트 이름 (extract_deck 의 slide_no)."""
        order = self.slide_order()
        if not 1 <= position <= len(order):
            raise IndexError(f"슬라이드 위치 {position} (총 {len(order)}장)")
        return order[position - 1]

    def chart_parts(self, slide_part: str) -> list[str]:
        """슬라이드가 참조하는 차트 파트 이름들 (문서 등장 순서)."""
        return self.related(slide_part, "/chart")

    def workbook_part(self, chart_part: str) -> str | None:
        """차트의 임베디드 워크시트 파트 (관계 Type 이 /package)."""
        found = self.related(chart_part, "/package")
        return found[0] if found else None

    # ---------- 테마 ----------

    def theme_part(self, slide_part: str | None = None) -> str | None:
        """슬라이드 -> 레이아웃 -> 마스터 -> 테마 체인을 따라간다.

        slide_part 를 안 주면 theme1.xml 로 폴백한다. 덱에 테마가 여러 개인
        (마스터가 여러 개인) 경우가 있으므로 가능하면 슬라이드를 넘긴다.
        """
        if slide_part:
            for layout in self.related(slide_part, "/slideLayout"):
                for master in self.related(layout, "/slideMaster"):
                    themes = self.related(master, "/theme")
                    if themes:
                        return themes[0]
        found = sorted(self.names(r"^ppt/theme/theme\d+\.xml$"))
        return found[0] if found else None

    def theme_font(self, slide_part: str | None = None, *, script: str = "ea") -> str:
        """테마의 minor 글꼴 이름. 라벨 폭 실측(textfit)에 쓴다.

        script 'ea' = 동아시아(한글), 'latin' = 라틴. ea 가 비어 있으면 latin 으로 떨어진다.
        """
        part = self.theme_part(slide_part)
        if part is None:
            return ""
        tree = self.xml(part)
        minor = tree.find(f".//{q(A, 'fontScheme')}/{q(A, 'minorFont')}")
        if minor is None:
            return ""
        for want in (script, "latin"):
            el = minor.find(q(A, want))
            if el is not None and (el.get("typeface") or "").strip():
                return el.get("typeface").strip()
        return ""

    # ---------- 도형 텍스트 (얇은 위임) ----------

    def set_shape_text(self, tree, shape_name: str, text: str, split: str | None = None) -> bool:
        return set_shape_text(tree, shape_name, text, split)

    def set_ph_text(self, tree, idx: str | int, text: str, split: str | None = None) -> bool:
        return set_ph_text(tree, idx, text, split)

    def drop_shape(self, tree, shape_name: str) -> bool:
        return drop_shape(tree, shape_name)

    def set_ment(self, tree, idx: str | int, lines: Sequence[str]) -> bool:
        return set_ment(tree, idx, lines)


# --------------------------------------------------------------------------- #
# 도형 찾기
# --------------------------------------------------------------------------- #

def iter_shapes(tree: etree._Element) -> Iterator[etree._Element]:
    """슬라이드의 모든 p:sp (그룹 안까지)."""
    return tree.iter(q(P, "sp"))


def shape_name(sp: etree._Element) -> str:
    nv = sp.find(f".//{q(P, 'cNvPr')}")
    return (nv.get("name") or "") if nv is not None else ""


def find_shape(tree: etree._Element, name: str) -> etree._Element | None:
    """cNvPr@name 이 정확히 일치하는 첫 도형.

    도형 이름은 제작팀이 자동 생성명('TextBox 15')을 쓰는 경우가 있어 신뢰도가 낮다.
    이름으로 못 찾으면 호출자가 텍스트 패턴으로 폴백해야 한다 (R-092).
    """
    for sp in iter_shapes(tree):
        if shape_name(sp) == name:
            return sp
    return None


def find_placeholder(tree: etree._Element, idx: str | int) -> etree._Element | None:
    """p:ph@idx 로 플레이스홀더 도형을 찾는다.

    ★ idx 의 의미는 덱마다 다르다 (R-090/DP-001). 여기서는 찾아 주기만 한다.
    """
    want = str(idx)
    for sp in iter_shapes(tree):
        ph = sp.find(f".//{q(P, 'ph')}")
        if ph is not None and (ph.get("idx") or "") == want:
            return sp
    return None


def placeholder_slots(tree: etree._Element) -> list[dict[str, Any]]:
    """슬라이드의 플레이스홀더를 top 오름차순으로. R-090 위치 판정용 재료다.

    반환 [{'idx','type','name','top_emu','top_cm','text'}]
    top_emu 가 None 이면 좌표를 레이아웃에서 상속받은 것이다 -- 그 경우 이 함수만으로는
    순서를 못 정하므로 extract_deck 의 placeholders_by_idx(레이아웃까지 해석함)를 쓴다.
    """
    out: list[dict[str, Any]] = []
    for sp in iter_shapes(tree):
        ph = sp.find(f".//{q(P, 'ph')}")
        if ph is None:
            continue
        off = sp.find(f"{q(P, 'spPr')}/{q(A, 'xfrm')}/{q(A, 'off')}")
        top = _int(off.get("y"), -1) if off is not None else -1
        out.append({
            "idx": ph.get("idx"),
            "type": ph.get("type"),
            "name": shape_name(sp),
            "top_emu": None if top < 0 else top,
            "top_cm": None if top < 0 else round(top / EMU_PER_CM, 3),
            "text": shape_text(sp),
        })
    out.sort(key=lambda d: (d["top_emu"] is None, d["top_emu"] or 0))
    return out


def shape_text(sp: etree._Element) -> str:
    """도형의 모든 a:t 를 이어 붙인다 (문단 사이는 \\n)."""
    body = sp.find(q(P, "txBody"))
    if body is None:
        return "".join(t.text or "" for t in sp.iter(q(A, "t")))
    paras = []
    for para in body.findall(q(A, "p")):
        paras.append("".join(t.text or "" for t in para.iter(q(A, "t"))))
    return "\n".join(paras)


# --------------------------------------------------------------------------- #
# 도형 텍스트 교체
# --------------------------------------------------------------------------- #

def _write_runs(ts: list[etree._Element], text: str, split: str | None = None) -> None:
    """런이 여러 개로 쪼개진 도형에 텍스트를 넣는다.

    split(예 '_')이 **원문과 새 문장 양쪽에** 있고 런이 2개 이상이면, split 이 걸친
    런을 경계로 잡아 접두/접미를 원래 런에 나눠 넣는다. 그래야 접미부의 다른 서식
    (작고 연한 글씨)이 살아남는다. 조건이 안 맞으면 첫 런에 몰고 나머지를 비운다.
    """
    if not ts:
        return
    old = "".join(t.text or "" for t in ts)
    boundary = None
    if split and split in old and split in text and len(ts) > 1:
        pos, acc = old.index(split), 0
        for i, t in enumerate(ts):
            acc += len(t.text or "")
            if acc > pos:
                boundary = i
                break
        if boundary == 0:          # 첫 런 안에서 갈리면 나눌 게 없다
            boundary = None

    if boundary is None:
        ts[0].text = text
        for extra in ts[1:]:
            extra.text = ""
    else:
        head, tail = text.split(split, 1)
        ts[0].text = head
        for i in range(1, len(ts)):
            ts[i].text = (split + tail) if i == boundary else ""


def set_shape_text(tree: etree._Element, name: str, text: str,
                   split: str | None = None) -> bool:
    """이름으로 도형을 찾아 텍스트를 갈아 끼운다. 못 찾으면 False."""
    sp = find_shape(tree, name)
    if sp is None:
        return False
    ts = sp.findall(f".//{q(A, 't')}")
    if not ts:
        return False
    _write_runs(ts, text, split)
    return True


def set_ph_text(tree: etree._Element, idx: str | int, text: str,
                split: str | None = None) -> bool:
    """플레이스홀더 idx 로 찾아 텍스트를 갈아 끼운다. ★ idx 는 호출자가 판정한다 (R-090)."""
    sp = find_placeholder(tree, idx)
    if sp is None:
        return False
    ts = sp.findall(f".//{q(A, 't')}")
    if not ts:
        return False
    _write_runs(ts, text, split)
    return True


def drop_shape(tree: etree._Element, name: str) -> bool:
    """도형을 통째로 제거한다. `Data 입력` 메모 삭제용 (R-082/R-098).

    ★ 삭제 전에 observations.jsonl 에 적재하는 것은 호출자의 책임이다 -- 학습 데이터다.
    """
    sp = find_shape(tree, name)
    if sp is None:
        return False
    sp.getparent().remove(sp)
    return True


def drop_shapes_matching(tree: etree._Element, predicate: Callable[[etree._Element], bool]) -> int:
    """조건에 맞는 도형을 전부 제거하고 개수를 돌려준다.

    도형 이름이 자동 생성명이라 못 믿을 때 텍스트로 거르는 용도.
        drop_shapes_matching(tree, lambda sp: 'Data 입력' in shape_text(sp))
    """
    victims = [sp for sp in iter_shapes(tree) if predicate(sp)]
    for sp in victims:
        sp.getparent().remove(sp)
    return len(victims)


# --------------------------------------------------------------------------- #
# 멘트 주입 (R-090)
# --------------------------------------------------------------------------- #

def set_ment(tree: etree._Element, idx: str | int, lines: Sequence[str]) -> bool:
    """멘트 플레이스홀더에 줄 목록을 넣는다. 한 줄 = 한 문단.

    ★ idx 를 인자로 받는 이유(R-090/DP-001): placeholder idx 의 역할이 덱마다 반대다.
      LG 는 12=sub_title / 13=ment, SKEE 는 13=sub_title / 12=ment.
      **위에서 몇 번째냐(top_cm)로 판정해** 그 idx 를 넘겨야 한다. 두 덱 모두 멘트가
      아래쪽(top_cm 큰 쪽)이다. placeholder_slots() 또는 extract_deck 결과를 쓴다.

    서식은 기존 문단에서 상속한다 -- 첫 문단의 pPr 과 첫 런의 rPr 을 원형으로 삼고,
    빈 슬롯이면 endParaRPr 을 rPr 로 승격한다. 원형이 아예 없을 때만 최소 rPr 을 만든다.
    """
    sp = find_placeholder(tree, idx)
    if sp is None:
        return False
    body = sp.find(q(P, "txBody"))
    if body is None:
        return False

    ppr_proto, rpr_proto = _para_prototypes(body)

    for old in body.findall(q(A, "p")):
        body.remove(old)

    for line in lines:
        para = etree.SubElement(body, q(A, "p"))
        if ppr_proto is not None:
            para.append(copy.deepcopy(ppr_proto))
        run = etree.SubElement(para, q(A, "r"))
        run.append(copy.deepcopy(rpr_proto) if rpr_proto is not None else _minimal_rpr())
        etree.SubElement(run, q(A, "t")).text = line
    return True


def _para_prototypes(body: etree._Element) -> tuple[etree._Element | None, etree._Element | None]:
    """txBody 에서 (pPr 원형, rPr 원형) 을 뽑는다. 없으면 (None, None)."""
    ppr = rpr = None
    for para in body.findall(q(A, "p")):
        if ppr is None:
            found = para.find(q(A, "pPr"))
            if found is not None:
                ppr = copy.deepcopy(found)
        if rpr is None:
            run = para.find(q(A, "r"))
            if run is not None and run.find(q(A, "rPr")) is not None:
                rpr = copy.deepcopy(run.find(q(A, "rPr")))
            else:
                end = para.find(q(A, "endParaRPr"))
                if end is not None:
                    rpr = copy.deepcopy(end)
                    rpr.tag = q(A, "rPr")
        if ppr is not None and rpr is not None:
            break
    return ppr, rpr


def _minimal_rpr() -> etree._Element:
    """상속할 서식이 하나도 없을 때의 최소 rPr. 나머지는 레이아웃에서 상속된다."""
    rpr = etree.Element(q(A, "rPr"))
    rpr.set("lang", "ko-KR")
    rpr.set("altLang", "en-US")
    rpr.set("dirty", "0")
    return rpr


# --------------------------------------------------------------------------- #
# 표 -- 구조 읽기
# --------------------------------------------------------------------------- #

def find_tables(tree: etree._Element) -> list[etree._Element]:
    """슬라이드의 a:tbl 전부 (문서 등장 순서)."""
    return list(tree.iter(q(A, "tbl")))


def first_table(tree: etree._Element) -> etree._Element:
    """본문 장표는 실측상 표가 1개다. 없으면 LookupError."""
    for tbl in tree.iter(q(A, "tbl")):
        return tbl
    raise LookupError("슬라이드에 표가 없다")


def table_rows(tbl: etree._Element) -> list[etree._Element]:
    return tbl.findall(q(A, "tr"))


def data_rows(tbl: etree._Element, header_rows: int = HEADER_ROWS) -> list[etree._Element]:
    """헤더 뒤의 데이터 행.

    ★ header_rows 를 상수로 가정하면 안 된다. 척도(±세로막대) 장표는 **헤더가 표 하단**에
      있어 위에서부터 3행을 잘라내면 데이터 행을 자른다 (LG s38 실측 6행 = TOP3 / 7점평균 /
      범례 / 배너라벨 / (n) / 배너그룹). 어느 행이 헤더인지는 호출자가 `(NN)` 패턴으로
      판정해 넘겨야 한다 -- fill_deck.banner_axis() / verify_deck.iter_table_rows() 담당.
    """
    return table_rows(tbl)[header_rows:]


def row_cells(tr: etree._Element) -> list[etree._Element]:
    return tr.findall(q(A, "tc"))


def grid_widths(tbl: etree._Element) -> list[int]:
    """tblGrid 의 열 너비(EMU). 열 수는 상수가 아니다 -- 실측 8/9/11/12/15/16/17/24 (R-095)."""
    grid = tbl.find(q(A, "tblGrid"))
    if grid is None:
        return []
    return [_int(gc.get("w")) for gc in grid.findall(q(A, "gridCol"))]


def n_cols(tbl: etree._Element) -> int:
    return len(grid_widths(tbl))


def table_frame(tbl: etree._Element) -> etree._Element | None:
    """표를 담은 p:graphicFrame. tbl -> graphicData -> graphic -> graphicFrame."""
    node = tbl
    for _ in range(4):
        node = node.getparent()
        if node is None:
            return None
        if node.tag == q(P, "graphicFrame"):
            return node
    return None


def frame_box(frame: etree._Element | None) -> dict[str, int | None]:
    """graphicFrame 의 (x, y, cx, cy). p:xfrm 을 정확히 짚는다.

    `.//a:ext` 로 뭉뚱그려 찾으면 a:extLst 안의 a:ext 를 먼저 물어 None 이 나온다 --
    실제로 밟았던 함정이라 경로를 고정한다.
    """
    empty = {"x": None, "y": None, "cx": None, "cy": None}
    if frame is None:
        return empty
    xfrm = frame.find(q(P, "xfrm"))
    if xfrm is None:
        return empty
    off = xfrm.find(q(A, "off"))
    ext = xfrm.find(q(A, "ext"))
    return {
        "x": _int(off.get("x")) if off is not None else None,
        "y": _int(off.get("y")) if off is not None else None,
        "cx": _int(ext.get("cx")) if ext is not None else None,
        "cy": _int(ext.get("cy")) if ext is not None else None,
    }


# --------------------------------------------------------------------------- #
# 표 -- 기하 실측 (R-105)
# --------------------------------------------------------------------------- #

@dataclass
class TableGeometry:
    """표 1장의 실측 기하. 상수를 박는 대신 여기서 읽는다."""
    n_rows: int
    header_rows: int
    n_data_rows: int
    row_heights: list[int]
    header_h: int
    data_area_h: int
    total_h: int
    frame_x: int | None = None
    frame_y: int | None = None
    frame_cx: int | None = None
    frame_cy: int | None = None
    n_cols: int = 0
    col_widths: list[int] = field(default_factory=list)

    @property
    def bottom(self) -> int | None:
        if self.frame_y is None:
            return None
        return self.frame_y + (self.frame_cy or self.total_h)

    @property
    def matches_r105_default(self) -> bool:
        """LG 표준 배너표(y=2220713 · h=3960000 · 헤더 864000)와 같은가."""
        return (self.frame_y == TABLE_Y and self.total_h == TABLE_H
                and self.header_h == HEADER_H and self.data_area_h == DATA_AREA_H)

    def as_dict(self) -> dict[str, Any]:
        d = {k: getattr(self, k) for k in (
            "n_rows", "header_rows", "n_data_rows", "row_heights", "header_h",
            "data_area_h", "total_h", "frame_x", "frame_y", "frame_cx", "frame_cy",
            "n_cols", "col_widths")}
        d["bottom"] = self.bottom
        d["matches_r105_default"] = self.matches_r105_default
        return d


def measure_table_geometry(tbl: etree._Element,
                           header_rows: int = HEADER_ROWS) -> TableGeometry:
    """표에서 R-105 값을 **실측**한다. 상수를 신뢰하지 않는 정상 경로다.

    같은 LG 덱 안에서도 표준 배너표(데이터 영역 3,096,000)와 가로누적100% 장표
    (s25: 총높이 2,376,000 · 데이터 영역 1,512,000)의 기하가 다르다. SKEE 는 또 다르다.
    그러므로 set_row_count 는 이 실측값의 **합을 보존**하는 방식으로 동작한다.
    """
    rows = table_rows(tbl)
    heights = [_int(tr.get("h")) for tr in rows]
    hr = max(0, min(header_rows, len(rows)))
    widths = grid_widths(tbl)
    box = frame_box(table_frame(tbl))
    return TableGeometry(
        n_rows=len(rows),
        header_rows=hr,
        n_data_rows=len(rows) - hr,
        row_heights=heights,
        header_h=sum(heights[:hr]),
        data_area_h=sum(heights[hr:]),
        total_h=sum(heights),
        frame_x=box["x"], frame_y=box["y"], frame_cx=box["cx"], frame_cy=box["cy"],
        n_cols=len(widths), col_widths=widths,
    )


def measure_chart_geometry(tree: etree._Element) -> list[dict[str, int | None]]:
    """슬라이드의 차트 graphicFrame 기하를 **읽기만** 한다 (R-106).

    이 모듈에는 차트 프레임을 쓰는 코드가 없다. 검수 쪽에서 '건드리지 않았음'을
    확인할 때 쓰라고 읽기 함수만 둔다.
    """
    out = []
    for frame in tree.iter(q(P, "graphicFrame")):
        data = frame.find(f"{q(A, 'graphic')}/{q(A, 'graphicData')}")
        if data is None or data.find(q(C, "chart")) is None:
            continue
        out.append(frame_box(frame))
    return out


def distribute_row_heights(total: int, n: int) -> list[int]:
    """데이터 영역 높이를 n행에 나눈다. **합이 정확히 total** 이 되게 (R-105).

    나머지 EMU 는 첫 행에 몰아 준다 -- 그래야 마지막 행의 하단 실선(lnB) 위치가
    표 bottom 과 정확히 맞는다.
        distribute_row_heights(3096000, 5) -> [619200]*5
        distribute_row_heights(3096000, 7) -> [442290, 442285, ... ]  (합 3096000)
    """
    if n < 1:
        raise ValueError("행 수는 1 이상")
    base, rem = divmod(total, n)
    out = [base] * n
    out[0] += rem
    return out


# --------------------------------------------------------------------------- #
# 표 -- 셀 텍스트 (R-096)
# --------------------------------------------------------------------------- #

def _first_para(tc: etree._Element) -> etree._Element | None:
    body = tc.find(q(A, "txBody"))
    if body is None:
        return None
    para = body.find(q(A, "p"))
    if para is None:
        para = etree.SubElement(body, q(A, "p"))
    return para


def cell_text(tc: etree._Element) -> str:
    """셀 텍스트. 문단 사이는 \\n, 수동 줄바꿈 <a:br/> 도 \\n 으로 읽는다."""
    body = tc.find(q(A, "txBody"))
    if body is None:
        return ""
    paras = []
    for para in body.findall(q(A, "p")):
        buf = []
        for node in para:
            if node.tag == q(A, "br"):
                buf.append("\n")
            elif node.tag in (q(A, "r"), q(A, "fld")):
                t = node.find(q(A, "t"))
                if t is not None and t.text:
                    buf.append(t.text)
        paras.append("".join(buf))
    return "\n".join(paras)


def clear_text(tc: etree._Element) -> None:
    """셀 텍스트를 비우고 endParaRPr 상태로 되돌린다 (서식 보존).

    원본 셀은 런이 여러 개이거나(자동 줄바꿈으로 쪼개짐) 문단이 2개인 경우가 있다
    (LG s23 라벨열 10칸 중 3칸이 런 3개 / 문단 2개). 첫 런의 rPr 만 남기고 나머지를
    전부 지워야 옛 글자가 남지 않는다 -- 문단만 지우면 두 번째 런의 글자가 살아남는다.
    """
    body = tc.find(q(A, "txBody"))
    if body is None:
        return
    paras = body.findall(q(A, "p"))
    if not paras:
        etree.SubElement(body, q(A, "p"))
        return

    keep = paras[0]
    for extra in paras[1:]:               # 둘째 문단부터 통째로 제거
        body.remove(extra)

    rpr = None
    for node in list(keep):
        if node.tag in (q(A, "r"), q(A, "fld")):
            if rpr is None:
                found = node.find(q(A, "rPr"))
                if found is not None:
                    rpr = found
            keep.remove(node)
        elif node.tag == q(A, "br"):
            keep.remove(node)

    if keep.find(q(A, "endParaRPr")) is not None:
        return
    if rpr is not None:
        end = copy.deepcopy(rpr)
        end.tag = q(A, "endParaRPr")
        keep.append(end)


def set_text(tc: etree._Element, text: str, lines: Sequence[str] | None = None) -> None:
    """셀에 텍스트를 넣는다. endParaRPr 을 run 으로 승격해 서식을 100% 보존한다.

    lines 를 주면 줄 사이에 <a:br/> 를 넣어 **수동 줄바꿈**한다 (R-108).
    PowerPoint 자동 줄바꿈은 greedy 라 마지막 줄에 한 어절만 남는 고아가 생긴다.
    줄 나누기 자체는 textfit.wrap_ko 가 정한다 -- 여기서는 받은 대로 박기만 한다.

    ※ 이미 글자가 있는 셀에 그냥 부르면 기존 런을 갈아엎는다. 서식이 여러 갈래인
      셀이라면 clear_text() 를 먼저 부르는 편이 예측 가능하다.
    """
    para = _first_para(tc)
    if para is None:
        return
    parts = [str(x) for x in lines] if lines else [text]
    if not parts:
        parts = [""]

    proto = None
    end = para.find(q(A, "endParaRPr"))
    if end is not None:
        proto = copy.deepcopy(end)
        proto.tag = q(A, "rPr")
        para.remove(end)
    else:
        for node in list(para):
            if node.tag in (q(A, "r"), q(A, "fld")):
                if proto is None:
                    found = node.find(q(A, "rPr"))
                    if found is not None:
                        proto = copy.deepcopy(found)
                para.remove(node)
            elif node.tag == q(A, "br"):
                para.remove(node)

    for i, line in enumerate(parts):
        if i:
            br = etree.SubElement(para, q(A, "br"))
            if proto is not None:
                br.append(copy.deepcopy(proto))
        run = etree.SubElement(para, q(A, "r"))
        if proto is not None:
            run.append(copy.deepcopy(proto))
        etree.SubElement(run, q(A, "t")).text = line


def put_cell(tc: etree._Element, text: str, lines: Sequence[str] | None = None) -> None:
    """clear_text + set_text. 일반적인 값 주입은 이걸 쓴다."""
    clear_text(tc)
    set_text(tc, text, lines)


# --------------------------------------------------------------------------- #
# 표 -- 라벨 폭 실측 (R-108)
# --------------------------------------------------------------------------- #

def label_metrics(tbl: etree._Element, *, header_rows: int = HEADER_ROWS,
                  row: int | None = None, col: int = 0) -> tuple[int, float]:
    """보기 라벨 열의 (가용 폭 EMU, 글꼴 크기 pt).

    가용 폭 = tblGrid[col].w - tcPr.marL - tcPr.marR   (R-108)
    셀이 가로 병합(gridSpan)돼 있으면 걸친 열 너비를 합산한다.
    row 는 측정할 행 인덱스(기본 = 첫 데이터 행). 헤더 행은 글꼴이 다르므로 쓰지 않는다.
    글꼴 이름은 표에 안 적혀 있고 테마 상속이라 Deck.theme_font() 로 따로 얻는다.
    """
    widths = grid_widths(tbl)
    if not widths:
        return 0, 10.0
    rows = table_rows(tbl)
    ridx = header_rows if row is None else row
    ridx = max(0, min(ridx, len(rows) - 1))
    cells = row_cells(rows[ridx])
    if col >= len(cells):
        return widths[col] if col < len(widths) else 0, 10.0
    tc = cells[col]

    span = _int(tc.get("gridSpan"), 1) or 1
    width = sum(widths[col:col + span])

    pr = tc.find(q(A, "tcPr"))
    mar_l = _int(pr.get("marL"), DEFAULT_CELL_MARGIN) if pr is not None else DEFAULT_CELL_MARGIN
    mar_r = _int(pr.get("marR"), DEFAULT_CELL_MARGIN) if pr is not None else DEFAULT_CELL_MARGIN

    size = None
    for el in tc.iter():
        if el.tag in (q(A, "rPr"), q(A, "endParaRPr"), q(A, "defRPr")) and el.get("sz"):
            size = _int(el.get("sz")) / 100.0
            break
    return max(0, width - mar_l - mar_r), size or 10.0


# --------------------------------------------------------------------------- #
# 표 -- 행 개수 조정 (R-105 / R-107)
# --------------------------------------------------------------------------- #

def row_border_dashes(tr: etree._Element, col: int = 0) -> tuple[str | None, str | None]:
    """행의 (lnT prstDash, lnB prstDash). R-107 행 서식 3종을 판정하는 실제 신호다.

    ★ lnB 엘리먼트는 **모든 행에 있다.** 'lnB 가 있느냐'로 마지막 행을 가리면 안 된다.
      갈리는 건 prstDash 값이다. LG s17/s18/s23 실측:
          첫 데이터행   lnT=solid  lnB=dot
          중간행        lnT=dot    lnB=dot
          마지막행      lnT=dot    lnB=solid   <- 표 하단 실선
    """
    cells = row_cells(tr)
    if col >= len(cells):
        return None, None
    pr = cells[col].find(q(A, "tcPr"))
    if pr is None:
        return None, None
    out = []
    for tag in ("lnT", "lnB"):
        ln = pr.find(q(A, tag))
        dash = ln.find(q(A, "prstDash")) if ln is not None else None
        out.append(dash.get("val") if dash is not None else None)
    return out[0], out[1]


def derive_mid_row(first: etree._Element, last: etree._Element) -> etree._Element:
    """중간 행 템플릿을 첫 행·마지막 행에서 합성한다 (R-107 폴백).

    데이터 행이 2개뿐인 표에는 중간 행 템플릿이 없다. 첫 행을 그대로 복제하면
    lnT=solid 가 따라와 **표 한가운데 헤더 밑줄이 그어진다.**
    중간 행의 두 테두리는 표 안에 이미 존재한다 -- 위쪽은 마지막 행의 lnT(=dot),
    아래쪽은 첫 행의 lnB(=dot) 다. 그 둘을 조립하면 서식을 지어내지 않고도
    정확한 중간 행이 나온다.

    ※ 그래도 R-107 정본은 '같은 클론 패밀리의 다른 장표에서 진짜 중간 행을 가져오는 것'
      이다. 어느 장표에서 가져올지는 미결이므로, 이 합성은 그 결정이 서기 전의 대안이다.
    """
    clone = copy.deepcopy(first)
    src_cells = row_cells(last)
    for i, tc in enumerate(row_cells(clone)):
        if i >= len(src_cells):
            break
        pr = tc.find(q(A, "tcPr"))
        src_pr = src_cells[i].find(q(A, "tcPr"))
        if pr is None or src_pr is None:
            continue
        src_top = src_pr.find(q(A, "lnT"))
        if src_top is None:
            continue
        old_top = pr.find(q(A, "lnT"))
        new_top = copy.deepcopy(src_top)
        if old_top is not None:
            pr.replace(old_top, new_top)
        else:
            pr.insert(0, new_top)
    return clone


def set_row_count(tbl: etree._Element, n: int, *,
                  header_rows: int = HEADER_ROWS,
                  data_area_h: int | None = None,
                  mid_template: etree._Element | None = None,
                  mid_style: str = "clone",
                  clear: bool = True,
                  on_warn: Callable[[str], None] | None = None) -> list[etree._Element]:
    """데이터 행을 n개로 맞추고 높이를 재분배한다. 첫/중간/마지막 서식을 유지한다.

    R-107  행 서식은 [첫행][중간행...][마지막행] 3종이고 **마지막 행만 lnB=solid** 다.
           그래서 축소는 중간 행부터 지우고(첫·마지막은 절대 안 건드린다),
           확대는 중간 행을 복제해 마지막 행 **앞에** 끼운다.
    R-105  데이터 영역 높이의 합을 보존하고 나머지 EMU 는 첫 행에 몰아 준다.
           data_area_h 를 안 주면 **현재 표에서 실측**한다 -- 그게 정상 경로다
           (표준 배너표 3,096,000 / 가로누적100% 1,512,000 / SKEE 는 또 다름).
    R-106  차트 graphicFrame 은 손대지 않는다. 이 함수는 차트를 아예 열지 않는다.

    mid_template  데이터 행이 2개뿐인 표에는 중간 행 템플릿이 없다. 그때는 같은 클론
                  패밀리의 다른 장표에서 <a:tr> 을 떼어 넘긴다(deepcopy 해서 쓴다).
                  R-107 정본 경로다.
    mid_style     mid_template 이 없고 중간 행도 없을 때의 폴백.
                  'clone'  첫 행을 그대로 복제한다(프로토타입 동작). 첫 행은
                           lnT=solid 라 **표 한가운데 헤더 밑줄이 그어진다** --
                           반드시 on_warn 을 받아 검수 태그로 사람에게 넘겨야 한다.
                  'derive' derive_mid_row() 로 첫 행 + 마지막 행에서 합성한다.
                           테두리를 지어내지 않고 표 안에 있는 선만 조립한다.
                  기본이 'clone' 인 이유: '어느 장표에서 중간 행을 가져올지'가 아직
                  미결이라, 모듈이 조용히 결정을 대신하지 않고 호출자에게 알린다.
    clear         복제한 행의 셀 텍스트를 비울지. False 면 원본 값이 따라온다.
    on_warn       경고 문자열 콜백. fill_deck 이 리포트 태그로 받아 간다.

    반환: 조정 후의 데이터 행 리스트 (호출자가 곧바로 값 주입에 쓴다)
    """
    warn = on_warn or (lambda _msg: None)
    rows = data_rows(tbl, header_rows)
    if not rows:
        raise ValueError(f"데이터 행이 없다 (전체 {len(table_rows(tbl))}행, header_rows={header_rows})")
    if n < 2:
        raise ValueError("데이터 행은 최소 2개 (R-107: 첫행/마지막행 서식이 다르다)")

    if data_area_h is None:
        data_area_h = sum(_int(tr.get("h")) for tr in rows)
        if data_area_h <= 0:
            data_area_h = DATA_AREA_H
            warn(f"R-105 데이터 영역 높이를 실측하지 못해 기본값 {DATA_AREA_H} 를 쓴다")

    first, last = rows[0], rows[-1]
    template = mid_template
    if template is None and len(rows) >= 3:
        template = rows[1]
    cur = len(rows)

    if n < cur:
        # 축소: 중간 행부터 제거. 첫 행과 마지막 행은 남긴다
        for tr in rows[1:1 + (cur - n)]:
            tbl.remove(tr)
    elif n > cur:
        if template is None:
            if mid_style == "derive":
                template = derive_mid_row(first, last)
                warn("R-107 중간 행 템플릿이 없어 첫 행+마지막 행에서 합성했다 "
                     "(데이터 행 2개짜리 표) -- 테두리 확인 권장")
            else:
                template = first
                warn("R-107 중간 행 템플릿이 없어 첫 행을 복제했다 (데이터 행 2개짜리 표) "
                     "-- 첫 행은 lnT=solid 라 중간에 실선이 그어진다. "
                     "mid_template 로 같은 패밀리의 중간 행을 넘기거나 mid_style='derive' 를 쓴다")
        want_cells = len(row_cells(first))
        got_cells = len(row_cells(template))
        if got_cells != want_cells:
            warn(f"R-107 중간 행 템플릿의 셀 수가 다르다 ({got_cells} != {want_cells})")
        idx = list(tbl).index(last)
        for _ in range(n - cur):
            clone = copy.deepcopy(template)
            if clear:
                for tc in row_cells(clone):
                    clear_text(tc)
            tbl.insert(idx, clone)
            idx += 1

    rows = data_rows(tbl, header_rows)
    for tr, h in zip(rows, distribute_row_heights(data_area_h, n)):
        tr.set("h", str(h))
    return rows


def verify_row_heights(tbl: etree._Element,
                       header_rows: int = HEADER_ROWS,
                       expect_total: int | None = None) -> dict[str, Any]:
    """행 높이 합이 맞는지 확인한다. 스모크 테스트와 검수(verify_deck)용.

    반환 {'header_h','data_area_h','total_h','row_heights','expect_total','ok'}
    expect_total 을 안 주면 데이터 영역 합만 보고하고 ok 는 None 이다.
    """
    geo = measure_table_geometry(tbl, header_rows)
    ok = None if expect_total is None else (geo.data_area_h == expect_total)
    return {
        "header_h": geo.header_h,
        "data_area_h": geo.data_area_h,
        "total_h": geo.total_h,
        "row_heights": geo.row_heights,
        "frame_cy": geo.frame_cy,
        "expect_total": expect_total,
        "ok": ok,
    }


# --------------------------------------------------------------------------- #
# 차트 -- 읽기
# --------------------------------------------------------------------------- #

def chart_series(chart_tree: etree._Element) -> list[etree._Element]:
    """c:ser 전부. 문서 등장 순서 = 겹침 순서다 (R-109: order=0 이 뒤에 깔린다)."""
    return chart_tree.findall(f".//{q(C, 'ser')}")


def series_names(chart_tree: etree._Element) -> list[str | None]:
    """계열 이름 목록.

    R-094: **계열의 개수·순서·이름은 제작본이 진리다.** 값만 갈아 끼운다.
    이름이 안 붙는 더미 계열(LG 실측 'do not know', '열2')은 호출자가 건너뛴다.
    """
    out: list[str | None] = []
    for ser in chart_series(chart_tree):
        tx = ser.find(q(C, "tx"))
        name = None
        if tx is not None:
            v = tx.find(f"{q(C, 'strRef')}/{q(C, 'strCache')}/{q(C, 'pt')}/{q(C, 'v')}")
            if v is None:
                v = tx.find(q(C, "v"))
            if v is not None and v.text:
                name = v.text
        out.append(name)
    return out


def series_values(chart_tree: etree._Element) -> list[list[float | None]]:
    """계열별 현재 값. 미기입 판정(더미 상수 검출)의 재료다."""
    out: list[list[float | None]] = []
    for ser in chart_series(chart_tree):
        cache = ser.find(f"{q(C, 'val')}/{q(C, 'numRef')}/{q(C, 'numCache')}")
        if cache is None:
            out.append([])
            continue
        count = _int(cache.find(q(C, "ptCount")).get("val")) if cache.find(q(C, "ptCount")) is not None else 0
        vals: list[float | None] = [None] * count
        for pt in cache.findall(q(C, "pt")):
            i = _int(pt.get("idx"), -1)
            v = pt.find(q(C, "v"))
            if 0 <= i < count and v is not None and v.text not in (None, ""):
                try:
                    vals[i] = float(v.text)
                except ValueError:
                    vals[i] = None
        out.append(vals)
    return out


def chart_categories(chart_tree: etree._Element) -> list[str]:
    """첫 계열의 카테고리 라벨."""
    sers = chart_series(chart_tree)
    if not sers:
        return []
    cat = sers[0].find(q(C, "cat"))
    if cat is None:
        return []
    count = 0
    cnt = cat.find(f".//{q(C, 'ptCount')}")
    if cnt is not None:
        count = _int(cnt.get("val"))
    out = [""] * count
    for pt in cat.findall(f".//{q(C, 'pt')}"):
        i = _int(pt.get("idx"), -1)
        v = pt.find(q(C, "v"))
        if 0 <= i < count:
            out[i] = (v.text or "") if v is not None else ""
    return out


# --------------------------------------------------------------------------- #
# 차트 -- 임베디드 워크시트 좌표계
#
# ★ 워크시트 배치가 차트타입마다 다르다. LG 실측 (프로토타입이 놓쳤던 지점):
#     chart64 겹침막대     tx=$B$1  cat=$A$2:$A$11  val=$B$2:$B$11   계열이 **열**로 선다
#     chart78 ±세로막대    tx=$A$2  cat=$B$1:$L$1   val=$B$2:$L$2    계열이 **행**으로 눕는다
#     chart66 가로누적100% tx=$A$2  cat=$B$1        val=$B$2         행 배치 + 카테고리 1개
#   전부 `$COL$2:$COL$(n+1)` 로 고쳐 쓰면 척도 장표(LG 36장 · SKEE 44장)의 워크시트
#   연결이 통째로 깨진다. 그래서 범위의 **모양**을 읽어 방향을 판정한다.
# --------------------------------------------------------------------------- #

_A1_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def col_to_index(col: str) -> int:
    """'A' -> 1, 'Z' -> 26, 'AA' -> 27."""
    idx = 0
    for ch in col.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx


def index_to_col(idx: int) -> str:
    """1 -> 'A', 27 -> 'AA'."""
    if idx < 1:
        raise ValueError(f"열 번호는 1 이상: {idx}")
    out = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        out = chr(65 + rem) + out
    return out


def parse_ref(text: str | None) -> dict[str, Any] | None:
    """`Sheet1!$B$2:$L$2` -> {'sheet','c1','r1','c2','r2','single'}.

    시트 이름은 원본 그대로 보존한다 (`Sheet1` 으로 박지 않는다).
    """
    if not text or "!" not in text:
        return None
    sheet, ref = text.rsplit("!", 1)
    parts = ref.split(":")
    m1 = _A1_RE.match(parts[0].strip())
    if m1 is None:
        return None
    c1, r1 = m1.group(1).upper(), int(m1.group(2))
    if len(parts) == 1:
        c2, r2 = c1, r1
    else:
        m2 = _A1_RE.match(parts[1].strip())
        if m2 is None:
            return None
        c2, r2 = m2.group(1).upper(), int(m2.group(2))
    return {"sheet": sheet, "c1": c1, "r1": r1, "c2": c2, "r2": r2,
            "single": len(parts) == 1}


def _series_ref(ser: etree._Element, kind: str) -> dict[str, Any] | None:
    """계열의 c:tx / c:cat / c:val 안에 있는 c:f 를 파싱한다."""
    node = ser.find(q(C, kind))
    if node is None:
        return None
    f = node.find(f".//{q(C, 'f')}")
    return parse_ref(f.text) if f is not None else None


def series_orientation(ser: etree._Element) -> str:
    """이 계열이 워크시트에서 'cols'(세로) 인지 'rows'(가로) 인지.

    ① val 범위 모양이 가장 확실하다 -- 한 열에 걸쳐 있으면 cols, 한 행이면 rows
    ② 범위가 단일 셀이면 계열명(tx) 위치로 본다 -- 1행에 있으면 cols, A열에 있으면 rows
    ③ 그래도 모르면 cat 범위 모양으로, 마지막엔 cols 로 떨어진다
    """
    val = _series_ref(ser, "val")
    if val and not val["single"]:
        if val["r1"] == val["r2"] and val["c1"] != val["c2"]:
            return "rows"
        if val["c1"] == val["c2"] and val["r1"] != val["r2"]:
            return "cols"
    tx = _series_ref(ser, "tx")
    if tx:
        if tx["r1"] == 1 and tx["c1"] != "A":
            return "cols"
        if tx["c1"] == "A":
            return "rows"
    cat = _series_ref(ser, "cat")
    if cat and not cat["single"]:
        if cat["r1"] == cat["r2"] and cat["c1"] != cat["c2"]:
            return "rows"
        if cat["c1"] == cat["c2"] and cat["r1"] != cat["r2"]:
            return "cols"
    return "cols"


@dataclass
class ChartLayout:
    """차트가 임베디드 워크시트의 어디를 보는지. update_embedded_workbook 의 좌표계다."""
    sheet: str = "Sheet1"
    orientation: str = "cols"        # 'cols' = 계열이 열, 'rows' = 계열이 행
    cat_col: str = "A"                              # cols 배치의 카테고리 라벨 열
    cat_row: int = 1                                # rows 배치의 카테고리 라벨 행
    first_data_col: str = WORKBOOK_FIRST_DATA_COL   # rows 배치에서 값이 시작하는 열
    first_data_row: int = WORKBOOK_FIRST_DATA_ROW   # cols 배치에서 값이 시작하는 행
    val_cols: list[str | None] = field(default_factory=list)
    val_rows: list[int | None] = field(default_factory=list)
    name_cells: list[tuple[str, int] | None] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return {k: getattr(self, k) for k in (
            "sheet", "orientation", "cat_col", "cat_row", "first_data_col",
            "first_data_row", "val_cols", "val_rows", "name_cells")}


def chart_layout(chart_tree: etree._Element) -> ChartLayout:
    """차트의 c:f 수식에서 워크시트 좌표계를 복원한다."""
    sers = chart_series(chart_tree)
    lay = ChartLayout()
    if not sers:
        return lay
    lay.orientation = series_orientation(sers[0])

    cat = _series_ref(sers[0], "cat")
    if cat:
        lay.sheet = cat["sheet"]
        if lay.orientation == "rows":
            lay.cat_row = cat["r1"]
            lay.first_data_col = cat["c1"]
        else:
            lay.cat_col = cat["c1"]
            lay.first_data_row = cat["r1"]

    for ser in sers:
        val = _series_ref(ser, "val")
        tx = _series_ref(ser, "tx")
        if val:
            lay.sheet = lay.sheet or val["sheet"]
            lay.val_cols.append(val["c1"])
            lay.val_rows.append(val["r1"])
            if lay.orientation == "rows":
                lay.first_data_col = val["c1"]
            else:
                lay.first_data_row = val["r1"]
        else:
            lay.val_cols.append(None)
            lay.val_rows.append(None)
        lay.name_cells.append((tx["c1"], tx["r1"]) if tx else None)
    return lay


# --------------------------------------------------------------------------- #
# 차트 -- 쓰기
# --------------------------------------------------------------------------- #

def sync_chart(chart_tree: etree._Element,
               series: Sequence[Sequence[float | None]] | Sequence[float | None],
               categories: Sequence[str] | None = None,
               *, strict: bool = True) -> list[str | None]:
    """차트 계열 값을 갈아 끼운다. 다계열 지원.

    series      [[v, ...], ...] 를 **문서 순서(c:ser 등장 순서)대로** 준다.
                단일 계열이면 [[...]] 또는 그냥 [...] 도 받는다.
                None 은 '데이터 없음'으로 넣는다 -- 0 이 아니다. 제작팀 메모
                `범례만큼 데이터가 없을때는 빈칸으로 유지해주세요` 가 근거다.
    categories  주면 카테고리 라벨까지 다시 쓴다. 안 주면 기존 라벨을 n개로 자른다
                (자르지 않으면 ptCount 와 pt 개수가 어긋나 PowerPoint 가 복구를 띄운다).
    strict      계열 수가 안 맞으면 ValueError. R-096 -- 계열 개수를 바꾸면 색·라벨
                서식이 어긋나므로, 안 맞으면 고치지 말고 사람에게 넘겨야 한다.

    반환: 계열별 값 열 문자 [`B`, `C`, ...] -- 임베디드 워크시트를 채울 때 쓴다.
          c:f 가 없는 계열은 None.
    """
    if series and not isinstance(series[0], (list, tuple)):
        series = [series]                                   # type: ignore[assignment]
    sers = chart_series(chart_tree)
    if len(sers) != len(series):
        msg = f"차트 계열 {len(sers)}개인데 값은 {len(series)}세트 (R-094: 계열은 제작본이 진리다)"
        if strict:
            raise ValueError(msg)
        series = list(series)[:len(sers)]

    cols: list[str | None] = []
    for ser, values in zip(sers, series):
        n = len(values)
        _write_num_cache(ser, values)
        _sync_categories(ser, n, categories)
        _prune_indexed(ser, n)
        cols.append(_rewrite_refs(ser, n, series_orientation(ser)))
    return cols


def _write_num_cache(ser: etree._Element, values: Sequence[float | None]) -> None:
    """c:val/c:numRef/c:numCache 를 다시 쓴다. None 은 pt 를 아예 안 만든다(빈칸)."""
    cache = ser.find(f"{q(C, 'val')}/{q(C, 'numRef')}/{q(C, 'numCache')}")
    if cache is None:
        raise ValueError("c:val/c:numRef/c:numCache 가 없는 계열이다 (지원 밖 차트타입)")
    for pt in cache.findall(q(C, "pt")):
        cache.remove(pt)
    cnt = cache.find(q(C, "ptCount"))
    if cnt is None:
        cnt = etree.Element(q(C, "ptCount"))
        cache.insert(0, cnt)
    cnt.set("val", str(len(values)))
    for i, v in enumerate(values):
        if v is None:
            continue                       # 빈칸: pt 를 만들지 않는다 (0 과 다르다)
        pt = etree.SubElement(cache, q(C, "pt"))
        pt.set("idx", str(i))
        etree.SubElement(pt, q(C, "v")).text = _num_text(v)


def _num_text(v: float | int) -> str:
    """차트 캐시에 넣을 숫자 문자열. -0.0 을 만들지 않는다.

    ±세로막대(척도)는 numFmt 가 음수 부호를 숨기지만 값 자체가 -0.0 이면
    PowerPoint 가 '-0.0' 을 계산에 쓴다. 0 은 언제나 0 으로 넣는다.
    """
    f = float(v)
    if f == 0:
        f = 0.0
    if f == int(f) and abs(f) < 1e15:
        return f"{f:.1f}" if isinstance(v, float) else str(int(f))
    return repr(f)


def _sync_categories(ser: etree._Element, n: int,
                     categories: Sequence[str] | None) -> None:
    """카테고리 캐시를 n개로 맞춘다. categories 를 주면 라벨까지 다시 쓴다."""
    cat = ser.find(q(C, "cat"))
    if cat is None:
        return
    cache = cat.find(f"{q(C, 'strRef')}/{q(C, 'strCache')}")
    if cache is None:
        cache = cat.find(f"{q(C, 'numRef')}/{q(C, 'numCache')}")
    if cache is None:
        return

    if categories is not None:
        for pt in cache.findall(q(C, "pt")):
            cache.remove(pt)
        cnt = cache.find(q(C, "ptCount"))
        if cnt is None:
            cnt = etree.Element(q(C, "ptCount"))
            cache.insert(0, cnt)
        cnt.set("val", str(len(categories)))
        for i, label in enumerate(categories):
            pt = etree.SubElement(cache, q(C, "pt"))
            pt.set("idx", str(i))
            etree.SubElement(pt, q(C, "v")).text = str(label)
        return

    # 라벨을 안 줬으면 개수만 맞춘다 -- 남는 pt 는 지운다
    for pt in cache.findall(q(C, "pt")):
        if _int(pt.get("idx"), -1) >= n:
            cache.remove(pt)
    cnt = cache.find(q(C, "ptCount"))
    if cnt is not None:
        cnt.set("val", str(n))


def _prune_indexed(ser: etree._Element, n: int) -> None:
    """범위를 벗어난 c:dPt / c:dLbl 을 지운다.

    보기 개수를 줄였는데 idx=7 짜리 데이터 라벨 override 가 남아 있으면
    PowerPoint 가 '복구가 필요합니다' 를 띄우거나 유령 라벨을 그린다.
    """
    for tag in ("dPt", "dLbl"):
        for node in ser.findall(f".//{q(C, tag)}"):
            idx = node.find(q(C, "idx"))
            if idx is not None and _int(idx.get("val"), -1) >= n:
                node.getparent().remove(node)


def _emit_ref(ref: dict[str, Any], n: int, orientation: str) -> str:
    """참조 범위를 데이터 n개짜리로 다시 쓴다. 시작점(c1/r1)과 방향은 그대로 둔다."""
    sheet, c1, r1 = ref["sheet"], ref["c1"], ref["r1"]
    if n <= 1:
        return f"{sheet}!${c1}${r1}"
    if orientation == "rows":                    # 한 행을 따라 가로로 뻗는다
        end = index_to_col(col_to_index(c1) + n - 1)
        return f"{sheet}!${c1}${r1}:${end}${r1}"
    return f"{sheet}!${c1}${r1}:${c1}${r1 + n - 1}"   # 한 열을 따라 세로로 뻗는다


def _rewrite_refs(ser: etree._Element, n: int, orientation: str) -> str | None:
    """c:cat / c:val 의 c:f 범위를 데이터 n개에 맞춘다. 값 범위의 시작 열 문자를 돌려준다.

    ★ 계열명(c:tx)의 참조는 건드리지 않는다 -- R-094, 계열 이름은 제작본이 진리다.
    ★ 방향(orientation)을 무시하고 전부 세로 범위로 고치면 척도 장표가 깨진다.
      chart78 실측: val=$B$2:$L$2 (가로) 를 $B$2:$B$12 로 바꾸면 워크시트 연결이 끊긴다.
    """
    val_col = None
    for kind in ("cat", "val"):
        node = ser.find(q(C, kind))
        if node is None:
            continue
        for f in node.iter(q(C, "f")):
            ref = parse_ref(f.text)
            if ref is None:
                continue
            f.text = _emit_ref(ref, n, orientation)
            if kind == "val":
                val_col = ref["c1"]
    return val_col


def _put(ws, row: int, col: int, value: Any) -> None:
    """셀에 값을 대입한다. None 이면 빈칸으로 만든다.

    ★ openpyxl 의 `ws.cell(row, col, value=None)` 은 **대입을 건너뛴다** -- None 을
      '값을 안 줬다'로 읽기 때문이다. 그래서 그 API 로는 셀을 비울 수 없고,
      '데이터 없음은 0 이 아니라 빈칸' 규칙이 조용히 깨진다. 반드시 .value 로 대입한다.
    """
    ws.cell(row=row, column=col).value = value


def _is_formula(ws, row: int, col: int) -> bool:
    v = ws.cell(row=row, column=col).value
    return isinstance(v, str) and v.startswith("=")


def workbook_grid(deck: Deck, chart_part: str) -> list[list[Any]]:
    """임베디드 워크시트를 2차원 리스트로 읽는다 (1행 1열이 [0][0]).

    척도 차트처럼 워크시트에 별도 입력 블록이 있는 경우, 호출자가 그 위치를
    찾아야 하므로 원본 그대로 보여 준다. 수식 셀은 `'=B15'` 같은 문자열로 나온다.
    """
    part = deck.workbook_part(chart_part)
    if part is None or part not in deck.parts:
        return []
    import openpyxl
    ws = openpyxl.load_workbook(io.BytesIO(deck.parts[part])).worksheets[0]
    return [[ws.cell(row=r, column=c).value for c in range(1, (ws.max_column or 0) + 1)]
            for r in range(1, (ws.max_row or 0) + 1)]


def write_workbook_cells(deck: Deck, chart_part: str,
                         cells: dict[tuple[int, int], Any]) -> str | None:
    """워크시트 셀을 직접 쓴다. {(행, 열): 값}, 둘 다 1-based.

    척도 차트의 `▼데이터 입력칸` 처럼 차트 범위 **밖**을 고쳐야 할 때 쓰는 원시 도구다.
    무엇을 어디에 쓸지는 호출자가 정한다 -- deckops 는 좌표를 해석하지 않는다.
    """
    part = deck.workbook_part(chart_part)
    if part is None or part not in deck.parts:
        return None
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(deck.parts[part]))
    ws = wb.worksheets[0]
    for (row, col), value in cells.items():
        _put(ws, row, col, value)
    buf = io.BytesIO()
    wb.save(buf)
    deck.put_bytes(part, buf.getvalue())
    return part


def update_embedded_workbook(deck: Deck, chart_part: str,
                             categories: Sequence[str],
                             series: Sequence[Sequence[float | None]],
                             *, layout: ChartLayout | None = None,
                             series_headers: Sequence[str | None] | None = None,
                             preserve_formulas: bool = True,
                             on_warn: Callable[[str], None] | None = None) -> str | None:
    """차트의 임베디드 워크시트를 같은 값으로 다시 쓴다.

    ★ 이걸 안 하면 **사용자가 차트를 더블클릭하는 순간 옛 값으로 되돌아간다.**
      차트 XML 의 캐시는 화면 표시용이고, 편집할 때의 진짜 소스는 이 xlsx 다.

    배치(cols/rows)는 layout 이 정한다. 안 주면 차트에서 읽는다.
    남는 옛 데이터는 행을 지우지 않고 **n 뒤의 셀만 비운다** -- 행 삭제는 가로 배치
    차트에서 다른 계열까지 날린다.

    preserve_formulas
      ★ 척도(±세로막대) 차트의 워크시트는 **수식으로 굴러간다.** LG chart78 실측:
            r1     ' ' | 전체 | 20대 | ...            <- 카테고리(배너)
            r2~r8  5점 | =B15 | =C15 ...              <- 차트가 보는 범위 (전부 수식)
                   4점 | =B14*-1 ...                  <- 중립점 아래는 -1 을 곱한다
            r10    ▼데이터 입력칸
            r11~17 1전혀 그렇지 않다 | 3 | ... | 실제 %  <- 사람이 넣는 원본
        여기에 생값을 덮어쓰면 발산형(±) 차트를 만들던 부호 반전 장치가 사라진다.
        그래서 기본값 True 면 **수식 셀은 건드리지 않고** 몇 개를 보존했는지 알린다.
        그 경우 호출자가 write_workbook_cells() 로 입력 블록(r11~17)을 채워야 한다.

    openpyxl 은 여기서만 필요하므로 지연 임포트한다.
    반환: 갱신한 워크시트 파트 이름 (임베디드가 없으면 None)
    """
    warn = on_warn or (lambda _m: None)
    part = deck.workbook_part(chart_part)
    if part is None or part not in deck.parts:
        return None

    import openpyxl

    lay = layout or chart_layout(deck.xml(chart_part))
    wb = openpyxl.load_workbook(io.BytesIO(deck.parts[part]))
    ws = wb.worksheets[0]
    n = len(categories)
    kept = 0

    def write(row: int, col: int, value: Any) -> None:
        nonlocal kept
        if preserve_formulas and _is_formula(ws, row, col):
            kept += 1
            return
        _put(ws, row, col, value)

    if lay.orientation == "rows":
        base = col_to_index(lay.first_data_col)
        last_col = max(ws.max_column or base, base + n - 1)
        for i, label in enumerate(categories):                 # 카테고리는 가로로
            write(lay.cat_row, base + i, label)
        for c in range(base + n, last_col + 1):
            write(lay.cat_row, c, None)
        for s, values in enumerate(series):                    # 계열마다 자기 행을 따라
            row = lay.val_rows[s] if s < len(lay.val_rows) else None
            if row is None:
                row = lay.cat_row + 1 + s
            for i, v in enumerate(values):
                write(row, base + i, v)
            for c in range(base + len(values), last_col + 1):
                write(row, c, None)
            _write_series_header(ws, lay, s, series_headers)
    else:
        base = lay.first_data_row
        cat_idx = col_to_index(lay.cat_col)
        last_row = max(ws.max_row or base, base + n - 1)
        for i, label in enumerate(categories):                 # 카테고리는 세로로
            write(base + i, cat_idx, label)
        for r in range(base + n, last_row + 1):
            write(r, cat_idx, None)
        for s, values in enumerate(series):                    # 계열마다 자기 열을 따라
            col = lay.val_cols[s] if s < len(lay.val_cols) else None
            idx = col_to_index(col) if col else (cat_idx + 1 + s)
            for i, v in enumerate(values):
                write(base + i, idx, v)
            for r in range(base + len(values), last_row + 1):
                write(r, idx, None)
            _write_series_header(ws, lay, s, series_headers)

    if kept:
        warn(f"임베디드 워크시트의 수식 셀 {kept}개를 보존했다 (척도 차트의 부호 반전 장치). "
             f"입력 블록은 write_workbook_cells() 로 따로 채워야 한다")

    buf = io.BytesIO()
    wb.save(buf)
    deck.put_bytes(part, buf.getvalue())
    return part


def _write_series_header(ws, lay: ChartLayout, s: int,
                         series_headers: Sequence[str | None] | None) -> None:
    """계열명을 tx 셀에 쓴다. series_headers 를 준 경우에만.

    기본은 안 쓴다 -- R-094 대로 계열 이름은 제작본이 진리이고 워크시트에 이미 맞게
    들어 있다. 새 슬라이드를 만들 때만 넘긴다.
    """
    if not series_headers or s >= len(series_headers) or series_headers[s] is None:
        return
    cell = lay.name_cells[s] if s < len(lay.name_cells) else None
    if cell is None:
        return
    _put(ws, cell[1], col_to_index(cell[0]), series_headers[s])


def apply_chart(deck: Deck, chart_part: str,
                series: Sequence[Sequence[float | None]],
                categories: Sequence[str] | None = None,
                *, strict: bool = True,
                update_workbook: bool = True,
                preserve_formulas: bool = True,
                series_headers: Sequence[str | None] | None = None,
                on_warn: Callable[[str], None] | None = None) -> dict[str, Any]:
    """차트 XML + 임베디드 워크시트를 한 번에 맞춘다. 정상 경로는 이 함수다.

    반환 {'val_cols','layout','workbook','n','series_names','warnings'}
    """
    warnings: list[str] = []
    sink = warnings.append if on_warn is None else (
        lambda m: (warnings.append(m), on_warn(m))[0])

    tree = deck.xml(chart_part)
    lay = chart_layout(tree)                     # 고치기 전 좌표계를 먼저 읽는다
    val_cols = sync_chart(tree, series, categories, strict=strict)
    deck.put(chart_part, tree)

    workbook = None
    if update_workbook:
        cats = list(categories) if categories is not None else chart_categories(tree)
        workbook = update_embedded_workbook(
            deck, chart_part, cats, series, layout=lay,
            series_headers=series_headers,
            preserve_formulas=preserve_formulas, on_warn=sink)
    return {
        "val_cols": val_cols,
        "layout": lay,
        "workbook": workbook,
        "n": len(series[0]) if series else 0,
        "series_names": series_names(tree),
        "warnings": warnings,
    }


__all__ = [
    # 네임스페이스·상수
    "A", "P", "C", "R", "PR", "NS", "EMU_PER_CM", "EMU_PER_PT",
    "HEADER_ROWS", "HEADER_H", "DATA_AREA_H", "TABLE_Y", "TABLE_H", "TABLE_BOTTOM",
    "CHART_Y", "CHART_H", "DEFAULT_CELL_MARGIN",
    "WORKBOOK_FIRST_DATA_ROW", "WORKBOOK_FIRST_DATA_COL", "q",
    # Deck
    "Deck",
    # 도형
    "iter_shapes", "shape_name", "find_shape", "find_placeholder", "placeholder_slots",
    "shape_text", "set_shape_text", "set_ph_text", "drop_shape", "drop_shapes_matching",
    "set_ment",
    # 표 구조
    "find_tables", "first_table", "table_rows", "data_rows", "row_cells",
    "grid_widths", "n_cols", "table_frame", "frame_box",
    # 표 기하
    "TableGeometry", "measure_table_geometry", "measure_chart_geometry",
    "distribute_row_heights", "set_row_count", "verify_row_heights", "label_metrics",
    "row_border_dashes", "derive_mid_row",
    # 셀
    "cell_text", "clear_text", "set_text", "put_cell",
    # 차트
    "chart_series", "series_names", "series_values", "chart_categories",
    "ChartLayout", "chart_layout", "series_orientation", "parse_ref",
    "col_to_index", "index_to_col",
    "sync_chart", "update_embedded_workbook", "apply_chart",
    "workbook_grid", "write_workbook_cells",
]
