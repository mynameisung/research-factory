#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""브랜드 치환 — 테마색 · 글꼴 · 로고 · 조사명 문자열을 클라이언트 프로파일로 갈아 끼운다.

왜 프로파일인가
---------------
예시 덱(LG 홈로봇)을 템플릿 삼아 에스원 보고서를 만들지만, **다음 프로젝트는 또 다른
클라이언트다.** 에스원 색을 코드에 박으면 그때 이 파일을 통째로 다시 쓰게 된다. 그래서
둘로 쪼갠다.

    BrandProfile   목표 브랜드       — 클라이언트마다 바뀐다 (색 1개 + 글꼴 + 이름)
    DeckMap        템플릿 덱의 지도  — LG 예시덱을 계속 쓰는 한 그대로다 (옛 자산 위치)

    apply_brand(deck, profile, deckmap)   치환 실행
    audit_brand(deck, terms=[...])        잔존 검사

색 파생 — 공식은 딱 하나다
--------------------------
에스원 CI 가 공식으로 규정한 것은 **PANTONE 286 C(#0033A0) · 검정 · (어두운 배경) 흰색**
셋뿐이다(https://www.s1.co.kr/company/introduction/ci). 보조색·강조색 규정은 공개돼 있지
않다. 그러므로 이 모듈이 만드는 나머지 색은 **전부 파생값이며 CI 규정이 아니다.**

파생 방식은 "원본 테마가 dk2 와 맺고 있던 관계를 새 주색 위로 그대로 옮긴다" 이다.
하드코딩한 목적지 hex 가 아니라 `derive_palette()` 가 원본 clrScheme 을 읽어 계산한다 —
그래야 다음 클라이언트 색을 넣어도 같은 논리로 팔레트가 따라온다.
각 색의 유도 근거는 PaletteEntry.rationale 에 문자열로 남아 리포트에 그대로 실린다.

★ 명도만은 상대변환하지 않는다 (이 판단을 되돌리지 마라)
   원본 lt2(#C74558)는 차트에서 lumMod/lumOff 로 3단 틴트(40/60·60/40·20/80)를 만들어
   쓰인다. lumMod 는 HSL 의 L 에 **곱하므로**, lt2 의 절대 명도가 바뀌면 틴트 사다리
   전체의 밝기가 같이 무너진다. 그래서 색상각·채도는 상대변환하되 **명도는 원본 lt2 의
   절대값(52.55%)을 그대로 쓴다.** 상대변환(dk2 대비 ×1.145)을 택하면 #233995(L 35.9%)가
   되고 lumMod40/lumOff60 결과가 L 74.4% 로 LG 의 81.0% 보다 눈에 띄게 진해진다.
   "상대변환이 더 논리적" 이라며 되돌리면 차트 2계열 막대가 전부 어두워진다.

기하 불변 (R-105 / R-106)
-------------------------
이 모듈은 표·차트 프레임 기하를 **단 하나도 건드리지 않는다.** 색·글꼴·문자열·미디어만
바꾼다. 수치를 만지는 곳은 두 군데뿐이고 둘 다 표/차트와 무관하다.

    (a) 표 라벨 셀의 자간 spc  -50 → -100   (pt 는 그대로라 행높이 체계에 영향 없음)
    (b) spAutoFit + wrap="none" 인 footer/PART 텍스트 상자의 ext.cx

`assert_geometry_unchanged()` 가 치환 전후 표/차트 프레임 좌표와 tblGrid·행높이를
전수 대조해 이를 증명한다. rebrand_file() 은 이 대조가 깨지면 저장하지 않고 죽는다.

검증
----
    python rebrand.py --selftest
        색 파생 산술 자기검증 (원본 덱 없이 즉시)
    python rebrand.py --src <template.pptx> --audit-only
    python rebrand.py --src <template.pptx> --out <out.pptx> --measure --render 18
        치환 → 기하 대조 → 저장 → 잔존 검사 → 라벨 폭 실측 → PowerPoint COM 렌더
"""

from __future__ import annotations

import argparse
import colorsys
import posixpath
import re
import subprocess
import sys
import tempfile
import warnings
from dataclasses import dataclass, field, replace as dc_replace
from pathlib import Path
from typing import Any, Sequence

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckops as D                                  # noqa: E402
import textfit as TF                                 # noqa: E402
import kobreak as KB                                 # noqa: E402

A = D.A                          # drawingml
P = D.P                          # presentationml
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

EMU_PER_PT = 12700
SLIDE_MID_X = 6096000            # 16:9 12192000 EMU 의 가로 중앙
IMAGE_REL = R_NS + "/image"


# =========================================================================== #
# 1. 색 대수 — HSL 왕복과 PowerPoint lumMod/lumOff
# =========================================================================== #
#
# PowerPoint 의 lumMod/lumOff 는 HSL 의 L 에 대해 L' = L*lumMod + lumOff 로 동작한다.
# 그래서 팔레트 파생도 RGB 가 아니라 HSL 위에서 해야 틴트 결과를 예측할 수 있다.

def hex_to_hsl(value: str) -> tuple[float, float, float]:
    """'#RRGGBB' 또는 'RRGGBB' -> (H 0~360, S 0~100, L 0~100)."""
    h = value.lstrip("#").strip()
    if len(h) != 6:
        raise ValueError(f"6자리 hex 가 아니다: {value!r}")
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    hh, ll, ss = colorsys.rgb_to_hls(r, g, b)
    return hh * 360.0, ss * 100.0, ll * 100.0


def hsl_to_hex(h: float, s: float, l: float) -> str:
    """(H, S, L) -> 'RRGGBB' (대문자, # 없음 — OOXML srgbClr/@val 형식)."""
    s = min(100.0, max(0.0, s))
    l = min(100.0, max(0.0, l))
    r, g, b = colorsys.hls_to_rgb((h % 360.0) / 360.0, l / 100.0, s / 100.0)
    return "%02X%02X%02X" % tuple(round(v * 255) for v in (r, g, b))


def lum_tint(value: str, lum_mod: float, lum_off: float) -> str:
    """PowerPoint 가 <a:lumMod/><a:lumOff/> 로 만들어 내는 색.

    lum_mod / lum_off 는 0~1 비율로 준다 (XML 의 40000 은 0.40).
    검증용이다 — 이 모듈은 틴트를 XML 에 쓰지 않는다. 테마 색만 바꾸면
    PowerPoint 가 알아서 계산하기 때문이다.
    """
    h, s, l = hex_to_hsl(value)
    return hsl_to_hex(h, s, l * lum_mod + lum_off * 100.0)


def rehue(value: str, hue: float) -> str:
    """채도·명도는 그대로 두고 색상각만 목표값으로 옮긴다.

    '따뜻한 아이보리 요약행' 처럼 눈대중으로 고른 하드코딩 색을 브랜드 색상환으로
    끌고 올 때 쓴다. 밝기와 탁도가 보존되므로 원래 디자인의 대비 관계가 유지된다.
    """
    _, s, l = hex_to_hsl(value)
    return hsl_to_hex(hue, s, l)


def _bare(value: str) -> str:
    return value.lstrip("#").upper()


# =========================================================================== #
# 2. 팔레트 파생
# =========================================================================== #

@dataclass(frozen=True)
class PaletteEntry:
    """테마 clrScheme 한 칸. rationale 이 이 색의 존재 근거다."""
    slot: str                 # 'dk2' | 'lt2' | 'accent1' ...
    value: str                # 'RRGGBB'
    official: bool            # True = 클라이언트 CI 가 규정한 색
    rationale: str            # 어떻게 나온 값인지 (리포트에 그대로 실린다)

    def comment(self) -> str:
        tag = "CI 공식" if self.official else "CI 규정 아님 · 파생"
        return f"{self.slot} #{self.value}  [{tag}] {self.rationale}"


# 강조색 사다리 — 원본 accent3~6 (보라·올리브·앰버·붉은색) 은 브랜드와 무관한 값이라
# '옮길 관계' 가 없다. 그래서 새 lt2 의 색상각 위에 놓인 채도/명도 사다리로 갈아엎는다.
# (S%, L%) 는 계산이 아니라 **고른 값**이다 — 근거는 각 줄 주석에.
#
#   accent3  실사용 12회(chart26 계열 '주거 운영/관리'). 다른 차트의 2계열 막대와
#            같은 톤으로 읽혀야 하므로 새 lt2 의 lumMod60/lumOff40 틴트(L 71.5%)와
#            같은 밝기대(L 61.6%)에 두고 채도는 lt2 수준(62%)으로 맞췄다.
#   accent4~6 실사용 0회. 비워 두면 안 되는 이유는 '누가 실수로 집었을 때 옛 클라이언트
#            톤이 새어 나오지 않게' 다. 그래서 브랜드 색상환 안 중성 사다리로 채운다.
ACCENT_LADDER: tuple[tuple[str, float, float, str], ...] = (
    ("accent3", 62.0, 61.6,
     "채도 lt2 수준 62% · 명도 61.6% (lt2 의 lumMod60/lumOff40 틴트와 같은 밝기대) ★실사용 12회"),
    ("accent4", 30.0, 55.0, "채도 ×0.30 중성 · 명도 55% (미사용 슬롯 안전값)"),
    ("accent5", 18.0, 72.0, "채도 ×0.18 중성 · 명도 72% (미사용 슬롯 안전값)"),
    ("accent6", 45.0, 40.0, "채도 ×0.45 · 명도 40% 어두운 쪽 (미사용 슬롯 안전값)"),
)

# 손대지 않는 칸. 검정은 대부분의 CI 가 공식으로 두고(에스원도 그렇다), hlink/folHlink 는
# 브랜드가 아니라 시스템 값이다. 여기에 브랜드 색을 밀어 넣으면 링크 색까지 물든다.
PALETTE_KEEP = ("dk1", "lt1", "hlink", "folHlink")


def derive_palette(primary: str, source_scheme: dict[str, str],
                   *, primary_source: str = "") -> dict[str, PaletteEntry]:
    """원본 clrScheme 이 dk2 와 맺고 있던 관계를 새 주색 위로 옮긴다.

    source_scheme 은 원본 테마에서 읽은 {'dk2': 'DA1047', 'lt2': 'C74558', ...} 이다.
    하드코딩한 목적지 hex 가 없다 — 전부 여기서 계산된다. 그래야 다음 클라이언트
    주색을 넣어도 같은 논리가 그대로 돈다.

    반환은 **바꿀 칸만** 담는다. PALETTE_KEEP 은 아예 나오지 않는다.
    """
    ph, ps, _pl = hex_to_hsl(primary)
    out: dict[str, PaletteEntry] = {}

    src_dk2 = source_scheme.get("dk2")
    src_lt2 = source_scheme.get("lt2")

    # --- dk2 = 주색 그 자체 -------------------------------------------------
    # 마스터의 clrMap 이 tx2=dk2 이므로 차트 주막대·PART 라벨·소제목 밴드가 전부
    # 이 한 칸을 본다. 즉 여기만 바꾸면 차트 파트 136개를 열 필요가 없다.
    out["dk2"] = PaletteEntry(
        "dk2", _bare(primary), official=True,
        rationale=f"클라이언트 CI 주색 그대로. {primary_source}".strip())

    # --- lt2 = 원본 lt2↔dk2 관계를 새 주색 위로 ----------------------------
    if src_dk2 and src_lt2:
        dh, ds, _dl = hex_to_hsl(src_dk2)
        lh, ls, ll = hex_to_hsl(src_lt2)
        d_hue = lh - dh                                   # 색상각 차 (LG 원본 +7.57°)
        r_sat = (ls / ds) if ds else 1.0                  # 채도 비   (LG 원본 ×0.6223)
        new_lt2 = hsl_to_hex(ph + d_hue, ps * r_sat, ll)  # ★ 명도는 원본 절대값 ll
        out["lt2"] = PaletteEntry(
            "lt2", new_lt2, official=False,
            rationale=(f"#{_bare(primary)} 파생: 색상각 {ph:.1f}°{d_hue:+.2f}°, "
                       f"채도 {ps:.0f}%×{r_sat:.4f}, "
                       f"명도는 원본 lt2 절대값 {ll:.2f}% 고정 "
                       f"— lumMod/lumOff 틴트 사다리 보존 (상대변환 금지)"))
    else:
        out["lt2"] = PaletteEntry("lt2", _bare(primary), official=False,
                                  rationale="원본 lt2 없음 — 주색으로 폴백")

    lt2_h, _lt2_s, _lt2_l = hex_to_hsl(out["lt2"].value)

    # --- accent1/2 = 주색·보조색 재사용 -------------------------------------
    # 실사용 0회지만 옛 값을 남기면 실수로 집힐 때 남의 브랜드 색이 나온다.
    out["accent1"] = PaletteEntry(
        "accent1", out["dk2"].value, official=False,
        rationale="dk2(주색) 재사용 — 실사용 0회 슬롯을 브랜드 안에 가둬 두는 안전값")
    out["accent2"] = PaletteEntry(
        "accent2", out["lt2"].value, official=False,
        rationale="lt2(보조색) 재사용 — 실사용 0회 슬롯 안전값")

    # --- accent3~6 = 중성 사다리 -------------------------------------------
    for slot, sat, lum, why in ACCENT_LADDER:
        out[slot] = PaletteEntry(
            slot, hsl_to_hex(lt2_h, sat, lum), official=False,
            rationale=f"#{_bare(primary)} 파생: 새 lt2 색상각 {lt2_h:.1f}° · {why}")
    return out


# =========================================================================== #
# 3. BrandProfile — 목표 브랜드 (클라이언트마다 바뀌는 것)
# =========================================================================== #

@dataclass(frozen=True)
class BrandProfile:
    """치환 목표. 여기 없는 건 전부 DeckMap(템플릿 지도) 쪽 일이다."""

    name: str                                  # '에스원'
    primary: str                               # '#0033A0' — CI 가 규정한 유일한 색
    primary_source: str = ""                   # 근거 문장 (리포트에 그대로 실린다)

    # ---- 글꼴 ----
    # 라틴/한글 한 이름으로 간다. Bold 는 별도 페이스가 아니라 런의 b="1" 로 처리되므로
    # major/minor 를 같은 이름으로 둔다.
    font_major: str = "맑은 고딕"
    font_minor: str = "맑은 고딕"
    font_scheme_name: str | None = None        # None -> font_minor
    color_scheme_name: str | None = None       # None -> 손대지 않음

    # ---- 문자열 ----
    survey_name: str = ""                      # footer 조사명 (짧은 쪽)
    string_map: tuple[tuple[str, str], ...] = ()   # 그 밖의 (원본→치환) 쌍

    # ---- 하드코딩 색 ----
    # theme 을 갈아도 안 바뀌는 srgbClr 직박음. 채도·명도는 두고 색상각만 브랜드로 옮긴다.
    rehue_srgb: tuple[str, ...] = ()

    # ---- 표 라벨 자간 ----
    # 맑은 고딕은 LG EI Text 보다 한글이 1.235배 넓다. pt 를 줄이면 R-105 행높이 체계와
    # 헤더 균형이 흔들리므로 **자간으로만** 흡수한다. None 이면 손대지 않는다.
    label_spc_from: int | None = -50
    label_spc_to: int | None = -100
    label_cols: int = 1                        # 라벨로 취급할 선두 열 수

    # ---- 로고 ----
    logo_light: Path | None = None             # 밝은 배경용 (원본 색)
    logo_dark: Path | None = None              # 어두운 배경용. None + auto_whiten 이면 자동 생성
    auto_whiten_dark: bool = True              # CI 「어두운 배경은 흰색으로 변환」 규정 이행
    logo_shape_name: str | None = None         # None -> f'고객사 로고({name})'

    # ---- docProps ----
    doc_title: str | None = None
    doc_last_modified_by: str | None = "Embrain"

    def palette(self, source_scheme: dict[str, str]) -> dict[str, PaletteEntry]:
        return derive_palette(self.primary, source_scheme,
                              primary_source=self.primary_source)

    def logo_name(self) -> str:
        return self.logo_shape_name or f"고객사 로고({self.name})"


# --------------------------------------------------------------------------- #
# 기성 프로파일 — 이 파일이 '에스원 전용' 이 되지 않도록 값만 모아 둔다.
# 다음 클라이언트는 여기에 한 덩이를 더 쓸 뿐 위쪽 코드는 손대지 않는다.
# --------------------------------------------------------------------------- #

S1 = BrandProfile(
    name="에스원",
    # ★ #0033A0 = PANTONE 286 C. 이 덱에서 유일하게 '규정된' 색이고 나머지는 전부 파생이다.
    primary="#0033A0",
    primary_source=("PANTONE 286 C = RGB(0,51,160). 출처 "
                    "https://www.s1.co.kr/company/introduction/ci "
                    "— 공식 규정은 286C·검정·(어두운 배경)흰색 3가지뿐이다"),
    # 맑은 고딕: Windows 기본 탑재라 의뢰처 PC 에서 대체 없이 열린다.
    # 에스원 CI 는 로고타입 규정만 있고 본문 서체 규정이 공개돼 있지 않다.
    font_major="맑은 고딕",
    font_minor="맑은 고딕",
    color_scheme_name="에스원 286C",
    survey_name="비상주 건물관리 조사",
    # 레이아웃13 의 자리표시자 문구. 조사명과 달리 브랜드가 아니라 템플릿 잔재다.
    string_map=(("목차를 입력하세요", "목차"),),
    # F6EEEA = 누적막대 장표 'TOP3(%)' 요약행 배경(38개 슬라이드 454회). tcPr 에 직접
    # 박혀 있어 theme 을 갈아도 안 바뀐다. 색상각 20.0°의 따뜻한 아이보리로,
    # dk2(343.7°)에서 계산으로 유도되지 않는 눈대중 값이었다.
    rehue_srgb=("F6EEEA",),
    label_spc_from=-50,
    label_spc_to=-100,
    doc_title="비상주 건물관리 조사",
)


# =========================================================================== #
# 4. DeckMap — 템플릿 덱의 지도 (LG 예시덱을 쓰는 한 그대로)
# =========================================================================== #

@dataclass
class LogoSlot:
    """옛 클라이언트 로고가 놓여 있던 자리 = 새 로고가 들어갈 자리. 기하는 EMU."""
    part: str
    shape: str
    media: str
    x: int
    y: int
    cx: int
    cy: int
    dark: bool                     # 어두운 배경 위인가 (흰색 변환본이 필요한 자리)
    align: str                     # 'left' | 'right' — cx 를 바꿀 때 어느 변을 고정하나

    @property
    def right(self) -> int:
        return self.x + self.cx

    @property
    def aspect(self) -> float:
        return self.cx / self.cy if self.cy else 0.0

    def as_dict(self) -> dict[str, Any]:
        return {"part": self.part, "shape": self.shape, "media": self.media,
                "off": (self.x, self.y), "ext": (self.cx, self.cy),
                "right": self.right, "aspect": round(self.aspect, 4),
                "dark": self.dark, "align": self.align}


@dataclass
class DeckMap:
    """치환 대상 덱에서 '어디를 고쳐야 하는지'. 값은 실측으로 채운다.

    BrandProfile 과 분리한 이유: 클라이언트가 바뀌어도 **템플릿이 같으면 이 지도는
    그대로** 다. 반대로 제작팀이 다른 템플릿을 주면 프로파일은 그대로 두고 이쪽만 바꾼다.
    """
    # 옛 클라이언트(=원 소유주) 로고로 판정된 media 파일 basename.
    # '어느 이미지가 남의 로고인가' 는 사람이 눈으로 볼 일이라 값으로 받는다.
    # 그 파일이 어느 슬롯에 몇 번 쓰였는지는 코드가 찾는다.
    client_logo_media: tuple[str, ...] = ()
    # 어두운 배경 위에 놓인 로고 media (흰색 변환본이 필요한 자리)
    dark_logo_media: tuple[str, ...] = ()
    # ★ 절대 건드리면 안 되는 도형 이름 — 우리(엠브레인) 로고
    keep_shapes: tuple[str, ...] = ("Embrain",)
    # footer 조사명이 들어 있는 도형 이름
    footer_shapes: tuple[str, ...] = ("TextBox 1",)
    # 원본 조사명. None 이면 detect() 가 레이아웃에서 최빈 문자열로 찾아낸다
    old_survey_name: str | None = None
    # 옛 소유주를 가리키는 낱말. 대체 텍스트(cNvPr/@descr·@title)에서 이게 걸리면 지운다.
    # 화면에는 안 보이지만 파일 안에 그대로 실려 나가고 접근성 창에는 뜬다.
    retire_terms: tuple[str, ...] = ()

    # ---- 아래는 detect() 가 실측으로 채운다 ----
    theme_parts: list[str] = field(default_factory=list)
    source_scheme: dict[str, str] = field(default_factory=dict)
    source_fonts: dict[str, str] = field(default_factory=dict)
    logo_slots: list[LogoSlot] = field(default_factory=list)

    # ---------------------------------------------------------------- detect
    @classmethod
    def detect(cls, deck: D.Deck, *, client_logo_media: Sequence[str] = (),
               dark_logo_media: Sequence[str] = (),
               keep_shapes: Sequence[str] = ("Embrain",),
               old_survey_name: str | None = None,
               retire_terms: Sequence[str] = ()) -> "DeckMap":
        """덱을 실제로 훑어 지도를 만든다.

        하드코딩한 좌표를 믿지 않는다 — 템플릿이 조금만 바뀌어도 조용히 틀리기 때문이다.
        사람 판단이 필요한 것(어느 media 가 남의 로고인가)만 인자로 받고 나머지는 찾는다.
        """
        m = cls(client_logo_media=tuple(client_logo_media),
                dark_logo_media=tuple(dark_logo_media),
                keep_shapes=tuple(keep_shapes),
                old_survey_name=old_survey_name,
                retire_terms=tuple(retire_terms))

        m.theme_parts = sorted(deck.names(r"^ppt/theme/theme(Override)?\d*\.xml$"))
        if m.theme_parts:
            tree = deck.xml(m.theme_parts[0])
            cs = tree.find(f".//{D.q(A, 'clrScheme')}")
            if cs is not None:
                for child in cs:
                    slot = etree.QName(child).localname
                    srgb = child.find(D.q(A, "srgbClr"))
                    if srgb is not None:
                        m.source_scheme[slot] = (srgb.get("val") or "").upper()
            fs = tree.find(f".//{D.q(A, 'fontScheme')}")
            if fs is not None:
                for grp in ("majorFont", "minorFont"):
                    g = fs.find(D.q(A, grp))
                    if g is None:
                        continue
                    for script in ("latin", "ea", "cs"):
                        el = g.find(D.q(A, script))
                        if el is not None:
                            m.source_fonts[f"{grp}.{script}"] = el.get("typeface") or ""

        if m.old_survey_name is None:
            m.old_survey_name = _detect_footer_text(deck, m.footer_shapes)

        m.logo_slots = _find_logo_slots(deck, m.client_logo_media,
                                        m.dark_logo_media, m.keep_shapes)
        return m

    def summary(self) -> dict[str, Any]:
        return {
            "theme_parts": self.theme_parts,
            "source_scheme": dict(self.source_scheme),
            "source_fonts": dict(self.source_fonts),
            "old_survey_name": self.old_survey_name,
            "logo_slots": [s.as_dict() for s in self.logo_slots],
        }


# LG 예시 덱(예시_제작본_LG홈로봇_최종.pptx) 실측 지도. 픽셀로 직접 확인한 값이다.
#   image7.png  205x72  #A50034+#6B6B6B  밝은 배경용   ← slideLayout 4/6/13 좌하단
#   image12.png 248x87  #A50034+#6B6B6B  밝은 배경용   ← slide1 표지
#   image9.png  256x90  #A50033+흰색     어두운 배경용 ← slideLayout5 우상단 배너
#   image41.png 309x108 #A50033+흰색     어두운 배경용 ← slide37/89 전면 밴드
# 네 파일 모두 가로세로비 2.85.
# ★ pic[Embrain](image2.emf, 레이아웃 11개)은 우리 회사 로고다. keep_shapes 로 보호한다.
#
# retire_terms 는 대체 텍스트 청소용이다. LG 덱 slide16/37/89 의 pic 'Picture 4' 에
#   descr="가사 스스로 판단·수행하는 AI 홈로봇 등장…LG전자, 홈로봇 'LG 클로이드' 공개 …"
# 라는 보도사진 캡션이 그대로 박혀 있다. 화면에는 안 보이지만 파일에 실려 나가고
# 접근성 창·대체텍스트 창에는 뜬다. 실측으로 찾아낸 잔재라 반드시 지운다.
LG_TEMPLATE_ARGS: dict[str, Any] = dict(
    client_logo_media=("image7.png", "image9.png", "image12.png", "image41.png"),
    dark_logo_media=("image9.png", "image41.png"),
    keep_shapes=("Embrain",),
    retire_terms=("LG", "엘지", "홈로봇"),
)


def _detect_footer_text(deck: D.Deck, shape_names: Sequence[str]) -> str | None:
    """레이아웃에서 footer 도형에 가장 여러 번 나오는 문자열 = 조사명.

    문자열을 코드에 박지 않으려는 것이다. 2곳 이상에서 같아야 조사명으로 인정한다.
    """
    counts: dict[str, int] = {}
    for part in deck.names(r"^ppt/slideLayouts/slideLayout\d+\.xml$"):
        tree = deck.xml(part)
        for sp in D.iter_shapes(tree):
            if D.shape_name(sp) not in shape_names:
                continue
            txt = D.shape_text(sp).strip()
            if txt:
                counts[txt] = counts.get(txt, 0) + 1
    if not counts:
        return None
    best, n = max(counts.items(), key=lambda kv: kv[1])
    return best if n >= 2 else None


def _xfrm_box(el: etree._Element) -> tuple[etree._Element | None,
                                           etree._Element | None,
                                           etree._Element | None]:
    """도형의 (xfrm, off, ext).

    ★ `.//a:ext` 를 그냥 쓰면 blipFill 의 <a:extLst><a:ext uri=..> 를 먼저 집어
      cx/cy 가 None 으로 나온다. 반드시 xfrm 아래에서만 찾는다.
    """
    xf = el.find(f".//{D.q(A, 'xfrm')}")
    if xf is None:
        return None, None, None
    return xf, xf.find(D.q(A, "off")), xf.find(D.q(A, "ext"))


def _pic_name(pic: etree._Element) -> str:
    nv = pic.find(f"{D.q(P, 'nvPicPr')}/{D.q(P, 'cNvPr')}")
    return (nv.get("name") or "") if nv is not None else ""


def _pic_media(pic: etree._Element, relmap: dict[str, str]) -> str | None:
    blip = pic.find(f"{D.q(P, 'blipFill')}/{D.q(A, 'blip')}")
    if blip is None:
        blip = pic.find(f".//{D.q(A, 'blip')}")
    if blip is None:
        return None
    return relmap.get(blip.get(f"{{{R_NS}}}embed") or "")


def _find_logo_slots(deck: D.Deck, media: Sequence[str], dark: Sequence[str],
                     keep: Sequence[str]) -> list[LogoSlot]:
    """지정한 media 를 참조하는 p:pic 을 전수로 찾는다 (그룹 안쪽 포함)."""
    want = {m.lower() for m in media}
    dark_set = {m.lower() for m in dark}
    slots: list[LogoSlot] = []
    pattern = (r"^ppt/(slides/slide|slideLayouts/slideLayout"
               r"|slideMasters/slideMaster)\d+\.xml$")
    for part in sorted(deck.names(pattern)):
        relmap = {r["Id"]: r["Part"] for r in deck.rels(part)}
        tree = deck.xml(part)
        for pic in tree.iter(D.q(P, "pic")):
            name = _pic_name(pic)
            if name in keep:                       # 우리 회사 로고 — 절대 손대지 않는다
                continue
            target = _pic_media(pic, relmap)
            if not target:
                continue
            base = posixpath.basename(target).lower()
            if base not in want:
                continue
            _xf, off, ext = _xfrm_box(pic)
            if off is None or ext is None:
                continue
            x, y = int(off.get("x", 0)), int(off.get("y", 0))
            cx, cy = int(ext.get("cx", 0)), int(ext.get("cy", 0))
            # 슬라이드 가로 중앙보다 오른쪽에 있으면 우변에 붙여 놓은 자리로 본다.
            # 로고 비율이 바뀌어 cx 가 달라질 때 어느 변을 고정할지가 여기서 갈린다.
            align = "right" if (x + cx / 2) > SLIDE_MID_X else "left"
            slots.append(LogoSlot(part, name, posixpath.basename(target),
                                  x, y, cx, cy, base in dark_set, align))
    return slots


# =========================================================================== #
# 5. 폭 실측 — 자간(spc) 을 반영한다
# =========================================================================== #
#
# ★ textfit.width_emu() 는 spc 를 모델링하지 않는다. 이 덱의 표 셀은 **전부**
#   spc="-50"(=-0.5pt/자) 이므로 textfit 만 쓰면 실폭을 5~8%p 과대평가한다.
#   LG EI Text 에서는 안전마진(1줄 최대 90.8% / 2줄 최소 111.2% = 20.4%p)에 가려
#   판정이 안 뒤집혔지만, 맑은 고딕에서는 마진이 좁아져 실제로 뒤집힌다.
#   (예: '소음/움직임으로 인한 생활 방해' — spc 무시 104.0% → 2줄 오판,
#         spc 반영 98.3% → 1줄 정답)
#   textfit.py 는 이번 과업 범위 밖이라 여기서 보정한다.
#   나중에 textfit 에 spc 인자가 생기면 이 두 함수는 지우고 그쪽을 불러라.

def width_emu_spc(text: str, font: str, sz_pt: float = 10.0, spc_hpt: int = 0) -> float:
    """자간을 반영한 렌더 폭(EMU).

    spc_hpt 는 OOXML rPr/@spc 그대로(100분의 1 pt). -50 = -0.5pt/자.
    PowerPoint 는 글자 **사이**가 아니라 글자마다 advance 에 spc 를 더하므로
    len(text) 를 곱한다 (마지막 글자 뒤 여백도 상자 폭에 계산된다).
    """
    return TF.width_emu(text, font, sz_pt) + spc_hpt / 100.0 * len(text) * EMU_PER_PT


def wrap_ko_spc(text: str, usable_emu: float, font: str, sz_pt: float = 10.0,
                max_lines: int = 4, spc_hpt: int = 0) -> list[str]:
    """textfit.wrap_ko 와 같은 알고리즘에 자간만 먹인다.

    textfit 의 DP(_dp_ko)와 글자분할(_char_split)은 폭 함수 W 를 인자로 받으므로
    재구현하지 않고 그대로 재사용한다 — 줄바꿈 벌점(kobreak, R-108)도 같이 따라온다.
    """
    text = " ".join(text.split())

    def W(s: str) -> float:
        return width_emu_spc(s, font, sz_pt, spc_hpt)

    if not text:
        return []
    if W(text) <= usable_emu:
        return [text]
    words = text.split(" ")
    if len(words) == 1:
        return TF._char_split(text, usable_emu, W)
    pen = KB.penalties(words)
    for n_lines in range(2, min(max_lines, len(words)) + 1):
        best = TF._dp_ko(words, pen, n_lines, usable_emu, W)
        if best:
            return best
    return TF._char_split(text, usable_emu, W)


def wrap_ko_hard(text: str, usable_emu: float, font: str, sz_pt: float = 10.0,
                 max_lines: int = 4, spc_hpt: int = 0) -> list[str]:
    """이미 <a:br/> 로 줄을 박아 둔 라벨을 그 줄을 지켜 가며 잰다.

    ★ 왜 필요한가 (실측으로 물린 자리)
      LG 덱의 긴 라벨은 제작팀이 이미 손으로 줄을 끊어 뒀다. 예를 들어 s23 의
      '집 안에 항상 존재하는 것에 대한\\n정서적 부담' 은 셀 안에 개행이 들어 있다.
      이걸 한 덩이 문자열로 재면 폭이 실제의 두 배 가까이 나오고, 라벨 이름이
      개행을 품은 채로 잡혀 다른 라벨과 매칭도 안 된다.
      하드 브레이크는 '저자가 이미 고른 자리' 이므로 존중하고, **각 조각이
      여전히 한 줄에 들어가는가** 만 따진다. 조각이 넘치면 그 조각만 다시 접는다.
    """
    parts = [p for p in re.split(r"[\r\n]+", text)]
    out: list[str] = []
    for piece in parts:
        piece = " ".join(piece.split())
        if not piece:
            continue
        out.extend(wrap_ko_spc(piece, usable_emu, font, sz_pt, max_lines, spc_hpt))
    return out


def label_width_hard(text: str, font: str, sz_pt: float, spc_hpt: int) -> float:
    """하드 브레이크를 가진 라벨의 '가장 넓은 줄' 폭. 넘침 판정의 기준이다."""
    pieces = [" ".join(p.split()) for p in re.split(r"[\r\n]+", text)]
    pieces = [p for p in pieces if p]
    if not pieces:
        return 0.0
    return max(width_emu_spc(p, font, sz_pt, spc_hpt) for p in pieces)


# =========================================================================== #
# 6. 치환 리포트
# =========================================================================== #

@dataclass
class Change:
    kind: str
    part: str
    detail: str
    count: int = 1


@dataclass
class BrandReport:
    profile: str
    palette: dict[str, PaletteEntry] = field(default_factory=dict)
    changes: list[Change] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    logo_slots: list[LogoSlot] = field(default_factory=list)

    def add(self, kind: str, part: str, detail: str, count: int = 1) -> None:
        if count:
            self.changes.append(Change(kind, part, detail, count))

    def total(self, kind: str | None = None) -> int:
        return sum(c.count for c in self.changes if kind is None or c.kind == kind)

    def kinds(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for c in self.changes:
            out[c.kind] = out.get(c.kind, 0) + c.count
        return out

    def of(self, kind: str) -> list[Change]:
        return [c for c in self.changes if c.kind == kind]

    def text(self, *, detail_kinds: Sequence[str] = ("geom.autofit", "logo.blank",
                                                     "logo.swap", "media.drop")) -> str:
        lines = [f"[브랜드 치환] {self.profile}", ""]
        lines.append("팔레트  ★ = 클라이언트 CI 공식 · 나머지는 전부 파생값이다")
        for e in self.palette.values():
            mark = "★" if e.official else " "
            lines.append(f"  {mark} {e.slot:8s} #{e.value}")
            lines.append(f"       {e.rationale}")
        lines.append("")
        lines.append("변경 요약")
        for kind, n in sorted(self.kinds().items(), key=lambda kv: -kv[1]):
            lines.append(f"  {n:6d}  {kind}")
        lines.append(f"  {'-' * 6}")
        lines.append(f"  {self.total():6d}  합계")
        for kind in detail_kinds:
            items = self.of(kind)
            if not items:
                continue
            lines.append("")
            lines.append(f"{kind}")
            for c in items:
                lines.append(f"  {c.part:40s} {c.detail}")
        if self.notes:
            lines.append("")
            lines.append("메모")
            for n in self.notes:
                lines.append(f"  · {n}")
        return "\n".join(lines)


# =========================================================================== #
# 7. apply_brand
# =========================================================================== #

_SLIDE_LIKE = (r"^ppt/(slides/slide|slideLayouts/slideLayout"
               r"|slideMasters/slideMaster)\d+\.xml$")


def apply_brand(deck: D.Deck, profile: BrandProfile,
                deckmap: DeckMap | None = None) -> BrandReport:
    """덱 전체에 브랜드를 입힌다. 표·차트 기하는 단 하나도 건드리지 않는다.

    순서에 뜻이 있다.
      ① 테마 색     이 한 파트로 차트 136개가 전부 따라온다 (색이 전부 schemeClr 참조라서)
      ② 테마 글꼴   레이아웃·마스터·차트는 이미 +mn-*/+mj-* 상속이라 손댈 게 없다
      ③ 직박음 글꼴 테마가 못 미치는 슬라이드 런의 typeface 를 테마 참조로 되돌린다
      ④ 직박음 색   theme 을 갈아도 안 바뀌는 srgbClr
      ⑤ 문자열      footer 조사명 · 자리표시자 문구 (+ spAutoFit 상자 폭 갱신)
      ⑥ 대체텍스트  cNvPr/@descr 에 박힌 옛 소유주 캡션 (화면엔 안 보이지만 파일에 남는다)
      ⑦ 로고        남의 로고를 걷어내고 자리를 남기거나 새 로고를 끼운다
      ⑧ 표 라벨 자간 글꼴이 넓어진 만큼 자간으로 흡수 (pt 는 그대로 → R-105 무영향)
      ⑨ docProps    파일 속성에 남는 흔적
    """
    if deckmap is None:
        deckmap = DeckMap.detect(deck, **LG_TEMPLATE_ARGS)

    rep = BrandReport(profile=profile.name)
    rep.palette = profile.palette(deckmap.source_scheme)
    rep.logo_slots = list(deckmap.logo_slots)

    _apply_theme_colors(deck, profile, deckmap, rep)
    _apply_theme_fonts(deck, profile, deckmap, rep)
    _strip_hardcoded_fonts(deck, profile, deckmap, rep)
    _recolor_hardcoded(deck, profile, deckmap, rep)
    _replace_strings(deck, profile, deckmap, rep)
    _scrub_alt_text(deck, deckmap, rep)
    _retire_logos(deck, profile, deckmap, rep)
    _tighten_label_spacing(deck, profile, rep)
    _clean_docprops(deck, profile, deckmap, rep)
    return rep


# ---------------------------------------------------------------- ① 테마 색
def _apply_theme_colors(deck: D.Deck, profile: BrandProfile,
                        deckmap: DeckMap, rep: BrandReport) -> None:
    """clrScheme 을 파생 팔레트로 갈아 끼운다.

    ★ 이 한 파트가 차트 색의 유일한 출처다. 차트 136개 색 전수조사 결과 하드코딩
      srgbClr 은 000000 2회뿐이고 나머지 2,243개가 전부 schemeClr(bg1/tx1/tx2/bg2/
      accent3) 참조였다. 그래서 ppt/charts/*.xml 은 한 바이트도 건드리지 않는다.
    """
    for part in deckmap.theme_parts:
        tree = deck.xml(part)
        cs = tree.find(f".//{D.q(A, 'clrScheme')}")
        if cs is None:
            continue
        n = 0
        for child in cs:
            slot = etree.QName(child).localname
            if slot in PALETTE_KEEP:                  # dk1/lt1/hlink/folHlink 는 유지
                continue
            entry = rep.palette.get(slot)
            if entry is None:
                continue
            srgb = child.find(D.q(A, "srgbClr"))
            if srgb is None:                          # sysClr(lt1) 등은 손대지 않는다
                continue
            if (srgb.get("val") or "").upper() != entry.value:
                srgb.set("val", entry.value)
                n += 1
        if profile.color_scheme_name:
            cs.set("name", profile.color_scheme_name)
        deck.put(part, tree)
        rep.add("theme.color", part, f"clrScheme {n}칸", n)


# -------------------------------------------------------------- ② 테마 글꼴
def _apply_theme_fonts(deck: D.Deck, profile: BrandProfile,
                       deckmap: DeckMap, rep: BrandReport) -> None:
    """fontScheme 의 major/minor latin·ea 를 새 글꼴로. cs 는 빈 문자열 그대로 둔다."""
    for part in deckmap.theme_parts:
        tree = deck.xml(part)
        fs = tree.find(f".//{D.q(A, 'fontScheme')}")
        if fs is None:
            continue
        n = 0
        for grp, want in (("majorFont", profile.font_major),
                          ("minorFont", profile.font_minor)):
            g = fs.find(D.q(A, grp))
            if g is None:
                continue
            for script in ("latin", "ea"):
                el = g.find(D.q(A, script))
                if el is not None and (el.get("typeface") or "") != want:
                    el.set("typeface", want)
                    n += 1
        fs.set("name", profile.font_scheme_name or profile.font_minor)
        deck.put(part, tree)
        rep.add("theme.font", part, f"fontScheme {n}칸 → {profile.font_minor}", n)


# ---------------------------------------------------- ③ 직박음 글꼴 → 테마 참조
def _strip_hardcoded_fonts(deck: D.Deck, profile: BrandProfile,
                           deckmap: DeckMap, rep: BrandReport) -> None:
    """런에 직접 박힌 옛 브랜드 전용 글꼴을 테마 참조(+mn-ea/+mn-lt)로 되돌린다.

    리터럴 '맑은 고딕' 을 박지 않는 이유: 다음 프로젝트에서 **테마만 갈면 따라오게**
    하려는 것이다. 원본 테마 글꼴 이름(source_fonts)과 그 브랜드 접두어를 가진 이름만
    건드리고, Arial·Wingdings 같은 의도적 지정은 그대로 둔다.
    (LG 덱 실측: 레이아웃 14개·마스터·차트 136개에는 직박음이 0건이고,
     'LG Smart UI Regular' 54회가 slide16/37/89 의 도형 'TextBox 14' 에만 있다.)
    """
    doomed = {v for v in deckmap.source_fonts.values() if v}
    # 원본 테마 글꼴과 같은 브랜드 계열('LG EI Text …' → 'LG …')의 직박음도 함께 잡는다.
    prefixes = {v.split()[0] for v in doomed if v and " " in v}
    if not doomed and not prefixes:
        return
    ref = {"latin": "+mn-lt", "ea": "+mn-ea", "cs": "+mn-cs", "sym": "+mn-ea"}
    for part in sorted(deck.names(_SLIDE_LIKE)):
        tree = deck.xml(part)
        n = 0
        hits: set[str] = set()
        for el in tree.iter():
            face = el.get("typeface")
            if not face or face.startswith("+"):
                continue
            if face not in doomed and face.split()[0] not in prefixes:
                continue
            script = etree.QName(el).localname
            el.set("typeface", ref.get(script, "+mn-ea"))
            hits.add(face)
            n += 1
        if n:
            deck.put(part, tree)
            rep.add("font.hardcoded", part, f"{sorted(hits)} → 테마 참조", n)


# ------------------------------------------------------------ ④ 직박음 색
def _recolor_hardcoded(deck: D.Deck, profile: BrandProfile,
                       deckmap: DeckMap, rep: BrandReport) -> None:
    """theme 을 갈아도 안 바뀌는 srgbClr 직박음을 브랜드 색상환으로 옮긴다.

    ★ 문자열 치환이 아니라 XML 노드 단위로 val 을 고친다 — 같은 hex 가 우연히 다른
      곳(예: p15:sldGuideLst 의 안내선 색)에 들어 있을 때 함께 물들지 않게 하려는 것이다.
    """
    if not profile.rehue_srgb:
        return
    hue, _s, _l = hex_to_hsl(profile.primary)
    mapping = {_bare(src): rehue(src, hue) for src in profile.rehue_srgb}
    targets = sorted(deck.names(_SLIDE_LIKE)) + ["ppt/presProps.xml"]
    for part in targets:
        if part not in deck:
            continue
        tree = deck.xml(part)
        n = 0
        for el in tree.iter(D.q(A, "srgbClr")):
            val = (el.get("val") or "").upper()
            if val in mapping:
                el.set("val", mapping[val])
                n += 1
        if n:
            deck.put(part, tree)
            rep.add("color.hardcoded", part,
                    " ".join(f"{k}→{v}" for k, v in mapping.items()), n)
    for src, dst in mapping.items():
        _h, s, l = hex_to_hsl(src)
        rep.notes.append(
            f"#{src} → #{dst}  [CI 규정 아님 · 파생] 채도 {s:.0f}% · 명도 {l:.1f}% 는 "
            f"원본 그대로 두고 색상각만 {hue:.1f}°(주색)로 옮겼다 — 원본은 계산으로 "
            f"유도되지 않는 눈대중 값이었고 새 값은 재현 가능하다")
    _refresh_color_mru(deck, rep)


def _refresh_color_mru(deck: D.Deck, rep: BrandReport) -> None:
    """presProps 의 '최근 사용 색' 을 새 팔레트로 갈아 둔다.

    렌더에는 안 나오지만 제작자가 색 선택기를 열면 남의 브랜드 색이 그대로 뜬다.
    """
    part = "ppt/presProps.xml"
    if part not in deck:
        return
    tree = deck.xml(part)
    mru = tree.find(D.q(P, "clrMru"))
    if mru is None:
        return
    seen: list[str] = []
    for e in rep.palette.values():
        if e.value not in seen:
            seen.append(e.value)
    for child in list(mru):
        mru.remove(child)
    for val in seen[:10]:                       # PowerPoint MRU 는 10칸이다
        etree.SubElement(mru, D.q(A, "srgbClr")).set("val", val)
    deck.put(part, tree)
    rep.add("color.mru", part, f"clrMru {len(seen[:10])}칸 교체", len(seen[:10]))


# --------------------------------------------------------------- ⑤ 문자열
def _replace_strings(deck: D.Deck, profile: BrandProfile,
                     deckmap: DeckMap, rep: BrandReport) -> None:
    """footer 조사명과 그 밖의 지정 문자열을 바꾸고, 자동맞춤 상자 폭을 갱신한다."""
    pairs: list[tuple[str, str]] = []
    if deckmap.old_survey_name and profile.survey_name:
        pairs.append((deckmap.old_survey_name, profile.survey_name))
    pairs.extend(profile.string_map)
    if not pairs:
        return

    old_font = deckmap.source_fonts.get("minorFont.ea") or profile.font_minor
    new_font = profile.font_minor

    for part in sorted(deck.names(_SLIDE_LIKE)):
        tree = deck.xml(part)
        n = 0
        for sp in D.iter_shapes(tree):
            txt = D.shape_text(sp)
            if not txt:
                continue
            new = txt
            for src, dst in pairs:
                if src in new:
                    new = new.replace(src, dst)
            if new == txt:
                continue
            if not D.set_shape_text(tree, D.shape_name(sp), new):
                continue
            n += 1
            _refit_autofit_width(sp, txt, new, old_font, new_font, rep, part)
        if n:
            deck.put(part, tree)
            rep.add("text", part, f"{n}개 도형", n)


def _refit_autofit_width(sp: etree._Element, old_text: str, new_text: str,
                         old_font: str, new_font: str,
                         rep: BrandReport, part: str) -> None:
    """spAutoFit + wrap="none" 텍스트 상자의 ext.cx 를 새 실폭에 맞춘다.

    왜 필요한가: 이 상자들은 PowerPoint 가 열 때 스스로 늘어나므로 XML 을 안 고쳐도
    '보이기엔' 멀쩡하다. 하지만 lxml 로 기하를 실측하는 우리 검수 파이프라인은
    stale 한 ext 를 읽어 조용히 틀린 값을 보고한다.

    폭은 **델타로** 옮긴다 (새폭 = 옛폭 + (새텍스트폭 − 옛텍스트폭)). 절대값으로 다시
    잡으면 원래 상자가 들고 있던 여백(footer 는 텍스트보다 2배 넓다)이 사라진다.

    ★ 대상은 p:sp 의 txBody 뿐이다. graphicFrame(표·차트)에는 bodyPr 이 없으므로
      R-105/R-106 기하에는 절대 닿지 않는다.
    """
    body = sp.find(f"{D.q(P, 'txBody')}/{D.q(A, 'bodyPr')}")
    if body is None:
        return
    if body.get("wrap") != "none" or body.find(D.q(A, "spAutoFit")) is None:
        return
    _xf, _off, ext = _xfrm_box(sp)
    if ext is None or ext.get("cx") is None:
        return
    sz_pt, spc = _run_style(sp)
    old_w = width_emu_spc(old_text, old_font, sz_pt, spc)
    new_w = width_emu_spc(new_text, new_font, sz_pt, spc)
    cur = int(ext.get("cx"))
    nxt = max(int(round(new_w)), int(round(cur + (new_w - old_w))))
    if nxt != cur:
        ext.set("cx", str(nxt))
        rep.add("geom.autofit", part,
                f"{D.shape_name(sp)!r} ext.cx {cur}→{nxt}  "
                f"[{old_font} {old_w:,.0f} → {new_font} {new_w:,.0f} EMU "
                f"@{sz_pt}pt spc={spc}]")


def _run_style(sp: etree._Element) -> tuple[float, int]:
    """도형 첫 런의 (sz pt, spc). 없으면 (18.0, 0) — PowerPoint 기본 본문 크기."""
    sz: float | None = None
    spc: int | None = None
    for el in sp.iter():
        if el.tag not in (D.q(A, "rPr"), D.q(A, "defRPr"), D.q(A, "endParaRPr")):
            continue
        if sz is None and el.get("sz"):
            sz = int(el.get("sz")) / 100.0
        if spc is None and el.get("spc"):
            spc = int(el.get("spc"))
        if sz is not None and spc is not None:
            break
    return (sz or 18.0), (spc or 0)


# ----------------------------------------------------------- ⑥ 대체 텍스트
def _scrub_alt_text(deck: D.Deck, deckmap: DeckMap, rep: BrandReport) -> None:
    """도형 대체 텍스트(cNvPr/@descr·@title)에 남은 옛 소유주 문자열을 지운다.

    ★ 실측으로 발견한 잔재다. LG 덱 slide16/37/89 의 pic 'Picture 4' 에
      descr="… AI 홈로봇 등장…LG전자, 홈로봇 'LG 클로이드' 공개 …" 라는 보도사진
      캡션이 그대로 박혀 있다. 화면 렌더에는 안 나오므로 눈으로 보는 검수와
      '텍스트 치환' 만으로는 절대 안 걸린다. 그런데 파일 안에는 그대로 실려 나가고
      PowerPoint 의 대체 텍스트 창·접근성 검사에는 뜬다.

    속성을 통째로 **뺀다**. 빈 문자열로 두면 PowerPoint 가 '설명 있음'으로 취급한다.
    """
    if not deckmap.retire_terms:
        return
    for part in sorted(deck.names(_SLIDE_LIKE)):
        tree = deck.xml(part)
        n = 0
        killed: list[str] = []
        for cnv in tree.iter(D.q(P, "cNvPr")):
            for attr in ("descr", "title"):
                val = cnv.get(attr)
                if not val:
                    continue
                if any(term in val for term in deckmap.retire_terms):
                    del cnv.attrib[attr]
                    killed.append(f"{cnv.get('name')!r}/@{attr}: {val[:40]}…")
                    n += 1
        if n:
            deck.put(part, tree)
            rep.add("alttext.scrub", part, " | ".join(killed), n)


# ------------------------------------------------------------------ ⑦ 로고
def _retire_logos(deck: D.Deck, profile: BrandProfile,
                  deckmap: DeckMap, rep: BrandReport) -> None:
    """남의 로고를 걷어낸다. 새 로고 파일이 있으면 끼우고, 없으면 자리를 비운다.

    ★ 우리 회사 로고 pic[Embrain] 은 DeckMap.keep_shapes 로 보호되며 슬롯 탐색
      단계에서 이미 제외됐다. 여기서도 이름을 한 번 더 확인한다 (이중 안전장치).
    """
    if not deckmap.logo_slots:
        return
    light = _load_logo(profile.logo_light)
    dark = _load_logo(profile.logo_dark)
    if dark is None and light is not None and profile.auto_whiten_dark:
        dark = _whiten(light)
        if dark is not None:
            rep.notes.append(
                "어두운 배경용 로고를 밝은 배경용에서 흰색 변환해 만들었다 — CI "
                "「어두운 배경에 활용할 경우 흰색으로 변환하여 사용할 수 있습니다」 이행. "
                "색 변환일 뿐 글자꼴·굵기·비례·자간은 그대로다")

    media_added: dict[str, str] = {}       # 'light'/'dark' -> 패키지 파트 이름 (1회만 추가)
    for slot in deckmap.logo_slots:
        tree = deck.xml(slot.part)
        pic = _locate_pic(tree, slot.shape, slot.media, deck.rels(slot.part))
        if pic is None:
            rep.notes.append(f"로고 슬롯을 못 찾았다: {slot.part} {slot.shape}")
            continue
        if _pic_name(pic) in deckmap.keep_shapes:
            continue
        art = dark if slot.dark else light
        if art is None:
            _blank_out(pic, profile.logo_name())
            # pic 이 통째로 사라졌으니 그 이미지 관계도 죽었다
            deck.put(slot.part, tree)
            _prune_dead_image_rels(deck, slot.part, rep)
            rep.add("logo.blank", slot.part,
                    f"{slot.shape!r} → 빈 자리표시자 {profile.logo_name()!r} "
                    f"[off=({slot.x},{slot.y}) ext=({slot.cx}x{slot.cy}) "
                    f"{'어두운배경' if slot.dark else '밝은배경'} {slot.align}정렬 보존]")
        else:
            key = "dark" if slot.dark else "light"
            if key not in media_added:
                media_added[key] = _add_media(deck, art[0], key)
            new_cx = _reflow_logo(pic, slot, art[1])
            _repoint_pic(deck, slot.part, pic, media_added[key], profile.logo_name())
            # 트리를 먼저 되돌려야 아래 rel 정리가 '새 r:embed' 를 보고 판정한다
            deck.put(slot.part, tree)
            _prune_dead_image_rels(deck, slot.part, rep)
            rep.add("logo.swap", slot.part,
                    f"{slot.shape!r} ← {profile.name} 로고"
                    f"({'흰색변환' if slot.dark else '원색'}) "
                    f"cy={slot.cy} 고정 · cx {slot.cx}→{new_cx} "
                    f"[aspect {slot.aspect:.3f}→{art[1]:.3f}, {slot.align}정렬 유지]")
        deck.put(slot.part, tree)

    _drop_unused_media(deck, deckmap, rep)


def _load_logo(path: Path | None) -> tuple[bytes, float] | None:
    """(파일 바이트, 가로세로비). 비율은 CI 가 임의 변경을 금하므로 원본 그대로 쓴다."""
    if path is None:
        return None
    p = Path(path)
    if not p.exists():
        warnings.warn(f"로고 파일이 없다 — 자리를 비운다: {p}")
        return None
    data = p.read_bytes()
    try:
        import io
        from PIL import Image
        with Image.open(io.BytesIO(data)) as im:
            w, h = im.size
        return data, (w / h if h else 1.0)
    except Exception:                                   # PIL 이 없어도 죽지 않는다
        warnings.warn(f"로고 크기를 못 읽었다 — 비율 1.0 으로 둔다: {p}")
        return data, 1.0


def _whiten(art: tuple[bytes, float]) -> tuple[bytes, float] | None:
    """불투명 픽셀을 흰색으로 바꾼다 (알파 보존). 형태·비례는 손대지 않는다.

    CI 가 명시적으로 허용하는 '색 변환' 이다. 비례 변경이 아니므로 규정 위반이 아니다.
    """
    try:
        import io
        from PIL import Image
        with Image.open(io.BytesIO(art[0])) as im:
            rgba = im.convert("RGBA")
            alpha = rgba.getchannel("A")
            white = Image.new("L", rgba.size, 255)
            out = Image.merge("RGBA", (white, white, white, alpha))
            buf = io.BytesIO()
            out.save(buf, format="PNG")
            return buf.getvalue(), art[1]
    except Exception as exc:
        warnings.warn(f"흰색 변환 실패 — 어두운 배경 슬롯은 비워 둔다: {exc}")
        return None


def _locate_pic(tree: etree._Element, shape: str, media: str,
                rels: list[dict[str, str]]) -> etree._Element | None:
    relmap = {r["Id"]: r["Part"] for r in rels}
    for pic in tree.iter(D.q(P, "pic")):
        if _pic_name(pic) != shape:
            continue
        target = _pic_media(pic, relmap)
        if target and posixpath.basename(target).lower() == media.lower():
            return pic
    return None


def _blank_out(pic: etree._Element, name: str) -> None:
    """p:pic 을 같은 자리·같은 크기의 **빈** p:sp 로 갈아 끼운다.

    지우지 않고 비우는 이유:
      · 렌더에는 아무것도 안 나오지만 선택 창에 잡혀서 나중에 로고를 끼우기 쉽다
      · 배경색을 넣지 않는다 — 같은 자리표시자가 흰 배경(레이아웃4/6/13)과
        어두운 밴드(레이아웃5·간지) 양쪽에 쓰이기 때문이다. 채우면 한쪽이 반드시 튄다
    """
    parent = pic.getparent()
    _xf, off, ext = _xfrm_box(pic)
    sp = etree.Element(D.q(P, "sp"), nsmap=pic.nsmap)
    nv = etree.SubElement(sp, D.q(P, "nvSpPr"))
    cnv = etree.SubElement(nv, D.q(P, "cNvPr"))
    src_cnv = pic.find(f"{D.q(P, 'nvPicPr')}/{D.q(P, 'cNvPr')}")
    cnv.set("id", (src_cnv.get("id") if src_cnv is not None else None) or "900")
    cnv.set("name", name)
    etree.SubElement(nv, D.q(P, "cNvSpPr"))
    etree.SubElement(nv, D.q(P, "nvPr")).set("userDrawn", "1")
    spr = etree.SubElement(sp, D.q(P, "spPr"))
    if off is not None and ext is not None:
        nxf = etree.SubElement(spr, D.q(A, "xfrm"))
        noff = etree.SubElement(nxf, D.q(A, "off"))
        noff.set("x", off.get("x", "0"))
        noff.set("y", off.get("y", "0"))
        next_ = etree.SubElement(nxf, D.q(A, "ext"))
        next_.set("cx", ext.get("cx", "0"))
        next_.set("cy", ext.get("cy", "0"))
    geom = etree.SubElement(spr, D.q(A, "prstGeom"))
    geom.set("prst", "rect")
    etree.SubElement(geom, D.q(A, "avLst"))
    etree.SubElement(spr, D.q(A, "noFill"))
    etree.SubElement(etree.SubElement(spr, D.q(A, "ln")), D.q(A, "noFill"))
    tx = etree.SubElement(sp, D.q(P, "txBody"))
    etree.SubElement(tx, D.q(A, "bodyPr"))
    etree.SubElement(tx, D.q(A, "lstStyle"))
    etree.SubElement(etree.SubElement(tx, D.q(A, "p")),
                     D.q(A, "endParaRPr")).set("lang", "ko-KR")
    parent.replace(pic, sp)


def _reflow_logo(pic: etree._Element, slot: LogoSlot, aspect: float) -> int:
    """세로(cy)를 고정하고 cx 만 새 로고 비율로 다시 계산한다.

    ★ CI 「글자의 꼴, 굵기, 비례, 자간 등의 임의 변경을 금합니다」.
      슬롯 폭에 늘여 맞추면 규정 위반이다. 정렬 기준점도 지킨다 —
      좌정렬 슬롯은 off.x 고정, 우정렬 슬롯은 우변 고정.
    """
    _xf, off, ext = _xfrm_box(pic)
    new_cx = int(round(slot.cy * aspect))
    if ext is not None:
        ext.set("cx", str(new_cx))
        ext.set("cy", str(slot.cy))
    if off is not None:
        off.set("x", str(slot.right - new_cx) if slot.align == "right" else str(slot.x))
        off.set("y", str(slot.y))
    # 원본 로고에 걸려 있던 크롭(srcRect)은 새 이미지에 아무 의미가 없다
    fill = pic.find(D.q(P, "blipFill"))
    if fill is not None:
        for sr in fill.findall(D.q(A, "srcRect")):
            fill.remove(sr)
    return new_cx


def _repoint_pic(deck: D.Deck, part: str, pic: etree._Element,
                 media_part: str, name: str) -> None:
    """pic 의 blip 이 새 media 파트를 가리키게 하고 도형 이름을 갈아 둔다."""
    rid = _ensure_rel(deck, part, IMAGE_REL, media_part)
    fill = pic.find(D.q(P, "blipFill"))
    for blip in (fill if fill is not None else pic).iter(D.q(A, "blip")):
        blip.set(f"{{{R_NS}}}embed", rid)
        # SVG 폴백이 붙어 있으면 옛 벡터가 그대로 그려진다 — 같이 떼어낸다
        for child in list(blip):
            if b"svgBlip" in etree.tostring(child):
                blip.remove(child)
    cnv = pic.find(f"{D.q(P, 'nvPicPr')}/{D.q(P, 'cNvPr')}")
    if cnv is not None:
        cnv.set("name", name)


def _prune_dead_image_rels(deck: D.Deck, part: str, rep: BrandReport) -> None:
    """파트 안 어떤 r:embed/r:link 도 안 가리키는 **이미지** 관계를 뺀다.

    ★ 이게 없으면 옛 로고 파일이 절대 안 지워진다. blip 의 r:embed 만 새 파트로
      돌려도 .rels 에는 옛 image7.png 항목이 그대로 남고, 참조 수를 .rels 로 세는
      GC 가 '아직 쓰인다' 고 판정해 버린다 (실측으로 물렸다).

    /image 타입만 건드린다. slideLayout·slideMaster 관계는 파트 안 r:id 로
    참조되지 않는 **암묵적** 관계라 같이 지우면 덱이 깨진다.
    """
    rels_name = D.Deck.rels_name(part)
    if rels_name not in deck:
        return
    tree = deck.xml(part)
    used = {v for el in tree.iter() for k, v in el.attrib.items()
            if k.startswith(f"{{{R_NS}}}")}
    root = etree.fromstring(deck.parts[rels_name])
    dropped: list[str] = []
    for rel in list(root):
        if not (rel.get("Type") or "").endswith("/image"):
            continue
        if rel.get("Id") in used:
            continue
        dropped.append(f"{rel.get('Id')}→{rel.get('Target')}")
        root.remove(rel)
    if dropped:
        deck.put_bytes(rels_name, etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True))
        rep.add("rel.prune", part, "참조 끊긴 이미지 관계 제거: " + ", ".join(dropped),
                len(dropped))


def _ensure_rel(deck: D.Deck, part: str, rel_type: str, target_part: str) -> str:
    """파트의 .rels 에 관계를 넣고 rId 를 준다. 이미 같은 게 있으면 그걸 쓴다."""
    rels_name = D.Deck.rels_name(part)
    if rels_name in deck:
        root = etree.fromstring(deck.parts[rels_name])
    else:
        root = etree.Element(f"{{{PR_NS}}}Relationships", nsmap={None: PR_NS})
    base = posixpath.dirname(part)
    target = posixpath.relpath(target_part, base).replace("\\", "/")
    used = set()
    for rel in root:
        used.add(rel.get("Id"))
        if rel.get("Type") == rel_type and rel.get("Target") == target:
            return rel.get("Id")
    n = 1
    while f"rId{n}" in used:
        n += 1
    rid = f"rId{n}"
    el = etree.SubElement(root, f"{{{PR_NS}}}Relationship")
    el.set("Id", rid)
    el.set("Type", rel_type)
    el.set("Target", target)
    deck.put_bytes(rels_name, etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True))
    return rid


def _add_media(deck: D.Deck, data: bytes, tag: str) -> str:
    """ppt/media 에 새 이미지를 넣고 파트 이름을 준다. Content_Types Default 도 챙긴다.

    imageN.png 채번에 끼어들지 않고 brand_* 라는 별도 이름공간을 쓴다 — 나중에
    deckasm 이 슬라이드를 복제·삭제할 때 media 채번과 충돌하지 않게 하려는 것이다.
    """
    n = 1
    while f"ppt/media/brand_{tag}{n}.png" in deck:
        n += 1
    name = f"ppt/media/brand_{tag}{n}.png"
    deck.put_bytes(name, data)
    _ensure_default_ct(deck, "png", "image/png")
    return name


def _ensure_default_ct(deck: D.Deck, ext: str, mime: str) -> None:
    """[Content_Types].xml 에 확장자 Default 가 있는지 확인하고 없으면 넣는다.

    (LG 덱에는 png/jpeg/jpg/emf/svg/xlsx/rels/xml Default 가 이미 다 있어서
     실제로는 아무것도 안 한다. 다른 템플릿을 대비한 방어다.)
    """
    part = "[Content_Types].xml"
    if part not in deck:
        return
    root = etree.fromstring(deck.parts[part])
    for el in root:
        if etree.QName(el).localname == "Default" and \
                (el.get("Extension") or "").lower() == ext:
            return
    el = etree.Element(f"{{{CT_NS}}}Default")
    el.set("Extension", ext)
    el.set("ContentType", mime)
    root.insert(0, el)
    deck.put_bytes(part, etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True))


def _drop_unused_media(deck: D.Deck, deckmap: DeckMap, rep: BrandReport) -> None:
    """참조가 0이 된 옛 로고 파일을 패키지에서 뺀다.

    ★ 남겨 두면 납품물 안에 **다른 고객사 로고가 그대로 실려 나간다** — 정보 유출이다.
      파일명만 보고 지우지 않는다. 패키지의 모든 .rels 를 훑어 참조 수를 실제로 세고
      0인 것만 지운다.
    """
    if not deckmap.client_logo_media:
        return
    live: set[str] = set()
    for part in deck.names(r"_rels/.*\.rels$"):
        root = etree.fromstring(deck.parts[part])
        base = posixpath.dirname(posixpath.dirname(part))
        for rel in root:
            if rel.get("TargetMode") == "External":
                continue
            live.add(posixpath.normpath(posixpath.join(base, rel.get("Target") or "")))
    wanted = {m.lower() for m in deckmap.client_logo_media}
    for name in list(deck.parts):
        if not name.startswith("ppt/media/"):
            continue
        if posixpath.basename(name).lower() not in wanted:
            continue
        if name in live:
            rep.notes.append(f"{name} 은 아직 참조가 남아 있어 지우지 않았다")
            continue
        size = len(deck.parts[name])
        _remove_part(deck, name)
        rep.add("media.drop", name, f"참조 0 — 옛 로고 파일 제거 ({size:,} B)")


def _remove_part(deck: D.Deck, name: str) -> None:
    """★ parts / _tree_cache / _dirty 세 곳에서 동시에 뺀다.

    parts 에서만 지우면 Deck.flush() 가 _dirty 를 돌며 _tree_cache 의 트리를 다시
    써 넣어 **지운 파트가 저장 시점에 되살아난다.**
    """
    deck.parts.pop(name, None)
    deck._tree_cache.pop(name, None)
    deck._dirty.discard(name)


# ------------------------------------------------------- ⑦ 표 라벨 자간
def _tighten_label_spacing(deck: D.Deck, profile: BrandProfile,
                           rep: BrandReport) -> None:
    """표 **라벨** 셀의 자간만 좁힌다.

    ★ 범위에 주의. 이 덱의 spc="-50" 은 12,137회 있고 그중 11,103회가 표 안이며,
      표 안에서도 수치 열까지 전부 -50 이다. 전역 문자열 치환을 하면 차트 범례·본문
      멘트·수치까지 같이 좁아진다. 그래서 **a:tbl 안 · 선두 label_cols 개 열 ·
      헤더를 뺀 데이터 행** 으로만 좁힌다. (col 0 만 = 2,205회)

    pt 는 건드리지 않는다 → R-105 의 데이터 행높이(3096000/n)·헤더 3행(864000)
    체계에 아무 영향이 없다.
    """
    if profile.label_spc_from is None or profile.label_spc_to is None:
        return
    src, dst = str(profile.label_spc_from), str(profile.label_spc_to)
    for part in sorted(deck.names(r"^ppt/slides/slide\d+\.xml$")):
        tree = deck.xml(part)
        n = 0
        for tbl in tree.iter(D.q(A, "tbl")):
            for tr in D.table_rows(tbl)[D.HEADER_ROWS:]:
                for ci, tc in enumerate(D.row_cells(tr)):
                    if ci >= profile.label_cols:
                        break
                    for el in tc.iter():
                        if el.get("spc") == src:
                            el.set("spc", dst)
                            n += 1
        if n:
            deck.put(part, tree)
            rep.add("table.spc", part, f"라벨 셀 spc {src}→{dst}", n)


# ------------------------------------------------------------- ⑧ docProps
def _clean_docprops(deck: D.Deck, profile: BrandProfile,
                    deckmap: DeckMap, rep: BrandReport) -> None:
    """파일 속성에 남는 흔적 정리.

    PowerPoint 는 app.xml 을 무시하고 잘 열지만, 탐색기 속성창·미리보기에는 옛
    클라이언트의 전용 글꼴 이름이 그대로 뜬다. 납품물에 남으면 곤란하다.
    """
    part = "docProps/app.xml"
    if part in deck:
        root = etree.fromstring(deck.parts[part])
        doomed = {v for v in deckmap.source_fonts.values() if v}
        prefixes = {v.split()[0] for v in doomed if v and " " in v}
        removed = _prune_font_titles(root, doomed, prefixes, profile.font_minor)
        if removed:
            deck.put_bytes(part, etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True))
            rep.add("docprops.font", part, f"옛 전용 글꼴 칸 {removed}개 정리", removed)

    part = "docProps/core.xml"
    if part in deck and (profile.doc_title or profile.doc_last_modified_by is not None):
        tree = deck.xml(part)
        n = 0
        # ★ 네임스페이스에 끝 슬래시를 붙이지 마라. OOXML core-properties 는
        #   'http://…/metadata/core-properties' 이고 슬래시를 붙이면 find() 가 조용히
        #   못 찾는다 — 예외도 안 나고 lastModifiedBy 에 옛 작업자 이름이 그대로 남는다
        #   (실측: LG 덱 cp:lastModifiedBy 가 치환되지 않고 납품본까지 따라갔다).
        #   그래도 파일마다 접두사가 다를 수 있으니 로컬명으로도 한 번 더 찾는다.
        core = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
        for tag, val in (("{http://purl.org/dc/elements/1.1/}title", profile.doc_title),
                         (f"{core}lastModifiedBy", profile.doc_last_modified_by)):
            if val is None:
                continue
            el = tree.find(tag)
            if el is None:
                local = tag.rsplit("}", 1)[-1]
                el = next((e for e in tree
                           if etree.QName(e).localname == local), None)
            if el is not None and (el.text or "") != val:
                el.text = val
                n += 1
        if n:
            deck.put(part, tree)
            rep.add("docprops.core", part, f"{n}개 항목 갱신", n)


def _prune_font_titles(root: etree._Element, doomed: set[str],
                       prefixes: set[str], replacement: str) -> int:
    """app.xml TitlesOfParts 의 옛 글꼴 이름 칸을 새 글꼴로 갈고 중복을 뺀다.

    HeadingPairs 의 '글꼴' 개수와 vector/@size 를 함께 줄여야 PowerPoint 가
    목록을 잘못 읽지 않는다.
    """
    changed = 0
    for lpstr in root.iter(f"{{{VT_NS}}}lpstr"):
        txt = lpstr.text or ""
        if txt in doomed or (txt and txt.split()[0] in prefixes):
            lpstr.text = replacement
            changed += 1
    if not changed:
        return 0
    for vec in root.iter(f"{{{VT_NS}}}vector"):
        kids = [k for k in vec if k.tag == f"{{{VT_NS}}}lpstr"]
        if len(kids) < 2:
            continue
        seen: set[str] = set()
        dropped = 0
        for k in kids:
            if (k.text or "") != replacement:
                continue
            if replacement in seen:
                vec.remove(k)
                dropped += 1
            else:
                seen.add(replacement)
        if dropped:
            vec.set("size", str(max(0, int(vec.get("size", "0")) - dropped)))
            _shrink_heading_pair(root, "글꼴", dropped)
    return changed


def _shrink_heading_pair(root: etree._Element, label: str, n: int) -> None:
    """HeadingPairs 의 <lpstr>글꼴</lpstr><i4>N</i4> 쌍에서 N 을 줄인다."""
    for vec in root.iter(f"{{{VT_NS}}}vector"):
        kids = list(vec)
        for i, k in enumerate(kids):
            lp = k.find(f"{{{VT_NS}}}lpstr")
            if lp is not None and (lp.text or "") == label and i + 1 < len(kids):
                i4 = kids[i + 1].find(f"{{{VT_NS}}}i4")
                if i4 is not None:
                    i4.text = str(max(0, int(i4.text or "0") - n))
                return


# =========================================================================== #
# 8. audit_brand — 잔존 검사
# =========================================================================== #

DEFAULT_AUDIT_TERMS = ("LG", "홈로봇", "LG EI Text", "LG Smart UI")

# 옛 테마의 dk2/lt2. 이미지 안에 구워진 이 색을 찾는 데 쓴다.
# LG 덱 실측: 아이콘 image35.png·image36.svg 가 dk2 와 동일한 #DA1047 이다.
DEFAULT_RETIRED_COLORS = ("DA1047", "C74558")

# 검사에서 빼는 파트. 임베디드 워크북(.xlsx)은 zip 안의 zip 이라 문자열 검사가 무의미하고,
# 거기 든 카테고리 문자열은 브랜드 잔재가 아니라 '아직 안 갈아 끼운 본문 데이터' 다.
AUDIT_SKIP = re.compile(r"^(ppt/embeddings/|ppt/media/|docProps/thumbnail)")


@dataclass
class AuditHit:
    part: str
    term: str
    count: int
    where: str                 # 'typeface'|'alttext'|'color'|'text'|'asset'|'xml'
    sample: str = ""
    # 정렬·우선순위용 무게. asset 은 count 를 '파일 1개' 로 세므로(픽셀 수를 세면
    # 합계가 수십만이 되어 요약이 무의미해진다) 심각도를 여기 따로 싣는다.
    weight: int = 0

    @property
    def bucket(self) -> str:
        """이 잔재를 누가 치워야 하는가. 셋을 섞으면 '0건' 이 거짓말이 된다.

        brand    브랜드 치환이 못 지운 것 — **여기 단계에서 0이어야 한다**
                 (테마·글꼴 지정·로고·대체텍스트·footer 문자열·직박음 색)
        asset    옛 브랜드 색이 **이미지 안에 구워진** 것 — 새 그림 파일이 있어야 풀린다.
                 XML 치환으로는 절대 0이 되지 않으므로 brand 와 섞으면 안 된다.
                 (설계 [아이콘-1]: 홈로봇 아이콘 image35/36 은 색만 바꿔도 주제가 틀리다)
        content  본문 카피·설문 문항 — 본문 데이터 교체 단계 소관
                 (설계 [범위밖]: '홈로봇' 카피 248회는 브랜드 치환 대상이 아니다)
        meta     docProps 의 슬라이드 제목 목록 — 조립 후 deckasm.refresh_app_xml() 소관
        """
        if self.where == "asset":
            return "asset"
        if self.part.startswith("docProps/"):
            return "meta"
        if self.where in ("typeface", "alttext", "color"):
            return "brand"
        if self.where == "text":
            return "content"
        return "brand"

    def as_dict(self) -> dict[str, Any]:
        return {"part": self.part, "term": self.term, "count": self.count,
                "where": self.where, "bucket": self.bucket, "sample": self.sample}


def audit_brand(deck: D.Deck, terms: Sequence[str] = DEFAULT_AUDIT_TERMS,
                *, include_charts: bool = True,
                include_docprops: bool = True,
                retired_colors: Sequence[str] = DEFAULT_RETIRED_COLORS
                ) -> list[AuditHit]:
    """옛 브랜드 흔적이 남았는지 전수로 훑는다.

    '어디서' 걸렸는지를 함께 준다 — 화면에 보이는 텍스트(text)인지, 글꼴 지정
    (typeface)인지, 색(color)인지에 따라 처리가 완전히 다르기 때문이다.
    """
    hits: list[AuditHit] = []
    for part in sorted(deck.parts):
        if not (part.endswith(".xml") or part.endswith(".rels")):
            continue
        if AUDIT_SKIP.search(part):
            continue
        if not include_charts and part.startswith("ppt/charts/"):
            continue
        if not include_docprops and part.startswith("docProps/"):
            continue
        try:
            raw = deck.parts[part].decode("utf-8")
        except UnicodeDecodeError:
            continue
        for term in terms:
            n = raw.count(term)
            if not n:
                continue
            where, sample = _classify_hit(deck, part, term)
            hits.append(AuditHit(part, term, n, where, sample))
    hits.extend(audit_media_colors(deck, retired_colors))
    return hits


# --------------------------------------------------------------------------- #
# 이미지 안에 구워진 옛 브랜드 색
# --------------------------------------------------------------------------- #
# 왜 따로 있는가: theme 을 갈면 차트 136개가 전부 따라오지만, **비트맵·SVG 안의 색은
# 안 따라온다.** LG 덱은 '집+로봇' 아이콘(image35.png + image36.svg)을 dk2 와 똑같은
# #DA1047 로 구워 놓고 85개 슬라이드에서 참조한다. 이걸 안 세면 감사 결과가
# '브랜드 잔재 0' 이라고 말하면서 정작 렌더에는 빨간 아이콘이 85장 남는다.
# 실제로 slide38 을 COM 으로 렌더해 아이콘 픽셀을 찍어 보니 #DA1047 이 그대로였다.

def audit_media_colors(deck: D.Deck, colors: Sequence[str],
                       *, tol: int = 16,
                       min_pixels: int = 12) -> list[AuditHit]:
    """ppt/media 에서 옛 브랜드 색을 찾는다. SVG 는 문자열로, 래스터는 픽셀로 본다.

    tol 은 RGB 채널별 허용 오차다. PNG 가 안티에일리어싱·감마로 정확히 같은 값을
    안 갖는 경우가 있어 완전일치로 보면 놓친다.
    """
    want = [_bare(c).upper() for c in colors if c]
    if not want:
        return []
    rgbs = [(int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)) for c in want]
    refs = _media_reference_counts(deck)
    hits: list[AuditHit] = []
    for part in sorted(deck.names(r"^ppt/media/")):
        data = deck.parts[part]
        ext = part.rsplit(".", 1)[-1].lower()
        if ext in ("svg", "emf", "wmf"):
            try:
                text = data.decode("utf-8", "ignore").upper()
            except Exception:
                continue
            for hexv in want:
                n = text.count(hexv)
                if n:
                    hits.append(AuditHit(
                        part, "#" + hexv, 1, "asset",
                        f"{ext.upper()} 안에 색 {n}곳 · {_media_use(refs, part)}",
                        weight=refs.get(part, 0) * 1000 + n))
            continue
        if ext not in ("png", "jpg", "jpeg", "bmp", "gif"):
            continue
        counts = _raster_color_hits(data, rgbs, tol)
        for hexv, n in zip(want, counts):
            if n >= min_pixels:
                hits.append(AuditHit(
                    part, "#" + hexv, 1, "asset",
                    f"{ext.upper()} 픽셀 {n:,} · {_media_use(refs, part)}",
                    weight=refs.get(part, 0) * 1000 + min(n, 999)))
    return hits


def _media_reference_counts(deck: D.Deck) -> dict[str, int]:
    """media 파트별 '몇 장의 슬라이드가 쓰는가'. 심각도의 실제 척도다.

    image35.png(85장)와 image13.png(작업용 slide1 1장)을 같은 무게로 보고하면
    정작 고쳐야 할 것이 묻힌다.
    """
    out: dict[str, int] = {}
    for part in deck.names(r"^ppt/slides/slide\d+\.xml$"):
        seen = {r.get("Part") for r in deck.rels(part)}
        for target in seen:
            if target and target.startswith("ppt/media/"):
                out[target] = out.get(target, 0) + 1
    return out


def _media_use(refs: dict[str, int], part: str) -> str:
    n = refs.get(part, 0)
    return f"슬라이드 {n}장 참조" if n else "슬라이드 참조 없음(레이아웃/마스터 전용)"


def _raster_color_hits(data: bytes, rgbs: Sequence[tuple[int, int, int]],
                       tol: int) -> list[int]:
    """이미지에서 각 목표색에 가까운 픽셀 수를 센다. PIL 이 없으면 전부 0."""
    try:
        from PIL import Image
    except Exception:
        return [0] * len(rgbs)
    import io as _io
    try:
        im = Image.open(_io.BytesIO(data))
        im = im.convert("RGBA")
    except Exception:
        return [0] * len(rgbs)
    out = [0] * len(rgbs)
    for count, px in im.getcolors(maxcolors=1 << 22) or []:
        r, g, b, a = px
        if a < 128:                       # 투명한 자리는 화면에 안 나온다
            continue
        for i, (tr, tg, tb) in enumerate(rgbs):
            if abs(r - tr) <= tol and abs(g - tg) <= tol and abs(b - tb) <= tol:
                out[i] += count
    return out


def _classify_hit(deck: D.Deck, part: str, term: str) -> tuple[str, str]:
    """걸린 자리가 어디인지 판정하고 예문을 뽑는다.

    순서가 곧 우선순위다. 글꼴 지정과 대체 텍스트를 본문 카피보다 먼저 본다 —
    둘 다 화면에 안 나와서 눈으로 하는 검수로는 못 잡는 잔재이기 때문이다.
    """
    try:
        tree = etree.fromstring(deck.parts[part])
    except Exception:
        return "xml", ""
    for el in tree.iter():
        face = el.get("typeface")
        if face and term in face:
            return "typeface", face
    for cnv in tree.iter(D.q(P, "cNvPr")):
        for attr in ("descr", "title"):
            val = cnv.get(attr)
            if val and term in val:
                return "alttext", f"{cnv.get('name')}/@{attr}={val[:50]}"
    for el in tree.iter(D.q(A, "t")):
        if el.text and term in el.text:
            return "text", el.text.strip()[:60]
    for el in tree.iter(D.q(A, "srgbClr")):
        if term in (el.get("val") or ""):
            return "color", el.get("val") or ""
    return "xml", ""


BUCKET_LABEL = {
    "brand": "브랜드 잔재 — 이 단계에서 0이어야 한다",
    "asset": "이미지에 구워진 옛 브랜드 색 — 새 그림 파일이 필요하다 (설계 [아이콘-1])",
    "content": "본문 카피 — 본문 데이터 교체 단계 소관 (설계 [범위밖])",
    "meta": "파일 속성 — 조립 후 deckasm.refresh_app_xml() 소관",
}


def audit_summary(hits: Sequence[AuditHit]) -> dict[str, int]:
    out = {"brand": 0, "asset": 0, "content": 0, "meta": 0}
    for h in hits:
        out[h.bucket] += h.count
    return out


def audit_text(hits: Sequence[AuditHit], *, show: Sequence[str] = ("brand", "asset"),
               top: int = 40) -> str:
    """등급별로 나눠 보고한다.

    셋을 한 덩이로 세면 '잔존 74건' 이 되어 브랜드 치환이 실패한 것처럼 보이고,
    반대로 전부 무시하면 진짜 잔재를 놓친다. 그래서 누가 치울 것인지로 가른다.
    """
    total = audit_summary(hits)
    lines = [f"[audit] 총 {sum(total.values())}건 / {len(hits)}곳"]
    for bucket in ("brand", "asset", "content", "meta"):
        mark = "★" if bucket == "brand" and total[bucket] else " "
        lines.append(f"  {mark} {total[bucket]:5d}  {bucket:8s} {BUCKET_LABEL[bucket]}")
    for bucket in show:
        rows = [h for h in hits if h.bucket == bucket]
        if not rows:
            lines.append(f"\n  [{bucket}] 0건")
            continue
        lines.append(f"\n  [{bucket}] 상세")
        for h in sorted(rows, key=lambda x: (-(x.weight or x.count), x.part))[:top]:
            s = f"    {h.count:5d}  {h.where:9s} {h.part:34s} {h.term!r}"
            if h.sample:
                s += f"  {h.sample!r}"
            lines.append(s)
    return "\n".join(lines)


# =========================================================================== #
# 9. 검증 — 기하 불변 · 테마 실제값 · 라벨 폭 영향 · 렌더
# =========================================================================== #

def geometry_fingerprint(deck: D.Deck) -> dict[str, Any]:
    """표·차트 기하 지문. R-105/R-106 이 지켜졌는지 대조하는 데 쓴다.

    graphicFrame 의 off/ext (= 표 y·h, 차트 y·h) 와 표의 tblGrid 열폭·행높이를
    전부 담는다. 하나라도 달라지면 대조에서 걸린다.
    """
    out: dict[str, Any] = {}
    for part in sorted(deck.names(r"^ppt/slides/slide\d+\.xml$")):
        tree = etree.fromstring(deck.parts[part])
        frames = []
        for gf in tree.iter(D.q(P, "graphicFrame")):
            xf = gf.find(D.q(P, "xfrm"))
            if xf is None:
                continue
            off, ext = xf.find(D.q(A, "off")), xf.find(D.q(A, "ext"))
            frames.append((off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy"))
                          if off is not None and ext is not None else None)
        tables = []
        for tbl in tree.iter(D.q(A, "tbl")):
            tables.append((tuple(D.grid_widths(tbl)),
                           tuple(r.get("h") for r in D.table_rows(tbl))))
        if frames or tables:
            out[part] = {"frames": frames, "tables": tables}
    return out


def assert_geometry_unchanged(before: dict[str, Any],
                              after: dict[str, Any]) -> list[str]:
    """치환 전후 표/차트 기하가 한 값도 안 바뀌었음을 증명한다. 문제 목록을 준다."""
    problems: list[str] = []
    for part in sorted(set(before) | set(after)):
        b, a = before.get(part), after.get(part)
        if b == a:
            continue
        if b is None or a is None:
            problems.append(f"{part}: 한쪽에만 있다")
            continue
        if b["frames"] != a["frames"]:
            problems.append(f"{part}: graphicFrame {b['frames']} != {a['frames']}")
        if b["tables"] != a["tables"]:
            problems.append(f"{part}: 표 격자/행높이가 바뀌었다")
    return problems


def geometry_count(fp: dict[str, Any]) -> int:
    return sum(len(v["frames"]) + len(v["tables"]) for v in fp.values())


def verify_theme(deck: D.Deck, deckmap: DeckMap) -> dict[str, Any]:
    """저장된 덱의 theme 파트를 **원본 바이트에서 다시 파싱해** 실제 값을 읽는다.

    put() 한 트리를 되읽으면 우리가 쓴 값을 그대로 보게 되므로 검증이 안 된다.
    """
    out: dict[str, Any] = {}
    for part in deckmap.theme_parts:
        if part not in deck:
            continue
        tree = etree.fromstring(deck.parts[part])
        cs = tree.find(f".//{D.q(A, 'clrScheme')}")
        scheme: dict[str, str] = {}
        if cs is not None:
            for child in cs:
                srgb = child.find(D.q(A, "srgbClr"))
                scheme[etree.QName(child).localname] = (
                    srgb.get("val") if srgb is not None else "(sysClr)")
        fs = tree.find(f".//{D.q(A, 'fontScheme')}")
        fonts: dict[str, str] = {}
        if fs is not None:
            for grp in ("majorFont", "minorFont"):
                g = fs.find(D.q(A, grp))
                if g is None:
                    continue
                for script in ("latin", "ea", "cs"):
                    el = g.find(D.q(A, script))
                    if el is not None:
                        fonts[f"{grp}.{script}"] = el.get("typeface") or ""
        out[part] = {"clrScheme_name": cs.get("name") if cs is not None else None,
                     "clrScheme": scheme,
                     "fontScheme_name": fs.get("name") if fs is not None else None,
                     "fontScheme": fonts}
    return out


# -------------------------------------------------------- 표 라벨 폭 영향
@dataclass
class LabelImpact:
    text: str
    usable: int
    pt: float
    old_w: float
    new_w: float
    old_break: tuple[str, ...]
    new_break: tuple[str, ...]
    cells: int
    suspect: bool          # 병합 셀 표라 가용폭 실측을 믿을 수 없다

    @property
    def old_lines(self) -> int:
        return len(self.old_break)

    @property
    def new_lines(self) -> int:
        return len(self.new_break)

    @property
    def old_pct(self) -> float:
        return 100.0 * self.old_w / self.usable if self.usable else 0.0

    @property
    def new_pct(self) -> float:
        return 100.0 * self.new_w / self.usable if self.usable else 0.0

    @property
    def grew(self) -> bool:
        return self.new_lines > self.old_lines

    @property
    def moved(self) -> bool:
        return (not self.grew) and self.old_break != self.new_break


def measure_label_impact(deck: D.Deck, *, old_font: str, new_font: str,
                         old_spc: int = -50, new_spc: int = -100,
                         label_col: int = 0,
                         max_lines: int = 4) -> list[LabelImpact]:
    """글꼴·자간 교체가 표 라벨 줄바꿈에 미치는 영향을 **실측**한다.

    추정하지 않는다 — 덱에서 가용폭(tblGrid − marL − marR)과 pt 를 직접 읽고,
    글꼴 파일에서 advance 를 직접 재고, textfit 의 DP 로 실제 줄바꿈을 돌린다.

    ★ 병합 셀(vMerge/rowSpan) 표는 deckops.label_metrics 가 tblGrid[0] 을 라벨 열로
      보아 가용폭을 좁게 잡는 알려진 오판이 있다(LG s6 응답자 특성 계열에서
      357,120 EMU 로 나온다 — 실제 렌더와 다르다). 지우지 않고 suspect 로 표시해
      통계에서 분리한다.
    """
    seen: dict[str, LabelImpact] = {}
    for part in sorted(deck.names(r"^ppt/slides/slide\d+\.xml$")):
        tree = deck.xml(part)
        for tbl in tree.iter(D.q(A, "tbl")):
            rows = D.table_rows(tbl)
            for ridx in range(D.HEADER_ROWS, len(rows)):
                cells = D.row_cells(rows[ridx])
                if label_col >= len(cells):
                    continue
                text = D.cell_text(cells[label_col]).strip()
                if not text or _is_numeric(text):
                    continue
                # ★ suspect 판정은 '이 라벨 칸이 병합에 걸려 있는가' 로만 한다.
                #   표 어딘가에 병합이 있으면 표 전체를 버리는 방식으로 재면 LG 덱은
                #   배너 헤더가 전부 gridSpan/vMerge 라 라벨 123개 중 97개가 통째로
                #   빠져 '줄 수 증가 0개' 라는 거짓 안전 신호가 나온다. 실측으로 물렸다.
                #   label_metrics 가 실제로 틀리는 자리는 라벨 열 자신이 병합된 표
                #   (s6 응답자 특성: tblGrid[0] 을 잡아 357,120 EMU)뿐이다.
                lab = cells[label_col]
                merged = any(lab.get(k) for k in
                             ("vMerge", "rowSpan", "hMerge", "gridSpan"))
                usable, pt = D.label_metrics(tbl, row=ridx, col=label_col)
                if usable <= 0:
                    continue
                key = f"{text}|{usable}|{pt}"
                if key in seen:
                    seen[key].cells += 1
                    continue
                seen[key] = LabelImpact(
                    text=text, usable=int(usable), pt=pt,
                    old_w=label_width_hard(text, old_font, pt, old_spc),
                    new_w=label_width_hard(text, new_font, pt, new_spc),
                    old_break=tuple(wrap_ko_hard(text, usable, old_font, pt,
                                                 max_lines, old_spc)),
                    new_break=tuple(wrap_ko_hard(text, usable, new_font, pt,
                                                 max_lines, new_spc)),
                    cells=1, suspect=merged)
    return list(seen.values())


def _is_numeric(text: str) -> bool:
    return bool(re.fullmatch(r"[\d.,%()\-\s]+", text))


def label_impact_text(items: Sequence[LabelImpact], *, old_font: str,
                      new_font: str, old_spc: int, new_spc: int,
                      top: int = 20) -> str:
    solid = [i for i in items if not i.suspect]
    grew = [i for i in solid if i.grew]
    moved = [i for i in solid if i.moved]
    ratio = (sum(i.new_w for i in solid) / sum(i.old_w for i in solid)) if solid else 0.0
    pct = (100.0 * len(grew) / len(solid)) if solid else 0.0
    lines = [
        f"[표 라벨 폭 실측] {old_font} spc={old_spc}  →  {new_font} spc={new_spc}",
        f"  유니크 라벨 {len(solid)}개 "
        f"(병합셀 표라 가용폭 신뢰불가로 분리한 것 {len(items) - len(solid)}개)",
        f"  폭 배율(가중 평균) {ratio:.3f}",
        f"  줄 수 증가 {len(grew)}개 ({pct:.1f}%) · 영향 셀 {sum(i.cells for i in grew)}개",
        f"  줄 수는 같은데 줄바꿈 자리가 바뀐 것 {len(moved)}개",
    ]
    if grew:
        lines.append("  넘치는 라벨 (현재% → 교체후%, 가용폭 EMU):")
        for i in sorted(grew, key=lambda x: -x.new_pct)[:top]:
            lines.append(f"    {i.old_pct:6.1f}→{i.new_pct:6.1f}  {i.usable:>9,}  "
                         f"{i.old_lines}→{i.new_lines}줄  ×{i.cells}셀  {i.text}")
    if moved:
        lines.append("  줄바꿈 자리가 바뀐 라벨 (줄 수만 세는 검수로는 못 잡는다):")
        for i in moved[:top]:
            lines.append(f"    {i.text}")
            lines.append(f"        전: {' / '.join(i.old_break)}")
            lines.append(f"        후: {' / '.join(i.new_break)}")
    return "\n".join(lines)


def compare_spc_options(deck: D.Deck, *, old_font: str, new_font: str,
                        old_spc: int = -50,
                        options: Sequence[tuple[float | None, int]] = (
                            (None, -50), (None, -100), (None, -150))) -> str:
    """자간/크기 후보별로 '줄 수가 늘어나는 라벨이 몇 개인지' 를 실측해 표로 준다.

    options 는 (pt 재지정 | None, spc) 쌍. pt=None 이면 덱에 적힌 pt 를 그대로 쓴다
    (= R-105 행높이 체계를 안 건드리는 안).
    """
    base = measure_label_impact(deck, old_font=old_font, new_font=new_font,
                                old_spc=old_spc, new_spc=old_spc)
    solid_keys = {i.text for i in base if not i.suspect}
    lines = [f"[자간/크기 후보 실측] 기준 {old_font} spc={old_spc}"]
    for pt_override, spc in options:
        items = measure_label_impact(deck, old_font=old_font, new_font=new_font,
                                     old_spc=old_spc, new_spc=spc)
        solid = [i for i in items if not i.suspect and i.text in solid_keys]
        grew = [i for i in solid if i.grew]
        ratio = (sum(i.new_w for i in solid) / sum(i.old_w for i in solid)) if solid else 0
        tag = f"{new_font} {'덱 pt' if pt_override is None else f'{pt_override}pt'} spc={spc}"
        lines.append(f"  {tag:34s} 배율 {ratio:.3f}  줄수증가 {len(grew):3d}개 "
                     f"({100.0 * len(grew) / max(1, len(solid)):5.1f}%)  "
                     f"영향셀 {sum(i.cells for i in grew):3d}")
    return "\n".join(lines)


# -------------------------------------------------- PowerPoint COM 렌더 검증
def render_slides(pptx: Path, positions: Sequence[int], outdir: Path,
                  *, width: int = 1600, height: int = 900) -> list[Path]:
    """PowerPoint COM 으로 실제 열어 슬라이드를 PNG 로 뽑는다.

    XML 을 아무리 읽어도 '막대가 파랗게 나오는가' 는 렌더로만 확인된다.
    PowerPoint 가 없는 환경에서는 RuntimeError 를 낸다 — 조용히 건너뛰지 않는다.

    ★ 경로에 한글이 있으면 -Command 로 넘기지 못한다. Windows PowerShell 5.1 이
      명령줄을 시스템 ANSI 코드페이지로 읽어 '2026_에스원_…' 이 깨지고
      FileNotFoundException(0x8007007B) 이 난다 — 실제로 물렸다.
      그래서 스크립트를 **UTF-8 BOM 파일로 써서 -File 로** 넘긴다. BOM 이 있어야
      5.1 이 UTF-8 로 인식한다.
    """
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    exports = "\n".join(
        f'$pres.Slides.Item({p}).Export('
        f'"{(outdir / f"slide{p}.png").resolve()}","PNG",{width},{height})'
        for p in positions)
    script = ("$ErrorActionPreference='Stop'\n"
              "$ppt = New-Object -ComObject PowerPoint.Application\n"
              f'$pres = $ppt.Presentations.Open("{Path(pptx).resolve()}",'
              "$true, $false, $false)\n"
              'Write-Output ("slides=" + $pres.Slides.Count)\n'
              f"{exports}\n"
              "$pres.Close()\n$ppt.Quit()\n")
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False,
                                     encoding="utf-8-sig") as fh:
        fh.write(script)
        script_path = fh.name
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-File", script_path],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
    finally:
        Path(script_path).unlink(missing_ok=True)
    if proc.returncode != 0:
        raise RuntimeError(f"PowerPoint COM 실패: {(proc.stderr or '').strip()[:600]}")
    missing = [p for p in positions if not (outdir / f"slide{p}.png").exists()]
    if missing:
        raise RuntimeError(f"렌더 파일이 안 생겼다: {missing}")
    return [outdir / f"slide{p}.png" for p in positions]


def png_palette(path: Path, top: int = 8, *, min_pixels: int = 200,
                drop_near_white: bool = True,
                drop_neutral: bool = True) -> list[tuple[str, int]]:
    """렌더 PNG 의 주요 유채색. '막대가 파란가' 를 눈이 아니라 숫자로 확인한다."""
    import collections
    from PIL import Image
    with Image.open(path) as im:
        counter = collections.Counter(im.convert("RGB").getdata())
    out: list[tuple[str, int]] = []
    for (r, g, b), n in counter.most_common(6000):
        if n < min_pixels:
            break
        if drop_near_white and r > 235 and g > 235 and b > 235:
            continue
        if drop_neutral and abs(r - g) < 14 and abs(g - b) < 14 and abs(r - b) < 14:
            continue                              # 무채색 = 글자·선·회색 배경
        out.append(("#%02X%02X%02X" % (r, g, b), n))
        if len(out) >= top:
            break
    return out


def hue_of(hex_value: str) -> float:
    """색상각만 뽑는다. 렌더 검증에서 '붉은 계열인가 파란 계열인가' 를 가른다."""
    return hex_to_hsl(hex_value)[0]


# =========================================================================== #
# 10. 파이프라인 · CLI
# =========================================================================== #

def rebrand_file(src: Path | str, out: Path | str, profile: BrandProfile = S1,
                 *, deckmap_args: dict[str, Any] | None = None,
                 verbose: bool = True) -> dict[str, Any]:
    """열기 → 지도 → 치환 → 기하 대조 → 저장 → 잔존 검사 → 테마 실제값 확인.

    ★ 기하 대조가 깨지면 **저장하지 않고 죽는다.** R-105/R-106 을 깬 파일이
      다음 단계로 흘러가는 것보다 여기서 멈추는 편이 싸다.
    """
    deck = D.Deck(src)
    before = geometry_fingerprint(deck)
    deckmap = DeckMap.detect(deck, **(deckmap_args or LG_TEMPLATE_ARGS))
    old_font = deckmap.source_fonts.get("minorFont.ea") or ""

    rep = apply_brand(deck, profile, deckmap)

    after = geometry_fingerprint(deck)
    problems = assert_geometry_unchanged(before, after)
    if problems:
        raise AssertionError("R-105/R-106 기하가 바뀌었다 — 저장하지 않는다:\n  "
                             + "\n  ".join(problems[:10]))

    saved = deck.save(out)
    check = D.Deck(saved)
    hits = audit_brand(check)
    theme = verify_theme(check, deckmap)

    if verbose:
        print(rep.text())
        print()
        print(f"[기하] 표·차트 지문 {geometry_count(before)}개 전수 대조 — 변경 0 "
              f"(R-105 표 y/h · R-106 차트 graphicFrame · tblGrid · 행높이)")
        print(f"[저장] {saved}  ({saved.stat().st_size:,} B)")
        print()
        for part, info in theme.items():
            print(f"[테마 실제값] {part}")
            print(f"  clrScheme name={info['clrScheme_name']!r}")
            for slot, val in info["clrScheme"].items():
                mark = ""
                entry = rep.palette.get(slot)
                if entry:
                    mark = " ★공식" if entry.official else " ·파생"
                print(f"    {slot:9s} {val}{mark}")
            print(f"  fontScheme name={info['fontScheme_name']!r}  {info['fontScheme']}")
        print()
        print(audit_text(hits))

    return {"out": saved, "report": rep, "deckmap": deckmap, "audit": hits,
            "theme": theme, "old_font": old_font,
            "geometry_checked": geometry_count(before)}


def _selftest() -> int:
    """색 파생 산술 자기검증. 원본 덱 없이 즉시 돈다."""
    ok = True

    def check(label: str, got: Any, want: Any) -> None:
        nonlocal ok
        good = got == want
        ok = ok and good
        print(f"  [{'OK ' if good else 'FAIL'}] {label}: {got}   (기대 {want})")

    print("색 대수 왕복")
    for hexv in ("0033A0", "DA1047", "C74558", "FFFFFF", "000000", "F6EEEA"):
        check(hexv, hsl_to_hex(*hex_to_hsl(hexv)), hexv)

    print("파생 팔레트 — LG 원본 clrScheme 에서 에스원 주색으로")
    pal = derive_palette("#0033A0", {"dk2": "DA1047", "lt2": "C74558"})
    check("dk2", pal["dk2"].value, "0033A0")
    check("lt2", pal["lt2"].value, "3B58D1")
    check("accent1", pal["accent1"].value, "0033A0")
    check("accent2", pal["accent2"].value, "3B58D1")
    check("accent3", pal["accent3"].value, "6078DA")
    check("accent4", pal["accent4"].value, "6A77AF")
    check("accent5", pal["accent5"].value, "ABB0C4")
    check("accent6", pal["accent6"].value, "384A94")
    check("공식으로 표시된 칸", [k for k, v in pal.items() if v.official], ["dk2"])
    check("dk1/lt1/hlink/folHlink 는 손 안 댐",
          [k for k in PALETTE_KEEP if k in pal], [])

    print("틴트 사다리 — 원본과 명도가 같아야 한다 (lt2 명도를 절대값으로 고정한 이유)")
    for lm, lo in ((0.4, 0.6), (0.6, 0.4), (0.2, 0.8)):
        a = round(hex_to_hsl(lum_tint("C74558", lm, lo))[2], 1)
        b = round(hex_to_hsl(lum_tint(pal["lt2"].value, lm, lo))[2], 1)
        check(f"lumMod{lm}/lumOff{lo} 명도", b, a)

    print("상대변환을 택했다면 어떻게 됐는지 (되돌리지 말라는 근거)")
    dh, ds, dl = hex_to_hsl("DA1047")
    lh, ls, ll = hex_to_hsl("C74558")
    ph, ps, pl = hex_to_hsl("0033A0")
    naive = hsl_to_hex(ph + (lh - dh), ps * (ls / ds), pl * (ll / dl))
    print(f"       상대변환 lt2 = #{naive} (L {hex_to_hsl(naive)[2]:.1f}%) → "
          f"lumMod40 틴트 L {hex_to_hsl(lum_tint(naive, 0.4, 0.6))[2]:.1f}% "
          f"vs 원본 {hex_to_hsl(lum_tint('C74558', 0.4, 0.6))[2]:.1f}%")

    print("rehue — 채도·명도 보존")
    got = rehue("F6EEEA", hex_to_hsl("0033A0")[0])
    check("F6EEEA → 브랜드 색상환", got, "EAEEF6")
    _h0, s0, l0 = hex_to_hsl("F6EEEA")
    _h1, s1, l1 = hex_to_hsl(got)
    check("채도 보존", round(s1, 1), round(s0, 1))
    check("명도 보존", round(l1, 1), round(l0, 1))

    print("자간 폭 모델")
    w0 = width_emu_spc("가나다라", "맑은 고딕", 10.0, 0)
    w1 = width_emu_spc("가나다라", "맑은 고딕", 10.0, -50)
    check("spc -50 = 4자 × 0.5pt 만큼 좁다", round(w0 - w1), 4 * 50 // 100 * EMU_PER_PT)

    # ---- 회귀 방지: 실측으로 물렸던 두 가지 ----
    # 둘 다 '조용히 틀린' 부류다. 예외가 안 나고 그럴듯한 숫자가 나오면서
    # 결론만 뒤집히기 때문에 검증을 박아 둔다.
    print("하드 브레이크 라벨 (셀 안에 <a:br/> 가 이미 있는 경우)")
    hard = "집 안에 항상 존재하는 것에 대한\n정서적 부담"
    flat = " ".join(hard.split())
    wh = label_width_hard(hard, "맑은 고딕", 10.0, -100)
    wf = width_emu_spc(flat, "맑은 고딕", 10.0, -100)
    check("가장 넓은 줄로 재야 한다 (통짜로 재면 과대평가)", wh < wf, True)
    check("이미 박힌 줄을 그대로 지킨다",
          wrap_ko_hard(hard, 1764000, "맑은 고딕", 10.0, 4, -100),
          ["집 안에 항상 존재하는 것에 대한", "정서적 부담"])
    check("교체 후에도 2줄 유지 (넘치지 않는다)",
          len(wrap_ko_hard(hard, 1764000, "맑은 고딕", 10.0, 4, -100)), 2)

    print("빈 문자열")
    check("빈 라벨은 빈 목록", wrap_ko_spc("", 100000, "맑은 고딕"), [])

    print("\n" + ("자기검증 통과" if ok else "자기검증 실패"))
    return 0 if ok else 1


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="브랜드 치환 — 테마색·글꼴·로고·조사명 문자열")
    ap.add_argument("--src", type=Path, help="원본 pptx")
    ap.add_argument("--out", type=Path, help="출력 pptx")
    ap.add_argument("--logo", type=Path, help="밝은 배경용 클라이언트 로고 (png)")
    ap.add_argument("--logo-dark", type=Path,
                    help="어두운 배경용 로고. 생략하면 --logo 를 흰색 변환해 쓴다")
    ap.add_argument("--no-logo", action="store_true",
                    help="로고를 넣지 않고 자리만 비운다")
    ap.add_argument("--audit-only", action="store_true", help="치환 없이 잔존만 검사")
    ap.add_argument("--measure", action="store_true", help="표 라벨 폭 영향 실측")
    ap.add_argument("--spc-options", action="store_true",
                    help="자간 후보(-50/-100/-150)별 줄 수 증가 비교")
    ap.add_argument("--render", type=int, nargs="*", default=None,
                    help="PowerPoint COM 으로 렌더할 슬라이드 위치 (예: --render 18)")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args(argv)

    if args.selftest:
        return _selftest()
    if not args.src:
        ap.error("--src 가 필요하다")

    if args.audit_only:
        print(audit_text(audit_brand(D.Deck(args.src))))
        return 0
    if not args.out:
        ap.error("--out 이 필요하다")

    profile = S1
    if args.logo and not args.no_logo:
        profile = dc_replace(profile, logo_light=args.logo)
    if args.logo_dark and not args.no_logo:
        profile = dc_replace(profile, logo_dark=args.logo_dark)

    res = rebrand_file(args.src, args.out, profile)

    if args.measure or args.spc_options:
        src_deck = D.Deck(args.src)
        old_font = res["old_font"] or profile.font_minor
        if args.measure:
            items = measure_label_impact(
                src_deck, old_font=old_font, new_font=profile.font_minor,
                old_spc=profile.label_spc_from or 0,
                new_spc=profile.label_spc_to or 0)
            print()
            print(label_impact_text(
                items, old_font=old_font, new_font=profile.font_minor,
                old_spc=profile.label_spc_from or 0,
                new_spc=profile.label_spc_to or 0))
        if args.spc_options:
            print()
            print(compare_spc_options(src_deck, old_font=old_font,
                                      new_font=profile.font_minor,
                                      old_spc=profile.label_spc_from or 0))

    if args.render:
        outdir = Path(args.out).parent / "render"
        print()
        for p in render_slides(Path(args.out), args.render, outdir):
            pal = png_palette(p)
            hues = [f"{c}(H{hue_of(c):.0f}°)" for c, _n in pal[:4]]
            print(f"[렌더] {p.name}  주요 유채색 {hues}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
