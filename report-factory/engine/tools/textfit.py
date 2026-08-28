#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""보기 라벨 줄바꿈 — 글꼴 실측 폭 + 균형 분할 (R-108).

PowerPoint 자동 줄바꿈은 greedy 다 — 첫 줄을 꽉 채우고 남은 걸 다음 줄로 흘린다.
그래서 마지막 줄에 한 어절만 떨어지는 고아(orphan)가 생긴다.

    자동:  로봇을 즉시 통제하기 어려울 것 같은
           느낌                                  <- 고아
    균형:  로봇을 즉시 통제하기
           어려울 것 같은 느낌

그래서 자동에 맡기지 않고 `<a:br/>` 를 직접 박는다. 어디서 끊을지는 사전식 3단 비교다.

    ① 문법 벌점 합   안긴문장·수식어구를 쪼개지 않는 자리     kobreak.penalties()
    ② 고아 여부      마지막 줄에 어절 하나만 남기지 않는다
    ③ 줄 폭 제곱합   남는 자유도는 줄 길이를 고르게

②③ 을 판정하려면 문자열이 실제로 몇 EMU 인지 알아야 하고, 그건 **글꼴 파일에서 직접
재는 수밖에 없다**. 가용 폭 = tblGrid[0].w − tcPr.marL − tcPr.marR 이고 글꼴은 테마의
minor East-Asian(예: `LG EI Text TTF Regular`) 이다.

글꼴 해석 (개선점)
-----------------
프로토타입은 글꼴명 → 파일 경로를 딕셔너리에 손으로 박아 뒀다(LG 3종). 프로젝트마다
테마 글꼴이 달라지므로 그대로 두면 다음 덱에서 조용히 맑은 고딕으로 재고 폭이 어긋난다.
그래서 4단 사다리로 자동 해석한다.

    ① 별칭        한글 이름처럼 레지스트리에 안 잡히는 것 (맑은 고딕)
    ② 레지스트리   HKLM/HKCU …\\CurrentVersion\\Fonts 의 표시 이름 (즉시)
    ③ 파일명      정규화한 파일 stem 과 대조 (즉시)
    ④ name 테이블  글꼴 파일의 name 레코드 전수 스캔 (~2초, 한 번만·캐시)

넷 다 실패하면 맑은 고딕으로 폴백하되 **폴백 사실을 호출자가 알 수 있게** 한다.
폭이 틀어지면 줄바꿈 자리가 통째로 바뀌므로 조용히 넘어가면 안 된다.

    - `warnings.warn(..., FontFallbackWarning)` 를 글꼴명당 한 번 띄운다
    - `resolve_font(name).exact` 로 개별 확인
    - `font_fallbacks()` 로 실행 전체의 폴백 목록을 받아 리포트에 싣는다

검증
----
    python textfit.py

  LG 예시 덱 s23 의 라벨 10칸을 덱에서 직접 읽어(가용폭·pt·테마 글꼴 포함) 우리 폭
  모델이 실제 렌더와 같은 줄 수를 내는지 대조한다. 줄바꿈 자리 회귀는 kobreak.py 쪽.
"""

from __future__ import annotations

import functools
import os
import struct
import sys
import warnings
from dataclasses import dataclass
from pathlib import Path

from PIL import ImageFont

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kobreak as KB  # noqa: E402

EMU_PER_PT = 12700
MEASURE_PT = 100          # 큰 크기로 재고 비례 축소해야 힌팅 오차가 줄어든다

WINDIR = Path(os.environ.get('SystemRoot', r'C:\Windows'))
FONT_DIRS = [
    WINDIR / 'Fonts',
    Path(os.environ.get('LOCALAPPDATA', '')) / 'Microsoft' / 'Windows' / 'Fonts',
]
FONT_SUFFIXES = ('.ttf', '.ttc', '.otf', '.otc')
FALLBACK_FILE = 'malgun.ttf'

# 레지스트리에도 파일명에도 안 잡히는 한글 표시 이름만 최소로 둔다.
# (나머지 한글 이름은 ④ name 테이블 스캔이 정확한 face index 까지 찾아 준다)
ALIASES = {
    '맑은고딕': ('malgun.ttf', 0),
    'malgungothic': ('malgun.ttf', 0),
    '맑은고딕semilight': ('malgunsl.ttf', 0),
}


class FontFallbackWarning(UserWarning):
    """요청한 글꼴을 못 찾아 대체 글꼴로 폭을 쟀다 — 줄바꿈 자리가 어긋날 수 있다."""


@dataclass(frozen=True)
class FontResolution:
    """글꼴명 해석 결과. `exact=False` 면 폭 계산을 신뢰하면 안 된다."""
    requested: str
    path: Path
    index: int          # TTC 안의 face 번호 (TTF 는 0)
    source: str         # alias | registry | filename | nametable | fallback
    exact: bool

    def describe(self) -> str:
        return '%s → %s%s [%s]' % (
            self.requested or '(이름 없음)', self.path.name,
            '' if self.index == 0 else '#%d' % self.index,
            self.source if self.exact else self.source + '·폭 신뢰 불가')


_FALLBACKS: dict[str, FontResolution] = {}


# ─────────────────────────────────────────────────────────────
# 글꼴명 → 파일 해석
# ─────────────────────────────────────────────────────────────

def _norm(name: str) -> str:
    """비교용 정규화 — 공백·하이픈·언더바를 떼고 소문자로."""
    return ''.join(ch for ch in (name or '').lower()
                   if ch.isalnum() and ch not in '_-').strip()


def _font_files() -> list[Path]:
    out = []
    for d in FONT_DIRS:
        if not d or not d.is_dir():
            continue
        for p in d.iterdir():
            if p.suffix.lower() in FONT_SUFFIXES and p.is_file():
                out.append(p)
    return sorted(out)


@functools.lru_cache(maxsize=1)
def _alias_map() -> dict[str, tuple[Path, int]]:
    out = {}
    for key, (fname, idx) in ALIASES.items():
        for d in FONT_DIRS:
            p = (d / fname) if d else None
            if p and p.is_file():
                out[key] = (p, idx)
                break
    return out


@functools.lru_cache(maxsize=1)
def _registry_map() -> dict[str, tuple[Path, int]]:
    """레지스트리의 글꼴 표시 이름 → 파일.

    값 이름은 `LG EI Text TTF Regular (TrueType)` 꼴이고, TTC 는
    `Gulim & GulimChe & Dotum & DotumChe (TrueType)` 처럼 `&` 로 이어 붙는다 —
    이때 나열 순서가 곧 face index 다. HKCU 쪽(사용자 설치 글꼴)은 값이 전체 경로다.
    """
    try:
        import winreg
    except ImportError:                       # 윈도 밖
        return {}

    out: dict[str, tuple[Path, int]] = {}
    sub = r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(hive, sub)
        except OSError:
            continue
        try:
            n = winreg.QueryInfoKey(key)[1]
            for i in range(n):
                try:
                    disp, val, _ = winreg.EnumValue(key, i)
                except OSError:
                    continue
                if not isinstance(val, str) or not val:
                    continue
                path = Path(val)
                if not path.is_absolute():
                    path = FONT_DIRS[0] / val
                if path.suffix.lower() not in FONT_SUFFIXES:
                    continue
                disp = disp.rsplit('(', 1)[0].strip() if disp.endswith(')') else disp
                for idx, part in enumerate(disp.split('&')):
                    k = _norm(part)
                    if k and k not in out:
                        out[k] = (path, idx)
        finally:
            key.Close()
    return out


@functools.lru_cache(maxsize=1)
def _filename_map() -> dict[str, tuple[Path, int]]:
    """파일 stem 정규화 대조 — `LG EI Text TTF Regular` ↔ `LGEITextTTF-Regular.ttf`."""
    out: dict[str, tuple[Path, int]] = {}
    for p in _font_files():
        out.setdefault(_norm(p.stem), (p, 0))
    return out


def _read_name_table(path: Path) -> list[dict[int, set[str]]]:
    """글꼴 파일의 name 테이블을 face 별로 읽는다 (fontTools 없이).

    필요한 건 nameID 1(가족)·2(하위 가족)·4(전체 이름)·16(타이포그래피 가족)뿐이다.
    한글 표시 이름은 platform 3(Windows)/언어 0x0412 레코드에 UTF-16BE 로 들어 있다 —
    그래서 레지스트리에 없는 `맑은 고딕` 도 여기서는 잡힌다.
    """
    data = path.read_bytes()
    if len(data) < 12:
        return []
    tag = data[:4]
    if tag in (b'ttcf',):
        n = struct.unpack('>I', data[8:12])[0]
        bases = list(struct.unpack('>%dI' % n, data[12:12 + 4 * n]))
    else:
        bases = [0]

    faces = []
    for base in bases:
        recs: dict[int, set[str]] = {}
        num = struct.unpack('>H', data[base + 4:base + 6])[0]
        for i in range(num):
            p = base + 12 + 16 * i
            if data[p:p + 4] != b'name':
                continue
            off, _ln = struct.unpack('>II', data[p + 8:p + 16])
            count, str_off = struct.unpack('>HH', data[off + 2:off + 6])
            for j in range(count):
                q = off + 6 + 12 * j
                pid, _eid, _lid, nid, slen, soff = struct.unpack('>6H', data[q:q + 12])
                if nid not in (1, 2, 4, 16):
                    continue
                raw = data[off + str_off + soff: off + str_off + soff + slen]
                try:
                    txt = raw.decode('utf-16-be') if pid in (0, 3) else raw.decode('latin-1')
                except UnicodeDecodeError:
                    continue
                txt = txt.strip('\x00').strip()
                if txt:
                    recs.setdefault(nid, set()).add(txt)
            break
        faces.append(recs)
    return faces


@functools.lru_cache(maxsize=1)
def _nametable_map() -> dict[str, tuple[Path, int]]:
    """설치된 글꼴 전수 스캔. 느리므로(수백 개 ≈ 2초) 마지막 수단이고 한 번만 돈다."""
    best: dict[str, tuple[tuple[int, bool], Path, int]] = {}
    for path in _font_files():
        try:
            faces = _read_name_table(path)
        except Exception:                     # 깨진 글꼴 하나가 전체를 막으면 안 된다
            continue
        for idx, recs in enumerate(faces):
            regular = any(_norm(s) in ('regular', 'normal', 'book')
                          for s in recs.get(2, ()))
            # 전체 이름(4) > 가족(1) > 타이포그래피 가족(16). 같은 등급이면 Regular 우선.
            for nid, strength in ((4, 3), (1, 2), (16, 1)):
                for s in recs.get(nid, ()):
                    k = _norm(s)
                    if not k:
                        continue
                    score = (strength, regular)
                    cur = best.get(k)
                    if cur is None or score > cur[0]:
                        best[k] = (score, path, idx)
    return {k: (v[1], v[2]) for k, v in best.items()}


def _fallback_path() -> Path:
    for d in FONT_DIRS:
        p = (d / FALLBACK_FILE) if d else None
        if p and p.is_file():
            return p
    raise RuntimeError(
        '대체 글꼴 %s 조차 찾지 못했다 (%s). 폭 계산을 할 수 없다.'
        % (FALLBACK_FILE, ' · '.join(str(d) for d in FONT_DIRS if d)))


@functools.lru_cache(maxsize=64)
def resolve_font(name: str) -> FontResolution:
    """글꼴명 → FontResolution. 못 찾으면 맑은 고딕으로 폴백하고 그 사실을 남긴다."""
    key = _norm(name)
    if key:
        for source, table in (('alias', _alias_map()),
                              ('registry', _registry_map()),
                              ('filename', _filename_map()),
                              ('nametable', _nametable_map())):
            hit = table.get(key)
            if hit and hit[0].is_file():
                return FontResolution(name, hit[0], hit[1], source, True)

    res = FontResolution(name, _fallback_path(), 0, 'fallback', False)
    _FALLBACKS[name or ''] = res
    warnings.warn(
        '글꼴 %r 을 찾지 못해 %s 로 폭을 잰다 — 줄바꿈 자리가 실제 렌더와 어긋날 수 있다.'
        % (name, res.path.name), FontFallbackWarning, stacklevel=3)
    return res


def font_fallbacks() -> list[FontResolution]:
    """이번 실행에서 폴백된 글꼴 목록. 검수리포트에 그대로 실을 것."""
    return sorted(_FALLBACKS.values(), key=lambda r: r.requested)


def reset_font_cache() -> None:
    """해석 캐시와 폴백 기록을 비운다 (테스트용)."""
    for fn in (_alias_map, _registry_map, _filename_map, _nametable_map, resolve_font, _font):
        fn.cache_clear()
    _FALLBACKS.clear()


@functools.lru_cache(maxsize=16)
def _font(name):
    res = resolve_font(name)
    return ImageFont.truetype(str(res.path), MEASURE_PT, index=res.index)


# ─────────────────────────────────────────────────────────────
# 폭 계산 · 줄바꿈
# ─────────────────────────────────────────────────────────────

def width_emu(text, font_name, sz_pt=10.0):
    """문자열의 렌더 폭(EMU)."""
    return _font(font_name).getlength(text) * sz_pt / MEASURE_PT * EMU_PER_PT


def wrap_ko(text, usable_emu, font_name, sz_pt=10.0, max_lines=4):
    """보기 라벨 줄바꿈 — 안긴문장/수식어구 경계 우선, 같은 등급이면 균형.

    우선순위(사전식):
      ① 문법 벌점 합   — 안긴문장·수식어구를 쪼개지 않는 자리를 고른다 (kobreak)
      ② 고아 여부      — 마지막 줄이 어절 하나만 남지 않게
      ③ 줄 폭 제곱합   — 남는 자유도는 줄 길이를 고르게
    줄 수는 가능한 한 적게 (2줄로 되면 2줄).
    """
    text = ' '.join(text.split())
    W = lambda s: width_emu(s, font_name, sz_pt)
    if W(text) <= usable_emu:
        return [text]

    words = text.split(' ')
    if len(words) == 1:
        return _char_split(text, usable_emu, W)

    pen = KB.penalties(words)
    for L in range(2, min(max_lines, len(words)) + 1):
        best = _dp_ko(words, pen, L, usable_emu, W)
        if best:
            return best
    return _char_split(text, usable_emu, W)


def _dp_ko(words, pen, L, usable, W):
    n = len(words)
    INF = (float('inf'),) * 3

    @functools.lru_cache(maxsize=None)
    def line_w(i, j):
        return W(' '.join(words[i:j]))

    dp = [[INF] * (L + 1) for _ in range(n + 1)]
    back = [[None] * (L + 1) for _ in range(n + 1)]
    dp[0][0] = (0, 0, 0.0)
    for k in range(1, L + 1):
        for j in range(1, n + 1):
            for i in range(k - 1, j):
                if dp[i][k - 1] == INF:
                    continue
                w = line_w(i, j)
                if w > usable:
                    continue
                p, orph, sq = dp[i][k - 1]
                p += pen[i - 1] if i > 0 else 0          # words[i-1] 뒤에서 끊었다
                if k == L and j == n and (j - i) == 1 and L > 1:
                    orph += 1                             # 마지막 줄이 한 어절
                cost = (p, orph, sq + w * w)
                if cost < dp[j][k]:
                    dp[j][k] = cost
                    back[j][k] = i
    if dp[n][L] == INF:
        return None
    out, j, k = [], n, L
    while k > 0:
        i = back[j][k]
        out.append(' '.join(words[i:j]))
        j, k = i, k - 1
    return out[::-1]


def _char_split(text, usable, W):
    """어절이 하나뿐이라 어쩔 수 없을 때만 — 글자 단위 greedy."""
    lines, cur = [], ''
    for ch in text:
        if cur and W(cur + ch) > usable:
            lines.append(cur)
            cur = ch
        else:
            cur += ch
    if cur:
        lines.append(cur)
    return lines


# ─────────────────────────────────────────────────────────────
# R-108 이력 — 걷어낸 옛 구현
# ─────────────────────────────────────────────────────────────
# 프로토타입에는 `wrap_balanced()` / `_greedy_split()` / `_dp_split()` / `_is_orphan()`
# 이 더 있었다. 그때의 가설은 "기본은 PowerPoint 자동(greedy)을 그대로 두고, 마지막 줄이
# 어절 하나뿐인 고아일 때만 균형 분할로 갈아탄다" 였다. 실제로 s23 의
# `집 안에 항상 존재하는 것에 대한 / 정서적 부담` 은 greedy 와 사람 손이 일치했다.
#
# 폐기한 이유: greedy 는 폭만 보므로 **문법 경계를 우연히 맞힌 것**이었다. 폭이 조금만
# 달라지면 `집 안에 항상 존재하는 / 것에 대한 정서적 부담` 처럼 관형절과 의존명사를 찢는다.
# 그래서 문법 벌점(kobreak)을 1순위, 고아를 2순위, 폭 균형을 3순위로 둔 `wrap_ko()` 로
# 일원화했고 greedy 계열은 전부 걷어냈다. 두 경로를 남겨 두면 어느 쪽이 슬라이드에 박혔는지
# 사후에 알 수 없다 — 줄바꿈은 결정론이어야 한다.
#
# 되살릴 일이 생기면 `_dp_ko()` 에서 벌점 항(pen)만 0으로 두면 옛 `_dp_split()` 과 같다.


# ─────────────────────────────────────────────────────────────
# self-test — 글꼴 폭 모델 검증
# ─────────────────────────────────────────────────────────────

ENGINE = Path(__file__).resolve().parents[1]
SELFTEST_DECK = ENGINE / 'refs' / '예시_제작본_LG홈로봇_최종.pptx'
SELFTEST_SLIDE = 23
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

# LG 예시 덱 s23 을 실제로 렌더한 결과(육안 실측)의 줄 수. 폭 모델의 정답지다.
# 10칸 중 2칸만 두 줄이며, 그중 `집 안에 항상…` 만 원본에 <a:br/> 로 명시돼 있고
# `로봇을 즉시…` 는 PowerPoint 자동 줄바꿈이 접은 것이다 — 그래서 XML 문단 수로는
# 알 수 없고, 폭 모델이 맞아야만 같은 판정이 나온다.
_S23_RENDERED_LINES = {
    '사생활 침해 우려': 1,
    '비용 부담': 1,
    '로봇의 예상치 못한 행동': 1,
    '소음/움직임으로 인한 생활 방해': 1,
    '관리 및 유지보수': 1,
    '로봇을 즉시 통제하기 어려울 것 같은 느낌': 2,
    '집 안 공간 차지': 1,
    '사람/반려동물 안전에 대한 불안감': 1,
    '집 안에 항상 존재하는 것에 대한 정서적 부담': 2,
    '특별히 꺼려지는 점 없음': 1,
}


def _s23_label_metrics(deck: Path, slide_no: int):
    """예시 덱에서 (테마 글꼴, 가용 폭 EMU, pt, 라벨 목록) 을 직접 읽는다.

    상수를 박지 않고 덱에서 꺼내는 이유: 가용 폭 계산식
    (tblGrid[0].w − marL − marR) 자체가 검증 대상이기 때문이다.
    """
    import xml.etree.ElementTree as ET
    import zipfile

    with zipfile.ZipFile(deck) as z:
        theme = ET.fromstring(z.read('ppt/theme/theme1.xml'))
        slide = ET.fromstring(z.read('ppt/slides/slide%d.xml' % slide_no))

    ea = theme.find('.//{%s}fontScheme/{%s}minorFont/{%s}ea' % (A_NS, A_NS, A_NS))
    font = ea.get('typeface')

    tbl = next(slide.iter('{%s}tbl' % A_NS))
    col_w = int(tbl.find('{%s}tblGrid' % A_NS)[0].get('w'))

    labels, usable, sz = [], None, None
    for tr in tbl.findall('{%s}tr' % A_NS):
        tc = tr.findall('{%s}tc' % A_NS)[0]
        paras = tc.findall('.//{%s}txBody/{%s}p' % (A_NS, A_NS))
        text = ' '.join(''.join(r.find('{%s}t' % A_NS).text or ''
                                for r in p.findall('{%s}r' % A_NS)).strip()
                        for p in paras).strip()
        text = ' '.join(text.split())
        if text not in _S23_RENDERED_LINES:
            continue                                  # 헤더·범례·(Base) 행
        pr = tc.find('{%s}tcPr' % A_NS)
        mar_l = int((pr.get('marL') if pr is not None else None) or 91440)
        mar_r = int((pr.get('marR') if pr is not None else None) or 91440)
        usable = col_w - mar_l - mar_r
        for e in tc.iter():
            if e.tag in ('{%s}rPr' % A_NS, '{%s}endParaRPr' % A_NS) and e.get('sz'):
                sz = int(e.get('sz')) / 100.0
                break
        labels.append(text)
    return font, usable, sz or 10.0, labels


def _selftest() -> int:
    fail = 0

    print('① 글꼴 해석 사다리')
    for name in ('LG EI Text TTF Regular', '맑은 고딕', 'Malgun Gothic',
                 '없는글꼴 Nonexistent Face'):
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', FontFallbackWarning)
            res = resolve_font(name)
        print('   %s' % res.describe())
    miss = [r for r in font_fallbacks()]
    print('   폴백 %d건 — %s' % (len(miss), ', '.join(repr(r.requested) for r in miss) or '없음'))
    if len(miss) != 1 or miss[0].requested != '없는글꼴 Nonexistent Face':
        print('   FAIL  폴백 보고가 어긋난다 (없는 글꼴 1건만 잡혀야 한다)')
        fail += 1

    print()
    print('② 폭 모델 검증 — %s s%d 라벨열' % (SELFTEST_DECK.name, SELFTEST_SLIDE))
    if not SELFTEST_DECK.is_file():
        print('   SKIP  예시 덱이 없다: %s' % SELFTEST_DECK)
        return 1 if fail else 0

    font, usable, sz, labels = _s23_label_metrics(SELFTEST_DECK, SELFTEST_SLIDE)
    res = resolve_font(font)
    print('   테마 글꼴 %r · 가용폭 %s EMU · %gpt · 라벨 %d칸'
          % (font, format(usable, ','), sz, len(labels)))
    print('   %s' % res.describe())
    if not res.exact:
        print('   SKIP  테마 글꼴이 폴백됐다 — 폭 기준이 달라 대조하지 않는다.')
        return 1 if fail else 0
    if len(labels) != len(_S23_RENDERED_LINES):
        print('   FAIL  라벨 %d칸을 기대했는데 %d칸을 읽었다'
              % (len(_S23_RENDERED_LINES), len(labels)))
        fail += 1

    fit, over = [], []
    for text in labels:
        ratio = width_emu(text, font, sz) / usable * 100
        lines = wrap_ko(text, usable, font, sz)
        want = _S23_RENDERED_LINES[text]
        ok = len(lines) == want
        fail += not ok
        (fit if ratio < 100 else over).append(ratio)
        print('   %s %6.1f%%  %d줄  %s'
              % ('OK ' if ok else 'FAIL', ratio, len(lines), ' / '.join(lines)))
        if not ok:
            print('          실측은 %d줄이다' % want)

    print()
    print('   1줄 %d칸(최대 %.1f%%) · 2줄 %d칸(최소 %.1f%%) · 안전마진 %.1f%%p'
          % (len(fit), max(fit) if fit else 0, len(over),
             min(over) if over else 0,
             (min(over) - max(fit)) if fit and over else 0))
    if fit and max(fit) >= 100:
        print('   FAIL  한 줄로 렌더된 칸이 100%% 를 넘는다')
        fail += 1
    if over and min(over) <= 100:
        print('   FAIL  두 줄로 렌더된 칸이 100%% 를 못 넘는다')
        fail += 1

    print()
    print('textfit self-test: %s' % ('통과' if not fail else '실패 %d건' % fail))
    return 1 if fail else 0


if __name__ == '__main__':
    sys.exit(_selftest())
