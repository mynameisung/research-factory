# -*- coding: utf-8 -*-
"""
deck_lint.py -- 납품 직전 덱의 **내용 위생** 검사

  ★ deckbuild.validate 와 무엇이 다른가 ★
  validate 는 「PowerPoint 가 여는가」를 본다. 이 도구는 「사람이 보기에 틀린 데가
  없는가」를 본다. 파일은 멀쩡한데 오타가 있거나, 제작 메모가 남아 있거나,
  이전 고객사 이름이 섞여 있는 것 — 전부 열리기는 잘 열린다.

  검사는 전부 **코퍼스 실측에서 나왔다**(engine/refs/corpus_20260824, 5사례).
  근거 없는 검사는 넣지 않는다 — 오탐 한 번이면 다음부터 아무도 안 본다.

사용법:
    python deck_lint.py <deck.pptx> [--json out.json]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
SLIDE_RE = re.compile(r"ppt/slides/slide(\d+)\.xml$")

#: 한글 낱말 — 오탈자 짝짓기 대상. 숫자·영문이 섞인 토큰은 뺀다.
_HANGUL_WORD = re.compile(r"[가-힣]{3,}")

#: 제작팀 메모 문구. 실측(코퍼스 3사례 제작본 66·16·72건 / 최종본 전부 0건).
MEMO_MARKERS = ("Data 입력", "데이터 입력", "보기갯수", "보기 개수", "장 분리",
                "레이블은 조정", "조정 필요", "확인 부탁", "수정 부탁")

#: 미기입 잔존 마커
UNFILLED_MARKERS = ("???", "내용을 입력하세요", "내용을입력하세요", "TBD", "XXX",
                    "텍스트를 입력하십시오")


def slide_texts(path: Path) -> dict[int, list[str]]:
    """슬라이드 번호 → 그 장의 텍스트 조각들."""
    out: dict[int, list[str]] = {}
    z = zipfile.ZipFile(path)
    for name in z.namelist():
        m = SLIDE_RE.match(name)
        if not m:
            continue
        tree = etree.fromstring(z.read(name))
        bits = [t.text for t in tree.iter(f"{{{A}}}t") if t.text and t.text.strip()]
        out[int(m.group(1))] = bits
    return out


def jamo(ch: str) -> tuple[int, int, int] | None:
    """한글 음절을 (초성, 중성, 종성) 로 쪼갠다. 한글이 아니면 None."""
    o = ord(ch) - 0xAC00
    if not (0 <= o < 11172):
        return None
    return (o // 588, (o % 588) // 28, o % 28)


def diff_at(a: str, b: str) -> int | None:
    """길이가 같고 **한 글자만** 다를 때 그 자리. 아니면 None."""
    if len(a) != len(b) or a == b:
        return None
    idx = [i for i, (x, y) in enumerate(zip(a, b)) if x != y]
    return idx[0] if len(idx) == 1 else None


def looks_like_typo(a: str, b: str) -> bool:
    """`a` 와 `b` 가 **오타 짝**으로 보이는가.

    ★ 빈도로는 못 가른다. 실측에서 진짜 오타(`청년암심주택` 3회)가 흔한 오탐
      (`단지명` 2회 / `거주지` 1회)보다 오히려 자주 나왔다 — 빈도 문턱을 낮추면
      오탐이 쏟아지고, 높이면 정답이 걸러진다. 그래서 **글자 구조**로 가른다.

    세 조건을 다 만족해야 오타로 본다.
      ① 낱말이 4글자 이상 — 짧을수록 서로 다른 낱말이 우연히 한 글자만 다를 확률이 높다
      ② 다른 자리가 **마지막 글자가 아니다** — 마지막은 접미사가 갈리는 자리다
         (`단지명`/`단지별` · `거주지`/`거주자` 는 서로 다른 낱말이지 오타가 아니다)
      ③ 그 글자의 초·중·종성 중 **하나만** 다르다 — 오타는 자모 하나가 어긋난다
         (`안`ㅇㅏㄴ / `암`ㅇㅏㅁ 는 종성만 다르다. `명`ㅁㅕㅇ / `별`ㅂㅕㄹ 은 둘이 다르다)
    """
    i = diff_at(a, b)
    if i is None or len(a) < 4 or i == len(a) - 1:
        return False
    ja, jb = jamo(a[i]), jamo(b[i])
    if not ja or not jb:
        return False
    return sum(1 for x, y in zip(ja, jb) if x != y) == 1


def check_typos(texts: dict[int, list[str]], *, min_major: int = 3) -> list[dict]:
    """덱 안에서 **거의 같은 낱말이 두 가지로** 나타나면 드문 쪽이 오타다.

    ★ 근거: 코퍼스 2(SH) 제작본 — `청년암심주택`(3회) 과 `청년안심주택`(다수)이
      한 덱에 공존했고, 최종본에서 전자가 전부 후자로 고쳐졌다(실측 3장).
      맞춤법 사전으로는 못 잡는다(둘 다 사전에 없는 고유명사다). **자기 일관성**이
      유일한 단서다 — 같은 문서가 같은 것을 두 이름으로 부를 리 없다.

    드문 쪽이 1~2회이고 흔한 쪽이 min_major 회 이상일 때만 든다. 둘 다 여러 번
    나오면 진짜 다른 낱말일 수 있다(예: 「거주자」/「거주지」).
    """
    where: dict[str, set[int]] = defaultdict(set)
    freq: Counter = Counter()
    for no, bits in texts.items():
        for b in bits:
            for w in _HANGUL_WORD.findall(b):
                freq[w] += 1
                where[w].add(no)
    by_len: dict[int, list[str]] = defaultdict(list)
    for w in sorted(freq):
        if len(w) >= 4:
            by_len[len(w)].append(w)

    out = []
    for _, group in sorted(by_len.items()):
        for i, a in enumerate(group):
            for b in group[i + 1:]:
                if not looks_like_typo(a, b):
                    continue
                rare, major = (a, b) if freq[a] < freq[b] else (b, a)
                if freq[major] < min_major or freq[major] < freq[rare] * 2:
                    continue
                out.append({
                    "check": "TYPO_NEAR_DUP",
                    "severity": "high",
                    "ratio": round(freq[major] / max(freq[rare], 1), 1),
                    "slides": sorted(where[rare]),
                    "message": f"「{rare}」({freq[rare]}회) 는 「{major}」({freq[major]}회) 의 "
                               f"오타로 보인다 — 한 덱이 같은 것을 두 이름으로 부르고 있다",
                })
    return out


def check_memos(texts: dict[int, list[str]]) -> list[dict]:
    """제작 메모가 남아 있는가.

    ★ 근거: 코퍼스 제작본 3벌에 66·16·72건 있고 **최종본 5벌 전부 0건**이다.
      납품본에 남으면 「미완성을 보냈다」로 읽힌다.
    """
    out = []
    for no, bits in sorted(texts.items()):
        for b in bits:
            for mk in MEMO_MARKERS:
                if mk in b:
                    out.append({"check": "PRODUCTION_MEMO", "severity": "fatal",
                                "slides": [no],
                                "message": f"제작 메모가 남아 있다: {b.strip()[:60]!r}"})
                    break
    return out


#: 연구원 검토 메모의 말투. **제작팀 메모(`Data 입력`)와 다른 부류다** —
#: 제작팀은 우리에게 지시를 남기고, 연구원은 자기가 할 일을 적어 둔다.
REVIEW_MEMO_RE = re.compile(
    r"(검토|수정|추가|확인|보완|재확인|정렬|변경)\s*(필요|요망|바람|부탁|고민)"
    r"|워딩\s*수정|자릿수\s*변경|내림차순\s*정렬|바꾸는\s*게"
    r"|해\s*주세요|부탁\s*드립니다|(으로|로)\s*내리기"
)


def check_review_memos(path: Path) -> list[dict]:
    """**연구원 검토 메모**가 남아 있는가.

    ★ `check_memos` 와 근거가 다르다. 저건 제작팀이 우리에게 남긴 지시(`Data 입력`)이고
      이건 연구원이 자기에게 남긴 할 일이다(`멘트 검토 필요` · `소수점 자릿수 변경`).
      실측(코퍼스 3 한옥 **최종본**): 5건이 그대로 납품본에 남았다 —
      s4 「방문관련 내용 추가」 · s9 「구매의향은 30% 미만으로 바꾸는 게..」 ·
      s10 「희망면적/구매가격 워딩 수정 / 소수점 자릿수 변경」 · s12 · s14 「멘트 검토 필요」.
    ★ **자리로 못 찾는다.** 제작팀 메모는 장 맨 위(top≈0)에 핀으로 꽂지만
      연구원 메모는 관련 내용 옆(0.26~1.83cm)에 둔다. 3벌 제작본 154건은 전부 top≈0,
      한옥 최종본 5건은 전부 top>0.05 였다 — 그래서 **말투로 찾는다.**
    """
    from pptx import Presentation
    from pptx.util import Emu                                    # noqa: F401
    out = []
    prs = Presentation(str(path))

    def walk(shapes):
        for sh in shapes:
            yield sh
            if sh.shape_type == 6:                               # GROUP
                yield from walk(sh.shapes)

    for i, sl in enumerate(prs.slides, 1):
        for sh in walk(sl.shapes):
            if sh.shape_type != 1 or not sh.has_text_frame:      # AUTO_SHAPE 만
                continue
            t = sh.text_frame.text.strip()
            if not t or len(t) > 200:
                continue
            if REVIEW_MEMO_RE.search(t):
                out.append({"check": "REVIEW_MEMO", "severity": "high", "slides": [i],
                            "message": f"연구원 검토 메모가 남아 있다: "
                                       f"{t[:60].replace(chr(10), ' / ')!r}"})
    return out


def check_unfilled(texts: dict[int, list[str]]) -> list[dict]:
    out = []
    for no, bits in sorted(texts.items()):
        for b in bits:
            for mk in UNFILLED_MARKERS:
                if mk in b:
                    out.append({"check": "UNFILLED_MARKER", "severity": "fatal",
                                "slides": [no],
                                "message": f"미기입 자리가 남아 있다: {b.strip()[:60]!r}"})
                    break
    return out


def check_decimals(texts: dict[int, list[str]]) -> list[dict]:
    """퍼센트 값의 소수 자릿수가 튀는가.

    ★ 근거: 코퍼스 3(한옥) 제작본 소제목 `'150평 이상(24.15)'` → 최종본에서 교정.
      통계표는 소수 1자리로 내려오므로 2자리 이상은 **붙여쓰기 사고**이거나
      다른 지표를 잘못 옮긴 것이다. 평균·금액은 자릿수가 다를 수 있으므로
      백분율 문맥(`%` 또는 괄호 안 단독 수)만 본다.
    """
    pat = re.compile(r"(?<![\d.])(\d{1,3}\.\d{2,})\s*%")
    out = []
    for no, bits in sorted(texts.items()):
        for b in bits:
            for m in pat.finditer(b):
                out.append({"check": "DECIMAL_PLACES", "severity": "low", "slides": [no],
                            "message": f"백분율 소수 {m.group(1)}% — 통계표는 1자리다"})
    return out


#: 표의 Base 칸. `(500)` · `(13*)` — 별표는 저베이스 표식(R-128)
_BASE_CELL = re.compile(r"^\(\s*(\d+)\s*\*?\s*\)$")


def check_tiny_base(path: Path) -> list[dict]:
    """**응답자 한 명짜리 배너 열**이 장표에 실려 있는가.

    ★ 저베이스 표시(R-128 별표)와 다른 문제다. 별표는 「이 수치를 믿지 말라」는 말이고,
      이건 **「이 사람이 누군지 좁혀진다」**는 말이다. n=1 열은 그 한 명의 응답이
      열 전체로 그대로 읽힌다 — 실측(코퍼스 사례 2 제작본 s41): 어느 집단 열이
      「매우 만족한다 100.0 · 평균 4.00」 이라 한 명의 답이 고스란히 드러났고,
      **최종본에서 그 열(n=1)과 n=6 열만 빠졌다.**
      다만 「저베이스면 뺀다」가 일반 규칙은 아니다 — 3사례에서 n=8·9·10·11·12 는
      전부 남는다. 그래서 **차단하지 않고 사람에게 묻는다.**
    """
    out = []
    z = zipfile.ZipFile(path)
    for name in sorted(z.namelist()):
        m = SLIDE_RE.match(name)
        if not m:
            continue
        tree = etree.fromstring(z.read(name))
        for tbl in tree.iter(f"{{{A}}}tbl"):
            for tr in tbl.findall(f"{{{A}}}tr"):
                for tc in tr.findall(f"{{{A}}}tc"):
                    txt = "".join(t.text or "" for t in tc.iter(f"{{{A}}}t")).strip()
                    mm = _BASE_CELL.match(txt)
                    if not mm:
                        continue
                    n = int(mm.group(1))
                    if n <= 2:
                        out.append({
                            "check": "TINY_BASE_COLUMN", "severity": "high",
                            "slides": [int(m.group(1))],
                            "message": f"응답자 {n}명짜리 열이 실려 있다 {txt} — "
                                       f"그 사람의 응답이 열에서 그대로 읽힌다. "
                                       f"수치 신뢰도(별표)와 별개로 노출 문제다",
                        })
    # 장·열 단위로 접어 준다 — 한 장에 여러 개면 한 줄로
    folded: dict[int, int] = {}
    for f in out:
        folded[f["slides"][0]] = folded.get(f["slides"][0], 0) + 1
    return [{"check": "TINY_BASE_COLUMN", "severity": "high", "slides": [sl],
             "message": f"응답자 1~2명짜리 열이 {c}개 실려 있다 — 그 사람의 응답이 "
                        f"열에서 그대로 읽힌다(수치 신뢰도와 별개인 노출 문제)"}
            for sl, c in sorted(folded.items())]


def check_foreign_client(texts: dict[int, list[str]], *, own: str = "") -> list[dict]:
    """**우리 템플릿을 물려준 조사**의 고유명사가 남아 있는가.

    ★ 아무 고객사 이름이나 넣으면 안 된다. 검사 대상은 「이 덱이 베껴 온 원본」의
      이름뿐이다 — 우리 템플릿은 LG 홈로봇·SKEE 제작본에서 왔다. 남의 조사 이름을
      목록에 넣으면 그 조사를 검사할 때 전 장이 지적으로 뜬다(실측: 청년안심주택을
      넣었더니 SH 최종본 110장이 지적됐다. 그 덱에서는 그게 **정답**이다).
    """
    known = ("LG전자", "LG 전자", "홈로봇", "SK에코엔지니어링", "SKEE")
    out = []
    for no, bits in sorted(texts.items()):
        for b in bits:
            for k in known:
                if own and k in own:
                    continue
                if k in b:
                    out.append({"check": "FOREIGN_CLIENT", "severity": "high", "slides": [no],
                                "message": f"다른 조사의 고유명사 「{k}」 가 남아 있다: {b.strip()[:50]!r}"})
                    break
    return out


CHECKS = (check_memos, check_unfilled, check_typos, check_decimals)


def lint(path: Path, *, own: str = "") -> dict[str, Any]:
    texts = slide_texts(path)
    findings: list[dict] = []
    for fn in CHECKS:
        findings.extend(fn(texts))
    findings.extend(check_foreign_client(texts, own=own))
    findings.extend(check_tiny_base(path))
    findings.extend(check_review_memos(path))
    # 빈도 차가 클수록 오타일 확률이 높다 — 확실한 것부터 보여 준다.
    # 실측 정밀도(코퍼스 최종본 5벌 13건 중 명백한 오타 9건):
    #   비 10 이상 = 사실상 확정(`출입동제` 1:423 · `물리모안` 2:374)
    #   비 5 미만  = 진짜 다른 낱말일 수 있다(`이상적인`/`일상적인`) — 사람이 본다
    order = {"fatal": 0, "high": 1, "low": 2}
    findings.sort(key=lambda f: (order.get(f["severity"], 9), -f.get("ratio", 0), f["slides"]))
    return {"deck": path.name, "slides": len(texts), "findings": findings,
            "counts": dict(Counter(f["check"] for f in findings))}


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--own", default="", help="이번 조사의 고객사·주제어(오탐 방지)")
    ap.add_argument("--json")
    a = ap.parse_args(argv[1:])
    res = lint(Path(a.deck), own=a.own)
    if a.json:
        Path(a.json).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{res['deck']} · {res['slides']}장 · 지적 {len(res['findings'])}건 {res['counts']}")
    for f in res["findings"][:40]:
        print(f"  [{f['severity']:5}] {f['check']:16} s{f['slides'][:6]} {f['message']}")
    return 1 if any(f["severity"] == "fatal" for f in res["findings"]) else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
