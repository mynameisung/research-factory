# -*- coding: utf-8 -*-
"""파트 팩 생성기 — RULES.md 를 담당 에이전트별 슬라이스로 자동 추출한다.

왜: 규칙 131장을 통째로 주면 컨텍스트가 희석된다(guide/18 §3-3). 담당마다 자기
파트의 카드만 로드한다. **손으로 두 벌 유지하지 않는다** — 이 스크립트가 유일한
생성 경로이고, /회고 가 RULES.md 를 갱신할 때마다 재실행한다.

사용: python engine/tools/make_packs.py            # engine/learning/packs/ 재생성
"""
import io
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]          # engine/
RULES = ROOT / "learning" / "RULES.md"
OUT = ROOT / "learning" / "packs"

#: 팩 → 그 팩이 가져갈 분류 **키워드**. 태그를 완전일치로 적으면 표기 변이 하나마다 카드가
#: 통째로 빠진다 — 실측 2026-09-01: `guide/요약결론` 만 적어 둔 탓에 `guide / 요약` ·
#: `guide / 요약·결론` · `guide / 결론` · `guide / 해석` · `guide / 표기 · 요약결론` ·
#: `guide / 요약결론 · 검수` · `guide / 요약 · 오케스트레이션` 일곱 변이가 배달되지 않았고,
#: 그 안에 R-185·R-186(모수·자 규율) · R-199(피라미드) · R-203 · R-207 이 들어 있었다.
#: 그래서 **부분일치**로 바꾼다. 한 카드가 여러 팩에 실릴 수 있다(목차·요약결론 경계 카드).
PACK_MAP = {
    "목차": ["목차"],
    "멘트": ["멘트", "행수식어", "표기"],
    # 「해석」은 R-185(모수)·R-186(자) 자리다 — 요약이 표를 읽기 전에 반드시 딛는 규율이다.
    # 「표해석」은 R-230(배너 인사이트)·R-231(구간 접기) 자리다 — 요약이 표에서 무엇을
    # 집을지 정하는 규율이라 요약결론 담당에게 반드시 가야 한다 (배달 누락 실측 2026-09-01).
    "요약결론": ["요약", "결론", "해석"],
    "조사개요": [],                       # 분류 태그가 없어 EXTRA_CARDS 로만 채운다
    "공통": ["공통/", "산출물스키마"],
}

#: 분류 태그 경계와 안 맞는 카드의 명시 등재 (guide-judge 심사 F-3·F-4 반영).
#: 태그를 고치는 게 아니라 여기 더한다 — 카드 분류는 원저자 표기 그대로 둔다.
EXTRA_CARDS = {
    # 배너 설계는 guide-toc 임무다 (guide/18 §2 — 목차·배너 쌍 산출)
    "목차": ["R-055", "R-056", "R-057", "R-089"],
    # guide-front (Part Ⅰ + Cover) 가 실제로 딛는 카드들
    "조사개요": ["R-011", "R-013", "R-022", "R-046", "R-124", "R-127"],
    # 카드 포맷(분류·근거·신뢰도)을 안 따르는 산문 카드와, 요약 담당이 실제로 딛는 경계 카드.
    # 태그가 없다는 이유로 빠지면 「설계의 닻은 조사 목적」(R-141) 같은 1급 규칙이 배달되지 않는다.
    "요약결론": ["R-119", "R-121", "R-126", "R-128", "R-141", "R-142", "R-143",
              "R-144", "R-145", "R-146", "R-215", "R-216", "R-217"],
}

#: 팩별로 담당이 함께 읽어야 하는 knowledge 문서 (포인터 — 내용 복제 금지)
PACK_DOCS = {
    "목차": ["knowledge/guide/10_목차-설계규칙.md"],
    "멘트": ["knowledge/guide/13_멘트-작성규칙.md"],
    "요약결론": ["knowledge/guide/14_요약결론-작성규칙.md"],
    "조사개요": ["knowledge/00_업무개요.md"],
    "공통": ["knowledge/guide/17_팀구조-제작가이드.md",
             "knowledge/guide/18_오케스트레이션-제작가이드.md"],
}


def norm(tag: str) -> str:
    return re.sub(r"\s+", "", tag).replace("·", "·")


def parse_cards(text: str):
    """RULES.md → [(id, title, 분류, 신뢰도, 본문)]"""
    cards = []
    # 카드 헤더는 두 형식이 섞여 있다 — `### R-070 · 제목`(초기)과 `## R-160 제목`(R-135 이후).
    # ID 이름공간도 둘이다(`R-NNN` · `R-PPT-NNN`). 한 형식만 잡으면 카드가 통째로 팩에서
    # 빠지고 담당 에이전트는 그 규칙을 영영 못 본다 (실측 2026-09-01: 228장 중 97장 누락).
    for m in re.finditer(r"^#{2,3} (R-[A-Z]*-?\d+)(?:\s*·\s*|\s+)(.+?)$\n(.*?)(?=^#{1,3} |\Z)",
                         text, re.M | re.S):
        rid, title, body = m.group(1), m.group(2).strip(), m.group(3)
        cat = re.search(r"^- 분류:\s*(.+?)\s*$", body, re.M)
        trust = re.search(r"^- 신뢰도:\s*(.+?)\s*$", body, re.M)
        cards.append((rid, title, norm(cat.group(1)) if cat else "",
                      (trust.group(1) if trust else "").strip(), body.rstrip()))
    return cards


def main():
    cards = parse_cards(io.open(RULES, encoding="utf-8").read())
    OUT.mkdir(exist_ok=True)
    report = []
    used = set()
    by_id = {c[0]: c for c in cards}
    def _deprecated(card):
        # 폐기된 카드는 팩 맨 뒤로 보낸다 — 앞에 섞여 있으면 담당이 그대로 적용한다.
        # 지우지 않는 이유는 폐기 사유를 함께 보여 규칙이 되살아나는 것을 막기 위해서다.
        return "폐기" in (card[3] or "") or "【폐기" in (card[1] or "")

    for pack, tags in PACK_MAP.items():
        keys = [norm(t) for t in tags]
        def _hit(cat):
            # guide 계열은 키워드 부분일치, 공통 계열은 접두로 가른다.
            return bool(cat) and any(k in cat for k in keys)
        picked = [c for c in cards if _hit(c[2])]
        for rid in EXTRA_CARDS.get(pack, []):
            if rid not in by_id:
                raise SystemExit(f"EXTRA_CARDS 의 {rid} 가 RULES.md 에 없다 — 오타이거나 폐기됨")
            if by_id[rid] not in picked:
                picked.append(by_id[rid])
        # 폐기 카드는 맨 뒤로, 나머지는 번호순. ID 이름공간이 둘이라(R-NNN · R-PPT-NNN)
        # 숫자만 뽑아 정렬한다 — int(c[0][2:]) 는 R-PPT-138 에서 터진다.
        def _key(c):
            digits = re.sub(r"\D", "", c[0])
            return (1 if _deprecated(c) else 0, int(digits or 0))
        picked.sort(key=_key)
        used.update(c[0] for c in picked)
        lines = [f"# 팩: {pack} — 담당 에이전트용 규칙 슬라이스",
                 "",
                 f"자동 생성 (make_packs.py). 원본은 `engine/learning/RULES.md` — **여기를 고치지 말 것.**",
                 f"카드 {len(picked)}장. 신뢰도 `가설` 카드는 적용 전 사용자 확인이 필요하다.",
                 ""]
        docs = PACK_DOCS.get(pack) or []
        if docs:
            lines.append("**먼저 읽을 knowledge** (내용은 원문으로 — 복제하지 않는다):")
            lines += [f"- `engine/{d}`" for d in docs]
            lines.append("")
        for rid, title, cat, trust, body in picked:
            lines.append(f"### {rid} · {title}")
            lines.append(body)
            lines.append("")
        (OUT / f"{pack}.md").write_text("\n".join(lines), encoding="utf-8", newline="\n")
        report.append(f"  {pack:<6} 카드 {len(picked):>3}장  {sum(len(c[4]) for c in picked):>7,}자")

    print(f"팩 {len(PACK_MAP)}종 생성 → {OUT}")
    print("\n".join(report))
    uncovered = [(c[0], c[2]) for c in cards if c[0] not in used]
    print(f"어느 팩에도 안 실린 카드 {len(uncovered)}장 (도구·통계표·report 소관 — 정상):")
    from collections import Counter
    print("  분류 분포:", dict(Counter(c for _, c in uncovered).most_common(8)))


if __name__ == "__main__":
    main()
