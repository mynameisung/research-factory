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

#: 팩 → 포함할 분류 태그 (RULES 카드의 `분류:` 값, 공백 정규화 후 비교)
#: 한 카드가 여러 팩에 실릴 수 있다 (예: 목차·요약결론 경계 카드)
PACK_MAP = {
    "목차": ["guide/목차", "guide/목차·구조", "guide/목차·요약결론"],
    "멘트": ["guide/멘트", "guide/행수식어", "guide/표기"],
    "요약결론": ["guide/요약결론", "guide/목차·요약결론"],
    "조사개요": [],                       # 분류 태그가 없어 EXTRA_CARDS 로만 채운다
    "공통": ["공통/프로세스", "공통/설계원칙", "공통/용어", "guide/산출물스키마"],
}

#: 분류 태그 경계와 안 맞는 카드의 명시 등재 (guide-judge 심사 F-3·F-4 반영).
#: 태그를 고치는 게 아니라 여기 더한다 — 카드 분류는 원저자 표기 그대로 둔다.
EXTRA_CARDS = {
    # 배너 설계는 guide-toc 임무다 (guide/18 §2 — 목차·배너 쌍 산출)
    "목차": ["R-055", "R-056", "R-057", "R-089"],
    # guide-front (Part Ⅰ + Cover) 가 실제로 딛는 카드들
    "조사개요": ["R-011", "R-013", "R-022", "R-046", "R-124", "R-127"],
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
    for m in re.finditer(r"^### (R-\d+) · (.+?)$\n(.*?)(?=^### R-\d+|\Z)",
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
    for pack, tags in PACK_MAP.items():
        tags_n = {norm(t) for t in tags}
        picked = [c for c in cards if c[2] in tags_n]
        for rid in EXTRA_CARDS.get(pack, []):
            if rid not in by_id:
                raise SystemExit(f"EXTRA_CARDS 의 {rid} 가 RULES.md 에 없다 — 오타이거나 폐기됨")
            if by_id[rid] not in picked:
                picked.append(by_id[rid])
        picked.sort(key=lambda c: int(c[0][2:]))
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
