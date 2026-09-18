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
    # 2026-09-10 신설: /보고서작성 이 RULES.md 전문(308KB) 대신 읽는 팩. 분류 report/* 와 R-PPT-* 카드.
    "report": ["report/"],
}

#: 분류 태그 경계와 안 맞는 카드의 명시 등재 (guide-judge 심사 F-3·F-4 반영).
#: 태그를 고치는 게 아니라 여기 더한다 — 카드 분류는 원저자 표기 그대로 둔다.
EXTRA_CARDS = {
    # 배너 설계는 guide-toc 임무다 (guide/18 §2 — 목차·배너 쌍 산출).
    # R-138·140·142·195 는 감사 F-4(2026-09-02)가 배달 결손으로 잡은 목차 카드.
    "목차": ["R-055", "R-056", "R-057", "R-089", "R-138", "R-140", "R-142", "R-195"],
    # guide-body-ment 가 정의에서 직접 인용하는 카드 (감사 F-5 — 팩에 0장이었다)
    "멘트": ["R-088", "R-139", "R-185"],
    # guide-front (Part Ⅰ + Cover) 가 실제로 딛는 카드들
    "조사개요": ["R-011", "R-013", "R-022", "R-046", "R-124", "R-127"],
    # 카드 포맷(분류·근거·신뢰도)을 안 따르는 산문 카드와, 요약 담당이 실제로 딛는 경계 카드.
    # 태그가 없다는 이유로 빠지면 「설계의 닻은 조사 목적」(R-141) 같은 1급 규칙이 배달되지 않는다.
    # 2026-09-10: R-128(저베이스 별표 — PPT 소관으로 정정) 대신 R-188(가이드에는 넣지 않는다)·R-135(계 행 제외)를 싣는다 (감사 A-1 · F-6).
    "요약결론": ["R-119", "R-121", "R-126", "R-135", "R-141", "R-142", "R-143",
              "R-144", "R-145", "R-146", "R-188", "R-215", "R-216", "R-217"],
}

#: 팩별로 담당이 함께 읽어야 하는 knowledge 문서 (포인터 — 내용 복제 금지)
PACK_DOCS = {
    "목차": ["knowledge/guide/10_목차-설계규칙.md"],
    "멘트": ["knowledge/guide/13_멘트-작성규칙.md"],
    "요약결론": ["knowledge/guide/14_요약결론-작성규칙.md"],
    "조사개요": ["knowledge/00_업무개요.md"],
    "공통": ["knowledge/guide/17_팀구조-제작가이드.md",
             "knowledge/guide/18_오케스트레이션-제작가이드.md"],
    "report": ["knowledge/report/23_팀구조-PPT제작.md",
               "knowledge/report/22_검수-체크리스트.md",
               "knowledge/report/25_차트-카탈로그.md"],
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

    def _excluded(card):
        # 2026-09-10 규칙 다이어트: 코퍼스 관찰로 강등된 카드(【지식)와 대표 카드에 흡수된 카드(【흡수)는
        # 팩에 싣지 않는다. 관찰은 knowledge 부록에, 흡수분은 대표 카드의 「흡수」 줄에 있다.
        # 꼬리에 번호만 남겨 담당이 「그 카드는 어디 갔나」를 알 수 있게 한다.
        t = card[1] or ""
        return "【지식" in t or "【흡수" in t

    def _render(rid, title, cat, trust, body):
        # 카드에 「- 요지:」 줄이 있으면 팩에는 요지·적용·정정·흡수 줄만 싣고 본문은 원본 포인터로 남긴다.
        # 담당이 읽는 양을 줄이는 자리다 — 예시·실측 서사·이력은 RULES.md 에서 필요할 때 읽는다.
        m = re.search(r"^- 요지:\s*(.+?)\s*$", body, re.M)
        if not m:
            return body
        keep = [ln for ln in body.splitlines()
                if re.match(r"^- (요지|적용|정정|흡수|예외|신뢰도|분류):", ln)]
        keep.append(f"- 전문: `engine/learning/RULES.md` {rid}")
        return "\n".join(keep)

    for pack, tags in PACK_MAP.items():
        keys = [norm(t) for t in tags]
        def _hit(cat):
            # guide 계열은 키워드 부분일치, 공통 계열은 접두로 가른다.
            return bool(cat) and any(k in cat for k in keys)
        picked = [c for c in cards if _hit(c[2])]
        if pack == "report":
            picked += [c for c in cards if c[0].startswith("R-PPT-") and c not in picked]
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
        dropped = [c for c in picked if _excluded(c)]
        picked = [c for c in picked if not _excluded(c)]
        used.update(c[0] for c in picked)
        used.update(c[0] for c in dropped)
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
            lines.append(_render(rid, title, cat, trust, body))
            lines.append("")
        if dropped:
            lines.append("---")
            lines.append("이 팩에서 뺀 카드 (코퍼스 관찰로 강등 · 대표 카드에 흡수 — 본문은 RULES.md 와 knowledge 부록): "
                         + " · ".join(f"{c[0]}" for c in dropped))
            lines.append("")
        (OUT / f"{pack}.md").write_text("\n".join(lines), encoding="utf-8", newline="\n")
        n_slim = sum(1 for c in picked if re.search(r"^- 요지:", c[4], re.M))
        n_chars = sum(len(_render(*c)) for c in picked)
        report.append(f"  {pack:<6} 카드 {len(picked):>3}장 (요지형 {n_slim:>3}) 배달 {n_chars:>7,}자 · 뺀 카드 {len(dropped)}장")

    # 카드 제목 색인 — /회고 · guide-judge · retro-miner 가 RULES.md 전문(300KB) 대신 읽는다.
    # 한 줄 = 번호 · 상태(폐기/지식/흡수/정정) · 분류 · 제목. 본문이 필요하면 번호로 RULES.md 를 grep 한다.
    idx = ["# 규칙 카드 색인 (make_packs.py 자동 생성 — 고치지 말 것)", "",
           f"카드 {len(cards)}장. 상태 표기: 【폐기】 대체됨 · 【지식】 코퍼스 관찰 부록으로 · 【흡수】 대표 카드에 합쳐짐 · (정정) 조항 정정 있음", "",
           "| 번호 | 상태 | 분류 | 제목 |", "|---|---|---|---|"]
    for rid, title, cat, trust, body in cards:
        state = "폐기" if _deprecated((rid, title, cat, trust, body)) else                 "지식" if "【지식" in (title or "") else                 "흡수" if "【흡수" in (title or "") else                 "정정" if re.search(r"^- 정정 \(", body, re.M) else ""
        clean = re.sub(r"^【[^】]*】\s*", "", title or "")
        idx.append(f"| {rid} | {state} | {cat or ''} | {clean[:80]} |")
    (OUT / "_색인.md").write_text("\n".join(idx) + "\n", encoding="utf-8", newline="\n")
    print(f"팩 {len(PACK_MAP)}종 + 색인 생성 → {OUT}")
    print("\n".join(report))
    uncovered = [(c[0], c[2]) for c in cards if c[0] not in used]
    print(f"어느 팩에도 안 실린 카드 {len(uncovered)}장 (도구·통계표·report 소관 — 정상):")
    from collections import Counter
    print("  분류 분포:", dict(Counter(c for _, c in uncovered).most_common(8)))


if __name__ == "__main__":
    main()
