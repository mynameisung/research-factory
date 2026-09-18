#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gen_preamble.py — 스킬 파일의 PREAMBLE 블록을 원본(templates/SKILL_PREAMBLE.md)으로 채운다.

gstack 의 gen-skill-docs 와 같은 자리다. 프리앰블은 한 곳(원본)에서만 고치고, 이 도구가
`.claude/skills/*/SKILL.md` 의 `<!-- PREAMBLE:START -->` … `<!-- PREAMBLE:END -->` 사이를 갈아 끼운다.
표식이 없는 스킬(업그레이드·개선공유)은 건드리지 않는다 — 그 둘은 판을 올리거나 보내는 스킬이라 판 확인이 없다.

사용: python gen_preamble.py            채운다 (바뀐 파일만 쓴다)
      python gen_preamble.py --check    원본과 어긋난 스킬이 있으면 목록을 내고 exit 1 (setup 스모크 · /개선공유 선행 확인)
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]                       # <repo>/research-factory/engine/tools → <repo>
TEMPLATE = HERE.parent / "templates" / "SKILL_PREAMBLE.md"
SKILLS = ROOT / ".claude" / "skills"
START, END = "<!-- PREAMBLE:START -->", "<!-- PREAMBLE:END -->"
RE_NAME = re.compile(r"^name:\s*(.+?)\s*$", re.M)


def render(skill: str) -> str:
    body = TEMPLATE.read_text(encoding="utf-8").rstrip("\n")
    return f"{START}\n{body.replace('{{SKILL}}', skill)}\n{END}"


def main(argv: list[str]) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    check = "--check" in argv
    stale, filled, skipped = [], [], []
    for f in sorted(SKILLS.glob("*/SKILL.md")):
        text = f.read_text(encoding="utf-8")
        if START not in text or END not in text:
            skipped.append(f.parent.name); continue
        m = RE_NAME.search(text.split("---", 2)[1] if text.startswith("---") else "")
        skill = m.group(1) if m else f.parent.name
        s, e = text.index(START), text.index(END) + len(END)
        new = text[:s] + render(skill) + text[e:]
        if new != text:
            stale.append(f.parent.name)
            if not check:
                f.write_text(new, encoding="utf-8"); filled.append(f.parent.name)
    if check:
        if stale:
            print("PREAMBLE 이 원본과 다른 스킬: " + " · ".join(stale) + "  → python engine/tools/gen_preamble.py")
            return 1
        print(f"PREAMBLE 일치 (스킬 {len(list(SKILLS.glob('*/SKILL.md'))) - len(skipped)}개 · 표식 없음 {' · '.join(skipped) or '없음'})")
        return 0
    print(f"채움 {len(filled)}개 ({' · '.join(filled) or '변경 없음'}) · 표식 없음 {' · '.join(skipped) or '없음'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
