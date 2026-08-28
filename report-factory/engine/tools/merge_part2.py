# -*- coding: utf-8 -*-
"""
merge_part2.py — summary_map.json + conclusion_map.json → 목차.md 의 Part Ⅱ 절(첫 펜스 Part Ⅱ 항목 · `### 요약 꼭지` · `### 결론 및 제언`)을 다시 쓴다.

  python engine/tools/merge_part2.py <work_dir> [--summary summary_map.json] [--conclusion conclusion_map.json] [--toc 목차.md]

지금까지 메인이 손으로 하던 병합을 고정 포맷으로 만든다(parse_toc / plan_summary / conclusion_layout 이 읽는 형식 그대로):
  **{no}. {title}** ← {refs}
  ```
  [{position}]   {ref} {topic}      {chart} · Base {base_n}
  ```
  - {headline}
  {body 줄…}
결론:
  레이아웃: {layout}
  **{no}. {title}** ← 요약 {n} · 요약 {m}
  {body 줄…}
"""
from __future__ import annotations
import argparse, json, re
from pathlib import Path


def summary_section(smap: dict) -> str:
    out = ["### 요약 꼭지", ""]
    for s in smap["summaries"]:
        refs = " · ".join(p["ref"] for p in s["panels"])
        out.append(f"**{s['no']}. {s['title']}** ← {refs}")
        out.append("```")
        for p in s["panels"]:
            base = f" · Base {p['base_n']}" if p.get("base_n") is not None else ""
            out.append(f"[{p['position']}]   {p['ref']} {p['topic']}      {p['chart']}{base}")
        out.append("```")
        head = str(s.get("headline") or "").strip()
        out.append(head if head.startswith("- ") else f"- {head}")
        for b in s.get("body") or []:
            out.append(str(b).strip())
        out.append("")
    return "\n".join(out)


def conclusion_section(cmap: dict) -> str:
    out = ["### 결론 및 제언", ""]
    lay = str(cmap.get("layout") or "").strip()
    if lay:
        out.append(f"레이아웃: {lay}")
        out.append("")
    for it in cmap["items"]:
        src = " · ".join(f"요약 {n}" for n in (it.get("from_summary") or []))
        out.append(f"**{it['no']}. {it['title']}**" + (f" ← {src}" if src else ""))
        for b in it.get("body") or []:
            out.append(str(b).strip())
        out.append("")
    return "\n".join(out)


def part2_fence_line(smap: dict, cmap: dict) -> str:
    items = [f"{s['no']}. {s['title']}" for s in smap["summaries"]]
    items.append(f"{len(items) + 1}. 결론 및 제언")
    return "    " + "    ".join(items)


def merge(toc_text: str, smap: dict, cmap: dict) -> str:
    # ① 첫 펜스의 Part Ⅱ 항목 줄
    m = re.search(r"(Part Ⅱ\. 요약 및 결론\n)(    .*\n)?", toc_text)
    if m:
        toc_text = toc_text[:m.start()] + m.group(1) + part2_fence_line(smap, cmap) + "\n" + toc_text[m.end():]
    # ② 요약 꼭지 · 결론 절 교체 (Part Ⅱ 헤더 아래 ### 절 두 개)
    def replace_section(text: str, head: str, body: str) -> str:
        i = text.find(head)
        if i < 0:
            return text.rstrip("\n") + "\n\n" + body + "\n"
        j = text.find("\n### ", i + len(head))
        k = text.find("\n## ", i + len(head))
        end = min(x for x in (j, k, len(text)) if x >= 0)
        return text[:i] + body.rstrip("\n") + "\n" + text[end:]
    toc_text = replace_section(toc_text, "### 요약 꼭지", summary_section(smap))
    toc_text = replace_section(toc_text, "### 결론 및 제언", conclusion_section(cmap))
    return toc_text


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("work", type=Path)
    ap.add_argument("--summary", type=Path)
    ap.add_argument("--conclusion", type=Path)
    ap.add_argument("--toc", type=Path)
    ap.add_argument("--out", type=Path, help="기본: --toc 자체를 덮어쓴다")
    a = ap.parse_args()
    W = a.work
    smap = json.loads((a.summary or W / "summary_map.json").read_text(encoding="utf-8"))
    cmap = json.loads((a.conclusion or W / "conclusion_map.json").read_text(encoding="utf-8"))
    toc = a.toc or W / "목차.md"
    text = merge(toc.read_text(encoding="utf-8"), smap, cmap)
    (a.out or toc).write_text(text, encoding="utf-8", newline="\n")
    print(f"→ {a.out or toc}  (요약 {len(smap['summaries'])}꼭지 · 결론 {len(cmap['items'])}박스)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
