# -*- coding: utf-8 -*-
"""
diff_decks.py -- 제작본 ↔ 최종본 정답지 대조기

  ★ 왜 만드나 ★
  제작본은 우리가 만들 수 있는 것이고, 최종본은 **연구원이 손을 댄 뒤**의 것이다.
  그 차이가 곧 「제작본이 아직 못 맞춘 것」의 목록이다 — 반복되는 수정 유형은
  렌더러가 미리 맞춰야 할 규칙이고, 한 번뿐인 수정은 그 조사의 사정이다.
  그래서 **차이의 종류를 세는 것**이 이 도구의 일이다. 개별 차이가 아니라 분포를 본다.

  짝짓기는 `join_key`(제목+소제목 정규화)로 한다. 제목을 고친 장은 짝이 깨지므로
  2차로 **차트·표 구성 지문**(fingerprint)으로 다시 붙인다 — 제목만 다듬은 장이
  「삭제 + 추가」로 세어지면 수정 유형 통계가 망가진다.

사용법:
    python diff_decks.py <제작본.json> <최종본.json> [out.json]
"""

from __future__ import annotations

import json
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any


def norm(s: str | None) -> str:
    """공백·문장부호를 지운 비교용 키. 줄바꿈 위치만 다른 문장을 같게 본다."""
    return re.sub(r"[\s​·,.\-—–…()\[\]{}'\"]+", "", str(s or ""))


def fingerprint(sl: dict) -> str:
    """제목을 빼고 **그림의 구성**만으로 만든 지문.

    제목·소제목을 고친 장을 다시 붙이는 데 쓴다. 차트 종류와 계열·범주 수,
    표 크기까지 같으면 같은 장으로 본다 — 그 정도면 우연이 아니다.
    """
    parts = [sl.get("derived_chart_type") or "", str(sl.get("chart_count") or 0),
             str(sl.get("table_count") or 0), sl.get("classification") or ""]
    for ch in (sl.get("charts") or []):
        parts.append("%s/%s/%s" % (ch.get("chart_type"), len(ch.get("series") or []),
                                   len(ch.get("categories") or [])))
    for tb in (sl.get("tables") or []):
        parts.append("T%sx%s" % (tb.get("row_count"), tb.get("col_count")))
    return "|".join(parts)


def slide_text(sl: dict) -> str:
    """장표가 말하는 문장 전부. 문구 수정을 잡는 데 쓴다."""
    bits = [sl.get("slide_title"), sl.get("sub_title"), sl.get("top_ment"), sl.get("qu_text")]
    bits += [t.get("text") for t in (sl.get("text_shapes") or [])]
    return norm("\n".join(b for b in bits if b))


def chart_sig(sl: dict) -> list[str]:
    out = []
    for ch in (sl.get("charts") or []):
        out.append("%s(계열%d·범주%d)" % (ch.get("chart_type"),
                                        len(ch.get("series") or []),
                                        len(ch.get("categories") or [])))
    return out


def pair_slides(a: list[dict], b: list[dict]) -> tuple[list[tuple], list[dict], list[dict]]:
    """(짝, 제작본에만, 최종본에만). 1차 join_key · 2차 지문."""
    used_b: set[int] = set()
    pairs: list[tuple] = []
    by_key: dict[str, list[int]] = {}
    for j, s in enumerate(b):
        by_key.setdefault(norm(s.get("join_key")), []).append(j)
    left: list[dict] = []
    for s in a:
        cand = by_key.get(norm(s.get("join_key")) or "\x00", [])
        hit = next((j for j in cand if j not in used_b), None)
        if hit is None:
            left.append(s)
        else:
            used_b.add(hit)
            pairs.append((s, b[hit], "제목"))
    # 2차 — 지문으로 다시 붙인다
    by_fp: dict[str, list[int]] = {}
    for j, s in enumerate(b):
        if j not in used_b:
            by_fp.setdefault(fingerprint(s), []).append(j)
    only_a: list[dict] = []
    for s in left:
        cand = by_fp.get(fingerprint(s), [])
        hit = next((j for j in cand if j not in used_b), None)
        if hit is None:
            only_a.append(s)
        else:
            used_b.add(hit)
            pairs.append((s, b[hit], "지문"))
    only_b = [s for j, s in enumerate(b) if j not in used_b]
    pairs.sort(key=lambda p: p[0].get("slide_no") or 0)
    return pairs, only_a, only_b


#: 차이 종류. 값은 사람이 읽는 이름 — 리포트에 그대로 쓴다.
KINDS = {
    "title": "제목 문구",
    "subtitle": "소제목·헤드라인 문구",
    "ment": "멘트 문구",
    "chart_type": "차트 종류 교체",
    "chart_shape": "계열·범주 수",
    "chart_count": "차트 개수",
    "table_count": "표 개수",
    "base": "Base·Unit 표기",
    "qu": "Qu 문구",
    "memo": "제작 메모 제거",
    "order": "장 순서",
    "text": "그 밖 문구",
}


def diff_pair(x: dict, y: dict) -> list[dict]:
    out = []

    def add(kind, before, after):
        out.append({"kind": kind, "kind_name": KINDS[kind],
                    "before": (str(before) or "")[:300], "after": (str(after) or "")[:300]})

    if norm(x.get("slide_title")) != norm(y.get("slide_title")):
        add("title", x.get("slide_title"), y.get("slide_title"))
    if norm(x.get("sub_title")) != norm(y.get("sub_title")):
        add("subtitle", x.get("sub_title"), y.get("sub_title"))
    if norm(x.get("top_ment")) != norm(y.get("top_ment")):
        add("ment", x.get("top_ment"), y.get("top_ment"))
    if (x.get("derived_chart_type") or "") != (y.get("derived_chart_type") or ""):
        add("chart_type", x.get("derived_chart_type"), y.get("derived_chart_type"))
    if (x.get("chart_count") or 0) != (y.get("chart_count") or 0):
        add("chart_count", x.get("chart_count"), y.get("chart_count"))
    elif chart_sig(x) != chart_sig(y):
        add("chart_shape", " · ".join(chart_sig(x)), " · ".join(chart_sig(y)))
    if (x.get("table_count") or 0) != (y.get("table_count") or 0):
        add("table_count", x.get("table_count"), y.get("table_count"))
    if norm(x.get("base_unit")) != norm(y.get("base_unit")):
        add("base", x.get("base_unit"), y.get("base_unit"))
    if norm(x.get("qu_text")) != norm(y.get("qu_text")):
        add("qu", x.get("qu_text"), y.get("qu_text"))
    if (x.get("production_memo_count") or 0) != (y.get("production_memo_count") or 0):
        add("memo", x.get("production_memo_count"), y.get("production_memo_count"))
    if not out and slide_text(x) != slide_text(y):
        add("text", "(본문 문구 차이)", "(본문 문구 차이)")
    return out


def is_unfinished(sl: dict) -> bool:
    """제작본의 그 장이 **아직 안 채워진 장**인가.

    ★ 차이를 한 덩어리로 세면 안 된다. 「Data 입력」 메모가 붙은 반복 장표는
      제작팀이 일부러 틀만 준 것이고, 최종본과의 차이는 **연구원이 채운 결과**다.
      우리는 늘 다 채워서 내므로 그건 렌더러가 배울 규칙이 아니다.
      완성된 장을 연구원이 고친 것만이 「제작본이 아직 못 맞춘 것」이다.
    """
    if sl.get("production_memo_count"):
        return True
    if sl.get("unfilled_markers"):
        return True
    # 본문인데 차트도 표도 없다 = 틀만 있는 장
    if sl.get("classification") == "본문" and not (sl.get("chart_count") or sl.get("table_count")):
        return True
    return False


def compare(a_path: Path, b_path: Path) -> dict:
    A = json.loads(a_path.read_text(encoding="utf-8"))
    B = json.loads(b_path.read_text(encoding="utf-8"))
    pairs, only_a, only_b = pair_slides(A["slides"], B["slides"])

    changed = []
    kinds_done, kinds_todo = Counter(), Counter()
    moved = n_unfinished = 0
    for x, y, how in pairs:
        ds = diff_pair(x, y)
        if (x.get("slide_no") or 0) != (y.get("slide_no") or 0):
            moved += 1
        unfin = is_unfinished(x)
        n_unfinished += bool(unfin)
        if ds:
            for d in ds:
                (kinds_todo if unfin else kinds_done)[d["kind"]] += 1
            changed.append({"제작본": x.get("slide_no"), "최종본": y.get("slide_no"),
                            "짝": how, "분류": x.get("classification"),
                            "미완성": unfin,
                            "제목": x.get("slide_title"), "차이": ds})

    return {
        "제작본": A["source_file"], "최종본": B["source_file"],
        "장수": {"제작본": A["slide_count"], "최종본": B["slide_count"]},
        "판형": {"제작본": [A.get("slide_width_cm"), A.get("slide_height_cm")],
                 "최종본": [B.get("slide_width_cm"), B.get("slide_height_cm")]},
        "짝": len(pairs), "제작본에만": len(only_a), "최종본에만": len(only_b),
        "자리이동": moved,
        "손댄 장": len(changed),
        "제작본 미완성 장": n_unfinished,
        "차이 종류 · 완성장(렌더러가 배울 것)": {KINDS[k]: n for k, n in kinds_done.most_common()},
        "차이 종류 · 미완성장(채운 결과)": {KINDS[k]: n for k, n in kinds_todo.most_common()},
        "분류별 손댄 장": dict(Counter(c["분류"] for c in changed).most_common()),
        "최종본에만_분류": dict(Counter(s.get("classification") for s in only_b).most_common()),
        "제작본에만_분류": dict(Counter(s.get("classification") for s in only_a).most_common()),
        "changed": changed,
        "only_a": [{"no": s.get("slide_no"), "분류": s.get("classification"),
                    "제목": s.get("slide_title")} for s in only_a],
        "only_b": [{"no": s.get("slide_no"), "분류": s.get("classification"),
                    "제목": s.get("slide_title")} for s in only_b],
    }


def main(argv: list[str]) -> int:
    if len(argv) < 3:
        print(__doc__)
        return 2
    res = compare(Path(argv[1]), Path(argv[2]))
    if len(argv) > 3:
        Path(argv[3]).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    brief = {k: v for k, v in res.items() if k not in ("changed", "only_a", "only_b")}
    print(json.dumps(brief, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
