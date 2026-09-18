#!/usr/bin/env python3
"""좌표 단위 대조 — 결정론 팩트가 정답지다 (V14 · V15 · V16).

**왜 따로 있나**

`verify.py` 의 V2 는 「이 값이 통계표 **어딘가에** 있는가」만 본다(집합 소속).
그래서 **같은 표 안에서 값 두 개를 맞바꾸면 그대로 통과한다.** 배너 라벨을 맞바꿔도,
차트 타입을 다른 것으로 적어도 마찬가지다.

2026-08-22 독립 검증관이 이 구멍을 실제 파일로 증명했다 — 작업본에서 셀 7개의
**텍스트만** 바꾼 변종(`MUTANT2.xlsm`)이 서식·구조·개봉 검사 45+19+60 항목과
현행 `verify.py` 를 **치명 0** 으로 통과했다. 바뀐 것은:

    F508↔G508   배너 라벨 두 개 맞바꿈 → 용도 6열 전체가 엉뚱한 라벨에 붙음
    E510↔F510   같은 행 안에서 값 두 개 맞바꿈
    B499        `전체 : 가로막대, 배너 : 표` → `전체 : 원그래프, 배너: 표`

세 검사 전부 **좌표**로 맞댄다. 「어딘가에 있는가」가 아니라 「그 자리에 있는가」다.
정답지는 사람 눈이 아니라 `tables.json`(통계표 원본)과
`slide_facts.json`(룰엔진 산출)이다.
"""

from __future__ import annotations

from typing import Any

NL = chr(10)


def _panel_offset(block: dict) -> int:
    """데이터 셀 목록에서 **그리드 0번(전체)** 이 놓인 자리.

    보통 0 이다. 순위 겹침 슬라이드만 앞에 `1순위` 칸이 하나 더 붙어 1 이 된다
    (`build_guide`: overlay 면 first=1순위, first+1=전체).
    """
    head = list(block.get("banner_labels") or [])[:1]
    grp = list(block.get("banner_groups") or [])[:1]
    first = next((str(x).strip() for x in (grp + head) if x), "")
    return 1 if first.startswith("1순위") else 0


def _src_rows(table: dict) -> dict[str, dict]:
    """행 라벨 → 통계표 행. 가이드는 `label_raw`(들여쓰기 포함)를 그대로 적는다."""
    out: dict[str, dict] = {}
    for r in table.get("rows") or []:
        for key in (r.get("label_raw"), r.get("label")):
            if key:
                out.setdefault(str(key).strip(), r)
    return out


def check_grid_coordinates(rep: Any, blocks: list[dict], tables: list[dict],
                           block_where) -> None:
    """V14 배너 순서 · V15 값 좌표."""
    by_index = {int(t["index"]): t for t in tables if t.get("index") is not None}
    n_cells = n_blocks = 0

    for b in blocks:
        if b.get("is_struct"):
            continue
        slide = b.get("slide") or {}
        srcs = [int(i) for i in (slide.get("source_tables") or [])]
        if not srcs or srcs[0] not in by_index:
            continue
        t = by_index[srcs[0]]
        banners = [x for x in (t.get("banners") or []) if x.get("in_report")]
        if not banners:
            continue

        off = _panel_offset(b)
        n_blocks += 1
        where = block_where(b)

        # ── V14 · 배너 라벨이 통계표 배너와 **순서까지** 같은가
        #    빌더는 라벨==그룹인 칸을 빈 칸으로 둔다(전체 열이 그렇다) — 그래서
        #    기대값도 같은 규칙으로 만든다.
        got = [str(x).strip() if x is not None else ""
               for x in (b.get("banner_labels") or [])]
        want = [(x["label"] if x["label"] != x["group"] else "") for x in banners[1:]]
        got_b = got[off + 1: off + 1 + len(want)]
        if len(got_b) == len(want) and got_b != want:
            k = next(i for i in range(len(want)) if got_b[i] != want[i])
            rep.add("V14", "치명", where,
                    "배너 라벨이 통계표 표%d 와 어긋납니다 (%d번째 배너)%s"
                    "      가이드: %r%s      통계표: %r%s"
                    "      → 열 하나가 밀리면 그 아래 수치가 **통째로 다른 배너**에 붙습니다"
                    % (t["index"], k + 1, NL, got_b[k], NL, want[k], NL))

        # ── V15 · 각 셀이 (행 라벨, 배너) 좌표에서 통계표와 같은가
        rows = _src_rows(t)
        for row in b["rows"]:
            lbl = row["label"].strip()
            if lbl == "Base for %":
                continue
            sr = rows.get(lbl)
            if sr is None:
                continue          # 파생 행(계·순위 병합 등) — 여기서 판정하지 않는다
            vals = sr.get("values") or []
            for k, bn in enumerate(banners):
                # off=1(순위 겹침)이면 전체 칸이 두 개라 배너는 한 칸 더 뒤에서 시작한다
                col = k + off + (1 if (off and k) else 0)
                if col >= len(row["cells"]):
                    break
                pos = bn.get("pos")
                pos = int(pos) if pos is not None else None
                src = vals[pos] if (pos is not None and pos < len(vals)) else None
                want_v = src.get("display") if isinstance(src, dict) else src
                got_v = row["cells"][col]
                if got_v is None and want_v in (None, "", "-"):
                    continue
                n_cells += 1
                if str(got_v).strip() != str(want_v).strip():
                    rep.add("V15", "치명", "본문!%d행" % row["row"],
                            "표%d `%s` × `%s` 의 값이 다릅니다%s"
                            "      가이드: %s   통계표: %s"
                            % (t["index"], lbl[:26], str(bn["label"])[:16], NL,
                               got_v, want_v))

    rep.stats["coord_cells"] = n_cells
    rep.stats["coord_blocks"] = n_blocks


def check_chart_type(rep: Any, blocks: list[dict], block_where) -> None:
    """V16 — `Chart type` 이 룰엔진이 정한 문자열과 같은가.

    이 칸은 제작팀이 눈으로 읽는 **자유 서술**이라 어휘 검사를 걸 수 없다(R-034).
    그래서 대신 룰엔진 산출(`slide_facts.chart_type`)과 1:1 로 맞댄다.
    사람이 손으로 고쳤다면 그건 룰엔진이 틀렸다는 신호이므로 `/회고` 로 올려야 한다.
    """
    n = 0
    for b in blocks:
        if b.get("is_struct"):
            continue
        want = (b.get("slide") or {}).get("chart_type")
        if not want:
            continue
        got = str(b["values"].get("Chart type") or "").strip()
        n += 1
        if got != str(want).strip():
            rep.add("V16", "치명", block_where(b),
                    "`Chart type` 이 룰엔진 판정과 다릅니다%s"
                    "      가이드: %r%s      룰엔진: %r"
                    % (NL, got, NL, str(want).strip()))
    rep.stats["chart_types_checked"] = n
