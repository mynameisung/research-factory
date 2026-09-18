#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""구조화 엑셀(.xlsx) → structure.json + structure.md (사람·Claude 가 읽는 펼침본).

    python read_structure.py <구조화.xlsx> <출력폴더>

구조화 파일은 조사마다 열 구성이 다르다(대분류/중분류/소분류/문번/문항/보기/응답 로직 …, NO/PART/주제/Details/구성/BASE …,
파트/질문/문번/문항/보기/응답대상/비고 …). 이 도구는
    · 시트마다 머리 행을 찾고(알려진 머리말이 3개 이상인 첫 행) 열 역할을 붙인다
    · 병합 셀 값을 병합 범위 전체로 채운다 (파트·중분류가 세로 병합인 경우 · 머리말이 가로 병합인 경우)
    · 같은 머리말 열(`보기 | 보기 | 보기`)은 같은 역할로 보고 값을 ` | ` 로 잇는다
    · 빈 행 뒤 번호·파트 없이 한 칸뿐인 줄은 「본표 아래 메모」로 뺀다 (둘 이상 차면 본표가 이어진 것)
    · 머리 행이 없는 시트(메일 본문·이전 버전 초안·보기 초안 등)는 「참고 시트」로 원문 줄만 남긴다
    · 셀 배경색·글자색을 남긴다 (노랑 = 핵심 지표 · 빨강 = 삭제/보류 표시인 경우가 많다 — 뜻은 사람에게 확인)
문장을 해석하지 않는다. 행마다 `ref`(시트!행) 를 달아 설문 명세의 `src` 가 이것을 가리킨다.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

# 열 역할 ← 머리말 (정규화: 공백·줄바꿈·괄호 내용 제거 후 소문자)
ROLE_WORDS: list[tuple[str, tuple[str, ...]]] = [
    ("seq", ("no", "번호", "순번", "연번")),
    ("code", ("문번", "문항번호", "문항no", "qno", "q번호")),
    ("part", ("대분류", "파트", "part", "영역")),
    ("section", ("중분류", "차원", "주제", "질문", "구분항목")),
    ("attribute", ("소분류", "속성", "세부항목", "항목")),
    ("question", ("문항", "details", "질문내용", "문항내용", "설문문항")),
    ("options", ("보기", "응답보기", "보기안")),
    ("format", ("구성", "응답형태", "응답유형", "응답조건", "척도")),
    ("logic", ("응답로직", "로직", "prog", "분기")),
    ("base", ("응답대상", "base", "대상", "베이스")),
    ("note", ("비고", "비고ref", "참고", "메모", "rfp상내용", "출처", "구분", "코멘트")),
]
KNOWN = {w for _, ws in ROLE_WORDS for w in ws}


def _norm(s: Any) -> str:
    s = re.sub(r"\(.*?\)", "", str(s or ""))
    return re.sub(r"[\s_\-./]+", "", s).lower()


def _role(header: str) -> str | None:
    h = _norm(header)
    if not h:
        return None
    for role, words in ROLE_WORDS:
        if h in words:
            return role
    for role, words in ROLE_WORDS:          # 앞부분 일치 (예: '응답 로직(필터)')
        if any(h.startswith(w) and len(w) >= 2 for w in words):
            return role
    return None


def _color(col) -> str | None:
    """openpyxl 색 → `FFFF00` · 테마색은 `theme5+0.4` (색표 위 격자) · 인덱스색(.xls 변환본)은 RGB 로."""
    try:
        if col is None:
            return None
        if col.type == "theme":
            tint = round(col.tint or 0, 1)
            return f"theme{col.theme}" + (f"{tint:+.1f}" if tint else "")
        if col.type == "indexed":
            from openpyxl.styles.colors import COLOR_INDEX
            rgb = COLOR_INDEX[col.indexed] if col.indexed < len(COLOR_INDEX) else None   # 64·65 = 시스템 기본
        else:
            rgb = col.rgb
        if isinstance(rgb, str) and rgb != "00000000":
            return rgb[-6:]
    except Exception:
        return None
    return None


def _fill(cell) -> str | None:
    """배경색 + 글자색(`글자FF0000`). 흰 배경은 빼고, 글자는 검정·회색(테마 기본 글자색 포함)을 뺀다."""
    out = []
    try:
        f = cell.fill
        if f is not None and f.fill_type not in (None, "none"):
            bg = _color(f.fgColor)
            if bg not in (None, "FFFFFF", "theme0"):
                out.append(bg)
    except Exception:
        pass
    try:
        fg = _color(cell.font.color) if cell.font is not None else None
        if fg and not re.fullmatch(r"theme[01](?:[+-][\d.]+)?|([0-9A-Fa-f]{2})\1\1", fg):
            out.append(f"글자{fg}")
    except Exception:
        pass
    return "/".join(out) or None


def read(path: Path) -> dict[str, Any]:
    import openpyxl
    wb = openpyxl.load_workbook(str(path), data_only=True)
    sheets: list[dict[str, Any]] = []
    for ws in wb.worksheets:
        # 병합 셀 채움 지도
        merged: dict[tuple[int, int], tuple[int, int]] = {}
        for rng in ws.merged_cells.ranges:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    merged[(r, c)] = (rng.min_row, rng.min_col)

        def val(r: int, c: int):
            rr, cc = merged.get((r, c), (r, c))
            v = ws.cell(rr, cc).value
            return None if v is None or (isinstance(v, str) and not v.strip()) else v

        max_r, max_c = ws.max_row, ws.max_column
        header_row, roles = None, {}
        for r in range(1, min(max_r, 15) + 1):
            hits = {}
            for c in range(1, max_c + 1):
                v = ws.cell(r, c).value
                role = _role(v) if isinstance(v, str) else None
                if role and role not in hits.values():
                    hits[c] = role
            vals = set(hits.values())
            # 문항 열 + 문항을 규정하는 열(문번·보기·구성·대상·로직·주제)이 있어야 구조화다.
            # `파트 | 문항 수 | 비고` 같은 집계 시트는 참고 시트로 둔다.
            defining = vals & {"code", "options", "format", "base", "logic", "section", "attribute"}
            if len(hits) >= 3 and ("question" in vals or "section" in vals) and defining:
                header_row, roles = r, hits
                break

        sh: dict[str, Any] = {"name": ws.title, "kind": "structure" if header_row else "reference",
                              "header_row": header_row, "columns": {}, "rows": [], "lines": []}
        if not header_row:
            for r in range(1, max_r + 1):
                cells = [str(val(r, c)).strip() for c in range(1, max_c + 1) if val(r, c) is not None]
                if cells:
                    sh["lines"].append({"ref": f"{ws.title}!{r}", "text": " | ".join(cells)})
            sheets.append(sh)
            continue

        # 머리말 — 가로 병합 머리(`중분류` B:C · `보기` C:E)의 나머지 칸도 같은 머리말이다.
        # 같은 머리말이 여러 열이면 `보기#2` 로 이름을 가르고 역할도 같이 준다 (값은 ` | ` 로 잇는다).
        headers: dict[int, str] = {}
        base_name: dict[int, str] = {}
        count: dict[str, int] = {}
        for c in range(1, max_c + 1):
            v = val(header_row, c)
            if v is None:
                continue
            h = re.sub(r"\s+", " ", str(v)).strip()
            count[h] = count.get(h, 0) + 1
            headers[c] = h if count[h] == 1 else f"{h}#{count[h]}"
            base_name[c] = h
        role_by_name = {base_name[c]: rl for c, rl in roles.items() if c in base_name}
        for c in headers:
            if c not in roles and base_name[c] in role_by_name:
                roles[c] = role_by_name[base_name[c]]
        # 본표 = 첫 역할 열부터 머리말이 끊기기 전까지. 빈 머리 열 뒤의 열은 곁표(보기 목록 등)로 뺀다.
        # 단 문항 열(없으면 주제 열)까지는 빈 머리 칸이 있어도 본표다.
        first = min(roles)
        qcol = min((c for c, rl in roles.items() if rl == "question"), default=None) \
            or min(c for c, rl in roles.items() if rl == "section")
        main_end = first
        while main_end + 1 <= max_c and ((main_end + 1) in headers or main_end + 1 <= qcol):
            main_end += 1
        side_cols = [c for c in headers if c > main_end]
        sh["columns"] = {headers[c]: roles.get(c) for c in headers if c <= main_end}
        if side_cols:
            sh["side_columns"] = [headers[c] for c in side_cols]
        side = []                                   # 곁표 줄
        below = []                                  # 본표 아래(빈 행 뒤) 메모 — 대안 문구·코멘트
        seen_data, gap = False, False
        anchor_roles = {c for c, rl in roles.items() if rl in ("seq", "code", "part") and c <= main_end}
        body_roles = {c for c, rl in roles.items() if rl not in ("seq", "code", "part", "note") and c <= main_end}
        for r in range(header_row + 1, max_r + 1):
            rec: dict[str, Any] = {"ref": f"{ws.title}!{r}", "fields": {}, "extra": {}, "fills": {}}
            own = False                             # 병합 채움이 아닌 자기 값이 있는가
            side_vals = [str(val(r, c)).strip() for c in range(main_end + 1, max_c + 1) if val(r, c) is not None]
            if side_vals:
                side.append({"ref": rec["ref"], "text": " | ".join(side_vals)})
            main_vals = [str(val(r, c)).strip() for c in range(1, main_end + 1) if val(r, c) is not None]
            if not main_vals:
                gap = gap or seen_data
                continue
            if gap:
                # 빈 행 뒤라도 번호·파트(병합 채움 포함)가 있거나 역할 칸이 둘 이상 차면 본표가 이어진 것이다.
                # 역할 칸 하나(대안 문구 한 줄)뿐인 줄만 본표 아래 메모로 뺀다.
                if (any(val(r, c) is not None for c in anchor_roles)
                        or sum(val(r, c) is not None for c in body_roles) >= 2):
                    gap = False
                else:
                    below.append({"ref": rec["ref"], "text": " | ".join(main_vals)})
                    continue
            seen_data = True
            for c in range(1, main_end + 1):
                v = val(r, c)
                if v is None:
                    continue
                anchor = merged.get((r, c))
                if anchor and anchor[0] == r and anchor != (r, c):
                    continue                        # 가로 병합의 나머지 칸 — 첫 칸에서만 읽는다
                if anchor and anchor[0] == r:
                    span = [cc for cc in range(c, main_end + 1) if merged.get((r, cc)) == anchor]
                    if len(span) >= 3:              # 여러 열을 가로지르는 안내 띠
                        rec["banner"] = str(v).strip()
                        own = True
                        continue
                name = headers.get(c)
                role = roles.get(c)
                text = str(v).strip()
                if (r, c) not in merged or merged[(r, c)] == (r, c):
                    if role not in ("part", "section") and name is not None:
                        own = True
                fl = _fill(ws.cell(r, c))
                if fl:
                    rec["fills"][name or f"col{c}"] = fl
                if role:
                    rec["fields"][role] = f"{rec['fields'][role]} | {text}" if role in rec["fields"] else text
                elif name:
                    rec["extra"][name] = text
                else:
                    rec["extra"][f"col{c}"] = text
            if not own and not rec["fields"].get("question"):
                continue
            sh["rows"].append(rec)
        if side:
            sh["side_lines"] = side
        if below:
            sh["below_lines"] = below
        sheets.append(sh)
    return {"source": str(path.resolve()), "sheets": sheets}


def to_markdown(doc: dict[str, Any]) -> str:
    out = [f"# 구조화 펼침본 — {Path(doc['source']).name}", ""]
    for sh in doc["sheets"]:
        if sh["kind"] == "structure":
            out.append(f"## 시트 `{sh['name']}` — 구조화 (머리 {sh['header_row']}행)")
            out.append("열 역할: " + " · ".join(f"{k}→{v or '?'}" for k, v in sh["columns"].items()))
            out.append("")
            order = ["seq", "part", "section", "attribute", "code", "question", "options", "format", "logic", "base", "note"]
            for rec in sh["rows"]:
                f = rec["fields"]
                bits = [f"**{rec['ref']}**"] + ([f"띠=「{rec['banner']}」"] if rec.get("banner") else [])
                for k in order:
                    if f.get(k):
                        bits.append(f"{k}=「{f[k]}」")
                for k, v in rec["extra"].items():
                    bits.append(f"{k}=「{v}」")
                if rec["fills"]:
                    bits.append("색=" + ",".join(f"{k}:{v}" for k, v in rec["fills"].items()))
                out.append("- " + " ".join(bits).replace("\n", " ⏎ "))
            if sh.get("side_lines"):
                out.append("")
                heads = " · ".join(sh.get("side_columns") or []) or "머리말 없음"
                out.append(f"곁표 · 열 밖 메모 (본표 오른쪽 — {heads}). 보기 목록이거나 수정 지시다:")
                out += [f"- {x['ref']}: {x['text']}".replace("\n", " ⏎ ") for x in sh["side_lines"]]
            if sh.get("below_lines"):
                out.append("")
                out.append("본표 아래 메모 (빈 행 뒤, 번호·파트 없이 한 칸뿐). 대안 문구·코멘트다 — 쓰면 그 문항 src 에, 안 쓰면 skipped_src 에:")
                out += [f"- {x['ref']}: {x['text']}".replace("\n", " ⏎ ") for x in sh["below_lines"]]
        else:
            out.append(f"## 시트 `{sh['name']}` — 참고 시트 (머리 행 없음)")
            out += [f"- {x['ref']}: {x['text']}".replace("\n", " ⏎ ") for x in sh["lines"]]
        out.append("")
    return "\n".join(out)


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="구조화 엑셀 → structure.json / structure.md")
    ap.add_argument("xlsx")
    ap.add_argument("outdir")
    a = ap.parse_args(argv)
    doc = read(Path(a.xlsx))
    od = Path(a.outdir)
    od.mkdir(parents=True, exist_ok=True)
    (od / "structure.json").write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    (od / "structure.md").write_text(to_markdown(doc), encoding="utf-8")
    for sh in doc["sheets"]:
        if sh["kind"] == "structure":
            unk = [k for k, v in sh["columns"].items() if not v]
            print(f"[구조화] {sh['name']}: 행 {len(sh['rows'])}개 · 열 역할 "
                  + " · ".join(f"{k}→{v}" for k, v in sh["columns"].items() if v)
                  + (f" · 역할 미상 {unk}" if unk else ""))
        else:
            print(f"[참고]   {sh['name']}: 줄 {len(sh['lines'])}개")
    print(f"→ {od / 'structure.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
