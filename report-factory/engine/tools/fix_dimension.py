#!/usr/bin/env python3
"""시트의 `<dimension>` 선언을 실제 내용 범위와 일치시킨다.

**왜 필요한가**
`<dimension ref="A3:V2007">` 로 선언해 놓고 실제로 2008행에 셀이 있으면 Excel 이
"내용에 문제가 있습니다" 복구 안내를 띄운다. openpyxl 로 쓴 뒤 zip 수술로 시트를
갈아 끼우면 dimension 이 갱신되지 않아 이 어긋남이 생긴다.

우리 검사기(파트 선언·rel·요소 순서·스타일 범위)는 전부 통과시키는데 Excel 은 손상이라고
하던 그 간극의 후보 하나다. **이것만으로 복구 안내가 사라진다고 단정하지 않는다** —
고친 뒤 실제로 열어 확인해야 한다.

사용법
------
    python fix_dimension.py <파일.xlsm> [--out 새파일.xlsm] [--check]
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

COL = re.compile(r"<c r=\"([A-Z]+)\d+\"")
ROW = re.compile(r"<row r=\"(\d+)\"")
DIM = re.compile(r"<dimension ref=\"([^\"]+)\"\s*/>")


def col_idx(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def col_name(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def actual_range(xml: str) -> str | None:
    rows = [int(m) for m in ROW.findall(xml)]
    cols = [col_idx(m) for m in COL.findall(xml)]
    if not rows:
        return None
    return f"A1:{col_name(max(cols) if cols else 1)}{max(rows)}"


def scan(path: Path) -> list[tuple[str, str, str]]:
    """(시트파트, 선언, 실제) 중 어긋난 것만."""
    out = []
    with zipfile.ZipFile(path) as z:
        for n in sorted(x for x in z.namelist()
                        if re.match(r"xl/worksheets/sheet\d+\.xml$", x)):
            xml = z.read(n).decode("utf-8")
            m = DIM.search(xml)
            act = actual_range(xml)
            if not m or not act:
                continue
            decl = m.group(1)
            end = decl.split(":")[-1]
            mm = re.match(r"([A-Z]+)(\d+)", end)
            if not mm:
                continue
            dc, dr = col_idx(mm.group(1)), int(mm.group(2))
            am = re.match(r"([A-Z]+)(\d+)", act.split(":")[1])
            ac, ar = col_idx(am.group(1)), int(am.group(2))
            if ar > dr or ac > dc:
                out.append((n, decl, act))
    return out


def fix(target: Path, out: Path | None = None) -> list[tuple[str, str, str]]:
    """어긋난 시트를 실제 범위로 고치고, 고친 목록을 돌려준다.

    빌드마다 다시 어긋난다 — openpyxl 이 쓴 뒤 zip 수술로 시트를 갈아 끼우면
    `dimension` 이 갱신되지 않기 때문이다. 그래서 손으로 부르는 도구로 두지 않고
    `build_guide.py` 가 매번 부르게 했다. 잊으면 Excel 복구 안내가 그대로 돌아온다.
    """
    bad = scan(target)
    if not bad:
        return []
    out = out or target
    tmp = out.with_suffix(out.suffix + ".tmp")
    repl = {n: act for n, _, act in bad}
    with zipfile.ZipFile(target) as zin,             zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in repl:
                # 선언을 실제 범위로 갈아끼운다. 시작 셀도 A1 로 정규화한다 —
                # `A3` 처럼 앞을 잘라 두면 그 위 행의 서식을 Excel 이 놓칠 수 있다.
                data = DIM.sub(f'<dimension ref="{repl[item.filename]}"/>',
                               data.decode("utf-8"), count=1).encode("utf-8")
            zout.writestr(item, data)
    shutil.move(str(tmp), str(out))
    return bad


def main() -> int:
    ap = argparse.ArgumentParser(description="dimension 선언을 실제 범위에 맞춘다")
    ap.add_argument("target", type=Path)
    ap.add_argument("--out", type=Path, default=None)
    ap.add_argument("--check", action="store_true", help="고치지 않고 검사만")
    args = ap.parse_args()

    bad = scan(args.target)
    if not bad:
        print("dimension 어긋남 없음")
        return 0
    for n, decl, act in bad:
        print(f"  ✗ {n}  선언={decl}  실제={act}")
    if args.check:
        return 1

    out = args.out or args.target
    fix(args.target, out)
    print(f"{chr(10)}{len(bad)}개 시트 수정 → {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
