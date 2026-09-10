"""통계표 .xls(BIFF8 OLE2 복합문서) → tables.json.

    python parse_table_biff.py <input.xls> <output.json> [--strict] [--quiet]

읽기 경로는 두 단계 폴백이다.

1. **xlrd 2.x** — `.xls` 전용으로 남아 있는 버전이며 `formatting_info=True`로
   XF/FORMAT 레코드까지 준다. 셀별 표시 형식(`###0.0` vs `###0.00`)을 그대로
   얻을 수 있으므로 1순위.
2. **olefile + 자체 BIFF8 파서** — xlrd가 없을 때. `Workbook` 스트림에서
   BOF/SST/CONTINUE/LABELSST/LABEL/RK/MULRK/NUMBER/FORMULA/XF/FORMAT/
   MERGEDCELLS/BOUNDSHEET 레코드를 직접 해석한다.

두 경로 모두 `parse_table_common.Grid`를 만들고, 이후 처리는 공통 모듈이 한다.

이 파일은 그대로 import 해서 `parse(path)` / `read_grids(path)`를 쓸 수 있다.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import struct
import sys
from pathlib import Path
from typing import Any

try:  # 모듈로 import 될 때와 스크립트로 실행될 때 모두 동작하게
    from parse_table_common import (
        BANNER_GROUPS_IN_REPORT,
        Grid,
        TableParseError,
        build_document,
        normalize_grid_columns,
        summarize,
    )
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from parse_table_common import (  # type: ignore
        BANNER_GROUPS_IN_REPORT,
        Grid,
        TableParseError,
        build_document,
        normalize_grid_columns,
        summarize,
    )

OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
INDEX_SHEET_HINTS = ("목", "차", "index", "toc", "contents")
TABLE_SHEET_HINTS = ("table", "표")


# ── 내장 서식 코드 (BIFF 기본 제공, FORMAT 레코드로 안 나옴) ──────────────────
BUILTIN_FORMATS: dict[int, str] = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    5: '"$"#,##0_);("$"#,##0)', 6: '"$"#,##0_);[Red]("$"#,##0)',
    7: '"$"#,##0.00_);("$"#,##0.00)', 8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "M/D/YY", 15: "D-MMM-YY", 16: "D-MMM", 17: "MMM-YY",
    18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
    22: "M/D/YY h:mm", 37: "#,##0_);(#,##0)", 38: "#,##0_);[Red](#,##0)",
    39: "#,##0.00_);(#,##0.00)", 40: "#,##0.00_);[Red](#,##0.00)",
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@",
}


# ═══════════════════════════════════════════════════════════════════════════
# 1순위 경로 — xlrd
# ═══════════════════════════════════════════════════════════════════════════
def _grids_via_xlrd(path: Path) -> tuple[dict[str, Grid], str]:
    import xlrd  # type: ignore

    bk = xlrd.open_workbook(str(path), formatting_info=True)
    grids: dict[str, Grid] = {}
    for sh in bk.sheets():
        values: list[list[Any]] = []
        formats: list[list[str]] = []
        for r in range(sh.nrows):
            vrow: list[Any] = []
            frow: list[str] = []
            for c in range(sh.ncols):
                ctype = sh.cell_type(r, c)
                if ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    vrow.append(None)
                    frow.append("General")
                    continue
                vrow.append(sh.cell_value(r, c))
                try:
                    xf = bk.xf_list[sh.cell_xf_index(r, c)]
                    frow.append(bk.format_map[xf.format_key].format_str)
                except Exception:  # noqa: BLE001 — 서식 정보가 없는 셀
                    frow.append("General")
            values.append(vrow)
            formats.append(frow)
        grids[sh.name] = Grid(name=sh.name, values=values, formats=formats)
    return grids, f"xlrd {xlrd.__version__}"


# ═══════════════════════════════════════════════════════════════════════════
# 2순위 경로 — olefile + 자체 BIFF8 파서
# ═══════════════════════════════════════════════════════════════════════════
class _BiffReader:
    """BIFF8 `Workbook` 스트림 최소 파서.

    필요한 레코드만 처리한다:
      BOF(0x0809) EOF(0x000A) BOUNDSHEET(0x0085) SST(0x00FC) CONTINUE(0x003C)
      FORMAT(0x041E) XF(0x00E0) MERGEDCELLS(0x00E5)
      LABELSST(0x00FD) LABEL(0x0204) RSTRING(0x00D6) RK(0x027E) MULRK(0x00BD)
      NUMBER(0x0203) FORMULA(0x0006) STRING(0x0207) BOOLERR(0x0205)
      BLANK(0x0201) MULBLANK(0x00BE)
    """

    def __init__(self, data: bytes) -> None:
        self.data = data
        self.sst: list[str] = []
        self.formats: dict[int, str] = dict(BUILTIN_FORMATS)
        self.xfs: list[int] = []                   # xf index -> format key
        self.sheets: list[tuple[str, int]] = []    # (name, stream offset)
        self.merged: dict[str, list[tuple[int, int, int, int]]] = {}

    # ── 레코드 순회 ────────────────────────────────────────────────────────
    def _records(self, start: int = 0):
        d = self.data
        pos = start
        n = len(d)
        while pos + 4 <= n:
            rid, size = struct.unpack_from("<HH", d, pos)
            body = d[pos + 4: pos + 4 + size]
            pos = pos + 4 + size
            yield rid, body

    # ── 문자열 ─────────────────────────────────────────────────────────────
    @staticmethod
    def _read_xl_unicode(buf: bytes, pos: int, nchars: int,
                         boundaries: frozenset[int] = frozenset()) -> tuple[str, int]:
        """BIFF8 XLUnicodeString 본문(길이 필드 제외)을 읽는다.

        `boundaries`는 CONTINUE 조각이 시작하는 절대 오프셋 집합이다. 문자열이
        경계를 넘어가면 조각 선두에 **option flags 바이트가 하나 더** 들어가고
        압축(1바이트)/비압축(2바이트) 여부가 바뀔 수 있다. 이 처리가 빠지면
        SST 인덱스가 통째로 어긋난다.
        """
        flags = buf[pos]
        pos += 1
        high = flags & 0x01
        ext = flags & 0x04
        rich = flags & 0x08
        nrun = cbext = 0
        if rich:
            nrun = struct.unpack_from("<H", buf, pos)[0]
            pos += 2
        if ext:
            cbext = struct.unpack_from("<I", buf, pos)[0]
            pos += 4

        out: list[str] = []
        need = nchars
        while need > 0:
            nxt = min((b for b in boundaries if b > pos), default=len(buf))
            avail = max(min(nxt, len(buf)) - pos, 0)
            if high:
                take = min(need, avail // 2)
                out.append(buf[pos:pos + take * 2].decode("utf-16-le", "replace"))
                pos += take * 2
            else:
                take = min(need, avail)
                out.append(buf[pos:pos + take].decode("latin-1"))
                pos += take
            need -= take
            if need <= 0:
                break
            if pos >= len(buf):
                break
            # 경계에 도달 — 새 조각의 option flags 바이트를 소비한다.
            high = buf[pos] & 0x01
            pos += 1
        # rich-run / phonetic 데이터는 경계를 넘어도 flags 바이트가 붙지 않는다.
        return "".join(out), pos + nrun * 4 + cbext

    def _read_sst(self, chunks: list[bytes]) -> None:
        """SST + 이어지는 CONTINUE 조각들에서 공유 문자열 테이블을 복원한다."""
        if not chunks:
            return
        buf = b"".join(chunks)
        boundaries: set[int] = set()
        acc = 0
        for ch in chunks[:-1]:
            acc += len(ch)
            boundaries.add(acc)
        frozen = frozenset(boundaries)

        total = struct.unpack_from("<I", buf, 4)[0]
        pos = 8
        for _ in range(total):
            if pos + 3 > len(buf):
                break
            nchars = struct.unpack_from("<H", buf, pos)[0]
            pos += 2
            s, pos = self._read_xl_unicode(buf, pos, nchars, frozen)
            self.sst.append(s)

    @classmethod
    def _unicode_body(cls, buf: bytes, pos: int, nchars: int) -> tuple[str, int, int]:
        """단일 레코드 안에 들어 있는 문자열용 얇은 래퍼."""
        s, newpos = cls._read_xl_unicode(buf, pos, nchars)
        return s, newpos, nchars - len(s)

    @staticmethod
    def _short_unicode(body: bytes, pos: int) -> tuple[str, int]:
        nchars = body[pos]
        pos += 1
        flags = body[pos]
        pos += 1
        if flags & 0x01:
            s = body[pos:pos + nchars * 2].decode("utf-16-le", "replace")
            pos += nchars * 2
        else:
            s = body[pos:pos + nchars].decode("latin-1")
            pos += nchars
        return s, pos

    # ── 글로벌 서브스트림 ──────────────────────────────────────────────────
    def read_globals(self) -> None:
        sst_chunks: list[bytes] = []
        in_sst = False
        for rid, body in self._records():
            if rid == 0x000A:  # EOF
                break
            if in_sst and rid == 0x003C:  # CONTINUE
                sst_chunks.append(body)
                continue
            if in_sst:
                self._read_sst(sst_chunks)
                sst_chunks, in_sst = [], False
            if rid == 0x00FC:  # SST
                sst_chunks = [body]
                in_sst = True
            elif rid == 0x0085:  # BOUNDSHEET
                off = struct.unpack_from("<I", body, 0)[0]
                name, _ = self._short_unicode(body, 6)
                self.sheets.append((name, off))
            elif rid == 0x041E:  # FORMAT
                idx, nchars = struct.unpack_from("<HH", body, 0)
                flags = body[4]
                if flags & 0x01:
                    s = body[5:5 + nchars * 2].decode("utf-16-le", "replace")
                else:
                    s = body[5:5 + nchars].decode("latin-1")
                self.formats[idx] = s
            elif rid == 0x00E0:  # XF
                self.xfs.append(struct.unpack_from("<H", body, 2)[0])
        if in_sst:
            self._read_sst(sst_chunks)

    # ── 시트 서브스트림 ────────────────────────────────────────────────────
    @staticmethod
    def _rk_to_float(rk: int) -> float:
        is_mul100 = rk & 0x01
        is_int = rk & 0x02
        v = rk & 0xFFFFFFFC
        if is_int:
            num = float(struct.unpack("<i", struct.pack("<I", v))[0] >> 2)
        else:
            num = struct.unpack("<d", struct.pack("<Q", v << 32))[0]
        return num / 100.0 if is_mul100 else num

    def read_sheet(self, name: str, offset: int) -> Grid:
        cells: dict[tuple[int, int], tuple[Any, int]] = {}
        merged: list[tuple[int, int, int, int]] = []
        maxr = maxc = -1
        pending_formula: tuple[int, int, int] | None = None

        for rid, body in self._records(offset):
            if rid == 0x000A:  # EOF
                break
            val: Any
            if rid == 0x00FD:  # LABELSST
                r, c, xf, isst = struct.unpack_from("<HHHI", body, 0)
                val = self.sst[isst] if isst < len(self.sst) else ""
            elif rid in (0x0204, 0x00D6):  # LABEL / RSTRING
                r, c, xf = struct.unpack_from("<HHH", body, 0)
                nchars = struct.unpack_from("<H", body, 6)[0]
                val, _, _ = self._unicode_body(body, 8, nchars)
            elif rid == 0x027E:  # RK
                r, c, xf, rk = struct.unpack_from("<HHHI", body, 0)
                val = self._rk_to_float(rk)
            elif rid == 0x00BD:  # MULRK
                r, c1 = struct.unpack_from("<HH", body, 0)
                c2 = struct.unpack_from("<H", len(body) >= 2 and body or b"\0\0",
                                        len(body) - 2)[0]
                p = 4
                for c in range(c1, c2 + 1):
                    xf, rk = struct.unpack_from("<HI", body, p)
                    p += 6
                    cells[(r, c)] = (self._rk_to_float(rk), xf)
                    maxr, maxc = max(maxr, r), max(maxc, c)
                continue
            elif rid == 0x0203:  # NUMBER
                r, c, xf = struct.unpack_from("<HHH", body, 0)
                val = struct.unpack_from("<d", body, 6)[0]
            elif rid == 0x0006:  # FORMULA
                r, c, xf = struct.unpack_from("<HHH", body, 0)
                raw = body[6:14]
                if raw[6:8] == b"\xff\xff":
                    kind = raw[0]
                    if kind == 0x00:      # 결과가 문자열 → 뒤의 STRING 레코드
                        pending_formula = (r, c, xf)
                        continue
                    val = bool(raw[2]) if kind == 0x01 else None
                else:
                    val = struct.unpack("<d", raw)[0]
            elif rid == 0x0207 and pending_formula is not None:  # STRING
                r, c, xf = pending_formula
                pending_formula = None
                nchars = struct.unpack_from("<H", body, 0)[0]
                val, _, _ = self._unicode_body(body, 2, nchars)
            elif rid == 0x0205:  # BOOLERR
                r, c, xf = struct.unpack_from("<HHH", body, 0)
                val = bool(body[6]) if body[7] == 0 else None
            elif rid == 0x00E5:  # MERGEDCELLS
                n = struct.unpack_from("<H", body, 0)[0]
                for i in range(n):
                    r1, r2, c1, c2 = struct.unpack_from("<HHHH", body, 2 + i * 8)
                    merged.append((r1, r2 + 1, c1, c2 + 1))
                continue
            else:
                continue

            if val is None or val == "":
                continue
            cells[(r, c)] = (val, xf)
            maxr, maxc = max(maxr, r), max(maxc, c)

        nrows, ncols = maxr + 1, maxc + 1
        values: list[list[Any]] = [[None] * ncols for _ in range(nrows)]
        formats: list[list[str]] = [["General"] * ncols for _ in range(nrows)]
        for (r, c), (v, xf) in cells.items():
            values[r][c] = v
            fk = self.xfs[xf] if xf < len(self.xfs) else 0
            formats[r][c] = self.formats.get(fk, "General")
        self.merged[name] = merged
        return Grid(name=name, values=values, formats=formats)


def _grids_via_olefile(path: Path) -> tuple[dict[str, Grid], str]:
    import olefile  # type: ignore

    ole = olefile.OleFileIO(str(path))
    try:
        stream_name = next((s for s in ("Workbook", "Book") if ole.exists(s)), None)
        if stream_name is None:
            listing = ["/".join(e) for e in ole.listdir()]
            raise TableParseError(
                f"{path.name}: OLE2 안에 Workbook/Book 스트림이 없습니다. (스트림: {listing})"
            )
        data = ole.openstream(stream_name).read()
    finally:
        ole.close()

    rd = _BiffReader(data)
    rd.read_globals()
    if not rd.sheets:
        raise TableParseError(f"{path.name}: BOUNDSHEET 레코드를 찾지 못했습니다.")
    grids = {name: rd.read_sheet(name, off) for name, off in rd.sheets}
    return grids, f"olefile {olefile.__version__} + 내장 BIFF8 파서"


# ═══════════════════════════════════════════════════════════════════════════
def read_grids(path: Path, *, prefer: str = "auto") -> tuple[dict[str, Grid], str]:
    """워크북 전체를 시트명 → Grid 로 읽는다. 반환 (grids, 사용한 리더 설명)."""
    with path.open("rb") as fh:
        head = fh.read(8)
    if head != OLE2_MAGIC:
        raise TableParseError(
            f"{path.name}: OLE2 복합문서 매직({OLE2_MAGIC.hex().upper()})이 아닙니다 "
            f"(실제 {head.hex().upper()}). HTML 위장 .xls 이거나 .xlsx 라면 "
            "parse_table_xlsx.py 를 쓰세요."
        )
    errors: list[str] = []
    if prefer in ("auto", "xlrd"):
        try:
            return _grids_via_xlrd(path)
        except ImportError as e:
            errors.append(f"xlrd 사용 불가: {e}")
        except Exception as e:  # noqa: BLE001
            errors.append(f"xlrd 실패: {type(e).__name__}: {e}")
            if prefer == "xlrd":
                raise
    if prefer in ("auto", "biff"):
        try:
            return _grids_via_olefile(path)
        except ImportError as e:
            errors.append(f"olefile 사용 불가: {e}")
        except Exception as e:  # noqa: BLE001
            errors.append(f"자체 BIFF8 파서 실패: {type(e).__name__}: {e}")
            if prefer == "biff":
                raise
    raise TableParseError(
        f"{path.name}: 어떤 리더로도 읽지 못했습니다.\n  " + "\n  ".join(errors) +
        "\n  → `pip install xlrd` 또는 `pip install olefile` 후 다시 시도하세요."
    )


def pick_sheets(grids: dict[str, Grid]) -> tuple[Grid, Grid | None]:
    """(본표 시트, 목차 시트)를 고른다."""
    table = next((g for g in grids.values()
                  if any(h in g.name.lower() for h in TABLE_SHEET_HINTS)), None)
    if table is None:  # 이름으로 못 찾으면 행이 가장 많은 시트를 본표로 본다.
        table = max(grids.values(), key=lambda g: g.nrows)
    index = next((g for g in grids.values()
                  if g is not table
                  and any(h in g.name.lower() for h in INDEX_SHEET_HINTS)), None)
    return table, index


def parse(path: str | Path, *, prefer: str = "auto",
          in_report_groups: tuple[str, ...] = BANNER_GROUPS_IN_REPORT,
          skip_bad_tables: bool = False) -> dict[str, Any]:
    p = Path(path)
    if not p.is_file():
        raise TableParseError(f"입력 파일이 없습니다: {p}")
    grids, reader = read_grids(p, prefer=prefer)
    table_grid, index_grid = pick_sheets(grids)
    # 수작업 열이 끼어 밀린 파일 보정 (정상 파일에서는 항등)
    table_grid, col_fix = normalize_grid_columns(table_grid)
    if col_fix["dropped_leading"] or col_fix["dropped_middle"]:
        print(f"[parse_table_biff] 열 밀림 보정: 앞 {col_fix['dropped_leading']}열 제거"
              f" · 중간 열 {col_fix['dropped_middle']} 제거", file=sys.stderr)
    return build_document(
        table_grid, index_grid,
        source=str(p.resolve()),
        source_format="biff8-xls (OLE2)",
        parser=Path(__file__).name,
        in_report_groups=in_report_groups,
        skip_bad_tables=skip_bad_tables,
        extra_meta={
            "reader": reader,
            "sheets": list(grids),
            "column_normalization": col_fix,
            "generated_at": _dt.datetime.now().isoformat(timespec="seconds"),
        },
    )


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="통계표 .xls(BIFF8 OLE2) → tables.json")
    ap.add_argument("input", help="통계표 .xls 경로")
    ap.add_argument("output", help="출력 tables.json 경로")
    ap.add_argument("--prefer", choices=["auto", "xlrd", "biff"], default="auto",
                    help="리더 선택 (기본 auto: xlrd → 자체 BIFF8 파서)")
    ap.add_argument("--strict", action="store_true",
                    help="목차 교차검증 경고도 실패로 취급")
    ap.add_argument("--skip-bad", action="store_true",
                    help="파싱 불가한 표를 중단 없이 스킵하고 meta.skipped_tables 에 기록 "
                         "(제로베이스 표 등 — 스킵 내역은 경고로도 남는다)")
    ap.add_argument("--in-groups", default="",
                    help="보고서에 싣는 배너군 이름을 세미콜론으로 구분해 준다 (R-056 그룹명 화이트리스트). "
                         "예: \"전체;건물 소유 형태;지역;연면적\". 생략하면 비상주 기본값을 쓴다")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    try:
        in_groups = (tuple(g.strip() for g in args.in_groups.split(";") if g.strip())
                     or BANNER_GROUPS_IN_REPORT)
        doc = parse(args.input, prefer=args.prefer, skip_bad_tables=args.skip_bad,
                    in_report_groups=in_groups)
    except TableParseError as e:
        print(f"[parse_table_biff] 파싱 실패: {e}", file=sys.stderr)
        return 2

    warnings = doc["meta"]["warnings"]
    if args.strict and warnings:
        print("[parse_table_biff] --strict: 경고가 있어 출력하지 않습니다.", file=sys.stderr)
        for w in warnings:
            print(f"  - {w}", file=sys.stderr)
        return 3

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    if not args.quiet:
        print(summarize(doc))
        print(f"→ {out}  ({out.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
