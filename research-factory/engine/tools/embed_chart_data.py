# -*- coding: utf-8 -*-
"""차트 데이터를 PPT 안에 포함(embedded)시킨다 — 외부 파일에 연결(linked)된 차트를 찾아 임베디드 워크북으로 바꾼다.

증상: 다른 컴퓨터에서 차트 「데이터 편집」을 누르면 「연결된 파일을 사용할 수 없습니다」 대화상자가 뜬다.
원인: 차트의 <c:externalData r:id> 가 가리키는 관계가 TargetMode="External"(로컬 xlsx 경로)이거나,
      임베디드 파트가 패키지 안에 없다.

처리: 차트마다
  ① 연결 대상 파일이 이 컴퓨터에 있으면 그 파일을 그대로 embeddings/ 에 넣는다.
  ② 없으면 차트 XML 의 캐시(카테고리·계열 값·계열 이름)로 워크북을 새로 만든다 — c:f 참조 셀 좌표 그대로 쓴다.
  그런 다음 관계를 내부 package 로 바꾸고 [Content_Types] 를 맞춘다.

    python embed_chart_data.py <src.pptx> <out.pptx>          # 변환
    python embed_chart_data.py <src.pptx> --check             # 진단만 (변환하지 않음)
"""
from __future__ import annotations
import io, re, sys, urllib.parse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckops as D
import deckbuild as B
from lxml import etree

C = D.C
R_NS = D.R
PKG_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
OLE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"


def _ext_rel(deck: D.Deck, chart_part: str):
    """(rId, rel dict | None, kind) — kind: 'embedded' | 'external' | 'missing-rel' | 'missing-part' | 'no-externalData'"""
    tree = deck.xml(chart_part)
    ed = tree.find(f".//{D.q(C, 'externalData')}")
    if ed is None:
        return None, None, "no-externalData"
    rid = ed.get(D.q(R_NS, "id"))
    for rel in deck.rels(chart_part):
        if rel.get("Id") == rid:
            if rel.get("TargetMode", "").lower() == "external":
                return rid, rel, "external"
            target = rel.get("Target", "")
            part = D.resolve_target(chart_part, target) if hasattr(D, "resolve_target") else _resolve(chart_part, target)
            return rid, rel, ("embedded" if part in deck.parts else "missing-part")
    return rid, None, "missing-rel"


def _resolve(base_part: str, target: str) -> str:
    base_dir = base_part.rsplit("/", 1)[0]
    segs = (base_dir + "/" + target).split("/")
    out: list[str] = []
    for s in segs:
        if s == "..":
            out.pop()
        elif s and s != ".":
            out.append(s)
    return "/".join(out)


def _parse_ref(f: str):
    """'Sheet1!$B$2:$B$8' → (sheet, [(row, col), ...]) — 1-based."""
    m = re.match(r"^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", f.strip())
    if not m:
        return None, []
    sheet, c1, r1, c2, r2 = m.groups()
    from openpyxl.utils import column_index_from_string as ci
    c1i, r1i = ci(c1), int(r1)
    c2i, r2i = (ci(c2), int(r2)) if c2 else (c1i, r1i)
    cells = [(r, c) for r in range(r1i, r2i + 1) for c in range(c1i, c2i + 1)]
    return sheet, cells


def _workbook_from_cache(tree: etree._Element) -> bytes:
    """차트 캐시로 워크북을 만든다. c:f 가 있으면 그 좌표에, 없으면 A1 부터 표 모양으로 쓴다."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws0 = wb.active
    sheets = {}

    def sheet(name):
        if not name:
            return ws0
        if name not in sheets:
            if ws0.title == "Sheet" and not sheets:
                ws0.title = name
                sheets[name] = ws0
            else:
                sheets[name] = wb.create_sheet(name)
        return sheets[name]

    def write(ref_el, kind):
        """ref_el = c:strRef | c:numRef. kind 'str'|'num'. 반환: 썼는지."""
        if ref_el is None:
            return False
        f = ref_el.findtext(D.q(C, "f"))
        cache = ref_el.find(D.q(C, "strCache" if kind == "str" else "numCache"))
        pts = {}
        if cache is not None:
            for pt in cache.findall(D.q(C, "pt")):
                v = pt.findtext(D.q(C, "v"))
                pts[int(pt.get("idx"))] = v
        sh, cells = _parse_ref(f) if f else (None, [])
        if not cells:
            return False
        ws = sheet(sh)
        for i, (r, c) in enumerate(cells):
            v = pts.get(i)
            if v is None:
                continue
            if kind == "num":
                try:
                    v = float(v)
                except ValueError:
                    pass
            ws.cell(row=r, column=c, value=v)
        return True

    wrote = False
    for ser in D.chart_series(tree):
        tx = ser.find(D.q(C, "tx"))
        if tx is not None:
            wrote |= write(tx.find(D.q(C, "strRef")), "str")
        cat = ser.find(D.q(C, "cat"))
        if cat is not None:
            wrote |= write(cat.find(D.q(C, "strRef")), "str") or write(cat.find(D.q(C, "numRef")), "num")
        val = ser.find(D.q(C, "val"))
        if val is not None:
            wrote |= write(val.find(D.q(C, "numRef")), "num")
        for tag in ("xVal", "yVal"):
            el = ser.find(D.q(C, tag))
            if el is not None:
                wrote |= write(el.find(D.q(C, "numRef")), "num") or write(el.find(D.q(C, "strRef")), "str")
    if not wrote:                                   # c:f 없음 → 캐시를 표로 쓴다
        cats = D.chart_categories(tree)
        names = D.series_names(tree)
        vals = D.series_values(tree)
        ws = ws0
        for j, n in enumerate(names):
            ws.cell(row=1, column=j + 2, value=n)
        for i, c in enumerate(cats):
            ws.cell(row=i + 2, column=1, value=c)
        for j, col in enumerate(vals):
            for i, v in enumerate(col):
                if v is not None:
                    ws.cell(row=i + 2, column=j + 2, value=v)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _local_file(target: str) -> Path | None:
    t = target
    if t.lower().startswith("file:"):
        t = urllib.parse.unquote(t[5:].lstrip("/"))
        if re.match(r"^[A-Za-z]:", t) is None and re.match(r"^[A-Za-z]%3A", t) is None:
            t = "/" + t
    p = Path(t)
    return p if p.exists() and p.is_file() else None


def check(deck: D.Deck) -> list[tuple[str, str, str]]:
    """[(chart_part, kind, target)] — kind 가 'embedded' 가 아닌 것만."""
    out = []
    for part in deck.names(r"ppt/charts/chart\d+\.xml$"):
        rid, rel, kind = _ext_rel(deck, part)
        if kind not in ("embedded", "no-externalData"):
            out.append((part, kind, (rel or {}).get("Target", "")))
    return out


def embed(deck: D.Deck) -> list[tuple[str, str, str]]:
    """외부 연결·결손 차트를 임베디드로 바꾼다. 반환 [(chart_part, 방법, 상세)]."""
    done = []
    n = 0
    existing = set(deck.names(r"ppt/embeddings/"))
    for part in deck.names(r"ppt/charts/chart\d+\.xml$"):
        rid, rel, kind = _ext_rel(deck, part)
        if kind in ("embedded", "no-externalData"):
            continue
        tree = deck.xml(part)
        method, detail = "", ""
        data = None
        if kind == "external":
            lf = _local_file(rel.get("Target", ""))
            if lf is not None:
                data = lf.read_bytes(); method, detail = "로컬 파일 복사", str(lf)
        if data is None:
            data = _workbook_from_cache(tree); method, detail = "캐시로 재구성", f"계열 {len(D.chart_series(tree))}"
        while True:
            n += 1
            new_part = f"ppt/embeddings/Microsoft_Excel_Worksheet_embed{n}.xlsx"
            if new_part not in existing and new_part not in deck.parts:
                break
        deck.put_bytes(new_part, data)
        existing.add(new_part)
        # 관계 재작성
        rels_name = deck.rels_name(part)
        rtree = deck.xml(rels_name)
        RELNS = "http://schemas.openxmlformats.org/package/2006/relationships"
        found = None
        for r in rtree.findall(f"{{{RELNS}}}Relationship"):
            if r.get("Id") == rid:
                found = r
        if found is None:
            found = etree.SubElement(rtree, f"{{{RELNS}}}Relationship")
            found.set("Id", rid)
        found.set("Type", PKG_TYPE)
        found.set("Target", "../embeddings/" + new_part.rsplit("/", 1)[-1])
        if "TargetMode" in found.attrib:
            del found.attrib["TargetMode"]
        deck.put(rels_name, rtree)
        # externalData 의 autoUpdate 는 0 으로
        ed = tree.find(f".//{D.q(C, 'externalData')}")
        au = ed.find(D.q(C, "autoUpdate"))
        if au is None:
            au = etree.SubElement(ed, D.q(C, "autoUpdate"))
        au.set("val", "0")
        deck.put(part, tree)
        done.append((part, method, detail))
    if done:
        B.sync_content_types(deck)
    return done


def main(argv):
    src = Path(argv[1])
    deck = D.Deck(src)
    if len(argv) > 2 and argv[2] == "--check":
        bad = check(deck)
        print(f"차트 {len(deck.names(r'ppt/charts/chart\d+\.xml$'))}개 · 포함되지 않은 차트 {len(bad)}개")
        for p, k, t in bad:
            print(f"  {p}  {k}  {t[:100]}")
        return 1 if bad else 0
    out = Path(argv[2])
    done = embed(deck)
    deck.save(out)
    print(f"포함 처리 {len(done)}개 → {out}")
    for p, m, d in done:
        print(f"  {p}  {m}  {d[:100]}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
