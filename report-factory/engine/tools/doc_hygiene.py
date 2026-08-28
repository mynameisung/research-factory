# -*- coding: utf-8 -*-
"""
doc_hygiene.py — 납품 문서(xlsx/xlsm/pptx/docx)의 개인정보·품질 위생.

사용자 지시(2026-08-23): 「콘텐츠의 개인 정보 보호 및 품질 관리」. 참고한 외부 도구(watermarks-remover · claude-watermark-cleaner)
중 **정당한 위생 부분만** 가져왔다 — 보이지 않는 유니코드 제거와 문서 속성 정리. AI 생성물의 출처 표식(통계 워터마크·C2PA)을
지우거나 재작성해 탐지를 피하는 기능은 넣지 않는다(투명성 장치 우회).

무엇을 하나
  ① 문서 속성(docProps/core.xml · app.xml): 작성자·최종 수정자를 조직명으로, lastPrinted·Company·Manager 제거, 생성일을 지금으로.
     실측: 제작가이드 템플릿에서 「lastModifiedBy=개인 이름 · creator=Registered User · created=2018」 이 납품본까지 따라왔다.
  ② 본문 텍스트의 보이지 않는 문자(제로폭·양방향 제어·태그 문자·Hangul filler 등)를 제거하고 이형 공백(NBSP 등)을 보통 공백으로.
     이런 문자는 라벨 원문 대조 게이트·PPT 텍스트를 깨뜨리고, 양방향 제어는 보안상 위험하다.

  python engine/tools/doc_hygiene.py scan  <file>...
  python engine/tools/doc_hygiene.py clean <file>... [--creator 엠브레인] [--inplace | --out DIR]

zip 을 통째로 다시 쓰되 손대는 part 만 교체한다(나머지는 바이트 그대로 — build_guide 의 zip surgery 와 같은 원칙).
"""
from __future__ import annotations
import argparse, datetime as _dt, io, json, re, shutil, sys, unicodedata, zipfile
from pathlib import Path

# 보이지 않는·조작용 문자 — 제로폭, 양방향 제어, 태그, 필러, 소프트하이픈 (watermarks-remover · claude-watermark-cleaner 의 공개 목록과 같다)
SUSPICIOUS_RANGES = (
    (0x00AD, 0x00AD), (0x034F, 0x034F), (0x061C, 0x061C), (0x115F, 0x1160), (0x17B4, 0x17B5),
    (0x180B, 0x180F), (0x200B, 0x200F), (0x202A, 0x202E), (0x2060, 0x206F), (0x3164, 0x3164),
    (0xFEFF, 0xFEFF), (0xFFF9, 0xFFFB), (0xE0000, 0xE007F),
)
TEXT_PARTS = re.compile(r"^(xl/sharedStrings\.xml|xl/worksheets/sheet\d+\.xml|ppt/slides/slide\d+\.xml|ppt/notesSlides/.*\.xml|word/document\.xml|word/(header|footer)\d*\.xml|docProps/core\.xml|docProps/app\.xml)$")
PROP_PARTS = ("docProps/core.xml", "docProps/app.xml")
PERSONAL_TAGS = ("dc:creator", "cp:lastModifiedBy")
DROP_CORE = ("cp:lastPrinted", "dc:description", "cp:keywords", "cp:category", "cp:contentStatus", "dc:subject")
DROP_APP = ("Company", "Manager", "HyperlinkBase")
ORG_ALIASES = {"엠브레인", "Embrain", "EMBRAIN", "embrain", "(주)엠브레인", "마크로밀엠브레인"}   #: 조직명으로 인정하는 표기


def is_suspicious(ch: str) -> bool:
    cp = ord(ch)
    return any(a <= cp <= b for a, b in SUSPICIOUS_RANGES)


def sanitize_text(text: str) -> tuple[str, dict]:
    """보이지 않는 문자 제거 + 이형 공백 정규화. (정리된 문자열, {문자명: 개수})"""
    removed: dict[str, int] = {}
    out: list[str] = []
    for ch in text:
        if is_suspicious(ch):
            k = f"U+{ord(ch):04X} {unicodedata.name(ch, '?')}"
            removed[k] = removed.get(k, 0) + 1
            continue
        if ch != " " and unicodedata.category(ch) == "Zs":      # NBSP · 전각 공백 등
            k = f"U+{ord(ch):04X} {unicodedata.name(ch, '?')}→space"
            removed[k] = removed.get(k, 0) + 1
            out.append(" ")
            continue
        out.append(ch)
    return "".join(out), removed


def _xml_text_nodes(xml: str):
    """<t>…</t> · <a:t>…</a:t> · <w:t>…</w:t> 안의 텍스트만 건드린다 (태그·속성은 그대로)."""
    return re.finditer(r"(<(?:t|a:t|w:t)(?:\s[^>]*)?>)([^<]*)(</(?:t|a:t|w:t)>)", xml)


def scan_props(core: str | None, app: str | None, org: str) -> list[str]:
    issues = []
    if core:
        for tag in PERSONAL_TAGS:
            m = re.search(rf"<{tag}[^>]*>([^<]*)</{tag}>", core)
            if m and m.group(1).strip() and m.group(1).strip() != org and m.group(1).strip() not in ORG_ALIASES:
                issues.append(f"{tag} = {m.group(1).strip()!r} (조직명 '{org}' 이 아님)")
        for tag in DROP_CORE:
            if re.search(rf"<{tag}[^>]*>[^<]+</{tag}>", core):
                issues.append(f"{tag} 값 있음")
        m = re.search(r"<dcterms:created[^>]*>([^<]*)</dcterms:created>", core)
        if m and m.group(1)[:4].isdigit() and int(m.group(1)[:4]) < _dt.date.today().year - 1:
            issues.append(f"created = {m.group(1)} (템플릿 생성일 잔재)")
    if app:
        for tag in DROP_APP:
            if re.search(rf"<{tag}>[^<]+</{tag}>", app):
                issues.append(f"app.xml {tag} 값 있음")
    return issues


def clean_props(core: str, app: str | None, org: str) -> tuple[str, str | None]:
    now = _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    for tag in PERSONAL_TAGS:
        if re.search(rf"<{tag}[^>]*>", core):
            core = re.sub(rf"(<{tag}[^>]*>)[^<]*(</{tag}>)", rf"\g<1>{org}\g<2>", core)
        else:
            ns = ' xmlns:dc="http://purl.org/dc/elements/1.1/"' if tag.startswith("dc:") else ""
            core = core.replace("</cp:coreProperties>", f"<{tag}{ns}>{org}</{tag}></cp:coreProperties>")
    for tag in DROP_CORE:
        core = re.sub(rf"<{tag}[^>]*>[^<]*</{tag}>", "", core)
        core = re.sub(rf"<{tag}[^>]*/>", "", core)
    core = re.sub(r"(<dcterms:created[^>]*>)[^<]*(</dcterms:created>)", rf"\g<1>{now}\g<2>", core)
    core = re.sub(r"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)", rf"\g<1>{now}\g<2>", core)
    if app:
        for tag in DROP_APP:
            app = re.sub(rf"<{tag}>[^<]*</{tag}>", "", app)
    return core, app


def scan_file(path: Path, org: str) -> dict:
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    core = z.read("docProps/core.xml").decode("utf-8") if "docProps/core.xml" in names else None
    app = z.read("docProps/app.xml").decode("utf-8") if "docProps/app.xml" in names else None
    props = scan_props(core, app, org)
    chars: dict[str, int] = {}
    parts_hit: list[str] = []
    for n in names:
        if not TEXT_PARTS.match(n) or n in PROP_PARTS:
            continue
        xml = z.read(n).decode("utf-8", "replace")
        hit = False
        for m in _xml_text_nodes(xml):
            _, removed = sanitize_text(m.group(2))
            for k, v in removed.items():
                chars[k] = chars.get(k, 0) + v
                hit = True
        if hit:
            parts_hit.append(n)
    return {"file": str(path), "props": props, "invisible": chars, "parts": parts_hit}


def clean_file(path: Path, out: Path, org: str) -> dict:
    z = zipfile.ZipFile(path)
    report = {"file": str(path), "out": str(out), "props_changed": False, "invisible": {}, "parts": []}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zo:
        core = z.read("docProps/core.xml").decode("utf-8") if "docProps/core.xml" in z.namelist() else None
        app = z.read("docProps/app.xml").decode("utf-8") if "docProps/app.xml" in z.namelist() else None
        new_core, new_app = (clean_props(core, app, org) if core else (None, app))
        for info in z.infolist():
            data = z.read(info.filename)
            n = info.filename
            if n == "docProps/core.xml" and new_core is not None:
                if new_core != core:
                    report["props_changed"] = True
                data = new_core.encode("utf-8")
            elif n == "docProps/app.xml" and new_app is not None:
                if new_app != app:
                    report["props_changed"] = True
                data = new_app.encode("utf-8")
            elif TEXT_PARTS.match(n):
                xml = data.decode("utf-8", "replace")
                changed = False
                def _fix(m):
                    nonlocal changed
                    t, removed = sanitize_text(m.group(2))
                    if removed:
                        changed = True
                        for k, v in removed.items():
                            report["invisible"][k] = report["invisible"].get(k, 0) + v
                    return m.group(1) + t + m.group(3)
                xml2 = re.sub(r"(<(?:t|a:t|w:t)(?:\s[^>]*)?>)([^<]*)(</(?:t|a:t|w:t)>)", _fix, xml)
                if changed:
                    report["parts"].append(n)
                    data = xml2.encode("utf-8")
            zi = zipfile.ZipInfo(n, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zo.writestr(zi, data)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(buf.getvalue())
    return report


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["scan", "clean"])
    ap.add_argument("files", nargs="+", type=Path)
    ap.add_argument("--creator", default="엠브레인", help="작성자·최종 수정자에 넣을 조직명")
    ap.add_argument("--inplace", action="store_true")
    ap.add_argument("--out", type=Path, help="clean 결과를 둘 폴더 (기본: 원본 옆 *_clean)")
    ap.add_argument("--json", type=Path)
    a = ap.parse_args()
    results = []
    for f in a.files:
        if a.cmd == "scan":
            r = scan_file(f, a.creator)
            flag = "문제" if (r["props"] or r["invisible"]) else "깨끗"
            print(f"[{flag}] {f.name}")
            for p in r["props"]:
                print(f"    속성: {p}")
            for k, v in r["invisible"].items():
                print(f"    문자: {k} ×{v}  ({', '.join(r['parts'][:3])})")
        else:
            if a.inplace:
                tmp = f.with_suffix(f.suffix + ".tmp")
                r = clean_file(f, tmp, a.creator)
                shutil.move(tmp, f)
                r["out"] = str(f)
            else:
                out = (a.out or f.parent) / (f.stem + ("" if a.out else "_clean") + f.suffix)
                r = clean_file(f, out, a.creator)
            print(f"[clean] {f.name} → 속성 {'정리' if r['props_changed'] else '변경 없음'} · 제거 문자 {sum(r['invisible'].values())}개 → {r['out']}")
        results.append(r)
    if a.json:
        a.json.write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
