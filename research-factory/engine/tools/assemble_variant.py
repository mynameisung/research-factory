# -*- coding: utf-8 -*-
"""
assemble_variant.py — 같은 파이프라인(build_facts → build_guide → verify)으로 **문장 담당만 바꾼** 제작가이드 변형본을 만든다.
A/B 비교용: 본문 멘트·요약 헤드라인/해석·결론 제목/본문·서브 타이틀을 외부 산출(JSON)로 치환하고 나머지(목차 구조·패널·표)는 정본 그대로.

  python engine/tools/assemble_variant.py <work_dir> <name> \
      [--ments ments.json]            # {ments:[{seq, lines:[...]}]}
      [--summary summary.json]        # {"1": {"headline": "...", "body": ["..."]}, ...}
      [--conclusion conclusion.json]  # {"1": {"title": "...", "body": ["...", "→ ..."]}, ...}
      [--toc-names names.json]        # {"보유규모": "소유 건물 수", ...}  (서브 타이틀 어휘만)
      --build-args "<build_guide.py 공통 인자>"  --facts-args "<build_facts.py 공통 인자>"

산출: <work_dir>/variants/<name>/{목차.md, 멘트.json, slide_facts.json, KPI…json}
      <work_dir>/../out/variants/<name>/… .xlsm (+ _제작팀송부본)   · verify 결과 <variants/<name>/verify_*.json>
"""
from __future__ import annotations
import argparse, json, re, shutil, subprocess, sys
from pathlib import Path

TOOLS = Path(__file__).resolve().parent


def patch_toc(toc: str, summary: dict | None, conclusion: dict | None, names: dict | None) -> str:
    lines = toc.split("\n")
    out = []
    i = 0
    section = None
    cur_no = None
    while i < len(lines):
        ln = lines[i]
        if ln.startswith("### 요약 꼭지"):
            section = "summary"
        elif ln.startswith("### 결론 및 제언"):
            section = "conclusion"
        elif ln.startswith("## ") or (ln.startswith("### ") and section):
            if not (ln.startswith("### 요약") or ln.startswith("### 결론")):
                section = None
        m = re.match(r"\*\*(\d+)\. (.+?)\*\*( ← .*)?$", ln)
        if section and m:
            cur_no = m.group(1)
            if section == "conclusion" and conclusion and cur_no in conclusion:
                ln = f"**{cur_no}. {conclusion[cur_no]['title']}**{m.group(3) or ''}"
                out.append(ln)
                # 본문 줄 치환: 다음 빈 줄 전까지
                i += 1
                while i < len(lines) and lines[i].strip() and not lines[i].startswith("**"):
                    i += 1
                out.extend(conclusion[cur_no]["body"])
                continue
        if section == "summary" and summary and cur_no in summary and ln.startswith("- "):
            out.append("- " + summary[cur_no]["headline"].lstrip("- ").strip())
            i += 1
            # body 줄(헤드라인 다음, 빈 줄 전까지) 치환
            while i < len(lines) and lines[i].strip() and not lines[i].startswith("**"):
                i += 1
            out.extend(summary[cur_no]["body"])
            continue
        out.append(ln)
        i += 1
    text = "\n".join(out)
    if names:
        # 배정표의 서브 타이틀 셀 「n) 이름」 만 치환 (구조 불변)
        for old, new in names.items():
            text = re.sub(rf"(\|\s*\d+\)\s*){re.escape(old)}(\s*\|)", rf"\g<1>{new}\2", text)
    return text


def run(cmd: list[str], log: Path) -> int:
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    log.write_text((p.stdout or "") + "\n" + (p.stderr or ""), encoding="utf-8")
    return p.returncode


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("work", type=Path)
    ap.add_argument("name")
    ap.add_argument("--ments", type=Path)
    ap.add_argument("--summary", type=Path)
    ap.add_argument("--conclusion", type=Path)
    ap.add_argument("--toc-names", type=Path)
    ap.add_argument("--facts-args", default="", help="build_facts 추가 인자(--kpi·--exclude-tables 등, 공백 구분)")
    ap.add_argument("--build-args", default="", help="build_guide 추가 인자(--logo·--title·--date 등)")
    ap.add_argument("--out-name", default=None, help="xlsm 파일명(확장자 제외)")
    a = ap.parse_args()

    W = a.work.resolve()
    V = W / "variants" / a.name
    V.mkdir(parents=True, exist_ok=True)
    O = W.parent / "out" / "variants" / a.name
    O.mkdir(parents=True, exist_ok=True)

    summary = json.loads(a.summary.read_text(encoding="utf-8")) if a.summary else None
    conclusion = json.loads(a.conclusion.read_text(encoding="utf-8")) if a.conclusion else None
    names = json.loads(a.toc_names.read_text(encoding="utf-8")) if a.toc_names else None
    toc = patch_toc((W / "목차.md").read_text(encoding="utf-8"), summary, conclusion, names)
    (V / "목차.md").write_text(toc, encoding="utf-8", newline="\n")
    ments_path = V / "멘트.json"
    if a.ments:
        shutil.copy(a.ments, ments_path)
    else:
        shutil.copy(W / "멘트_확정.json", ments_path)
    for f in ("KPI요약표_직접생성.json", "조사개요.json", "summary_map.json"):
        if (W / f).exists():
            shutil.copy(W / f, V / f)

    import shlex
    facts_cmd = [sys.executable, str(TOOLS / "build_facts.py"), str(W / "questionnaire.json"), str(W / "tables.json"),
                 str(V / "slide_facts.json"), "--toc", str(V / "목차.md"), "--ments", str(ments_path)] + shlex.split(a.facts_args)
    rc = run(facts_cmd, V / "log_build_facts.txt")
    print(f"[{a.name}] build_facts rc={rc}")
    if rc:
        print((V / "log_build_facts.txt").read_text(encoding="utf-8")[-1500:]); return rc

    out_name = a.out_name or f"비상주_제작가이드_{a.name}"
    xlsm = O / f"{out_name}.xlsm"
    build_cmd = [sys.executable, str(TOOLS / "build_guide.py"), str(V / "slide_facts.json"), str(xlsm), "--toc", str(V / "목차.md"),
                 "--tables", str(W / "tables.json"), "--overview", str(W / "조사개요.json")] + shlex.split(a.build_args)
    rc = run(build_cmd, V / "log_build_guide.txt")
    print(f"[{a.name}] build_guide rc={rc} → {xlsm}")
    if rc:
        print((V / "log_build_guide.txt").read_text(encoding="utf-8")[-1500:]); return rc

    send = xlsm.with_name(xlsm.stem + "_제작팀송부본" + xlsm.suffix)
    ver_cmd = [sys.executable, str(TOOLS / "verify.py"), str(send if send.exists() else xlsm), str(W / "tables.json"),
               "--facts", str(V / "slide_facts.json"), "--questionnaire", str(W / "questionnaire.json"), "--toc", str(V / "목차.md"),
               "--json", str(V / "verify_guide.json")]
    rc_v = run(ver_cmd, V / "log_verify_guide.txt")
    ment_cmd = [sys.executable, str(TOOLS / "verify.py"), str(ments_path), "--mode", "ments", "--facts", str(W / "slide_facts.json"),
                "--json", str(V / "verify_ments.json")]
    rc_m = run(ment_cmd, V / "log_verify_ments.txt")
    print(f"[{a.name}] verify guide rc={rc_v} · ments rc={rc_m}  (variants/{a.name}/verify_*.json)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
