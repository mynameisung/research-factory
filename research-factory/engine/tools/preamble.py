#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""preamble.py — 모든 작업 스킬이 본 절차보다 먼저 한 번 부르는 진입점 (gstack 의 preamble 을 옮긴 것).

한 번의 호출로 ① 새 판 확인(update_check) ② 활성 조사 지정(capture_prompt set-active) ③ 조사 메모리 위치
④ 스킬 실행 기록을 처리하고, 기계가 읽기 쉬운 `KEY: 값` 줄로만 답한다. 스킬 본문은 이 줄들을 보고 움직인다.

사용: python preamble.py --skill <스킬이름> [<조사 폴더>]
      조사 폴더는 `2026_에스원/02_비상주` · `projects/…` · `research-factory/projects/…` · 절대 경로 어느 꼴이든 된다.

출력 (항상 이 순서, 항상 exit 0 — 프리앰블은 스킬을 멈추지 않는다)
    SKILL: 제작가이드
    VERSION: 0.3.1
    UPDATE: NONE | UPGRADE_AVAILABLE <내판> <새판> | JUST_UPGRADED <이전판> <내판> | CHECK_FAILED <이유> | DISABLED
    MIGRATE: <폴더 이관 요약>     ← 0.5.0 폴더 이름 변경 뒷정리에서 옮긴 것·남은 것이 있을 때만
    ACTIVE: <프로젝트>/<조사> | NONE
    MEMORY: <MEMORY.md 경로> | NONE
    MEMORY_NEW: yes | no          ← 이번에 처음 만들었으면 yes (읽을 것이 없다)
    LAYOUT: v2 | v1 | none        ← 조사 폴더 배치 (paths.py). v1 은 옛 이름(input/·1_제작가이드…)
    PATHS: {"dirs": {...}, "files": {...}, "files_all": {...}}   ← 역할 → 경로 (한 줄 JSON). 폴더 이름을 추측하지 말고 여기서 찾는다
    TRACKING: <조사약칭> 회차 <회차> (회차 폴더 N개) 이전 회차 <이름|NONE>   ← 트래킹 조사(폴더 이름에 회차 꼬리표)일 때만
    PREV_PATHS: {"dirs": {...}, "files": {...}}   ← 이전 회차의 경로. 전차 비교(통계표·최종보고서·제작가이드)에 쓴다
    PLAYWRIGHT: yes | no          ← 웹테스트 스킬일 때만

원본 블록은 `engine/templates/SKILL_PREAMBLE.md`, 스킬 파일에 채우는 것은 `gen_preamble.py`.
"""
from __future__ import annotations

import contextlib
import io
import json
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))


def _utf8() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass


def _update_line() -> str:
    try:
        import update_check as uc
        if not uc.load_config().get("update_check", True):
            return "DISABLED"
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            uc.cmd_check(force=False)
        lines = [ln.strip() for ln in buf.getvalue().splitlines() if ln.strip()]
        return lines[0] if lines else "NONE"
    except Exception as e:  # 프리앰블은 절대 실패하지 않는다
        return f"CHECK_FAILED {type(e).__name__}: {e}"


def _migrate() -> str:
    """0.5.0 폴더 이름 변경의 뒷정리(migrate_folder). 옮길 것도 남은 것도 없으면 빈 문자열."""
    try:
        import migrate_folder as mf
        return mf.notice(mf.run(), repeat_problems=True)   # 옮긴 일은 한 번만, 남은 문제는 스킬마다
    except Exception:
        return ""


def _version() -> str:
    try:
        import update_check as uc
        return uc.read_text(uc.repo_root() / "VERSION") or "?"
    except Exception:
        return "?"


def _active(arg: str | None) -> tuple[str, Path | None, bool]:
    """(ACTIVE 값, MEMORY.md 경로, 새로 만들었는가)."""
    try:
        import capture_prompt as cp
        if arg:
            mem_before = None
            rel = arg.strip().replace("\\", "/").rstrip("/")
            for prefix in (str(cp.PROJECTS).replace("\\", "/") + "/", str(cp.ROOT.parent / "report-factory" / "projects").replace("\\", "/") + "/",
                           "research-factory/projects/", "report-factory/projects/", "projects/"):
                if rel.startswith(prefix):
                    rel = rel[len(prefix):]
            target = (cp.PROJECTS / rel)
            mem_before = (target / "_memory" / "MEMORY.md").exists()
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
                rc = cp.set_active(rel)
            if rc != 0:
                return "NONE", None, False
            mem = target / "_memory" / "MEMORY.md"
            return rel, mem, (not mem_before and mem.exists())
        rel = cp.ACTIVE_FILE.read_text(encoding="utf-8").strip() if cp.ACTIVE_FILE.exists() else ""
        if not rel:
            return "NONE", None, False
        mem = cp.PROJECTS / rel / "_memory" / "MEMORY.md"
        return rel, (mem if mem.exists() else None), False
    except Exception:
        return "NONE", None, False


def _log(skill: str, arg: str | None, mem: Path | None) -> None:
    if not mem:
        return
    try:
        rec = {"ts": time.strftime("%Y-%m-%dT%H:%M:%S%z"), "skill": skill, "arg": arg or ""}
        with (mem.parent / "스킬실행.jsonl").open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except Exception:
        pass


def _rel_to_root(p: Path) -> str:
    try:
        import update_check as uc
        return p.resolve().relative_to(uc.repo_root().resolve()).as_posix()
    except Exception:
        return p.as_posix()


def main(argv: list[str]) -> int:
    _utf8()
    skill = "?"
    arg = None
    it = iter(argv)
    for a in it:
        if a == "--skill":
            skill = next(it, "?")
        elif not a.startswith("-"):
            arg = a
    print(f"SKILL: {skill}")
    print(f"VERSION: {_version()}")
    print(f"UPDATE: {_update_line()}")
    mig = _migrate()
    if mig:
        print(f"MIGRATE: {mig}")
    active, mem, new = _active(arg)
    print(f"ACTIVE: {active}")
    print(f"MEMORY: {_rel_to_root(mem) if mem else 'NONE'}")
    print(f"MEMORY_NEW: {'yes' if new else 'no'}")
    if active != "NONE":
        try:
            import paths as ps
            info = ps.resolve(ps.resolve_survey(active))
            print(f"LAYOUT: {info['layout']}")
            print("PATHS: " + json.dumps({"dirs": info["dirs"], "files": info["files"], "files_all": info.get("files_all") or {}}, ensure_ascii=False))
            tr = info.get("tracking")
            if tr:
                prev = tr.get("prev") or {}
                print(f"TRACKING: {tr['base']} 회차 {tr['wave']} (회차 폴더 {len(tr['waves'])}개) 이전 회차 {prev.get('name') or 'NONE'}")
                if prev:
                    print("PREV_PATHS: " + json.dumps({"dirs": prev["dirs"], "files": prev["files"]}, ensure_ascii=False))
        except Exception as e:
            print(f"LAYOUT: CHECK_FAILED {type(e).__name__}: {e}")
    if skill == "웹테스트":
        try:
            import playwright  # noqa: F401
            print("PLAYWRIGHT: yes")
        except Exception:
            print("PLAYWRIGHT: no")
    _log(skill, arg, mem)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
