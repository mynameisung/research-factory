#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""학습 루프 Layer 0 — 사용자가 이 폴더에서 입력한 지시를 **조사 프로젝트별로** 남긴다.

Claude Code 의 `UserPromptSubmit` 훅으로 붙는다. 사용자가 무엇을 쳤든 그 전문을
활성 프로젝트의 `_memory/사용자입력.jsonl` 에 한 줄로 덧붙인다.

왜 값진가: `/회고` 의 관찰 소스 표에서 「사용자가 대화에서 직접 지적」은 ★★★ 인데
지금까지는 「파일에 안 남는다 — 반드시 물어서 받는다」였다. 이 훅이 그 구멍을 메운다.
그리고 연구원마다 맡는 조사가 다르므로, 기억은 리포 전체가 아니라 **조사 프로젝트 폴더**에
쌓여야 다음 세션·다음 담당이 그 조사의 맥락을 이어받는다.

어디에 남기나
    report-factory/projects/<조사>/_memory/사용자입력.jsonl   원문 (한 줄 = 입력 하나)
    report-factory/projects/<조사>/_memory/MEMORY.md         스킬·회고가 다듬어 쓰는 정리본
    활성 프로젝트가 없으면 report-factory/projects/_공통/_memory/ 에 남긴다.

활성 프로젝트는 `report-factory/projects/.active` 한 줄(프로젝트 상대 경로)이 정한다.
스킬이 조사 폴더 인자를 받으면 `capture_prompt.py set-active <경로>` 로 적어 둔다.
`projects/` 는 .gitignore 라 이 기록은 리포에 올라가지 않는다 — 고객 조사의 맥락은 로컬에만 남는다.

훅이므로 지켜야 할 것 (capture_edit.py 와 같다):
- 절대 실패하지 않는다. 무슨 일이 있어도 exit 0.
- stdout 에 아무것도 쓰지 않는다 (훅 출력은 모델 컨텍스트에 들어간다).
- 빠르다.

사용:
    (훅)      python capture_prompt.py                 stdin 의 훅 JSON 을 읽는다
    (스킬)    python capture_prompt.py set-active 2026_에스원_건물현황실사/02_비상주
    (스킬)    python capture_prompt.py active           활성 프로젝트 경로를 출력 (없으면 빈 줄)
    (회고)    python capture_prompt.py tail [N]         활성 프로젝트의 최근 입력 N개를 출력
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]          # report-factory/
PROJECTS = ROOT / "projects"
ACTIVE_FILE = PROJECTS / ".active"
COMMON = "_공통"
MEMORY_TEMPLATE = """# 이 조사의 기억

이 파일은 이 조사 폴더에서만 쓰는 정리본이다. `/제작가이드`·`/보고서작성`·`/회고` 가
시작할 때 읽고, 끝날 때 갱신한다. 원문 입력은 `사용자입력.jsonl` 에 자동으로 쌓인다.

## 결정된 것
(사용자가 정한 것. 날짜와 함께 한 줄씩)

## 지시의 요지
(`사용자입력.jsonl` 에서 반복되거나 방향을 바꾼 지시를 추려 적는다)

## 다음 세션이 알아야 할 것
(어디까지 됐고, 무엇이 대기 중인가)
"""


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _active_project() -> Path:
    try:
        rel = ACTIVE_FILE.read_text(encoding="utf-8").strip()
        if rel:
            p = (PROJECTS / rel).resolve()
            if p.is_dir() and PROJECTS.resolve() in p.parents:
                return p
    except Exception:
        pass
    return PROJECTS / COMMON


def _memory_dir(project: Path) -> Path:
    d = project / "_memory"
    d.mkdir(parents=True, exist_ok=True)
    mem = d / "MEMORY.md"
    if not mem.exists():
        mem.write_text(MEMORY_TEMPLATE, encoding="utf-8")
    return d


def hook() -> int:
    try:
        raw = sys.stdin.read()
        data = json.loads(raw) if raw.strip() else {}
    except Exception:
        return 0
    prompt = str(data.get("prompt") or data.get("user_prompt") or "").strip()
    if not prompt:
        return 0
    try:
        project = _active_project()
        d = _memory_dir(project)
        rec = {
            "ts": _now(),
            "session": str(data.get("session_id") or "")[:12],
            "project": project.name if project.name != COMMON else None,
            "chars": len(prompt),
            "prompt": prompt,
        }
        with (d / "사용자입력.jsonl").open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except Exception:
        pass
    return 0


def set_active(rel: str) -> int:
    rel = rel.strip().replace("\\", "/")
    # 절대 경로나 report-factory/projects/ 접두가 와도 받는다
    for prefix in (str(PROJECTS).replace("\\", "/") + "/", "report-factory/projects/", "projects/"):
        if rel.startswith(prefix):
            rel = rel[len(prefix):]
    target = (PROJECTS / rel).resolve()
    if not (target.is_dir() and PROJECTS.resolve() in target.parents):
        print(f"활성 프로젝트로 쓸 수 없는 경로: {rel}", file=sys.stderr)
        return 1
    PROJECTS.mkdir(parents=True, exist_ok=True)
    ACTIVE_FILE.write_text(rel + "\n", encoding="utf-8")
    _memory_dir(target)
    print(rel)
    return 0


def tail(n: int) -> int:
    d = _active_project() / "_memory" / "사용자입력.jsonl"
    if not d.exists():
        return 0
    lines = d.read_text(encoding="utf-8").splitlines()[-n:]
    for ln in lines:
        try:
            r = json.loads(ln)
            print(f"[{r.get('ts','')[:16]}] {r.get('prompt','')}")
        except Exception:
            continue
    return 0


def main(argv: list[str]) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    if not argv:
        return hook()
    cmd = argv[0]
    if cmd == "set-active" and len(argv) > 1:
        return set_active(argv[1])
    if cmd == "active":
        p = _active_project()
        print("" if p.name == COMMON else p.relative_to(PROJECTS).as_posix())
        return 0
    if cmd == "tail":
        try:
            return tail(int(argv[1]) if len(argv) > 1 else 20)
        except Exception:
            return 0
    print(__doc__.strip().splitlines()[0], file=sys.stderr)
    return 2


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except SystemExit:
        raise
    except Exception:
        raise SystemExit(0)   # 훅은 절대 실패하지 않는다
