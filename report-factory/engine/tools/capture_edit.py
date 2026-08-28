#!/usr/bin/env python3
"""학습 루프 Layer 1 — 자동 포착.

Claude Code의 PostToolUse 훅으로 붙어서, 사람이 프로젝트 산출물을 손으로 고칠 때마다
그 사실을 `engine/learning/observations.jsonl` 에 남긴다.

이게 왜 값진가: AI가 만든 제작가이드를 사람이 고친 자리는 곧 **AI가 모르던 규칙**이다.
`/회고` 가 이 로그와 `out/.baseline/` 스냅샷 diff를 함께 읽어 규칙 후보를 뽑는다.

훅이므로 지켜야 할 것:
- 절대 실패하지 않는다. 무슨 일이 있어도 exit 0.
- 빠르다. 편집 한 번에 수십 ms를 넘기지 않는다.
- stdout에 아무것도 쓰지 않는다 (훅 출력은 사용자 피드백으로 해석된다).

설정: embrain/.claude/settings.json
    "PostToolUse": [{"matcher": "Edit|Write",
      "hooks": [{"type": "command",
                 "command": "python report-factory/engine/tools/capture_edit.py"}]}]
"""

from __future__ import annotations

import hashlib
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

# report-factory/engine/tools/capture_edit.py → report-factory/
ROOT = Path(__file__).resolve().parents[2]
LOG = ROOT / "engine" / "learning" / "observations.jsonl"

# 이 경로 밑의 편집만 학습 신호로 본다. engine/ 는 우리가 만드는 것이므로 제외.
WATCHED = ("projects",)

# 중간 산출물은 사람이 고치는 대상이 아니다 — 노이즈만 늘린다.
IGNORED_PARTS = {".baseline", "__pycache__", ".git"}
IGNORED_SUFFIXES = {".pyc", ".log", ".tmp"}


def _digest(text: str) -> str:
    """전문 대신 지문만 남긴다. 로그가 비대해지면 아무도 안 읽는다."""
    return hashlib.sha1(text.encode("utf-8", "replace")).hexdigest()[:12]


def _classify(rel: Path) -> str:
    """어떤 종류의 편집인지. `/회고` 가 이걸로 관찰을 묶는다."""
    parts = set(rel.parts)
    if "1_제작가이드" in parts:
        return "guide"
    if "2_보고서작성" in parts:
        return "report"
    if "00_공통" in parts or "input" in parts:
        return "input"
    if rel.name in {"ORIENTATION.md", "README.md"}:
        return "orientation"
    return "other"


def main() -> int:
    try:
        # stdin을 바이트로 읽어 UTF-8로 직접 디코드한다.
        # Windows에서 sys.stdin 은 cp949 로 열리기 때문에, 한글이 든 경로가 오면
        # json.load(sys.stdin) 이 UnicodeDecodeError 로 조용히 죽는다.
        payload = json.loads(sys.stdin.buffer.read().decode("utf-8", "replace"))
    except Exception:
        return 0  # 훅 입력이 예상과 달라도 조용히 넘어간다

    try:
        tool = payload.get("tool_name") or payload.get("tool") or ""
        tool_input = payload.get("tool_input") or {}
        raw_path = tool_input.get("file_path") or tool_input.get("path")
        if not raw_path:
            return 0

        path = Path(raw_path)
        try:
            rel = path.resolve().relative_to(ROOT)
        except ValueError:
            return 0  # report-factory 밖의 편집은 우리 관심사가 아니다

        if not rel.parts or rel.parts[0] not in WATCHED:
            return 0
        if IGNORED_PARTS & set(rel.parts) or path.suffix in IGNORED_SUFFIXES:
            return 0

        record: dict[str, object] = {
            "ts": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "tool": tool,
            "path": rel.as_posix(),
            "kind": _classify(rel),
        }

        # Edit 은 무엇이 무엇으로 바뀌었는지가 곧 규칙의 단서다.
        old, new = tool_input.get("old_string"), tool_input.get("new_string")
        if old is not None and new is not None:
            record["change"] = "edit"
            record["old"] = old[:400]
            record["new"] = new[:400]
            record["truncated"] = len(old) > 400 or len(new) > 400
        else:
            content = tool_input.get("content")
            record["change"] = "write"
            if content is not None:
                record["bytes"] = len(content)
                record["digest"] = _digest(content)

        LOG.parent.mkdir(parents=True, exist_ok=True)
        with LOG.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception:
        pass  # 학습 로깅이 실패해도 사용자의 작업을 막지 않는다

    return 0


if __name__ == "__main__":
    sys.exit(main())
