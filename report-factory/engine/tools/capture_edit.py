#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""전환용 대리 파일 (0.5.0) — 이 폴더는 research-factory/ 로 옮겨졌다. 진짜는 research-factory/engine/tools/capture_edit.py.

업그레이드 전에 열린 Claude 세션과 옛 절대경로를 쓰는 조사 스크립트가 이 경로를 부른다.
실행하면 새 경로의 같은 도구를 돌리고, import 하면 새 모듈로 바꿔 끼운다. 가는 길에 옛 폴더에 남은 데이터를
새 폴더로 옮긴다(migrate_folder). 만든 곳: research-factory/engine/tools/migrate_folder.py — 여기서 고치지 않는다.
"""
import runpy
import sys
from pathlib import Path

_OLD = Path(__file__).resolve().parent
_NEW = _OLD.parents[2] / "research-factory" / "engine" / "tools"
_REAL = _NEW / "capture_edit.py"

if __name__ == "__main__":
    # 훅은 어떤 경우에도 exit 0 — UserPromptSubmit 훅이 2로 끝나면 사용자 입력이 막힌다
    sys.path[:] = [str(_NEW)] + [p for p in sys.path if Path(p or ".").resolve() != _OLD]
    sys.argv[0] = str(_REAL)
    _blocked = False
    try:
        import migrate_folder
        _res = migrate_folder.run()
        _blocked = migrate_folder.blocked()
    except Exception:
        pass
    try:
        if not _blocked:         # 조사 데이터가 아직 옛 폴더에 있으면 기록을 건너뛴다(갈라지지 않게)
            runpy.run_path(str(_REAL), run_name="__main__")
    except BaseException:
        pass
    sys.exit(0)
else:
    # 옛 경로를 sys.path 에 넣고 import 하는 조사 스크립트 — 새 모듈로 바꿔 끼운다
    import importlib.util
    if str(_NEW) not in sys.path:
        sys.path.insert(0, str(_NEW))
    _spec = importlib.util.spec_from_file_location(__name__, _REAL)
    _mod = importlib.util.module_from_spec(_spec)
    sys.modules[__name__] = _mod
    _spec.loader.exec_module(_mod)
