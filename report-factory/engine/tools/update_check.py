#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""전환용 대리 파일 (0.5.0) — 이 폴더는 research-factory/ 로 옮겨졌다. 진짜는 research-factory/engine/tools/update_check.py.

업그레이드 전에 열린 Claude 세션과 옛 절대경로를 쓰는 조사 스크립트가 이 경로를 부른다.
실행하면 새 경로의 같은 도구를 돌리고, import 하면 새 모듈로 바꿔 끼운다. 가는 길에 옛 폴더에 남은 데이터를
새 폴더로 옮긴다(migrate_folder). 만든 곳: research-factory/engine/tools/migrate_folder.py — 여기서 고치지 않는다.
"""
import runpy
import sys
from pathlib import Path

_OLD = Path(__file__).resolve().parent
_NEW = _OLD.parents[2] / "research-factory" / "engine" / "tools"
_REAL = _NEW / "update_check.py"

if __name__ == "__main__":
    sys.path[:] = [str(_NEW)] + [p for p in sys.path if Path(p or ".").resolve() != _OLD]
    sys.argv[0] = str(_REAL)
    try:
        import migrate_folder
        migrate_folder.run()
    except Exception:
        pass
    try:
        runpy.run_path(str(_REAL), run_name="__main__")
    finally:
        if len(sys.argv) > 1 and sys.argv[1] == "mark":   # 옛 /업그레이드 가 부른다 — 옛 스킬 절차로 계속 가지 않게
            sys.stdout.reconfigure(encoding="utf-8")
            print("[research-factory] 0.5.0 부터 폴더가 research-factory/ 로 바뀌었다. 이 세션이 읽어 둔 스킬 절차는 "
                  "옛 경로를 가리키므로, 업그레이드 보고를 마치면 본 절차로 가지 말고 사용자에게 Claude 세션을 닫고 "
                  "새로 열어 달라고 한다.")
else:
    # 옛 경로를 sys.path 에 넣고 import 하는 조사 스크립트 — 새 모듈로 바꿔 끼운다
    import importlib.util
    if str(_NEW) not in sys.path:
        sys.path.insert(0, str(_NEW))
    _spec = importlib.util.spec_from_file_location(__name__, _REAL)
    _mod = importlib.util.module_from_spec(_spec)
    sys.modules[__name__] = _mod
    _spec.loader.exec_module(_mod)
