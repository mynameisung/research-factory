#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""migrate_folder.py — 0.5.0 폴더 이름 변경(report-factory/ → research-factory/)의 뒷정리.

git 은 추적 파일만 옮긴다. git 밖에 있던 것 — 조사 데이터(projects/) · 코퍼스(engine/refs/) ·
인수인계 문서 · 개인 로그 · 아직 추적되지 않은 새 파일 — 은 옛 폴더에 남으므로 이 도구가 새 폴더로 옮긴다.

- 멱등이다. 옮길 것이 없으면 아무것도 하지 않는다. 훅(capture_prompt) · 프리앰블 · setup · /업그레이드 가
  매번 불러도 된다. 두 프로세스가 동시에 부르면 한쪽은 조용히 물러난다(잠금 파일).
- 내용이 다른 파일은 지우지 않는다. 새 폴더 같은 자리에 다른 파일이 이미 있으면 옛 쪽을 그대로 두고 보고한다.
  똑같은 파일이면 옛 사본만 지우고, 기록 파일(`*.jsonl`)은 옛 기록을 앞에 이어 붙인다.
- 조사 데이터(projects/)가 아직 옛 폴더에 있으면(열린 파일 때문에 못 옮겼으면) `blocked()` 가 참이다.
  훅은 그동안 기록하지 않는다 — 새 폴더에 따로 쌓이면 기록이 둘로 갈라진다.
- 옛 경로가 계속 돌게 한다(옛 폴더에서 실제로 옮긴 것이 있었던 PC 만 — 새로 설치한 PC 는 깨끗하게 둔다). 옛 자리 `report-factory/projects` 와 `report-factory/engine/<도구 밖 폴더>` 에
  새 폴더를 가리키는 연결(Windows 정션)을 걸고, `report-factory/engine/tools/` 에는 도구마다 대리 파일을 만든다
  (실행하면 새 경로의 같은 도구를 돌리고, import 하면 새 모듈로 바꿔 끼운다). 옛 절대경로를 쓰는 조사 스크립트가
  그대로 돈다. 연결과 대리 파일은 .gitignore 가 옛 폴더 전체를 무시하므로 git 에 보이지 않는다.
- 정리할 때 연결 안으로 들어가지 않는다. 연결 너머는 새 폴더의 진짜 데이터다.
- git 이 추적하는 대리 파일 6개(훅·판 확인 등)는 건드리지 않는다. 거기에 로컬 수정이 섞였으면(업그레이드의
  stash pop 충돌) 보고한다 — 대리 파일은 되돌리고 수정은 새 경로의 같은 파일에 다시 적용해야 한다.
- 원격 주소가 옛 주소(github.com/FineHourTeam/report-factory)면 새 주소로 바꾼다.

사용:
    python migrate_folder.py            옮기고 요약
    python migrate_folder.py --check    옮길 것이 있는지만 본다 (exit 1 = 남은 것 있음)
훅·프리앰블은 run() · notice() 를 부른다 — 무슨 일이 있어도 예외를 올리지 않는다.
"""
from __future__ import annotations

import filecmp
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import time
from pathlib import Path

OLD_NAME = "report-factory"
NEW_NAME = "research-factory"
NEW_REMOTE = "https://github.com/mynameisung/research-factory.git"
RE_OLD_REMOTE = re.compile(r"github\.com[/:]FineHourTeam/report-factory(?:\.git)?/?\s*$", re.I)

HERE = Path(__file__).resolve().parent            # <리포>/research-factory/engine/tools
REPO = HERE.parents[2]
OLD = REPO / OLD_NAME
NEW = REPO / NEW_NAME
TOOLS = Path("engine", "tools")

# git 이 추적하는 대리 파일 — 업그레이드 직후 이관 전에도 옛 경로 훅·도구가 돌게 한다
STUBS = ("capture_prompt.py", "capture_edit.py", "update_check.py", "verify.py", "preamble.py", "gen_preamble.py")
HOOKS = ("capture_prompt.py", "capture_edit.py")
SHIM_MARK = "전환용 대리 파일 (0.5.0)"
LOCK, REPORT, NOTICE, DONE = ".이관잠금", ".이관보고", ".이관알림", ".이관함"   # 옛 폴더에 둔다(.gitignore 가 무시한다)
MARKS = {LOCK, REPORT, NOTICE, DONE}


# ─────────────────────────────────────────────────────────────
# 대리 파일
# ─────────────────────────────────────────────────────────────

def shim_source(name: str) -> str:
    """옛 경로 대리 파일의 내용. 추적하는 6개와 이 도구가 만드는 나머지가 같은 틀을 쓴다."""
    hook = name in HOOKS
    head = f'''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""{SHIM_MARK} — 이 폴더는 research-factory/ 로 옮겨졌다. 진짜는 research-factory/engine/tools/{name}.

업그레이드 전에 열린 Claude 세션과 옛 절대경로를 쓰는 조사 스크립트가 이 경로를 부른다.
실행하면 새 경로의 같은 도구를 돌리고, import 하면 새 모듈로 바꿔 끼운다. 가는 길에 옛 폴더에 남은 데이터를
새 폴더로 옮긴다(migrate_folder). 만든 곳: research-factory/engine/tools/migrate_folder.py — 여기서 고치지 않는다.
"""
import runpy
import sys
from pathlib import Path

_OLD = Path(__file__).resolve().parent
_NEW = _OLD.parents[2] / "research-factory" / "engine" / "tools"
_REAL = _NEW / "{name}"
'''
    main_hook = '''
if __name__ == "__main__":
    # 훅은 어떤 경우에도 exit 0 — UserPromptSubmit 훅이 2로 끝나면 사용자 입력이 막힌다
    sys.path[:] = [str(_NEW)] + [p for p in sys.path if Path(p or ".").resolve() != _OLD]
    sys.argv[0] = str(_REAL)
    _blocked = False
    try:
        import migrate_folder
        _res = migrate_folder.run()
        _blocked = migrate_folder.blocked()
''' + ('''        _line = migrate_folder.notice(_res)
        if _line:                # 옮긴 일이 있을 때 한 번 — 훅 출력은 Claude 의 맥락이 된다
            sys.stdout.reconfigure(encoding="utf-8")
            print(_line)
''' if name == "capture_prompt.py" else '') + '''    except Exception:
        pass
    try:
        if not _blocked:         # 조사 데이터가 아직 옛 폴더에 있으면 기록을 건너뛴다(갈라지지 않게)
            runpy.run_path(str(_REAL), run_name="__main__")
    except BaseException:
        pass
    sys.exit(0)
'''
    main_tool = '''
if __name__ == "__main__":
    sys.path[:] = [str(_NEW)] + [p for p in sys.path if Path(p or ".").resolve() != _OLD]
    sys.argv[0] = str(_REAL)
    try:
        import migrate_folder
        migrate_folder.run()
    except Exception:
        pass
''' + ('''    try:
        runpy.run_path(str(_REAL), run_name="__main__")
    finally:
        if len(sys.argv) > 1 and sys.argv[1] == "mark":   # 옛 /업그레이드 가 부른다 — 옛 스킬 절차로 계속 가지 않게
            sys.stdout.reconfigure(encoding="utf-8")
            print("[research-factory] 0.5.0 부터 폴더가 research-factory/ 로 바뀌었다. 이 세션이 읽어 둔 스킬 절차는 "
                  "옛 경로를 가리키므로, 업그레이드 보고를 마치면 본 절차로 가지 말고 사용자에게 Claude 세션을 닫고 "
                  "새로 열어 달라고 한다.")
''' if name == "update_check.py" else '''    runpy.run_path(str(_REAL), run_name="__main__")
''')
    as_module = '''else:
    # 옛 경로를 sys.path 에 넣고 import 하는 조사 스크립트 — 새 모듈로 바꿔 끼운다
    import importlib.util
    if str(_NEW) not in sys.path:
        sys.path.insert(0, str(_NEW))
    _spec = importlib.util.spec_from_file_location(__name__, _REAL)
    _mod = importlib.util.module_from_spec(_spec)
    sys.modules[__name__] = _mod
    _spec.loader.exec_module(_mod)
'''
    return head + (main_hook if hook else main_tool) + as_module


def _is_shim(p: Path) -> bool:
    try:
        with p.open("r", encoding="utf-8", errors="replace") as fh:
            return SHIM_MARK in fh.read(400)
    except OSError:
        return False


# ─────────────────────────────────────────────────────────────
# 판정
# ─────────────────────────────────────────────────────────────

def _is_link(p: Path) -> bool:
    """심볼릭 링크나 Windows 정션(재분석 지점)인가. 따라 들어가지 않는다."""
    try:
        st = os.lstat(p)
    except OSError:
        return False
    if stat.S_ISLNK(st.st_mode):
        return True
    return bool(getattr(st, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))


def _same_target(a: Path, b: Path) -> bool:
    try:
        return os.path.normcase(os.path.realpath(a)) == os.path.normcase(os.path.realpath(b))
    except OSError:
        return False


def _plan(src: Path, dst: Path, rel: Path, out: list[tuple[str, Path, Path | None]]) -> None:
    try:
        entries = sorted(os.scandir(src), key=lambda e: e.name)
    except OSError:
        return
    for e in entries:
        r = rel / e.name
        p = Path(e.path)
        d = dst / e.name
        if rel == Path() and e.name in MARKS:
            continue
        if _is_link(p):
            if _same_target(p, d):
                continue                                   # 우리가 건 연결
            out.append(("move" if not os.path.lexists(d) else "conflict", p, d))   # 사용자가 건 연결 — 연결째 옮긴다
            continue
        if rel == TOOLS and (e.name in STUBS or (e.is_file() and _is_shim(p))):
            continue                                       # 대리 파일
        if e.name == "__pycache__":
            if rel != TOOLS:
                out.append(("drop", p, None))              # 옛 도구 폴더의 것은 대리 파일이 다시 만든다 — 그대로 둔다
            continue
        if not os.path.lexists(d):
            out.append(("move", p, d))
        elif e.is_dir(follow_symlinks=False) and d.is_dir() and not _is_link(d):
            _plan(p, d, r, out)
        elif e.is_file(follow_symlinks=False) and d.is_file() and filecmp.cmp(p, d, shallow=False):
            out.append(("same", p, d))
        elif e.is_file(follow_symlinks=False) and d.is_file() and p.suffix == ".jsonl":
            out.append(("append", p, d))
        else:
            out.append(("conflict", p, d))


def _ready() -> bool:
    return OLD.is_dir() and not _is_link(OLD) and NEW.is_dir()


def pending() -> list[tuple[str, Path, Path | None]]:
    if not _ready():
        return []
    out: list[tuple[str, Path, Path | None]] = []
    _plan(OLD, NEW, Path(), out)
    return out


def blocked() -> bool:
    """조사 데이터가 아직 옛 폴더에 진짜 폴더로 남아 있는가."""
    old = OLD / "projects"
    return _ready() and old.is_dir() and not _is_link(old)


def _stub_trouble() -> list[str]:
    """추적하는 대리 파일에 로컬 수정이 섞였는가(업그레이드 stash pop 충돌 등)."""
    bad = []
    for n in STUBS:
        p = OLD / TOOLS / n
        try:
            t = p.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        if SHIM_MARK not in t or "<<<<<<<" in t:
            bad.append(n)
    return bad


# ─────────────────────────────────────────────────────────────
# 실행
# ─────────────────────────────────────────────────────────────

def _prune_empty(root: Path) -> None:
    """옛 폴더의 빈 폴더를 아래부터 지운다. 연결 안으로는 들어가지 않는다(os.walk 는 Windows 정션을 따라간다)."""
    dirs: list[Path] = []
    for dirpath, dirnames, _ in os.walk(root, topdown=True):
        dirnames[:] = [d for d in dirnames if not _is_link(Path(dirpath, d))]
        dirs.append(Path(dirpath))
    for p in reversed(dirs):
        if p == root or p == root / TOOLS:
            continue
        try:
            if not any(p.iterdir()):
                p.rmdir()
        except OSError:
            pass


def _unlink(p: Path) -> None:
    try:
        p.unlink()
    except PermissionError:
        os.chmod(p, stat.S_IWRITE)          # 읽기 전용 파일
        p.unlink()


def _make_link(link: Path, target: Path) -> bool:
    if os.path.lexists(link) or not target.is_dir():
        return False
    link.parent.mkdir(parents=True, exist_ok=True)
    try:
        if os.name == "nt":
            import _winapi
            _winapi.CreateJunction(str(target), str(link))
        else:
            os.symlink(target, link, target_is_directory=True)
        return True
    except Exception:
        try:
            r = subprocess.run(["cmd", "/c", "mklink", "/J", str(link), str(target)],
                               capture_output=True, timeout=20)
            return r.returncode == 0
        except Exception:
            return False


def _bridge() -> None:
    """옛 경로가 계속 돌게 연결과 대리 파일을 둔다. 조용히 한다(알릴 일이 아니다)."""
    targets = [Path("projects")]
    try:
        targets += [Path("engine", e.name) for e in os.scandir(NEW / "engine")
                    if e.is_dir() and e.name not in ("tools", "__pycache__")]
        targets += [TOOLS / e.name for e in os.scandir(NEW / TOOLS)
                    if e.is_dir() and e.name != "__pycache__"]
    except OSError:
        pass
    for rel in targets:
        _make_link(OLD / rel, NEW / rel)
    tools = OLD / TOOLS
    if not tools.is_dir() or _is_link(tools):
        return
    try:
        names = [e.name for e in os.scandir(NEW / TOOLS) if e.is_file() and e.name.endswith(".py")]
    except OSError:
        return
    for n in names:
        p = tools / n
        if n in STUBS or os.path.lexists(p):
            continue
        try:
            p.write_text(shim_source(n), encoding="utf-8", newline="\n")
        except OSError:
            pass


def _fix_remote() -> bool:
    """.git/config 를 직접 읽는다(훅에서 git 을 띄우지 않으려고). 워크트리(.git 이 파일)는 건너뛴다."""
    try:
        text = (REPO / ".git" / "config").read_text(encoding="utf-8")
    except OSError:
        return False
    if "FineHourTeam/report-factory" not in text:
        return False
    try:
        url = subprocess.run(["git", "-C", str(REPO), "config", "--get", "remote.origin.url"],
                             capture_output=True, text=True, timeout=10).stdout.strip()
        if RE_OLD_REMOTE.search(url):
            subprocess.run(["git", "-C", str(REPO), "remote", "set-url", "origin", NEW_REMOTE],
                           capture_output=True, timeout=10)
            return True
    except Exception:
        pass
    return False


def _lock() -> int | None:
    p = OLD / LOCK
    for _ in range(2):
        try:
            return os.open(p, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                if time.time() - p.stat().st_mtime > 600:      # 죽은 프로세스가 남긴 잠금
                    p.unlink()
                    continue
            except OSError:
                pass
            return None
        except OSError:
            return None
    return None


def run() -> dict:
    """옮기고 결과를 돌려준다. 예외를 올리지 않는다."""
    res: dict = {"moved": [], "same": 0, "dropped": 0, "conflict": [], "locked": [], "denied": [],
                 "remote": False, "stubs": []}
    if not _ready():
        return res
    fd = _lock()
    if fd is None:
        res["busy"] = True                                  # 다른 프로세스가 옮기는 중
        return res
    try:
        plan = pending()
        for kind, p, d in plan:
            rel = p.relative_to(OLD).as_posix()
            try:
                if kind == "move":
                    d.parent.mkdir(parents=True, exist_ok=True)
                    os.rename(p, d)
                    res["moved"].append(rel)
                elif kind == "same":
                    _unlink(p)
                    res["same"] += 1
                elif kind == "append":
                    old_b, new_b = p.read_bytes(), d.read_bytes()
                    if old_b and not old_b.endswith(b"\n"):
                        old_b += b"\n"
                    d.write_bytes(old_b + new_b)
                    _unlink(p)
                    res["moved"].append(rel + "(이어 붙임)")
                elif kind == "drop":
                    shutil.rmtree(p, ignore_errors=True)
                    res["dropped"] += 1
                else:
                    res["conflict"].append(rel)
            except OSError as e:
                # 32·33 = 다른 프로그램이 파일을 열고 있음. 폴더 이름 바꾸기는 안에 열린 파일이 있으면 5(거부)로 온다.
                # 그 밖의 거부(권한 등)는 파일을 닫아도 풀리지 않으므로 따로 알린다
                opened = getattr(e, "winerror", None) in (32, 33) or (isinstance(e, PermissionError) and p.is_dir())
                res["locked" if opened else "denied"].append(rel)
        if plan:
            _prune_empty(OLD)
        if res["moved"]:
            (OLD / DONE).touch()                            # 이 PC 에는 옛 폴더 데이터가 있었다
        if not blocked() and (OLD / DONE).exists():        # 새로 설치한 PC 에는 연결·대리 파일을 두지 않는다
            _bridge()
        res["remote"] = _fix_remote()
        res["stubs"] = _stub_trouble()
        if res["moved"] or res["remote"]:
            _report(res)
    except Exception as e:                              # noqa: BLE001 — 훅에서 불린다
        res["error"] = f"{type(e).__name__}: {e}"
    finally:
        os.close(fd)
        try:
            (OLD / LOCK).unlink()
        except OSError:
            pass
    return res


# ─────────────────────────────────────────────────────────────
# 알림
# ─────────────────────────────────────────────────────────────

def _events(res: dict) -> list[str]:
    bits = []
    if res.get("moved"):
        m = res["moved"]
        bits.append(f"옮김 {len(m)}건({', '.join(m[:5])}{' …' if len(m) > 5 else ''}) — 옛 경로는 연결로 그대로 쓴다")
    if res.get("remote"):
        bits.append(f"원격 주소 → {NEW_REMOTE}")
    return bits


def _problems(res: dict) -> list[str]:
    bits = []
    if res.get("locked"):
        bits.append(f"열린 파일 때문에 못 옮김 {len(res['locked'])}건({', '.join(res['locked'][:3])}) — "
                    "엑셀 등을 닫으면 다음에 다시 옮긴다")
    if res.get("denied"):
        bits.append(f"권한 문제로 못 옮김 {len(res['denied'])}건({', '.join(res['denied'][:3])}) — 파일 속성·권한을 확인한다")
    if res.get("conflict"):
        bits.append(f"같은 자리에 다른 파일이 있어 옛 쪽을 남김 {len(res['conflict'])}건"
                    f"({', '.join(res['conflict'][:3])}) — 사람이 비교해 정한다")
    if res.get("stubs"):
        n = res["stubs"]
        bits.append(f"옛 경로 대리 파일에 로컬 수정이 섞임({', '.join(n)}) — `git checkout HEAD -- "
                    f"{OLD_NAME}/engine/tools/{n[0]}` 로 되돌리고 수정은 {NEW_NAME}/engine/tools/ 의 같은 파일에 다시 적용한다")
    if blocked() and not res.get("locked"):
        bits.append("조사 데이터(projects/)가 아직 옛 폴더에 있다")
    if res.get("error"):
        bits.append(f"오류 {res['error']}")
    return bits


def _line(bits: list[str]) -> str:
    return f"[{NEW_NAME}] 폴더 이관({OLD_NAME}/ → {NEW_NAME}/): " + " · ".join(bits) if bits else ""


def _report(res: dict) -> None:
    """옮긴 일을 적어 둔다. 조용히 부른 쪽(편집 훅·도구)이 옮겼어도 다음 입력 훅·프리앰블이 한 번 알린다."""
    try:
        with (OLD / REPORT).open("a", encoding="utf-8") as fh:
            fh.write(" · ".join(_events(res)) + "\n")
    except OSError:
        pass


def summary(res: dict) -> str:
    """이번 실행의 옮긴 일과 남은 문제를 한 줄로. 아무 일도 없으면 빈 문자열."""
    return _line(_events(res) + _problems(res))


def notice(res: dict, repeat_problems: bool = False) -> str:
    """입력 훅·프리앰블용 — 쌓인 옮긴 일은 한 번만, 남은 문제는 지난번과 달라졌을 때만(repeat_problems 면 늘) 알린다."""
    bits: list[str] = []
    try:
        rp = OLD / REPORT
        if rp.exists():
            bits += [ln.strip() for ln in rp.read_text(encoding="utf-8").splitlines() if ln.strip()]
            rp.unlink()
    except OSError:
        pass
    probs = _problems(res)
    if probs and not repeat_problems:
        sig = json.dumps(probs, ensure_ascii=False)
        mark = OLD / NOTICE
        try:
            seen = mark.read_text(encoding="utf-8") == sig
        except OSError:
            seen = False
        if seen:
            probs = []
        elif _ready():
            try:
                mark.write_text(sig, encoding="utf-8")
            except OSError:
                pass
    return _line(bits + probs)


def main(argv: list[str]) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    if "-h" in argv or "--help" in argv:
        print(__doc__.strip())
        return 0
    if "--check" in argv:
        plan = [x for x in pending() if x[0] != "drop"]
        for kind, p, _ in plan:
            print(f"{kind}\t{p.relative_to(OLD).as_posix()}")
        return 1 if plan else 0
    res = run()
    print(notice(res, repeat_problems=True) or "이관할 것 없음")
    left = [x for x in pending() if x[0] in ("move", "conflict")]
    return 1 if left or res.get("stubs") else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
