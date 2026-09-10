#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""update_check.py — 판 확인. gstack 의 bin/gstack-update-check 를 report-factory 에 옮긴 것.

스킬(/제작가이드·/보고서작성·/회고)이 시작할 때 이 스크립트를 먼저 부른다.
**작업 폴더는 건드리지 않는다.** `git fetch`(읽기 전용)로 origin/main 의 VERSION 한 줄만
읽고, pull·merge·checkout 은 하지 않는다. 실제로 올리는 일은 /업그레이드 가 한다.

출력 (표준출력 한 줄, 또는 침묵):
    UPGRADE_AVAILABLE <내판> <원격판>   원격 VERSION 이 더 높다
    JUST_UPGRADED <이전판> <내판>       직전에 /업그레이드 가 끝났다 (표식 파일이 있었다)
    CHECK_FAILED <이유>                 확인 자체가 실패했다 (오프라인 등)
    (없음)                              최신 · 미룸 중 · 끔 · 캐시 유효

침묵이 곧 「최신」 신호이므로, 실패는 반드시 CHECK_FAILED 로 소리 낸다. gstack 은 이 구분이
없던 시절에 45개 판이 지나도록 아무도 낡은 줄 모르는 사고를 겪었다.

부명령:
    check [--force]                  기본. --force 는 캐시·미룸을 무시하고 원격을 다시 본다
    snooze <원격판>                  「나중에」. 1단계 24시간 → 2단계 48시간 → 3단계부터 7일
    mark --from <이전판>             업그레이드 직후 표식. 다음 check 가 JUST_UPGRADED 를 낸다
    config get <키> | set <키> <값>  update_check(true/false) · auto_upgrade(true/false)
    whats-new <판>                   원격 CHANGELOG 에서 그 판의 요약 줄을 뽑아 출력

상태는 리포 밖 `~/.report-factory/` 에 둔다 — 업그레이드로 코드가 통째로 갈려도 살아남는다.
    last-update-check   캐시. 「최신」이면 60분, 「새 판 있음」이면 12시간 동안 원격을 다시 안 본다
    update-snoozed      미룸. `<판> <단계> <시각>`. 그 사이 더 새 판이 나오면 미룸을 무시한다
    just-upgraded-from  표식. /업그레이드 가 남기고 check 가 지운다
    config.json         설정. {"update_check": true, "auto_upgrade": false}

환경변수(시험용): RF_STATE_DIR(상태 폴더) · RF_REPO_ROOT(리포 루트) · RF_REMOTE_REF(기본 origin/main)
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path

CACHE_TTL_UP_TO_DATE_MIN = 60
CACHE_TTL_UPGRADE_MIN = 720
SNOOZE_SECONDS = {1: 86_400, 2: 172_800}   # 3단계부터는 7일
SNOOZE_MAX_SECONDS = 604_800
FETCH_TIMEOUT_SEC = 20
DEFAULT_CONFIG = {"update_check": True, "auto_upgrade": False}


def _utf8_stdout() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass


def repo_root() -> Path:
    env = os.environ.get("RF_REPO_ROOT")
    if env:
        return Path(env)
    here = Path(__file__).resolve()
    # engine/tools/update_check.py → engine/tools → engine → report-factory → 리포 루트
    guess = here.parents[3]
    if (guess / "VERSION").exists():
        return guess
    try:
        out = subprocess.run(["git", "-C", str(here.parent), "rev-parse", "--show-toplevel"],
                             capture_output=True, text=True, timeout=10)
        if out.returncode == 0 and out.stdout.strip():
            return Path(out.stdout.strip())
    except Exception:
        pass
    return guess


def state_dir() -> Path:
    d = Path(os.environ.get("RF_STATE_DIR") or (Path.home() / ".report-factory"))
    d.mkdir(parents=True, exist_ok=True)
    return d


def read_text(p: Path) -> str:
    try:
        return p.read_text(encoding="utf-8").strip()
    except Exception:
        return ""


def parse_version(v: str) -> tuple[int, ...]:
    nums = re.findall(r"\d+", v)
    return tuple(int(n) for n in nums) if nums else (0,)


def load_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    p = state_dir() / "config.json"
    try:
        cfg.update(json.loads(p.read_text(encoding="utf-8")))
    except Exception:
        pass
    return cfg


def save_config(cfg: dict) -> None:
    (state_dir() / "config.json").write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")


def to_bool(v: str) -> bool:
    return str(v).strip().lower() in ("1", "true", "yes", "on", "y")


# ─────────────────────────────────────────────────────────────
# 원격 읽기 — fetch 는 origin/main 포인터만 갱신한다. 작업 폴더·현재 브랜치는 그대로다.
# ─────────────────────────────────────────────────────────────

def remote_ref() -> str:
    return os.environ.get("RF_REMOTE_REF", "origin/main")


def git(root: Path, *args: str, timeout: int = 15) -> subprocess.CompletedProcess:
    return subprocess.run(["git", "-C", str(root), *args], capture_output=True, text=True,
                          timeout=timeout, encoding="utf-8", errors="replace")


def fetch_remote(root: Path) -> str | None:
    """원격 갱신. 실패 사유 문자열을 돌려주고, 성공이면 None."""
    remote, _, branch = remote_ref().partition("/")
    try:
        r = git(root, "fetch", "--quiet", "--no-tags", remote, branch, timeout=FETCH_TIMEOUT_SEC)
    except subprocess.TimeoutExpired:
        return f"git fetch 가 {FETCH_TIMEOUT_SEC}초 안에 끝나지 않음 (오프라인?)"
    except FileNotFoundError:
        return "git 명령을 찾을 수 없음"
    if r.returncode != 0:
        reason = (r.stderr or r.stdout).strip().splitlines()
        return "git fetch 실패: " + (reason[-1] if reason else "이유 없음")
    return None


def remote_file(root: Path, relpath: str) -> str | None:
    r = git(root, "show", f"{remote_ref()}:{relpath}")
    if r.returncode != 0:
        return None
    return r.stdout


# ─────────────────────────────────────────────────────────────
# 미룸
# ─────────────────────────────────────────────────────────────

def snoozed(remote_ver: str) -> bool:
    p = state_dir() / "update-snoozed"
    parts = read_text(p).split()
    if len(parts) != 3:
        return False
    ver, level, epoch = parts
    if not (level.isdigit() and epoch.isdigit()):
        return False
    if ver != remote_ver:            # 더 새 판이 나왔으면 미룸을 무시한다
        return False
    lvl = int(level)
    duration = SNOOZE_SECONDS.get(lvl, SNOOZE_MAX_SECONDS)
    return time.time() < int(epoch) + duration


def write_snooze(remote_ver: str) -> int:
    p = state_dir() / "update-snoozed"
    parts = read_text(p).split()
    level = 0
    if len(parts) == 3 and parts[0] == remote_ver and parts[1].isdigit():
        level = int(parts[1])
    level = min(level + 1, 3)
    p.write_text(f"{remote_ver} {level} {int(time.time())}", encoding="utf-8")
    return level


# ─────────────────────────────────────────────────────────────
# check
# ─────────────────────────────────────────────────────────────

def cmd_check(force: bool) -> int:
    cfg = load_config()
    if not cfg.get("update_check", True):
        return 0

    root = repo_root()
    local = read_text(root / "VERSION")
    if not local:
        print("CHECK_FAILED VERSION 파일이 없음")
        return 0

    sd = state_dir()
    cache = sd / "last-update-check"
    snooze = sd / "update-snoozed"
    marker = sd / "just-upgraded-from"

    if force:
        for p in (cache, snooze):
            try:
                p.unlink()
            except FileNotFoundError:
                pass

    # 업그레이드 직후 표식 — 한 번만 알리고 지운다. 그 뒤로도 원격은 다시 본다.
    old = read_text(marker)
    if marker.exists():
        for p in (marker, snooze):
            try:
                p.unlink()
            except FileNotFoundError:
                pass
        if old:
            print(f"JUST_UPGRADED {old} {local}")

    # 캐시가 살아 있으면 원격을 다시 보지 않는다
    parts = read_text(cache).split()
    if len(parts) == 3 and parts[1] == local:
        kind, _, cached_remote = parts
        ttl_min = CACHE_TTL_UP_TO_DATE_MIN if kind == "UP_TO_DATE" else CACHE_TTL_UPGRADE_MIN
        age_min = (time.time() - cache.stat().st_mtime) / 60
        if age_min < ttl_min:
            if kind == "UPGRADE_AVAILABLE" and not snoozed(cached_remote):
                print(f"UPGRADE_AVAILABLE {local} {cached_remote}")
            return 0

    # 느린 길 — 원격을 본다
    err = fetch_remote(root)
    if err:
        print(f"CHECK_FAILED {err}")
        return 0
    remote = (remote_file(root, "VERSION") or "").strip()
    if not remote:
        print(f"CHECK_FAILED {remote_ref()} 에 VERSION 이 없음")
        return 0

    if parse_version(remote) > parse_version(local):
        cache.write_text(f"UPGRADE_AVAILABLE {local} {remote}", encoding="utf-8")
        if not snoozed(remote):
            print(f"UPGRADE_AVAILABLE {local} {remote}")
    else:
        cache.write_text(f"UP_TO_DATE {local} {remote}", encoding="utf-8")
    return 0


# ─────────────────────────────────────────────────────────────
# whats-new — 원격 CHANGELOG 에서 해당 판의 절을 뽑는다
# ─────────────────────────────────────────────────────────────

def cmd_whats_new(version: str, max_lines: int = 6) -> int:
    root = repo_root()
    if fetch_remote(root):
        pass  # 이미 캐시된 origin/main 으로도 읽을 수 있다
    text = remote_file(root, "CHANGELOG.md") or read_text(root / "CHANGELOG.md")
    if not text:
        print("(CHANGELOG 를 읽지 못함)")
        return 0
    lines = text.splitlines()
    start = next((i for i, ln in enumerate(lines)
                  if re.match(rf"^##\s+v?{re.escape(version)}\b", ln.strip())), None)
    if start is None:
        print(f"(CHANGELOG 에 {version} 항목이 없음)")
        return 0
    out: list[str] = []
    for ln in lines[start + 1:]:
        if ln.startswith("## "):
            break
        s = ln.strip()
        if s:
            out.append(s)
        if len(out) >= max_lines:
            break
    print("\n".join(out))
    return 0


# ─────────────────────────────────────────────────────────────
# 진입점
# ─────────────────────────────────────────────────────────────

def usage() -> int:
    print(__doc__.strip().splitlines()[0])
    print("사용: update_check.py [check [--force] | snooze <판> | mark --from <판> | "
          "config get <키> | config set <키> <값> | whats-new <판>]")
    return 2


def main(argv: list[str]) -> int:
    _utf8_stdout()
    args = list(argv)
    cmd = args[0] if args else "check"
    try:
        if cmd == "check":
            return cmd_check(force="--force" in args)
        if cmd == "snooze":
            if len(args) < 2:
                return usage()
            level = write_snooze(args[1])
            hours = {1: "24시간", 2: "48시간"}.get(level, "1주일")
            print(f"SNOOZED {args[1]} level={level} next={hours}")
            return 0
        if cmd == "mark":
            if len(args) < 3 or args[1] != "--from":
                return usage()
            (state_dir() / "just-upgraded-from").write_text(args[2].strip(), encoding="utf-8")
            # 캐시도 비운다 — 다음 check 가 곧바로 원격을 다시 본다
            try:
                (state_dir() / "last-update-check").unlink()
            except FileNotFoundError:
                pass
            print(f"MARKED {args[2].strip()}")
            return 0
        if cmd == "config":
            cfg = load_config()
            if len(args) >= 3 and args[1] == "get":
                v = cfg.get(args[2])
                print("true" if v is True else "false" if v is False else str(v))
                return 0
            if len(args) >= 4 and args[1] == "set":
                key = args[2]
                if key not in DEFAULT_CONFIG:
                    print(f"CHECK_FAILED 모르는 설정 키: {key}")
                    return 0
                cfg[key] = to_bool(args[3])
                save_config(cfg)
                print(f"SET {key}={'true' if cfg[key] else 'false'}")
                return 0
            return usage()
        if cmd == "whats-new":
            if len(args) < 2:
                return usage()
            return cmd_whats_new(args[1])
        return usage()
    except Exception as e:  # 어떤 예외도 침묵으로 끝나면 안 된다
        print(f"CHECK_FAILED {type(e).__name__}: {e}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
