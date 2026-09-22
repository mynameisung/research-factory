#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""release_notes.py — 판마다 「업데이트 내역」 한 장. 링크를 만들고, 빠진 노트를 잡고, GitHub Release 로 올린다.

CHANGELOG.md 는 한 줄짜리 「바뀐 점」이다. 연구원이 업그레이드를 끝냈을 때 링크를 눌러 **무엇이 달라졌고 · 무엇을 새로
써볼 수 있고 · 그 기능이 무엇인지** 를 읽는 자리는 `releases/<판>.md` 한 장이고, 같은 내용이 GitHub Release 페이지에 올라간다.

    <리포 루트>/releases/
        _TEMPLATE.md      새 노트의 틀 (절 이름이 고정이다 — check 가 본다)
        0.5.1.md          판마다 한 장. 파일 이름 = VERSION

부명령 (어느 것도 작업 폴더의 다른 파일을 건드리지 않는다):
    url [<판>] [--from <이전판>]   업데이트 내역 링크를 `KEY: 값` 줄로 낸다. 판을 안 주면 VERSION.
                                   --from 을 주면 건너뛴 중간 판의 링크도 함께 낸다 (0.4.1 → 0.5.1 이면 0.5.0 도)
    show [<판>] [--remote]         노트 본문을 그대로 출력. --remote 는 origin/main 의 것을 읽는다 (아직 안 받았을 때)
    check [<판>]                   노트가 있고 절이 갖춰졌는지. 없으면 exit 1 (setup 스모크 · 판 올릴 때 선행 확인)
    new <판>                       틀에서 releases/<판>.md 를 만든다 (이미 있으면 그대로 두고 exit 1)
    publish [<판>] [--target <커밋>]  gh 로 GitHub Release(태그 v<판>)를 만든다. 이미 있으면 본문만 갱신한다. 관리자 전용
    -h · --help

링크 규칙: origin 주소에서 `https://github.com/<owner>/<repo>/releases/tag/v<판>` 을 만든다. 태그가 아직 없어도 주소는 같다
(publish 가 그 태그를 만든다). 옛 주소(FineHourTeam/report-factory)는 GitHub 이 새 주소로 넘긴다.

환경변수(시험용): RF_REPO_ROOT(리포 루트) · RF_REMOTE_REF(기본 origin/main)
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

REQUIRED_SECTIONS = ("달라진 점", "새로 써볼 수 있는 것", "기능 설명")
RELEASES_DIRNAME = "releases"
TEMPLATE_NAME = "_TEMPLATE.md"


def _utf8() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass


def _root() -> Path:
    import update_check as uc
    return uc.repo_root()


def _version() -> str:
    import update_check as uc
    return uc.read_text(_root() / "VERSION")


def _releases() -> Path:
    return _root() / RELEASES_DIRNAME


def _note_path(version: str) -> Path:
    return _releases() / f"{version}.md"


def _git(*args: str, timeout: int = 15) -> subprocess.CompletedProcess:
    return subprocess.run(["git", "-C", str(_root()), *args], capture_output=True, text=True,
                          timeout=timeout, encoding="utf-8", errors="replace")


def parse_version(v: str) -> tuple[int, ...]:
    import update_check as uc
    return uc.parse_version(v)


# ─────────────────────────────────────────────────────────────
# 링크
# ─────────────────────────────────────────────────────────────

def repo_web_url() -> str | None:
    """origin → https://github.com/<owner>/<repo>. https·ssh 두 꼴 다 읽는다."""
    r = _git("remote", "get-url", "origin")
    if r.returncode != 0:
        return None
    url = r.stdout.strip()
    m = re.match(r"^(?:https?://|git@|ssh://git@)([^/:]+)[/:]([^/]+)/([^/]+?)(?:\.git)?/?$", url)
    if not m:
        return None
    host, owner, repo = m.groups()
    return f"https://{host}/{owner}/{repo}"


def release_url(version: str) -> str | None:
    base = repo_web_url()
    return f"{base}/releases/tag/v{version}" if base else None


def releases_index_url() -> str | None:
    base = repo_web_url()
    return f"{base}/releases" if base else None


def local_versions() -> list[str]:
    d = _releases()
    if not d.is_dir():
        return []
    vers = [p.stem for p in d.glob("*.md") if not p.name.startswith("_")]
    return sorted(vers, key=parse_version)


def cmd_url(version: str, since: str | None) -> int:
    url = release_url(version)
    if not url:
        print("RELEASE_NOTES: NONE (origin 주소를 읽지 못함)")
        return 0
    print(f"RELEASE_NOTES: {url}")
    p = _note_path(version)
    print(f"RELEASE_NOTES_FILE: {p.relative_to(_root()).as_posix() if p.exists() else 'NONE'}")
    if since:
        lo, hi = parse_version(since), parse_version(version)
        skipped = [v for v in local_versions() if lo < parse_version(v) < hi]
        if skipped:
            print("RELEASE_NOTES_SKIPPED: " + " ".join(f"{v}={release_url(v)}" for v in skipped))
            idx = releases_index_url()
            if idx:
                print(f"RELEASE_NOTES_ALL: {idx}")
    return 0


# ─────────────────────────────────────────────────────────────
# show · check · new
# ─────────────────────────────────────────────────────────────

def read_note(version: str, remote: bool = False) -> str | None:
    if remote:
        import update_check as uc
        uc.fetch_remote(_root())          # 실패해도 캐시된 origin/main 으로 읽는다
        text = uc.remote_file(_root(), f"{RELEASES_DIRNAME}/{version}.md")
        if text:
            return text
    p = _note_path(version)
    return p.read_text(encoding="utf-8") if p.exists() else None


def cmd_show(version: str, remote: bool) -> int:
    text = read_note(version, remote)
    if text is None:
        print(f"(releases/{version}.md 가 없음)")
        return 1
    print(text.rstrip("\n"))
    return 0


def check_note(version: str) -> list[str]:
    """문제 목록. 비어 있으면 통과."""
    p = _note_path(version)
    if not p.exists():
        return [f"releases/{version}.md 가 없음  → python research-factory/engine/tools/release_notes.py new {version}"]
    text = p.read_text(encoding="utf-8")
    problems: list[str] = []
    first = next((ln for ln in text.splitlines() if ln.strip()), "")
    if not re.match(rf"^#\s+.*\b{re.escape(version)}\b", first):
        problems.append(f"첫 줄 제목에 판 {version} 이 없음: {first[:60]!r}")
    heads = {m.group(1).strip() for m in re.finditer(r"^##\s+(.+?)\s*$", text, re.M)}
    for sec in REQUIRED_SECTIONS:
        if sec not in heads:
            problems.append(f"절 「## {sec}」 이 없음")
    if "<!-- " in text and "-->" in text and re.search(r"<!--\s*(여기에|예:)", text):
        problems.append("틀의 안내 주석이 남아 있음 (채운 뒤 지운다)")
    return problems


def cmd_check(version: str) -> int:
    problems = check_note(version)
    if problems:
        print(f"업데이트 내역 미비 ({version}):")
        for pr in problems:
            print("  - " + pr)
        return 1
    print(f"업데이트 내역 확인 (releases/{version}.md · 절 {len(REQUIRED_SECTIONS)}개)")
    return 0


def cmd_new(version: str) -> int:
    p = _note_path(version)
    if p.exists():
        print(f"이미 있음: {p.relative_to(_root()).as_posix()}")
        return 1
    tpl = _releases() / TEMPLATE_NAME
    if not tpl.exists():
        print(f"틀이 없음: releases/{TEMPLATE_NAME}")
        return 1
    body = tpl.read_text(encoding="utf-8").replace("{{VERSION}}", version).replace("{{DATE}}", date.today().isoformat())
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(body, encoding="utf-8")
    print(f"만듦: {p.relative_to(_root()).as_posix()}  (절을 채운 뒤 check)")
    return 0


# ─────────────────────────────────────────────────────────────
# publish — GitHub Release
# ─────────────────────────────────────────────────────────────

def cmd_publish(version: str, target: str | None) -> int:
    problems = check_note(version)
    if problems:
        cmd_check(version)
        return 1
    p = _note_path(version)
    tag = f"v{version}"
    title = f"research-factory {version}"
    try:
        exists = subprocess.run(["gh", "release", "view", tag, "--json", "tagName"], capture_output=True,
                                text=True, cwd=str(_root()), encoding="utf-8", errors="replace").returncode == 0
    except FileNotFoundError:
        print("gh 명령을 찾을 수 없음 — GitHub CLI 를 설치하고 `gh auth login` 뒤 다시")
        return 1
    if exists:
        args = ["gh", "release", "edit", tag, "--title", title, "--notes-file", str(p)]
    else:
        args = ["gh", "release", "create", tag, "--title", title, "--notes-file", str(p), "--latest"]
        if target:
            args += ["--target", target]
    r = subprocess.run(args, capture_output=True, text=True, cwd=str(_root()), encoding="utf-8", errors="replace")
    if r.returncode != 0:
        print("gh 실패: " + (r.stderr or r.stdout).strip())
        return 1
    print(("갱신: " if exists else "만듦: ") + (release_url(version) or tag))
    return 0


# ─────────────────────────────────────────────────────────────
# 진입점
# ─────────────────────────────────────────────────────────────

def usage() -> int:
    print(__doc__.strip().splitlines()[0])
    print("사용: release_notes.py [url [<판>] [--from <이전판>] | show [<판>] [--remote] | check [<판>] | "
          "new <판> | publish [<판>] [--target <커밋>]]")
    return 2


def main(argv: list[str]) -> int:
    _utf8()
    args = list(argv)
    cmd = args[0] if args else "url"
    if cmd in ("-h", "--help"):
        usage()
        return 0
    rest = args[1:]
    opts: dict[str, str | bool] = {}
    pos: list[str] = []
    i = 0
    while i < len(rest):
        a = rest[i]
        if a in ("--from", "--target"):
            if i + 1 >= len(rest) or rest[i + 1].startswith("-"):
                return usage()
            opts[a] = rest[i + 1]
            i += 2
            continue
        if a == "--remote":
            opts[a] = True
        elif a.startswith("-"):
            return usage()
        else:
            pos.append(a)
        i += 1
    try:
        if cmd == "new":
            return cmd_new(pos[0]) if pos else usage()
        version = pos[0] if pos else _version()
        if not version:
            print("VERSION 파일이 없음")
            return 1
        if cmd == "url":
            return cmd_url(version, opts.get("--from"))          # type: ignore[arg-type]
        if cmd == "show":
            return cmd_show(version, bool(opts.get("--remote")))
        if cmd == "check":
            return cmd_check(version)
        if cmd == "publish":
            return cmd_publish(version, opts.get("--target"))    # type: ignore[arg-type]
        return usage()
    except Exception as e:
        print(f"실패 {type(e).__name__}: {e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
