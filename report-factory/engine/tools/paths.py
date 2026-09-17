#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""paths.py — 조사 폴더의 「역할 → 경로」 해석기. 스킬과 도구는 폴더 이름을 추측하지 않고 이것을 부른다.

조사 폴더 배치(v2)는 README 1절 업무 흐름의 단계 번호를 그대로 쓴다. 폴더는 우리가 하는 단계에만 있고,
다른 팀(온라인실사팀·DP·제작팀)의 산출물은 그것을 받는 우리 단계 폴더에 들어온다.

    projects/<연도_클라이언트_프로젝트>/
    ├─ README.md · ORIENTATION.md
    ├─ 00_협의/            1 클라이언트 협의 · 2 To-Know-List (RFP · 제안서 · 조사개요 · 회의록 · ToKnowList)
    └─ <NN_조사명>/
        ├─ 03_설문지/       3 설문지 구성 (설문지 .docx 판별) · 4 온라인 구성 요청/회신
        ├─ 05_웹테스트/     5 웹테스트 — work/ (시나리오 · 덤프 · 캡처 · 세션) · out/ (findings · 수정요청서)
        ├─ 06_실사/         6 설문 실사 — 실사 현황 · 완료 통보 · 원자료
        ├─ 07_테이블/       7 테이블 가이드 작성 · 8 DP 통계표 전달본 · 전처리본
        ├─ 09_제작가이드/   9 제작가이드 작성 — work/ (정본 JSON · 목차.md) · out/ (.xlsm · 검수리포트 · .baseline/)
        ├─ 11_보고서/       10 제작팀 회신본(in/) · 11 보고서 완료 — work/ · out/ (최종 .pptx · 검수리포트)
        ├─ 12_납품/         12 납품본 · 납품 기록
        ├─ _반송대장/       가이드 팀 ↔ PPT 팀 반송
        └─ _memory/         MEMORY.md · 사용자입력 · 편집기록 · 스킬실행

옛 배치(v1: input/ · 1_제작가이드 · 2_보고서작성 · 3_웹테스트 · 00_공통)도 같은 역할 이름으로 풀어 준다.

프로젝트 하나에 조사가 여럿이면 조사 폴더를 여럿 둔다 (`01_TA풀` · `02_비상주`). **트래킹 조사는 회차마다 조사 폴더 하나**이고
이름 끝에 회차를 붙인다 (`01_고객만족도_2025` · `02_고객만족도_2026` · `03_고객만족도_2027`, 또는 `_1차`·`_2차`).
그러면 `tracking.prev` 에 이전 회차의 경로(통계표·최종보고서·제작가이드 out …)가 함께 풀려 전차 비교에 바로 쓴다.
설문지·통계표가 한 조사 안에 여러 벌이면(유형 A/B 등) `files_all` 에 전부 나온다.

사용
    python paths.py <조사 폴더>                     역할 → 경로 JSON (layout · 폴더 · 찾은 파일)
    python paths.py <조사 폴더> --role 설문지        경로 한 줄
    python paths.py new <프로젝트>/<조사>            v2 뼈대 생성 (프로젝트 README·ORIENTATION 은 템플릿에서, 없을 때만)
    python paths.py migrate <조사 폴더> [--apply]    v1 → v2 이름 바꾸기. --apply 없으면 계획만 보인다
"""
from __future__ import annotations

import json
import re
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
PROJECTS = ROOT / "report-factory" / "projects"
TEMPLATE = HERE.parent / "templates" / "TEMPLATE_프로젝트"

V2 = {"협의": "00_협의", "설문지_dir": "03_설문지", "웹테스트": "05_웹테스트", "실사": "06_실사", "테이블_dir": "07_테이블",
      "제작가이드": "09_제작가이드", "보고서": "11_보고서", "납품": "12_납품", "반송대장": "_반송대장", "memory": "_memory"}
V1 = {"협의": "00_공통", "설문지_dir": "input", "웹테스트": "3_웹테스트", "실사": "input", "테이블_dir": "input",
      "제작가이드": "1_제작가이드", "보고서": "2_보고서작성", "납품": "2_보고서작성/out", "반송대장": "_반송대장", "memory": "_memory"}
SUBDIRS = {"설문지_dir": ["work"], "웹테스트": ["work", "out"], "제작가이드": ["work", "out"], "보고서": ["in", "work", "out"]}
PROJECT_LEVEL = {"협의"}
FILE_ROLES = {   # 역할: (폴더 역할, 글롭 패턴들 — 앞의 것이 우선)
    "설문지": ("설문지_dir", ["*설문지*.docx", "*설문*.docx"]),
    "구조화": ("설문지_dir", ["*구조화*.xlsx", "*구조*.xlsx"]),
    "설문명세": ("설문지_dir", ["work/설문명세.json"]),
    "통계표": ("테이블_dir", ["*통계표*.xlsx", "*통계표*.xls", "*.xlsx", "*.xls"]),
    "테이블가이드": ("테이블_dir", ["*테이블가이드*.xlsx", "*테이블*가이드*.xls*"]),
    "원자료": ("실사", ["*원자료*.xlsx", "*원자료*.xls", "*raw*.xlsx", "*raw*.csv"]),
    "조사개요": ("협의", ["*조사개요*.docx", "*조사개요*.pdf"]),
    "RFP": ("협의", ["*RFP*", "*입찰*", "*제안요청*"]),
    "ToKnowList": ("협의", ["*ToKnow*", "*To_Know*", "*알고싶은*"]),
    "회의록": ("협의", ["*회의록*"]),
    "questionnaire_json": ("제작가이드", ["work/questionnaire.json"]),
    "tables_json": ("제작가이드", ["work/tables.json"]),
    "slide_facts": ("제작가이드", ["work/slide_facts.json"]),
    "목차": ("제작가이드", ["work/목차.md"]),
    "제작가이드_out": ("제작가이드", ["out/*제작팀송부본*.xlsm", "out/*.xlsm"]),
    "제작본": ("보고서", ["in/*.pptx"]),
    "최종보고서": ("보고서", ["out/*최종*.pptx", "out/*.pptx"]),
    "시나리오": ("웹테스트", ["work/시나리오.json"]),
    "수정요청서": ("웹테스트", ["out/수정요청서_*.xlsx"]),
}
RE_INPUT_SPLIT = [  # v1 input/ 파일을 v2 로 나누는 규칙 (앞의 것이 우선)
    (re.compile(r"설문", re.I), "설문지_dir"),
    (re.compile(r"원자료|raw", re.I), "실사"),
    (re.compile(r"통계표|테이블|table|전처리", re.I), "테이블_dir"),
]


def _utf8() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass


def resolve_survey(arg: str) -> Path:
    """`2026_에스원/02_비상주` · `projects/…` · `report-factory/projects/…` · 절대 경로 어느 꼴이든 조사 폴더로."""
    s = arg.strip().replace("\\", "/").rstrip("/")
    for prefix in (str(PROJECTS).replace("\\", "/") + "/", "report-factory/projects/", "projects/"):
        if s.startswith(prefix):
            s = s[len(prefix):]
    p = Path(s)
    return p.resolve() if p.is_absolute() else (PROJECTS / s).resolve()


def detect_layout(survey: Path) -> str:
    if any((survey / V2[k]).is_dir() for k in ("제작가이드", "웹테스트", "설문지_dir", "테이블_dir")):
        return "v2"
    if any((survey / V1[k]).is_dir() for k in ("제작가이드", "보고서", "웹테스트")) or (survey / "input").is_dir():
        return "v1"
    return "none"


def _latest(paths: list[Path]) -> Path | None:
    if not paths:
        return None
    return sorted(paths, key=lambda p: (p.name, p.stat().st_mtime))[-1]


RE_SURVEY = re.compile(r"^(\d{2})_(.+?)(?:_([^_]+))?$")   # NN_조사약칭[_회차]
MULTI_ROLES = ("설문지", "구조화", "통계표", "테이블가이드", "제작본")   # 여러 벌일 수 있는 것은 전부 목록으로도 준다


def parse_survey_name(name: str) -> tuple[int | None, str, str | None]:
    """`02_고객만족도_2026` → (2, '고객만족도', '2026'). 회차가 없으면 (NN, 이름, None). 규약 밖이면 (None, name, None)."""
    m = RE_SURVEY.match(name)
    if not m:
        return None, name, None
    nn, base, wave = m.groups()
    return int(nn), base, wave


def tracking_info(survey: Path, _deep: bool = True) -> dict | None:
    """트래킹 조사(같은 조사약칭 + 회차 꼬리표가 붙은 형제 폴더)면 회차 목록과 이전 회차의 경로를 준다."""
    nn, base, wave = parse_survey_name(survey.name)
    if nn is None or wave is None:
        return None
    waves = []
    for sib in sorted(survey.parent.iterdir()):
        if not sib.is_dir():
            continue
        snn, sbase, swave = parse_survey_name(sib.name)
        if snn is not None and swave is not None and sbase == base:
            waves.append((snn, sib))
    prev = max((s for n, s in waves if n < nn), key=lambda s: parse_survey_name(s.name)[0], default=None)
    info = {"base": base, "wave": wave, "waves": [s.name for _, s in waves], "prev": None}
    if prev is not None and _deep:
        p = resolve(prev, _deep=False)
        info["prev"] = {"name": p["name"], "survey": p["survey"], "dirs": p["dirs"], "files": p["files"]}
    return info


def resolve(survey: Path, _deep: bool = True) -> dict:
    layout = detect_layout(survey)
    table = V2 if layout != "v1" else V1
    project = survey.parent
    dirs: dict[str, str] = {}
    for role, rel in table.items():
        base = project if role in PROJECT_LEVEL else survey
        dirs[role] = (base / rel).as_posix()
    files: dict[str, str | None] = {}
    files_all: dict[str, list[str]] = {}
    for role, (dir_role, pats) in FILE_ROLES.items():
        base = Path(dirs[dir_role])
        hit = None
        for pat in pats:
            cands = [p for p in base.glob(pat) if p.is_file() and not p.name.startswith("~$")]
            hit = _latest(cands)
            if hit:
                if role in MULTI_ROLES:
                    files_all[role] = [c.as_posix() for c in sorted(cands, key=lambda p: p.name)]
                break
        files[role] = hit.as_posix() if hit else None
    out = {"layout": layout, "name": f"{project.name}/{survey.name}", "project": project.as_posix(), "survey": survey.as_posix(),
           "dirs": dirs, "files": files, "files_all": files_all}
    if _deep:
        out["tracking"] = tracking_info(survey)
    return out


def cmd_new(rel: str) -> int:
    survey = resolve_survey(rel)
    project = survey.parent
    if not PROJECTS.resolve() in survey.parents or survey.parent == PROJECTS.resolve():
        print("형식: new <프로젝트>/<조사>  예) new 2026_KB손보_자동차CSI/01_본조사", file=sys.stderr); return 2
    project.mkdir(parents=True, exist_ok=True)
    (project / V2["협의"]).mkdir(exist_ok=True)
    for name in ("README.md", "ORIENTATION.md"):
        if not (project / name).exists() and (TEMPLATE / name).exists():
            shutil.copy(TEMPLATE / name, project / name)
    for role, rel_dir in V2.items():
        if role in PROJECT_LEVEL:
            continue
        d = survey / rel_dir
        d.mkdir(parents=True, exist_ok=True)
        for sub in SUBDIRS.get(role, []):
            (d / sub).mkdir(exist_ok=True)
    print(json.dumps(resolve(survey), ensure_ascii=False, indent=1))
    return 0


def cmd_migrate(rel: str, apply: bool) -> int:
    survey = resolve_survey(rel)
    layout = detect_layout(survey)
    if layout != "v1":
        print(f"{survey.name}: 배치가 {layout} 라 옮길 것이 없다"); return 0
    plan: list[tuple[Path, Path]] = []
    for role in ("제작가이드", "보고서", "웹테스트"):
        src = survey / V1[role]
        if src.is_dir():
            plan.append((src, survey / V2[role]))
    inp = survey / "input"
    if inp.is_dir():
        for f in sorted(inp.iterdir()):
            if f.name == ".gitkeep":
                continue
            dest_role = next((r for rx, r in RE_INPUT_SPLIT if rx.search(f.name)), None)
            if dest_role:
                plan.append((f, survey / V2[dest_role] / f.name))
            else:
                print(f"  ? 어느 단계인지 모름 — input/ 에 남긴다: {f.name}")
    common = survey.parent / V1["협의"]
    if common.is_dir() and not (survey.parent / V2["협의"]).exists():
        plan.append((common, survey.parent / V2["협의"]))
    for src, dst in plan:
        print(f"  {src.relative_to(survey.parent).as_posix()}  →  {dst.relative_to(survey.parent).as_posix()}")
    if not apply:
        print(f"계획 {len(plan)}건. 실제로 옮기려면 --apply. 옮긴 뒤 조사 폴더 안 스크립트(work/*.py)의 옛 경로는 손으로 고친다"); return 0
    for src, dst in plan:
        if dst.exists():
            print(f"  ! 이미 있음, 건너뜀: {dst.name}"); continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        src.rename(dst)
    for role, rel_dir in V2.items():   # 빈 단계 폴더도 만들어 둔다
        if role in PROJECT_LEVEL:
            continue
        d = survey / rel_dir
        d.mkdir(exist_ok=True)
        for sub in SUBDIRS.get(role, []):
            (d / sub).mkdir(exist_ok=True)
    if inp.is_dir() and not any(p for p in inp.iterdir() if p.name != ".gitkeep"):
        shutil.rmtree(inp)
    print(f"옮김 {len(plan)}건 → {detect_layout(survey)}")
    return 0


def main(argv: list[str]) -> int:
    _utf8()
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__.strip()); return 0
    if argv[0] == "new" and len(argv) > 1:
        return cmd_new(argv[1])
    if argv[0] == "migrate" and len(argv) > 1:
        return cmd_migrate(argv[1], "--apply" in argv)
    survey = resolve_survey(argv[0])
    if not survey.is_dir():
        print(f"조사 폴더가 없다: {survey}", file=sys.stderr); return 2
    info = resolve(survey)
    if "--role" in argv:
        role = argv[argv.index("--role") + 1]
        v = info["dirs"].get(role) or info["files"].get(role)
        if v is None:
            print(f"역할 {role} 에 해당하는 것이 없다 (폴더 역할: {', '.join(info['dirs'])} / 파일 역할: {', '.join(info['files'])})", file=sys.stderr); return 1
        print(v); return 0
    print(json.dumps(info, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
