# -*- coding: utf-8 -*-
"""
clova_draft.py — 하이브리드 멘트 초안: CLOVA 가 쓰고, 게이트가 거르고, 반려분은 Claude 담당이 고친다.

사용자 지시(2026-08-23): 「클로바를 워크플로에 넣되 숫자가 틀리거나 하면 네(Claude)가 수정하는 방식」.
요약·결론은 대상이 아니다(5차 실측 — 요약 층 탈락·근거 창작). 본문 멘트만.

  python engine/tools/clova_draft.py <work_dir> [--model HCX-007] [--reasoning low] [--max-slides 6]

하는 일
  1. work/샤드/S*.json 을 grid 없는 슬림 반샤드(≤ max-slides 장)로 쪼개 CLOVA 에 보낸다(출력 자체 절단 회피 — 실측).
  2. 반환을 합쳐 `verify --mode ments` 게이트에 건다.
  3. 수용분 → work/검수/clova_draft_accepted.json ({ments:[{seq,lines}]})
     반려분 → work/샤드/재발주_clova.json ({slides:[{seq, 반려줄:[{line,reasons}], 슬라이스…}]}) — guide-body-ment 에게 그대로 준다.
  4. 둘을 합친 결과가 멘트_확정.json 의 후보가 된다(메인이 Claude 재작성분을 게이트 후 병합).

실측 메모: 비상주 69장에서 CLOVA HCX-007 단독 수용은 3/69 였다(보기명 의역·수치 따옴표·마침표). 하이브리드의 값어치는
「수용분만큼 Claude 호출을 아낀다」인데 그 수가 작으면 사실상 Claude 단독과 같다 — 수용률을 로그로 남겨 판단한다.
키: 환경변수 CLOVA_STUDIO_API_KEY (저장소·대화에 넣지 않는다). PowerShell 에서 User 스코프 변수를 프로세스에 물려 실행한다.
"""
from __future__ import annotations
import argparse, json, os, subprocess, sys
from pathlib import Path

TOOLS = Path(__file__).resolve().parent
AGENTS = TOOLS.parents[2] / ".claude" / "agents"

TPL = ("각 줄 ≤60자 · 개조식 ㅁ받침 종결(~임/~함/~순임) · 마침표 없음 · 보기명은 ment_facts 의 label 원문을 작은따옴표로(의역·축약 금지) · "
       "수치는 display 문자열 그대로 + '%' (수치에 따옴표 금지) · 1줄 = 「{주제}는 '{1위}'이 {v}%로 가장 높고, '{2위}'({v}%) 등의 순임」 · "
       "척도는 「{항목} 만족도는 {top}%(TOP2 기준)로 {수준어} 수준임」(90 매우 우수/80 우수/70 긍정적/60 보통/50 다소 낮은/미만 낮은) · "
       "순위 문항은 「(1~N순위 기준)」 한정구 · 2줄 = line2 에 있는 배너만 「{group}별로 보면 {label}의 '{row}' 응답이 상대적으로 {direction}」 · "
       "line2 가 비면 2줄 없음 · 장당 3줄 이하 · 문항 번호(DQ1. 등)를 문장에 넣지 않음 · 모든 seq 에 ment")


def slim(shard: dict, max_slides: int) -> list[dict]:
    out = []
    slides = []
    for s in shard["slides"]:
        s = {k: v for k, v in s.items() if k != "grid"}
        slides.append(s)
    for i in range(0, len(slides), max_slides):
        out.append({"shard": shard.get("shard"), "slides": slides[i:i + max_slides], "요구": TPL})
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("work", type=Path)
    ap.add_argument("--model", default="HCX-007")
    ap.add_argument("--reasoning", default="low")
    ap.add_argument("--max-slides", type=int, default=6)
    a = ap.parse_args()
    if not os.environ.get("CLOVA_STUDIO_API_KEY"):
        sys.exit("CLOVA_STUDIO_API_KEY 가 없습니다 — PowerShell: $env:CLOVA_STUDIO_API_KEY=[Environment]::GetEnvironmentVariable('CLOVA_STUDIO_API_KEY','User')")
    W = a.work.resolve()
    tmp = W / "검수" / "_clova_draft"; tmp.mkdir(parents=True, exist_ok=True)
    env = {**os.environ, "PYTHONUTF8": "1"}
    merged: dict[int, list[str]] = {}
    for sp in sorted((W / "샤드").glob("S*.json")):
        shard = json.loads(sp.read_text(encoding="utf-8"))
        for k, part in enumerate(slim(shard, a.max_slides)):
            pin = tmp / f"{sp.stem}_{k}.json"; pout = tmp / f"{sp.stem}_{k}.out.json"
            pin.write_text(json.dumps(part, ensure_ascii=False, indent=1), encoding="utf-8")
            cmd = [sys.executable, str(TOOLS / "llm_write.py"), "--provider", "clova", "--model", a.model, "--task", "ments",
                   "--reasoning", a.reasoning, "--agent", str(AGENTS / "guide-body-ment.md"), "--input", str(pin), "--out", str(pout)]
            r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)
            if r.returncode or not pout.exists():
                print(f"[clova] {pin.name} 실패: {(r.stderr or '')[-200:]}")
                continue
            for s in json.loads(pout.read_text(encoding="utf-8")).get("slides", []):
                if s.get("ment"):
                    merged[int(s["seq"])] = [str(x).strip() for x in s["ment"] if str(x).strip()]
            print(f"[clova] {pin.name} → {len(merged)} slides 누적")
    draft = tmp / "clova_ments_all.json"
    draft.write_text(json.dumps({"ments": [{"seq": k, "lines": v} for k, v in sorted(merged.items())]}, ensure_ascii=False, indent=1), encoding="utf-8")
    gate = tmp / "gate.json"
    subprocess.run([sys.executable, str(TOOLS / "verify.py"), str(draft), "--mode", "ments", "--facts", str(W / "slide_facts.json"), "--json", str(gate)],
                   capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)
    g = json.loads(gate.read_text(encoding="utf-8"))
    acc = {int(k): v for k, v in (g.get("accepted") or {}).items()}
    rej = {int(k): v for k, v in (g.get("rejected") or {}).items()}
    facts = {s["seq"]: s for s in json.loads((W / "slide_facts.json").read_text(encoding="utf-8"))["slides"]}
    missing = [seq for seq in facts if seq not in acc and seq not in rej]
    (W / "검수" / "clova_draft_accepted.json").write_text(
        json.dumps({"ments": [{"seq": k, "lines": v} for k, v in sorted(acc.items())]}, ensure_ascii=False, indent=1), encoding="utf-8")
    reissue = {"지시": "CLOVA 초안이 게이트에 반려된 장이다. 반려 사유를 보고 ment_facts 로 접지해 다시 쓴다 — 수치·보기명은 슬라이스의 것만. 빠진 장(누락)은 처음부터 쓴다.",
               "slides": []}
    for seq in sorted(set(rej) | set(missing)):
        s = {k: v for k, v in facts[seq].items() if k != "grid"}
        s["반려줄"] = rej.get(seq, [{"line": "", "reasons": ["CLOVA 누락"]}])
        reissue["slides"].append(s)
    (W / "샤드" / "재발주_clova.json").write_text(json.dumps(reissue, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\nCLOVA 초안 {len(merged)}/{len(facts)}장 · 게이트 수용 {len(acc)} · 반려 {len(rej)} · 누락 {len(missing)}")
    print(f"→ 수용분 {W / '검수' / 'clova_draft_accepted.json'}\n→ Claude 재작성 패킷 {W / '샤드' / '재발주_clova.json'} (guide-body-ment 에 배정)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
