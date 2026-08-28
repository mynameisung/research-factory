#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""외부 LLM(CLOVA Studio 등)에 **문장화 담당**을 맡기는 어댑터.

오케스트레이터(Claude)가 프롬프트를 짜고 → 이 도구가 API 를 쏘고 → 반환 JSON 을 **기존 기계 게이트**
(`verify.py --mode ments|summary|conclusion`)가 수용한다. 모델이 바뀌어도 계약은 같다:
  · 수치는 팩트 슬라이스에 있는 것만 (게이트가 강제) · 문형 규칙(≤60자·ㅁ종결·라벨 원문) · JSON 스키마

사용
----
    set CLOVA_STUDIO_API_KEY=nv-...            (키는 저장소에 넣지 않는다)
    python llm_write.py --provider clova --model HCX-005 --task ments \\
        --agent .claude/agents/guide-body-ment.md --input work/샤드/S1.json --out work/검수/clova_ments_S1.json
    python llm_write.py --provider clova --model HCX-007 --task conclusion --reasoning medium \\
        --agent .claude/agents/guide-conclusion.md --input work/summary_map.json --context work/조사개요.json --out work/검수/clova_conclusion.json
    --dry-run 이면 호출 없이 프롬프트와 스키마만 출력한다.

공급자
------
  clova      OpenAI 호환 엔드포인트 https://clovastudio.stream.ntruss.com/v1/openai  (response_format json_schema · reasoning_effort 지원)
  anthropic  비교군. ANTHROPIC_API_KEY. 같은 프롬프트·스키마로 A/B 한다
"""
from __future__ import annotations

import argparse
import io
import json
import os
import sys
from pathlib import Path

SCHEMAS = {
    "ments": {
        "type": "object",
        "properties": {
            "slides": {"type": "array", "items": {"type": "object", "properties": {
                "seq": {"type": "integer"}, "ment": {"type": "array", "items": {"type": "string"}}},
                "required": ["seq", "ment"]}},
            "decisions": {"type": "array", "items": {"type": "object", "properties": {"seq": {"type": "integer"}, "issue": {"type": "string"}}, "required": ["issue"]}},
            "observations": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["slides", "decisions", "observations"],
    },
    "summary": {
        "type": "object",
        "properties": {
            "summaries": {"type": "array", "items": {"type": "object", "properties": {
                "no": {"type": "integer"}, "title": {"type": "string"}, "headline": {"type": "string"},
                "body": {"type": "array", "items": {"type": "string"}},
                "panels": {"type": "array", "items": {"type": "object", "properties": {
                    "position": {"type": "string"}, "ref": {"type": "string"}, "topic": {"type": "string"},
                    "chart": {"type": "string"}, "base_n": {"type": "integer"}}, "required": ["position", "ref", "topic", "chart", "base_n"]}},
                "claims": {"type": "array", "items": {"type": "object", "properties": {
                    "text": {"type": "string"},
                    "evidence": {"type": "array", "items": {"type": "object", "properties": {
                        "table": {"type": "integer"}, "row_label": {"type": "string"}, "col_label": {"type": "string"}, "value": {"type": "string"}},
                        "required": ["table", "row_label", "col_label", "value"]}}}, "required": ["text", "evidence"]}},
                "source_pages": {"type": "array", "items": {"type": "string"}}},
                "required": ["no", "title", "headline", "body", "panels", "claims"]}},
            "decisions": {"type": "array", "items": {"type": "object"}},
            "observations": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["summaries", "decisions", "observations"],
    },
    "summary_text": {
        "type": "object",
        "properties": {
            "summaries": {"type": "array", "items": {"type": "object", "properties": {
                "no": {"type": "integer"}, "headline": {"type": "string"},
                "body": {"type": "array", "items": {"type": "string"}}},
                "required": ["no", "headline", "body"]}},
            "observations": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["summaries", "observations"],
    },
    "toc": {
        "type": "object",
        "properties": {
            "chapters": {"type": "array", "items": {"type": "object", "properties": {
                "no": {"type": "integer"}, "title": {"type": "string"},
                "subs": {"type": "array", "items": {"type": "object", "properties": {
                    "no": {"type": "integer"}, "title": {"type": "string"},
                    "seqs": {"type": "array", "items": {"type": "integer"}}}, "required": ["no", "title", "seqs"]}}},
                "required": ["no", "title", "subs"]}},
            "excluded": {"type": "array", "items": {"type": "object", "properties": {"code": {"type": "string"}, "reason": {"type": "string"}}, "required": ["code", "reason"]}},
            "rationale": {"type": "array", "items": {"type": "string"}},
            "observations": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["chapters", "excluded", "rationale", "observations"],
    },
    "polish": {
        "type": "object",
        "properties": {
            "items": {"type": "array", "items": {"type": "object", "properties": {
                "id": {"type": "string"}, "text": {"type": "string"}, "changed": {"type": "boolean"}, "note": {"type": "string"}},
                "required": ["id", "text", "changed"]}},
            "observations": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["items", "observations"],
    },
    "toc_names": {
        "type": "object",
        "properties": {
            "names": {"type": "array", "items": {"type": "object", "properties": {
                "id": {"type": "string"}, "current": {"type": "string"},
                "alternatives": {"type": "array", "items": {"type": "string"}},
                "verdict": {"type": "string", "enum": ["유지", "교체 권고"]}, "why": {"type": "string"}},
                "required": ["id", "current", "alternatives", "verdict", "why"]}},
            "observations": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["names", "observations"],
    },
    "storyline": {
        "type": "object",
        "properties": {
            "headers": {"type": "array", "items": {"type": "object", "properties": {
                "no": {"type": "integer"}, "headline": {"type": "string"}, "interpretation": {"type": "string"},
                "link_to_prev": {"type": "string"}}, "required": ["no", "headline", "interpretation"]}},
            "conclusions": {"type": "array", "items": {"type": "object", "properties": {
                "no": {"type": "integer"}, "title": {"type": "string"}, "recommendation": {"type": "string"},
                "from_headers": {"type": "array", "items": {"type": "integer"}}}, "required": ["no", "title", "recommendation", "from_headers"]}},
            "story_in_one_paragraph": {"type": "string"},
            "observations": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["headers", "conclusions", "story_in_one_paragraph", "observations"],
    },
    "conclusion": {
        "type": "object",
        "properties": {
            "layout": {"type": "string"},
            "items": {"type": "array", "items": {"type": "object", "properties": {
                "no": {"type": "integer"}, "title": {"type": "string"},
                "body": {"type": "array", "items": {"type": "string"}},
                "from_summary": {"type": "array", "items": {"type": "integer"}},
                "toknow": {"type": "array", "items": {"type": "string"}},
                "evidence": {"type": "array", "items": {"type": "object", "properties": {
                    "table": {"type": "integer"}, "row_label": {"type": "string"}, "col_label": {"type": "string"}, "value": {"type": "string"}},
                    "required": ["table", "row_label", "col_label", "value"]}}},
                "required": ["no", "title", "body", "from_summary", "evidence"]}},
            "decisions": {"type": "array", "items": {"type": "object"}},
            "observations": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["layout", "items", "decisions", "observations"],
    },
}

HOUSE_RULES = {
    "ments": ("각 줄 ≤60자(공백 포함) · 개조식 ㅁ받침 종결(~임/~함/~순임/~편임, 꼬리 괄호 허용) · 마침표 금지 · "
              "작은따옴표 보기명은 grid.rows[].label 원문 그대로(축약 금지, 길면 나열 절단) · 수치는 슬라이스 표시값만 · "
              "1행 = 전체값(ment_facts.line1 / level), 2행 = 배너 튀는 값(ment_facts.line2 후보 안에서만, 「상대적으로 높음/낮음」) · "
              "line2 가 비면 2행 생략 · base_n<30 은 % 없이 패턴 D · chart_type null 이면 문장화하지 말고 decisions"),
    "summary": ("요약 멘트 = 수치 없는 판단 헤드라인(≤40자, `- ` 접두) + 패널을 가로질러 종합한 해석 1~2줄(≤60자, ㅁ종결) — 수치는 패널이 보여준다 · "
                "패널은 상위·격차 행만(`· 상위 N` / `· 행 'A'|'B'`, 구간형·척도 예외), 배너 표는 `배너 표 (군명)` 로 열 선별 · "
                "모든 주장은 claims 에 근거 셀(표·행 라벨·열 라벨·표시값) 등재, 합산값 금지 · 꼭지 제목 2~3어절 · 표본 단위(건물) 유지 · 「에스원」·제품명 금지"),
    "summary_text": ("요약 꼭지마다 **판단 헤드라인 1줄(≤40자, 수치 금지, `- ` 접두) + 종합 해석 1줄(≤60자, 수치 금지, ㅁ받침 종결 — ~임/~함/~음)** 만 쓴다 · "
                     "패널(표) 여러 개를 가로질러 「그래서 무엇인가」를 말한다 — 수치 읽기·나열 금지, 수치는 패널이 보여준다 · 핵심 보기명만 작은따옴표(라벨 원문) · "
                     "마침표 금지 · 「에스원」·제품명 금지 · 표본 단위는 건물(「건물주」「응답자」로 바꾸지 말 것)"),
    "polish": ("【문체 교정 전용】 내용은 이미 확정됐다 — **의미·주장·수치·보기명(라벨 원문)·문장 수·종결 형태를 바꾸지 않는다**. 어휘 선택·어순·리듬·군더더기만 다듬어 보고서답게 만든다 · "
               "숫자·%·n=·괄호 안 기준 표기는 글자 그대로 보존 · 종결은 원문과 같은 형태(ㅁ받침 개조식이면 ㅁ받침, 마침표 평서문이면 평서문) · 길이는 원문 이하 또는 +5자 이내 · "
               "고칠 것이 없으면 원문 그대로 두고 changed=false · 새 해석·제언·접속 논리를 얹지 않는다 · id 는 입력과 동일, 전부 응답"),
    "toc_names": ("구조(챕터·서브·배정)는 이미 확정됐다 — 바꾸지 않는다. 각 항목의 **제목 어휘만** 다듬는다 · 제목은 한 어절 명사(구) ≤12자, `및`·`/`·`·`·쉼표로 두 개념을 잇지 않는다, 어미·조사 종결 금지, 수치 금지 · "
                  "항목이 담은 문항(문항 텍스트·응답 유형)을 읽고 독자가 목차만 보고 무엇을 묻는 장인지 알 수 있는 이름을 제안한다 · 보고서 목차 어휘(현황·실태·평가·의향·수용도·구조·요인·범위 등) · "
                  "현재 이름이 충분하면 verdict=유지 하고 alternatives 는 비워도 된다(억지 교체 금지) · 교체 권고는 이유를 한 줄로"),
    "toc": ("Part Ⅲ 조사 결과 목차를 3단(챕터 → 서브 타이틀 → 장)으로 설계한다 · 입력의 본문 단위 목록(seq 1..N)을 **전부 정확히 한 번씩** 어느 서브에 배정한다(누락·중복 금지) · "
            "챕터·서브 타이틀은 **한 어절 명사(구)**, `및`·`/`·`·` 로 두 개념을 잇지 않는다(잇고 싶으면 항목을 쪼갠다), 어미·조사 종결 금지 · "
            "챕터 = 측정 대상 단위(설문 파트 맹신 금지 — 같은 대상의 사전/사후·만족/이유 문항은 한 묶음) · 매트릭스 묶음(같은 문항코드 여러 장)과 순위 쌍은 같은 서브에 연속 · "
            "보고서 순서는 의뢰인이 읽는 논리 순서(설문 순서와 달라도 됨) · excluded 에는 입력의 제외 문항(통계표 없음·SQ 등)을 사유와 함께 전부 · rationale 에 분리·병합 근거를 표 번호·실측값으로"),
    "storyline": ("【종결 규칙 — 가장 먼저 지킨다】 모든 헤더·해석·결론 제목은 한국어 개조식 **ㅁ받침 종결**로 끝난다: ~임 / ~함 / ~음 / ~짐 / ~됨. "
                  "예) 「가격을 제시하는 순간 의향이 절반 아래로 내려앉음」 「다음에 맡길 영역은 현장이 아니라 자산·행정임」. "
                  "명사형 종결은 금지 — 「…필요」「…과제」「…포착」「…전략」「…예상」 ✗ → 「…필요함」「…과제임」「…보임」 ✓. 마침표 금지 · "
                  "목표: 요약 장표 헤더 5개 + 결론 박스 3개 + 총결론 1개를 **표를 보지 않아도 누구나 이해되는 한 줄기 이야기**로 쓴다 · "
                  "각 헤더는 앞 헤더를 받아 잇는다(link_to_prev 에 연결 논리 한 줄) · 헤더 ≤40자·수치 금지·판단문(ㅁ받침 종결: ~임/~함/~음/~짐) · "
                  "interpretation 은 헤더를 뒷받침하는 해석 1줄 ≤60자·수치 금지·ㅁ종결 · 결론 title 은 판단 한 문장(≤30자), recommendation 은 「→ …」 제언 1줄 ≤60자 · "
                  "총결론(마지막)은 헤더 5개 전체를 한 문장으로 닫는다 · story_in_one_paragraph 에 헤더+결론을 이어 읽은 한 문단(검증용) · "
                  "핵심 보기명만 작은따옴표(표의 라벨 원문) · 「에스원」·제품명 금지 · 표본 단위는 건물"),
    "conclusion": ("결론 = 박스 3개 + 총결론 1개(마지막) · 박스 = 판단 한 문장 제목 + 결정적 수치 1~2개 + `→` 제언, 박스당 ≤150자 · "
                   "결론은 요약에서만(from_summary 에 요약 번호) · 본문 수치는 전부 evidence 에 좌표 · 「유일한」「반드시」는 반대 증거 0 일 때만 · "
                   "박스 순서 = 의뢰인 결정 순서(진입 → 상품·타겟 → 가격 → 총결론)"),
}


def load_text(p: Path | None) -> str:
    if not p or not p.exists():
        return ""
    return io.open(p, encoding="utf-8").read()


def build_messages(task: str, agent_md: str, input_json: str, context: str) -> list[dict]:
    system = (f"너는 리서치 보고서 제작가이드의 문장화 담당이다. 아래 담당 정의를 그대로 따른다. 반환은 JSON 하나만.\n\n"
              f"## 담당 정의\n{agent_md}\n\n## 하우스 문장 규칙 (기계 게이트가 전수 검사)\n{HOUSE_RULES[task]}")
    user = f"## 입력 (팩트 슬라이스 — 여기 없는 수치는 쓰지 않는다)\n```json\n{input_json}\n```"
    if context:
        user += f"\n\n## 맥락 (조사 목적·결정)\n{context}"
    user += "\n\n위 스키마에 맞는 JSON 만 반환하라."
    return [{"role": "system", "content": system}, {"role": "user", "content": user}]


def call_clova(model: str, messages: list[dict], schema: dict, reasoning: str | None, temperature: float) -> str:
    key = os.environ.get("CLOVA_STUDIO_API_KEY")
    if not key:
        raise SystemExit("CLOVA_STUDIO_API_KEY 환경변수가 없습니다 (CLOVA Studio → API 키 → 테스트 API 키, nv-…)")
    from openai import OpenAI  # pip install openai
    client = OpenAI(api_key=key, base_url="https://clovastudio.stream.ntruss.com/v1/openai")
    base = dict(model=model, messages=messages, temperature=temperature)
    # 추론형(HCX-007)은 max_tokens 를 거부하고 max_completion_tokens 를 받는다 (실측 40001)
    limit = {"max_completion_tokens": 8192} if model.upper().startswith("HCX-007") else {"max_tokens": 4096}
    rf = {"type": "json_schema", "json_schema": {"name": "guide_output", "schema": schema}}
    # 호환 엔드포인트는 모델·조합에 따라 reasoning / response_format 을 거부한다(40001). 단계적으로 뺀다.
    attempts = []
    if reasoning:
        attempts.append({**base, **limit, "response_format": rf, "reasoning_effort": reasoning})
        attempts.append({**base, **limit, "reasoning_effort": reasoning})
    attempts.append({**base, **limit, "response_format": rf})
    attempts.append({**base, **limit})
    attempts.append(dict(base))
    last = None
    for kw in attempts:
        if "response_format" not in kw:
            # 스키마를 프롬프트로 강제
            kw = {**kw, "messages": messages[:-1] + [{"role": "user", "content": messages[-1]["content"]
                  + "\n\n## 반환 JSON 스키마 (이 스키마의 JSON 객체 하나만, 설명 없이)\n" + json.dumps(schema, ensure_ascii=False)}]}
        try:
            r = None
            for wait in (0, 20, 60, 120):            # 429(rate limit)·5xx 는 대기 후 재시도
                if wait:
                    import time; time.sleep(wait)
                try:
                    r = client.chat.completions.create(**kw)
                    break
                except Exception as e:
                    msg = str(e)
                    if ("429" in msg or "rate" in msg.lower() or "50" in msg[:12]) and wait < 120:
                        print(f"[clova] {msg[:60]} → {wait+20}s 후 재시도", file=sys.stderr)
                        continue
                    raise
            used = [k for k in ("response_format", "reasoning_effort") if k in kw]
            print(f"[clova] ok with {used or 'plain'}", file=sys.stderr)
            return r.choices[0].message.content
        except Exception as e:                       # openai.BadRequestError 등
            last = e
            if "40001" not in str(e) and "Invalid parameter" not in str(e):
                raise
            print(f"[clova] retry — {str(e)[:120]}", file=sys.stderr)
    raise last


def call_anthropic(model: str, messages: list[dict], schema: dict, temperature: float) -> str:
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        raise SystemExit("ANTHROPIC_API_KEY 환경변수가 없습니다")
    import anthropic  # pip install anthropic
    client = anthropic.Anthropic(api_key=key)
    system = messages[0]["content"] + f"\n\n## 반환 JSON 스키마\n{json.dumps(schema, ensure_ascii=False)}"
    r = client.messages.create(model=model, max_tokens=8000, temperature=temperature, system=system,
                               messages=[{"role": "user", "content": messages[1]["content"]}])
    return "".join(b.text for b in r.content if getattr(b, "type", "") == "text")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--provider", choices=["clova", "anthropic"], default="clova")
    ap.add_argument("--model", default="HCX-005", help="clova: HCX-005 | HCX-007(추론) | HCX-DASH-002 / anthropic: claude-opus-5 등")
    ap.add_argument("--task", choices=list(SCHEMAS), required=True)
    ap.add_argument("--agent", type=Path, required=True, help=".claude/agents/guide-*.md (담당 정의 = system prompt)")
    ap.add_argument("--input", type=Path, required=True, help="팩트 슬라이스 JSON (샤드 / summary_map / conclusion 입력)")
    ap.add_argument("--context", type=Path, help="조사개요.json · ORIENTATION.md 등 맥락 파일")
    ap.add_argument("--reasoning", choices=["none", "low", "medium", "high"], help="clova HCX-007 / 호환 파라미터")
    ap.add_argument("--temperature", type=float, default=0.3)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    agent_md = load_text(args.agent)
    body = agent_md.split("---", 2)[-1] if agent_md.startswith("---") else agent_md   # frontmatter 제거
    messages = build_messages(args.task, body.strip(), load_text(args.input), load_text(args.context))
    schema = SCHEMAS[args.task]
    if args.dry_run:
        print(json.dumps({"messages": messages, "schema": schema}, ensure_ascii=False, indent=1)[:6000])
        return 0
    raw = call_clova(args.model, messages, schema, args.reasoning, args.temperature) if args.provider == "clova" \
        else call_anthropic(args.model, messages, schema, args.temperature)
    try:
        obj = json.loads(raw)
    except json.JSONDecodeError:
        start, end = raw.find("{"), raw.rfind("}")
        try:
            obj = json.loads(raw[start:end + 1])
        except json.JSONDecodeError:
            rawp = args.out.with_suffix(".raw.txt")
            rawp.write_text(raw, encoding="utf-8")
            sys.exit(f"[{args.provider}] JSON 파싱 실패 — 원문을 {rawp} 에 저장 (재발주 또는 수동 복구)")
    obj.setdefault("_meta", {})["provider"] = args.provider
    obj["_meta"]["model"] = args.model
    obj["_meta"]["task"] = args.task
    args.out.parent.mkdir(parents=True, exist_ok=True)
    with io.open(args.out, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=1)
    print(f"→ {args.out}  (다음: verify.py --mode {args.task} 로 게이트)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
