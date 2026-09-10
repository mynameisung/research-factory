# 30. 외부 LLM 연동 — CLOVA Studio (HyperCLOVA X)

문장화 담당(본문 멘트 · 요약 · 결론)을 외부 한국어 모델에 맡길 때의 기준 문서. 어댑터는 `engine/tools/llm_write.py`.
원칙은 바뀌지 않는다 — **오케스트레이터(Claude)가 프롬프트·슬라이스를 주고, 모델은 문장만 쓰고, 수용은 기계 게이트**
(`verify.py --mode ments|summary|conclusion`)와 검토관(guide-summary-reviewer)이 결정한다. 모델이 바뀌어도 계약은 같다.

## 1. 접속

| 항목 | 값 |
|---|---|
| 콘솔 | NCP → AI Services → **CLOVA Studio** (AI·NAVER API 의 CLOVA Voice/CSR 과 다르다) |
| 키 | CLOVA Studio → 「API 키」 → 테스트 API 키 (`nv-…`). **최초 발급 시에만 표시**. 환경변수 `CLOVA_STUDIO_API_KEY` — 저장소·대화에 넣지 않는다 |
| OpenAI 호환 | `https://clovastudio.stream.ntruss.com/v1/openai` — `openai` SDK `base_url` 교체. `response_format: json_schema` · `reasoning_effort`(none/low/medium/high) 지원. 미지원: frequency/presence_penalty |
| 네이티브 v3 | `POST /v3/chat-completions/{model}` · 헤더 `Authorization: Bearer` · `X-NCP-CLOVASTUDIO-REQUEST-ID` |
| 모델 | **HCX-007** 추론형(결론·제언 — thinking short/medium/deep) · **HCX-005** 일반(입력 128k·출력 4k) · **HCX-DASH-002** 경량(입력 32k) |

## 2. 요금 (2026-08 기준, VAT 별도, 1,000 토큰당 원) — 사용자 제공 2026-08-23

### 인퍼런스 (실시간)

| 구분 | 모델/도구 | 입력 | 출력 | 비고 |
|---|---|---:|---:|---|
| 기본 | HCX-007 | 1.25 | 5 | 입·출력 분리 과금 |
| 기본 | HCX-005 | 1.25 | 5 | 〃 |
| 기본 | HCX-DASH-002 | 0.25 | 1 | 〃 |
| 기본 | HCX-003 | 5 (입출력 합산) | | 구형 |
| 기본 | HCX-DASH-001 | 1 (입출력 합산) | | 구형 |
| 튜닝 | HCX-005 | 2.5 | 10 | 튜닝 모델 인퍼런스 |
| 튜닝 | HCX-DASH-002 | 0.5 | 2 | |
| 튜닝 | HCX-003 | 30 | | |
| 튜닝 | HCX-DASH-001 | 3 | | |
| 스킬셋 | HCX-003 / 학습모델 | 5 / 30 | | |
| 익스플로러 | 리랭커 | 1.25 | 5 | |
| 익스플로러 | RAG Reasoning | 2.5 | 10 | |
| 익스플로러 | 요약 | 22.5 (입력 기준) | | 요약 도구만 입력 토큰 과금 |
| 익스플로러 | 임베딩v2 / 임베딩 / 문단 나누기 | 0.2 / 0.1 / 0.4 | | |
| 라우터 | 라우터 전용 모델 | 0.5 | | |

- 플레이그라운드·튜닝·스킬셋·익스플로러 사용 시 선택 엔진 요금 부과. 익스플로러의 데이터 확장·일괄 생성은 사용 엔진의 인퍼런스 요금.

### 학습 (튜닝)

| 모델 | 1,000 토큰당 |
|---|---:|
| HCX-005 | 6.25 |
| HCX-DASH-002 | 1.25 |
| HCX-003 | 15 |
| HCX-DASH-001 | 5 |
| 스킬셋 HCX-003 | 15 |

## 3. 우리 파이프라인 기준 비용 추정 (비상주 실측 파일 크기 × 한국어·JSON ≈ 2토큰/자, 보수적)

| 작업 | 입력 규모 | HCX-007/005 | HCX-DASH-002 |
|---|---|---:|---:|
| 결론·제언 1회 (summary_map 58k자 + 조사개요 + 담당 정의) | ≈125k 토큰 | 입력 ≈157원 + 출력(≈4k) ≈20원 ≈ **180원** | ≈35원 |
| 본문 멘트 샤드 1개 (슬라이스 57k자 + 정의) | ≈116k 토큰 | ≈ **160원** | ≈32원 |
| 본문 멘트 7샤드 + 재발주 1회 | | ≈ **1,300원** | ≈260원 |
| 요약 5꼭지 (tables.json 전수 3.9MB 는 못 넣는다 → 표별 요약 통계로 압축 입력 필요) | 압축 후 ≈100k | ≈ **150원** | |

→ 프로젝트 한 건 전체를 CLOVA 에 맡겨도 **수천 원대**. 비용보다 **슬라이스 압축**(HCX-005 입력 128k·DASH 32k 상한)이 설계 변수다.
  본문 샤드는 grid 전체(17열×행)를 주지 말고 `ment_facts` + 라벨 목록만 주면 1/5 로 준다.

## 4. 파일럿 설계 (A/B)

1. 같은 슬라이스·같은 프롬프트(담당 정의)로 Claude 담당 vs CLOVA(HCX-007 / HCX-005) 각각 생성
2. 비교 지표: ① 기계 게이트 1차 통과율(길이·종결·라벨 원문·수치⊆슬라이스) ② 검토관 findings 수·심각도 ③ 사용자 선호(문체)
3. 먼저 **결론·제언**(설득 문장, HCX-007 reasoning medium) → **요약 해석 2줄**(HCX-005) → 본문 멘트 순. 결과는 `work/검수/clova_*.json` + `work/검수/ab_*.md`
4. 규칙 준수는 모델이 아니라 게이트가 보장한다 — 모델 비교는 **통과 후 문체·설득력**에서만 의미가 있다

## 5. 실행

```bash
# 키는 사용자가 직접: setx CLOVA_STUDIO_API_KEY "nv-…"   (대화·저장소에 넣지 않는다)
pip install openai
python engine/tools/llm_write.py --provider clova --model HCX-007 --task conclusion --reasoning medium \
  --agent ../.claude/agents/guide-summary.md --input work/summary_map.json --context work/조사개요.json \
  # (guide-conclusion 은 2026-09-01 폐지 — 결론은 총괄 guide-summary 가 쓴다, R-208)
  --out work/검수/clova_conclusion.json
python engine/tools/verify.py work/검수/clova_conclusion.json work/tables.json --mode conclusion
```

## 6. 1차 파일럿 결과 (2026-08-23, 결론·제언)
Claude 담당 치명 0·경고 0 vs HCX-007 치명 2·경고 8 vs HCX-005 치명 0·경고 8. 구조 규칙(박스 3+총결론·→ 제언)·근거 좌표·개조식 문체에서 CLOVA 가 게이트를 못 넘음.
상세: `projects/…/02_비상주/1_제작가이드/work/검수/ab_conclusion_clova.md`. API 실측: HCX-007 은 `max_completion_tokens`+`reasoning_effort`, `response_format` 은 양쪽 40001(프롬프트 강제로 대체).
다음 시도 조건: few-shot(확정본) · claims 압축 입력 · 반려 재발주 1회.

## 7. 2차 파일럿 (컨텍스트 팩 + few-shot) — 요약 헤드라인 0/5 통과(두 모델), 결론 제언 누락. few-shot 은 HCX-005 가 예시문을 복사하는 부작용.
판정: 판단·제언 층은 현행 Claude 담당 유지. CLOVA 재시도 조건 = **튜닝 데이터 축적**(확정 요약·결론 ≥ 수십 건) 후 HCX-005 튜닝(학습 6.25원/1k).

## 8. 3~4차 (요약 장표·스토리라인) — HCX-007 은 실제 표를 받으면 해석은 하나 **명사형 보고서체**를 고수(ㅁ종결 규칙 강조·예시·금지형에도 0/4)하고 근거 없는 제언을 창작. 해석 줄에 수치 나열.
최종 판정(2026-08-23): 단독 담당 불가. 재시험 조건 = 튜닝(확정 요약·결론 데이터 축적). 상세 `ab_conclusion_clova.md`.

## 9. 5차 — 두 접근법 실증(① 완성본 문체 교정 vs ② 구간 생성) — 둘 다 불가
①: HCX-005 는 거의 안 고침(6/60), HCX-007 은 59/60 고치며 수치 16·보기명 23·ㅁ종결 13 훼손 + 배너명 치환 등 사실 오류. **문체만 바꾸는 모델이 아니다.**
②: 목차 배정 최대 49/69, 작명 보조는 대안이 더 길고 느슨, 요약·결론은 요약 층 탈락(so-what 만).
공통: 컨텍스트 팩의 내용을 산출로 복사하는 경향 → 외부 모델 입력에 정답 후보를 넣지 않는다.
어댑터 task 현황: ments·summary·summary_text·conclusion·storyline·toc·toc_names·polish. 상세 `ab_conclusion_clova.md` §5차.

## 10. 파이프라인 A/B 실측 (2026-08-23) — 같은 빌더·같은 게이트로 3벌 생성
`engine/tools/assemble_variant.py <work> <name> --ments --summary --conclusion [--toc-names]` 가 문장 담당만 바꾼 변형본을 `work/variants/<name>/` · `out/variants/<name>/` 에 만든다(build_facts → build_guide → verify guide/ments).
| 변형 | 멘트 게이트 수용 | 가이드 검수 |
|---|---|---|
| A_Claude (정본 체인 재생성) | 69/69 | 치명 0 · 경고 0 |
| B1 CLOVA 문체교정 (HCX-007, 8청크) | 22/69 — ㅁ종결 55 · 보기명 37 · 마침표 32 · 수치 부재 8 | 치명 9 · 경고 32 |
| B2 CLOVA 생성 (HCX-007 멘트 반샤드 14·재발주 1회 + 스토리라인 + 서브명 HCX-005) | 3/69 — 보기명 91 · ㅁ종결 47 · 마침표 27 · 길이 20 | 치명 8 · 경고 27 |
실측 메모: HCX-007 은 출력이 길면 스스로 잘라 버린다(「답변 길이 제한으로 생략」 주석) → 샤드 ≤6장·교정 ≤16줄로 쪼개야 JSON 이 완성된다. 어댑터는 파싱 실패 시 `.raw.txt` 를 남긴다.
질적 심사(사실 정합·문체·스토리라인)는 `work/variants/비교판정_보고서.md`.

## 11. 하이브리드 설계 방향 (사용자 지시 2026-08-23: 「클로바를 워크플로에 넣되 숫자가 틀리면 Claude 가 수정」)
CLOVA 초안(멘트) → `verify --mode ments`(수치·보기명·종결·길이 + V11 장 단위 검사) → 반려 줄만 Claude 담당이 ment_facts 로 접지·수정 → 재게이트.
전제: 요약·결론은 생성 대상에서 제외(5차 실측 — 요약 층 탈락·근거 창작). 구현 시 `llm_write.py --task ments` + `assemble_variant.py` 경로 재사용.
