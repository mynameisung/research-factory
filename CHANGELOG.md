# Changelog

## 0.2.0 — 2026-09-10

정본 리포가 **FineHourTeam/report-factory** 로 옮겨 갔고, 스킬이 스스로 새 판을 알아채 묻기 시작한다.

- **자동 판 확인** — `/제작가이드`·`/보고서작성`·`/회고` 가 시작할 때 원격 `VERSION` 을 먼저 읽는다(작업 폴더는 건드리지 않는다). 새 판이 있으면 「올릴까요?」를 묻고, 「나중에」는 24시간 → 48시간 → 1주일씩 미룬다. 엔진 `engine/tools/update_check.py`, 상태는 `~/.report-factory/`
- **요약·결론 팀 재편** — 꼭지별 문장 담당(guide-summary-slide)과 의뢰인 자리 심사(guide-client-review)를 신설하고, 결론은 총괄(guide-summary)이 직접 쓴다. guide-summary-text·guide-conclusion 은 폐지
- **규칙 R-198~R-243** — 요약·결론 피라미드(R-199), 요약 세 줄은 걸음(R-200), 설계는 본문 헤더부터(R-201) 등 46개 등재
- **도구 5종** — `refill_deck`(완성 덱에 정본 재반영·경로 C) · `audit_deck_values`(표·차트를 장마다 정본과 대조) · `refresh_facts` · `embed_chart_data`(차트 데이터를 PPT 안에) · `banner_insight`
- **통계표전처리 .xlam** — 리본·배너 팝업 개정, 설치용 xlam 추가
- 지식 문서: 19_요약그래프-데이터설계 · summary_chart_rules.yaml 신설

## 0.1.0 — 2026-08-28

첫 판. 에스원 건물현황실사(TA·비상주) 2개 조사를 완주한 상태의 엔진을 그대로 담았다.

- 스킬 5개: `/제작가이드` `/보고서작성` `/회고` + 배포용 `/업그레이드` `/개선공유`
- 에이전트 12명 (가이드 설계 3 · 요약결론 3 · 심사검증 3 · PPT·회고 3)
- 규칙 대장 R-197 까지 (비상주 r15 에서 R-193 개정 — 멘트 한 문장 한 통계량, K-16 검사 신설)
- 기계 검수(verify) · 칸 대 칸 대조(crosscheck) · 문서 위생(doc_hygiene) · 학습 캡처 훅
- 고객 데이터(`projects/`)·코퍼스 원자료(`refs/`)는 리포 밖 (.gitignore)
