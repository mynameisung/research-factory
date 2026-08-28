#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""줄바꿈 자리 문법 판정 — 안은문장 / 안긴문장 경계 룰베이스 (R-108).

원리: **안긴문장은 통째로 한 줄에 둔다.** 끊는 자리는 안긴문장이 완결되는 지점이다.

    [로봇을 즉시 통제하기]  어려울 것 같은 느낌        명사절이 끝나는 자리
    [집 안에 항상 존재하는 것에 대한]  정서적 부담      관형절+의존명사+조사구가 끝나는 자리

`존재하는` 뒤에서 끊으면 관형절이 피수식어 `것` 과 찢어진다 — 그래서 금지다.

형태소 분석기 없이 **어미·조사 표층형**으로만 판정하는 휴리스틱이다.
파서가 아니므로 오판이 있을 수 있고, 그래서 금지(AVOID)는 '가능하면 피함'이지
'절대 불가'가 아니다 — 대안이 없으면 벌점을 물고 쓴다. 실제 선택은 textfit.wrap_ko()
의 DP 가 (문법 벌점 합 → 고아 여부 → 줄 폭 제곱합) 사전식 비교로 끝낸다.

여기는 **판정만** 한다. 폭 계산도 줄 나누기도 하지 않는다 — 그건 textfit.py 몫이다.

회귀 확인
--------
    python kobreak.py

  ① 벌점 단위 검사 — 글꼴 없이도 도는 판정 고정 (아래 _PENALTY_CASES)
  ② 줄바꿈 결과 검사 — LG 예시 덱 실측 조건(가용폭 1,764,000 EMU ·
     `LG EI Text TTF Regular` 10pt)에서 사람이 손으로 끊은 자리를 재현하는가
     (아래 _WRAP_CASES). 글꼴이 폴백되면 폭이 달라지므로 SKIP 으로 빠진다.
"""

from __future__ import annotations

import sys
from pathlib import Path

FREE = 0        # 안긴문장이 완결되는 자리 — 최우선
OK = 3          # 무난한 자리 (체언 + 격조사 등)
AVOID = 1000    # 안긴문장 내부 — 끊으면 수식어와 피수식어가 찢어진다

# ── 안긴문장을 완결시키는 표층형 ──────────────────────────────
END_PHRASE = ('대한', '대해', '대하여', '관한', '관해', '관하여', '위한', '위해', '위하여',
              '따른', '따라', '의한', '의해', '통한', '통해', '인한', '인해', '비해')
NOMINAL = ('하기', '되기', '지기', '주기', '보기', '쓰기', '읽기', '듣기', '말하기')
CONNECT = ('하고', '되고', '하며', '이며', '하면서', '지만', '하지만', '이지만',
           '하되', '되되', '거나', '든지', '하여', '해서', '아서', '어서')

# ── 뒤 어절에 반드시 붙는 것들 (그 앞에서 끊으면 안 된다) ────────
BOUND_NEXT = ('것', '수', '바', '줄', '데', '리', '때', '뿐', '만큼', '채', '즈음', '무렵',
              '나름', '따름', '지', '터', '적')                      # 의존명사
AUX = ('주는', '주기', '주고', '있는', '있기', '있고', '없는', '보는', '두는', '놓는',
       '버리는', '가는', '오는', '대는', '같은', '같이', '같다', '싶은', '만한')  # 보조용언·의존형
BOUND_NEXT_EXACT = set(END_PHRASE) | set(AUX)

DETERMINER = ('그', '이', '저', '각', '전', '총', '새', '첫', '매', '온', '온갖', '여러',
              '모든', '어떤', '무슨', '한', '두', '세', '네', '몇', '본', '동')

# 부사도 뒤 용언을 수식한다 — 뒤에서 끊으면 수식어구가 찢어진다 (못 / 하는)
ADVERB = ('못', '안', '잘', '더', '덜', '다', '또', '다시', '함께', '서로', '직접', '바로',
          '매우', '아주', '너무', '가장', '좀', '늘', '자주', '항상', '즉시', '이미', '아직',
          '곧', '결국', '오히려', '특히', '주로', '대체로', '전혀', '거의', '모두', '전부',
          '혼자', '스스로', '미리', '먼저', '반드시', '꼭', '그냥', '일부러', '충분히')

# 체언 + 격조사/보조사 (은/는/의 는 제외 — 아래에서 따로 처리)
PARTICLE = ('을', '를', '이', '가', '에', '에서', '에게', '에겐', '와', '과', '도', '만',
            '로', '으로', '부터', '까지', '처럼', '보다', '이나', '나', '한테', '께', '랑')

DEPENDENT_NOUNS = ('것', '수', '바', '줄', '데', '리', '때', '뿐', '만큼', '채')


def _has_jong(ch, jong=None):
    o = ord(ch)
    if not (0xAC00 <= o <= 0xD7A3):
        return False
    j = (o - 0xAC00) % 28
    return j != 0 if jong is None else j == jong


JONG_N = 4    # ㄴ
JONG_L = 8    # ㄹ


def break_penalty(word, nxt):
    """word 뒤에서 줄을 끊을 때의 벌점. 낮을수록 좋은 자리."""
    w = word.strip()
    bare = w.rstrip(',.·')

    # ⓪ 뒤 어절이 **여는 괄호로 시작하면** 거기서 끊는 게 옳다.
    #    괄호 부연은 그 자체로 완결된 덩어리라 앞 어절에 매달리지 않는다 —
    #    「환경적 요인 | (지반조건, 지형의 변화…」 처럼 사람이 실제로 그렇게 끊는다.
    #    실측(코퍼스 5사례 최종본, 사람 손 줄바꿈 979자리): 이 자리가 20건 있고
    #    SKEE 7 · 물리보안 11 · SH 2 로 **3사례에 걸쳐 재현**된다.
    #    옛 판정은 이걸 AVOID 로 깎아 실제 관행과 어긋났다.
    if nxt and nxt[0] in '(（[［{':
        return FREE

    # ① 뒤 어절이 앞에 매달리는 것이면 끊을 수 없다
    if nxt:
        if nxt in BOUND_NEXT_EXACT:
            return AVOID
        if nxt[0] in ('것', '수', '바', '줄', '데', '리', '때', '뿐', '채') and len(nxt) <= 3:
            return AVOID                       # 의존명사(+조사): 것에 / 수가 / 때만
    # ② 관형사·부사 뒤 — 둘 다 뒤를 수식한다
    if bare in DETERMINER or bare in ADVERB:
        return AVOID

    # ③ 안긴문장이 완결되는 자리
    if w.endswith((',', '·')):
        return FREE
    if bare.endswith(END_PHRASE):
        return FREE
    if bare.endswith(NOMINAL) and len(bare) >= 3:
        return FREE
    if bare.endswith(CONNECT) and len(bare) >= 3:
        return FREE

    # ④ 관형격 '의' / 관형형 '는·던·은' — 뒤를 수식하므로 끊으면 안 된다
    if bare.endswith('의') and len(bare) >= 2:
        return AVOID
    if bare.endswith(('는', '던')) and len(bare) >= 2:
        return AVOID
    if bare.endswith('은') and len(bare) >= 2:
        return AVOID                            # 보조사일 수도 있으나 안전하게 피한다
    if bare.endswith('적'):
        return AVOID                            # 접미사 '-적' 은 뒤 체언을 수식

    # ⑤ 체언 + 격조사 — 무난한 자리
    if bare.endswith(PARTICLE):
        return OK

    # ⑥ 남은 ㄴ/ㄹ 종성은 관형형으로 본다 (중요한 / 어려울 / 받을 …)
    if _has_jong(bare[-1], JONG_N) or _has_jong(bare[-1], JONG_L):
        return AVOID

    return OK


def penalties(words):
    """words[i] 뒤에서 끊을 때의 벌점 리스트 (마지막 어절 뒤는 해당 없음)."""
    return [break_penalty(words[i], words[i + 1] if i + 1 < len(words) else None)
            for i in range(len(words) - 1)]


# ─────────────────────────────────────────────────────────────
# self-test — 회귀 고정
# ─────────────────────────────────────────────────────────────

# LG 예시 덱 s23 라벨열 실측 조건. tblGrid[0].w 1,800,000 − marL 0 − marR 36,000.
SELFTEST_USABLE_EMU = 1_764_000
SELFTEST_FONT = 'LG EI Text TTF Regular'
SELFTEST_PT = 10.0

# ① 벌점 단위 — 판정 근거를 한 줄씩 고정한다. 글꼴이 없어도 돈다.
_PENALTY_CASES = [
    ('통제하기', '어려울',   FREE,  '명사형 어미 -기: 안긴 명사절이 여기서 끝난다'),
    ('대한',     '정서적',   FREE,  'END_PHRASE: 관형절+의존명사+조사구가 여기서 끝난다'),
    ('존재하는', '것에',     AVOID, '관형형 -는 + 뒤가 의존명사 것: 두 번 금지'),
    ('것에',     '대한',     AVOID, '뒤 어절이 END_PHRASE — 앞에 매달린다'),
    ('어려울',   '것',       AVOID, 'ㄹ 관형형 + 뒤가 의존명사'),
    ('항상',     '존재하는', AVOID, '부사는 뒤 용언을 수식한다'),
    ('즉시',     '통제하기', AVOID, '부사'),
    ('시킬',     '때만',     AVOID, 'ㄹ 관형형 + 뒤가 의존명사(+조사) 때만'),
    ('대신해',   '주는',     AVOID, '뒤 어절이 보조용언 주는'),
    ('로봇이',   '그',       OK,    '체언 + 주격조사 — 무난'),
    ('일을',     '대신해',   OK,    '체언 + 목적격조사'),
    ('때만',     '로봇이',   OK,    '의존명사 + 보조사로 구가 닫힌다'),
]

# ② 줄바꿈 결과 — 사람이 손으로 끊은 자리(LG 제작본 실측)를 재현해야 한다.
_WRAP_CASES = [
    ('로봇을 즉시 통제하기 어려울 것 같은 느낌',
     ['로봇을 즉시 통제하기', '어려울 것 같은 느낌']),
    ('집 안에 항상 존재하는 것에 대한 정서적 부담',
     ['집 안에 항상 존재하는 것에 대한', '정서적 부담']),
    ('내가 시킬 때만 로봇이 그 일을 대신해 주는 방식',
     ['내가 시킬 때만 로봇이', '그 일을 대신해 주는 방식']),
]

_GRADE = {FREE: 'FREE', OK: 'OK', AVOID: 'AVOID'}


def _selftest() -> int:
    fail = 0

    print('① 벌점 단위 판정')
    for word, nxt, want, why in _PENALTY_CASES:
        got = break_penalty(word, nxt)
        ok = got == want
        fail += not ok
        print('  %s  %-8s | %-8s  %-5s  %s'
              % ('OK ' if ok else 'FAIL', word, nxt,
                 _GRADE.get(got, got), why if ok else '기대 %s' % _GRADE[want]))

    print()
    print('② 줄바꿈 결과 (가용폭 %s EMU · %s %gpt)'
          % (format(SELFTEST_USABLE_EMU, ','), SELFTEST_FONT, SELFTEST_PT))
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    try:
        import textfit as TF
    except Exception as exc:                     # PIL 미설치 등
        print('  SKIP  textfit 을 불러오지 못했다 — %s' % exc)
        return 1 if fail else 0

    res = TF.resolve_font(SELFTEST_FONT)
    if not res.exact:
        # 폭이 달라지면 어느 자리가 후보인지 자체가 바뀐다 — 비교가 무의미하다.
        print('  SKIP  글꼴 %r 을 못 찾아 %s 로 폴백했다. 폭 기준이 달라져 비교하지 않는다.'
              % (SELFTEST_FONT, res.path.name))
        return 1 if fail else 0

    for text, want in _WRAP_CASES:
        got = TF.wrap_ko(text, SELFTEST_USABLE_EMU, SELFTEST_FONT, SELFTEST_PT)
        ok = got == want
        fail += not ok
        print('  %s  %s' % ('OK ' if ok else 'FAIL', ' / '.join(got)))
        if not ok:
            print('        기대: %s' % ' / '.join(want))

    print()
    print('kobreak self-test: %s' % ('통과' if not fail else '실패 %d건' % fail))
    return 1 if fail else 0


if __name__ == '__main__':
    sys.exit(_selftest())
