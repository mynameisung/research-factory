# 엠브레인 설문 엔진 어댑터

`survey.panel.co.kr` 계열 설문 엔진의 실측 사양이다. 2026-08 기준으로 SurveyID 151508에서
직접 확인했다. 같은 엔진을 쓰는 설문이면 이 문서만으로 어댑터가 완성된다.

## 기본 구조

| 항목 | 값 |
|---|---|
| 문자셋 | **EUC-KR** (UTF-8 아님) |
| 진입 URL | `https://survey.panel.co.kr/SurveyTest.asp?alias=<alias>` |
| 이후 모든 페이지 | `https://survey.panel.co.kr/rSurveyTest.asp` 로 자기순환 POST |
| 폼 | `name="form2"`, `method="post"`, `onsubmit="return embPageSet()"` |
| 다음 버튼 | `<a id="smtBtn" onclick="if (embPageSet()){ form2.submit(); }">다 음</a>` |
| 진입 버튼 | 첫 페이지만 `<a id="smtBtn" href="javascript:form2.submit();">조사참여 하기</a>` |

**URL이 페이지마다 바뀌지 않는다.** 현재 위치를 URL로 알 수 없으므로 히든필드와 본문으로 판별한다.

EUC-KR이라 raw HTTP fetch로 긁으면 한글이 깨진다. 반드시 브라우저(iframe 또는 Playwright)로
띄워서 브라우저가 charset을 처리하게 한다.

## 상태 히든필드

세션 상태가 **쿠키가 아니라 폼 히든필드**에 들어있다. 이것이 이 엔진의 가장 중요한 성질이다.

| 필드 | 의미 |
|---|---|
| `id` | 응답자 테스트 ID. `CT`로 시작 (예 `CT08252338474YP71`) |
| `ProgressPage` | 현재 페이지 번호 |
| `last_page` | 총 페이지 수 (예 45). 진입 시점에 이미 알 수 있다 |
| `nextPageAction` | **기본 다음 페이지 파일명** (예 `m7.asp`) |
| `tmpCnt` | 진행 카운터 |
| `gType` | 분기 그룹 (A / B) |
| `v1` ~ `vN` | **지금까지의 모든 응답이 누적된다** |
| `SurveyID`, `project_no`, `SurveyName`, `Skin_State` | 메타 |

### 이 성질이 주는 이점

**1. 병렬 세션이 안 섞인다.** 쿠키 기반이 아니므로 같은 브라우저에서 iframe을 10개 띄워
동시에 돌려도 각자 다른 `id`를 받아 독립적으로 진행된다. 실측으로 확인했다.

**2. 응답 이력을 매 페이지에서 통째로 읽을 수 있다.** `v*` 필드를 읽으면 지금까지
무엇을 골랐는지 전부 알 수 있어 파이핑·분기 검증에 쓸 수 있다.

**3. 현재 페이지를 식별할 수 있다.** `nextPageAction`은 서버가 렌더링 시점에 넣는
"기본 다음 페이지"다. 즉 현재 페이지 ≈ `nextPageAction`의 직전 페이지다.
문항코드를 본문에서 못 뽑는 안내문 페이지도 `p:<nextPageAction>`으로 고유 식별할 수 있다.

**주의**: `nextPageAction`은 페이지 로드 시점의 **기본값**이지 실제 분기 결과가 아니다.
실제 다음 페이지는 응답을 POST한 뒤 서버가 정한다. 따라서 **분기 검증은
"제출 후 실제로 도착한 페이지"로 해야 하며**, `nextPageAction`만 보고 판단하면 틀린다.

## 응답 위젯

### 라디오 / 체크박스

```html
<td><p><font><span class="input_rdo">
  <input type="radio" name="v4" id="v4_1_00" value="1"><i></i>
</span></font></p></td>
```

- id 규칙: `{name}_{value}_{행인덱스}`. 행렬형은 `_00`, `_01`, `_02`...로 행이 구분된다
- **척도형 문항에는 `<label>` 태그가 하나도 없다.** 보기 텍스트가 다른 셀/행에 있어
  라디오와 구조적으로 분리되어 있다
- 리스트형(단일선택 목록, 예 D1·F3·F4·H2)은 보기 텍스트가 같은 행에 있어 추출된다

### 순위형 / 복수선택 (RankTxt)

```html
<input class="RankTxt" id="RankTxt7" onclick="jFunRank('','7')">
<input class="RankTxt" id="RankTxt11" onclick="jFunRank('t1','11')">  <!-- 기타 -->
<input class="RankTxt" id="RankTxt12" onclick="jFunRank('y','12')">   <!-- 없음(배타) -->
```

`jFunRank(type, n)`의 첫 인자로 보기 성격을 구분한다:

| type | 의미 |
|---|---|
| `''` | 일반 보기 |
| `'t1'` | 기타 — 텍스트 입력란 연동 |
| `'y'` | 배타 보기 (없음 / 해당사항 없음) — 고르면 다른 선택이 해제되고 후속 문항을 스킵 |

클릭하면 `input.value`에 선택 순위(`1`, `2`, ...)가 들어간다. 선택 상태 확인은
`value`가 비어있지 않은지로 한다.

## 함정 — 반드시 지킬 것

### 1. 순위형에서 같은 보기를 두 번 누르면 해제된다

랜덤 인덱스로 두 번 뽑으면 같은 값이 나올 수 있고, 그러면 토글되어 **0개 선택**이 된다.
그 상태로 다음을 누르면 "2개를 선택해 주십시오" 경고에 막힌다.

실측에서 이 버그로 50개 시나리오 중 **33개가 막혔고**, 그게 전부 설문 결함으로 잘못
집계됐다. 서로 다른 인덱스를 반드시 보장한다.

```js
const a = norm[i % norm.length];
const b = norm[(i + 3) % norm.length];   // i와 겹치지 않도록
```

더 안전하게는 클릭 후 `value`가 채워진 개수를 세어 목표에 도달할 때까지 다음 보기로 넘어간다.

### 2. "기타"는 클릭한 뒤에야 텍스트란이 활성화된다

기타 보기를 클릭하기 **전에** 텍스트를 채우면 무효 처리된다. 순서를 지켜야 한다:

```js
etcInput.click();
await sleep(250);
const texts = [...doc.querySelectorAll('input[type=text],textarea')]
                .filter(e => !e.classList.contains('RankTxt'));   // 클릭 후 재조회
texts.forEach(x => setValue(x, '기타내용'));
```

반대로 **기타가 아닌 보기를 골랐는데 기타란을 채우면** 엔진이 차단한다
(`"SQ2번에 기타를 선택한 분만 기타를 써 주십시오."`). 텍스트란을 무조건 채우면 안 된다.

권장 전략은 **적응형 재시도**다: 먼저 선택형만 답하고 제출 → 막히면 경고 메시지를 보고
그때 텍스트를 채워 재시도. 이러면 기타 연동 여부를 몰라도 알아서 맞춰진다.

### 3. 라벨 추출은 자주 실패하고, 실패하면 옆 보기와 뭉친다

척도형에는 label이 없고, 리스트형도 스킨에 따라 `<td>` 하나에 두 보기가 들어가
`"예 아니오"` 같이 뭉쳐 잡힌다.

**전이표는 반드시 `input.value`로 기록한다.** 실측에서 라벨 기준으로 기록했더니
`SQ1: "아니오" → SQ2` (실제는 정반대)처럼 전이표 전체가 틀리게 나왔다.

보기 텍스트는 보고서 가독성용으로만 쓰고, 판정은 value로 한다.

### 4. 로드 직후 읽으면 빈 페이지로 보인다

`load` 이벤트 직후 `body.innerText`를 읽으면 아직 비어 있을 수 있다.
실측에서 이 경쟁 조건으로 "빈 페이지" 오탐이 대량 발생했다.

```js
let txt = doc.body ? doc.body.innerText : '';
if (txt.trim().length < 10) { await sleep(220); txt = doc.body.innerText; }  // 재확인
```

재확인 후에도 비어 있으면 진짜 빈 페이지다. 이 엔진에는 `m27_1.asp`, `m33_1.asp` 같은
**내용 없는 라우팅 전용 페이지**가 실제로 존재한다.

### 5. `doc.body`가 null인 순간이 있다

네비게이션 도중 `contentDocument.body`가 `null`일 수 있다. 항상 가드한다.

## 검증 기법

### 필수응답 검증을 경고창 없이 전수 확인

`window.alert`를 덮어쓰면 경고창에 막히지 않고 메시지를 수집할 수 있다.
`embPageSet()`을 직접 호출하거나, 무응답 상태로 버튼을 눌러 alert가 뜨는지 본다.

```js
const captured = [];
iframe.contentWindow.alert = m => captured.push(String(m));
iframe.contentWindow.confirm = () => true;

const before = captured.length;
btn.click();
if (captured.length > before) {
  // 동기적으로 alert가 떴다 = 검증이 작동함, 페이지 이동 없음
} else {
  await waitForLoad();   // 이동함
}
```

`embPageSet()`이 `false`를 반환하면 alert가 **동기적으로** 뜨고 폼은 제출되지 않는다.
따라서 클릭 직후 캡처 배열 길이만 비교하면 이동 여부를 즉시 알 수 있어, 로드 타임아웃을
기다릴 필요가 없다.

이 방법으로 35개 문항의 필수응답 검증을 한 세션에서 전부 확인했다.

### 경고 문구의 문항번호 대조

경고 메시지는 `"E1-1번에 입력해 주십시오."` 형태로 문항번호를 포함한다.
현재 페이지의 문항코드와 비교하면 **문항번호 오표기**를 잡을 수 있다.
실측에서 E2-1 페이지가 `"E2-2번에 입력해 주십시오."`를 띄우는 결함을 이 방법으로 찾았다.

### 페이지 전환 대기

전체 POST 네비게이션이므로 `load` 이벤트를 기다린다. iframe이면 부모 컨텍스트가
살아 있어 안정적이다.

```js
const waitLoad = (ms) => new Promise(res => {
  let done = false;
  const h = () => { if (!done) { done = true; ifr.removeEventListener('load', h); res(true); } };
  ifr.addEventListener('load', h);
  setTimeout(() => { if (!done) { done = true; ifr.removeEventListener('load', h); res(false); } }, ms);
});
```

## 페이지 종류 판별

| 종류 | 판별 |
|---|---|
| 문항 | 본문에 `SQ\d+` 또는 `[A-H]\d+(-\d+)?` 뒤에 `.` 이 오고, 응답 위젯이 있음 |
| 안내문 | 다음 버튼은 있는데 입력 요소가 0개. 본문의 "응답해 주세요"에 속으면 안 된다 |
| 라우팅 전용 | 본문이 비어 있고 다음 버튼만 있음 |
| 종료 | `smtBtn`이 없음. 본문에 "대상이 아닙니다"면 이탈, "모두 응답해 주셔서"면 정상 완료 |

문항코드 정규식: `/\b(SQ\d+|[A-H]\d+(?:[-_]\d+)?)\s*[.)]/`

## 성능 실측

| 항목 | 값 |
|---|---|
| 페이지당 평균 | **178 ms** |
| 50 시나리오 / 1,358 페이지 | 185 초 (10개 동시) |
| 완주 1회 (45~51 페이지) | 5~9 초 |

스크린샷 기반 에이전트 제어는 클릭당 5~15초이므로 **30배 이상** 차이가 난다.

## 새 엔진 어댑터 만들기

다른 설문 엔진이면 아래 순서로 30분 안에 어댑터를 만들 수 있다.

1. **진입 페이지 DOM 덤프** — `document.forms`, 모든 `input`의 name/id/type/value,
   버튼 후보(`a`, `button`, `input[type=button]`)의 id/class/onclick/href, `document.characterSet`
2. **다음 버튼과 제출 방식 확인** — onclick에 검증 함수가 있는지, 폼 이름은 무엇인지
3. **한 페이지 진행 후 히든필드 비교** — 어떤 필드가 상태를 나르는지, 페이지 식별자가 있는지
4. **문항 페이지 덤프** — 라디오 name/id 규칙, label 유무, 보기 텍스트의 위치
5. **특수 위젯 확인** — 순위형·슬라이더·드래그 등. onclick 핸들러 이름으로 성격 구분
6. **무응답 제출** — 검증 함수 이름과 경고 메시지 형식 파악, alert 가로채기 가능 여부
7. **종료 페이지** — 정상 완료와 이탈을 어떻게 구분하는지

각 단계는 `javascript_tool`로 JSON을 반환받으면 된다. 위 7가지가 채워지면
`assets/runner.js`의 어댑터 상수만 바꿔 그대로 재사용할 수 있다.
