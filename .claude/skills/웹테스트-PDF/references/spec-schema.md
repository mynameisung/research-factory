# spec.json 스키마 — 설문 원안(Word)의 구조화 규격

이 문서는 설문 원안(Word/한글 문서)을 읽어 `spec.json` 으로 옮길 때의 규격이다.
`spec.json` 은 이후 모든 검증 단계의 **정답지(expected)** 이고, 실제 웹 설문에서 크롤링한
**관측치(actual)** 와 대조하는 기준이 된다.

---

## 0. 가장 중요한 원칙 — 파서를 만들지 마라

**원안에서 spec.json을 뽑는 정규식 파서를 작성하지 말 것.** 모델이 원안을 직접 읽고 판단해서 채운다.

이유는 단순하다. 원안 포맷은 매번 다르다.

- 프로젝트마다 문항 번호 체계가 다르다 (`SQ1`, `S1`, `문1`, `Q1`, `A1`, `DQ1`…)
- 같은 회사 안에서도 연구원마다 지시문 표기가 다르다 ("1개 선택" / "단일응답" / "(SA)" / 표기 없음)
- 로테이션·이탈 조건은 본문 텍스트가 아니라 **서식**(빨간색, 취소선, 음영, 각주, 표 바깥 메모)으로 표현되는 경우가 많다
- 표 안에 들어간 문항, 표 없이 줄바꿈만으로 된 문항이 한 문서에 섞여 있다
- 앞 문항 지시문이 뒤 문항에 상속되는 암묵적 규칙이 있다 ("이하 동일", "위와 같음")

정규식 파서는 이 중 하나만 어긋나도 조용히 틀린 spec을 만들어내고, 그러면 검증 리포트 전체가
거짓 양성/거짓 음성으로 오염된다. 잘못된 spec은 검증을 안 한 것보다 나쁘다.

대신 이렇게 한다.

1. 원안 전체를 읽는다 (docx는 `anthropic-skills:docx` 스킬로 텍스트 + 서식 정보를 함께 추출).
2. 문항 단위로 끊어서 아래 스키마 필드를 하나씩 채운다.
3. **판단이 애매한 항목은 추측해서 채우지 말고 `"_uncertain"` 배열에 기록**하고 사용자에게 확인받는다.
4. 다 채운 뒤 §6 체크리스트로 자기 검증한다.

---

## 1. 최상위 구조

```json
{
  "survey": { },
  "questions": [ ]
}
```

파일은 **UTF-8 (BOM 없음)** 으로 저장한다.
엔진 페이지는 EUC-KR이지만 그건 크롤링 쪽 문제이고, spec.json은 UTF-8로 통일한다.
비교 시점에는 크롤링 결과를 EUC-KR → UTF-8로 디코딩한 뒤 spec과 맞춘다.

---

## 2. `survey` 오브젝트

| 필드 | 타입 | 필수 | 의미 | 원안 어디서 가져오나 |
|---|---|---|---|---|
| `id` | string | O | 프로젝트 관리번호. 파일명·리포트 제목에 쓰인다. | 원안 표지의 프로젝트 코드 (예: `"2026-0812"`). 없으면 원안 파일명에서 딴다. |
| `name` | string | O | 설문 정식 명칭. | 표지 제목 (예: `"OO 브랜드 U&A 조사"`). |
| `alias` | string | X | 짧은 식별자. 산출물 폴더명·로그 접두어로 쓴다. 영문/숫자/하이픈만. | 사람이 정한다. 없으면 `id`를 그대로 쓴다. |
| `url` | string | O | 테스트 설문 진입 URL. | 원안이 아니라 **개발자가 준 링크**에서 온다. 엠브레인 엔진은 보통 `https://survey.panel.co.kr/rSurveyTest.asp` 계열이며, 진입 시 쿼리스트링으로 테스트 ID가 붙는 형태다. |

```json
"survey": {
  "id": "2026-0812",
  "name": "OO 브랜드 U&A 조사",
  "alias": "ooua",
  "url": "https://survey.panel.co.kr/rSurveyTest.asp"
}
```

> 엔진 실측: 모든 페이지가 같은 URL(`/rSurveyTest.asp`)로 POST 자기순환한다. 페이지가 넘어가도
> 주소창은 바뀌지 않는다. 따라서 `url`은 "진입점" 하나만 적으면 되고, 문항별 URL 같은 건 없다.

---

## 3. `questions[]` 오브젝트

### 3.1 필드 요약

| 필드 | 타입 | 필수 | 기본값 |
|---|---|---|---|
| `code` | string | O | — |
| `page` | string \| null | X | `null` |
| `text` | string | O | — |
| `type` | enum | O | — |
| `options` | array | 조건부 | `[]` |
| `required` | boolean | O | `true` |
| `minSelect` | number \| null | 조건부 | `null` |
| `maxSelect` | number \| null | 조건부 | `null` |
| `rotation` | enum | O | `"none"` |
| `piping` | object \| null | X | `null` |
| `logic` | array | X | `[]` |

### 3.2 `code`

원안에 적힌 문항 번호를 **그대로** 쓴다. `SQ1`, `D1`, `E1`, `E1-1`, `문5` 등.

- 대소문자·하이픈·공백을 임의로 정규화하지 말 것. 원안이 `E1-1`이면 `E1_1`로 바꾸지 않는다.
  로직의 `target`, 파이핑의 `sourceQuestion` 이 이 문자열로 서로를 참조하기 때문에
  표기가 흔들리면 참조가 깨진다.
- 원안에 번호가 없는 안내문(인사말, 섹션 구분 화면)은 `INFO1`, `INFO2` 처럼 붙이고 `type: "info"` 로 둔다.

### 3.3 `page`

원안에 페이지 파일명이 명시돼 있을 때만 채운다 (예: `"m7.asp"`). 없으면 `null`.

엔진은 히든필드 `nextPageAction` 에 다음 페이지 파일명(예 `m7.asp`)을, `ProgressPage` 에 현재
페이지 번호를 담아 다닌다. 원안에 파일명이 적혀 있으면 크롤링한 `nextPageAction` 과 대조해서
**문항이 의도한 페이지에 실렸는지**를 직접 검증할 수 있다. 없다고 억지로 지어내지 말 것.

### 3.4 `text`

문항 지시문 본문. 보기 목록과 응답 개수 지시("1개 선택" 등)는 **빼고** 질문 문장만 넣는다.

- 파이핑 자리는 원안 표기를 살려서 `{E1_1순위}` 같은 플레이스홀더로 남긴다. 실제 값은 `piping` 이 정의한다.
- 줄바꿈은 `\n` 으로 보존한다. 강조(볼드)는 텍스트에 반영하지 않는다 — 텍스트 비교 시 노이즈가 된다.

### 3.5 `type`

| 값 | 의미 | 엔진에서의 모습 |
|---|---|---|
| `single` | 단일선택 (하나만) | 리스트형 라디오, 또는 `RankTxt` 위젯 1개 제한 |
| `multi` | 복수선택 | `RankTxt` 위젯 (클릭 순서대로 1,2,3… 이 value에 들어감) |
| `rank` | 순위형 (1순위/2순위…) | `RankTxt` 위젯. `multi` 와 위젯은 같고 **순서에 의미가 있다**는 점이 다르다 |
| `scale` | 척도형 (5점/7점/11점 등) | `<td><p><font><span class="input_rdo"><input type="radio" …>` 구조. **`<label>` 태그가 없다** |
| `text` | 주관식 (단답/장문/숫자) | `<input type="text">` / `<textarea>` |
| `matrix` | 행렬형 (여러 행 × 같은 척도) | 척도형과 같은 구조가 여러 행. id가 `_00`, `_01`, `_02` … 로 증가 |
| `info` | 안내문 (응답 없음) | 다음 버튼만 있는 페이지 |

`multi` 와 `rank` 의 구분은 원안 문구로 판단한다. "3개 선택"은 `multi`, "1순위/2순위/3순위" 또는
"우선순위대로"는 `rank`. 위젯이 동일하므로 웹만 봐서는 구분이 안 되고, **원안이 유일한 근거**다.

> **척도형 함정 (실측)**: 척도형·행렬형 페이지에는 `<label>` 태그가 하나도 없고 보기 텍스트가
> 다른 셀/행에 흩어져 있다. 라벨 기반 추출은 실패하거나 옆 보기 텍스트까지 뭉쳐 잡힌다.
> 따라서 `scale` / `matrix` 문항의 `options[].code` 는 반드시 **input의 `value` 속성과 맞아떨어지는
> 값**(척도 점수 1,2,3,4,5)으로 적어야 한다. 라벨 텍스트로 대조하려 들면 안 된다.

### 3.6 `options[]`

`single` / `multi` / `rank` / `scale` / `matrix` 에서 필수. `text` / `info` 에서는 `[]`.

| 필드 | 타입 | 의미 |
|---|---|---|
| `code` | string | 보기 번호. **엔진 input의 `value` 속성과 일치해야 하는 값.** 원안의 보기 번호(1,2,3…)를 문자열로 넣는다. |
| `text` | string | 보기 문구. |
| `fixed` | boolean | 로테이션에서 제외되어 자리가 고정되는 보기인가. |
| `exclusive` | boolean | 배타 보기인가 (선택 시 다른 보기가 모두 해제됨). |
| `hasTextInput` | boolean | 보기 옆에 직접입력 칸이 붙는가 ("기타(____)"). |

```json
{ "code": "11", "text": "기타", "fixed": true, "exclusive": false, "hasTextInput": true }
```

> **`code` 는 라벨이 아니라 value다 (실측 근거)**: 라벨 추출이 실패하는 페이지가 실재하므로,
> 전이표(어떤 보기를 고르면 어디로 가는지)와 모든 대조는 **input의 `value` 속성** 기준으로 기록한다.
> `options[].code` 를 보기 텍스트나 임의 슬러그로 채우면 전이표와 조인이 안 된다.

> **RankTxt 위젯 타입 대응 (실측)**: 복수/순위 문항의 각 보기는
> `<input class="RankTxt" id="RankTxt{n}" onclick="jFunRank(type,'{n}')">` 형태이고, `type` 인자가
> 보기 성격을 드러낸다.
> - `type=''` → 일반 보기 → `exclusive:false, hasTextInput:false`
> - `type='t1'` → 기타(텍스트 연동) 보기 → `hasTextInput:true`
> - `type='y'` → 배타 보기(없음/해당사항 없음) → `exclusive:true`
>
> spec을 채운 뒤 이 세 가지가 실제 페이지의 `jFunRank` 인자와 맞는지 대조하면 배타/기타 속성
> 구현 오류를 바로 잡아낼 수 있다.

### 3.7 `required`, `minSelect`, `maxSelect`

- `required`: 무응답으로 다음 버튼을 눌렀을 때 경고가 떠야 하면 `true`.
  원안에 "응답 안 해도 통과" 같은 명시가 있을 때만 `false`.
- `minSelect` / `maxSelect`: 선택 개수 제약. `single` 이면 둘 다 `1`.
  `text` / `info` 는 둘 다 `null`.
- "3개 선택"처럼 정확히 N개면 `min=max=N`.
- "최대 3개"면 `min=1, max=3` (required가 true인 경우).
- "해당되는 것 모두"면 `min=1, max=` 보기 개수.

> 엔진의 `embPageSet()` 이 필수응답·개수 제약 검증 함수다. 반환값이 false면 `window.alert` 로
> 경고를 띄운다. 검증 스크립트에서 `window.alert` 를 덮어쓰면 경고창에 막히지 않고 메시지를
> 수집할 수 있으므로, `required` / `minSelect` / `maxSelect` 는 **자동 검증이 잘 되는 필드**다.
> 그만큼 정확히 채워둘 값어치가 있다.

### 3.8 `rotation`

| 값 | 의미 |
|---|---|
| `"none"` | 보기 순서 고정 |
| `"all"` | 모든 보기 로테이션 |
| `"exceptFixed"` | `fixed: true` 인 보기를 제자리에 두고 나머지만 로테이션 |

실무에서 가장 흔한 건 `"exceptFixed"` 다. "기타", "없음", "잘 모르겠음" 은 거의 항상 맨 끝 고정이므로
로테이션 지시가 있으면 그 보기들의 `fixed` 를 `true` 로 놓고 `exceptFixed` 를 쓴다.
원안에 "보기 전체 랜덤"이 명시된 경우에만 `"all"`.

### 3.9 `piping`

앞 문항의 응답을 이 문항 텍스트/보기에 끌어다 쓰는 경우.

```json
"piping": {
  "sourceQuestion": "E1",
  "rank": 1,
  "fallbackWhenEtc": "textInput"
}
```

| 필드 | 타입 | 의미 |
|---|---|---|
| `sourceQuestion` | string | 값을 끌어오는 원본 문항의 `code`. |
| `rank` | number \| null | 원본이 순위형일 때 몇 순위 응답을 쓰는가. 순위형이 아니면 `null`. |
| `fallbackWhenEtc` | string \| null | 원본에서 "기타"가 선택됐을 때 대신 노출할 값. 응답자가 입력한 기타 텍스트를 그대로 쓰면 `"textInput"`, 고정 문구로 대체되면 그 문구 문자열, 처리 규칙이 없으면 `null`. |

**ASP 소스에서 파이핑을 읽는 법.** 개발자가 준 ASP 소스를 볼 수 있다면 이 두 패턴이 근거가 된다.

- `<%=ArrE1(rr("v20"))%>`
  → 보기 배열 `ArrE1` 에서 히든필드 `v20` 값에 해당하는 보기 텍스트를 꺼내 쓴다는 뜻.
  `v20` 이 E1의 1순위 응답을 담는 필드라면 `sourceQuestion:"E1", rank:1`.
- `If rr("v20")="11" Then` … 기타 텍스트 필드를 대신 출력
  → 기타(보기 코드 11) 선택 시 분기 처리가 있다는 뜻 → `fallbackWhenEtc` 를 채운다.

ASP 소스가 없으면 원안의 표기("E1에서 1순위로 응답한 브랜드")로 채우고, `vN` 대응은
**히든필드 `v1..vN` 이 지금까지의 모든 응답을 누적**하므로 실제 주행 중 폼 상태를 덤프해서 확인한다.

### 3.10 `logic[]`

조건부 흐름. 각 원소는 다음 형태다.

```json
{ "when": { "option": "1" }, "action": "terminate", "target": null }
```

| 필드 | 타입 | 의미 |
|---|---|---|
| `when.option` | string | 이 문항에서 선택된 보기 `code`(=value). 조건이 되는 보기가 여러 개면 원소를 여러 개 만든다. |
| `action` | `"skip"` \| `"jumpTo"` \| `"terminate"` | 아래 참조 |
| `target` | string \| null | `skip` → 건너뛸 문항 code. `jumpTo` → 이동할 문항 code. `terminate` → `null`. |

- `skip` : `target` 문항(들)을 **건너뛰고** 그 다음 문항으로 간다. 원안의 "D2 생략", "D2는 묻지 않음".
- `jumpTo` : `target` 문항으로 **점프**한다. 원안의 "→ Q7로", "Q7로 이동", "문7번으로".
- `terminate` : 설문 **이탈/탈락**. 원안의 "종료", "중단", "조사 종료", "탈락".

한 문항에 여러 조건이 걸리면 배열로 나열한다. 순서는 원안 기재 순서를 따른다.

```json
"logic": [
  { "when": { "option": "1" }, "action": "terminate", "target": null },
  { "when": { "option": "2" }, "action": "terminate", "target": null },
  { "when": { "option": "3" }, "action": "jumpTo",    "target": "D5" }
]
```

---

## 4. 한국어 설문 원안 표기 → 필드 매핑표

원안에서 아래 표기를 만나면 이렇게 옮긴다. **표기는 회사·연구원마다 다르므로 이 표는 출발점일 뿐,
같은 뜻의 다른 표현이 나오면 뜻을 보고 판단한다.**

### 4.1 응답 개수 / 문항 유형

| 원안 표기 | 채울 필드 |
|---|---|
| "1개 선택", "하나만 선택", "단일응답", "(SA)" | `type:"single"`, `minSelect:1`, `maxSelect:1` |
| "3개 선택", "3개만 선택" | `type:"multi"`, `minSelect:3`, `maxSelect:3` |
| "최대 3개 선택", "3개까지" | `type:"multi"`, `minSelect:1`, `maxSelect:3` |
| "복수 선택", "해당되는 것 모두", "(MA)" | `type:"multi"`, `minSelect:1`, `maxSelect:`보기 수 |
| "1순위, 2순위, 3순위", "우선순위대로 3개" | `type:"rank"`, `minSelect:3`, `maxSelect:3` |
| "직접 입력", "주관식", "________", "( )원" | `type:"text"`, `options:[]` |
| "매우 그렇다 ~ 전혀 그렇지 않다", "5점 척도", "7점" | `type:"scale"`, 보기 code = 척도 점수 |
| 여러 행 × 같은 척도 표 | `type:"matrix"` |
| 번호 없는 인사말/안내 화면 | `type:"info"`, `required:false` |

### 4.2 보기 속성

| 원안 표기 | 채울 필드 |
|---|---|
| "기타 ( )", "기타(직접 입력)", "기타: ______" | 해당 보기 `hasTextInput:true`, 보통 `fixed:true` |
| "없음", "해당사항 없음", "해당 없음" | `exclusive:true`, `fixed:true` |
| "잘 모르겠음", "무응답", "응답 거부" | `exclusive:true`, `fixed:true` (배타 여부가 원안에 없으면 사용자에게 확인) |
| "[보기] 로테이션", "보기 순환", "보기 랜덤" | `rotation:"exceptFixed"` (기타/없음은 `fixed:true`) |
| "보기 전체 랜덤", "고정 없음" | `rotation:"all"` |
| "보기 순서 고정", 표기 없음 | `rotation:"none"` |
| 보기 옆 "(고정)" | 그 보기만 `fixed:true` |

### 4.3 흐름 / 이탈

| 원안 표기 | 채울 필드 |
|---|---|
| **취소선 또는 빨간 글씨로 표시된 보기** | 대개 이탈 조건 → `action:"terminate"` |
| "종료", "중단", "조사 종료", "탈락", "SCREEN OUT", "S/O" | `action:"terminate"`, `target:null` |
| "→ Q7로", "→ 문7", "Q7로 이동", "GO TO Q7" | `action:"jumpTo"`, `target:"Q7"` |
| "D2 생략", "D2 SKIP", "D2는 묻지 않음" | `action:"skip"`, `target:"D2"` |
| "1번 응답자만", "2번 선택자에 한해" | 해당 문항이 아니라 **앞 문항의 logic** 으로 표현한다 |

> **취소선·빨간색 주의.** 취소선은 두 가지 뜻으로 쓰인다.
> (a) 이 보기를 고르면 이탈 — 로직으로 옮긴다.
> (b) 최종안에서 **삭제된 보기** — `options` 에서 아예 빼야 한다.
> 문맥으로 확신이 안 서면 절대 추측하지 말고 `_uncertain` 에 넣고 연구원에게 묻는다.
> 두 해석은 검증 결과가 정반대로 나온다.

### 4.4 ASP 소스에 나타나는 표기

개발자가 준 ASP 소스나 원안에 삽입된 개발 메모에서 아래 패턴이 보이면 이렇게 읽는다.

| 소스 표기 | 의미 | 채울 필드 |
|---|---|---|
| `<%=ArrE1(rr("v20"))%>` | 히든필드 `v20` 값으로 보기배열 `ArrE1` 을 조회해 텍스트 출력 = 파이핑 | `piping.sourceQuestion:"E1"`, 해당 순위를 `piping.rank` |
| `<%=ArrX(rr("vN"))%>` (일반형) | 앞 문항 `X` 의 응답 `vN` 을 파이핑 | `piping.sourceQuestion:"X"` |
| `If rr("v20")="11" Then` | 보기 코드 `11`(보통 기타) 선택 시 분기 | `piping.fallbackWhenEtc` 또는 `logic` |
| `rr("v21")` 이 텍스트 필드 | 기타 직접입력 값 | `piping.fallbackWhenEtc:"textInput"` |
| `Response.Redirect` / 종료 페이지 이동 | 이탈 처리 | `action:"terminate"` |

`vN` 번호 자체는 spec.json에 적지 않는다. spec은 **원안 기준의 논리 모델**이고, `vN` 은 엔진 구현
세부사항이다. 다만 주행 중 히든필드 `v1..vN` 을 덤프해두면 파이핑·분기 검증 때 실제 값을 대조할 수 있다.

---

## 5. 완성 예시 (샘플 설문 기준)

아래는 실제 검증에 쓰인 형태의 `spec.json` 조각이다.
SQ1 이탈, D1 스킵, E1 순위형 + 배타보기 + 기타, E1-1 파이핑을 모두 담고 있다.

```json
{
  "survey": {
    "id": "2026-0812",
    "name": "OO 브랜드 U&A 조사",
    "alias": "ooua",
    "url": "https://survey.panel.co.kr/rSurveyTest.asp"
  },
  "questions": [
    {
      "code": "SQ1",
      "page": "m1.asp",
      "text": "귀하 또는 귀하의 가족 중 다음 업종에 종사하는 분이 있습니까?",
      "type": "single",
      "options": [
        { "code": "1", "text": "광고 대행업",             "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "2", "text": "시장조사 / 여론조사업",    "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "3", "text": "언론 / 방송업",            "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "4", "text": "해당 업종 종사자 없음",    "fixed": true,  "exclusive": true,  "hasTextInput": false }
      ],
      "required": true,
      "minSelect": 1,
      "maxSelect": 1,
      "rotation": "none",
      "piping": null,
      "logic": [
        { "when": { "option": "1" }, "action": "terminate", "target": null },
        { "when": { "option": "2" }, "action": "terminate", "target": null },
        { "when": { "option": "3" }, "action": "terminate", "target": null }
      ]
    },

    {
      "code": "D1",
      "page": "m2.asp",
      "text": "귀하의 결혼 여부를 선택해 주십시오.",
      "type": "single",
      "options": [
        { "code": "1", "text": "미혼",            "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "2", "text": "기혼 (자녀 없음)", "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "3", "text": "기혼 (자녀 있음)", "fixed": false, "exclusive": false, "hasTextInput": false }
      ],
      "required": true,
      "minSelect": 1,
      "maxSelect": 1,
      "rotation": "none",
      "piping": null,
      "logic": [
        { "when": { "option": "1" }, "action": "skip", "target": "D2" },
        { "when": { "option": "2" }, "action": "skip", "target": "D2" }
      ]
    },

    {
      "code": "D2",
      "page": "m2.asp",
      "text": "함께 거주하는 자녀는 몇 명입니까?",
      "type": "text",
      "options": [],
      "required": true,
      "minSelect": null,
      "maxSelect": null,
      "rotation": "none",
      "piping": null,
      "logic": []
    },

    {
      "code": "E1",
      "page": "m7.asp",
      "text": "다음 브랜드 중 평소 가장 자주 이용하시는 브랜드를\n1순위부터 3순위까지 순서대로 선택해 주십시오.",
      "type": "rank",
      "options": [
        { "code": "1",  "text": "A 브랜드", "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "2",  "text": "B 브랜드", "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "3",  "text": "C 브랜드", "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "4",  "text": "D 브랜드", "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "5",  "text": "E 브랜드", "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "6",  "text": "F 브랜드", "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "7",  "text": "G 브랜드", "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "8",  "text": "H 브랜드", "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "9",  "text": "I 브랜드", "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "10", "text": "이용하는 브랜드 없음", "fixed": true, "exclusive": true,  "hasTextInput": false },
        { "code": "11", "text": "기타",     "fixed": true,  "exclusive": false, "hasTextInput": true  }
      ],
      "required": true,
      "minSelect": 3,
      "maxSelect": 3,
      "rotation": "exceptFixed",
      "piping": null,
      "logic": [
        { "when": { "option": "10" }, "action": "jumpTo", "target": "E2" }
      ]
    },

    {
      "code": "E1-1",
      "page": "m8.asp",
      "text": "앞서 1순위로 선택하신 {E1_1순위}을(를) 이용하시는 가장 큰 이유는 무엇입니까?",
      "type": "single",
      "options": [
        { "code": "1", "text": "가격이 저렴해서",       "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "2", "text": "품질이 좋아서",         "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "3", "text": "접근성이 좋아서",       "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "4", "text": "브랜드 이미지가 좋아서", "fixed": false, "exclusive": false, "hasTextInput": false },
        { "code": "9", "text": "기타",                  "fixed": true,  "exclusive": false, "hasTextInput": true  }
      ],
      "required": true,
      "minSelect": 1,
      "maxSelect": 1,
      "rotation": "exceptFixed",
      "piping": {
        "sourceQuestion": "E1",
        "rank": 1,
        "fallbackWhenEtc": "textInput"
      },
      "logic": []
    }
  ]
}
```

### 예시 해설

- **SQ1 이탈** — 보기 1~3은 원안에서 빨간 글씨 + "종료"로 표기된 이탈 조건이다. 각각 별도
  `logic` 원소로 풀어 썼다. 보기 4는 "없음"이므로 `exclusive:true, fixed:true`.
  단일선택이라 배타 동작이 눈에 안 띄지만, 엔진이 이 보기를 `jFunRank(...,'y')` 로 구현했는지
  대조할 근거가 되므로 그대로 적는다.
- **D1 스킵** — "미혼/자녀 없음이면 D2 생략"이 원안 표기다. `skip` 의 `target` 은 **건너뛸 문항**
  (D2)이지 이동할 문항이 아니다. `jumpTo` 와 헷갈리지 말 것.
- **E1 순위형** — `type:"rank"` + `min=max=3`. "이용하는 브랜드 없음"(코드 10)은 배타 보기라
  선택 시 다른 보기가 다 풀리고, 순위 3개를 채울 수 없으므로 별도 `jumpTo` 로직이 함께 온다.
  "기타"(코드 11)는 배타는 아니고 텍스트 입력이 붙는다. 둘 다 맨 끝 고정이라 `rotation:"exceptFixed"`.
- **E1-1 파이핑** — 텍스트에 `{E1_1순위}` 플레이스홀더를 남기고, 실제 대응은 `piping` 이 정의한다.
  ASP 소스에 `<%=ArrE1(rr("v20"))%>` 와 `If rr("v20")="11" Then` 이 함께 있으면
  `rank:1` + `fallbackWhenEtc:"textInput"` 이 맞다는 근거가 된다.

---

## 6. 작성 후 자기 검증 체크리스트

spec.json을 다 채운 뒤, 실제 웹 검증에 넘기기 **전에** 아래를 확인한다.
여기서 걸리는 오류는 전부 spec 쪽 오류이고, 그대로 두면 리포트가 오염된다.

**참조 무결성**
- [ ] 모든 `logic[].target` 이 존재하는 `questions[].code` 와 정확히 일치한다 (`terminate` 는 `null`).
- [ ] 모든 `piping.sourceQuestion` 이 존재하는 `code` 이고, **자기보다 앞에** 나온다.
- [ ] `piping.rank` 가 있으면 `sourceQuestion` 의 `type` 이 `rank` 다.
- [ ] `code` 에 중복이 없다.

**개수 제약**
- [ ] `type:"single"` 인 문항은 `minSelect == maxSelect == 1`.
- [ ] `minSelect <= maxSelect`.
- [ ] `maxSelect <=` 해당 문항의 `options` 개수.
- [ ] `type:"text"` / `"info"` 는 `minSelect`, `maxSelect` 가 모두 `null` 이고 `options` 가 빈 배열.

**보기 속성**
- [ ] `exclusive:true` 인 보기는 `fixed:true` 다 (배타 보기가 로테이션에 섞이는 설계는 사실상 없다).
- [ ] `rotation:"exceptFixed"` 인데 `fixed:true` 보기가 하나도 없으면 `"all"` 을 의도한 것 아닌지 재확인.
- [ ] `hasTextInput:true` 보기가 두 개 이상이면 원안을 다시 본다 (드문 설계다).
- [ ] `code` 가 보기 텍스트나 슬러그가 아니라 **숫자 문자열**이다.

**흐름**
- [ ] 이탈(`terminate`)이 최소 한 군데는 있다. 스크리닝 문항이 있는 설문에서 이탈이 0건이면 놓친 것이다.
- [ ] `skip` 으로 건너뛰는 문항이 실제로 `questions[]` 에 존재한다.
- [ ] 배타 보기("없음")를 가진 복수/순위 문항에 후속 흐름(`jumpTo` 또는 다음 문항의 조건)이 정의돼 있다.

**미확정 항목**
- [ ] `_uncertain` 에 남은 항목을 연구원에게 확인받았고, 확인된 것은 본문에 반영하고 배열에서 지웠다.

`_uncertain` 은 이런 형태로 최상위에 임시로 둔다. 최종 spec.json에는 남기지 않는다.

```json
"_uncertain": [
  { "code": "E3", "field": "options[7].text", "issue": "취소선 표기 — 이탈 조건인지 삭제된 보기인지 불명" },
  { "code": "F2", "field": "rotation", "issue": "'보기 순환' 메모가 F2 전체인지 F2-1만인지 불명" }
]
```

---

## 7. spec.json 이 이후 단계에서 쓰이는 곳

| 검증 항목 | 쓰이는 필드 | 방식 |
|---|---|---|
| 문항 누락/추가 | `code`, `page` | 주행하며 만난 문항 집합과 대조 |
| 문항 지시문 오타 | `text` | EUC-KR 디코딩 후 텍스트 비교 |
| 보기 누락/오타/순서 | `options[].code`, `.text` | input `value` 기준 조인 후 텍스트 비교 |
| 필수응답 | `required` | `window.alert` 를 덮어쓰고 무응답 상태로 다음 버튼 → 경고 유무 수집 |
| 선택 개수 제약 | `minSelect`, `maxSelect` | 경계값(min-1, max+1)으로 시도 → 경고 유무 수집 |
| 배타 보기 동작 | `options[].exclusive` | `jFunRank` 의 `type='y'` 인자 대조 + 실제 클릭 후 다른 보기 해제 확인 |
| 기타 직접입력 | `options[].hasTextInput` | `jFunRank` 의 `type='t1'` 대조. **반드시 기타 클릭 → 텍스트 요소 재조회 → 입력 순서**로 확인 (클릭 전 입력값은 무효) |
| 로테이션 | `rotation`, `options[].fixed` | 서로 다른 테스트 ID로 같은 문항을 여러 번 열어 보기 순서 분포 확인 |
| 분기/이탈 | `logic[]` | 보기별 전이표(**value 기준**)를 만들어 spec과 대조 |
| 파이핑 | `piping` | 앞 문항 응답을 고정하고 주행 → 뒤 문항 텍스트에 기대값이 노출되는지 확인. 히든필드 `v1..vN` 덤프로 교차 확인 |

주행 관련 실측 사항 두 가지는 spec을 쓸 때가 아니라 스크립트를 쓸 때 필요하지만, 여기 적어둔다.

- 히든필드 기반 세션이라 같은 브라우저에서 iframe 여러 개를 동시에 돌려도 세션이 섞이지 않는다.
  10개 동시 주행으로 50 시나리오/1358 페이지를 185초에 완주한 실측이 있다 (페이지당 평균 178ms).
- 순위/복수선택을 무작위로 채울 때 **서로 다른 보기 인덱스를 보장**해야 한다. 같은 인덱스를 두 번
  뽑으면 토글이 해제되어 0개 선택 상태가 되고, 필수응답 경고에 걸려 시나리오가 통째로 어긋난다.
