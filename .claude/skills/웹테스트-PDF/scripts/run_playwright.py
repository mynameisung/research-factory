"""설문 웹테스트 주행 엔진 — full 모드 (Playwright)

브라우저 컨텍스트가 완전히 분리되어 세션 충돌 위험이 없고, 수백~수천 시나리오를
감당하며, 이상 발생 시 스크린샷을 남긴다.

사용법:
    python run_playwright.py --url <TEST_URL> --scenarios scenarios.json \
        --workers 10 --out results.json [--shots shots/] [--headed]

scenarios.json 형식:
    [
      {"id":"P001","label":"SQ1 아니오(이탈)","policy":"first",
       "decide":{"SQ1":{"value":"2"},"E1":{"rank":"excl"}}},
      {"id":"N001","label":"필수응답 전수검증","policy":"first","probeRequired":true},
      {"id":"S001","label":"strict 클릭 검증","policy":"first","strict":true},
      {"id":"R001","label":"랜덤 seed1","policy":"rand","seed":1}
    ]

decide 는 {문항코드: 지시} 매핑이고, 지시는 다음 중 하나다:
    {"index": 0}            보기 순번 (0부터)
    {"value": "2"}          보기의 value 속성 — 가장 안전하다
    {"text": "해당 사항"}    보기 텍스트 정규식
    {"rank": "norm"|"etc"|"excl"}   순위형 위젯 전략
"""

import argparse, asyncio, json, re, sys, time
from pathlib import Path

from playwright.async_api import async_playwright

# ── 엔진 어댑터 (다른 엔진이면 여기만 교체) ──────────────────────────────────
QCODE_RE = re.compile(r"\b(SQ\d+|[A-H]\d+(?:[-_]\d+)?)\s*[.)]")
WARN_QCODE_RE = re.compile(r"\b(SQ\d+|[A-H]\d+(?:[-_]\d+)?)\s*번")
NEED_RE = re.compile(r"(\d+)\s*개\s*선택")
PCT_RE = re.compile(r"(\d+)%")
DROPOUT_RE = re.compile(r"대상이\s?아닙니다|대상자가\s?아니")
COMPLETE_RE = re.compile(r"모두\s?응답해|참여해\s?주셔서\s?감사")
GARBLED_RE = re.compile(r"[�]|\?{3,}|<[a-z]+\s|&nbsp;|undefined|NaN", re.I)
ERROR_PAGE_RE = re.compile(r"오류가\s?발생|잘못된\s?접근|세션이\s?만료")

NEXT_BTN = "#smtBtn"
RANK_SEL = "input.RankTxt"
PAGE_ID_FIELD = "nextPageAction"

SETTLE_MS = 220        # 빈 본문 재확인 대기
ETC_SETTLE_MS = 250    # 기타 클릭 후 텍스트란 활성화 대기
SLOW_PAGE_MS = 6000
MAX_STEPS = 80


class Rng:
    """시드 고정 난수 — 재현 가능해야 한다."""

    def __init__(self, seed=1):
        self.s = seed or 1

    def next(self, n):
        self.s = (self.s * 1103515245 + 12345) & 0x7FFFFFFF
        return self.s % n if n else self.s


async def page_state(page):
    """현재 페이지의 상태를 한 번의 evaluate로 긁어온다 (왕복 최소화)."""
    return await page.evaluate(
        """() => {
        const d = document;
        if (!d || !d.body) return null;
        const hidden = {};
        const f = d.forms[0];
        if (f) [...f.elements].forEach(e => { if (e.type === 'hidden') hidden[e.name] = e.value; });
        const esc = s => (window.CSS && CSS.escape) ? CSS.escape(s) : String(s).replace(/[^\\w-]/g, '\\\\$&');
        const optText = el => {
          let t = '';
          if (el.id) { const l = d.querySelector('label[for="' + esc(el.id) + '"]'); if (l) t = l.innerText; }
          if (!t) { const p = el.closest('label'); if (p) t = p.innerText; }
          if (!t) { const td = el.closest('td'); if (td && td.nextElementSibling) t = td.nextElementSibling.innerText; }
          if (!t) { const p = el.closest('li,tr'); if (p) t = p.innerText; }
          return (t || '').replace(/\\s+/g, ' ').trim().slice(0, 45);
        };
        const ranks = [...d.querySelectorAll('input.RankTxt')].map((e, i) => ({
          i, id: e.id, value: e.value,
          kind: ((e.getAttribute('onclick') || '').match(/jFunRank\\('([^']*)'/) || [,''])[1],
          text: optText(e)
        }));
        const radios = [...d.querySelectorAll('input[type=radio],input[type=checkbox]')];
        const groupName = radios.length ? radios[0].name : null;
        const group = radios.filter(e => e.name === groupName).map((e, i) => ({
          i, id: e.id, value: e.value, type: e.type, disabled: !!e.disabled,
          checked: e.checked, text: optText(e)
        }));
        const otherGroups = {};
        radios.forEach(e => { if (e.name !== groupName) (otherGroups[e.name] = otherGroups[e.name] || []).push(e.id); });
        const texts = [...d.querySelectorAll('input[type=text],textarea')]
          .filter(e => !e.classList.contains('RankTxt'))
          .map(e => ({ id: e.id, name: e.name, disabled: !!e.disabled, readOnly: !!e.readOnly,
                       value: e.value, visible: e.offsetParent !== null }));
        return {
          text: (d.body.innerText || '').replace(/\\s+/g, ' ').trim(),
          hidden, ranks, group, groupName, otherGroups, texts,
          hasNext: !!d.querySelector('#smtBtn'),
          selects: [...d.querySelectorAll('select')].map(e => ({ id: e.id, name: e.name, n: e.options.length, value: e.value }))
        };
      }"""
    )


async def install_hooks(page):
    """alert/confirm을 가로채고 JS 에러를 모은다. 경고창에 막히지 않게 된다."""
    await page.evaluate(
        """() => {
        window.__alerts = []; window.__jserr = [];
        window.alert = m => { window.__alerts.push(String(m)); };
        window.confirm = () => true;
        window.onerror = m => { window.__jserr.push(String(m).slice(0, 70)); return true; };
      }"""
    )


async def drain(page, key):
    try:
        return await page.evaluate(f"() => {{ const a = window.{key} || []; window.{key} = []; return a; }}")
    except Exception:
        return []


async def click_next(page):
    """다음 버튼을 누르고, 동기 alert가 떴는지(=검증에 막혔는지) 알려준다.

    누르기 전에 window.__nav 표식을 심는다. 페이지가 실제로 넘어가면 새 문서에는
    이 표식이 없으므로, 표식이 사라지는 것을 기다리면 네비게이션 완료를 정확히 알 수 있다.
    wait_for_load_state 는 이전 페이지가 이미 load 상태라 즉시 반환해버려 쓸 수 없다.
    """
    return await page.evaluate(
        """() => {
        window.__nav = 1;
        const before = (window.__alerts || []).length;
        const b = document.querySelector('#smtBtn');
        if (!b) return { clicked: false, blocked: false };
        b.click();
        return { clicked: true, blocked: (window.__alerts || []).length > before };
      }"""
    )


async def wait_navigated(page, timeout=20000):
    """click_next 가 심은 표식이 사라질 때까지 기다린다."""
    try:
        await page.wait_for_function("() => typeof window.__nav === 'undefined'", timeout=timeout)
        await page.wait_for_load_state("load", timeout=timeout)
        return True
    except Exception:
        return False


async def choose_radio(page, el_id, strict):
    """strict면 순수 클릭만 시도한다 — 사람이 실제로 누를 수 있는지 검증."""
    return await page.evaluate(
        """([id, strict]) => {
        const d = document, el = d.getElementById(id);
        if (!el) return 'MISSING';
        if (el.checked) return 'already';
        const esc = s => (window.CSS && CSS.escape) ? CSS.escape(s) : s;
        try { el.click(); } catch (e) {}
        if (el.checked) return 'click';
        const l = d.querySelector('label[for="' + esc(id) + '"]');
        if (l) { try { l.click(); } catch (e) {} if (el.checked) return 'label'; }
        const p = el.closest('label');
        if (p) { try { p.click(); } catch (e) {} if (el.checked) return 'parentLabel'; }
        if (strict) return 'FAIL';
        el.checked = true;
        el.dispatchEvent(new Event('click', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        return el.checked ? 'forced' : 'FAIL';
      }""",
        [el_id, strict],
    )


async def fill_texts(page, value, only_empty=True):
    return await page.evaluate(
        """([v, onlyEmpty]) => {
        const els = [...document.querySelectorAll('input[type=text],textarea')]
          .filter(e => !e.classList.contains('RankTxt'));
        let n = 0;
        els.forEach(x => {
          if (x.offsetParent === null || x.disabled || x.readOnly) return;
          if (onlyEmpty && x.value) return;
          x.value = v;
          ['input', 'change', 'keyup', 'blur'].forEach(ev => x.dispatchEvent(new Event(ev, { bubbles: true })));
          n++;
        });
        return n;
      }""",
        [value, only_empty],
    )


async def run_scenario(context, plan, url, shots_dir):
    page = await context.new_page()
    rng = Rng(plan.get("seed", 1))
    strict = bool(plan.get("strict"))
    decide = plan.get("decide") or {}
    policy = plan.get("policy", "first")
    etc_value = plan.get("etcValue", "기타_" + plan["id"])

    path, anomalies, answers = [], [], []
    ended, end_text, last_pct = "MAXSTEP", "", -1
    req_ok = req_miss = 0
    seen_q = {}
    t0 = time.time()

    def flag(step, key, typ, detail):
        anomalies.append({"step": step, "key": key, "type": typ, "detail": str(detail)[:180]})

    # 진입. 서버가 간헐적으로 "설문 진행 중 오류" 페이지를 주므로 몇 번 재시도한다.
    entered = False
    for attempt in range(3):
        await page.goto(url, wait_until="load")
        await page.wait_for_timeout(150)
        await install_hooks(page)
        if not await page.query_selector(NEXT_BTN):
            await page.wait_for_timeout(400 * (attempt + 1))
            continue
        await click_next(page)
        if not await wait_navigated(page):
            await page.wait_for_timeout(400 * (attempt + 1))
            continue
        await install_hooks(page)
        body = await page.evaluate("() => (document.body && document.body.innerText || '').slice(0, 300)")
        if ERROR_PAGE_RE.search(body):
            await page.wait_for_timeout(600 * (attempt + 1))
            continue
        entered = True
        break
    if not entered:
        flag(0, "", "진입 실패", "재시도 3회 후에도 설문 진입 불가 (서버 오류 페이지 또는 버튼 없음)")
        await page.close()
        return {
            "id": plan["id"], "label": plan.get("label", plan["id"]), "steps": 0,
            "ms": int((time.time() - t0) * 1000), "endType": "진입실패", "endText": "",
            "path": "", "answers": [], "anomalies": anomalies, "reqOk": 0, "reqMiss": 0,
        }

    for step in range(1, MAX_STEPS + 1):
        st = await page_state(page)
        if st is None:
            flag(step, "", "문서 없음", "body null")
            break

        # 안정화: 로드 직후 빈 본문은 경쟁 조건일 수 있다
        if len(st["text"]) < 10:
            await page.wait_for_timeout(SETTLE_MS)
            st = await page_state(page)
            if st is None:
                break

        txt = st["text"]
        hid = st["hidden"]
        page_id = "p:" + str(hid.get(PAGE_ID_FIELD, "")).replace(".asp", "")
        m = QCODE_RE.search(txt)
        qcode = m.group(1) if m else ""
        key = qcode or page_id
        pm = PCT_RE.search(txt)
        pct = int(pm.group(1)) if pm else -1

        jserr = await drain(page, "__jserr")
        n_widget = len(st["ranks"]) + len(st["group"])

        # ── 검수 레이어 ──
        if len(txt) < 10 and st["hasNext"]:
            flag(step, key, "빈 페이지", f"안정화 후에도 본문 0자, next={hid.get(PAGE_ID_FIELD)}")
        if GARBLED_RE.search(txt[:400]):
            flag(step, key, "문자깨짐/태그노출", txt[:90])
        if jserr:
            flag(step, key, "JS 에러", " | ".join(jserr[:2]))
        if pct >= 0:
            if pct < last_pct:
                flag(step, key, "진행률 역행", f"{last_pct}% → {pct}%")
            if pct > 100:
                flag(step, key, "진행률 100% 초과", f"{pct}%")
            last_pct = pct
        if qcode:
            if qcode in seen_q:
                flag(step, key, "문항 중복 노출", f"{seen_q[qcode] + 1}회")
            seen_q[qcode] = seen_q.get(qcode, 0) + 1
        # 안내문 오탐 방지: 문항코드가 있을 때만 위젯 없음을 의심한다
        if st["hasNext"] and qcode and n_widget == 0 and not st["texts"]:
            flag(step, key, "응답 위젯 없음", txt[:90])

        if not st["hasNext"]:
            ended, end_text = "종료", txt[-160:]
            path.append(key + "#END")
            break
        path.append(key)

        # ── 음성 테스트: 무응답 제출 ──
        if plan.get("probeRequired") and n_widget > 0:
            r = await click_next(page)
            if r["blocked"]:
                req_ok += 1
                msgs = await page.evaluate("() => window.__alerts || []")
                if msgs:
                    wm = WARN_QCODE_RE.search(msgs[-1])
                    if qcode and wm and wm.group(1) != qcode:
                        flag(step, key, "경고문구 문항번호 불일치",
                             f'페이지={qcode} 경고문="{msgs[-1][:60]}"')
            else:
                req_miss += 1
                flag(step, key, "필수응답 검증 없음", "무응답으로 통과됨")
                try:
                    await page.wait_for_load_state("load", timeout=20000)
                except Exception:
                    pass
                await install_hooks(page)
                continue

        # ── 응답 ──
        nm = NEED_RE.search(txt)
        need = int(nm.group(1)) if nm else 1
        want = decide.get(qcode) or decide.get(key) or {}
        chosen_v, chosen_t, how = "-", "-", ""

        if st["ranks"]:
            norm = [r for r in st["ranks"] if r["kind"] == ""]
            etc = [r for r in st["ranks"] if r["kind"] == "t1"]
            excl = [r for r in st["ranks"] if r["kind"] == "y"]
            mode = want.get("rank", "norm")

            if mode == "excl" and excl:
                await page.click("#" + excl[0]["id"])
                await page.wait_for_timeout(60)
                chosen_v, chosen_t = "excl", excl[0]["text"]
            elif mode == "etc" and etc:
                await page.click("#" + norm[0]["id"])
                await page.wait_for_timeout(60)
                await page.click("#" + etc[0]["id"])
                # 기타는 클릭한 뒤에야 텍스트란이 활성화된다. 미리 채우면 무효다.
                await page.wait_for_timeout(ETC_SETTLE_MS)
                await fill_texts(page, etc_value, only_empty=False)
                chosen_v, chosen_t = "etc", "일반1+기타"
            else:
                # 서로 다른 인덱스를 보장한다. 같은 걸 두 번 누르면 토글되어 해제된다.
                start = rng.next(len(norm)) if policy == "rand" else 0
                picked, i, guard = 0, 0, 0
                while picked < min(need, len(norm)) and guard < len(norm) * 2:
                    guard += 1
                    el = norm[(start + i) % len(norm)]
                    i += 1
                    cur = await page.evaluate(
                        "id => { const e = document.getElementById(id); return e ? e.value : ''; }", el["id"]
                    )
                    if not cur:
                        await page.click("#" + el["id"])
                        await page.wait_for_timeout(50)
                    picked = await page.evaluate(
                        "() => [...document.querySelectorAll('input.RankTxt')].filter(e => e.value).length"
                    )
                chosen_v, chosen_t = "norm", f"일반{picked}개"

        elif st["group"]:
            g = st["group"]
            idx = 0
            if "index" in want:
                idx = want["index"]
            elif "value" in want:
                j = next((k for k, o in enumerate(g) if o["value"] == str(want["value"])), -1)
                idx = j if j >= 0 else 0
            elif "text" in want:
                rx = re.compile(want["text"])
                j = next((k for k, o in enumerate(g) if rx.search(o["text"] or "")), -1)
                idx = j if j >= 0 else 0
            elif policy == "last":
                idx = len(g) - 1
            elif policy == "second":
                idx = min(1, len(g) - 1)
            elif policy == "cycle":
                # 결정점이 아닌 문항의 보기 커버리지를 시나리오 번호로 순환시켜 채운다
                idx = plan.get("optionOffset", 0) % len(g)
            elif policy == "rand":
                idx = rng.next(len(g))
            idx = max(0, min(idx, len(g) - 1))

            how = await choose_radio(page, g[idx]["id"], strict)
            if how == "FAIL":
                flag(step, key, "보기 선택 불가", f"보기 {idx + 1}을 클릭으로 선택할 수 없음")
            chosen_v, chosen_t = g[idx]["value"], g[idx]["text"]

            if g[0]["type"] == "checkbox":
                cnt = 1
                for off in range(1, len(g)):
                    if cnt >= min(need, len(g)):
                        break
                    if await choose_radio(page, g[(idx + off) % len(g)]["id"], strict) != "FAIL":
                        cnt += 1
            for _, ids in (st["otherGroups"] or {}).items():   # 행렬형 나머지 행
                if ids:
                    await choose_radio(page, ids[0], strict)

        elif st["texts"]:
            chosen_v, chosen_t = "text", "(주관식)"
        else:
            chosen_v, chosen_t = "info", "(안내문)"

        for s in st["selects"]:
            if s["n"] > 1 and not s["value"]:
                await page.evaluate(
                    "n => { const e = document.getElementsByName(n)[0]; if (e) { e.selectedIndex = 1; e.dispatchEvent(new Event('change', {bubbles:true})); } }",
                    s["name"],
                )

        # decide 에 fill 이 있으면 주관식 값을 지정한다 (출생년도 등 경계값 테스트용)
        if want.get("fill") is not None and st["texts"]:
            await fill_texts(page, str(want["fill"]), only_empty=False)
            chosen_t = f'입력={want["fill"]}'

        # ── 제출: 적응형 재시도. 텍스트는 막혔을 때만 채운다 ──
        moved, tried = False, []
        for attempt in range(1, 5):
            if attempt == 2:
                await fill_texts(page, "5")
            elif attempt == 3:
                await fill_texts(page, "1", only_empty=False)
            elif attempt == 4:
                await fill_texts(page, "테스트응답", only_empty=False)
            t1 = time.time()
            r = await click_next(page)
            if r["blocked"]:
                msgs = await page.evaluate("() => window.__alerts || []")
                if msgs:
                    tried.append(msgs[-1][:60])
                await page.evaluate("() => { window.__alerts = []; window.__nav = undefined; }")
                continue
            moved = await wait_navigated(page)
            if moved:
                elapsed = (time.time() - t1) * 1000
                if elapsed > SLOW_PAGE_MS:
                    flag(step, key, "페이지 로드 지연", f"{int(elapsed)}ms")
            break

        answers.append({"step": step, "key": key, "qcode": qcode, "value": chosen_v,
                        "text": chosen_t, "how": how, "pct": pct, "next": None})
        if len(answers) > 1:
            answers[-2]["next"] = key

        if not moved:
            ended = "막힘"
            flag(step, key, "진행 불가", " / ".join(tried) or "응답 후에도 이동 실패")
            if shots_dir:
                await page.screenshot(path=str(Path(shots_dir) / f"{plan['id']}_{step}_{key.replace(':','_')}.png"))
            break

        await install_hooks(page)

    if answers:
        answers[-1]["next"] = path[-1] if path else None

    if ended == "종료":
        end_type = "이탈" if DROPOUT_RE.search(end_text) else ("완료" if COMPLETE_RE.search(end_text) else "기타종료")
    else:
        end_type = ended
    # 완주 후 이탈은 조기 이탈과 구분해야 잡힌다
    if end_type == "이탈" and len(path) > 10:
        flag(len(path), path[-2] if len(path) > 1 else "", "후반부 이탈",
             f"{len(path)}단계까지 응답 후 대상아님 종료")

    await page.close()
    return {
        "id": plan["id"], "label": plan.get("label", plan["id"]),
        "steps": len(path), "ms": int((time.time() - t0) * 1000),
        "endType": end_type, "endText": end_text[:120], "path": ">".join(path),
        "answers": answers, "anomalies": anomalies, "reqOk": req_ok, "reqMiss": req_miss,
    }


def aggregate(sessions):
    grouped, transitions, end_types, paths = {}, {}, {}, {}
    for s in sessions:
        end_types[s.get("endType", "ERROR")] = end_types.get(s.get("endType", "ERROR"), 0) + 1
        if s.get("path"):
            paths[s["path"]] = paths.get(s["path"], 0) + 1
        for a in s.get("anomalies", []):
            k = f"{a['type']}|{a.get('key','')}|{str(a.get('detail',''))[:60]}"
            g = grouped.setdefault(k, {"type": a["type"], "key": a.get("key", ""),
                                       "detail": a.get("detail", ""), "count": 0, "sessions": []})
            g["count"] += 1
            if len(g["sessions"]) < 3:
                g["sessions"].append(s["id"])
        for a in s.get("answers", []):
            if a.get("next"):
                transitions.setdefault(a["key"], {}).setdefault(a["value"], {})
                t = transitions[a["key"]][a["value"]]
                t[a["next"]] = t.get(a["next"], 0) + 1
    return {
        "summary": {
            "scenarios": len(sessions),
            "pages": sum(s.get("steps", 0) for s in sessions),
            "rawAlerts": sum(len(s.get("anomalies", [])) for s in sessions),
            "uniqueAlerts": len(grouped),
            "endTypes": end_types,
            "uniquePaths": len(paths),
        },
        "alerts": sorted(grouped.values(), key=lambda x: -x["count"]),
        "transitions": transitions,
        "paths": [{"n": n, "len": p.count(">") + 1, "path": p}
                  for p, n in sorted(paths.items(), key=lambda kv: -kv[1])],
        "sessions": sessions,
    }


async def main_async(args):
    plans = json.loads(Path(args.scenarios).read_text(encoding="utf-8"))
    if args.shots:
        Path(args.shots).mkdir(parents=True, exist_ok=True)

    queue = asyncio.Queue()
    for p in plans:
        queue.put_nowait(p)
    results, done = [], [0]
    t0 = time.time()

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=not args.headed)

        async def worker():
            while not queue.empty():
                try:
                    plan = queue.get_nowait()
                except asyncio.QueueEmpty:
                    break
                # 시나리오마다 새 컨텍스트를 쓴다. 재사용하면 이전 응답의 쿠키가 남아
                # 서버가 중복/만료로 판단해 오류 페이지를 준다 (실측 확인).
                ctx = await browser.new_context(locale="ko-KR")
                try:
                    results.append(await run_scenario(ctx, plan, args.url, args.shots))
                except Exception as e:
                    results.append({"id": plan["id"], "error": str(e)[:200],
                                    "anomalies": [], "answers": [], "steps": 0})
                finally:
                    await ctx.close()
                done[0] += 1
                print(f"  {done[0]}/{len(plans)}  {plan['id']}", file=sys.stderr, flush=True)

        await asyncio.gather(*[worker() for _ in range(min(args.workers, len(plans)))])
        await browser.close()

    out = aggregate(results)
    out["summary"]["seconds"] = round(time.time() - t0, 1)
    Path(args.out).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    s = out["summary"]
    print(f"\n완료: {s['scenarios']} 시나리오 / {s['pages']} 페이지 / {s['seconds']}초", file=sys.stderr)
    print(f"경보 {s['rawAlerts']}건 (고유 {s['uniqueAlerts']}종) · 경로 {s['uniquePaths']}종", file=sys.stderr)
    print(f"종료유형: {s['endTypes']}", file=sys.stderr)
    print(f"→ {args.out}", file=sys.stderr)
    print("\n※ 원시 경보를 그대로 보고하지 말 것. 반드시 재현 검증으로 오탐을 걸러낸다.", file=sys.stderr)


def main():
    ap = argparse.ArgumentParser(description="설문 웹테스트 주행 (Playwright)")
    ap.add_argument("--url", required=True, help="테스트 URL (SurveyTest.asp?alias=...)")
    ap.add_argument("--scenarios", required=True, help="시나리오 JSON 파일")
    ap.add_argument("--out", default="results.json", help="결과 JSON 출력 경로")
    ap.add_argument("--workers", type=int, default=10, help="동시 브라우저 컨텍스트 수")
    ap.add_argument("--shots", default=None, help="이상 발생 시 스크린샷을 저장할 디렉터리")
    ap.add_argument("--headed", action="store_true", help="브라우저를 화면에 띄운다")
    args = ap.parse_args()
    asyncio.run(main_async(args))


if __name__ == "__main__":
    main()
