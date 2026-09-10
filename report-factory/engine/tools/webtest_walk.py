#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""webtest_walk.py — 웹조사(한 페이지 한 문항)를 시나리오대로 순회하며 페이지를 덤프한다.

연구원이 브라우저에서 「다음 → 다음 → …」 하며 눈으로 보던 것을 Playwright 가 대신 걷는다.
사람이 하는 일은 로그인뿐이다(엠브레인 패널 계정). 순회는 기계가, 판단은 대조기(webtest_compare.py)와 사람이 한다.

사용
    python webtest_walk.py <URL> <시나리오.json> <out_dir> [--scenario S0] [--profile generic|isurvey]
                           [--headed] [--login-wait 90] [--storage-state auth.json] [--save-state auth.json]
                           [--screenshots] [--save-html 3] [--max-pages 400]

    --headed --login-wait 90   브라우저를 띄우고 90초 동안 사람이 로그인·조사 진입할 시간을 준다
    --save-state auth.json     로그인 뒤 세션을 저장해 두면 다음부터 --storage-state 로 로그인 없이 돈다
    --save-html N              처음 N 페이지의 HTML 을 남긴다 — 사이트 선택자를 맞출 때 본다

산출: <out_dir>/덤프_<시나리오>.json  {scenario, url, pages:[{n, url, code_guess, text, options, inputs, images, answered, next_clicked}], ended_by}
      (--screenshots 면 <out_dir>/캡처_<시나리오>/p<n>.png)

주의: 뒤로가기가 안 되는 엔진이 많다. 시나리오 하나가 틀리면 처음부터 다시 돈다 — 그래서 이 도구가 필요하다.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
from pathlib import Path

RE_CODE = re.compile(r"(?<![A-Za-z가-힣])(SQ\d+|[A-K]{1,2}\d+(?:-\d+)?)(?=[\.\)\s:：])")
PROFILES = {
    "generic": {
        "next": ["button:has-text('다음')", "input[type=submit]", "input[type=button][value*='다음']", "a:has-text('다음')",
                 "button:has-text('NEXT')", "#next", ".next", "button[type=submit]"],
        "end_keywords": ["설문이 종료", "종료되었습니다", "참여해 주셔서 감사", "감사합니다", "대상이 아닙", "조사 대상이 아", "설문을 마쳤", "완료되었습니다"],
        "wait_ms": 400,
    },
    # 엠브레인 isurvey (classic ASP). 선택자는 첫 순회의 --save-html 로 맞춘 뒤 여기에 고정한다.
    "isurvey": {
        "next": ["input[type=image]", "input[type=submit]", "input[type=button][value*='다음']", "a:has-text('다음')", "img[alt*='다음']", "button:has-text('다음')"],
        "end_keywords": ["설문이 종료", "종료되었습니다", "참여해 주셔서 감사", "감사합니다", "대상이 아닙", "조사 대상이 아", "설문을 마쳤", "완료되었습니다"],
        "wait_ms": 600,
    },
}

JS_EXTRACT = """
() => {
  const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el); return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const lab = el => {
    let t = '';
    if (el.labels && el.labels.length) t = el.labels[0].innerText;
    if (!t) { const l = el.closest('label'); if (l) t = l.innerText; }
    if (!t && el.id) { const l = document.querySelector('label[for="' + el.id + '"]'); if (l) t = l.innerText; }
    if (!t) { const p = el.parentElement; if (p) t = p.innerText; }
    return (t || '').replace(/\\s+/g, ' ').trim().slice(0, 200);
  };
  const options = [];
  document.querySelectorAll('input[type=radio], input[type=checkbox]').forEach(el => {
    if (!vis(el)) return;
    options.push({ kind: el.type, name: el.name, value: el.value, label: lab(el), checked: el.checked });
  });
  document.querySelectorAll('select').forEach(el => {
    if (!vis(el)) return;
    Array.from(el.options).forEach((o, i) => { if (o.value !== '') options.push({ kind: 'select', name: el.name, value: o.value, label: o.text.trim(), checked: o.selected }); });
  });
  const inputs = [];
  document.querySelectorAll('input[type=text], input[type=number], input[type=tel], textarea').forEach(el => {
    if (!vis(el)) return;
    inputs.push({ name: el.name, placeholder: el.placeholder || '', maxlength: el.maxLength > 0 ? el.maxLength : null, value: el.value });
  });
  const images = [];
  document.querySelectorAll('img').forEach(el => { if (vis(el) && el.naturalWidth > 40) images.push({ src: el.getAttribute('src') || '', alt: el.alt || '' }); });
  return { text: document.body.innerText.replace(/\\s+\\n/g, '\\n').trim(), options, inputs, images, title: document.title };
}
"""


def log(msg: str) -> None:
    print(msg, flush=True)


def page_snapshot(page) -> dict:
    return page.evaluate(JS_EXTRACT)


def guess_code(text: str) -> str | None:
    head = text[:400]
    m = RE_CODE.search(head)
    return m.group(1) if m else None


def choose_answers(page, snap: dict, spec: dict | None, code: str | None) -> list[str]:
    """시나리오가 정한 답(보기 번호·입력값)을 페이지에 넣는다. 없으면 첫 보기. 무엇을 골랐는지 라벨 목록을 돌려준다."""
    done = []
    opts = snap["options"]
    radios = [o for o in opts if o["kind"] == "radio"]
    checks = [o for o in opts if o["kind"] == "checkbox"]
    selects = [o for o in opts if o["kind"] == "select"]
    want = list((spec or {}).get("choose") or [])
    text_val = (spec or {}).get("text")

    def click_nth(kind: str, group: list, nos: list[int]) -> None:
        """group 은 (페이지 전체 인덱스, 보기) 목록. nos 는 그 그룹 안의 보기 번호(1부터)."""
        loc = page.locator(f"input[type={kind}]:visible")
        for no in nos:
            idx = no - 1
            if 0 <= idx < len(group):
                gi, o = group[idx]
                try:
                    loc.nth(gi).check(force=True, timeout=3000)
                    done.append(o["label"] or f"{kind}#{no}")
                except Exception as e:
                    log(f"    ! 보기 {no} 선택 실패: {e}")

    def by_name(items: list) -> list[list]:
        """name 별로 묶는다(행렬 문항은 행마다 name 이 다르다). 순서는 페이지 순."""
        groups: dict[str, list] = {}
        for gi, o in enumerate(items):
            groups.setdefault(o.get("name") or "", []).append((gi, o))
        return list(groups.values())

    if radios:
        for grp in by_name(radios):                 # 행렬이면 행마다 하나씩 고른다
            click_nth("radio", grp, want[:1] if want else [1])
    elif checks:
        grps = by_name(checks)
        if len(grps) == len(checks):                # 보기마다 name 이 다른 엔진 → 한 묶음으로 본다
            grps = [[(gi, o) for gi, o in enumerate(checks)]]
        for grp in grps:
            click_nth("checkbox", grp, want if want else [1])
    elif selects:
        sel = page.locator("select:visible").first
        idx = (want[0] - 1) if want else 0
        try:
            sel.select_option(index=idx)
            done.append(selects[idx]["label"] if idx < len(selects) else f"select#{idx+1}")
        except Exception as e:
            log(f"    ! select 실패: {e}")
    if snap["inputs"]:
        val = text_val if text_val is not None else ("1" if not radios and not checks else "")
        if val:
            for i, _ in enumerate(snap["inputs"]):
                try:
                    page.locator("input[type=text]:visible, input[type=number]:visible, input[type=tel]:visible, textarea:visible").nth(i).fill(str(val), timeout=3000)
                    done.append(f"입력={val}")
                except Exception as e:
                    log(f"    ! 입력 실패: {e}")
    return done


def find_next(page, profile: dict):
    for sel in profile["next"]:
        loc = page.locator(sel)
        try:
            if loc.count() and loc.first.is_visible():
                return loc.first
        except Exception:
            continue
    return None


def click_next(page, profile: dict) -> bool:
    loc = find_next(page, profile)
    if loc is None:
        return False
    try:
        loc.click(timeout=4000)
        return True
    except Exception:
        return False


def is_end(page, snap: dict, profile: dict) -> str | None:
    """입력 요소가 없는 페이지만 종료 후보다. 종료 문구가 있으면 종료, 다음 버튼마저 없으면 막다른 페이지."""
    if snap["options"] or snap["inputs"]:
        return None
    for k in profile["end_keywords"]:
        if k in snap["text"]:
            return f"종료 문구 「{k}」"
    if find_next(page, profile) is None:
        return "입력 요소도 다음 버튼도 없음"
    return None


def walk(url: str, scenario: dict, out_dir: Path, profile_name: str, headed: bool, login_wait: int,
         storage_state: str | None, save_state: str | None, screenshots: bool, save_html: int, max_pages: int) -> dict:
    from playwright.sync_api import sync_playwright
    prof = PROFILES[profile_name]
    sid = scenario["id"]
    cap_dir = out_dir / f"캡처_{sid}"
    if screenshots:
        cap_dir.mkdir(parents=True, exist_ok=True)
    result = {"scenario": sid, "url": url, "profile": profile_name, "pages": [], "ended_by": None, "started": time.strftime("%Y-%m-%d %H:%M:%S")}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not headed)
        ctx = browser.new_context(storage_state=storage_state) if storage_state and Path(storage_state).exists() else browser.new_context()
        page = ctx.new_page()
        page.goto(url, wait_until="domcontentloaded")
        if headed and login_wait > 0:
            log(f"  브라우저에서 로그인하고 첫 문항 화면까지 진입하세요 — {login_wait}초 뒤 순회를 시작합니다.")
            time.sleep(login_wait)
        if save_state:
            ctx.storage_state(path=save_state)
            log(f"  세션 저장 → {save_state}")
        stuck = 0
        for n in range(1, max_pages + 1):
            page.wait_for_timeout(prof["wait_ms"])
            snap = page_snapshot(page)
            h = hashlib.sha1(snap["text"].encode("utf-8", "replace")).hexdigest()[:10]
            code = guess_code(snap["text"])
            rec = {"n": n, "url": page.url, "code_guess": code, "text": snap["text"][:4000], "options": snap["options"],
                   "inputs": snap["inputs"], "images": snap["images"], "answered": [], "next_clicked": False}
            if screenshots:
                try:
                    page.screenshot(path=str(cap_dir / f"p{n:03d}.png"), full_page=True); rec["screenshot"] = f"캡처_{sid}/p{n:03d}.png"
                except Exception:
                    pass
            if n <= save_html:
                (out_dir / f"html_{sid}_p{n:03d}.html").write_text(page.content(), encoding="utf-8")
            end = is_end(page, snap, prof)
            if end:
                rec["end"] = end; result["pages"].append(rec); result["ended_by"] = end
                log(f"  p{n:03d} {code or '-':<8} 끝: {end}")
                break
            spec = scenario["answers"].get(code) if code else None
            rec["answered"] = choose_answers(page, snap, spec, code)
            clicked = click_next(page, prof)
            rec["next_clicked"] = clicked
            result["pages"].append(rec)
            log(f"  p{n:03d} {code or '-':<8} 보기 {len(snap['options'])} 입력 {len(snap['inputs'])} → {('다음' if clicked else '다음 없음')}  {rec['answered'][:2]}")
            if not clicked:
                result["ended_by"] = "다음 버튼을 찾지 못함"
                break
            try:
                page.wait_for_load_state("domcontentloaded", timeout=8000)
            except Exception:
                pass
            page.wait_for_timeout(prof["wait_ms"])
            new_hash = hashlib.sha1(page_snapshot(page)["text"].encode("utf-8", "replace")).hexdigest()[:10]
            if new_hash == h:
                stuck += 1
                if stuck >= 2:
                    result["ended_by"] = f"p{n} 에서 화면이 바뀌지 않음 (필수 응답 미충족이거나 선택자 불일치)"
                    break
            else:
                stuck = 0
        else:
            result["ended_by"] = f"최대 페이지 {max_pages} 도달"
        if save_state:
            ctx.storage_state(path=save_state)
        browser.close()
    (out_dir / f"덤프_{sid}.json").write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
    return result


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("url"); ap.add_argument("scenarios"); ap.add_argument("out_dir")
    ap.add_argument("--scenario", default="S0", help="시나리오 ID (또는 all)")
    ap.add_argument("--profile", default="generic", choices=sorted(PROFILES))
    ap.add_argument("--headed", action="store_true"); ap.add_argument("--login-wait", type=int, default=0)
    ap.add_argument("--storage-state"); ap.add_argument("--save-state")
    ap.add_argument("--screenshots", action="store_true"); ap.add_argument("--save-html", type=int, default=0)
    ap.add_argument("--max-pages", type=int, default=400)
    a = ap.parse_args(argv)
    doc = json.loads(Path(a.scenarios).read_text(encoding="utf-8"))
    out = Path(a.out_dir); out.mkdir(parents=True, exist_ok=True)
    todo = doc["scenarios"] if a.scenario == "all" else [s for s in doc["scenarios"] if s["id"] == a.scenario]
    if not todo:
        print(f"시나리오 {a.scenario} 없음. 있는 것: {[s['id'] for s in doc['scenarios']]}"); return 2
    for s in todo:
        log(f"▶ {s['id']} — {s['purpose']}")
        r = walk(a.url, s, out, a.profile, a.headed, a.login_wait, a.storage_state, a.save_state, a.screenshots, a.save_html, a.max_pages)
        log(f"  페이지 {len(r['pages'])} · 끝: {r['ended_by']} → 덤프_{s['id']}.json")
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    raise SystemExit(main())
