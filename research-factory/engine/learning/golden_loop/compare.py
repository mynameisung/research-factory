# -*- coding: utf-8 -*-
"""골든 대조기 — 예측(가이드 통계표 기반) vs 정답(제작본 역산).

R-104: 배치·서식·명명·순서는 제작본이 정답, 수치·Base 표기는 아니다.
그래서 여기서 채점하는 축은 **형식**뿐이다.
"""
import json, sys, re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
REFS = ROOT / "engine" / "refs"


def load(tag):
    gt = json.loads((REFS / f"ground_truth_{tag}.json").read_text(encoding="utf-8"))
    pr = json.loads((Path(__file__).parent / "out" / f"pred_{'G1' if tag.startswith('LG') else 'G2'}.json")
                    .read_text(encoding="utf-8"))
    return gt, pr


#: 차트타입별 **보기의 자리**. R-113 — 갈림길은 「보기가 카테고리축인가 계열인가」다.
#: 이걸 안 갈면 정상인 장도 전부 틀렸다고 나온다(실측: 척도 19장 × 3축 = 54건 오탐).
#:
#:   가로막대·겹침막대  보기=범주축  표에 보기 행 O  차트 범주=보기
#:   누적 100%         보기=계열    표에 보기 행 O  차트 범주=「전체」 1개
#:   ±세로막대(척도)    보기=계열    표에 보기 행 X  차트 범주=배너 열
AXIS_SPEC = {
    "bar":     {"opt_axis": "category", "table_has_options": True},
    "rank":    {"opt_axis": "category", "table_has_options": True},
    "stacked": {"opt_axis": "series",   "table_has_options": True},
    "scale":   {"opt_axis": "series",   "table_has_options": False},
}


def chart_family(ct: str) -> str:
    t = ct or ""
    if "누적" in t:
        return "stacked"
    if "세로막대" in t or "+-" in t or "±" in t:
        return "scale"
    if "겹침" in t:
        return "rank"
    return "bar"


def is_unfilled(s):
    """제작본이 **채워지지 않은** 장인가. 정답이 없으므로 채점에서 뺀다.

    ★ 하네스의 근본 전제가 「정답은 제작본」인데, LG 79장 중 25장이 비어 있다.
      빈 장을 정답으로 채점하면 우리가 맞은 것도 틀렸다고 나온다.
      비었다는 사실 자체는 결함이 아니라 **정답 부재**다.
    """
    t = (s.get("tables") or [{}])[0]
    labs = [x for x in (t.get("data_row_labels") or []) if str(x).strip()]
    if not labs:
        return True
    if s.get("unfilled_markers"):
        return True
    # ★ **라벨은 있는데 값만 비어 있는** 장이 따로 있다(G1 s26~29·s83, 5장).
    #   라벨만 보고 「채워짐」으로 판정하면 값이 없는 표를 정답 삼아 채점하게 되고,
    #   정렬 축이 통째로 오탐한다(실측: s83 라벨 17건 차이 → 전부 허깨비).
    rows = t.get("rows") or []
    bri = t.get("base_row_index")
    if bri is not None:
        data = rows[bri + 1:]
        if not data:
            return True
        filled = sum(1 for r in data for c in r[1:]
                     if str(c).strip() and str(c).strip() != "　")
        if filled == 0:
            return True
    return False


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def main(tag, lo, hi):
    gt, pr = load(tag)
    body = [s for s in gt["slides"] if s["classification"] == "본문"]
    window = [s for s in body if lo <= s["slide_no"] <= hi]
    preds = pr["predictions"]

    # 정답 본문 순서 == 예측 순서 로 짝짓는다 (둘 다 가이드/제작본의 진행 순서)
    idx = {s["slide_no"]: i for i, s in enumerate(body)}
    rows, skipped = [], []
    for s in window:
        i = idx[s["slide_no"]]
        p = preds[i] if i < len(preds) else {}
        if is_unfilled(s):
            skipped.append({"slide": s["slide_no"], "code": s.get("question_code"),
                            "why": "제작본 미기입 — 정답이 없다"})
            continue
        rows.append((s, p))

    diffs = []
    geo = []
    shape = {"행수": [], "열수": [], "범례": [], "라벨차": [], "병합": [],
             "서브": [], "멘트": [], "정렬": []}
    for s, p in rows:
        n = s["slide_no"]
        def add(axis, got, want, note=""):
            if norm(got) != norm(want):
                diffs.append({"slide": n, "axis": axis, "예측": got, "정답": want, "note": note})

        fam = chart_family(s.get("derived_chart_type"))
        spec = AXIS_SPEC[fam]
        add("차트타입", p.get("predicted_chart_type"), s.get("derived_chart_type"))
        add("슬라이드제목", p.get("slide_title"), s.get("slide_title"))
        add("Qu", p.get("qu"), s.get("qu_text"))
        add("Base문자열", p.get("base_unit"), s.get("base_unit"))
        # 표 행 순서
        tbls = s.get("tables") or []
        gt_rows = [norm(x) for x in ((tbls[0].get("data_row_labels") or []) if tbls else []) if norm(x)]
        pred_rows = [norm(x) for x in (p.get("table_row_order") or [])]
        if not spec["table_has_options"]:
            pred_rows = []          # 척도 표엔 보기 행이 없다 — 채점 대상 아님
        if gt_rows and pred_rows:
            head = gt_rows[len(gt_rows) - len(pred_rows):] if len(gt_rows) > len(pred_rows) else gt_rows
            if pred_rows != head and pred_rows != gt_rows:
                diffs.append({"slide": n, "axis": "표 행 순서",
                              "예측": pred_rows[:6], "정답": gt_rows[:8],
                              "note": f"예측 {len(pred_rows)}행 vs 정답 {len(gt_rows)}행"})
        # 표 헤더 행 수 — 추출기의 header_rows 는 **(Base) 행을 뺀** 수다.
        # 우리 HEADER_ROWS_TOP=3 은 group+label+base 를 센다. 정의가 다를 뿐 같은 구조라
        # base_row_index 를 더해 맞춘 뒤 비교한다.
        # ★ 척도는 표 구조가 아예 다르다(요약행 + 범례행). 이 축을 걸지 않는다.
        if tbls and fam != "scale":
            hr = tbls[0].get("header_rows")
            n_hr = len(hr) if isinstance(hr, list) else (hr or 0)
            bri = tbls[0].get("base_row_index")
            if bri is not None and bri >= n_hr:
                n_hr = bri + 1
            add("표 헤더행수", 3, n_hr, "(Base) 포함 기준")
        # ---- 기하: 표 총높이는 고정, 헤더가 커지면 데이터 행이 줄어든다 ----
        if tbls:
            t0 = tbls[0]
            hs = t0.get("row_heights_cm") or []
            nhdr = 3 if fam != "scale" else None
            if hs and nhdr and len(hs) > nhdr:
                hdr_h = round(sum(hs[:nhdr]), 2)
                dat_h = round(sum(hs[nhdr:]), 2)
                tot = round((t0.get("height_cm") or 0), 2)
                geo.append({"slide": n, "열": t0.get("n_cols"), "총높이": tot,
                            "헤더": hdr_h, "데이터영역": dat_h,
                            "행높이": round(hs[nhdr], 3) if len(hs) > nhdr else None,
                            "데이터행": len(hs) - nhdr})
            # 데이터 행 수 상한 관찰
            labs = [x for x in (t0.get("data_row_labels") or []) if norm(x)]
            if labs:
                shape["행수"].append(len(labs))
            shape["열수"].append(t0.get("n_cols"))
        c_leg = (s.get("charts") or [{}])[0].get("has_legend")
        shape["범례"].append((fam, c_leg))

        # ---- 보기 라벨 표기 — 통계표 라벨을 그대로 쓰는가, 손질하는가 ----
        if spec["table_has_options"] and gt_rows and p.get("table_row_order"):
            for a, b in zip([norm(x) for x in p["table_row_order"]], gt_rows):
                if a != b:
                    shape["라벨차"].append({"slide": n, "예측": a, "정답": b})
        # ---- 배너 병합 구조 ----
        if tbls:
            grp = tbls[0].get("banner_group_row") or []
            runs, cur, cnt = [], None, 0
            for cell in grp:
                v = norm(cell)
                if v:
                    if cur is not None:
                        runs.append((cur, cnt))
                    cur, cnt = v, 1
                else:
                    cnt += 1
            if cur is not None:
                runs.append((cur, cnt))
            shape["병합"].append({"slide": n, "그룹": [(g, c) for g, c in runs if g not in ("　", "")]})
        # ---- 서브타이틀 명명 ----
        shape["서브"].append({"slide": n, "정답": norm(s.get("sub_title")),
                              "제목": norm(s.get("slide_title"))})
        # ---- 멘트 ----
        tm = norm(s.get("top_ment"))
        if tm:
            shape["멘트"].append({"slide": n, "줄수": len([x for x in (s.get("top_ment") or "").split(chr(10)) if x.strip()]),
                                  "글자": len(tm), "끝": tm[-6:]})
        # ---- 값 정렬 ----
        if spec["opt_axis"] == "category" and p.get("option_values"):
            vs = [v for v in p["option_values"] if isinstance(v, (int, float))]
            if len(vs) >= 3:
                shape["정렬"].append("내림차순" if vs == sorted(vs, reverse=True)
                                     else ("오름차순" if vs == sorted(vs) else "통계표순"))

        # 보기가 차트의 어디로 가는가 — 범주축인가 계열인가
        # ★ 누적 100% 막대는 **보기가 계열**이다(범주는 「전체」 하나).
        #   보기 수를 범주 수와 비교하면 정상인 것도 틀렸다고 나온다.
        chs = s.get("charts") or []
        if chs and p.get("option_order"):
            c0 = chs[0]
            as_series = spec["opt_axis"] == "series"
            n_axis = c0.get("real_series_count") or c0.get("series_count")                 if as_series else c0.get("category_count")
            axis_name = f"보기 수({'계열' if as_series else '범주'})"
            if n_axis and n_axis != len(p["option_order"]):
                diffs.append({"slide": n, "axis": axis_name,
                              "예측": len(p["option_order"]), "정답": n_axis,
                              "note": f"{fam}"})
            if fam == "stacked":
                # R-068: 누적은 통계표 역순으로 쌓는다.
                # ★ 예측기의 option_order 는 **이미 계열 순서**(역순 적용 후)다.
                #   여기서 또 뒤집으면 맞는 것도 틀렸다고 나온다.
                names = [norm(x) for x in (c0.get("series_names") or [])]
                want = [norm(x) for x in p["option_order"]]
                if names and names != want:
                    diffs.append({"slide": n, "axis": "누적 적층 순서",
                                  "예측": want, "정답": names, "note": "R-068"})
    import collections as _c
    out = {"tag": tag, "범위": f"s{lo}~s{hi}", "대조": len(rows),
           "정답없음": len(skipped), "차이": len(diffs),
           "관찰": {"행수_최대": max(shape["행수"]) if shape["행수"] else None,
                    "행수_분포": dict(sorted(_c.Counter(shape["행수"]).items())),
                    "열수_최대": max([x for x in shape["열수"] if x] or [0]),
                    "열수_분포": dict(sorted(_c.Counter(x for x in shape["열수"] if x).items())),
                    "범례": {f"{k[0]}/{k[1]}": v for k, v in _c.Counter(shape["범례"]).items()},
                    "라벨차": shape["라벨차"][:20], "라벨차_수": len(shape["라벨차"]),
                    "병합": shape["병합"][:6],
                    "서브": shape["서브"][:20],
                    "멘트": shape["멘트"][:20],
                    "정렬": dict(_c.Counter(shape["정렬"]))},
           "기하": geo, "not_checked": skipped, "diffs": diffs}
    # ★ 결과는 **파일로** 쓴다. 표준출력으로 넘기면 길이 제한에 잘려
    #   다음 단계의 json.loads 가 깨진다(실측: 6000자에서 절단).
    outdir = Path(__file__).parent / "out"
    outdir.mkdir(exist_ok=True)
    f = outdir / f"cmp_{'G1' if tag.startswith('LG') else 'G2'}_{lo}_{hi}.json"
    f.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")

    import collections
    c = collections.Counter(x["axis"] for x in diffs)
    print(f"범위 {out['범위']} · 대조 {out['대조']}장 · 정답없음 {out['정답없음']}장 · 차이 {len(diffs)}건")
    for k, v in c.most_common():
        print(f"   {k:<18} {v}건")
    o = out["관찰"]
    print(f"   관찰 · 데이터행 최대 {o['행수_최대']} · 열 최대 {o['열수_최대']} · 범례 {o['범례']}")
    if geo:
        tots = {g["총높이"] for g in geo}
        print(f"   기하 · 표 총높이 {sorted(tots)} · 헤더 {sorted({g['헤더'] for g in geo})}"
              f" · 행높이 {sorted({g['행높이'] for g in geo})}")
    if o.get("라벨차_수"):
        print(f"   라벨 표기 차이 {o['라벨차_수']}건")
    if o.get("정렬"):
        print(f"   값 정렬 {o['정렬']}")
    if o.get("멘트"):
        ls = [m["줄수"] for m in o["멘트"]]
        cs = [m["글자"] for m in o["멘트"]]
        print(f"   멘트 {len(ls)}장 · 줄수 {sorted(set(ls))} · 글자 {min(cs)}~{max(cs)}")
    print(f"→ {f}")
    return out


if __name__ == "__main__":
    main(sys.argv[1], int(sys.argv[2]), int(sys.argv[3]))
