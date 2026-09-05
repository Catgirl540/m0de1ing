# -*- coding: utf-8 -*-
"""solve_q1.py — 问题一 MILP（模型卡片-模型I §5）
case1 滞销 / case2 五折；输出 plan CSV 与图表数据 CSV。
运行：python solve_q1.py（依赖 load_data.py 生成的 data.json）
季槽表示：(plot, year, season) 元组；2023 年槽为常量（fixed 集合），空季=空集（休耕断开重茬链）。
"""
import csv
import json
from pathlib import Path
import numpy as np
from scipy import sparse
from scipy.optimize import milp, LinearConstraint, Bounds

HERE = Path(__file__).resolve().parent
D = json.loads((HERE / "data.json").read_text(encoding="utf-8"))
YEARS, PLOTS, CROPS = D["years"], D["plots"], D["crops"]
COMPAT = {tuple(k.split("|")): v for k, v in D["compat"].items()}
PARAMS = {tuple(int(a) if a.isdigit() else a for a in k.split("|")): v
          for k, v in D["params"].items()}
BEANS, ROOTS, RICE = set(D["beans"]), set(D["roots"]), D["rice"]
MIN_A = D["min_area"]
DEMAND = {int(k): v for k, v in D["demand"].items()}
PLOT = {p["name"]: p for p in PLOTS}
GH = {p["name"]: "大棚" in p["type"] for p in PLOTS}


def season_slots(ptype):
    if ptype in ("平旱地", "梯田", "山坡地"):
        return ["单季"]
    return ["第一季", "第二季"]


def build_chain():
    """每地块时序季槽链；2023 槽含实种作物集合（缺=空集=休耕）"""
    chains, fixed = {}, {}
    for p in PLOTS:
        name, ptype = p["name"], p["type"]
        h = D["hist23"].get(name, {})
        seq = []
        if ptype in ("平旱地", "梯田", "山坡地"):
            slots23 = [("单季", set(h.get("单季", [])))]
        else:
            if ptype == "水浇地" and "单季" in h:      # 水稻整年
                slots23 = [("第一季", set(h["单季"])), ("第二季", set())]
            else:
                slots23 = [(s, set(h.get(s, []))) for s in ("第一季", "第二季")]
        for s, cs in slots23:
            k = (name, 2023, s)
            seq.append(k)
            if cs:
                fixed[k] = cs
        for t in YEARS:
            for s in season_slots(ptype):
                seq.append((name, t, s))
        chains[name] = seq
    return chains, fixed


CHAINS, FIXED = build_chain()


def compat_crops(plot_name, season):
    return COMPAT.get((PLOT[plot_name]["type"], season), [])


def ycp(cid, plot_name, season):
    t = PLOT[plot_name]["type"]
    key = (cid, t, season)
    if key not in PARAMS and cid == RICE:
        key = (cid, t, "单季")          # 统计表里水稻记为"单季"
    if key not in PARAMS and t == "智慧大棚":
        key = (cid, t, "第二季")        # 统计表智慧大棚只录一组参数，两季共用
    return PARAMS[key]


def solve(case, time_limit=240, gap=0.005, scen=None):
    """scen: {'demand_m'/'yield_m'/'cost_m'/'price_m': {(cid, year): 乘数}}，缺省 1.0（问题一口径）"""
    meta = []
    for p in PLOTS:
        for k in CHAINS[p["name"]]:
            if k[1] < YEARS[0]:          # 2023 历史季槽不建变量（含空=休耕）
                continue
            for c in compat_crops(k[0], k[2]):
                # case2 恒等变换：收入 = 0.5·p·q + 0.5·p·min(q,D) → 消掉降价段变量 σ2，
                # 只保留 σ1 = min(q,D)，LP 松弛更紧（直接消元，不再建 s2 变量）
                kinds = ("x", "y", "s1")
                for kind in kinds:
                    meta.append((kind, c, k))
    idx = {m: i for i, m in enumerate(meta)}
    NV = len(meta)

    def mult(key, c, t):
        return scen[key].get((c, t), 1.0) if scen else 1.0

    obj = np.zeros(NV); ub = np.zeros(NV); integ = np.zeros(NV)
    for i, (kind, c, k) in enumerate(meta):
        par = ycp(c, k[0], k[2]); A = PLOT[k[0]]["area"]
        if kind == "x":
            ub[i] = A
            obj[i] = par["cost"] * mult("cost_m", c, k[1]) \
                if case == 1 else \
                par["cost"] * mult("cost_m", c, k[1]) \
                - 0.5 * par["price"] * mult("price_m", c, k[1]) \
                * par["yield"] * mult("yield_m", c, k[1])
        elif kind == "y":
            ub[i] = 1.0; integ[i] = 1
        else:
            ub[i] = DEMAND[c] * mult("demand_m", c, k[1])
            obj[i] = -par["price"] * mult("price_m", c, k[1]) \
                if case == 1 else -0.5 * par["price"] * mult("price_m", c, k[1])

    rows, cols, vals, hi = [], [], [], []
    nv = [0]
    def add_row(terms, upper):
        for j, v in terms:
            rows.append(nv[0]); cols.append(j); vals.append(v)
        hi.append(upper); nv[0] += 1

    # (1-1)(1-2) 面积上限 + x-y 联动（含最小面积；水稻=整地块）
    for p in PLOTS:
        for k in CHAINS[p["name"]]:
            if k[1] < YEARS[0]:
                continue
            A = p["area"]
            add_row([(idx[("x", c, k)], 1.0) for c in compat_crops(k[0], k[2])], A)
            for c in compat_crops(k[0], k[2]):
                m = A if (c == RICE and p["type"] == "水浇地") else MIN_A["大棚" if GH[k[0]] else "露地"]
                add_row([(idx[("x", c, k)], -1.0), (idx[("y", c, k)], m)], 0.0)
                add_row([(idx[("x", c, k)], 1.0), (idx[("y", c, k)], -A)], 0.0)

    # (1-3) 重茬：相邻季槽同作物禁止；2023 槽为常量（空槽=休耕，不约束）
    for p in PLOTS:
        ch = CHAINS[p["name"]]
        for a, b in zip(ch[:-1], ch[1:]):
            shared = set(compat_crops(a[0], a[2])) & set(compat_crops(b[0], b[2]))
            fa, fb = FIXED.get(a, set()), FIXED.get(b, set())
            a_free, b_free = a[1] >= YEARS[0], b[1] >= YEARS[0]
            if not a_free and not b_free:
                continue
            for c in shared:
                if a_free and b_free:
                    add_row([(idx[("y", c, a)], 1.0), (idx[("y", c, b)], 1.0)], 1.0)
                elif a_free:                  # b 是 2023 常量
                    if c in fb:
                        add_row([(idx[("y", c, a)], 1.0)], 0.0)
                elif c in fa:                 # a 是 2023 常量
                    add_row([(idx[("y", c, b)], 1.0)], 0.0)

    # (1-4) 豆类三年滑窗
    for p in PLOTS:
        for w in range(YEARS[0], YEARS[-1] - 1):
            terms = []
            for t in (w, w + 1, w + 2):
                for k in CHAINS[p["name"]]:
                    if k[1] == t and k not in FIXED:
                        for c in set(BEANS) & set(compat_crops(k[0], k[2])):
                            terms.append((idx[("y", c, k)], 1.0))
            if terms:
                add_row(terms, np.inf)

    # 分散度：单(作物,年,季) ≤ 8 地块
    from collections import defaultdict
    by_slot = defaultdict(list)
    for (kind, c, k) in meta:
        if kind == "y":
            by_slot[(c, k[1], k[2])].append(k)
    for (c, t, s), ks in by_slot.items():
        if len(ks) > 1:
            add_row([(idx[("y", c, k)], 1.0) for k in ks], D["max_plots_per_crop_season"])

    # 水稻整年：y_rice(S1,t) + y_c(S2,t) ≤ 1
    for p in PLOTS:
        if p["type"] != "水浇地":
            continue
        for t in YEARS:
            k1, k2 = (p["name"], t, "第一季"), (p["name"], t, "第二季")
            for c in compat_crops(p["name"], "第二季"):
                add_row([(idx[("y", RICE, k1)], 1.0), (idx[("y", c, k2)], 1.0)], 1.0)

    # 产量-销量线性化 (1-5)：σ1 ≤ min(γ·x, D)；case2 的 0.5·p·q 项已并入 x 的目标系数
    for (kind, c, k) in meta:
        if kind == "s1":
            g = ycp(c, k[0], k[2])["yield"] * mult("yield_m", c, k[1])
            add_row([(idx[("s1", c, k)], 1.0), (idx[("x", c, k)], -g)], 0.0)

    A = sparse.csc_matrix((vals, (rows, cols)), shape=(nv[0], NV))
    res = milp(c=obj, constraints=LinearConstraint(A, -np.inf, np.array(hi)),
               integrality=integ, bounds=Bounds(np.zeros(NV), ub),
               options={"time_limit": time_limit, "mip_rel_gap": gap, "disp": False})
    print(f"[case{case}] status={res.status} ({res.message})  目标值={-res.fun:,.0f} 元")
    return res, meta, idx


def extract(res, meta, case):
    rows = [{"年份": k[1], "地块": k[0], "季次": k[2], "作物编号": c,
             "面积": round(float(res.x[i]), 4)}
            for i, (kind, c, k) in enumerate(meta)
            if kind == "x" and res.x[i] > 1e-6]
    with open(HERE / f"plan_case{case}.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["年份", "地块", "季次", "作物编号", "面积"])
        w.writeheader(); w.writerows(rows)
    return rows


def summary(rows, case):
    cname = {c["id"]: c["name"] for c in CROPS}
    def ctype(cid):
        t = next(c["type"] for c in CROPS if c["id"] == cid)
        return "食用菌" if t == "食用菌" else ("粮食" if t.startswith("粮食") else "蔬菜")
    per_year, struct = {}, {}
    for r in rows:
        par = ycp(r["作物编号"], r["地块"], r["季次"])
        q = r["面积"] * par["yield"]
        sold = min(q, DEMAND[r["作物编号"]])
        rev = par["price"] * sold + (0.5 * par["price"] * (q - sold) if case == 2 else 0.0)
        per_year[r["年份"]] = per_year.get(r["年份"], 0.0) + rev - par["cost"] * r["面积"]
        struct[(r["年份"], ctype(r["作物编号"]))] = \
            struct.get((r["年份"], ctype(r["作物编号"])), 0.0) + r["面积"]
    with open(HERE / f"F01_历年收入对比_case{case}.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["年份", f"case{case}_净收益_元"])
        for y in YEARS:
            w.writerow([y, round(per_year.get(y, 0.0), 1)])
    with open(HERE / f"F02_面积结构_case{case}.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["年份", "作物类型", "面积亩"])
        for (y, t), v in sorted(struct.items()):
            w.writerow([y, t, round(v, 2)])
    return per_year


if __name__ == "__main__":
    out = {}
    for case in (1, 2):
        res, meta, idx = solve(case)
        rows = extract(res, meta, case)
        py = summary(rows, case)
        out[f"case{case}_total"] = round(sum(py.values()), 1)
        out[f"case{case}_by_year"] = {str(k): round(v, 1) for k, v in py.items()}
        out[f"case{case}_n_plantings"] = len(rows)
    with open(HERE / "Q1_结论.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["指标", "数值"])
        for k, v in out.items():
            if "by_year" in k:
                for y, val in v.items():
                    w.writerow([f"{k}_{y}", val])
            else:
                w.writerow([k, v])
    print(json.dumps(out, ensure_ascii=False, indent=1))
