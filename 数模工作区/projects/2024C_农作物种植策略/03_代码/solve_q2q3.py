# -*- coding: utf-8 -*-
"""solve_q2q3.py — 问题二（情景规划）与问题三（相关模拟对比）
口径：Q2/Q3 目标沿用问题一情形(2)（超产五折）。
Q2：题给区间中枢路径 → 最优方案（result2.xlsx）；20 个均匀抽样情景评估方案收益，
    3 个情景重优化对比。
Q3：高斯相关模拟（蔬菜类需求共享市场因子 ρ≈0.6；需求冲击与蔬菜价格涨幅负相关 −0.2）
    ×8 组，每组重优化，与"沿用 Q2 中枢方案"比较。
运行：python solve_q2q3.py   （随机种子固定，结果可复现）
"""
import csv
import json
from pathlib import Path
import numpy as np
import solve_q1 as S

HERE = Path(__file__).resolve().parent
OUT = HERE.parent / "06_交付"
rng = np.random.default_rng(20240905)
YEARS = S.YEARS
CIDS = [c["id"] for c in S.CROPS]
VEG = [c["id"] for c in S.CROPS if c["type"].startswith("蔬菜")]
MUSH = [c["id"] for c in S.CROPS if c["type"] == "食用菌"]
WHEAT, CORN, YDZ = 6, 7, 41


def mults_from_rates(rates):
    """rates: {c: dict(demand_r, cost_r, price_r)} → 年度乘数路径"""
    dm, ym, cm, pm = {}, {}, {}, {}
    for c in CIDS:
        rd, rc, rp = rates[c]
        for i, t in enumerate(YEARS):
            k = i + 1
            dm[(c, t)] = (1 + rd) ** k
            ym[(c, t)] = float(rng.uniform(0.9, 1.1))     # 亩产量每年 ±10% 独立波动
            cm[(c, t)] = (1 + rc) ** k
            pm[(c, t)] = (1 + rp) ** k
    return {"demand_m": dm, "yield_m": ym, "cost_m": cm, "price_m": pm}


def sample_rates(correlated=False):
    rates = {}
    f_shared = float(rng.normal()) if correlated else 0.0
    for c in CIDS:
        if c in (WHEAT, CORN):
            rd = float(rng.uniform(0.05, 0.10)); rp = 0.0
        elif c in VEG:
            shock = 0.6 * f_shared + 0.8 * float(rng.normal()) if correlated else 0.0
            rd = max(-0.05, min(0.05, 0.02 + 0.02 * shock)) if correlated \
                else float(rng.uniform(-0.05, 0.05))
            rp = max(-0.02, 0.05 - 0.2 * shock) if correlated else float(rng.uniform(0.04, 0.06))
        elif c in MUSH:
            rd = float(rng.uniform(-0.05, 0.05))
            rp = -0.05 if c == YDZ else float(-rng.uniform(0.01, 0.05))
        else:
            rd = float(rng.uniform(-0.05, 0.05)); rp = 0.0
        rates[c] = (rd, float(rng.uniform(0.04, 0.06)), rp)
    return rates


def mid_rates():
    rates = {}
    for c in CIDS:
        if c in (WHEAT, CORN): rd = 0.075
        elif c in VEG: rd = 0.0; rp = 0.05
        elif c in MUSH: rp = -0.05 if c == YDZ else -0.03; rd = 0.0
        else: rd = 0.0; rp = 0.0
        rates[c] = (rd, 0.05, rp)
    return rates


def extract_to(res, meta, path):
    rows = [{"年份": k[1], "地块": k[0], "季次": k[2], "作物编号": c,
             "面积": round(float(res.x[i]), 4)}
            for i, (kind, c, k) in enumerate(meta)
            if kind == "x" and res.x[i] > 1e-6]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["年份", "地块", "季次", "作物编号", "面积"])
        w.writeheader(); w.writerows(rows)
    return rows


def eval_plan(rows, scen):
    """固定方案在情景参数下的净收益"""
    total = 0.0
    for r in rows:
        c, t = r["作物编号"], r["年份"]
        par = S.ycp(c, r["地块"], r["季次"])
        y = par["yield"] * scen["yield_m"].get((c, t), 1.0)
        p = par["price"] * scen["price_m"].get((c, t), 1.0)
        cst = par["cost"] * scen["cost_m"].get((c, t), 1.0)
        d = S.DEMAND[c] * scen["demand_m"].get((c, t), 1.0)
        q = r["面积"] * y
        total += p * min(q, d) + 0.5 * p * max(q - d, 0) - cst * r["面积"]
    return total


if __name__ == "__main__":
    log = {}
    # ---------- Q2 中枢 ----------
    mid = mults_from_rates(mid_rates())
    res, meta, idx = S.solve(2, time_limit=300, gap=0.005, scen=mid)
    plan_q2 = extract_to(res, meta, HERE / "plan_q2.csv")
    log["Q2_中枢_总净收益"] = round(-res.fun, 1)
    log["Q2_中枢_求解状态"] = int(res.status)

    # 回填 result2.xlsx
    import fill_result as FR
    FR.fill(HERE / "plan_q2.csv", HERE / "结果模板/result2.xlsx", OUT / "result2.xlsx")

    # 20 情景评估 + 3 情景重优化
    fixed_profits, reopt_profits = [], []
    for i in range(20):
        sc = mults_from_rates(sample_rates(correlated=False))
        fixed_profits.append(eval_plan(plan_q2, sc))
        if i < 3:
            r2, m2, i2 = S.solve(2, time_limit=180, gap=0.02, scen=sc)
            reopt_profits.append(-r2.fun)
    log["Q2_20情景_沿用方案_均值"] = round(float(np.mean(fixed_profits)), 1)
    log["Q2_20情景_沿用方案_标准差"] = round(float(np.std(fixed_profits)), 1)
    log["Q2_3情景_重优化_均值"] = round(float(np.mean(reopt_profits)), 1)

    # ---------- Q3 相关模拟 ----------
    q3_rows = []
    for i in range(8):
        sc = mults_from_rates(sample_rates(correlated=True))
        fixed_p = eval_plan(plan_q2, sc)
        r3, m3, i3 = S.solve(2, time_limit=120, gap=0.03, scen=sc)
        q3_rows.append([i + 1, round(fixed_p, 1), round(-r3.fun, 1)])
    with open(HERE / "F05_问题三对比.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["模拟组", "沿用Q2方案_净收益", "Q3重优化_净收益"])
        w.writerows(q3_rows)
    fx = [r[1] for r in q3_rows]; ro = [r[2] for r in q3_rows]
    log["Q3_8组_沿用方案_均值"] = round(float(np.mean(fx)), 1)
    log["Q3_8组_重优化_均值"] = round(float(np.mean(ro)), 1)
    log["Q3_8组_重优化_提升"] = round(float(np.mean(ro) - np.mean(fx)), 1)

    # 情景收益分布（箱线图数据）
    with open(HERE / "F04_收益分布.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["组别", "净收益"])
        for v in fixed_profits: w.writerow(["沿用Q2方案", round(v, 1)])
        for v in reopt_profits: w.writerow(["按情景重优化", round(v, 1)])

    (HERE / "Q2Q3_结论.json").write_text(json.dumps(log, ensure_ascii=False, indent=1),
                                         encoding="utf-8")
    print(json.dumps(log, ensure_ascii=False, indent=1))
