# -*- coding: utf-8 -*-
"""load_data.py — 2024C 数据加载与清洗 → data.json
口径（见 01_审题/题目分析.md 假设 A1/A2）：
  · 过滤表尾注释行；销售单价区间取中值；
  · 预期销售量 D_c = 该作物 2023 年总产量（Σ 种植面积 × 对应地块类型亩产量），未种者取 0。
"""
import json
from pathlib import Path
import pandas as pd

HERE = Path(__file__).resolve().parent
YEARS = list(range(2024, 2031))

def num(s):
    return pd.to_numeric(pd.Series([s]), errors="coerce").iloc[0]

def main():
    # ---------- 地块 ----------
    land = pd.read_csv(HERE / "A1_乡村的现有耕地.csv").dropna(subset=["地块类型"])
    plots = []
    for _, r in land.iterrows():
        plots.append({"name": r["地块名称"], "type": str(r["地块类型"]).strip(),
                      "area": float(r["地块面积/亩"])})
    # ---------- 作物 ----------
    crops_raw = pd.read_csv(HERE / "A1_乡村种植的农作物.csv").dropna(subset=["作物编号"])
    crops_raw = crops_raw[pd.to_numeric(crops_raw["作物编号"], errors="coerce").notna()]
    crops = [{"id": int(r["作物编号"]), "name": str(r["作物名称"]).strip(), "type": str(r["作物类型"]).strip()}
             for _, r in crops_raw.iterrows()]
    ctype = {c["id"]: c["type"] for c in crops}
    beans = [c["id"] for c in crops if "豆类" in c["type"]]
    # ---------- 参数表（作物×地块类型×季次）----------
    st = pd.read_csv(HERE / "A2_2023年统计的相关数据.csv")
    st = st[pd.to_numeric(st["作物编号"], errors="coerce").notna()].copy()
    st["作物编号"] = st["作物编号"].astype(int)

    def mid_price(v):
        s = str(v)
        if "-" in s:
            a, b = s.split("-")
            return (float(a) + float(b)) / 2
        return float(s)

    params = {}   # (cid, 地块类型, 季次) -> dict
    for _, r in st.iterrows():
        key = (int(r["作物编号"]), str(r["地块类型"]).strip(), str(r["种植季次"]).strip())
        params[key] = {"yield": float(r["亩产量/斤"]), "cost": float(r["种植成本/(元/亩)"]),
                       "price": mid_price(r["销售单价/(元/斤)"])}
    # ---------- 2023 实种（合并单元格：第二季行地块名为 NaN → 前向填充归属上一地块）----------
    p23 = pd.read_csv(HERE / "A2_2023年的农作物种植情况.csv")
    p23["种植地块"] = p23["种植地块"].ffill()
    hist23, demand = {}, {}
    for _, r in p23.iterrows():
        pid, cid = r["种植地块"], int(r["作物编号"])
        s = str(r["种植季次"]).strip()
        hist23.setdefault(pid, {}).setdefault(s, []).append(cid)
        key = (cid, next(p["type"] for p in plots if p["name"] == pid), s)
        if key in params:
            demand[cid] = demand.get(cid, 0.0) + float(r["种植面积/亩"]) * params[key]["yield"]
    for c in crops:
        demand.setdefault(c["id"], 0.0)

    # ---------- 适配矩阵：slot类型 -> 可种作物 ----------
    grains = [c["id"] for c in crops if c["type"].startswith("粮食")]        # 1-16（含水稻）
    grains_dry = [i for i in grains if i != 16]                              # 旱地粮食 1-15
    vegs = [c["id"] for c in crops if c["type"].startswith("蔬菜")]          # 17-37
    vegs_free = [i for i in vegs if i not in (35, 36, 37)]                   # 可上大棚/水浇地S1
    mushrooms = [c["id"] for c in crops if c["type"] == "食用菌"]            # 38-41
    roots = [35, 36, 37]

    compat = {}
    for t in ("平旱地", "梯田", "山坡地"):
        compat[(t, "单季")] = grains_dry
    compat[("水浇地", "第一季")] = [16] + vegs_free          # 水稻走整年通道，另加约束
    compat[("水浇地", "第二季")] = roots
    compat[("普通大棚", "第一季")] = vegs_free
    compat[("普通大棚", "第二季")] = mushrooms
    for s in ("第一季", "第二季"):
        compat[("智慧大棚", s)] = vegs_free

    data = {"years": YEARS, "plots": plots, "crops": crops, "beans": beans,
            "params": {f"{k[0]}|{k[1]}|{k[2]}": v for k, v in params.items()},
            "compat": {f"{k[0]}|{k[1]}": v for k, v in compat.items()},
            "hist23": hist23, "demand": demand,
            "roots": roots, "rice": 16,
            "min_area": {"露地": 0.5, "大棚": 0.3},
            "max_plots_per_crop_season": 8}
    (HERE / "data.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    print(f"地块 {len(plots)} | 作物 {len(crops)} | 豆类 {len(beans)} | 参数键 {len(params)}")
    print("预期销售量=0 的作物:", [c['name'] for c in crops if demand[c['id']] == 0])
    print("2023 总产量前5:", sorted(demand.items(), key=lambda kv: -kv[1])[:5])

if __name__ == "__main__":
    main()
