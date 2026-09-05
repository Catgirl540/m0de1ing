# -*- coding: utf-8 -*-
"""
make_chart.py — 数模论文一键图表工具 v2（18 种图型，统一风格：fig_style.py）

设计原则（借鉴 GitHub 高星图表 skill：modelviz-skill / sci-box / SciencePlots）：
  · 从表达目标选图：趋势→line/area，对比→bar/waterfall/pareto，分布→hist/box/violin，
    相关→scatter/heat/bubble，机理→stream/contour，综合→panel/radar；
  · 低饱和科研版式、色盲安全配色、成图自检（缺字警告）；
  · 每个命令都接受 --data 真实数据，缺省用内置演示数据。

用法（参数可全部放在子命令之后，顺序无关）：
    python tools/make_chart.py --outdir OUT bar --data d.csv --title T --out F01
    python tools/make_chart.py gallery                    # 一次生成全部演示图库

各图型 CSV 格式（UTF-8，首行表头）：
    line        首列=X（类目或数值），其余列=系列
    bar         同 line（分组柱，自动净空+数值标签）
    stackarea   同 line；--pct 归一化为 100% 堆叠
    dualaxis    三列：X, 左轴系列, 右轴系列
    errbar      三列：类目, 均值, 误差
    waterfall   两列：类目, 数值（"合计/总"行渲染为总柱）
    pareto      两列：类目, 数值（自动降序 + 累积%线）
    hist        一列数值（多列则同图叠加）
    box         每列=一组
    violin      每列=一组
    scatter     两列：X, Y（线性拟合 + 95% 置信带 + r 值）
    bubble      三列：X, Y, 规模（气泡大小=第三维）
    heat        矩阵数值表（相关系数等；发散配色 + 数值注释）
    radar       首列=指标名，其余列=系列；--norm 逐轴归一化
    contour     三列长表：X, Y, Z（三角网格等高线填充）
    sensitivity 首列=扰动(%)，其余列=参数（中心基准线）
    panel       长表：面板, X, Y[, 系列]（自动分面 + (a)(b)(c) 编号）
    stream      无需数据：Lotka-Volterra 相场演示（机理题向量场模板）
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import numpy as np
import matplotlib.pyplot as plt

from fig_style import (PALETTES, apply_style, get_colors, grid_on, headroom,
                       kfmt, label_axes, save_fig)

DEFAULT_OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")


def _parse_formats(s: str):
    return tuple(t.strip() for t in s.split(",") if t.strip())


def load_table(path: str):
    """读 CSV/XLSX → (列名列表, DataFrame)。"""
    import pandas as pd
    if str(path).lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path, encoding="utf-8-sig")
    return [str(c) for c in df.columns], df


def num(df, col):
    import pandas as pd
    return pd.to_numeric(df[col], errors="coerce").fillna(0).to_numpy(dtype=float)


def short_num(v):
    v = float(v)
    if abs(v) >= 1e8:
        return f"{v / 1e8:.1f}亿"
    if abs(v) >= 1e4:
        return f"{v / 1e4:.0f}万"
    return f"{v:,.0f}"


def demo(name):
    """内置演示数据（缺省 --data 时使用）。"""
    import pandas as pd
    rng = np.random.default_rng(42)
    if name == "line":
        x = np.arange(2019, 2025)
        return pd.DataFrame({"年份": x, "方案甲": 500 + 40 * np.arange(6) + rng.normal(0, 8, 6),
                             "方案乙": 480 + 62 * np.arange(6) + rng.normal(0, 8, 6)})
    if name == "bar":
        return pd.DataFrame({"方案": ["方案一", "方案二", "方案三", "方案四"],
                             "模型Ⅰ（本文）": [86, 78, 92, 81], "模型Ⅱ（对比）": [79, 85, 88, 90]})
    if name == "stackarea":
        x = np.arange(2019, 2025)
        return pd.DataFrame({"年份": x, "粮食": 60 + 5 * np.arange(6),
                             "蔬菜": 40 + 8 * np.arange(6), "食用菌": 25 + 3 * np.arange(6)})
    if name == "dualaxis":
        x = np.arange(2019, 2025)
        return pd.DataFrame({"年份": x, "产量万吨": 120 + 15 * np.arange(6),
                             "均价元每吨": 3200 + 180 * np.arange(6)})
    if name == "errbar":
        return pd.DataFrame({"组别": ["对照组", "低剂量", "中剂量", "高剂量"],
                             "均值": [42.0, 51.3, 60.8, 66.2], "误差": [3.1, 2.7, 3.9, 2.4]})
    if name == "waterfall":
        return pd.DataFrame({"项目": ["基线收益", "需求增长", "成本上升", "结构优化", "合计"],
                             "数值": [7800, 1240, -860, 426, 8606]})
    if name == "pareto":
        return pd.DataFrame({"缺陷类型": ["尺寸", "外观", "装配", "材料", "其他"],
                             "数量": [62, 34, 21, 12, 6]})
    if name == "hist":
        return pd.DataFrame({"观测值": rng.normal(50, 12, 300)})
    if name == "box":
        return pd.DataFrame({"基线": rng.normal(50, 10, 40), "方案A": rng.normal(58, 8, 40),
                             "方案B": rng.normal(63, 6, 40)})
    if name == "violin":
        return pd.DataFrame({"基线": rng.normal(50, 10, 60), "方案A": rng.normal(58, 8, 60),
                             "方案B": rng.normal(63, 6, 60)})
    if name == "scatter":
        x = rng.uniform(10, 90, 40)
        return pd.DataFrame({"投入": x, "产出": 20 + 0.8 * x + rng.normal(0, 8, 40)})
    if name == "bubble":
        x = rng.uniform(10, 90, 12)
        return pd.DataFrame({"投入": x, "产出": 15 + 0.9 * x + rng.normal(0, 6, 12),
                             "规模": rng.uniform(80, 600, 12)})
    if name == "heat":
        m = np.array([[1.00, 0.82, -0.63, 0.41], [0.82, 1.00, -0.55, 0.36],
                      [-0.63, -0.55, 1.00, -0.28], [0.41, 0.36, -0.28, 1.00]])
        return pd.DataFrame(m, columns=["产量", "面积", "成本", "价格"])
    if name == "radar":
        return pd.DataFrame({"维度": ["精度", "速度", "稳健", "成本", "可解释"],
                             "本文模型": [0.9, 0.7, 0.8, 0.6, 0.85],
                             "对比模型": [0.8, 0.85, 0.6, 0.7, 0.6]})
    if name == "contour":
        g = np.linspace(0, 10, 30)
        xs, ys = np.meshgrid(g, g)
        z = np.sin(xs / 2) * np.cos(ys / 3) + 0.02 * (xs * ys) / 10
        return pd.DataFrame({"X": xs.ravel(), "Y": ys.ravel(), "Z": z.ravel()})
    if name == "sensitivity":
        d = np.array([-20, -10, 0, 10, 20])
        return pd.DataFrame({"扰动%": d, "销售价格": 86 + 0.9 * d, "种植成本": 86 - 0.6 * d,
                             "预期销量": 86 + 0.45 * d})
    if name == "panel":
        rows = []
        for name_s, base in [("情景一", 50), ("情景二", 58)]:
            xs = np.arange(2024, 2031)
            for x in xs:
                rows.append({"面板": name_s, "年份": x,
                             "净收益": base + 3 * (x - 2023) + rng.normal(0, 4)})
                rows.append({"面板": name_s, "年份": x,
                             "净收益下界": base + 2 * (x - 2023) + rng.normal(0, 4)})
        return pd.DataFrame(rows)
    raise KeyError(name)


# ---------------- 各图型实现 ----------------

def c_line(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("line")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    xs = df.iloc[:, 0]
    markers = ["o", "s", "^", "D", "v", "P", "X", "*"]
    for i, col in enumerate(df.columns[1:]):
        ax.plot(np.arange(len(xs)), num(df, col), marker=markers[i % len(markers)],
                markevery=max(1, len(xs) // 10), label=col)
    ax.set_xticks(np.arange(len(xs)), [str(v) for v in xs])
    ax.set_xlabel(args.xlabel or (cols[0] if cols else df.columns[0]))
    ax.set_ylabel(args.ylabel or "")
    if args.kfmt != "off":
        kfmt(ax, "y", args.kfmt)
    ax.set_title(args.title or "多系列趋势对比")
    ax.legend(loc="best", ncols=min(len(df.columns) - 1, 3))
    grid_on(ax)
    save_fig(fig, args.out or "fig_line", args.outdir, args.formats)


def c_bar(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("bar")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    cats = [str(v) for v in df.iloc[:, 0]]
    series = [num(df, c) for c in df.columns[1:]]
    labels = list(df.columns[1:])
    n = max(len(series), 1)
    colors = get_colors(n, args.palette)
    x = np.arange(len(cats))
    w = min(0.8 / n, 0.5)
    for i, s in enumerate(series):
        b = ax.bar(x - 0.4 + w * (i + 0.5), s, w, label=labels[i], color=colors[i % len(colors)])
        ax.bar_label(b, fmt=short_num, padding=2, fontsize=8.5)
    ax.set_xticks(x, cats, rotation=max(0, 15 * (len(cats) > 8)),
                  ha="right" if len(cats) > 8 else "center")
    ax.set_ylabel(args.ylabel or "数值")
    headroom(ax, 1.15)
    kfmt(ax, "y", args.kfmt)
    ax.set_title(args.title or "分组对比")
    ax.legend(loc="upper left", ncols=min(n, 3))
    grid_on(ax)
    save_fig(fig, args.out or "fig_bar", args.outdir, args.formats)


def c_stackarea(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("stackarea")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    xs = df.iloc[:, 0]
    series = [num(df, c) for c in df.columns[1:]]
    labels = list(df.columns[1:])
    A = np.vstack(series)
    if args.pct:
        A = A / A.sum(axis=0, keepdims=True) * 100
    colors = get_colors(len(series), args.palette)
    ax.stackplot(np.arange(len(xs)), A, labels=labels, colors=colors, alpha=0.88)
    ax.set_xticks(np.arange(len(xs)), [str(v) for v in xs])
    ax.set_ylabel("占比（%）" if args.pct else (args.ylabel or "数值"))
    ax.set_ylim(0, 100 if args.pct else None)
    ax.set_title(args.title or "构成随时间变化")
    ax.legend(loc="upper left", ncols=min(len(labels), 4), reverse=True)
    ax.set_xlim(0, len(xs) - 1)
    grid_on(ax)
    save_fig(fig, args.out or "fig_stackarea", args.outdir, args.formats)


def c_dualaxis(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("dualaxis")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    xs = df.iloc[:, 0]
    l, r = num(df, df.columns[1]), num(df, df.columns[2])
    ax.plot(np.arange(len(xs)), l, "o-", color=get_colors(1, args.palette)[0], label=df.columns[1])
    ax2 = ax.twinx()
    ax2.spines["right"].set_visible(True)
    ax2.plot(np.arange(len(xs)), r, "s--", color=get_colors(2, args.palette)[1], label=df.columns[2])
    ax2.set_ylabel(df.columns[2])
    ax.set_xticks(np.arange(len(xs)), [str(v) for v in xs])
    ax.set_ylabel(df.columns[1])
    ax.set_title(args.title or "双轴对照")
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper left")
    grid_on(ax)
    save_fig(fig, args.out or "fig_dualaxis", args.outdir, args.formats)


def c_errbar(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("errbar")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    cats = [str(v) for v in df.iloc[:, 0]]
    mean, err = num(df, df.columns[1]), num(df, df.columns[2])
    x = np.arange(len(cats))
    ax.bar(x, mean, 0.55, yerr=err, capsize=5, color=get_colors(1, args.palette)[0],
           error_kw={"elinewidth": 1.4})
    for xi, m, e in zip(x, mean, err):
        ax.text(xi, m + e, short_num(m), ha="center", va="bottom", fontsize=9)
    ax.set_xticks(x, cats)
    ax.set_ylabel(args.ylabel or "均值 ± 误差")
    headroom(ax, 1.22)
    ax.set_title(args.title or "均值与误差对比")
    grid_on(ax)
    save_fig(fig, args.out or "fig_errbar", args.outdir, args.formats)


def c_waterfall(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("waterfall")))
    fig, ax = plt.subplots(figsize=(6.6, 3.9))
    cats = [str(v) for v in df.iloc[:, 0]]
    vals = num(df, df.columns[1])
    colors = get_colors(3, args.palette)
    cum = 0.0
    for i, (cat, v) in enumerate(zip(cats, vals)):
        is_total = ("总" in cat) or ("合计" in cat)
        if is_total:
            ax.bar(i, v, 0.6, color=colors[0])
            ax.text(i, v, short_num(v), ha="center", va="bottom", fontsize=9)
            cum = v
            continue
        bottom = cum if v >= 0 else cum + v
        ax.bar(i, abs(v), 0.6, bottom=bottom, color=colors[2] if v >= 0 else colors[1])
        ax.text(i, cum + max(v, 0), short_num(v), ha="center", va="bottom", fontsize=9)
        cum += v
        ax.plot([i + 0.3, i + 0.7], [cum, cum], color="#888888", lw=0.8, ls="--")
    ax.set_xticks(np.arange(len(cats)), cats, rotation=15, ha="right")
    ax.set_ylabel(args.ylabel or "数值")
    headroom(ax, 1.15)
    kfmt(ax, "y", args.kfmt)
    ax.set_title(args.title or "贡献分解（瀑布图）")
    grid_on(ax)
    save_fig(fig, args.out or "fig_waterfall", args.outdir, args.formats)


def c_pareto(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("pareto")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    d = df.sort_values(df.columns[1], ascending=False).reset_index(drop=True)
    cats = [str(v) for v in d.iloc[:, 0]]
    vals = num(d, d.columns[1])
    x = np.arange(len(cats))
    ax.bar(x, vals, 0.55, color=get_colors(1, args.palette)[0])
    for xi, v in zip(x, vals):
        ax.text(xi, v, short_num(v), ha="center", va="bottom", fontsize=9)
    ax.set_xticks(x, cats)
    ax.set_ylabel(args.ylabel or "频数/数值")
    headroom(ax, 1.12)
    cum = np.cumsum(vals) / vals.sum() * 100
    ax2 = ax.twinx()
    ax2.spines["right"].set_visible(True)
    ax2.plot(x, cum, "o-", color=get_colors(2, args.palette)[1], lw=1.8)
    ax2.set_ylim(0, 110)
    ax2.set_ylabel("累积占比（%）")
    ax2.axhline(80, color="#999999", lw=0.8, ls=":")
    ax2.text(len(cats) - 0.55, 81.5, "80%", fontsize=9, color="#666666")
    ax.set_title(args.title or "帕累托分析")
    grid_on(ax)
    save_fig(fig, args.out or "fig_pareto", args.outdir, args.formats)


def c_hist(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("hist")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    colors = get_colors(len(df.columns), args.palette)
    try:
        from scipy.stats import gaussian_kde
        has_kde = True
    except ImportError:
        has_kde = False
    for i, col in enumerate(df.columns):
        v = num(df, col)
        ax.hist(v, bins=18, density=True, alpha=0.55, color=colors[i % len(colors)], label=col)
        if has_kde:
            g = np.linspace(v.min(), v.max(), 200)
            ax.plot(g, gaussian_kde(v)(g), color=colors[i % len(colors)], lw=2)
    ax.set_xlabel(args.xlabel or "取值")
    ax.set_ylabel("概率密度")
    ax.set_title(args.title or "分布形态")
    ax.legend()
    grid_on(ax)
    save_fig(fig, args.out or "fig_hist", args.outdir, args.formats)


def c_box(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("box")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    data = [num(df, c) for c in df.columns]
    bp = ax.boxplot(data, patch_artist=True, widths=0.5, tick_labels=list(df.columns),
                    showmeans=True,
                    meanprops={"marker": "D", "markerfacecolor": "white",
                               "markeredgecolor": "#333333", "markersize": 5},
                    medianprops={"color": "#333333", "lw": 1.6},
                    flierprops={"marker": "o", "markersize": 3.5, "alpha": 0.6})
    for patch, c in zip(bp["boxes"], get_colors(len(data), args.palette)):
        patch.set_facecolor(c)
        patch.set_alpha(0.65)
    ax.set_ylabel(args.ylabel or "取值")
    ax.set_title(args.title or "分布对比（◆均值）")
    grid_on(ax)
    save_fig(fig, args.out or "fig_box", args.outdir, args.formats)


def c_violin(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("violin")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    data = [num(df, c) for c in df.columns]
    vp = ax.violinplot(data, showmeans=False, showmedians=False, widths=0.75)
    for body, c in zip(vp["bodies"], get_colors(len(data), args.palette)):
        body.set_facecolor(c)
        body.set_alpha(0.6)
    for i, v in enumerate(data, 1):
        ax.scatter([i], [np.mean(v)], marker="D", s=28, color="#333333", zorder=3)
        ax.vlines(i, np.percentile(v, 25), np.percentile(v, 75), color="#333333", lw=1.4)
    ax.set_xticks(np.arange(1, len(df.columns) + 1), list(df.columns))
    ax.set_ylabel(args.ylabel or "取值")
    ax.set_title(args.title or "分布形态对比（◆均值）")
    grid_on(ax)
    save_fig(fig, args.out or "fig_violin", args.outdir, args.formats)


def c_scatter(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("scatter")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    x, y = num(df, df.columns[0]), num(df, df.columns[1])
    ax.scatter(x, y, s=34, alpha=0.75, color=get_colors(1, args.palette)[0],
               edgecolors="white", lw=0.6)
    k, b = np.polyfit(x, y, 1)
    xs = np.linspace(x.min(), x.max(), 100)
    resid = y - (k * x + b)
    se = resid.std(ddof=2) * np.sqrt(1 / len(x) + (xs - x.mean()) ** 2 / ((x - x.mean()) ** 2).sum())
    ax.plot(xs, k * xs + b, color=get_colors(2, args.palette)[1], lw=2,
            label=f"y = {k:.3f}x + {b:.1f}")
    ax.fill_between(xs, k * xs + b - 1.96 * se, k * xs + b + 1.96 * se,
                    color=get_colors(2, args.palette)[1], alpha=0.15, label="95% 置信带")
    r = np.corrcoef(x, y)[0, 1]
    ax.text(0.02, 0.96, f"r = {r:.3f}", transform=ax.transAxes, fontsize=10, va="top")
    ax.set_xlabel(args.xlabel or (cols[0] if cols else "X"))
    ax.set_ylabel(args.ylabel or (cols[1] if cols else "Y"))
    ax.set_title(args.title or "相关与拟合")
    ax.legend(loc="lower right")
    grid_on(ax)
    save_fig(fig, args.out or "fig_scatter", args.outdir, args.formats)


def c_bubble(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("bubble")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    x, y, s = num(df, df.columns[0]), num(df, df.columns[1]), num(df, df.columns[2])
    sizes = 40 + (s - s.min()) / (np.ptp(s) + 1e-9) * 620
    ax.scatter(x, y, s=sizes, alpha=0.55, color=get_colors(1, args.palette)[0],
               edgecolors="white", lw=0.8)
    ax.set_xlabel(args.xlabel or (cols[0] if cols else "X"))
    ax.set_ylabel(args.ylabel or (cols[1] if cols else "Y"))
    ax.set_title(args.title or "气泡图（气泡大小=第三维）")
    grid_on(ax)
    save_fig(fig, args.out or "fig_bubble", args.outdir, args.formats)


def c_heat(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("heat")))
    cols = cols if cols else [str(c) for c in df.columns]
    fig, ax = plt.subplots(figsize=(6.0, 5.0))
    M = df.apply(lambda c: __import__("pandas").to_numeric(c, errors="coerce")).to_numpy(dtype=float)
    has_neg = bool((M < 0).any())
    im = ax.imshow(M, cmap="RdBu_r" if has_neg else "YlGnBu",
                   vmin=-1 if has_neg else float(np.nanmin(M)),
                   vmax=1 if has_neg else float(np.nanmax(M)))
    fig.colorbar(im, ax=ax, shrink=0.85)
    if M.shape[0] <= 15 and M.shape[1] <= 15:
        for i in range(M.shape[0]):
            for j in range(M.shape[1]):
                ax.text(j, i, f"{M[i, j]:.2f}", ha="center", va="center", fontsize=8.5,
                        color="white" if abs(M[i, j]) > 0.6 * np.nanmax(np.abs(M)) else "#222222")
    ax.set_xticks(range(len(cols)), cols, rotation=30, ha="right")
    ax.set_yticks(range(len(df.index)), [str(v) for v in df.index])
    ax.set_title(args.title or "相关系数矩阵")
    save_fig(fig, args.out or "fig_heat", args.outdir, args.formats)


def c_radar(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("radar")))
    cats = [str(v) for v in df.iloc[:, 0]]
    series = [num(df, c) for c in df.columns[1:]]
    if args.norm:
        series = [s / (np.abs(s).max() + 1e-12) for s in series]
    n = len(cats)
    ang = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist()
    ang += ang[:1]
    fig, ax = plt.subplots(figsize=(5.6, 5.2), subplot_kw={"polar": True})
    colors = get_colors(len(series), args.palette)
    for s, c, lab in zip(series, colors, df.columns[1:]):
        v = s.tolist() + s.tolist()[:1]
        ax.plot(ang, v, "o-", lw=1.8, color=c, label=lab, markersize=4)
        ax.fill(ang, v, color=c, alpha=0.12)
    ax.set_xticks(ang[:-1], cats)
    ax.set_title(args.title or "综合评价雷达", pad=18)
    ax.legend(loc="upper right", bbox_to_anchor=(1.28, 1.10))
    save_fig(fig, args.out or "fig_radar", args.outdir, args.formats)


def c_contour(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("contour")))
    fig, ax = plt.subplots(figsize=(6.4, 4.4))
    x, y, z = num(df, df.columns[0]), num(df, df.columns[1]), num(df, df.columns[2])
    cf = ax.tricontourf(x, y, z, levels=18, cmap="viridis")
    cs = ax.tricontour(x, y, z, levels=10, colors="white", linewidths=0.5)
    ax.clabel(cs, inline=True, fontsize=7.5, fmt="%.1f")
    fig.colorbar(cf, ax=ax, label=args.ylabel or "Z")
    ax.set_xlabel(args.xlabel or (cols[0] if cols else "X"))
    ax.set_ylabel(cols[1] if cols else "Y")
    ax.set_title(args.title or "响应曲面与等高线")
    save_fig(fig, args.out or "fig_contour", args.outdir, args.formats)


def c_sensitivity(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("sensitivity")))
    fig, ax = plt.subplots(figsize=(6.4, 3.9))
    d = num(df, df.columns[0])
    colors = get_colors(len(df.columns) - 1, args.palette)
    for i, col in enumerate(df.columns[1:]):
        ax.plot(d, num(df, col), "o-", lw=1.8, color=colors[i], label=col, markersize=4.5)
    ax.axvline(0, color="#888888", lw=0.9, ls=":")
    ax.text(0.3, ax.get_ylim()[1], "基准", fontsize=9, va="top", color="#666666")
    ax.set_xlabel(args.xlabel or "参数扰动（%）")
    ax.set_ylabel(args.ylabel or "目标值")
    ax.set_title(args.title or "敏感性分析")
    ax.legend(loc="best")
    grid_on(ax)
    save_fig(fig, args.out or "fig_sensitivity", args.outdir, args.formats)


def c_panel(args):
    cols, df = (load_table(args.data) if args.data else (None, demo("panel")))
    panels = list(dict.fromkeys(df.iloc[:, 0].astype(str)))
    n = len(panels)
    ncol = 2 if n > 1 else 1
    nrow = int(np.ceil(n / ncol))
    fig, axs = plt.subplots(nrow, ncol, figsize=(6.4 * ncol, 3.6 * nrow), squeeze=False)
    n_series = max(len(df.columns) - 3, 1)
    colors = get_colors(n_series, args.palette)
    val_col = df.columns[2]
    x_col = df.columns[1]
    series_cols = list(df.columns[3:]) if len(df.columns) > 3 else [None]
    for i, p in enumerate(panels):
        ax = axs[i // ncol][i % ncol]
        sub = df[df.iloc[:, 0].astype(str) == p]
        for j, sc in enumerate(series_cols):
            sub2 = sub if sc is None else sub[sub[sc].notna()]
            grp = sub2.groupby(sub2.columns[1], sort=True)[val_col].mean()
            ax.plot(grp.index.astype(str), grp.values, "o-", lw=1.8, ms=4,
                    color=colors[j % len(colors)], label=sc or val_col)
        ax.set_title(str(p), fontsize=11)
        ax.set_xlabel(args.xlabel or x_col)
        ax.set_ylabel(args.ylabel or val_col)
        if len(series_cols) > 1:
            ax.legend(fontsize=9)
        grid_on(ax)
    for j in range(n, nrow * ncol):
        axs[j // ncol][j % ncol].axis("off")
    label_axes(axs)
    fig.suptitle(args.title or "分面板对比", fontweight="bold")
    save_fig(fig, args.out or "fig_panel", args.outdir, args.formats)


def c_stream(args):
    fig, ax = plt.subplots(figsize=(6.2, 5.2))
    g = np.linspace(0, 4, 24)
    X, Y = np.meshgrid(g, g)
    a, b, c, dd = 1.1, 0.4, 0.9, 0.35
    U = a * X - b * X * Y
    V = -c * Y + dd * X * Y
    speed = np.hypot(U, V)
    ax.streamplot(X, Y, U, V, color=speed, cmap="viridis", density=1.2,
                  linewidth=0.9, arrowsize=1.1)
    for x0 in (1.0, 2.0, 3.2):
        xs, ys = [x0], [0.8]
        for _ in range(400):
            dx = a * xs[-1] - b * xs[-1] * ys[-1]
            dy = -c * ys[-1] + dd * xs[-1] * ys[-1]
            xs.append(xs[-1] + 0.02 * dx)
            ys.append(ys[-1] + 0.02 * dy)
        ax.plot(xs, ys, color=get_colors(2, args.palette)[1], lw=1.6)
    ax.set_xlabel("猎物 x")
    ax.set_ylabel("捕食者 y")
    ax.set_title(args.title or "Lotka-Volterra 相场与轨线")
    save_fig(fig, args.out or "fig_stream", args.outdir, args.formats)


def c_gallery(args):
    made = []
    for name, fn in FUNCS.items():
        if name == "gallery":
            continue
        try:
            fn(argparse.Namespace(**{**vars(args), "data": None, "title": None, "out": None,
                                     "ylabel": None, "xlabel": None, "kfmt": "auto",
                                     "pct": False, "norm": False}))
            made.append(name)
        except Exception as e:
            print(f"[gallery] {name} 失败：{e}")
    print(f"\n[gallery] 演示图库完成（{len(made)}/18，PNG+PDF 双格式）→ {args.outdir}")


FUNCS = {"line": c_line, "bar": c_bar, "stackarea": c_stackarea, "dualaxis": c_dualaxis,
         "errbar": c_errbar, "waterfall": c_waterfall, "pareto": c_pareto, "hist": c_hist,
         "box": c_box, "violin": c_violin, "scatter": c_scatter, "bubble": c_bubble,
         "heat": c_heat, "radar": c_radar, "contour": c_contour, "sensitivity": c_sensitivity,
         "panel": c_panel, "stream": c_stream, "gallery": c_gallery}


def main():
    p = argparse.ArgumentParser(description="数模论文一键图表工具 v2（18 种图型，统一风格：fig_style.py）",
                                formatter_class=argparse.RawDescriptionHelpFormatter,
                                epilog=__doc__)
    p.add_argument("--outdir", default=DEFAULT_OUT, help="输出目录（默认 tools/output）")
    p.add_argument("--palette", default="academic", choices=list(PALETTES), help="主题配色")
    p.add_argument("--formats", default="png,pdf", type=_parse_formats, help="导出格式")
    sub = p.add_subparsers(dest="chart", required=True)
    helps = {"line": "多系列趋势折线", "bar": "分组柱状对比", "stackarea": "堆叠面积（--pct 百分比）",
             "dualaxis": "双轴对照", "errbar": "均值±误差棒", "waterfall": "贡献分解瀑布图",
             "pareto": "帕累托（柱+累积%）", "hist": "直方图+核密度", "box": "箱线分布",
             "violin": "小提琴分布", "scatter": "散点+拟合+置信带", "bubble": "气泡图（三维）",
             "heat": "矩阵/相关热图", "radar": "综合评价雷达", "contour": "等高线响应面",
             "sensitivity": "敏感性分析", "panel": "多面板分面", "stream": "向量场/相图",
             "gallery": "一次生成全部演示图库"}
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--data", default=None, help="CSV/XLSX 数据文件（缺省用演示数据）")
    common.add_argument("--title", default=None, help="图标题")
    common.add_argument("--out", default=None, help="输出文件名主干（不含扩展名）")
    common.add_argument("--ylabel", default=None, help="y 轴标签")
    common.add_argument("--xlabel", default=None, help="x 轴标签")
    common.add_argument("--kfmt", default="auto", choices=["auto", "plain", "wan", "off"],
                        help="数字刻度：auto=量级≥1万自动转万")
    for name, help_ in helps.items():
        sp = sub.add_parser(name, help=help_, parents=[common])
        if name == "stackarea":
            sp.add_argument("--pct", action="store_true", help="归一化为 100% 堆叠")
        if name == "radar":
            sp.add_argument("--norm", action="store_true", help="逐轴归一化到 [0,1]")
        sp.set_defaults(func=FUNCS[name])
    args = p.parse_args()
    apply_style(args.palette)
    args.func(args)


if __name__ == "__main__":
    main()
