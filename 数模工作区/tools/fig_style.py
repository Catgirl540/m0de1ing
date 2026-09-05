# -*- coding: utf-8 -*-
"""
fig_style.py — 数模论文统一图表风格库 v2（全套工具的风格基石）

v2 重点（针对"越位爆格、字号不合适、不够高级"的根治）：
  1. 版面：默认 constrained_layout，元素永不越出画布；
  2. 字号物理正确：基准字号 11pt、图宽 6.4in≈16.3cm——按论文 15cm 插图宽度换算，
     正文字号缩水后仍 ≈10pt，不再出现"贴进 Word 字小成一团"；
  3. 数字轴自动千分位、禁用 1e7 科学计数偏移（旧版偏移文字会压住轴标签）；
  4. 四套低饱和科研配色（Okabe-Ito 色盲安全系），默认 academic 十色；
  5. save_fig 爆格自检：捕获字体缺字警告并显式提示；tight bbox + 0.12in 内边距。

用法：
    from fig_style import apply_style, save_fig, grid_on, headroom, kfmt
    apply_style()                                # 一行生效（palette=academic）
    fig, ax = plt.subplots(constrained_layout=True)
    ...
    save_fig(fig, "fig_demo", "output")          # PNG(300dpi)+PDF 双格式
"""
from __future__ import annotations

import os
import warnings

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# 中西文字体栈：按可用性依次回退（Windows 优先微软雅黑）
FONT_STACK = [
    "Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "Source Han Sans SC",
    "PingFang SC", "Arial Unicode MS", "DejaVu Sans",
]

# 四套低饱和/色盲安全配色（学术审美：克制、重点突出）
PALETTES = {
    # 默认：Okabe-Ito 色盲安全系扩展（10 色，供分组/堆叠/多系列）
    "academic": ["#0072B2", "#D55E00", "#009E73", "#CC79A7", "#E69F00",
                 "#56B4E9", "#8C6D31", "#999999", "#666666", "#B2182B"],
    # 经典：v1 兼容（沉稳六色）
    "classic": ["#2E5A87", "#C44E52", "#55A868", "#8172B2", "#CCB974", "#64B5CD"],
    # 低饱和 Nature 风（柔和、适合大面积填充）
    "muted": ["#7C9CBF", "#C98A86", "#8FBCA0", "#A88FBE", "#D4B26A",
              "#86B8CF", "#B08968", "#9AA5B1"],
    # 高对比（答辩 PPT / 需要强区分时）
    "vivid": ["#1F77B4", "#D62728", "#2CA02C", "#FF7F0E", "#9467BD",
              "#8C564B", "#E377C2", "#7F7F7F", "#BCBD22", "#17BECF"],
}


def apply_style(palette: str = "academic", font_scale: float = 1.0) -> list:
    """应用论文统一风格，返回当前配色列表。"""
    colors = PALETTES.get(palette, PALETTES["academic"])
    plt.rcParams.update({
        # 字体（11pt 基准：按 15cm 插图宽度换算后仍 ≈10pt）
        "font.family": "sans-serif",
        "font.sans-serif": FONT_STACK,
        "axes.unicode_minus": False,
        "font.size": 11.0 * font_scale,
        # 坐标轴
        "axes.titlesize": 12.0 * font_scale,
        "axes.titleweight": "bold",
        "axes.titlepad": 8,
        "axes.labelsize": 11.0 * font_scale,
        "axes.linewidth": 0.8,
        "axes.edgecolor": "#3A3A3A",
        "axes.prop_cycle": plt.cycler(color=colors),
        "axes.spines.top": False,
        "axes.spines.right": False,
        "axes.formatter.useoffset": False,      # 禁用 1eN 偏移（压轴标签元凶）
        "axes.grid": False,
        # 刻度
        "xtick.direction": "out",
        "ytick.direction": "out",
        "xtick.labelsize": 10.0 * font_scale,
        "ytick.labelsize": 10.0 * font_scale,
        "xtick.major.size": 3.5,
        "ytick.major.size": 3.5,
        # 图例与线条
        "legend.fontsize": 10.0 * font_scale,
        "legend.frameon": False,
        "legend.handlelength": 1.6,
        "lines.linewidth": 2.0,
        "lines.markersize": 5.5,
        # 填充
        "patch.linewidth": 0.0,
        "hatch.linewidth": 0.6,
        # 网格（默认关，按需 grid_on 打开）
        "grid.linestyle": "--",
        "grid.linewidth": 0.5,
        "grid.alpha": 0.35,
        # 版面与导出
        "figure.constrained_layout.use": True,   # 元素永不越出画布
        "figure.dpi": 110,
        "figure.facecolor": "white",
        "savefig.dpi": 300,
        "savefig.bbox": "tight",
        "savefig.pad_inches": 0.12,
        "pdf.fonttype": 42,                      # PDF 字体可编辑、防缺字
    })
    return colors


def get_colors(n=None, palette: str = "academic") -> list:
    """取 n 个主题色（不足则循环取用）。"""
    colors = PALETTES.get(palette, PALETTES["academic"])
    if n is None:
        return list(colors)
    return [colors[i % len(colors)] for i in range(n)]


def grid_on(ax, axis: str = "y") -> None:
    """打开网格并置底（不被数据遮挡）。"""
    ax.grid(True, axis=axis)
    ax.set_axisbelow(True)


def headroom(ax, factor: float = 1.18, axis: str = "y") -> None:
    """为数值标签留出净空——根治'柱顶数字压图例/越位'。"""
    if axis == "y":
        top = ax.get_ylim()[1]
        ax.set_ylim(top=top * factor if top > 0 else top)
    else:
        right = ax.get_xlim()[1]
        ax.set_xlim(right=right * factor if right > 0 else right)


def kfmt(ax, axis: str = "y", mode: str = "auto") -> None:
    """数字刻度格式：plain=千分位；wan=以万为单位；auto=量级≥1e4 时自动万。"""
    ax_axis = ax.yaxis if axis == "y" else ax.xaxis
    if mode == "plain":
        ax_axis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
        return
    series = ax.get_ylim() if axis == "y" else ax.get_xlim()
    big = max(abs(min(series)), abs(max(series)))
    if mode == "wan" or (mode == "auto" and big >= 1e4):
        ax_axis.set_major_formatter(FuncFormatter(lambda v, _: f"{v / 1e4:g}"))
        unit = "（万元）" if axis == "y" else ""
        lbl = ax.get_ylabel() if axis == "y" else ax.get_xlabel()
        if unit and unit not in lbl:
            ax.set_ylabel(lbl + unit) if axis == "y" else ax.set_xlabel(lbl + unit)
    else:
        ax_axis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))


def label_axes(axs, start: str = "a") -> None:
    """多子图编号：(a)(b)(c)… 加粗置于各子图左上角。"""
    import string

    import numpy as _np
    flat = _np.atleast_1d(_np.asarray(axs)).ravel()
    letters = list(string.ascii_lowercase)
    for ax, tag in zip(flat, letters[letters.index(start):]):
        ax.text(-0.08, 1.06, f"({tag})", transform=ax.transAxes,
                fontsize=12, fontweight="bold", va="bottom", ha="left")


def save_fig(fig, out_stem: str, out_dir: str = "output",
             formats=("png", "pdf"), dpi: int = 300) -> list:
    """统一导出：PNG(300dpi)+PDF 双格式；爆格自检（捕获字体缺字警告）。"""
    os.makedirs(out_dir, exist_ok=True)
    paths = []
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        for fmt in formats:
            path = os.path.join(out_dir, f"{out_stem}.{fmt}")
            fig.savefig(path, format=fmt, dpi=dpi)
            paths.append(path)
    glyph = [str(w.message) for w in caught if "Glyph" in str(w.message) or "missing from" in str(w.message)]
    if glyph:
        print("⚠️ [fig_style] 疑似字体缺字（爆格风险）：", glyph[:2],
              "→ 检查特殊字符或换字体栈")
    plt.close(fig)
    print("[fig_style] 已保存: " + "  ".join(paths))
    return paths
