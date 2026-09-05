# -*- coding: utf-8 -*-
"""
make_flowchart.py — 论文级流程图生成器（技术路线图 / 算法流程图）

一份 JSON 描述「节点 + 连线」，一键渲染出版级流程图（PNG 300dpi + PDF 矢量）：

    python tools/make_flowchart.py --demo route       --outdir tools/output   # 全文技术路线图示例
    python tools/make_flowchart.py --demo algorithm   --outdir tools/output   # 含判断/回环的算法流程图示例
    python tools/make_flowchart.py --spec my_flow.json --outdir 04_图表/figs   # 用自己的 JSON

JSON 格式（模板见 tools/flowchart_spec_example.json）：
{
  "title": "图1 全文技术路线",
  "nodes": [
    {"id": "s", "text": "开始", "type": "start"},          # type: start/process/decision/data/sub/end
    {"id": "p", "text": "处理", "col": 0, "row": 1}         # col/row 可选，用于精调自动布局
  ],
  "edges": [ ["s", "p", ""], ["p", "d", "是"], ["d", "p", "否"] ]   # [起点, 终点, 连线标注]
}

自动布局规则：从 start 起按主干 BFS 分行（主干居中、分支就近占列、回边绕弧）；
给节点显式 col/row 可覆盖。节点类型决定形状与配色（起止胶囊/处理圆角/判断菱形/数据平行四边形/子过程）。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from collections import deque

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Polygon

from fig_style import apply_style, save_fig

DEFAULT_OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

# 节点类型 → 形状/填充/边框/文字色（低饱和学术风，黑白打印亦可辨）
STYLE = {
    "start":   dict(shape="stadium", fill="#2E5A87", edge="#1F3A5C", tc="white"),
    "end":     dict(shape="stadium", fill="#C44E52", edge="#8F3236", tc="white"),
    "process": dict(shape="round",   fill="#EAF1F8", edge="#2E5A87", tc="#1F3A5C"),
    "decision":dict(shape="diamond", fill="#FFF6DF", edge="#C9862B", tc="#6B4A00"),
    "data":    dict(shape="para",    fill="#E8F5EC", edge="#3E7C4F", tc="#24512F"),
    "sub":     dict(shape="sub",     fill="#F3EDF7", edge="#7B5296", tc="#3F2753"),
}
ARROW = "#37474F"
W, H0 = 3.4, 0.92      # 节点基础尺寸
DX, DY = 4.9, 2.35     # 列距 / 行距
WRAP = 13              # 中文自动换行宽度（字符数）


def _wrap(text: str, w: int = WRAP) -> str:
    if "\n" in text:
        return text
    return "\n".join(text[i:i + w] for i in range(0, len(text), w))


def _parse_formats(s: str):
    return tuple(t.strip() for t in s.split(",") if t.strip())


# ---------------------------------------------------------------- 布局
def layout(spec):
    """BFS 定行（主干）、就近占列（分支），显式 col/row 优先。"""
    nodes = spec["nodes"]
    edges = [list(e) for e in spec["edges"]]
    N = {n["id"]: n for n in nodes}
    ids = [n["id"] for n in nodes]
    start = next((i for i in ids if N[i].get("type") == "start"), ids[0])

    succ = {i: [] for i in ids}
    for u, v, *_ in edges:
        if u in succ and v in N:
            succ[u].append(v)

    row, order = {start: 0}, [start]
    q = deque([start])
    while q:
        u = q.popleft()
        for v in succ[u]:
            if v not in row:
                row[v] = row[u] + 1
                order.append(v)
                q.append(v)
    for i in ids:
        if i not in row:
            row[i] = row.get(next(iter(succ[i]), start), 0) if succ[i] else 0
            order.append(i)

    col, occ = {}, {}
    for i in order:
        r = row[i]
        explicit = N[i].get("col")
        if explicit is not None:
            c = int(explicit)
        elif i == start:
            c = 0
        else:
            used = occ.get(r, set())
            c, k = 0, 1
            while c in used:
                if k not in used:
                    c = k
                elif -k not in used:
                    c = -k
                else:
                    k += 1
                    continue
                break
        col[i] = c
        occ.setdefault(r, set()).add(c)

    geom = {}
    for i in ids:
        n = N[i]
        st = STYLE.get(n.get("type", "process"), STYLE["process"])
        text = _wrap(n.get("text", ""))
        lines = text.count("\n") + 1
        if st["shape"] == "diamond":
            w, h = W + 0.6, 1.55
        else:
            w, h = W, H0 + 0.34 * (lines - 1)
        geom[i] = dict(x=col[i] * DX, y=-row[i] * DY, w=w, h=h,
                       text=text, st=st, shape=st["shape"], col=col[i], row=row[i])
    return ids, geom, edges


def _anchor(g, side):
    if side == "bottom":
        return (g["x"], g["y"] - g["h"] / 2)
    if side == "top":
        return (g["x"], g["y"] + g["h"] / 2)
    if side == "right":
        return (g["x"] + g["w"] / 2, g["y"])
    return (g["x"] - g["w"] / 2, g["y"])


# ---------------------------------------------------------------- 绘制
def draw(spec, out, outdir, formats):
    apply_style()  # 中文与统一风格（流程图与图表全文同源）
    ids, G, edges = layout(spec)

    # 画布范围与动态比例：布局越"高"，等比缩放越小，需按最宽文本行反推
    # 最小 figsize，保证节点文字不溢出边框（CJK 全角字符宽 ≈ 1 em = fontsize/72 in）
    xs = [G[i]["x"] for i in ids]
    ys = [G[i]["y"] for i in ids]
    back_any = any(G[v]["row"] <= G[u]["row"] for u, v, *_ in edges if u in G and v in G)
    x0, x1 = min(xs) - W / 2 - 1.1, max(xs) + W / 2 + (2.6 if back_any else 1.1)
    y0, y1 = min(ys) - 1.2, max(ys) + 1.9
    max_line = max((len(ln) for g in G.values() for ln in g["text"].split("\n")), default=8)
    scale = max(0.55, max_line * 9.8 / 72.0 / (W * 0.86))
    fig, ax = plt.subplots(figsize=(min(13.0, (x1 - x0) * scale),
                                    min(15.0, (y1 - y0) * scale)))
    for i in ids:
        g = G[i]
        st, shape = g["st"], g["shape"]
        x, y, w, h = g["x"], g["y"], g["w"], g["h"]
        if shape == "diamond":
            ax.add_patch(Polygon([(x, y + h / 2), (x + w / 2, y), (x, y - h / 2), (x - w / 2, y)],
                                 closed=True, fc=st["fill"], ec=st["edge"], lw=1.3, zorder=3))
            fs = 9.2
        elif shape == "stadium":
            ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h,
                                        boxstyle=f"round,pad=0,rounding_size={h / 2:.3f}",
                                        fc=st["fill"], ec=st["edge"], lw=1.2, zorder=3))
            fs = 9.8
        elif shape == "para":
            k = 0.38
            ax.add_patch(Polygon([(x - w / 2 + k, y + h / 2), (x + w / 2, y + h / 2),
                                  (x + w / 2 - k, y - h / 2), (x - w / 2, y - h / 2)],
                                 closed=True, fc=st["fill"], ec=st["edge"], lw=1.2, zorder=3))
            fs = 9.8
        elif shape == "sub":
            ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h,
                                        boxstyle="round,pad=0,rounding_size=0.05",
                                        fc=st["fill"], ec=st["edge"], lw=1.2, zorder=3))
            for dx in (-w / 2 + 0.16, w / 2 - 0.16):
                ax.plot([x + dx, x + dx], [y - h / 2, y + h / 2],
                        color=st["edge"], lw=1.2, zorder=4)
            fs = 9.8
        else:
            ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h,
                                        boxstyle="round,pad=0,rounding_size=0.14",
                                        fc=st["fill"], ec=st["edge"], lw=1.2, zorder=3))
            fs = 9.8
        if g["text"]:
            ax.text(x, y, g["text"], ha="center", va="center", fontsize=fs,
                    color=st["tc"], linespacing=1.3, zorder=4)

    for u, v, *rest in edges:
        if u not in G or v not in G:
            continue
        label = rest[0] if rest else ""
        gu, gv = G[u], G[v]
        ru, rv, cu, cv = gu["row"], gv["row"], gu["col"], gv["col"]
        if ru == rv:                                   # 同层横向
            go_right = cv > cu
            posA = _anchor(gu, "right" if go_right else "left")
            posB = _anchor(gv, "left" if go_right else "right")
            cs = "arc3,rad=0"
        elif rv <= ru:                                 # 回边：从外侧绕弧
            side = 1 if cu >= 0 else -1
            posA = _anchor(gu, "right" if side > 0 else "left")
            posB = _anchor(gv, "right" if side > 0 else "left")
            cs = f"arc3,rad={0.5 * side:.2f}"
        else:                                          # 下行直连/微弯
            posA = _anchor(gu, "bottom")
            posB = _anchor(gv, "top")
            rad = 0.0 if cu == cv else (0.12 if cv > cu else -0.12)
            cs = f"arc3,rad={rad:.2f}"
        ax.add_patch(FancyArrowPatch(posA, posB, arrowstyle="-|>", mutation_scale=15,
                                     lw=1.3, color=ARROW, connectionstyle=cs,
                                     shrinkA=2, shrinkB=2, zorder=2))
        if label:
            mx, my = (posA[0] + posB[0]) / 2, (posA[1] + posB[1]) / 2
            if abs(posA[0] - posB[0]) < 1e-6:
                lx, ly, ha = mx + 0.22, my, "left"
            elif abs(posA[1] - posB[1]) < 1e-6:
                lx, ly, ha = mx, my + 0.22, "center"
            else:
                lx, ly, ha = mx + 0.18, my + 0.14, "left"
            ax.text(lx, ly, label, fontsize=9, color="#B03A2E", ha=ha, va="center",
                    bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.85),
                    zorder=5)

    ax.set_xlim(x0, x1)
    ax.set_ylim(y0, y1)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.set_title(spec.get("title", ""), fontsize=13, fontweight="bold", pad=8)
    save_fig(fig, out, outdir, formats)


# ---------------------------------------------------------------- 内置示例
def demo_route():
    return {
        "title": "全文技术路线图",
        "nodes": [
            {"id": "s",  "text": "问题重述与数据预处理", "type": "start"},
            {"id": "p1", "text": "问题一：运输方案优化\n（整数规划模型）", "type": "process"},
            {"id": "p2", "text": "问题二：需求预测\n（灰色预测＋LSTM）", "type": "process"},
            {"id": "p3", "text": "问题三：方案综合评价\n（熵权－TOPSIS）", "type": "process"},
            {"id": "s1", "text": "求解与结果分析", "type": "process"},
            {"id": "s2", "text": "求解与结果分析", "type": "process"},
            {"id": "s3", "text": "求解与结果分析", "type": "process"},
            {"id": "e",  "text": "灵敏度分析与模型评价", "type": "end"},
        ],
        "edges": [["s", "p1", ""], ["s", "p2", ""], ["s", "p3", ""],
                  ["p1", "s1", ""], ["p2", "s2", ""], ["p3", "s3", ""],
                  ["s1", "e", ""], ["s2", "e", ""], ["s3", "e", ""]],
    }


def demo_algorithm():
    return {
        "title": "遗传算法求解流程",
        "nodes": [
            {"id": "a", "text": "开始：编码与参数初始化", "type": "start"},
            {"id": "b", "text": "生成初始种群", "type": "process"},
            {"id": "c", "text": "计算个体适应度", "type": "process"},
            {"id": "d", "text": "达到最大迭代次数？", "type": "decision"},
            {"id": "e", "text": "输出最优解与收敛曲线", "type": "end"},
            {"id": "f", "text": "选择·交叉·变异\n生成新一代种群", "type": "process"},
        ],
        "edges": [["a", "b", ""], ["b", "c", ""], ["c", "d", ""],
                  ["d", "e", "是"], ["d", "f", "否"], ["f", "c", ""]],
    }


def main():
    ap = argparse.ArgumentParser(description="论文级流程图生成器（JSON 进，出版级图出）",
                                 formatter_class=argparse.RawDescriptionHelpFormatter,
                                 epilog=__doc__)
    ap.add_argument("--spec", help="流程图 JSON 文件")
    ap.add_argument("--demo", choices=["route", "algorithm"], help="内置示例")
    ap.add_argument("--out", default=None, help="输出文件名主干（不含扩展名）")
    ap.add_argument("--outdir", default=DEFAULT_OUT, help="输出目录（默认 tools/output）")
    ap.add_argument("--formats", default="png,pdf", type=_parse_formats, help="导出格式")
    a = ap.parse_args()
    if not a.spec and not a.demo:
        ap.error("需要 --spec <json> 或 --demo route|algorithm")
    if a.spec:
        with open(a.spec, encoding="utf-8") as f:
            spec = json.load(f)
        stem = a.out or os.path.splitext(os.path.basename(a.spec))[0]
    else:
        spec = demo_route() if a.demo == "route" else demo_algorithm()
        stem = a.out or f"flowchart_{a.demo}"
    draw(spec, stem, a.outdir, a.formats)


if __name__ == "__main__":
    main()
