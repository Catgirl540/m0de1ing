# -*- coding: utf-8 -*-
"""
绘制问题一图表（论文手图表车间 · 遵循 AGENTS.md 图表铁律）
  F01 遮蔽判据时间历程   —— 数据来自 03_代码/F01_视线距离.csv（单一事实源）
  F02 三维几何态势图     —— 云团/真目标做示意性放大（图题注明），几何同源重构
  F03 遮蔽几何放大图     —— 云团局部视图：导弹→真目标视线束穿云可视化（等比例）
输出：figs/ 下 PNG(300dpi)+PDF 双格式
"""
import csv
import os
import sys

import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "03_代码"))
from fig_style import apply_style, grid_on, save_fig  # noqa: E402
from 烟幕内核 import (P_FY1, R_CLOUD, V_CLOUD, cloud_pos,  # noqa: E402
                   drop_state, heading_unit, missile_pos as missile)

HERE = os.path.dirname(os.path.abspath(__file__))
FIGS = os.path.join(HERE, "figs")
CODE = os.path.join(os.path.dirname(HERE), "03_代码")

# ---- 绘图几何：单一事实源 = 仿真内核（常数不再复制，杜绝文图不符）----
V_UAV = 120.0                       # 问题一策略参数
u1 = heading_unit(180.0)
_st = drop_state(180.0, V_UAV, 1.5, 3.6)
D, B, T_DET = _st["D"], _st["B"], _st["t_det"]


def cloud(t):
    return cloud_pos(t, B, T_DET)


# ================= F01 遮蔽判据时间历程（数据：F01_视线距离.csv） =================
with open(os.path.join(CODE, "F01_视线距离.csv"), encoding="utf-8-sig") as f:
    rows = list(csv.reader(f))[1:]
t = np.array([float(r[0]) for r in rows])
d = np.array([float(r[1]) for r in rows])
b = np.array([int(r[2]) for r in rows]) == 1
w0, w1 = t[b].min(), t[b].max()

with open(os.path.join(CODE, "Q1_结论.csv"), encoding="utf-8-sig") as f:
    concl = {r[0]: r[1] for r in list(csv.reader(f))[1:]}
duration = float(concl["有效遮蔽时长(s)"])

apply_style()
fig, ax = plt.subplots(figsize=(7.4, 4.4))
ax.axvspan(w0, w1, color="#55A868", alpha=0.16,
           label=f"有效遮蔽窗口 [{w0:.2f}, {w1:.2f}] s")
ax.axhline(10, color="#C44E52", ls="--", lw=1.2,
           label="有效遮蔽判据（10 m）")
ax.axvline(T_DET, color="#888888", ls=":", lw=1.2, label="云团起爆 t=5.1 s")
ax.plot(t, d, color="#2E5A87", lw=1.9, label="最坏视线距离（真目标全柱146点取最大）")
imin = int(np.argmin(np.where(b, d, np.inf)))
ax.annotate(f"窗口时长 {duration:.2f} s", xy=(t[imin], d[imin]), xytext=(10.1, 48),
            fontsize=9.5, color="#2F5D3A",
            arrowprops=dict(arrowstyle="->", color="#55A868"))
ax.set_xlim(4.6, 12)
ax.set_ylim(0, 145)
ax.set_xlabel("受领任务后时间 t (s)")
ax.set_ylabel("最坏视线距离 (m)")
ax.set_title("问题一：遮蔽判据时间历程")
ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1.0))
grid_on(ax)
save_fig(fig, "F01_遮蔽判据时间历程", FIGS)

# ================= F02 三维几何态势图：双面板（全景 + 局部） =================
# 说明：假目标/真目标在 x≈0，全场量程必然 ~2 万米，单面板下 FY1 的 180 m 航线与云团细节不可辨，
# 故拆 (a) 全景 / (b) 局部两面板；(b) 中云团按 3× 示意放大。
R_CLOUD_A, R_T_S, H_T_S = 150.0, 120.0, 180.0    # (a) 云团示意球半径 / 真目标示意尺寸
R_CLOUD_B = 30.0                                  # (b) 云团示意半径（3×）


def sphere(ax, c, r, color, alpha):
    u, v = np.mgrid[0:2 * np.pi:20j, 0:np.pi:10j]
    ax.plot_surface(c[0] + r * np.cos(u) * np.sin(v), c[1] + r * np.sin(u) * np.sin(v),
                    c[2] + r * np.cos(v), color=color, alpha=alpha,
                    linewidth=0, shade=False)


apply_style()
fig = plt.figure(figsize=(12.8, 5.8))

# ---- (a) 全景 ----
WB = dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.75)  # 标注白底衬
ax = fig.add_subplot(121, projection="3d")
tm = np.linspace(0.0, 66.0, 80)
Mm = np.array([missile(x) for x in tm])
ax.plot(Mm[:, 0], Mm[:, 1], Mm[:, 2], ls="--", color="#C44E52", lw=1.6)
ax.text(7000, -500, 1930, "M1 弹道（300 m/s，直指假目标）", fontsize=9.5,
        color="#8F3236", bbox=WB)
tf = np.linspace(0.0, T_DET, 30)
Ff = np.array([P_FY1 + V_UAV * x * u1 for x in tf])
ax.plot(Ff[:, 0], Ff[:, 1], Ff[:, 2], color="#2E5A87", lw=3.0)
ax.scatter([P_FY1[0]], [P_FY1[1]], [P_FY1[2]], marker="^", color="#2E5A87", s=70)
ax.text(P_FY1[0] + 150, P_FY1[1] + 1300, P_FY1[2] - 550, "FY1 起点", fontsize=9.5,
        color="#1F3A5C", bbox=WB)
ax.plot([P_FY1[0] + 90, P_FY1[0] + 15], [900, 50], [P_FY1[2] - 350, P_FY1[2] - 40],
        color="#2E5A87", lw=0.9, alpha=0.75)          # 引线：FY1 起点 → 三角
ax.scatter(*D, color="#2E5A87", s=40)
ax.text(19300, -1300, 1020, "投放点", fontsize=9.5, color="#1F3A5C", bbox=WB)
ax.plot([19220, D[0] + 70], [-1200, -60], [960, D[2] - 50],
        color="#2E5A87", lw=0.9, alpha=0.75)          # 引线：投放点 → 圆点
ax.scatter(*B, marker="*", color="#C44E52", s=130)
ax.text(19400, -1100, 680, "起爆点", fontsize=9.5, color="#8F3236", bbox=WB)
ax.plot([19320, B[0] + 80], [-1000, -80], [740, B[2] - 70],
        color="#8F3236", lw=0.9, alpha=0.75)          # 引线：起爆点 → 星标
sphere(ax, cloud(T_DET), R_CLOUD_A, "#4C72B0", 0.22)
ax.text(16700, -500, 1230, "烟幕云团\n（示意放大）", fontsize=9,
        color="#2F5D8A", ha="center", bbox=WB)
th = np.linspace(0.0, 2.0 * np.pi, 60)
for z in (0.0, H_T_S):
    ax.plot(R_T_S * np.cos(th), 200 + R_T_S * np.sin(th), np.full_like(th, z),
            color="#8F3236", lw=1.4)
for a in (0.0, np.pi / 2, np.pi, 3 * np.pi / 2):
    ax.plot([R_T_S * np.cos(a)] * 2, [200 + R_T_S * np.sin(a)] * 2, [0.0, H_T_S],
            color="#8F3236", lw=1.0)
ax.text(0, 640, 260, "真目标\n（示意放大）", fontsize=9, color="#8F3236", ha="center", bbox=WB)
ax.scatter([0], [0], [0], marker="x", color="k", s=70)
ax.text(900, -1500, 150, "假目标", fontsize=9.5, ha="center", bbox=WB)
ax.set_xticks([0, 5000, 10000, 15000, 20000])
ax.set_yticks([-400, 0, 400])
ax.set_zticks([0, 1000, 2000])
ax.set_xlabel("x (m)", labelpad=10)
ax.set_ylabel("y (m)", labelpad=10)
ax.set_zlabel("z (m)", labelpad=10)
ax.set_title("(a) 全场几何态势（云团/真目标未按比例）", fontsize=11, pad=2)
ax.view_init(elev=16, azim=-62)
ax.set_box_aspect((8, 1.4, 4.6))

# ---- (b) FY1–云团局部（含遮蔽窗口段弹道） ----
ax = fig.add_subplot(122, projection="3d")
tm2 = np.linspace(8.0, 10.5, 30)
Mm2 = np.array([missile(x) for x in tm2])
ax.plot(Mm2[:, 0], Mm2[:, 1], Mm2[:, 2], ls="--", color="#C44E52", lw=1.6,
        label="M1 弹道（遮蔽窗口段）")
ax.plot(Ff[:, 0], Ff[:, 1], Ff[:, 2], color="#2E5A87", lw=3.0,
        label="FY1 航线（120 m/s，等高）")
ax.scatter([P_FY1[0]], [P_FY1[1]], [P_FY1[2]], marker="^", color="#2E5A87", s=70)
ax.text(P_FY1[0] + 10, P_FY1[1] - 180, P_FY1[2] + 130, "FY1 起点", fontsize=9.5,
        color="#1F3A5C", bbox=WB)
ax.scatter(*D, color="#2E5A87", s=40)
ax.text(D[0] - 30, D[1] - 360, D[2] + 110, "投放点", fontsize=9.5,
        color="#1F3A5C", bbox=WB)
ax.scatter(*B, marker="*", color="#C44E52", s=130)
ax.text(B[0] + 40, B[1] + 300, B[2] - 330, "起爆点", fontsize=9.5,
        color="#8F3236", bbox=WB)
ax.plot([B[0] + 25, B[0] + 8], [240, 40], [B[2] - 280, B[2] - 85],
        color="#8F3236", lw=0.9, alpha=0.75)          # 引线：起爆点 → 星标
for off in (0.0, 10.0, 20.0):
    sphere(ax, cloud(T_DET + off), R_CLOUD_B, "#4C72B0", 0.25)
ax.plot([], [], ls="none", marker="o", ms=9, mfc="#4C72B0", mec="#4C72B0", alpha=0.45,
        label="烟幕云团（r=10 m，三时刻；示意放大 3×）")
ax.set_xticks([17000, 17500, 18000])
ax.set_yticks([0, 300])
ax.set_zticks([1500, 1750, 2000])
ax.set_xlabel("x (m)", labelpad=10)
ax.set_ylabel("y (m)", labelpad=10)
ax.text2D(1.115, 0.30, "z (m)", transform=ax.transAxes, rotation=90,
          fontsize=10.5, color="#404040")   # (b) 面板靠右，z 标签用 2D 文本防裁切
ax.set_title("(b) FY1–云团局部（起爆后 0/10/20 s 云团）", fontsize=11, pad=2)
ax.legend(loc="upper left", fontsize=8.5, frameon=True, framealpha=0.95, edgecolor="none")
ax.view_init(elev=14, azim=-62)
ax.set_box_aspect((8, 3.2, 3.0))

fig.subplots_adjust(left=0.02, right=0.98, top=0.93, bottom=0.04, wspace=0.06)
fig.canvas.draw()          # 3D 轴标签位置在绘制后才确定：先画一遍再导出，防 tight bbox 裁切
save_fig(fig, "F02_几何态势图", FIGS)

# ================= F03 遮蔽几何放大（三时刻视线扫过云团，x-z 投影，等比例） =================
C = cloud(8.75)   # 云团投影圆取窗口中点时刻位置（下沉 11 m 内圆心仅降 ~2 m，示意可忽略）
# 几何事实：真目标全部采样点的视线在云区近乎重合（散布 <1.5 m），故每时刻画一条代表视线；
# 遮蔽窗口的形成 = 视线扫入云团（t=8.06），结束 = 导弹飞抵云团边缘（t=9.45，弹目视线缩短至云心距≈10 m）
apply_style()
fig, ax = plt.subplots(figsize=(7.8, 3.9))
ax.add_patch(plt.Circle((C[0], C[2]), R_CLOUD, color="#4C72B0", alpha=0.28,
                        label="烟幕云团投影（r=10 m，t=8.75 s）"))
for tk, cc in [(8.057, "#C9862B"), (8.75, "#C44E52"), (9.449, "#8F3236")]:
    Mt = missile(tk)
    ax.plot([Mt[0], 0.0], [Mt[2], 5.0], color=cc, lw=1.3, alpha=0.9, zorder=1)
    ax.scatter([Mt[0]], [Mt[2]], marker="^", color=cc, s=60, zorder=4)
ax.plot([], [], color="#C9862B", lw=1.3, label="视线（至真目标，146 条近乎重合为一条）")
ax.scatter([], [], marker="^", s=60, color="#555555",
           label="M1 三个时刻位置：t=8.06 / 8.75 / 9.45 s")
ax.annotate("t=8.06 s 视线扫入云团\n→ 遮蔽开始", xy=(17188, 1719), xytext=(17066, 1655),
            fontsize=9, color="#7A5200",
            bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="none", alpha=0.9),
            arrowprops=dict(arrowstyle="->", color="#C9862B"))
ax.annotate("t=9.45 s 导弹飞抵云团边缘\n（弹目视线缩短）→ 遮蔽结束",
            xy=(17186, 1714), xytext=(17196, 1655),
            fontsize=9, color="#6B2B2E",
            bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="none", alpha=0.9),
            arrowprops=dict(arrowstyle="->", color="#8F3236"))
ax.set_aspect("equal")
ax.set_xlim(17060, 17650)
ax.set_ylim(1640, 1795)
ax.set_xlabel("x (m)")
ax.set_ylabel("z (m)")
ax.set_title("遮蔽窗口机理放大（x–z 投影，等比例）")
ax.legend(loc="upper left", fontsize=9, frameon=True, framealpha=0.95, edgecolor="none")
grid_on(ax, axis="both")
save_fig(fig, "F03_遮蔽几何放大", FIGS)

print("[done] 问题一图表全部生成 →", FIGS)
