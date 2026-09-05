# -*- coding: utf-8 -*-
"""
2025A 问题一：烟幕干扰弹有效遮蔽时长计算（公共仿真内核 · 标定运行）
  模型：02_建模/模型卡片-模型I.md（论文式 (1-1)~(1-5)）
  内核：烟幕内核.py（常数与判据的单一事实源；本脚本只做驱动与 CSV 导出）
  运行：python q1_遮蔽时长.py   预期 <2s；无随机过程，结果唯一可复现
  输出：F01_视线距离.csv（判据时间序列，dt=0.001 精扫，供 F01 绘图）
        F02_几何态势.csv（关键轨迹点存档）/ Q1_结论.csv（论文引用的唯一数值来源）
"""
import csv
import os
import sys

import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from 烟幕内核 import (EZ, G, H_T, P_FY1, R_CLOUD, T_CLOUD, TARGET_C,  # noqa: F401
                   V_CLOUD, _unit, cloud_pos, drop_state,
                   heading_unit, missile_pos, sample_target,
                   shielding_windows)

HERE = os.path.dirname(os.path.abspath(__file__))

# ---------- 问题一给定策略（题面常数经由内核同一来源） ----------
V_UAV = 120.0       # FY1 飞行速度（问题一给定，非内核常数）
HEADING, T_DROP, TAU = 180.0, 1.5, 3.6
DT = 0.001          # 报告口径：判据曲线精扫步长（窗口时长由细化边界解析给出）

st = drop_state(HEADING, V_UAV, T_DROP, TAU)
D, B, T_DET = st["D"], st["B"], st["t_det"]

# ---------- 判据扫描（dt=0.001 精扫供 F01 绘图；时长用细化值） ----------
duration, windows, ts, worst = shielding_windows(B, T_DET, dt=DT, refine=True)
blocked = worst <= R_CLOUD

# ---------- 派生几何量（供论文与 F03 引用，数字唯一来源） ----------
t_mid = 0.5 * (windows[0][0] + windows[0][1])
M_mid, C_mid = missile_pos(t_mid), cloud_pos(t_mid, B, T_DET)
# 视线束在云心距处的横向散布：max_p |u_p − u_c| × |C−M|（u 为单位视线方向）
u_c = _unit((TARGET_C + 0.5 * H_T * EZ) - M_mid)
MC = float(np.linalg.norm(C_mid - M_mid))
spread = float(max(np.linalg.norm(_unit(p - M_mid) - u_c) for p in sample_target()) * MC)

# ---------- 结果导出（CSV 交接约定：UTF-8、首行表头、文件名=图号） ----------
with open(os.path.join(HERE, "F01_视线距离.csv"), "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["时间s", "最坏视线距离m", "是否被遮"])
    for t, d, b in zip(ts, worst, blocked):
        w.writerow([f"{t:.3f}", f"{d:.4f}", int(b)])

with open(os.path.join(HERE, "F02_几何态势.csv"), "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["类别", "x", "y", "z", "序"])
    seq = 0
    for t in np.arange(0.0, 66.0, 2.0):
        p = missile_pos(t)
        w.writerow(["导弹M1弹道", f"{p[0]:.1f}", f"{p[1]:.1f}", f"{p[2]:.1f}", seq]); seq += 1
    for t in np.arange(0.0, T_DET + 0.5, 0.5):
        p = P_FY1 + V_UAV * t * heading_unit(HEADING)
        w.writerow(["FY1航线", f"{p[0]:.1f}", f"{p[1]:.1f}", f"{p[2]:.1f}", seq]); seq += 1
    for name, p in [("投放点", D), ("起爆点", B),
                    ("云团@起爆", cloud_pos(T_DET, B, T_DET)),
                    ("云团@+10s", cloud_pos(T_DET + 10, B, T_DET)),
                    ("云团@+20s", cloud_pos(T_DET + 20, B, T_DET)),
                    ("真目标下底", TARGET_C), ("真目标上底", TARGET_C + H_T * EZ),
                    ("假目标", np.zeros(3))]:
        w.writerow([name, f"{p[0]:.1f}", f"{p[1]:.1f}", f"{p[2]:.1f}", seq]); seq += 1

with open(os.path.join(HERE, "Q1_结论.csv"), "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["指标", "数值"])
    for k, v in [
        ("FY1航向角(度,x正向逆时针)", f"{st['heading_deg']:.2f}"),
        ("投放点(m)", "({:.2f}, {:.2f}, {:.2f})".format(*D)),
        ("起爆点(m)", "({:.2f}, {:.2f}, {:.2f})".format(*B)),
        ("起爆时刻(s)", f"{T_DET:.2f}"),
        ("遮蔽窗口-细化(s)", "; ".join(f"[{a:.4f}, {b:.4f}]" for a, b in windows)),
        ("有效遮蔽时长(s)", f"{duration:.4f}"),
        ("窗口中点时刻(s)", f"{t_mid:.3f}"),
        ("窗口中点导弹位置(m)", "({:.2f}, {:.2f}, {:.2f})".format(*M_mid)),
        ("窗口中点云团中心(m)", "({:.2f}, {:.2f}, {:.2f})".format(*C_mid)),
        ("视线束云心处横向散布(m)", f"{spread:.3f}"),
        ("判据口径", "视线段-球心距离≤10m, 全柱146点采样, g=9.8, 曲线dt=0.001s, 边界二分细化1e-4s"),
    ]:
        w.writerow([k, v])

# ---------- 自检三问（角色卡 03）：量纲/量级/边界 + 回归断言 ----------
print("=== 自检 ===")
assert abs(st["heading_deg"] - 180.0) < 1e-9 and np.isclose(D[2], P_FY1[2]), "等高度检查失败"
print(f"航向角 {st['heading_deg']:.2f}°（期望 180°）；等高度：投放点 z={D[2]:.1f}（应=1800）")
print(f"投放点 {D.round(2)}  起爆点 {B.round(2)}（z 低于投放点、x 更靠近原点）")
assert 0 < duration < T_CLOUD, "时长超出云团有效窗口，量级异常"
print(f"起爆时刻 {T_DET:.1f}s，云团窗口 [{T_DET:.1f}, {T_DET + T_CLOUD:.1f}]s")
print("=== 问题一结果（细化口径） ===")
for a, b in windows:
    print(f"遮蔽窗口: [{a:.4f}, {b:.4f}] s")
print(f"有效遮蔽时长 = {duration:.4f} s（粗扫口径 1.392 s，正文报告 1.392 s）")
print(f"窗口中点 t={t_mid:.3f}s：导弹 {M_mid.round(2)}，云团中心 {C_mid.round(2)}，视线束散布 ≤{spread:.2f} m")
