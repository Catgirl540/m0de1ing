# -*- coding: utf-8 -*-
"""
烟幕内核.py — 2025A 公共仿真内核（单一事实源：运动学常数与遮蔽判据全工作区仅此一份）

消费方：
  · q1_遮蔽时长.py          问题一驱动（写 F01/F02/Q1_结论 三张 CSV）
  · 04_图表/绘制问题一图表.py  绘图常数同源 import，杜绝"文图不符"
  · Q2–Q5 优化              duration() 即差分进化的目标函数（决策变量→遮蔽时长）

遮蔽判据口径（模型卡片 H4/H5，论文 §5.1 声明）：
  导弹 M → 真目标表面任一采样点的视线段与云团球心距离 ≤ R_CLOUD；
  真目标按上/下底圆周各 n_theta 点加两底圆心采样，全部视线被遮方计有效。
"""
from __future__ import annotations

import numpy as np

# ---------- 题面常数（唯一副本，勿在其他文件复制） ----------
V_MISS, V_CLOUD, R_CLOUD, T_CLOUD = 300.0, 3.0, 10.0, 20.0
G = 9.8
R_T, H_T = 7.0, 10.0
TARGET_C = np.array([0.0, 200.0, 0.0])          # 真目标下底圆心
P_FY1 = np.array([17800.0, 0.0, 1800.0])
EZ = np.array([0.0, 0.0, 1.0])

MISSILES = {"M1": np.array([20000.0, 0.0, 2000.0]),
            "M2": np.array([19000.0, 600.0, 2100.0]),
            "M3": np.array([18000.0, -600.0, 1900.0])}
UVS = {"FY1": np.array([17800.0, 0.0, 1800.0]),
       "FY2": np.array([12000.0, 1400.0, 1400.0]),
       "FY3": np.array([6000.0, -3000.0, 700.0]),
       "FY4": np.array([11000.0, 2000.0, 1800.0]),
       "FY5": np.array([13000.0, -2000.0, 1300.0])}


def _unit(v):
    return v / np.linalg.norm(v)


def heading_unit(heading_deg):
    """方向角（x 正向逆时针 0~360°，result 模板口径）→ 水平面单位矢量（等高度）。"""
    r = np.radians(heading_deg)
    return np.array([np.cos(r), np.sin(r), 0.0])


def missile_pos(t, missile_id="M1"):
    return MISSILES[missile_id] + V_MISS * t * _unit(-MISSILES[missile_id])


def cloud_pos(t, B, t_det):
    return B + (t - t_det) * (-V_CLOUD) * EZ


def drop_state(heading_deg, v_uav, t_drop, tau, uav_id="FY1"):
    """(方向角, 速度, 投放时刻, 引信延时) → 投放/起爆几何。Q2 的决策变量入口。"""
    u = heading_unit(heading_deg)
    P = UVS[uav_id]
    D = P + v_uav * t_drop * u
    v_bomb = v_uav * u
    B = D + v_bomb * tau - 0.5 * G * tau ** 2 * EZ
    return dict(D=D, B=B, v_bomb=v_bomb, t_det=t_drop + tau,
                heading_deg=heading_deg % 360.0)


def sample_target(n_theta=72):
    """真目标全柱采样：上/下底圆周各 n_theta 点 + 两底圆心 → (2n+2, 3)。"""
    th = np.linspace(0.0, 2.0 * np.pi, n_theta, endpoint=False)
    circ = [np.column_stack([TARGET_C[0] + R_T * np.cos(th),
                             TARGET_C[1] + R_T * np.sin(th),
                             np.full(th.shape, TARGET_C[2] + z)])
            for z in (0.0, H_T)]
    return np.vstack(circ + [TARGET_C, TARGET_C + H_T * EZ])


def _seg_dist(C, A, pts):
    """球心 C 到视线段 A→p 的距离，p ∈ pts（向量化点到线段）。"""
    AB = pts - A
    AC = C - A
    denom = np.maximum((AB * AB).sum(axis=1), 1e-12)
    tt = np.clip((AB @ AC) / denom, 0.0, 1.0)
    foot = A + tt[:, None] * AB
    return np.linalg.norm(foot - C, axis=1)


def max_los_dist(t, B, t_det, missile_id="M1", pts=None):
    """t 时刻的"最坏视线距离"（全部采样视线的最大球心距）。标量时刻。"""
    pts = sample_target() if pts is None else pts
    d = _seg_dist(cloud_pos(t, B, t_det), missile_pos(t, missile_id), pts)
    return float(d.max())


def worst_series(ts, B, t_det, missile_id="M1", pts=None, chunk=2000):
    """向量化版：一次性计算 ts 全部时刻的最坏视线距离 (T,)。

    分块广播 (T_c, N, 3)，避免 2 万时刻 × 146 点一次性占内存过大。
    """
    pts = sample_target() if pts is None else pts
    um = _unit(-MISSILES[missile_id])
    out = np.empty(ts.size)
    for i in range(0, ts.size, chunk):
        tc = ts[i:i + chunk]
        M = MISSILES[missile_id][None, :] + V_MISS * tc[:, None] * um[None, :]   # (Tc,3)
        C = B[None, :] + (tc - t_det)[:, None] * (-V_CLOUD) * EZ[None, :]        # (Tc,3)
        AB = pts[None, :, :] - M[:, None, :]                                     # (Tc,N,3)
        AC = C[:, None, :] - M[:, None, :]
        denom = np.maximum((AB * AB).sum(-1), 1e-12)
        tt = np.clip((AB * AC).sum(-1) / denom, 0.0, 1.0)
        foot = M[:, None, :] + tt[..., None] * AB
        out[i:i + chunk] = np.linalg.norm(foot - C[:, None, :], axis=-1).max(-1)
    return out


def _worst_series_loop(ts, B, t_det, missile_id="M1", pts=None):
    """参考实现（逐时步循环，仅用于内核自检的性能基准对照）。"""
    pts = sample_target() if pts is None else pts
    um = _unit(-MISSILES[missile_id])
    M0 = MISSILES[missile_id]
    out = np.empty(ts.size)
    for i, t in enumerate(ts):
        M = M0 + V_MISS * t * um
        C = cloud_pos(t, B, t_det)
        out[i] = _seg_dist(C, M, pts).max()
    return out


def _blocked(t, B, t_det, missile_id, pts):
    return max_los_dist(t, B, t_det, missile_id, pts) <= R_CLOUD


def _refine_edge(t_keep, t_flip, B, t_det, missile_id, pts, tol=1e-4):
    """窗口边界二分细化：t_keep 侧被遮、t_flip 侧未遮，收敛到 tol 秒。"""
    for _ in range(40):
        if abs(t_flip - t_keep) <= tol:
            break
        mid = 0.5 * (t_keep + t_flip)
        if _blocked(mid, B, t_det, missile_id, pts):
            t_keep = mid
        else:
            t_flip = mid
    return t_keep


def shielding_windows(B, t_det, missile_id="M1", dt=0.001, refine=True, n_theta=72):
    """扫描云团有效窗口 → (时长, 细化窗口列表, ts, worst)。

    dt 为粗扫步长；refine=True 时窗口边界二分细化到 1e-4 s，
    时长由细化边界直接解析得出（消除网格量化误差）。
    """
    pts = sample_target(n_theta)
    ts = np.arange(t_det, t_det + T_CLOUD, dt)
    worst = worst_series(ts, B, t_det, missile_id, pts)
    blocked = worst <= R_CLOUD

    raw, start = [], None
    for i, b in enumerate(blocked):
        if b and start is None:
            start = ts[i]
        elif not b and start is not None:
            raw.append((start, ts[i - 1]))
            start = None
    if start is not None:
        raw.append((start, ts[-1]))

    windows = []
    for a, b in raw:
        a2 = _refine_edge(a, a - dt, B, t_det, missile_id, pts) if refine else a
        b2 = _refine_edge(b, b + dt, B, t_det, missile_id, pts) if refine else b
        windows.append((a2, b2))
    duration = float(sum(b - a for a, b in windows))
    return duration, windows, ts, worst


def duration(heading_deg, v_uav, t_drop, tau, uav_id="FY1", missile_id="M1",
             dt=0.01, refine=True):
    """一行调用：投放策略 → 有效遮蔽时长。Q2–Q5 优化目标函数（越大越好）。

    性能口径（实测）：dt=0.01+refine 单次约 26 ms，与 dt=0.001 结果一致到 1e-4 s
    ——粗扫步长只影响"漏检宽度 <dt 的微窗口"，边界精度由二分细化保证；
    优化阶段用默认 dt=0.01，正文报告口径用 dt=0.001。单次 26 ms 时
    差分进化 4000 次评估约 1.7 分钟。
    """
    st = drop_state(heading_deg, v_uav, t_drop, tau, uav_id)
    d, _, _, _ = shielding_windows(st["B"], st["t_det"], missile_id,
                                   dt=dt, refine=refine)
    return d


# ---------------------------------------------------------------- 自检与基准
if __name__ == "__main__":
    import time

    st = drop_state(180.0, 120.0, 1.5, 3.6)
    assert abs(st["heading_deg"] - 180.0) < 1e-9
    assert np.allclose(st["D"], [17620.0, 0.0, 1800.0]), "投放点回归失败"
    assert np.allclose(st["B"], [17188.0, 0.0, 1736.496], atol=5e-3), "起爆点回归失败"

    pts = sample_target()
    ts = np.arange(st["t_det"], st["t_det"] + T_CLOUD, 0.001)

    t0 = time.perf_counter()
    ref = _worst_series_loop(ts, st["B"], st["t_det"], pts=pts)
    t_loop = time.perf_counter() - t0
    t0 = time.perf_counter()
    vec = worst_series(ts, st["B"], st["t_det"], pts=pts)
    t_vec = time.perf_counter() - t0
    assert np.allclose(ref, vec, atol=1e-9), "向量化与参考实现不一致"

    dur, wins, _, _ = shielding_windows(st["B"], st["t_det"], n_theta=72)
    print(f"[内核自检] 循环版 {t_loop*1e3:.0f} ms | 向量化(dt=0.001) {t_vec*1e3:.0f} ms | "
          f"加速 {t_loop/t_vec:.1f}×")
    t0 = time.perf_counter()
    dur_fast, wins_fast, _, _ = shielding_windows(st["B"], st["t_det"],
                                                  dt=0.01, n_theta=72)
    t_fast = time.perf_counter() - t0
    assert abs(dur_fast - dur) < 1e-4, "dt=0.01+refine 与 dt=0.001 结果不一致"
    print(f"[内核自检] dt=0.01+refine {t_fast*1e3:.0f} ms（Q2 目标函数口径），"
          f"与 dt=0.001 时长一致：{dur_fast:.4f} s")
    print(f"[内核自检] 细化窗口: " + "; ".join(f"[{a:.4f}, {b:.4f}]" for a, b in wins))
    print(f"[内核自检] 有效遮蔽时长 = {dur:.4f} s")
