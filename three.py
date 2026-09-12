"""2026 CUMCM C题问题三：多时点光伏预报下的 MPC-LP 购电策略。

在问题二模型基础上：
- 负荷继承严格前向贝叶斯优化后的年度谐波—岭回归—傅里叶—Haar小波模型，并用当日已观测负荷在线纠偏；
- 光伏使用附件3在0:00、6:00、12:00、18:00发布的未来24小时预报；
- 每次求解未来24小时，只执行到下一次预报发布时刻；
- 增购部分按1.5倍电价，减购部分退回原价并支付50%违约费；
- 真实供电缺口按5倍电价紧急购电。

程序比较四种预报组合，并将总费用最低的策略写入正式结果。默认生成：
  C题/result3.xlsx
  C题/问题三逐日汇总.csv
  C题/问题三预报时刻对比.csv
其中策略只按验证集费用选择，测试集仅用于最终泛化检验。

依赖：two.py、numpy、scipy、openpyxl
运行：python three.py
"""

from __future__ import annotations

import argparse
import csv
import math
import os
import shutil
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
from openpyxl import load_workbook

import two


ROOT = Path(__file__).resolve().parent
DEFAULT_FORECAST = ROOT / "C题" / "附件" / "csv" / "附件3.csv"
DEFAULT_TEMPLATE = ROOT / "C题" / "附件" / "附件5" / "result3.xlsx"
DEFAULT_OUTPUT = ROOT / "C题" / "result3.xlsx"
DEFAULT_SUMMARY = ROOT / "C题" / "问题三逐日汇总.csv"
DEFAULT_COMPARISON = ROOT / "C题" / "问题三预报时刻对比.csv"
DEFAULT_SPLIT_REPORT = ROOT / "C题" / "问题三_训练验证测试检验.csv"


@dataclass(frozen=True)
class Config(two.Config):
    adjustment_up_multiplier: float = 1.50
    adjustment_down_penalty: float = 0.50
    load_bias_half_life_hours: float = 6.0


@dataclass
class ForecastData:
    values: dict[tuple[date, int], np.ndarray]


@dataclass
class DecisionInput:
    load_kw: np.ndarray
    pv_kw: np.ndarray
    reserve_kwh: np.ndarray


@dataclass
class AdjustmentSolution:
    grid_kwh: np.ndarray
    charge_kwh: np.ndarray
    discharge_kwh: np.ndarray
    curtailment_kwh: np.ndarray
    increase_kwh: np.ndarray
    decrease_kwh: np.ndarray
    soc_kwh: np.ndarray


@dataclass
class DailyResult:
    day: date
    plan_grid_kwh: np.ndarray
    final_grid_kwh: np.ndarray
    charge_kwh: np.ndarray
    discharge_kwh: np.ndarray
    emergency_kwh: np.ndarray
    actual_curtailment_kwh: np.ndarray
    soc_start_kwh: float
    soc_end_kwh: float
    plan_cost_yuan: float
    settled_purchase_cost_yuan: float
    emergency_cost_yuan: float

    @property
    def adjustment_delta_cost_yuan(self) -> float:
        return self.settled_purchase_cost_yuan - self.plan_cost_yuan

    @property
    def total_cost_yuan(self) -> float:
        return self.settled_purchase_cost_yuan + self.emergency_cost_yuan


@dataclass
class StrategySummary:
    name: str
    release_hours: tuple[int, ...]
    results: list[DailyResult]
    plan_energy_kwh: float
    final_energy_kwh: float
    increase_energy_kwh: float
    decrease_energy_kwh: float
    emergency_energy_kwh: float
    plan_cost_yuan: float
    settled_purchase_cost_yuan: float
    emergency_cost_yuan: float
    total_cost_yuan: float


def read_forecasts(path: Path) -> ForecastData:
    if not path.exists():
        raise FileNotFoundError(f"找不到附件3 CSV：{path}")
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file)
        header = next(reader, None)
        rows = list(reader)
    if header is None or len(header) != 26:
        raise ValueError("附件3应包含日期、预报时刻和未来24小时预报共26列。")
    if len(rows) != 365 * 4:
        raise ValueError("附件3应包含365天、每天4个发布时间的数据。")

    values: dict[tuple[date, int], np.ndarray] = {}
    current_date: date | None = None
    for row_number, row in enumerate(rows, start=2):
        if len(row) != 26:
            raise ValueError(f"附件3第{row_number}行列数不是26。")
        if row[0].strip():
            try:
                current_date = datetime.strptime(row[0].strip(), "%Y-%m-%d").date()
            except ValueError as exc:
                raise ValueError(f"附件3第{row_number}行日期无法解析：{row[0]!r}") from exc
        if current_date is None:
            raise ValueError(f"附件3第{row_number}行缺少所属日期。")
        try:
            release_hour = int(row[1].strip().split(":", 1)[0])
        except (ValueError, IndexError) as exc:
            raise ValueError(f"附件3第{row_number}行预报时刻无法解析。") from exc
        if release_hour not in (0, 6, 12, 18):
            raise ValueError(f"附件3第{row_number}行预报时刻不是0、6、12或18时。")
        forecast = np.asarray(
            [two._float(value, f"附件3第{row_number}行预报值") for value in row[2:]],
            dtype=float,
        )
        if np.any(forecast < 0):
            raise ValueError(f"附件3第{row_number}行包含负光伏功率。")
        key = (current_date, release_hour)
        if key in values:
            raise ValueError(f"附件3存在重复预报：{key}")
        values[key] = forecast

    expected_dates = [date(2025, 1, 1) + timedelta(days=i) for i in range(365)]
    expected_keys = {(day, hour) for day in expected_dates for hour in (0, 6, 12, 18)}
    if set(values) != expected_keys:
        raise ValueError("附件3没有完整覆盖2025年每天的四个预报时刻。")
    return ForecastData(values)


def forecast_load_profile(
    historical: two.HistoricalData,
    baseline: two.BaselineData,
    cutoff: int,
    target_date: date,
    hours: np.ndarray,
    cfg: Config,
    allowed_indices: set[int] | None = None,
) -> np.ndarray:
    profile = two._weighted_history_profile(
        historical.load_kw,
        historical.dates,
        cutoff,
        target_date,
        baseline.load_kw,
        cfg,
        allowed_indices,
    )
    dynamic_load = two.fit_dynamic_load_profile(
        historical,
        cutoff,
        target_date,
        baseline.load_kw,
        cfg,
        allowed_indices,
    )
    blend = float(np.clip(cfg.dynamic_load_blend, 0.0, 1.0))
    profile = (1.0 - blend) * profile + blend * dynamic_load
    fourier = two.fit_load_fourier(hours, profile)
    residual = two.wavelet_correct_residual(profile, fourier, cfg.wavelet_level)
    return np.maximum(0.0, fourier + residual)


def _horizon_from_daily(
    current: np.ndarray, following: np.ndarray, start_index: int
) -> np.ndarray:
    if start_index == 0:
        return current.copy()
    return np.concatenate([current[start_index:], following[:start_index]])


def interpolate_pv_forecast(
    hourly_forecast: np.ndarray, anchor_power_kw: float
) -> np.ndarray:
    """将未来1至24小时整点预报线性插值为144个10分钟功率值。"""
    xp = np.arange(25, dtype=float)
    fp = np.concatenate([[max(0.0, anchor_power_kw)], hourly_forecast])
    target = np.arange(1, 145, dtype=float) / 6.0
    return np.maximum(0.0, np.interp(target, xp, fp))


def actual_horizon(matrix: np.ndarray, day_index: int, start_index: int) -> np.ndarray:
    if start_index == 0:
        return matrix[day_index].copy()
    if day_index + 1 >= len(matrix):
        raise IndexError("年末最后一天之后没有真实数据，不能构造误差回放窗口。")
    return np.concatenate(
        [matrix[day_index, start_index:], matrix[day_index + 1, :start_index]]
    )


def prepare_decision_inputs(
    baseline: two.BaselineData,
    historical: two.HistoricalData,
    forecasts: ForecastData,
    cfg: Config,
    training_cutoff: int | None = None,
    training_indices: set[int] | None = None,
) -> dict[tuple[int, int], DecisionInput]:
    hours = np.arange(144, dtype=float) * cfg.dt_hours
    raw: dict[tuple[int, int], tuple[np.ndarray, np.ndarray]] = {}

    for day_index, current_date in enumerate(historical.dates):
        # 训练期逐日因果拟合；验证和测试阶段固定使用训练集样本。
        model_cutoff = (
            day_index
            if training_cutoff is None
            else min(day_index, training_cutoff)
        )
        allowed_indices = training_indices
        forecast_cfg = two.bayesian_load_config(cfg, current_date)
        current_load = forecast_load_profile(
            historical, baseline, model_cutoff, current_date, hours, forecast_cfg,
            allowed_indices=allowed_indices,
        )
        next_load = forecast_load_profile(
            historical,
            baseline,
            model_cutoff,
            current_date + timedelta(days=1),
            hours,
            forecast_cfg,
            allowed_indices=training_indices,
        )
        for release_hour in (0, 6, 12, 18):
            start = release_hour * 6
            load_horizon = _horizon_from_daily(current_load, next_load, start)

            # 利用发布时刻之前最近1小时的真实负荷偏差进行在线修正。
            if start > 0:
                observed = historical.load_kw[day_index, max(0, start - 6) : start]
                predicted = current_load[max(0, start - 6) : start]
                bias = float(np.mean(observed - predicted))
                half_life_steps = cfg.load_bias_half_life_hours / cfg.dt_hours
                decay = np.exp(-math.log(2.0) * np.arange(144) / half_life_steps)
                load_horizon = np.maximum(0.0, load_horizon + bias * decay)

            anchor = float(historical.pv_kw[day_index, start])
            pv_horizon = interpolate_pv_forecast(
                forecasts.values[(current_date, release_hour)], anchor
            )
            raw[(day_index, release_hour)] = (load_horizon, pv_horizon)

    prepared: dict[tuple[int, int], DecisionInput] = {}
    for day_index, _ in enumerate(historical.dates):
        for release_hour in (0, 6, 12, 18):
            load_horizon, pv_horizon = raw[(day_index, release_hour)]
            residuals: list[np.ndarray] = []
            first = max(0, day_index - cfg.risk_days)
            for historical_index in range(first, day_index):
                if training_indices is not None and historical_index not in training_indices:
                    continue
                # 历史同一发布时间的24小时误差在当前发布时间已经全部可观测。
                hist_load, hist_pv = raw[(historical_index, release_hour)]
                actual_load = actual_horizon(
                    historical.load_kw,
                    historical_index,
                    release_hour * 6,
                )
                actual_pv = actual_horizon(
                    historical.pv_kw,
                    historical_index,
                    release_hour * 6,
                )
                residuals.append(
                    (actual_load - actual_pv - hist_load + hist_pv) * cfg.dt_hours
                )
            if len(residuals) >= cfg.minimum_risk_days:
                reserve = np.maximum(
                    0.0,
                    np.quantile(np.asarray(residuals), cfg.risk_quantile, axis=0),
                )
            else:
                reserve = np.zeros(144, dtype=float)
            prepared[(day_index, release_hour)] = DecisionInput(
                load_horizon, pv_horizon, reserve
            )
    return prepared


def solve_adjustment_lp(
    price: np.ndarray,
    load_kw: np.ndarray,
    pv_kw: np.ndarray,
    initial_soc_kwh: float,
    original_plan_current_day: np.ndarray,
    current_day_slots: int,
    cfg: Config,
) -> AdjustmentSolution:
    try:
        from scipy.optimize import linprog
        from scipy.sparse import coo_matrix
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError("缺少scipy，请先执行：python -m pip install scipy") from exc

    n = 144
    if any(len(array) != n for array in (price, load_kw, pv_kw)):
        raise ValueError("MPC调整模型的预测窗口必须包含144个10分钟时段。")
    if len(original_plan_current_day) != current_day_slots:
        raise ValueError("原计划参考量长度与当天剩余时段数不一致。")
    load_energy = load_kw * cfg.dt_hours
    pv_energy = pv_kw * cfg.dt_hours

    g0, c0, d0, w0, u0, v0, e0 = (
        0,
        n,
        2 * n,
        3 * n,
        4 * n,
        5 * n,
        6 * n,
    )
    variable_count = 7 * n + 1
    objective = np.zeros(variable_count, dtype=float)
    objective[c0 : c0 + n] = cfg.throughput_tie_breaker
    objective[d0 : d0 + n] = cfg.throughput_tie_breaker
    objective[u0 : u0 + current_day_slots] = (
        cfg.adjustment_up_multiplier * price[:current_day_slots]
    )
    objective[v0 : v0 + current_day_slots] = (
        -cfg.adjustment_down_penalty * price[:current_day_slots]
    )
    # 跨过午夜的远期时段尚未形成下一日原计划，按普通购电价格计入终端规划。
    objective[g0 + current_day_slots : g0 + n] = price[current_day_slots:]

    rows: list[int] = []
    columns: list[int] = []
    coefficients: list[float] = []
    equality_count = 2 * n + current_day_slots
    b_eq = np.zeros(equality_count, dtype=float)
    for t in range(n):
        rows.extend([t, t, t, t])
        columns.extend([g0 + t, c0 + t, d0 + t, w0 + t])
        coefficients.extend([1.0, -1.0, 1.0, -1.0])
        b_eq[t] = load_energy[t] - pv_energy[t]

        row = n + t
        rows.extend([row, row, row, row])
        columns.extend([c0 + t, d0 + t, e0 + t, e0 + t + 1])
        coefficients.extend(
            [-cfg.charge_efficiency, 1.0 / cfg.discharge_efficiency, -1.0, 1.0]
        )

    for t in range(current_day_slots):
        row = 2 * n + t
        # g = G_plan + u - v
        rows.extend([row, row, row])
        columns.extend([g0 + t, u0 + t, v0 + t])
        coefficients.extend([1.0, -1.0, 1.0])
        b_eq[row] = original_plan_current_day[t]

    a_eq = coo_matrix(
        (coefficients, (rows, columns)), shape=(equality_count, variable_count)
    ).tocsr()

    bounds: list[tuple[float | None, float | None]] = []
    bounds.extend([(0.0, None)] * n)
    bounds.extend([(0.0, cfg.max_charge_kwh)] * n)
    bounds.extend([(0.0, cfg.max_discharge_kwh)] * n)
    bounds.extend([(0.0, None)] * n)
    bounds.extend([(0.0, None)] * current_day_slots)
    bounds.extend([(0.0, 0.0)] * (n - current_day_slots))
    bounds.extend([(0.0, None)] * current_day_slots)
    bounds.extend([(0.0, 0.0)] * (n - current_day_slots))
    bounds.extend([(cfg.soc_min_kwh, cfg.soc_max_kwh)] * (n + 1))
    bounds[e0] = (initial_soc_kwh, initial_soc_kwh)
    bounds[e0 + n] = (initial_soc_kwh, initial_soc_kwh)

    result = linprog(
        objective,
        A_eq=a_eq,
        b_eq=b_eq,
        bounds=bounds,
        method="highs",
        options={"presolve": True},
    )
    if not result.success:
        raise RuntimeError(f"MPC调整线性规划失败：{result.status}, {result.message}")

    x = np.asarray(result.x, dtype=float)
    solution = AdjustmentSolution(
        x[g0 : g0 + n],
        x[c0 : c0 + n],
        x[d0 : d0 + n],
        x[w0 : w0 + n],
        x[u0 : u0 + n],
        x[v0 : v0 + n],
        x[e0 : e0 + n + 1],
    )
    validate_adjustment_solution(
        solution,
        load_energy,
        pv_energy,
        initial_soc_kwh,
        original_plan_current_day,
        current_day_slots,
        cfg,
    )
    return solution


def validate_adjustment_solution(
    solution: AdjustmentSolution,
    load_energy: np.ndarray,
    pv_energy: np.ndarray,
    initial_soc_kwh: float,
    original_plan: np.ndarray,
    current_day_slots: int,
    cfg: Config,
) -> None:
    tolerance = 1.0e-5
    balance = (
        solution.grid_kwh
        + pv_energy
        + solution.discharge_kwh
        - load_energy
        - solution.charge_kwh
        - solution.curtailment_kwh
    )
    next_soc = (
        solution.soc_kwh[:-1]
        + cfg.charge_efficiency * solution.charge_kwh
        - solution.discharge_kwh / cfg.discharge_efficiency
    )
    adjustment = (
        solution.grid_kwh[:current_day_slots]
        - solution.increase_kwh[:current_day_slots]
        + solution.decrease_kwh[:current_day_slots]
        - original_plan
    )
    checks = {
        "电量平衡": float(np.max(np.abs(balance))),
        "SOC递推": float(np.max(np.abs(solution.soc_kwh[1:] - next_soc))),
        "调整量关系": float(np.max(np.abs(adjustment))),
        "初始SOC": abs(float(solution.soc_kwh[0]) - initial_soc_kwh),
        "终端SOC": abs(float(solution.soc_kwh[-1]) - initial_soc_kwh),
    }
    failed = {name: value for name, value in checks.items() if value > tolerance}
    if failed:
        raise RuntimeError(f"MPC调整结果校验失败：{failed}")
    if np.min(solution.soc_kwh) < cfg.soc_min_kwh - tolerance:
        raise RuntimeError("MPC调整结果SOC低于下限。")
    if np.max(solution.soc_kwh) > cfg.soc_max_kwh + tolerance:
        raise RuntimeError("MPC调整结果SOC高于上限。")
    if np.max(np.minimum(solution.charge_kwh, solution.discharge_kwh)) > 1.0e-4:
        raise RuntimeError("MPC调整结果出现同一时段同时充放电。")


def settled_purchase_cost(
    price: np.ndarray,
    original_plan: np.ndarray,
    final_grid: np.ndarray,
    cfg: Config,
) -> float:
    increase = np.maximum(0.0, final_grid - original_plan)
    decrease = np.maximum(0.0, original_plan - final_grid)
    # 计划费用 + 1.5倍增购费用 - 已取消电量原价退款 + 50%违约费
    return float(
        np.dot(price, original_plan)
        + np.dot(cfg.adjustment_up_multiplier * price, increase)
        - np.dot((1.0 - cfg.adjustment_down_penalty) * price, decrease)
    )


def execute_soc(
    initial_soc: float,
    charge: np.ndarray,
    discharge: np.ndarray,
    cfg: Config,
) -> float:
    soc = float(initial_soc)
    for c_value, d_value in zip(charge, discharge):
        soc += cfg.charge_efficiency * float(c_value)
        soc -= float(d_value) / cfg.discharge_efficiency
        if not (cfg.soc_min_kwh - 1.0e-4 <= soc <= cfg.soc_max_kwh + 1.0e-4):
            raise RuntimeError(f"执行阶段SOC越界：{soc}")
    return soc


def simulate_strategy(
    name: str,
    release_hours: tuple[int, ...],
    baseline: two.BaselineData,
    historical: two.HistoricalData,
    decision_inputs: dict[tuple[int, int], DecisionInput],
    cfg: Config,
) -> StrategySummary:
    if 0 not in release_hours:
        raise ValueError("所有策略都必须包含0:00计划。")
    results: list[DailyResult] = []
    soc = cfg.initial_soc_kwh

    for day_index, current_date in enumerate(historical.dates):
        soc_start = soc
        input_zero = decision_inputs[(day_index, 0)]
        risk_load = input_zero.load_kw + input_zero.reserve_kwh / cfg.dt_hours
        plan_solution = two.solve_rolling_lp(
            baseline.price,
            risk_load,
            input_zero.pv_kw,
            soc,
            cfg,
        )
        plan_grid = plan_solution.grid_kwh.copy()
        final_grid = plan_solution.grid_kwh.copy()
        final_charge = plan_solution.charge_kwh.copy()
        final_discharge = plan_solution.discharge_kwh.copy()

        for release_hour in (0, 6, 12, 18):
            start = release_hour * 6
            if release_hour > 0 and release_hour in release_hours:
                decision = decision_inputs[(day_index, release_hour)]
                risk_load = decision.load_kw + decision.reserve_kwh / cfg.dt_hours
                price_horizon = np.concatenate(
                    [baseline.price[start:], baseline.price[:start]]
                )
                current_slots = 144 - start
                adjusted = solve_adjustment_lp(
                    price_horizon,
                    risk_load,
                    decision.pv_kw,
                    soc,
                    plan_grid[start:],
                    current_slots,
                    cfg,
                )
                final_grid[start:] = adjusted.grid_kwh[:current_slots]
                final_charge[start:] = adjusted.charge_kwh[:current_slots]
                final_discharge[start:] = adjusted.discharge_kwh[:current_slots]

            stop = min(start + 36, 144)
            soc = execute_soc(
                soc,
                final_charge[start:stop],
                final_discharge[start:stop],
                cfg,
            )

        emergency, actual_curtailment = two.replay_actual_day(
            historical.load_kw[day_index],
            historical.pv_kw[day_index],
            final_grid,
            final_charge,
            final_discharge,
            cfg,
        )
        plan_cost = float(np.dot(baseline.price, plan_grid))
        purchase_cost = settled_purchase_cost(
            baseline.price, plan_grid, final_grid, cfg
        )
        emergency_cost = float(
            np.dot(cfg.emergency_price_multiplier * baseline.price, emergency)
        )
        results.append(
            DailyResult(
                current_date,
                plan_grid,
                final_grid,
                final_charge,
                final_discharge,
                emergency,
                actual_curtailment,
                soc_start,
                soc,
                plan_cost,
                purchase_cost,
                emergency_cost,
            )
        )

        if (day_index + 1) % 90 == 0 or day_index == len(historical.dates) - 1:
            print(
                f"{name} 已完成 {current_date.isoformat()}："
                f"SOC={soc:.2f} kWh，紧急购电={np.sum(emergency):.2f} kWh"
            )

    output_results = [item for item in results if item.day >= date(2025, 2, 1)]
    plan_energy = float(sum(np.sum(item.plan_grid_kwh) for item in output_results))
    final_energy = float(sum(np.sum(item.final_grid_kwh) for item in output_results))
    increase_energy = float(
        sum(
            np.sum(np.maximum(0.0, item.final_grid_kwh - item.plan_grid_kwh))
            for item in output_results
        )
    )
    decrease_energy = float(
        sum(
            np.sum(np.maximum(0.0, item.plan_grid_kwh - item.final_grid_kwh))
            for item in output_results
        )
    )
    emergency_energy = float(
        sum(np.sum(item.emergency_kwh) for item in output_results)
    )
    plan_cost = float(sum(item.plan_cost_yuan for item in output_results))
    purchase_cost = float(
        sum(item.settled_purchase_cost_yuan for item in output_results)
    )
    emergency_cost = float(sum(item.emergency_cost_yuan for item in output_results))
    return StrategySummary(
        name,
        release_hours,
        results,
        plan_energy,
        final_energy,
        increase_energy,
        decrease_energy,
        emergency_energy,
        plan_cost,
        purchase_cost,
        emergency_cost,
        purchase_cost + emergency_cost,
    )


def validate_strategy(summary: StrategySummary, cfg: Config) -> None:
    if len(summary.results) != 365:
        raise RuntimeError(f"{summary.name}没有生成365天结果。")
    if abs(summary.results[0].soc_start_kwh - cfg.initial_soc_kwh) > 1.0e-4:
        raise RuntimeError(f"{summary.name}的1月1日初始SOC错误。")
    for index, item in enumerate(summary.results):
        if index > 0:
            previous = summary.results[index - 1]
            if abs(item.soc_start_kwh - previous.soc_end_kwh) > 1.0e-4:
                raise RuntimeError(f"{summary.name}在{item.day}的跨日SOC不连续。")
        if np.any(item.plan_grid_kwh < -1.0e-5):
            raise RuntimeError(f"{summary.name}存在负计划购电量。")
        if np.any(item.final_grid_kwh < -1.0e-5):
            raise RuntimeError(f"{summary.name}存在负调整购电量。")


def _copy_cell_style(source, target) -> None:
    target._style = copy(source._style)
    target.number_format = source.number_format
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.alignment = copy(source.alignment)
    target.protection = copy(source.protection)


def write_result_workbook(
    template: Path,
    output: Path,
    summary: StrategySummary,
    baseline: two.BaselineData,
    chronological_order: list[int],
    cfg: Config,
) -> None:
    if not template.exists():
        raise FileNotFoundError(f"找不到result3模板：{template}")
    if template.resolve() == output.resolve():
        raise ValueError("输出路径不能覆盖官方模板。")
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(f".{output.stem}.building{output.suffix}")
    if temporary.exists():
        temporary.unlink()
    shutil.copy2(template, temporary)

    try:
        workbook = load_workbook(temporary)
        plan_sheet = workbook["计划购电量"]
        adjust_sheet = workbook["调整购电量"]
        storage_sheet = workbook["充放电量"]
        emergency_sheet = workbook["紧急购电量"]
        output_results = [
            item for item in summary.results if item.day >= date(2025, 2, 1)
        ]
        if len(output_results) != 334:
            raise RuntimeError("result3应包含2025-02-01至2025-12-31共334天。")

        for sheet in (plan_sheet, adjust_sheet):
            template_dates = [
                sheet.cell(row, 1).value.date()
                if isinstance(sheet.cell(row, 1).value, datetime)
                else sheet.cell(row, 1).value
                for row in range(2, sheet.max_row + 1)
            ]
            if template_dates != [item.day for item in output_results]:
                raise ValueError(f"{sheet.title}日期与计算日期不一致。")

        for row, item in enumerate(output_results, start=2):
            plan_values = two._to_template_order(
                item.plan_grid_kwh, chronological_order
            )
            adjusted_values = two._to_template_order(
                item.final_grid_kwh, chronological_order
            )
            for column, value in enumerate(plan_values, start=2):
                plan_sheet.cell(row, column, round(float(value), 4))
            for column, value in enumerate(adjusted_values, start=2):
                adjust_sheet.cell(row, column, round(float(value), 4))
            plan_sheet.cell(row, 146, round(float(np.sum(item.plan_grid_kwh)), 4))
            plan_sheet.cell(row, 147, round(item.plan_cost_yuan, 4))
            adjust_sheet.cell(row, 146, round(float(np.sum(item.final_grid_kwh)), 4))
            adjust_sheet.cell(
                row, 147, round(item.settled_purchase_cost_yuan, 4)
            )

        storage_styles = [
            [copy(storage_sheet.cell(row, column)) for column in range(1, 7)]
            for row in (2, 3)
        ]
        if storage_sheet.max_row > 1:
            storage_sheet.delete_rows(2, storage_sheet.max_row - 1)
        periods = [
            "0:00-4:00",
            "4:00-8:00",
            "8:00-12:00",
            "12:00-16:00",
            "16:00-20:00",
            "20:00-24:00",
        ]
        storage_row = 2
        for item in output_results:
            for block, period in enumerate(periods):
                style_row = storage_styles[0 if block == 0 else 1]
                for column in range(1, 7):
                    _copy_cell_style(
                        style_row[column - 1], storage_sheet.cell(storage_row, column)
                    )
                start, stop = block * 24, (block + 1) * 24
                storage_sheet.cell(storage_row, 1, item.day if block == 0 else None)
                storage_sheet.cell(storage_row, 2, period)
                storage_sheet.cell(
                    storage_row,
                    3,
                    round(float(np.sum(item.charge_kwh[start:stop])), 4),
                )
                storage_sheet.cell(
                    storage_row,
                    4,
                    round(float(np.sum(item.discharge_kwh[start:stop])), 4),
                )
                if block == 0:
                    storage_sheet.cell(storage_row, 5, "0:00")
                    storage_sheet.cell(storage_row, 6, round(item.soc_start_kwh, 4))
                elif block == 1:
                    storage_sheet.cell(storage_row, 5, "24:00")
                    storage_sheet.cell(storage_row, 6, round(item.soc_end_kwh, 4))
                storage_row += 1

        emergency_styles = [
            [copy(emergency_sheet.cell(row, column)) for column in range(1, 4)]
            for row in (2, 3)
        ]
        if emergency_sheet.max_row > 1:
            emergency_sheet.delete_rows(2, emergency_sheet.max_row - 1)
        emergency_row = 2
        for item in output_results:
            groups = two.group_emergency(
                item.emergency_kwh, cfg.emergency_tolerance_kwh
            )
            for group_index, (period, amount) in enumerate(groups):
                style_row = emergency_styles[0 if group_index == 0 else 1]
                for column in range(1, 4):
                    _copy_cell_style(
                        style_row[column - 1],
                        emergency_sheet.cell(emergency_row, column),
                    )
                emergency_sheet.cell(
                    emergency_row, 1, item.day if group_index == 0 else None
                )
                emergency_sheet.cell(emergency_row, 2, period)
                emergency_sheet.cell(emergency_row, 3, round(amount, 4))
                emergency_row += 1

        workbook.save(temporary)
        try:
            os.replace(temporary, output)
        except PermissionError as exc:
            raise PermissionError(
                f"无法写入{output}，请关闭正在打开该文件的Excel窗口后重试。"
            ) from exc
    finally:
        if temporary.exists():
            temporary.unlink()


def write_daily_summary(path: Path, summary: StrategySummary) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, lineterminator="\n")
        writer.writerow(
            [
                "日期",
                "0点计划购电量(kWh)",
                "最终调整购电量(kWh)",
                "增加购电量(kWh)",
                "减少购电量(kWh)",
                "0点计划购电费(元)",
                "调整后购电结算费(元)",
                "调整相关费用变化(元)",
                "紧急购电量(kWh)",
                "紧急购电费(元)",
                "总购电费(元)",
                "0:00储电量(kWh)",
                "24:00储电量(kWh)",
            ]
        )
        for item in summary.results:
            increase = float(
                np.sum(np.maximum(0.0, item.final_grid_kwh - item.plan_grid_kwh))
            )
            decrease = float(
                np.sum(np.maximum(0.0, item.plan_grid_kwh - item.final_grid_kwh))
            )
            writer.writerow(
                [
                    item.day.isoformat(),
                    f"{np.sum(item.plan_grid_kwh):.6f}",
                    f"{np.sum(item.final_grid_kwh):.6f}",
                    f"{increase:.6f}",
                    f"{decrease:.6f}",
                    f"{item.plan_cost_yuan:.6f}",
                    f"{item.settled_purchase_cost_yuan:.6f}",
                    f"{item.adjustment_delta_cost_yuan:.6f}",
                    f"{np.sum(item.emergency_kwh):.6f}",
                    f"{item.emergency_cost_yuan:.6f}",
                    f"{item.total_cost_yuan:.6f}",
                    f"{item.soc_start_kwh:.6f}",
                    f"{item.soc_end_kwh:.6f}",
                ]
            )


def write_strategy_comparison(path: Path, summaries: list[StrategySummary]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    baseline_total = summaries[0].total_cost_yuan
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, lineterminator="\n")
        writer.writerow(
            [
                "策略",
                "使用预报时刻",
                "计划购电量(kWh)",
                "最终购电量(kWh)",
                "增加购电量(kWh)",
                "减少购电量(kWh)",
                "紧急购电量(kWh)",
                "计划购电费(元)",
                "调整后购电结算费(元)",
                "紧急购电费(元)",
                "总购电费(元)",
                "相对仅0点策略节省(元)",
                "相对仅0点策略节省率",
            ]
        )
        for item in summaries:
            saving = baseline_total - item.total_cost_yuan
            writer.writerow(
                [
                    item.name,
                    "/".join(f"{hour}:00" for hour in item.release_hours),
                    f"{item.plan_energy_kwh:.6f}",
                    f"{item.final_energy_kwh:.6f}",
                    f"{item.increase_energy_kwh:.6f}",
                    f"{item.decrease_energy_kwh:.6f}",
                    f"{item.emergency_energy_kwh:.6f}",
                    f"{item.plan_cost_yuan:.6f}",
                    f"{item.settled_purchase_cost_yuan:.6f}",
                    f"{item.emergency_cost_yuan:.6f}",
                    f"{item.total_cost_yuan:.6f}",
                    f"{saving:.6f}",
                    f"{saving / baseline_total:.8%}",
                ]
            )


def strategy_period_metrics(
    summary: StrategySummary,
    split: two.DatasetSplit,
    label: str,
) -> dict[str, float | int | str]:
    groups = split.cycle_labels([item.day for item in summary.results])
    selected_indices = set(groups[label])
    selected = [item for index, item in enumerate(summary.results) if index in selected_indices]
    if not selected:
        raise ValueError(f"{summary.name}的{label}为空。")
    plan_energy = float(sum(np.sum(item.plan_grid_kwh) for item in selected))
    final_energy = float(sum(np.sum(item.final_grid_kwh) for item in selected))
    increase_energy = float(
        sum(
            np.sum(np.maximum(0.0, item.final_grid_kwh - item.plan_grid_kwh))
            for item in selected
        )
    )
    decrease_energy = float(
        sum(
            np.sum(np.maximum(0.0, item.plan_grid_kwh - item.final_grid_kwh))
            for item in selected
        )
    )
    emergency_energy = float(sum(np.sum(item.emergency_kwh) for item in selected))
    plan_cost = float(sum(item.plan_cost_yuan for item in selected))
    purchase_cost = float(sum(item.settled_purchase_cost_yuan for item in selected))
    emergency_cost = float(sum(item.emergency_cost_yuan for item in selected))
    purchased = final_energy + emergency_energy
    return {
        "开始日期": selected[0].day.isoformat(),
        "结束日期": selected[-1].day.isoformat(),
        "天数": len(selected),
        "初始计划购电量(kWh)": plan_energy,
        "最终购电量(kWh)": final_energy,
        "增购量(kWh)": increase_energy,
        "退购量(kWh)": decrease_energy,
        "紧急购电量(kWh)": emergency_energy,
        "紧急购电占比": 0.0 if purchased <= 0.0 else emergency_energy / purchased,
        "初始计划购电费(元)": plan_cost,
        "调整后购电费(元)": purchase_cost,
        "紧急购电费(元)": emergency_cost,
        "总购电费(元)": purchase_cost + emergency_cost,
    }


def select_strategy_on_validation(
    summaries: list[StrategySummary], split: two.DatasetSplit
) -> StrategySummary:
    """只依据验证集成本选策略，测试集不参与模型选择。"""
    return min(
        summaries,
        key=lambda item: float(
            strategy_period_metrics(item, split, "验证集")["总购电费(元)"]
        ),
    )


def write_split_evaluation(
    path: Path,
    summaries: list[StrategySummary],
    selected: StrategySummary,
    split: two.DatasetSplit,
) -> None:
    rows: list[dict[str, object]] = []
    for label in ("训练集", "验证集", "测试集"):
        for summary in summaries:
            row: dict[str, object] = {
                "数据集": label,
                "策略": summary.name,
                "验证集选定策略": "是" if summary.name == selected.name else "否",
            }
            row.update(strategy_period_metrics(summary, split, label))
            rows.append(row)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0].keys()), lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)

    validation_cost = strategy_period_metrics(selected, split, "验证集")["总购电费(元)"]
    test_cost = strategy_period_metrics(selected, split, "测试集")["总购电费(元)"]
    print("\n问题三训练/验证/测试检验")
    print(f"  验证集选定策略：{selected.name}，验证集费用={validation_cost:.4f} 元")
    print(f"  冻结策略后的测试集费用：{test_cost:.4f} 元")
    print(f"数据集检验报告：{path.resolve()}")


def print_summary(
    summaries: list[StrategySummary], selected: StrategySummary, output: Path
) -> None:
    print("\n问题三不同预报时刻组合对比（2025-02-01至2025-12-31）")
    previous: StrategySummary | None = None
    for item in summaries:
        marginal = 0.0 if previous is None else previous.total_cost_yuan - item.total_cost_yuan
        print(
            f"  {item.name}：总费用={item.total_cost_yuan:.4f} 元，"
            f"紧急购电={item.emergency_energy_kwh:.4f} kWh，"
            f"相对上一策略边际节省={marginal:.4f} 元"
        )
        previous = item
    print(f"  最低费用策略：{selected.name}")

    specified = {date(2025, 3, 20), date(2025, 6, 21), date(2025, 9, 23), date(2025, 12, 21)}
    print(f"\n正式采用的{selected.name}策略指定日期紧急购电")
    for item in selected.results:
        if item.day not in specified:
            continue
        groups = two.group_emergency(item.emergency_kwh, 1.0e-5)
        if not groups:
            print(f"  {item.day.isoformat()}：无")
        else:
            print(f"  {item.day.isoformat()}：")
            for period, amount in groups:
                print(f"    {period}：{amount:.4f} kWh")
    print(f"\n正式结果：{output.resolve()}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="C题问题三MPC-LP滚动购电策略")
    parser.add_argument("--baseline", type=Path, default=two.DEFAULT_BASELINE)
    parser.add_argument("--load", type=Path, default=two.DEFAULT_LOAD)
    parser.add_argument("--pv", type=Path, default=two.DEFAULT_PV)
    parser.add_argument("--forecast", type=Path, default=DEFAULT_FORECAST)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--summary", type=Path, default=DEFAULT_SUMMARY)
    parser.add_argument("--comparison", type=Path, default=DEFAULT_COMPARISON)
    parser.add_argument("--split-report", type=Path, default=DEFAULT_SPLIT_REPORT)
    parser.add_argument("--train-end", type=date.fromisoformat, default=date(2025, 7, 31))
    parser.add_argument(
        "--validation-end", type=date.fromisoformat, default=date(2025, 9, 30)
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    cfg = Config()
    baseline, chronological_order = two.read_baseline(args.baseline)
    historical = two.read_historical_data(args.load, args.pv, chronological_order)
    forecasts = read_forecasts(args.forecast)
    split = two.DatasetSplit(args.train_end, args.validation_end)
    groups = split.cycle_labels(historical.dates)
    training_indices = set(groups["训练集"])
    print("正在构造四个发布时间的因果预测与历史误差安全裕度……")
    decision_inputs = prepare_decision_inputs(
        baseline,
        historical,
        forecasts,
        cfg,
        training_indices=training_indices,
    )

    strategy_specs = [
        ("S0_仅0点", (0,)),
        ("S1_增加6点", (0, 6)),
        ("S2_增加12点", (0, 6, 12)),
        ("S3_完整MPC", (0, 6, 12, 18)),
    ]
    summaries: list[StrategySummary] = []
    for name, release_hours in strategy_specs:
        print(f"\n开始计算{name}……")
        summary = simulate_strategy(
            name,
            release_hours,
            baseline,
            historical,
            decision_inputs,
            cfg,
        )
        validate_strategy(summary, cfg)
        summaries.append(summary)

    selected_strategy = select_strategy_on_validation(summaries, split)
    write_result_workbook(
        args.template,
        args.output,
        selected_strategy,
        baseline,
        chronological_order,
        cfg,
    )
    write_daily_summary(args.summary, selected_strategy)
    write_strategy_comparison(args.comparison, summaries)
    write_split_evaluation(args.split_report, summaries, selected_strategy, split)
    print_summary(summaries, selected_strategy, args.output)


if __name__ == "__main__":
    main()
