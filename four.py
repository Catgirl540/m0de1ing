"""2026 CUMCM C题问题四：傅里叶—小波实时电价预测下的 LP 与 MPC 购电策略。

电价预测链：
1. 对最近28天加权电价轮廓作六阶傅里叶拟合，提取日周期；
2. 对拟合残差作三级Haar小波软阈值重构，恢复局部尖峰与突变；
3. 使用前一日同一时刻的预测偏差滚动修正，并在线估计修正系数；
4. 对最终预测作非负截断。

问题4-2沿用问题二的48小时滚动线性规划；问题4-3沿用问题三的MPC-LP，
分别生成官方模板要求的result4-2.xlsx和result4-3.xlsx。

依赖：two.py、three.py、numpy、scipy、openpyxl（仅读取附件4）、Node.js。
运行：python four.py
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import subprocess
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
from openpyxl import load_workbook

import three
import two


ROOT = Path(__file__).resolve().parent
DEFAULT_PRICE = ROOT / "C题" / "附件" / "附件4.xlsx"
DEFAULT_FORECAST = ROOT / "C题" / "附件" / "csv" / "附件3.csv"
DEFAULT_TEMPLATE_42 = ROOT / "C题" / "附件" / "附件5" / "result4-2.xlsx"
DEFAULT_TEMPLATE_43 = ROOT / "C题" / "附件" / "附件5" / "result4-3.xlsx"
DEFAULT_OUTPUT_42 = ROOT / "C题" / "result4-2.xlsx"
DEFAULT_OUTPUT_43 = ROOT / "C题" / "result4-3.xlsx"
DEFAULT_SUMMARY_42 = ROOT / "C题" / "问题四_对应问题二逐日汇总.csv"
DEFAULT_SUMMARY_43 = ROOT / "C题" / "问题四_对应问题三逐日汇总.csv"
DEFAULT_COMPARISON = ROOT / "C题" / "问题四_MPC策略对比.csv"
DEFAULT_PRICE_EVALUATION = ROOT / "C题" / "问题四_电价预测评估.csv"
DEFAULT_MODEL = ROOT / "C题" / "问题四_动态电价数学模型.md"
DEFAULT_WRITER = ROOT / "four_workbooks.mjs"


@dataclass(frozen=True)
class Config(three.Config):
    price_history_days: int = 28
    price_half_life_days: float = 7.0
    price_same_weekday_multiplier: float = 2.0
    price_fourier_harmonics: int = 6
    price_wavelet_level: int = 3
    price_bias_days: int = 28
    minimum_bias_pairs: int = 7
    default_bias_weight: float = 0.60


@dataclass
class PriceData:
    dates: list[date]
    price: np.ndarray


@dataclass
class PriceForecast:
    value: np.ndarray
    periodic: np.ndarray
    wavelet_residual: np.ndarray
    previous_day_bias: np.ndarray
    bias_weight: float


@dataclass
class Daily42:
    day: date
    forecast_price: np.ndarray
    grid_kwh: np.ndarray
    charge_kwh: np.ndarray
    discharge_kwh: np.ndarray
    emergency_kwh: np.ndarray
    soc_start_kwh: float
    soc_end_kwh: float
    plan_cost_yuan: float
    emergency_cost_yuan: float

    @property
    def total_cost_yuan(self) -> float:
        return self.plan_cost_yuan + self.emergency_cost_yuan


def read_dynamic_prices(path: Path, chronological_order: list[int]) -> PriceData:
    if not path.exists():
        raise FileNotFoundError(f"找不到附件4：{path}")
    workbook = load_workbook(path, read_only=True, data_only=True)
    worksheet = workbook.active
    if worksheet.max_row != 366 or worksheet.max_column != 145:
        raise ValueError("附件4应包含1行表头、365天和每天144个实时电价。")

    dates: list[date] = []
    rows: list[list[float]] = []
    order = np.asarray(chronological_order, dtype=int)
    for row_number, row in enumerate(
        worksheet.iter_rows(min_row=2, max_row=366, values_only=True), start=2
    ):
        raw_date = row[0]
        if isinstance(raw_date, datetime):
            current_date = raw_date.date()
        elif isinstance(raw_date, date):
            current_date = raw_date
        else:
            current_date = datetime.fromisoformat(str(raw_date)).date()
        values = np.asarray(
            [two._float(value, f"附件4第{row_number}行电价") for value in row[1:]],
            dtype=float,
        )[order]
        dates.append(current_date)
        rows.append(values.tolist())
    workbook.close()

    expected = [date(2025, 1, 1) + timedelta(days=i) for i in range(365)]
    if dates != expected:
        raise ValueError("附件4日期应连续覆盖2025-01-01至2025-12-31。")
    matrix = np.asarray(rows, dtype=float)
    if not np.all(np.isfinite(matrix)) or np.any(matrix < 0):
        raise ValueError("附件4包含负值或非有限电价。")
    return PriceData(dates, matrix)


def _price_fourier_design(hours: np.ndarray, harmonics: int) -> np.ndarray:
    omega = 2.0 * np.pi / 24.0
    columns = [np.ones_like(hours)]
    for order in range(1, harmonics + 1):
        columns.extend(
            [np.cos(order * omega * hours), np.sin(order * omega * hours)]
        )
    return np.column_stack(columns)


def fit_price_fourier(
    hours: np.ndarray, profile: np.ndarray, harmonics: int
) -> np.ndarray:
    design = _price_fourier_design(hours, harmonics)
    coefficients, *_ = np.linalg.lstsq(design, profile, rcond=None)
    return design @ coefficients


def _price_history_profile(
    prices: PriceData,
    cutoff: int,
    target_date: date,
    baseline_price: np.ndarray,
    cfg: Config,
) -> np.ndarray:
    """只使用索引小于cutoff的已观测价格构造目标日轮廓。"""
    if cutoff <= 0:
        return baseline_price.copy()
    first = max(0, cutoff - cfg.price_history_days)
    indices = np.arange(first, cutoff, dtype=int)
    ages = cutoff - 1 - indices
    weights = np.exp(-math.log(2.0) * ages / cfg.price_half_life_days)
    same_weekday = np.asarray(
        [prices.dates[index].weekday() == target_date.weekday() for index in indices],
        dtype=float,
    )
    weights *= 1.0 + cfg.price_same_weekday_multiplier * same_weekday
    return np.average(prices.price[indices], axis=0, weights=weights)


def _causal_price_bases(
    prices: PriceData,
    baseline_price: np.ndarray,
    cfg: Config,
) -> tuple[list[np.ndarray], list[np.ndarray], list[np.ndarray]]:
    hours = np.arange(144, dtype=float) * cfg.dt_hours
    bases: list[np.ndarray] = []
    periodic_parts: list[np.ndarray] = []
    wavelet_parts: list[np.ndarray] = []
    for cutoff, current_date in enumerate(prices.dates):
        profile = _price_history_profile(
            prices, cutoff, current_date, baseline_price, cfg
        )
        periodic = fit_price_fourier(
            hours, profile, cfg.price_fourier_harmonics
        )
        wavelet = two.wavelet_correct_residual(
            profile, periodic, cfg.price_wavelet_level
        )
        bases.append(periodic + wavelet)
        periodic_parts.append(periodic)
        wavelet_parts.append(wavelet)
    return bases, periodic_parts, wavelet_parts


def _estimate_bias_weight(
    prices: PriceData,
    causal_bases: list[np.ndarray],
    cutoff: int,
    cfg: Config,
) -> float:
    """用此前日期残差的一阶关系估计前一日偏差修正系数。"""
    pair_end = cutoff - 1
    pair_start = max(1, pair_end - cfg.price_bias_days + 1)
    if pair_end - pair_start + 1 < cfg.minimum_bias_pairs:
        return cfg.default_bias_weight
    previous: list[np.ndarray] = []
    current: list[np.ndarray] = []
    for index in range(pair_start, pair_end + 1):
        previous.append(prices.price[index - 1] - causal_bases[index - 1])
        current.append(prices.price[index] - causal_bases[index])
    x = np.concatenate(previous)
    y = np.concatenate(current)
    denominator = float(np.dot(x, x))
    if denominator <= 1.0e-12:
        return cfg.default_bias_weight
    return float(np.clip(np.dot(x, y) / denominator, 0.0, 1.0))


def make_price_forecast(
    prices: PriceData,
    baseline_price: np.ndarray,
    causal_bases: list[np.ndarray],
    cutoff: int,
    target_date: date,
    cfg: Config,
) -> PriceForecast:
    hours = np.arange(144, dtype=float) * cfg.dt_hours
    profile = _price_history_profile(
        prices, cutoff, target_date, baseline_price, cfg
    )
    periodic = fit_price_fourier(hours, profile, cfg.price_fourier_harmonics)
    wavelet = two.wavelet_correct_residual(
        profile, periodic, cfg.price_wavelet_level
    )
    latest_observed = cutoff - 1
    if latest_observed >= 0:
        bias = prices.price[latest_observed] - causal_bases[latest_observed]
        weight = _estimate_bias_weight(prices, causal_bases, cutoff, cfg)
    else:
        bias = np.zeros(144, dtype=float)
        weight = 0.0
    value = np.maximum(0.0, periodic + wavelet + weight * bias)
    return PriceForecast(value, periodic, wavelet, bias, weight)


def prepare_daily_price_forecasts(
    prices: PriceData,
    baseline_price: np.ndarray,
    cfg: Config,
) -> tuple[list[PriceForecast], list[np.ndarray]]:
    causal_bases, _, _ = _causal_price_bases(prices, baseline_price, cfg)
    current: list[PriceForecast] = []
    following: list[np.ndarray] = []
    for cutoff, current_date in enumerate(prices.dates):
        current.append(
            make_price_forecast(
                prices,
                baseline_price,
                causal_bases,
                cutoff,
                current_date,
                cfg,
            )
        )
        following.append(
            make_price_forecast(
                prices,
                baseline_price,
                causal_bases,
                cutoff,
                current_date + timedelta(days=1),
                cfg,
            ).value
        )
    return current, following


def simulate_42(
    baseline: two.BaselineData,
    historical: two.HistoricalData,
    prices: PriceData,
    current_prices: list[PriceForecast],
    following_prices: list[np.ndarray],
    cfg: Config,
) -> list[Daily42]:
    hours = np.arange(144, dtype=float) * cfg.dt_hours
    error_history: list[np.ndarray] = []
    results: list[Daily42] = []
    soc = cfg.initial_soc_kwh

    for day_index, current_date in enumerate(historical.dates):
        current_forecast = two.make_forecast(
            historical, baseline, day_index, current_date, hours, cfg
        )
        next_forecast = two.make_forecast(
            historical,
            baseline,
            day_index,
            current_date + timedelta(days=1),
            hours,
            cfg,
        )
        reserve = two.risk_reserve(error_history, cfg)
        price_horizon = np.concatenate(
            [current_prices[day_index].value, following_prices[day_index]]
        )
        solution = two.solve_rolling_lp(
            price_horizon,
            np.concatenate(
                [
                    current_forecast.load_kw + reserve / cfg.dt_hours,
                    next_forecast.load_kw + reserve / cfg.dt_hours,
                ]
            ),
            np.concatenate([current_forecast.pv_kw, next_forecast.pv_kw]),
            soc,
            cfg,
        )
        grid = solution.grid_kwh[:144]
        charge = solution.charge_kwh[:144]
        discharge = solution.discharge_kwh[:144]
        emergency, _ = two.replay_actual_day(
            historical.load_kw[day_index],
            historical.pv_kw[day_index],
            grid,
            charge,
            discharge,
            cfg,
        )
        soc_start = soc
        soc = float(solution.soc_kwh[144])
        actual_price = prices.price[day_index]
        plan_cost = float(np.dot(actual_price, grid))
        emergency_cost = float(
            np.dot(cfg.emergency_price_multiplier * actual_price, emergency)
        )
        results.append(
            Daily42(
                current_date,
                current_prices[day_index].value,
                grid,
                charge,
                discharge,
                emergency,
                soc_start,
                soc,
                plan_cost,
                emergency_cost,
            )
        )
        actual_net = (
            historical.load_kw[day_index] - historical.pv_kw[day_index]
        ) * cfg.dt_hours
        forecast_net = (
            current_forecast.load_kw - current_forecast.pv_kw
        ) * cfg.dt_hours
        error_history.append(actual_net - forecast_net)
        if (day_index + 1) % 60 == 0 or day_index == 364:
            print(f"问题4-2已完成{current_date}，SOC={soc:.2f} kWh")
    return results


def price_horizon(
    day_index: int,
    release_hour: int,
    current_prices: list[PriceForecast],
    following_prices: list[np.ndarray],
) -> np.ndarray:
    start = release_hour * 6
    if start == 0:
        return current_prices[day_index].value.copy()
    return np.concatenate(
        [current_prices[day_index].value[start:], following_prices[day_index][:start]]
    )


def simulate_43_strategy(
    name: str,
    release_hours: tuple[int, ...],
    historical: two.HistoricalData,
    prices: PriceData,
    decision_inputs: dict[tuple[int, int], three.DecisionInput],
    current_prices: list[PriceForecast],
    following_prices: list[np.ndarray],
    cfg: Config,
) -> three.StrategySummary:
    results: list[three.DailyResult] = []
    soc = cfg.initial_soc_kwh
    for day_index, current_date in enumerate(historical.dates):
        soc_start = soc
        decision_zero = decision_inputs[(day_index, 0)]
        plan_solution = two.solve_rolling_lp(
            price_horizon(
                day_index, 0, current_prices, following_prices
            ),
            decision_zero.load_kw + decision_zero.reserve_kwh / cfg.dt_hours,
            decision_zero.pv_kw,
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
                adjusted = three.solve_adjustment_lp(
                    price_horizon(
                        day_index,
                        release_hour,
                        current_prices,
                        following_prices,
                    ),
                    decision.load_kw + decision.reserve_kwh / cfg.dt_hours,
                    decision.pv_kw,
                    soc,
                    plan_grid[start:],
                    144 - start,
                    cfg,
                )
                final_grid[start:] = adjusted.grid_kwh[: 144 - start]
                final_charge[start:] = adjusted.charge_kwh[: 144 - start]
                final_discharge[start:] = adjusted.discharge_kwh[: 144 - start]
            stop = min(start + 36, 144)
            soc = three.execute_soc(
                soc,
                final_charge[start:stop],
                final_discharge[start:stop],
                cfg,
            )

        emergency, curtailment = two.replay_actual_day(
            historical.load_kw[day_index],
            historical.pv_kw[day_index],
            final_grid,
            final_charge,
            final_discharge,
            cfg,
        )
        actual_price = prices.price[day_index]
        plan_cost = float(np.dot(actual_price, plan_grid))
        settled_cost = three.settled_purchase_cost(
            actual_price, plan_grid, final_grid, cfg
        )
        emergency_cost = float(
            np.dot(cfg.emergency_price_multiplier * actual_price, emergency)
        )
        results.append(
            three.DailyResult(
                current_date,
                plan_grid,
                final_grid,
                final_charge,
                final_discharge,
                emergency,
                curtailment,
                soc_start,
                soc,
                plan_cost,
                settled_cost,
                emergency_cost,
            )
        )
        if (day_index + 1) % 120 == 0 or day_index == 364:
            print(f"{name}已完成{current_date}，SOC={soc:.2f} kWh")

    selected = [item for item in results if item.day >= date(2025, 2, 1)]
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
    settled_cost = float(sum(item.settled_purchase_cost_yuan for item in selected))
    emergency_cost = float(sum(item.emergency_cost_yuan for item in selected))
    return three.StrategySummary(
        name,
        release_hours,
        results,
        plan_energy,
        final_energy,
        increase_energy,
        decrease_energy,
        emergency_energy,
        plan_cost,
        settled_cost,
        emergency_cost,
        settled_cost + emergency_cost,
    )


def validate_42(results: list[Daily42], cfg: Config) -> None:
    if len(results) != 365:
        raise RuntimeError("问题4-2没有生成365天结果。")
    for index, item in enumerate(results):
        if index and abs(item.soc_start_kwh - results[index - 1].soc_end_kwh) > 1e-4:
            raise RuntimeError(f"问题4-2在{item.day}的跨日SOC不连续。")
        if np.any(item.grid_kwh < -1e-6) or np.any(item.emergency_kwh < -1e-6):
            raise RuntimeError(f"问题4-2在{item.day}出现负购电量。")
        if not cfg.soc_min_kwh - 1e-4 <= item.soc_end_kwh <= cfg.soc_max_kwh + 1e-4:
            raise RuntimeError(f"问题4-2在{item.day}的SOC越界。")


def _storage_rows(results: list, q43: bool = False) -> list[list[object]]:
    rows: list[list[object]] = []
    periods = [
        "0:00-4:00",
        "4:00-8:00",
        "8:00-12:00",
        "12:00-16:00",
        "16:00-20:00",
        "20:00-24:00",
    ]
    for item in results:
        charge = item.charge_kwh
        discharge = item.discharge_kwh
        for block, period in enumerate(periods):
            start, stop = block * 24, (block + 1) * 24
            rows.append(
                [
                    item.day.isoformat() if block == 0 else None,
                    period,
                    round(float(np.sum(charge[start:stop])), 4),
                    round(float(np.sum(discharge[start:stop])), 4),
                    "0:00" if block == 0 else ("24:00" if block == 1 else None),
                    round(
                        float(item.soc_start_kwh if block == 0 else item.soc_end_kwh),
                        4,
                    )
                    if block in (0, 1)
                    else None,
                ]
            )
    return rows


def _emergency_rows(results: list, tolerance: float) -> list[list[object]]:
    rows: list[list[object]] = []
    for item in results:
        groups = two.group_emergency(item.emergency_kwh, tolerance)
        for group_index, (period, amount) in enumerate(groups):
            rows.append(
                [
                    item.day.isoformat() if group_index == 0 else None,
                    period,
                    round(amount, 4),
                ]
            )
    return rows or [["无紧急购电", None, 0.0]]


def make_workbook_payload(
    results42: list[Daily42],
    selected43: three.StrategySummary,
    chronological_order: list[int],
    cfg: Config,
) -> dict[str, object]:
    selected42 = [item for item in results42 if item.day >= date(2025, 2, 1)]
    selected43_results = [
        item for item in selected43.results if item.day >= date(2025, 2, 1)
    ]
    plan42 = [
        [
            round(float(value), 4)
            for value in two._to_template_order(
                item.grid_kwh, chronological_order
            )
        ]
        + [round(float(np.sum(item.grid_kwh)), 4), round(item.plan_cost_yuan, 4)]
        for item in selected42
    ]
    plan43 = [
        [
            round(float(value), 4)
            for value in two._to_template_order(
                item.plan_grid_kwh, chronological_order
            )
        ]
        + [
            round(float(np.sum(item.plan_grid_kwh)), 4),
            round(item.plan_cost_yuan, 4),
        ]
        for item in selected43_results
    ]
    adjust43 = [
        [
            round(float(value), 4)
            for value in two._to_template_order(
                item.final_grid_kwh, chronological_order
            )
        ]
        + [
            round(float(np.sum(item.final_grid_kwh)), 4),
            round(item.settled_purchase_cost_yuan, 4),
        ]
        for item in selected43_results
    ]
    return {
        "q42": {
            "plan": plan42,
            "storage": _storage_rows(selected42),
            "emergency": _emergency_rows(
                selected42, cfg.emergency_tolerance_kwh
            ),
        },
        "q43": {
            "plan": plan43,
            "adjust": adjust43,
            "storage": _storage_rows(selected43_results, q43=True),
            "emergency": _emergency_rows(
                selected43_results, cfg.emergency_tolerance_kwh
            ),
        },
    }


def write_csv_outputs(
    results42: list[Daily42],
    summaries43: list[three.StrategySummary],
    selected43: three.StrategySummary,
    prices: PriceData,
    price_forecasts: list[PriceForecast],
    args: argparse.Namespace,
) -> None:
    with args.summary42.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, lineterminator="\n")
        writer.writerow(
            [
                "日期",
                "计划购电量(kWh)",
                "计划购电费(元)",
                "紧急购电量(kWh)",
                "紧急购电费(元)",
                "总购电费(元)",
                "0:00储电量(kWh)",
                "24:00储电量(kWh)",
                "电价预测MAE(元/kWh)",
            ]
        )
        for index, item in enumerate(results42):
            writer.writerow(
                [
                    item.day.isoformat(),
                    f"{np.sum(item.grid_kwh):.6f}",
                    f"{item.plan_cost_yuan:.6f}",
                    f"{np.sum(item.emergency_kwh):.6f}",
                    f"{item.emergency_cost_yuan:.6f}",
                    f"{item.total_cost_yuan:.6f}",
                    f"{item.soc_start_kwh:.6f}",
                    f"{item.soc_end_kwh:.6f}",
                    f"{np.mean(np.abs(prices.price[index] - item.forecast_price)):.8f}",
                ]
            )
    three.write_daily_summary(args.summary43, selected43)
    three.write_strategy_comparison(args.comparison, summaries43)
    with args.price_evaluation.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, lineterminator="\n")
        writer.writerow(
            [
                "日期",
                "MAE(元/kWh)",
                "RMSE(元/kWh)",
                "MAPE",
                "前日偏差权重",
                "实际均价(元/kWh)",
                "预测均价(元/kWh)",
            ]
        )
        for index, forecast in enumerate(price_forecasts):
            actual = prices.price[index]
            error = actual - forecast.value
            writer.writerow(
                [
                    prices.dates[index].isoformat(),
                    f"{np.mean(np.abs(error)):.8f}",
                    f"{np.sqrt(np.mean(error**2)):.8f}",
                    f"{np.mean(np.abs(error) / np.maximum(actual, 0.05)):.8%}",
                    f"{forecast.bias_weight:.8f}",
                    f"{np.mean(actual):.8f}",
                    f"{np.mean(forecast.value):.8f}",
                ]
            )


def write_workbooks(
    payload: dict[str, object], args: argparse.Namespace
) -> None:
    payload_path = ROOT / "C题" / ".four_workbook_payload.json"
    payload_path.write_text(
        json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    command = [
        "node",
        str(args.writer),
        str(payload_path),
        str(args.template42),
        str(args.output42),
        str(args.template43),
        str(args.output43),
    ]
    try:
        subprocess.run(command, cwd=ROOT, check=True)
    finally:
        payload_path.unlink(missing_ok=True)


def print_summary(
    results42: list[Daily42],
    summaries43: list[three.StrategySummary],
    selected43: three.StrategySummary,
    prices: PriceData,
    price_forecasts: list[PriceForecast],
    args: argparse.Namespace,
) -> None:
    output42 = [item for item in results42 if item.day >= date(2025, 2, 1)]
    forecast_matrix = np.asarray(
        [item.value for item in price_forecasts[31:]], dtype=float
    )
    actual_matrix = prices.price[31:]
    mae = float(np.mean(np.abs(actual_matrix - forecast_matrix)))
    rmse = float(np.sqrt(np.mean((actual_matrix - forecast_matrix) ** 2)))
    total42 = float(sum(item.total_cost_yuan for item in output42))
    emergency42 = float(sum(np.sum(item.emergency_kwh) for item in output42))
    print("\n问题四动态电价预测（2025-02-01至2025-12-31）")
    print(f"  MAE={mae:.6f} 元/kWh，RMSE={rmse:.6f} 元/kWh")
    print(f"  问题4-2总费用={total42:.4f} 元，紧急购电={emergency42:.4f} kWh")
    print("\n问题4-3策略比较")
    for item in summaries43:
        print(
            f"  {item.name}：总费用={item.total_cost_yuan:.4f} 元，"
            f"紧急购电={item.emergency_energy_kwh:.4f} kWh"
        )
    print(f"  正式采用：{selected43.name}")
    print(f"\n结果文件：{args.output42.resolve()}")
    print(f"结果文件：{args.output43.resolve()}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="C题问题四傅里叶—小波动态电价预测与LP/MPC购电策略"
    )
    parser.add_argument("--baseline", type=Path, default=two.DEFAULT_BASELINE)
    parser.add_argument("--load", type=Path, default=two.DEFAULT_LOAD)
    parser.add_argument("--pv", type=Path, default=two.DEFAULT_PV)
    parser.add_argument("--forecast", type=Path, default=DEFAULT_FORECAST)
    parser.add_argument("--price", type=Path, default=DEFAULT_PRICE)
    parser.add_argument("--template42", type=Path, default=DEFAULT_TEMPLATE_42)
    parser.add_argument("--template43", type=Path, default=DEFAULT_TEMPLATE_43)
    parser.add_argument("--output42", type=Path, default=DEFAULT_OUTPUT_42)
    parser.add_argument("--output43", type=Path, default=DEFAULT_OUTPUT_43)
    parser.add_argument("--summary42", type=Path, default=DEFAULT_SUMMARY_42)
    parser.add_argument("--summary43", type=Path, default=DEFAULT_SUMMARY_43)
    parser.add_argument("--comparison", type=Path, default=DEFAULT_COMPARISON)
    parser.add_argument(
        "--price-evaluation", type=Path, default=DEFAULT_PRICE_EVALUATION
    )
    parser.add_argument("--writer", type=Path, default=DEFAULT_WRITER)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    cfg = Config()
    baseline, chronological_order = two.read_baseline(args.baseline)
    historical = two.read_historical_data(args.load, args.pv, chronological_order)
    prices = read_dynamic_prices(args.price, chronological_order)
    if prices.dates != historical.dates:
        raise ValueError("附件2与附件4的日期不一致。")
    forecasts = three.read_forecasts(args.forecast)

    print("正在构造傅里叶—小波—前日偏差动态电价预测……")
    current_prices, following_prices = prepare_daily_price_forecasts(
        prices, baseline.price, cfg
    )
    print("正在计算问题4-2……")
    results42 = simulate_42(
        baseline,
        historical,
        prices,
        current_prices,
        following_prices,
        cfg,
    )
    validate_42(results42, cfg)

    print("正在构造问题4-3负荷、光伏决策输入……")
    decision_inputs = three.prepare_decision_inputs(
        baseline, historical, forecasts, cfg
    )
    strategy_specs = [
        ("S0_仅0点", (0,)),
        ("S1_增加6点", (0, 6)),
        ("S2_增加12点", (0, 6, 12)),
        ("S3_完整MPC", (0, 6, 12, 18)),
    ]
    summaries43: list[three.StrategySummary] = []
    for name, release_hours in strategy_specs:
        print(f"正在计算问题4-3策略{name}……")
        summary = simulate_43_strategy(
            name,
            release_hours,
            historical,
            prices,
            decision_inputs,
            current_prices,
            following_prices,
            cfg,
        )
        three.validate_strategy(summary, cfg)
        summaries43.append(summary)
    selected43 = min(summaries43, key=lambda item: item.total_cost_yuan)

    for path in (
        args.summary42,
        args.summary43,
        args.comparison,
        args.price_evaluation,
    ):
        path.parent.mkdir(parents=True, exist_ok=True)
    write_csv_outputs(
        results42,
        summaries43,
        selected43,
        prices,
        current_prices,
        args,
    )
    payload = make_workbook_payload(
        results42, selected43, chronological_order, cfg
    )
    write_workbooks(payload, args)
    print_summary(
        results42,
        summaries43,
        selected43,
        prices,
        current_prices,
        args,
    )


if __name__ == "__main__":
    main()
