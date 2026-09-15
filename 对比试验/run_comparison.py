"""当前正式模型的基准对比、模块消融与论文图表生成。

运行方式（项目根目录）：python 对比试验/run_comparison.py

实验严格沿用 two.py/four.py 的“训训训验训测”划分及因果可用样本约束；
所有模型贡献率和论文结论均以测试集为准。
"""

from __future__ import annotations

import csv
import math
import sys
from dataclasses import replace
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from scipy.stats import wilcoxon


ROOT = Path(__file__).resolve().parents[1]
OUT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import four  # noqa: E402
import two  # noqa: E402


SPLIT_ORDER = ("训练集", "验证集", "测试集")


def configure_plot() -> None:
    plt.rcParams.update(
        {
            "font.sans-serif": ["Microsoft YaHei", "SimHei", "Arial Unicode MS", "DejaVu Sans"],
            "axes.unicode_minus": False,
            "figure.dpi": 130,
            "savefig.dpi": 320,
            "axes.spines.top": False,
            "axes.spines.right": False,
            "axes.grid": True,
            "grid.alpha": 0.22,
            "grid.linestyle": "--",
        }
    )


def metrics(actual: np.ndarray, predicted: np.ndarray, *, mape_floor: float, mask=None) -> dict[str, float]:
    y = np.asarray(actual, dtype=float)
    yhat = np.asarray(predicted, dtype=float)
    if mask is not None:
        keep = np.asarray(mask, dtype=bool)
        y, yhat = y[keep], yhat[keep]
    error = yhat - y
    denominator = np.maximum(np.abs(y), mape_floor)
    ss_res = float(np.sum(error**2))
    ss_tot = float(np.sum((y - np.mean(y)) ** 2))
    return {
        "样本数": int(y.size),
        "MAE": float(np.mean(np.abs(error))),
        "RMSE": float(np.sqrt(np.mean(error**2))),
        "MAPE(%)": float(100.0 * np.mean(np.abs(error) / denominator)),
        "WMAPE(%)": float(100.0 * np.sum(np.abs(error)) / max(np.sum(np.abs(y)), 1e-12)),
        "R2": float("nan") if ss_tot <= 0 else 1.0 - ss_res / ss_tot,
    }


def write_rows(path: Path, rows: list[dict[str, object]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def load_and_pv_predictions(baseline: two.BaselineData, historical: two.HistoricalData, training: set[int]):
    cfg = two.Config()
    cfg_static = replace(cfg, use_bayesian_load_schedule=False)
    hours = np.arange(144, dtype=float) * cfg.dt_hours
    n = len(historical.dates)
    load_names = [
        "典型日静态基准",
        "前一日持续性基准",
        "仅历史加权轮廓",
        "去除周期回归",
        "去除小波修正",
        "去除贝叶斯更新",
        "完整模型",
    ]
    pv_names = ["典型日静态基准", "前一日持续性基准", "历史加权轮廓", "截断正弦完整模型"]
    load_pred = {name: np.zeros_like(historical.load_kw) for name in load_names}
    pv_pred = {name: np.zeros_like(historical.pv_kw) for name in pv_names}

    for i, day in enumerate(historical.dates):
        daily_cfg = two.bayesian_load_config(cfg, day)
        history_profile = two._weighted_history_profile(
            historical.load_kw, historical.dates, i, day, baseline.load_kw, daily_cfg, training
        )
        dynamic = two.fit_dynamic_load_profile(
            historical, i, day, baseline.load_kw, daily_cfg, training
        )
        blend = float(np.clip(daily_cfg.dynamic_load_blend, 0.0, 1.0))
        blended = (1.0 - blend) * history_profile + blend * dynamic
        fourier = two.fit_load_fourier(hours, blended)
        full = np.maximum(
            0.0, fourier + two.wavelet_correct_residual(blended, fourier, daily_cfg.wavelet_level)
        )

        static_history = two._weighted_history_profile(
            historical.load_kw, historical.dates, i, day, baseline.load_kw, cfg_static, training
        )
        static_dynamic = two.fit_dynamic_load_profile(
            historical, i, day, baseline.load_kw, cfg_static, training
        )
        static_blend = (1.0 - cfg_static.dynamic_load_blend) * static_history + cfg_static.dynamic_load_blend * static_dynamic
        static_fourier = two.fit_load_fourier(hours, static_blend)
        no_bayes = np.maximum(
            0.0,
            static_fourier
            + two.wavelet_correct_residual(static_blend, static_fourier, cfg_static.wavelet_level),
        )
        history_fourier = two.fit_load_fourier(hours, history_profile)
        no_calendar = np.maximum(
            0.0,
            history_fourier
            + two.wavelet_correct_residual(history_profile, history_fourier, daily_cfg.wavelet_level),
        )

        load_pred["典型日静态基准"][i] = baseline.load_kw
        load_pred["前一日持续性基准"][i] = baseline.load_kw if i == 0 else historical.load_kw[i - 1]
        load_pred["仅历史加权轮廓"][i] = history_profile
        load_pred["去除周期回归"][i] = no_calendar
        load_pred["去除小波修正"][i] = fourier
        load_pred["去除贝叶斯更新"][i] = no_bayes
        load_pred["完整模型"][i] = full

        pv_profile = two._weighted_history_profile(
            historical.pv_kw, historical.dates, i, day, baseline.pv_kw, cfg, training
        )
        sine, *_ = two.fit_pv_truncated_sine(hours, pv_profile)
        pv_pred["典型日静态基准"][i] = baseline.pv_kw
        pv_pred["前一日持续性基准"][i] = baseline.pv_kw if i == 0 else historical.pv_kw[i - 1]
        pv_pred["历史加权轮廓"][i] = pv_profile
        pv_pred["截断正弦完整模型"][i] = sine
        if (i + 1) % 60 == 0:
            print(f"  功率预测消融：{i + 1}/{n} 天")
    return load_pred, pv_pred


def price_predictions(
    baseline: two.BaselineData,
    prices: four.PriceData,
    training: set[int],
) -> dict[str, np.ndarray]:
    cfg = four.Config()
    fixed = replace(cfg, use_bayesian_price_schedule=False)
    hours = np.arange(144, dtype=float) * cfg.dt_hours
    n = len(prices.dates)
    names = [
        "典型日静态基准",
        "前一日持续性基准",
        "历史加权轮廓",
        "仅傅里叶周期项",
        "傅里叶+小波",
        "增加前日偏差修正",
        "完整模型（含贝叶斯更新）",
    ]
    predicted = {name: np.zeros_like(prices.price) for name in names}

    fixed_bases, _, _ = four._causal_price_bases(
        prices, baseline.price, fixed, training_indices=training
    )
    fixed_forecasts = [
        four.make_price_forecast(
            prices, baseline.price, fixed_bases, i, day, fixed, training_indices=training
        ).value
        for i, day in enumerate(prices.dates)
    ]
    full_forecasts, _ = four.prepare_daily_price_forecasts(
        prices, baseline.price, cfg, training_indices=training
    )

    for i, day in enumerate(prices.dates):
        profile = four._price_history_profile(prices, i, day, baseline.price, fixed, training)
        periodic = four.fit_price_fourier(hours, profile, fixed.price_fourier_harmonics)
        wavelet = two.wavelet_correct_residual(profile, periodic, fixed.price_wavelet_level)
        predicted["典型日静态基准"][i] = baseline.price
        predicted["前一日持续性基准"][i] = baseline.price if i == 0 else prices.price[i - 1]
        predicted["历史加权轮廓"][i] = np.maximum(0.0, profile)
        predicted["仅傅里叶周期项"][i] = np.maximum(0.0, periodic)
        predicted["傅里叶+小波"][i] = np.maximum(0.0, periodic + wavelet)
        predicted["增加前日偏差修正"][i] = fixed_forecasts[i]
        predicted["完整模型（含贝叶斯更新）"][i] = full_forecasts[i].value
    return predicted


def evaluate_predictions(
    target: str,
    actual: np.ndarray,
    predictions: dict[str, np.ndarray],
    groups: dict[str, list[int]],
    mape_floor: float,
    daylight_only: bool = False,
) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for split_name in SPLIT_ORDER:
        idx = groups[split_name]
        y = actual[idx]
        for model, forecast in predictions.items():
            yhat = forecast[idx]
            mask = y > 1.0 if daylight_only else None
            row: dict[str, object] = {"对象": target, "数据集": split_name, "模型": model}
            row.update(metrics(y, yhat, mape_floor=mape_floor, mask=mask))
            rows.append(row)
    return rows


def read_strategy_rows(path: Path, question: str) -> list[dict[str, object]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        source = list(csv.DictReader(file))
    rows: list[dict[str, object]] = []
    for item in source:
        if item.get("数据集") != "测试集":
            continue
        raw_name = item.get("策略") or item.get("模型或策略", "")
        if question == "问题四":
            if not raw_name.startswith("问题4-3_"):
                continue
            raw_name = raw_name.removeprefix("问题4-3_")
        cost = float(item["总购电费(元)"])
        emergency = float(item["紧急购电量(kWh)"])
        share = float(item["紧急购电占比"])
        rows.append(
            {
                "问题": question,
                "策略": raw_name,
                "测试集总费用(元)": cost,
                "测试集紧急购电量(kWh)": emergency,
                "测试集紧急购电占比(%)": 100.0 * share,
            }
        )
    baseline_cost = float(rows[0]["测试集总费用(元)"])
    baseline_emergency = float(rows[0]["测试集紧急购电量(kWh)"])
    for row in rows:
        row["相对仅0点节省率(%)"] = 100.0 * (baseline_cost - float(row["测试集总费用(元)"])) / baseline_cost
        row["紧急购电削减率(%)"] = 100.0 * (baseline_emergency - float(row["测试集紧急购电量(kWh)"])) / baseline_emergency
    return rows


def release_marginal_rows(strategy_rows: list[dict[str, object]]) -> list[dict[str, object]]:
    q3 = [row for row in strategy_rows if row["问题"] == "问题三"]
    output: list[dict[str, object]] = []
    for before, after, hour in zip(q3[:-1], q3[1:], ("6:00", "12:00", "18:00")):
        saved = float(before["测试集总费用(元)"]) - float(after["测试集总费用(元)"])
        emergency_reduction = float(before["测试集紧急购电量(kWh)"]) - float(after["测试集紧急购电量(kWh)"])
        output.append(
            {
                "新增预报时刻": hour,
                "原策略": before["策略"],
                "新策略": after["策略"],
                "测试集边际费用节省(元)": saved,
                "边际费用节省率(%)": 100.0 * saved / float(before["测试集总费用(元)"]),
                "紧急购电减少量(kWh)": emergency_reduction,
                "是否建议引入": "是" if saved > 0 and emergency_reduction >= 0 else "否",
            }
        )
    return output


def test_rows(rows: list[dict[str, object]], target: str) -> list[dict[str, object]]:
    return [row for row in rows if row["对象"] == target and row["数据集"] == "测试集"]


def plot_metric_comparison(rows: list[dict[str, object]], target: str, filename: str, unit: str) -> None:
    selected = test_rows(rows, target)
    labels = [str(row["模型"]) for row in selected]
    rmse = np.asarray([float(row["RMSE"]) for row in selected])
    r2 = np.asarray([float(row["R2"]) for row in selected])
    colors = ["#B8C4CE"] * len(labels)
    colors[-1] = "#3E6B89"
    fig, ax = plt.subplots(figsize=(10.8, 5.8))
    x = np.arange(len(labels))
    bars = ax.bar(x, rmse, color=colors, width=0.68, label="RMSE")
    ax.set_ylabel(f"RMSE（{unit}）")
    ax.set_xticks(x, labels, rotation=18, ha="right")
    ax.bar_label(bars, labels=[f"{v:.2f}" for v in rmse], padding=3, fontsize=8)
    ax2 = ax.twinx()
    ax2.plot(x, r2, color="#C26B4A", marker="o", linewidth=2, label="$R^2$")
    ax2.set_ylabel("$R^2$")
    lower = min(0.0, float(np.nanmin(r2)) - 0.06)
    ax2.set_ylim(lower, 1.04)
    ax.set_title(f"{target}：测试集基准与消融实验")
    handles1, labels1 = ax.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(handles1 + handles2, labels1 + labels2, loc="upper left", frameon=False)
    fig.tight_layout()
    fig.savefig(OUT / filename, bbox_inches="tight")
    plt.close(fig)


def plot_mpc(rows: list[dict[str, object]]) -> None:
    fig, axes = plt.subplots(1, 2, figsize=(12, 5.2), sharey=False)
    for ax, question in zip(axes, ("问题三", "问题四")):
        part = [row for row in rows if row["问题"] == question]
        labels = [str(row["策略"]) for row in part]
        costs = np.asarray([float(row["测试集总费用(元)"]) / 1e6 for row in part])
        shares = np.asarray([float(row["测试集紧急购电占比(%)"]) for row in part])
        x = np.arange(len(labels))
        colors = ["#B8C4CE", "#8FA9B8", "#3E6B89", "#6F8795"]
        bars = ax.bar(x, costs, color=colors, width=0.65, label="总费用")
        ax.set_xticks(x, labels, rotation=15, ha="right")
        ax.set_ylabel("测试集总费用（百万元）")
        ax.bar_label(bars, labels=[f"{v:.3f}" for v in costs], padding=3, fontsize=8)
        twin = ax.twinx()
        twin.plot(x, shares, color="#C26B4A", marker="o", linewidth=2, label="紧急购电占比")
        twin.set_ylabel("紧急购电占比（%）")
        ax.set_title(f"{question}滚动策略对比")
    fig.tight_layout()
    fig.savefig(OUT / "图4_MPC滚动策略经济性对比.png", bbox_inches="tight")
    plt.close(fig)


def plot_example(
    dates,
    test_indices: list[int],
    actual_load: np.ndarray,
    predicted_load: np.ndarray,
    actual_price: np.ndarray,
    predicted_price: np.ndarray,
) -> str:
    daily_rmse = np.sqrt(np.mean((actual_load[test_indices] - predicted_load[test_indices]) ** 2, axis=1))
    chosen_position = int(np.argsort(daily_rmse)[len(daily_rmse) // 2])
    index = test_indices[chosen_position]
    hours = np.arange(144) / 6.0
    fig, axes = plt.subplots(2, 1, figsize=(10.8, 7.0), sharex=True)
    axes[0].plot(hours, actual_load[index], color="#263238", linewidth=1.8, label="实际值")
    axes[0].plot(hours, predicted_load[index], color="#3E6B89", linewidth=1.8, linestyle="--", label="预测值")
    axes[0].set_ylabel("负荷（kW）")
    axes[0].legend(frameon=False, ncol=2)
    axes[1].plot(hours, actual_price[index], color="#263238", linewidth=1.8, label="实际值")
    axes[1].plot(hours, predicted_price[index], color="#C26B4A", linewidth=1.8, linestyle="--", label="预测值")
    axes[1].set_ylabel("电价（元/kWh）")
    axes[1].set_xlabel("时刻（h）")
    axes[1].legend(frameon=False, ncol=2)
    fig.suptitle(f"测试集代表日预测曲线（{dates[index]}）")
    fig.tight_layout()
    fig.savefig(OUT / "图5_测试集代表日预测曲线.png", bbox_inches="tight")
    plt.close(fig)
    return str(dates[index])


def significance_rows(
    actual: np.ndarray,
    predictions: dict[str, np.ndarray],
    test_indices: list[int],
    full_name: str,
    target: str,
) -> list[dict[str, object]]:
    full_daily = np.mean(np.abs(actual[test_indices] - predictions[full_name][test_indices]), axis=1)
    rows = []
    for name, values in predictions.items():
        if name == full_name:
            continue
        other_daily = np.mean(np.abs(actual[test_indices] - values[test_indices]), axis=1)
        result = wilcoxon(full_daily, other_daily, alternative="less", zero_method="wilcox")
        rows.append(
            {
                "对象": target,
                "完整模型": full_name,
                "对照模型": name,
                "测试集日均MAE差值": float(np.mean(other_daily - full_daily)),
                "Wilcoxon单侧p值": float(result.pvalue),
                "完整模型显著更优(p<0.05)": "是" if result.pvalue < 0.05 else "否",
            }
        )
    return rows


def model_row(rows, target, model):
    return next(row for row in rows if row["对象"] == target and row["数据集"] == "测试集" and row["模型"] == model)


def build_contributions(all_metrics: list[dict[str, object]], strategy_rows: list[dict[str, object]]):
    """将测试集上的逐模块变化转成正负贡献率；正值代表误差或费用下降。"""
    rows: list[dict[str, object]] = []

    def add_prediction(target: str, module: str, before: str, after: str) -> None:
        old = model_row(all_metrics, target, before)
        new = model_row(all_metrics, target, after)
        old_value, new_value = float(old["RMSE"]), float(new["RMSE"])
        rows.append(
            {
                "对象": target,
                "模块": module,
                "对照模型": before,
                "加入模块后模型": after,
                "评价指标": "RMSE",
                "加入前": old_value,
                "加入后": new_value,
                "贡献率(%；正值为改善)": 100.0 * (old_value - new_value) / old_value,
            }
        )

    add_prediction("负荷预测", "历史信息", "典型日静态基准", "仅历史加权轮廓")
    add_prediction("负荷预测", "年度周期与星期回归", "去除周期回归", "完整模型")
    add_prediction("负荷预测", "小波残差修正", "去除小波修正", "完整模型")
    add_prediction("负荷预测", "贝叶斯参数更新", "去除贝叶斯更新", "完整模型")
    add_prediction("光伏预测（有效发电时段）", "截断幂正弦", "历史加权轮廓", "截断正弦完整模型")
    add_prediction("实时电价预测", "历史信息", "典型日静态基准", "历史加权轮廓")
    add_prediction("实时电价预测", "傅里叶周期提取", "历史加权轮廓", "仅傅里叶周期项")
    add_prediction("实时电价预测", "小波残差重构", "仅傅里叶周期项", "傅里叶+小波")
    add_prediction("实时电价预测", "前日偏差修正", "傅里叶+小波", "增加前日偏差修正")
    add_prediction("实时电价预测", "贝叶斯参数更新", "增加前日偏差修正", "完整模型（含贝叶斯更新）")

    for question in ("问题三", "问题四"):
        base = next(row for row in strategy_rows if row["问题"] == question and str(row["策略"]).startswith("S0"))
        selected = next(row for row in strategy_rows if row["问题"] == question and str(row["策略"]).startswith("S2"))
        rows.append(
            {
                "对象": question + "购电优化",
                "模块": "6:00/12:00 MPC滚动更新",
                "对照模型": "S0_仅0点",
                "加入模块后模型": "S2_增加12点",
                "评价指标": "测试集总费用(元)",
                "加入前": float(base["测试集总费用(元)"]),
                "加入后": float(selected["测试集总费用(元)"]),
                "贡献率(%；正值为改善)": float(selected["相对仅0点节省率(%)"]),
            }
        )
    return rows


def plot_contributions(rows: list[dict[str, object]]) -> None:
    selected = [row for row in rows if row["评价指标"] == "RMSE"]
    labels = [f"{row['对象'].replace('预测（有效发电时段）', '')}—{row['模块']}" for row in selected]
    values = np.asarray([float(row["贡献率(%；正值为改善)"]) for row in selected])
    colors = ["#3E6B89" if value >= 0 else "#C26B4A" for value in values]
    y = np.arange(len(labels))
    fig, ax = plt.subplots(figsize=(10.8, 6.8))
    bars = ax.barh(y, values, color=colors, height=0.62)
    ax.axvline(0, color="#263238", linewidth=0.9)
    ax.set_yticks(y, labels)
    ax.invert_yaxis()
    ax.set_xlabel("测试集RMSE改善率（%）")
    ax.set_title("预测模型各组成模块的边际贡献")
    ax.bar_label(bars, labels=[f"{value:+.2f}%" for value in values], padding=4, fontsize=9)
    span = max(abs(values.min()), abs(values.max()))
    ax.set_xlim(min(values.min() - 8, -8), values.max() + max(8, span * 0.18))
    fig.tight_layout()
    fig.savefig(OUT / "图6_预测模块边际贡献率.png", bbox_inches="tight")
    plt.close(fig)


def write_paper_text(
    all_metrics: list[dict[str, object]],
    strategy_rows: list[dict[str, object]],
    significance: list[dict[str, object]],
    example_day: str,
) -> None:
    load_full = model_row(all_metrics, "负荷预测", "完整模型")
    load_static = model_row(all_metrics, "负荷预测", "典型日静态基准")
    load_no_bayes = model_row(all_metrics, "负荷预测", "去除贝叶斯更新")
    load_no_wavelet = model_row(all_metrics, "负荷预测", "去除小波修正")
    pv_full = model_row(all_metrics, "光伏预测（有效发电时段）", "截断正弦完整模型")
    price_full = model_row(all_metrics, "实时电价预测", "完整模型（含贝叶斯更新）")
    price_persist = model_row(all_metrics, "实时电价预测", "前一日持续性基准")
    price_no_bayes = model_row(all_metrics, "实时电价预测", "增加前日偏差修正")
    price_history = model_row(all_metrics, "实时电价预测", "历史加权轮廓")
    price_fourier = model_row(all_metrics, "实时电价预测", "仅傅里叶周期项")
    price_wavelet = model_row(all_metrics, "实时电价预测", "傅里叶+小波")
    q3_s0 = next(row for row in strategy_rows if row["问题"] == "问题三" and str(row["策略"]).startswith("S0"))
    q3_s2 = next(row for row in strategy_rows if row["问题"] == "问题三" and str(row["策略"]).startswith("S2"))
    q4_s0 = next(row for row in strategy_rows if row["问题"] == "问题四" and str(row["策略"]).startswith("S0"))
    q4_s2 = next(row for row in strategy_rows if row["问题"] == "问题四" and str(row["策略"]).startswith("S2"))
    load_gain = 100 * (float(load_static["RMSE"]) - float(load_full["RMSE"])) / float(load_static["RMSE"])
    bayes_gain = 100 * (float(load_no_bayes["RMSE"]) - float(load_full["RMSE"])) / float(load_no_bayes["RMSE"])
    wavelet_gain = 100 * (float(load_no_wavelet["RMSE"]) - float(load_full["RMSE"])) / float(load_no_wavelet["RMSE"])
    price_gain = 100 * (float(price_persist["RMSE"]) - float(price_full["RMSE"])) / float(price_persist["RMSE"])
    price_bayes_gain = 100 * (float(price_no_bayes["RMSE"]) - float(price_full["RMSE"])) / float(price_no_bayes["RMSE"])
    price_fourier_change = 100 * (float(price_history["RMSE"]) - float(price_fourier["RMSE"])) / float(price_history["RMSE"])
    price_wavelet_gain = 100 * (float(price_fourier["RMSE"]) - float(price_wavelet["RMSE"])) / float(price_fourier["RMSE"])
    price_bias_gain = 100 * (float(price_wavelet["RMSE"]) - float(price_no_bayes["RMSE"])) / float(price_wavelet["RMSE"])
    significant_count = sum(row["完整模型显著更优(p<0.05)"] == "是" for row in significance)
    text = f"""# 模型对比、消融实验与评价（论文可直接使用）

## 1 实验设置

为定量评价模型各组成模块的有效性，本文在统一数据划分下开展基准对比与消融实验。全年样本按照“训练、训练、训练、验证、训练、测试”的六日循环进行划分，共得到244个训练日、61个验证日和60个测试日。模型参数及MPC策略仅依据训练集和验证集确定，测试集只用于最终泛化评价。预测性能采用MAE、RMSE、MAPE、WMAPE和决定系数 $R^2$ 衡量；考虑到光伏功率在夜间为零，光伏百分比误差仅在实际功率大于1 kW的有效发电时段计算。对于可能接近零或穿越零点的净负荷，不将普通MAPE作为主要判据。

## 2 负荷与光伏预测消融实验

测试集上，完整负荷预测模型的MAE、RMSE和 $R^2$ 分别为{float(load_full['MAE']):.2f} kW、{float(load_full['RMSE']):.2f} kW和{float(load_full['R2']):.4f}。相较典型日静态基准，RMSE降低{load_gain:.2f}%，说明引入历史数据、周期规律和动态修正能够显著提升预测精度。移除贝叶斯参数更新后RMSE为{float(load_no_bayes['RMSE']):.2f} kW，完整模型相对降低{bayes_gain:.2f}%；移除小波修正后RMSE为{float(load_no_wavelet['RMSE']):.2f} kW，对应改进幅度为{wavelet_gain:.2f}%。这表明贝叶斯更新用于适应季节变化，小波模块用于保留日内局部波动，二者对最终精度均可通过消融结果单独辨识。

光伏预测采用历史加权轮廓与截断幂正弦模型。有效发电时段内，完整模型的RMSE为{float(pv_full['RMSE']):.2f} kW、WMAPE为{float(pv_full['WMAPE(%)']):.2f}%、$R^2$ 为{float(pv_full['R2']):.4f}。相较未经形状约束的历史加权轮廓，截断正弦使RMSE小幅下降，表明其主要作用是保证夜间功率为零并提高曲线的物理合理性，而非带来大幅数值增益。前一日持续性基准在MAE、MAPE和WMAPE上略占优势，但完整模型具有更低的RMSE和更高的 $R^2$，说明其对较大偏差和整体峰形的控制更好。MAPE受日出和日落附近小分母影响明显，因此本文主要结合RMSE、WMAPE和 $R^2$ 评价光伏预测效果。

负荷与光伏组合形成的净负荷是后续购电优化的直接输入。完整模型在测试集上的净负荷MAE为{float(model_row(all_metrics, '净负荷预测', '完整模型')['MAE']):.2f} kW、RMSE为{float(model_row(all_metrics, '净负荷预测', '完整模型')['RMSE']):.2f} kW、WMAPE为{float(model_row(all_metrics, '净负荷预测', '完整模型')['WMAPE(%)']):.2f}%、$R^2$ 为{float(model_row(all_metrics, '净负荷预测', '完整模型')['R2']):.4f}。由于净负荷在光伏大发时段可能接近或穿越零点，普通MAPE被小分母显著放大，故不作为净负荷预测的核心结论。

## 3 实时电价预测消融实验

测试集上，完整电价预测模型取得MAE={float(price_full['MAE']):.4f}元/kWh、RMSE={float(price_full['RMSE']):.4f}元/kWh、MAPE={float(price_full['MAPE(%)']):.2f}%和 $R^2={float(price_full['R2']):.4f}$。相较前一日持续性基准，RMSE降低{price_gain:.2f}%。消融结果并非单调：仅保留傅里叶周期项时，RMSE相对历史加权轮廓上升{abs(price_fourier_change):.2f}%，表明单独的低阶周期拟合会过度平滑局部尖峰；在此基础上加入小波重构和前日偏差修正后，RMSE分别改善{price_wavelet_gain:.2f}%和{price_bias_gain:.2f}%，逐步补回局部和短期扰动信息。最后，贝叶斯更新使固定参数链的RMSE由{float(price_no_bayes['RMSE']):.4f}元/kWh降至{float(price_full['RMSE']):.4f}元/kWh，进一步改善{price_bayes_gain:.2f}%。因此，傅里叶项的价值主要在于提供可解释的周期骨架，必须与小波、偏差修正及自适应参数联合使用，不能将其单独解释为精度增益。

## 4 MPC滚动优化策略评价

在固定电价条件下，问题三测试集的仅0点策略总费用为{float(q3_s0['测试集总费用(元)']):.2f}元；增加6:00和12:00滚动修正后的S2策略费用为{float(q3_s2['测试集总费用(元)']):.2f}元，相对节省{float(q3_s2['相对仅0点节省率(%)']):.2f}%，紧急购电量削减{float(q3_s2['紧急购电削减率(%)']):.2f}%。在动态电价条件下，问题四S2策略相对仅0点策略节省{float(q4_s2['相对仅0点节省率(%)']):.2f}%，并使紧急购电量减少{float(q4_s2['紧急购电削减率(%)']):.2f}%。结果表明，多时刻滚动更新能够利用新到达的负荷、光伏和电价信息修正日前计划，显著降低预测偏差导致的紧急购电成本。

需要指出的是，增加更新次数并不必然单调降低总费用。18:00时剩余调节窗口较短，且频繁调整可能产生额外的高价增购，因此完整四时刻MPC策略未必优于0:00、6:00和12:00三时刻组合。本文据此使用验证集费用选择S2策略，并在测试集上冻结检验，避免根据测试结果反向选取策略。

进一步定义新增预报时刻的边际信息价值为 $V(τ|T)=C(T)-C(T∪{{τ}})$。测试集结果表明，增加6:00和12:00预报的边际价值分别为{float(q3_s0['测试集总费用(元)'])-float(next(row for row in strategy_rows if row['问题']=='问题三' and str(row['策略']).startswith('S1'))['测试集总费用(元)']):.2f}元和{float(next(row for row in strategy_rows if row['问题']=='问题三' and str(row['策略']).startswith('S1'))['测试集总费用(元)'])-float(q3_s2['测试集总费用(元)']):.2f}元，而继续增加18:00预报的边际价值为{float(q3_s2['测试集总费用(元)'])-float(next(row for row in strategy_rows if row['问题']=='问题三' and str(row['策略']).startswith('S3'))['测试集总费用(元)']):.2f}元。这说明随着发布时间推迟，剩余调节窗口缩短，新增信息的经济价值逐步下降。附件3只提供0:00、6:00、12:00和18:00四个独立预报时刻，缺少其他时刻的真实新增信息，因而不宜通过插值虚构额外预报。在现有数据和结算规则下，无需再引入其他时刻；若实际系统能够提供新的实时预报，可优先在验证集考察15:00更新点，并仅在费用、紧急购电量和跨月份稳定性均改善时将其纳入策略。

## 5 综合评价

消融与基准对比表明，当前模型的性能提升并非来自单一模块，而是由周期特征提取、局部残差重构、短期偏差修正、贝叶斯参数更新和MPC滚动决策共同形成。配对Wilcoxon检验中共有{significant_count}组对照达到 $p<0.05$，具体结果见“统计显著性检验.csv”。测试集代表日{example_day}的实际—预测曲线进一步表明，模型能够跟踪主要峰谷和日内变化。总体而言，模型具有较好的预测精度与样本外稳定性，并能将预测精度提升转化为购电费用和紧急购电量的下降；其主要局限是对突发尖峰、日出日落低功率区间以及晚间短调节窗口的刻画仍有改进空间。

## 图表引用建议

- 图1：负荷预测基准与消融实验。
- 图2：光伏预测基准对比实验。
- 图2b：净负荷预测基准对比实验。
- 图3：实时电价预测基准与消融实验。
- 图4：问题三、问题四MPC滚动策略经济性对比。
- 图5：测试集代表日负荷和实时电价预测曲线。
- 图6：预测模型各组成模块的测试集边际贡献率；负值表示该模块单独加入时造成误差上升。
"""
    (OUT / "论文可用_模型检验与评价.md").write_text(text, encoding="utf-8")


def main() -> None:
    configure_plot()
    OUT.mkdir(parents=True, exist_ok=True)
    print("读取正式数据与当前模型参数……")
    baseline, order = two.read_baseline(two.DEFAULT_BASELINE)
    historical = two.read_historical_data(two.DEFAULT_LOAD, two.DEFAULT_PV, order)
    split = two.DatasetSplit()
    groups = split.cycle_labels(historical.dates)
    training = set(groups["训练集"])

    print("计算问题二基准与模块消融……")
    load_pred, pv_pred = load_and_pv_predictions(baseline, historical, training)
    power_rows = evaluate_predictions("负荷预测", historical.load_kw, load_pred, groups, 1.0)
    power_rows += evaluate_predictions(
        "光伏预测（有效发电时段）", historical.pv_kw, pv_pred, groups, 1.0, daylight_only=True
    )
    net_pred = {
        "典型日静态基准": load_pred["典型日静态基准"] - pv_pred["典型日静态基准"],
        "前一日持续性基准": load_pred["前一日持续性基准"] - pv_pred["前一日持续性基准"],
        "仅历史轮廓": load_pred["仅历史加权轮廓"] - pv_pred["历史加权轮廓"],
        "完整模型": load_pred["完整模型"] - pv_pred["截断正弦完整模型"],
    }
    power_rows += evaluate_predictions(
        "净负荷预测", historical.load_kw - historical.pv_kw, net_pred, groups, 1.0
    )
    write_rows(OUT / "问题二_基准与消融指标.csv", power_rows)

    print("计算问题四电价基准与模块消融……")
    prices = four.read_dynamic_prices(four.DEFAULT_PRICE, order)
    price_pred = price_predictions(baseline, prices, training)
    price_rows = evaluate_predictions("实时电价预测", prices.price, price_pred, groups, 0.05)
    write_rows(OUT / "问题四_电价基准与消融指标.csv", price_rows)

    all_rows = power_rows + price_rows
    significance = significance_rows(
        historical.load_kw, load_pred, groups["测试集"], "完整模型", "负荷预测"
    )
    significance += significance_rows(
        prices.price,
        price_pred,
        groups["测试集"],
        "完整模型（含贝叶斯更新）",
        "实时电价预测",
    )
    write_rows(OUT / "统计显著性检验.csv", significance)

    strategy_rows = read_strategy_rows(ROOT / "C题" / "问题三_训练验证测试检验.csv", "问题三")
    strategy_rows += read_strategy_rows(ROOT / "C题" / "问题四_训练验证测试检验.csv", "问题四")
    write_rows(OUT / "问题三四_MPC测试集策略对比.csv", strategy_rows)
    write_rows(OUT / "问题三_新增预报时刻边际价值.csv", release_marginal_rows(strategy_rows))
    contributions = build_contributions(all_rows, strategy_rows)
    write_rows(OUT / "各组成模块测试集贡献率.csv", contributions)

    print("生成论文图表……")
    plot_metric_comparison(all_rows, "负荷预测", "图1_负荷预测基准与消融.png", "kW")
    plot_metric_comparison(all_rows, "光伏预测（有效发电时段）", "图2_光伏预测基准对比.png", "kW")
    plot_metric_comparison(all_rows, "净负荷预测", "图2b_净负荷预测基准对比.png", "kW")
    plot_metric_comparison(all_rows, "实时电价预测", "图3_实时电价预测基准与消融.png", "元/kWh")
    plot_mpc(strategy_rows)
    plot_contributions(contributions)
    example_day = plot_example(
        historical.dates,
        groups["测试集"],
        historical.load_kw,
        load_pred["完整模型"],
        prices.price,
        price_pred["完整模型（含贝叶斯更新）"],
    )
    write_paper_text(all_rows, strategy_rows, significance, example_day)
    print(f"完成：{OUT}")


if __name__ == "__main__":
    main()
