# -*- coding: utf-8 -*-
"""review_crosscheck.py — 总揽审查底稿生成器（角色卡 09 配套）
用法：python tools/review_crosscheck.py projects/<赛题>
输出：<赛题>/06_交付/审查底稿.md（机器可查项全部自动查；判断由审查官完成）

四类自动检查：
  A 求解健康：独立重算意识 / 求解状态记录 / 随机种子 / TODO 残留
  B 结果文件健康：NaN / inf / 空文件 / 重复列
  C 题意覆盖：题目问题数 vs 草稿小节 vs result 交付表
  D 建模-代码信号：模型卡片约束条款数 vs 代码约束行数（量级对照，非精确匹配）
"""
import json
import re
import sys
from pathlib import Path

import pandas as pd

WS = Path(__file__).resolve().parents[1]


def scan_solve_health(code_dir: Path, lines):
    lines.append("## A 求解健康（代码信号）\n")
    pys = sorted(code_dir.glob("*.py"))
    if not pys:
        lines.append("- ⚠️ 未发现任何 .py 求解脚本\n")
        return
    all_text = {p.name: p.read_text(encoding="utf-8", errors="ignore") for p in pys}
    checks = [
        ("独立重算/校验意识", ["重算", "独立校验", "独立重算", "diff 一致", "守恒"]),
        ("求解状态记录", ["status", "Optimal", "res.fun"]),
        ("随机种子/可复现", ["default_rng", "seed", "random_state"]),
    ]
    joined = "\n".join(all_text.values())
    for name, kws in checks:
        hit = [k for k in kws if k in joined]
        lines.append(f"- {'✅' if hit else '⚠️'} {name}：{'信号 ' + str(hit[:3]) if hit else '未发现——请人工确认'}")
    for name, text in all_text.items():
        bad = re.findall(r"TODO|FIXME|待填|待补", text)
        if bad:
            lines.append(f"- ⚠️ {name}：残留 {len(bad)} 处 TODO/待填")
    lines.append("")


def scan_results(code_dir: Path, lines):
    lines.append("## B 结果文件健康（CSV 扫描）\n")
    csvs = [p for p in sorted(code_dir.glob("*.csv"))]
    if not csvs:
        lines.append("- ⚠️ 03_代码 下无结果 CSV\n")
        return
    bad_total = 0
    for p in csvs:
        raw_attachment = p.name.startswith(("A1_", "A2_"))   # 原始附件提取件：含注释行属正常
        try:
            df = pd.read_csv(p, encoding="utf-8-sig")
        except Exception as e:
            lines.append(f"- ⚠️ {p.name}：读取失败 {e}")
            bad_total += 1
            continue
        issues = []
        if df.empty:
            issues.append("空表")
        if df.columns.duplicated().any():
            issues.append("重复列")
        num = df.select_dtypes("number")
        n_nan = int(num.isna().sum().sum())
        n_inf = int(np_inf_count(num))
        if n_inf:
            issues.append(f"inf×{n_inf}")
        wide = p.stem.endswith("_wide")
        if n_nan and not raw_attachment and not wide:
            issues.append(f"NaN×{n_nan}")
        note = ("（原始附件提取件，表尾注释行属正常）" if raw_attachment and n_nan
                else ("（宽表空格填充属正常）" if n_nan and wide else ""))
        if issues:
            bad_total += 1
            lines.append(f"- ⚠️ {p.name}：{', '.join(issues)} {note}")
        elif note:
            lines.append(f"- ✅ {p.name}：健康 {note}")
    if bad_total == 0:
        lines.append(f"- ✅ {len(csvs)} 个 CSV 全部健康（无 inf/意外 NaN/空表/重复列）")
    lines.append("")


def np_inf_count(num_df):
    import numpy as np
    return int(np.isinf(num_df).to_numpy().sum()) if len(num_df) else 0


def scan_coverage(proj: Path, lines):
    lines.append("## C 题意覆盖（题目问题数 vs 草稿小节 vs 交付表）\n")
    brief = proj / "01_审题" / "题目分析.md"
    draft = proj / "05_论文" / "章节草稿.md"
    n_q = 0
    if brief.exists():
        text = brief.read_text(encoding="utf-8")
        n_q = len(set(re.findall(r"问题\s*[一二三123]", text)) )
    checks = [
        ("result1_1.xlsx", draft, "result1_1"),
        ("result1_2.xlsx", draft, "result1_2"),
        ("result2.xlsx", draft, "result2"),
    ]
    for fname, src, kw in checks:
        delivered = (proj / "06_交付" / fname).exists()
        mentioned = src.exists() and (kw in src.read_text(encoding="utf-8"))
        state = "✅" if (delivered and mentioned) else "⚠️"
        lines.append(f"- {state} {fname}：已交付={'是' if delivered else '否'}，草稿提及={'是' if mentioned else '否'}")
    if draft.exists():
        dt = draft.read_text(encoding="utf-8")
        for i, zh in [(1, "一"), (2, "二"), (3, "三")]:
            ok = bool(re.search(rf"5\.{i}\s*问题{zh}", dt)) or f"5.{i}" in dt
            lines.append(f"- {'✅' if ok else '⚠️'} 题目问题{i} ↔ 论文小节 5.{i}")
    lines.append(f"- 审题纪要覆盖问题数（信号）：{n_q}")
    lines.append("")


def scan_model_code_signal(proj: Path, lines):
    lines.append("## D 建模-代码信号对照（量级参考，精确对照由审查官逐条做）\n")
    cards = sorted((proj / "02_建模").glob("*.md"))
    code = (proj / "03_代码").glob("*.py")
    card_terms = 0
    for c in cards:
        t = c.read_text(encoding="utf-8")
        card_terms += len(re.findall(r"式\(|约束|假设|≤|≥", t))
    code_terms = 0
    for c in code:
        t = c.read_text(encoding="utf-8", errors="ignore")
        code_terms += len(re.findall(r"add_row|add_constraint|constraint|subject to", t, re.I))
    lines.append(f"- 模型卡片约束/公式信号 {card_terms} 处 vs 代码约束信号 {code_terms} 处")
    lines.append("- ⚠️ 请逐条执行'卡片条款→代码实现行'对照（角色卡 09 线二第 1 问），"
                 "本工具只给量级信号，不做精确匹配。\n")


def main():
    proj = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    if not proj or not proj.exists():
        raise SystemExit("用法：python tools/review_crosscheck.py projects/<赛题>")
    out = proj / "06_交付" / "审查底稿.md"
    lines = [f"# 审查底稿（自动生成 · {proj.name}）\n",
             "> 机器可查项；判断与深查由总揽审查官（prompts/09）按三条对抗线完成。\n"]
    code_dir = proj / "03_代码"
    scan_solve_health(code_dir, lines)
    scan_results(code_dir, lines)
    scan_coverage(proj, lines)
    scan_model_code_signal(proj, lines)
    lines.append("## 后续动作\n")
    lines.append("1. 审查官按 `prompts/09_总揽审查官.md` 三条对抗线深查，产出《审查意见单》")
    lines.append("2. 图册一致性：`python tools/check_figures.py <赛题>`")
    lines.append("3. 意见单投 `协作接口/论文手回执/`， teammates 销项\n")
    out.write_text("\n".join(lines), encoding="utf-8")
    print(f"OK -> {out}")


if __name__ == "__main__":
    main()
