# -*- coding: utf-8 -*-
"""
new_project.py — 一键建赛脚手架

用法：
    python tools/new_project.py 2026C_管道运输
    python tools/new_project.py 2026MCM_C --force   # 目录已存在时合并式补齐

在 projects/<名称>/ 下生成标准目录结构、全套模板副本、图册登记表与进度看板。
模板只复制不覆盖——你在项目里改过的内容永远安全。
"""
from __future__ import annotations

import argparse
import csv
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TPL = ROOT / "templates"

DIRS = ["01_审题", "02_建模", "03_代码", "03_代码/附件", "03_代码/结果模板",
        "04_图表/figs", "05_论文/figs", "06_交付", "07_复盘"]
COPIES = [
    ("题目分析.md",     "01_审题/题目分析.md"),
    ("模型卡片.md",     "02_建模/模型卡片-模型I.md"),
    ("图表需求单.md",   "04_图表/图表需求单.md"),
    ("论文骨架.md",     "05_论文/论文骨架.md"),
    ("摘要模板.md",     "05_论文/摘要模板.md"),
    ("00_标准/论文模板/CUMCM_国赛/CUMCM论文模板_官方格式_2026.docx",
                        "05_论文/论文模板_官方格式.docx"),
    ("提交检查清单.md", "06_交付/提交检查清单.md"),
]

BOARD = """# 进度看板 · {name}

> 72h 时间轴与各阶段门禁详见工作区 `WORKFLOW.md`；本页只追踪状态，勾选即验收通过。

## 里程碑（对应 WORKFLOW.md 的 Gates）
- [ ] **G1 选题审题**：`01_审题/题目分析.md` 完成，全队对齐
- [ ] **G2 骨架先行**：论文骨架 + 技术路线图 v1 + 符号表开张
- [ ] **G3 模型一轮**：问题一有图、有数、有公式
- [ ] **G4 模型滚动**：问题二/三初稿 + 图册 6–10 张
- [ ] **G5 补件深化**：敏感性分析 + 模型评价章
- [ ] **G6 摘要冲刺**：五要素摘要完成三轮打磨
- [ ] **G7 质检交付**：P0 清零 + `06_交付/提交检查清单.md` 全勾

## 图册登记（figure_index.csv 单一事实源）
04_图表/figure_index.csv —— 每张正式图登记：图号/文件名/标题/caption/引用章节/版本/状态。

## 常用命令
```bash
# 流程图
python tools/make_flowchart.py --spec flow.json --outdir "projects/{name}/04_图表/figs"
# 图表（CSV 进，出版级图出）
python tools/make_chart.py line --data "03_代码/F01_收敛.csv" --outdir "projects/{name}/04_图表/figs"
```
"""


def main():
    ap = argparse.ArgumentParser(description="一键建赛：生成标准项目目录与全套模板")
    ap.add_argument("name", help="项目名，如 2026C_管道运输")
    ap.add_argument("--force", action="store_true", help="目录已存在时合并式补齐缺失文件")
    a = ap.parse_args()

    proj = ROOT / "projects" / a.name
    if proj.exists() and not a.force:
        sys.exit(f"[!] {proj} 已存在；确认要补齐请加 --force")

    for d in DIRS:
        (proj / d).mkdir(parents=True, exist_ok=True)

    created = []
    for src, dst in COPIES:
        s = ROOT / src if src.startswith("00_标准") else TPL / src
        t = proj / dst
        if s.exists() and not t.exists():
            shutil.copy(s, t)
            created.append(dst)

    idx = proj / "04_图表" / "figure_index.csv"
    if not idx.exists():
        with open(idx, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow(
                ["图号", "文件名", "标题", "caption", "引用章节", "版本", "状态"])
        created.append("04_图表/figure_index.csv")

    board = proj / "进度看板.md"
    if not board.exists():
        board.write_text(BOARD.format(name=a.name), encoding="utf-8")
        created.append("进度看板.md")

    print(f"[OK] 项目已创建：{proj}")
    for c in created:
        print(f"     + {c}")
    print("     下一步：把题目原文投给 /审题 → 填写 01_审题/题目分析.md")


if __name__ == "__main__":
    main()
