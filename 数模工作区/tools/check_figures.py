# -*- coding: utf-8 -*-
"""
check_figures.py — 「文-图-册」一致性自动检查（审图十查第⑩条的自动化）

用法：
    python tools/check_figures.py                          # 检查 projects/ 下全部项目
    python tools/check_figures.py projects/2025A_烟幕干扰弹  # 只检查指定项目

检查项（对应 WORKFLOW.md 图表生产 SOP 第 6 步）：
  A. figure_index.csv 每行：图文件在 04_图表/figs/ 中存在（PNG）、引用章节非空、图号唯一
  B. figs/ 下每个 PNG 都已登记入册（防"画了图没登记"）
  C. 每张登记的图都在论文源（05_论文/**/*.tex|md）中被引用——按「图号」或「文件名主干」检索

输出：按项目打印问题清单（[P0]阻断 / [P1]警告）；全部通过退出码 0，否则 1。
"""
from __future__ import annotations

import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def check_project(proj: Path) -> list[str]:
    problems: list[str] = []
    index_csv = proj / "04_图表" / "figure_index.csv"
    figs_dir = proj / "04_图表" / "figs"
    if not index_csv.exists():
        return [f"[P0] 缺少图册登记表：{index_csv}"]

    with open(index_csv, encoding="utf-8-sig", newline="") as f:
        rows = [r for r in csv.reader(f) if r]
    header, entries = rows[0], rows[1:]
    col = {name: i for i, name in enumerate(header)}

    # ---- A. 逐行检查 ----
    seen_ids: set[str] = set()
    registered_stems: set[str] = set()
    for r in entries:
        if len(r) < len(header):
            problems.append(f"[P0] 登记行列数不足：{r}")
            continue
        fid = r[col["图号"]].strip()
        fname = r[col["文件名"]].strip()
        section = r[col["引用章节"]].strip()
        if fid in seen_ids:
            problems.append(f"[P0] 图号重复：{fid}")
        seen_ids.add(fid)
        png = figs_dir / fname
        if not png.exists():
            problems.append(f"[P0] {fid} 登记的文件不存在：{fname}")
        if not section:
            problems.append(f"[P1] {fid} 未填写引用章节")
        registered_stems.add(png.stem)

    # ---- B. 反查：figs/ 里的图是否都登记了 ----
    if figs_dir.exists():
        for png in sorted(figs_dir.glob("*.png")):
            if png.stem not in registered_stems:
                problems.append(f"[P1] 图未登记入册：{png.name}")

    # ---- C. 论文源引用检查 ----
    sources = ""
    for pat in ("*.md", "*.tex"):
        for p in (proj / "05_论文").rglob(pat):
            if "figs" in p.parts:
                continue
            sources += p.read_text(encoding="utf-8", errors="ignore")
    for r in entries:
        if len(r) < len(header):
            continue
        fid = r[col["图号"]].strip()
        stem = (figs_dir / r[col["文件名"]].strip()).stem
        if fid in sources or stem in sources:
            continue
        problems.append(f"[P1] {fid} 未在论文源（05_论文）中被引用（检索：图号或文件名主干）")

    return problems


def main():
    targets = []
    if len(sys.argv) > 1:
        targets.append(Path(sys.argv[1]))
    else:
        pj = ROOT / "projects"
        targets = sorted(p for p in pj.iterdir() if p.is_dir()) if pj.exists() else []
    if not targets:
        print("没有找到可检查的项目目录")
        return

    total = 0
    for proj in targets:
        problems = check_project(proj)
        print(f"\n=== 图册一致性检查 · {proj.name} ===")
        if not problems:
            print("  全部通过：登记/文件/引用三者一致 ✓")
            continue
        for p in problems:
            print("  " + p)
        total += len(problems)

    print(f"\n[结果] 共 {total} 个问题")
    sys.exit(1 if total else 0)


if __name__ == "__main__":
    main()
