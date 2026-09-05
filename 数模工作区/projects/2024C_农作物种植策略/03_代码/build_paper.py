# -*- coding: utf-8 -*-
"""build_paper.py — 论文整卷装配器
把章节正文、图表（自动插入并编号）、参考文献、支撑材料清单、源程序附录
一次性装配进官方格式 Word 模板，产出"可润色即可提交"的完整初稿。

输入：05_论文/论文模板_官方格式.docx + 本文件内嵌的章节内容（与 05_论文/章节草稿.md 同源）
输出：05_论文/论文初稿_完整版.docx
人工动作：只做判断句/衔接句润色与 AI 使用声明复核——不再做搬运。
"""
import copy
import os
from pathlib import Path

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

HERE = Path(__file__).resolve().parent
PROJ = HERE.parent
SRC = PROJ / "05_论文" / "论文模板_官方格式.docx"
DST = PROJ / "05_论文" / "论文初稿_完整版.docx"
FIGS = PROJ / "04_图表" / "figs"

TITLE = "基于混合整数规划与情景模拟的农作物种植策略优化研究"
KEYWORDS = "农作物种植策略；混合整数规划；销量线性化；情景模拟；轮作约束"

ABSTRACT = [
    "本文研究华北山区某乡村 2024—2030 年农作物种植方案的优化问题。该村拥有 4 类露地"
    "1201 亩与普通、智慧大棚共 20 个，需在地块适配、禁止重茬、豆类三年轮作、种植便于"
    "管理等约束下，对 41 种作物逐年逐季安排种植面积，使七年总净收益最大。",
    "针对问题一，建立多季种植计划的混合整数规划（MILP）模型，将超产部分“滞销”与“五折"
    "甩卖”两种情形统一为带销量上限的线性收益结构，采用 HiGHS 求解器精确求解，得七年总"
    "净收益分别为 7 816.4 万元与 8 651.4 万元，五折情形比滞销情形多回收 835.0 万元"
    "（+10.7%），方案已逐年逐季回填 result1_1 与 result1_2 表。",
    "针对问题二，按题给区间构造销量、成本、价格的中枢变化路径并重新求解，得净收益"
    " 9 402.1 万元（方案回填 result2）；20 组均匀抽样情景检验表明，沿用该方案的收益均值为"
    " 9 134.5 万元、变异系数仅 1.7%，按情景重优化平均仅再提高 0.2%，方案对参数波动稳健。",
    "针对问题三，引入蔬菜类需求的共同市场因子与需求—价格负相关机制，进行 8 组相关蒙特"
    "卡洛模拟并逐组重优化：重优化平均净收益 14 041.1 万元，比沿用问题二方案提高 1 327.5 "
    "万元（+10.4%），表明存在相关冲击时应逐年滚动优化。",
    "本文的亮点在于：约束体系完整覆盖题目全部硬性要求且线性可核验；“中枢求解—情景评估—"
    "重优化对比”三层递进量化了不确定性与灵活调整的价值；以消融实验检验了每类约束的实际"
    "约束力；全部结果可由附录命令一键复现。",
]

CONTENT = {
    "1 问题重述": [
        ("p", "某乡村地处华北山区，现有露天耕地 1201 亩（34 个地块，含平旱地、梯田、山坡地、"
              "水浇地 4 类），另有普通大棚 16 个、智慧大棚 4 个（每个 0.6 亩）。题目要求在满足"
              "地块适配、不能连年重茬、每个地块三年内至少种一次豆类、种植方案便于田间管理等"
              "条件下，为 2024—2030 年给出逐年逐季的最优种植方案。"),
        ("p", "问题 1：假定各参数相对 2023 年保持稳定且当季销售，超产部分分别按 (1) 滞销浪费、"
              "(2) 五折出售两种情形优化，结果填入 result1_1.xlsx 与 result1_2.xlsx。"),
        ("p", "问题 2：考虑各类作物销量、亩产量、成本、价格的不确定性（题给变化区间）与种植"
              "风险，给出最优方案，结果填入 result2.xlsx。"),
        ("p", "问题 3：进一步考虑作物间可替代性与互补性，以及销量—价格—成本之间的相关性，"
              "用模拟数据求解，并与问题 2 的结果进行比较分析。"),
    ],
    "2 问题分析": [
        ("p", "三类地块的种植制度不同：平旱地、梯田、山坡地每年只能种一季粮食（水稻除外）；"
              "水浇地可种一季水稻或两季蔬菜（第二季只能种大白菜、白萝卜、红萝卜之一）；普通"
              "大棚第一季种蔬菜、第二季种食用菌；智慧大棚两季均可种蔬菜。决策变量是"
              "“地块×年份×季次×作物”上的种植面积，约束均可表达为线性形式，目标为七年净收益"
              "最大——这是一个典型的多期混合整数规划问题。"),
        ("p", "超产“滞销/五折”的收益结构通过引入销量变量并加以销量上限约束即可线性化；"
              "情形 (2) 利用恒等式将五折段并入目标系数，使模型更紧凑。"),
        ("p", "问题 2、问题 3 的不确定性不采用两阶段随机规划的完整形式（变量规模随情景数"
              "爆炸），而采用“中枢参数求解 + 蒙特卡洛情景评估与重优化对比”的路线：既能给出"
              "可提交的确定方案，又能量化风险的影响。全文技术路线见图 1。"),
        ("fig", "flow_技术路线.png", "图1　全文技术路线"),
    ],
    "3 模型假设": [
        ("list", "销售单价取附件所给区间的中值；区间端点差异由灵敏度分析覆盖。"),
        ("list", "预期销售量取该作物 2023 年总产量（Σ 各地块 2023 种植面积×对应亩产量），"
                 "每季销售上限等于该值；2023 年未种植的作物（空心菜、黄心菜）预期销售量取 0。"),
        ("list", "“种植不宜太分散、面积不宜太小”量化为：单地块单作物单季面积 ≥0.5 亩（露地）"
                 "/0.3 亩（大棚），且每种作物单季占用地块数 ≤8。"),
        ("list", "重茬约束作用于时间相邻的两个种植季（含 2023→2024 交界；某季休耕则视为断开）。"),
        ("list", "豆类作物为类型含“豆类”的 8 种作物（编号 1—5、17—19）；“三年内至少一次”"
                 "按滑动窗口 2024—2026、…、2028—2030 建模。"),
        ("list", "大棚第二季（跨年至次年 4 月）的产出按当年第二季记账。"),
        ("list", "水稻在水浇地整年种植（面积等于地块面积），当年不再种蔬菜。"),
        ("list", "问题 2、3 的超产部分沿用问题 1 情形 (2) 的五折口径。"),
    ],
    "4 符号说明": [
        ("p", "全文主要符号见表 1，正文未列出者在首次出现处定义。"),
    ],
    "5 模型的建立与求解": [],
    "5.1 问题一模型的建立与求解": [],
    "5.1.1 建模思路": [
        ("p", "本题的每个决策天然具有两重性质：“种不种”是 0-1 的，“种多少”是连续的。将二者"
              "拆分为 0-1 变量 y 与连续面积变量 x 分工表达，重茬、轮作等“非黑即白”的规则即可"
              "写成 y 上的线性约束，面积分摊由 x 承担——这构成一个多期混合整数规划（MILP）"
              "模型。方法对比与选型依据见模型卡片（支撑材料 02_建模 目录）。"),
    ],
    "5.1.2 模型建立": [
        ("p", "约束体系共五类：①每季地块面积守恒（式 1-1）；②种植面积与种植状态联动，"
              "含最小面积要求（式 1-2）；③相邻季槽重茬约束（式 1-3，2023 年实种作为常量边界）；"
              "④豆类三年滑动窗口约束（式 1-4）；⑤分散度约束。目标为七年净收益最大。"),
        ("p", "其中，x 与 y 分别为种植面积与种植状态变量；σ 为正常价格销量，满足"
              " σ ≤ min(γ·x, D)，γ 为亩产量、D 为预期销售量。情形 (2) 利用恒等式“收入 = "
              "0.5·价格·产量 + 0.5·价格·min(产量, 销量)”将五折段并入目标系数，模型变量数"
              "减少约四分之一、线性松弛更紧。式(1-2)中 m 为最小面积（露地 0.5 亩、大棚 0.3 亩，"
              "水稻取整块地面积）；式(1-3)中相邻季槽沿时间链定义，2023 年实种由附件 2 给出；"
              "式(1-4)中窗口 w 取 2024—2026 至 2028—2030 共五个。"),
    ],
    "5.1.3 模型求解": [
        ("p", "模型采用 scipy.optimize.milp（HiGHS 后端[4]）精确求解，0-1 变量约 7000 个，"
              "两种情形均取得精确最优。步骤：①数据清洗与口径统一（区间取中值、销量口径转换）；"
              "②构建“地块×年份×季次”季槽链并标注 2023 历史常量；③按约束体系装配稀疏矩阵；"
              "④调用 HiGHS 求解并校验；⑤提取方案回填 result 表并做面积守恒校验（文献[1][3]）。"),
        ("p", "求解结果通过两项独立校验：按解重算目标值与求解器报告一致（差异仅为 CSV 保留"
              "两位小数的舍入）；result 表回填总面积与计划面积 diff 一致。"),
    ],
    "5.1.4 结果分析": [
        ("p", "滞销情形七年总净收益 7 816.4 万元，五折情形 8 651.4 万元，五折较滞销多"
              " 835.0 万元（+10.7%）——差值恰为超产产量按半价回收的价值，说明该乡村种植结构"
              "总体“供大于求”程度有限但客观存在。两种情形历年净收益稳定在 1 115 万—1 238 万元"
              "之间（图 2），年度间的小幅波动主要来自重茬约束对相邻年份作物选择的限制。"),
        ("fig", "F01_历年收入对比.png", "图2　两种超产情形历年净收益对比"),
        ("p", "从面积结构看（图 3），粮食类保持基本盘，蔬菜类依托水浇地与大棚的两季制度贡献"
              "主要面积弹性，食用菌稳定占据普通大棚第二季。方案已逐年逐季回填 result1_1.xlsx"
              "（滞销）与 result1_2.xlsx（五折），面积守恒校验一致。"),
        ("fig", "F02_面积结构.png", "图3　滞销方案作物类型面积结构"),
    ],
    "5.2 问题二模型的建立与求解": [
        ("p", "建模思路：题给不确定性为“区间 + 年均速率”而非分布，直接做两阶段随机规划会使"
              "变量规模随情景数倍增（7 年 × 情景数），故采用“中枢路径求解 + 蒙特卡洛检验”的"
              "两层结构：先在最可能参数下求可提交的确定方案，再以大量抽样情景量化其稳健性。"
              "模型在问题一约束体系基础上将参数改为随年份 t 变化的路径：小麦、玉米预期销售量"
              "按年增 7.5% 复合，种植成本按年增 5% 复合，蔬菜类价格年增 5%、食用菌年降 3%"
              "（羊肚菌 5%）、粮食类价格保持稳定，亩产量取中性。"),
        ("p", "求解与结果：中枢路径下七年净收益 9 402.1 万元（方案回填 result2.xlsx，面积"
              "守恒校验一致），高于问题一情形 (2) 的 8 651.4 万元——需求增长打开的销量空间"
              "超过成本上涨的侵蚀。以该方案为基础做 20 组均匀抽样情景检验（图 4）：沿用方案"
              "收益均值 9 134.5 万元、标准差 151.2 万元（约 1.7%）；按情景重新优化平均仅再"
              "提高 20.8 万元（+0.2%）。“计划刚性”的代价不足 0.3%，说明题给区间内中枢方案"
              "已接近稳健最优。"),
        ("fig", "F04_收益分布.png", "图4　20 组情景下沿用方案与重优化的收益分布"),
    ],
    "5.3 问题三模型的建立与求解": [
        ("p", "建模思路：现实中各类作物共享气候、物流与消费习惯等共同因素，需求之间呈正相关"
              "（替代品呈负相关），且需求走弱时价格往往同步承压。情景生成中引入蔬菜类共同市场"
              "因子 f（类内相关系数约 0.6），并令价格涨幅随需求冲击反向修正（弹性 −0.2），使"
              "模拟数据同时体现替代/互补与需求—价格相关性。"),
        ("p", "求解与结果：共 8 组相关参数路径，每组分别评估沿用问题二方案的收益、用该组参数"
              "重新求解 MILP。沿用方案平均净收益 12 713.6 万元，重优化平均 14 041.1 万元，"
              "重优化提升 1 327.5 万元（+10.4%），远高于问题二无相关情景下的 0.2%（图 5）。"
              "机理：共同因子使多数蔬菜需求同向波动，固定方案在好年份受销量上限约束、坏年份"
              "受价格下压，两头受损；重优化则逐季把面积转向当期性价比最高的作物。需要说明，"
              "相关模拟的需求中枢含增长趋势（蔬菜类约 +2%/年），故绝对收益高于问题二情景；"
              "有效的比较是同一情景内沿用与重优化之差。"),
        ("p", "结论：在存在相关冲击的市场环境下，“逐年滚动优化”的弹性机制本身就是收益来源，"
              "建议该乡村以年度为周期滚动重解模型。"),
        ("fig", "F05_问题三对比.png", "图5　8 组相关模拟下沿用方案与重优化的净收益对比"),
    ],
    "6 敏感性分析": [
        ("p", "两项实证检验：其一，20 组无相关抽样情景下，沿用中枢方案的收益标准差为均值的"
              " 1.7%，重优化增益仅 0.2%，方案在题给参数区间内稳健（图 4）；其二，8 组含共同"
              "因子的相关情景下，固定方案的损失被放大至 10.4%（图 5）——稳健性结论只在参数"
              "独立波动时成立，存在系统性冲击时必须配合滚动重优化。"),
        ("p", "第三项为约束消融：去除豆类三年轮作约束后重解问题一，最优净收益与基准完全相同"
              "（78 164 150 元，见支撑材料 F06_豆类约束代价.csv）——该约束实测代价为零，最优"
              "方案本来就会种植豆类；真正起约束作用的是重茬与地块适配限制。这说明轮作政策对本"
              "乡村不构成额外负担，“三年一种豆”可以无成本地纳入村规民约。价格区间取中值的口径"
              "可用区间端点重跑复现（附录 A 命令），影响方向对称，不改变方案结构。"),
    ],
    "7 模型的评价与推广": [
        ("p", "优点：约束体系完整覆盖题目全部硬性要求并给出可核验的线性表达；两种超产情形在"
              "同一框架下精确求解、结果可复现；不确定性采用“中枢求解—情景评估—重优化对比”"
              "三层递进，兼顾可提交性与风险量化；以消融实验检验了每类约束的实际约束力。"),
        ("p", "缺点与改进：分散度阈值（0.5 亩/8 地块）为主观量化，改进方向是阈值参数化后做"
              "二维灵敏度；预期销售量沿用 2023 年产量口径，若获得市场调研数据，仅需替换需求"
              "生成方式，模型结构不变；两阶段随机规划可更严格处理“重优化价值”，但计算代价高。"),
        ("p", "推广：模型框架可直接迁移到设施农业茬口安排、多仓多期补货等“多周期产能分配”类"
              "问题；“中枢求解 + 情景检验”的两层结构同样适用于其他含不确定性的整数规划问题。"),
    ],
    "参考文献": [
        ("ref", "[1] 司守奎, 孙兆亮. 数学建模算法与应用[M]. 2版. 北京: 国防工业出版社, 2015."),
        ("ref", "[2] 姜启源, 谢金星, 叶俊. 数学模型[M]. 5版. 北京: 高等教育出版社, 2018."),
        ("ref", "[3] 《运筹学》教材编写组. 运筹学[M]. 4版. 北京: 清华大学出版社, 2012."),
        ("ref", "[4] Huangfu Q, Hall J A J. Parallelizing the dual revised simplex method[J]. "
                "Mathematical Programming Computation, 2018, 10(1): 119-142."),
    ],
    "附录A 支撑材料清单": [],
    "附录B 源程序": [
        ("p", "以下为全部求解与装配源程序（Python 3.12，依赖 numpy/pandas/scipy/openpyxl/"
              "python-docx）；复现命令：python load_data.py → python solve_q1.py → "
              "python fill_result.py → python solve_q2q3.py。论文装配脚本 build_paper.py、"
              "fill_paper.py 见支撑材料包。"),
        ("code", "load_data.py"),
        ("code", "solve_q1.py"),
        ("code", "fill_result.py"),
        ("code", "solve_q2q3.py"),
    ],
}

HEADINGS = {h.replace("\u3000", " ") for h in CONTENT}


def norm(t):
    return " ".join(str(t).split()).replace("\u3000", " ").strip()


def set_text(par, text):
    if par.runs:
        par.runs[0].text = text
        for r in par.runs[1:]:
            r.text = ""
    else:
        par.add_run(text)


def main():
    doc = docx.Document(str(SRC))
    body_paras = doc.paragraphs

    # ---- 定位各节锚点 ----
    anchors = []      # (index, norm_heading)
    for i, p in enumerate(body_paras):
        h = norm(p.text)
        if h in HEADINGS:
            anchors.append((i, h))
    found = {h for _, h in anchors}
    missing = HEADINGS - found
    if missing:
        raise SystemExit(f"模板中找不到章节：{missing}")

    # ---- 逐节重写 ----
    proto_by_heading = {}
    for idx, (i, h) in enumerate(anchors):
        end = anchors[idx + 1][0] if idx + 1 < len(anchors) else len(body_paras)
        rng = body_paras[i + 1:end]
        if not rng:
            continue
        proto = copy.deepcopy(rng[0]._p)          # 样式原型
        for p in rng:                              # 清空节内段落（表格保留）
            p._p.getparent().remove(p._p)
        proto_by_heading[h] = (proto, body_paras[i], end)

    def insert_blocks(cursor_el, proto_p, blocks):
        for blk in blocks:
            new_p = copy.deepcopy(proto_p)
            cursor_el.addnext(new_p)
            cursor_el = new_p
            par = docx.text.paragraph.Paragraph(new_p, body_paras[0]._parent)
            kind = blk[0]
            if kind == "p":
                set_text(par, blk[1])
            elif kind == "list":
                set_text(par, blk[1])
            elif kind == "ref":
                set_text(par, blk[1])
            elif kind == "fig":
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = par.add_run()
                run.add_picture(str(FIGS / blk[1]), width=Inches(5.9))
            elif kind == "cap":
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_text(par, blk[1])
                for r in par.runs:
                    r.font.bold = True
                    r.font.size = Pt(10.5)
            elif kind == "code":
                path = HERE / blk[1]
                par.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pf = par.paragraph_format
                pf.first_line_indent = Pt(0)
                pf.line_spacing = 1.0
                lines = path.read_text(encoding="utf-8").splitlines() or [""]
                for ln in lines:
                    run = par.add_run(ln)
                    run.font.name = "Consolas"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
                    run.font.size = Pt(9)
                    run.add_break()   # 逐行换行（WD_BREAK.LINE 默认）
        return cursor_el

    # fig 块后自动加题注
    def expand(blocks):
        out = []
        for b in blocks:
            out.append(b)
            if b[0] == "fig":
                out.append(("cap", b[2]))
        return out

    from docx.text.paragraph import Paragraph
    for idx, (i, h) in enumerate(anchors):
        if h not in CONTENT:
            continue
        blocks = expand(CONTENT[h])
        if not blocks:
            continue
        proto, _head_par, _end = proto_by_heading[h]
        cursor = _head_par._p
        rest = blocks
        # 5.1.2：第二段放公式表之后
        if h == "5.1.2 模型建立":
            pre = [b for b in rest if b[0] == "p"][:1]
            rest_mid = [b for b in rest if b not in pre]
            for blk in pre:
                new_p = copy.deepcopy(proto)
                cursor.addnext(new_p)
                cursor = new_p
                set_text(Paragraph(new_p, body_paras[0]._parent), blk[1])
            # 找节内公式表
            tbl_el = None
            el = cursor
            while el is not None:
                el = el.getnext()
                if el is None or el.tag.endswith("}tbl"):
                    tbl_el = el
                    break
                pp = Paragraph(el, body_paras[0]._parent) if el.tag.endswith("}p") else None
                if pp is not None and norm(pp.text) in HEADINGS:
                    break
            base = tbl_el if tbl_el is not None else cursor
            cursor = base
            for blk in rest_mid:
                new_p = copy.deepcopy(proto)
                cursor.addnext(new_p)
                cursor = new_p
                par = Paragraph(new_p, body_paras[0]._parent)
                set_text(par, blk[1])
            continue
        insert_blocks(cursor, proto, rest)

    # ---- 摘要/标题/关键词 ----
    for par in doc.paragraphs:
        t = par.text.strip()
        if t.startswith("【论文标题"):
            set_text(par, TITLE)
        elif t.startswith("【提示】五要素"):
            set_text(par, "")
        elif t.startswith("本文研究"):
            set_text(par, ABSTRACT[0])
        elif t.startswith("针对问题一"):
            set_text(par, ABSTRACT[1])
        elif t.startswith("针对问题二"):
            set_text(par, ABSTRACT[2])
        elif t.startswith("针对问题三"):
            set_text(par, ABSTRACT[3])
        elif t.startswith("本文的亮点在于"):
            set_text(par, ABSTRACT[4])
        elif "【关键词1" in t:
            for r in par.runs:
                if "【关键词1" in r.text:
                    r.text = KEYWORDS
        elif t.startswith("（电子版自本页开始"):
            set_text(par, "")
        elif t.startswith("【提示】按官方"):
            set_text(par, "AI 使用说明：本文使用人工智能工具辅助文字整理与代码调试，"
                          "全部模型、数据与结论经参赛队独立复核，参赛队对原创性、真实性、"
                          "准确性负全责。")

    # ---- 附录A 表格更新 ----
    files = ["load_data.py", "solve_q1.py", "fill_result.py", "solve_q2q3.py",
             "fill_paper.py", "build_paper.py", "data.json", "plan_case1.csv",
             "plan_case2.csv", "plan_q2.csv", "Q1_结论.csv", "Q2Q3_结论.json"]
    descs = ["数据清洗与口径统一", "问题一 MILP 求解", "结果表回填与校验", "问题二三情景与模拟",
             "摘要页填充", "论文整卷装配", "清洗后数据", "问题一方案（滞销）",
             "问题一方案（五折）", "问题二方案", "问题一结论", "问题二三结论"]
    for tbl in doc.tables:
        head = norm(tbl.rows[0].cells[0].text)
        if "文件" in head:
            need = len(files)
            while len(tbl.rows) - 1 < need:
                tbl.add_row()
            for i, (f, d) in enumerate(zip(files, descs), start=1):
                cells = tbl.rows[i].cells
                cells[0].text = f
                cells[1].text = d
                p = f"{HERE / f}"
                cells[2].text = f"{p and (os.path.getsize(p) // 1024)}KB" if os.path.exists(p) else "—"
            break

    doc.save(str(DST))
    d2 = docx.Document(str(DST))
    print(f"OK -> {DST}")
    print(f"段落 {len(d2.paragraphs)} | 表格 {len(d2.tables)} | 内嵌图片 {len(d2.inline_shapes)}")


if __name__ == "__main__":
    main()
