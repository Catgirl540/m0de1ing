# -*- coding: utf-8 -*-
"""向 发析生物ppt0901.pptx 插入两页新幻灯片：
A. 市场分析 TAM/SAM/SOM 三层市场模型（插在第1页后）
B. 商业模式 三段式获客漏斗（追加在末尾）
风格完全继承原 deck：母版深蓝背景 + 金色强调 + 原页眉页脚。
"""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION

SRC = r"C:\Users\yan\Desktop\发析生物ppt0901.pptx"
OUT = r"C:\Users\yan\Desktop\发析生物ppt0902.pptx"

# ---- palette (sampled from deck) ----
GOLD_HEAD = RGBColor(0xFF, 0xC0, 0x00)   # 标题金
GOLD_TAG  = RGBColor(0xFF, 0xD0, 0x41)   # 标签金
GOLD_BODY = RGBColor(0xFA, 0xE8, 0x9C)   # 正文金
GOLD_CHART= RGBColor(0xFF, 0xE7, 0x44)   # 图表金
YELLOW    = RGBColor(0xFF, 0xFF, 0x00)   # 行内高亮
GOLD_DEEP = RGBColor(0xC9, 0xA9, 0x4E)   # 深金
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
MUTED     = RGBColor(0xC7, 0xD3, 0xEE)   # 注释淡蓝白
NAVY      = RGBColor(0x13, 0x2B, 0x5E)   # 金底上的深蓝字
CARD_LINE = RGBColor(0x4F, 0x8F, 0xE0)   # 卡片描边蓝
CARD_FILL = RGBColor(0x0B, 0x2A, 0x6B)   # 卡片底色(半透明)
BLUE1     = RGBColor(0x12, 0x3C, 0x8C)   # 漏斗1
BLUE2     = RGBColor(0x1D, 0x55, 0xB8)   # 漏斗2
SEP_LINE  = RGBColor(0x3C, 0x5A, 0xA6)   # 分隔细线

F_HEAD = "微软雅黑"
F_HEAVY = "思源宋体 CN Heavy"
F_CARD = "方正公文黑体"

def set_run(r, text, name, size, bold=False, color=WHITE):
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = name
    rPr = r._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", name)

def add_text(slide, x, y, w, h, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             wrap=True, spacing=1.0):
    """paras: list of (runs, opts); runs = list of (text,name,size,bold,color)"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, (runs, opts) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = opts.get("align", align)
        if spacing != 1.0:
            p.line_spacing = spacing
        if opts.get("space_after"):
            p.space_after = Pt(opts["space_after"])
        for spec in runs:
            set_run(p.add_run(), *spec)
    return tb

def add_box(slide, x, y, w, h, line=None, line_w=1.25, fill=None, fill_alpha=None,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sp.adjustments[0] = radius
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
        if fill_alpha is not None:
            clr = sp.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
            a = clr.makeelement(qn("a:alpha"), {"val": str(int(fill_alpha * 1000))})
            clr.append(a)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp

def add_hline(slide, x, y, w, color=SEP_LINE, weight=0.75):
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Emu(9525))
    ln.fill.solid(); ln.fill.fore_color.rgb = color
    ln.line.fill.background(); ln.shadow.inherit = False
    return ln

def clone_shapes(src_slide, dst_slide, keep_ids):
    for sp in src_slide.shapes:
        if sp.shape_id not in keep_ids:
            continue
        el = copy.deepcopy(sp._element)
        for cust in el.findall(".//" + qn("p:custDataLst")):
            cust.getparent().remove(cust)   # tags 部件不随克隆，防止悬空 rId
        for blip in el.iter(qn("a:blip")):
            rid = blip.get(qn("r:embed"))
            if rid:
                part = src_slide.part.rels[rid].target_part
                blip.set(qn("r:embed"), dst_slide.part.relate_to(
                    part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"))
        dst_slide.shapes._spTree.append(el)

def replace_text(slide, shape_id, text):
    for sp in slide.shapes:
        if sp.shape_id == shape_id:
            p = sp.text_frame.paragraphs[0]
            runs = p.runs
            runs[0].text = text
            for r in runs[1:]:
                r._r.getparent().remove(r._r)
            return sp

def fix_chart_ea(chart, name=F_HEAD):
    for defRPr in chart._chartSpace.iter(qn("a:defRPr")):
        for tag in ("a:ea", "a:cs"):
            e = defRPr.find(qn(tag))
            if e is None:
                e = defRPr.makeelement(qn(tag), {})
                defRPr.append(e)
            e.set("typeface", name)

def chart_noFill(chart):
    cs = chart._chartSpace
    spPr = cs.makeelement(qn("c:spPr"), {})
    cs.insert(list(cs).index(cs.find(qn("c:chart"))) + 1, spPr)
    spPr.append(spPr.makeelement(qn("a:noFill"), {}))
    ln = spPr.makeelement(qn("a:ln"), {})
    ln.append(ln.makeelement(qn("a:noFill"), {}))
    spPr.append(ln)

prs = Presentation(SRC)
s0, s2 = prs.slides[0], prs.slides[2]
layout = s0.slide_layout

# ================= SLIDE A : TAM / SAM / SOM =================
A = prs.slides.add_slide(layout)
clone_shapes(s0, A, {89, 7, 9, 10, 2, 19})   # 页眉发光图/渐变条/副标题/章节标签/校徽/来源行
replace_text(A, 19, "数据来源：《2025中国头皮健康消费白皮书》、弗若斯特沙利文；SAM/SOM 为团队测算")

add_text(A, 0.33, 0.92, 12.7, 0.71, [([
    ("市场规模测算：", F_HEAD, 32, True, GOLD_HEAD),
    ("TAM·SAM·SOM", F_HEAD, 32, True, YELLOW),
    (" 三层锁定切入路径", F_HEAD, 32, True, GOLD_HEAD),
], {})], anchor=MSO_ANCHOR.MIDDLE)

# 左：同心圆示意
add_box(A, 0.75, 1.95, 4.55, 4.55, line=GOLD_TAG, line_w=2.25, shape=MSO_SHAPE.OVAL, radius=None)
add_box(A, 1.90, 3.35, 2.85, 2.85, line=GOLD_CHART, line_w=1.75, shape=MSO_SHAPE.OVAL, radius=None)
add_box(A, 2.62, 4.62, 1.42, 1.42, fill=GOLD_TAG, shape=MSO_SHAPE.OVAL, radius=None)
add_text(A, 1.00, 2.30, 4.05, 0.85, [
    ([("TAM", F_HEAD, 17, True, GOLD_TAG), ("  中国头皮健康市场", F_CARD, 12, False, WHITE)], {"align": PP_ALIGN.CENTER}),
    ([("≈950 亿元", F_HEAD, 22, True, YELLOW)], {"align": PP_ALIGN.CENTER}),
])
add_text(A, 1.95, 3.60, 2.75, 0.70, [
    ([("SAM", F_HEAD, 14, True, GOLD_CHART), ("  B端研发检测", F_CARD, 11.5, False, WHITE)], {"align": PP_ALIGN.CENTER}),
    ([("≈85 亿元", F_HEAD, 18, True, YELLOW)], {"align": PP_ALIGN.CENTER}),
])
add_text(A, 2.62, 5.00, 1.42, 0.66, [
    ([("SOM", F_HEAD, 13, True, NAVY)], {"align": PP_ALIGN.CENTER}),
    ([("≈0.2 亿元", F_HEAD, 14, True, NAVY)], {"align": PP_ALIGN.CENTER}),
])
add_text(A, 0.75, 6.60, 4.6, 0.30, [([("示意图：同心圆仅示层级关系，非等比例", F_CARD, 12, False, MUTED)], {})])

# 右：三行定义（无容器分组 + 细分隔线）
rows = [
    (1.74, "TAM", "总可及市场", "≈950 亿元", "100%", "2026E 中国头皮健康市场：防脱美妆 / 脱发药物 / 头皮管理全盘"),
    (3.02, "SAM", "可服务市场", "≈85 亿元", "≈9%", "可被毛囊类器官替代的B端研发与检测支出，按产业研发检测投入约 9% 测算"),
    (4.30, "SOM", "可获得市场", "≈0.2 亿元", "3年累计", "前3年服务+产品收入测算：首年 5–8 家 × 8–15 万/单，复购 66% 滚动放量"),
]
for y, tag, cn, num, pct, desc in rows:
    add_text(A, 6.35, y, 6.55, 0.5, [([
        (tag, F_HEAD, 20, True, GOLD_TAG), ("  " + cn, F_HEAVY, 16, True, WHITE),
        ("　", F_HEAD, 16, True, WHITE), (num, F_HEAD, 20, True, YELLOW),
        ("　" + pct, F_CARD, 12.5, False, MUTED),
    ], {})])
    add_text(A, 6.35, y + 0.52, 6.55, 0.30, [([(desc, F_CARD, 12.5, False, WHITE)], {})])
add_hline(A, 6.35, 2.88, 6.55)
add_hline(A, 6.35, 4.16, 6.55)

# 右下：SAM 构成横向条形图（原生图表）
add_text(A, 6.35, 5.22, 6.5, 0.28, [([
    ("SAM ≈85 亿元 构成测算", F_HEAVY, 13.5, True, GOLD_BODY), ("（亿元，示意）", F_CARD, 12, False, MUTED)], {})])
cd = CategoryChartData()
cd.categories = ["防脱美妆功效评价", "药物临床前筛选", "科研试剂与服务"]
cd.add_series("SAM构成", (45, 28, 12))
gf = A.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(6.30), Inches(5.50), Inches(6.60), Inches(1.00), cd)
ch = gf.chart
ch.has_legend = False
ch.has_title = False
plot = ch.plots[0]
plot.gap_width = 60
plot.has_data_labels = True
dl = plot.data_labels
dl.position = XL_LABEL_POSITION.OUTSIDE_END
dl.font.size = Pt(12); dl.font.bold = True; dl.font.name = F_HEAD
dl.font.color.rgb = WHITE
ser = plot.series[0]
for i, c in enumerate([GOLD_CHART, GOLD_TAG, GOLD_DEEP]):
    pt = ser.points[i]
    pt.format.fill.solid(); pt.format.fill.fore_color.rgb = c
cax = ch.category_axis
cax.has_major_gridlines = False
cax.tick_labels.font.size = Pt(12); cax.tick_labels.font.name = F_HEAD
cax.tick_labels.font.color.rgb = WHITE
cax.format.line.color.rgb = SEP_LINE
vax = ch.value_axis
vax.has_major_gridlines = False
vax.maximum_scale = 55; vax.minimum_scale = 0
d = vax._element.find(qn("c:delete"))
if d is None:
    d = vax._element.makeelement(qn("c:delete"), {"val": "1"})
    vax._element.find(qn("c:scaling")).addnext(d)   # schema: axId, scaling, delete
else:
    d.set("val", "1")
fix_chart_ea(ch)
chart_noFill(ch)

# 右下结论
add_text(A, 6.35, 6.62, 6.55, 0.45, [([
    ("先切 ", F_HEAVY, 18, False, GOLD_BODY), ("SAM 头部场景", F_HEAVY, 18, True, YELLOW),
    ("，以标杆客户向 ", F_HEAVY, 18, False, GOLD_BODY), ("TAM", F_HEAVY, 18, True, YELLOW),
    (" 延展", F_HEAVY, 18, False, GOLD_BODY),
], {})])

# ================= SLIDE B : 获客漏斗 =================
B = prs.slides.add_slide(layout)
clone_shapes(s2, B, {2, 3, 7, 9, 20})   # 页眉发光图/渐变条/章节标签(商业模式)/副标题/校徽
replace_text(B, 9, "营销与获客——三段式漏斗精准转化 B 端客户")

add_text(B, 0.33, 0.92, 12.7, 0.71, [([
    ("营销与获客：", F_HEAD, 32, True, GOLD_HEAD),
    ("三段式漏斗", F_HEAD, 32, True, YELLOW),
    ("，从 200–300 家触达到 8 家深度绑定", F_HEAD, 32, True, GOLD_HEAD),
], {})], anchor=MSO_ANCHOR.MIDDLE)

# 漏斗三段（chevron 互锁）
add_box(B, 0.40, 1.90, 4.30, 1.50, fill=BLUE1, line=CARD_LINE, line_w=1.0, shape=MSO_SHAPE.PENTAGON, radius=None)
add_box(B, 4.45, 1.90, 4.00, 1.50, fill=BLUE2, line=CARD_LINE, line_w=1.0, shape=MSO_SHAPE.CHEVRON, radius=None)
add_box(B, 8.20, 1.90, 3.55, 1.50, fill=GOLD_TAG, line=None, shape=MSO_SHAPE.CHEVRON, radius=None)
add_text(B, 0.75, 2.02, 3.10, 1.26, [
    ([("① 认知触达 ", F_HEAD, 15, True, WHITE), ("Top of Funnel", F_CARD, 11, False, MUTED)], {"space_after": 2}),
    ([("目标池 200–300 家", F_HEAD, 19, True, YELLOW)], {"space_after": 2}),
    ([("头部美妆日化 / 药企 / 科研机构", F_CARD, 11.5, False, WHITE)], {}),
])
add_text(B, 5.35, 2.02, 2.55, 1.26, [
    ([("② 意向培育 ", F_HEAD, 15, True, WHITE), ("Middle", F_CARD, 11, False, MUTED)], {"space_after": 2}),
    ([("有效线索 50–80 家", F_HEAD, 17, True, YELLOW)], {"space_after": 2}),
    ([("主动询价 / 提供研发痛点", F_CARD, 11.5, False, WHITE)], {}),
])
add_text(B, 9.00, 2.02, 2.10, 1.26, [
    ([("③ 转化复购 ", F_HEAD, 15, True, NAVY), ("Bottom", F_CARD, 11, False, NAVY)], {"space_after": 2}),
    ([("首年签约 5–8 家", F_HEAD, 15, True, NAVY)], {"space_after": 2}),
    ([("客单价 8–15 万/单", F_HEAD, 13, True, NAVY)], {}),
])
# 转化率 chips
add_text(B, 3.55, 3.47, 2.00, 0.30, [([("线索转化 ≈25% ▼", F_HEAD, 13, True, GOLD_CHART)], {"align": PP_ALIGN.CENTER})])
add_text(B, 7.60, 3.47, 2.00, 0.30, [([("商务转化 ≈10% ▼", F_HEAD, 13, True, GOLD_CHART)], {"align": PP_ALIGN.CENTER})])
# 复购徽章
add_box(B, 11.80, 2.02, 1.22, 1.22, fill=GOLD_TAG, line=None, shape=MSO_SHAPE.OVAL, radius=None)
add_text(B, 11.80, 2.32, 1.22, 0.66, [
    ([("66%", F_HEAD, 20, True, NAVY)], {"align": PP_ALIGN.CENTER}),
    ([("复购目标", F_CARD, 11.5, True, NAVY)], {"align": PP_ALIGN.CENTER}),
])

# 三张手段卡
cards = [
    (0.40, 4.30, "营销触达组合", [
        [("行业展会：", GOLD_BODY, True), ("In-Cosmetics 原料展触达", WHITE, False)],
        [("学术论坛：", GOLD_BODY, True), ("皮肤科/再生医学专家背书", WHITE, False)],
        [("行业白皮书：", GOLD_BODY, True), ("联合协会发布建立信任", WHITE, False)],
        [("线上阵地：", GOLD_BODY, True), ("科研电商 + 行业媒体内容", WHITE, False)],
    ], "行业曝光 · 总监触达 · 需求验证"),
    (4.85, 3.90, "免费小试培育", [
        [("免费小试/低价小样", YELLOW, True), ("，降低尝鲜门槛", WHITE, False)],
        [("交付毛囊生长数据", YELLOW, True), ("，验证可行性", WHITE, False)],
        [("打消「100天未见效」", YELLOW, True), ("决策顾虑", WHITE, False)],
        [("科研电商承接下单", YELLOW, True), ("，转化即时跟进", WHITE, False)],
    ], "有效线索 50–80 家 · 询价/痛点"),
    (8.90, 3.45, "转化与生态锁定", [
        [("首年签约 5–8 家", YELLOW, True), ("，客单 8–15 万", WHITE, False)],
        [("数据入 AI 模型", YELLOW, True), (" + 原料备案库", WHITE, False)],
        [("使用数据反哺模型", YELLOW, True), ("，迭代优化", WHITE, False)],
        [("单次采购→年度订阅", YELLOW, True), ("，复购 66%", WHITE, False)],
    ], "复购率 · 续约金额 · 生态绑定"),
]
for x, w, title, bullets, metrics in cards:
    add_box(B, x, 3.90, w, 2.42, line=CARD_LINE, line_w=1.25, fill=CARD_FILL, fill_alpha=38, radius=0.05)
    add_text(B, x + 0.20, 4.06, w - 0.40, 0.32, [([(title, F_HEAVY, 15, True, GOLD_TAG)], {})])
    paras = []
    for lead, rest in bullets:
        runs = [(lead[0], F_CARD, 12.5, lead[2], lead[1]), (rest[0], F_CARD, 12.5, rest[2], rest[1])]
        paras.append((runs, {"space_after": 5}))
    add_text(B, x + 0.20, 4.46, w - 0.40, 1.28, paras, spacing=1.05)
    add_hline(B, x + 0.20, 5.82, w - 0.40)
    add_text(B, x + 0.20, 5.92, w - 0.40, 0.30, [([
        ("指标｜", F_CARD, 12, True, GOLD_BODY), (metrics, F_CARD, 12, False, MUTED)], {})])

# 底部获客逻辑总结
add_text(B, 0.40, 6.62, 12.53, 0.45, [([
    ("获客逻辑：", F_HEAVY, 18, True, GOLD_TAG),
    ("免费小试破冰 → 数据建立信任 → 生态锁定复购", F_HEAVY, 18, False, WHITE),
    ("，把一次性订单变成 ", F_HEAVY, 18, False, WHITE),
    ("年度订阅", F_HEAVY, 18, True, YELLOW),
], {})], align=PP_ALIGN.CENTER)
add_text(B, 0.33, 7.18, 12.6, 0.26, [([
    ("口径说明：目标池 / 线索 / 签约数为第 1 年经营目标（团队测算），转化率 ≈25% / ≈10% 为目标值，复购 66% 为生态锁定目标", F_CARD, 12, False, MUTED)], {})])

# ================= 插入位置：A 移到第2页 =================
lst = prs.slides._sldIdLst
ids = list(lst)
a_id = ids[3]          # A 追加在末尾 index 3
lst.remove(a_id)
lst.insert(1, a_id)    # 位于原市场分析页之后

prs.save(OUT)
print("SAVED", OUT, "slides:", len(prs.slides._sldIdLst))
