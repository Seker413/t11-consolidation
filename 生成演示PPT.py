"""
生成演示PPT V2 (.pptx) — 专业深蓝主题
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from datetime import datetime
from pathlib import Path

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ── 颜色方案 ──
DARK_BLUE = RGBColor(0x1A, 0x3C, 0x6E)
MED_BLUE  = RGBColor(0x2C, 0x5F, 0x9E)
LIGHT_BLUE = RGBColor(0xE8, 0xF0, 0xFE)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
MID_GRAY  = RGBColor(0x66, 0x66, 0x66)
ACCENT_GREEN = RGBColor(0x2E, 0xCC, 0x71)
ACCENT_RED   = RGBColor(0xE7, 0x4C, 0x3C)
ACCENT_ORANGE = RGBColor(0xF3, 0x9C, 0x12)

def add_blank_slide():
    return prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

def add_rect(slide, left, top, width, height, color, alpha=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape

def add_text_box(slide, left, top, width, height, text, size=14, bold=False, color=DARK_GRAY, align=PP_ALIGN.LEFT, font_name='Microsoft YaHei'):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = align
    return txBox

def add_bullet_frame(slide, left, top, width, height, items, size=13, color=DARK_GRAY):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = item
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.name = 'Microsoft YaHei'
        p.space_after = Pt(6)
    return txBox

def add_table_shape(slide, left, top, width, height, headers, rows, col_widths=None):
    n_rows = 1 + len(rows)
    n_cols = len(headers)
    table_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    table = table_shape.table
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w
    # Header
    for ci, h in enumerate(headers):
        cell = table.cell(0, ci)
        cell.text = h
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.font.name = 'Microsoft YaHei'
            p.alignment = PP_ALIGN.CENTER
        cell.fill.solid()
        cell.fill.fore_color.rgb = DARK_BLUE
    # Rows
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.text = str(val)
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(10)
                p.font.color.rgb = DARK_GRAY
                p.font.name = 'Microsoft YaHei'
            if ri % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = LIGHT_BLUE
    return table_shape

def add_footer(slide, text="DRP T11 · 合并报表与一键出表 · 内部存货交易未实现利润自动抵消系统"):
    add_rect(slide, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4), DARK_BLUE)
    add_text_box(slide, Inches(0.5), Inches(7.12), Inches(12), Inches(0.35), text, size=9, color=WHITE)

def add_slide_title(slide, title, subtitle=None):
    add_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(1.2), DARK_BLUE)
    add_text_box(slide, Inches(0.6), Inches(0.2), Inches(12), Inches(0.7), title, size=28, bold=True, color=WHITE)
    if subtitle:
        add_text_box(slide, Inches(0.6), Inches(0.75), Inches(12), Inches(0.4), subtitle, size=13, color=RGBColor(0xBB, 0xCC, 0xDD))

# ============================================================
# SLIDE 1: 封面
# ============================================================
s = add_blank_slide()
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), DARK_BLUE)
add_rect(s, Inches(0), Inches(3.2), Inches(13.333), Inches(0.06), ACCENT_GREEN)
add_text_box(s, Inches(1.5), Inches(1.0), Inches(10.5), Inches(1.2),
    '基于AI规则引擎的\n内部存货交易未实现利润自动抵消系统', size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text_box(s, Inches(1.5), Inches(3.6), Inches(10.5), Inches(0.8),
    'DRP系统 T11 合并报表与一键出表 — 子模块', size=20, color=RGBColor(0xAA, 0xBB, 0xCC), align=PP_ALIGN.CENTER)
add_text_box(s, Inches(1.5), Inches(4.8), Inches(10.5), Inches(0.6),
    '中国社会科学院大学 · 2025级会计专硕二班 · 尚浩然、李金阳', size=16, color=RGBColor(0x88, 0x99, 0xAA), align=PP_ALIGN.CENTER)
add_text_box(s, Inches(1.5), Inches(5.4), Inches(10.5), Inches(0.6),
    f'指导老师：张金昌  |  {datetime.now().strftime("%Y年%m月")}', size=14, color=RGBColor(0x77, 0x88, 0x99), align=PP_ALIGN.CENTER)

# ============================================================
# SLIDE 2: 项目背景
# ============================================================
s = add_blank_slide()
add_slide_title(s, '项目背景与业务痛点', 'DRP平台 T11 合并报表  —  为什么需要这个系统')
add_bullet_frame(s, Inches(0.8), Inches(1.6), Inches(5.5), Inches(5.0), [
    '📋  政策驱动',
    '  • 国资委1号文：应用AI大模型、OCR技术实现管理智能化',
    '  • 国资委2号文：数据自动采集、模型自动分析、风险自动预警',
    '  • CAS33：合并报表必须抵消内部交易未实现利润',
    '',
    '📊  业务规模',
    '  • 集团型企业数十家法人主体',
    '  • 月度内部交易数百笔',
    '  • 涉及顺流/逆流/平流三种交易类型',
], size=14, color=DARK_GRAY)

# 右侧：痛点
add_rect(s, Inches(7.0), Inches(1.5), Inches(5.8), Inches(5.3), LIGHT_BLUE)
add_text_box(s, Inches(7.3), Inches(1.7), Inches(5.2), Inches(0.5), '🔴 传统手工抵消的四大痛点', size=16, bold=True, color=DARK_BLUE)
add_bullet_frame(s, Inches(7.3), Inches(2.3), Inches(5.2), Inches(4.2), [
    '① 工作量大',
    '   数十家主体、数百笔交易，Excel逐笔勾对耗时漫长',
    '',
    '② 易出错',
    '   顺流/逆流/平流抵消比例不同，人工计算容易遗漏',
    '',
    '③ 追溯困难',
    '   审计时无法完整还原每笔抵消的计算过程',
    '',
    '④ 监管脱节',
    '   手工抵消无法满足国资委穿透式监管要求',
], size=12, color=DARK_GRAY)
add_footer(s)

# ============================================================
# SLIDE 3: 解决方案
# ============================================================
s = add_blank_slide()
add_slide_title(s, '解决方案概述', 'AI规则引擎 + 自动计算 + 智能校验  —  核心流程')
# 流程步骤
steps = ['① 上传Excel\n三表导入', '② 智能识别\n交易类型', '③ 自动计算\n未实现利润', '④ 生成分录\nJSON规则引擎', '⑤ 试算平衡\n借贷验证', '⑥ 导出报告\nExcel/PDF']
for i, step in enumerate(steps):
    x = Inches(0.8 + i * 2.1)
    add_rect(s, x, Inches(2.0), Inches(1.8), Inches(1.4), MED_BLUE if i % 2 == 0 else DARK_BLUE)
    add_text_box(s, x + Inches(0.1), Inches(2.2), Inches(1.6), Inches(1.0), step, size=11, color=WHITE, align=PP_ALIGN.CENTER)
    if i < 5:
        add_text_box(s, x + Inches(1.8), Inches(2.45), Inches(0.3), Inches(0.5), '→', size=18, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)

add_bullet_frame(s, Inches(0.8), Inches(3.8), Inches(5.8), Inches(2.8), [
    '✅ 覆盖场景',
    '  • 顺流交易（母→子）：未实现利润全部归母',
    '  • 逆流交易（子→母）：少数股东按持股比例分摊',
    '  • 平流交易（子↔子）：按销售方少数股东比例分摊',
    '  • 递延所得税自动计算（25%税率可调节）',
], size=13)
add_bullet_frame(s, Inches(7.0), Inches(3.8), Inches(5.8), Inches(2.8), [
    '⚙️ 技术特点',
    '  • JSON可配置规则引擎（非硬编码）',
    '  • 本地SHA256哈希链保证可追溯',
    '  • Python + Streamlit，极简部署一键运行',
    '  • 模块化架构：engine / utils / templates',
], size=13)
add_footer(s)

# ============================================================
# SLIDE 4: 系统架构
# ============================================================
s = add_blank_slide()
add_slide_title(s, '系统架构（V2.0 模块化）', 'Python + Streamlit + SQLite + Matplotlib')
# 架构图用文本框模拟
layers = [
    ('Streamlit Web 界面', '文件上传 | 参数配置 | 结果展示 | 仪表盘 | 导出', MED_BLUE, Inches(1.5)),
    ('核心计算引擎 (engine/)', '数据校验 → 类型识别 → 利润计算 → 分录生成 → 哈希存证', DARK_BLUE, Inches(2.8)),
    ('规则配置 (rules.json)', '三套抵消模板（DOWN/UP/FLAT）| 安全公式求值 | 声明式规则', LIGHT_BLUE, Inches(4.1)),
    ('数据层 (SQLite + Excel)', '股权配置表 | 交易明细表 | 抵消分录表 | calc_log 计算日志', RGBColor(0xF0, 0xF0, 0xF0), Inches(5.4)),
]
for label, desc, color, y in layers:
    add_rect(s, Inches(0.8), y, Inches(11.5), Inches(1.0), color)
    text_color = WHITE if color in (MED_BLUE, DARK_BLUE) else DARK_GRAY
    add_text_box(s, Inches(1.0), y + Inches(0.05), Inches(3.5), Inches(0.4), label, size=14, bold=True, color=text_color)
    add_text_box(s, Inches(1.0), y + Inches(0.5), Inches(10.5), Inches(0.4), desc, size=11, color=text_color)
    add_text_box(s, Inches(0.5), y + Inches(0.25), Inches(0.25), Inches(0.3), '⬆', size=10, color=DARK_GRAY)

add_text_box(s, Inches(0.8), Inches(6.5), Inches(11.5), Inches(0.4),
    '技术栈：Python 3.11+ | Streamlit | pandas | openpyxl | SQLite | Matplotlib | ReportLab | Git', size=11, color=MID_GRAY)
add_footer(s)

# ============================================================
# SLIDE 5: 核心算法 — 利润计算
# ============================================================
s = add_blank_slide()
add_slide_title(s, '核心算法：未实现利润计算', '毛利率 × 未售比例 × 销售收入')

add_rect(s, Inches(0.8), Inches(1.6), Inches(5.5), Inches(2.0), LIGHT_BLUE)
add_text_box(s, Inches(1.0), Inches(1.7), Inches(5.0), Inches(0.5), '🔢 核心公式', size=16, bold=True, color=DARK_BLUE)
add_text_box(s, Inches(1.0), Inches(2.2), Inches(5.0), Inches(1.2),
    '未实现利润 = 销售收入 × 毛利率 × 未对外出售比例\n\n'
    '毛利率 = (收入 - 成本) / 收入\n'
    '未售比例 = 期末结存数量 / 内部购入总数量', size=14, color=DARK_GRAY)

# 三种场景
add_text_box(s, Inches(7.0), Inches(1.6), Inches(5.5), Inches(0.5), '📊 三种计算场景', size=16, bold=True, color=DARK_BLUE)
scenarios = [
    ('🟢 全未售（结存=购入）', '全额抵消\n未实现利润 = 收入 - 成本'),
    ('🟡 部分售出', '仅抵消未售部分\n未实现利润 = 收入×毛利率×未售%'),
    ('⚪ 全售出（结存=0）', '已实现，无需抵消\n自动跳过'),
]
for i, (title, desc) in enumerate(scenarios):
    y = Inches(3.8 + i * 1.2)
    add_rect(s, Inches(0.8), y, Inches(5.5), Inches(1.0), WHITE)
    add_text_box(s, Inches(1.0), y + Inches(0.05), Inches(2.5), Inches(0.4), title, size=13, bold=True, color=DARK_BLUE)
    add_text_box(s, Inches(3.6), y + Inches(0.05), Inches(2.5), Inches(0.9), desc, size=12, color=MID_GRAY)

# 右侧示例
add_rect(s, Inches(7.0), Inches(2.3), Inches(5.5), Inches(4.0), DARK_BLUE)
add_text_box(s, Inches(7.3), Inches(2.5), Inches(5.0), Inches(0.5), '💡 计算示例', size=14, bold=True, color=WHITE)
add_text_box(s, Inches(7.3), Inches(3.1), Inches(5.0), Inches(2.5),
    'P(母公司) → S(子公司，持股80%)\n'
    '销售A产品100件，收入100万，成本60万\n'
    'S期末对外售出30件，剩余70件\n\n'
    '毛利率 = (100万-60万)/100万 = 40%\n'
    '未售比例 = 70/100 = 70%\n'
    '未实现利润 = 100万 × 40% × 70%\n'
    '            = 280,000 元', size=13, color=RGBColor(0xDD, 0xEE, 0xFF))
add_footer(s)

# ============================================================
# SLIDE 6: 分录生成
# ============================================================
s = add_blank_slide()
add_slide_title(s, '核心算法：抵消分录自动生成', 'JSON规则引擎驱动 — 三套模板')

add_table_shape(s, Inches(0.5), Inches(1.6), Inches(12.3), Inches(5.0),
    ['交易类型', '方向', '借：营业收入     贷：营业成本     贷：存货', '借：递延所得税资产     贷：所得税费用', '少数股东分摊'],
    [
        ['顺流交易\n(母→子)', 'DOWN', '借：营业收入\n  贷：营业成本（倒挤）\n  贷：存货（未实现利润）', '借：递延所得税资产\n  贷：所得税费用\n  (未实现利润×25%)', '❌ 不涉及\n（利润全归母）'],
        ['逆流交易\n(子→母)', 'UP', '同上', '同上', '借：少数股东权益\n  贷：少数股东损益\n  (利润×(1-25%)×少数%)'],
        ['平流交易\n(子↔子)', 'FLAT', '同上', '同上', '同上\n（按销售方少数股东%）'],
    ],
    col_widths=[Inches(2.0), Inches(1.0), Inches(4.0), Inches(3.0), Inches(2.3)]
)
add_footer(s)

# ============================================================
# SLIDE 7: AI规则引擎
# ============================================================
s = add_blank_slide()
add_slide_title(s, 'AI规则引擎设计', '会计逻辑与程序代码解耦 — 规则即配置')

items = [
    ('🧠  智能识别', '基于股权关系图谱的LCA算法，自动判定每笔交易是顺流/逆流/平流，O(1)复杂度'),
    ('📋  规则驱动', '抵消模板以JSON声明式规则存储（rules.json），系统根据规则自动推导分录，而非传统if-else堆砌'),
    ('🔌  可扩展', '新增固定资产内部交易抵消等场景，只需增加JSON配置条目，核心引擎代码无需修改'),
    ('🤖  AI可读', '规则文件可被大模型读取和理解，未来可通过自然语言指令让AI自动生成新规则条目'),
]
for i, (title, desc) in enumerate(items):
    y = Inches(1.8 + i * 1.3)
    add_rect(s, Inches(0.8), y, Inches(11.5), Inches(1.1), LIGHT_BLUE if i % 2 == 0 else WHITE)
    add_text_box(s, Inches(1.0), y + Inches(0.1), Inches(2.5), Inches(0.4), title, size=15, bold=True, color=DARK_BLUE)
    add_text_box(s, Inches(3.6), y + Inches(0.15), Inches(8.5), Inches(0.8), desc, size=13, color=DARK_GRAY)

add_rect(s, Inches(0.8), Inches(7.0), Inches(11.5), Inches(0.06), ACCENT_GREEN)
add_footer(s)

# ============================================================
# SLIDE 8: 仪表盘
# ============================================================
s = add_blank_slide()
add_slide_title(s, '可视化仪表盘与数据分析', 'KPI指标卡 + 图表 + 对比表 + 试算平衡')
dash_items = [
    ('📊 KPI卡片 × 5', '交易额/未实现利润/抵消笔数/冲减存货/净利润影响'),
    ('🚀 快捷操作 × 4', '银行对账/关联交易/合并抵消/生成报告'),
    ('🕐 最近操作', '时间线展示最近8条分录记录'),
    ('📈 图表 × 2', '抵消金额构成柱状图 + 交易类型分布柱状图'),
    ('📋 对比表', '抵消前后：营收/成本/存货/递延所得税/净利润 五项'),
    ('⚖️ 试算平衡', '借方合计/贷方合计/差额 — 自动验证借贷平衡'),
    ('🏷️ 分录筛选', '按交易类型（全部/顺流/逆流/平流）过滤分录'),
    ('🔍 历史追溯', '按批次查看历次计算过程 + SHA256哈希链'),
]
for i, (title, desc) in enumerate(dash_items):
    col = i % 2
    row = i // 2
    x = Inches(0.8 + col * 6.2)
    y = Inches(1.6 + row * 1.35)
    add_rect(s, x, y, Inches(5.8), Inches(1.15), LIGHT_BLUE)
    add_text_box(s, x + Inches(0.2), y + Inches(0.1), Inches(5.4), Inches(0.4), title, size=13, bold=True, color=DARK_BLUE)
    add_text_box(s, x + Inches(0.2), y + Inches(0.55), Inches(5.4), Inches(0.4), desc, size=11, color=MID_GRAY)
add_footer(s)

# ============================================================
# SLIDE 9: 试算平衡 + 手动调整
# ============================================================
s = add_blank_slide()
add_slide_title(s, '试算平衡验证与手动调整', 'V4.0新增功能 — 对标企业级合并报表系统')

add_rect(s, Inches(0.8), Inches(1.6), Inches(5.8), Inches(2.5), LIGHT_BLUE)
add_text_box(s, Inches(1.0), Inches(1.7), Inches(5.4), Inches(0.5), '⚖️ 试算平衡表', size=16, bold=True, color=DARK_BLUE)
add_text_box(s, Inches(1.0), Inches(2.2), Inches(5.4), Inches(1.8),
    '• 自动汇总全部抵消分录的借方与贷方合计\n'
    '• 差额 = 借方合计 - 贷方合计\n'
    '• 差额为0 → ✅ 借贷平衡（正常）\n'
    '• 差额不为0 → ⚠️ 借贷不平（需检查手动调整）\n'
    '• 对称分录模板设计，系统自动生成永远平衡', size=12, color=DARK_GRAY)

add_rect(s, Inches(7.0), Inches(1.6), Inches(5.8), Inches(2.5), LIGHT_BLUE)
add_text_box(s, Inches(7.2), Inches(1.7), Inches(5.4), Inches(0.5), '✏️ 分录手动调整', size=16, bold=True, color=DARK_BLUE)
add_text_box(s, Inches(7.2), Inches(2.2), Inches(5.4), Inches(1.8),
    '• 修改金额：每行分录旁有数字输入框\n'
    '• 删除分录行：点击行尾 🗑️ 按钮\n'
    '• 追加分录：输入借贷科目 + 金额 → 点击添加\n'
    '• 删除整笔：点击卡片底部删除按钮\n'
    '• 人工追加项标注 "✏️ 人工追加" 备注', size=12, color=DARK_GRAY)

add_text_box(s, Inches(0.8), Inches(4.4), Inches(12), Inches(0.5), '📜 历史记录追溯', size=16, bold=True, color=DARK_BLUE)
add_text_box(s, Inches(0.8), Inches(4.9), Inches(12), Inches(1.5),
    '• 所有计算批次自动存入SQLite calc_log表\n'
    '• 按批次分组展示：批次ID / 分录笔数 / 计算时间\n'
    '• 展开查看每笔分录的完整计算过程（卖方/买方/毛利率/未售比例/未实现利润）\n'
    '• 每笔分录附带SHA256哈希值，形成防篡改哈希链 → ELIM-0001 → ELIM-0002 → ELIM-0003 → ...', size=12, color=DARK_GRAY)
add_footer(s)

# ============================================================
# SLIDE 10: 可追溯性
# ============================================================
s = add_blank_slide()
add_slide_title(s, '可追溯性设计', '计算日志 + SHA256哈希链 — 为穿透监管提供技术基础')
add_text_box(s, Inches(0.8), Inches(1.6), Inches(5.8), Inches(0.5), '📝 JSON计算日志（每笔分录同步输出）', size=14, bold=True, color=DARK_BLUE)
add_rect(s, Inches(0.8), Inches(2.2), Inches(5.8), Inches(4.0), RGBColor(0xF5, 0xF5, 0xF5))
add_text_box(s, Inches(1.0), Inches(2.3), Inches(5.5), Inches(3.8),
    '{\n'
    '  "elim_no": "ELIM-202605-0003",\n'
    '  "input": { 卖方, 买方, 收入, 成本, 数量, 结存 },\n'
    '  "classification": { 交易类型, 判定依据 },\n'
    '  "calculation": {\n'
    '    毛利率: 0.40,\n'
    '    未售比例: 0.70,\n'
    '    未实现利润: 280000\n'
    '  },\n'
    '  "entries": [\n'
    '    { D, 营业收入, 700000 },\n'
    '    { C, 营业成本, 420000 },\n'
    '    { C, 存货, 280000 },\n'
    '    ...\n'
    '  ],\n'
    '  "hash": "sha256:a3f8b2..."\n'
    '}', size=10, color=DARK_GRAY)

add_text_box(s, Inches(7.0), Inches(1.6), Inches(5.8), Inches(0.5), '🔗 哈希链机制', size=14, bold=True, color=DARK_BLUE)
add_bullet_frame(s, Inches(7.0), Inches(2.2), Inches(5.8), Inches(3.0), [
    '• 每笔分录的哈希 = SHA256(上一笔哈希 + 本笔数据)',
    '• 形成链式结构：ELIM-0001 → ELIM-0002 → ELIM-0003 → ...',
    '• 类似区块链的简化版，任何篡改都会导致链条断裂',
    '',
    '🔮 未来扩展',
    '• hash字段预留了区块链存证接口',
    '• 对接T33区块链服务时：',
    '  直接将哈希值通过RESTful API上链',
    '  保存链上返回的tx_id',
    '  核心计算逻辑无需任何修改',
], size=12, color=DARK_GRAY)
add_footer(s)

# ============================================================
# SLIDE 11: 完整示例
# ============================================================
s = add_blank_slide()
add_slide_title(s, '完整计算示例', '顺流交易 — 从原始数据到抵消分录的全过程')

add_rect(s, Inches(0.8), Inches(1.6), Inches(3.8), Inches(2.0), LIGHT_BLUE)
add_text_box(s, Inches(1.0), Inches(1.7), Inches(3.4), Inches(0.4), '📋 输入数据', size=14, bold=True, color=DARK_BLUE)
add_text_box(s, Inches(1.0), Inches(2.1), Inches(3.4), Inches(1.4),
    'P(母公司) → S(子公司,80%)\n'
    '销售A产品100件\n'
    '收入：1,000,000 元\n'
    '成本：600,000 元\n'
    'S期末库存：70件\n'
    '所得税率：25%', size=11, color=DARK_GRAY)

add_text_box(s, Inches(5.0), Inches(1.6), Inches(3.5), Inches(0.4), '🔢 计算过程', size=14, bold=True, color=DARK_BLUE)
add_text_box(s, Inches(5.0), Inches(2.1), Inches(3.5), Inches(1.4),
    '类型识别：P→S = 顺流\n'
    '毛利率 = (100万-60万)/100万\n'
    '       = 40%\n'
    '未售比 = 70/100 = 70%\n'
    '利润 = 100万×40%×70%\n'
    '     = 280,000 元', size=11, color=DARK_GRAY)

add_rect(s, Inches(8.8), Inches(1.6), Inches(4.0), Inches(2.0), DARK_BLUE)
add_text_box(s, Inches(9.0), Inches(1.7), Inches(3.6), Inches(0.4), '📝 抵消分录', size=14, bold=True, color=WHITE)
add_text_box(s, Inches(9.0), Inches(2.1), Inches(3.6), Inches(1.4),
    '借：营业收入   700,000\n'
    '  贷：营业成本   420,000\n'
    '  贷：存货       280,000\n\n'
    '借：递延所得税资产  70,000\n'
    '  贷：所得税费用     70,000\n\n'
    '净利润影响：-210,000', size=11, color=RGBColor(0xCC, 0xDD, 0xEE))

add_text_box(s, Inches(0.8), Inches(4.0), Inches(12.0), Inches(0.5), '📱 系统实际输出界面', size=14, bold=True, color=DARK_BLUE)
add_rect(s, Inches(0.8), Inches(4.5), Inches(12.0), Inches(2.3), RGBColor(0xF5, 0xF5, 0xF5))
add_text_box(s, Inches(1.0), Inches(4.6), Inches(11.5), Inches(2.0),
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n'
    'ELIM-202605-0001 | 顺流交易 P001→S001 | A产品 | 未实现利润：280,000 元\n'
    '毛利率=40.00% | 未售=70.00% | 收入=1,000,000 | 成本=600,000\n'
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n'
    '借：营业收入                    700,000\n'
    '  贷：营业成本                  420,000\n'
    '  贷：存货                      280,000\n'
    '借：递延所得税资产               70,000\n'
    '  贷：所得税费用                 70,000\n'
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', size=10, color=DARK_GRAY)
add_footer(s)

# ============================================================
# SLIDE 12: 与老师平台对比
# ============================================================
s = add_blank_slide()
add_slide_title(s, '与DRP平台"集团合并报表"功能对比', '对标教师T11系统，取长补短')

add_table_shape(s, Inches(0.5), Inches(1.6), Inches(12.3), Inches(4.5),
    ['对比维度', '教师DRP平台 T11', '本系统 V4.0', '评价'],
    [
        ['未实现利润抵消', '空白表单，需手动填写', '✅ 自动计算+JSON规则引擎驱动', '⭐ 本系统更深'],
        ['投资成本抵消', '✅ 有选项', '❌ 不在此范围', '不同子模块'],
        ['内部往来抵消', '✅ 有选项', '❌ 不在此范围', '不同子模块'],
        ['试算平衡表', '✅ 有展示区域', '✅ 自动汇总+借贷差额验证', '⭐ 本系统更强'],
        ['分录手动调整', '✅ 新增/编辑/删除', '✅ 修改金额/删行/追加入工分录', '⚖️ 持平'],
        ['历史记录', '✅ 处理历史', '✅ 批次分组+完整计算过程+哈希追溯', '⭐ 本系统更强'],
        ['仪表盘', '✅ 4个KPI+对账趋势图', '✅ 5个KPI+2图表+对比表+时间线', '⚖️ 各有所长'],
        ['筛选过滤', '✅ 按类型/状态/期间', '✅ 按交易类型标签过滤', '⚖️ 持平'],
        ['接口管理', '✅ 银企/ERP/金蝶/浪潮', '❌ 原型阶段预留', 'V3.0扩展'],
    ],
    col_widths=[Inches(2.2), Inches(3.5), Inches(4.0), Inches(2.6)]
)
add_footer(s)

# ============================================================
# SLIDE 13: 交付范围
# ============================================================
s = add_blank_slide()
add_slide_title(s, 'V2.0 交付范围与未来扩展', '缩实现，保专业 — 文档像大系统，实现像精巧Demo')

add_text_box(s, Inches(0.8), Inches(1.6), Inches(5.8), Inches(0.5), '✅ V2.0 已实现（本课程交付）', size=16, bold=True, color=ACCENT_GREEN)
items_v20 = [
    'Excel三表上传 + 手动录入', '数据校验 + 4种异常检测',
    '顺流/逆流/平流自动识别', '未实现利润自动计算',
    '抵消分录自动生成（JSON规则引擎）', '递延所得税+少数股东分摊',
    '试算平衡验证', '分录筛选与手动调整',
    '可视化仪表盘（KPI+图表+对比表）', '历史记录+哈希追溯',
    'Excel/PDF导出', '一键加载示例数据',
]
for i, item in enumerate(items_v20):
    col = i % 2; row = i // 2
    add_text_box(s, Inches(0.8 + col * 6.2), Inches(2.2 + row * 0.45), Inches(5.8), Inches(0.4),
        f'  ✅ {item}', size=11, color=DARK_GRAY)

add_text_box(s, Inches(0.8), Inches(5.2), Inches(5.8), Inches(0.5), '🔮 V3.0+ 未来扩展', size=16, bold=True, color=ACCENT_ORANGE)
items_v30 = [
    '加权平均法', '存货跌价准备抵消', '连续编报期初滚调',
    'Java Spring Boot微服务', 'Vue 3前端', '国产数据库（达梦/金仓）',
    '区块链存证（T33）', 'RESTful API对接T10/T05', '多会计准则切换',
]
for i, item in enumerate(items_v30):
    col = i % 3; row = i // 3
    add_text_box(s, Inches(0.8 + col * 4.2), Inches(5.7 + row * 0.45), Inches(3.8), Inches(0.4),
        f'  ⬜ {item}', size=11, color=MID_GRAY)

add_rect(s, Inches(7.2), Inches(1.6), Inches(5.5), Inches(5.0), DARK_BLUE)
add_text_box(s, Inches(7.4), Inches(1.8), Inches(5.0), Inches(0.5), '📦 交付物清单', size=16, bold=True, color=WHITE)
add_text_box(s, Inches(7.4), Inches(2.4), Inches(5.0), Inches(3.5),
    '📘 系统需求说明书 V4.0 (.docx)\n'
    '  GB/T 9385规范，含用例/数据模型/接口设计\n\n'
    '📗 用户操作手册 V2.0 (.docx)\n'
    '  分步详解，含常见问题排查\n\n'
    '📊 介绍PPT V2 (.pptx)\n'
    '  14页深蓝专业主题\n\n'
    '💻 系统源代码\n'
    '  635行 Python，模块化架构\n'
    '  app.py + engine/ + utils/ + rules.json\n\n'
    '📋 Excel模板 × 3 + 演示数据', size=11, color=RGBColor(0xCC, 0xDD, 0xEE))
add_footer(s)

# ============================================================
# SLIDE 14: 总结
# ============================================================
s = add_blank_slide()
add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), DARK_BLUE)
add_rect(s, Inches(0), Inches(2.8), Inches(13.333), Inches(0.06), ACCENT_GREEN)
add_text_box(s, Inches(1.5), Inches(0.8), Inches(10.5), Inches(0.8),
    '总结与展望', size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text_box(s, Inches(1.5), Inches(3.2), Inches(10.5), Inches(3.0),
    '🔹 解决了一个真实的会计难题：合并报表内部存货交易抵消\n'
    '🔹 用AI规则引擎替代硬编码，体现财务数智化设计思想\n'
    '🔹 完整覆盖顺流/逆流/平流 + 递延所得税 + 少数股东\n'
    '🔹 可追溯计算日志 + 哈希链，为穿透监管提供技术基础\n'
    '🔹 模块化架构，核心计算逻辑可平滑迁移至企业级Java平台\n\n'
    '课程收获：通过AI协作完成从需求分析、系统设计到原型实现的完整流程\n'
    '体现了MPAcc培养目标：既懂会计又懂技术的复合型人才', size=16, color=RGBColor(0xCC, 0xDD, 0xEE), align=PP_ALIGN.CENTER)
add_text_box(s, Inches(1.5), Inches(5.5), Inches(10.5), Inches(0.8),
    '谢谢！欢迎批评指正', size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text_box(s, Inches(1.5), Inches(6.3), Inches(10.5), Inches(0.6),
    '中国社会科学院大学 · 2025级MPAcc二班 · 尚浩然、李金阳 · 指导老师：张金昌', size=14, color=RGBColor(0x88, 0x99, 0xAA), align=PP_ALIGN.CENTER)

# ── 保存 ──
output = Path(__file__).parent / 'T11-介绍PPT-v2.pptx'
prs.save(str(output))
print(f'✅ PPT已生成：{output}')
print(f'   共 {len(prs.slides)} 页')
