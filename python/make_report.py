"""
生成 Olist 差评归因分析 Word 报告（v2：大字号 + 多元建模 + 建议 + 简历）。
"""
import os
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image

BASE = r'C:\Users\28223\Desktop\巴西olist数据分析项目'
RES = os.path.join(BASE, 'docs', 'sql_results')
FIG = os.path.join(BASE, 'python', 'output')
OUT = os.environ.get('REPORT_OUT') or os.path.join(
    BASE, 'Olist巴西电商_差评归因分析报告.docx')

FONT_CN = '微软雅黑'
FONT_BODY = '宋体'
FONT_CODE = 'Consolas'

# ---- 字号（整体放大，提升阅读舒适度）----
SZ_BODY = 12.0
SZ_TABLE = 10.0
SZ_CODE = 9.5
SZ_CAP = 10.0
SZ_H1 = 19.0
SZ_H2 = 15.0
SZ_H3 = 13.0

CONTENT_W = 16.2

doc = Document()

sec0 = doc.sections[0]
sec0.page_width = Cm(21.0)
sec0.page_height = Cm(29.7)
sec0.left_margin = Cm(2.4)
sec0.right_margin = Cm(2.4)
sec0.top_margin = Cm(2.2)
sec0.bottom_margin = Cm(2.2)

normal = doc.styles['Normal']
normal.font.name = FONT_BODY
normal.font.size = Pt(SZ_BODY)
normal.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_BODY)
normal.paragraph_format.line_spacing = 1.62
normal.paragraph_format.space_after = Pt(6)

# 去掉模板自带的最佳适应缩放
_zoom = doc.settings.element.find(qn('w:zoom'))
if _zoom is not None:
    _zoom.set(qn('w:val'), '100')


def _font(run, name, size=None, bold=None, color=None, italic=None):
    run.font.name = name
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rPr.append(rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia'):
        rf.set(qn(a), name)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def set_table_fixed(tbl, widths_cm):
    total = sum(widths_cm)
    if total > CONTENT_W:
        widths_cm = [w * CONTENT_W / total for w in widths_cm]
    tbl.autofit = False
    tblPr = tbl._tbl.tblPr
    for tag in ('w:tblW', 'w:tblLayout'):
        for el in tblPr.findall(qn(tag)):
            tblPr.remove(el)
    tw = OxmlElement('w:tblW')
    tw.set(qn('w:w'), str(int(sum(widths_cm) / 2.54 * 1440)))
    tw.set(qn('w:type'), 'dxa')
    tblPr.append(tw)
    lay = OxmlElement('w:tblLayout')
    lay.set(qn('w:type'), 'fixed')
    tblPr.append(lay)
    grid = tbl._tbl.find(qn('w:tblGrid'))
    if grid is not None:
        for gc, w in zip(grid.findall(qn('w:gridCol')), widths_cm):
            gc.set(qn('w:w'), str(int(w / 2.54 * 1440)))
    for row in tbl.rows:
        for cell, w in zip(row.cells, widths_cm):
            cell.width = Cm(w)
    return widths_cm


def shade(obj, hex_color):
    """给段落或表格单元格加底色。_Cell 会代理属性，必须按内部属性区分。"""
    if hasattr(obj, '_tc'):                     # _Cell
        pr = obj._tc.get_or_add_tcPr()
    elif hasattr(obj, '_p'):                    # Paragraph
        pr = obj._p.get_or_add_pPr()
    else:
        el = obj
        pr = el.get_or_add_pPr() if hasattr(el, 'get_or_add_pPr') else el.get_or_add_tcPr()
    sh = OxmlElement('w:shd')
    sh.set(qn('w:val'), 'clear')
    sh.set(qn('w:color'), 'auto')
    sh.set(qn('w:fill'), hex_color)
    pr.append(sh)


def para(text='', size=SZ_BODY, bold=False, color=None, align=None,
         space_before=0, space_after=6, indent=None, italic=False):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if indent is not None:
        p.paragraph_format.first_line_indent = Cm(indent)
    if text:
        _font(p.add_run(text), FONT_BODY, size, bold, color, italic)
    return p


def h1(text, num=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(12)
    if num:
        _font(p.add_run(num + '  '), FONT_CN, SZ_H1, True, '1E3A5F')
    _font(p.add_run(text), FONT_CN, SZ_H1, True, '1E3A5F')
    pPr = p._element.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single')
    bot.set(qn('w:sz'), '10')
    bot.set(qn('w:color'), '1E3A5F')
    pbdr.append(bot)
    pPr.append(pbdr)
    return p


def h2(text, num=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(8)
    if num:
        _font(p.add_run(num + '  '), FONT_CN, SZ_H2, True, '2E5C8A')
    _font(p.add_run(text), FONT_CN, SZ_H2, True, '2E5C8A')
    return p


def h3(text, num=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(5)
    if num:
        _font(p.add_run(num + '  '), FONT_CN, SZ_H3, True, '374151')
    _font(p.add_run(text), FONT_CN, SZ_H3, True, '374151')
    return p


def bullet(text, size=SZ_BODY):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(0.9)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.5
    _font(p.add_run(text), FONT_BODY, size)
    return p


def code_block(code, caption=None, size=SZ_CODE):
    if caption:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(9)
        p.paragraph_format.space_after = Pt(3)
        _font(p.add_run('▍' + caption), FONT_CN, 10.5, True, '6B7280')
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_fixed(tbl, [CONTENT_W])
    cell = tbl.cell(0, 0)
    shade(cell, 'F7F8FA')
    tcPr = cell._tc.get_or_add_tcPr()
    bd = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), '4')
        e.set(qn('w:color'), 'E5E7EB')
        bd.append(e)
    tcPr.append(bd)
    cell.text = ''
    first = True
    for line in code.split('\n'):
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        _font(p.add_run(line if line else ' '), FONT_CODE, size, color='1F2937')
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return tbl


def add_figure(path, caption, max_w=None, max_h=11.0):
    if not os.path.exists(path):
        para('[图片缺失: %s]' % os.path.basename(path), color='C0392B')
        return
    im = Image.open(path)
    w_cm = min(max_w or CONTENT_W, CONTENT_W)
    h_cm = w_cm * im.height / im.width
    if h_cm > max_h:
        h_cm = max_h
        w_cm = h_cm * im.width / im.height
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.add_run().add_picture(path, width=Cm(w_cm))
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_after = Pt(12)
    _font(cp.add_run(caption), FONT_CN, SZ_CAP, color='6B7280')


def add_table(headers, rows, widths=None, caption=None, font_size=SZ_TABLE,
              header_bg='1E3A5F', zebra=True, right_cols=None, first_left=True):
    if caption:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(9)
        p.paragraph_format.space_after = Pt(4)
        _font(p.add_run('表  ' + caption), FONT_CN, 10.5, True, '6B7280')
    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixed = set_table_fixed(tbl, widths if widths else [CONTENT_W / len(headers)] * len(headers))
    hdr = tbl.rows[0]
    for i, htxt in enumerate(headers):
        c = hdr.cells[i]
        c.text = ''
        shade(c, header_bg)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        _font(p.add_run(str(htxt)), FONT_CN, font_size, True, 'FFFFFF')
    for ri, row in enumerate(rows):
        cells = tbl.add_row().cells
        for ci, val in enumerate(row):
            c = cells[ci]
            c.text = ''
            c.width = Cm(fixed[ci])
            if zebra and ri % 2 == 1:
                shade(c, 'F7F9FB')
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.2
            if right_cols and ci in right_cols:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif ci == 0 and first_left:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _font(p.add_run(str(val)), FONT_BODY, font_size)
    for ri in range(len(tbl.rows)):
        for ci, w in enumerate(fixed):
            tbl.rows[ri].cells[ci].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return tbl


def callout(title, text, bg='FFF8E1', border='E8A33D', tcolor='8A5A00', size=11.0):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_fixed(tbl, [CONTENT_W])
    cell = tbl.cell(0, 0)
    shade(cell, bg)
    tcPr = cell._tc.get_or_add_tcPr()
    bd = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), '4' if edge != 'left' else '20')
        e.set(qn('w:color'), border)
        bd.append(e)
    tcPr.append(bd)
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(3)
    _font(p.add_run(title), FONT_CN, size + 0.5, True, tcolor)
    for line in text.split('\n'):
        pp = cell.add_paragraph()
        pp.paragraph_format.space_before = Pt(0)
        pp.paragraph_format.space_after = Pt(2)
        pp.paragraph_format.line_spacing = 1.45
        _font(pp.add_run(line), FONT_BODY, size)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return tbl


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ===========================================================================
# 封面
# ===========================================================================
for _ in range(4):
    doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_font(p.add_run('Olist 巴西电商数据集'), FONT_CN, 28, True, '1E3A5F')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(8)
_font(p.add_run('客户差评归因分析报告'), FONT_CN, 22, True, '2E5C8A')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(34)
_font(p.add_run('—— 从流程拆解到多元统计建模的完整分析链路 ——'), FONT_CN, 12.5, False, '6B7280')

info = [
    ('分析对象', 'Olist 巴西电商平台公开数据集（2016-09 至 2018-10）'),
    ('数据规模', '9 张表 · 1,550,922 行 · 99,441 笔订单 · 96,096 名客户'),
    ('技术栈', 'PostgreSQL 18 · VS Code · Python 3.11（pandas / SciPy / scikit-learn）'),
    ('分析方法', '流程拆解 → 单因素归因 → 混杂因素分层 → 相关分析 → 多元回归建模'),
    ('分析产出', 'SQL 31 条 · 特征 10 维 · 建模样本 93,047 条 · 41 张图表'),
]
tbl = doc.add_table(rows=0, cols=2)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_fixed(tbl, [3.8, 12.4])
for k, v in info:
    cells = tbl.add_row().cells
    cells[0].width = Cm(3.8)
    cells[1].width = Cm(12.4)
    _font(cells[0].paragraphs[0].add_run(k), FONT_CN, 11.5, True, '374151')
    pv = cells[1].paragraphs[0]
    pv.paragraph_format.line_spacing = 1.35
    _font(pv.add_run(v), FONT_BODY, 11.5, color='1F2937')
set_table_fixed(tbl, [3.8, 12.4])

for _ in range(5):
    doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_font(p.add_run(datetime.now().strftime('%Y 年 %m 月 %d 日')), FONT_CN, 11.5, color='9CA3AF')

page_break()

# ===========================================================================
# 目录
# ===========================================================================
h1('目  录')
toc = [
    ('第一章', '项目背景与研究问题', 0),
    ('1.1', '数据集是什么', 1), ('1.2', '研究问题是如何一步步收敛的', 1),
    ('1.3', '报告框架', 1),
    ('第二章', '数据理解与环境搭建', 0),
    ('2.1', '数据集结构：9 张表的分工', 1), ('2.2', '数据规模与时间跨度', 1),
    ('2.3', '数据导入与环境搭建', 1), ('2.4', '数据质量与两个必须注意的陷阱', 1),
    ('第三章', '分析方法论：我是如何“想到”这些因素的', 0),
    ('3.1', '第一步：精确定义结果指标', 1), ('3.2', '第二步：沿业务流程拆解', 1),
    ('3.3', '第三步：把因素翻译成可测量字段', 1), ('3.4', '第四步：分组对比与控制变量', 1),
    ('3.5', '第五步：混杂因素的分层验证', 1), ('3.6', '六条贯穿全程的分析思维', 1),
    ('第四章', '分析过程：六个步骤的实证推进', 0),
    ('4.1', '第 0 步：建立差评率基线', 1), ('4.2', '第 1 步：流程拆解，定位瓶颈', 1),
    ('4.3', '第 2 步：地理因素', 1), ('4.4', '第 3 步：卖家因素', 1),
    ('4.5', '第 4 步：商品因素', 1), ('4.6', '第 5 步：时间因素', 1),
    ('4.7', '第 6 步：五因素汇总对比', 1),
    ('第五章', '统计学验证：相关性与显著性', 0),
    ('5.1', '相关系数矩阵（Pearson / Spearman）', 1), ('5.2', '显著性检验与 p 值', 1),
    ('5.3', '效应量 Cohen\'s d', 1), ('5.4', '分类变量卡方检验', 1),
    ('5.5', '阈值效应验证：为什么相关系数会低估', 1),
    ('第六章', '多元统计建模：控制混杂后的净效应', 0),
    ('6.1', '为什么必须做多元模型', 1), ('6.2', '特征工程与真实地理距离', 1),
    ('6.3', '多元线性回归（OLS）', 1), ('6.4', '二元逻辑回归（Logit）与优势比', 1),
    ('6.5', '随机森林对照与预测力评估', 1), ('6.6', '交互效应检验', 1),
    ('6.7', '稳健性检验', 1), ('6.8', '两处颠覆单因素结论的发现', 1),
    ('6.9', '模型的可解释性与边界', 1),
    ('第七章', '结论与公司建议', 0),
    ('7.1', '三条核心结论', 1), ('7.2', '公司可执行建议（分优先级）', 1),
    ('7.3', '指标体系建议', 1), ('7.4', '落地路线图', 1),
    ('第八章', '简历表述参考', 0),
    ('8.1', '简历项目条目（完整版）', 1), ('8.2', '简历精简版（推荐）', 1),
    ('8.3', 'STAR 面试讲述版', 1), ('8.4', '技能关键词与注意事项', 1),
    ('第九章', '数据局限与分析反思', 0),
    ('附录', '完整代码与文件清单', 0),
]
for no, title, lvl in toc:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.25
    if lvl == 0:
        p.paragraph_format.space_before = Pt(10)
        _font(p.add_run('%s  ' % no), FONT_CN, 12.5, True, '1E3A5F')
        _font(p.add_run(title), FONT_CN, 12.5, True, '1E3A5F')
    else:
        _font(p.add_run('        %s  ' % no), FONT_BODY, 11, color='4B5563')
        _font(p.add_run(title), FONT_BODY, 11, color='4B5563')

page_break()

# ===========================================================================
# 第一章
# ===========================================================================
h1('项目背景与研究问题', '第一章')

h2('数据集是什么', '1.1')
para('Olist 是巴西的一家电商聚合平台，帮助中小商家把商品铺到各大电商渠道。本次分析使用的'
     '是其公开的脱敏交易数据，覆盖 2016 年 9 月至 2018 年 10 月共 25 个月的真实订单记录。', indent=0.85)
para('这份数据的价值在于：它不仅记录了「卖了什么、卖了多少钱」，还完整记录了订单从下单到签收的'
     '每一个时间节点，以及客户留下的评分和文字评论。这意味着可以回答「为什么客户不满意」这类'
     '归因问题，而不只是做销售统计。', indent=0.85)

h2('研究问题是如何一步步收敛的', '1.2')
para('这一点值得单独说明，因为它是整个项目最有价值的部分——研究问题不是一开始就明确的，'
     '而是在分析过程中被数据不断修正的。', indent=0.85)

add_table(
    ['阶段', '原始问题', '遇到的问题', '修正后的问题'],
    [['起点', '配送速度是否影响客户满意度？', '问题太宽泛，“配送速度”本身包含多个环节',
      '拆解为：下单→审核→备货→运输，哪一段是瓶颈？'],
     ['中段', '哪些因素导致配送慢？', '发现“慢”和“迟到”不是同一件事', '区分：绝对时长 vs 是否违约'],
     ['收敛', '哪些因素导致差评？权重如何？', '“导致”需要因果证据，不能只看相关',
      '哪些因素与差评相关？强度排序？业务该抓哪个？']],
    widths=[1.8, 4.4, 4.4, 5.6], caption='研究问题的三次收敛')

callout('问题收敛的关键转折',
        '当分析进行到一半时，我发现原来的问题「配送速度是否影响满意度」已经不足以支撑业务决策。\n'
        '因为它只回答“有没有影响”，却不能回答“该改哪个”。于是问题被重新表述为：\n'
        '「影响差评的因素有哪些？各因素权重如何？业务该优先抓哪个？」\n'
        '这个重新表述，把一份描述性分析变成了一份可执行的诊断报告。')

h2('报告框架', '1.3')
para('本报告按真实分析推进顺序组织，而不是按“结论优先”的倒叙，目的是完整呈现思考过程：', indent=0.85)
bullet('第二章：看清楚数据是什么、有什么坑（数据理解）')
bullet('第三章：说明“如何想到要查这些因素”的方法论（思维框架）')
bullet('第四章：六个步骤的实证推进，每步配 SQL 与结果（分析过程）')
bullet('第五章：相关系数、显著性、效应量的统计验证（相关性证据）')
bullet('第六章：多元回归控制混杂因素后的净效应（严谨建模）')
bullet('第七章：结论与可执行的公司建议')
bullet('第八章：可直接使用的简历表述')
bullet('第九章：诚实记录数据局限与未解之谜')

page_break()

# ===========================================================================
# 第二章
# ===========================================================================
h1('数据理解与环境搭建', '第二章')

h2('数据集结构：9 张表的分工', '2.1')
para('数据集由 9 张表组成，围绕「订单」这一核心事实表展开。理解表间关系是后续所有分析的前提。', indent=0.85)
add_table(
    ['表名', '行数', '业务含义', '在分析中的作用'],
    [['orders', '99,441', '订单主表，含 5 个关键时间戳', '★核心：计算各环节时长'],
     ['order_items', '112,650', '订单明细：商品、单价、运费、发货期限', '★核心：超期发货判定'],
     ['order_reviews', '99,224', '评价：1-5 星评分 + 文字评论', '★核心：被解释变量'],
     ['customers', '99,441', '客户所在城市与州、邮编', '地理因素'],
     ['sellers', '3,095', '卖家所在城市与州、邮编', '卖家因素'],
     ['products', '32,951', '商品品类、重量、尺寸', '商品因素'],
     ['order_payments', '103,886', '支付方式、分期数、金额', '辅助（证明影响很小）'],
     ['geolocation', '1,000,163', '邮编对应的经纬度', '★计算真实距离'],
     ['product_category_translation', '71', '品类葡语→英语翻译', '品类名称可读性']],
    widths=[4.0, 1.9, 6.2, 4.1], caption='9 张表的分工')

para('表间关系如下（外键约束已验证，孤儿记录为 0）：', indent=0.85)
code_block(
    'customers ──1:N──> orders ──1:N──> order_items <──N:1── products\n'
    '                      │                 └────N:1─────────> sellers\n'
    '                      ├──1:N──> order_payments\n'
    '                      └──1:N──> order_reviews\n'
    '\n'
    'products.product_category_name ──N:1──> product_category_translation',
    caption='表间关系（ER 结构）')

h2('数据规模与时间跨度', '2.2')
add_table(['指标', '数值'],
          [['总行数', '1,550,922 行'],
           ['订单数', '99,441 笔'],
           ['真实客户数（customer_unique_id 去重）', '96,096 人'],
           ['时间跨度', '2016-09-04 至 2018-10-17，共 25 个月'],
           ['平均每单商品数', '1.14 件（最多 21 件）'],
           ['订单状态为 delivered 的比例', '97.02%'],
           ['已评价订单数', '98,673 单'],
           ['差评（1-2 星）基线比例', '14.63%']],
          widths=[8.4, 7.8], caption='数据集基本信息')

h2('数据导入与环境搭建', '2.3')
para('环境搭建过程中解决了几个具体问题，记录如下：', indent=0.85)
h3('（1）导入 PostgreSQL', '2.3.1')
para('使用 psql 的 \\copy 命令（客户端读取），因为服务端 COPY 需要 pg_read_server_files 超级用户'
     '权限，普通角色无法使用。9 个 CSV 全部导入，逐表核对行数与源文件一致。', indent=0.85)
code_block(
    '\\copy customers (customer_id, customer_unique_id, customer_zip_code_prefix,\n'
    '                 customer_city, customer_state)\n'
    "FROM 'C:/Users/.../olist_customers_dataset.csv'\n"
    "WITH (FORMAT csv, HEADER true, ENCODING 'UTF8', NULL '');",
    caption='导入命令示例（\\copy 是 psql 元命令，需单行书写）')

h3('（2）建立分析视图', '2.3.2')
para('为避免每次分析重复写多表 JOIN，创建了 3 个视图。其中 v_orders_full 是订单宽表，'
     '把订单、客户、金额合计、支付合计、评分合并成一行一单。', indent=0.85)
code_block(
    'CREATE VIEW v_orders_full AS\n'
    'SELECT o.order_id, o.order_status,\n'
    "       o.order_purchase_timestamp AS purchase_at,\n"
    "       date_trunc('month', o.order_purchase_timestamp) AS purchase_month,\n"
    '       o.order_delivered_customer_date AS delivered_at,\n'
    '       o.order_estimated_delivery_date AS estimated_at,\n'
    '       c.customer_unique_id, c.customer_city, c.customer_state,\n'
    '       SUM(i.price)                AS item_total,\n'
    '       SUM(i.freight_value)        AS freight_total,\n'
    '       SUM(i.price + i.freight_value) AS order_total,\n'
    '       count(i.order_item_id)      AS item_count,\n'
    '       p.payment_total, r.review_score\n'
    'FROM orders o\n'
    'JOIN customers c        ON c.customer_id = o.customer_id\n'
    'LEFT JOIN order_items i ON i.order_id    = o.order_id\n'
    'LEFT JOIN (...) p       ON p.order_id    = o.order_id\n'
    'LEFT JOIN (...) r       ON r.order_id    = o.order_id\n'
    'GROUP BY ...;',
    caption='v_orders_full 视图（订单宽表）')

h2('数据质量与两个必须注意的陷阱', '2.4')
para('正式分析前做了数据质量检查。这份数据的参照完整性很好（6 组外键孤儿记录全为 0），'
     '但仍有两个陷阱足以让后续结论完全错误。', indent=0.85)
add_figure(os.path.join(RES, '28_7.1关键字段缺失率.png'),
           '图 2-1  关键字段缺失率：评论标题缺失 88.3%、评论内容缺失 58.7%')
add_figure(os.path.join(RES, '31_7.4顾客标识的唯一性（说明为何必须用customer_unique_id）.png'),
           '图 2-2  顾客标识唯一性：99,441 个 customer_id 各不相同，真实客户仅 96,096 人')

callout('陷阱一：customer_id ≠ customer_unique_id',
        'customers 表有 99,441 行，唯一 customer_id 也是 99,441 个——每笔订单都生成新的 customer_id。\n'
        '真正代表“同一个人”的是 customer_unique_id（96,096 个）。\n\n'
        '如果误用 customer_id 计算复购率，会得到 0% 的错误结论。真实复购率约 3.1%。')

callout('陷阱二：平均值会骗人',
        '运输时长平均值 9.33 天，但中位数只有 7.10 天，p90 达 18.90 天，最大 205 天。\n'
        '均值 > 中位数说明分布右偏——平均值被少数极端慢单拉高了。\n\n'
        '因此本报告在描述时长时同时给出均值与中位数，并在相关系数分析中同时计算 Pearson 与 Spearman。')
para('另外两处细节：评论文本缺失率高，本分析只使用评分；存在 23 条“签收早于发货”的负值记录'
     '（占 0.02%），分析时剔除。', indent=0.85)

page_break()

# ===========================================================================
# 第三章
# ===========================================================================
h1('分析方法论：我是如何“想到”这些因素的', '第三章')
para('这一章是本报告的核心。新手常见的错误是“拍脑袋列一堆因素，然后逐个试”。'
     '正确做法是有推导路径的——本章说明这条路径。', indent=0.85)

h2('第一步：精确定义结果指标', '3.1')
para('在分析“配送是否影响满意度”之前，必须定义什么算“慢”、什么算“差评”。'
     '这里有一个 90% 新手会混淆的区分：', indent=0.85)
add_table(['概念', '定义', '衡量什么', '业务含义'],
          [['配送时长', '签收时间 − 下单时间', '绝对快慢', '客户实际等待多久'],
           ['是否迟到', '签收时间 > 预估送达时间', '是否违约', '平台是否兑现承诺'],
           ['差评', 'review_score ≤ 2', '主观评价', '客户是否满意']],
          widths=[3.0, 5.4, 3.2, 4.6], caption='三个核心指标的严格定义')
para('两者不是一回事：一单送了 20 天，如果平台预估 25 天，它不算迟到；另一单送出 10 天、预估 7 天，'
     '它迟到了。这个区分在第五章被证明极其关键。', indent=0.85)

h2('第二步：沿业务流程拆解', '3.2')
para('配送不是一个动作，而是一条流水线。把时间轴摊开，每个环节都可能堵：', indent=0.85)
code_block(
    '下单 ──①──> 审核通过 ──②──> 交给物流 ──③──> 客户签收\n'
    'purchase    approved      carrier      delivered\n'
    '   └──── ① 支付确认 ────┘\n'
    '         └──── ② 备货 ────┘\n'
    '                    └──── ③ 运输 ────┘',
    caption='配送流程的时间轴拆解')
para('这条时间轴在数据中是现成的——orders 表提供四个时间戳，可以精确算出每一段耗时。'
     '这一步的价值在于：不知道瓶颈在哪一段，后面的归因都是瞎猜。', indent=0.85)

h2('第三步：把因素翻译成可测量字段', '3.3')
para('每段流程问“什么东西会让它变慢”，然后去数据里找对应的列。关键原则是：'
     '先问业务机制，再找数据字段，而不是反过来（纯数据挖矿容易得到伪相关）。', indent=0.85)
add_table(['环节', '候选因素', '为什么可能影响（机制）', '对应字段'],
          [['① 审核', '支付方式', 'boleto（巴西银行票据）需线下付款，到账慢', 'payment_type'],
           ['① 审核', '订单金额', '高金额可能触发风控复核', 'payment_value'],
           ['② 备货', '卖家能力', '库存管理差、拣货慢', 'seller_id'],
           ['② 备货', '商品品类', '定制家具需备货期，标品现货即发', 'product_category_name'],
           ['② 备货', '商品重量', '重货需特殊包装与安排', 'product_weight_g'],
           ['② 备货', '多卖家订单', '需等所有卖家都发货', 'seller_id 计数'],
           ['② 备货', '是否超发货期限', '卖家未按自己承诺的期限发货', 'shipping_limit_date'],
           ['③ 运输', '客户所在州', '巴西物流基建差异极大', 'customer_state'],
           ['③ 运输', '是否跨州', '同州距离短、中转少', 'customer_state vs seller_state'],
           ['③ 运输', '真实距离', '客户与卖家的直线距离', 'geolocation 经纬度'],
           ['③ 运输', '下单季节', '黑五/圣诞订单暴增致物流超载', 'order_purchase_timestamp'],
           ['③ 运输', '下单周几', '周末下单需等周一处理', 'EXTRACT(DOW ...)']],
          widths=[1.6, 3.0, 6.8, 4.8], caption='因素清单：从业务机制到数据字段的映射')

h2('第四步：分组对比与控制变量', '3.4')
para('数据分析的核心思维是控制变量。举个例子：如果整体看“21 天以上送达的订单评分很低”，'
     '不能直接说“配送慢导致差评”——因为那批订单里可能同时集中了偏远州、重货、boleto 支付'
     '三个因素，你分不清是哪个在起作用。', indent=0.85)
callout('控制变量的操作原则',
        '每次只让一个因素变化，其他因素尽量固定。\n\n'
        '错误做法：SELECT avg(...) FROM orders;                -- 只看整体\n'
        '正确做法：SELECT customer_state, 迟到率 FROM ... GROUP BY customer_state;')

h2('第五步：混杂因素的分层验证', '3.5')
para('发现“办公家具品类备货要 9.89 天”之后，必须回头检查：那些“慢卖家”是不是恰好都在卖家具？'
     '如果是，那就不是卖家能力问题，而是品类特性。', indent=0.85)
para('实际结果是：6 个最慢卖家中，2 个确实卖办公家具（品类特性），但另外 4 个卖的是床品、美妆'
     '——这些标品本该 2-3 天发货，却慢了 11-16 天，属于真实的运营问题。'
     '排除家具后，卖家备货仍从 2.75 天分化到 16.1 天。', indent=0.85)

h2('六条贯穿全程的分析思维', '3.6')
add_table(['序号', '思维原则', '在本次分析中的具体体现'],
          [['1', '先拆流程，再找原因', '不做四段耗时拆解，就不知道瓶颈在运输（74%）而非支付（3%）'],
           ['2', '警惕指标被补偿机制掩盖', '平台给偏远州极宽松的预估（38.8 天），压低了迟到率，但实际等待仍达 20 天'],
           ['3', '最大的问题 ≠ 最该改的问题', '运输占 74% 但由距离决定、平台改不了；备货占 22% 却完全可控'],
           ['4', '按影响总量排序，不按严重程度', '备货 16.1 天/76 单影响面 1,010 天；11.6 天/980 单影响面 8,623 天'],
           ['5', '结果越完美越要怀疑', '多卖家订单迟到率 0.0%、评分 2.24，三处反常 → 查出时间戳失真'],
           ['6', '不要为异常编故事', '2018-02/03 异常只能说“运输环节出事，原因需业务方确认”，不臆测罢工']],
          widths=[1.2, 5.2, 9.8], caption='六条分析思维原则')
callout('第六章将补充第七条思维',
        '单因素相关可能被混杂、抑制、中介三种机制扭曲。\n'
        '只有把多个变量同时放进模型，才能得到“净效应”。详见第六章两处颠覆性发现。',
        bg='EDF4FB', border='4A90D9', tcolor='1E3A5F')

page_break()

# ===========================================================================
# 第四章
# ===========================================================================
h1('分析过程：六个步骤的实证推进', '第四章')
para('本章按实际执行顺序展开。每一步都遵循相同结构：提出假设 → 写 SQL → 看结果 → 修正判断。', indent=0.85)

h2('第 0 步：建立差评率基线', '4.1')
para('任何归因分析都需要基准。先看评分整体分布，再算出差评率基线——后面所有分组都要和这个数比。', indent=0.85)
add_figure(os.path.join(RES, '02_0.2评分分布（看清差评的规模）.png'),
           '图 4-1  评分分布：5 星占 57.78%，但 1 星也有 11.51%，呈现两极分化')
add_figure(os.path.join(RES, '03_0.3差评基线★后面所有分组都要和这个数比.png'),
           '图 4-2  差评率基线 14.63%（98,673 单已评价，其中 14,437 单为 1-2 星）')
code_block(
    '-- 差评基线：所有已评价订单中 1-2 星的比例\n'
    'SELECT count(*) AS 已评订单数,\n'
    '       count(*) FILTER (WHERE review_score <= 2) AS 差评订单数,\n'
    '       round(100.0 * count(*) FILTER (WHERE review_score <= 2)\n'
    '             / count(*), 2) AS 差评率基线\n'
    'FROM (SELECT order_id, max(review_score) AS review_score\n'
    '      FROM olist.order_reviews GROUP BY order_id) t;',
    caption='SQL 0.3  建立差评率基线')

h2('第 1 步：流程拆解，定位瓶颈', '4.2')
para('这是最关键的一步，因为它决定了后面所有分析的重心。把订单按四个环节拆开，看各段平均耗时。', indent=0.85)
add_figure(os.path.join(RES, '04_1.1四段耗时拆解★这一步决定后面查哪里.png'),
           '图 4-3  四段耗时拆解：审核 0.43 天、备货 2.80 天、运输 9.33 天、全程 12.56 天')
code_block(
    'SELECT round(avg(EXTRACT(EPOCH FROM (order_approved_at\n'
    '             - order_purchase_timestamp))/86400)::numeric, 2) AS 下单到审核_天,\n'
    '       round(avg(EXTRACT(EPOCH FROM (order_delivered_carrier_date\n'
    '             - order_approved_at))/86400)::numeric, 2)     AS 审核到发货_天,\n'
    '       round(avg(EXTRACT(EPOCH FROM (order_delivered_customer_date\n'
    '             - order_delivered_carrier_date))/86400)::numeric, 2) AS 发货到签收_天,\n'
    '       round(avg(EXTRACT(EPOCH FROM (order_delivered_customer_date\n'
    '             - order_purchase_timestamp))/86400)::numeric, 2) AS 全程_天,\n'
    '       count(*) AS 样本数\n'
    'FROM olist.orders\n'
    'WHERE order_delivered_customer_date IS NOT NULL\n'
    '  AND order_approved_at IS NOT NULL\n'
    '  AND order_delivered_carrier_date IS NOT NULL;',
    caption='SQL 1.1  四段耗时拆解')
add_table(['环节', '平均耗时', '占全程', '判断'],
          [['下单 → 审核', '0.43 天', '3.4%', '可忽略'],
           ['审核 → 发货（备货）', '2.80 天', '22.3%', '中等，且可控'],
           ['发货 → 签收（运输）', '9.33 天', '74.3%', '★ 主要瓶颈'],
           ['全程', '12.56 天', '100%', '—']],
          widths=[5.2, 3.0, 2.6, 5.4], caption='四段耗时与占比（样本 96,461 单）',
          right_cols=[1, 2])
callout('这一步推翻了此前的一个预判',
        '我原本推测 boleto（巴西银行票据支付）会因为要等银行到账而拖慢审核，进而拖慢全流程。\n'
        '但数据显示审核环节只占 0.43 天（约 10 小时），占全程 3.4%——基本可以忽略。')
para('由于平均值容易被极端值拉高，进一步检查分布：', indent=0.85)
add_figure(os.path.join(RES, '05_1.2运输时长分布（验证平均值是否被极端值拉高）.png'),
           '图 4-4  运输时长分布：平均 9.33 天，中位数仅 7.10 天，p90 达 18.90 天')
para('运输时长的标准差（8.76）是备货（3.54）的 2.5 倍，说明运输环节不仅更慢，而且高度不稳定。', indent=0.85)

h2('第 2 步：地理因素', '4.3')
bullet('客户在哪 —— 哪个州的客户收货慢？（反映目的地物流基建）')
bullet('是否跨州 —— 同州 vs 跨州差多少？（隔离“距离”这个变量）')
h3('2.1 各州运输时长与迟到率', '4.3.1')
add_figure(os.path.join(RES, '07_2.1各州运输时长与迟到率（样本_=200）.png'),
           '图 4-5  各州运输时长与迟到率：AL 21.1 天/23.9%、PA 20.3 天、MA 18.0 天/19.7% 居前')
para('运输最慢的州全部集中在东北部（AL、MA、SE、CE、PB、PI、BA、RN、PE）和北部（PA、RO），'
     '与巴西物流基建的区域差异一致。', indent=0.85)

h3('2.2 同州 vs 跨州', '4.3.2')
add_figure(os.path.join(RES, '08_2.2同州vs跨州（隔离_距离_这个变量）★.png'),
           '图 4-6  跨州 11.87 天 vs 同州 4.80 天，运输时间相差 2.5 倍')
callout('这里出现了一个反直觉现象',
        '运输时间差了 2.5 倍（11.87 vs 4.80），但迟到率只差 1.5 倍（9.26% vs 6.06%）。\n'
        '如果距离是主因，迟到率不该差得这么少。这提示：平台的预估送达时间本身可能按距离动态给。',
        bg='EDF4FB', border='4A90D9', tcolor='1E3A5F')

h3('2.3 验证：平台预估是否按距离调整', '4.3.3')
add_figure(os.path.join(RES, '09_2.3平台预估窗口是否按距离动态调整★.png'),
           '图 4-7  平台预估天数：RO 38.8 天、PA 37.2 天，而 SP 仅 19.1 天')
add_table(['州', '平台预估', '实际全程', '平均富余', '迟到率'],
          [['RO', '38.8 天', '19.4 天', '19.4 天', '2.9%'],
           ['PA', '37.2 天', '23.8 天', '13.4 天', '12.4%'],
           ['AL', '32.6 天', '24.5 天', '8.0 天', '23.9%'],
           ['SP', '19.1 天', '8.8 天', '10.3 天', '5.9%']],
          widths=[2.4, 3.2, 3.2, 3.2, 3.6], caption='平台预估窗口随距离变化（部分州）',
          right_cols=[1, 2, 3, 4])
callout('迟到率被“保守预估”掩盖了真实问题',
        '偏远州的预估是 SP 的 2 倍（38.8 vs 19.1 天）。RO 收到 38.8 天预估、实际 19.4 天就到，\n'
        '自然几乎不“迟到”（2.9%）——但客户实际仍要等近 20 天。\n'
        '所以：迟到率低 ≠ 送得快。分析必须同时看“绝对时长”和“是否违约”。')

h3('2.4 圣保罗 vs 其他州', '4.3.4')
add_figure(os.path.join(RES, '10_2.4圣保罗vs其他州（对比标杆）.png'),
           '图 4-8  SP 8.8 天/评分 4.25 vs 其他州 15.3 天/评分 4.09')

h2('第 3 步：卖家因素', '4.4')
h3('3.1 备货最慢的卖家', '4.4.1')
add_figure(os.path.join(RES, '11_3.1备货最慢的卖家（样本_=50，避免小样本噪音）.png'),
           '图 4-9  备货最慢的卖家（仅统计订单数≥50 的卖家，避免小样本噪音）')
para('关键在 HAVING 设最小样本量，否则会看到一堆“1 单卖了 60 天”的噪音卖家。', indent=0.85)

h3('3.2 是否超出发货期限', '4.4.2')
para('shipping_limit_date 是卖家承诺的发货截止时间。这是一个二元、可问责的指标——'
     '卖家要么按时发货，要么违约，没有模糊地带。', indent=0.85)
add_figure(os.path.join(RES, '12_3.2卖家是否超出发货期限（shipping_limit_date）★硬违约证据.png'),
           '图 4-10  9.35% 的订单明细超出了卖家承诺的发货期限')
add_figure(os.path.join(RES, '13_3.3超期发货的代价★★全项目最强发现.png'),
           '图 4-11  单因素视角：超期发货订单全程 19.6 vs 11.9 天，评分 3.70 vs 4.20')
callout('注意：这个"最强发现"在第六章被修正',
        '单因素对比显示超期发货影响巨大（评分 3.70 vs 4.20）。\n'
        '但第六章的多元回归表明：控制“是否迟到”后，超期发货的直接效应几乎为零（β=-0.0002, p=0.973）。\n'
        '它是通过“导致迟到”间接影响评分的。这正是多元建模的价值。',
        bg='FDECEA', border='C0392B', tcolor='922B21')

h3('3.3 按“影响面”排序', '4.4.3')
add_figure(os.path.join(RES, '14_3.4按_影响面_排序（超额延误总量=超出均值的天数×订单量）.png'),
           '图 4-12  按影响面排序：备货 11.6 天/980 单的卖家影响面 8,623 天，远超 16.1 天/76 单的 1,010 天')
para('问题该按“影响总量”排序，而不是“单笔最长”。一个拖 60 天但只影响 1 个客户的卖家，'
     '不如一个拖 3 天但影响 1 万个客户的卖家值得改。', indent=0.85)

h3('3.4 多卖家订单：一个必须怀疑的反常结果', '4.4.4')
add_figure(os.path.join(RES, '15_3.5单卖家vs多卖家订单（多卖家需等所有卖家发货）.png'),
           '图 4-13  多卖家订单：全程反而更短（9.2 天）、迟到率 1.5%，但差评率高达 46.34%')
callout('三处反常同时出现，必须怀疑而不是解释',
        '① 多卖家订单反而更快？逻辑上要等多个卖家发货，应该更慢才对。\n'
        '② 三个卖家以上的订单迟到率 0.0%？59 单一个都没迟到，太完美。\n'
        '③ 评分低到 2.24？比超期发货的 3.70 还低得多。\n\n'
        '诊断结论：评分低是真实信号（35%~48% 给一星，全国平均仅 11.5%），\n'
        '但“更快、零迟到”是数据假象——推测分批发货时 order_delivered_customer_date\n'
        '只记录了第一个包裹的签收时间。处理：不采信该指标，写入数据局限。',
        bg='FDECEA', border='C0392B', tcolor='922B21')

h2('第 4 步：商品因素', '4.5')
h3('4.1 商品重量与备货时长', '4.5.1')
add_figure(os.path.join(RES, '16_4.1商品重量→备货时长.png'),
           '图 4-14  商品重量与备货时长单调递增：<0.5kg 为 2.58 天，>15kg 为 4.94 天')
h3('4.2 品类与备货时长', '4.5.2')
add_figure(os.path.join(RES, '17_4.2品类→备货时长（样本_=300）.png'),
           '图 4-15  品类备货时长：office_furniture 达 9.89 天，是全国均值 2.80 天的 3.5 倍')
h3('4.3 混杂因素分层验证', '4.5.3')
add_figure(os.path.join(RES, '19_4.4排除_办公家具_后，卖家备货分化还剩多少★混杂因素分层验证.png'),
           '图 4-16  排除办公家具后：423 个卖家平均备货 2.75 天，但最慢的仍有 16.1 天')

h2('第 5 步：时间因素', '4.6')
add_table(['类型', '问的问题', '怎么看'],
          [['长期趋势', '业务在变好还是变差？', '按月看，注意剔除数据不完整的月份'],
           ['季节性', '每年固定时段是否有波动？', '对比不同年份同一月份'],
           ['短期周期', '一周中哪天、一天中几点下单更慢？', '按周几、按小时看']],
          widths=[3.0, 6.6, 6.6], caption='三种时间效应的区分')
h3('5.1 月度趋势', '4.6.1')
add_figure(os.path.join(RES, '20_5.1月度趋势.png'),
           '图 4-17  月度趋势：2017-11 黑五与 2018-02/03 出现两次异常峰值，2018 下半年显著改善')
para('注意：2016-09（4 单）、2016-12（1 单）、2018-09（16 单）、2018-10（4 单）数据不完整，'
     '分析时应剔除。', indent=0.85)
h3('5.2 两个关键异常', '4.6.2')
add_table(['时间', '现象', '归因'],
          [['2017-11（黑五）', '订单量 7,544（环比 +63%），迟到率 5.3% → 14.3%', '容量冲击：物流扛不住突增'],
           ['2018-02~03', '迟到率 16.0% → 21.4%（全年最差），评分 3.76',
            '运输环节异常：运输天数 10.5 → 13.4（+28%），备货稳定 3.0-3.2']],
          widths=[3.4, 6.8, 6.0], caption='两次时间异常及其归因')
add_figure(os.path.join(RES, '23_5.42018-02_03异常归因：是哪一环节出的问题★.png'),
           '图 4-18  2018-02/03 归因：备货天数始终稳定，是运输天数从 10.5 涨到 13.4 天')
callout('思维原则：不要为异常编故事',
        '可以猜“巴西 2018 年初可能有卡车司机罢工/暴雨/狂欢节”，但数据里没有证据。\n'
        '正确写法：「2018-02~03 运输环节出现全国性异常（运输天数 +28%），原因需业务方确认」。\n'
        '把“不知道”写清楚，比编一个合理的故事更专业。')
h3('5.3 迟到率由“平台承诺”决定', '4.6.3')
add_figure(os.path.join(RES, '21_5.2迟到率由_平台承诺_决定★反直觉的关键洞察.png'),
           '图 4-19  迟到率由“平台承诺”决定：2018-08 送更快（7.7 天）但预估压到 15.2 天，迟到率反升至 10.4%')
add_table(['月份', '平台预估', '实际全程', '平均富余', '迟到率'],
          [['2018-06', '27.7 天', '9.2 天', '18.5 天', '1.4%'],
           ['2018-07', '19.7 天', '9.0 天', '10.7 天', '4.5%'],
           ['2018-08', '15.2 天', '7.7 天', '7.5 天', '10.4%'],
           ['2018-03', '22.1 天', '16.3 天', '5.7 天', '21.4%']],
          widths=[2.6, 3.4, 3.4, 3.4, 3.4], caption='预估窗口 vs 实际速度 vs 迟到率',
          right_cols=[1, 2, 3, 4])
callout('整个项目最反直觉的洞察',
        '迟到率的分子是“物流表现”，分母却是“平台的承诺”。\n'
        '所以降低迟到率有两条路：提速（难），或者放宽承诺（易但损害体验）。\n'
        '这就解释了为什么“迟到率低”不等于“服务好”——2018-06 迟到率仅 1.4%，但承诺的是 27.7 天。')
h3('5.4 周几与时段', '4.6.4')
add_figure(os.path.join(RES, '22_5.3周几下单的影响.png'),
           '图 4-20  周五/周六下单需 13.4-13.6 天，比周日多 1.4 天')
para('下单时段（小时）影响很小（凌晨 13.0 天 vs 上午 12.4 天，仅差 0.6 天），可以忽略。', indent=0.85)

h2('第 6 步：五因素汇总对比', '4.7')
add_figure(os.path.join(RES, '25_6.1五个因素并列对比（统一口径：差评率=1-2星占已评订单比例）.png'),
           '图 4-21  五因素对差评率的影响对比（基线 14.63%）')
add_table(['排名', '因素', '高分组差评率', '低分组差评率', '差距'],
          [['1', '是否迟到', '53.95%（迟到）', '9.18%（准时）', '+44.8 pp'],
           ['2', '运输时长', '47.97%（21天以上）', '8.40%（3天内）', '+39.6 pp'],
           ['3', '备货时长', '41.53%（14天以上）', '9.84%（1天内）', '+31.7 pp'],
           ['4', '卖家超期发货', '23.71%（超期）', '11.68%（按期）', '+12.0 pp'],
           ['5', '是否跨州', '14.07%（跨州）', '10.43%（同州）', '+3.6 pp']],
          widths=[1.4, 4.2, 4.4, 4.4, 2.4], caption='五因素对差评率的影响排序',
          right_cols=[2, 3, 4])

page_break()

# ===========================================================================
# 第五章
# ===========================================================================
h1('统计学验证：相关性与显著性', '第五章')
para('第四章的结论都是“分组对比”，属于描述性统计。本章用相关系数矩阵、显著性检验和效应量，'
     '给出数学层面的证据。分析脚本：analysis/stat_analysis.py，清洗后样本 94,138 单。', indent=0.85)

h2('相关系数矩阵（Pearson / Spearman）', '5.1')
para('同时计算两种相关系数，原因是：Pearson 衡量线性关系，Spearman 衡量单调关系。'
     '本数据多个变量偏度绝对值大于 1（审核时长 4.48、运费 12.90），分布明显右偏，'
     '此时 Pearson 会失效，必须对照 Spearman。', indent=0.85)
add_table(['变量', '偏度', '峰度', '结论'],
          [['审核时长', '4.478', '69.503', '严重右偏'],
           ['备货时长', '3.882', '25.040', '严重右偏'],
           ['运输时长', '1.853', '4.632', '右偏'],
           ['运费', '12.902', '716.354', '极端右偏'],
           ['评分', '−1.504', '1.045', '左偏']],
          widths=[4.0, 3.2, 3.2, 5.6], caption='关键变量的偏度与峰度（|偏度|>1 说明非正态）',
          right_cols=[1, 2])
add_figure(os.path.join(FIG, 'fig1_相关系数热力图.png'),
           '图 5-1  相关系数热力图（左：Pearson 线性；右：Spearman 单调）', max_h=10.5)
add_table(['因素', 'Pearson r', 'p 值', 'Spearman ρ', '显著性'],
          [['是否迟到', '−0.3514', '≈0', '−0.2823', '***'],
           ['全程时长（天）', '−0.3351', '≈0', '−0.2266', '***'],
           ['运输时长（天）', '−0.3015', '≈0', '−0.1877', '***'],
           ['预估富余（天）', '+0.2559', '≈0', '+0.1684', '***'],
           ['备货时长（天）', '−0.1515', '≈0', '−0.1004', '***'],
           ['商品件数', '−0.1224', '≈0', '−0.1063', '***'],
           ['卖家超期发货', '−0.1058', '≈0', '−0.0924', '***'],
           ['运费', '−0.0867', '≈0', '−0.0840', '***'],
           ['是否跨州', '−0.0580', '≈0', '−0.0557', '***'],
           ['商品金额', '−0.0316', '≈0', '−0.0267', '***'],
           ['商品重量（g）', '−0.0248', '≈0', '−0.0093', '***'],
           ['运费占比', '−0.0239', '≈0', '−0.0312', '***'],
           ['审核时长（天）', '−0.0198', '≈0', '−0.0090', '***']],
          widths=[4.6, 3.0, 2.6, 3.0, 3.0], caption='各因素与「评分」的相关系数（n=93,526）',
          right_cols=[1, 2, 3, 4])
add_figure(os.path.join(FIG, 'fig2_各因素相关强度.png'),
           '图 5-2  各因素与「评分」/「是否差评」的相关强度对比')
callout('如何正确解读相关系数',
        '① p 值 < 0.001（***）表示相关显著不为 0，即“有相关”，但不表示“相关很强”。\n'
        '② 本分析样本达 9.4 万，几乎所有 p 值都≈0——大样本下极易显著，必须看 |r| 大小。\n'
        '③ 相关系数只证明共变，不证明因果。',
        bg='EDF4FB', border='4A90D9', tcolor='1E3A5F')

h2('显著性检验与 p 值', '5.2')
add_table(['因素', 'Pearson r', 'p 值', 'Spearman ρ', '判定'],
          [['是否迟到', '−0.3514', '≈0', '−0.2823', '强负相关'],
           ['全程时长', '−0.3351', '≈0', '−0.2266', '强负相关'],
           ['运输时长', '−0.3015', '≈0', '−0.1877', '中等负相关'],
           ['预估富余', '+0.2559', '≈0', '+0.1684', '中等正相关'],
           ['备货时长', '−0.1515', '≈0', '−0.1004', '弱负相关'],
           ['卖家超期发货', '−0.1058', '≈0', '−0.0924', '弱负相关'],
           ['是否跨州', '−0.0580', '≈0', '−0.0557', '极弱']],
          widths=[3.8, 2.8, 2.2, 2.8, 3.2], caption='与「评分」的显著性检验结果',
          right_cols=[1, 2, 3])

h2('效应量 Cohen\'s d', '5.3')
para('p 值只回答“有没有差异”，效应量回答“差异有多大”。用 Cohen\'s d 衡量组间差异的标准化大小：', indent=0.85)
code_block(
    'def cohens_d(a, b):\n'
    '    """效应量：均值差 / 合并标准差。0.2 小 / 0.5 中 / 0.8 大。"""\n'
    '    a, b = np.asarray(a, float), np.asarray(b, float)\n'
    '    na, nb = len(a), len(b)\n'
    '    pooled = np.sqrt(((na - 1) * a.var(ddof=1)\n'
    '                      + (nb - 1) * b.var(ddof=1)) / (na + nb - 2))\n'
    '    return (a.mean() - b.mean()) / pooled if pooled > 0 else np.nan',
    caption='Cohen\'s d 计算函数（Python）')
add_table(['分组变量', '组1（是）', '组0（否）', '均值差', "Cohen's d", '判读'],
          [['是否迟到', 'n=7,125  均值 2.61', 'n=86,401  均值 4.30', '−1.69', '−1.41', '极大'],
           ['卖家超期发货', 'n=8,385  均值 3.74', 'n=85,141  均值 4.21', '−0.47', '−0.37', '小-中'],
           ['是否跨州', 'n=59,722  均值 4.11', 'n=33,804  均值 4.27', '−0.15', '−0.12', '可忽略']],
          widths=[3.2, 4.0, 4.0, 1.8, 2.2, 2.0], caption='组间差异的效应量',
          right_cols=[3, 4, 5])
callout('效应量给出了最直接的数学证据',
        '“是否迟到”的 Cohen\'s d = −1.41，意味着迟到组与准时组的评分分布相差 1.41 个标准差。\n'
        '而“是否跨州”的 d 仅 −0.12，几乎可以忽略。\n'
        '这从数学上否证了“距离是满意度主因”的说法——距离影响的是速度，不是满意度本身。',
        bg='EDF4FB', border='4A90D9', tcolor='1E3A5F')
add_figure(os.path.join(FIG, 'fig5_组间对比箱线图.png'),
           '图 5-3  迟到 / 超期发货两组的评分箱线图（标题标注 Cohen\'s d）')

h2('分类变量卡方检验', '5.4')
add_table(['变量 1', '变量 2', '卡方值', 'p 值', "Cramér's V", '解读'],
          [['是否迟到', '是否差评', '11,482.3', '0.000e+00', '0.3504', '中等关联'],
           ['卖家超期发货', '是否差评', '885.3', '1.541e−194', '0.0973', '弱关联'],
           ['是否跨州', '是否差评', '213.4', '2.458e−48', '0.0478', '极弱'],
           ['卖家超期发货', '是否迟到', '3,170.5', '0.000e+00', '0.1835', '弱-中'],
           ['是否跨州', '是否迟到', '234.6', '6.014e−53', '0.0499', '极弱']],
          widths=[3.2, 3.2, 2.8, 3.0, 2.6, 2.4],
          caption='卡方检验与 Cramér\'s V（0.1 弱 / 0.3 中 / 0.5 强）', right_cols=[2, 3, 4])

h2('阈值效应验证：为什么相关系数会低估', '5.5')
para('这是本次分析在统计学上最有价值的部分。它说明：即使相关系数看着不高（r=0.256），'
     '真实关系可能非常强——只是形态不是线性的。', indent=0.85)
add_figure(os.path.join(FIG, 'fig3_关键散点与趋势.png'),
           '图 5-4  散点与分箱均值曲线：趋势在 20 天左右开始加速下滑，呈明显非线性')
add_table(['子集', 'Pearson r', 'p 值', '解读'],
          [['全部样本（n=94,138）', '0.2559', '≈0', '看起来中等相关'],
           ['仅未超期子集（n=87,013）', '0.0241', '1.38e−12', '★ 关系几乎消失']],
          widths=[6.2, 3.2, 3.0, 3.8], caption='分段相关系数：整体 vs 未超期子集', right_cols=[1, 2])
add_table(['预估富余档', '订单数', '差评率', '平均评分'],
          [['已超期（负富余）', '7,125', '52.77%', '2.61'],
           ['富余 0–3 天', '4,653', '11.11%', '4.14'],
           ['富余 3–7 天', '12,371', '9.82%', '4.23'],
           ['富余 7–15 天', '40,718', '8.77%', '4.32'],
           ['富余 15–30 天', '26,390', '8.92%', '4.33'],
           ['富余 30 天以上', '2,269', '10.71%', '4.26']],
          widths=[5.0, 3.6, 3.6, 4.0], caption='预估富余分档与差评率（阈值效应的直接证据）',
          right_cols=[1, 2, 3])
add_figure(os.path.join(FIG, 'fig4_悬崖效应.png'),
           '图 5-5  悬崖效应：已超期组差评率 52.8%，未超期各档仅在 8.8%–11.1% 之间持平')
callout('统计学的核心结论',
        '整体相关系数 r = 0.2559 完全由“已超期”这 7.6% 的样本驱动。\n'
        '在未超期范围内，富余天数与评分基本无关（r = 0.0241）。\n\n'
        '真实关系是「阈值型（悬崖）」，不是「线性型（斜坡）」。\n'
        '业务含义完全不同：若是斜坡 → 应追求“越送越快”；若是悬崖 → 应追求“绝不违约”。',
        bg='FDECEA', border='C0392B', tcolor='922B21')

page_break()

# ===========================================================================
# 第六章（新增）多元建模
# ===========================================================================
h1('多元统计建模：控制混杂后的净效应', '第六章')
para('第五章的相关分析是“两两之间”的关系。但现实中多个因素同时起作用，彼此纠缠。'
     '本章用多元模型同时控制所有变量，估计每个因素的“净效应”。', indent=0.85)

h2('为什么必须做多元模型', '6.1')
para('单因素相关可能被三种机制扭曲，这是本章的出发点：', indent=0.85)
add_table(['机制', '含义', '本次分析中的实例'],
          [['混杂（Confounding）', '第三个变量同时影响自变量和因变量，制造出虚假关联',
            '偏远州既距离远、物流也差'],
           ['中介（Mediation）', '自变量通过中间变量起作用，直接效应为零但有间接效应',
            '★超期发货 → 迟到 → 差评（第六章将证明）'],
           ['抑制（Suppression）', '控制某变量后，相关系数的符号发生反转',
            '★距离与评分的关系（第六章将证明）']],
          widths=[3.6, 6.4, 6.2], caption='三种会扭曲单因素结论的机制')
para('这三种机制在本次数据中全部出现了。如果只做单因素分析，会得出三个错误结论。', indent=0.85)

h2('特征工程与真实地理距离', '6.2')
para('第四章的“距离”只能用“是否跨州”粗略代替。本章用 geolocation 表的经纬度，'
     '通过 Haversine 公式计算客户与卖家的真实直线距离。', indent=0.85)
code_block(
    'def haversine(lat1, lon1, lat2, lon2):\n'
    '    """Haversine 公式：返回两点间大圆距离（km）。"""\n'
    '    R = 6371.0\n'
    '    p1, p2 = np.radians(lat1), np.radians(lat2)\n'
    '    dp = p2 - p1\n'
    '    dl = np.radians(lon2 - lon1)\n'
    '    a = (np.sin(dp / 2) ** 2\n'
    '         + np.cos(p1) * np.cos(p2) * np.sin(dl / 2) ** 2)\n'
    '    return 2 * R * np.arcsin(np.sqrt(a))',
    caption='A-1  真实距离计算（Haversine）')
add_table(['指标', '数值'],
          [['邮编坐标表覆盖', '19,015 个邮编'],
           ['成功匹配经纬度的订单', '95,983 条（99.5%）'],
           ['距离均值', '601 km'],
           ['距离中位数', '434 km'],
           ['距离 P90', '1,457 km'],
           ['距离最大值', '8,678 km'],
           ['建模样本（清洗后）', '93,047 条'],
           ['特征维度', '10 维']],
          widths=[8.4, 7.8], caption='距离特征与建模样本概况')
para('对距离、重量、金额做了 log(1+x) 变换以降低右偏；所有连续变量标准化，'
     '使回归系数可直接比较大小。', indent=0.85)

h2('多元线性回归（OLS）', '6.3')
para('模型：评分 = β₀ + β₁·备货时长 + β₂·运输时长 + β₃·是否迟到 + β₄·超期发货 + '
     'β₅·距离 + β₆·跨州 + β₇·重量 + β₈·金额 + β₉·件数 + β₁₀·运费占比 + ε', indent=0.85)
add_table(['整体指标', '数值', '含义'],
          [['样本量 N', '93,047', '—'],
           ['R²', '0.1695', '模型解释了 16.95% 的评分方差'],
           ['调整 R²', '0.1694', '几乎无衰减，说明没有过拟合'],
           ['F 统计量', '1,898.8', '整体模型高度显著（p < 1e-300）'],
           ['残差标准误', '1.1601', '评分残差约 ±1.16 分']],
          widths=[3.6, 3.2, 9.4], caption='OLS 模型整体拟合情况', right_cols=[1])
callout('R² 只有 0.17，是不是模型没用？',
        '不是。评分为 1-5 分的整数主观打分，本身含大量随机性（个人偏好、商品质量、'
        '客服态度等未观测因素）。\n'
        '在消费者行为研究中，R² 在 0.10–0.25 属于正常且可用的水平。\n'
        '关键在于：模型揭示的方向和相对重要性是稳定的（第六章 6.7 稳健性检验可证），\n'
        '而不是用 R² 去预测单个客户的评分。')

para('标准化回归系数（可直接比较影响大小，按 |β| 降序）：', indent=0.85)
add_table(['排名', '变量', '标准化 β', '标准误', 't 值', 'p 值', '显著性'],
          [['1', '是否迟到', '−0.2771', '0.0049', '−56.0', '<1e-300', '***'],
           ['2', '运输时长(天)', '−0.2749', '0.0056', '−49.1', '<1e-300', '***'],
           ['3', '商品件数', '−0.1642', '0.0041', '−40.4', '<1e-300', '***'],
           ['4', '备货时长(天)', '−0.1310', '0.0048', '−27.2', '<1e-300', '***'],
           ['5', '距离(对数km)', '+0.0636', '0.0061', '10.5', '<1e-300', '***'],
           ['6', '是否跨州', '+0.0151', '0.0058', '2.6', '0.0088', '**'],
           ['7', '重量(对数g)', '+0.0123', '0.0045', '2.7', '0.0063', '**'],
           ['8', '金额(对数)', '−0.0063', '0.0062', '−1.0', '0.3042', 'ns'],
           ['9', '运费占比', '−0.0041', '0.0055', '−0.7', '0.4564', 'ns'],
           ['10', '卖家超期发货', '−0.0002', '0.0048', '−0.0', '0.9730', 'ns']],
          widths=[1.2, 4.0, 2.4, 2.2, 1.8, 2.4, 2.2],
          caption='OLS 标准化系数（因变量：评分）', right_cols=[2, 3, 4, 5, 6])

para('共线性诊断（VIF）：所有变量 VIF 均小于 3，远低于经验阈值 10，说明不存在多重共线性问题，'
     '系数估计稳定可靠。', indent=0.85)
add_table(['变量', 'VIF', '判定'],
          [['金额(对数)', '2.63', '可接受'],
           ['距离(对数km)', '2.55', '可接受'],
           ['是否跨州', '2.30', '可接受'],
           ['运输时长(天)', '2.16', '可接受'],
           ['运费占比', '2.09', '可接受'],
           ['是否迟到', '1.69', '可接受'],
           ['备货时长(天)', '1.60', '可接受'],
           ['卖家超期发货', '1.56', '可接受'],
           ['重量(对数g)', '1.40', '可接受'],
           ['商品件数', '1.14', '可接受']],
          widths=[5.2, 3.2, 7.8], caption='方差膨胀因子（VIF < 10 为安全）', right_cols=[1])

para('业务量纲解读（非标准化系数）：', indent=0.85)
add_table(['变量', '系数', '业务含义'],
          [['是否迟到', '−1.045', '迟到订单较准时订单，评分平均低 1.05 分'],
           ['商品件数', '−0.313', '每多买 1 件商品，评分下降 0.31 分'],
           ['备货时长', '−0.040', '备货每多 1 天，评分下降 0.04 分'],
           ['运输时长', '−0.037', '运输每多 1 天，评分下降 0.04 分'],
           ['距离(对数)', '+0.046', '控制运输时长后，距离越远评分反而略高（见 6.8）']],
          widths=[3.4, 2.4, 10.4], caption='关键变量的业务量纲解释', right_cols=[1])

h2('二元逻辑回归（Logit）与优势比', '6.4')
para('OLS 回答“评分降多少”，但业务更关心“差评的几率翻几倍”。逻辑回归给出优势比（Odds Ratio）。', indent=0.85)
add_table(['整体指标', '数值', '含义'],
          [['样本量 N', '93,047', '—'],
           ['McFadden 伪 R²', '0.1518', '相当于中等拟合优度'],
           ['LR 统计量', '10,629.7', 'df=10，p < 1e-300'],
           ['基准差评率', '12.47%', '—']],
          widths=[3.8, 3.2, 9.2], caption='Logit 模型整体情况', right_cols=[1])

add_table(['变量', 'OR', '95% 置信区间', 'p 值', '显著性'],
          [['运输时长(天)', '1.603', '[1.562, 1.646]', '<1e-300', '***'],
           ['是否迟到', '1.457', '[1.428, 1.486]', '<1e-300', '***'],
           ['商品件数', '1.351', '[1.327, 1.375]', '<1e-300', '***'],
           ['备货时长(天)', '1.252', '[1.225, 1.280]', '<1e-300', '***'],
           ['距离(对数km)', '0.848', '[0.820, 0.878]', '<1e-300', '***'],
           ['金额(对数)', '1.089', '[1.054, 1.125]', '2.81e−07', '***'],
           ['重量(对数g)', '0.947', '[0.924, 0.971]', '1.67e−05', '***'],
           ['运费占比', '1.034', '[1.008, 1.062]', '0.0108', '*'],
           ['卖家超期发货', '1.013', '[0.990, 1.037]', '0.2756', 'ns'],
           ['是否跨州', '0.996', '[0.964, 1.030]', '0.8237', 'ns']],
          widths=[4.2, 2.2, 4.6, 2.6, 2.4],
          caption='Logit 优势比（每增加 1 个标准差，差评几率的变化倍数）', right_cols=[1, 2, 3, 4])

callout('优势比怎么读',
        'OR = 1.603（运输时长）表示：运输时长每增加 1 个标准差，差评的几率变为原来的 1.60 倍，即增加 60%。\n'
        'OR = 1.457（是否迟到）表示：迟到订单的差评几率是准时订单的 1.46 倍。\n'
        'OR = 0.848（距离）表示：距离每增加 1 个标准差，差评几率反而降低 15%（这违反直觉，见 6.8）。',
        bg='EDF4FB', border='4A90D9', tcolor='1E3A5F')

para('在均值处的边际效应（对差评概率的百分点影响）：', indent=0.85)
add_table(['变量', '差评概率变化'],
          [['运输时长(天)', '+5.15 个百分点 / 每标准差'],
           ['是否迟到', '+4.11 个百分点 / 每标准差'],
           ['商品件数', '+3.28 个百分点 / 每标准差'],
           ['备货时长(天)', '+2.45 个百分点 / 每标准差'],
           ['距离(对数km)', '−1.80 个百分点 / 每标准差']],
          widths=[6.4, 9.8], caption='边际效应（在均值处）', right_cols=[1])

h2('随机森林对照与预测力评估', '6.5')
para('前两个模型是线性的。随机森林不假设函数形式，可作为稳健性对照，并给出特征重要性排序。', indent=0.85)
add_table(['变量', '重要性', '占比', '累计占比'],
          [['是否迟到', '0.3960', '39.60%', '39.60%'],
           ['运输时长(天)', '0.3033', '30.33%', '69.93%'],
           ['商品件数', '0.1003', '10.03%', '79.96%'],
           ['备货时长(天)', '0.0710', '7.10%', '87.06%'],
           ['距离(对数km)', '0.0408', '4.08%', '91.14%'],
           ['运费占比', '0.0241', '2.41%', '93.55%'],
           ['金额(对数)', '0.0228', '2.28%', '95.83%'],
           ['重量(对数g)', '0.0200', '2.00%', '97.83%'],
           ['卖家超期发货', '0.0135', '1.35%', '99.18%'],
           ['是否跨州', '0.0081', '0.81%', '99.99%']],
          widths=[4.2, 3.0, 3.0, 3.4],
          caption='随机森林特征重要性（n_estimators=300, max_depth=12）',
          right_cols=[1, 2, 3])
add_table(['评估指标', '数值'],
          [['5 折交叉验证 AUC', '0.7439 ± 0.0034'],
           ['基线 AUC（随机猜测）', '0.5000'],
           ['提升幅度', '+48.8%']],
          widths=[8.4, 7.8], caption='随机森林预测性能', right_cols=[1])
callout('三种方法结论高度一致，这是最强的证据',
        'OLS 标准化系数排序：迟到(−0.277) > 运输(−0.275) > 件数(−0.164) > 备货(−0.131)\n'
        'Logit 优势比排序：运输(1.60) > 迟到(1.46) > 件数(1.35) > 备货(1.25)\n'
        '随机森林重要性：迟到(39.6%) + 运输(30.3%) = 70%\n\n'
        '三种完全不同原理的方法，都把“迟到”和“运输时长”排在最前两位，\n'
        '且都认为“跨州”和“超期发货”几乎不重要。结论稳健。',
        bg='E8F5E9', border='43A047', tcolor='1B5E20')

h2('交互效应检验', '6.6')
para('前面假设各因素独立作用。但现实可能是“距离远时，迟到更不可接受”。用交互项检验：', indent=0.85)
add_table(['模型', '变量', 'β', 't 值', 'p 值', '显著性'],
          [['模型 A', '是否迟到', '−0.779', '−11.61', '<1e-300', '***'],
           ['', '距离(对数)', '−0.028', '−9.39', '<1e-300', '***'],
           ['', '★ 迟到 × 距离', '−0.150', '−13.70', '<1e-300', '***'],
           ['模型 B', '是否迟到', '−1.691', '−98.78', '<1e-300', '***'],
           ['', '超期发货', '−0.227', '−14.67', '<1e-300', '***'],
           ['', '★ 迟到 × 超期', '+0.186', '5.26', '1.46e−07', '***']],
          widths=[1.8, 4.0, 2.4, 2.2, 2.8, 2.4],
          caption='交互效应检验（因变量：评分）', right_cols=[2, 3, 4, 5])
callout('交互效应说明了两件事',
        '① 迟到 × 距离 = −0.150（显著为负）：距离越远，迟到对评分的伤害越大。\n'
        '   业务含义：远距离订单一旦迟到，客户更愤怒——因为等待本来就长。\n'
        '   建议：对远距离订单要额外谨慎，宁可放宽承诺也不要违约。\n\n'
        '② 迟到 × 超期 = +0.186（显著为正）：两者同时发生时，负面影响小于“简单相加”。\n'
        '   说明这两个因素有部分重叠（超期本身就是迟到的原因），不是独立的双重打击。')

h2('稳健性检验', '6.7')
para('为确认结论不是某一年数据偶然造成，按年份拆分重新估计：', indent=0.85)
add_table(['子样本', 'N', '迟到 β', '距离 β', '超期 β', 'R²'],
          [['2017 年', '42,331', '−0.2528', '+0.0612', '−0.0051', '0.1408'],
           ['2018 年', '50,456', '−0.2937', '+0.0626', '+0.0041', '0.1931'],
           ['全样本', '93,047', '−0.2771', '+0.0636', '−0.0002', '0.1695']],
          widths=[3.0, 2.8, 2.8, 2.6, 2.6, 2.4],
          caption='分年度稳健性检验（标准化系数）', right_cols=[1, 2, 3, 4, 5])
para('三个关键结论在两年间完全一致：迟到系数稳定在 −0.25 至 −0.29，距离系数稳定在 +0.061 至 '
     '+0.064，超期发货系数都不显著。说明结论稳健，不是特定年份的偶然现象。', indent=0.85)

h2('两处颠覆单因素结论的发现', '6.8')
para('这是本章最重要的部分——多元建模直接推翻了第四章的两个判断。', indent=0.85)

h3('发现一：超期发货的效应被“迟到”完全中介', '6.8.1')
add_table(['分析方式', '结果', '结论'],
          [['单因素相关', 'r = −0.1058（p≈0）', '看似显著，评分低 0.5 分'],
           ['单因素分组', '评分 3.70 vs 4.20', '看起来影响很大'],
           ['多元回归（控制迟到）', 'β = −0.0002，p = 0.973', '★ 直接效应几乎为零'],
           ['Logit（控制迟到）', 'OR = 1.013，p = 0.276', '★ 优势比不显著']],
          widths=[4.4, 5.6, 6.2], caption='超期发货在单因素与多因素下的对比')
callout('因果路径：超期发货 → 迟到 → 差评',
        '第四章认为“超期发货是最大痛点”，这个判断不准确。\n'
        '真实机制是：卖家超期发货（shipping_limit_date）会导致订单迟到（estimated_delivery_date），\n'
        '而客户真正在意的是“迟到”本身。一旦控制住“是否迟到”，超期发货就失去了独立解释力。\n\n'
        '这叫做“完全中介效应”。业务含义：\n'
        '  · 考核卖家发货时效仍然有意义——因为它是迟到的上游原因；\n'
        '  · 但真正的 KPI 应该是“准时送达率”，而不是“按时发货率”；\n'
        '  · 如果只考核发货率，卖家可能提前发货但物流仍慢，客户依然不满意。',
        bg='FDECEA', border='C0392B', tcolor='922B21')

h3('发现二：距离的符号反转（抑制效应）', '6.8.2')
add_table(['分析方式', '距离与评分的关系', '解释'],
          [['Pearson 相关', 'r = −0.0580', '距离越远，评分越低（弱负相关）'],
           ['单因素 d', '−0.12', '跨州 vs 同州差异可忽略'],
           ['多元回归（控制运输时长）', 'β = +0.0636（p<1e-300）', '★ 符号反转为正'],
           ['Logit（控制运输时长）', 'OR = 0.848（p<1e-300）', '★ 距离越远，差评几率反而更低']],
          widths=[4.6, 5.4, 6.2], caption='距离效应在单因素与多因素下的对比')
callout('为什么会符号反转？（典型抑制效应）',
        '距离与运输时长高度相关（r = 0.44）。距离远 → 运输慢 → 评分低，这是距离的“间接负效应”。\n'
        '但在多元模型中，当我们已经控制了“运输时长”这个中介变量，剩下的就是“同等运输时长下，\n'
        '远距离 vs 近距离还有什么差别”——结果发现远距离订单评分反而略高。\n\n'
        '可能的解释（需进一步验证）：\n'
        '  · 偏远地区客户对物流预期更低，同样的等待时间容忍度更高；\n'
        '  · 平台给偏远地区的预估窗口更宽松，客户看到“38 天”时就已经调整了期待；\n'
        '  · 能配送到偏远地区的卖家可能本身服务质量更好。\n\n'
        '这个发现的方法论意义：不能用单因素相关系数断言“距离导致不满”。',
        bg='EDF4FB', border='4A90D9', tcolor='1E3A5F')

h2('模型的可解释性与边界', '6.9')
para('必须诚实说明模型能做什么、不能做什么：', indent=0.85)
add_table(['能做', '不能做'],
          [['判断各因素的相对重要性排序（三种方法一致）', '预测单个客户会给几星（R² 仅 0.17）'],
           ['估计控制混杂后的净效应方向和大小', '证明严格因果（缺少随机实验）'],
           ['识别中介路径和交互效应', '解释未观测因素（商品质量、客服态度等）'],
           ['为业务改进提供优先级依据', '适用于其他平台（地域与文化差异大）']],
          widths=[8.0, 8.2], caption='模型的适用边界')
add_figure(os.path.join(FIG, 'fig8_模型解释力对比.png'),
           '图 6-1  逐步加入变量后的解释力提升：仅时间变量 R²=0.11 → 加入违约 0.16 → 全模型 0.17')

page_break()

# ===========================================================================
# 第七章
# ===========================================================================
h1('结论与公司建议', '第七章')

h2('三条核心结论', '7.1')
h3('结论一：客户不是不能等，而是不能接受承诺落空', '7.1.1')
para('“是否迟到”的标准化 β = −0.277（OLS 排名第一）、OR = 1.457（Logit 排名第二）、'
     '随机森林重要性 39.6%（排名第一）、Cohen\'s d = −1.41（极大效应）、'
     '差评率 53.95% vs 9.18%（5.9 倍）。', indent=0.85)
para('多种方法一致指向同一结论。且在未超期范围内，多等几天对评分几乎没影响（r = 0.0241）。'
     '满意度管理的重点是“守住承诺”，而不是“盲目提速”。', indent=0.85)

h3('结论二：最大瓶颈是运输，但最该改的是卖家备货', '7.1.2')
para('运输占全程 74.3%，是绝对瓶颈；但它主要由距离决定，平台无法改变。'
     '而备货只占 22.3%，却完全在平台和卖家的控制范围内（β = −0.131，OR = 1.252）。', indent=0.85)
para('特别是 9.35% 的订单超出了卖家承诺的发货期限——这是明确违约，责任清晰、可考核、可改进。'
     '虽然它的直接效应被“迟到”中介，但作为迟到的上游原因，仍是最可控的干预点。', indent=0.85)

h3('结论三：问题高度集中，精准打击收益最大', '7.1.3')
para('按“影响面”排序，一个卖家（980 单、备货 11.6 天）贡献了 8,623 天的超额延误，'
     '是第二名的 4.5 倍。不需要大规模整改，锁定少数高影响卖家即可获得大部分改善。', indent=0.85)

h2('公司可执行建议（分优先级）', '7.2')
add_table(['优先级', '措施', '数据依据', '责任方', '预期效果', '难度'],
          [['P0', '建立卖家发货时效考核，超期扣分/降权', '9.35% 超期；超期单差评率 23.71%（基线 14.63%）',
            '商家运营', '超期率降到 5%，可减少约 1/3 相关差评', '低'],
           ['P0', '停止为“好看”而收紧预估窗口', '2018-08 预估压至 15.2 天，送更快(7.7 天)但迟到率反升至 10.4%',
            '产品/算法', '迟到率可降 5-8pp', '低'],
           ['P1', '大件/家具品类设独立承诺标准', 'office_furniture 备货 9.89 天（均值 3.5 倍）',
            '品类运营', '避免卖家被误判为差卖家', '低'],
           ['P1', '旺季（黑五/圣诞）提前扩容运力', '2017-11 订单 +63%，迟到率 5.3%→14.3%',
            '供应链', '旺季迟到率减半', '中'],
           ['P2', '聚焦 Top 卖家治理（按影响面排序）', '单一卖家贡献 8,623 天超额延误（第二名 4.5 倍）',
            '商家运营', '集中资源，ROI 最高', '低'],
           ['P2', '多卖家订单体验专项', '2 卖家订单差评率 46.3%（全国 12.5%）',
            '产品', '该场景占 1.3%，但体验极差', '中'],
           ['P3', '远距离订单主动沟通预期', '迟到×距离交互 β=−0.150，远距离迟到伤害更大',
            '客服', '降低期望落差', '低'],
           ['—', '不建议盲目在偏远州投运力', '距离净效应 +0.064，d=−0.12',
            '—', '投入产出比低', '—']],
          widths=[1.3, 4.4, 5.2, 1.8, 2.6, 1.1],
          caption='按优先级排列的公司建议')

callout('三条建议背后的同一个逻辑',
        'P0 的两条建议成本都极低（一个考核规则、一个参数调整），却能直接影响最大的因素。\n'
        '而“在偏远州建仓/加运力”这种高投入方案，数据表明收益有限（距离净效应甚至是正的）。\n\n'
        '分析的价值就在于：把有限资源投到“影响大且可控”的地方。',
        bg='E8F5E9', border='43A047', tcolor='1B5E20')

h2('指标体系建议', '7.3')
add_table(['层级', '指标', '建议目标', '说明'],
          [['北极星指标', '差评率（1-2 星占比）', '从 14.63% 降至 10% 以内', '直接反映客户满意度'],
           ['过程指标', '准时送达率', '从 92.28% 提升至 96%', '★ 最重要，对应“迟到”这个最强因素'],
           ['过程指标', '卖家超期发货率', '从 9.35% 降至 5% 以内', '迟到的上游原因，可考核'],
           ['过程指标', '平均备货天数', '从 2.80 天降至 2.2 天', '完全可控'],
           ['预警指标', '预估窗口收紧幅度 ÷ 实际提速幅度', '比值 ≤ 1', '★ 防止为数据好看而收紧承诺'],
           ['分群指标', '各州 / 各品类准时率', '差异化目标', '避免用全国平均掩盖结构问题']],
          widths=[2.4, 5.4, 4.4, 4.0], caption='建议的指标体系')
para('特别说明“预警指标”这一项：2018 年 8 月的数据显示，平台把预估从 19.7 天压到 15.2 天，'
     '实际配送反而更快了（9.0 → 7.7 天），但迟到率从 4.5% 涨到 10.4%。'
     '如果不监控“承诺 vs 实际的相对变化”，这类问题会被误判为“物流变差”。', indent=0.85)

h2('落地路线图', '7.4')
add_table(['阶段', '时间', '动作', '预期收益'],
          [['第一阶段\n（快速见效）', '0–1 个月',
            '① 上线卖家超期发货考核\n② 冻结预估窗口收紧操作\n③ 建立差评率周报',
            '超期率 -2pp，迟到率 -3pp'],
           ['第二阶段\n（结构优化）', '1–3 个月',
            '① 按影响面排序治理 Top 50 卖家\n② 大件品类独立承诺标准\n③ 多卖家订单体验排查',
            '差评率 -2pp'],
           ['第三阶段\n（长期建设）', '3–6 个月',
            '① 旺季运力弹性预案\n② 分州差异化承诺策略\n③ 远距离订单主动通知机制',
            '旺季迟到率减半']],
          widths=[2.6, 2.2, 7.4, 4.0], caption='三阶段落地路线图')

page_break()

# ===========================================================================
# 第八章（新增）简历表述
# ===========================================================================
h1('简历表述参考', '第八章')
para('本章提供三种颗粒度的表述，按投递场景选用。所有数字均来自本报告的实际分析结果，'
     '不含虚构，面试时可被追问。', indent=0.85)

h2('简历项目条目（完整版）', '8.1')
para('适合：简历中“项目经历”栏，篇幅允许 4-5 行时使用。', indent=0.85)
callout('电商客户满意度归因分析（个人项目）',
        '技术栈：PostgreSQL · SQL · Python（pandas / SciPy / scikit-learn / Matplotlib）\n'
        '数据规模：9 张表 / 155 万行真实交易数据（Olist 巴西电商公开数据集）\n\n'
        '· 数据工程：将 9 个 CSV（155 万行）导入 PostgreSQL，设计 schema、主外键与索引，'
        '建立 3 个分析视图，沉淀 31 条可复用分析 SQL。\n'
        '· 分析框架：提出“沿业务流程拆解”的归因方法，把配送拆为审核/备货/运输三段，'
        '识别运输环节占全程 74.3%、是核心瓶颈。\n'
        '· 统计验证：对 9.4 万笔已送达订单计算 Pearson/Spearman 相关系数矩阵、'
        "Cohen's d 效应量与卡方检验；用 OLS 与 Logistic 回归控制混杂因素，"
        '10 维特征模型交叉验证 AUC 0.744。\n'
        '· 核心发现：定位“是否迟到”为差评最强预测因子（标准化 β=−0.277，OR=1.457，'
        "Cohen's d=−1.41），迟到订单差评率 53.95% vs 准时 9.18%；"
        '并证明该关系为阈值型而非线性（超期组 52.8%，未超期仅 8.8%–11.1%）。\n'
        '· 产出：41 张图表、10 页统计报告，提出卖家时效考核与承诺策略两项 P0 级建议。',
        bg='F3F4F6', border='9CA3AF', tcolor='1F2937', size=10.5)

h2('简历精简版（推荐）', '8.2')
para('适合：简历空间紧张时的标准写法，2-4 行。', indent=0.85)
callout('电商客户满意度归因分析 | PostgreSQL · Python · 统计建模',
        '基于 155 万行真实交易数据构建完整分析链路：把配送流程拆解为三段并定位运输为瓶颈'
        '（占 74.3%）；通过相关分析与多元回归控制混杂因素，发现“是否履约”是差评最强'
        "预测因子（标准化 β=−0.277，Cohen's d=−1.41，迟到订单差评率 53.95% vs 准时 9.18%），"
        '并证伪“距离是满意度主因”的直觉判断；据此提出卖家时效考核等 P0 级建议，'
        '产出 41 张图表分析报告。',
        bg='F3F4F6', border='9CA3AF', tcolor='1F2937', size=10.5)

h2('STAR 面试讲述版', '8.3')
para('适合：面试被问“介绍一个你做的数据分析项目”时口头讲述，时长约 2 分钟。', indent=0.85)
add_table(['要素', '内容'],
          [['Situation\n背景',
            '看到一份巴西电商的公开数据集（99,441 笔订单），想练习完整的数据分析流程，'
            '但一开始不知道该分析什么。'],
           ['Task\n任务',
            '把一个模糊的“看看这份数据”转成一个可交付的问题：哪些因素导致客户差评？'
            '权重如何？业务该优先改哪个？'],
           ['Action\n行动',
            '① 先把数据导入 PostgreSQL，建视图减少重复 JOIN；\n'
            '② 用“沿流程拆解”的方法定位瓶颈，发现运输占 74%，不是支付；\n'
            '③ 逐个排查地理、卖家、商品、时间四类因素，用分组对比控制变量；\n'
            '④ 发现单因素结论不可靠（超期发货看似影响大），于是补做多元回归控制混杂；\n'
            '⑤ 用 OLS、Logit、随机森林三种方法交叉验证，确保结论稳健。'],
           ['Result\n结果',
            '① 找到最强因素：是否迟到（β=−0.277，d=−1.41）；\n'
            '② 推翻两个错误直觉：超期发货的效应被“迟到”完全中介、距离的相关系数符号反转；\n'
            '③ 证明关系是阈值型而非线性——业务含义从“追求更快”变成“绝不违约”；\n'
            '④ 产出 41 张图表 + 10 页报告，给出 P0 级建议。']],
          widths=[2.6, 13.6], caption='STAR 结构讲述稿', font_size=9.8)

h2('技能关键词与注意事项', '8.4')
print('')
add_table(['类别', '可写进简历的关键词'],
          [['数据库', 'PostgreSQL、SQL（多表 JOIN、CTE、窗口函数、聚合、索引优化）'],
           ['Python', 'pandas、NumPy、SciPy、scikit-learn、Matplotlib、Seaborn'],
           ['统计方法', '描述性统计、相关系数（Pearson/Spearman）、显著性检验、效应量（Cohen\'s d）、'
                        '卡方检验、OLS 回归、Logistic 回归、随机森林、交叉验证'],
           ['工程能力', '数据清洗、特征工程（对数变换/标准化）、schema 设计、视图建模、'
                        '结果可视化、报告撰写'],
           ['业务能力', '指标体系设计、A/B 思维、归因分析、优先级排序、ROI 评估']],
          widths=[2.4, 13.8], caption='技能关键词建议')

callout('写简历时的四点注意',
        '① 数字要真实可复现。本文所有数字都来自实际代码，面试官追问时能打开脚本解释。\n'
        '② 不要写“精通”。用“使用过 / 完成过”更安全，被追问时才不会露怯。\n'
        '③ 强调“为什么”而不只是“做了什么”。比如“因为单因素结论不可靠，所以补做多元回归”\n'
        '   这句话本身就能体现分析思维，比列十个工具名更有说服力。\n'
        '④ 主动说明局限。说“这个模型只能做群体归因、不能预测个体评分（R²=0.17）”，\n'
        '   反而显得专业，因为说明你真的理解模型边界。',
        bg='FFF8E1', border='E8A33D', tcolor='8A5A00')

page_break()

# ===========================================================================
# 第九章
# ===========================================================================
h1('数据局限与分析反思', '第九章')
para('一份诚实的分析必须说明自己没做到什么、哪里可能出错。', indent=0.85)

h2('数据本身的局限', '9.1')
add_table(['局限', '具体影响', '本报告的处理'],
          [['订单只有一个签收时间戳', '分批发货时无法反映最晚送达，多卖家订单指标失真',
            '不采信该指标，并说明原因'],
           ['缺少物流商字段', '无法归因到具体承运方', '无法分析，作为局限说明'],
           ['评论缺失率高（标题 88.3%、内容 58.7%）', '无法做文本情感分析', '只使用评分，不用文本'],
           ['4 个月份数据不完整', '月度趋势首尾失真', '分析时剔除'],
           ['23 条负值时间戳', '污染时长计算', '剔除（占 0.02%）'],
           ['非 delivered 订单占 3%', '销售额口径需明确', '分析仅针对已送达订单'],
           ['geolocation 重复采样', '100 万行仅 19,015 个邮编', '按邮编聚合取均值']],
          widths=[4.8, 6.2, 5.2], caption='数据局限性清单')

h2('分析方法的局限', '9.2')
bullet('因果推断有限：虽然“迟到 → 差评”具备时间先后和剂量反应关系，但缺乏随机对照实验'
       '或工具变量，严格意义上只能说“强相关 + 合理机制”，不能断言因果。')
bullet('R² 仅 0.17：评分为主观打分，大量未观测因素（商品质量、客服态度、包装）无法进入模型。'
       '模型适合做群体归因，不适合预测个体。')
bullet('未纳入卖家/品类固定效应：本报告用距离和件数等变量控制，但没有为每个卖家、'
       '每个品类加虚拟变量。加入后 R² 会提升，但会牺牲可解释性。')
bullet('未验证的异常：2018-02/03 的运输异常原因不明（数据内无解释变量），'
       '需要业务方提供当时的外部事件信息。')

h2('如果重做一次，我会怎么做', '9.3')
bullet('第一步就把“差评率”定为核心因变量，而不是中途才从“配送时长”切换过来。'
       '前期有一半精力花在过程指标上，最后还要换算成差评才能比较权重。')
bullet('更早做多元建模。我在做完所有单因素分析后才补回归，结果发现两个结论需要修正'
       '（超期发货被中介、距离符号反转）。如果先建模再细查，能少走弯路。')
bullet('对多卖家订单的异常，应该第一时间查数据字典而不是先做统计——'
       '这类“指标失效”问题，业务理解比统计方法更有效。')
bullet('把随机种子、清洗阈值等参数外置成配置，便于复现和敏感性分析。目前是硬编码在脚本里。')

page_break()

# ===========================================================================
# 附录
# ===========================================================================
h1('完整代码与文件清单', '附录')

h2('SQL 查询文件', 'A.1')
para('全部 SQL 汇总在 sql/06_factor_analysis.sql，共 31 条查询、7 个部分，全部经过逐条执行验证。'
     '所有表名带 olist. 前缀，因此每条查询都能独立运行。', indent=0.85)
add_table(['部分', '内容', '条数'],
          [['第 0 部分', '基础信息与差评率基线', '3'],
           ['第 1 部分', '流程拆解（定位瓶颈）', '3'],
           ['第 2 部分', '地理因素', '4'],
           ['第 3 部分', '卖家因素（备货 + 违约）', '5'],
           ['第 4 部分', '商品因素（品类/重量/件数）', '4'],
           ['第 5 部分', '时间因素（趋势/旺季/异常）', '5'],
           ['第 6 部分', '汇总对比：各因素影响权重', '3'],
           ['第 7 部分', '数据质量与局限', '4'],
           ['合计', '—', '31']],
          widths=[3.0, 9.4, 4.0], caption='SQL 文件结构', right_cols=[2])

h2('Python 建模关键代码', 'A.2')
para('手写 OLS 完整推断（不依赖 statsmodels），以及 Logit 的优势比与置信区间计算：', indent=0.85)
code_block(
    'def ols(y, X, names):\n'
    '    """最小二乘 + 完整统计推断。X 为设计矩阵（不含常数项）。"""\n'
    '    n = len(y)\n'
    '    Xd = np.column_stack([np.ones(n), X])\n'
    '    beta, *_ = np.linalg.lstsq(Xd, y, rcond=None)\n'
    '    resid = y - Xd @ beta\n'
    '    k = Xd.shape[1]; dof = n - k\n'
    '    sigma2 = resid @ resid / dof\n'
    '    XtX_inv = np.linalg.pinv(Xd.T @ Xd)\n'
    '    se = np.sqrt(np.diag(XtX_inv) * sigma2)\n'
    '    tval = beta / se\n'
    '    pval = 2 * (1 - stats.t.cdf(np.abs(tval), dof))\n'
    '    r2 = 1 - (resid @ resid) / ((y - y.mean()) ** 2).sum()\n'
    '    adj_r2 = 1 - (1 - r2) * (n - 1) / dof\n'
    '    return dict(beta=beta, se=se, t=tval, p=pval, r2=r2, adj_r2=adj_r2)',
    caption='A-2  手写 OLS 推断（系数/标准误/t/p/R²）')

code_block(
    '# Logit：拟合 + Fisher 信息矩阵求标准误 + 优势比\n'
    'lr = LogisticRegression(penalty=None, solver=\'lbfgs\', max_iter=5000)\n'
    'lr.fit(Xstd, y_bad)\n'
    'coef = lr.coef_[0]\n'
    'p_hat = lr.predict_proba(Xstd)[:, 1]\n'
    'W = p_hat * (1 - p_hat)\n'
    'H = Xstd.T @ (Xstd * W[:, None])\n'
    'cov = np.linalg.pinv(H)\n'
    'se_lr = np.sqrt(np.diag(cov))\n'
    'z = coef / se_lr\n'
    'p_lr = 2 * (1 - stats.norm.cdf(np.abs(z)))\n'
    'or_ = np.exp(coef)\n'
    'ci_lo = np.exp(coef - 1.96 * se_lr)\n'
    'ci_hi = np.exp(coef + 1.96 * se_lr)',
    caption='A-3  Logit 优势比与 95% 置信区间')

h2('文件清单', 'A.3')
add_table(['文件 / 目录', '说明'],
          [['sql/06_factor_analysis.sql', '31 条分析查询（7 个部分，全部验证通过）'],
           ['sql/01~04_*.sql', '建表、导入、外键视图、数据探索脚本'],
           ['analysis/stat_analysis.py', '相关性分析脚本（相关系数矩阵、显著性、效应量、5 张图）'],
           ['analysis/rigorous_model.py', '★多元建模脚本（OLS/Logit/随机森林/交互/稳健性）'],
           ['analysis/shoot_results.py', 'SQL 结果截图生成脚本'],
           ['analysis/make_report.py', '本报告的生成脚本'],
           ['analysis/output/*.png', '11 张统计与建模图表'],
           ['analysis/output/*.csv', '13 个结果表（utf-8-sig，Excel 可直接打开）'],
           ['results/*.png', '31 张 SQL 查询结果截图 + 索引页'],
           ['Olist巴西电商_差评归因分析报告.docx', '本报告 Word 版'],
           ['Olist巴西电商_差评归因分析报告.pdf', '本报告 PDF 版']],
          widths=[6.6, 9.6], caption='项目文件清单')

doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_font(p.add_run('— 报告完 —'), FONT_CN, 12, True, '9CA3AF')

doc.save(OUT)
print('已生成: %s' % OUT)
print('大小: %.2f MB' % (os.path.getsize(OUT) / 1024 / 1024))
