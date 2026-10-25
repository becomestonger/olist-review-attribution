"""
生成 GitHub 版分析报告（在 59 页原报告基础上重构：结论前置 + 仪表盘章节 + 工程化内容）。
"""
import os
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image

BASE = r'C:\Users\28223\Desktop\巴西olist数据分析项目'
RES = os.path.join(BASE, 'results')
FIG = os.path.join(BASE, 'analysis', 'output')
PBI = os.path.join(BASE, 'powerbi', 'screenshots', 'pbi')
OUT = os.environ.get('REPORT_OUT') or os.path.join(
    BASE, 'Olist巴西电商_差评归因分析报告_GitHub版.docx')

FONT_CN = '微软雅黑'
FONT_BODY = '宋体'
FONT_CODE = 'Consolas'

SZ_BODY = 11.5
SZ_TABLE = 9.5
SZ_CODE = 8.8
SZ_CAP = 9.5

CONTENT_W = 16.2

doc = Document()
s0 = doc.sections[0]
s0.page_width = Cm(21.0)
s0.page_height = Cm(29.7)
s0.left_margin = Cm(2.4)
s0.right_margin = Cm(2.4)
s0.top_margin = Cm(2.2)
s0.bottom_margin = Cm(2.2)

normal = doc.styles['Normal']
normal.font.name = FONT_BODY
normal.font.size = Pt(SZ_BODY)
normal.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_BODY)
normal.paragraph_format.line_spacing = 1.6
normal.paragraph_format.space_after = Pt(6)
_z = doc.settings.element.find(qn('w:zoom'))
if _z is not None:
    _z.set(qn('w:val'), '100')


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


def set_table_fixed(tbl, widths):
    tot = sum(widths)
    if tot > CONTENT_W:
        widths = [w * CONTENT_W / tot for w in widths]
    tbl.autofit = False
    tblPr = tbl._tbl.tblPr
    for tag in ('w:tblW', 'w:tblLayout'):
        for el in tblPr.findall(qn(tag)):
            tblPr.remove(el)
    tw = OxmlElement('w:tblW')
    tw.set(qn('w:w'), str(int(sum(widths) / 2.54 * 1440)))
    tw.set(qn('w:type'), 'dxa')
    tblPr.append(tw)
    lay = OxmlElement('w:tblLayout')
    lay.set(qn('w:type'), 'fixed')
    tblPr.append(lay)
    grid = tbl._tbl.find(qn('w:tblGrid'))
    if grid is not None:
        for gc, w in zip(grid.findall(qn('w:gridCol')), widths):
            gc.set(qn('w:w'), str(int(w / 2.54 * 1440)))
    for row in tbl.rows:
        for cell, w in zip(row.cells, widths):
            cell.width = Cm(w)
    return widths


def shade(obj, color):
    if hasattr(obj, '_tc'):
        pr = obj._tc.get_or_add_tcPr()
    elif hasattr(obj, '_p'):
        pr = obj._p.get_or_add_pPr()
    else:
        pr = obj.get_or_add_pPr() if hasattr(obj, 'get_or_add_pPr') else obj.get_or_add_tcPr()
    sh = OxmlElement('w:shd')
    sh.set(qn('w:val'), 'clear')
    sh.set(qn('w:color'), 'auto')
    sh.set(qn('w:fill'), color)
    pr.append(sh)


def para(text='', size=SZ_BODY, bold=False, color=None, align=None,
         before=0, after=6, indent=None, italic=False):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
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
        _font(p.add_run(num + '  '), FONT_CN, 18, True, '1E3A5F')
    _font(p.add_run(text), FONT_CN, 18, True, '1E3A5F')
    pPr = p._element.get_or_add_pPr()
    bd = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single')
    b.set(qn('w:sz'), '10')
    b.set(qn('w:color'), '1E3A5F')
    bd.append(b)
    pPr.append(bd)


def h2(text, num=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(15)
    p.paragraph_format.space_after = Pt(7)
    if num:
        _font(p.add_run(num + '  '), FONT_CN, 14.5, True, '2E5C8A')
    _font(p.add_run(text), FONT_CN, 14.5, True, '2E5C8A')


def h3(text, num=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(11)
    p.paragraph_format.space_after = Pt(4)
    if num:
        _font(p.add_run(num + '  '), FONT_CN, 12.5, True, '374151')
    _font(p.add_run(text), FONT_CN, 12.5, True, '374151')


def bullet(text, size=SZ_BODY, indent=0.9):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.5
    _font(p.add_run(text), FONT_BODY, size)


def bullet_b(text, size=SZ_BODY):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(0.9)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.5
    _font(p.add_run(text), FONT_BODY, size, True)


def code_block(code, caption=None, size=SZ_CODE):
    if caption:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        _font(p.add_run('▍' + caption), FONT_CN, 9.5, True, '6B7280')
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
        e.set(qn('w:color'), 'E0E3E8')
        bd.append(e)
    tcPr.append(bd)
    cell.text = ''
    first = True
    for line in code.split('\n'):
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.12
        st = line.strip()
        is_c = st.startswith('--') or st.startswith('#')
        _font(p.add_run(line if line else ' '), FONT_CODE, size,
              color='1B7A3E' if is_c else '1F2937')
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def add_figure(path, caption, max_w=None, max_h=10.5, size=SZ_CAP):
    if not os.path.exists(path):
        para('[图片缺失: %s]' % os.path.basename(path), color='C0392B')
        return
    im = Image.open(path)
    w = min(max_w or CONTENT_W, CONTENT_W)
    h = w * im.height / im.width
    if h > max_h:
        h = max_h
        w = h * im.width / im.height
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.add_run().add_picture(path, width=Cm(w))
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_after = Pt(11)
    _font(cp.add_run(caption), FONT_CN, size, color='6B7280')


def add_table(headers, rows, widths=None, caption=None, fs=SZ_TABLE,
              hbg='1E3A5F', right_cols=None):
    if caption:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        _font(p.add_run('表  ' + caption), FONT_CN, 9.5, True, '6B7280')
    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    fixed = set_table_fixed(tbl, widths if widths else [CONTENT_W / len(headers)] * len(headers))
    for i, ht in enumerate(headers):
        c = tbl.rows[0].cells[i]
        c.text = ''
        shade(c, hbg)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        _font(p.add_run(str(ht)), FONT_CN, fs, True, 'FFFFFF')
    for ri, row in enumerate(rows):
        cells = tbl.add_row().cells
        for ci, val in enumerate(row):
            c = cells[ci]
            c.text = ''
            c.width = Cm(fixed[ci])
            if ri % 2 == 1:
                shade(c, 'F7F9FB')
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.space_after = Pt(1.5)
            p.paragraph_format.line_spacing = 1.15
            if right_cols and ci in right_cols:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci == 0 else WD_ALIGN_PARAGRAPH.CENTER
            _font(p.add_run(str(val)), FONT_BODY, fs)
    for ri in range(len(tbl.rows)):
        for ci, w in enumerate(fixed):
            tbl.rows[ri].cells[ci].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def callout(title, text, bg='FFF8E1', border='E8A33D', tcolor='8A5A00', size=10.5):
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
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    _font(p.add_run(title), FONT_CN, size + 0.5, True, tcolor)
    for line in text.split('\n'):
        pp = cell.add_paragraph()
        pp.paragraph_format.space_before = Pt(0)
        pp.paragraph_format.space_after = Pt(1)
        pp.paragraph_format.line_spacing = 1.42
        _font(pp.add_run(line), FONT_BODY, size)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def landscape_section():
    """切换到横向页面（用于放宽幅仪表盘截图）。"""
    global CONTENT_W
    guard = doc.add_section(WD_SECTION.NEW_PAGE)
    guard.orientation = WD_ORIENT.LANDSCAPE
    guard.page_width = Cm(29.7)
    guard.page_height = Cm(21.0)
    guard.left_margin = Cm(1.9)
    guard.right_margin = Cm(1.9)
    guard.top_margin = Cm(1.5)
    guard.bottom_margin = Cm(1.5)
    CONTENT_W = 25.9


def portrait_section():
    """切回纵向页面。"""
    global CONTENT_W
    guard = doc.add_section(WD_SECTION.NEW_PAGE)
    guard.orientation = WD_ORIENT.PORTRAIT
    guard.page_width = Cm(21.0)
    guard.page_height = Cm(29.7)
    guard.left_margin = Cm(2.4)
    guard.right_margin = Cm(2.4)
    guard.top_margin = Cm(2.2)
    guard.bottom_margin = Cm(2.2)
    CONTENT_W = 16.2


# ===========================================================================
# 封面
# ===========================================================================
for _ in range(4):
    doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_font(p.add_run('Olist 巴西电商'), FONT_CN, 27, True, '1E3A5F')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
_font(p.add_run('客户差评归因分析'), FONT_CN, 33, True, '2E5C8A')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(30)
_font(p.add_run('从流程拆解到多元统计建模的完整分析链路'), FONT_CN, 13, False, '6B7280')

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
_font(p.add_run('155 万行真实交易数据 · 99,441 笔订单 · 25 个月'), FONT_CN, 11.5, color='4B5563')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(34)
_font(p.add_run('PostgreSQL · Python · Power BI · 统计建模'), FONT_CN, 11.5, color='4B5563')

core = [
    ('核心结论', ''),
    ('①', '「是否迟到」是差评最强预测因子\n   迟到 53.94%  vs  准时 9.18%（5.9 倍）'),
    ('②', '关系是阈值型，不是线性\n   已超期 53.98%，未超期仅 8.79%–11.16%'),
    ('③', '运输占全程 74.2%，是瓶颈但不可控\n   备货仅占 22.3%，却完全可控'),
    ('④', '超期发货的效应被「迟到」完全中介\n   控制迟到后 β = −0.0002（p = 0.97）'),
]
tbl = doc.add_table(rows=0, cols=2)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_fixed(tbl, [1.6, 14.0])
for k, v in core:
    cells = tbl.add_row().cells
    cells[0].width = Cm(1.6)
    cells[1].width = Cm(14.0)
    if k == '核心结论':
        _font(cells[0].paragraphs[0].add_run(''), FONT_CN, 11)
        _font(cells[1].paragraphs[0].add_run('核 心 结 论'), FONT_CN, 12.5, True, '1E3A5F')
        continue
    p0 = cells[0].paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _font(p0.add_run(k), FONT_CN, 12, True, 'C0392B')
    for i, line in enumerate(v.split('\n')):
        pv = cells[1].paragraphs[0] if i == 0 else cells[1].add_paragraph()
        pv.paragraph_format.line_spacing = 1.4
        pv.paragraph_format.space_after = Pt(2)
        _font(pv.add_run(line), FONT_BODY, 10.5, color='1F2937')
set_table_fixed(tbl, [1.6, 14.0])

for _ in range(4):
    doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_font(p.add_run(datetime.now().strftime('%Y 年 %m 月')), FONT_CN, 10.5, color='9CA3AF')

page_break()

# ===========================================================================
# 摘要
# ===========================================================================
h1('摘要')
para('本项目基于巴西电商平台 Olist 的公开交易数据（155 万行、99,441 笔订单、25 个月），'
     '回答一个具体的业务问题：**什么因素导致客户差评？各因素权重如何？业务该优先改哪一个？**',
     indent=0.8)
para('分析按「沿业务流程拆解 → 单因素归因 → 相关分析 → 多元回归验证」四层递进。'
     '先用四段耗时拆解定位瓶颈，再逐个排查地理、卖家、商品、时间四类因素，'
     '最后用 OLS、Logistic 回归和随机森林三种方法交叉验证，确保结论不依赖单一模型假设。',
     indent=0.8)

h2('五个核心发现')
add_table(['#', '发现', '关键证据'],
          [['1', '「是否迟到」是差评最强预测因子',
            '迟到 53.94% vs 准时 9.18%（5.9 倍）\n标准化 β = −0.277，Cohen\'s d = −1.41，OR = 1.457'],
           ['2', '关系是阈值型（悬崖），不是线性（斜坡）',
            '已超期 53.98%；未超期各档仅 8.79%–11.16%\n分段相关：全样本 r=0.256 → 未超期子集 r=0.024'],
           ['3', '运输占全程 74.2%，是瓶颈但不可控',
            '审核 0.43 + 备货 2.81 + 运输 9.33 = 12.56 天\n运输由距离决定，平台无法改变'],
           ['4', '超期发货的效应被「迟到」完全中介',
            '单因素 r=−0.106 → 控制迟到后 β=−0.0002（p=0.97）\n因果路径：超期发货 → 迟到 → 差评'],
           ['5', '距离的作用被严重高估',
            '相关系数 r=−0.058 → 控制运输时长后 β=+0.064（符号反转）\n跨州 Cohen\'s d 仅 −0.12，可忽略']],
          widths=[1.0, 5.6, 9.6], caption='五个核心发现与对应证据')

h2('一个可执行的结论')
callout('业务目标应该是「绝不违约」，而不是「盲目提速」',
        '如果关系是线性斜坡，业务应追求"越送越快"；但数据显示是悬崖——\n'
        '未超期时多等几天几乎不影响评分（8.79%→11.16%），一旦超期就断崖式恶化（→53.98%）。\n\n'
        '所以正确的动作是：\n'
        '· 停止为"数据好看"而收紧预估窗口（2018-08 曾因此让迟到率从 4.5% 涨到 10.4%）\n'
        '· 建立卖家发货时效考核（9.06% 的超期发货是可控的上游原因）\n'
        '· 大件品类设独立承诺标准（office_furniture 备货 9.92 天，是均值的 3.5 倍）',
        bg='E8F5E9', border='43A047', tcolor='1B5E20')

h2('项目规模')
add_table(['维度', '内容'],
          [['数据', '9 张表 / 1,550,922 行 / 99,441 笔订单 / 96,096 名客户'],
           ['时间跨度', '2016-09 至 2018-10（25 个月）'],
           ['分析查询', '31 条 SQL（7 个部分，逐条验证通过）'],
           ['统计方法', '相关分析（Pearson/Spearman）、效应量、卡方检验、OLS 回归、Logistic 回归、随机森林'],
           ['建模样本', '93,047 条记录、10 维特征，交叉验证 AUC = 0.744'],
           ['交付物', '分析报告 + 6 页 Power BI 仪表盘 + 31 张 SQL 结果图 + 可复现代码']],
          widths=[2.8, 13.4], caption='项目规模概览')

page_break()

# ===========================================================================
# 第一章 数据理解与工程
# ===========================================================================
h1('数据理解与工程实现', '第一章')

h2('数据集结构', '1.1')
para('9 张表围绕「订单」这一核心事实表展开。理解表间关系是后续所有分析的前提。', indent=0.8)
add_table(['表名', '行数', '业务含义', '在分析中的作用'],
          [['orders', '99,441', '订单主表，含 5 个关键时间戳', '★ 计算各环节时长'],
           ['order_items', '112,650', '订单明细：商品、单价、运费、发货期限', '★ 超期发货判定'],
           ['order_reviews', '99,224', '评价：1-5 星评分 + 文字评论', '★ 被解释变量'],
           ['customers', '99,441', '客户城市、州、邮编', '地理因素'],
           ['sellers', '3,095', '卖家城市、州、邮编', '卖家因素'],
           ['products', '32,951', '商品品类、重量、尺寸', '商品因素'],
           ['order_payments', '103,886', '支付方式、分期数、金额', '辅助（证明影响很小）'],
           ['geolocation', '1,000,163', '邮编对应经纬度', '计算真实距离'],
           ['category_translation', '71', '品类葡语→英语', '品类可读性']],
          widths=[4.0, 1.9, 6.2, 4.1], caption='9 张表的分工')
code_block(
    'customers ──1:N──> orders ──1:N──> order_items <──N:1── products\n'
    '                      │                 └────N:1────────> sellers\n'
    '                      ├──1:N──> order_payments\n'
    '                      └──1:N──> order_reviews\n'
    '\n'
    'products.product_category_name ──N:1──> product_category_translation',
    caption='表间关系（外键约束已验证，孤儿记录为 0）')

h2('两个必须注意的数据陷阱', '1.2')
add_figure(os.path.join(RES, '28_7.1关键字段缺失率.png'),
           '图 1-1  关键字段缺失率：评论标题缺失 88.3%、评论内容缺失 58.7%', max_h=7.5)
add_figure(os.path.join(RES, '31_7.4顾客标识的唯一性（说明为何必须用customer_unique_id）.png'),
           '图 1-2  顾客标识唯一性：99,441 个 customer_id 各不相同，真实客户仅 96,096 人', max_h=7.5)
callout('陷阱一：customer_id ≠ customer_unique_id',
        'customers 表有 99,441 行，唯一 customer_id 也是 99,441 个——每笔订单都生成新的 customer_id。\n'
        '真正代表"同一个人"的是 customer_unique_id（96,096 个）。\n\n'
        '如果误用 customer_id 计算复购率，会得到 0% 的错误结论。真实复购率约 3.1%。')
callout('陷阱二：平均值会骗人',
        '运输时长平均值 9.33 天，但中位数只有 7.10 天，p90 达 18.90 天，最大 205 天。\n'
        '均值 > 中位数说明分布右偏——平均值被少数极端慢单拉高了。\n\n'
        '因此本报告在描述时长时同时给出均值与中位数，并在相关分析中同时计算 Pearson 与 Spearman。')

h2('数据导入与视图设计', '1.3')
para('使用 psql 的 \\copy 命令（客户端读取）导入 9 个 CSV，因为服务端 COPY 需要 '
     'pg_read_server_files 超级用户权限。逐表核对行数与源文件一致。', indent=0.8)
code_block(
    '\\copy customers (customer_id, customer_unique_id, customer_zip_code_prefix,\n'
    '                 customer_city, customer_state)\n'
    "FROM 'C:/Users/.../olist_customers_dataset.csv'\n"
    "WITH (FORMAT csv, HEADER true, ENCODING 'UTF8', NULL '');",
    caption='导入命令示例（\\copy 是 psql 元命令，需在脚本文件中执行）')
para('为减少重复 JOIN，建立了 3 个分析视图。其中 v_orders_full 是订单宽表，'
     '把订单、客户、金额合计、支付合计、评分合并成一行一单。', indent=0.8)

h2('为 BI 构建宽表', '1.4')
para('直接给 Power BI 提供 9 张表会让建模变得复杂（需手工建 6 条关系、且一对多 JOIN 容易重复计数）。'
     '因此专门建了一个订单级宽表 v_pbi_orders，把 7 张表的信息压成一行一单。', indent=0.8)
code_block(
    'CREATE VIEW olist.v_pbi_orders AS\n'
    'WITH item_agg AS (               -- ① 明细聚合成一单一行，避免重复计数\n'
    '    SELECT order_id, sum(price) AS items_total,\n'
    '           count(*) AS item_count, max(shipping_limit_date) AS ship_limit,\n'
    '           count(DISTINCT seller_id) AS seller_count,\n'
    '           min(seller_id) AS main_seller_id, min(product_id) AS main_product_id\n'
    '    FROM olist.order_items GROUP BY order_id\n'
    '), ...\n'
    'SELECT o.order_id, o.order_status,\n'
    '       -- 过程指标全部预计算，BI 端零 DAX 即可拖拽\n'
    '       CASE WHEN o.order_delivered_customer_date > o.order_estimated_delivery_date\n'
    '            THEN 1 ELSE 0 END AS is_late,\n'
    '       round((EXTRACT(EPOCH FROM (o.order_delivered_carrier_date\n'
    '               - o.order_approved_at))/86400)::numeric, 2) AS prep_days, ...\n'
    'FROM olist.orders o\n'
    'LEFT JOIN item_agg a ON a.order_id = o.order_id   -- ★ LEFT JOIN 保留全部订单\n'
    'JOIN olist.customers c ON c.customer_id = o.customer_id ...;',
    caption='v_pbi_orders 视图（订单级宽表，30 个字段）')
callout('一个必须用 LEFT JOIN 的细节',
        '第一版用 INNER JOIN 连商品明细，结果只有 98,666 单——775 笔订单没有商品明细\n'
        '（603 unavailable、164 canceled 等）被丢掉了，导致差评率从 14.63% 变成 14.13%。\n\n'
        '改用 LEFT JOIN 后全量 99,441 单都在，差评率与报告完全一致。\n'
        '同时把各标记位在"无法判断"时设为 NULL 而非 0，这样 BI 里的 DIVIDE 会自动正确排除。')

page_break()

# ===========================================================================
# 第二章 核心结论
# ===========================================================================
h1('五个核心结论', '第二章')
para('本章是报告主体。每个结论都给出「现象 → 证据 → 业务含义」三段。', indent=0.8)

h2('结论一：「是否迟到」是差评最强预测因子', '2.1')
add_table(['指标', '准时', '迟到', '差距'],
          [['差评率', '9.18%', '53.94%', '5.9 倍'],
           ['平均评分', '4.30', '2.61', '−1.69 分'],
           ['样本量', '88,649', '7,827', '—']],
          widths=[4.0, 4.0, 4.0, 4.2], caption='迟到 vs 准时的直接对比')
para('三种完全不同原理的方法都把「是否迟到」排在第一或第二：', indent=0.8)
add_table(['方法', '该因素的排名与数值'],
          [['OLS 标准化系数', '排名 1：β = −0.277（t = −56.0，p < 1e-300）'],
           ['Logistic 优势比', '排名 2：OR = 1.457（95% CI 1.428–1.486）'],
           ['随机森林重要性', '排名 1：39.60%（第二名的运输为 30.33%）'],
           ['效应量', "Cohen's d = −1.41（极大效应，>0.8 即为大）"],
           ['卡方检验', "χ² = 11,482.3，Cramér's V = 0.3504（中等关联）"]],
          widths=[4.4, 11.8], caption='五种方法对「是否迟到」的估计')
callout('业务含义',
        '客户不是不能等，而是不能接受承诺落空。\n'
        '满意度管理的重点是"守住承诺"，而不是"盲目提速"。')

h2('结论二：关系是阈值型（悬崖），不是线性（斜坡）', '2.2')
add_figure(os.path.join(FIG, 'fig4_悬崖效应.png'),
           '图 2-1  悬崖效应：已超期组差评率 52.8%，未超期各档仅在 8.8%–11.1% 之间持平', max_h=9.0)
add_table(['预估富余档', '订单数', '差评率', '平均评分'],
          [['已超期（负富余）', '7,821', '52.77%', '2.61'],
           ['富余 0–3 天', '4,744', '11.16%', '4.14'],
           ['富余 3–7 天', '12,586', '9.83%', '4.23'],
           ['富余 7–15 天', '41,719', '8.79%', '4.32'],
           ['富余 15 天以上', '29,606', '9.14%', '4.26']],
          widths=[5.0, 3.6, 3.6, 4.0], caption='预估富余分档与差评率', right_cols=[1, 2, 3])
para('这个结论可以用一个反证来验证：把「已超期」的样本剔除后重新计算相关系数，'
     '整体相关性几乎消失。', indent=0.8)
add_table(['子集', 'Pearson r', 'p 值', '解读'],
          [['全部样本（n = 94,138）', '0.2559', '≈0', '看起来中等相关'],
           ['仅未超期子集（n = 87,013）', '0.0241', '1.38e−12', '★ 关系几乎消失']],
          widths=[6.2, 3.2, 3.0, 3.8], caption='分段相关系数：整体 vs 未超期子集', right_cols=[1, 2])
add_figure(os.path.join(FIG, 'fig3_关键散点与趋势.png'),
           '图 2-2  散点与分箱均值曲线：趋势在 20 天左右开始加速下滑，呈明显非线性', max_h=8.5)
callout('这是全项目最重要的统计发现',
        '整体相关系数 r = 0.2559 完全由"已超期"这 7.6% 的样本驱动。\n'
        '在未超期范围内，富余天数与评分基本无关（r = 0.0241）。\n\n'
        '真实关系是「阈值型（悬崖）」，不是「线性型（斜坡）」。\n'
        '业务含义完全不同：若是斜坡 → 应追求"越送越快"；若是悬崖 → 应追求"绝不违约"。',
        bg='FDECEA', border='C0392B', tcolor='922B21')

h2('结论三：运输是瓶颈但不可控，备货可控', '2.3')
add_figure(os.path.join(RES, '04_1.1四段耗时拆解★这一步决定后面查哪里.png'),
           '图 2-3  四段耗时拆解：审核 0.43 天、备货 2.81 天、运输 9.33 天、全程 12.56 天', max_h=7.5)
add_table(['环节', '平均耗时', '占全程', '可控性', '判断'],
          [['下单 → 审核', '0.43 天', '3.5%', '高', '可忽略'],
           ['审核 → 发货（备货）', '2.81 天', '22.3%', '★ 高', '中等，完全可控'],
           ['发货 → 签收（运输）', '9.33 天', '74.2%', '低（由距离决定）', '★ 主要瓶颈']],
          widths=[4.4, 2.8, 2.4, 3.4, 3.2], caption='四段耗时、占比与可控性', right_cols=[1, 2])
callout('最大的问题 ≠ 最该改的问题',
        '运输占 74.2%，是绝对的流程瓶颈；但它主要由客户与卖家的距离决定，平台改变不了。\n'
        '备货只占 22.3%，却完全在平台和卖家的控制范围内。\n\n'
        '所以改进优先级应该是「可控性 × 影响力」，而不是只看影响力。\n'
        '特别是 9.06% 的超期发货订单——责任清晰、可考核、可改进。')

h2('结论四：距离的作用被严重高估', '2.4')
add_table(['分析方式', '距离与评分的关系', '解读'],
          [['Pearson 相关', 'r = −0.0580', '弱负相关，看似"越远越不满"'],
           ['Cohen\'s d（跨州 vs 同州）', '−0.12', '效应量可忽略'],
           ['★ 多元回归（控制运输时长）', 'β = +0.0636（p < 1e-300）', '符号反转为正'],
           ['★ Logit（控制运输时长）', 'OR = 0.848（p < 1e-300）', '距离越远，差评几率反而更低']],
          widths=[5.0, 5.4, 5.8], caption='距离效应在单因素与多因素下的对比')
callout('为什么会符号反转？（典型抑制效应）',
        '距离与运输时长高度相关（r = 0.44）。距离远 → 运输慢 → 评分低，这是距离的"间接负效应"。\n'
        '但在多元模型中，控制住"运输时长"后，问的就变成「同等运输时长下，远近还有什么差别」——\n'
        '结果发现远距离订单评分反而略高。\n\n'
        '可能原因：偏远地区客户预期更低、平台预估窗口更宽松。\n'
        '方法论意义：不能用单因素相关系数断言"距离导致不满"。',
        bg='EDF4FB', border='4A90D9', tcolor='1E3A5F')

h2('结论五：品类差评分两类，处置方式完全不同', '2.5')
para('把「差评率」和「平均备货天数」并排看，高差评品类会分成两种截然不同的类型：', indent=0.8)
add_table(['品类', '订单数', '差评率', '平均备货', '类型', '处置方向'],
          [['office_furniture', '1,267', '22.51%', '9.92 天', '物流型', '大件品类设独立承诺标准'],
           ['audio', '347', '21.51%', '2.51 天', '★ 商品型', '查商品描述、质检、包装'],
           ['fixed_telephony', '215', '18.87%', '3.41 天', '商品型', '同上'],
           ['home_confort', '388', '18.65%', '2.3 天', '商品型', '同上'],
           ['bed_bath_table', '9,297', '16.32%', '2.71 天', '混合', '需进一步拆分']],
          widths=[4.0, 1.8, 2.2, 2.2, 2.2, 3.8], caption='高差评品类的两种类型',
          right_cols=[1, 2, 3])
callout('这条洞察比「差评率排行」有价值得多',
        'office_furniture 高差评可以用备货慢解释（9.92 天，全国均值 2.81 的 3.5 倍）——改物流有效。\n'
        '但 audio 差评率 21.51%，备货只有 2.51 天（比均值还快），说明差评来自商品本身——'
        '描述不符、质量问题、到手损坏，靠改物流完全无效。\n\n'
        '同样是"高差评品类"，根因不同，处置方式完全不同。')

h2('五因素汇总对比', '2.6')
add_table(['排名', '因素', '高分组差评率', '低分组差评率', '差距'],
          [['1', '是否迟到', '53.94%（迟到）', '9.18%（准时）', '+44.8 pp'],
           ['2', '运输时长', '48.01%（21天+）', '8.46%（3天内）', '+39.6 pp'],
           ['3', '备货时长', '41.56%（14天+）', '9.84%（1天内）', '+31.7 pp'],
           ['4', '超期发货', '23.69%（超期）', '11.68%（按期）', '+12.0 pp'],
           ['5', '是否跨州', '14.07%（跨州）', '10.43%（同州）', '+3.6 pp']],
          widths=[1.4, 4.2, 4.4, 4.4, 2.4], caption='五因素对差评率的影响排序',
          right_cols=[2, 3, 4])

page_break()

# ===========================================================================
# 第三章 统计与建模验证
# ===========================================================================
h1('统计与建模验证', '第三章')
para('第二章的结论全部经过统计检验和多元建模验证。本章说明验证过程与两处颠覆性发现。', indent=0.8)

h2('相关分析：Pearson 与 Spearman', '3.1')
para('同时计算两种相关系数：Pearson 衡量线性关系，Spearman 衡量单调关系。'
     '本数据多个变量偏度绝对值大于 1（审核时长 4.48、运费 12.90），分布明显右偏，'
     '此时 Pearson 会失效，必须对照 Spearman。', indent=0.8)
add_figure(os.path.join(FIG, 'fig1_相关系数热力图.png'),
           '图 3-1  相关系数热力图（左：Pearson 线性；右：Spearman 单调）', max_h=9.5)
add_table(['因素', 'Pearson r', 'Spearman ρ', '强度判断'],
          [['是否迟到', '−0.3514', '−0.2823', '强负相关'],
           ['全程时长', '−0.3351', '−0.2266', '强负相关'],
           ['运输时长', '−0.3015', '−0.1877', '中等负相关'],
           ['预估富余', '+0.2559', '+0.1684', '中等正相关'],
           ['备货时长', '−0.1515', '−0.1004', '弱负相关'],
           ['卖家超期发货', '−0.1058', '−0.0924', '弱负相关'],
           ['是否跨州', '−0.0580', '−0.0557', '极弱']],
          widths=[4.4, 3.2, 3.2, 3.8], caption='各因素与「评分」的相关系数（n = 93,526）',
          right_cols=[1, 2])
add_figure(os.path.join(FIG, 'fig2_各因素相关强度.png'),
           '图 3-2  各因素与「评分」/「是否差评」的相关强度对比', max_h=8.5)

h2('效应量：p 值之外的另一半', '3.2')
para('p 值只回答"有没有差异"，效应量回答"差异有多大"。本项目样本达 9.3 万，'
     '几乎所有 p 值都趋近于 0，因此必须同时报告效应量。', indent=0.8)
add_table(['分组变量', 'Cohen\'s d', '判读', '说明'],
          [['是否迟到', '−1.41', '★ 极大效应', '迟到组评分分布整体左移 1.41 个标准差'],
           ['卖家超期发货', '−0.37', '小-中效应', '单因素看有影响，但多因素下消失'],
           ['是否跨州', '−0.12', '可忽略', '远低于 0.2 的"小效应"门槛']],
          widths=[3.6, 2.6, 3.2, 6.8], caption='组间差异的效应量', right_cols=[1])
add_figure(os.path.join(FIG, 'fig5_组间对比箱线图.png'),
           '图 3-3  迟到 / 超期发货两组的评分箱线图（标题标注 Cohen\'s d）', max_h=7.5)
callout('大样本陷阱（容易被追问）',
        '样本量越大，p 值越容易变小。本项目几乎所有 p 值都 < 1e-300，\n'
        '连相关系数只有 0.02 的变量都"显著"。\n\n'
        '所以不能只看 p 值，必须同时看效应大小。\n'
        '例如"是否跨州"的 p 值也 < 0.001，但 d 只有 −0.12，实际可忽略。')

h2('多元线性回归（OLS）', '3.3')
add_table(['整体指标', '数值', '含义'],
          [['样本量 N', '93,047', '—'],
           ['R²', '0.1695', '模型解释了 16.95% 的评分方差'],
           ['调整 R²', '0.1694', '几乎无衰减，无过拟合'],
           ['F 统计量', '1,898.8', '整体模型高度显著'],
           ['VIF', '全部 < 3', '无多重共线性问题']],
          widths=[3.6, 3.2, 9.4], caption='OLS 模型整体拟合情况', right_cols=[1])
add_table(['排名', '变量', '标准化 β', 't 值', 'p 值', '显著性'],
          [['1', '是否迟到', '−0.2771', '−56.0', '<1e-300', '***'],
           ['2', '运输时长(天)', '−0.2749', '−49.1', '<1e-300', '***'],
           ['3', '商品件数', '−0.1642', '−40.4', '<1e-300', '***'],
           ['4', '备货时长(天)', '−0.1310', '−27.2', '<1e-300', '***'],
           ['5', '距离(对数km)', '+0.0636', '10.5', '<1e-300', '***'],
           ['6', '是否跨州', '+0.0151', '2.6', '0.0088', '**'],
           ['7', '重量(对数g)', '+0.0123', '2.7', '0.0063', '**'],
           ['8', '金额(对数)', '−0.0063', '−1.0', '0.3042', 'ns'],
           ['9', '运费占比', '−0.0041', '−0.7', '0.4564', 'ns'],
           ['10', '卖家超期发货', '−0.0002', '−0.0', '0.9730', 'ns']],
          widths=[1.2, 4.0, 2.6, 2.0, 2.6, 2.2],
          caption='OLS 标准化系数（因变量：评分）', right_cols=[2, 3, 4, 5])

h2('逻辑回归：优势比', '3.4')
add_table(['整体指标', '数值'],
          [['McFadden 伪 R²', '0.1518'],
           ['LR 统计量', '10,629.7（df = 10，p < 1e-300）'],
           ['基准差评率', '12.47%']],
          widths=[6.4, 9.8], caption='Logit 模型整体情况', right_cols=[1])
add_table(['变量', 'OR', '95% 置信区间', 'p 值'],
          [['运输时长(天)', '1.603', '[1.562, 1.646]', '<1e-300'],
           ['是否迟到', '1.457', '[1.428, 1.486]', '<1e-300'],
           ['商品件数', '1.351', '[1.327, 1.375]', '<1e-300'],
           ['备货时长(天)', '1.252', '[1.225, 1.280]', '<1e-300'],
           ['距离(对数km)', '0.848', '[0.820, 0.878]', '<1e-300'],
           ['金额(对数)', '1.089', '[1.054, 1.125]', '2.81e−07'],
           ['重量(对数g)', '0.947', '[0.924, 0.971]', '1.67e−05'],
           ['运费占比', '1.034', '[1.008, 1.062]', '0.0108'],
           ['卖家超期发货', '1.013', '[0.990, 1.037]', '0.2756'],
           ['是否跨州', '0.996', '[0.964, 1.030]', '0.8237']],
          widths=[4.2, 2.2, 4.6, 2.6],
          caption='Logit 优势比（每增加 1 个标准差，差评几率的变化倍数）',
          right_cols=[1, 2, 3])

h2('随机森林对照', '3.5')
add_table(['变量', '重要性', '占比'],
          [['是否迟到', '0.3960', '39.60%'],
           ['运输时长(天)', '0.3033', '30.33%'],
           ['商品件数', '0.1003', '10.03%'],
           ['备货时长(天)', '0.0710', '7.10%'],
           ['距离(对数km)', '0.0408', '4.08%'],
           ['其余 5 个变量合计', '0.0986', '9.86%']],
          widths=[6.0, 4.0, 4.0], caption='随机森林特征重要性（5 折交叉验证 AUC = 0.744）',
          right_cols=[1, 2])
add_figure(os.path.join(FIG, 'fig6_多元模型系数对比.png'),
           '图 3-4  三种方法的系数/重要性对比：结论高度一致', max_h=8.0)

h2('两处颠覆单因素结论的发现', '3.6')
h3('发现一：超期发货的效应被「迟到」完全中介', '3.6.1')
add_table(['分析方式', '结果', '结论'],
          [['单因素相关', 'r = −0.1058（p ≈ 0）', '看似显著'],
           ['单因素分组', '评分 3.70 vs 4.20', '看起来影响很大'],
           ['★ 多元回归（控制迟到）', 'β = −0.0002，p = 0.973', '直接效应几乎为零'],
           ['★ Logit（控制迟到）', 'OR = 1.013，p = 0.276', '优势比不显著']],
          widths=[4.4, 5.6, 6.2], caption='超期发货在单因素与多因素下的对比')
callout('因果路径：超期发货 → 迟到 → 差评',
        '客户并不知道卖家有没有超期（那是平台内部指标），客户只感受得到"有没有按时收到"。\n'
        '所以一旦控制住"是否迟到"，超期发货就失去了独立解释力。\n\n'
        '业务含义：考核卖家发货时效仍有意义（它是上游原因），\n'
        '但真正的 KPI 应该是"准时送达率"，而不是"按时发货率"。',
        bg='FDECEA', border='C0392B', tcolor='922B21')

h3('发现二：距离的符号反转（抑制效应）', '3.6.2')
para('见第二章 2.4 节。这是典型的抑制效应，也是"只看相关系数会出错"的最好例证。', indent=0.8)

h2('稳健性检验', '3.7')
add_table(['子样本', 'N', '迟到 β', '距离 β', '超期 β', 'R²'],
          [['2017 年', '42,331', '−0.2528', '+0.0612', '−0.0051', '0.1408'],
           ['2018 年', '50,456', '−0.2937', '+0.0626', '+0.0041', '0.1931'],
           ['全样本', '93,047', '−0.2771', '+0.0636', '−0.0002', '0.1695']],
          widths=[3.0, 2.8, 2.8, 2.6, 2.6, 2.4],
          caption='分年度稳健性检验（标准化系数）', right_cols=[1, 2, 3, 4, 5])
para('三个关键结论在两年间完全一致，说明不是特定年份的偶然现象。', indent=0.8)
add_figure(os.path.join(FIG, 'fig8_模型解释力对比.png'),
           '图 3-5  逐步加入变量后的解释力：仅时间变量 R²=0.11 → 加入违约 0.16 → 全模型 0.17',
           max_h=7.0)

landscape_section()

# ===========================================================================
# 第四章 交互式仪表盘
# ===========================================================================
h1('交互式仪表盘（Power BI）', '第四章')
para('为让结论可以被业务方自助探索，用 Power BI 构建了 6 页交互式仪表盘，'
     '数据源是从 PostgreSQL 直接连接的订单级宽表 v_pbi_orders（99,441 行）。', indent=0.8)
add_table(['页面', '内容', '核心图表'],
          [['P1 经营总览', '5 项核心 KPI + 趋势 + 分布', '多行卡、折线图、环形图、柱形图'],
           ['P2 差评归因', '★ 六因素对差评率的对比', '6 个簇状柱形图 + 结论文本框'],
           ['P3 流程拆解', '三段耗时占比 + 月度迟到率', '环形图、明细表、折线图'],
           ['P4 地域分析', '各州订单量/迟到率/差评率', '3 个条形图 + 气泡图'],
           ['P5 品类分析', '品类销售额/差评率/备货天数', '3 个条形图 + 矩阵表'],
           ['P6 数据说明', '口径定义与局限', '文本框']],
          widths=[4.6, 10.0, 11.3], caption='仪表盘页面结构')

page_break()
h2('P1 经营总览', '4.1')
add_figure(os.path.join(PBI, 'page-0001.png'),
           '仪表盘 P1  经营总览：99,441 单 / R$ 15.84M / 客单价 R$ 159.33 / 评分 4.09 / 差评率 14.63%',
           max_h=13.3)

page_break()
h2('P2 差评归因（核心页）', '4.2')
para('这一页是整份报告核心结论的可视化：六个因素分别作为一个柱形图，'
     'Y 轴统一为差评率、统一量程 0–60%，方便横向对比。', indent=0.8)
add_figure(os.path.join(PBI, 'page-0002.png'),
           '仪表盘 P2  差评归因：迟到 53.94% vs 准时 9.18%；超期 23.69% vs 按期 11.68%；'
           '跨州 14.07% vs 同州 10.43%', max_h=12.5)

page_break()
h2('P3 流程拆解', '4.3')
add_figure(os.path.join(PBI, 'page-0003.png'),
           '仪表盘 P3  流程拆解：运输占 74.23%，是绝对瓶颈；月度迟到率在 2018-03 达到 21.4%',
           max_h=13.3)

page_break()
h2('P4 地域分析', '4.4')
add_figure(os.path.join(PBI, 'page-0004.png'),
           '仪表盘 P4  地域分析：AL 迟到率 23.93% 最高，RO 仅 2.9%；'
           '气泡图展示"运输天数 × 迟到率 × 订单量"', max_h=13.3)

page_break()
h2('P5 品类分析', '4.5')
add_figure(os.path.join(PBI, 'page-0005.png'),
           '仪表盘 P5  品类分析：office_furniture 差评 22.51% 且备货 9.92 天（物流型）；'
           'audio 差评 21.51% 但备货仅 2.51 天（商品型）', max_h=13.3)

page_break()
h2('P6 数据说明', '4.6')
add_figure(os.path.join(PBI, 'page-0006.png'),
           '仪表盘 P6  数据说明：数据来源、指标口径、分析方法、已知局限与修正记录', max_h=13.3)

portrait_section()

# ===========================================================================
# 第五章 业务建议
# ===========================================================================
h1('业务建议', '第五章')

h2('三条核心建议', '5.1')
add_table(['优先级', '措施', '数据依据', '预期效果', '成本'],
          [['P0', '停止为"数据好看"而收紧预估窗口',
            '2018-08 预估压至 15.2 天，实际送更快（7.7 天）但迟到率从 4.5% 涨到 10.4%',
            '迟到率可降 5–8 pp', '极低（改参数）'],
           ['P0', '建立卖家发货时效考核，超期扣分/降权',
            '9.06% 超期发货；超期单差评率 23.69%（基线 14.63%）',
            '超期率降到 5%，可减少约 1/3 相关差评', '低（定规则）'],
           ['P1', '大件/家具品类设独立承诺标准',
            'office_furniture 备货 9.92 天（均值 3.5 倍），差评 22.51%',
            '避免卖家被误判，减少品类性差评', '低'],
           ['P1', '旺季（黑五/圣诞）提前扩容运力',
            '2017-11 订单 +63%，迟到率 5.3% → 14.3%；2018-03 迟到 21.4%',
            '旺季迟到率减半', '中'],
           ['P2', '聚焦 Top 卖家治理（按影响面排序）',
            '单一卖家贡献 8,623 天超额延误（是第二名的 4.5 倍）',
            '集中资源，ROI 最高', '低'],
           ['P2', '多卖家订单体验专项',
            '2 卖家订单差评率 46.34%（全国 12.5%）',
            '该场景仅占 1.3%，但体验极差', '中'],
           ['—', '不建议盲目在偏远州投运力',
            '距离净效应 +0.064（控制运输时长后为正），d = −0.12',
            '投入产出比低', '—']],
          widths=[1.3, 4.2, 5.4, 3.4, 1.9], caption='按优先级排列的业务建议')

h2('指标体系建议', '5.2')
add_table(['层级', '指标', '建议目标', '说明'],
          [['北极星', '差评率（1-2 星占比）', '14.63% → 10% 以内', '直接反映客户满意度'],
           ['过程', '准时送达率', '92.28% → 96%', '★ 最重要，对应最强因素'],
           ['过程', '卖家超期发货率', '9.06% → 5% 以内', '迟到的上游原因，可考核'],
           ['过程', '平均备货天数', '2.81 → 2.2 天', '完全可控'],
           ['预警', '预估收紧幅度 ÷ 实际提速幅度', '比值 ≤ 1', '★ 防止为数据好看而收紧承诺'],
           ['分群', '各州 / 各品类准时率', '差异化目标', '避免全国平均掩盖结构问题']],
          widths=[2.4, 5.4, 4.4, 4.0], caption='建议的指标体系')
para('特别说明「预警指标」：2018 年 8 月的数据显示，平台把预估从 19.7 天压到 15.2 天，'
     '实际配送反而更快了（9.0 → 7.7 天），但迟到率从 4.5% 涨到 10.4%。'
     '如果不监控「承诺 vs 实际的相对变化」，这类问题会被误判为「物流变差」。', indent=0.8)

h2('落地路线图', '5.3')
add_table(['阶段', '时间', '动作', '预期收益'],
          [['第一阶段\n快速见效', '0–1 个月',
            '① 上线卖家超期发货考核\n② 冻结预估窗口收紧操作\n③ 建立差评率周报',
            '超期率 −2 pp，迟到率 −3 pp'],
           ['第二阶段\n结构优化', '1–3 个月',
            '① 按影响面排序治理 Top 50 卖家\n② 大件品类独立承诺标准\n③ 多卖家订单体验排查',
            '差评率 −2 pp'],
           ['第三阶段\n长期建设', '3–6 个月',
            '① 旺季运力弹性预案\n② 分州差异化承诺策略\n③ 远距离订单主动通知机制',
            '旺季迟到率减半']],
          widths=[2.6, 2.2, 7.4, 4.0], caption='三阶段落地路线图')

page_break()

# ===========================================================================
# 第六章 局限与修正记录
# ===========================================================================
h1('局限与修正记录', '第六章')
para('一份诚实的分析必须说明自己没做到什么、哪里可能出错。这一章也是本项目最贴近真实工作的部分。',
     indent=0.8)

h2('数据本身的局限', '6.1')
add_table(['局限', '具体影响', '本报告的处理'],
          [['订单只有一个签收时间戳', '分批发货时无法反映最晚送达，多卖家订单指标失真',
            '不采信该指标，并说明原因'],
           ['缺少物流商字段', '无法归因到具体承运方', '作为局限说明'],
           ['评论缺失率高（标题 88.3%、内容 58.7%）', '无法做文本情感分析', '只使用评分'],
           ['4 个月份数据不完整', '月度趋势首尾失真', '分析时剔除（HAVING 门槛）'],
           ['23 条负值时间戳', '污染时长计算', '剔除（占 0.02%）'],
           ['geolocation 重复采样', '100 万行仅 19,015 个邮编', '按邮编聚合取均值']],
          widths=[4.8, 6.2, 5.2], caption='数据局限性清单')

h2('分析方法的局限', '6.2')
bullet('因果推断有限：虽然具备时间先后、剂量反应和合理机制，但缺乏随机对照实验，'
       '严格意义上只能说"强相关 + 合理机制"，不能断言因果。')
bullet('R² 仅 0.17：评分为主观打分，大量未观测因素（商品质量、客服态度、包装）无法进入模型。'
       '模型适合做群体归因，不适合预测个体。')
bullet('未纳入卖家/品类固定效应：本报告用距离和件数等变量控制，但没有为每个卖家、'
       '每个品类加虚拟变量。加入后 R² 会提升，但会牺牲可解释性。')
bullet('main_category 取订单首个商品：约 1–2% 的订单品类归类不精确。')

h2('★ 修正记录：一个被推翻的结论', '6.3')
para('这是本项目最有价值的一段，因为它体现了「分析过程中如何自我纠错」。', indent=0.8)
add_table(['阶段', '当时的结论', '依据'],
          [['初版（单因素）', '"卖家超期发货是最大痛点"', '相关系数 r = −0.106，评分 3.70 vs 4.20'],
           ['补做多元回归后', '"超期发货的效应被迟到完全中介"',
            '控制迟到后 β = −0.0002（p = 0.97），OR = 1.013（p = 0.28）'],
           ['最终结论', '"真正的 KPI 是准时送达率，而非按时发货率"',
            '因果路径：超期发货 → 迟到 → 差评']],
          widths=[3.4, 6.4, 6.4], caption='结论的修正过程')
callout('这个修正为什么比原结论更有价值',
        '如果报告停在"超期发货是最大痛点"，后续建议就会写成"考核卖家发货率"，\n'
        '而客户真正在意的"是否按时收到"反而被忽略。\n\n'
        '补做多元回归后才发现：超期发货是"原因的原因"，真正的直接原因是迟到。\n'
        '这直接改变了建议的方向——从"考核发货率"改为"考核准时送达率 + 管住预估窗口"。',
        bg='E8F5E9', border='43A047', tcolor='1B5E20')

h2('如果重做一次', '6.4')
bullet('一开始就把「差评率」定为核心因变量，而不是中途从「配送时长」切换——'
       '前期有一半精力花在过程指标上，最后还要换算才能比较权重。')
bullet('更早做多元建模。做完所有单因素分析后才补回归，才发现两个结论需要修正。'
       '如果先建模再细查，能少走弯路。')
bullet('对多卖家订单的异常，应该第一时间查数据字典而不是先做统计——'
       '这类「指标失效」问题，业务理解比统计方法更有效。')
bullet('把清洗阈值、随机种子等参数外置成配置，便于复现和敏感性分析。')

page_break()

# ===========================================================================
# 第七章 技术实现与复现
# ===========================================================================
h1('技术实现与复现', '第七章')

h2('技术栈', '7.1')
add_table(['层次', '工具', '用途'],
          [['数据存储', 'PostgreSQL 18', '导入 9 张 CSV，建索引、外键、视图'],
           ['数据查询', 'SQL（psql / VS Code）', '31 条分析查询，含 CTE、窗口函数、FILTER 聚合'],
           ['统计分析', 'Python 3.11（pandas / SciPy / NumPy）', '数据清洗、相关分析、手写 OLS 推断'],
           ['机器学习', 'scikit-learn', 'Logistic 回归、随机森林、交叉验证'],
           ['可视化', 'Matplotlib / Seaborn', '8 张统计图表'],
           ['BI 报表', 'Power BI Desktop', '6 页交互式仪表盘'],
           ['文档', 'python-docx', '自动生成报告（便于数据更新后重跑）']],
          widths=[2.6, 5.8, 7.8], caption='技术栈')

h2('仓库结构', '7.2')
code_block(
    'olist-review-attribution/\n'
    '├── README.md                    # 项目说明与核心结论\n'
    '├── sql/\n'
    '│   ├── 01_create_schema.sql     # 建 schema 与表\n'
    '│   ├── 02_import_data.sql       # 导入 9 个 CSV\n'
    '│   ├── 03_constraints_views.sql # 外键、索引、分析视图\n'
    '│   ├── 04_data_exploration.sql  # 数据探索与质量检查\n'
    '│   ├── 06_factor_analysis.sql   # ★ 31 条归因查询（7 个部分）\n'
    '│   └── 07_pbi_view.sql          # ★ BI 专用订单级宽表\n'
    '├── python/\n'
    '│   ├── stat_analysis.py         # 相关分析、效应量、卡方检验\n'
    '│   ├── rigorous_model.py        # ★ OLS / Logit / 随机森林\n'
    '│   └── output/                  # 8 张统计图 + 13 个结果 CSV\n'
    '├── powerbi/\n'
    '│   ├── olist_dashboard.pbix\n'
    '│   ├── DAX_度量值与计算列.dax\n'
    '│   └── screenshots/             # 6 页仪表盘截图\n'
    '├── docs/\n'
    '│   └── report.pdf               # 分析报告\n'
    '└── .gitignore                   # 排除原始 CSV（120 MB）',
    caption='推荐的项目结构')

h2('关键代码片段', '7.3')
para('四段耗时拆解——整个分析的分诊环节，决定了后续所有分析方向：', indent=0.8)
code_block(
    'SELECT\n'
    '  -- ① 下单 → 审核\n'
    '  round(avg(EXTRACT(EPOCH FROM (order_approved_at - order_purchase_timestamp))\n'
    '            / 86400)::numeric, 2) AS 下单到审核_天,\n'
    '  -- ② 审核 → 交给物流（= 卖家备货时间）\n'
    '  round(avg(EXTRACT(EPOCH FROM (order_delivered_carrier_date - order_approved_at))\n'
    '            / 86400)::numeric, 2) AS 审核到发货_天,\n'
    '  -- ③ 交给物流 → 客户签收（= 运输时间）\n'
    '  round(avg(EXTRACT(EPOCH FROM (order_delivered_customer_date\n'
    '            - order_delivered_carrier_date)) / 86400)::numeric, 2) AS 发货到签收_天,\n'
    '  count(*) AS 样本数\n'
    'FROM olist.orders\n'
    '-- 过滤三个时间戳都非空：否则各列分母不同，四段加起来对不上全程\n'
    'WHERE order_delivered_customer_date IS NOT NULL\n'
    '  AND order_approved_at IS NOT NULL\n'
    '  AND order_delivered_carrier_date IS NOT NULL;',
    caption='SQL：四段耗时拆解')

para('手写 OLS 完整统计推断（不依赖 statsmodels，便于理解每一步）：', indent=0.8)
code_block(
    'def ols(y, X, names):\n'
    '    """最小二乘 + 完整统计推断。X 为设计矩阵（不含常数项）。"""\n'
    '    n = len(y)\n'
    '    Xd = np.column_stack([np.ones(n), X])       # 加截距项\n'
    '    beta, *_ = np.linalg.lstsq(Xd, y, rcond=None)  # β = (X\'X)^-1 X\'y\n'
    '    resid = y - Xd @ beta\n'
    '    k = Xd.shape[1]; dof = n - k\n'
    '    sigma2 = resid @ resid / dof\n'
    '    XtX_inv = np.linalg.pinv(Xd.T @ Xd)         # 伪逆，防共线性奇异\n'
    '    se = np.sqrt(np.diag(XtX_inv) * sigma2)\n'
    '    tval = beta / se\n'
    '    pval = 2 * (1 - stats.t.cdf(np.abs(tval), dof))\n'
    '    r2 = 1 - (resid @ resid) / ((y - y.mean()) ** 2).sum()\n'
    '    adj_r2 = 1 - (1 - r2) * (n - 1) / dof\n'
    '    return dict(beta=beta, se=se, t=tval, p=pval, r2=r2, adj_r2=adj_r2)',
    caption='Python：手写 OLS 推断（系数/标准误/t 值/p 值/R²）')

h2('复现步骤', '7.4')
para('1. 准备环境', bold=True, after=2)
bullet('安装 PostgreSQL 18，创建数据库 analysis')
bullet('Python 3.11 + pip install pandas numpy scipy scikit-learn matplotlib seaborn sqlalchemy psycopg2-binary python-docx')
para('2. 下载数据', bold=True, before=8, after=2)
bullet('从 Kaggle 下载 Olist Brazilian E-Commerce Public Dataset（9 个 CSV）')
bullet('放到项目 data/ 目录（该目录已加入 .gitignore，不进仓库）')
para('3. 建库导入', bold=True, before=8, after=2)
code_block(
    'psql -h 127.0.0.1 -U postgres -d analysis -f sql/01_create_schema.sql\n'
    'psql -h 127.0.0.1 -U postgres -d analysis -f sql/02_import_data.sql\n'
    'psql -h 127.0.0.1 -U postgres -d analysis -f sql/03_constraints_views.sql\n'
    'psql -h 127.0.0.1 -U postgres -d analysis -f sql/07_pbi_view.sql',
    caption='命令行导入')
para('4. 跑分析', bold=True, before=8, after=2)
code_block(
    'python python/stat_analysis.py     # 相关分析 → 5 张图 + 7 个 CSV\n'
    'python python/rigorous_model.py     # 多元建模 → 3 张图 + 6 个 CSV',
    caption='运行分析脚本')
para('5. 打开仪表盘', bold=True, before=8, after=2)
bullet('用 Power BI Desktop 打开 powerbi/olist_dashboard.pbix')
bullet('数据源指向本地 PostgreSQL 的 analysis 库（连接信息见 DAX 文件头部）')

page_break()

# ===========================================================================
# 附录
# ===========================================================================
h1('附录', '')

h2('A. 31 条分析查询的结构', 'A.1')
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
          widths=[3.0, 9.4, 4.0], caption='SQL 文件结构（全部逐条验证通过）', right_cols=[2])

h2('B. 关键数字速查', 'B.1')
add_table(['数字', '含义'],
          [['99,441 / 1,550,922', '订单数 / 总行数'],
           ['14.63%', '差评率基线（1-2 星占比）'],
           ['12.56 天', '平均全程配送时长'],
           ['0.43 / 2.81 / 9.33 天', '审核 / 备货 / 运输 三段耗时'],
           ['74.23%', '运输环节占全程比例'],
           ['53.94% vs 9.18%', '迟到 vs 准时 的差评率'],
           ['52.77% vs 8.79%', '已超期 vs 未超期 的差评率'],
           ['23.69% vs 11.68%', '超期发货 vs 按期发货 的差评率'],
           ['14.07% vs 10.43%', '跨州 vs 同州 的差评率'],
           ['9.06%', '卖家超期发货率'],
           ['8.11%', '迟到率'],
           ['−1.41', "「是否迟到」的 Cohen's d"],
           ['0.1695 / 0.1518 / 0.744', 'OLS R² / Logit 伪 R² / RF AUC'],
           ['93,047 / 10 维', '建模样本 / 特征数']],
          widths=[5.4, 10.8], caption='报告中出现的关键数字')

h2('C. 数据来源与许可', 'C.1')
para('数据来自 Olist 巴西电商平台公开数据集（Kaggle），采用 CC BY-NC-SA 4.0 许可。'
     '本仓库仅包含分析代码与结果，不含原始数据文件。'
     '数据集地址：https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce', indent=0.8)

para('')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(10)
_font(p.add_run('— 报告完 —'), FONT_CN, 11, True, '9CA3AF')

doc.save(OUT)
print('已生成: %s' % OUT)
print('大小: %.2f MB' % (os.path.getsize(OUT) / 1024 / 1024))
