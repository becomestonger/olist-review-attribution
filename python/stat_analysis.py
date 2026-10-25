"""
================================================================================
 Olist 巴西电商 —— 差评归因的统计学分析
================================================================================

研究问题：
    哪些因素影响客户满意度（评分 / 差评）？各因素的影响强度如何？
    用相关系数矩阵 + 显著性检验 + 效应量，给出数学层面的证据。

分析内容：
    第 1 部分  从 PostgreSQL 读取订单级特征数据
    第 2 部分  描述性统计（均值/标准差/分布）
    第 3 部分  Pearson 相关系数矩阵（线性关系）
    第 4 部分  Spearman 秩相关系数矩阵（单调关系，抗非线性/异常值）
    第 5 部分  各因素与目标变量的相关 + 显著性检验（p 值）
    第 6 部分  组间差异检验（t 检验 / Mann-Whitney U / Cohen's d 效应量）
    第 7 部分  分类变量关联检验（卡方）
    第 8 部分  阈值效应验证（为什么 Pearson 会低估）
    第 9 部分  可视化（热力图 / 散点 / 柱状 / 悬崖效应图）
    第 10 部分 导出结果到 analysis/output/

运行方式（Windows，本机 Python 3.11）：
    python python/stat_analysis.py

依赖：numpy pandas matplotlib seaborn scipy sqlalchemy psycopg2-binary
================================================================================
"""

import os
import warnings

import numpy as np
import pandas as pd
from sqlalchemy import create_engine, text

import matplotlib
matplotlib.use('Agg')                      # 无界面后端，直接存图
import matplotlib.pyplot as plt
import seaborn as sns

from scipy import stats

warnings.filterwarnings('ignore')

# ---------------------------------------------------------------------------
# 0. 全局配置
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 输出目录跟着脚本走，与文件夹改名无关
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
os.makedirs(OUT_DIR, exist_ok=True)

DB = dict(host='127.0.0.1', port=5432, db='analysis',
          user='postgres', pwd='123456')
ENGINE_URL = 'postgresql+psycopg2://{user}:{pwd}@{host}:{port}/{db}'.format(**DB)

# 中文显示
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 120
plt.rcParams['savefig.dpi'] = 150
plt.rcParams['axes.facecolor'] = 'white'
sns.set_style('whitegrid', {'font.sans-serif': ['Microsoft YaHei', 'SimHei']})

SEP = '=' * 78


def section(title):
    print('\n' + SEP)
    print(' ' + title)
    print(SEP)


# ---------------------------------------------------------------------------
# 1. 读取数据
# ---------------------------------------------------------------------------
section('第 1 部分  读取订单级特征数据')

SQL = """
WITH first_item AS (
    SELECT DISTINCT ON (order_id) order_id, seller_id, product_id
    FROM olist.order_items
    ORDER BY order_id, order_item_id
),
order_agg AS (
    SELECT order_id,
           sum(price)         AS items_total,
           sum(freight_value) AS freight_total,
           count(*)           AS item_count,
           max(shipping_limit_date) AS ship_limit
    FROM olist.order_items
    GROUP BY order_id
),
review AS (
    SELECT order_id, max(review_score) AS review_score
    FROM olist.order_reviews
    GROUP BY order_id
)
SELECT o.order_id,
       EXTRACT(EPOCH FROM (o.order_approved_at - o.order_purchase_timestamp)) / 86400.0          AS approve_days,
       EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at)) / 86400.0     AS prep_days,
       EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_delivered_carrier_date)) / 86400.0 AS ship_days,
       EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_purchase_timestamp)) / 86400.0     AS total_days,
       EXTRACT(EPOCH FROM (o.order_estimated_delivery_date - o.order_delivered_customer_date)) / 86400.0 AS buffer_days,
       CASE WHEN o.order_delivered_customer_date > o.order_estimated_delivery_date THEN 1 ELSE 0 END AS is_late,
       CASE WHEN o.order_delivered_carrier_date > a.ship_limit THEN 1 ELSE 0 END                     AS is_overdue_ship,
       CASE WHEN c.customer_state <> s.seller_state THEN 1 ELSE 0 END                                AS is_cross_state,
       a.items_total,
       a.freight_total,
       a.item_count,
       p.product_weight_g AS weight_g,
       r.review_score
FROM olist.orders o
JOIN olist.customers c   ON c.customer_id = o.customer_id
JOIN first_item f        ON f.order_id    = o.order_id
JOIN olist.sellers s     ON s.seller_id   = f.seller_id
JOIN order_agg a         ON a.order_id    = o.order_id
LEFT JOIN olist.products p ON p.product_id = f.product_id
LEFT JOIN review r         ON r.order_id   = o.order_id
WHERE o.order_delivered_customer_date IS NOT NULL
  AND o.order_delivered_carrier_date IS NOT NULL
  AND o.order_approved_at IS NOT NULL
  AND o.order_estimated_delivery_date IS NOT NULL
"""

engine = create_engine(ENGINE_URL, pool_pre_ping=True)
with engine.connect() as conn:
    df = pd.read_sql(text(SQL), conn)

print('原始记录数: %s' % format(len(df), ','))

# 派生变量
df['is_bad_review'] = np.where(df['review_score'].notna(),
                               (df['review_score'] <= 2).astype(float), np.nan)
df['freight_ratio'] = df['freight_total'] / df['items_total'].replace(0, np.nan)

# 清洗：去掉负值时间戳与极端离群（保留 99.5% 分位以内，避免极端值主导相关系数）
before = len(df)
for col in ['approve_days', 'prep_days', 'ship_days', 'total_days']:
    df = df[df[col] >= 0]
df = df[df['total_days'] <= df['total_days'].quantile(0.995)]
df = df[df['items_total'].between(0, df['items_total'].quantile(0.995))]
print('清洗后记录数: %s（剔除 %s 条异常/极端值）'
      % (format(len(df), ','), format(before - len(df), ',')))

# 特征列定义
CONT_COLS = ['approve_days', 'prep_days', 'ship_days', 'total_days', 'buffer_days',
             'items_total', 'freight_total', 'freight_ratio', 'item_count', 'weight_g']
BIN_COLS = ['is_late', 'is_overdue_ship', 'is_cross_state', 'is_bad_review']
TARGETS = ['review_score', 'is_bad_review']

LABELS = {
    'approve_days':    '审核时长(天)',
    'prep_days':       '备货时长(天)',
    'ship_days':       '运输时长(天)',
    'total_days':      '全程时长(天)',
    'buffer_days':     '预估富余(天)',
    'items_total':     '商品金额',
    'freight_total':   '运费',
    'freight_ratio':   '运费占比',
    'item_count':      '商品件数',
    'weight_g':        '商品重量(g)',
    'is_late':         '是否迟到',
    'is_overdue_ship': '卖家超期发货',
    'is_cross_state':  '是否跨州',
    'review_score':    '评分',
    'is_bad_review':   '是否差评',
}


# ---------------------------------------------------------------------------
# 2. 描述性统计
# ---------------------------------------------------------------------------
section('第 2 部分  描述性统计')

desc = df[CONT_COLS + ['review_score']].describe().T
desc = desc[['count', 'mean', 'std', 'min', '25%', '50%', '75%', 'max']]
desc.index = [LABELS.get(i, i) for i in desc.index]
desc.columns = ['样本数', '均值', '标准差', '最小值', 'P25', '中位数', 'P75', '最大值']
print(desc.round(2).to_string())

print('\n二元变量占比:')
for c in BIN_COLS:
    v = df[c].dropna()
    if len(v):
        print('  %-14s 为1的比例 = %6.2f%%   (n=%s)'
              % (LABELS.get(c, c), 100 * v.mean(), format(len(v), ',')))

print('\n偏度与峰度（判断是否近似正态）:')
sk = pd.DataFrame({
    '偏度 skew': df[CONT_COLS + ['review_score']].skew(),
    '峰度 kurt': df[CONT_COLS + ['review_score']].kurt(),
})
sk.index = [LABELS.get(i, i) for i in sk.index]
print(sk.round(3).to_string())
print('\n注：偏度绝对值 > 1 说明分布明显偏斜，此时 Pearson 相关不可靠，应看 Spearman。')


# ---------------------------------------------------------------------------
# 3. Pearson 相关系数矩阵
# ---------------------------------------------------------------------------
section('第 3 部分  Pearson 相关系数矩阵（线性关系）')

corr_cols = CONT_COLS + BIN_COLS + ['review_score']
pearson = df[corr_cols].corr(method='pearson')

pearson_named = pearson.copy()
pearson_named.index = [LABELS.get(i, i) for i in pearson_named.index]
pearson_named.columns = [LABELS.get(i, i) for i in pearson_named.columns]
print(pearson_named.round(3).to_string())


# ---------------------------------------------------------------------------
# 4. Spearman 秩相关系数矩阵
# ---------------------------------------------------------------------------
section('第 4 部分  Spearman 秩相关系数矩阵（单调关系）')

spearman = df[corr_cols].corr(method='spearman')
spearman_named = spearman.copy()
spearman_named.index = [LABELS.get(i, i) for i in spearman_named.index]
spearman_named.columns = [LABELS.get(i, i) for i in spearman_named.columns]
print(spearman_named.round(3).to_string())


# ---------------------------------------------------------------------------
# 5. 各因素与目标变量的相关 + 显著性
# ---------------------------------------------------------------------------
section('第 5 部分  各因素 vs 目标变量：相关系数与显著性(p 值)')

rows = []
for feat in CONT_COLS + BIN_COLS:
    if feat in TARGETS:
        continue
    for tgt in TARGETS:
        sub = df[[feat, tgt]].dropna()
        if len(sub) < 100 or sub[feat].nunique() < 2:
            continue
        r_p, p_p = stats.pearsonr(sub[feat], sub[tgt])
        r_s, p_s = stats.spearmanr(sub[feat], sub[tgt])
        rows.append({
            '因素': LABELS.get(feat, feat),
            '目标': LABELS.get(tgt, tgt),
            '样本数': len(sub),
            'Pearson r': r_p,
            'p值(Pearson)': p_p,
            'Spearman ρ': r_s,
            'p值(Spearman)': p_s,
            '显著(p<0.001)': '***' if p_p < 0.001 else ('*' if p_p < 0.05 else 'ns'),
        })

corr_target = pd.DataFrame(rows)
corr_target['|r|'] = corr_target['Pearson r'].abs()

print('\n【与「评分」的相关，按 |Pearson r| 降序】')
show = corr_target[corr_target['目标'] == LABELS['review_score']].sort_values('|r|', ascending=False)
print(show[['因素', '样本数', 'Pearson r', 'p值(Pearson)', 'Spearman ρ', '显著(p<0.001)']]
      .round(4).to_string(index=False))

print('\n【与「是否差评」的相关，按 |Pearson r| 降序】')
show2 = corr_target[corr_target['目标'] == LABELS['is_bad_review']].sort_values('|r|', ascending=False)
print(show2[['因素', '样本数', 'Pearson r', 'p值(Pearson)', 'Spearman ρ', '显著(p<0.001)']]
      .round(4).to_string(index=False))

print('\n统计口径说明：')
print('  * Pearson r 衡量线性相关，取值 [-1, 1]，绝对值越大关系越强')
print('  * Spearman ρ 衡量单调相关（基于秩），对异常值和非线性更稳健')
print('  * p 值 < 0.001 表示在 99.9% 置信水平下相关显著不为 0（***）')
print('  * 样本量很大时 p 值极易显著，因此需同时看效应大小 |r|')


# ---------------------------------------------------------------------------
# 6. 组间差异检验 + 效应量
# ---------------------------------------------------------------------------
section('第 6 部分  组间差异检验（t 检验 / Mann-Whitney U / Cohen\'s d）')


def cohens_d(a, b):
    """效应量：均值差除以合并标准差。0.2 小 / 0.5 中 / 0.8 大。"""
    a, b = np.asarray(a, float), np.asarray(b, float)
    na, nb = len(a), len(b)
    if na < 2 or nb < 2:
        return np.nan
    va, vb = a.var(ddof=1), b.var(ddof=1)
    pooled = np.sqrt(((na - 1) * va + (nb - 1) * vb) / (na + nb - 2))
    return (a.mean() - b.mean()) / pooled if pooled > 0 else np.nan


group_rows = []
for col, label in [('is_late', '是否迟到'), ('is_overdue_ship', '是否超期发货'),
                   ('is_cross_state', '是否跨州')]:
    g1 = df.loc[df[col] == 1, 'review_score'].dropna()
    g0 = df.loc[df[col] == 0, 'review_score'].dropna()
    if len(g1) < 30 or len(g0) < 30:
        continue
    t_stat, t_p = stats.ttest_ind(g1, g0, equal_var=False)
    u_stat, u_p = stats.mannwhitneyu(g1, g0, alternative='two-sided')
    d = cohens_d(g1, g0)
    group_rows.append({
        '分组变量': label,
        '组1(=是)': 'n=%s 均值=%.2f' % (format(len(g1), ','), g1.mean()),
        '组0(=否)': 'n=%s 均值=%.2f' % (format(len(g0), ','), g0.mean()),
        '均值差': g1.mean() - g0.mean(),
        "Cohen's d": d,
        't检验p值': t_p,
        'Mann-Whitney p': u_p,
    })

grp = pd.DataFrame(group_rows)
print(grp.round(4).to_string(index=False))

print("\n效应量判读（Cohen 标准）: |d|<0.2 微小, 0.2-0.5 小, 0.5-0.8 中, >0.8 大")


# ---------------------------------------------------------------------------
# 7. 分类变量关联检验（卡方）
# ---------------------------------------------------------------------------
section('第 7 部分  二元变量关联检验（卡方 + Cramér\'s V）')

print('\n  %-26s %-22s %-10s %-10s %s' % ('变量1', '变量2', '卡方值', 'p值', "Cramér's V"))
print('  ' + '-' * 86)

pairs = [('is_late', 'is_bad_review'), ('is_overdue_ship', 'is_bad_review'),
         ('is_cross_state', 'is_bad_review'), ('is_overdue_ship', 'is_late'),
         ('is_cross_state', 'is_late')]

chi_rows = []
for a, b in pairs:
    sub = df[[a, b]].dropna()
    if len(sub) < 100:
        continue
    table_ = pd.crosstab(sub[a], sub[b])
    if table_.shape != (2, 2):
        continue
    chi2, p, dof, _ = stats.chi2_contingency(table_, correction=False)
    n = table_.values.sum()
    v = np.sqrt(chi2 / (n * (min(table_.shape) - 1)))
    print('  %-26s %-22s %-10.1f %-10.3e %.4f'
          % (LABELS.get(a, a), LABELS.get(b, b), chi2, p, v))
    chi_rows.append({'变量1': LABELS.get(a, a), '变量2': LABELS.get(b, b),
                     '卡方值': chi2, 'p值': p, "Cramér's V": v})

print("\n  Cramér's V：0.1 弱, 0.3 中, 0.5 强（分类变量间的关联强度）")


# ---------------------------------------------------------------------------
# 8. 阈值效应验证：为什么 Pearson 会低估
# ---------------------------------------------------------------------------
section('第 8 部分  阈值效应验证（Pearson 低估的关键原因）')

d8 = df.dropna(subset=['buffer_days', 'is_bad_review']).copy()
bins = [-100, -0.0001, 3, 7, 15, 30, 1000]
labels8 = ['已超期', '富余0-3天', '富余3-7天', '富余7-15天', '富余15-30天', '富余30天以上']
d8['富余分档'] = pd.cut(d8['buffer_days'], bins=bins, labels=labels8)

grp8 = d8.groupby('富余分档').agg(
    订单数=('is_bad_review', 'size'),
    差评率=('is_bad_review', lambda s: 100 * s.mean()),
    平均评分=('review_score', 'mean'),
).round(2)
print(grp8.to_string())

# 分段相关：分界线两侧各算一次 Pearson
late_part = d8[d8['buffer_days'] < 0]
ontime_part = d8[d8['buffer_days'] >= 0]

r_all, _ = stats.pearsonr(d8['buffer_days'], d8['review_score'].fillna(d8['review_score'].mean()))
print('\n全部样本  富余 vs 评分  Pearson r = %.4f' % r_all)
if len(ontime_part) > 100:
    r_on, p_on = stats.pearsonr(ontime_part['buffer_days'], ontime_part['review_score'].dropna())
    print('仅未超期  富余 vs 评分  Pearson r = %.4f (p=%.3g)  ← 关系明显变弱'
          % (r_on, p_on))

print('\n结论：整体相关系数被"已超期"这一小部分样本主导。')
print('      在"未超期"的子集内，富余天数与评分的线性关系大幅减弱，')
print('      说明真实关系是阈值型（悬崖），不是线性型（斜坡）。')


# ---------------------------------------------------------------------------
# 9. 可视化
# ---------------------------------------------------------------------------
section('第 9 部分  生成可视化图表')

# ---- 图1：Pearson / Spearman 热力图 --------------------------------------
fig, axes = plt.subplots(1, 2, figsize=(18, 7.5))

for ax, mat, title in [(axes[0], pearson_named, 'Pearson 相关系数（线性）'),
                       (axes[1], spearman_named, 'Spearman 秩相关（单调）')]:
    mask = np.triu(np.ones_like(mat, dtype=bool), k=1)
    sns.heatmap(mat, mask=mask, annot=True, fmt='.2f', cmap='RdBu_r',
                center=0, vmin=-1, vmax=1, square=True, ax=ax,
                annot_kws={'size': 7}, linewidths=.4, linecolor='white',
                cbar_kws={'shrink': .75, 'label': '相关系数'})
    ax.set_title(title, fontsize=13, pad=12)
    ax.tick_params(axis='x', rotation=45, labelsize=9)
    ax.tick_params(axis='y', rotation=0, labelsize=9)

plt.tight_layout()
p1 = os.path.join(OUT_DIR, 'fig1_相关系数热力图.png')
plt.savefig(p1, bbox_inches='tight')
plt.close()
print('  已保存: %s' % os.path.basename(p1))

# ---- 图2：与评分/差评的相关强度条形图 -----------------------------------
fig, axes = plt.subplots(1, 2, figsize=(15, 6))

for ax, tgt in [(axes[0], 'review_score'), (axes[1], 'is_bad_review')]:
    sub = corr_target[corr_target['目标'] == LABELS[tgt]].copy()
    sub = sub.sort_values('Pearson r')
    colors = ['#c0392b' if v > 0 else '#2980b9' for v in sub['Pearson r']]
    ax.barh(sub['因素'], sub['Pearson r'], color=colors)
    ax.axvline(0, color='#333', lw=.9)
    ax.set_xlabel('Pearson r')
    ax.set_title('各因素与「%s」的相关系数' % LABELS[tgt], fontsize=12)
    for i, (v, s) in enumerate(zip(sub['Pearson r'], sub['Spearman ρ'])):
        ax.text(v + (0.008 if v >= 0 else -0.008), i,
                'r=%.3f (ρ=%.3f)' % (v, s),
                va='center', ha='left' if v >= 0 else 'right', fontsize=7.5)
    ax.margins(x=0.28)

plt.tight_layout()
p2 = os.path.join(OUT_DIR, 'fig2_各因素相关强度.png')
plt.savefig(p2, bbox_inches='tight')
plt.close()
print('  已保存: %s' % os.path.basename(p2))

# ---- 图3：关键关系的散点 + 回归 ------------------------------------------
fig, axes = plt.subplots(1, 3, figsize=(17, 5))
pairs3 = [('total_days', 'review_score', '全程时长 vs 评分'),
          ('ship_days', 'review_score', '运输时长 vs 评分'),
          ('prep_days', 'review_score', '备货时长 vs 评分')]

for ax, (x, y, title) in zip(axes, pairs3):
    sub = df[[x, y]].dropna()
    # 抽样绘制，避免 10 万个点糊成一团
    samp = sub.sample(min(len(sub), 12000), random_state=42)
    ax.scatter(samp[x], samp[y], s=3, alpha=.18, color='#2c7fb8', edgecolors='none')
    # 分箱均值线（更能反映真实趋势）
    qb = pd.qcut(sub[x], 20, duplicates='drop')
    bm = sub.groupby(qb, observed=True).agg(mx=(x, 'mean'), my=(y, 'mean'))
    ax.plot(bm['mx'], bm['my'], color='#d94701', lw=2.2, marker='o', ms=3.5,
            label='分箱均值')
    r_p, _ = stats.pearsonr(sub[x], sub[y])
    r_s, _ = stats.spearmanr(sub[x], sub[y])
    ax.set_xlabel(LABELS[x])
    ax.set_ylabel('评分')
    ax.set_title('%s\nr=%.3f  ρ=%.3f' % (title, r_p, r_s), fontsize=11)
    ax.legend(fontsize=8)
    ax.set_ylim(0, 5.4)

plt.tight_layout()
p3 = os.path.join(OUT_DIR, 'fig3_关键散点与趋势.png')
plt.savefig(p3, bbox_inches='tight')
plt.close()
print('  已保存: %s' % os.path.basename(p3))

# ---- 图4：悬崖效应（阈值型关系）------------------------------------------
fig, axes = plt.subplots(1, 2, figsize=(15, 5.2))

ax = axes[0]
x = range(len(grp8))
ax.bar(x, grp8['差评率'], color=['#c0392b'] + ['#7fb3d5'] * (len(grp8) - 1))
ax.set_xticks(list(x))
ax.set_xticklabels(grp8.index, rotation=25, ha='right', fontsize=9)
ax.set_ylabel('差评率 (%)')
ax.set_title('悬崖效应：一旦超期，差评率骤升', fontsize=12)
for i, v in enumerate(grp8['差评率']):
    ax.text(i, v + 1.2, '%.1f%%' % v, ha='center', fontsize=9)
ax.axvline(0.5, color='#333', ls='--', lw=1)
ax.set_ylim(0, max(grp8['差评率']) * 1.25)

ax = axes[1]
ax.bar(x, grp8['平均评分'], color=['#c0392b'] + ['#5dade2'] * (len(grp8) - 1))
ax.set_xticks(list(x))
ax.set_xticklabels(grp8.index, rotation=25, ha='right', fontsize=9)
ax.set_ylabel('平均评分')
ax.set_ylim(0, 5)
ax.set_title('同一现象在评分上的表现', fontsize=12)
for i, v in enumerate(grp8['平均评分']):
    ax.text(i, v + .1, '%.2f' % v, ha='center', fontsize=9)
ax.axvline(0.5, color='#333', ls='--', lw=1)

plt.tight_layout()
p4 = os.path.join(OUT_DIR, 'fig4_悬崖效应.png')
plt.savefig(p4, bbox_inches='tight')
plt.close()
print('  已保存: %s' % os.path.basename(p4))

# ---- 图5：组间对比箱线图 -------------------------------------------------
fig, axes = plt.subplots(1, 2, figsize=(13, 5.2))

for ax, col, label in [(axes[0], 'is_late', '是否迟到'),
                       (axes[1], 'is_overdue_ship', '卖家是否超期发货')]:
    plot_df = df[[col, 'review_score']].dropna().copy()
    plot_df[label] = plot_df[col].map({0: '否', 1: '是'})
    sns.boxplot(data=plot_df, x=label, y='review_score', ax=ax,
                palette={'否': '#5dade2', '是': '#c0392b'}, width=.45,
                showfliers=False)
    sns.stripplot(data=plot_df.sample(min(len(plot_df), 4000), random_state=1),
                  x=label, y='review_score', ax=ax, size=1.6, alpha=.18,
                  color='#333')
    g1 = plot_df.loc[plot_df[col] == 1, 'review_score']
    g0 = plot_df.loc[plot_df[col] == 0, 'review_score']
    d = cohens_d(g1, g0)
    ax.set_title('%s（Cohen\'s d = %.2f）' % (label, d), fontsize=12)
    ax.set_xlabel('')
    ax.set_ylabel('评分')
    ax.set_ylim(0, 5.5)

plt.tight_layout()
p5 = os.path.join(OUT_DIR, 'fig5_组间对比箱线图.png')
plt.savefig(p5, bbox_inches='tight')
plt.close()
print('  已保存: %s' % os.path.basename(p5))


# ---------------------------------------------------------------------------
# 10. 导出结果
# ---------------------------------------------------------------------------
section('第 10 部分  导出结果')

exports = {
    'corr_pearson.csv': pearson_named.round(4),
    'corr_spearman.csv': spearman_named.round(4),
    'target_correlations.csv': corr_target.round(6),
    'group_tests.csv': grp.round(6),
    'chi_square.csv': pd.DataFrame(chi_rows).round(6),
    'buffer_bins.csv': grp8,
    'descriptive_stats.csv': desc.round(3),
}
for fname, obj in exports.items():
    p = os.path.join(OUT_DIR, fname)
    obj.to_csv(p, encoding='utf-8-sig')
    print('  已导出: %s' % fname)

section('分析完成')
print('输出目录: %s' % OUT_DIR)
print('图表: 5 张 PNG')
print('数据: 7 个 CSV（可用 Excel 打开）')
