"""
================================================================================
 Olist 差评归因 —— 严谨多元统计建模
================================================================================

本脚本补充单因素分析的不足，用多元模型控制混杂因素后估计各因素的净效应。

包含：
    A. 特征扩充：用 geolocation 经纬度计算客户与卖家的真实直线距离（Haversine）
    B. 多元线性回归（OLS）：因变量 = 评分，输出标准化系数、t 值、p 值、R²、VIF
    C. 二元逻辑回归（Logit）：因变量 = 是否差评，输出优势比 OR 与 95% 置信区间
    D. 随机森林特征重要性：交叉验证，作为非参数稳健性对照
    E. 交互效应检验：迟到 × 距离、迟到 × 超期发货
    F. 稳健性检验：分年度子样本 + 剔除极端值
    G. 边际效应与业务解读

依赖：numpy pandas scipy sklearn matplotlib seaborn sqlalchemy psycopg2
运行：python analysis/rigorous_model.py
================================================================================
"""

import os
import warnings

import numpy as np
import pandas as pd
from scipy import stats
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import cross_val_score
from sklearn.preprocessing import StandardScaler
from sqlalchemy import create_engine, text

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns

warnings.filterwarnings('ignore')

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, 'analysis', 'output')
os.makedirs(OUT, exist_ok=True)

DB = dict(host='127.0.0.1', port=5432, db='analysis', user='postgres', pwd='123456')
URL = 'postgresql+psycopg2://{user}:{pwd}@{host}:{port}/{db}'.format(**DB)

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 120
plt.rcParams['savefig.dpi'] = 150
sns.set_style('whitegrid', {'font.sans-serif': ['Microsoft YaHei', 'SimHei']})

SEP = '=' * 80
def sec(t):
    print('\n' + SEP + '\n ' + t + '\n' + SEP)


# =============================================================================
# A. 数据准备：订单特征 + 真实地理距离
# =============================================================================
sec('A. 读取数据并计算真实地理距离')

SQL = """
WITH first_item AS (
    SELECT DISTINCT ON (order_id) order_id, seller_id, product_id
    FROM olist.order_items ORDER BY order_id, order_item_id
),
order_agg AS (
    SELECT order_id, sum(price) AS items_total, sum(freight_value) AS freight_total,
           count(*) AS item_count, max(shipping_limit_date) AS ship_limit
    FROM olist.order_items GROUP BY order_id
),
review AS (
    SELECT order_id, max(review_score) AS review_score
    FROM olist.order_reviews GROUP BY order_id
)
SELECT o.order_id,
       EXTRACT(EPOCH FROM (o.order_approved_at - o.order_purchase_timestamp))/86400.0          AS approve_days,
       EXTRACT(EPOCH FROM (o.order_delivered_carrier_date - o.order_approved_at))/86400.0     AS prep_days,
       EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_delivered_carrier_date))/86400.0 AS ship_days,
       EXTRACT(EPOCH FROM (o.order_delivered_customer_date - o.order_purchase_timestamp))/86400.0     AS total_days,
       EXTRACT(EPOCH FROM (o.order_estimated_delivery_date - o.order_delivered_customer_date))/86400.0 AS buffer_days,
       CASE WHEN o.order_delivered_customer_date > o.order_estimated_delivery_date THEN 1 ELSE 0 END AS is_late,
       CASE WHEN o.order_delivered_carrier_date > a.ship_limit THEN 1 ELSE 0 END                     AS is_overdue_ship,
       CASE WHEN c.customer_state <> s.seller_state THEN 1 ELSE 0 END                                AS is_cross_state,
       c.customer_zip_code_prefix AS czip,
       s.seller_zip_code_prefix   AS szip,
       a.items_total, a.freight_total, a.item_count,
       p.product_weight_g AS weight_g,
       t.product_category_name_english AS category,
       r.review_score,
       date_part('year', o.order_purchase_timestamp) AS yr
FROM olist.orders o
JOIN olist.customers c   ON c.customer_id = o.customer_id
JOIN first_item f        ON f.order_id    = o.order_id
JOIN olist.sellers s     ON s.seller_id   = f.seller_id
JOIN order_agg a         ON a.order_id    = o.order_id
LEFT JOIN olist.products p ON p.product_id = f.product_id
LEFT JOIN olist.product_category_translation t ON t.product_category_name = p.product_category_name
LEFT JOIN review r         ON r.order_id = o.order_id
WHERE o.order_delivered_customer_date IS NOT NULL
  AND o.order_delivered_carrier_date IS NOT NULL
  AND o.order_approved_at IS NOT NULL
  AND o.order_estimated_delivery_date IS NOT NULL
"""

engine = create_engine(URL, pool_pre_ping=True)
with engine.connect() as conn:
    df = pd.read_sql(text(SQL), conn)
    geo = pd.read_sql(text("""
        SELECT geolocation_zip_code_prefix AS zip,
               avg(geolocation_lat) AS lat,
               avg(geolocation_lng) AS lng
        FROM olist.geolocation
        GROUP BY 1
    """), conn)

print('订单记录数: %s' % format(len(df), ','))
print('邮编坐标表: %s 个邮编' % format(len(geo), ','))

geo['zip'] = geo['zip'].astype(float)
df['czip'] = pd.to_numeric(df['czip'], errors='coerce')
df['szip'] = pd.to_numeric(df['szip'], errors='coerce')

g = geo.rename(columns={'zip': 'czip', 'lat': 'clat', 'lng': 'clng'})
df = df.merge(g, on='czip', how='left')
g2 = geo.rename(columns={'zip': 'szip', 'lat': 'slat', 'lng': 'slng'})
df = df.merge(g2, on='szip', how='left')


def haversine(lat1, lon1, lat2, lon2):
    """Haversine 公式：返回两点间大圆距离（km）。"""
    R = 6371.0
    p1, p2 = np.radians(lat1), np.radians(lat2)
    dp = p2 - p1
    dl = np.radians(lon2 - lon1)
    a = np.sin(dp / 2) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(dl / 2) ** 2
    return 2 * R * np.arcsin(np.sqrt(a))


df['distance_km'] = haversine(df['clat'], df['clng'], df['slat'], df['slng'])
have_d = df['distance_km'].notna()
print('能匹配到经纬度的订单: %s（%.1f%%）'
      % (format(int(have_d.sum()), ','), 100 * have_d.mean()))
print('距离统计（km）: 均值 %.0f  中位数 %.0f  p90 %.0f  最大 %.0f'
      % (df['distance_km'].mean(), df['distance_km'].median(),
         df['distance_km'].quantile(.9), df['distance_km'].max()))

# 清洗
before = len(df)
for col in ['approve_days', 'prep_days', 'ship_days', 'total_days']:
    df = df[df[col] >= 0]
df = df[df['total_days'] <= df['total_days'].quantile(0.995)]
df = df[df['items_total'].between(0, df['items_total'].quantile(0.995))]
print('清洗后: %s 条（剔除 %s 条异常值）'
      % (format(len(df), ','), format(before - len(df), ',')))

# 派生变量（对数变换降低偏度）
df['is_bad_review'] = np.where(df['review_score'].notna(),
                               (df['review_score'] <= 2).astype(float), np.nan)
df['log_distance'] = np.log1p(df['distance_km'])
df['log_weight'] = np.log1p(df['weight_g'])
df['log_price'] = np.log1p(df['items_total'])
df['freight_ratio'] = df['freight_total'] / df['items_total'].replace(0, np.nan)

FEATURES = ['prep_days', 'ship_days', 'is_late', 'is_overdue_ship',
            'log_distance', 'is_cross_state', 'log_weight', 'log_price',
            'item_count', 'freight_ratio']

LABEL = {
    'prep_days': '备货时长(天)', 'ship_days': '运输时长(天)',
    'is_late': '是否迟到', 'is_overdue_ship': '卖家超期发货',
    'log_distance': '距离(对数km)', 'is_cross_state': '是否跨州',
    'log_weight': '重量(对数g)', 'log_price': '金额(对数)',
    'item_count': '商品件数', 'freight_ratio': '运费占比',
}

model_df = df.dropna(subset=FEATURES + ['review_score', 'is_bad_review']).copy()
print('建模样本: %s 条，特征 %d 个' % (format(len(model_df), ','), len(FEATURES)))


# =============================================================================
# 通用 OLS 推断（手写，不依赖 statsmodels）
# =============================================================================
def ols(y, X, names):
    """最小二乘 + 完整统计推断。X 为设计矩阵（不含常数项）。"""
    n = len(y)
    Xd = np.column_stack([np.ones(n), X])
    all_names = ['常数项'] + names
    beta, *_ = np.linalg.lstsq(Xd, y, rcond=None)
    resid = y - Xd @ beta
    k = Xd.shape[1]
    dof = n - k
    sigma2 = resid @ resid / dof
    XtX_inv = np.linalg.pinv(Xd.T @ Xd)
    se = np.sqrt(np.diag(XtX_inv) * sigma2)
    tval = beta / se
    pval = 2 * (1 - stats.t.cdf(np.abs(tval), dof))
    tcrit = stats.t.ppf(0.975, dof)
    ss_tot = ((y - y.mean()) ** 2).sum()
    ss_res = resid @ resid
    r2 = 1 - ss_res / ss_tot
    adj_r2 = 1 - (1 - r2) * (n - 1) / dof
    fstat = (r2 / (k - 1)) / ((1 - r2) / dof)
    fp = 1 - stats.f.cdf(fstat, k - 1, dof)
    return {
        'names': all_names, 'beta': beta, 'se': se, 't': tval, 'p': pval,
        'ci_lo': beta - tcrit * se, 'ci_hi': beta + tcrit * se,
        'n': n, 'dof': dof, 'r2': r2, 'adj_r2': adj_r2,
        'f': fstat, 'f_p': fp, 'sigma': np.sqrt(sigma2),
    }


def stars(p):
    return '***' if p < 0.001 else ('**' if p < 0.01 else ('*' if p < 0.05 else 'ns'))


def vif_table(X, names):
    """方差膨胀因子：VIF_i = 1 / (1 - R²_i)。>10 提示严重共线性。"""
    out = []
    for i, nm in enumerate(names):
        y = X[:, i]
        others = np.delete(X, i, axis=1)
        r = ols(y, others, ['x'] * others.shape[1])
        v = 1 / (1 - r['r2']) if r['r2'] < 1 else np.inf
        out.append((nm, r['r2'], v))
    return out


def fmt_p(p):
    if p < 1e-300:
        return '<1e-300'
    if p < 0.001:
        return '%.2e' % p
    return '%.4f' % p


# =============================================================================
# B. 多元线性回归（OLS）
# =============================================================================
sec('B. 多元线性回归：因变量 = 评分（review_score）')

Xraw = model_df[FEATURES].values.astype(float)
y_score = model_df['review_score'].values.astype(float)

# 标准化，便于比较系数大小
scaler = StandardScaler()
Xstd = scaler.fit_transform(Xraw)

res_std = ols(y_score, Xstd, FEATURES)

print('\n模型整体:')
print('  样本量 N = %s' % format(res_std['n'], ','))
print('  R² = %.4f   调整 R² = %.4f   F = %.1f (p = %s)'
      % (res_std['r2'], res_std['adj_r2'], res_std['f'], fmt_p(res_std['f_p'])))
print('  残差标准误 = %.4f' % res_std['sigma'])
print('\n标准化回归系数（可直接比较影响大小）:')
print('  %-16s %9s %9s %9s %12s %6s' % ('变量', '标准化β', '标准误', 't 值', 'p 值', '显著性'))
print('  ' + '-' * 72)

rows_std = []
order = np.argsort(-np.abs(res_std['beta'][1:]))     # 按 |β| 排序
for idx in order:
    i = idx + 1
    print('  %-16s %9.4f %9.4f %9.1f %12s %6s'
          % (LABEL[FEATURES[idx]], res_std['beta'][i], res_std['se'][i],
             res_std['t'][i], fmt_p(res_std['p'][i]), stars(res_std['p'][i])))
    rows_std.append({
        '变量': LABEL[FEATURES[idx]], 'code': FEATURES[idx],
        '标准化β': res_std['beta'][i], '标准误': res_std['se'][i],
        't': res_std['t'][i], 'p': res_std['p'][i], '显著性': stars(res_std['p'][i]),
        'CI下限': res_std['ci_lo'][i], 'CI上限': res_std['ci_hi'][i],
    })

print('\n共线性诊断（VIF，>10 表示严重共线）:')
vifs = vif_table(Xraw, FEATURES)
for nm, r2, v in sorted(vifs, key=lambda x: -x[2]):
    flag = ' ⚠' if v > 10 else (' △' if v > 5 else '')
    print('  %-16s R²=%.3f  VIF=%.2f%s' % (LABEL[nm], r2, v, flag))

# 非标准化模型（用于业务量纲解读）
res_raw = ols(y_score, Xraw, FEATURES)
print('\n非标准化系数（业务量纲）:')
for i, nm in enumerate(FEATURES, start=1):
    print('  %-16s β=%+.5f   （%s每增加1单位，评分变化 %+.5f 分）'
          % (LABEL[nm], res_raw['beta'][i], LABEL[nm], res_raw['beta'][i]))


# =============================================================================
# C. 逻辑回归（Logit）—— 优势比 OR
# =============================================================================
sec('C. 二元逻辑回归：因变量 = 是否差评（1-2 星）')

y_bad = model_df['is_bad_review'].values.astype(int)

lr = LogisticRegression(penalty=None, solver='lbfgs', max_iter=5000)
lr.fit(Xstd, y_bad)
coef = lr.coef_[0]
p_hat = lr.predict_proba(Xstd)[:, 1]

# Fisher 信息矩阵求标准误
W = p_hat * (1 - p_hat)
H = Xstd.T @ (Xstd * W[:, None])
cov = np.linalg.pinv(H)
se_lr = np.sqrt(np.diag(cov))
z = coef / se_lr
p_lr = 2 * (1 - stats.norm.cdf(np.abs(z)))
or_ = np.exp(coef)
ci_lo = np.exp(coef - 1.96 * se_lr)
ci_hi = np.exp(coef + 1.96 * se_lr)

# 拟合优度
ll_full = np.sum(y_bad * np.log(p_hat + 1e-12) + (1 - y_bad) * np.log(1 - p_hat + 1e-12))
p0 = y_bad.mean()
ll_null = np.sum(y_bad * np.log(p0) + (1 - y_bad) * np.log(1 - p0))
mcf = 1 - ll_full / ll_null
lr_stat = 2 * (ll_full - ll_null)
df_lr = len(coef)
print('\n模型整体: N = %s   McFadden R² = %.4f   LR 统计量 = %.1f (df=%d, p≈0)'
      % (format(len(y_bad), ','), mcf, lr_stat, df_lr))
print('  基准差评率 p0 = %.2f%%' % (100 * p0))

print('\n优势比 OR（>1 表示增加差评几率，<1 表示降低）:')
print('  %-16s %9s %8s %-20s %12s %6s' % ('变量', 'OR', '标准误', '95% 置信区间', 'p 值', '显著性'))
print('  ' + '-' * 84)

rows_or = []
for idx in np.argsort(-np.abs(coef)):
    print('  %-16s %9.4f %8.4f [%7.4f, %7.4f] %12s %6s'
          % (LABEL[FEATURES[idx]], or_[idx], se_lr[idx],
             ci_lo[idx], ci_hi[idx], fmt_p(p_lr[idx]), stars(p_lr[idx])))
    rows_or.append({
        '变量': LABEL[FEATURES[idx]], 'code': FEATURES[idx],
        'OR': or_[idx], '标准误': se_lr[idx],
        'CI下限': ci_lo[idx], 'CI上限': ci_hi[idx],
        'p': p_lr[idx], '显著性': stars(p_lr[idx]),
    })

# 边际效应（在均值处）
print('\n在均值处的边际效应（对差评概率的百分点影响）:')
for idx in np.argsort(-np.abs(coef))[:5]:
    me = coef[idx] * p0 * (1 - p0)     # 标准化变量，近似
    print('  %-16s 每增加 1 个标准差，差评概率变化 %+.2f 个百分点'
          % (LABEL[FEATURES[idx]], 100 * me))


# =============================================================================
# D. 随机森林特征重要性
# =============================================================================
sec('D. 随机森林特征重要性（非参数稳健性对照）')

rf = RandomForestClassifier(n_estimators=300, max_depth=12, min_samples_leaf=50,
                            random_state=42, n_jobs=-1)
rf.fit(Xraw, y_bad)
imp = pd.DataFrame({'变量': [LABEL[f] for f in FEATURES],
                    'code': FEATURES, '重要性': rf.feature_importances_}
                   ).sort_values('重要性', ascending=False)
imp['占比%'] = 100 * imp['重要性']
print('  %-18s %10s %8s' % ('变量', '重要性', '占比%'))
print('  ' + '-' * 40)
for _, r in imp.iterrows():
    print('  %-18s %10.4f %8.2f' % (r['变量'], r['重要性'], r['占比%']))

cv = cross_val_score(rf, Xraw, y_bad, cv=5, scoring='roc_auc', n_jobs=-1)
print('\n  5 折交叉验证 AUC = %.4f ± %.4f' % (cv.mean(), cv.std()))


# =============================================================================
# E. 交互效应检验
# =============================================================================
sec('E. 交互效应检验')

# 迟到 × 距离
d = model_df.dropna(subset=['is_late', 'log_distance', 'review_score']).copy()
d['late_x_dist'] = d['is_late'] * d['log_distance']
X_int = d[['is_late', 'log_distance', 'late_x_dist']].values.astype(float)
r_int = ols(d['review_score'].values, X_int, ['is_late', 'log_distance', 'late_x_dist'])
print('\n模型：评分 ~ 迟到 + 距离 + 迟到×距离')
for i, nm in enumerate(r_int['names']):
    print('  %-14s β=%+.5f  t=%6.2f  p=%s  %s'
          % (nm, r_int['beta'][i], r_int['t'][i], fmt_p(r_int['p'][i]),
             stars(r_int['p'][i])))
print('  R² = %.4f' % r_int['r2'])
print('  解读：若交互项显著且为正，说明距离会放大迟到的负面影响（或反之）。')

# 迟到 × 超期发货
d2 = model_df.dropna(subset=['is_late', 'is_overdue_ship', 'review_score']).copy()
d2['late_x_over'] = d2['is_late'] * d2['is_overdue_ship']
X_int2 = d2[['is_late', 'is_overdue_ship', 'late_x_over']].values.astype(float)
r_int2 = ols(d2['review_score'].values, X_int2, ['is_late', 'is_overdue_ship', 'late_x_over'])
print('\n模型：评分 ~ 迟到 + 超期发货 + 迟到×超期')
for i, nm in enumerate(r_int2['names']):
    print('  %-14s β=%+.5f  t=%6.2f  p=%s  %s'
          % (nm, r_int2['beta'][i], r_int2['t'][i], fmt_p(r_int2['p'][i]),
             stars(r_int2['p'][i])))
print('  R² = %.4f' % r_int2['r2'])


# =============================================================================
# F. 稳健性检验
# =============================================================================
sec('F. 稳健性检验：分年度子样本')

print('  %-10s %9s %11s %11s %11s %9s'
      % ('子样本', 'N', '迟到β(标准化)', '距离β', '超期β', 'R²'))
print('  ' + '-' * 68)
for yr in sorted(model_df['yr'].dropna().unique()):
    sub = model_df[model_df['yr'] == yr]
    if len(sub) < 500:
        continue
    ys = sub['review_score'].values.astype(float)
    Xs = StandardScaler().fit_transform(sub[FEATURES].values.astype(float))
    rr = ols(ys, Xs, FEATURES)
    b_late = rr['beta'][1 + FEATURES.index('is_late')]
    b_dist = rr['beta'][1 + FEATURES.index('log_distance')]
    b_over = rr['beta'][1 + FEATURES.index('is_overdue_ship')]
    print('  %-10d %9s %11.4f %11.4f %11.4f %9.4f'
          % (int(yr), format(len(sub), ','), b_late, b_dist, b_over, rr['r2']))

print('\n  说明：若各年系数符号一致且量级接近，说明结论稳健，不是某一年偶然。')


# =============================================================================
# G. 可视化
# =============================================================================
sec('G. 生成图表')

# 图1：标准化系数 vs OR 重要性对比
fig, axes = plt.subplots(1, 3, figsize=(17, 5.6))

sd = pd.DataFrame(rows_std).sort_values('标准化β')
colors = ['#c0392b' if v < 0 else '#2980b9' for v in sd['标准化β']]
axes[0].barh(sd['变量'], sd['标准化β'], color=colors)
axes[0].axvline(0, color='#333', lw=.9)
axes[0].set_title('OLS 标准化系数 β\n（因变量：评分）', fontsize=11)
axes[0].set_xlabel('β')
for i, v in enumerate(sd['标准化β']):
    axes[0].text(v + (0.004 if v >= 0 else -0.004), i, '%.3f' % v,
                 va='center', ha='left' if v >= 0 else 'right', fontsize=7.5)
axes[0].margins(x=0.22)

od = pd.DataFrame(rows_or).copy()
od['logOR'] = np.log(od['OR'])
od = od.sort_values('logOR')
colors = ['#c0392b' if v > 0 else '#2980b9' for v in od['logOR']]
axes[1].barh(od['变量'], od['logOR'], color=colors)
axes[1].axvline(0, color='#333', lw=.9)
axes[1].set_title('Logit 系数 ln(OR)\n（因变量：是否差评）', fontsize=11)
axes[1].set_xlabel('ln(OR)')
for i, (v, o) in enumerate(zip(od['logOR'], od['OR'])):
    axes[1].text(v + (0.02 if v >= 0 else -0.02), i, 'OR=%.2f' % o,
                 va='center', ha='left' if v >= 0 else 'right', fontsize=7.5)
axes[1].margins(x=0.28)

im = imp.sort_values('重要性')
axes[2].barh(im['变量'], im['重要性'], color='#8e44ad')
axes[2].set_title('随机森林特征重要性\n（AUC=%.3f）' % cv.mean(), fontsize=11)
axes[2].set_xlabel('重要性')
for i, v in enumerate(im['重要性']):
    axes[2].text(v, i, ' %.3f' % v, va='center', fontsize=7.5)
axes[2].margins(x=0.18)

plt.tight_layout()
p1 = os.path.join(OUT, 'fig6_多元模型系数对比.png')
plt.savefig(p1, bbox_inches='tight')
plt.close()
print('  已保存: %s' % os.path.basename(p1))

# 图2：距离与评分/差评率的分箱关系
fig, axes = plt.subplots(1, 2, figsize=(14, 5))
d3 = model_df.dropna(subset=['distance_km', 'review_score'])
d3['dist_bin'] = pd.qcut(d3['distance_km'], 10, duplicates='drop')
gb = d3.groupby('dist_bin', observed=True).agg(
    距离=('distance_km', 'median'),
    评分=('review_score', 'mean'),
    差评率=('is_bad_review', lambda s: 100 * s.mean()),
    样本=('review_score', 'size')).reset_index()

axes[0].plot(gb['距离'], gb['评分'], marker='o', color='#2c7fb8', lw=2)
axes[0].set_xlabel('客户—卖家直线距离 (km)')
axes[0].set_ylabel('平均评分')
axes[0].set_title('距离与评分（十分位分箱）', fontsize=11)
axes[0].grid(alpha=.35)

axes[1].plot(gb['距离'], gb['差评率'], marker='s', color='#c0392b', lw=2)
axes[1].set_xlabel('客户—卖家直线距离 (km)')
axes[1].set_ylabel('差评率 (%)')
axes[1].set_title('距离与差评率（十分位分箱）', fontsize=11)
axes[1].grid(alpha=.35)

plt.tight_layout()
p2 = os.path.join(OUT, 'fig7_距离非线性关系.png')
plt.savefig(p2, bbox_inches='tight')
plt.close()
print('  已保存: %s' % os.path.basename(p2))

# 图3：模型拟合优度对比
fig, ax = plt.subplots(figsize=(8, 4.4))
models = ['仅时间变量\n(备货+运输)', '时间+违约\n(+迟到+超期)', '全模型\n(+距离+商品)']
X1 = model_df[['prep_days', 'ship_days']].values.astype(float)
X2 = model_df[['prep_days', 'ship_days', 'is_late', 'is_overdue_ship']].values.astype(float)
r1 = ols(y_score, StandardScaler().fit_transform(X1), ['a', 'b'])
r2 = ols(y_score, StandardScaler().fit_transform(X2), ['a', 'b', 'c', 'd'])
r3 = ols(y_score, Xstd, FEATURES)
r2s = [r1['r2'], r2['r2'], r3['r2']]
bars = ax.bar(models, r2s, color=['#95a5a6', '#5dade2', '#2e86c1'])
for b, v in zip(bars, r2s):
    ax.text(b.get_x() + b.get_width() / 2, v + 0.004, 'R²=%.4f' % v,
            ha='center', fontsize=9)
ax.set_ylabel('R²（解释力）')
ax.set_ylim(0, max(r2s) * 1.35)
ax.set_title('逐步加入变量后的解释力提升', fontsize=12)
plt.tight_layout()
p3 = os.path.join(OUT, 'fig8_模型解释力对比.png')
plt.savefig(p3, bbox_inches='tight')
plt.close()
print('  已保存: %s' % os.path.basename(p3))


# =============================================================================
# H. 导出
# =============================================================================
sec('H. 导出结果')

pd.DataFrame(rows_std).to_csv(os.path.join(OUT, 'ols_standardized.csv'),
                              index=False, encoding='utf-8-sig')
pd.DataFrame(rows_or).to_csv(os.path.join(OUT, 'logit_odds_ratios.csv'),
                             index=False, encoding='utf-8-sig')
imp.to_csv(os.path.join(OUT, 'rf_importance.csv'), index=False, encoding='utf-8-sig')
pd.DataFrame([{'模型': m, 'R2': v} for m, v in zip(models, r2s)]).to_csv(
    os.path.join(OUT, 'model_r2.csv'), index=False, encoding='utf-8-sig')
pd.DataFrame([{'变量': LABEL[n], 'VIF': v} for n, _, v in vifs]).to_csv(
    os.path.join(OUT, 'vif.csv'), index=False, encoding='utf-8-sig')
gb.to_csv(os.path.join(OUT, 'distance_bins.csv'), index=False, encoding='utf-8-sig')
print('  已导出 6 个 CSV')

sec('建模完成')
print('样本: %s 条   特征: %d 个' % (format(len(model_df), ','), len(FEATURES)))
print('OLS 调整 R² = %.4f    Logit 伪 R² = %.4f    RF AUC = %.4f'
      % (res_std['adj_r2'], mcf, cv.mean()))
