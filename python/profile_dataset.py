"""
Olist 数据集完整剖析：表结构、行数、字段类型、缺失情况、关键分布。
"""
import csv
import io
import os
import subprocess

PSQL = r'C:\Users\28223\pgsql\bin\psql.exe'
env = os.environ.copy()
env['PGPASSWORD'] = '123456'
env['PGCLIENTENCODING'] = 'UTF8'


def q(sql):
    cmd = [PSQL, '-h', '127.0.0.1', '-p', '5432', '-U', 'postgres',
           '-d', 'analysis', '-X', '-q', '-A', '-F', '\t', '--csv',
           '-v', 'ON_ERROR_STOP=1', '-c', sql]
    p = subprocess.run(cmd, capture_output=True, env=env, timeout=300)
    if p.returncode != 0:
        raise RuntimeError(p.stderr.decode('utf-8', 'replace')[-400:])
    out = p.stdout.decode('utf-8', 'replace')
    if not out.strip():
        return [], []
    rows = list(csv.reader(io.StringIO(out)))
    return rows[0], rows[1:]


print('=' * 78)
print(' 一、数据表清单与规模')
print('=' * 78)

tables = ['customers', 'geolocation', 'orders', 'order_items', 'order_payments',
          'order_reviews', 'products', 'sellers', 'product_category_translation']
views = ['v_orders_full', 'v_product_sales', 'v_seller_performance']

total = 0
for t in tables:
    _, r = q('SELECT count(*) FROM olist.%s' % t)
    n = int(r[0][0])
    total += n
    print('  %-32s %10s 行' % (t, format(n, ',')))
print('  %-32s %10s 行' % ('【总计】', format(total, ',')))
print()
for v in views:
    _, r = q('SELECT count(*) FROM olist.%s' % v)
    print('  %-32s %10s 行  (视图)' % (v, format(int(r[0][0]), ',')))

print()
print('=' * 78)
print(' 二、每张表的字段结构')
print('=' * 78)

for t in tables:
    cols, rows = q("""
        SELECT column_name, data_type,
               character_maximum_length, numeric_precision, numeric_scale,
               is_nullable
        FROM information_schema.columns
        WHERE table_schema='olist' AND table_name='%s'
        ORDER BY ordinal_position
    """ % t)
    _, cnt = q('SELECT count(*) FROM olist.%s' % t)
    print('\n▶ %s  (%s 行)' % (t, format(int(cnt[0][0]), ',')))
    print('  %-34s %-16s %s' % ('字段名', '类型', '可空'))
    print('  ' + '-' * 66)
    for r in rows:
        name, dtype, maxlen, prec, scale, nullable = r
        if maxlen:
            dtype = '%s(%s)' % ('CHAR' if 'character' in dtype else 'VARCHAR', maxlen)
        elif prec and scale:
            dtype = 'NUMERIC(%s,%s)' % (prec, scale)
        print('  %-34s %-16s %s' % (name, dtype, 'NULL' if nullable == 'YES' else 'NOT NULL'))

print()
print('=' * 78)
print(' 三、缺失值情况（NULL 占比）')
print('=' * 78)

null_checks = [
    ('orders', 'order_approved_at', '审核时间'),
    ('orders', 'order_delivered_carrier_date', '交给物流时间'),
    ('orders', 'order_delivered_customer_date', '客户签收时间'),
    ('orders', 'order_estimated_delivery_date', '预估送达时间'),
    ('products', 'product_category_name', '商品品类'),
    ('order_reviews', 'review_comment_title', '评论标题'),
    ('order_reviews', 'review_comment_message', '评论内容'),
    ('customers', 'customer_city', '客户城市'),
]
for t, c, label in null_checks:
    sql = ("SELECT count(*) AS total, count(*) FILTER (WHERE %s IS NULL) AS nulls "
           "FROM olist.%s" % (c, t))
    _, r = q(sql)
    tot, nul = int(r[0][0]), int(r[0][1])
    pct = 100.0 * nul / tot if tot else 0
    print('  %-22s.%-30s %8s / %-8s = %5.1f%% 空' % (t, c, format(nul, ','), format(tot, ','), pct))

print()
print('=' * 78)
print(' 四、关键维度分布')
print('=' * 78)

def dist(title, sql, maxn=12):
    print('\n▶ %s' % title)
    cols, rows = q(sql)
    if not cols:
        print('  （无数据）')
        return
    w = max(len(str(c)) for c in cols)
    for i, r in enumerate(rows[:maxn]):
        if i == 0:
            print('  ' + '  '.join(str(c).ljust(14) for c in cols))
            print('  ' + '-' * (16 * len(cols)))
        print('  ' + '  '.join(str(v).ljust(14) for v in r))

dist('订单状态分布', """
    SELECT order_status AS 状态, count(*) AS 数量,
           round(100.0*count(*)/sum(count(*)) OVER (), 2) AS 占比 FROM olist.orders
    GROUP BY 1 ORDER BY 2 DESC
""")

dist('支付方式分布', """
    SELECT payment_type AS 支付方式, count(*) AS 笔数,
           round(sum(payment_value)::numeric,0) AS 总金额 FROM olist.order_payments
    GROUP BY 1 ORDER BY 2 DESC
""")

dist('评分分布', """
    SELECT review_score AS 星级, count(*) AS 数量,
           round(100.0*count(*)/sum(count(*)) OVER (), 2) AS 占比
    FROM olist.order_reviews GROUP BY 1 ORDER BY 1
""")

dist('客户所在州 Top 10', """
    SELECT customer_state AS 州, count(*) AS 客户数 FROM olist.customers
    GROUP BY 1 ORDER BY 2 DESC LIMIT 10
""")

dist('卖家所在州 Top 10', """
    SELECT seller_state AS 州, count(*) AS 卖家数 FROM olist.sellers
    GROUP BY 1 ORDER BY 2 DESC LIMIT 10
""")

dist('商品品类 Top 10（按商品数）', """
    SELECT p.product_category_name AS 品类_葡语,
           t.product_category_name_english AS 品类_英语,
           count(*) AS 商品数
    FROM olist.products p
    LEFT JOIN olist.product_category_translation t
           ON t.product_category_name = p.product_category_name
    GROUP BY 1,2 ORDER BY 3 DESC LIMIT 10
""")

print()
print('=' * 78)
print(' 五、数值字段统计')
print('=' * 78)

stats = [
    ('订单金额', 'olist.v_orders_full', 'order_total'),
    ('商品单价 price', 'olist.order_items', 'price'),
    ('运费 freight', 'olist.order_items', 'freight_value'),
    ('支付金额', 'olist.order_payments', 'payment_value'),
    ('商品重量(g)', 'olist.products', 'product_weight_g'),
]
print('  %-16s %12s %12s %12s %12s' % ('字段', '平均', '最小', '最大', '非空数'))
print('  ' + '-' * 68)
for label, tbl, col in stats:
    _, r = q("""
        SELECT round(avg(%s)::numeric,2), min(%s), max(%s), count(%s)
        FROM %s
    """ % (col, col, col, col, tbl))
    avg, mn, mx, n = r[0]
    print('  %-16s %12s %12s %12s %12s' % (label, avg, mn, mx, format(int(n), ',')))

print()
print('=' * 78)
print(' 六、时间跨度')
print('=' * 78)
_, r = q("""
    SELECT min(order_purchase_timestamp)::date, max(order_purchase_timestamp)::date,
           count(DISTINCT date_trunc('month', order_purchase_timestamp))
    FROM olist.orders
""")
print('  下单时间: %s  至  %s   共 %s 个月' % (r[0][0], r[0][1], r[0][2]))

print()
print('=' * 78)
print(' 七、表间关系（外键）')
print('=' * 78)
_, r = q("""
    SELECT conrelid::regclass::text AS 子表,
           a.attname AS 外键字段,
           confrelid::regclass::text AS 父表
    FROM pg_constraint c
    JOIN pg_attribute a ON a.attrelid = c.conrelid AND a.attnum = ANY(c.conkey)
    WHERE c.contype='f' AND c.connamespace='olist'::regnamespace
    ORDER BY 1,2
""")
for r_ in r:
    print('  %-40s %-24s -> %s' % (r_[0].replace('olist.', ''), r_[1], r_[2].replace('olist.', '')))
