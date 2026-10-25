"""
连接测试 —— 确认 Python 能连上 PostgreSQL 并读到 Olist 数据。
用法（在项目目录下）：
    .venv\\Scripts\\python.exe analysis\\connect_test.py
"""
import sys

try:
    import pandas as pd
    from sqlalchemy import create_engine, text
except ImportError as e:
    print('缺少依赖:', e)
    print('请先执行: .venv\\Scripts\\python.exe -m pip install pandas sqlalchemy psycopg2-binary')
    sys.exit(1)

# ---------------------------------------------------------------------------
# 连接信息
#   密码 123456 里的特殊字符需要 URL 编码；这里都是普通字符，可直接写
# ---------------------------------------------------------------------------
DB = dict(host='127.0.0.1', port=5432, db='analysis', user='postgres', pwd='123456')

URL = 'postgresql+psycopg2://{user}:{pwd}@{host}:{port}/{db}'.format(**DB)

print('正在连接 %s:%s/%s ...' % (DB['host'], DB['port'], DB['db']))
engine = create_engine(URL, pool_pre_ping=True)

try:
    with engine.connect() as conn:
        ver = conn.execute(text('SELECT version()')).scalar()
        print('连接成功 ✓')
        print('  ' + ver.split(',')[0])
except Exception as e:
    print('连接失败 ✗:', e)
    sys.exit(2)

# ---------------------------------------------------------------------------
# 确认数据在
# ---------------------------------------------------------------------------
with engine.connect() as conn:
    tables = pd.read_sql(
        "SELECT table_name, table_type FROM information_schema.tables "
        "WHERE table_schema = 'olist' ORDER BY table_type, table_name", conn)
    print('\nolist 下的对象（%d 个）:' % len(tables))
    for _, r in tables.iterrows():
        kind = '视图' if 'VIEW' in r['table_type'] else '表  '
        print('  [%s] %s' % (kind, r['table_name']))

# ---------------------------------------------------------------------------
# 跑一个真实查询，验证端到端可用
# ---------------------------------------------------------------------------
sql = """
SELECT customer_state            AS 州,
       count(*)                   AS 订单数,
       round(sum(order_total)::numeric, 2) AS 销售额
FROM olist.v_orders_full
GROUP BY 1
ORDER BY 销售额 DESC
LIMIT 5
"""
df = pd.read_sql(sql, engine)
print('\n各州销售额 Top 5:')
print(df.to_string(index=False))

print('\n环境就绪，可以开始分析了。')
