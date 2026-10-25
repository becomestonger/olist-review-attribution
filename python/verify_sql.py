"""逐条验证 SQL 文件里的每条语句能否执行。"""
import csv
import io
import os
import re
import subprocess
import sys

PSQL = r'C:\Users\28223\pgsql\bin\psql.exe'
SQL_FILE = sys.argv[1] if len(sys.argv) > 1 else \
    r'C:\Users\28223\Desktop\巴西olist数据分析项目\sql\06_factor_analysis.sql'

env = os.environ.copy()
env['PGPASSWORD'] = '123456'
env['PGCLIENTENCODING'] = 'UTF8'


def run(sql):
    cmd = [PSQL, '-h', '127.0.0.1', '-p', '5432', '-U', 'postgres',
           '-d', 'analysis', '-X', '-q', '-A', '-F', '\t', '--csv',
           '-v', 'ON_ERROR_STOP=1', '-c', sql]
    p = subprocess.run(cmd, capture_output=True, env=env, timeout=600)
    out = p.stdout.decode('utf-8', 'replace')
    err = p.stderr.decode('utf-8', 'replace')
    rows = list(csv.reader(io.StringIO(out))) if out.strip() else []
    return p.returncode, rows, err


def split_statements(text):
    lines = []
    for line in text.splitlines():
        if line.strip().startswith('--'):
            continue
        lines.append(line)
    body = '\n'.join(lines)
    return [s.strip() for s in body.split(';') if s.strip()]


with open(SQL_FILE, encoding='utf-8') as f:
    content = f.read()

stmts = split_statements(content)
print('文件: %s' % os.path.basename(SQL_FILE))
print('语句数: %d\n' % len(stmts))

ok = fail = 0
for i, s in enumerate(stmts, 1):
    label = ' '.join(s.split())[:64]
    rc, rows, err = run(s)
    if rc == 0:
        ok += 1
        n = max(0, len(rows) - 1)
        print('[%2d] OK   %4d 行  %s' % (i, n, label))
        if len(rows) > 1:
            print('       %s' % (' | '.join(rows[1][:5]))[:100])
    else:
        fail += 1
        msg = [l for l in err.splitlines()
               if l.strip().startswith(('ERROR', '错误', 'FATAL'))]
        print('[%2d] FAIL        %s' % (i, label))
        print('       %s' % (msg[0] if msg else err.strip()[:180]))

print('\n结果: 成功 %d / 失败 %d' % (ok, fail))
