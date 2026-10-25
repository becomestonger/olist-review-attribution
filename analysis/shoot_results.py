"""
把 06_factor_analysis.sql 里所有查询执行一遍，把结果渲染成 HTML 再截图成 PNG。

输出：<项目>/results/*.png   每条查询一张图
"""
import csv
import html
import io
import os
import re
import subprocess
import sys
import time

PSQL = r'C:\Users\28223\pgsql\bin\psql.exe'
EDGE = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
BASE = r'C:\Users\28223\Desktop\巴西olist数据分析项目'
SQL_FILE = os.path.join(BASE, 'sql', '06_factor_analysis.sql')
OUT_DIR = os.path.join(BASE, 'results')
TMP_DIR = os.path.join(os.environ['TEMP'], 'olist_shot')

env = os.environ.copy()
env['PGPASSWORD'] = '123456'
env['PGCLIENTENCODING'] = 'UTF8'

os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(TMP_DIR, exist_ok=True)


# ---------------------------------------------------------------------------
# 1. 解析 SQL 文件：提取 部分名 / 编号 / 标题 / SQL 语句
# ---------------------------------------------------------------------------
def parse_sql(path):
    with open(path, encoding='utf-8') as f:
        lines = f.readlines()

    items = []
    section = ''
    title = ''
    buf = []
    pending_comment = []

    for raw in lines:
        line = raw.rstrip('\n')

        m = re.search(r'第\s*(\d+)\s*部分\s+(.+?)\s*$', line)
        if m and line.strip().startswith('--'):
            section = '第 %s 部分  %s' % (m.group(1), m.group(2).strip())
            title = ''
            pending_comment = []
            continue

        stripped = line.strip()

        # 形如 "-- 0.1 数据规模（...）"
        m2 = re.match(r'^--\s*(\d+\.\d+)\s+(.+)$', stripped)
        if m2 and not buf:
            num = m2.group(1)
            t = m2.group(2).strip()
            # 去掉后面的 ★ 标记等，保留原样
            title = '%s  %s' % (num, t)
            pending_comment = []
            continue

        # 普通说明注释：作为副标题累积
        if stripped.startswith('--') and not buf:
            body = stripped.lstrip('-').strip()
            if body and not re.fullmatch(r'[█═\-=\s]+', body):
                pending_comment.append(body)
            continue

        if not stripped:
            continue

        buf.append(line)

        # 语句以分号结尾 -> 收集完一条
        if stripped.endswith(';'):
            sql = '\n'.join(buf).strip().rstrip(';').strip()
            if sql:
                items.append({
                    'section': section,
                    'title': title or ('查询 %d' % (len(items) + 1)),
                    'note': ' '.join(pending_comment)[:160],
                    'sql': sql,
                })
            buf = []
            title = ''
            pending_comment = []

    return items


# ---------------------------------------------------------------------------
# 2. 执行查询
# ---------------------------------------------------------------------------
def run_sql(sql):
    cmd = [PSQL, '-h', '127.0.0.1', '-p', '5432', '-U', 'postgres',
           '-d', 'analysis', '-X', '-q', '-A', '-F', '\t', '--csv',
           '-v', 'ON_ERROR_STOP=1', '-c', sql + ';']
    t0 = time.time()
    p = subprocess.run(cmd, capture_output=True, env=env, timeout=900)
    cost = (time.time() - t0) * 1000
    if p.returncode != 0:
        raise RuntimeError(p.stderr.decode('utf-8', 'replace')[-300:])
    out = p.stdout.decode('utf-8', 'replace')
    rows = list(csv.reader(io.StringIO(out))) if out.strip() else []
    cols = rows[0] if rows else []
    data = rows[1:] if rows else []
    return cols, data, cost


# ---------------------------------------------------------------------------
# 3. 渲染 HTML
# ---------------------------------------------------------------------------
def esc(s):
    return html.escape('' if s is None else str(s))


def is_num(v):
    return bool(re.fullmatch(r'-?\d+(\.\d+)?', v or ''))


def render_html(item, cols, rows, cost):
    head = ''.join('<th>%s</th>' % esc(c) for c in cols)
    body = []
    for r in rows:
        tds = []
        for i in range(len(cols)):
            v = r[i] if i < len(r) else ''
            if v == '':
                tds.append('<td class="n">NULL</td>')
            elif is_num(v):
                tds.append('<td class="num">%s</td>' % esc(v))
            else:
                tds.append('<td>%s</td>' % esc(v))
        body.append('<tr>%s</tr>' % ''.join(tds))

    if not cols:
        grid = '<div class="empty">(无结果)</div>'
    else:
        grid = ('<table><thead><tr>%s</tr></thead><tbody>%s</tbody></table>'
                % (head, ''.join(body)))

    note = ('<div class="note">%s</div>' % esc(item['note'])) if item['note'] else ''

    # 标题首部是编号，抽出来做徽标，正文只留描述，避免 "6.1 6.1 ..." 重复
    full = item['title']
    num = full.split()[0] if full.split() else ''
    if re.fullmatch(r'\d+\.\d+', num):
        heading = full[len(num):].strip()
    else:
        num, heading = '', full

    return """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<style>
*{box-sizing:border-box}
body{margin:0;padding:18px 24px;background:#fff;
     font-family:"Segoe UI","Microsoft YaHei",Arial,sans-serif;color:#1f2937}
.top{display:flex;align-items:center;gap:10px;margin-bottom:4px;flex-wrap:wrap}
.sec{background:#1e3a5f;color:#fff;font-size:12px;padding:3px 10px;border-radius:4px}
.no{font-size:19px;font-weight:700;color:#1e3a5f}
h1{font-size:14px;font-weight:600;margin:0;color:#374151}
.note{font-size:12px;color:#6b7280;margin:4px 0 0;line-height:1.5}
.meta{font-size:11.5px;color:#9ca3af;margin:7px 0 10px;padding-bottom:7px}
.hr{height:1px;background:#e5e7eb;margin:0 0 10px;max-width:760px}
table{border-collapse:collapse;font-size:12.5px;white-space:nowrap}
th{position:sticky;top:0;background:#f1f5f9;text-align:left;padding:7px 11px;
   border-bottom:2px solid #cbd5e1;border-right:1px solid #e2e8f0;font-weight:600}
td{padding:5px 11px;border-bottom:1px solid #f1f5f9;border-right:1px solid #f8fafc;
   max-width:420px;overflow:hidden;text-overflow:ellipsis}
tbody tr:nth-child(even){background:#fafbfc}
td.num{text-align:right;font-variant-numeric:tabular-nums}
td.n{color:#cbd5e1;font-style:italic}
.empty{padding:26px;color:#9ca3af;text-align:center}
</style></head><body>
<div class="top">
  <span class="sec">%s</span>
  <span class="no">%s</span>
  <h1>%s</h1>
</div>
%s
<div class="meta">共 %d 行 · 耗时 %d ms · 数据源 PostgreSQL 18 / analysis / olist</div>
<div class="hr"></div>
%s
</body></html>""" % (
        esc(item['section']), esc(num), esc(heading),
        note, len(rows), int(cost), grid)


# ---------------------------------------------------------------------------
# 4. 截图
# ---------------------------------------------------------------------------
def shot(html_path, png_path, cols, nrows):
    width = min(320 + max(1, len(cols)) * 175, 2400)
    height = min(240 + min(nrows, 50) * 31, 1700)
    height = max(height, 340)
    prof = os.path.join(TMP_DIR, 'edgeprof')
    subprocess.run([
        EDGE, '--headless=new', '--disable-gpu', '--no-sandbox',
        '--hide-scrollbars', '--force-device-scale-factor=1.5',
        '--virtual-time-budget=3000',
        '--user-data-dir=%s' % prof,
        '--window-size=%d,%d' % (width, height),
        '--screenshot=%s' % png_path,
        'file:///' + html_path.replace('\\', '/'),
    ], capture_output=True, timeout=120)
    autocrop(png_path)


def autocrop(png_path, pad=20):
    """裁掉图片四周的空白，让内容贴边。"""
    try:
        from PIL import Image, ImageChops
    except ImportError:
        return
    try:
        img = Image.open(png_path).convert('RGB')
        bg = Image.new('RGB', img.size, (255, 255, 255))
        bbox = ImageChops.difference(img, bg).getbbox()
        if not bbox:
            return
        l, t, r, b = bbox
        l = max(0, l - pad)
        t = max(0, t - pad)
        r = min(img.width, r + pad)
        b = min(img.height, b + pad)
        if (r - l) < 120 or (b - t) < 60:
            return
        if (r - l, b - t) != img.size:
            img.crop((l, t, r, b)).save(png_path)
    except Exception:
        pass


def safe_name(s):
    s = re.sub(r'[\\/:*?"<>|]', '_', s)
    s = re.sub(r'\s+', '', s)
    return s[:48]


def main():
    items = parse_sql(SQL_FILE)
    print('解析到 %d 条查询\n' % len(items))

    ok = 0
    for i, item in enumerate(items, 1):
        try:
            cols, rows, cost = run_sql(item['sql'])
        except Exception as e:
            print('[%2d] 执行失败 %s -> %s' % (i, item['title'][:30], str(e)[:120]))
            continue

        doc = render_html(item, cols, rows, cost)
        hp = os.path.join(TMP_DIR, 'q%02d.html' % i)
        with open(hp, 'w', encoding='utf-8') as f:
            f.write(doc)

        name = '%02d_%s.png' % (i, safe_name(item['title']))
        png = os.path.join(OUT_DIR, name)
        shot(hp, png, cols, len(rows))

        size = os.path.getsize(png) if os.path.exists(png) else 0
        print('[%2d] %-46s %2d 行  ->  %s (%d KB)'
              % (i, item['title'][:44], len(rows), name, size // 1024))
        ok += 1

    print('\n完成: %d / %d 张截图' % (ok, len(items)))
    print('输出目录: %s' % OUT_DIR)

    # 生成索引页
    idx = ['<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">',
           '<title>Olist 差评归因分析 - 结果索引</title>',
           '<style>body{font-family:"Segoe UI","Microsoft YaHei";margin:30px;'
           'background:#f5f6f8;color:#1f2937}',
           'h1{font-size:20px} h2{font-size:15px;color:#1e3a5f;margin-top:26px}',
           '.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:14px}',
           '.card{background:#fff;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden}',
           '.card img{width:100%;display:block;border-bottom:1px solid #f1f5f9}',
           '.card .t{padding:8px 10px;font-size:12px;color:#374151}',
           'a{text-decoration:none;color:inherit}</style></head><body>',
           '<h1>Olist 巴西电商 —— 差评归因分析结果</h1>',
           '<p style="font-size:13px;color:#6b7280">数据库 PostgreSQL 18 · analysis · schema olist · '
           '共 %d 张结果图</p>' % len(items)]

    cur = None
    for i, item in enumerate(items, 1):
        if item['section'] != cur:
            if cur is not None:
                idx.append('</div>')
            cur = item['section']
            idx.append('<h2>%s</h2><div class="grid">' % esc(cur))
        name = '%02d_%s.png' % (i, safe_name(item['title']))
        idx.append('<a href="%s"><div class="card"><img src="%s"><div class="t">%s</div></div></a>'
                   % (name, name, esc(item['title'])))
    if cur is not None:
        idx.append('</div>')
    idx.append('</body></html>')

    ip = os.path.join(OUT_DIR, 'index.html')
    with open(ip, 'w', encoding='utf-8') as f:
        f.write('\n'.join(idx))
    print('索引页: %s' % ip)


if __name__ == '__main__':
    main()
