# -*- coding: utf-8 -*-
"""
build_lc_sched_charts.py — 下游排产（加密 xls）→ lithium-charts 两个数据集

源：readdata-10.xls（**加密，密码 yongfu**；OLE2/RC4，需 msoffcrypto 解密后 xlrd 读）
  三个 sheet：
    · 排产      环节(电池/正极/碳酸锂/前驱体/电解液/负极) × 企业 × 2024-10~2026-11 月度
                电池行为「x三元+y铁锂」文本（GWh）；正极/碳酸锂/前驱体/电解液/负极为吨；
                另有 4 行合计：电池合计（全样本 GWh）、正极合计（万吨，文本）、
                碳酸锂国内合计（万吨）、以及各企业「备注」列（第 28 列，分析师原话）
    · 磷酸铁锂  45 家企业 2023-01~2026-09 月度产量（吨）+ 末行合计
    · 碳酸锂    87 家企业 2024-01~2026-09 月度产量（吨）+ 末行合计

产出两个 dataset：
  · lc_sched     下游排产     —— 图版式对标用户给的「下游排产（月初）」表：
                    顶部汇总表（电池六家 / 正极全样本 × 2026-09·10·11 三元/铁锂及其环比）
                    + 4 张季节性图（头部六家三元电池 / 三元正极 / 头部六家铁锂电池 / 铁锂正极）
  · lc_sched_co  分企业排产   —— 43 家企业逐家一张季节性图（环节分组），图下带
                    2026-11 / 2026-10 / 环比 指标条，卡片副标题放源表「备注」原话

口径要点（务必保留在 note 里）：
  1. 电池「合计」行是全样本（2026-11 铁锂 246 GWh），而分企业只列了 6 家
     （六家合计 175.3 GWh）——图片版式用的是**六家**口径，两者不可混比。
  2. 正极「合计」行以万吨文本给出（如「62.41铁锂+9.83三元」），与磷酸铁锂 sheet、
     排产 sheet 里企业行（吨）是同一口径换算，量级一致。
  3. **源表历史只到 2024-10**，故 2024 年曲线只有 10~12 月三个点；不要拿 2024
     的季节形状做同比结论。磷酸铁锂/碳酸锂明细表另有 2023-01 / 2024-01 起的历史，
     但与排产 sheet 的合计在 2024-10~2025-02 有 0.8~3.0 万吨级差异（源表版本不同），
     **刻意不混用**，避免把口径差读成趋势。
  4. 环比一律按「源表当前版本的相邻两月」计算；历史月会被源表回溯修订，
     与上月已发布版本算出来的环比可能差 0.5 个点左右。

用法：
  python scripts/build_lc_sched_charts.py               # 落盘 data/lc_sched*.json + data.js + meta.json
  python scripts/build_lc_sched_charts.py --check       # 只打印结构，不落盘
  python scripts/build_lc_sched_charts.py --src X.xls --pwd yongfu
"""
import argparse
import datetime
import json
import os
import re
import tempfile
from decimal import Decimal, ROUND_HALF_UP

import msoffcrypto
import xlrd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SRC = r'C:\Users\Administrator\Nutstore\1\我的坚果云\碳酸锂\readdata-10.xls'
PWD = 'yongfu'

# 排产 sheet 的月列：C 列（0-based 2）起 26 个月
M0 = (2024, 10)
NMON = 26
MONTHS = []
_y, _m = M0
for _ in range(NMON):
    MONTHS.append((_y, _m))
    _m += 1
    if _m > 12:
        _y, _m = _y + 1, 1

COLOR = {'2024': '#A5A5A5', '2025': '#0070C0', '2026': '#FF0000'}
NUM_RE = re.compile(r'([\d.]+)\s*三元')
LFP_RE = re.compile(r'([\d.]+)\s*铁锂')


# ---------------------------------------------------------------- 读源表
def open_workbook(path, pwd):
    """加密 xls：先解密到临时文件再交给 xlrd（xlrd 不支持带密码打开）。"""
    f = open(path, 'rb')
    of = msoffcrypto.OfficeFile(f)
    if not of.is_encrypted():
        f.close()
        return xlrd.open_workbook(path)
    of.load_key(password=pwd)
    tmp = tempfile.NamedTemporaryFile(suffix='.xls', delete=False)
    with open(tmp.name, 'wb') as g:
        of.decrypt(g)
    f.close()
    wb = xlrd.open_workbook(tmp.name)          # 一次性读入内存
    try:
        os.unlink(tmp.name)
    except Exception:
        pass
    return wb


def clean(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if v != v:
            return None
        return float(v)
    s = str(v).strip()
    if s in ('', '/', '-', '—', '#N/A', '#VALUE!', '#DIV/0!'):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_pair(v):
    """「17三元+86.5铁锂」→ (三元, 铁锂)；缺项返回 None。"""
    s = str(v)
    a = NUM_RE.search(s)
    b = LFP_RE.search(s)
    return (float(a.group(1)) if a else None, float(b.group(1)) if b else None)


def txt(v):
    return re.sub(r'\s+', ' ', str(v)).strip() if v is not None else ''


def read_paichan(wb):
    sh = wb.sheet_by_name('排产')
    blocks, cur = [], None
    for r in range(2, sh.nrows):
        seg = txt(sh.cell_value(r, 0))
        name = txt(sh.cell_value(r, 1))
        if not name:
            continue
        if seg:
            cur = {'seg': seg, 'rows': []}
            blocks.append(cur)
        if cur is None:
            continue
        note = txt(sh.cell_value(r, 28)) if sh.ncols > 28 else ''
        if '合计' in name:
            cur['total'] = {
                'name': name, 'note': note,
                'raw': [sh.cell_value(r, 2 + i) for i in range(NMON)],
            }
            continue
        row = {'name': name, 'note': note}
        if '三元' in txt(sh.cell_value(r, 2)) or '铁锂' in txt(sh.cell_value(r, 2)) \
                or any('三元' in txt(sh.cell_value(r, 2 + i)) for i in range(3)):
            row['kind'] = 'battery'
            row['tri'], row['lfp'] = [], []
            for i in range(NMON):
                a, b = parse_pair(sh.cell_value(r, 2 + i))
                row['tri'].append(a)
                row['lfp'].append(b)
        else:
            row['kind'] = 'num'
            row['vals'] = [clean(sh.cell_value(r, 2 + i)) for i in range(NMON)]
        cur['rows'].append(row)
    return blocks


# ---------------------------------------------------------------- 数据工具
def by_year(vals, year):
    """26 月序列 → 12 槽（按月）该年数组，缺月 None。"""
    out = [None] * 12
    for i, (y, m) in enumerate(MONTHS):
        if y == year:
            out[m - 1] = vals[i]
    return out


def rnd(v, nd):
    """四舍五入（半进一，与 Excel 一致；Python 内置 round 是银行家舍入，9.85→9.8 会错）。"""
    if v is None:
        return None
    d = Decimal(str(v)).quantize(Decimal('1' if not nd else '1.' + '0' * nd),
                                rounding=ROUND_HALF_UP)
    return int(d) if not nd else float(d)


def ser(name, data, marker='none', width=1.5, dash='solid', smooth=False):
    return {'name': name, 'color': COLOR[name], 'width': width, 'dash': dash,
            'smooth': smooth, 'marker': marker, 'data': data}


def chart(key, title, group, series, nd, unit='', sub=None, tbl=None):
    c = {'key': key, 'title': title, 'type': 'line', 'axis': 0, 'group': group,
         'series': series, 'fmt': 'v', 'nd': nd}
    if tbl:
        c['table'] = tbl
    if sub:
        c['sub'] = sub
    return c


def sum_rows(rows, field, idx=None):
    """按行求和（field='tri'/'lfp'/'vals'），返回 26 月列表。"""
    n = NMON
    out = [0.0] * n
    seen = [False] * n
    for r in rows:
        arr = r[field]
        for i in range(n):
            v = arr[i] if idx is None else (arr[i] if arr[i] is not None else None)
            if v is None:
                continue
            out[i] += v
            seen[i] = True
    return [out[i] if seen[i] else None for i in range(n)]


def pct_mom(cur, prev):
    if cur is None or prev in (None, 0):
        return None
    return rnd((cur / prev - 1) * 100, 1)


# ---------------------------------------------------------------- 数据集 1：下游排产
def build_sched(blocks):
    bat = next(b for b in blocks if b['seg'] == '电池')
    pos = next(b for b in blocks if b['seg'] == '正极')

    six_tri = sum_rows(bat['rows'], 'tri')
    six_lfp = sum_rows(bat['rows'], 'lfp')
    tot_tri_c = None
    # 正极合计（万吨文本）
    p_lfp, p_tri = [], []
    for v in pos['total']['raw']:
        a, b = parse_pair(v)      # parse_pair 返回 (三元, 铁锂)
        p_tri.append(a)
        p_lfp.append(b)
    # 电池 全样本合计（GWh 文本）
    all_tri, all_lfp = [], []
    for v in bat['total']['raw']:
        a, b = parse_pair(v)
        all_tri.append(a)
        all_lfp.append(b)

    def idx(y, m):
        return MONTHS.index((y, m))

    i9, i10, i11 = idx(2026, 9), idx(2026, 10), idx(2026, 11)
    i8 = idx(2026, 8)

    six9, six10, six11 = six_tri[i9], six_tri[i10], six_tri[i11]
    lfp9, lfp10, lfp11 = six_lfp[i9], six_lfp[i10], six_lfp[i11]
    p_tri9, p_tri10, p_tri11 = p_tri[i9], p_tri[i10], p_tri[i11]
    p_lfp9, p_lfp10, p_lfp11 = p_lfp[i9], p_lfp[i10], p_lfp[i11]

    cols = []
    for i, (lab, grp) in enumerate([('2026/9/1', '电池 六家（GWh）'),
                                    ('2026/10/1', '电池 六家（GWh）'),
                                    ('2026/11/1', '电池 六家（GWh）'),
                                    ('2026/9/1', '正极 全样本（万吨）'),
                                    ('2026/10/1', '正极 全样本（万吨）'),
                                    ('2026/11/1', '正极 全样本（万吨）')]):
        cols.append({'key': 'c%d' % i, 'label': lab, 'group': grp})

    summary = {
        'columns': cols,
        'rows': {
            '三元合计': [rnd(six9, 1), rnd(six10, 1), rnd(six11, 1),
                     rnd(p_tri9, 1), rnd(p_tri10, 1), rnd(p_tri11, 1)],
            '铁锂合计': [rnd(lfp9, 1), rnd(lfp10, 1), rnd(lfp11, 1),
                     rnd(p_lfp9, 1), rnd(p_lfp10, 1), rnd(p_lfp11, 1)],
            '铁锂增速%': [pct_mom(lfp9, six_lfp[i8]), pct_mom(lfp10, lfp9), pct_mom(lfp11, lfp10),
                      pct_mom(p_lfp9, p_lfp[i8]), pct_mom(p_lfp10, p_lfp9),
                      pct_mom(p_lfp11, p_lfp10)],
            '三元增速%': [pct_mom(six9, six_tri[i8]), pct_mom(six10, six9), pct_mom(six11, six10),
                      pct_mom(p_tri9, p_tri[i8]), pct_mom(p_tri10, p_tri9),
                      pct_mom(p_tri11, p_tri10)],
        },
        'rowOrder': ['三元合计', '铁锂合计', '铁锂增速%', '三元增速%'],
        'rowMeta': {'铁锂增速%': {'signed': True, 'suffix': '%'},
                    '三元增速%': {'signed': True, 'suffix': '%'}},
        'currentWeek': '2026/11/1', 'previousWeek': '2026/10/1',
        'unit': 'GWh / 万吨', 'nd': 1,
    }

    charts = []
    # 1) 头部六家三元电池
    charts.append(chart(
        'sched-bat-tri', '头部六家三元电池排产（GWh）', '电池 六家（GWh·月）',
        [ser('%d' % y, [rnd(v, 2) for v in by_year(six_tri, y)],
             marker='circle' if y == 2026 else 'none') for y in (2024, 2025, 2026)],
        2, 'GWh',
        tbl={'labels': ['2026-11', '2026-10', '环比%'],
             'kinds': ['val', 'val', 'pct'],
             'values': [rnd(six11, 2), rnd(six10, 2), pct_mom(six11, six10)]}))
    # 2) 头部六家铁锂电池
    charts.append(chart(
        'sched-bat-lfp', '头部六家铁锂电池排产（GWh）', '电池 六家（GWh·月）',
        [ser('%d' % y, [rnd(v, 1) for v in by_year(six_lfp, y)],
             marker='circle' if y == 2026 else 'none') for y in (2024, 2025, 2026)],
        1, 'GWh',
        tbl={'labels': ['2026-11', '2026-10', '环比%'],
             'kinds': ['val', 'val', 'pct'],
             'values': [rnd(lfp11, 1), rnd(lfp10, 1), pct_mom(lfp11, lfp10)]}))
    # 3) 三元正极
    charts.append(chart(
        'sched-pos-tri', '三元正极材料排产（万吨）', '正极 全样本（万吨·月）',
        [ser('%d' % y, [rnd(v, 2) for v in by_year(p_tri, y)],
             marker='circle' if y == 2026 else 'none') for y in (2024, 2025, 2026)],
        2, '万吨',
        tbl={'labels': ['2026-11', '2026-10', '环比%'],
             'kinds': ['val', 'val', 'pct'],
             'values': [rnd(p_tri11, 2), rnd(p_tri10, 2), pct_mom(p_tri11, p_tri10)]}))
    # 4) 磷酸铁锂正极
    charts.append(chart(
        'sched-pos-lfp', '磷酸铁锂正极材料排产（万吨）', '正极 全样本（万吨·月）',
        [ser('%d' % y, [rnd(v, 2) for v in by_year(p_lfp, y)],
             marker='circle' if y == 2026 else 'none') for y in (2024, 2025, 2026)],
        2, '万吨',
        tbl={'labels': ['2026-11', '2026-10', '环比%'],
             'kinds': ['val', 'val', 'pct'],
             'values': [rnd(p_lfp11, 2), rnd(p_lfp10, 2), pct_mom(p_lfp11, p_lfp10)]}))

    note = '｜'.join([
        '口径：源表《readdata-10.xls·排产》2026-11 月初版本（数据为上下游综合评估的预测值）',
        '电池按「头部六家」（C/B/E/G/Z/瑞浦兰钧）加总，与源表「合计」行（全样本，2026-11 铁锂 246 / 三元 54 GWh）'
        '口径不同——图与表统一用六家，勿与全样本互换',
        '正极取源表「合计」行（全样本，万吨；与《磷酸铁锂》明细表 2026-09 的 61.33 万吨对得上），'
        '而源表正极分企业仅列 11 家铁锂（2026-11 合计 43.75 万吨），差额为未列名企业',
        '增速 = 环比（同一源表版本相邻两月），历史月会被源表回溯修订，'
        '与上期已发布版本算出的环比可能差 0.5 个点左右',
        '⚠ 源表历史起于 2024-10，2024 年曲线仅 10~12 月三点，不可用于季节形状比较',
    ])

    return {
        'id': 'lc_sched', 'name': '下游排产',
        'axes': [[str(i + 1) for i in range(12)]],
        'asOf': '2026-11-01',
        'charts': charts, 'summary': summary, 'note': note,
        'unit': 'GWh / 万吨',
        '_six9': six9, '_six11': six11, '_all11': all_lfp[i11],
    }


# ---------------------------------------------------------------- 数据集 2：分企业排产
SEG_ORDER = [
    ('电池', '电池分企业（GWh·月）', 'GWh', 2, 1, 'bat'),
    ('正极', '正极·磷酸铁锂（吨·月）', '吨', 0, 0, 'pos_lfp'),
    ('正极', '正极·三元（吨·月）', '吨', 0, 0, 'pos_ncm'),
    ('碳酸锂', '碳酸锂（吨·月）', '吨', 0, 0, 'lc'),
    ('前驱体', '前驱体（吨·月）', '吨', 0, 0, 'pre'),
    ('电解液', '电解液（吨·月）', '吨', 0, 0, 'elec'),
    ('负极', '负极（吨·月）', '吨', 0, 0, 'anode'),
]


def seg_of_pos(name):
    u = name.upper()
    if 'NCM' in u:
        return 'pos_ncm'
    return 'pos_lfp'


def build_sched_co(blocks):
    charts = []
    for seg, gtitle, unit, nd, nd_chart, tag in SEG_ORDER:
        blk = next((b for b in blocks if b['seg'] == seg), None)
        if not blk:
            continue
        for r in blk['rows']:
            if seg == '正极' and seg_of_pos(r['name']) != tag:
                continue
            if r['kind'] == 'battery':
                vals = [None if (a is None and b is None) else round((a or 0) + (b or 0), 2)
                        for a, b in zip(r['tri'], r['lfp'])]
                sub = r['note'] or None
                if sub and len(sub) > 90:
                    sub = sub[:88] + '…'
                if r['tri'][-1] is not None and r['lfp'][-1] is not None:
                    sub = ('%s（11月：三元 %.2f + 铁锂 %.2f GWh）'
                           % (('' if not sub else sub + '｜'),
                              r['tri'][-1] or 0, r['lfp'][-1] or 0))
            else:
                vals = r['vals']
                sub = r['note'] or None
            cur, prev = vals[-1], vals[-2]
            key = '%s-%d-%s' % (tag, len(charts), re.sub(r'\W+', '', r['name'])[:12])
            charts.append(chart(
                key, r['name'].replace('（GWH）', '（GWh）'), gtitle,
                [ser('%d' % y, [rnd(v, nd_chart) for v in by_year(vals, y)],
                     marker='circle' if y == 2026 else 'none') for y in (2025, 2026)],
                nd_chart, unit, sub=sub,
                tbl={'labels': ['2026-11', '2026-10', '环比%'],
                     'kinds': ['val', 'val', 'pct'],
                     'values': [rnd(cur, nd_chart), rnd(prev, nd_chart),
                                pct_mom(cur, prev)]}))

    note = '｜'.join([
        '口径：源表《readdata-10.xls·排产》分企业月度排产（2026-11 月初版本，预测值）',
        '电池企业为「三元+铁锂」合计（GWh）；正极/碳酸锂/前驱体/电解液/负极为吨',
        '仅列源表列示企业：正极铁锂 11 家 2026-11 合计 43.75 万吨，源表「合计」行（全样本）为 62.41 万吨，'
        '碳酸锂列示企业合计与源表「国内合计」亦有缺口——差额为未列名企业，勿把列示合计当全行业',
        '企业名沿用源表字母代码（C=宁德时代、B=比亚迪、E=亿纬、G=国轩、Z=中创新航、瑞浦兰钧），'
        '正极各行字母为源表自用代号，含义以卡片副标题的「备注」为准',
        '「新增富临」11 月源表备注「未回复」、当月留空，图上是断点不是 0',
        '2024 年只有 10~12 月，故分企业图只画 2025 / 2026 两年',
    ])
    return {
        'id': 'lc_sched_co', 'name': '分企业排产',
        'axes': [[str(i + 1) for i in range(12)]],
        'asOf': '2026-11-01',
        'charts': charts, 'note': note, 'unit': 'GWh / 吨',
    }


# ---------------------------------------------------------------- 落盘
def write_all(ds_list):
    meta_path = os.path.join(DATA, 'meta.json')
    order = []
    if os.path.exists(meta_path):
        try:
            with open(meta_path, encoding='utf-8') as f:
                order = [d['id'] for d in json.load(f).get('datasets', [])]
        except Exception:
            pass
    for d in ds_list:
        if d['id'] not in order:
            order.append(d['id'])

    all_ds = []
    for did in order:
        hit = next((d for d in ds_list if d['id'] == did), None)
        if hit:
            all_ds.append(hit)
        else:
            p = os.path.join(DATA, did + '.json')
            if os.path.exists(p):
                with open(p, encoding='utf-8') as f:
                    all_ds.append(json.load(f))
            else:
                print('[warn] 缺少 %s.json，跳过' % did)

    stamp = datetime.datetime.now().isoformat(timespec='seconds')
    for d in ds_list:
        d = {k: v for k, v in d.items() if not k.startswith('_')}
        p = os.path.join(DATA, d['id'] + '.json')
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, separators=(',', ':'))
        print('[ok] 写 %s.json (%.0f KB) charts=%d asOf=%s'
              % (d['id'], os.path.getsize(p) / 1024, len(d['charts']), d['asOf']))

    # 内存里的下划线字段不能进 data.js
    clean_all = [{k: v for k, v in d.items() if not k.startswith('_')} for d in all_ds]
    with open(os.path.join(DATA, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('// 自动生成，勿手改。build_lc_sched_charts.py @ ' + stamp + '\n')
        f.write('window.CHART_DATA = ')
        json.dump({'updated': stamp, 'datasets': clean_all}, f,
                  ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    print('[ok] 写 data.js (datasets=%d)' % len(clean_all))

    meta = {'updated': stamp,
            'datasets': [{'id': d['id'], 'name': d['name'],
                          'charts': len(d['charts']), 'asOf': d['asOf']}
                         for d in clean_all]}
    with open(meta_path, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print('[ok] 写 meta.json datasets=%d' % len(meta['datasets']))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    ap.add_argument('--src', default=SRC)
    ap.add_argument('--pwd', default=PWD)
    args = ap.parse_args()

    wb = open_workbook(args.src, args.pwd)
    blocks = read_paichan(wb)
    print('排产 环节块：', [(b['seg'], len(b['rows']), 'total' in b) for b in blocks])
    for b in blocks:
        if 'total' in b:
            print('  合计[%s] 首月/末月：%s / %s'
                  % (b['seg'], txt(b['total']['raw'][0])[:28], txt(b['total']['raw'][-1])[:28]))

    ds1 = build_sched(blocks)
    ds2 = build_sched_co(blocks)

    for ds in (ds1, ds2):
        print('\n== %s(%s) | asOf %s | 图 %d'
              % (ds['name'], ds['id'], ds['asOf'], len(ds['charts'])))
        for c in ds['charts'][:6]:
            nz = sum(1 for v in c['series'][-1]['data'] if v is not None)
            print('   [%s] %s | 年=%s | 2026非空=%d'
                  % (c['group'], c['title'], ','.join(s['name'] for s in c['series']), nz))
        if len(ds['charts']) > 6:
            print('   … 共 %d 张' % len(ds['charts']))
        s = ds.get('summary')
        if s:
            for k in s['rowOrder']:
                print('   %-8s %s' % (k, s['rows'][k]))

    if args.check:
        return
    write_all([ds1, ds2])


if __name__ == '__main__':
    main()
