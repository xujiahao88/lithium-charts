# -*- coding: utf-8 -*-
"""
build_lc_sched_charts.py — 下游排产（则言）→ lithium-charts 五个数据集

历史数据（2026-11-03 起）：**双源合并**，解决「排产 sheet 只有 2024-10 起」的历史缺口
  ① 则言排产.xlsx（不加密，openpyxl 读）
     · sheet《锂电排产-则言月初》——电池六家 / 正极 16 家 / 各环节备注，
       月度 **2024-01 ~ 2026-11**（原 readdata 只有 2024-10 起，2024 年只有 Q4 三个点）
     · sheet《锂盐-则言-（月）》——碳酸锂分企业（77 家）月度 **2023-01 起**
  ② readdata-N.xls（**加密，密码 yongfu**，msoffcrypto 解密 → xlrd）
     · sheet《排产》——2024-10 ~ 2026-11，用于：碳酸锂后三个月（则言锂盐表只到 2026-08）、
       前驱体/电解液/负极（则言里只有文字备注没有数）、电池全样本合计

合并规则
  · 电池 / 正极：整段用则言（单一来源，不会出现拼接跳变；已核对 2024-10~2026-11 与
    readdata 逐企业一致，仅 E/G 有 0.1~0.6 GWh 级修订）
  · 碳酸锂：2023-01~2024-09 取《锂盐》表，2024-10 起改用 readdata（两源在 2024-10 完全对齐，
    越往后越发散——2025-06 志存 4300 vs 1300，属不同复古版本，**只在 2024-10 这个接缝拼接**）
  · 前驱体 / 电解液 / 负极：仅 readdata（2024-10 起）

产出五个 dataset（各自一个 tab）
  · lc_sched      下游排产     —— 汇总表（电池六家 / 正极全样本 × 最近三月）+ 4 张季节性图
  · lc_sched_bat  电池分企业   —— 6 家，GWh
  · lc_sched_pos  正极分企业   —— 磷酸铁锂 11 家 + 三元 6 家
  · lc_sched_lc   碳酸锂分企业 —— 10 家 + 智利出口中国（进口参考），吨
  · lc_sched_co   其他材料     —— 前驱体 4 / 电解液 3 / 负极 2

口径要点（务必保留在 note 里）
  1. 电池「合计」行是全样本（2026-11 铁锂 246 / 三元 54 GWh），分企业只列 6 家
     （六家合计 175.3 GWh）——图片版式用的是**六家**，两者不可混比。
  2. 正极「合计」行是全样本（万吨，2026-11 铁锂 62.41 / 三元 9.83），分企业列示 11 家铁锂
     合计 43.75 万吨，差额为未列名企业。
  3. 环比一律按同一源表版本的相邻两月计算；历史月会被源表回溯修订。
  4. 碳酸锂分企业 2023-01~2024-09 来自《锂盐-则言-（月）》，2024-10 起来自 readdata；
     两源接缝在 2024-10 对齐（逐家一致，华友 3000 vs 3100 差 100 吨）。
  5. 前驱体/电解液/负极只有 2024-10 起，2024 曲线仍为 Q4 三点。

用法：
  python scripts/build_lc_sched_charts.py            # 落盘 data/lc_sched*.json + data.js + meta.json
  python scripts/build_lc_sched_charts.py --check    # 只打印结构与自检，不落盘
  python scripts/build_lc_sched_charts.py --zy X.xlsx --rd Y.xls --pwd yongfu
"""
import argparse
import datetime
import json
import os
import re
import tempfile
from decimal import Decimal, ROUND_HALF_UP

import msoffcrypto
import openpyxl
import xlrd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
BASE = r'C:\Users\Administrator\Nutstore\1\我的坚果云\碳酸锂'
ZY_SRC = os.path.join(BASE, '则言排产.xlsx')
RD_SRC = os.path.join(BASE, 'readdata-10.xls')
PWD = 'yongfu'

M_START = (2023, 1)
M_END = (2026, 11)
MONTHS = []


def _seq(a, b):
    out, (y, m) = [], a
    while (y, m) <= b:
        out.append((y, m))
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return out


MONTHS = _seq(M_START, M_END)
YEARS = (2024, 2025, 2026)
COLOR = {'2024': '#A5A5A5', '2025': '#0070C0', '2026': '#FF0000'}
NUM_RE = re.compile(r'([\d.]+)\s*三元')
LFP_RE = re.compile(r'([\d.]+)\s*铁锂')

# 则言《锂电排产-则言月初》正极区：行 30~45，个别行 col3 名称为空（第 40 行 = BP（NCM））
POS_FALLBACK = ['D（LFP）', 'Y（LFP）', 'L（LFP）', 'W（LFP）', 'R（LFP）',
                'RB（NCM）', 'B（NCM）', 'D（NCM）', 'C（NCM）', 'RX（NCM）',
                'BP（NCM）', 'G（LFP）', 'YS（LFP）', '新增邦普（LFP）',
                '新增富临（LFP）', '新增协鑫（LFP）']

# 电池六家：源表只给字母代号，补中文名（C/B 两行源表已带全名）
BAT_NAME = {'C': '宁德时代', 'B': '比亚迪', 'E': '亿纬锂能', 'G': '国轩高科',
            'Z': '中创新航', 'R': '瑞浦兰钧'}

# readdata 碳酸锂代号 → 则言《锂盐》表企业名
LC_MAP = {
    'T（LC）': '天齐锂业',
    'G（LC）': '赣锋锂业',
    'Z（LC）': '志存锂业',
    'L（LC）': '蓝科（盐湖股份）',
    'Y（LC）': '永兴特钢',
    'S（LC）': '盛新锂能',
    'J（LC）': '九岭锂业',
    'H（LC）': '华友锂业',
    '天宜（不含奉新时代）（LC）': '天宜锂业',
    '中矿（LC）': '中矿资源',
}


# ---------------------------------------------------------------- 通用工具
def clean(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return None if v != v else float(v)
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


def rnd(v, nd):
    """四舍五入（半进一，与 Excel 一致；内置 round 是银行家舍入，9.85→9.8 会错）。"""
    if v is None:
        return None
    d = Decimal(str(v)).quantize(Decimal('1' if not nd else '1.' + '0' * nd),
                                rounding=ROUND_HALF_UP)
    return int(d) if not nd else float(d)


def pct_mom(cur, prev):
    if cur is None or prev in (None, 0):
        return None
    return rnd((cur / prev - 1) * 100, 1)


def by_year(md, year):
    """{(y,m): v} → 该年 12 槽数组。"""
    out = [None] * 12
    for (y, m), v in md.items():
        if y == year:
            out[m - 1] = v
    return out


def ser(name, data, marker='none', width=1.5, dash='solid'):
    return {'name': name, 'color': COLOR[name], 'width': width, 'dash': dash,
            'smooth': False, 'marker': marker, 'data': data}


def yrs_series(md, nd, marker_year=2026, years=YEARS):
    return [ser('%d' % y, [rnd(v, nd) for v in by_year(md, y)],
                marker='circle' if y == marker_year else 'none') for y in years]


def chart(key, title, group, series, nd, unit='', sub=None, tbl=None):
    c = {'key': key, 'title': title, 'type': 'line', 'axis': 0, 'group': group,
         'series': series, 'fmt': 'v', 'nd': nd}
    if tbl:
        c['table'] = tbl
    if sub:
        c['sub'] = sub
    return c


def metric(cur, prev, nd):
    return {'labels': ['2026-11', '2026-10', '环比%'], 'kinds': ['val', 'val', 'pct'],
            'values': [rnd(cur, nd), rnd(prev, nd), pct_mom(cur, prev)]}


# ---------------------------------------------------------------- 源 ①：则言排产
def load_zy(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['锂电排产-则言月初']
    mcols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(2, c).value
        if isinstance(v, datetime.datetime):
            mcols[(v.year, v.month)] = c
    note_col = ws.max_column

    def cell(r, m):
        c = mcols.get(m)
        return ws.cell(r, c).value if c else None

    def note(r):
        return txt(ws.cell(r, note_col).value)

    bat, pos, tot, pos_tot = [], [], None, None
    pos_idx = 0
    for r in range(3, ws.max_row + 1):
        raw3 = ws.cell(r, 3).value
        lines = [txt(x) for x in str(raw3).split('\n') if txt(x)] if raw3 else []
        label = lines[-1] if lines else ''          # 「瑞浦兰钧\nR（GWH）」→ 取末行代号
        firm = txt(ws.cell(r, 2).value) or (lines[0] if len(lines) > 1 else '')
        if label in ('合计（GWH）',):
            tot = {'raw': {m: cell(r, m) for m in mcols}}
            continue
        if label in ('合计（万吨）',):
            pos_tot = {'raw': {m: cell(r, m) for m in mcols}, 'note': note(r)}
            continue
        if '（GWH）' in label.upper() and '合计' not in label:
            row = {'code': label, 'name': firm, 'note': note(r), 'tri': {}, 'lfp': {}}
            for m in mcols:
                a, b = parse_pair(cell(r, m))
                if a is not None:
                    row['tri'][m] = a
                if b is not None:
                    row['lfp'][m] = b
            bat.append(row)
            continue
        # 正极区：名称含（LFP）/（NCM）；第 40 行名称为空
        if ('（LFP）' in label) or ('（NCM）' in label) or (
                not label and 30 <= r <= 45):
            name = label or (POS_FALLBACK[pos_idx] if pos_idx < len(POS_FALLBACK) else '未命名')
            row = {'code': name, 'name': name, 'note': note(r), 'vals': {}}
            for m in mcols:
                v = clean(cell(r, m))
                if v is not None:
                    row['vals'][m] = v
            pos.append(row)
            pos_idx += 1
    return {'bat': bat, 'pos': pos, 'bat_total': tot, 'pos_total': pos_tot,
            'months': sorted(mcols)}


def load_ls(path):
    """则言《锂盐-则言-（月）》：碳酸锂分企业月度产量（吨），2023-01 起。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['锂盐-则言-（月）']
    cols = {}
    for c in range(4, ws.max_column + 1):
        v = ws.cell(5, c).value
        if isinstance(v, (int, float)) and 30000 < v < 80000:
            d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))
            cols[(d.year, d.month)] = c
    out = {}
    for r in range(6, ws.max_row + 1):
        n = txt(ws.cell(r, 2).value)
        if not n or n == '合计':
            continue
        out[n] = {m: clean(ws.cell(r, c).value) for m, c in cols.items()
                  if clean(ws.cell(r, c).value) is not None}
    return out, sorted(cols)


# ---------------------------------------------------------------- 源 ②：readdata（加密）
def open_rd(path, pwd):
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
    wb = xlrd.open_workbook(tmp.name)
    try:
        os.unlink(tmp.name)
    except Exception:
        pass
    return wb


def load_rd(path, pwd):
    wb = open_rd(path, pwd)
    sh = wb.sheet_by_name('排产')
    months = [(2024 + (9 + i) // 12, (9 + i) % 12 + 1) for i in range(26)]

    def cell(r, i):
        return sh.cell_value(r, 2 + i)

    lc, mat, bat_total, pos_extra = [], [], None, []
    for r in range(1, sh.nrows):
        label = txt(sh.cell_value(r, 1))
        note = txt(sh.cell_value(r, sh.ncols - 1))
        if '合计（GWH）' in label:
            bat_total = {'raw': {m: cell(r, i) for i, m in enumerate(months)}}
            continue
        if '万华' in label:
            row = {'code': label, 'note': note, 'vals': {}}
            for i, m in enumerate(months):
                v = clean(cell(r, i))
                if v is not None:
                    row['vals'][m] = v
            pos_extra.append(row)
            continue
        if ('（LC）' in label) or ('智利' in label):
            row = {'code': label, 'note': note, 'vals': {}}
            for i, m in enumerate(months):
                v = clean(cell(r, i))
                if v is not None:
                    row['vals'][m] = v
            lc.append(row)
            continue
        if ('（FP）' in label or '（NCM)' in label or label in ('T', 'X', 'D（6F）', 'P', 'B')):
            row = {'code': label, 'note': note, 'vals': {}}
            for i, m in enumerate(months):
                v = clean(cell(r, i))
                if v is not None:
                    row['vals'][m] = v
            mat.append(row)
    # 归属分组：前驱体 4 行 → 电解液 3 行 → 负极 2 行（源表顺序）
    order = ['前驱体', '前驱体', '前驱体', '前驱体', '电解液', '电解液', '电解液', '负极', '负极']
    for i, row in enumerate(mat):
        row['seg'] = order[i] if i < len(order) else '其他'
    return {'lc': lc, 'mat': mat, 'bat_total': bat_total,
            'pos_extra': pos_extra, 'months': months}


# ---------------------------------------------------------------- 合并
def merge_lc(zy, ls, rd):
    """碳酸锂：2023-01~2024-09 取则言锂盐表，2024-10 起取 readdata。"""
    ls_map, ls_months = ls
    seam = (2024, 10)
    out = []
    for row in rd['lc']:
        code = row['code']
        if code.startswith('S（智利出口中国）'):
            out.append({'code': code, 'name': '智利出口中国', 'note': row['note'],
                        'vals': dict(row['vals']), 'seg': '进口参考'})
            continue
        zname = LC_MAP.get(code)
        md = {}
        if zname and zname in ls_map:
            md.update({m: v for m, v in ls_map[zname].items() if m < seam})
        md.update(row['vals'])          # 2024-10 起以 readdata 为准
        out.append({'code': code, 'name': zname or code, 'note': row['note'],
                    'vals': md, 'seg': '国内'})
    return out


def six_sum(bat, field):
    md = {}
    for m in MONTHS:
        tot, hit = 0.0, False
        for row in bat:
            v = row[field].get(m)
            if v is not None:
                tot += v
                hit = True
        if hit:
            md[m] = tot
    return md


# ---------------------------------------------------------------- DS1 汇总
def build_sched(zy, rd):
    six_tri = six_sum(zy['bat'], 'tri')
    six_lfp = six_sum(zy['bat'], 'lfp')
    p_tri, p_lfp = {}, {}
    for m, v in zy['pos_total']['raw'].items():
        a, b = parse_pair(v)
        if a is not None:
            p_tri[m] = a
        if b is not None:
            p_lfp[m] = b

    idx = [(2026, 8), (2026, 9), (2026, 10), (2026, 11)]
    c8, c9, c10, c11 = idx

    cols = []
    for lab, grp in [('2026/9/1', '电池 六家（GWh）'), ('2026/10/1', '电池 六家（GWh）'),
                     ('2026/11/1', '电池 六家（GWh）'), ('2026/9/1', '正极 全样本（万吨）'),
                     ('2026/10/1', '正极 全样本（万吨）'), ('2026/11/1', '正极 全样本（万吨）')]:
        cols.append({'key': 'c%d' % len(cols), 'label': lab, 'group': grp})

    summary = {
        'columns': cols,
        'rows': {
            '三元合计': [rnd(six_tri.get(c9), 1), rnd(six_tri.get(c10), 1), rnd(six_tri.get(c11), 1),
                     rnd(p_tri.get(c9), 1), rnd(p_tri.get(c10), 1), rnd(p_tri.get(c11), 1)],
            '铁锂合计': [rnd(six_lfp.get(c9), 1), rnd(six_lfp.get(c10), 1), rnd(six_lfp.get(c11), 1),
                     rnd(p_lfp.get(c9), 1), rnd(p_lfp.get(c10), 1), rnd(p_lfp.get(c11), 1)],
            '铁锂增速%': [pct_mom(six_lfp.get(c9), six_lfp.get(c8)),
                      pct_mom(six_lfp.get(c10), six_lfp.get(c9)),
                      pct_mom(six_lfp.get(c11), six_lfp.get(c10)),
                      pct_mom(p_lfp.get(c9), p_lfp.get(c8)),
                      pct_mom(p_lfp.get(c10), p_lfp.get(c9)),
                      pct_mom(p_lfp.get(c11), p_lfp.get(c10))],
            '三元增速%': [pct_mom(six_tri.get(c9), six_tri.get(c8)),
                      pct_mom(six_tri.get(c10), six_tri.get(c9)),
                      pct_mom(six_tri.get(c11), six_tri.get(c10)),
                      pct_mom(p_tri.get(c9), p_tri.get(c8)),
                      pct_mom(p_tri.get(c10), p_tri.get(c9)),
                      pct_mom(p_tri.get(c11), p_tri.get(c10))],
        },
        'rowOrder': ['三元合计', '铁锂合计', '铁锂增速%', '三元增速%'],
        'rowMeta': {'铁锂增速%': {'signed': True, 'suffix': '%'},
                    '三元增速%': {'signed': True, 'suffix': '%'}},
        'currentWeek': '2026/11/1', 'previousWeek': '2026/10/1',
        'unit': 'GWh / 万吨', 'nd': 1,
    }

    charts = [
        chart('sched-bat-tri', '头部六家三元电池排产（GWh）', '电池 六家（GWh·月）',
              yrs_series(six_tri, 2), 2, 'GWh',
              tbl=metric(six_tri.get(c11), six_tri.get(c10), 2)),
        chart('sched-bat-lfp', '头部六家铁锂电池排产（GWh）', '电池 六家（GWh·月）',
              yrs_series(six_lfp, 1), 1, 'GWh',
              tbl=metric(six_lfp.get(c11), six_lfp.get(c10), 1)),
        chart('sched-pos-tri', '三元正极材料排产（万吨）', '正极 全样本（万吨·月）',
              yrs_series(p_tri, 2), 2, '万吨',
              tbl=metric(p_tri.get(c11), p_tri.get(c10), 2)),
        chart('sched-pos-lfp', '磷酸铁锂正极材料排产（万吨）', '正极 全样本（万吨·月）',
              yrs_series(p_lfp, 2), 2, '万吨',
              tbl=metric(p_lfp.get(c11), p_lfp.get(c10), 2)),
    ]

    all_lfp = None
    if rd.get('bat_total'):
        all_lfp = parse_pair(rd['bat_total']['raw'].get((2026, 11)))[1]

    note = '｜'.join([
        '口径：则言排产（2026-11 月初版本，上下游综合评估的预测值），历史回溯至 2024-01',
        '电池按「头部六家」（C/B/E/G/Z/瑞浦兰钧）加总，与源表「合计」行（全样本，2026-11 铁锂 %s GWh）'
        '口径不同——图与表统一用六家，勿与全样本互换' % (rnd(all_lfp, 0) if all_lfp else '246'),
        '正极取源表「合计」行（全样本万吨），而分企业仅列 11 家铁锂（2026-11 合计 43.75 万吨），'
        '差额为未列名企业',
        '增速 = 环比（同一源表版本相邻两月），历史月会被源表回溯修订',
        '⚠ 2024 年起为完整 12 个月，季节形状可与 2025/2026 直接对比（此前版本 2024 只有 Q4 三点）',
    ])
    return {
        'id': 'lc_sched', 'name': '下游排产',
        'axes': [[str(i + 1) for i in range(12)]],
        'asOf': '2026-11-01',
        'charts': charts, 'summary': summary, 'note': note,
        'unit': 'GWh / 万吨',
    }


# ---------------------------------------------------------------- DS2 电池分企业
def build_bat(zy):
    charts = []
    for i, row in enumerate(zy['bat']):
        code = row['code'].replace('（GWH）', '').replace('（GWh）', '').strip()
        firm = row['name'] or BAT_NAME.get(code, '')
        vals = {m: (row['tri'].get(m) or 0) + (row['lfp'].get(m) or 0)
                for m in set(row['tri']) | set(row['lfp'])}
        cur, prev = vals.get((2026, 11)), vals.get((2026, 10))
        sub = row['note'] or None
        if sub and len(sub) > 90:
            sub = sub[:88] + '…'
        t, l = row['tri'].get((2026, 11)), row['lfp'].get((2026, 11))
        if t is not None and l is not None:
            sub = '%s（11月：三元 %.2f + 铁锂 %.2f GWh）' % (
                (sub + '｜') if sub else '', t, l)
        charts.append(chart(
            'bat-%d' % i, '%s（%s）' % (firm, code), '电池分企业（GWh·月）',
            yrs_series(vals, 2), 2, 'GWh', sub=sub, tbl=metric(cur, prev, 2)))

    note = '｜'.join([
        '口径：则言排产《锂电排产-则言月初》电池分企业，2024-01 ~ 2026-11（预测值）',
        '纵轴为该企业「三元 + 铁锂」合计（GWh）；卡片副标题为源表「备注」原话 + 11 月拆分',
        '六家合计 2026-11 = 175.3 GWh，而源表「合计」行（全样本）= 246 GWh，差额为未列名企业',
        '企业代码：C=宁德时代、B=比亚迪、E=亿纬锂能、G=国轩高科、Z=中创新航、R=瑞浦兰钧',
    ])
    return {'id': 'lc_sched_bat', 'name': '电池分企业',
            'axes': [[str(i + 1) for i in range(12)]], 'asOf': '2026-11-01',
            'charts': charts, 'note': note, 'unit': 'GWh'}


# ---------------------------------------------------------------- DS3 正极分企业
def pos_group(name):
    return '正极·三元（吨·月）' if 'NCM' in name.upper() else '正极·磷酸铁锂（吨·月）'


def build_pos(zy, rd):
    rows = [dict(r) for r in zy['pos']]
    # 则言月初表未列「新增万华」，用 readdata 补（2025-01 起）
    for r in rd.get('pos_extra', []):
        rows.append({'code': '新增万华（LFP）', 'name': '新增万华',
                     'note': (r['note'] + '｜则言月初表未列示，2025-01 起') if r['note']
                             else '则言月初表未列示，2025-01 起',
                     'vals': dict(r['vals'])})
    charts = []
    for i, row in enumerate(rows):
        name = row['code']
        vals = row['vals']
        cur, prev = vals.get((2026, 11)), vals.get((2026, 10))
        sub = row['note'] or None
        if sub and len(sub) > 90:
            sub = sub[:88] + '…'
        charts.append(chart(
            'pos-%d' % i, name.replace('（LFP）', '（磷酸铁锂）').replace('（NCM）', '（三元）'),
            pos_group(name), yrs_series(vals, 0), 0, '吨',
            sub=sub, tbl=metric(cur, prev, 0)))

    note = '｜'.join([
        '口径：则言排产《锂电排产-则言月初》正极分企业，2024-01 ~ 2026-11（预测值，吨）',
        '磷酸铁锂 11 家 + 三元 6 家；企业代号为源表自用字母，卡片副标题为源表「备注」原话',
        '列示 11 家铁锂 2026-11 合计 43.75 万吨，源表「合计」行（全样本）为 62.41 万吨——'
        '差额为未列名企业，勿把列示合计当全行业',
        '「新增富临」11 月源表备注「未回复」、当月留空，图上是断点不是 0',
    ])
    return {'id': 'lc_sched_pos', 'name': '正极分企业',
            'axes': [[str(i + 1) for i in range(12)]], 'asOf': '2026-11-01',
            'charts': charts, 'note': note, 'unit': '吨'}


# ---------------------------------------------------------------- DS4 碳酸锂分企业
def build_lc(rows, rd):
    # 国内企业在前，「智利出口中国」放最后（性质是进口节奏参考）
    rows = [r for r in rows if r['seg'] != '进口参考'] + \
           [r for r in rows if r['seg'] == '进口参考']
    charts = []
    for i, row in enumerate(rows):
        vals = row['vals']
        cur, prev = vals.get((2026, 11)), vals.get((2026, 10))
        grp = '碳酸锂·进口参考（吨·月）' if row['seg'] == '进口参考' else '碳酸锂分企业（吨·月）'
        sub = row['note'] or None
        if sub and len(sub) > 90:
            sub = sub[:88] + '…'
        charts.append(chart(
            'lc-%d' % i, '%s  %s' % (row['name'], row['code'].replace('（LC）', ''))
            if row['seg'] != '进口参考' else '智利出口中国（吨）',
            grp, yrs_series(vals, 0), 0, '吨', sub=sub, tbl=metric(cur, prev, 0)))

    note = '｜'.join([
        '口径：碳酸锂分企业月度排产（吨）——2023-01~2024-09 取则言《锂盐-则言-（月）》，'
        '2024-10 起取 readdata 排产表；两源在 2024-10 接缝处逐家对齐（华友差 100 吨）',
        'Z = 志存（含西海/原新疆志存），Y = 永兴特钢，H = 华友，S = 盛新，J = 九岭，L = 蓝科，'
        'T = 天齐，G = 赣锋，另有天宜（不含奉新时代）、中矿',
        '「智利出口中国」为智利发往中国的锂盐量，非国内产量，只作进口节奏参考，量级不可直接比',
        '列示 10 家不等于全国：源表「国内合计」2026-11 为 12.18 万吨，列示企业合计明显更小',
    ])
    return {'id': 'lc_sched_lc', 'name': '碳酸锂分企业',
            'axes': [[str(i + 1) for i in range(12)]], 'asOf': '2026-11-01',
            'charts': charts, 'note': note, 'unit': '吨'}


# ---------------------------------------------------------------- DS5 其他材料
def build_mat(rd):
    charts = []
    for i, row in enumerate(rd['mat']):
        vals = row['vals']
        cur, prev = vals.get((2026, 11)), vals.get((2026, 10))
        sub = row['note'] or None
        if sub and len(sub) > 90:
            sub = sub[:88] + '…'
        charts.append(chart(
            'mat-%d' % i, '%s（%s）' % (row['code'], row['seg']),
            '%s（吨·月）' % row['seg'], yrs_series(vals, 0), 0, '吨',
            sub=sub, tbl=metric(cur, prev, 0)))

    note = '｜'.join([
        '口径：readdata 排产表 前驱体 / 电解液 / 负极，2024-10 ~ 2026-11（预测值，吨）',
        '则言《锂电排产-则言月初》这三块只有文字备注、没有数值，故历史只能到 2024-10，'
        '2024 曲线仍为 Q4 三点，不可用于季节形状比较',
        '企业代号沿用源表：前驱体 T/H/Z/B，电解液 T/X/D（6F），负极 P/B',
    ])
    return {'id': 'lc_sched_co', 'name': '其他材料',
            'axes': [[str(i + 1) for i in range(12)]], 'asOf': '2026-11-01',
            'charts': charts, 'note': note, 'unit': '吨'}


# ---------------------------------------------------------------- 落盘
SCHED_IDS = ['lc_sched', 'lc_sched_bat', 'lc_sched_pos', 'lc_sched_lc', 'lc_sched_co']


def write_all(ds_list):
    meta_path = os.path.join(DATA, 'meta.json')
    order = []
    if os.path.exists(meta_path):
        try:
            with open(meta_path, encoding='utf-8') as f:
                order = [d['id'] for d in json.load(f).get('datasets', [])]
        except Exception:
            pass
    pos = order.index('lc_sched') if 'lc_sched' in order else len(order)
    order = [x for x in order if x not in SCHED_IDS]
    order[pos:pos] = SCHED_IDS

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
        dd = {k: v for k, v in d.items() if not k.startswith('_')}
        p = os.path.join(DATA, dd['id'] + '.json')
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(dd, f, ensure_ascii=False, separators=(',', ':'))
        print('[ok] 写 %s.json (%.0f KB) charts=%d asOf=%s'
              % (dd['id'], os.path.getsize(p) / 1024, len(dd['charts']), dd['asOf']))

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
    ap.add_argument('--zy', default=ZY_SRC)
    ap.add_argument('--rd', default=RD_SRC)
    ap.add_argument('--pwd', default=PWD)
    args = ap.parse_args()

    zy = load_zy(args.zy)
    ls = load_ls(args.zy)
    rd = load_rd(args.rd, args.pwd)

    print('则言月初：电池 %d 家 / 正极 %d 家 / 月份 %s~%s'
          % (len(zy['bat']), len(zy['pos']), zy['months'][0], zy['months'][-1]))
    print('则言锂盐：%d 家 / 月份 %s~%s' % (len(ls[0]), ls[1][0], ls[1][-1]))
    print('readdata：碳酸锂 %d 行 / 材料 %d 行 / 月份 %s~%s'
          % (len(rd['lc']), len(rd['mat']), rd['months'][0], rd['months'][-1]))

    lc_rows = merge_lc(zy, ls, rd)

    ds = [build_sched(zy, rd), build_bat(zy), build_pos(zy, rd),
          build_lc(lc_rows, rd), build_mat(rd)]
    for d in ds:
        yrs = {s['name']: sum(1 for v in s['data'] if v is not None)
               for s in d['charts'][0]['series']}
        print('  %-14s %-12s 图 %2d | 首图各年非空 %s'
              % (d['id'], d['name'], len(d['charts']), yrs))
    if ds[0].get('summary'):
        for k in ds[0]['summary']['rowOrder']:
            print('   %-8s %s' % (k, ds[0]['summary']['rows'][k]))

    # 自检：碳酸锂 2023 覆盖、正极 2024 覆盖
    lc0 = ds[3]['charts'][0]['series']
    print('   [自检] 碳酸锂首图年覆盖：',
          {s['name']: sum(1 for v in s['data'] if v is not None) for s in lc0})
    pos0 = ds[2]['charts'][0]['series']
    print('   [自检] 正极首图年覆盖：',
          {s['name']: sum(1 for v in s['data'] if v is not None) for s in pos0})

    if args.check:
        for d in ds:
            for c in d['charts']:
                print('     [%s] %s | %s' % (c['group'], c['title'], c.get('sub') or ''))
        return
    write_all(ds)


if __name__ == '__main__':
    main()
