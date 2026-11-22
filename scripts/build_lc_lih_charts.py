# -*- coding: utf-8 -*-
"""把《氢氧化锂.xlsx》做成锂站 5 个 tab。

tab:
  lc_lih_price      氢氧化锂价格
  lc_lih_capacity   氢氧化锂产能
  lc_lih_output     氢氧化锂产量
  lc_lih_inventory  氢氧化锂库存
  lc_lih_cost       氢氧化锂成本利润

约定与 build_lc_mine_charts.py 一致：
  - 周轴 52 槽（ISO 周 → 基准年同周周四，53 周退到第 52 周）
  - 月轴 12 槽（'1'..'12'）
  - 每张图都带 table，前端才会渲染「⤓ 下载图片」按钮
"""
import os
import io
import sys
import json
import argparse
import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SRC = r'C:\Users\Administrator\Nutstore\1\我的坚果云\碳酸锂\氢氧化锂.xlsx'
BASE_YEAR = 2026
# 2026-11-22 用户要求「做近五年就行」：周轴/月度产量由 6 年收为 5 年（2022–2026）。
# 产能月度、库存月度源表本身从 2023 才起步，保持原样（不超过五年）。
WEEK_YEARS = [2022, 2023, 2024, 2025, 2026]
MONTH_YEARS = [2023, 2024, 2025, 2026]
PROD_MONTH_YEARS = [2022, 2023, 2024, 2025, 2026]


# ---------------------------------------------------------------- 基础工具
def to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        s = v.strip()[:10]
        for fmt in ('%Y-%m-%d', '%Y/%m/%d'):
            try:
                return datetime.datetime.strptime(s, fmt).date()
            except ValueError:
                continue
    return None


def num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        return f if f == f and abs(f) != float('inf') else None
    return None


def rnd(v, nd):
    if v is None:
        return None
    return int(round(v)) if nd == 0 else round(v, nd)


def iso_key(d):
    iy, iw, _ = d.isocalendar()
    return (iy, iw)


def iso_slot(d):
    iy, iw, _ = d.isocalendar()
    try:
        return datetime.date.fromisocalendar(BASE_YEAR, iw, 4)
    except ValueError:
        return datetime.date.fromisocalendar(BASE_YEAR, 52, 4)


def series_style(i, n):
    latest, prev = (i == n - 1), (i == n - 2)
    if latest:
        color = '#FF0000'
    elif prev:
        color = '#0070C0'
    elif i == n - 3:
        color = '#A5A5A5'
    else:
        color = '#9DC3E6'
    return {'color': color, 'width': 1.6,
            'dash': 'solid' if i >= n - 3 else 'dash',
            'marker': 'circle' if latest else 'none'}


def read_series(ws, date_col, val_col, min_year=2021):
    """读 (日期列, 数值列) → {date: value}；重复日期取最后一条。"""
    di = ord(date_col) - 65
    vi = ord(val_col) - 65
    out = {}
    for r in ws.iter_rows(min_row=5, values_only=True):
        d = to_date(r[di]) if di < len(r) else None
        if d is None or d.year < min_year:
            continue
        v = num(r[vi]) if vi < len(r) else None
        if v is None:
            continue
        out[d] = v          # 源表按日期倒序，后写覆盖 = 取较早行？ 需按日期排序后取最后
    # 按日期升序重建，重复日期取「行号更大 = 日期更晚」的一条
    # 上面倒序覆盖会让较早行胜出，修正：重新按 (date, rowidx) 处理
    return out


def read_series_ordered(ws, date_col, val_col, min_year=2021):
    """源表日期倒序；逐行覆盖后，最后写入的是最早行。改为显式保留最晚日期那条。"""
    di = ord(date_col) - 65
    vi = ord(val_col) - 65
    tmp = {}
    for idx, r in enumerate(ws.iter_rows(min_row=5, values_only=True)):
        d = to_date(r[di]) if di < len(r) else None
        if d is None or d.year < min_year:
            continue
        v = num(r[vi]) if vi < len(r) else None
        if v is None:
            continue
        tmp.setdefault(d, []).append((idx, v))
    out = {}
    for d, lst in tmp.items():
        # 同一天取行号最大（源表越靠上行日期越晚）
        out[d] = min(lst)[1]
    return out


def prev_obs(vals, last):
    cand = [d for d in vals if d < last]
    return max(cand) if cand else None


def yoy_date(vals, last, period):
    if period == 'week':
        iy, iw = iso_key(last)
        cand = [d for d in vals if iso_key(d) == (iy - 1, iw)]
    elif period == 'month':
        cand = [d for d in vals if d.year == last.year - 1 and d.month == last.month]
    else:
        cand = []
    return max(cand) if cand else None


def card_table(vals, nd, period='week', with_yoy=True):
    if not vals:
        return None
    last = max(vals)
    p = prev_obs(vals, last)
    cur = vals.get(last)
    if cur is None:
        return None
    prv = vals.get(p) if p else None
    labels = ['本期', '上期', '环比']
    kinds = ['val', 'val', 'delta']
    values = [rnd(cur, nd), rnd(prv, nd), rnd(cur - prv, nd)]
    if with_yoy and period in ('week', 'month'):
        y = yoy_date(vals, last, period)
        yv = vals.get(y) if y else None
        labels.append('同比')
        kinds.append('delta')
        values.append(rnd(cur - yv, nd) if yv is not None else None)
    return {'labels': labels, 'kinds': kinds, 'values': values}


# ---------------------------------------------------------------- 轴/图构造
WEEK_AXIS = [datetime.date.fromisocalendar(BASE_YEAR, w, 4) for w in range(1, 53)]
WEEK_LABEL = [d.strftime('%m-%d') for d in WEEK_AXIS]
WEEK_POS = {d: i for i, d in enumerate(WEEK_AXIS)}
WEEK_FIRST = [d.strftime('%m-%d') for d in WEEK_AXIS if d.month not in
              [] ]  # 占位，下面覆盖
_seen = set()
WEEK_FIRST = []
for d in WEEK_AXIS:
    if d.month not in _seen:
        _seen.add(d.month)
        WEEK_FIRST.append(d.strftime('%m-%d'))

MONTH_AXIS = [str(i) for i in range(1, 13)]


def week_seasonal(vals, years):
    out = {}
    for d, v in vals.items():
        iy, _ = iso_key(d)
        if iy not in years:
            continue
        out.setdefault(iy, {})[WEEK_POS[iso_slot(d)]] = v
    return out


def month_seasonal(vals, years):
    out = {}
    for d, v in vals.items():
        if d.year not in years:
            continue
        out.setdefault(d.year, {})[d.month - 1] = v
    return out


def seasonal_chart(key, title, group, vals, axis, nd, period, years,
                   table=True, with_yoy=True, multi=None):
    """axis: 0=周 1=月；multi=None 表示单序列，否则 [(名称, vals, 颜色)]。"""
    if multi is None:
        byyear = week_seasonal(vals, years) if period == 'week' else month_seasonal(vals, years)
        present = [y for y in years if y in byyear and
                   (y == BASE_YEAR or len(byyear[y]) >= 2)]
        if not present:
            return None
        n = len(years)
        ser = []
        for y in present:
            st = series_style(years.index(y), n)
            row = [None] * (52 if period == 'week' else 12)
            for i, v in byyear[y].items():
                row[i] = rnd(v, nd)
            ser.append({'name': str(y), 'color': st['color'], 'width': st['width'],
                        'dash': st['dash'], 'marker': st['marker'], 'data': row})
        ch_table = card_table(vals, nd, period, with_yoy) if table else None
    else:
        ser = []
        for name, vmap, color in multi:
            byyear = week_seasonal(vmap, years) if period == 'week' else month_seasonal(vmap, years)
            y = max(byyear) if byyear else None
            if y is None:
                continue
            row = [None] * (52 if period == 'week' else 12)
            for i, v in byyear[y].items():
                row[i] = rnd(v, nd)
            ser.append({'name': name, 'color': color, 'width': 1.6,
                        'dash': 'solid', 'marker': 'circle', 'data': row})
            vals = vmap
        if not ser:
            return None
        ch_table = card_table(vals, nd, period, with_yoy) if table else None
    ch = {'key': key, 'title': title, 'type': 'line', 'axis': axis,
          'group': group, 'series': ser, 'fmt': 'v', 'nd': nd}
    if ch_table:
        ch['table'] = ch_table
    return ch


def date_axis_chart(key, title, group, series_defs, nd=2):
    """series_defs: [(名称, {date:value}, 颜色)]，横轴为年度日期。"""
    alldates = sorted({d for _, m, _ in series_defs for d in m})
    if not alldates:
        return None
    # 年度轴：前端 date 轴对 01 月标签显示年份，对其它月显示「N月」；
    # 源表年度值都落在 12-31，直接用会全部显示成「12月」，故统一投影到 01-01。
    def annual_label(d):
        return '%04d-01-01' % d.year

    axis = [annual_label(d) for d in alldates]
    pos = {annual_label(d): i for i, d in enumerate(alldates)}
    ser = []
    for name, m, color in series_defs:
        row = [None] * len(axis)
        for d, v in m.items():
            row[pos[annual_label(d)]] = rnd(v, nd)
        ser.append({'name': name, 'color': color, 'width': 1.8,
                    'dash': 'solid', 'marker': 'circle', 'data': row})
    last = alldates[-1]
    p = prev_obs({d: 1 for d in alldates}, last)
    def vals_at(name):
        for nm, m, _ in series_defs:
            if nm == name:
                return m
        return {}
    v0 = vals_at(series_defs[0][0])
    cur = v0.get(last)
    prv = v0.get(p) if p else None
    tbl = {'labels': ['本期', '上期', '环比'], 'kinds': ['val', 'val', 'delta'],
           'values': [rnd(cur, nd), rnd(prv, nd),
                      rnd(cur - prv, nd) if cur is not None and prv is not None else None]}
    return {'key': key, 'title': title, 'type': 'line', 'axis': 1,
            'group': group, 'series': ser, 'fmt': 'v', 'nd': nd, 'table': tbl}


# ---------------------------------------------------------------- 数据读取
def load(wb):
    return {
        'price': wb['价格'], 'cap': wb['供给-产能'], 'out': wb['供给-产量'],
        'inv': wb['库存'], 'cost': wb['成本利润'],
    }


def build_price(ws):
    p = 'lc_lih_price'
    S = {
        '电氢': read_series_ordered(ws, 'A', 'B'),
        '价差': read_series_ordered(ws, 'A', 'C'),
        '电碳': read_series_ordered(ws, 'D', 'F'),
        'CIF': read_series_ordered(ws, 'G', 'H'),
        '中日价差': read_series_ordered(ws, 'J', 'L'),
        '汇率': read_series_ordered(ws, 'M', 'N'),
    }
    charts = []
    g1 = '价格（周，元/吨）'
    g2 = 'CIF 与汇率'
    charts.append(seasonal_chart(p + '-lih', '电池级氢氧化锂（粗颗粒）平均价', g1,
                                 S['电氢'], 0, 0, 'week', WEEK_YEARS))
    charts.append(seasonal_chart(p + '-spread', '电氢-电碳价差', g1,
                                 S['价差'], 0, 0, 'week', WEEK_YEARS))
    charts.append(seasonal_chart(p + '-lc', '电池级碳酸锂价格（对照）', g1,
                                 S['电碳'], 0, 0, 'week', WEEK_YEARS))
    charts.append(seasonal_chart(p + '-cnjp', '中国-日韩电氢价差', g2,
                                 S['中日价差'], 0, 0, 'week', WEEK_YEARS))
    charts.append(seasonal_chart(p + '-cif', 'CIF 中日韩电氢（美元/公斤）', g2,
                                 S['CIF'], 0, 2, 'week', WEEK_YEARS))
    charts.append(seasonal_chart(p + '-fx', '美元兑人民币汇率中间价', g2,
                                 S['汇率'], 0, 4, 'week', WEEK_YEARS))
    charts = [c for c in charts if c]
    return {'id': p, 'name': '氢氧化锂价格', 'axes': [WEEK_LABEL],
            'axisTypes': ['week'], 'monthFirst': WEEK_FIRST,
            'asOf': max(max(m) for m in S.values()).strftime('%Y-%m-%d'),
            'charts': charts,
            'note': ('单位：电氢/电碳/价差 元/吨，CIF 美元/公斤，汇率 元/美元。源表为日频，'
                     '按 ISO 周取周内最后一个交易日投影到周轴（基准年 2026）；仅展示近五年（2022–2026）。'
                     'CIF 与中日价差源表含重复日期，已按最后一条去重。')}


def build_capacity(ws):
    p = 'lc_lih_capacity'
    m = {
        '冶炼': read_series_ordered(ws, 'A', 'B'),
        '苛化': read_series_ordered(ws, 'A', 'C'),
        '总': read_series_ordered(ws, 'A', 'D'),
        '江西': read_series_ordered(ws, 'O', 'P'),
        '其他': read_series_ordered(ws, 'O', 'Q'),
        '四川': read_series_ordered(ws, 'O', 'R'),
    }
    a = {
        '国内冶炼': read_series_ordered(ws, 'F', 'G', 2021),
        '国内苛化': read_series_ordered(ws, 'F', 'H', 2021),
        '国内总': read_series_ordered(ws, 'F', 'I', 2021),
        '海外冶炼': read_series_ordered(ws, 'F', 'J', 2021),
        '海外苛化': read_series_ordered(ws, 'F', 'K', 2021),
        '海外总': read_series_ordered(ws, 'F', 'L', 2021),
        '全球总': read_series_ordered(ws, 'F', 'M', 2021),
    }
    charts = []
    g1 = '月度产能（吨）'
    for key, title, k in [
        ('smelt', '冶炼端产能', '冶炼'), ('caust', '苛化端产能', '苛化'),
        ('total', '总产能', '总'), ('jx', '江西产能', '江西'),
        ('sc', '四川产能', '四川'), ('other', '其他地区产能', '其他'),
    ]:
        c = seasonal_chart(p + '-' + key, title, g1, m[k], 0, 0, 'month', MONTH_YEARS)
        if c: charts.append(c)
    g2 = '年度产能（万吨）'
    c = date_axis_chart(p + '-annual', '国内 / 海外 / 全球总产能',
                        g2, [(n, a[n], col) for n, col in
                             [('国内总', '#FF0000'), ('海外总', '#0070C0'), ('全球总', '#7F7F7F')]])
    if c: charts.append(c)
    c = date_axis_chart(p + '-annual2', '分工艺年度产能（国内 / 海外）',
                        g2, [(n, a[n], col) for n, col in
                             [('国内冶炼', '#FF0000'), ('国内苛化', '#FFC000'),
                              ('海外冶炼', '#0070C0'), ('海外苛化', '#5B9BD5')]])
    if c: charts.append(c)
    mx = max([max(v) for v in m.values()] + [max(v) for v in a.values()])
    return {'id': p, 'name': '氢氧化锂产能',
            'axes': [MONTH_AXIS, sorted({('%04d-01-01' % d.year) for v in a.values() for d in v})],
            'axisTypes': ['month', 'date'], 'monthFirst': [],
            'asOf': mx.strftime('%Y-%m-%d'), 'charts': charts,
            'note': ('月度产能源表单位：吨；年度产能源表单位虽标注「实物吨」，'
                     '但数值量级与月度产能×12 一致（2025 国内总 76.85 ↔ 2025-12 月产能 7.02 万吨），'
                     '按万吨呈现。江西/四川/其他为月度分地区产能。')}


def build_output(ws):
    p = 'lc_lih_output'
    w = {
        '冶炼': read_series_ordered(ws, 'A', 'B'),
        '苛化': read_series_ordered(ws, 'A', 'C'),
        '总': read_series_ordered(ws, 'A', 'D'),
        '表需': read_series_ordered(ws, 'A', 'E'),
    }
    m = {
        '总': read_series_ordered(ws, 'G', 'H'),
        '冶炼': read_series_ordered(ws, 'G', 'I'),
        '苛化': read_series_ordered(ws, 'G', 'J'),
        '江西': read_series_ordered(ws, 'G', 'K'),
        '四川': read_series_ordered(ws, 'G', 'L'),
        '浙江': read_series_ordered(ws, 'G', 'M'),
        '广西': read_series_ordered(ws, 'G', 'N'),
        '湖南': read_series_ordered(ws, 'G', 'Q'),
        '其他': read_series_ordered(ws, 'G', 'R'),
        'CR5': read_series_ordered(ws, 'G', 'S'),
    }
    charts = []
    g1 = '周度产量（实物吨）'
    for key, title, k in [('smelt', '冶炼端产量', '冶炼'), ('caust', '苛化端产量', '苛化'),
                          ('total', '总产量', '总'), ('demand', '表需', '表需')]:
        c = seasonal_chart(p + '-w-' + key, title, g1, w[k], 0, 0, 'week', WEEK_YEARS)
        if c: charts.append(c)
    g2 = '月度产量（吨）'
    for key, title, k, nd in [
        ('total', '总产量', '总', 0), ('smelt', '冶炼端产量', '冶炼', 0),
        ('caust', '苛化端产量', '苛化', 0), ('jx', '江西', '江西', 0),
        ('sc', '四川', '四川', 0), ('zj', '浙江', '浙江', 0),
        ('gx', '广西', '广西', 0), ('hn', '湖南', '湖南', 0),
        ('other', '其他地区', '其他', 0), ('cr5', 'CR5 占比（%）', 'CR5', 2),
    ]:
        c = seasonal_chart(p + '-m-' + key, title, g2, m[k], 1, nd, 'month', PROD_MONTH_YEARS)
        if c: charts.append(c)
    mx = max([max(v) for v in w.values()] + [max(v) for v in m.values()])
    return {'id': p, 'name': '氢氧化锂产量',
            'axes': [WEEK_LABEL, MONTH_AXIS], 'axisTypes': ['week', 'month'],
            'monthFirst': WEEK_FIRST, 'asOf': mx.strftime('%Y-%m-%d'),
            'charts': charts,
            'note': ('周度产能源表仅 2026-05-21 起，为单年走势，按 ISO 周对齐到 52 槽基准年轴；'
                     '月度产量展示近五年（2022–2026）。山东、河北产量长期为 0，未单列。')}


def build_inventory(ws):
    p = 'lc_lih_inventory'
    w = {
        '上游': read_series_ordered(ws, 'A', 'B'),
        '下游': read_series_ordered(ws, 'A', 'C'),
        '其他': read_series_ordered(ws, 'A', 'D'),
        '总': read_series_ordered(ws, 'A', 'E'),
    }
    m = {
        '下游': read_series_ordered(ws, 'H', 'I'),
        '冶炼厂': read_series_ordered(ws, 'H', 'J'),
        '总': read_series_ordered(ws, 'H', 'K'),
    }
    charts = []
    g1 = '周度样本库存（实物吨）'
    for key, title, k in [('up', '上游锂盐厂', '上游'), ('down', '下游三元材料厂', '下游'),
                          ('other', '其他', '其他'), ('total', '总库存', '总')]:
        c = seasonal_chart(p + '-w-' + key, title, g1, w[k], 0, 0, 'week', WEEK_YEARS)
        if c: charts.append(c)
    g2 = '月度库存（实物吨）'
    for key, title, k in [('down', '下游', '下游'), ('smelter', '冶炼厂', '冶炼厂'),
                          ('total', '总计', '总')]:
        c = seasonal_chart(p + '-m-' + key, title, g2, m[k], 1, 0, 'month', MONTH_YEARS)
        if c: charts.append(c)
    mx = max([max(v) for v in w.values()] + [max(v) for v in m.values()])
    return {'id': p, 'name': '氢氧化锂库存',
            'axes': [WEEK_LABEL, MONTH_AXIS], 'axisTypes': ['week', 'month'],
            'monthFirst': WEEK_FIRST, 'asOf': mx.strftime('%Y-%m-%d'),
            'charts': charts,
            'note': ('周度样本库存源表仅 2026-05-21 起；月度库存展示 2023 年以来。'
                     '源表 F:G 的「上游/中游占比」仅 3 行、不成序列，未纳入。')}


def build_cost(ws):
    p = 'lc_lih_cost'
    d = {
        '出口利润': read_series_ordered(ws, 'A', 'B'),
        '进口0': read_series_ordered(ws, 'A', 'C'),
        '进口5': read_series_ordered(ws, 'A', 'D'),
        '冶炼利润': read_series_ordered(ws, 'A', 'E'),
        '苛化利润': read_series_ordered(ws, 'A', 'F'),
        '苛化成本': read_series_ordered(ws, 'A', 'G'),
        '冶炼成本': read_series_ordered(ws, 'A', 'H'),
    }
    m = {
        '锂辉石成本': read_series_ordered(ws, 'J', 'K'),
        '苛化加工费': read_series_ordered(ws, 'J', 'L'),
        '碳化加工费': read_series_ordered(ws, 'J', 'M'),
    }
    charts = []
    g1 = '日度利润 / 成本（元/吨，按周取周末值）'
    for key, title, k in [
        ('export', '出口利润', '出口利润'), ('imp0', '进口盈亏（零关税来源国）', '进口0'),
        ('imp5', '进口盈亏（5%关税来源国）', '进口5'),
        ('smelt', '生产利润（冶炼法）', '冶炼利润'),
        ('caust', '生产利润（苛化法）', '苛化利润'),
        ('scost', '现金生产成本（冶炼法）', '冶炼成本'),
        ('ccost', '现金生产成本（苛化法）', '苛化成本'),
    ]:
        c = seasonal_chart(p + '-' + key, title, g1, d[k], 0, 0, 'week', WEEK_YEARS)
        if c: charts.append(c)
    g2 = '月度成本 / 加工费（元/吨）'
    for key, title, k in [
        ('spod', '锂辉石-氢氧化锂生产成本', '锂辉石成本'),
        ('cau', '苛化加工费（碳酸锂-氢氧化锂）', '苛化加工费'),
        ('carb', '碳化加工费（氢氧化锂-碳酸锂）', '碳化加工费'),
    ]:
        c = seasonal_chart(p + '-' + key, title, g2, m[k], 1, 0, 'month', MONTH_YEARS)
        if c: charts.append(c)
    mx = max([max(v) for v in d.values()] + [max(v) for v in m.values()])
    return {'id': p, 'name': '氢氧化锂成本利润',
            'axes': [WEEK_LABEL, MONTH_AXIS], 'axisTypes': ['week', 'month'],
            'monthFirst': WEEK_FIRST, 'asOf': mx.strftime('%Y-%m-%d'),
            'charts': charts,
            'note': ('日频指标按 ISO 周取周内最后一个交易日投影到周轴；'
                     '进口盈亏仅 2026 年有值；现金成本源表含 2014 年孤立值，'
                     '已随近五年（2022 年起）窗口过滤。')}


# ---------------------------------------------------------------- 写盘
def write_all(ds_list):
    meta_path = os.path.join(DATA, 'meta.json')
    order = []
    if os.path.exists(meta_path):
        try:
            with open(meta_path, encoding='utf-8') as f:
                order = [x['id'] for x in json.load(f).get('datasets', [])]
        except Exception:
            pass
    for d in ds_list:
        if d['id'] not in order:
            order.append(d['id'])
    all_ds = []
    for did in order:
        hit = next((x for x in ds_list if x['id'] == did), None)
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
        p = os.path.join(DATA, d['id'] + '.json')
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, separators=(',', ':'))
        print('[ok] %s.json  charts=%d  asOf=%s  %.0fKB'
              % (d['id'], len(d['charts']), d['asOf'], os.path.getsize(p) / 1024))
    with open(os.path.join(DATA, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('// 自动生成，勿手改。build_lc_lih_charts.py @ ' + stamp + '\n')
        f.write('window.CHART_DATA = ')
        json.dump({'updated': stamp, 'datasets': all_ds}, f,
                  ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    print('[ok] data.js datasets=%d' % len(all_ds))
    meta = {'updated': stamp,
            'datasets': [{'id': d['id'], 'name': d['name'],
                          'charts': len(d['charts']), 'asOf': d['asOf']}
                         for d in all_ds]}
    with open(meta_path, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print('[ok] meta.json datasets=%d' % len(meta['datasets']))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    ap.add_argument('--src', default=SRC)
    args = ap.parse_args()
    wb = openpyxl.load_workbook(args.src, data_only=True)
    try:
        s = load(wb)
        ds_list = [build_price(s['price']), build_capacity(s['cap']),
                   build_output(s['out']), build_inventory(s['inv']),
                   build_cost(s['cost'])]
    finally:
        wb.close()
    for ds in ds_list:
        print('\n== %s (%s) asOf=%s 图=%d' %
              (ds['name'], ds['id'], ds['asOf'], len(ds['charts'])))
        for c in ds['charts']:
            n0 = sum(1 for v in c['series'][0]['data'] if v is not None)
            print('   [%s] %-34s 年=%s 主系列非空=%d table=%s' %
                  (c['group'], c['title'], ','.join(x['name'] for x in c['series']),
                   n0, 'Y' if c.get('table') else '-'))
    if args.check:
        return
    write_all(ds_list)


if __name__ == '__main__':
    main()
