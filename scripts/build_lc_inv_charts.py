# -*- coding: utf-8 -*-
"""
build_lc_inv_charts.py — 「碳酸锂库存（钢联）」数据集生成器（季节性年对比版）

源：《下游周度数据.xlsx》的「钢联」sheet（我的钢铁网 Mysteel 周度库存调研，周三口径）
  表结构（固定 8 行表头 + 数据行倒序）：
    第1行  标题「钢联数据」
    第2行  指标名称：碳酸锂：库存：中国（周） / 贸易商 / 冶炼厂 / 材料厂及终端
    第3行  单位：吨
    第4行  数据来源、第5行 指标编码、第6行 频度、第7行 时间区间、第8行 指标描述
    第9行起 数据：A=日期（新→旧倒序），B..E=四个分项
  注意：同工作簿的「正极 / 表需 / 电芯 / 碳酸锂产量」等 sheet 是 **SMM** 口径，
        与本数据集（钢联）不是同一来源，刻意不纳入。

产出 lithium-charts 一个 dataset「碳酸锂库存」(id=lc_inv)：
  · 横轴 = MM-DD 季节性轴（出现过的周度日期并集，按 (月,日) 升序）
  · 每个指标按年份叠线（最新年红色、前一年蓝…），与钢材站同构；
    2026 年起步只有 1 年 → 单条红线；明年起自动变成 2 年对比，无需改代码。
  · 组1「碳酸锂库存（周）」：4 张 —— 库存总量 / 贸易商 / 冶炼厂 / 材料厂及终端
  · 组2「库存结构（派生）」：3 张 —— 各环节占总量比重（%），由组1 四列算出
  · 汇总表：4 个绝对量指标的 本期/上期/环比/较年初/较年初%

单位：源表即「吨」，本站直接按吨呈现（不换算、取整），图上口径已在 note 注明；占比图另按 % 显示。

用法：
  python scripts/build_lc_inv_charts.py            # 生成 data/lc_inv.json + data.js + meta.json
  python scripts/build_lc_inv_charts.py --check    # 只看统计与结构，不落盘
"""
import argparse
import datetime
import json
import os

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SRC = r'C:\Users\Administrator\Nutstore\1\我的坚果云\碳酸锂\下游周度数据.xlsx'
SHEET = '钢联'

DS_ID = 'lc_inv'
DS_NAME = '碳酸锂库存'

MAX_YEARS = 5
SCALE = 1.0           # 单位＝吨（源表原值，不做换算）
ND = 0                # 吨级取整（源值本身即整数吨）

# ---- 指标定义：(源表列号, 显示名, 简称) ----
# 列号固定（B=1 .. E=4，0-based），但下方 parse() 会按表头名重新定位并校验。
ITEMS = [
    (1, '碳酸锂库存总量', '总量'),
    (2, '贸易商库存', '贸易商'),
    (3, '冶炼厂库存', '冶炼厂'),
    (4, '材料厂及终端库存', '材料厂'),
]
# 结构占比（分子列号, 显示名）；分母恒为总量列(1)
RATIOS = [
    (3, '冶炼厂库存占比'),
    (4, '材料厂及终端占比'),
    (2, '贸易商库存占比'),
]

ERR_TOKENS = {'#N/A', '#N/A!', '#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', ''}


def clean_num(v):
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s in ERR_TOKENS:
            return None
        try:
            v = float(s)
        except ValueError:
            return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        if v != v:
            return None
        return float(v)
    return None


def rnd(v, nd=ND):
    """按 nd 位四舍五入；nd=0 时返回 int（避免 JSON 里出现 103330.0）"""
    if v is None:
        return None
    return int(round(v)) if nd == 0 else round(v, nd)


def series_style(i, n):
    latest = (i == n - 1)
    prev = (i == n - 2)
    if latest:
        color = '#FF0000'
    elif prev:
        color = '#0070C0'
    elif i == n - 3:
        color = '#A5A5A5'
    else:
        color = '#9DC3E6'
    return {
        'color': color,
        'width': 1.5,
        'dash': 'dash' if i <= n - 4 else 'solid',
        'smooth': False,
        'marker': 'circle' if latest else 'none',
    }


def parse(wb):
    """返回 (headers, dates, data) ；data[col] = {date: 吨值}"""
    if SHEET not in wb.sheetnames:
        raise SystemExit('源工作簿缺少「%s」sheet' % SHEET)
    ws = wb[SHEET]
    rows = list(ws.iter_rows(values_only=True))

    # 定位「指标名称」表头行
    hdr_i = None
    for i, r in enumerate(rows[:20]):
        if r and r[0] is not None and str(r[0]).strip() == '指标名称':
            hdr_i = i
            break
    if hdr_i is None:
        raise SystemExit('未找到「指标名称」表头行')
    header = rows[hdr_i]

    # 校验列号：ITEMS 里写死的列必须与表头一致，不一致直接报错（防止源表插列后静默错位）
    mism = []
    for col, name, _ in ITEMS:
        got = header[col] if col < len(header) else None
        if got is None or str(got).strip() == '':
            mism.append('第%d列(%s) 表头为空' % (col, name))
    if mism:
        raise SystemExit('列定位失败：' + '；'.join(mism))
    print('[ok] 表头校验通过：')
    for col, name, short in ITEMS:
        print('     第%d列 = %s（简称 %s）' % (col, str(header[col]).strip(), short))

    data = {col: {} for col, _, _ in ITEMS}
    dates = []
    for r in rows[hdr_i + 1:]:
        if not r:
            continue
        ds_raw = r[0]
        if ds_raw is None:
            continue
        if isinstance(ds_raw, datetime.datetime):
            d = ds_raw.date()
        elif isinstance(ds_raw, datetime.date):
            d = ds_raw
        else:
            try:
                d = datetime.datetime.strptime(str(ds_raw)[:10], '%Y-%m-%d').date()
            except ValueError:
                continue
        got_any = False
        for col, _, _ in ITEMS:
            v = clean_num(r[col] if col < len(r) else None)
            if v is not None:
                data[col][d] = v
                got_any = True
        if got_any:
            dates.append(d)
    dates = sorted(set(dates))
    return header, dates, data


def build_dataset():
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    try:
        header, dates, data = parse(wb)
    finally:
        wb.close()
    if not dates:
        return None, {}

    # 勾稽自检：总量 = 贸易商 + 冶炼厂 + 材料厂及终端
    bad = []
    for d in dates:
        tot = data[1].get(d)
        parts = [data[c].get(d) for c, _, _ in ITEMS[1:]]
        if tot is None or any(p is None for p in parts):
            continue
        s = sum(parts)
        if abs(s - tot) > max(1.0, tot * 1e-6):
            bad.append((d, tot, s))
    checks = {'sum_rows': len([d for d in dates if data[1].get(d) is not None]),
              'sum_mismatch': len(bad)}
    if bad:
        print('[warn] 总量与分项不勾稽 %d 期，前 3 期：%s' % (len(bad), bad[:3]))
    else:
        print('[ok] 勾稽自检通过：总量 = 贸易商 + 冶炼厂 + 材料厂及终端')

    # —— 季节性横轴：出现过的 MM-DD 并集 ——
    md_set = {(d.month, d.day) for d in dates}
    md_axis = ['%02d-%02d' % (mm, dd) for mm, dd in sorted(md_set)]

    month_first, seen_mm = [], set()
    for mm, dd in sorted(md_set):
        if mm not in seen_mm:
            seen_mm.add(mm)
            month_first.append('%02d-%02d' % (mm, dd))

    years = sorted({d.year for d in dates})
    if len(years) > MAX_YEARS:
        years = years[-MAX_YEARS:]
    n = len(years)

    def seasonal(series_map, scale=1.0, nd=2):
        byyear = {}
        for d in dates:
            v = series_map.get(d)
            if v is None or d.year not in years:
                continue
            byyear.setdefault(d.year, {})['%02d-%02d' % (d.month, d.day)] = rnd(v * scale, nd)
        return byyear

    charts = []

    def add_chart(key, title, group, byyear, fmt='v'):
        present = [y for y in years if y in byyear]
        if not present:
            return
        ser = []
        for y in present:
            st = series_style(years.index(y), n)
            ser.append({
                'name': str(y), 'color': st['color'], 'width': st['width'],
                'dash': st['dash'], 'smooth': st['smooth'], 'marker': st['marker'],
                'data': [byyear[y].get(md) for md in md_axis],
            })
        charts.append({'key': key, 'title': title, 'type': 'line', 'axis': 0,
                       'group': group, 'series': ser, 'fmt': fmt})

    # —— 组1：碳酸锂库存（周）——
    for col, name, short in ITEMS:
        add_chart('lc_inv-%s' % short, name + '（周）', '碳酸锂库存（周）',
                  seasonal(data[col], SCALE, ND))

    # —— 组2：库存结构（派生）——
    ratio_maps = {}
    for col, name in RATIOS:
        m = {}
        for d in dates:
            num, den = data[col].get(d), data[1].get(d)
            if num is not None and den not in (None, 0):
                m[d] = num / den * 100.0
        ratio_maps[col] = m
    for col, name in RATIOS:
        add_chart('lc_inv-%s' % name.replace('库存占比', '占比'), name + '（%）',
                  '库存结构（派生）', seasonal(ratio_maps[col], 1.0, 2))

    # —— 汇总表（仅 4 个绝对量指标；占比不混入，单位不同）——
    last = dates[-1]
    prev = last - datetime.timedelta(days=7)
    year_start = None
    for d in dates:
        if d.year == last.year:
            year_start = d
            break

    def val_at(series_map, target):
        best = None
        for d, v in series_map.items():
            if d <= target and (best is None or d > best[0]):
                best = (d, v)
        return best[1] if best else None

    def date_at(series_map, target):
        """target 当天或之前最近一期真实有数据的日子。
        源表偶有缺周（如 2026-10-09），此时「上期」必须显示实际取到的日期，
        否则标签写 10-09、数值却来自 10-02，会误导环比口径。"""
        best = None
        for d in series_map:
            if d <= target and (best is None or d > best):
                best = d
        return best

    def val_first(series_map, year):
        for d in sorted(series_map):
            if d.year == year:
                return series_map[d]
        return None

    columns, cur = [], []
    for col, name, short in ITEMS:
        columns.append({'key': 'c%d' % len(columns), 'label': name, 'group': '碳酸锂库存'})
        c = val_at(data[col], last)
        p = val_at(data[col], prev)
        ys = val_first(data[col], last.year)
        cur.append({
            'cur': c, 'prev': p, 'ys': ys,
            'vs_ys': (c - ys) if (c is not None and ys is not None) else None,
            'vs_ys_pct': (round((c - ys) / abs(ys) * 100, 1)
                          if (c is not None and ys not in (None, 0)) else None),
        })

    g = lambda v: rnd(v * SCALE)
    rows = {
        '本期': [g(x['cur']) for x in cur],
        '上期': [g(x['prev']) for x in cur],
        '环比': [(None if (x['cur'] is None or x['prev'] is None)
                  else rnd((x['cur'] - x['prev']) * SCALE)) for x in cur],
        '较年初': [g(x['vs_ys']) for x in cur],
        '较年初%': [x['vs_ys_pct'] for x in cur],
    }
    # 上期基准：源表若缺周，则回退到最近的真实期，并同步用于标签显示
    prev_actual = date_at(data[1], prev) or prev

    summary = {
        'columns': columns, 'rows': rows,
        'rowOrder': ['本期', '上期', '环比', '较年初', '较年初%'],
        'currentWeek': last.strftime('%Y-%m-%d'),
        'previousWeek': prev_actual.strftime('%Y-%m-%d'),
        'unit': '吨',
    }

    # —— 卡片指标条（本期 / 上期 / 环比 / 较年初）——
    for ch in charts:
        short = ch['key'].split('-', 1)[1]
        col = next((c for c, _, s in ITEMS if s == short), None)
        if col is None:                       # 占比图不挂 card table
            continue
        c = val_at(data[col], last)
        p = val_at(data[col], prev)
        ys = val_first(data[col], last.year)
        if c is None:
            continue
        ch['table'] = {
            'labels': ['本期', '上期', '环比', '较年初'],
            'kinds': ['val', 'val', 'delta', 'delta'],
            'values': [g(c), g(p),
                       (None if p is None else rnd(c - p)),
                       (None if ys is None else rnd(c - ys))],
        }

    # —— 口径备注 ——
    missing = []
    wk = {d for d in dates}
    d0 = dates[0]
    while d0 <= last:
        if d0 not in wk:
            missing.append(d0.strftime('%Y-%m-%d'))
        d0 += datetime.timedelta(days=7)
    dup = []
    for i in range(1, len(dates)):
        a, b = dates[i - 1], dates[i]
        if all(data[c].get(a) == data[c].get(b) for c, _, _ in ITEMS):
            dup.append('%s 与 %s' % (a.strftime('%m-%d'), b.strftime('%m-%d')))
    note_parts = ['单位：吨（源表原值，未做换算）',
                  '口径：Mysteel 周度库存调研，样本 207 家（冶炼厂 74 / 贸易商 61 / 材料厂及终端 72），货权口径']
    if missing:
        note_parts.append('源表缺周：' + '、'.join(missing))
    span = (last - prev_actual).days
    # 只在「跨期超过一周」时提示：说明源表缺周、环比不是单周口径。
    # （源表日期记的是下载日/周五，与 Mysteel 调研周标签（周三）相差 1~2 天，
    #   span 偶尔 < 7 属正常，不能当成缺周报警。）
    if span > 7:
        note_parts.append('本期「环比」实为 %s→%s 共 %d 天变动（非单周）'
                          % (prev_actual.strftime('%m-%d'), last.strftime('%m-%d'), span))
    if dup:
        note_parts.append('源表两期数值完全相同：' + '、'.join(dup) + '（按源表原样呈现，未插补）')
    note = '｜'.join(note_parts)

    as_of = last.strftime('%Y-%m-%d')
    ds = {
        'id': DS_ID, 'name': DS_NAME,
        'axes': [md_axis], 'monthFirst': month_first,
        'asOf': as_of, 'charts': charts, 'summary': summary,
        'note': note, 'unit': 'mixed',
    }
    return ds, checks


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    args = ap.parse_args()

    ds, checks = build_dataset()
    if ds is None:
        print('[warn] 无数据')
        return

    if args.check:
        print('\n== %s(%s) | asOf %s | charts %d | axis %d | years %s'
              % (ds['name'], ds['id'], ds['asOf'], len(ds['charts']),
                 len(ds['axes'][0]),
                 ','.join(sorted({s['name'] for s in ds['charts'][0]['series']}))))
        for c in ds['charts']:
            yrs = [s['name'] for s in c.get('series', [])]
            npt = len([v for v in c['series'][0]['data'] if v is not None]) if c.get('series') else 0
            print('   [%s] %s | 年份=%s | 非空点=%d | table=%s'
                  % (c.get('group'), c['title'], ','.join(yrs), npt,
                     'Y' if c.get('table') else '-'))
        print('   summary cols:', [c['label'] for c in ds['summary']['columns']])
        print('   rows:', {k: v for k, v in ds['summary']['rows'].items()})
        print('   monthFirst:', ds['monthFirst'])
        print('   note:', ds['note'])
        return

    # ---- 合并并重写 data.js + meta.json ----
    meta_path = os.path.join(DATA, 'meta.json')
    order = []
    if os.path.exists(meta_path):
        try:
            with open(meta_path, encoding='utf-8') as f:
                order = [d['id'] for d in json.load(f).get('datasets', [])]
        except Exception:
            pass
    if ds['id'] not in order:
        order.append(ds['id'])

    all_ds = []
    for did in order:
        if did == ds['id']:
            all_ds.append(ds)
        else:
            p = os.path.join(DATA, did + '.json')
            if os.path.exists(p):
                with open(p, encoding='utf-8') as f:
                    all_ds.append(json.load(f))
                print('[ok] 读取既有 %s.json' % did)
            else:
                print('[warn] 缺少 %s.json，跳过' % did)

    stamp = datetime.datetime.now().isoformat(timespec='seconds')
    p = os.path.join(DATA, ds['id'] + '.json')
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(ds, f, ensure_ascii=False, separators=(',', ':'))
    print('[ok] 写 %s.json (%.0f KB) charts=%d asOf=%s'
          % (ds['id'], os.path.getsize(p) / 1024, len(ds['charts']), ds['asOf']))

    with open(os.path.join(DATA, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('// 自动生成，勿手改。build_lc_inv_charts.py @ ' + stamp + '\n')
        f.write('window.CHART_DATA = ')
        json.dump({'updated': stamp, 'datasets': all_ds}, f,
                  ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    print('[ok] 写 data.js (datasets=%d)' % len(all_ds))

    meta = {'updated': stamp,
            'datasets': [{'id': d['id'], 'name': d['name'],
                          'charts': len(d['charts']), 'asOf': d['asOf']}
                         for d in all_ds]}
    with open(meta_path, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print('[ok] 写 meta.json datasets=%d' % len(meta['datasets']))


if __name__ == '__main__':
    main()
