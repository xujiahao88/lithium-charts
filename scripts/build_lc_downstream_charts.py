# -*- coding: utf-8 -*-
"""
build_lc_downstream_charts.py — 碳酸锂下游三张 SMM 表 → lithium-charts 数据集

源：《下游周度数据.xlsx》的三个 sheet（均为 **SMM 上海有色网** 口径，与「钢联」sheet 不同源）
  · 正极         ：三元材料/磷酸铁锂 周度产量 + 两环节厂内库存（吨）
  · 电芯         ：动力电芯周度产量 合计/铁锂/三元（GWh）
  · 碳酸锂产量   ：碳酸锂周度产量 总计/锂辉石/锂云母/盐湖/回收料（实物吨）
                    ⚠ 该 sheet 同时含「SMM 碳酸锂样本库存」4 列（G~J），按需求**不纳入**，
                      库存仍以「碳酸锂库存」tab 的 Mysteel 口径为准，避免两套口径混读。

产出 3 个 dataset：
  · lc_output  碳酸锂产量  —— 5 图（1 组），勾稽：总计 = 辉石+云母+盐湖+回收料
  · lc_cathode 正极材料    —— 4 图（2 组：产量 / 厂库存）
  · lc_cell    动力电芯    —— 3 图（1 组），勾稽：合计 = 铁锂 + 三元

横轴口径（与「碳酸锂库存」有意不同）：
  这三张表的 A 列日期是**下载日**而非周标签，实测间隔在 5/6/7/8/14 天之间漂移，
  年份之间还会左右挪 1~2 天。若直接按原始 (月,日) 建槽位，4 年叠线会横向错开、
  轴长膨胀到 160+ 槽位（碳酸锂产量 164 期 → 163 槽位），季节性完全看不出来。
  因此这里统一把每期归一到**该 ISO 周的周四**再取 (月,日)：同一周号跨年落点一致，
  轴收敛到 53~56 槽位，年份间可比。真实日期仍用于 asOf 与「较年初」判定。

用法：
  python scripts/build_lc_downstream_charts.py            # 生成 3 个 json + 重建 data.js/meta.json
  python scripts/build_lc_downstream_charts.py --check    # 只打印结构统计，不落盘
  python scripts/build_lc_downstream_charts.py --only lc_cell
"""
import argparse
import datetime
import json
import os

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SRC = r'C:\Users\Administrator\Nutstore\1\我的坚果云\碳酸锂\下游周度数据.xlsx'

MAX_YEARS = 5
ERR_TOKENS = {'#N/A', '#N/A!', '#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', ''}

# tab 顺序（上游→下游；「碳酸锂库存」保持首位，是日常主看的那一页）
ORDER = ['lc_inv', 'lc_output', 'lc_cathode', 'lc_cell']

# ---------------------------------------------------------------- 数据集定义
# items: (源表列号 0-based, 显示名, 简称, 表头关键字用于校验)
SPECS = [
    {
        'id': 'lc_output', 'name': '碳酸锂产量', 'sheet': '碳酸锂产量',
        'unit': '吨', 'nd': 0,
        'groups': [('碳酸锂产量（周）', [
            (1, '碳酸锂产量总计', '总计', '碳酸锂产量: 总计'),
            (2, '锂辉石产', '辉石', '锂辉石产'),
            (3, '锂云母产', '云母', '锂云母产'),
            (4, '盐湖产', '盐湖', '盐湖产'),
            (5, '回收料产', '回收料', '回收料产'),
        ])],
        # 总计 = 各工艺路线之和
        'checks': [(1, [2, 3, 4, 5])],
        'src_note': 'SMM 碳酸锂周度产量，含总计与 4 条工艺路线（实物吨）',
    },
    {
        'id': 'lc_cathode', 'name': '正极材料', 'sheet': '正极',
        'unit': '吨', 'nd': 0,
        'groups': [
            ('正极材料产量（周）', [
                (1, '三元材料产量', '三元', '三元材料周度产量'),
                (2, '磷酸铁锂产量', '铁锂', '磷酸铁锂周度产量'),
            ]),
            ('正极材料厂库存（周）', [
                (3, '三元材料厂库存', '三元库存', '三元材料周度库存'),
                (4, '铁锂厂库存', '铁锂库存', '磷酸铁锂周度库存'),
            ]),
        ],
        'checks': [],
        'src_note': 'SMM 正极材料周度产量与厂内库存（吨）',
    },
    {
        'id': 'lc_cell', 'name': '动力电芯', 'sheet': '电芯',
        'unit': 'GWh', 'nd': 1,
        'groups': [('动力电芯产量（周）', [
            (1, '动力电芯产量合计', '合计', 'SMM动力电芯周度产量'),
            (2, '铁锂电芯', '铁锂', '铁锂'),
            (3, '三元电芯', '三元', '三元'),
        ])],
        # 合计 = 铁锂 + 三元
        'checks': [(1, [2, 3])],
        'src_note': 'SMM 动力电芯周度产量，按铁锂/三元拆分（GWh）',
    },
]


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


def rnd(v, nd=0):
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
    return {'color': color, 'width': 1.5,
            'dash': 'dash' if i <= n - 4 else 'solid',
            'smooth': False, 'marker': 'circle' if latest else 'none'}


def iso_week(d):
    """ISO 周号（1~53）。跨年周按 ISO 年归属，避免 12-31/01-01 错周。"""
    return d.isocalendar()[1]


def iso_year(d):
    """ISO 年。12-29~12-31 常属次年 W01，用日历年会把它们错排到本年 W01，
    与本年 1 月初的同周点撞槽（实测正极/电芯各有 3~4 个点因此被覆盖）。"""
    return d.isocalendar()[0]


def week_label(base_year, week):
    """把周号映射到「基准年该周的周四」的 (月,日)。
    直接对原始日期取 (月,日) 不行：每年同一周号的周四会左右漂 1~2 天，
    4 年叠线后槽位会膨胀到 160+（碳酸锂产量实测 164 期→164 槽），季节性看不出来。
    锚定到同一年份后，同周号恒为同一槽位，轴收敛到 52~53。"""
    for w in (week, week - 1, 52):
        try:
            thu = datetime.date.fromisocalendar(base_year, w, 4)
            return '%02d-%02d' % (thu.month, thu.day)
        except ValueError:
            continue
    return None


def parse_sheet(wb, spec):
    """返回 {col: {date: value}}，并做表头/单位校验"""
    if spec['sheet'] not in wb.sheetnames:
        raise SystemExit('源工作簿缺少「%s」sheet' % spec['sheet'])
    ws = wb[spec['sheet']]
    rows = list(ws.iter_rows(values_only=True))

    hdr_i = None
    for i, r in enumerate(rows[:12]):
        if r and r[0] is not None and str(r[0]).strip() == '指标名称':
            hdr_i = i
            break
    if hdr_i is None:
        raise SystemExit('[%s] 未找到「指标名称」表头行' % spec['sheet'])
    header = rows[hdr_i]

    items = [it for _, its in spec['groups'] for it in its]

    # 列定位校验：表头必须含关键字，防止源表插列后静默错位
    bad = []
    for col, name, short, kw in items:
        got = str(header[col]).strip() if (col < len(header) and header[col] is not None) else ''
        if not got or kw not in got:
            bad.append('第%d列期望含「%s」，实际「%s」' % (col, kw, got))
    if bad:
        raise SystemExit('[%s] 列定位失败：%s' % (spec['sheet'], '；'.join(bad)))

    # 单位行（位置各表不一，按 A 列关键词找）
    unit_row = None
    for r in rows[:hdr_i + 6]:
        if r and r[0] is not None and str(r[0]).strip() == '单位':
            unit_row = r
            break
    units = {}
    if unit_row:
        for col, name, short, kw in items:
            units[col] = str(unit_row[col]).strip() if col < len(unit_row) and unit_row[col] else ''
    print('[ok] %s 表头校验通过' % spec['sheet'])
    for col, name, short, kw in items:
        u = units.get(col, '')
        print('     第%d列 %-18s 单位=%s' % (col, name, u or '(空)'))

    data = {col: {} for col, _, _, _ in items}
    for r in rows[hdr_i + 1:]:
        if not r:
            continue
        raw = r[0]
        if isinstance(raw, datetime.datetime):
            d = raw.date()
        elif isinstance(raw, datetime.date):
            d = raw
        elif isinstance(raw, str):
            try:
                d = datetime.datetime.strptime(raw.strip()[:10], '%Y-%m-%d').date()
            except ValueError:
                continue
        else:
            continue
        for col, _, _, _ in items:
            v = clean_num(r[col] if col < len(r) else None)
            if v is not None:
                data[col][d] = v
    return data


def build_one(wb, spec):
    data = parse_sheet(wb, spec)
    items = [it for _, its in spec['groups'] for it in its]
    nd = spec['nd']

    all_dates = sorted({d for m in data.values() for d in m})
    if not all_dates:
        print('[warn] %s 无数据' % spec['name'])
        return None
    last = all_dates[-1]

    # —— 勾稽自检 ——
    check_note = []
    for tot_col, part_cols in spec.get('checks', []):
        bad = []
        for d in sorted(data[tot_col]):
            tot = data[tot_col].get(d)
            parts = [data[c].get(d) for c in part_cols]
            if tot is None or any(p is None for p in parts):
                continue
            s = sum(parts)
            tol = 1.0 if nd == 0 else 0.05
            if abs(s - tot) > tol:
                bad.append((d.strftime('%Y-%m-%d'), round(tot, 2), round(s, 2)))
        tname = next(n for c, n, _, _ in items if c == tot_col)
        if bad:
            print('[warn] %s：%s 与分项不勾稽 %d 期，前 3 期：%s'
                  % (spec['name'], tname, len(bad), bad[:3]))
            check_note.append('%s 与分项合计有 %d 期不等（最大差 %.2f %s，源表口径差异，未做调整）'
                              % (tname, len(bad), max(abs(t - s) for _, t, s in bad), spec['unit']))
        else:
            print('[ok] %s：%s = 分项之和（%d 期全部通过）'
                  % (spec['name'], tname, len(data[tot_col])))

    # —— 季节性横轴：周号 → 基准年同周周四的 (月,日)，按周号排列 ——
    years = sorted({iso_year(d) for d in all_dates})
    if len(years) > MAX_YEARS:
        years = years[-MAX_YEARS:]
    base_year = iso_year(last)
    weeks = sorted({iso_week(d) for d in all_dates if iso_year(d) in years})
    md_axis, w2label = [], {}
    for w in weeks:
        lb = week_label(base_year, w)
        if lb is None:
            continue
        w2label[w] = lb
        if md_axis and md_axis[-1] == lb:
            continue
        md_axis.append(lb)
    month_first, seen = [], set()
    for lb in md_axis:
        mm = lb[:2]
        if mm not in seen:
            seen.add(mm)
            month_first.append(lb)
    n = len(years)

    # 槽位冲突检测（6 天间隔可能两点同属一个 ISO 周，后者会覆盖前者）
    clash = 0
    for col, _, _, _ in items:
        seen_pt = set()
        for d in sorted(data[col]):
            if iso_year(d) not in years:
                continue
            k = (iso_year(d), w2label.get(iso_week(d)))
            if k in seen_pt:
                clash += 1
            seen_pt.add(k)
    if clash:
        print('[warn] %s：%d 个点与同周另一点重叠（仅保留较晚一期）' % (spec['name'], clash))

    def seasonal(m):
        out = {}
        for d, v in m.items():
            y = iso_year(d)
            if y not in years:
                continue
            lb = w2label.get(iso_week(d))
            if lb is None:
                continue
            out.setdefault(y, {})[lb] = rnd(v, nd)
        return out

    charts = []

    def add_chart(key, title, group, byyear):
        # 非最新年份若只有 1 个点，画不出线段却仍占一条图例（如铁锂厂库存 2023 年仅 12-28 一期），
        # 会让人以为该年缺数据，故剔除；最新年份哪怕只有一个点也保留。
        present = [y for y in years
                   if y in byyear and (y == base_year or len(byyear[y]) >= 2)]
        if not present:
            return
        ser = []
        for y in present:
            st = series_style(years.index(y), n)
            ser.append({'name': str(y), 'color': st['color'], 'width': st['width'],
                        'dash': st['dash'], 'smooth': st['smooth'],
                        'marker': st['marker'],
                        'data': [byyear[y].get(m) for m in md_axis]})
        charts.append({'key': key, 'title': title, 'type': 'line', 'axis': 0,
                       'group': group, 'series': ser, 'fmt': 'v', 'nd': nd})

    for group, its in spec['groups']:
        for col, name, short, kw in its:
            add_chart('%s-%s' % (spec['id'], short), '%s（周）' % name, group,
                      seasonal(data[col]))

    # —— 汇总表（每个 item 用自身最新/上期/本年初首值，不按日期推算）——
    columns, cur = [], []
    for col, name, short, kw in items:
        columns.append({'key': 'c%d' % len(columns), 'label': name, 'group': spec['name']})
        ds = sorted(data[col])
        c = data[col].get(ds[-1]) if ds else None
        p = data[col].get(ds[-2]) if len(ds) > 1 else None
        ys = next((data[col][d] for d in ds if iso_year(d) == base_year), None)
        cur.append({'cur': c, 'prev': p, 'ys': ys, 'last_d': ds[-1] if ds else None,
                    'prev_d': ds[-2] if len(ds) > 1 else None,
                    'gap': (ds[-1] - ds[-2]).days if len(ds) > 1 else None,
                    'vs_ys': (c - ys) if (c is not None and ys is not None) else None,
                    'vs_ys_pct': (round((c - ys) / abs(ys) * 100, 1)
                                  if (c is not None and ys not in (None, 0)) else None)})

    g = lambda v: rnd(v, nd)
    rows = {
        '本期': [g(x['cur']) for x in cur],
        '上期': [g(x['prev']) for x in cur],
        '环比': [None if (x['cur'] is None or x['prev'] is None) else rnd(x['cur'] - x['prev'], nd)
                 for x in cur],
        '较年初': [g(x['vs_ys']) for x in cur],
        '较年初%': [x['vs_ys_pct'] for x in cur],
    }
    summary = {'columns': columns, 'rows': rows,
               'rowOrder': ['本期', '上期', '环比', '较年初', '较年初%'],
               'currentWeek': last.strftime('%Y-%m-%d'),
               'previousWeek': (cur[0]['prev_d'].strftime('%Y-%m-%d')
                                if cur[0]['prev_d'] else None),
               'unit': spec['unit'], 'nd': spec['nd']}

    # —— 卡片指标条 ——
    for ch in charts:
        short = ch['key'].split('-', 1)[1]
        it = next((x for x in items if x[2] == short), None)
        if it is None:
            continue
        x = cur[items.index(it)]
        if x['cur'] is None:
            continue
        ch['table'] = {'labels': ['本期', '上期', '环比', '较年初'],
                       'kinds': ['val', 'val', 'delta', 'delta'],
                       'values': [g(x['cur']), g(x['prev']),
                                  None if x['prev'] is None else rnd(x['cur'] - x['prev'], nd),
                                  g(x['vs_ys'])]}

    # —— 口径备注 ——
    parts = ['单位：%s' % spec['unit'], spec['src_note'],
             '横轴按周号对齐（锚定 %d 年同周日期）' % base_year]
    # 各指标起点差异（源表前期未提供该项，非本站缺失）
    base = min(all_dates)
    late = []
    for col, name, short, kw in items:
        ds = sorted(data[col])
        if ds and (ds[0] - base).days > 60:
            late.append('%s 自 %s 起' % (name, ds[0].strftime('%Y-%m')))
    if late:
        parts.append('；'.join(late) + '（源表前期无该指标）')
    gaps = sorted({x['gap'] for x in cur if x['gap'] and x['gap'] > 7})
    if gaps:
        parts.append('源表存在 %s 天的跨周采样，个别环比非严格单周' % '/'.join(str(v) for v in gaps))
    if check_note:
        parts.append('；'.join(check_note))

    ds = {'id': spec['id'], 'name': spec['name'],
          'axes': [md_axis], 'monthFirst': month_first,
          'asOf': last.strftime('%Y-%m-%d'), 'charts': charts, 'summary': summary,
          'note': '｜'.join(parts), 'unit': spec['unit']}

    print('   asOf=%s | 轴=%d 槽位 | 年份=%s | 图=%d'
          % (ds['asOf'], len(md_axis), ','.join(str(y) for y in years), len(charts)))
    return ds


def write_all(dss, stamp):
    """写 json + 重建 data.js/meta.json（保留未重算的既有数据集）"""
    for ds in dss:
        p = os.path.join(DATA, ds['id'] + '.json')
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(ds, f, ensure_ascii=False, separators=(',', ':'))
        print('[ok] 写 %s.json (%.0f KB) charts=%d asOf=%s'
              % (ds['id'], os.path.getsize(p) / 1024, len(ds['charts']), ds['asOf']))

    all_ds = []
    for did in ORDER:
        p = os.path.join(DATA, did + '.json')
        if not os.path.exists(p):
            print('[warn] 缺少 %s.json，跳过' % did)
            continue
        with open(p, encoding='utf-8') as f:
            all_ds.append(json.load(f))
    # 不在 ORDER 里的历史数据集追加到末尾，避免丢 tab
    known = {d['id'] for d in all_ds}
    for fn in sorted(os.listdir(DATA)):
        if not fn.endswith('.json') or fn == 'meta.json':
            continue
        did = fn[:-5]
        if did in known or did in ORDER:
            continue
        with open(os.path.join(DATA, fn), encoding='utf-8') as f:
            all_ds.append(json.load(f))

    with open(os.path.join(DATA, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('// 自动生成，勿手改。build_lc_downstream_charts.py @ %s\n' % stamp)
        f.write('window.CHART_DATA = ')
        json.dump({'updated': stamp, 'datasets': all_ds}, f,
                  ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    print('[ok] 写 data.js (datasets=%d)' % len(all_ds))

    meta = {'updated': stamp,
            'datasets': [{'id': d['id'], 'name': d['name'],
                          'charts': len(d['charts']), 'asOf': d['asOf']} for d in all_ds]}
    with open(os.path.join(DATA, 'meta.json'), 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print('[ok] 写 meta.json datasets=%d：%s'
          % (len(meta['datasets']), ' / '.join(d['name'] for d in meta['datasets'])))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    ap.add_argument('--only', default='', help='只重算某个 id')
    args = ap.parse_args()

    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    try:
        dss = []
        for spec in SPECS:
            if args.only and spec['id'] != args.only:
                continue
            print('\n== %s（%s）' % (spec['name'], spec['sheet']))
            ds = build_one(wb, spec)
            if ds:
                dss.append(ds)
    finally:
        wb.close()

    if not dss:
        raise SystemExit('无数据')

    if args.check:
        for ds in dss:
            print('\n---- %s(%s) | asOf %s | charts %d | axis %d'
                  % (ds['name'], ds['id'], ds['asOf'], len(ds['charts']), len(ds['axes'][0])))
            for c in ds['charts']:
                npt = len([v for v in c['series'][0]['data'] if v is not None])
                print('   [%s] %s | 年份=%s | 首年非空=%d | table=%s'
                      % (c['group'], c['title'],
                         ','.join(s['name'] for s in c['series']), npt,
                         'Y' if c.get('table') else '-'))
            print('   本期:', dict(zip([c['label'] for c in ds['summary']['columns']],
                                       ds['summary']['rows']['本期'])))
            print('   环比:', ds['summary']['rows']['环比'])
            print('   note:', ds['note'])
        return

    write_all(dss, datetime.datetime.now().isoformat(timespec='seconds'))


if __name__ == '__main__':
    main()
