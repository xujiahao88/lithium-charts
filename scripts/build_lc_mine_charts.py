# -*- coding: utf-8 -*-
"""
build_lc_mine_charts.py — 中联金《锂原料数据库》→ lithium-charts 三个数据集

源：C:\\Users\\Administrator\\Nutstore\\1\\我的坚果云\\碳酸锂\\中联金锂原料数据库.xlsx
    （中联金新能源，每周四更新；本版 26.10.30）

产出：
  lc_ore_ship   全球锂矿周度发运   —— sheet「1.0 全球锂矿周度发运」
  lc_chile_ship 智利锂盐周度发运   —— sheet「1.3 智利锂盐周度发运」
  lc_ore_inv    锂精矿港口库存     —— sheet「3.1 锂精矿港口库存」

横轴口径：ISO 周号 → 基准年（最新年）同周周四，两/多年叠线可对齐；
          周日期每年会漂 1~2 天，直接用 MM-DD 会槽位膨胀且对不齐。

派生指标（均自算，不直接沿用源表 D/E/F 列，便于填满源表留空的早期周）：
  MA4 / MA13  按**日历年**分组计算（跨年间隔一年，不能连续滚）
  年内累计     按年从该年首个观测累加
  周度变化     本期 − 上一期

用法：
  python scripts/build_lc_mine_charts.py --check    # 只试算
  python scripts/build_lc_mine_charts.py            # 落盘 data/*.json + data.js + meta.json
"""
import argparse
import datetime
import json
import os

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SRC = r'C:\Users\Administrator\Nutstore\1\我的坚果云\碳酸锂\中联金锂原料数据库.xlsx'
BASE_YEAR = 2026          # ISO 周锚定年（当前最新数据年）
MAX_YEARS = 5

ERR_TOKENS = {'#N/A', '#N/A!', '#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', '', '-'}


# ---------------------------------------------------------------- 基础工具
def clean_num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, str):
        s = v.strip()
        if s in ERR_TOKENS:
            return None
        try:
            v = float(s)
        except ValueError:
            return None
    if isinstance(v, (int, float)):
        if v != v:
            return None
        return float(v)
    return None


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


def rnd(v, nd):
    if v is None:
        return None
    return int(round(v)) if nd == 0 else round(v, nd)


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


def iso_slot(d):
    """ISO 周 → 基准年同周周四。53 周在基准年不存在时退到第 52 周。"""
    iy, iw, _ = d.isocalendar()
    try:
        return datetime.date.fromisocalendar(BASE_YEAR, iw, 4)
    except ValueError:
        return datetime.date.fromisocalendar(BASE_YEAR, 52, 4)


def iso_key(d):
    """(ISO 年, ISO 周) —— 年份归属必须用 ISO 年，否则 12/29~31 会串到上一年。"""
    iy, iw, _ = d.isocalendar()
    return iy, iw


# ---------------------------------------------------------------- 读表
def read_sheet(wb, sheet, hdr_row, colmap):
    """colmap: {name: 1-based col}；返回 (dates, data) —— data[name][date] = float"""
    if sheet not in wb.sheetnames:
        raise SystemExit('源工作簿缺少「%s」' % sheet)
    ws = wb[sheet]
    data = {k: {} for k in colmap}
    dates = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        d = to_date(ws.cell(r, colmap['date']).value)
        if d is None:
            continue
        got = False
        for name, c in colmap.items():
            if name == 'date':
                continue
            v = clean_num(ws.cell(r, c).value)
            if v is not None:
                data[name][d] = v
                got = True
        if got:
            dates.append(d)
    dates = sorted(set(dates))
    if not dates:
        raise SystemExit('「%s」无有效数据行' % sheet)
    return dates, data


def moving_avg(dates, vals, n):
    """按日历年分组滚动均值（窗口不足或含缺失则不产出）。"""
    out = {}
    by_year = {}
    for d in dates:
        if d in vals:
            by_year.setdefault(d.year, []).append(d)
    for _, ds in by_year.items():
        ds = sorted(ds)
        for i in range(n - 1, len(ds)):
            win = ds[i - n + 1:i + 1]
            if any(x not in vals for x in win):
                continue
            out[ds[i]] = sum(vals[x] for x in win) / float(n)
    return out


def cum_by_year(dates, vals):
    out, acc, cur_year = {}, 0.0, None
    for d in sorted(dates):
        if d not in vals:
            continue
        if d.year != cur_year:
            cur_year, acc = d.year, 0.0
        acc += vals[d]
        out[d] = acc
    return out


def prev_obs(dates, d):
    """上一期实际观测日（不按 −7 天推算）。"""
    idx = dates.index(d)
    return dates[idx - 1] if idx > 0 else None


def same_week_last_year(dates, d):
    """上年同 ISO 周的实际观测日；找不到返回 None。"""
    iy, iw = iso_key(d)
    for x in dates:
        if iso_key(x) == (iy - 1, iw):
            return x
    return None


# ---------------------------------------------------------------- 数据集装配
def assemble(spec, dates, series, axis_dates):
    """series: {name: {date: value}}（已换算单位）；axis_dates 决定横轴槽位"""
    slot_of = {d: iso_slot(d) for d in axis_dates}
    axis = sorted({slot_of[d] for d in axis_dates})
    label = [s.strftime('%m-%d') for s in axis]
    pos = {s: i for i, s in enumerate(axis)}

    month_first, seen = [], set()
    for s in axis:
        if s.month not in seen:
            seen.add(s.month)
            month_first.append(s.strftime('%m-%d'))

    years = sorted({iso_key(d)[0] for d in axis_dates})
    if len(years) > MAX_YEARS:
        years = years[-MAX_YEARS:]
    n = len(years)
    nd = spec['nd']
    scale = spec.get('scale', 1.0)

    def seasonal(vals):
        out = {}
        for d, v in vals.items():
            iy, _ = iso_key(d)
            if iy not in years or d not in slot_of:
                continue
            out.setdefault(iy, {})[pos[slot_of[d]]] = rnd(v * scale, nd)
        return out

    charts = []

    def add(key, title, group, vals, unit_hint=None, table=True):
        byyear = seasonal(vals)
        present = [y for y in years if y in byyear and
                   (y == BASE_YEAR or len(byyear[y]) >= 2)]
        if not present:
            return
        ser = []
        for y in present:
            st = series_style(years.index(y), n)
            row = [None] * len(axis)
            for i, v in byyear[y].items():
                row[i] = v
            ser.append({'name': str(y), 'color': st['color'], 'width': st['width'],
                        'dash': st['dash'], 'marker': st['marker'], 'data': row})
        ch = {'key': '%s-%s' % (spec['id'], key), 'title': title, 'type': 'line',
              'axis': 0, 'group': group, 'series': ser, 'fmt': 'v', 'nd': nd}
        if table:
            ch['table'] = card_table(spec, dates, vals, unit_hint)
        charts.append(ch)

    for item in spec['charts']:
        key, title, group, vals, hint = item[:5]
        # 第 6 项可选：False = 不挂图下指标条（占比类图单位与主指标不同，挂上去会丢 % ）
        want_table = item[5] if len(item) > 5 else True
        add(key, title, group, vals, hint, table=want_table)

    summary = build_summary(spec, dates, series)

    ds = {
        'id': spec['id'], 'name': spec['name'],
        'axes': [label], 'monthFirst': month_first,
        'asOf': dates[-1].strftime('%Y-%m-%d'),
        'charts': charts, 'summary': summary,
        'note': spec['note'], 'unit': spec['unit'], 'nd': nd,
    }
    return ds


def card_table(spec, dates, vals, hint=None):
    """图下指标条：本期/上期/环比/(同比)。单位随数据集精度。"""
    nd = spec['nd']
    scale = spec.get('scale', 1.0)
    last = dates[-1]
    p = prev_obs(dates, last)
    cur = vals.get(last)
    if cur is None:
        return None
    prv = vals.get(p) if p else None
    labels = ['本期', '上期', '环比']
    kinds = ['val', 'val', 'delta']
    values = [rnd(cur * scale, nd), rnd(prv * scale, nd) if prv is not None else None,
              rnd((cur - prv) * scale, nd) if prv is not None else None]
    if spec.get('yoy'):
        ly = same_week_last_year(dates, last)
        lyv = vals.get(ly) if ly else None
        labels.append('同比')
        kinds.append('delta')
        values.append(rnd((cur - lyv) * scale, nd) if lyv is not None else None)
    return {'labels': labels, 'kinds': kinds, 'values': values}


def build_summary(spec, dates, series):
    nd = spec['nd']
    scale = spec.get('scale', 1.0)
    last = dates[-1]
    p = prev_obs(dates, last)
    ly = same_week_last_year(dates, last)
    cols, rows = [], {k: [] for k in spec['rows']}

    for item in spec['summary_cols']:
        key, label, group = item[:3]
        is_cum = bool(item[3]) if len(item) > 3 else False
        vals = series[key]
        cur = vals.get(last)
        prv = vals.get(p) if p else None
        cols.append({'key': 'c%d' % len(cols), 'label': label, 'group': group})
        rows['本期'].append(rnd(cur * scale, nd) if cur is not None else None)

        if '上期' in rows:
            rows['上期'].append(rnd(prv * scale, nd) if prv is not None else None)
        if '环比' in rows:
            rows['环比'].append(rnd((cur - prv) * scale, nd)
                              if (cur is not None and prv is not None) else None)
        if '环比%' in rows:
            rows['环比%'].append(round((cur - prv) / abs(prv) * 100, 1)
                               if (cur is not None and prv not in (None, 0)) else None)
        if '同比' in rows:
            lyv = vals.get(ly) if ly else None
            rows['同比'].append(rnd((cur - lyv) * scale, nd)
                              if (cur is not None and lyv is not None) else None)
        if '同比%' in rows:
            lyv = vals.get(ly) if ly else None
            rows['同比%'].append(round((cur - lyv) / abs(lyv) * 100, 1)
                               if (cur is not None and lyv not in (None, 0)) else None)
        if '较年初' in rows:
            first = next((vals[d] for d in sorted(vals) if d.year == last.year), None)
            rows['较年初'].append(rnd((cur - first) * scale, nd)
                               if (cur is not None and first is not None) else None)
        if '较年初%' in rows:
            first = next((vals[d] for d in sorted(vals) if d.year == last.year), None)
            rows['较年初%'].append(round((cur - first) / abs(first) * 100, 1)
                                if (cur is not None and first not in (None, 0)) else None)

        # 累计列：环比/环比% 只是"本周增量"的重复，清掉避免误读；
        #          同比/同比% 是「累计同比」，有分析价值，保留。
        if is_cum:
            idx = len(cols) - 1
            for k in ('上期', '环比', '环比%'):
                if k in rows:
                    rows[k][idx] = None

    return {'columns': cols, 'rows': rows, 'rowOrder': spec['rows'],
            'currentWeek': last.strftime('%Y-%m-%d'),
            'previousWeek': p.strftime('%Y-%m-%d') if p else '—',
            'unit': spec['unit'], 'nd': nd}


# ---------------------------------------------------------------- 三张表
def build_ore_ship(wb):
    dates, raw = read_sheet(wb, '1.0 全球锂矿周度发运', 2,
                            {'date': 1, 'val': 3})
    vals = raw['val']
    series = {
        'val': vals,
        'ma4': moving_avg(dates, vals, 4),
        'ma13': moving_avg(dates, vals, 13),
        'cum': cum_by_year(dates, vals),
    }
    # 源表表头标「万吨」，但数值量级（周 5~27 万）与「1.0 全球月度发运（万吨）」表
    # （月 60~85 万）交叉核对后为**吨**；此处统一 ÷10000 显示为万吨。
    spec = {
        'id': 'lc_ore_ship', 'name': '全球锂矿发运',
        'nd': 1, 'scale': 1e-4, 'unit': '万吨', 'yoy': False,
        'charts': [
            ('val', '全球锂矿周度发运（万吨）', '全球锂矿周度发运', vals, '周'),
            ('ma4', '4 周移动平均（万吨）', '全球锂矿周度发运', series['ma4'], 'MA4'),
            ('ma13', '13 周移动平均（万吨）', '全球锂矿周度发运', series['ma13'], 'MA13'),
            ('cum', '年内累计发运（万吨）', '全球锂矿周度发运', series['cum'], '累计'),
        ],
        'summary_cols': [('val', '周度发运量', '发运量'),
                         ('ma4', '4周MA', '发运量'),
                         ('ma13', '13周MA', '发运量'),
                         ('cum', '年内累计', '发运量', True)],
        'rows': ['本期', '上期', '环比', '环比%'],
        'note': '｜'.join([
            '单位：万吨（源表数值实为吨，÷10000 显示；源表表头「万吨」标注有误，'
            '已与「全球锂矿月度发运」表交叉核对）',
            '口径：中联金新能源全球锂矿散货周度发运，含澳洲/津巴布韦/尼日利亚/南非/马里/巴西/加拿大',
            'MA4 / MA13 为本站按日历年自算（源表 D/E 列仅后期有值）',
            '⚠ 2026-11-01 期（10.26-11.01）仅 1.64 万吨、环比 −88%，'
            '显著低于年内均值，疑为当期发运/统计未完整，按源表原样呈现，解读时建议剔除',
            '⚠ 2026-09-20 期（09.14-09.20）2.05 万吨为年内次低，次周反弹至 21.98 万吨',
        ]),
    }
    return assemble(spec, dates, series, dates)


def build_chile_ship(wb):
    dates, raw = read_sheet(wb, '1.3 智利锂盐周度发运', 1,
                            {'date': 1, 'val': 3})
    vals = raw['val']
    series = {
        'val': vals,
        'ma4': moving_avg(dates, vals, 4),
        'ma13': moving_avg(dates, vals, 13),
        'cum': cum_by_year(dates, vals),
    }
    spec = {
        'id': 'lc_chile_ship', 'name': '智利锂盐发运',
        'nd': 0, 'scale': 1.0, 'unit': '吨', 'yoy': True,
        'charts': [
            ('val', '智利锂盐周度发运（吨）', '智利锂盐周度发运', vals, '周'),
            ('ma4', '4 周移动平均（吨）', '智利锂盐周度发运', series['ma4'], 'MA4'),
            ('ma13', '13 周移动平均（吨）', '智利锂盐周度发运', series['ma13'], 'MA13'),
            ('cum', '年内累计发运（吨）', '智利锂盐周度发运', series['cum'], '累计'),
        ],
        'summary_cols': [('val', '周度发运量', '发运量'),
                         ('ma4', '4周MA', '发运量'),
                         ('ma13', '13周MA', '发运量'),
                         ('cum', '年内累计', '发运量', True)],
        'rows': ['本期', '上期', '环比', '环比%', '同比', '同比%'],
        'note': '｜'.join([
            '单位：吨（源表原值）',
            '口径：智利锂盐（碳酸锂 + 氢氧化锂 + 硫酸锂）周度发运，中联金新能源',
            '⚠ 每季度会按实际发货情况回溯修正历史周值，故历史值可能与上期版本不同——'
            '做环比/同比前请以同版本重算，勿直接对比上期已发布图',
            'MA4 / MA13 为本站按日历年自算',
        ]),
    }
    return assemble(spec, dates, series, dates)


def build_ore_inv(wb):
    dates, raw = read_sheet(wb, '3.1 锂精矿港口库存', 1,
                            {'date': 1, 'qz': 2, 'zj': 3, 'cq': 4, 'other': 5, 'total': 6})
    series = {'total': raw['total'], 'qz': raw['qz'], 'zj': raw['zj'],
              'cq': raw['cq'], 'other': raw['other']}
    # 派生：镇江港占合计比重
    series['zj_share'] = {d: raw['zj'][d] / raw['total'][d] * 100.0
                          for d in dates
                          if d in raw['zj'] and raw['total'].get(d)}

    # 勾稽自检：合计 = 四港之和
    bad = []
    for d in dates:
        parts = [raw[k].get(d) for k in ('qz', 'zj', 'cq', 'other')]
        tot = raw['total'].get(d)
        if tot is None or any(v is None for v in parts):
            continue
        if abs(sum(parts) - tot) > max(0.05, tot * 1e-3):
            bad.append((d.strftime('%Y-%m-%d'), tot, round(sum(parts), 2)))
    if bad:
        print('[warn] 港口库存合计与分港不勾稽 %d 期：%s' % (len(bad), bad[:3]))
    else:
        print('[ok] 港口库存勾稽通过：合计 = 钦州 + 镇江 + 重庆 + 其他')

    spec = {
        'id': 'lc_ore_inv', 'name': '锂精矿港口库存',
        'nd': 2, 'scale': 1.0, 'unit': '万吨', 'yoy': False,
        'charts': [
            ('total', '锂精矿港口库存合计（万吨）', '锂精矿港口库存', series['total'], '合计'),
            ('qz', '钦州港（万吨）', '锂精矿港口库存', series['qz'], '钦州'),
            ('zj', '镇江港（万吨）', '锂精矿港口库存', series['zj'], '镇江'),
            ('cq', '重庆港（万吨）', '锂精矿港口库存', series['cq'], '重庆'),
            ('other', '其他港（万吨）', '锂精矿港口库存', series['other'], '其他'),
            ('zj_share', '镇江港占合计比重（%）', '锂精矿港口库存', series['zj_share'], '镇江占比', False),
        ],
        'summary_cols': [('total', '库存合计', '港口库存'),
                         ('qz', '钦州港', '分港口'),
                         ('zj', '镇江港', '分港口'),
                         ('cq', '重庆港', '分港口'),
                         ('other', '其他港', '分港口')],
        'rows': ['本期', '上期', '环比', '较年初', '较年初%'],
        'note': '｜'.join([
            '单位：万吨',
            '口径：中国锂精矿港口库存（钦州港 / 镇江港 / 重庆港 / 其他港），中联金新能源周度调研',
            '结构变化：镇江港由 2025 年中 3~6 万吨升至 2026-10 的 26.8 万吨，'
            '其他港由 30 万吨以上降至 14.2 万吨——港间转移明显，看「合计」比看单港更可靠',
        ]),
    }
    return assemble(spec, dates, series, dates)


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
        p = os.path.join(DATA, d['id'] + '.json')
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, separators=(',', ':'))
        print('[ok] 写 %s.json (%.0f KB) charts=%d asOf=%s'
              % (d['id'], os.path.getsize(p) / 1024, len(d['charts']), d['asOf']))

    with open(os.path.join(DATA, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('// 自动生成，勿手改。build_lc_mine_charts.py @ ' + stamp + '\n')
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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    ap.add_argument('--src', default=SRC)
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.src, data_only=True)
    try:
        ds_list = [build_ore_ship(wb), build_chile_ship(wb), build_ore_inv(wb)]
    finally:
        wb.close()

    for ds in ds_list:
        print('\n== %s(%s) | asOf %s | 图 %d | 轴 %d 槽 | 年 %s'
              % (ds['name'], ds['id'], ds['asOf'], len(ds['charts']),
                 len(ds['axes'][0]),
                 ','.join(sorted({s['name'] for c in ds['charts'] for s in c['series']}))))
        for c in ds['charts']:
            yrs = [s['name'] for s in c['series']]
            npt = sum(1 for v in c['series'][0]['data'] if v is not None)
            print('   [%s] %s | 年份=%s | 主系列非空=%d | table=%s'
                  % (c['group'], c['title'], ','.join(yrs), npt,
                     'Y' if c.get('table') else '-'))
        print('   cols:', [c['label'] for c in ds['summary']['columns']])
        for k in ds['summary']['rowOrder']:
            print('   %-6s %s' % (k, ds['summary']['rows'][k]))

    if args.check:
        return
    write_all(ds_list)


if __name__ == '__main__':
    main()
