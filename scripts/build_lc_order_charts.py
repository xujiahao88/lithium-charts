# -*- coding: utf-8 -*-
"""
build_lc_order_charts.py — 《新能源汽车订单.xlsx》→ lithium-charts 数据集 lc_order

源：C:\\Users\\Administrator\\Nutstore\\1\\我的坚果云\\碳酸锂\\新能源汽车订单.xlsx
    单 sheet「电车-订单（周）」，数据来自高盛《China New Energy Vehicle Weekly Chartbook》。
    表结构（非标准长表，按行列定位）：
      · 第 4 行 = 日期行；列 2~42 为 2025 年（c2=2025-01-05 起），
        列 55~98 为 2026 年（c55=2026-01-04 起），列 43~53 与 54 为分隔/表头列。
      · 品牌行：5 BYD / 6 HIMA / 12 Li Auto / 13 Nio / 17 Tesla / 18 Xiaomi / 19 XPeng
        / 20 2026SUM（源表合计）。7~11、14~16 为子品牌行（2026 年 3 月后停更，不用）。

⚠ 两处口径陷阱（已处理，勿随手删）：
  1) **2025 年日期标签比 2026 年早 7 天**。2026 年标签 = GS 该周结束日（如 W41=10/5-10/11 标 10-11）；
     2025 年标签 = 该周结束日 −7 天（GS 2025 W39=9/22-9/28 的 244.5 被记在 09-21）。
     实测：源表 09-21=244.5k、10-05=205.3k 分别精确等于 GS 2025 W39/W41，
     而表头周号（c39=第38周）与日期一致 → 是**数据行整体早一周**，不是表头错。
     故对 2025 年日期 +7 天后再取 ISO 周，否则同比对标整体错位一周。
  2) **合计行中途换口径**。2025 全年合计 = 7 个品牌之和；2026 年 1/04~5/03 也是 7 品牌之和，
     **自 2026-05-10 起**才并入 Geely、Leapmotor（合计比 7 品牌和多出 3.3~4.5 万辆）。
     故「9 品牌合计」无任何历史同口径：只画 2026-05-10 之后、同比留空；
     跨年可比只用自算的「7 品牌可比合计」列（同比 −41.3% 与此吻合）。

横轴：与下游三表一致，ISO 周号 → 基准年（最新年份）同周周四的 (月,日)，
      多年同周号落同一槽位，轴收敛到 52~53 槽。

用法：
  python scripts/build_lc_order_charts.py           # 落盘 + 重建 data.js/meta.json
  python scripts/build_lc_order_charts.py --check   # 只打印，不落盘
"""
import argparse
import calendar
import datetime
import json
import os

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SRC = r'C:\Users\Administrator\Nutstore\1\我的坚果云\碳酸锂\新能源汽车订单.xlsx'
SHEET = '电车-订单（周）'

DS_ID = 'lc_order'
DS_NAME = '新能源车订单'
UNIT = '辆'
ND = 0
MAX_YEARS = 5

DATE_ROW = 4
# 2025 段 c2~c54、2026 段 c55~c98 全扫，靠「第 4 行是否 datetime」筛掉分隔空列。
# ⚠ 原先写死 range(2,43)，2026-10-29 补 2025 W43 到 c43 后整列读不到，
#   导致本期同比全为 None —— 不要再改回固定上界。
COLS = list(range(2, 55)) + list(range(55, 99))

# 行号 -> (显示名, 简称, 行名列校验关键字)
BRANDS = [
    (5,  '比亚迪 BYD',      'BYD',    'BYD'),
    (6,  '鸿蒙智行 HIMA',   'HIMA',   'HIMA'),
    (12, '理想 Li Auto',    '理想',   'Li Auto'),
    (13, '蔚来 Nio',        '蔚来',   'Nio'),
    (17, '特斯拉 Tesla',    '特斯拉', 'Tesla'),
    (18, '小米 Xiaomi',     '小米',   'Xiaomi'),
    (19, '小鹏 XPeng',      '小鹏',   'XPeng'),
]
SUM_ROW = 20          # 源表 SUM 行
# ⚠ 合计行自 2026-05-10 起才含 Geely / Leapmotor（2026-01-04~05-03 仍等于 7 品牌之和），
#   2025 全年也是 7 品牌 → 9 品牌口径**没有任何历史年份可比**，故该序列只从 2026-05-10 起画，
#   且不给同比（给了就是拿 7 品牌的旧值当基数，−8.9% 是假数）。
SUM_CUT = datetime.date(2026, 5, 10)
SUM_NAME, SUM_SHORT = '9 品牌合计（含吉利/零跑）', '9 品牌合计'

ERR_TOKENS = {'#N/A', '#N/A!', '#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', ''}


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


def rnd(v, nd=0):
    return None if v is None else (int(round(v)) if nd == 0 else round(v, nd))


def iso_week(d):
    return d.isocalendar()[1]


def iso_year(d):
    return d.isocalendar()[0]


def norm_date(d):
    """2025 年标签统一 +7 天到「该周结束日」口径（见文件头说明 1）。"""
    return d + datetime.timedelta(days=7) if d.year <= 2025 else d


def week_label(base_year, week):
    for w in (week, week - 1, 52):
        try:
            thu = datetime.date.fromisocalendar(base_year, w, 4)
            return '%02d-%02d' % (thu.month, thu.day)
        except ValueError:
            continue
    return None


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
    return {'color': color, 'width': 1.5,
            'dash': 'dash' if i <= n - 4 else 'solid',
            'smooth': False, 'marker': 'circle' if latest else 'none'}


def parse_sheet(wb):
    if SHEET not in wb.sheetnames:
        raise SystemExit('源工作簿缺少「%s」sheet' % SHEET)
    ws = wb[SHEET]

    # 行名校验，防止插行后静默错位
    bad = []
    for r, name, short, kw in BRANDS + [(SUM_ROW, SUM_NAME, SUM_SHORT, 'SUM')]:
        got = ws.cell(r, 1).value
        got = '' if got is None else str(got).strip()
        if kw.upper() not in got.upper():
            bad.append('第%d行期望「%s」，实际「%s」' % (r, kw, got))
    if bad:
        raise SystemExit('行定位失败：%s' % '；'.join(bad))
    print('[ok] 品牌行定位通过（8 行）')

    # 日期行：表头周号（第3行）与日期（第4行）一致性抽检
    dates, skipped = {}, 0
    for c in COLS:
        raw = ws.cell(DATE_ROW, c).value
        if isinstance(raw, datetime.datetime):
            d = raw.date()
        elif isinstance(raw, datetime.date):
            d = raw
        else:
            skipped += 1
            continue
        dates[c] = d
    if not dates:
        raise SystemExit('日期行（第%d行）无有效日期' % DATE_ROW)
    # 抽检：第3行「第N周」应等于该日期所在 ISO 周
    mism = []
    for c, d in dates.items():
        tag = ws.cell(DATE_ROW - 1, c).value
        if not isinstance(tag, str) or not tag.startswith('第'):
            continue
        try:
            n = int(tag.replace('第', '').replace('周', ''))
        except ValueError:
            continue
        if n != iso_week(d):
            mism.append((c, tag, d.isoformat(), iso_week(d)))
    if mism:
        print('[warn] 表头周号与日期 ISO 周不一致 %d 处，前 3：%s' % (len(mism), mism[:3]))
    else:
        print('[ok] 表头周号与日期 ISO 周一致（%d 列）' % len([1 for c in dates]))

    data = {k: {} for k in ['SUM'] + [s for _, _, s, _ in BRANDS]}
    for c, d in dates.items():
        nd = norm_date(d)
        for r, name, short, kw in BRANDS:
            v = clean_num(ws.cell(r, c).value)
            if v is not None:
                data[short][nd] = v
        v = clean_num(ws.cell(SUM_ROW, c).value)
        if v is not None:
            data['SUM'][nd] = v
    return data


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    args = ap.parse_args()

    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    try:
        data = parse_sheet(wb)
    finally:
        wb.close()

    shorts = [s for _, _, s, _ in BRANDS]
    all_dates = sorted({d for m in data.values() for d in m})
    if not all_dates:
        raise SystemExit('无数据')
    last = all_dates[-1]
    base_year = iso_year(last)

    # 7 品牌可比合计（自算）：2025/2026 同口径，用于同比
    comp = {}
    for d in all_dates:
        vals = [data[s].get(d) for s in shorts]
        if all(v is not None for v in vals):
            comp[d] = sum(vals)
    data['COMP'] = comp

    years = sorted({iso_year(d) for d in all_dates})
    if len(years) > MAX_YEARS:
        years = years[-MAX_YEARS:]
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

    def seasonal(m, min_date=None, only_year=None):
        """min_date：早于该日的点不输出（口径起点截断）；only_year：只保留该 ISO 年。"""
        out = {}
        for d, v in m.items():
            y = iso_year(d)
            if y not in years:
                continue
            if only_year is not None and y != only_year:
                continue
            if min_date is not None and d < min_date:
                continue
            lb = w2label.get(iso_week(d))
            if lb is None:
                continue
            out.setdefault(y, {})[lb] = rnd(v, ND)
        return out

    charts = []

    def add_chart(key, title, group, byyear, no_yoy=False):
        present = [y for y in years if y in byyear and (y == base_year or len(byyear[y]) >= 2)]
        if not present:
            return
        ser = []
        for y in present:
            st = series_style(years.index(y), n)
            ser.append({'name': str(y), 'color': st['color'], 'width': st['width'],
                        'dash': st['dash'], 'smooth': False,
                        'marker': 'circle' if y == base_year else 'none',
                        'data': [byyear[y].get(m) for m in md_axis]})
        ch = {'key': key, 'title': title, 'type': 'line', 'axis': 0,
              'group': group, 'series': ser, 'fmt': 'v', 'nd': ND}
        if no_yoy:
            ch['noYoy'] = True
        charts.append(ch)

    # 9 品牌合计：无历史同口径 → 只画基准年、从 2026-05-10 起；不给同比。
    add_chart('%s-%s' % (DS_ID, '合计'), '9 品牌合计（2026-05 起含吉利、零跑）', '周订单合计',
              seasonal(data['SUM'], min_date=SUM_CUT, only_year=base_year), no_yoy=True)
    # 7 品牌可比：唯一能跨年叠线的口径
    add_chart('%s-%s' % (DS_ID, '可比合计'), '7 品牌可比合计（跨年可叠）', '周订单合计', seasonal(comp))
    for _, name, short, _ in BRANDS:
        add_chart('%s-%s' % (DS_ID, short), '%s（周）' % name, '分品牌周订单', seasonal(data[short]))

    # —— 汇总表：本期 / 上期 / 环比 / 上年同周 / 同比% ——
    cols_def = [(s, next(nm for _, nm, s2, _ in BRANDS if s2 == s), '分品牌')
                for s in shorts]
    cols_def.insert(0, ('COMP', '7 品牌可比合计', '合计'))
    cols_def.append(('SUM', SUM_NAME, '合计'))

    def yoy_of(m, d):
        """上年同 ISO 周的值（源按周号匹配，不看具体日期）。"""
        if d is None:
            return None
        w = iso_week(d)
        for dd, vv in m.items():
            if iso_year(dd) == iso_year(d) - 1 and iso_week(dd) == w:
                return vv
        return None

    columns, cur, rows_cur, rows_prev, rows_wow, rows_yoy, rows_yoy_pct = [], [], [], [], [], [], []
    last_d_used, prev_d_used = None, None
    for key, label, group in cols_def:
        m = data[key]
        ds = sorted(m)
        columns.append({'key': 'c%d' % len(columns), 'label': label, 'group': group})
        c = m[ds[-1]] if ds else None
        p = m[ds[-2]] if len(ds) > 1 else None
        y = yoy_of(m, ds[-1]) if ds else None
        if key == 'SUM':
            y = None   # 9 品牌无上年同口径（2025 只有 7 品牌），同比留空而不是给假数
        cur.append({'cur': c, 'prev': p, 'yoy': y, 'last_d': ds[-1] if ds else None,
                    'prev_d': ds[-2] if len(ds) > 1 else None})
        rows_cur.append(rnd(c, ND))
        rows_prev.append(rnd(p, ND))
        rows_wow.append(None if (c is None or p is None) else rnd(c - p, ND))
        rows_yoy.append(None if (c is None or y is None) else rnd(c - y, ND))
        rows_yoy_pct.append(None if (c is None or not y) else round((c - y) / abs(y) * 100, 1))
        if ds:
            last_d_used = ds[-1]
            if len(ds) > 1:
                prev_d_used = ds[-2]

    summary = {'columns': columns,
               'rows': {'本期': rows_cur, '上期': rows_prev, '环比': rows_wow,
                        '同比': rows_yoy, '同比%': rows_yoy_pct},
               'rowOrder': ['本期', '上期', '环比', '同比', '同比%'],
               'currentWeek': last_d_used.strftime('%Y-%m-%d'),
               'previousWeek': prev_d_used.strftime('%Y-%m-%d') if prev_d_used else None,
               'unit': UNIT, 'nd': ND}

    # —— 卡片指标条 ——
    for ch in charts:
        short = ch['key'].split('-', 1)[1]
        idx = next((i for i, (k, _, _) in enumerate(cols_def) if
                    (k == 'COMP' and short == '可比合计') or (k == short)), None)
        if idx is None:
            continue
        x = cur[idx]
        if x['cur'] is None:
            continue
        ch['table'] = {'labels': ['本期', '上期', '环比'],
                       'kinds': ['val', 'val', 'delta'],
                       'values': [rnd(x['cur'], ND), rnd(x['prev'], ND),
                                  None if x['prev'] is None else rnd(x['cur'] - x['prev'], ND)]}
        if not ch.get('noYoy'):
            ch['table']['labels'].append('同比')
            ch['table']['kinds'].append('delta')
            ch['table']['values'].append(
                None if x['yoy'] is None else rnd(x['cur'] - x['yoy'], ND))

    # ================= 月度（第二根轴：'1'..'12'） =================
    # 口径：把每周值按 7 天日均摊到它覆盖的自然日（窗口 [d-6, d]），再按自然月累加。
    #  · cov < 当月天数 80% → 丢弃（2025-02 只有 19 天、2025-12 是 3 天碎片）
    #  · 80%~100% → 按 cov 折算成整月并记 est（当月没过完 / 源表上年停在同周）
    #    折算不改变同月同比，因为两年系数相同（都覆盖到 10-25）。
    MONTH_LB = [str(i) for i in range(1, 13)]

    def to_month(m, min_ym=None, only_year=None):
        acc = {}
        for d, v in m.items():
            per = v / 7.0
            for k in range(7):
                dd = d - datetime.timedelta(days=6 - k)
                a = acc.setdefault((dd.year, dd.month), [0.0, set()])
                a[0] += per
                a[1].add(dd.day)
        by_year, est = {}, set()
        for (y, mo), (v, days) in acc.items():
            if only_year is not None and y != only_year:
                continue
            ym = '%04d-%02d' % (y, mo)
            if min_ym is not None and ym < min_ym:
                continue
            full = calendar.monthrange(y, mo)[1]
            cov = len(days)
            if cov < full * 0.8:
                continue
            val = v * full / cov if cov < full else v
            if cov < full:
                est.add(ym)
            by_year.setdefault(y, {})[str(mo)] = rnd(val, ND)
        return by_year, est

    est_all = set()

    def add_month_chart(key, title, group, by_year, no_yoy=False):
        present = [y for y in years if y in by_year and (y == base_year or len(by_year[y]) >= 2)]
        if not present:
            return
        ser = []
        for y in present:
            st = series_style(years.index(y), n)
            ser.append({'name': str(y), 'color': st['color'], 'width': st['width'],
                        'dash': st['dash'], 'smooth': False,
                        'marker': 'circle' if y == base_year else 'none',
                        'data': [by_year[y].get(mo) for mo in MONTH_LB]})
        ch = {'key': key, 'title': title, 'type': 'line', 'axis': 1,
              'group': group, 'series': ser, 'fmt': 'v', 'nd': ND}
        if no_yoy:
            ch['noYoy'] = True
        # 卡片指标条：本月 / 上月 / 环比 / 同比（同比按上年同月）
        ms = sorted([int(x) for x in by_year.get(base_year, {})])
        if ms:
            cm = ms[-1]
            cv = by_year[base_year].get(str(cm))
            pv = by_year[base_year].get(str(cm - 1))
            yv = None if no_yoy else by_year.get(base_year - 1, {}).get(str(cm))
            labels = ['本月', '上月', '环比']
            kinds = ['val', 'val', 'delta']
            vals = [cv, pv, None if (cv is None or pv is None) else rnd(cv - pv, ND)]
            if not no_yoy:
                labels.append('同比')
                kinds.append('delta')
                vals.append(None if (cv is None or yv is None) else rnd(cv - yv, ND))
            ch['table'] = {'labels': labels, 'kinds': kinds, 'values': vals}
        charts.append(ch)

    sum_m, e1 = to_month(data['SUM'], min_ym='2026-06', only_year=base_year)
    est_all |= e1
    add_month_chart(DS_ID + '-合计-m', '9 品牌合计（月度，2026-06 起）', '月度合计',
                    sum_m, no_yoy=True)
    comp_m, e2 = to_month(comp)
    est_all |= e2
    add_month_chart(DS_ID + '-可比合计-m', '7 品牌可比合计（月度，跨年可叠）', '月度合计', comp_m)
    for _, name, short, _ in BRANDS:
        bm, e3 = to_month(data[short])
        est_all |= e3
        add_month_chart('%s-%s-m' % (DS_ID, short), '%s（月度）' % name, '分品牌月度订单', bm)

    note = '｜'.join([
        '单位：%s' % UNIT,
        '高盛《中国新能源汽车周度图表》Exhibit 1（原表 k units ×1000）',
        '周度横轴按 ISO 周号对齐（锚定 %d 年同周日期）；月度横轴为 1~12 月' % base_year,
        '⚠ 源表 2025 年日期标签较 2026 年早 7 天，已统一到「周结束日」后再对齐周号',
        '⚠ 9 品牌合计自 2026-05-10 起含吉利（Geely）与零跑（Leapmotor）；2026-01~04 与 2025 全年'
        '均为 7 品牌之和 → 9 品牌无上年同口径，故该图只画 2026-05（月度取 06）之后、同比留空；'
        '跨年比较一律看「7 品牌可比合计」',
        '月度口径：周值按 7 天日均摊入自然月后逐月累加（跨月周按天数分摊），'
        '避免 4/5 周月带来的机械波动；覆盖不足 80% 的月份已剔除',
        ('⚠ 按已覆盖天数折算为整月的月份：%s（两年均覆盖至 10-25，折算系数一致，同月同比不受影响）'
         % '、'.join(sorted(est_all))) if est_all else '',
        '2025 年数据自 2 月中旬起（2026 年自 1 月起），故 1 月无同比、2025 年 2 月不绘',
    ])

    ds = {'id': DS_ID, 'name': DS_NAME,
          'axes': [md_axis, MONTH_LB], 'axisTypes': ['week', 'month'],
          'monthFirst': month_first,
          'asOf': last.strftime('%Y-%m-%d'), 'charts': charts, 'summary': summary,
          'note': note, 'unit': UNIT}

    print('\n== %s | asOf %s | 周轴 %d 槽位 | 年份 %s | 图 %d（周 %d / 月 %d）'
          % (DS_NAME, ds['asOf'], len(md_axis), ','.join(str(y) for y in years),
             len(charts), len(charts) - 9, 9))
    print('   本周:', dict(zip([c['label'] for c in columns], rows_cur)))
    print('   环比:', rows_wow)
    print('   同比%:', rows_yoy_pct)
    print('   月度折算月:', sorted(est_all))
    for ch in charts:
        print('   [%s] %s' % (ch['group'], ch['title']))

    if args.check:
        return

    p = os.path.join(DATA, DS_ID + '.json')
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(ds, f, ensure_ascii=False, separators=(',', ':'))
    print('[ok] 写 %s.json (%.0f KB)' % (DS_ID, os.path.getsize(p) / 1024))

    # 复用下游脚本的合并逻辑，保证 tab 顺序与其它 datasets 一致
    import importlib.util
    spec_p = os.path.join(ROOT, 'scripts', 'build_lc_downstream_charts.py')
    spec = importlib.util.spec_from_file_location('bld', spec_p)
    bld = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(bld)
    if DS_ID not in bld.ORDER:
        bld.ORDER.append(DS_ID)
    bld.write_all([ds], datetime.datetime.now().isoformat(timespec='seconds'))


if __name__ == '__main__':
    main()
