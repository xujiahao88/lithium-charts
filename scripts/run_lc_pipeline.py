# -*- coding: utf-8 -*-
"""
run_lc_pipeline.py — 碳酸锂下游周度「一条龙」

build（钢联 → lc_inv；SMM 三表 → lc_output / lc_cathode / lc_cell）
  → bump index.html ?v=NN
  → 推 GitHub Pages（deploy_repo.py）
  → Edge headless 出 2K 长图（默认只出「碳酸锂库存」）
  → 发微信（当前打开的会话）

注：4 个数据集每次都重建并推站（保证站点数据一致）；微信默认仍只发库存那张，
    避免同一周连发多条。要发下游拼接图时加 --downstream。

用法：
  python scripts/run_lc_pipeline.py                 # 全链，默认发微信（只发库存图）
  python scripts/run_lc_pipeline.py --downstream    # 另出「产量+正极+电芯」拼接长图并发送
  python scripts/run_lc_pipeline.py --no-send       # 只更新站点 + 出图
  python scripts/run_lc_pipeline.py --no-push       # 不推公网
  python scripts/run_lc_pipeline.py --no-shot       # 不出图
  python scripts/run_lc_pipeline.py --check-only    # 只试算，不落盘
  python scripts/run_lc_pipeline.py --force         # 忽略 asOf 去重，强制重跑重发
"""
import argparse
import datetime
import json
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PY = sys.executable
BUILD = os.path.join(ROOT, 'scripts', 'build_lc_inv_charts.py')
BUILD2 = os.path.join(ROOT, 'scripts', 'build_lc_downstream_charts.py')
STACK = os.path.join(ROOT, 'scripts', 'stack_images.py')
SHOT = os.path.join(ROOT, 'scripts', 'shot_lithium.py')
DEPLOY = os.path.join(ROOT, 'deploy_repo.py')
SEND = r'C:\Users\Administrator\iron-ore-charts\scripts\send_to_wechat.py'
STATE = os.path.join(ROOT, 'shot', '.lc_inv_state.json')
IMG = os.path.join(ROOT, 'shot', '碳酸锂库存_长图.png')
IMG_DN = os.path.join(ROOT, 'shot', '碳酸锂下游_长图.png')

# 下游三段（产量 / 正极 / 电芯）→ 拼成一张长图
# 2026-10-21：用户要求「产量 / 电芯」一行三张。CSS 断点已放宽到 1000px，
# 这里再把视口放宽到 1440（>1360 不触发塌列）并降 scale 到 1.5，
# 输出仍是 2160px 宽，但每张卡从 330 → 440 CSS px，X 轴月份标签不再挤。
DN_SEGS = [
    ('lc_output',  '碳酸锂产量', 'shot/_dn_output.png',  1200),
    ('lc_cathode', '正极材料',   'shot/_dn_cathode.png', 1400),
    ('lc_cell',    '动力电芯',   'shot/_dn_cell.png',     800),
]
DN_WIDTH = 1440
DN_SCALE = 1.5
REPO = 'lithium-charts'
DS_ID = 'lc_inv'

PUSH_FILES = [
    'index.html', 'README.md',
    'assets/app.js', 'assets/style.css', 'assets/vendor/echarts.min.js',
    'data/data.js', 'data/meta.json', 'data/lc_inv.json',
    'data/lc_output.json', 'data/lc_cathode.json', 'data/lc_cell.json',
    'scripts/build_lc_inv_charts.py', 'scripts/build_lc_downstream_charts.py',
    'scripts/shot_lithium.py', 'scripts/stack_images.py',
    'scripts/run_lc_pipeline.py',
]


def run(cmd, **kw):
    print('\n$ ' + ' '.join(str(c) for c in cmd))
    r = subprocess.run([str(c) for c in cmd], **kw)
    if r.returncode != 0:
        raise SystemExit('命令失败 rc=%d: %s' % (r.returncode, cmd[1:3]))
    return r


def read_signature():
    # 4 个数据集 asOf 的组合签名。任一换期都应重发 ——
    # 否则只更新了 SMM 三表、lc_inv 没变时，会被判成「已发送」直接跳过。
    p = os.path.join(ROOT, 'data', 'meta.json')
    try:
        with open(p, encoding='utf-8') as f:
            ds = json.load(f).get('datasets', [])
        return '|'.join('%s=%s' % (d.get('id'), d.get('asOf')) for d in ds)
    except Exception:
        return None


def as_of_of(ds_id):
    p = os.path.join(ROOT, 'data', 'meta.json')
    try:
        with open(p, encoding='utf-8') as f:
            for d in json.load(f).get('datasets', []):
                if d.get('id') == ds_id:
                    return d.get('asOf')
    except Exception:
        pass
    return None


def bump_cache():
    p = os.path.join(ROOT, 'index.html')
    s = open(p, encoding='utf-8').read()
    old = sorted(set(re.findall(r'\?v=(\d+)', s)), key=int)
    if not old:
        print('[warn] index.html 无 ?v=NN，跳过 bump')
        return
    new = str(int(old[-1]) + 1)
    open(p, 'w', encoding='utf-8').write(s.replace('?v=' + old[-1], '?v=' + new))
    print('[ok] bump ?v=%s → %s' % (old[-1], new))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--no-send', action='store_true')
    ap.add_argument('--no-push', action='store_true')
    ap.add_argument('--no-shot', action='store_true')
    ap.add_argument('--check-only', action='store_true')
    ap.add_argument('--force', action='store_true')
    ap.add_argument('--downstream', action='store_true',
                    help='另出「产量+正极+电芯」拼接长图并发送')
    args = ap.parse_args()

    # 1) build：钢联库存 + SMM 三表
    run([PY, BUILD] + (['--check'] if args.check_only else []), cwd=ROOT)
    run([PY, BUILD2] + (['--check'] if args.check_only else []), cwd=ROOT)
    if args.check_only:
        return

    as_of = as_of_of(DS_ID)
    sig = read_signature()
    print('\n📅 asOf =', as_of)
    print('🔖 签名 =', sig)

    st = {}
    if os.path.exists(STATE):
        try:
            st = json.load(open(STATE, encoding='utf-8'))
        except Exception:
            st = {}
    if not args.force and sig and st.get('sent_sig') == sig and not args.no_send:
        print('⏭️ 该期（%s）已发送过，跳过。需重发请加 --force' % as_of)
        return

    # 2) bump
    bump_cache()

    # 3) push
    if not args.no_push:
        env = dict(os.environ)
        env['COMMIT_MSG'] = 'data: 碳酸锂产业链数据更新（库存至 %s）' % as_of
        run([PY, DEPLOY, REPO] + PUSH_FILES, cwd=ROOT, env=env)
    else:
        print('[skip] 不推公网')

    # 4) shot
    if not args.no_shot:
        run([PY, SHOT, '--ds', DS_ID, '--out', IMG,
             '--width', '1080', '--height', '4200', '--scale', '2',
             '--title', '碳酸锂库存',
             '--subtitle', '数据截至 %s｜我的钢铁网（Mysteel）周度调研｜单位：吨' % as_of],
            cwd=ROOT)
    else:
        print('[skip] 不出图')

    # 4b) 下游拼接图（可选）
    if args.downstream and not args.no_shot:
        parts = []
        for ds_id, _name, out, h in DN_SEGS:
            run([PY, SHOT, '--ds', ds_id, '--out', os.path.join(ROOT, out),
                 '--width', str(DN_WIDTH), '--height', str(h),
                 '--scale', str(DN_SCALE)], cwd=ROOT)
            parts.append(os.path.join(ROOT, out))
        run([PY, STACK, '--out', IMG_DN, '--inputs'] + parts + ['--gap', '30'], cwd=ROOT)

    # 5) send
    if args.no_send:
        print('\n[skip] 不发微信（--no-send）')
    else:
        if not os.path.exists(IMG):
            raise SystemExit('长图不存在：%s' % IMG)
        r = subprocess.run([PY, SEND, IMG, '--no-countdown'], cwd=ROOT)
        ok = (r.returncode == 0)
        if ok and args.downstream and os.path.exists(IMG_DN):
            r2 = subprocess.run([PY, SEND, IMG_DN, '--no-countdown'], cwd=ROOT)
            ok = (r2.returncode == 0)
        if ok:
            st['sent_asOf'] = as_of
            st['sent_sig'] = sig
            st['sentAt'] = datetime.datetime.now().isoformat(timespec='seconds')
            os.makedirs(os.path.dirname(STATE), exist_ok=True)
            json.dump(st, open(STATE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
            print('✅ 已发送并记账')
        else:
            print('⚠️ 微信发送未确认（rc=%d），未记账 → 下次运行会重试' % r.returncode)


if __name__ == '__main__':
    main()
