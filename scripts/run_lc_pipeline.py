# -*- coding: utf-8 -*-
"""
run_lc_pipeline.py — 碳酸锂下游周度「一条龙」

build（解析钢联 → data/lc_inv.json + data.js + meta.json）
  → bump index.html ?v=NN
  → 推 GitHub Pages（deploy_repo.py）
  → Edge headless 出 2K 长图
  → 发微信（当前打开的会话）

用法：
  python scripts/run_lc_pipeline.py                 # 全链，默认发微信
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
SHOT = os.path.join(ROOT, 'scripts', 'shot_lithium.py')
DEPLOY = os.path.join(ROOT, 'deploy_repo.py')
SEND = r'C:\Users\Administrator\iron-ore-charts\scripts\send_to_wechat.py'
STATE = os.path.join(ROOT, 'shot', '.lc_inv_state.json')
IMG = os.path.join(ROOT, 'shot', '碳酸锂下游_长图.png')
REPO = 'lithium-charts'
DS_ID = 'lc_inv'

PUSH_FILES = [
    'index.html', 'README.md',
    'assets/app.js', 'assets/style.css', 'assets/vendor/echarts.min.js',
    'data/data.js', 'data/meta.json', 'data/lc_inv.json',
    'scripts/build_lc_inv_charts.py', 'scripts/shot_lithium.py',
    'scripts/run_lc_pipeline.py',
]


def run(cmd, **kw):
    print('\n$ ' + ' '.join(str(c) for c in cmd))
    r = subprocess.run([str(c) for c in cmd], **kw)
    if r.returncode != 0:
        raise SystemExit('命令失败 rc=%d: %s' % (r.returncode, cmd[1:3]))
    return r


def read_as_of():
    p = os.path.join(ROOT, 'data', 'meta.json')
    try:
        with open(p, encoding='utf-8') as f:
            m = json.load(f)
        for d in m.get('datasets', []):
            if d.get('id') == DS_ID:
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
    args = ap.parse_args()

    # 1) build
    bc = [PY, BUILD] + (['--check'] if args.check_only else [])
    run(bc, cwd=ROOT)
    if args.check_only:
        return

    as_of = read_as_of()
    print('\n📅 数据集 asOf =', as_of)

    st = {}
    if os.path.exists(STATE):
        try:
            st = json.load(open(STATE, encoding='utf-8'))
        except Exception:
            st = {}
    if not args.force and st.get('sent_asOf') == as_of and not args.no_send:
        print('⏭️ 该期（%s）已发送过，跳过。需重发请加 --force' % as_of)
        return

    # 2) bump
    bump_cache()

    # 3) push
    if not args.no_push:
        env = dict(os.environ)
        env['COMMIT_MSG'] = 'data(lc_inv): 碳酸锂下游（钢联）库存 更新至 %s' % as_of
        run([PY, DEPLOY, REPO] + PUSH_FILES, cwd=ROOT, env=env)
    else:
        print('[skip] 不推公网')

    # 4) shot
    if not args.no_shot:
        run([PY, SHOT, '--ds', DS_ID, '--out', IMG,
             '--width', '1080', '--height', '4200', '--scale', '2',
             '--title', '碳酸锂 · 下游周度库存',
             '--subtitle', '数据截至 %s｜我的钢铁网（Mysteel）周度调研｜单位：万吨' % as_of],
            cwd=ROOT)
    else:
        print('[skip] 不出图')

    # 5) send
    if args.no_send:
        print('\n[skip] 不发微信（--no-send）')
    else:
        if not os.path.exists(IMG):
            raise SystemExit('长图不存在：%s' % IMG)
        r = subprocess.run([PY, SEND, IMG, '--no-countdown'], cwd=ROOT)
        if r.returncode == 0:
            st['sent_asOf'] = as_of
            st['sentAt'] = datetime.datetime.now().isoformat(timespec='seconds')
            os.makedirs(os.path.dirname(STATE), exist_ok=True)
            json.dump(st, open(STATE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
            print('✅ 已发送并记账')
        else:
            print('⚠️ 微信发送未确认（rc=%d），未记账 → 下次运行会重试' % r.returncode)


if __name__ == '__main__':
    main()
