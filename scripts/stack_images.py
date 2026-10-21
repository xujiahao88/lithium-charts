# -*- coding: utf-8 -*-
"""
stack_images.py — 多张长图纵向拼接（宽度不一致时统一缩放）

用途：站点按 tab 分页，一个 tab 出一张图；发微信时把几个 tab 拼成一张，
      避免连发多条。每段之间可加一条浅灰分隔线。

用法：
  python scripts/stack_images.py --out shot/碳酸锂下游_长图.png ^
      --inputs shot/_a.png shot/_b.png shot/_c.png --gap 28
"""
import argparse
import os

from PIL import Image, ImageDraw


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', required=True)
    ap.add_argument('--inputs', nargs='+', required=True)
    ap.add_argument('--gap', type=int, default=24, help='段间留白像素')
    ap.add_argument('--bg', default='#F4F6FA')
    ap.add_argument('--no-rule', dest='rule', action='store_false', help='不画分隔线')
    args = ap.parse_args()

    paths = [p for p in args.inputs if os.path.exists(p)]
    missing = [p for p in args.inputs if not os.path.exists(p)]
    if missing:
        print('[warn] 跳过不存在的图：%s' % '，'.join(missing))
    if not paths:
        raise SystemExit('没有可拼接的图')

    imgs = [Image.open(p).convert('RGB') for p in paths]
    width = max(i.width for i in imgs)
    for idx, im in enumerate(imgs):
        if im.width != width:
            h = int(round(im.height * width / im.width))
            imgs[idx] = im.resize((width, h), Image.LANCZOS)
            print('[ok] 缩放 %s → %dx%d' % (os.path.basename(paths[idx]), width, h))

    total_h = sum(i.height for i in imgs) + args.gap * (len(imgs) - 1) + args.gap
    canvas = Image.new('RGB', (width, total_h), args.bg)
    draw = ImageDraw.Draw(canvas)

    y = 0
    for i, im in enumerate(imgs):
        canvas.paste(im, (0, y))
        y += im.height
        if i < len(imgs) - 1:
            if args.rule:
                draw.line([(40, y + args.gap // 2), (width - 40, y + args.gap // 2)],
                          fill='#D8DEE9', width=2)
            y += args.gap

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    canvas.save(args.out)
    print('✅ 拼接完成: %s (%dx%d, %.1f MB)'
          % (args.out, canvas.width, canvas.height,
             os.path.getsize(args.out) / 1024 / 1024))


if __name__ == '__main__':
    main()
