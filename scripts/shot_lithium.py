# -*- coding: utf-8 -*-
"""
shot_lithium.py — lithium-charts 长图截图（Edge headless 单图工具）

用法：
  python shot_lithium.py --ds lc_inv --out "C:/path/to/long.png"
  python shot_lithium.py --ds lc_inv --out "C:/path/to/long.png" --width 1920 --height 2600

说明：
  · lithium-charts 前端（沿用钢材站骨架）支持 ?ds=<id> 切 tab，?shot=1 截图模式
  · 这里用 Edge headless 一次性出图（7 张图 + 汇总表）
  · Edge 路径自动探测（程序文件/Microsoft Edge + Program Files (x86)/Microsoft Edge）
"""
import argparse
import os
import shutil
import subprocess

from PIL import Image, ImageDraw, ImageFont


def _load_font(size, bold=False):
    candidates = []
    if bold:
        candidates += [
            "C:/Windows/Fonts/msyhbd.ttc",
            "C:/Windows/Fonts/simhei.ttf",
            "C:/Windows/Fonts/simsun.ttc",
            "C:/Windows/Fonts/msyh.ttc",
        ]
    candidates += [
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/simsun.ttc",
    ]
    for p in candidates:
        if os.path.isfile(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()


def add_title_overlay(png_path, title, subtitle=None, title_height=180,
                      bg="#FFFFFF", title_color="#1F2937", subtitle_color="#DC2626"):
    """在 PNG 顶部加标题条，返回新路径（覆盖原文件）"""
    img = Image.open(png_path).convert("RGB")
    w = img.width
    new = Image.new("RGB", (w, img.height + title_height), bg)
    draw = ImageDraw.Draw(new)
    # 标题字体
    font_title = _load_font(48, bold=True)
    font_sub = _load_font(26)
    # 主标题居中
    bbox = draw.textbbox((0, 0), title, font=font_title)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    y_title = (title_height - th) // 2 - (20 if subtitle else 0)
    draw.text(((w - tw) // 2, y_title), title, fill=title_color, font=font_title)
    # 副标题
    if subtitle:
        bbox2 = draw.textbbox((0, 0), subtitle, font=font_sub)
        tw2 = bbox2[2] - bbox2[0]
        draw.text(((w - tw2) // 2, y_title + th + 10), subtitle,
                  fill=subtitle_color, font=font_sub)
    # 底部灰线
    draw.line([(0, title_height - 1), (w, title_height - 1)], fill="#E5E7EB", width=1)
    # 贴上原图
    new.paste(img, (0, title_height))
    new.save(png_path)
    return png_path


def trim_bottom_whitespace(png_path, margin=45, dark_thresh=180):
    """裁掉 PNG 底部空白（页面背景色区域）。

    原理：从底向上扫，找最后一行含"深色内容像素"(min通道<=dark_thresh，
    即文字/边框/柱体/轴线) 的行，截到该行 + margin。页面浅灰背景(≈244)
    与卡片白底(255) 都不会被当作内容，从而精准去掉底部留白。
    返回输出路径。
    """
    import numpy as np
    img = Image.open(png_path).convert("RGB")
    a = np.asarray(img)
    h, w, _ = a.shape
    dark = a.min(axis=2) <= dark_thresh      # 每行是否有内容像素
    last = 0
    for y in range(h - 1, -1, -1):
        if dark[y].any():
            last = y
            break
    if last == 0:                            # 没找到内容（异常），原样返回
        return png_path
    crop_h = min(h, last + margin)
    if crop_h >= h - 2:                       # 本就没空白
        return png_path
    img.crop((0, 0, w, crop_h)).save(png_path)
    print(f"✂️ 裁掉底部空白: {png_path} 高度 {h}→{crop_h} (去掉 {h - crop_h}px)")
    return png_path


def find_edge():
    for p in ("C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
              "C:/Program Files/Microsoft/Edge/Application/msedge.exe"):
        if os.path.isfile(p):
            return p
    return shutil.which("msedge") or shutil.which("msedge.exe")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ds", required=True, help="数据集 id (如 gangyin_stock / juanluo_luowen)")
    ap.add_argument("--out", required=True, help="输出 PNG 绝对路径")
    ap.add_argument("--site", default=r"C:\Users\Administrator\lithium-charts",
                    help="本地站点根目录（含 index.html）")
    ap.add_argument("--width", type=int, default=1920)
    ap.add_argument("--height", type=int, default=3000, help="7 图 + 汇总表约 3000")
    ap.add_argument("--scale", type=float, default=1.0,
                    help="设备缩放比例（不影响 CSS 布局，仅提高像素密度）。"
                         "1.333 → 1920x2600 出 2560x3467（2K）；2.0 → 3840x5200（4K）")
    ap.add_argument("--url", default=None, help="自定义 URL（覆盖默认 site 构造）")
    ap.add_argument("--title", default=None, help="截图顶部大标题")
    ap.add_argument("--subtitle", default=None, help="截图顶部副标题（默认空）")
    ap.add_argument("--no-trim", dest="trim", action="store_false",
                    help="关闭底部自动裁白边（默认开启）")
    ap.add_argument("--trim-margin", type=int, default=45,
                    help="裁白边时在最后内容行下方保留的留白像素，默认45")
    args = ap.parse_args()

    edge = find_edge()
    if not edge:
        raise SystemExit("未找到 Edge，无法截图")

    # ⚠️ 2026-09-14 踩坑：--out 传相对路径时，Edge 会按「自己的 cwd」写文件
    #    → 图写到别处，脚本却报"截图未生成"，且旧图已在下面被删掉。
    #    统一转绝对路径。
    args.out = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(args.out), exist_ok=True)

    # 先截到临时文件，成功后再原子替换 —— 失败也不会毁掉上一版好图。
    tmp_out = args.out + ".tmp.png"
    if os.path.exists(tmp_out):
        try:
            os.remove(tmp_out)
        except Exception:
            pass

    if args.url:
        url = args.url
    else:
        index = os.path.join(args.site, "index.html").replace("\\", "/")
        url = "file:///" + index + f"?ds={args.ds}&shot=1"

    cmd = [
        edge, "--headless=new", "--disable-gpu", "--hide-scrollbars",
        "--virtual-time-budget=9000",
        f"--screenshot={tmp_out}",
        f"--window-size={args.width},{args.height}",
    ]
    # --force-device-scale-factor 只提高输出像素密度，不改变 CSS 布局
    if args.scale and args.scale != 1.0:
        cmd.append(f"--force-device-scale-factor={args.scale}")
    cmd.append(url)
    print("$ " + " ".join(cmd[:5]) + f" … --window-size={args.width},{args.height}"
          + (f" scale={args.scale}" if args.scale != 1.0 else "") + f" {url}")
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if r.returncode != 0:
        raise SystemExit(f"Edge 截图失败 rc={r.returncode}: {(r.stderr or r.stdout)[:300]}")
    if not os.path.exists(tmp_out) or os.path.getsize(tmp_out) < 5000:
        raise SystemExit(f"截图未生成或过小：{tmp_out}")
    # 成功 → 原子替换目标文件
    try:
        os.replace(tmp_out, args.out)
    except Exception:
        import shutil
        shutil.move(tmp_out, args.out)
    if args.title:
        add_title_overlay(args.out, args.title, args.subtitle)
    if args.trim:
        trim_bottom_whitespace(args.out, margin=args.trim_margin)
    sz = os.path.getsize(args.out)
    print(f"✅ 截图完成: {args.out} ({sz:,} bytes, {args.width}x{args.height})")


if __name__ == "__main__":
    main()
