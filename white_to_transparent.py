#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
把白色/近白色背景的图标转成透明 PNG。

依赖：
    pip install pillow numpy

用法：
    python white_to_transparent.py bell.png
    python white_to_transparent.py bell.png -o bell_transparent.png
"""

import argparse
from pathlib import Path
import numpy as np
from PIL import Image


def estimate_bg(rgb, border_ratio=0.08):
    h, w, _ = rgb.shape
    bh = max(1, int(h * border_ratio))
    bw = max(1, int(w * border_ratio))

    samples = np.concatenate([
        rgb[:bh].reshape(-1, 3),
        rgb[-bh:].reshape(-1, 3),
        rgb[:, :bw].reshape(-1, 3),
        rgb[:, -bw:].reshape(-1, 3),
    ], axis=0)

    brightness = samples.mean(axis=1)
    bright = samples[brightness >= np.percentile(brightness, 60)]
    if len(bright) < 10:
        bright = samples

    return np.median(bright, axis=0)


def remove_white_bg(img, hard=10.0, soft=38.0, border_ratio=0.08):
    rgba = img.convert("RGBA")
    arr = np.array(rgba).astype(np.float32)
    rgb = arr[:, :, :3]

    bg = estimate_bg(rgb, border_ratio)

    diff = rgb - bg.reshape(1, 1, 3)
    dist = np.sqrt(np.sum(diff * diff, axis=2))

    alpha = np.zeros_like(dist, dtype=np.float32)
    alpha[dist >= soft] = 255

    mid = (dist > hard) & (dist < soft)
    alpha[mid] = (dist[mid] - hard) / (soft - hard) * 255

    original_alpha = arr[:, :, 3]
    alpha *= original_alpha / 255.0

    # 对近白、低饱和像素进一步减弱，减少白边
    maxc = rgb.max(axis=2)
    minc = rgb.min(axis=2)
    saturation = maxc - minc
    near_white = (maxc > 235) & (saturation < 22)
    alpha[near_white] *= 0.35

    out = arr.copy()
    out[:, :, 3] = np.clip(alpha, 0, 255)

    return Image.fromarray(out.astype(np.uint8), "RGBA"), tuple(int(round(x)) for x in bg)


def trim_transparent(img, padding=4):
    alpha = np.array(img.getchannel("A"))
    ys, xs = np.where(alpha > 2)

    if len(xs) == 0:
        return img

    left = max(0, int(xs.min()) - padding)
    top = max(0, int(ys.min()) - padding)
    right = min(img.width, int(xs.max()) + 1 + padding)
    bottom = min(img.height, int(ys.max()) + 1 + padding)

    return img.crop((left, top, right, bottom))


def main():
    parser = argparse.ArgumentParser(description="白底图标转透明 PNG")
    parser.add_argument("input", help="输入图片，例如 bell.png")
    parser.add_argument("-o", "--output", help="输出 PNG")
    parser.add_argument("--hard", type=float, default=10.0,
                        help="完全透明阈值，默认 10")
    parser.add_argument("--soft", type=float, default=38.0,
                        help="完全不透明阈值，默认 38")
    parser.add_argument("--border-ratio", type=float, default=0.08)
    parser.add_argument("--no-crop", action="store_true")
    parser.add_argument("--padding", type=int, default=4)
    args = parser.parse_args()

    src = Path(args.input)
    if not src.exists():
        raise SystemExit(f"输入文件不存在：{src}")

    if args.soft <= args.hard:
        raise SystemExit("--soft 必须大于 --hard")

    dst = Path(args.output) if args.output else src.with_name(src.stem + "_transparent.png")

    img = Image.open(src)
    result, bg = remove_white_bg(
        img,
        hard=args.hard,
        soft=args.soft,
        border_ratio=args.border_ratio
    )

    if not args.no_crop:
        result = trim_transparent(result, args.padding)

    dst.parent.mkdir(parents=True, exist_ok=True)
    result.save(dst, "PNG")

    print("=" * 52)
    print("White Background -> Transparent PNG")
    print("=" * 52)
    print(f"Input       : {src}")
    print(f"Detected BG : RGB{bg}")
    print(f"Hard        : {args.hard}")
    print(f"Soft        : {args.soft}")
    print(f"Output size : {result.width} x {result.height}")
    print(f"Output      : {dst}")
    print("=" * 52)


if __name__ == "__main__":
    main()
