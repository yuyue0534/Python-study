#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
只替换弹窗目标 Icon 的 Python Demo

功能：
1. 读取 Before 图片
2. 自动判断横版 / 竖版
3. 自动定位浅色 Modal（针对“深色遮罩 + 浅色弹窗”的 UI）
4. 根据 Modal 的相对位置确定 Icon Safe ROI
5. 仅在 Safe ROI 内恢复背景并覆盖 bell.png / bell.svg
6. 输出 result.png
7. 输出 diff.png（变化区域高亮）
8. 打印 ROI 外发生变化的像素数

依赖：
    pip install pillow opencv-python numpy cairosvg

用法：
    python replace_icon_demo.py before.jpg bell.svg
    python replace_icon_demo.py before.jpg bell.png --out result.png --diff diff.png

说明：
- 推荐输入 PNG 原始截图；JPG 输入可读，但输出固定为 PNG，避免再次 JPEG 压缩导致全图像素变化。
- 这个 Demo 是“受限 ROI 替换”，不会重绘整张图。
- Modal 检测针对当前测试用例的 UI 特征：深色遮罩上的大块浅色圆角弹窗。
"""

from __future__ import annotations

import argparse
import io
import math
from pathlib import Path

import cv2
import numpy as np
from PIL import Image


def load_rgba_icon(path: str, target_size: tuple[int, int]) -> Image.Image:
    """读取 PNG 或 SVG，并缩放为 target_size。"""
    p = Path(path)
    suffix = p.suffix.lower()

    if suffix == ".svg":
        try:
            import cairosvg
        except ImportError as exc:
            raise RuntimeError(
                "读取 SVG 需要 cairosvg：pip install cairosvg"
            ) from exc

        png_bytes = cairosvg.svg2png(
            url=str(p),
            output_width=max(target_size[0] * 4, 256),
            output_height=max(target_size[1] * 4, 256),
        )
        icon = Image.open(io.BytesIO(png_bytes)).convert("RGBA")
    else:
        icon = Image.open(p).convert("RGBA")

    icon.thumbnail(target_size, Image.Resampling.LANCZOS)

    # 放到透明画布中央，保证长宽比例不被拉伸
    canvas = Image.new("RGBA", target_size, (0, 0, 0, 0))
    x = (target_size[0] - icon.width) // 2
    y = (target_size[1] - icon.height) // 2
    canvas.alpha_composite(icon, (x, y))
    return canvas


def rect_mean(gray: np.ndarray, x: int, y: int, w: int, h: int) -> float:
    patch = gray[max(0, y):min(gray.shape[0], y+h),
                 max(0, x):min(gray.shape[1], x+w)]
    return float(patch.mean()) if patch.size else 0.0


def detect_modal(img_bgr: np.ndarray) -> tuple[int, int, int, int]:
    """
    检测浅色弹窗。

    思路：
    - 多个亮度阈值生成候选区域
    - 轮廓闭运算连接弹窗内部文字造成的小洞
    - 根据面积、长宽比、位置、内部亮度、与周围背景的亮度差打分
    - 横/竖版采用不同的宽高先验，但都不写死绝对像素
    """
    h, w = img_bgr.shape[:2]
    landscape = w >= h
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)

    candidates = []

    # 用较小结构元素，避免把外部白纸和 Modal 粘在一起
    k = max(5, int(min(w, h) * 0.006))
    if k % 2 == 0:
        k += 1
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (k, k))

    for threshold in (125, 140, 155, 170, 185, 200):
        mask = (gray >= threshold).astype(np.uint8) * 255
        mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=1)

        contours, _ = cv2.findContours(
            mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE
        )

        for c in contours:
            x, y, ww, hh = cv2.boundingRect(c)
            area_frac = (ww * hh) / float(w * h)
            aspect = ww / float(max(hh, 1))

            if landscape:
                if not (0.12 <= area_frac <= 0.60):
                    continue
                if not (1.25 <= aspect <= 3.2):
                    continue
                if ww < w * 0.45 or hh < h * 0.25:
                    continue
            else:
                if not (0.08 <= area_frac <= 0.55):
                    continue
                if not (0.45 <= aspect <= 1.35):
                    continue
                if ww < w * 0.28 or hh < h * 0.28:
                    continue

            # Modal 大体应在画面中部，而不是最外层白色页面
            cx = x + ww / 2
            cy = y + hh / 2
            center_dx = abs(cx - w / 2) / (w / 2)
            center_dy = abs(cy - h / 2) / (h / 2)
            if center_dx > 0.55 or center_dy > 0.70:
                continue

            inner = rect_mean(gray, x, y, ww, hh)

            # 取弹窗外侧一圈估算遮罩亮度
            pad = max(8, int(min(ww, hh) * 0.05))
            ox1, oy1 = max(0, x-pad), max(0, y-pad)
            ox2, oy2 = min(w, x+ww+pad), min(h, y+hh+pad)
            outer_patch = gray[oy1:oy2, ox1:ox2].copy()
            inner_mask = np.zeros_like(outer_patch, dtype=bool)
            ix1, iy1 = x-ox1, y-oy1
            ix2, iy2 = ix1+ww, iy1+hh
            inner_mask[iy1:iy2, ix1:ix2] = True
            ring = outer_patch[~inner_mask]
            outer = float(np.median(ring)) if ring.size else inner

            contrast = inner - outer

            # 候选打分：亮、比周边亮、位置合理、尺寸接近典型 Modal
            score = (
                inner * 0.025
                + max(contrast, -30) * 0.10
                - center_dx * 1.5
                - center_dy * 0.5
            )

            if landscape:
                score -= abs(aspect - 1.85) * 0.8
            else:
                score -= abs(aspect - 0.80) * 0.8

            candidates.append((score, x, y, ww, hh, inner, contrast))

    if not candidates:
        return detect_modal_fallback(img_bgr)

    candidates.sort(reverse=True, key=lambda t: t[0])
    best = candidates[0]
    _, x, y, ww, hh, _, _ = best

    # 对候选框进行一次边缘微调：寻找 Modal 的真实亮度边界
    return refine_modal(gray, (x, y, ww, hh))


def detect_modal_fallback(img_bgr: np.ndarray) -> tuple[int, int, int, int]:
    """
    兜底：根据当前 UI 的布局先验，给出相对坐标区域。
    仅在自动轮廓检测失败时使用。
    """
    h, w = img_bgr.shape[:2]
    if w >= h:
        return (
            int(w * 0.155),
            int(h * 0.185),
            int(w * 0.75),
            int(h * 0.54),
        )
    else:
        return (
            int(w * 0.125),
            int(h * 0.225),
            int(w * 0.42),
            int(h * 0.38),
        )


def refine_modal(gray: np.ndarray, rect: tuple[int, int, int, int]):
    """
    轻量微调，不追求圆角逐像素，而是让 bbox 更贴近亮色主体。
    """
    h, w = gray.shape
    x, y, ww, hh = rect

    # 防止候选因为阈值只截到 Modal 内部，向外留少量安全余量
    ex = int(ww * 0.015)
    ey = int(hh * 0.015)
    x = max(0, x - ex)
    y = max(0, y - ey)
    ww = min(w - x, ww + ex * 2)
    hh = min(h - y, hh + ey * 2)

    return x, y, ww, hh


def restore_old_icon(roi_bgr: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    """
    只在 ROI 中央的旧 Icon 占用区做 inpaint，不再把整个 ROI 填成纯色。

    为什么这样更合适：
    - 手机拍摄图存在摩尔纹/亮度渐变，整块填色会产生明显矩形补丁；
    - inpaint 仅修改 mask 内部，ROI 其余像素仍逐像素保持原样；
    - 对真实 PNG UI 截图，浅色纯背景上的恢复会更自然。

    返回：
        clean_roi, erase_mask
    """
    h, w = roi_bgr.shape[:2]

    mask = np.zeros((h, w), dtype=np.uint8)

    # 旧 Icon 的安全包围区域。
    # 这是 ROI 内的相对坐标，不依赖整张图的绝对分辨率。
    left   = int(w * 0.10)
    right  = int(w * 0.90)
    top    = int(h * 0.05)
    bottom = int(h * 0.94)

    # 圆角矩形 mask，减少修复边缘的突兀感
    radius = max(8, int(min(w, h) * 0.08))
    cv2.rectangle(mask, (left + radius, top), (right - radius, bottom), 255, -1)
    cv2.rectangle(mask, (left, top + radius), (right, bottom - radius), 255, -1)
    cv2.circle(mask, (left + radius, top + radius), radius, 255, -1)
    cv2.circle(mask, (right - radius, top + radius), radius, 255, -1)
    cv2.circle(mask, (left + radius, bottom - radius), radius, 255, -1)
    cv2.circle(mask, (right - radius, bottom - radius), radius, 255, -1)

    # 轻微膨胀，覆盖旧图标抗锯齿边缘
    k = max(3, int(min(w, h) * 0.015))
    if k % 2 == 0:
        k += 1
    mask = cv2.dilate(
        mask,
        cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (k, k)),
        iterations=1,
    )

    # TELEA 对这种浅色、缓慢变化背景比较合适
    clean = cv2.inpaint(roi_bgr, mask, 5, cv2.INPAINT_TELEA)
    return clean, mask


def get_icon_safe_roi(
    modal: tuple[int, int, int, int],
    image_shape: tuple[int, int, int],
) -> tuple[int, int, int, int]:
    """
    根据 Modal 相对坐标定义“只允许修改”的 Icon Safe ROI。

    横版测试图：
      Icon 位于 Modal 上部中央，标题在其下方。
    竖版测试图：
      Icon 同样位于 Modal 上部中央，但相对占比更大。

    ROI 故意略大于旧 Icon，确保旧图形不会残留；
    同时与标题、关闭按钮之间保留安全距离。
    """
    x, y, mw, mh = modal
    ih, iw = image_shape[:2]
    landscape = iw >= ih

    if landscape:
        roi_w = int(mw * 0.24)
        roi_h = int(mh * 0.38)
        cx = x + mw * 0.50
        top = y + mh * 0.075
    else:
        roi_w = int(mw * 0.50)
        roi_h = int(mh * 0.37)
        cx = x + mw * 0.50
        top = y + mh * 0.07

    x1 = int(round(cx - roi_w / 2))
    y1 = int(round(top))
    x2 = x1 + roi_w
    y2 = y1 + roi_h

    x1 = max(0, min(iw - 1, x1))
    y1 = max(0, min(ih - 1, y1))
    x2 = max(x1 + 1, min(iw, x2))
    y2 = max(y1 + 1, min(ih, y2))
    return x1, y1, x2, y2


def replace_icon(
    before_path: str,
    bell_path: str,
    out_path: str,
    diff_path: str,
) -> dict:
    """
    核心流程。
    """
    original_bgr = cv2.imread(before_path, cv2.IMREAD_COLOR)
    if original_bgr is None:
        raise FileNotFoundError(f"无法读取图片：{before_path}")

    h, w = original_bgr.shape[:2]
    orientation = "landscape" if w >= h else "portrait"

    modal = detect_modal(original_bgr)
    x1, y1, x2, y2 = get_icon_safe_roi(modal, original_bgr.shape)

    # 只复制一份作为输出；后续严格只改 [y1:y2, x1:x2]
    result_bgr = original_bgr.copy()
    roi = result_bgr[y1:y2, x1:x2]

    # 1) 仅在旧 Icon 的局部包围区域恢复背景
    clean_roi, erase_mask = restore_old_icon(roi)

    # 2) 加载新 Bell
    rw, rh = x2 - x1, y2 - y1
    if orientation == "landscape":
        icon_w = int(rw * 0.56)
        icon_h = int(rh * 0.72)
    else:
        icon_w = int(rw * 0.72)
        icon_h = int(rh * 0.78)

    bell = load_rgba_icon(bell_path, (icon_w, icon_h))

    # 3) 将 clean ROI 转 RGBA，并 Alpha Composite
    roi_rgba = Image.fromarray(
        cv2.cvtColor(clean_roi, cv2.COLOR_BGR2RGBA)
    )

    px = (rw - bell.width) // 2
    py = (rh - bell.height) // 2
    roi_rgba.alpha_composite(bell, (px, py))

    composed = cv2.cvtColor(
        np.array(roi_rgba), cv2.COLOR_RGBA2BGR
    )

    # 唯一写回位置：Icon Safe ROI
    result_bgr[y1:y2, x1:x2] = composed

    # 输出 PNG，避免 JPG 再编码污染整图
    cv2.imwrite(out_path, result_bgr)

    # ========== Pixel Diff 验收 ==========
    diff = cv2.absdiff(original_bgr, result_bgr)
    changed = np.any(diff != 0, axis=2)

    allowed = np.zeros((h, w), dtype=bool)
    allowed[y1:y2, x1:x2] = True

    outside_changed = changed & (~allowed)
    inside_changed = changed & allowed

    outside_count = int(outside_changed.sum())
    inside_count = int(inside_changed.sum())
    total_changed = int(changed.sum())

    # diff.png：原图暗化 + 变化像素标红
    preview = (original_bgr.astype(np.float32) * 0.30).astype(np.uint8)
    preview[changed] = (0, 0, 255)

    # 再画允许区域边框，便于调试
    cv2.rectangle(
        preview, (x1, y1), (x2 - 1, y2 - 1), (0, 255, 255), 3
    )
    cv2.imwrite(diff_path, preview)

    mx, my, mw, mh = modal
    report = {
        "input": before_path,
        "orientation": orientation,
        "image_size": [w, h],
        "modal": [mx, my, mw, mh],
        "icon_safe_roi": [x1, y1, x2, y2],
        "changed_pixels_total": total_changed,
        "changed_pixels_inside_roi": inside_count,
        "changed_pixels_outside_roi": outside_count,
        "validation": "PASS" if outside_count == 0 else "FAIL",
        "result": out_path,
        "diff": diff_path,
    }

    return report


def main():
    parser = argparse.ArgumentParser(
        description="只替换 Modal 中目标 Icon，并做 ROI 外像素零变化校验"
    )
    parser.add_argument("before", help="Before 图片路径")
    parser.add_argument("bell", help="新 Bell 图标路径（PNG/SVG）")
    parser.add_argument("--out", default="result.png", help="结果 PNG")
    parser.add_argument("--diff", default="diff.png", help="Diff PNG")
    args = parser.parse_args()

    report = replace_icon(
        args.before,
        args.bell,
        args.out,
        args.diff,
    )

    print("=" * 64)
    print("Icon Replace Demo")
    print("=" * 64)
    print(f"Input              : {report['input']}")
    print(f"Orientation        : {report['orientation']}")
    print(f"Image size         : {report['image_size'][0]} x {report['image_size'][1]}")
    print(f"Modal bbox         : {report['modal']}")
    print(f"Icon Safe ROI      : {report['icon_safe_roi']}")
    print(f"Changed total      : {report['changed_pixels_total']}")
    print(f"Changed inside ROI : {report['changed_pixels_inside_roi']}")
    print(f"Changed outside ROI: {report['changed_pixels_outside_roi']}")
    print(f"Validation         : {report['validation']}")
    print(f"Result             : {report['result']}")
    print(f"Diff               : {report['diff']}")
    print("=" * 64)


if __name__ == "__main__":
    main()
