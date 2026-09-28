# -*- coding: utf-8 -*-
"""生成桌面快捷方式图标 转写.ico（圆角渐变底 + 声波）。

需要 Pillow：pip install pillow
重新生成：python make_icon.py
"""
from pathlib import Path

from PIL import Image, ImageDraw

OUT = Path(__file__).resolve().parent / "转写.ico"

SUPER = 1024                      # 超采样画布，最后降采样得到锐利边缘
TOP = (79, 70, 229)               # indigo-600
BOTTOM = (124, 58, 237)           # violet-600
BAR_RATIOS = (0.34, 0.62, 0.86, 0.62, 0.34)
ICON_SIZES = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]


def rounded_gradient(size, radius_ratio=0.22):
    grad = Image.new("RGB", (1, size))
    for y in range(size):
        t = y / max(size - 1, 1)
        grad.putpixel((0, y), tuple(
            int(TOP[i] + (BOTTOM[i] - TOP[i]) * t) for i in range(3)))
    grad = grad.resize((size, size))

    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, size - 1, size - 1], radius=int(size * radius_ratio), fill=255)

    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(grad, (0, 0), mask)
    return out


def waveform(size):
    layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    bar_w = size * 0.085
    gap = size * 0.052
    total = len(BAR_RATIOS) * bar_w + (len(BAR_RATIOS) - 1) * gap
    x, cy = (size - total) / 2, size / 2
    for ratio in BAR_RATIOS:
        half = size * ratio / 2
        draw.rounded_rectangle([x, cy - half, x + bar_w, cy + half],
                               radius=bar_w / 2, fill=(255, 255, 255, 242))
        x += bar_w + gap
    return layer


def main():
    icon = rounded_gradient(SUPER)
    icon.alpha_composite(waveform(SUPER))
    icon.resize((256, 256), Image.LANCZOS).save(OUT, format="ICO",
                                                sizes=ICON_SIZES)
    print(f"已生成：{OUT}  ({OUT.stat().st_size} 字节，含 {len(ICON_SIZES)} 种尺寸)")


if __name__ == "__main__":
    main()