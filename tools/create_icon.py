from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter


ROOT = Path(__file__).resolve().parent.parent
ASSETS_DIR = ROOT / "assets"
PNG_PATH = ASSETS_DIR / "app_icon.png"
ICO_PATH = ASSETS_DIR / "app_icon.ico"
SIZE = 1024


def rotate_point(x: float, y: float, angle_deg: float, cx: float, cy: float) -> tuple[float, float]:
    angle = math.radians(angle_deg)
    tx, ty = x - cx, y - cy
    rx = tx * math.cos(angle) - ty * math.sin(angle)
    ry = tx * math.sin(angle) + ty * math.cos(angle)
    return rx + cx, ry + cy


def rotated_rect(cx: float, cy: float, width: float, height: float, angle_deg: float):
    hw = width / 2
    hh = height / 2
    points = [
        (cx - hw, cy - hh),
        (cx + hw, cy - hh),
        (cx + hw, cy + hh),
        (cx - hw, cy + hh),
    ]
    return [rotate_point(x, y, angle_deg, cx, cy) for x, y in points]


def draw_gear(draw: ImageDraw.ImageDraw, center: tuple[int, int]) -> None:
    cx, cy = center
    outer = 275
    ring = 215
    inner = 108
    tooth_w = 76
    tooth_h = 70
    gear_color = (31, 187, 211, 255)
    shadow_color = (6, 22, 38, 150)

    for angle in range(0, 360, 30):
        tooth_cx = cx + math.cos(math.radians(angle)) * (outer - tooth_h / 2)
        tooth_cy = cy + math.sin(math.radians(angle)) * (outer - tooth_h / 2)
        draw.polygon(rotated_rect(tooth_cx + 8, tooth_cy + 10, tooth_w, tooth_h, angle), fill=shadow_color)

    for angle in range(0, 360, 30):
        tooth_cx = cx + math.cos(math.radians(angle)) * (outer - tooth_h / 2)
        tooth_cy = cy + math.sin(math.radians(angle)) * (outer - tooth_h / 2)
        draw.polygon(rotated_rect(tooth_cx, tooth_cy, tooth_w, tooth_h, angle), fill=gear_color)

    draw.ellipse((cx - outer, cy - outer, cx + outer, cy + outer), fill=gear_color)
    draw.ellipse((cx - ring, cy - ring, cx + ring, cy + ring), fill=(14, 40, 63, 255))
    draw.ellipse((cx - inner, cy - inner, cx + inner, cy + inner), fill=(31, 187, 211, 255))


def draw_wrench(draw: ImageDraw.ImageDraw) -> None:
    body = (255, 183, 77, 255)
    highlight = (255, 226, 172, 210)
    dark = (164, 96, 31, 255)

    draw.rounded_rectangle((280, 620, 700, 720), radius=46, fill=body)
    draw.rounded_rectangle((300, 640, 680, 676), radius=18, fill=highlight)
    draw.polygon(
        [
            (620, 450),
            (720, 350),
            (795, 425),
            (730, 490),
            (778, 538),
            (720, 596),
            (672, 548),
            (606, 614),
            (532, 540),
        ],
        fill=body,
    )
    draw.polygon([(696, 372), (750, 318), (860, 428), (806, 482)], fill=body)
    draw.polygon([(759, 340), (812, 393), (786, 419), (733, 366)], fill=highlight)
    draw.arc((705, 295, 915, 505), start=318, end=42, fill=(14, 40, 63, 255), width=42)
    draw.line((324, 662, 656, 662), fill=dark, width=8)


def draw_clock(base: Image.Image) -> None:
    draw = ImageDraw.Draw(base)
    cx, cy = 760, 250
    radius = 168
    rim = (14, 40, 63, 255)
    face = (247, 250, 252, 255)
    accent = (255, 183, 77, 255)
    teal = (31, 187, 211, 255)

    shadow = Image.new("RGBA", base.size, (0, 0, 0, 0))
    ImageDraw.Draw(shadow).ellipse((cx - radius + 14, cy - radius + 18, cx + radius + 14, cy + radius + 18), fill=(0, 0, 0, 90))
    shadow = shadow.filter(ImageFilter.GaussianBlur(18))
    base.alpha_composite(shadow)

    draw.ellipse((cx - radius, cy - radius, cx + radius, cy + radius), fill=face, outline=rim, width=28)
    draw.ellipse((cx - radius + 26, cy - radius + 26, cx + radius - 26, cy + radius - 26), outline=(201, 214, 224, 255), width=6)

    for hour in range(12):
        angle = math.radians(hour * 30 - 90)
        inner = radius - 34
        outer = radius - 14
        x1 = cx + math.cos(angle) * inner
        y1 = cy + math.sin(angle) * inner
        x2 = cx + math.cos(angle) * outer
        y2 = cy + math.sin(angle) * outer
        draw.line((x1, y1, x2, y2), fill=teal, width=10)

    draw.line((cx, cy, cx, cy - 72), fill=rim, width=18)
    draw.line((cx, cy, cx + 62, cy + 38), fill=accent, width=18)
    draw.ellipse((cx - 18, cy - 18, cx + 18, cy + 18), fill=accent, outline=rim, width=6)


def create_icon() -> None:
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)

    if PNG_PATH.exists():
        icon = Image.open(PNG_PATH).convert("RGBA")
        icon.save(PNG_PATH)
        icon.save(ICO_PATH, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
        print(f"Created {PNG_PATH}")
        print(f"Created {ICO_PATH}")
        return

    canvas = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    shadow = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).ellipse((92, 112, 932, 952), fill=(0, 0, 0, 70))
    shadow = shadow.filter(ImageFilter.GaussianBlur(36))
    canvas.alpha_composite(shadow)

    draw = ImageDraw.Draw(canvas)
    draw.ellipse((72, 72, 952, 952), fill=(14, 40, 63, 245))
    draw.ellipse((104, 104, 920, 920), outline=(75, 110, 145, 170), width=8)

    draw_gear(draw, (420, 565))
    draw_wrench(draw)
    draw_clock(canvas)

    canvas.save(PNG_PATH)
    canvas.save(ICO_PATH, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])

    print(f"Created {PNG_PATH}")
    print(f"Created {ICO_PATH}")


if __name__ == "__main__":
    create_icon()
