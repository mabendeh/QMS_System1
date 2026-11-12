"""Generate 4K blueprint-style frozen-bakery production-line artwork (original line art).

Outputs:
  hero.png     - title slide background (3840x2160)
  section.png  - dimmer variant for section / closing slides
These are placeholders for authentic plant photography.
"""
import math, sys
from PIL import Image, ImageDraw, ImageFilter, ImageFont

W, H = 3840, 2160
OUT = sys.argv[1] if len(sys.argv) > 1 else "."
FONT = "/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf"


def gradient(top, bottom):
    img = Image.new("RGB", (W, H))
    px = img.load()
    for y in range(H):
        t = y / (H - 1)
        c = tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3))
        for x in range(W):
            px[x, y] = c
    return img


def draw_scene(alpha_scale=1.0, machines=True):
    base = gradient((9, 26, 50), (17, 45, 84)).convert("RGBA")
    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    A = lambda a: int(a * alpha_scale)
    grid = (140, 175, 215, A(22))
    for x in range(0, W, 80):
        d.line([(x, 0), (x, H)], fill=grid, width=1)
    for y in range(0, H, 80):
        d.line([(0, y), (W, y)], fill=grid, width=1)
    grid2 = (140, 175, 215, A(38))
    for x in range(0, W, 400):
        d.line([(x, 0), (x, H)], fill=grid2, width=2)
    for y in range(0, H, 400):
        d.line([(0, y), (W, y)], fill=grid2, width=2)

    if not machines:
        img = Image.alpha_composite(base, layer)
        return img.convert("RGB")
    line = (190, 215, 240, A(170))
    soft = (190, 215, 240, A(90))
    green = (80, 190, 140, A(200))
    lw = 5
    floor = 1820
    d.line([(0, floor), (W, floor)], fill=soft, width=3)

    # Mixer (bowl on column)
    mx = 260
    d.rectangle([mx, floor - 540, mx + 90, floor], outline=line, width=lw)
    d.rectangle([mx - 40, floor - 600, mx + 420, floor - 520], outline=line, width=lw)
    d.chord([mx + 110, floor - 640, mx + 470, floor - 260], 0, 180, outline=line, width=lw)
    d.line([(mx + 110, floor - 450), (mx + 470, floor - 450)], fill=line, width=lw)
    d.line([(mx + 290, floor - 520), (mx + 290, floor - 330)], fill=line, width=lw)
    d.arc([mx + 220, floor - 420, mx + 360, floor - 300], 0, 360, fill=soft, width=4)
    d.rectangle([mx - 60, floor - 30, mx + 520, floor], outline=line, width=lw)

    # Dough sheeter rollers + conveyor
    cx0, cx1, cy = 820, 3020, floor - 420
    d.line([(cx0, cy), (cx1, cy)], fill=line, width=lw)
    d.line([(cx0, cy + 40), (cx1, cy + 40)], fill=line, width=lw)
    for x in range(cx0 + 40, cx1, 220):
        d.line([(x, cy + 40), (x - 30, floor)], fill=soft, width=4)
        d.line([(x, cy + 40), (x + 30, floor)], fill=soft, width=4)
    for (rx, r) in [(900, 70), (1060, 55)]:
        d.ellipse([rx - r, cy - 2 * r - 10, rx + r, cy - 10], outline=line, width=lw)
        d.ellipse([rx - 12, cy - r - 22, rx + 12, cy - r + 2], fill=line)
    # dough sheet then cut pieces
    d.line([(1150, cy - 18), (1400, cy - 18)], fill=line, width=10)
    for x in range(1440, 1900, 95):
        d.ellipse([x, cy - 46, x + 70, cy - 4], outline=line, width=4)

    # Tunnel oven
    ox0, ox1 = 1920, 2620
    d.rounded_rectangle([ox0, cy - 300, ox1, cy + 60], radius=24, outline=line, width=lw)
    for x in range(ox0 + 60, ox1 - 40, 120):
        for k in range(3):
            y0 = cy - 250 + k * 18
            d.arc([x, y0, x + 60, y0 + 40], 200, 340, fill=(230, 160, 90, A(150)), width=4)

    # Spiral freezer (cylinder with helix)
    sx, sw = 2820, 560
    top, bot = floor - 1240, floor - 40
    d.rectangle([sx, top, sx + sw, bot], outline=line, width=lw)
    d.ellipse([sx, top - 60, sx + sw, top + 60], outline=line, width=lw)
    for i in range(14):
        y = top + 60 + i * 80
        d.arc([sx + 30, y - 40, sx + sw - 30, y + 40], 0, 180, fill=soft, width=4)
    # snowflake
    fx, fy, fr = sx + sw // 2, top - 150, 80
    for k in range(6):
        a = math.radians(k * 60)
        x2, y2 = fx + fr * math.cos(a), fy + fr * math.sin(a)
        d.line([(fx, fy), (x2, y2)], fill=green, width=7)
        for s in (0.55,):
            bx, by = fx + fr * s * math.cos(a), fy + fr * s * math.sin(a)
            for da in (-40, 40):
                b = a + math.radians(180 + da)
                d.line([(bx, by), (bx + 35 * math.cos(b + math.pi), by + 35 * math.sin(b + math.pi))], fill=green, width=6)
    # Conveyor into freezer / out to pack
    d.line([(3020, cy), (sx, cy)], fill=line, width=lw)

    # Packaging cases
    px0 = 3460
    for i, (bx, by) in enumerate([(px0, floor - 220), (px0 + 200, floor - 220), (px0 + 100, floor - 420)]):
        d.rectangle([bx, by, bx + 180, by + 190], outline=line, width=lw)
        d.line([(bx, by + 50), (bx + 180, by + 50)], fill=soft, width=4)

    # Station labels (engineering-drawing style)
    f = ImageFont.truetype(FONT, 44)
    labels = [(mx + 120, "MIX"), (980, "SHEET"), (1640, "FORM"), (2200, "BAKE"), (sx + 200, "FREEZE"), (px0 + 90, "PACK")]
    for x, t in labels:
        d.line([(x + 30, floor + 40), (x + 30, floor + 110)], fill=soft, width=3)
        d.text((x, floor + 130), t, font=f, fill=(190, 215, 240, A(200)))
    # dimension line
    d.line([(260, floor + 290), (3640, floor + 290)], fill=soft, width=3)
    for x in (260, 3640):
        d.line([(x, floor + 260), (x, floor + 320)], fill=soft, width=3)
    d.text((1700, floor + 250), "PROCESS FLOW", font=ImageFont.truetype(FONT, 36), fill=(190, 215, 240, A(150)))

    img = Image.alpha_composite(base, layer)
    # darken upper-left for title legibility
    shade = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shade)
    for y in range(H):
        t = max(0.0, 1 - y / (H * 0.75))
        sd.line([(0, y), (W, y)], fill=(6, 18, 36, int(150 * t)))
    img = Image.alpha_composite(img, shade)
    return img.convert("RGB")


draw_scene(1.0).save(f"{OUT}/hero.png", optimize=True)
dim = draw_scene(1.0, machines=False)
dim.save(f"{OUT}/section.png", optimize=True)
print("ok")
