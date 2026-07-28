#!/usr/bin/env python3
"""Prepare processed images: gradients, title composite, rounded photos, cropped bags."""
from PIL import Image, ImageDraw, ImageFilter, ImageOps, ImageEnhance
import os, math

A = os.path.join(os.path.dirname(__file__), "assets")
os.makedirs(A, exist_ok=True)

# ---- palette ----
DEEP   = (16, 43, 32)     # 102B20 deep forest
DEEP2  = (10, 28, 21)     # 0A1C15 darker
GREEN  = (45, 156, 90)    # 2D9C5A
INK    = (21, 33, 27)

def lerp(a, b, t):
    return tuple(int(a[i] + (b[i]-a[i])*t) for i in range(3))

# ---- 1. Diagonal deep-green gradient (title / closing / section) ----
def diagonal_gradient(w, h, c1, c2, fname, angle_deg=30):
    img = Image.new("RGB", (w, h))
    px = img.load()
    ang = math.radians(angle_deg)
    dx, dy = math.cos(ang), math.sin(ang)
    # project each pixel onto direction, normalize
    maxp = w*abs(dx) + h*abs(dy)
    for y in range(h):
        for x in range(w):
            t = (x*dx + y*dy) / maxp
            px[x, y] = lerp(c1, c2, t)
    img.save(os.path.join(A, fname), quality=95)

# smaller compute: build at low res and upscale (smooth gradient anyway)
def diagonal_gradient_fast(w, h, c1, c2, fname, angle_deg=30):
    sw, sh = 320, 180
    small = Image.new("RGB", (sw, sh))
    px = small.load()
    ang = math.radians(angle_deg)
    dx, dy = math.cos(ang), math.sin(ang)
    maxp = sw*abs(dx) + sh*abs(dy)
    for y in range(sh):
        for x in range(sw):
            t = (x*dx + y*dy) / maxp
            px[x, y] = lerp(c1, c2, t)
    small.resize((w, h), Image.BILINEAR).save(os.path.join(A, fname), quality=95)

diagonal_gradient_fast(1920, 1080, (13,38,28), (7,20,15), "bg_dark.jpg", 35)
diagonal_gradient_fast(1920, 1080, (18,52,38), (9,26,19), "bg_section.jpg", 35)

# subtle light background with faint green vignette
def light_bg(w, h, fname):
    img = Image.new("RGB", (w, h), (247, 250, 246))
    img.save(os.path.join(A, fname), quality=95)
light_bg(1920, 1080, "bg_light.jpg")

# ---- 2. Rounded corner helper ----
def rounded(im, radius_frac=0.045):
    im = im.convert("RGBA")
    w, h = im.size
    r = int(min(w, h) * radius_frac)
    mask = Image.new("L", (w, h), 0)
    d = ImageDraw.Draw(mask)
    d.rounded_rectangle([0, 0, w-1, h-1], radius=r, fill=255)
    out = Image.new("RGBA", (w, h), (0,0,0,0))
    out.paste(im, (0,0), mask)
    return out

def cover_crop(im, tw, th):
    """crop-to-cover a target aspect ratio."""
    w, h = im.size
    ta = tw/th; a = w/h
    if a > ta:
        nw = int(h*ta); x = (w-nw)//2
        im = im.crop((x, 0, x+nw, h))
    else:
        nh = int(w/ta); y = (h-nh)//2
        im = im.crop((0, y, w, y+nh))
    return im

def save_rounded_cover(src, tw, th, fname, radius=0.05):
    im = Image.open(os.path.join(A, src)).convert("RGB")
    im = cover_crop(im, tw, th)
    im = im.resize((tw*2, th*2), Image.LANCZOS)
    rounded(im, radius).save(os.path.join(A, fname))

# process/diagram, plant render, maps, corn, yieldmap -> rounded cover crops
save_rounded_cover("plant3d.png",     900, 640, "r_plant3d.png")
save_rounded_cover("process.jpg",     820, 620, "r_process.png")
save_rounded_cover("corn.png",        760, 620, "r_corn.png")
save_rounded_cover("map_pukhalsk.png",560, 380, "r_map_pukhalsk.png")
save_rounded_cover("map_arykbalyk.png",560,380, "r_map_arykbalyk.png")
save_rounded_cover("yieldmap.jpeg",   700, 560, "r_yieldmap.png")

# ---- 3. Title hero composite: hero image cover + diagonal dark scrim from left ----
def title_bg(fname):
    W, H = 1920, 1080
    hero = Image.open(os.path.join(A, "hero.png")).convert("RGB")
    hero = cover_crop(hero, W, H).resize((W, H), Image.LANCZOS)
    hero = ImageEnhance.Brightness(hero).enhance(0.92)
    base = hero.convert("RGBA")
    # scrim: strong dark-green on the left fading to transparent on right, plus bottom darken
    scrim = Image.new("RGBA", (W, H), (0,0,0,0))
    sp = scrim.load()
    for x in range(W):
        # left->right fade
        t = x / W
        alpha_left = max(0.0, 0.92 - t*1.15)  # opaque left, clear ~x>0.8
        for y in range(H):
            # bottom gradient add
            tb = y / H
            alpha_bottom = max(0.0, (tb-0.55)*0.9)
            a = min(0.96, alpha_left + alpha_bottom*0.6)
            sp[x, y] = (9, 26, 19, int(a*255))
    out = Image.alpha_composite(base, scrim)
    out.convert("RGB").save(os.path.join(A, fname), quality=92)

title_bg("bg_title.jpg")

# ---- 4. Closing bg: plant render darkened, right side, with dark scrim ----
def closing_bg(fname):
    W, H = 1920, 1080
    base = Image.open(os.path.join(A, "bg_dark.jpg")).convert("RGBA")
    plant = Image.open(os.path.join(A, "plant3d.png")).convert("RGB")
    plant = cover_crop(plant, 1000, 1080).resize((1000, 1080), Image.LANCZOS)
    plant = ImageEnhance.Brightness(plant).enhance(0.7).convert("RGBA")
    # fade plant's left edge into the dark bg
    pmask = Image.new("L", (1000, 1080), 0)
    pd = pmask.load()
    for x in range(1000):
        for y in range(1080):
            t = x/1000
            pd[x,y] = int(min(1.0, max(0.0, (t-0.15)/0.35))*255)
    base.paste(plant, (920, 0), pmask)
    base.convert("RGB").save(os.path.join(A, fname), quality=92)

closing_bg("bg_closing.jpg")

# ---- 5. Bags: trim white, keep as-is on RGBA transparent for placing on cards ----
def trim_to_png(src, fname):
    im = Image.open(os.path.join(A, src)).convert("RGB")
    im.save(os.path.join(A, fname), quality=95)
trim_to_png("bag_npk.png", "p_bag_npk.png")
trim_to_png("bag_np.png",  "p_bag_np.png")

print("images done")
