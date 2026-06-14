"""
Generiert nesk3.ico aus dem Splash-Screen-Design:
  - Dunkler radialer Hintergrund
  - Teal-Arc (grosser Ring)
  - Gold-Arc (kleiner, gegenlaeufig)
  - Teal-Glow in der Mitte
  - "N" in der Mitte (Segoe UI / Fallback)
  - Goldstreifen oben

Speichert Daten/Logo/nesk3.ico mit den Groessen 256, 128, 64, 48, 32, 16.
"""
import math, os
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

BASE = Path(__file__).parent

# Farben (RGBA)
BG_OUTER  = (10,  21,  32, 255)
BG_INNER  = (24,  38,  52, 255)
TEAL      = (91, 138, 170, 255)
GOLD      = (192, 148,  74, 255)
WHITE_DIM = (180, 205, 220, 200)


def _arc_points(cx, cy, r, start_deg, span_deg, steps=120):
    pts = []
    for i in range(steps + 1):
        a = math.radians(start_deg + span_deg * i / steps)
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return pts


def draw_icon(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img, "RGBA")

    # ── Hintergrund (radial via Gradient-Simulation) ─────────────────────
    cx, cy = size / 2, size / 2
    # Zeichne konzentrische Kreise von aussen nach innen
    steps = 40
    for i in range(steps, -1, -1):
        t = i / steps
        r_step = cx * t
        r_val = int(BG_OUTER[0] + (BG_INNER[0] - BG_OUTER[0]) * (1 - t))
        g_val = int(BG_OUTER[1] + (BG_INNER[1] - BG_OUTER[1]) * (1 - t))
        b_val = int(BG_OUTER[2] + (BG_INNER[2] - BG_OUTER[2]) * (1 - t))
        draw.ellipse(
            [cx - r_step, cy - r_step, cx + r_step, cy + r_step],
            fill=(r_val, g_val, b_val, 255)
        )

    # ── Goldstreifen oben ─────────────────────────────────────────────────
    stripe_h = max(2, size // 85)
    draw.rectangle([0, 0, size, stripe_h], fill=GOLD)

    # ── Radialer Glow in der Mitte ────────────────────────────────────────
    glow_r = int(size * 0.28)
    for gi in range(glow_r, 0, -1):
        alpha = int(55 * (1 - gi / glow_r) ** 1.5)
        draw.ellipse(
            [cx - gi, cy - gi, cx + gi, cy + gi],
            fill=(91, 138, 170, alpha)
        )

    # ── Innerer dunkler Kreis ─────────────────────────────────────────────
    inner_r = int(size * 0.20)
    draw.ellipse(
        [cx - inner_r, cy - inner_r, cx + inner_r, cy + inner_r],
        fill=(17, 30, 42, 255)
    )

    # ── Teal-Ring (grosser, 230° Bogen) ──────────────────────────────────
    r1 = int(size * 0.265)
    lw1 = max(2, size // 40)
    # start bei -120°, span 230°
    draw.arc(
        [cx - r1, cy - r1, cx + r1, cy + r1],
        start=-120, end=110,
        fill=TEAL, width=lw1
    )

    # ── Gold-Ring (kleiner, 110° Bogen, versetzt) ─────────────────────────
    r2 = int(size * 0.315)
    lw2 = max(1, size // 65)
    draw.arc(
        [cx - r2, cy - r2, cx + r2, cy + r2],
        start=130, end=240,
        fill=GOLD, width=lw2
    )

    # ── "N" in der Mitte ──────────────────────────────────────────────────
    font_size = max(8, int(size * 0.28))
    font = None
    for fname in ["segoeuil.ttf", "segoeui.ttf", "arial.ttf"]:
        try:
            font = ImageFont.truetype(fname, font_size)
            break
        except Exception:
            pass
    if font is None:
        font = ImageFont.load_default()

    bb = draw.textbbox((0, 0), "N", font=font)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    nx = cx - tw / 2 - bb[0]
    ny = cy - th / 2 - bb[1]
    draw.text((nx, ny), "N", font=font, fill=TEAL)

    return img


# Erzeuge alle Größen und speichere als ICO
sizes   = [256, 128, 64, 48, 32, 16]
images  = [draw_icon(s) for s in sizes]

out_path = BASE / "Daten" / "Logo" / "nesk3.ico"
out_path.parent.mkdir(parents=True, exist_ok=True)

# Pillow speichert ICO mit mehreren Größen über append_images
images[0].save(
    str(out_path),
    format="ICO",
    sizes=[(s, s) for s in sizes],
    append_images=images[1:],
)
print(f"Icon gespeichert: {out_path}")
