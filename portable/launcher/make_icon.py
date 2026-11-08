"""Draw buybox.ico (a gold bar on a navy tile) for the Windows launcher.

    python portable/launcher/make_icon.py
"""

from pathlib import Path

from PIL import Image, ImageDraw

HERE = Path(__file__).resolve().parent
NAVY = (31, 78, 120, 255)


def draw(size: int = 256) -> Image.Image:
    s = size
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, s - 1, s - 1], radius=int(s * 0.2), fill=NAVY)
    # gold bar: front face (trapezoid) + top face
    bl, br = (s * 0.14, s * 0.74), (s * 0.86, s * 0.74)
    tl, tr = (s * 0.27, s * 0.47), (s * 0.73, s * 0.47)
    d.polygon([bl, br, tr, tl], fill=(212, 160, 23, 255))
    d.polygon([tl, tr, (s * 0.64, s * 0.33), (s * 0.36, s * 0.33)], fill=(247, 210, 90, 255))
    d.line([bl, br], fill=(150, 105, 10, 255), width=max(1, s // 64))
    # shine
    d.polygon([(s * 0.33, s * 0.69), (s * 0.43, s * 0.52), (s * 0.48, s * 0.52), (s * 0.38, s * 0.69)],
              fill=(255, 236, 160, 200))
    # up arrow (buy box won)
    d.polygon([(s * 0.5, s * 0.08), (s * 0.62, s * 0.24), (s * 0.38, s * 0.24)], fill=(80, 200, 120, 255))
    return img


if __name__ == "__main__":
    out = HERE / "buybox.ico"
    draw(256).save(out, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("wrote", out)
