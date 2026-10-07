"""Compose hero.png (1280 x 640, also the GitHub social preview) from before.png and after.png."""
import os

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))


def font(size, bold=False):
    names = (["DejaVuSans-Bold.ttf", "arialbd.ttf"] if bold else ["DejaVuSans.ttf", "arial.ttf"])
    for n in names + ["/usr/share/fonts/truetype/dejavu/" + names[0]]:
        try:
            return ImageFont.truetype(n, size)
        except OSError:
            pass
    return ImageFont.load_default()


W, H, PAD, TOP = 1280, 640, 32, 120
hero = Image.new("RGB", (W, H), (244, 246, 245))
d = ImageDraw.Draw(hero)
d.text((PAD, 24), "Same content, same data.", fill=(26, 26, 26), font=font(36, bold=True))
d.text((PAD, 70), "Left: a typical first draft. Right: the skill's rules applied.",
       fill=(95, 99, 104), font=font(24))
col_w = (W - 3 * PAD) // 2
frame_h = round(col_w * 9 / 16)  # both panels in the same 16:9 frame
for i, (name, label, color) in enumerate([("before.png", "Without the skill's rules", (176, 58, 46)),
                                          ("after.png", "With them", (11, 110, 79))]):
    x = PAD + i * (col_w + PAD)
    d.rectangle((x, TOP, x + col_w, TOP + frame_h), fill="white", outline=(200, 205, 203))
    img = Image.open(os.path.join(HERE, name)).convert("RGB")
    scale = min((col_w - 2) / img.width, (frame_h - 2) / img.height)
    img = img.resize((round(img.width * scale), round(img.height * scale)))
    hero.paste(img, (x + (col_w - img.width) // 2, TOP + (frame_h - img.height) // 2))
    d.text((x, TOP + frame_h + 14), label, fill=color, font=font(28, bold=True))
d.text((PAD, H - 64), "Claude Code skill  ·  github.com/PeterKalmstrom/claude-powerpoint-skill",
       fill=(95, 99, 104), font=font(24))
hero.save(os.path.join(HERE, "hero.png"))
print("hero.png", hero.size)
