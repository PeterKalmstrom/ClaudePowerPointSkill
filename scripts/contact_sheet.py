"""Every slide as a thumbnail grid in one PNG - shows colour drift, density imbalance,
monotone-anchor stretches and forgotten template slides that per-slide review misses.

    uvx --with pywin32 --with pillow python contact_sheet.py --file deck.pptx [--cols 6] [--slide-width 240]
    -> <deck>_contact.png beside the deck
"""
import argparse
import os
import tempfile

from PIL import Image, ImageDraw

from render_slides import render

ap = argparse.ArgumentParser()
ap.add_argument("--file", required=True)
ap.add_argument("--cols", type=int, default=6)
ap.add_argument("--slide-width", type=int, default=240)
ap.add_argument("--out")
a = ap.parse_args()

files = render(a.file, tempfile.mkdtemp(prefix="contact-"))
first = Image.open(files[0])
w = a.slide_width
h = round(w * first.height / first.width)  # aspect ratio from the deck itself
cap, pad = 18, 8
rows = -(-len(files) // a.cols)
sheet = Image.new("RGB", (a.cols * (w + pad) + pad, rows * (h + cap + pad) + pad), "white")
draw = ImageDraw.Draw(sheet)
for i, f in enumerate(files):
    x = pad + (i % a.cols) * (w + pad)
    y = pad + (i // a.cols) * (h + cap + pad)
    sheet.paste(Image.open(f).resize((w, h)), (x, y))
    draw.text((x, y + h + 2), str(i + 1), fill="black")
out = a.out or os.path.splitext(a.file)[0] + "_contact.png"
sheet.save(out)
print(out)
