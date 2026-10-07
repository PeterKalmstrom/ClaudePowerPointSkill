"""Before/after example: the same slide built twice, once breaking the skill's rules and once following them.

    uvx --with python-pptx --with pillow python examples/before-after/build_example.py
    -> before.pptx, after.pptx (beside this script)

Rules shown (see SKILL.md -> Core rules):
  before                                      after
  topic title ("Q1 revenue")                  claim title ("East leads Q1, up 8 %")
  9 bullets, 40+ words, 14 pt body            hero stat + one 24 pt caption (~10 words)
  photo stretched into its box (circle ->     photo cover-cropped to the box ratio
    ellipse)
  details on the slide                        details, Q&A and source in the speaker notes
  default 4:3, 720 x 540 pt                   Full HD, 1440 x 810 pt
"""
import os
import sys

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "scripts"))
from cover_crop import cover_crop  # noqa: E402

INK, ACCENT, MUTED = RGBColor(0x1A, 0x1A, 0x1A), RGBColor(0x0B, 0x6E, 0x4F), RGBColor(0x5F, 0x63, 0x68)


def photo(path):
    """Stand-in 'photo' with a perfect circle in it, so stretching is visible."""
    img = Image.new("RGB", (1500, 1000), (214, 228, 222))
    d = ImageDraw.Draw(img)
    d.ellipse((550, 300, 950, 700), fill=(11, 110, 79))
    d.rectangle((0, 820, 1500, 1000), fill=(180, 200, 192))
    img.save(path)


def text(slide, s, left, top, width, height, size, color=INK, bold=False, name="Text"):
    tb = slide.shapes.add_textbox(Pt(left), Pt(top), Pt(width), Pt(height))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(s.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        for r in p.runs:
            r.font.size, r.font.bold, r.font.color.rgb = Pt(size), bold, color
    return tb


def before(img):
    prs = Presentation()  # python-pptx default: 720 x 540 pt, 4:3
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Q1 revenue"
    bullets = "\n".join([
        "East region revenue was up 8% compared to the previous quarter",
        "West region was flat quarter over quarter",
        "North declined slightly due to seasonal effects",
        "South grew 3% after the new partner programme",
        "Total revenue reached 41.2 million",
        "Margins improved by 1.5 points",
        "Source: internal finance dashboard, April 2026",
        "Methodology: constant currency, excluding one-offs",
        "Next steps: review East playbook for other regions",
    ])
    text(s, bullets, 30, 130, 360, 380, 14, name="Bullets")
    # the classic mistake: a 3:2 photo forced into a 290 x 120 pt box - both sizes given, aspect ratio lost
    s.shapes.add_picture(img, Pt(410), Pt(150), Pt(290), Pt(120)).name = "Photo"
    prs.save(os.path.join(HERE, "before.pptx"))


def after(img):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(1440), Pt(810)
    s = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    text(s, "East leads Q1, up 8 %", 80, 60, 1280, 110, 54, bold=True, name="Title")
    text(s, "+8 %", 80, 260, 620, 260, 160, color=ACCENT, bold=True, name="HeroStat")
    text(s, "East revenue vs Q4 — the other regions were flat", 80, 540, 620, 120, 24, color=MUTED, name="Caption")
    box_w, box_h = 600, 480
    cropped = os.path.join(HERE, "_photo.crop.png")
    cover_crop(Image.open(img), box_w, box_h).save(cropped)
    s.shapes.add_picture(cropped, Pt(760), Pt(220), Pt(box_w), Pt(box_h)).name = "Photo"
    os.remove(cropped)
    s.notes_slide.notes_text_frame.text = (
        "Key fact: East grew 8 % quarter on quarter; every other region was within ±3 %.\n"
        "Facts: South +3 % after the partner programme; West flat; North down slightly (seasonal). "
        "Total 41.2 M; margin +1.5 pt.\n"
        "Q&A: Is it currency? No — constant currency, one-offs excluded.\n"
        "Pitfalls: don't present North as a trend; it is seasonal.\n"
        "Sources: internal finance dashboard, April 2026.")
    prs.save(os.path.join(HERE, "after.pptx"))


if __name__ == "__main__":
    img = os.path.join(HERE, "_photo.png")
    photo(img)
    before(img)
    after(img)
    os.remove(img)
    print("before.pptx, after.pptx written to", HERE)
