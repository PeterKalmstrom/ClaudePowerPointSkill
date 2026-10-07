"""Before/after example: the same slide built twice, once breaking the skill's rules and once following them.

    uvx --with python-pptx --with pillow python examples/before-after/build_example.py
    -> before.pptx, after.pptx (beside this script)

Rules shown (see SKILL.md -> Core rules):
  before                                        after
  topic title ("Q1 revenue")                    claim title ("East leads Q1, up 8 %")
  9 bullets, 60+ words, 14 pt body              hero stat + one 24 pt caption
  chart pasted as a picture and stretched       native, editable chart; the finding is the
    (labels squashed, default Office colours)     only accent colour
  details on the slide                          details, Q&A and source in the speaker notes
  default 4:3, 720 x 540 pt                     Full HD, 1440 x 810 pt
"""
import os

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION
from pptx.util import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
REGIONS = ["East", "South", "West", "North"]
GROWTH = [8.0, 3.0, 0.2, -1.5]  # % change vs Q4
INK, ACCENT, MUTED, QUIET = (RGBColor(0x1A, 0x1A, 0x1A), RGBColor(0x0B, 0x6E, 0x4F),
                             RGBColor(0x5F, 0x63, 0x68), RGBColor(0xC9, 0xCF, 0xCC))


def font(size):
    for f in ("DejaVuSans.ttf", "arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(f, size)
        except OSError:
            pass
    return ImageFont.load_default()


def chart_png(path):
    """A chart exported as a picture, the way it often arrives: Office-default blue, every bar labelled."""
    w, h = 900, 600
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    d.text((30, 20), "Revenue growth by region, Q1 vs Q4 (%)", fill="black", font=font(28))
    base, scale, bw = 420, 30, 140
    d.line((40, base, w - 40, base), fill=(160, 160, 160), width=2)
    for i, (r, g) in enumerate(zip(REGIONS, GROWTH)):
        x = 80 + i * 200
        top = base - g * scale
        d.rectangle((x, min(base, top), x + bw, max(base, top)), fill=(0x44, 0x72, 0xC4))
        d.text((x + 30, min(base, top) - 36), f"{g:+.1f}", fill="black", font=font(26))
        d.text((x + 30, base + 70), r, fill="black", font=font(26))
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


def before(png):
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
    text(s, bullets, 30, 130, 330, 380, 14, name="Bullets")
    # the classic mistake: a 3:2 chart picture forced into a wide, short box - labels squashed
    s.shapes.add_picture(png, Pt(380), Pt(170), Pt(320), Pt(130)).name = "ChartPicture"
    prs.save(os.path.join(HERE, "before.pptx"))


def after():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(1440), Pt(810)
    s = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only: a real title placeholder for outline and screen readers
    title = s.shapes.title
    title.left, title.top, title.width, title.height = Pt(80), Pt(60), Pt(1280), Pt(110)
    title.text = "East leads Q1, up 8 %"
    tf = title.text_frame
    tf.paragraphs[0].alignment = 1  # left
    for r in tf.paragraphs[0].runs:
        r.font.size, r.font.bold, r.font.color.rgb = Pt(54), True, INK
    text(s, "+8 %", 80, 250, 560, 260, 160, color=ACCENT, bold=True, name="HeroStat")
    text(s, "East revenue vs Q4. Every other region within ±3 %.", 80, 540, 560, 120, 24,
         color=MUTED, name="Caption")

    data = CategoryChartData()
    data.categories = REGIONS
    data.add_series("Growth vs Q4 (%)", GROWTH)
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Pt(720), Pt(220), Pt(640), Pt(480), data)
    gf.name = "GrowthChart"
    gf._element.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr").set(
        "descr", "Column chart of Q1 revenue growth vs Q4: East +8.0 %, South +3.0 %, West +0.2 %, North -1.5 %.")
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = False
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW  # below negative bars
    ch.category_axis.tick_labels.font.size = Pt(20)
    ch.category_axis.tick_labels.font.color.rgb = MUTED
    plot = ch.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.number_format, labels.number_format_is_linked = '+0.0;-0.0', False
    labels.position = XL_LABEL_POSITION.OUTSIDE_END
    labels.font.size, labels.font.bold = Pt(20), True
    for i, point in enumerate(plot.series[0].points):  # only the finding gets the accent colour
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = ACCENT if i == 0 else QUIET
    s.notes_slide.notes_text_frame.text = (
        "Key fact: East grew 8 % quarter on quarter; every other region was within ±3 %.\n"
        "Facts: South +3 % after the partner programme; West flat; North down 1.5 % (seasonal). "
        "Total 41.2 M; margin +1.5 pt.\n"
        "Q&A: Is it currency? No — constant currency, one-offs excluded.\n"
        "Pitfalls: don't present North as a trend; it is seasonal.\n"
        "Sources: internal finance dashboard, April 2026.")
    prs.save(os.path.join(HERE, "after.pptx"))


if __name__ == "__main__":
    png = os.path.join(HERE, "_chart.png")
    chart_png(png)
    before(png)
    after()
    os.remove(png)
    print("before.pptx, after.pptx written to", HERE)
