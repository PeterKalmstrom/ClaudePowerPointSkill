"""Build a deck from a JSON (or YAML) spec - any OS, python-pptx only.

    uvx --with python-pptx --with pillow python scripts/build_deck.py spec.json --out deck.pptx
    uvx --with python-pptx --with pillow python scripts/build_deck.py spec.json --out deck.pptx --lint
    python scripts/build_deck.py --list-directions

The spec names a design direction (scripts/directions.json) or a template (.pptx/.potx), then lists
slides by pattern. The builder applies the skill's rules by construction: 1440 x 810 pt, a real title
placeholder on every slide, theme colours and fonts (so the deck re-themes), named shapes, stable slide
ids, cover-cropped pictures, native charts with one highlighted finding, alt text and speaker notes.
Inputs are checked against each pattern's limits first; nothing is written if a limit is broken
(--force builds anyway). The spec format is in reference/BUILDER.md.

Patterns: title, section, statement, big_number, kpi, bullets, compare, process, timeline, quote,
chart, table, image, matrix.
"""
import argparse
import copy
import json
import os
import subprocess
import sys
import tempfile

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}

# Canvas and type scale for a 1440 x 810 pt slide (1920 x 1080 px). Body never below 18 pt.
W, H, M = 1440, 810, 80
TITLE_TOP, TITLE_H = 56, 100
BODY_TOP, BODY_BOTTOM = 196, 730
SIZE = {"title": 54, "section": 84, "statement": 72, "hero": 220, "kpi": 84, "h2": 46, "body": 34,
        "small": 30, "caption": 28, "label": 24}  # 1.5x a 960-pt slide: body 34 ~ 23 pt, floor 27 ~ 18 pt
GAP = 24

TEXT, MUTED, ACCENT, BG, QUIET = (MSO_THEME_COLOR.TEXT_1, MSO_THEME_COLOR.TEXT_2, MSO_THEME_COLOR.ACCENT_1,
                                  MSO_THEME_COLOR.BACKGROUND_1, MSO_THEME_COLOR.BACKGROUND_2)

LIMITS = {  # pattern: {field: max characters} plus list-length ranges, checked before building
    "title": {"title": 70, "subtitle": 120},
    "section": {"title": 60, "eyebrow": 30},
    "statement": {"title": 90, "support": 140},
    "big_number": {"title": 70, "number": 12, "unit": 10, "caption": 90},
    "kpi": {"title": 70, "metrics": (3, 6), "metrics.value": 12, "metrics.label": 28},
    "bullets": {"title": 70, "items": (1, 7), "items.*": 100},
    "compare": {"title": 70, "columns": (2, 3), "columns.heading": 30, "columns.points": (1, 4),
                "columns.points.*": 70},
    "process": {"title": 70, "steps": (3, 8), "steps.label": 30, "steps.detail": 60},
    "timeline": {"title": 70, "events": (3, 7), "events.date": 16, "events.label": 40},
    "quote": {"quote": 240, "attribution": 40, "role": 60},
    "chart": {"title": 70, "categories": (2, 24), "series": (1, 6)},
    "table": {"title": 70, "header": (2, 6), "rows": (1, 8)},
    "image": {"title": 70, "caption": 120},
    "matrix": {"title": 70, "quadrants": (4, 4), "quadrants.heading": 30, "quadrants.text": 90,
               "x_axis": 30, "y_axis": 30},
}


# ---------------------------------------------------------------- spec checks

def check_spec(spec):
    errors = []
    for i, sl in enumerate(spec.get("slides", []), 1):
        pat = sl.get("pattern")
        if pat not in LIMITS:
            errors.append(f"slide {i}: unknown pattern '{pat}' (one of: {', '.join(LIMITS)})")
            continue
        if pat not in ("title", "quote", "section") and not sl.get("title"):
            errors.append(f"slide {i} ({pat}): needs a 'title' - write it as a claim")
        for key, lim in LIMITS[pat].items():
            parts = key.split(".")
            for path, val in _values(sl, parts):
                if isinstance(lim, tuple):
                    n = len(val) if isinstance(val, list) else 0
                    if not lim[0] <= n <= lim[1]:
                        errors.append(f"slide {i} ({pat}): '{path}' has {n} items, needs {lim[0]}-{lim[1]}")
                elif isinstance(val, str) and len(val) > lim:
                    errors.append(f"slide {i} ({pat}): '{path}' is {len(val)} characters, max {lim}: '{val[:40]}…'")
        if pat in ("kpi", "process", "timeline", "compare", "matrix", "chart", "table", "bullets"):
            need = {"kpi": "metrics", "process": "steps", "timeline": "events", "compare": "columns",
                    "matrix": "quadrants", "chart": "series", "table": "rows", "bullets": "items"}[pat]
            if need not in sl:
                errors.append(f"slide {i} ({pat}): missing '{need}'")
        if pat == "image" and not os.path.exists(sl.get("image", "")):
            errors.append(f"slide {i} (image): image file not found: {sl.get('image')}")
        if pat == "chart":
            cats = len(sl.get("categories", []))
            for s in sl.get("series", []):
                if len(s.get("values", [])) != cats:
                    errors.append(f"slide {i} (chart): series '{s.get('name')}' has {len(s.get('values', []))} "
                                  f"values for {cats} categories")
    return errors


def _values(obj, parts, path=""):
    head, rest = parts[0], parts[1:]
    if head == "*":
        items = list(enumerate(obj)) if isinstance(obj, list) else []
    else:
        items = [(head, obj.get(head))] if isinstance(obj, dict) and head in obj else []
    for k, v in items:
        p = f"{path}.{k}" if path else str(k)
        if not rest:
            yield p, v
        elif isinstance(v, list) and rest[0] != "*":
            for j, el in enumerate(v):
                yield from _values(el, rest, f"{p}[{j}]")
        else:
            yield from _values(v, rest, p)


# ---------------------------------------------------------------- theme

def load_direction(name):
    ds = json.load(open(os.path.join(HERE, "directions.json"), encoding="utf-8"))["directions"]
    for d in ds:
        if d["id"] == name:
            return d
    sys.exit(f"unknown direction '{name}'; see --list-directions")


def mix(a, b, t):
    ca, cb = [int(a[i:i + 2], 16) for i in (0, 2, 4)], [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ca, cb))


def apply_direction(prs, d):
    """Write the direction into the theme so everything below uses theme colours and fonts."""
    part = prs.slide_master.part.part_related_by(RT.THEME)
    theme = etree.fromstring(part.blob)
    scheme = theme.find(".//a:clrScheme", NS)
    slots = {"dk1": d["text"], "lt1": d["background"], "dk2": d["muted"],
             "lt2": mix(d["background"], d["text"], 0.18), "accent1": d["accent"],
             "accent2": mix(d["accent"], d["background"], 0.45), "accent3": d["muted"],
             "accent4": mix(d["accent"], d["text"], 0.4), "accent5": mix(d["muted"], d["background"], 0.4),
             "accent6": mix(d["accent"], d["background"], 0.7)}
    for slot, hexv in slots.items():
        node = scheme.find(f"a:{slot}", NS)
        for child in list(node):
            node.remove(child)
        etree.SubElement(node, f"{{{A}}}srgbClr", val=hexv)
    scheme.set("name", d["id"])
    fonts = theme.find(".//a:fontScheme", NS)
    fonts.find("a:majorFont/a:latin", NS).set("typeface", d["heading"])
    fonts.find("a:minorFont/a:latin", NS).set("typeface", d["body"])
    fonts.set("name", d["id"])
    part._blob = etree.tostring(theme, xml_declaration=True, encoding="UTF-8", standalone=True)


def open_template(path):
    if path.lower().endswith((".potx", ".potm")):
        sys.path.insert(0, HERE)
        from extract_theme import open_any
        return open_any(path)
    return Presentation(path)


def title_only_layout(prs):
    for layout in prs.slide_layouts:
        types = [ph.placeholder_format.type for ph in layout.placeholders]
        if layout.name.lower().replace("-", " ").startswith("title only"):
            return layout
    for layout in prs.slide_layouts:  # fallback: a layout whose only text placeholder is the title
        kinds = {str(ph.placeholder_format.type) for ph in layout.placeholders}
        if any("TITLE" in k for k in kinds) and not any("BODY" in k or "OBJECT" in k for k in kinds):
            return layout
    return prs.slide_layouts[0]


# ---------------------------------------------------------------- drawing helpers

class Builder:
    def __init__(self, prs, themed):
        self.prs = prs
        self.themed = themed  # True when a direction was applied (we own fonts/colours)
        self.layout = title_only_layout(prs)

    def new_slide(self, spec, n):
        s = self.prs.slides.add_slide(self.layout)
        s._element.find("p:cSld", NS).set("name", spec.get("id", f"s{n:02d}"))
        if self.themed:
            s.background.fill.solid()
            s.background.fill.fore_color.theme_color = BG
        for ph in list(s.placeholders):
            if "TITLE" not in str(ph.placeholder_format.type):
                ph._element.getparent().remove(ph._element)
        return s

    def title(self, s, text, size=None, top=TITLE_TOP, height=TITLE_H, align=PP_ALIGN.LEFT, colour=TEXT):
        t = s.shapes.title
        t.left, t.top, t.width, t.height = Pt(M), Pt(top), Pt(W - 2 * M), Pt(height)
        t.text = text
        tf = t.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
        for p in tf.paragraphs:
            p.alignment = align
            for r in p.runs:
                r.font.size = Pt(size or SIZE["title"])
                r.font.bold = True
                if self.themed:
                    r.font.color.theme_color = colour
                    r.font.name = "+mj-lt"
        return t

    def text(self, s, name, txt, x, y, w, h, size, colour=TEXT, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, heading=False):
        tb = s.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
        tb.name = name
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        for side in ("margin_left", "margin_right"):
            setattr(tf, side, Pt(0))
        lines = txt if isinstance(txt, list) else [txt]
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.alignment = align
            if i:
                p.space_before = Pt(size * 0.5)
            for r in p.runs:
                r.font.size, r.font.bold = Pt(size), bold
                r.font.color.theme_color = colour
                if self.themed:
                    r.font.name = "+mj-lt" if heading else "+mn-lt"
        return tb

    def rect(self, s, name, x, y, w, h, colour=QUIET, shape=MSO_SHAPE.RECTANGLE):
        r = s.shapes.add_shape(shape, Pt(x), Pt(y), Pt(w), Pt(h))
        r.name = name
        r.fill.solid()
        r.fill.fore_color.theme_color = colour
        r.line.fill.background()
        r.shadow.inherit = False
        return r


def alt(shape, text):
    shape._element.find(".//p:cNvPr", NS).set("descr", text)


def notes_text(notes):
    if not notes:
        return ""
    if isinstance(notes, str):
        return notes
    out = []
    if notes.get("key_fact"):
        out.append(f"KEY FACT: {notes['key_fact']}")
    for label, key in (("FACTS", "facts"), ("PITFALLS", "pitfalls"), ("SOURCES", "sources")):
        if notes.get(key):
            out.append(f"{label}:\n" + "\n".join(f"- {x}" for x in notes[key]))
    if notes.get("qa"):
        out.append("Q&A:\n" + "\n".join(f"Q: {q['q']}\nA: {q['a']}" for q in notes["qa"]))
    return "\n\n".join(out)


# ---------------------------------------------------------------- patterns

def p_title(b, s, sl):
    b.title(s, sl.get("title", ""), size=SIZE["section"], top=260, height=200)
    if sl.get("subtitle"):
        b.text(s, "Subtitle", sl["subtitle"], M, 480, W - 2 * M, 120, SIZE["h2"], MUTED)


def p_section(b, s, sl):
    if sl.get("eyebrow"):
        b.text(s, "Eyebrow", sl["eyebrow"].upper(), M, 250, W - 2 * M, 50, SIZE["label"], ACCENT, bold=True)
    b.title(s, sl["title"], size=SIZE["section"], top=300, height=220)
    b.rect(s, "AccentRule", M, 540, 160, 8, ACCENT)


def p_statement(b, s, sl):
    b.title(s, sl["title"], size=SIZE["statement"], top=200, height=300)
    if sl.get("support"):
        b.text(s, "Support", sl["support"], M, 540, W - 2 * M, 120, SIZE["h2"], MUTED)


def p_big_number(b, s, sl):
    b.title(s, sl["title"])
    num = sl["number"] + (f" {sl['unit']}" if sl.get("unit") else "")
    b.text(s, "HeroNumber", num, M, 240, W - 2 * M, 280, SIZE["hero"], ACCENT, bold=True, heading=True)
    if sl.get("caption"):
        b.text(s, "Caption", sl["caption"], M, 560, W - 2 * M, 100, SIZE["h2"], MUTED)


def p_kpi(b, s, sl):
    b.title(s, sl["title"])
    ms = sl["metrics"]
    hi = sl.get("highlight", 0)
    w = (W - 2 * M - GAP * (len(ms) - 1)) / len(ms)
    for i, m in enumerate(ms):
        x = M + i * (w + GAP)
        b.rect(s, f"Card{i + 1}", x, 300, w, 300, QUIET if i != hi else ACCENT)
        fg = BG if i == hi else TEXT
        b.text(s, f"Value{i + 1}", m["value"], x + 24, 340, w - 48, 120, SIZE["kpi"], fg, bold=True, heading=True)
        b.text(s, f"Label{i + 1}", m["label"], x + 24, 480, w - 48, 90, SIZE["small"], fg if i == hi else MUTED)


def p_bullets(b, s, sl):
    b.title(s, sl["title"])
    b.text(s, "Points", ["• " + x for x in sl["items"]], M, BODY_TOP + 20, W - 2 * M - 300, BODY_BOTTOM - BODY_TOP,
           SIZE["body"])


def p_compare(b, s, sl):
    b.title(s, sl["title"])
    cols = sl["columns"]
    hi = sl.get("highlight")
    w = (W - 2 * M - GAP * 2 * (len(cols) - 1)) / len(cols)
    for i, c in enumerate(cols):
        x = M + i * (w + 2 * GAP)
        b.rect(s, f"Rule{i + 1}", x, 230, w, 8, ACCENT if i == hi else MUTED)
        b.text(s, f"Heading{i + 1}", c["heading"], x, 260, w, 60, SIZE["h2"], TEXT, bold=True, heading=True)
        b.text(s, f"Points{i + 1}", c["points"], x, 340, w, 360, SIZE["caption"], TEXT)


def p_process(b, s, sl):
    b.title(s, sl["title"])
    steps = sl["steps"]
    n = len(steps)
    w = (W - 2 * M - GAP * (n - 1)) / n
    hi = sl.get("highlight")
    for i, st in enumerate(steps):
        x = M + i * (w + GAP)
        b.rect(s, f"Step{i + 1}", x, 330, w, 120, ACCENT if i == hi else QUIET, MSO_SHAPE.CHEVRON if n <= 5 else MSO_SHAPE.RECTANGLE)
        b.text(s, f"StepNo{i + 1}", str(i + 1), x + (36 if n <= 5 else 16), 350, w - 60, 80, SIZE["h2"],
               BG if i == hi else TEXT, bold=True, anchor=MSO_ANCHOR.MIDDLE, heading=True)
        b.text(s, f"StepLabel{i + 1}", st["label"], x, 470, w, 70, SIZE["caption"] if n <= 5 else SIZE["label"], TEXT, bold=True)
        if st.get("detail"):
            b.text(s, f"StepDetail{i + 1}", st["detail"], x, 545, w, 140, SIZE["label"], MUTED)


def p_timeline(b, s, sl):
    b.title(s, sl["title"])
    ev = sl["events"]
    n = len(ev)
    y = 420
    b.rect(s, "Rail", M, y - 3, W - 2 * M, 6, QUIET)
    step = (W - 2 * M) / n
    hi = sl.get("highlight")
    for i, e in enumerate(ev):
        cx = M + step * i + step / 2
        b.rect(s, f"Dot{i + 1}", cx - 14, y - 14, 28, 28, ACCENT if i == hi or hi is None else MUTED, MSO_SHAPE.OVAL)
        top = i % 2 == 0
        b.text(s, f"Date{i + 1}", e["date"], cx - step / 2 + 8, y - 150 if top else y + 40, step - 16, 50,
               SIZE["h2"] - 2, ACCENT, bold=True, align=PP_ALIGN.CENTER, heading=True)
        b.text(s, f"Event{i + 1}", e["label"], cx - step / 2 + 8, y - 100 if top else y + 90, step - 16, 80,
               SIZE["label"], TEXT, align=PP_ALIGN.CENTER)


def p_quote(b, s, sl):
    # the title stays a real placeholder (outline, screen readers) but reads as a small label
    t = b.title(s, sl.get("title") or "In their words", size=SIZE["label"], top=56, height=50, colour=MUTED)
    t.text_frame.vertical_anchor = MSO_ANCHOR.TOP
    b.text(s, "QuoteMark", "\u201C", M - 10, 120, 120, 170, 180, ACCENT, bold=True, heading=True)
    b.text(s, "Quote", sl["quote"], M + 120, 220, W - 2 * M - 240, 360, SIZE["statement"] - 16, TEXT, heading=True)
    who = sl.get("attribution", "") + (f", {sl['role']}" if sl.get("role") else "")
    if who:
        b.text(s, "Attribution", "\u2014 " + who, M + 120, 620, W - 2 * M - 240, 60, SIZE["small"], MUTED)


CHART_TYPES = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
               "line": XL_CHART_TYPE.LINE_MARKERS, "pie": XL_CHART_TYPE.PIE}


def p_chart(b, s, sl):
    b.title(s, sl["title"])
    kind = sl.get("type", "column")
    data = CategoryChartData()
    data.categories = sl["categories"]
    for ser in sl["series"]:
        data.add_series(ser["name"], ser["values"])
    has_side = bool(sl.get("caption"))
    cw = W - 2 * M - (380 if has_side else 0)
    gf = s.shapes.add_chart(CHART_TYPES[kind], Pt(M), Pt(BODY_TOP), Pt(cw), Pt(BODY_BOTTOM - BODY_TOP), data)
    gf.name = "Chart"
    ch = gf.chart
    single = len(sl["series"]) == 1
    hi = sl.get("highlight")
    if isinstance(hi, str):
        hi = sl["categories"].index(hi) if hi in sl["categories"] else None
    ch.has_title = False
    ch.font.size = Pt(SIZE["label"])
    ch.font.color.theme_color = MUTED
    # legend: none for a single series; on the right otherwise (never top/bottom - it steals plot height)
    ch.has_legend = not single and kind != "pie"
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.RIGHT
        ch.legend.include_in_layout = False
    plot = ch.plots[0]
    # data labels: pie, or a single series of bars/columns; never on line charts
    labels = kind == "pie" or (single and kind in ("column", "bar"))
    plot.has_data_labels = labels
    if labels:
        dl = plot.data_labels
        dl.font.size, dl.font.bold = Pt(SIZE["label"]), True
        if sl.get("number_format"):
            dl.number_format, dl.number_format_is_linked = sl["number_format"], False
        if kind != "pie":
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
    if kind in ("column", "bar"):
        plot.gap_width = 75
    if kind != "pie":
        va = ch.value_axis
        va.has_major_gridlines = not labels
        if va.has_major_gridlines:
            va.major_gridlines.format.line.width = Pt(0.5)
            va.major_gridlines.format.line.color.theme_color = QUIET
        va.visible = not labels
        va.has_minor_gridlines = False
        if sl.get("number_format"):
            va.tick_labels.number_format, va.tick_labels.number_format_is_linked = sl["number_format"], False
        ca = ch.category_axis
        ca.tick_labels.font.size = Pt(SIZE["label"])
        ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        ca.has_major_gridlines = False
    # colour: one series -> highlight one point in the accent, the rest quiet; several -> accent ramp
    for si, ser in enumerate(plot.series):
        if kind == "line":
            ser.format.line.width = Pt(2.25)
            ser.format.line.color.theme_color = ACCENT if si == 0 else MUTED
            ser.smooth = False
            continue
        if single:
            for pi_, point in enumerate(ser.points):
                point.format.fill.solid()
                point.format.fill.fore_color.theme_color = ACCENT if (hi is None or pi_ == hi) else QUIET
        else:
            ser.format.fill.solid()
            ser.format.fill.fore_color.theme_color = ACCENT
            ser.format.fill.fore_color.brightness = [0, 0.35, -0.3, 0.6, -0.5, 0.75][si % 6]
    alt(gf, sl.get("alt") or chart_alt(sl))
    if has_side:
        b.text(s, "ChartNote", sl["caption"], W - M - 340, BODY_TOP + 40, 340, 400, SIZE["small"], MUTED)


def chart_alt(sl):
    parts = []
    for ser in sl["series"]:
        pairs = ", ".join(f"{c} {v:g}" for c, v in zip(sl["categories"], ser["values"]))
        parts.append(f"{ser['name']}: {pairs}")
    return f"{sl.get('type', 'column').title()} chart. " + "; ".join(parts) + "."


def p_table(b, s, sl):
    b.title(s, sl["title"])
    rows, cols = len(sl["rows"]) + 1, len(sl["header"])
    row_h = 56
    gf = s.shapes.add_table(rows, cols, Pt(M), Pt(BODY_TOP + 20), Pt(W - 2 * M), Pt(row_h * rows))
    gf.name = "Table"
    tbl = gf.table
    hi = sl.get("highlight_row")
    for c, head in enumerate(sl["header"]):
        cell = tbl.cell(0, c)
        cell.text = str(head)
        cell.fill.solid()
        cell.fill.fore_color.theme_color = TEXT
        for r in cell.text_frame.paragraphs[0].runs:
            r.font.size, r.font.bold = Pt(SIZE["label"]), True
            r.font.color.theme_color = BG
    for ri, row in enumerate(sl["rows"], 1):
        for c, val in enumerate(row):
            cell = tbl.cell(ri, c)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.theme_color = ACCENT if hi == ri - 1 else (QUIET if ri % 2 == 0 else BG)
            for r in cell.text_frame.paragraphs[0].runs:
                r.font.size = Pt(SIZE["label"])
                r.font.color.theme_color = BG if hi == ri - 1 else TEXT
            if c > 0 and str(val).replace(",", "").replace(".", "").replace("%", "").replace("-", "").strip().isdigit():
                cell.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
    alt(gf, sl.get("alt") or f"Table: {', '.join(map(str, sl['header']))}; {len(sl['rows'])} rows.")


def p_image(b, s, sl):
    b.title(s, sl["title"])
    from PIL import Image
    sys.path.insert(0, HERE)
    from cover_crop import cover_crop
    box_w, box_h = W - 2 * M, BODY_BOTTOM - BODY_TOP - (90 if sl.get("caption") else 0)
    img = Image.open(sl["image"])
    out = os.path.join(tempfile.mkdtemp(prefix="crop-"), "image.png")
    cover_crop(img, box_w, box_h, sl.get("focus_x", 0.5), sl.get("focus_y", 0.5)).save(out)
    pic = s.shapes.add_picture(out, Pt(M), Pt(BODY_TOP), Pt(box_w), Pt(box_h))
    pic.name = "Photo"
    alt(pic, sl.get("alt") or sl["title"])
    if sl.get("caption"):
        b.text(s, "Caption", sl["caption"], M, BODY_BOTTOM - 70, box_w, 70, SIZE["label"], MUTED)


def p_matrix(b, s, sl):
    b.title(s, sl["title"])
    q = sl["quadrants"]
    hi = sl.get("highlight")
    left, top = M + 70, BODY_TOP + 10
    w, h = (W - left - M - GAP) / 2, (BODY_BOTTOM - top - 50 - GAP) / 2
    for i, quad in enumerate(q):
        x = left + (i % 2) * (w + GAP)
        y = top + (i // 2) * (h + GAP)
        b.rect(s, f"Quadrant{i + 1}", x, y, w, h, ACCENT if i == hi else QUIET)
        fg = BG if i == hi else TEXT
        b.text(s, f"QHeading{i + 1}", quad["heading"], x + 28, y + 24, w - 56, 50, SIZE["h2"] - 2, fg, bold=True, heading=True)
        b.text(s, f"QText{i + 1}", quad.get("text", ""), x + 28, y + 88, w - 56, h - 100, SIZE["caption"], fg)
    if sl.get("x_axis"):
        b.text(s, "XAxis", sl["x_axis"] + " \u2192", left, BODY_BOTTOM - 40, W - left - M, 40, SIZE["label"], MUTED,
               align=PP_ALIGN.CENTER)
    if sl.get("y_axis"):
        tb = b.text(s, "YAxis", sl["y_axis"] + " \u2192", M - 200 + 30, top + h, 400, 40, SIZE["label"], MUTED,
                    align=PP_ALIGN.CENTER)
        tb.rotation = -90


PATTERNS = {"title": p_title, "section": p_section, "statement": p_statement, "big_number": p_big_number,
            "kpi": p_kpi, "bullets": p_bullets, "compare": p_compare, "process": p_process,
            "timeline": p_timeline, "quote": p_quote, "chart": p_chart, "table": p_table, "image": p_image,
            "matrix": p_matrix}


# ---------------------------------------------------------------- main

def build(spec, out):
    if spec.get("template"):
        prs = open_template(spec["template"])
        for sld in list(prs.slides._sldIdLst):  # start from the template's masters, not its slides
            prs.part.drop_rel(sld.rId)
            prs.slides._sldIdLst.remove(sld)
        themed = False
    else:
        prs = Presentation()
        apply_direction(prs, load_direction(spec.get("direction", "clean-corporate")))
        themed = True
    prs.slide_width, prs.slide_height = Pt(W), Pt(H)
    b = Builder(prs, themed)
    for n, sl in enumerate(spec["slides"], 1):
        s = b.new_slide(sl, n)
        PATTERNS[sl["pattern"]](b, s, sl)
        text = notes_text(sl.get("notes"))
        if text:
            s.notes_slide.notes_text_frame.text = text
    prs.save(out)
    return len(spec["slides"])


def load_spec(path):
    with open(path, encoding="utf-8") as fh:
        if path.lower().endswith((".yml", ".yaml")):
            import yaml  # uvx --with pyyaml
            return yaml.safe_load(fh)
        return json.load(fh)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", nargs="?")
    ap.add_argument("--out")
    ap.add_argument("--lint", action="store_true", help="run lint_deck.py on the result")
    ap.add_argument("--force", action="store_true", help="build even if the spec breaks a pattern limit")
    ap.add_argument("--list-directions", action="store_true")
    a = ap.parse_args()
    if a.list_directions:
        for d in json.load(open(os.path.join(HERE, "directions.json"), encoding="utf-8"))["directions"]:
            print(f"{d['id']:<18} {d['tone']:<10} {d['heading']} / {d['body']}  #{d['accent']} on #{d['background']}"
                  f"  — {d['mood']}")
        return
    if not a.spec or not a.out:
        ap.error("spec and --out are required")
    spec = load_spec(a.spec)
    base = os.path.dirname(os.path.abspath(a.spec))
    for sl in spec.get("slides", []):  # image paths are relative to the spec
        if sl.get("image") and not os.path.isabs(sl["image"]):
            sl["image"] = os.path.join(base, sl["image"])
    if spec.get("template") and not os.path.isabs(spec["template"]):
        spec["template"] = os.path.join(base, spec["template"])
    errors = check_spec(spec)
    for e in errors:
        print(f"spec: {e}", file=sys.stderr)
    if errors and not a.force:
        sys.exit(2)
    n = build(spec, a.out)
    print(f"{n} slides -> {a.out}")
    if a.lint:
        r = subprocess.run([sys.executable, os.path.join(HERE, "lint_deck.py"), a.out])
        sys.exit(r.returncode)


if __name__ == "__main__":
    main()
