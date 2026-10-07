"""Lint a .pptx for the defect codes in reference/AUDIT.md - from the file alone, on any OS.

    uvx --with python-pptx --with pillow python scripts/lint_deck.py deck.pptx
    uvx --with python-pptx --with pillow python scripts/lint_deck.py deck.pptx --json
    uvx --with python-pptx --with pillow python scripts/lint_deck.py deck.pptx --room-depth 45
    uvx --with python-pptx --with pillow python scripts/lint_deck.py deck.pptx --fix --out fixed.pptx

No PowerPoint and no rendering: geometry, text, fonts, colours, pictures, charts and alt text are
read from the OOXML. What it cannot see - real line breaks, text that overflows its box, how the
slide looks - still needs a render (render_slides.py on Windows, render_lo.py elsewhere).

Exit code: 1 if any finding has severity "error" (or "warn" with --fail-on warn), else 0.
--fix applies only safe fixes (deletes empty placeholders on slides that have other content)
and writes to --out; it never overwrites the input.
"""
import argparse
import io
import json
import re
import sys
from collections import Counter

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.util import Emu

from _rules import (CARTOON_HOSTS, DEFAULT_FACES, INSIGHT_WORDS, OFFICE_DEFAULT_SERIES, ORDINAL_RE,
                    STOCK_HOSTS, STOP_WORDS, chars_per_line, contains, contrast_ratio, floor_for_room,
                    has_emoji, is_large_text, looks_like_label, overlap_area)
from _theme import Theme, hue, saturation

PT = 12700  # EMU per point
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a16": "http://schemas.microsoft.com/office/drawing/2014/main"}
TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.VERTICAL_TITLE}
TEXT_PLACEHOLDERS = TITLE_TYPES | {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.OBJECT}
LABEL_MAX_PT = 12.0        # at or below: caption tier, never a body-floor finding
BODY_MIN_WORDS = 4         # a paragraph with fewer words is a label
TITLE_MAX_CHARS = 55
MAX_BULLETS = 7
MAX_COLOURS = 5
OVERLAP_WARN_PT2 = 4.0     # PointClaw thresholds: >= 4 pt2 warn, >= 200 pt2 error
OVERLAP_ERROR_PT2 = 200.0
EDGE_TOLERANCE_PT = 2.0
STRETCH_TOLERANCE = 0.03
SCALE = 1.0                # slide width / 960 pt, set per deck in lint()


# ---------------------------------------------------------------- helpers

def rect(sh):
    """(left, top, width, height) in pt as drawn - a box turned 90/270 degrees swaps width and height
    around its centre."""
    l, t, w, h = (Emu(sh.left or 0).pt, Emu(sh.top or 0).pt, Emu(sh.width or 0).pt, Emu(sh.height or 0).pt)
    rot = round(getattr(sh, "rotation", 0) or 0) % 180
    if rot == 90:
        cx, cy = l + w / 2, t + h / 2
        l, t, w, h = cx - h / 2, cy - w / 2, h, w
    return (l, t, w, h)


def walk(shapes):
    """Every shape, descending into groups (group children keep absolute coordinates in python-pptx)."""
    for sh in shapes:
        yield sh
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk(sh.shapes)


def ph_type(sh):
    try:
        return sh.placeholder_format.type if sh.is_placeholder else None
    except ValueError:
        return None


def is_title(sh):
    return ph_type(sh) in TITLE_TYPES


def level_size(lst_style, level):
    """Size from an <a:lstStyle>/<p:txStyles> list for paragraph level (1-based), or None."""
    if lst_style is None:
        return None
    node = lst_style.find(f"a:lvl{level}pPr/a:defRPr", NS)
    sz = node.get("sz") if node is not None else None
    return int(sz) / 100 if sz else None


def inherited_size(sh, para):
    """Font size of a paragraph when its runs don't set one: shape lstStyle -> layout placeholder
    -> master text style. Returns None when nothing in the chain sets a size."""
    level = (para.level or 0) + 1
    body = sh.text_frame._txBody
    size = level_size(body.find("a:lstStyle", NS), level)
    if size:
        return size
    pt = ph_type(sh)
    if pt is not None:
        try:
            layout_ph = sh.part.slide.slide_layout.placeholders.get(idx=sh.placeholder_format.idx)
        except (AttributeError, KeyError):
            layout_ph = None
        if layout_ph is not None and layout_ph.has_text_frame:
            size = level_size(layout_ph.text_frame._txBody.find("a:lstStyle", NS), level)
            if size:
                return size
    master = sh.part.slide.slide_layout.slide_master._element
    style = "titleStyle" if pt in TITLE_TYPES else ("bodyStyle" if pt is not None else "otherStyle")
    size = level_size(master.find(f"p:txStyles/p:{style}", NS), level)
    return size or (18.0 if pt is None else None)  # PowerPoint's default text box size is 18 pt


def para_size(sh, para):
    sizes = [r.font.size.pt for r in para.runs if r.font.size is not None]
    return min(sizes) if sizes else inherited_size(sh, para)


def solid_rgb(fill):
    try:
        if fill.type == 1:  # MSO_FILL.SOLID
            return str(fill.fore_color.rgb)
    except (AttributeError, TypeError, ValueError):
        pass
    return None


def run_rgb(run):
    try:
        return str(run.font.color.rgb) if run.font.color and run.font.color.type is not None else None
    except (AttributeError, TypeError, ValueError):
        return None


IMAGE_EXT = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".svg", ".webp")


def alt_text(sh):
    """Alt text, treating a bare file name (python-pptx and some tools write one) as missing."""
    nv = sh._element.find(".//p:cNvPr", NS)
    text = (nv.get("descr") or nv.get("title") or "").strip() if nv is not None else ""
    return "" if text.lower().endswith(IMAGE_EXT) and " " not in text else text


def is_decorative(sh):
    nv = sh._element.find(".//p:cNvPr", NS)
    if nv is None:
        return False
    dec = nv.find(".//a16:decorative", NS)
    return dec is not None and dec.get("val") in ("1", "true")


def picture_stretch(sh):
    """Relative difference between the picture's shown aspect and its cropped native aspect."""
    try:
        from PIL import Image
        w_px, h_px = Image.open(io.BytesIO(sh.image.blob)).size
    except Exception:  # unreadable or vector image (EMF/SVG): nothing to compare
        return 0.0
    cw = w_px * (1 - sh.crop_left - sh.crop_right)
    ch = h_px * (1 - sh.crop_top - sh.crop_bottom)
    if cw <= 0 or ch <= 0 or not sh.width or not sh.height:
        return 0.0
    return (sh.width / sh.height) / (cw / ch) - 1


# ---------------------------------------------------------------- checks

class Findings:
    def __init__(self):
        self.items = []

    def add(self, slide, severity, code, message, shape=None):
        self.items.append({"slide": slide, "severity": severity, "code": code,
                           "shape": shape, "message": message})


def shape_fill(s, theme):
    """Solid fill colour of a shape (theme-resolved), or None."""
    sp = s._element.find(".//p:spPr", NS)
    if sp is None:
        return None
    solid = sp.find("a:solidFill", NS)
    if solid is not None:
        return theme.colour_in(solid)
    ref = s._element.find("p:style/a:fillRef", NS)  # shape style default fill (idx 0 = none)
    if ref is not None and ref.get("idx", "0") != "0" and sp.find("a:noFill", NS) is None \
            and sp.find("a:gradFill", NS) is None and sp.find("a:blipFill", NS) is None:
        return theme.colour_in(ref)
    return None


def has_other_fill(s):
    sp = s._element.find(".//p:spPr", NS)
    return sp is not None and (sp.find("a:gradFill", NS) is not None or sp.find("a:blipFill", NS) is not None
                               or sp.find("a:pattFill", NS) is not None)


def run_colour(run, theme):
    rpr = run._r.find("a:rPr", NS)
    solid = rpr.find("a:solidFill", NS) if rpr is not None else None
    return theme.colour_in(solid) if solid is not None else None


def lint_slide(n, slide, sw, sh_, floor, budget, f, theme):
    shapes = list(walk(slide.shapes))
    titles = [s for s in shapes if is_title(s)]
    title_text = titles[0].text_frame.text.strip() if titles and titles[0].has_text_frame else ""
    slide_bg = theme.background(slide)

    # titles
    if not titles:
        f.add(n, "warn", "missing_title", "No title placeholder. Every slide needs a title, even a hidden one, "
              "for navigation and screen readers.")
    elif not title_text:
        f.add(n, "warn", "empty_title", "Title placeholder is empty.", titles[0].name)
    else:
        if len(title_text) > TITLE_MAX_CHARS:
            f.add(n, "warn", "headline_too_long",
                  f"Title is {len(title_text)} characters (> {TITLE_MAX_CHARS}); it will likely wrap.", titles[0].name)
        if "\n" in title_text or "\v" in title_text:
            f.add(n, "warn", "headline_two_line", "Title has a hard line break; make it one line or split "
                  "headline and subhead.", titles[0].name)
        if looks_like_label(title_text):
            f.add(n, "info", "title_is_label", f"'{title_text}' reads as a topic label, not a claim "
                  "(fine for dividers, agenda and Q&A).", titles[0].name)

    words, colours, raw_fills, shadows, accents = 0, set(), 0, 0, set()
    sizes, big_tokens, contrast_done = [], Counter(), False
    content = [s for s in shapes if not is_title(s) and s.shape_type != MSO_SHAPE_TYPE.GROUP]
    text_shapes = [s for s in shapes if s.has_text_frame and s.text_frame.text.strip()]
    for s in content:
        fill = shape_fill(s, theme)
        if fill:
            colours.add(fill)
        el = s._element
        sp = el.find(".//p:spPr", NS)
        if sp is not None:
            srgb = sp.find("a:solidFill/a:srgbClr", NS)
            if srgb is not None and not theme.is_theme_hex(srgb.get("val", "")):
                raw_fills += 1
            if sp.find("a:effectLst/a:outerShdw", NS) is not None:
                shadows += 1
            stops = sp.findall("a:gradFill/a:gsLst/a:gs/a:srgbClr", NS)
            if stops and max(saturation(c.get("val", "000000")) for c in stops) > 0.45:
                f.add(n, "warn", "gradient_high_chroma", "Saturated gradient that isn't built from theme colours "
                      "(a common machine-made look); use a solid accent or a two-stop brand gradient.", s.name)
        accents.update(c.get("val") for c in el.iter(f"{{{NS['a']}}}schemeClr") if c.get("val", "").startswith("accent"))

        # bounds
        l, t, w, h = rect(s)
        if l < -EDGE_TOLERANCE_PT or t < -EDGE_TOLERANCE_PT or l + w > sw + EDGE_TOLERANCE_PT or t + h > sh_ + EDGE_TOLERANCE_PT:
            f.add(n, "warn", "offslide_shape", "Shape extends outside the slide.", s.name)
        if el.find(".//a:hlinkClick", NS) is not None and el.find(".//p:cNvPr/a:hlinkClick", NS) is not None \
                and (w < 32 or h < 32):
            f.add(n, "warn", "tiny_click_target", f"Clickable shape is {w:.0f} × {h:.0f} pt; make it at least "
                  "44 × 44 pt for touch.", s.name)

        # placeholders left empty
        pt = ph_type(s)
        if pt in TEXT_PLACEHOLDERS and s.has_text_frame and not s.text_frame.text.strip():
            if any(o is not s and (o.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART, MSO_SHAPE_TYPE.TABLE)
                                   or (o.has_text_frame and o.text_frame.text.strip() and not is_title(o)))
                   for o in content):
                f.add(n, "error", "unused_placeholder", "Empty placeholder next to real content shows "
                      "'Click to add text' in edit view and confuses screen readers.", s.name)

        # text
        if s.has_text_frame and s.text_frame.text.strip():
            paras = [p for p in s.text_frame.paragraphs if p.text.strip()]
            words += sum(len(p.text.split()) for p in paras)
            text = s.text_frame.text
            if len(paras) > MAX_BULLETS:
                f.add(n, "warn", "too_many_bullets", f"{len(paras)} paragraphs (> {MAX_BULLETS}); split the slide or "
                      "reveal one at a time.", s.name)
            if re.search(r"\blorem ipsum\b|\bdolor sit amet\b", text, re.I):
                f.add(n, "error", "lorem_ipsum", "Placeholder 'lorem ipsum' text left in.", s.name)
            for p in paras:
                size = para_size(s, p)
                if size and LABEL_MAX_PT < size < floor and len(p.text.split()) >= BODY_MIN_WORDS:
                    f.add(n, "warn", "body_below_floor", f"{size:g} pt body text (floor {floor:g} pt): "
                          f"'{p.text.strip()[:40]}'", s.name)
                    break
            for p in paras:
                pt_text = p.text.strip()
                size = para_size(s, p) or 18.0
                sizes.append(size)
                if len(pt_text) <= 24 and has_emoji(pt_text):
                    f.add(n, "warn", "emoji_as_icon", f"Emoji used as an icon ('{pt_text}'); use a real icon "
                          "with a text label.", s.name)
                if len(pt_text) >= 8 and pt_text.endswith(("…", "...")):
                    f.add(n, "warn", "truncated_text", f"Text ends in an ellipsis ('…{pt_text[-24:]}'); "
                          "is something cut off?", s.name)
                if len(pt_text) >= 80 and p.alignment == 2:  # PP_ALIGN.CENTER
                    f.add(n, "warn", "centered_long_body", "Long centred text is hard to read; left-align body "
                          "text and keep centring for one-line headlines and quotes.", s.name)
                if len(pt_text) >= 90 and size <= 28 * SCALE:
                    cpl = chars_per_line(Emu(s.width or 0).pt, size)
                    if cpl > 75:
                        f.add(n, "warn", "measure_too_wide", f"About {cpl:.0f} characters per line (comfortable: "
                              "45–75); narrow the box or raise the size.", s.name)
                if size >= 24 * SCALE:
                    big_tokens.update(w.lower() for w in re.findall(r"[^\W\d_]{4,}", pt_text)
                                      if w.lower() not in STOP_WORDS)
            # contrast: run colour (or theme text colour) against shape fill (or slide background)
            backing = fill
            if backing is None and not has_other_fill(s):  # a card or band drawn behind the text box
                for o in content:
                    if o is s:
                        break
                    if contains(rect(o), rect(s), tol=2.0):
                        if has_other_fill(o):
                            backing = "?"
                        else:
                            backing = shape_fill(o, theme) or backing
            if not contrast_done and not has_other_fill(s) and backing != "?":
                bg = backing or slide_bg
                fill = backing
                default_fg = theme.colours.get(theme.map.get("tx1", "dk1"))
                for p in paras:
                    size = para_size(s, p) or 18.0
                    for r in p.runs:
                        if not r.text.strip():
                            continue
                        fg = run_colour(r, theme) or (default_fg if fill else None)
                        if not fg or not bg:
                            continue
                        need = 3.0 if is_large_text(size / SCALE, bool(r.font.bold)) else 4.5
                        ratio = contrast_ratio(fg, bg)
                        if ratio < need:
                            f.add(n, "error" if ratio < 3.0 else "warn", "a11y_low_text_contrast",
                                  f"Text #{fg} on #{bg} is {ratio:.1f}:1 (needs {need}:1).", s.name)
                            contrast_done = True
                            break
                    if contrast_done:
                        break

        # pictures and charts
        if s.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART) or getattr(s, "has_chart", False):
            if not alt_text(s) and not is_decorative(s):
                f.add(n, "warn", "a11y_missing_alt_text", "No alt text (and not marked decorative).", s.name)
        if s.shape_type == MSO_SHAPE_TYPE.PICTURE:
            stretch = picture_stretch(s)
            if abs(stretch) > STRETCH_TOLERANCE:
                f.add(n, "error" if abs(stretch) > 0.15 else "warn", "picture_stretched",
                      f"Picture is {stretch:+.0%} {'wider' if stretch > 0 else 'narrower'} than its source; "
                      "crop instead (scripts/cover_crop.py).", s.name)
            hay = (alt_text(s) + " " + s.name + " " + " ".join(
                r.target_ref for r in s.part.rels.values() if r.is_external)).lower()
            if any(h in hay for h in STOCK_HOSTS + CARTOON_HOSTS):
                f.add(n, "info", "stock_or_cartoon_image", "Looks like a stock photo or generic illustration; "
                      "real product, team or customer images carry more weight.", s.name)
        if getattr(s, "has_chart", False):
            lint_chart(n, s, f, theme)

    if words > budget:
        f.add(n, "info", "word_budget", f"{words} visible words (budget {budget}). Fine for gallery, matrix, "
              "chart, quote and reference slides; otherwise cut or move to the notes.")
    if len(colours) > MAX_COLOURS:
        f.add(n, "warn", "palette_too_many_colours", f"{len(colours)} distinct fill colours (> {MAX_COLOURS}).")
    if raw_fills >= 3:
        f.add(n, "warn", "off_palette_fill", f"{raw_fills} shapes use hard-coded colours outside the theme; "
              "use theme colours so the deck re-themes cleanly.")
    if shadows > 3:
        f.add(n, "warn", "shadow_overuse", f"{shadows} shapes have drop shadows; reserve elevation for one or two.")
    if len(accents) >= 4:
        f.add(n, "warn", "accent_overload", f"{len(accents)} accent colours on one slide ({', '.join(sorted(accents))}); "
              "keep to 3 or fewer, ideally one highlight.")
    for tok, c in big_tokens.items():
        if c >= 3:
            f.add(n, "warn", "repeated_word", f"'{tok}' appears {c} times in large type; demote the repeats.")
            break
    big = sorted(set(sizes), reverse=True)
    if len(big) >= 2 and 1.08 < big[0] / big[1] < 1.6:
        f.add(n, "info", "weak_focal_hierarchy", f"Largest text sizes {big[0]:g} and {big[1]:g} pt are too close; "
              "make one element clearly dominant (≥ 1.6×) or equal.")

    # grid monotony: 4+ body boxes, 3+ of them the same width and top as the first
    area = sw * sh_
    body = [rect(s) for s in text_shapes if not is_title(s) and rect(s)[2] > 1 and rect(s)[3] > 1
            and rect(s)[2] * rect(s)[3] < 0.92 * area]
    if len(body) >= 4:
        same = sum(1 for r in body[1:] if abs(r[2] - body[0][2]) <= 4 and abs(r[1] - body[0][1]) <= 4)
        if same >= 3:
            f.add(n, "info", "grid_monotony", f"{same + 1} identical boxes in a row; vary size or emphasis so "
                  "one item leads.")

    # overlaps between text-bearing shapes and pictures (containment = a card/backing, not a defect)
    boxes = [(s.name, rect(s)) for s in content
             if (s.has_text_frame and s.text_frame.text.strip()) or s.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART)]
    for i, (na, ra) in enumerate(boxes):
        for nb, rb in boxes[i + 1:]:
            ov = overlap_area(ra, rb)
            if ov >= OVERLAP_WARN_PT2 and not contains(ra, rb) and not contains(rb, ra):
                f.add(n, "error" if ov >= OVERLAP_ERROR_PT2 else "warn", "shape_overlap",
                      f"'{na}' and '{nb}' overlap by {ov:.0f} pt².", na)

    if not slide.has_notes_slide or not slide.notes_slide.notes_text_frame.text.strip():
        f.add(n, "info", "missing_notes", "No speaker notes.")
    return title_text


C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def lint_chart(n, s, f, theme):
    chart = s.chart
    cx = chart._chartSpace
    c = {"c": C_NS, "a": NS["a"]}
    series = cx.findall(".//c:ser", c)
    colours = []
    for ser in series[:6]:
        colours.append(theme.colour_in(ser.find("c:spPr/a:solidFill", c)))
    if sum(1 for col in colours if col in OFFICE_DEFAULT_SERIES) >= 2:
        f.add(n, "warn", "chart_default_palette", "Series use Office default colours; bind them to the "
              "theme, or highlight only the finding.", s.name)

    if chart.has_title and chart.chart_title.has_text_frame:
        t = chart.chart_title.text_frame.text.strip()
        if t and not any(w in f" {t.lower()} " for w in INSIGHT_WORDS):
            f.add(n, "info", "chart_descriptive_title", f"Chart title '{t}' names the data; headline the finding "
                  "instead (e.g. 'East leads Q1, up 8 %').", s.name)

    labels_on = any(d.find("c:showVal", c) is not None and d.find("c:showVal", c).get("val") in ("1", "true")
                    for d in cx.findall(".//c:dLbls", c))
    gridlines = cx.find(".//c:valAx/c:majorGridlines", c) is not None
    val_deleted = cx.find(".//c:valAx/c:delete", c)
    if labels_on and gridlines and not (val_deleted is not None and val_deleted.get("val") in ("1", "true")):
        f.add(n, "info", "chart_redundant_labels", "Data labels and a gridlined value axis say the same thing; "
              "keep one.", s.name)

    legend = cx.find(".//c:legend", c)
    if legend is not None:
        pos = legend.find("c:legendPos", c)
        overlay = legend.find("c:overlay", c)
        if pos is not None and pos.get("val") in ("t", "b") and (overlay is None or overlay.get("val") in ("0", "false")):
            f.add(n, "warn", "chart_legend_steals_plot", "Legend at the top/bottom takes height from the plot; move "
                  "it to the right, or drop it and colour-code the title.", s.name)

    names = [(ser.findtext(".//c:tx//c:v", namespaces=c) or "").strip().lower() for ser in series]
    real = [col for col in colours if col]
    if len(names) >= 3 and all(re.match(ORDINAL_RE, nm) for nm in names if nm) and all(names) \
            and len(real) >= 2 and len({round(hue(col), 2) for col in real}) >= 2:
        f.add(n, "info", "chart_ordinal_categorical_color", "Ordered series (months, quarters, years) in unrelated "
              "hues; use one hue from light to dark so the order reads at a glance.", s.name)

    for fmt in cx.findall(".//c:valAx/c:numFmt", c):
        code = fmt.get("formatCode", "")
        if "_(" in code or ('"-"' in code and "*" in code):
            f.add(n, "warn", "chart_accounting_zero_dash", "Accounting number format shows zero as '$-' on the "
                  "axis; use a currency or number format.", s.name)
            break

    points = max((len(ser.findall(".//c:val//c:pt", c)) for ser in series), default=0)
    if labels_on and (points > 12 or len(series) > 1):
        f.add(n, "info", "chart_label_collision", f"Data labels on {len(series)} series × {points} points will "
              "likely collide; label only the key points or rely on the axis. Check the render.", s.name)


def lint(path, floor, budget):
    """Sizes are judged relative to a standard 960-pt-wide slide: on a 1440-pt (Full HD) slide an
    18 pt floor becomes 27 pt, because the same text is two-thirds as big on screen."""
    global LABEL_MAX_PT, SCALE
    prs = Presentation(path)
    f = Findings()
    theme = Theme(prs.slide_master)
    sw, sh_ = Emu(prs.slide_width).pt, Emu(prs.slide_height).pt
    scale = SCALE = max(1.0, sw / 960)
    floor, LABEL_MAX_PT = round(floor * scale, 1), round(12.0 * scale, 1)
    if (round(sw), round(sh_)) != (1440, 810):
        ratio = sw / sh_
        f.add(0, "warn" if abs(ratio - 16 / 9) > 0.01 else "info", "slide_size",
              f"Slide size {sw:g} x {sh_:g} pt; Full HD is 1440 x 810 pt.")
    titles = Counter()
    fonts = Counter()
    resolved = Counter()
    for n, slide in enumerate(prs.slides, 1):
        if slide._element.get("show") == "0":
            continue
        title = lint_slide(n, slide, sw, sh_, floor, budget, f, theme)
        if title:
            titles[title.lower()] += 1
        for s in walk(slide.shapes):
            if s.has_text_frame:
                for p in s.text_frame.paragraphs:
                    for r in p.runs:
                        if not r.text.strip():
                            continue
                        name = r.font.name
                        if name and not name.startswith("+"):
                            fonts[name] += 1
                        face = theme.font(name) if name else (theme.major if is_title(s) else theme.minor)
                        if face:
                            resolved[face.lower()] += 1
    for t, c in titles.items():
        if c > 1:
            f.add(0, "warn", "duplicate_titles", f"{c} slides share the title '{t}'; screen-reader and "
                  "outline navigation can't tell them apart.")
    if len(fonts) > 3:
        f.add(0, "warn", "mixed_font_families", f"{len(fonts)} font families set directly on text: "
              f"{', '.join(sorted(fonts))}. Use the theme fonts (display + body).")
    if resolved and len(resolved) <= 1 and set(resolved) <= DEFAULT_FACES:
        f.add(0, "info", "default_font_only", f"Only {next(iter(resolved)).title()} is used. Fine for a quick "
              "internal deck; for anything branded, pair a distinctive heading face with the body font.")
    f.floor = floor
    return prs, f


def fix(prs, findings):
    """Safe fixes only: delete empty placeholders flagged as unused."""
    done = 0
    targets = {(i["slide"], i["shape"]) for i in findings.items if i["code"] == "unused_placeholder"}
    for n, slide in enumerate(prs.slides, 1):
        for s in list(slide.shapes):
            if (n, s.name) in targets:
                s._element.getparent().remove(s._element)
                done += 1
    return done


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--floor", type=float, help="body font floor in pt (default 18, or from --room-depth)")
    ap.add_argument("--room-depth", type=float, help="viewing distance in feet; sets the floor (20->14, 30->18, 50->24, more->28)")
    ap.add_argument("--budget", type=int, default=12, help="visible words per slide before an info finding (default 12)")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--fail-on", choices=["error", "warn"], default="error")
    ap.add_argument("--fix", action="store_true", help="apply safe fixes; needs --out")
    ap.add_argument("--out")
    a = ap.parse_args()
    if a.fix and not a.out:
        ap.error("--fix needs --out (the input is never overwritten)")
    floor = a.floor or (floor_for_room(a.room_depth) if a.room_depth else 18.0)

    prs, f = lint(a.file, floor, a.budget)
    sev = Counter(i["severity"] for i in f.items)
    if a.json:
        print(json.dumps({"file": a.file, "floor_pt": f.floor, "counts": dict(sev), "findings": f.items},
                         ensure_ascii=False, indent=2))
    else:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
        for i in sorted(f.items, key=lambda x: (x["slide"], {"error": 0, "warn": 1, "info": 2}[x["severity"]])):
            where = f"slide {i['slide']}" if i["slide"] else "deck"
            shape = f" [{i['shape']}]" if i["shape"] else ""
            print(f"{i['severity']:<5}  {where:<9} {i['code']:<24}{shape} {i['message']}")
        print(f"\n{sev.get('error', 0)} error(s), {sev.get('warn', 0)} warning(s), {sev.get('info', 0)} info  "
              f"(body floor {f.floor:g} pt)")
    if a.fix:
        done = fix(prs, f)
        prs.save(a.out)
        print(f"fixed {done} issue(s) -> {a.out}", file=sys.stderr)
    failing = {"error"} if a.fail_on == "error" else {"error", "warn"}
    sys.exit(1 if any(i["severity"] in failing for i in f.items) else 0)


if __name__ == "__main__":
    main()
