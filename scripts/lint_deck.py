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
import sys
from collections import Counter

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.util import Emu

from _rules import (OFFICE_DEFAULT_SERIES, contains, contrast_ratio, floor_for_room,
                    looks_like_label, overlap_area)

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


# ---------------------------------------------------------------- helpers

def rect(sh):
    return (Emu(sh.left or 0).pt, Emu(sh.top or 0).pt, Emu(sh.width or 0).pt, Emu(sh.height or 0).pt)


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


def lint_slide(n, slide, sw, sh_, floor, budget, f):
    shapes = list(walk(slide.shapes))
    titles = [s for s in shapes if is_title(s)]
    title_text = titles[0].text_frame.text.strip() if titles and titles[0].has_text_frame else ""

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

    words, colours = 0, set()
    content = [s for s in shapes if not is_title(s) and s.shape_type != MSO_SHAPE_TYPE.GROUP]
    for s in content:
        fill = None
        try:
            fill = solid_rgb(s.fill)
        except (AttributeError, NotImplementedError):
            pass
        if fill:
            colours.add(fill)

        # bounds
        l, t, w, h = rect(s)
        if l < -EDGE_TOLERANCE_PT or t < -EDGE_TOLERANCE_PT or l + w > sw + EDGE_TOLERANCE_PT or t + h > sh_ + EDGE_TOLERANCE_PT:
            f.add(n, "warn", "offslide_shape", "Shape extends outside the slide.", s.name)

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
            if len(paras) > MAX_BULLETS:
                f.add(n, "warn", "too_many_bullets", f"{len(paras)} paragraphs (> {MAX_BULLETS}); split the slide or "
                      "reveal one at a time.", s.name)
            fill_bg = fill
            for p in paras:
                size = para_size(s, p)
                if size and LABEL_MAX_PT < size < floor and len(p.text.split()) >= BODY_MIN_WORDS:
                    f.add(n, "warn", "body_below_floor", f"{size:g} pt body text (floor {floor:g} pt): "
                          f"'{p.text.strip()[:40]}'", s.name)
                    break
            for p in paras:
                size = para_size(s, p) or 18.0
                for r in p.runs:
                    fg = run_rgb(r)
                    if fg and fill_bg:
                        need = 3.0 if size >= 24 or (size >= 18.7 and r.font.bold) else 4.5
                        ratio = contrast_ratio(fg, fill_bg)
                        if ratio < need:
                            f.add(n, "error" if ratio < 3.0 else "warn", "a11y_low_text_contrast",
                                  f"Text #{fg} on #{fill_bg} is {ratio:.1f}:1 (needs {need}:1).", s.name)
                            break
                else:
                    continue
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
        if getattr(s, "has_chart", False):
            lint_chart(n, s, f)

    if words > budget:
        f.add(n, "info", "word_budget", f"{words} visible words (budget {budget}). Fine for gallery, matrix, "
              "chart, quote and reference slides; otherwise cut or move to the notes.")
    if len(colours) > MAX_COLOURS:
        f.add(n, "warn", "palette_too_many_colours", f"{len(colours)} distinct fill colours (> {MAX_COLOURS}).")

    # overlaps between text-bearing shapes and pictures (containment = a card/backing, not a defect)
    boxes = [(s.name, rect(s)) for s in content
             if (s.has_text_frame and s.text_frame.text.strip()) or s.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART)]
    for i, (na, ra) in enumerate(boxes):
        for nb, rb in boxes[i + 1:]:
            area = overlap_area(ra, rb)
            if area >= OVERLAP_WARN_PT2 and not contains(ra, rb) and not contains(rb, ra):
                f.add(n, "error" if area >= OVERLAP_ERROR_PT2 else "warn", "shape_overlap",
                      f"'{na}' and '{nb}' overlap by {area:.0f} pt².", na)

    if not slide.has_notes_slide or not slide.notes_slide.notes_text_frame.text.strip():
        f.add(n, "info", "missing_notes", "No speaker notes.")
    return title_text


def lint_chart(n, s, f):
    chart = s.chart
    series = list(chart.plots[0].series) if len(chart.plots) else []
    colours = []
    for ser in series:
        rgb = solid_rgb(ser.format.fill)
        if rgb:
            colours.append(rgb)
    if colours and all(c in OFFICE_DEFAULT_SERIES for c in colours):
        f.add(n, "warn", "chart_default_palette", "Series use Office default colours; bind them to the "
              "theme or highlight only the finding.", s.name)
    if chart.has_title and chart.chart_title.has_text_frame:
        t = chart.chart_title.text_frame.text.strip()
        if t and not any(ch.isdigit() for ch in t) and not looks_like_label(t) and " by " in f" {t.lower()} ":
            f.add(n, "info", "chart_descriptive_title", f"Chart title '{t}' names the data, not the finding.", s.name)
    plot = chart.plots[0] if len(chart.plots) else None
    if plot is not None and plot.has_data_labels and chart.value_axis is not None:
        try:
            if chart.value_axis.visible and chart.value_axis.has_major_gridlines:
                f.add(n, "info", "chart_redundant_labels", "Data labels and a gridlined value axis say the "
                      "same thing; keep one.", s.name)
        except (AttributeError, ValueError):
            pass


def lint(path, floor, budget):
    prs = Presentation(path)
    f = Findings()
    sw, sh_ = Emu(prs.slide_width).pt, Emu(prs.slide_height).pt
    if (round(sw), round(sh_)) != (1440, 810):
        ratio = sw / sh_
        f.add(0, "warn" if abs(ratio - 16 / 9) > 0.01 else "info", "slide_size",
              f"Slide size {sw:g} x {sh_:g} pt; Full HD is 1440 x 810 pt.")
    titles = Counter()
    fonts = Counter()
    for n, slide in enumerate(prs.slides, 1):
        if slide._element.get("show") == "0":
            continue
        title = lint_slide(n, slide, sw, sh_, floor, budget, f)
        if title:
            titles[title.lower()] += 1
        for s in walk(slide.shapes):
            if s.has_text_frame:
                for p in s.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.name:
                            fonts[r.font.name] += 1
    for t, c in titles.items():
        if c > 1:
            f.add(0, "warn", "duplicate_titles", f"{c} slides share the title '{t}'; screen-reader and "
                  "outline navigation can't tell them apart.")
    if len(fonts) > 3:
        f.add(0, "warn", "mixed_font_families", f"{len(fonts)} font families set directly on text: "
              f"{', '.join(sorted(fonts))}. Use the theme fonts (display + body).")
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
        print(json.dumps({"file": a.file, "floor_pt": floor, "counts": dict(sev), "findings": f.items},
                         ensure_ascii=False, indent=2))
    else:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
        for i in sorted(f.items, key=lambda x: (x["slide"], {"error": 0, "warn": 1, "info": 2}[x["severity"]])):
            where = f"slide {i['slide']}" if i["slide"] else "deck"
            shape = f" [{i['shape']}]" if i["shape"] else ""
            print(f"{i['severity']:<5}  {where:<9} {i['code']:<24}{shape} {i['message']}")
        print(f"\n{sev.get('error', 0)} error(s), {sev.get('warn', 0)} warning(s), {sev.get('info', 0)} info  "
              f"(body floor {floor:g} pt)")
    if a.fix:
        done = fix(prs, f)
        prs.save(a.out)
        print(f"fixed {done} issue(s) -> {a.out}", file=sys.stderr)
    failing = {"error"} if a.fail_on == "error" else {"error", "warn"}
    sys.exit(1 if any(i["severity"] in failing for i in f.items) else 0)


if __name__ == "__main__":
    main()
