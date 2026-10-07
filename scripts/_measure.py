"""Estimate how text wraps in a box, from real font metrics (Pillow; any OS).

PowerPoint only knows the real line breaks when it lays the slide out, but a good estimate catches
most overflow: measure each word with the actual font (or a metric-compatible twin - Carlito for
Calibri, Liberation Sans for Arial ...), wrap greedily at the box width, and multiply lines by the line
height. Fonts that can't be found are approximated from Liberation Sans/DejaVu Sans with a width factor.
Treat results as estimates: they are typically within one line of PowerPoint's layout.
"""
import glob
import os
import re
from functools import lru_cache

from PIL import ImageFont

LINE_HEIGHT = 1.2      # PowerPoint single spacing is ~1.2 x the font size
METRIC_TWINS = {       # same advance widths as the original
    "calibri": "carlito", "calibri light": "carlito", "cambria": "caladea",
    "arial": "liberation sans", "helvetica": "liberation sans", "arial nova": "liberation sans",
    "times new roman": "liberation serif", "times": "liberation serif",
    "courier new": "liberation mono", "consolas": "liberation mono",
}
WIDTH_FACTOR = {       # average width relative to Arial, for fonts we can't find
    "segoe ui": 1.0, "segoe ui semibold": 1.03, "segoe ui light": 0.95, "aptos": 1.0, "aptos display": 0.98,
    "inter": 1.04, "roboto": 0.97, "verdana": 1.17, "tahoma": 0.97, "trebuchet ms": 0.98, "georgia": 1.06,
    "palatino linotype": 0.98, "book antiqua": 0.98, "garamond": 0.9, "arial black": 1.27, "gill sans": 0.93,
    "century gothic": 1.1, "franklin gothic": 0.95, "lato": 0.98, "open sans": 1.04, "source sans pro": 0.93,
}
FONT_DIRS = [os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
             os.path.expanduser(r"~\AppData\Local\Microsoft\Windows\Fonts"),
             "/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts"),
             os.path.expanduser("~/.local/share/fonts"), "/Library/Fonts", "/System/Library/Fonts",
             os.path.expanduser("~/Library/Fonts")]


@lru_cache(maxsize=1)
def _index():
    """family (lower case) -> {"regular": path, "bold": path}"""
    idx = {}
    for d in FONT_DIRS:
        for path in glob.glob(os.path.join(d, "**", "*.[tT][tT][fFcC]"), recursive=True) + \
                glob.glob(os.path.join(d, "**", "*.[oO][tT][fF]"), recursive=True):
            try:
                family, style = ImageFont.truetype(path, 10).getname()
            except (OSError, ValueError):
                continue
            slot = "bold" if "bold" in style.lower() and "italic" not in style.lower() else \
                ("regular" if style.lower() in ("regular", "book", "normal", "roman") else None)
            if slot:
                idx.setdefault(family.lower(), {}).setdefault(slot, path)
    return idx


@lru_cache(maxsize=256)
def _font(family, bold, size_px):
    """(ImageFont, width factor) for a family; factor scales a stand-in font's widths."""
    idx = _index()
    fam = (family or "").lower().strip()
    for name, factor in ((fam, 1.0), (METRIC_TWINS.get(fam, ""), 1.0),
                         ("liberation sans", WIDTH_FACTOR.get(fam, 1.0)),
                         ("dejavu sans", WIDTH_FACTOR.get(fam, 1.0) * 0.9)):
        files = idx.get(name)
        if files:
            path = files.get("bold" if bold else "regular") or files.get("regular") or next(iter(files.values()))
            return ImageFont.truetype(path, size_px), factor
    return None, WIDTH_FACTOR.get(fam, 1.0)


def text_width(text, family, size_pt, bold=False):
    """Advance width of `text` in points."""
    scale = 4  # measure at 4 px per pt for precision
    font, factor = _font(family, bool(bold), max(1, round(size_pt * scale)))
    if font is None:
        return len(text) * 0.5 * size_pt * factor
    return font.getlength(text) / scale * factor


def wrap(text, family, size_pt, width_pt, bold=False):
    """Greedy word wrap. Returns (lines, widest_word_pt)."""
    lines, widest = 0, 0.0
    space = text_width(" ", family, size_pt, bold)
    for para in re.split(r"[\r\n\v]+", text) or [""]:
        words = para.split()
        if not words:
            lines += 1
            continue
        cur = 0.0
        lines += 1
        for w in words:
            ww = text_width(w, family, size_pt, bold)
            widest = max(widest, ww)
            if cur and cur + space + ww > width_pt:
                lines += 1
                cur = ww
            else:
                cur = cur + (space if cur else 0) + ww
            while cur > width_pt and width_pt > 0:  # a word wider than the line breaks inside itself
                lines += 1
                cur -= width_pt
    return lines, widest


def text_height(paragraphs, width_pt):
    """paragraphs: [(text, family, size_pt, bold, space_before_pt)] -> (height_pt, widest_word_pt, lines)."""
    height, widest, total = 0.0, 0.0, 0
    for i, (text, family, size, bold, before) in enumerate(paragraphs):
        n, w = wrap(text, family, size, width_pt, bold)
        height += n * size * LINE_HEIGHT + (before if i else 0)
        widest = max(widest, w)
        total += n
    return height, widest, total
