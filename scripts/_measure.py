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

from kShared import kS

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


class kMeasure:
    """Stateless text measurement from real font metrics."""

    @staticmethod
    @lru_cache(maxsize=1)
    def Index():
        """family (lower case) -> {"regular": path, "bold": path}"""
        if kS.ErrorMode:
            return {}
        try:
            Idx = {}
            for D in FONT_DIRS:
                for Path in glob.glob(os.path.join(D, "**", "*.[tT][tT][fFcC]"), recursive=True) + \
                        glob.glob(os.path.join(D, "**", "*.[oO][tT][fF]"), recursive=True):
                    try:
                        Family, Style = ImageFont.truetype(Path, 10).getname()
                    except (OSError, ValueError):  # ERROR-SUPPRESSED-JUSTIFIED: an unreadable font file is skipped, the others still index
                        continue
                    Slot = "bold" if "bold" in Style.lower() and "italic" not in Style.lower() else \
                        ("regular" if Style.lower() in ("regular", "book", "normal", "roman") else None)
                    if Slot:
                        Idx.setdefault(Family.lower(), {}).setdefault(Slot, Path)
            return Idx
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMeasure.Index")
            return {}

    @staticmethod
    @lru_cache(maxsize=256)
    def Font(Family, Bold, SizePx):
        """(ImageFont, width factor) for a family; factor scales a stand-in font's widths."""
        if kS.ErrorMode:
            return None, 1.0
        try:
            Idx = kMeasure.Index()
            Fam = (Family or "").lower().strip()
            for Name, Factor in ((Fam, 1.0), (METRIC_TWINS.get(Fam, ""), 1.0),
                                 ("liberation sans", WIDTH_FACTOR.get(Fam, 1.0)),
                                 ("dejavu sans", WIDTH_FACTOR.get(Fam, 1.0) * 0.9)):
                Files = Idx.get(Name)
                if Files:
                    Path = Files.get("bold" if Bold else "regular") or Files.get("regular") or next(iter(Files.values()))
                    return ImageFont.truetype(Path, SizePx), Factor
            return None, WIDTH_FACTOR.get(Fam, 1.0)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMeasure.Font")
            return None, 1.0

    @staticmethod
    def TextWidth(Text, Family, SizePt, Bold=False):
        """Advance width of `Text` in points."""
        if kS.ErrorMode:
            return 0.0
        try:
            Scale = 4  # measure at 4 px per pt for precision
            Font, Factor = kMeasure.Font(Family, bool(Bold), max(1, round(SizePt * Scale)))
            if Font is None:
                return len(Text) * 0.5 * SizePt * Factor
            return Font.getlength(Text) / Scale * Factor
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMeasure.TextWidth")
            return 0.0

    @staticmethod
    def Wrap(Text, Family, SizePt, WidthPt, Bold=False):
        """Greedy word wrap. Returns (lines, widest_word_pt)."""
        if kS.ErrorMode:
            return 0, 0.0
        try:
            Lines, Widest = 0, 0.0
            Space = kMeasure.TextWidth(" ", Family, SizePt, Bold)
            for Para in re.split(r"[\r\n\v]+", Text) or [""]:
                Words = Para.split()
                if not Words:
                    Lines += 1
                    continue
                Cur = 0.0
                Lines += 1
                for W in Words:
                    Ww = kMeasure.TextWidth(W, Family, SizePt, Bold)
                    Widest = max(Widest, Ww)
                    if Cur and Cur + Space + Ww > WidthPt:
                        Lines += 1
                        Cur = Ww
                    else:
                        Cur = Cur + (Space if Cur else 0) + Ww
                    while Cur > WidthPt and WidthPt > 0:  # a word wider than the line breaks inside itself
                        Lines += 1
                        Cur -= WidthPt
            return Lines, Widest
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMeasure.Wrap")
            return 0, 0.0

    @staticmethod
    def TextHeight(Paras, Width):
        """Paras: [(text, family, size_pt, bold, space_before_pt)] -> (height_pt, widest_word_pt, lines)."""
        if kS.ErrorMode:
            return 0.0, 0.0, 0
        try:
            Height, Widest, Total = 0.0, 0.0, 0
            for I, (Text, Family, Size, Bold, Before) in enumerate(Paras):
                N, W = kMeasure.Wrap(Text, Family, Size, Width, Bold)
                Height += N * Size * LINE_HEIGHT + (Before if I else 0)
                Widest = max(Widest, W)
                Total += N
            return Height, Widest, Total
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMeasure.TextHeight")
            return 0.0, 0.0, 0
