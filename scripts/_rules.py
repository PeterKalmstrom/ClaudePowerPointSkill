"""Rule helpers shared by lint_deck.py and build_deck.py (no dependencies)."""
import re

from kShared import kS

VERBS = set("""is are was were be been has have had do does did can will must should may
need needs wins leads beats grows drops rises fails ships pays
comes goes stays decides""".split())

# Office's default chart series colours (Office 2013+ theme); seeing them verbatim usually
# means nobody chose the colours.
OFFICE_DEFAULT_SERIES = {"4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"}

# ---- taste / accessibility data (thresholds partly follow the author's PowerPoint add-in and Impeccable; see NOTICE)

# Faces that read as "nobody chose a font" when they are the only family in a deck.
DEFAULT_FACES = {"inter", "roboto", "arial", "helvetica", "calibri", "calibri light", "aptos",
                 "aptos display", "system-ui", "segoe ui"}
EMOJI_RANGES = [(0x1F300, 0x1F9FF), (0x2600, 0x27BF), (0x1FA70, 0x1FAFF)]
STOCK_HOSTS = ("unsplash", "shutterstock", "pexels", "istockphoto", "gettyimages", "pixabay", "freepik")
CARTOON_HOSTS = ("undraw", "humaaans", "openpeeps", "drawkit", "icons8")
INSIGHT_WORDS = ("leads", "trails", "beats", "exceeds", "drops", "rises", "grows", "shrinks", "climbs",
                 "falls", "declines", "jumps", "gains", "loses", "outperforms", "underperforms", "doubles",
                 "halves", " up ", " down ", " above ", " below ", " ahead ", " behind ", "%", " vs")
STOP_WORDS = set("""this that with from have will were been they their there what when which about
into than then them these those your more most over also only very just such each other some""".split())
ORDINAL_RE = (r"^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?$|^q[1-4]$|^fy\s?\d{2,4}\s*q[1-4]$"
              r"|^(19|20)\d{2}$|^(mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?$|^h[12]$|^week\s?\d+$")


class kRules:
    """Stateless rule helpers shared by lint_deck.py and build_deck.py."""

    @staticmethod
    def LooksLikeLabel(Title):
        """True when a title reads as a topic label, not a claim: 1-2 words, no common verb, not a question."""
        if kS.ErrorMode:
            return False
        try:
            Words = re.findall(r"[^\W\d_]{2,}", Title.lower())
            return bool(Words) and len(Words) <= 2 and "?" not in Title and not VERBS & set(Words)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.LooksLikeLabel")
            return False

    @staticmethod
    def FloorForRoom(DepthFeet):
        """Smallest readable body size (pt) for a viewing distance."""
        if kS.ErrorMode:
            return None
        try:
            if DepthFeet <= 20:
                return 14.0
            if DepthFeet <= 30:
                return 18.0
            if DepthFeet <= 50:
                return 24.0
            return 28.0
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.FloorForRoom")
            return None

    @staticmethod
    def Linear(Channel):
        """sRGB channel (0-255) to linear light."""
        if kS.ErrorMode:
            return 0.0
        try:
            Channel /= 255.0
            return Channel / 12.92 if Channel <= 0.03928 else ((Channel + 0.055) / 1.055) ** 2.4
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.Linear")
            return 0.0

    @staticmethod
    def Luminance(Hex):
        """WCAG relative luminance of an RRGGBB colour."""
        if kS.ErrorMode:
            return 0.0
        try:
            R, G, B = (int(Hex[I:I + 2], 16) for I in (0, 2, 4))
            return 0.2126 * kRules.Linear(R) + 0.7152 * kRules.Linear(G) + 0.0722 * kRules.Linear(B)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.Luminance")
            return 0.0

    @staticmethod
    def ContrastRatio(HexA, HexB):
        """WCAG 2.1 contrast ratio between two RRGGBB colours."""
        if kS.ErrorMode:
            return None
        try:
            La, Lb = sorted((kRules.Luminance(HexA), kRules.Luminance(HexB)), reverse=True)
            return (La + 0.05) / (Lb + 0.05)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.ContrastRatio")
            return None

    @staticmethod
    def OverlapArea(A, B):
        """Intersection area of two (left, top, width, height) rectangles."""
        if kS.ErrorMode:
            return 0.0
        try:
            W = min(A[0] + A[2], B[0] + B[2]) - max(A[0], B[0])
            H = min(A[1] + A[3], B[1] + B[3]) - max(A[1], B[1])
            return max(0.0, W) * max(0.0, H)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.OverlapArea")
            return 0.0

    @staticmethod
    def Contains(Outer, Inner, Tol=1.0):
        """True when Inner lies inside Outer (within Tol pt)."""
        if kS.ErrorMode:
            return False
        try:
            return (Inner[0] >= Outer[0] - Tol and Inner[1] >= Outer[1] - Tol
                    and Inner[0] + Inner[2] <= Outer[0] + Outer[2] + Tol
                    and Inner[1] + Inner[3] <= Outer[1] + Outer[3] + Tol)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.Contains")
            return False

    @staticmethod
    def HasEmoji(Text):
        """True when Text contains an emoji."""
        if kS.ErrorMode:
            return False
        try:
            return any(Lo <= ord(Ch) <= Hi for Ch in Text for Lo, Hi in EMOJI_RANGES)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.HasEmoji")
            return False

    @staticmethod
    def IsLargeText(SizePt, Bold):
        """WCAG 2.1 'large text': 18 pt, or 14 pt bold."""
        if kS.ErrorMode:
            return False
        try:
            return SizePt >= 18 or (Bold and SizePt >= 14)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.IsLargeText")
            return False

    @staticmethod
    def CharsPerLine(WidthPt, SizePt, InsetPt=7.2):
        """Rough characters per line: average glyph ~0.5 em wide."""
        if kS.ErrorMode:
            return 0.0
        try:
            return max(0.0, WidthPt - 2 * InsetPt) / (0.5 * SizePt) if SizePt else 0.0
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.CharsPerLine")
            return 0.0

    @staticmethod
    def TintRamp(N, R=0.4):
        """Tints for N series of one hue: evenly from -R (darker) to +R (lighter). Keep R <= 0.5:
        below -0.5 turns muddy, above +0.6 washes out."""
        if kS.ErrorMode:
            return []
        try:
            R = min(R, 0.9)
            return [0.0] if N <= 1 else [round(-R + 2 * R * I / (N - 1), 3) for I in range(N)]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.TintRamp")
            return []
