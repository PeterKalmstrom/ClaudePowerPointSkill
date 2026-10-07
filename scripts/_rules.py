"""Rule helpers shared by lint_deck.py and build_deck.py (no dependencies)."""
import re

VERBS = set("""is are was were be been has have had do does did can will must should may
need needs wins leads beats grows drops rises fails ships pays
comes goes stays decides""".split())

# Office's default chart series colours (Office 2013+ theme); seeing them verbatim usually
# means nobody chose the colours.
OFFICE_DEFAULT_SERIES = {"4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"}


def looks_like_label(title):
    """True when a title reads as a topic label, not a claim: 1-2 words, no common verb, not a question."""
    words = re.findall(r"[^\W\d_]{2,}", title.lower())
    return bool(words) and len(words) <= 2 and "?" not in title and not VERBS & set(words)


def floor_for_room(depth_feet):
    """Smallest readable body size (pt) for a viewing distance."""
    if depth_feet <= 20:
        return 14.0
    if depth_feet <= 30:
        return 18.0
    if depth_feet <= 50:
        return 24.0
    return 28.0


def _lin(c):
    c /= 255.0
    return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4


def contrast_ratio(hex_a, hex_b):
    """WCAG 2.1 contrast ratio between two RRGGBB colours."""
    def lum(h):
        r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
        return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)
    la, lb = sorted((lum(hex_a), lum(hex_b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def overlap_area(a, b):
    """Intersection area of two (left, top, width, height) rectangles."""
    w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
    h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return max(0.0, w) * max(0.0, h)


def contains(outer, inner, tol=1.0):
    return (inner[0] >= outer[0] - tol and inner[1] >= outer[1] - tol
            and inner[0] + inner[2] <= outer[0] + outer[2] + tol
            and inner[1] + inner[3] <= outer[1] + outer[3] + tol)


# ---- taste / accessibility data (thresholds partly follow PointClaw and Impeccable; see NOTICE)

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


def has_emoji(text):
    return any(lo <= ord(ch) <= hi for ch in text for lo, hi in EMOJI_RANGES)


def is_large_text(size_pt, bold):
    """WCAG 2.1 'large text': 18 pt, or 14 pt bold."""
    return size_pt >= 18 or (bold and size_pt >= 14)


def chars_per_line(width_pt, size_pt, inset_pt=7.2):
    """Rough characters per line: average glyph ~0.5 em wide."""
    return max(0.0, width_pt - 2 * inset_pt) / (0.5 * size_pt) if size_pt else 0.0


def tint_ramp(n, r=0.4):
    """Tints for n series of one hue: evenly from -r (darker) to +r (lighter). Keep r <= 0.5:
    below -0.5 turns muddy, above +0.6 washes out."""
    r = min(r, 0.9)
    return [0.0] if n <= 1 else [round(-r + 2 * r * i / (n - 1), 3) for i in range(n)]
