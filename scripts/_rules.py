"""Rule helpers shared by audit_deck.py and lint_deck.py (no dependencies)."""
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
