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

# Speaker notes are a spoken script first; this line separates it from the presenter-only reference part
# (assumptions, sources, pitfalls, Q&A) - build_deck writes it, lint_deck counts the script above it.
NOTES_DIVIDER = "--- For the presenter, not to be read out ---"
SCRIPT_MIN_WORDS = 12  # fewer words of prose than this before the divider: the notes have no spoken script


# A success target is measurable when, after any time window ("last 6 months", "by Q3", "in 2027") is set aside,
# a number remains with a unit or a comparator - and the row (or the target) names what is measured.
TIME_WINDOW = re.compile(r"\b(?:(?:the\s+)?(?:last|past|next|first|coming|within|in|over|for|after|before|by|until|"
                         r"from|per|each|every)\s+(?:the\s+)?)(?:\d+(?:[.,]\d+)?\s*)?(?:days?|weeks?|wks?|months?|"
                         r"mos?|quarters?|years?|yrs?|sprints?)\b|\b(?:(?:by|in|until|before|after|from|end of|"
                         r"start of)\s+)?(?:q[1-4]|h[12]|fy\s?\d{2,4}|(?:19|20)\d{2}|\d{1,2}\s+(?:jan|feb|mar|apr|may|"
                         r"jun|jul|aug|sep|oct|nov|dec)[a-z]*|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*"
                         r"(?:\s+\d{1,4})?)\b", re.I)
TARGET_UNIT = re.compile(r"\d\s*(?:%|\u2030|pp\b|pts?\b|x\b|[kKMB]\b|bn\b|ms\b|s\b|sec\w*|min\w*|h\b|hours?\b|"
                         r"[kM]?(?:USD|EUR|SEK|GBP|NOK|DKK|kr)\b|\w{3,})|[$\u20ac\u00a3\u00a5]\s*\d", re.I)
TARGET_COMPARATOR = re.compile(r"[<>\u2264\u2265=\u00b1]|\b(?:below|above|under|over|at least|at most|less than|"
                               r"more than|fewer than|up to|no more than|max(?:imum)?|min(?:imum)?|within|from|to|"
                               r"down|up|reduce\w*|cut|raise|increase\w*|decrease\w*)\b", re.I)
TARGET_FILLER = set("""below above under over least most less more fewer than up to no max maximum min minimum within
from down reduce reduced cut raise increase decrease the a an of and or at by in on per vs target goal level same
current clearly lower higher better improved""".split())
# An ask states its cost: money, people or time, or the word itself.
COST_CUE = re.compile(r"\b(?:costs?|costing|budget\w*|spend\w*|price\w*|invest\w*|funding|funds?|fte|headcount|"
                      r"salar\w*|man-?days?|person-?(?:days?|weeks?|months?))\b|[$\u20ac\u00a3\u00a5]\s*\d|"
                      r"\d\s*(?:[kKM]|bn)?\s*(?:USD|EUR|SEK|GBP|NOK|DKK|kr|kSEK|MSEK|kUSD|MUSD|kEUR|MEUR)\b|"
                      r"\d\s*(?:hours?|days?|weeks?|months?)\b", re.I)


class kRules:
    """Stateless rule helpers shared by lint_deck.py and build_deck.py."""

    @staticmethod
    def ScriptWords(Notes):
        """Words of spoken prose in speaker notes: the text before the presenter-reference divider, without
        labelled blocks ('KEY FACT:', 'SOURCES:'), bullet lines and Q&A lines - what a presenter would say."""
        if kS.ErrorMode:
            return 0
        try:
            Script = (Notes or "").split(NOTES_DIVIDER)[0]
            Count = 0
            for Line in Script.splitlines():
                Line = Line.strip()
                if not Line or re.match(r"^(?:[-*•]\s|Q:|A:)", Line):
                    continue
                Label = re.match(r"^([A-Z][A-Z &/]{2,}):\s*(.*)$", Line)
                if Label:
                    continue  # 'KEY FACT: ...', 'ASSUMPTIONS:' - a fill-in block, not something said
                Count += len(re.findall(r"\S+", Line))
            return Count
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.ScriptWords")
            return 0

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

    @staticmethod
    def WordCount(Text):
        """Visible words in Text: tokens with a letter or a digit (a bullet, a dash or a lone '%' is not a word).
        The one count both build_deck.py --plan and lint_deck.py (--check) judge the word budget by."""
        if kS.ErrorMode:
            return 0
        try:
            return sum(1 for Tok in str(Text).split() if any(Ch.isalnum() for Ch in Tok))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.WordCount")
            return 0

    @staticmethod
    def Measurable(Target, Metric=""):
        """True when a success target can be judged: once time windows are set aside, a number with a unit or a
        comparator is left, and the target or its row's Metric names what is measured ('Churn' + 'below 30 %').
        'Clearly lower' and '<= last 6 months' are not measurable."""
        if kS.ErrorMode:
            return True
        try:
            Rest = TIME_WINDOW.sub(" ", str(Target))
            if not re.search(r"\d", Rest):
                return False
            if re.search(r"\d\s*(?:of|out of|/)\s*\d", Rest):  # a count against a total: '4 of 4', '18/20'
                return True
            if not (TARGET_UNIT.search(Rest) or TARGET_COMPARATOR.search(Rest)):
                return False
            Nouns = [W for W in re.findall(r"[^\W\d_]{3,}", f"{Rest} {Metric}".lower()) if W not in TARGET_FILLER]
            return bool(Nouns) or bool(re.search(r"\d\s*(?:%|\u2030|[$\u20ac\u00a3\u00a5])", Rest))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.Measurable")
            return True

    @staticmethod
    def StatesCost(Text):
        """True when the text of an ask says what it costs: money, people or time, or the word 'cost' itself."""
        if kS.ErrorMode:
            return True
        try:
            return bool(COST_CUE.search(str(Text)))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRules.StatesCost")
            return True
