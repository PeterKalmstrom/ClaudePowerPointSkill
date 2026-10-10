"""What a slide says, read from an existing deck - shared by extract_spec.py and improve_deck.py (python-pptx).

kSlideContent turns a slide into paragraph items (text, size, bold, box in pt) plus its charts, tables and
pictures, and sorts out the furniture: slide numbers, footers, number badges ("1", "!") and small source lines.
kNotesDraft writes a spoken-script DRAFT from that content only - the title and the slide's own words and
figures, never anything that is not on the slide - for a presenter to review.
A second consumer (another reader of foreign decks) can use kSlideContent.Read as is; the thresholds are the
module constants below.
"""
import re

from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.util import Emu

from kShared import kS
from lint_deck import GROUP_TF, GROWN, kLintDeck

BADGE_RE = re.compile(r"^(?:\d{1,2}\.?|[!?•*✓✗×])$")
FIGURE_RE = re.compile(r"\d")
SOURCE_MAX_PT = 12.0
DRAFT_MARK = ("DRAFT script written from the slide's own text by improve_deck.py / extract_spec.py - check every "
              "sentence and add the figures' sources before presenting.")
FURNITURE_PH = (PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE)


class kSlideContent:
    """Stateless reader: one slide as {title, items, charts, tables, pictures, footer, sources, notes}."""

    @staticmethod
    def Clean(Text):
        """A paragraph's text on one line: soft breaks and runs of spaces become one space."""
        if kS.ErrorMode:
            return ""
        try:
            return re.sub(r"\s+", " ", (Text or "").replace("\x0b", " ")).strip()
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideContent.Clean")
            return ""

    @staticmethod
    def ReadingKey(Item):
        """Sort key: rows of 18 pt top to bottom, then left to right, then paragraph order in its box."""
        if kS.ErrorMode:
            return (0, 0, 0)
        try:
            return (round(Item["t"] / 18), round(Item["l"]), Item["seq"])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideContent.ReadingKey")
            return (0, 0, 0)

    @staticmethod
    def ParaItems(Sh, Box):
        """Every non-empty paragraph of a text shape as an item dict."""
        if kS.ErrorMode:
            return []
        try:
            L, T, W, H = Box
            Out = []
            for Seq, P in enumerate(Sh.text_frame.paragraphs):
                Text = kSlideContent.Clean(P.text)
                if not Text:
                    continue
                Bold = any(R.font.bold for R in P.runs if R.text.strip())
                Out.append({"text": Text, "size": kLintDeck.ParaSize(Sh, P) or 18.0, "bold": bool(Bold),
                            "l": L, "t": T + Seq * 0.01, "w": W, "h": H, "seq": Seq, "shape": Sh.name,
                            "box": Sh.shape_id, "title": kLintDeck.IsTitle(Sh), "ph": kLintDeck.PhType(Sh)})
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideContent.ParaItems")
            return []

    @staticmethod
    def ChartData(Sh):
        """A chart's type, categories and series."""
        if kS.ErrorMode:
            return None
        try:
            Ch = Sh.chart
            Plot = Ch.plots[0] if len(Ch.plots) else None
            Title = Ch.chart_title.text_frame.text.strip() if Ch.has_title and Ch.chart_title.has_text_frame else ""
            return {"type": str(Ch.chart_type).split(".")[-1].split(" ")[0].lower(), "title": Title,
                    "categories": [str(C) for C in Plot.categories] if Plot is not None else [],
                    "series": [{"name": S.name or "", "values": [V for V in S.values]} for S in Plot.series]
                    if Plot is not None else [], "box": kLintDeck.Rect(Sh)}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideContent.ChartData")
            return None

    @staticmethod
    def IsFurniture(Item, Number, SlideH):
        """True for a slide number, footer or date: not content."""
        if kS.ErrorMode:
            return False
        try:
            if Item["ph"] in FURNITURE_PH:
                return True
            return Item["text"] == str(Number) and Item["size"] <= 16 and Item["t"] > SlideH * 0.75
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideContent.IsFurniture")
            return False

    @staticmethod
    def IsBadge(Item, Others):
        """True for a number or symbol badge ('1', '2.', '!') no larger than the text around it - the builder
        numbers items itself. A large number is a figure, not a badge."""
        if kS.ErrorMode:
            return False
        try:
            if not BADGE_RE.match(Item["text"]):
                return False
            Sizes = sorted(X["size"] for X in Others if X is not Item and not BADGE_RE.match(X["text"]))
            Median = Sizes[len(Sizes) // 2] if Sizes else 18.0
            return Item["size"] <= Median * 1.3
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideContent.IsBadge")
            return False

    @staticmethod
    def LooksLikeFigure(Text):
        """True for a short text with a digit in it ('5.9 MUSD', '41 %'): a figure, not a title."""
        if kS.ErrorMode:
            return False
        try:
            return len(Text) <= 14 and bool(FIGURE_RE.search(Text))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideContent.LooksLikeFigure")
            return False

    @staticmethod
    def PickTitle(Items, Number, SlideH):
        """The title item: the title placeholder, else the largest text in the top 30 % (anywhere on slide 1)."""
        if kS.ErrorMode:
            return None
        try:
            Ph = [X for X in Items if X["title"]]
            if Ph:
                return Ph[0]
            Pool = [X for X in Items if X["size"] >= 24 and (Number == 1 or X["t"] < SlideH * 0.3)]
            Words = [X for X in Pool if not kSlideContent.LooksLikeFigure(X["text"])]
            Pool = Words or Pool
            if not Pool:
                return None
            if Number == 1:  # the cover: its largest text
                Best = max(X["size"] for X in Pool)
                return sorted([X for X in Pool if X["size"] == Best], key=kSlideContent.ReadingKey)[0]
            return sorted(Pool, key=kSlideContent.ReadingKey)[0]  # a content slide: the topmost large text
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideContent.PickTitle")
            return None

    @staticmethod
    def Shapes(Slide, Out):
        """Fill Out's raw items, charts, tables and pictures from the slide's shapes (groups included)."""
        if kS.ErrorMode:
            return None
        try:
            GROUP_TF.clear()
            GROWN.clear()
            for Sh in kLintDeck.Walk(Slide.shapes):
                if Sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                    continue
                Box = kLintDeck.Rect(Sh)
                if getattr(Sh, "has_chart", False) and Sh.has_chart:
                    Out["charts"].append(kSlideContent.ChartData(Sh))
                elif getattr(Sh, "has_table", False) and Sh.has_table:
                    Out["tables"].append([[kSlideContent.Clean(C.text) for C in Row.cells] for Row in Sh.table.rows])
                elif Sh.shape_type == MSO_SHAPE_TYPE.PICTURE and not kLintDeck.IsDecorative(Sh):
                    Out["pictures"].append({"shape": Sh, "box": Box, "alt": kLintDeck.AltText(Sh)})
                elif Sh.has_text_frame and Sh.text_frame.text.strip():
                    Out["raw"] += kSlideContent.ParaItems(Sh, Box)
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideContent.Shapes")
            return None

    @staticmethod
    def Read(Slide, Number, SlideW, SlideH):
        """One slide: title (text or ''), content items in reading order, charts, tables, pictures, the footer
        text, small source lines, badges dropped, and the speaker notes."""
        if kS.ErrorMode:
            return None
        try:
            Out = {"raw": [], "charts": [], "tables": [], "pictures": [], "footer": "", "sources": [],
                   "badges": [], "items": [], "title": "", "title_size": 0.0, "number": Number,
                   "notes": kSlideContent.Clean(Slide.notes_slide.notes_text_frame.text)
                   if Slide.has_notes_slide else "", "notes_raw": Slide.notes_slide.notes_text_frame.text
                   if Slide.has_notes_slide else "", "w": SlideW, "h": SlideH}
            kSlideContent.Shapes(Slide, Out)
            Content = []
            for X in Out["raw"]:
                if kSlideContent.IsFurniture(X, Number, SlideH):
                    if X["ph"] == PP_PLACEHOLDER.FOOTER and not Out["footer"]:
                        Out["footer"] = X["text"]
                    continue
                Content.append(X)
            Title = kSlideContent.PickTitle(Content, Number, SlideH)
            if Title is not None:
                Out["title_size"] = Title["size"]
                Rest = [X for X in Content if X["box"] != Title["box"]]
                Out["title"] = " ".join(X["text"] for X in Content if X["box"] == Title["box"])
            else:
                Rest = Content
            for X in sorted(Rest, key=kSlideContent.ReadingKey):
                if kSlideContent.IsBadge(X, Rest):
                    Out["badges"].append(X["text"])
                elif X["size"] <= SOURCE_MAX_PT and X["t"] > SlideH * 0.78 and len(X["text"].split()) >= 2:
                    Out["sources"].append(X["text"])
                else:
                    Out["items"].append(X)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSlideContent.Read(slide={Number})")
            return None

    @staticmethod
    def ReadDeck(Prs):
        """Every slide of a Presentation through Read, hidden slides included (marked)."""
        if kS.ErrorMode:
            return []
        try:
            W, H = Emu(Prs.slide_width).pt, Emu(Prs.slide_height).pt
            Out = []
            for N, Slide in enumerate(Prs.slides, 1):
                Info = kSlideContent.Read(Slide, N, W, H)
                if Info is None:
                    return []
                Info["hidden"] = Slide._element.get("show") == "0"
                Out.append(Info)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideContent.ReadDeck")
            return []


class kNotesDraft:
    """Stateless: a spoken-script draft from a slide's own words. Nothing is added that the slide does not say."""

    @staticmethod
    def Sentence(Text):
        """Text as a sentence: capital first letter, closing full stop."""
        if kS.ErrorMode:
            return ""
        try:
            Text = Text.strip().rstrip(";,:")
            if not Text:
                return ""
            Text = Text[0].upper() + Text[1:]
            return Text if Text[-1] in ".!?\"”'" else Text + "."
        except Exception as e:
            kS.GlobalErrorHandler(e, "kNotesDraft.Sentence")
            return ""

    @staticmethod
    def LabelUnder(Figure, Items, Used):
        """The text right under a figure in its own column (its label), marked used; None when there is none."""
        if kS.ErrorMode:
            return None
        try:
            Best = None
            for X in Items:
                if X is Figure or id(X) in Used or kSlideContent.LooksLikeFigure(X["text"]):
                    continue
                Centre = X["l"] + X["w"] / 2
                Under = X["seq"] > Figure["seq"] if X["box"] == Figure["box"] else X["t"] > Figure["t"]
                if Under and Figure["l"] - 24 <= Centre <= Figure["l"] + max(Figure["w"], 150) + 24 \
                        and X["w"] <= max(Figure["w"], 150) * 2.5 and (Best is None or X["t"] < Best["t"]):
                    Best = X
            if Best is not None:
                Used.add(id(Best))
            return Best
        except Exception as e:
            kS.GlobalErrorHandler(e, "kNotesDraft.LabelUnder")
            return None

    @staticmethod
    def ChartSentence(Chart):
        """The chart's figures, read out: 'Series: A 4.1, B 4.6.'"""
        if kS.ErrorMode:
            return ""
        try:
            Parts = []
            for S in Chart["series"][:3]:
                Pairs = ", ".join(f"{C} {V:g}" for C, V in zip(Chart["categories"], S["values"])
                                  if isinstance(V, (int, float)))
                Parts.append(f"{S['name']}: {Pairs}" if S["name"] else Pairs)
            return kNotesDraft.Sentence("The chart shows " + "; ".join(Parts)) if Parts else ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kNotesDraft.ChartSentence")
            return ""

    @staticmethod
    def Script(Info, MaxSentences=5):
        """2-5 sentences from Info (kSlideContent.Read): the title as the point, then the slide's own lines -
        short phrases grouped into one 'Key points' sentence - then its chart and table figures."""
        if kS.ErrorMode:
            return []
        try:
            Out = [kNotesDraft.Sentence(Info["title"])] if Info["title"] else []
            Short, Lines, Used = [], [], set()
            for X in Info["items"]:
                if id(X) in Used:
                    continue
                Label = kNotesDraft.LabelUnder(X, Info["items"], Used) \
                    if kSlideContent.LooksLikeFigure(X["text"]) else None
                if Label is not None:
                    Lines.append(kNotesDraft.Sentence(f"{Label['text'].rstrip('.')}: {X['text']}"))
                elif len(X["text"].split()) < 5 and not FIGURE_RE.search(X["text"]):
                    Short.append(X["text"].rstrip("."))
                else:
                    Lines.append(kNotesDraft.Sentence(X["text"]))
            if Short:
                Out.append(kNotesDraft.Sentence("Key points: " + ", ".join(Short)))
            Out += Lines[:max(0, MaxSentences - len(Out))]
            for C in Info["charts"]:
                if len(Out) < MaxSentences + 1:
                    Out.append(kNotesDraft.ChartSentence(C))
            for T in Info["tables"]:
                if T and len(Out) < MaxSentences + 1:
                    Out.append(kNotesDraft.Sentence(f"The table compares {', '.join(T[0])} across {len(T) - 1} rows"))
            return [S for S in Out if S]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kNotesDraft.Script")
            return []
