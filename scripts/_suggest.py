"""Ready-to-apply spec edits for what build_deck.py --plan / --check find, and the writer that applies them.

Loops cost more than scripts: an agent that is told "text too long" shortens it by hand, builds, and finds the
next thing. This module turns each finding it can fix without new facts into a concrete edit - the exact new text,
measured - and leaves the rest as questions for the author:

  kShorten      shortens one text until a test passes: filler words and short forms first (nothing lost), then a
                parenthesis, a cut at a clause boundary, then the last words. Never adds a word.
  kSuggest      walks a spec and collects the edits: over-limit texts, titles that wrap past two lines, bullet and
                compare points too long for their designed layout, a quiz answer slide's default title (taken from
                the author's own 'explain'), a summary trend that repeats a detail chart, and slides over their word
                budget (detail to the notes). Whatever still fails afterwards is a question (a cost, a target number,
                supporting points: content only the author has).
  kSpecPatcher  applies edits to the spec file as written (figure tokens kept; a field whose text differs, e.g.
                because it holds a {sum} token, is skipped and reported), writes removed text to notes.moved
                ("MOVED FROM SLIDE:"), and turns build_deck's in-memory auto-fixes into edits too.

Project-specific knowledge (limits, budgets, title measuring, the spec checks) comes in through a context object
that build_deck.py supplies (kSuggestContext), so this module imports nothing from it.
"""
import copy
import json
import os
import re
import shutil

from kShared import kS
from _autofix import BREAKS, kAutoFix

EXTRA_FILLERS = [  # words that carry no fact; matched case-insensitively, on word boundaries
    (r"\b(?:clearly|obviously|just|certainly|definitely|extremely|highly|single|whole)\s+", ""),
    (r"\ball of the\b", "all the"), (r"\bsome of the\b", "some"), (r"\s*\b(?:right now|at the moment)\b", ""),
    (r"\bthat we have\b", ""), (r"\b(?:going|getting) to be\b", "will be"),
]
SHORT_FORMS = [  # same meaning, fewer characters
    (r"\bapproximately\b", "~"), (r"\bfor example\b", "e.g."), (r"\bversus\b", "vs"), (r"\band so on\b", "etc."),
    (r"(\d)\s*minutes?\b", "\\1 min"), (r"(\d)\s*hours?\b", "\\1 h"), (r"(\d)\s*per cent\b", "\\1 %"),
    (r"\bas well as\b", "and"), (r"\bin addition to\b", "besides"), (r"\ba number of\b", "several"),
]
WEAK_END = {"and", "or", "the", "a", "an", "of", "to", "for", "with", "in", "on", "by", "at", "from", "that",
            "which", "so", "but", "our", "your", "its", "their", "is", "are", "before", "after", "as", "into",
            "than", "while", "because", "when", "if", "&", "-"}
PHRASE_START = {"in", "for", "with", "by", "at", "from", "before", "after", "during", "across", "within", "near",
                "so", "which", "who", "while", "because", "when", "if", "since", "until", "unless", "through",
                "via", "using", "without", "whenever", "including", "such"}  # a closing phrase may be dropped here
REVIEW = "cut mid-phrase - check the sense"  # Shorten's 'how' for a cut that is proposed, never applied
NEEDS = [  # (words in a remaining finding, what only the author can supply)
    ("visible words", "fewer words on the slide: cut or merge points (only the author knows which matter)"),
    ("over 80 characters", "shorter points"), ("characters", "a shorter wording"),
    ("cost", "the ask's cost or effort (money, time, FTE) from the brief - or say the figure comes from Finance"),
    ("measurable", "a target number with a unit or comparator"),
    ("sparse", "2-3 short supporting points from the brief"),
    ("one number and one caption", "what the number costs or drives (points), or the share it is"),
    ("falls outside", "a date inside the window, or 'after ...' in the label"),
    ("months jump", "the missing month, or why it is missing (notes)"),
    ("same data", "which slide keeps the chart"),
    ("answer slide's title", "what the answer teaches, as a claim ('answer_title')"),
    ("title", "a shorter claim title"),
    ("speaker notes", "what the presenter says (notes.say)"), ("notes.say", "what the presenter says (notes.say)"),
]
BUDGET_FIELDS = ["mitigation", "detail", "note", "caption", "support", "explain", "rule", "subtitle", "body"]
BUDGET_SLACK = 1.3  # --plan's threshold: a slide over budget x this is a finding; the move stops there
BUDGET_SKIP = {"table", "cost_table", "metrics", "quote", "email", "title", "section"}  # reference slides


class kTextFit:
    """A test a shortened text must pass: at most Limit characters and, with a Measure, Measure.Fits(text)."""

    def __init__(self, Limit, Measure=None):
        try:
            self.Limit = Limit
            self.Measure = Measure
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTextFit.__init__")

    def Fits(self, Text):
        """True when Text passes."""
        if kS.ErrorMode:
            return False
        try:
            if self.Limit and len(str(Text)) > self.Limit:
                return False
            return self.Measure is None or bool(self.Measure.TitleFits(str(Text)))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTextFit.Fits")
            return False


class kShorten:
    """Shortens text by removing or abbreviating words only - never by writing new ones."""

    @staticmethod
    def Squeeze(Text):
        """Text without filler words and with short forms; nothing a reader needs is lost."""
        if kS.ErrorMode:
            return Text
        try:
            New = kAutoFix.Tidy(Text)
            for Pattern, Repl in EXTRA_FILLERS + SHORT_FORMS:
                New = re.sub(Pattern, Repl, New, flags=re.I)
            New = re.sub(r"\s{2,}", " ", New).strip()
            New = re.sub(r"\s+([,.;:])", "\\1", New)
            if New and str(Text)[:1].isupper() and New[:1].islower():
                New = New[0].upper() + New[1:]
            return New
        except Exception as e:
            kS.GlobalErrorHandler(e, "kShorten.Squeeze")
            return Text

    @staticmethod
    def Trim(Text):
        """Text without trailing weak words ('... and the') and trailing punctuation."""
        if kS.ErrorMode:
            return Text
        try:
            Words = str(Text).rstrip(" ,;:-–—").split()
            while Words and Words[-1].lower().strip(",;:") in WEAK_END:
                Words.pop()
            return " ".join(Words).rstrip(" ,;:-–—")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kShorten.Trim")
            return Text

    @staticmethod
    def ClauseCut(Text, Fit, MinWords):
        """The longest prefix of Text that ends at a clause boundary (sentence end, semicolon, dash, colon, comma,
        or before 'with', 'because', 'so that'...), keeps MinWords words and passes Fit; '' when none does."""
        if kS.ErrorMode:
            return ""
        try:
            Best = ""
            for Pattern in BREAKS:
                for Hit in re.finditer(Pattern, Text):
                    Kept = kShorten.Trim(Text[:Hit.start()])
                    if len(Kept.split()) >= MinWords and len(Kept) > len(Best) and Fit.Fits(Kept):
                        Best = Kept
            return Best
        except Exception as e:
            kS.GlobalErrorHandler(e, "kShorten.ClauseCut")
            return ""

    @staticmethod
    def Shorten(Text, Fit, MinWords=2):
        """(new text, words lost, how) - the least change that makes Text pass Fit; (None, False, '') when even
        MinWords words do not. 'Words lost' means more than filler went: the caller keeps the original (notes)."""
        if kS.ErrorMode:
            return None, False, ""
        try:
            Text = str(Text)
            if Fit.Fits(Text):
                return Text, False, ""
            Squeezed = kShorten.Squeeze(Text)
            if Fit.Fits(Squeezed):
                return Squeezed, False, "dropped filler words / used short forms"
            Bare = re.sub(r"\s*\([^)]*\)", "", Squeezed).strip()
            if Bare != Squeezed and len(Bare.split()) >= MinWords and Fit.Fits(Bare):
                return Bare, True, "dropped the parenthesis"
            Cut = kShorten.ClauseCut(Squeezed, Fit, MinWords)
            if Cut:
                return Cut, True, "cut at a clause boundary"
            Words = Squeezed.split()
            for N in range(len(Words) - 1, MinWords - 1, -1):  # before a phrase: '... projects | in every region'
                Kept = kShorten.Trim(" ".join(Words[:N]))
                if Words[N].lower() in PHRASE_START and len(Kept.split()) >= MinWords and Fit.Fits(Kept):
                    return Kept, True, "dropped the closing phrase"
            for N in range(len(Words) - 1, MinWords - 1, -1):  # anywhere: may change the sense, so only proposed
                Kept = kShorten.Trim(" ".join(Words[:N]))
                if len(Kept.split()) >= MinWords and Fit.Fits(Kept):
                    return Kept, True, REVIEW
            return None, False, ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kShorten.Shorten")
            return None, False, ""


class kSuggest:
    """Collects ready-to-apply edits (Patches) and questions for the author (Questions) for one spec."""

    def __init__(self, Spec, Ctx):
        try:
            self.Spec = Spec
            self.Ctx = Ctx
            self.Work = copy.deepcopy(Spec)  # the spec as it would be with every patch so far applied
            self.Patches = []
            self.Questions = []
            self.Proposals = []  # (slide, path, text): a cut that may change the sense - shown, never applied
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.__init__")

    @staticmethod
    def PathText(Parts):
        """'columns[0].points[2]' for ['columns', 0, 'points', 2]."""
        if kS.ErrorMode:
            return ""
        try:
            Out = ""
            for P in Parts:
                Out += f"[{P}]" if isinstance(P, int) else (f".{P}" if Out else str(P))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.PathText")
            return ""

    @staticmethod
    def Walk(Obj, Keys, Trail=None):
        """Every (parts, value) a dotted LIMITS key ('columns.points.*') reaches, as kSpecCheck.Values does, but
        with the path as a list a patch can follow."""
        if kS.ErrorMode:
            return []
        try:
            Trail = Trail or []
            Head, Rest = Keys[0], Keys[1:]
            if Head == "*":
                Items = list(enumerate(Obj)) if isinstance(Obj, list) else []
            else:
                Items = [(Head, Obj.get(Head))] if isinstance(Obj, dict) and Head in Obj else []
            Out = []
            for K, V in Items:
                Here = Trail + [K]
                if not Rest:
                    Out.append((Here, V))
                elif isinstance(V, list) and Rest[0] != "*":
                    for J, El in enumerate(V):
                        Out += kSuggest.Walk(El, Rest, Here + [J])
                else:
                    Out += kSuggest.Walk(V, Rest, Here)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.Walk")
            return []

    @staticmethod
    def Get(Obj, Parts):
        """The value at Parts under Obj, or None."""
        if kS.ErrorMode:
            return None
        try:
            for P in Parts:
                if isinstance(P, int) and isinstance(Obj, list) and 0 <= P < len(Obj):
                    Obj = Obj[P]
                elif isinstance(P, str) and isinstance(Obj, dict):
                    Obj = Obj.get(P)
                else:
                    return None
            return Obj
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.Get")
            return None

    @staticmethod
    def Put(Obj, Parts, Value):
        """Set (or, with Value None, remove) the value at Parts under Obj. True when done."""
        if kS.ErrorMode:
            return False
        try:
            Box = kSuggest.Get(Obj, Parts[:-1]) if len(Parts) > 1 else Obj
            Key = Parts[-1]
            if isinstance(Box, dict) and isinstance(Key, str):
                if Value is None:
                    Box.pop(Key, None)
                else:
                    Box[Key] = Value
                return True
            if isinstance(Box, list) and isinstance(Key, int) and 0 <= Key < len(Box) and Value is not None:
                Box[Key] = Value
                return True
            return False
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.Put")
            return False

    def Add(self, No, Parts, New, Lost, Code, How):
        """Record one edit (merging with an earlier one on the same field, whose original it keeps) and apply it
        to the working copy, so later steps measure the shortened slide."""
        if kS.ErrorMode:
            return
        try:
            Sl = self.Work["slides"][No - 1]
            Old = kSuggest.Get(Sl, Parts)
            Path = kSuggest.PathText(Parts)
            for P in self.Patches:
                if P["slide"] == No and P["path"] == Path:
                    P["new"], P["how"] = New, f"{P['how']}; {How}"
                    P["moved"] = P["moved"] or (P["old"] if Lost else None)
                    kSuggest.Put(Sl, Parts, New)
                    return
            self.Patches.append({"slide": No, "id": Sl.get("id", ""), "parts": list(Parts), "path": Path,
                                 "old": Old, "new": New, "moved": Old if Lost and isinstance(Old, str) else None,
                                 "code": Code, "how": How})
            kSuggest.Put(Sl, Parts, New)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.Add")
            return

    def Merge(self, Patches):
        """Add patches made on top of this result (the dry build's auto-fixes): one on a field already edited
        updates that edit and keeps its original for the notes."""
        if kS.ErrorMode:
            return
        try:
            for New in Patches:
                Same = None
                for P in self.Patches:
                    if P["id"] == New["id"] and P["path"] == New["path"]:
                        Same = P
                if Same is None:
                    self.Patches.append(New)
                else:
                    Same["new"], Same["how"] = New["new"], f"{Same['how']}; {New['how']}"
                    Same["moved"] = Same["moved"] or New.get("moved")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.Merge")
            return

    def ShortenAt(self, No, Parts, Fit, Code):
        """Shorten the text at Parts on slide No until it passes Fit; True when an edit was recorded."""
        if kS.ErrorMode:
            return False
        try:
            Old = kSuggest.Get(self.Work["slides"][No - 1], Parts)
            if not isinstance(Old, str) or Fit.Fits(Old):
                return False
            New, Lost, How = kShorten.Shorten(Old, Fit, 3 if len(Old.split()) > 6 else 2)
            if New is None or New == Old:
                return False
            if How == REVIEW:
                self.Proposals.append((No, kSuggest.PathText(Parts), New))
                return False
            self.Add(No, Parts, New, Lost, Code, How)
            return True
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.ShortenAt")
            return False

    def LimitTexts(self, No, Sl):
        """Every text over its pattern's character limit (a spec error) - shortened to the limit; a title also to
        two lines at the title size."""
        if kS.ErrorMode:
            return
        try:
            Pat = Sl.get("pattern")
            for Key, Lim in self.Ctx.Limits.get(Pat, {}).items():
                if isinstance(Lim, tuple):
                    continue
                for Parts, Val in kSuggest.Walk(Sl, Key.split(".")):
                    if isinstance(Val, str) and len(Val) > Lim:
                        Measure = self.Ctx if Parts == ["title"] and Pat not in self.Ctx.NoChrome else None
                        self.ShortenAt(No, Parts, kTextFit(Lim, Measure), "spec")
            if Pat == "compare":
                Cols = Sl.get("columns", [])
                Lim = self.Ctx.CompareHeading.get(len(Cols), 40)
                for C in range(len(Cols)):
                    self.ShortenAt(No, ["columns", C, "heading"], kTextFit(Lim), "spec")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.LimitTexts")
            return

    def TitleFit(self, No, Sl):
        """A title within its limit that still wraps past two lines at the title size (a plan finding)."""
        if kS.ErrorMode:
            return
        try:
            Title = Sl.get("title")
            if isinstance(Title, str) and Title and Sl.get("pattern") not in self.Ctx.NoChrome \
                    and not self.Ctx.TitleFits(Title):
                Lim = self.Ctx.Limits.get(Sl.get("pattern"), {}).get("title", 90)
                self.ShortenAt(No, ["title"], kTextFit(Lim, self.Ctx), "plan_title")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.TitleFit")
            return

    def LayoutTexts(self, No, Sl):
        """Points too long for their designed layout (spec warnings): a bullets item over the band length, a
        compare point over the card length."""
        if kS.ErrorMode:
            return
        try:
            Pat = Sl.get("pattern")
            if Pat == "bullets" and any(len(str(X)) > self.Ctx.BulletBand for X in Sl.get("items", [])):
                for I in range(len(Sl.get("items", []))):
                    self.ShortenAt(No, ["items", I], kTextFit(self.Ctx.BulletBand), "spec_warning")
            if Pat == "compare":
                for C, Col in enumerate(Sl.get("columns", [])):
                    for P in range(len(Col.get("points", [])) if isinstance(Col, dict) else 0):
                        self.ShortenAt(No, ["columns", C, "points", P], kTextFit(self.Ctx.CompareCard),
                                       "spec_warning")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.LayoutTexts")
            return

    def AnswerTitle(self, No, Sl):
        """A quiz answer slide with the default 'Answer: C' title: the author's own 'explain', first clause,
        shortened to the answer_title limit. Without an 'explain' it stays a question."""
        if kS.ErrorMode:
            return
        try:
            if Sl.get("pattern") != "quiz" or not self.Ctx.LabelAnswer(Sl) or not Sl.get("explain"):
                return
            First = re.split(r"(?<=[.!?])\s+|;\s+", str(Sl["explain"]).strip())[0].rstrip(".!? ")
            Lim = self.Ctx.Limits.get("quiz", {}).get("answer_title", 55)
            New, _, How = kShorten.Shorten(First, kTextFit(Lim), 3)
            if New and How == REVIEW:
                self.Proposals.append((No, "answer_title", New))
            elif New and not self.Ctx.LooksLikeLabel(New):
                self.Add(No, ["answer_title"], New, False, "spec_warning",
                         f"from the slide's own 'explain'" + (f" ({How})" if How else ""))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.AnswerTitle")
            return

    def Budget(self, No, Sl):
        """A slide well over its word budget (--plan's finding): secondary detail moved to the notes, longest first
        (mitigations, step details, captions, support lines - never titles, values, points, items or quiz options),
        until the slide is back under the finding's threshold or nothing more can move."""
        if kS.ErrorMode:
            return
        try:
            Limit = self.Ctx.Budget(Sl) * BUDGET_SLACK
            if Sl.get("pattern") in BUDGET_SKIP or self.Ctx.Words(Sl) <= Limit:
                return
            Trial = copy.deepcopy(Sl)
            for _ in range(8):
                Moved = False
                for Fd in BUDGET_FIELDS:
                    if kAutoFix.MoveDetail(Trial, Fd, []):
                        Moved = True
                        break
                if not Moved or self.Ctx.Words(Trial) <= Limit:
                    break
            Trial.pop("_moved", None)
            for Parts, Old in kSpecPatcher.Leaves(Sl):
                New = kSuggest.Get(Trial, Parts)
                if isinstance(New, str) and New != Old:
                    self.Add(No, Parts, New, True, "word_budget", "detail to the notes (word budget)")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.Budget")
            return

    def RepeatTrends(self):
        """A kpi metric whose trend repeats data another slide charts: drop the trend (the summary states the
        figure; the detail slide keeps the chart)."""
        if kS.ErrorMode:
            return
        try:
            Seen = {}
            for No, Sl in enumerate(self.Work["slides"], 1):
                for Name, Vals in self.Ctx.ChartedSeries(Sl):
                    Seen.setdefault(Vals, []).append((No, Name))
            for Vals, Where in Seen.items():
                if len({No for No, _ in Where}) < 2:
                    continue
                for No, Name in Where:
                    Sl = self.Work["slides"][No - 1]
                    if Sl.get("pattern") != "kpi":
                        continue
                    for M, Mt in enumerate(Sl.get("metrics", [])):
                        Trend = Mt.get("trend") if isinstance(Mt, dict) else None
                        if isinstance(Trend, list) and tuple(round(float(V), 6) for V in Trend) == Vals:
                            if "trend_labels" in Mt:
                                self.Add(No, ["metrics", M, "trend_labels"], None, False, "spec_warning",
                                         "removed with the repeated trend")
                            self.Add(No, ["metrics", M, "trend"], None, False, "spec_warning",
                                     "removed: another slide charts the same numbers")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.RepeatTrends")
            return

    @staticmethod
    def Need(Message):
        """What only the author can supply for a remaining finding."""
        if kS.ErrorMode:
            return ""
        try:
            Low = str(Message).lower()
            for Words, What in NEEDS:
                if Words in Low:
                    return What
            return "an edit by hand, as the message says"
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.Need")
            return ""

    def Run(self):
        """Every edit, then every finding the edits leave as a question. Returns self."""
        if kS.ErrorMode:
            return self
        try:
            for No in range(1, len(self.Work.get("slides", [])) + 1):
                Sl = self.Work["slides"][No - 1]
                if not isinstance(Sl, dict):
                    continue
                self.LimitTexts(No, Sl)
                self.TitleFit(No, Sl)
                self.LayoutTexts(No, Sl)
                self.AnswerTitle(No, Sl)
                self.Budget(No, Sl)
            self.RepeatTrends()
            for Msg in self.Ctx.Remaining(self.Work):
                Hit = re.match(r"slide (\d+)", Msg)
                No = int(Hit.group(1)) if Hit else 0
                Sl = self.Work["slides"][No - 1] if 0 < No <= len(self.Work["slides"]) else {}
                self.Questions.append({"slide": No, "id": Sl.get("id", "") if isinstance(Sl, dict) else "",
                                       "needs": kSuggest.Need(Msg), "message": Msg, "proposals": []})
            self.AttachProposals()
            return self
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.Run")
            return self

    def AttachProposals(self):
        """Each proposed (sense-changing) cut joins the question about its slide and field, or becomes one."""
        if kS.ErrorMode:
            return
        try:
            for No, Path, Text in self.Proposals:
                Key = re.split(r"[.\[]", Path)[0]
                Home = None
                for Q in self.Questions:
                    if Q["slide"] == No and (f"'{Key}" in Q["message"] or f"({Key}" in Q["message"]
                                             or Key in ("items", "points") and "point" in Q["message"]):
                        Home = Q
                        break
                if Home is None:
                    Sl = self.Work["slides"][No - 1]
                    Home = {"slide": No, "id": Sl.get("id", ""), "needs": "a shorter wording",
                            "message": f"slide {No}: '{Path}' needs a shorter wording", "proposals": []}
                    self.Questions.append(Home)
                Home["proposals"].append({"path": Path, "text": Text})
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.AttachProposals")
            return

    @staticmethod
    def Line(Patch):
        """One patch as a report line."""
        if kS.ErrorMode:
            return ""
        try:
            Where = f"slide {Patch['slide']} ({Patch['id']}) {Patch['path']}"
            if Patch["new"] is None:
                return f"{Where}: remove ({Patch['how']})"
            Old = "" if Patch["old"] is None else kAutoFix.Short(Patch["old"], 70)
            Tail = "; the full text goes to the notes (MOVED FROM SLIDE:)" if Patch.get("moved") else ""
            return f"{Where}: '{Old}' -> '{Patch['new']}' ({Patch['how']}{Tail})"
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.Line")
            return ""

    def Print(self, Applied=False):
        """The suggest: / applied: lines, then the question: lines."""
        if kS.ErrorMode:
            return
        try:
            Tag = "applied" if Applied else "suggest"
            for P in self.Patches:
                print(f"{Tag}: {kSuggest.Line(P)}")
            for Q in self.Questions:
                print(f"question: slide {Q['slide']} ({Q['id']}) needs {Q['needs']} - {Q['message']}")
                for Pr in Q.get("proposals", []):
                    print(f"          proposal {Pr['path']}: '{Pr['text']}' (a cut mid-phrase: use it only if the "
                          "sense holds)")
            if self.Patches and not Applied:
                print(f"suggest: {len(self.Patches)} ready-to-apply edit(s) - rerun with --apply to write them into "
                      "the spec (a .before-apply copy is kept) and build in the same call")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.Print")
            return

    def AsJson(self):
        """Patches and questions for the CHECK-JSON / SUGGEST-JSON line."""
        if kS.ErrorMode:
            return {}
        try:
            return {"edits": [{K: V for K, V in P.items() if K != "parts"} for P in self.Patches],
                    "questions": self.Questions}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSuggest.AsJson")
            return {}


class kSpecPatcher:
    """Applies kSuggest patches (and build_deck's auto-fixes) to the spec file as written."""

    Saved = set()  # spec files this process has already backed up

    @staticmethod
    def Load(Path):
        """The spec file as written (no figure tokens filled, no paths resolved)."""
        if kS.ErrorMode:
            return None
        try:
            with open(Path, encoding="utf-8") as Fh:
                if Path.lower().endswith((".yml", ".yaml")):
                    import yaml  # uvx --with pyyaml
                    return yaml.safe_load(Fh)
                return json.load(Fh)
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSpecPatcher.Load(path={Path})")
            return None

    @staticmethod
    def Save(Raw, Path):
        """Write the patched spec over Path, keeping the previous file as <name>.before-apply<ext>. Returns the
        backup's path."""
        if kS.ErrorMode:
            return ""
        try:
            Base, Ext = os.path.splitext(Path)
            Backup = f"{Base}.before-apply{Ext}"
            Key = os.path.abspath(Path)
            if Key not in kSpecPatcher.Saved:  # one backup per call: the file as it was before this call's edits
                shutil.copyfile(Path, Backup)
                kSpecPatcher.Saved.add(Key)
            with open(Path, "w", encoding="utf-8", newline="\n") as Fh:
                if Ext.lower() in (".yml", ".yaml"):
                    import yaml
                    yaml.safe_dump(Raw, Fh, sort_keys=False, allow_unicode=True, width=120)
                else:
                    Fh.write(json.dumps(Raw, ensure_ascii=False, indent=1) + "\n")
            return Backup
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSpecPatcher.Save(path={Path})")
            return ""

    @staticmethod
    def Leaves(Obj, Trail=None):
        """Every (parts, text) string leaf under a slide, notes and ids left out."""
        if kS.ErrorMode:
            return []
        try:
            Trail = Trail or []
            Out = []
            if isinstance(Obj, dict):
                for K, V in Obj.items():
                    if not Trail and K in ("notes", "id", "pattern", "_moved"):
                        continue
                    Out += kSpecPatcher.Leaves(V, Trail + [K])
            elif isinstance(Obj, list):
                for I, V in enumerate(Obj):
                    Out += kSpecPatcher.Leaves(V, Trail + [I])
            elif isinstance(Obj, str):
                Out.append((Trail, Obj))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecPatcher.Leaves")
            return []

    @staticmethod
    def SlideFor(Raw, Patch):
        """The raw slide a patch is for: by id, else by number when the slide counts agree."""
        if kS.ErrorMode:
            return None
        try:
            Slides = Raw.get("slides", [])
            if Patch.get("id"):
                for Sl in Slides:
                    if isinstance(Sl, dict) and Sl.get("id") == Patch["id"]:
                        return Sl
            No = Patch.get("slide", 0)
            Sl = Slides[No - 1] if 0 < No <= len(Slides) else None
            return Sl if isinstance(Sl, dict) and not Sl.get("id") else None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecPatcher.SlideFor")
            return None

    @staticmethod
    def AddMoved(Sl, Path, Text):
        """Keep text removed from the slide in its notes ('moved', written as MOVED FROM SLIDE:)."""
        if kS.ErrorMode:
            return
        try:
            Notes = Sl.get("notes")
            if not isinstance(Notes, dict):
                Notes = {"say": Notes} if isinstance(Notes, str) and Notes.strip() else {}
            Moved = Notes.get("moved") or []
            Moved = [Moved] if isinstance(Moved, str) else list(Moved)
            Line = f"({Path}) {Text}"
            if not any(str(X).startswith(f"({Path}) ") and Text in str(X) for X in Moved):  # a later, shorter cut
                Moved.append(Line)                                                          # is in the first one
            Notes["moved"] = Moved
            Sl["notes"] = Notes
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecPatcher.AddMoved")
            return

    @staticmethod
    def Apply(Raw, Patches):
        """Apply each patch whose field still holds the text it was made from. Returns (applied, skipped), each
        a list of (patch, reason)."""
        if kS.ErrorMode:
            return [], []
        try:
            Applied, Skipped = [], []
            for P in Patches:
                Sl = kSpecPatcher.SlideFor(Raw, P)
                if Sl is None:
                    Skipped.append((P, "slide not found in the spec file (no matching id)"))
                    continue
                Cur = kSuggest.Get(Sl, P["parts"])
                if (P["old"] is None and Cur is not None) or (P["old"] is not None and str(Cur) != str(P["old"])):
                    Skipped.append((P, "the field in the file differs (a figure token?) - edit it by hand"))
                    continue
                if not kSuggest.Put(Sl, P["parts"], P["new"]):
                    Skipped.append((P, "could not set the field"))
                    continue
                if P.get("moved"):
                    kSpecPatcher.AddMoved(Sl, P["path"], P["moved"])
                Applied.append((P, ""))
            return Applied, Skipped
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecPatcher.Apply")
            return [], []

    @staticmethod
    def MovedOf(Sl):
        """{path: original} from the '_moved' lines kAutoFix leaves on a slide."""
        if kS.ErrorMode:
            return {}
        try:
            Out = {}
            for Line in Sl.get("_moved") or []:
                Hit = re.match(r"MOVED FROM SLIDE \(([^)]*)\): (.*)$", str(Line), re.S)
                if Hit:
                    Out.setdefault(Hit.group(1), Hit.group(2))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecPatcher.MovedOf")
            return {}

    @staticmethod
    def AutoPatches(Before, After):
        """build_deck's auto-fixes (Before: the spec as built, After: as kAutoFix left it) as patches. A slide
        added by a split is not written back (it is reported instead)."""
        if kS.ErrorMode:
            return [], []
        try:
            Patches, Splits = [], []
            Old = {Sl.get("id"): Sl for Sl in Before.get("slides", []) if isinstance(Sl, dict) and Sl.get("id")}
            for No, Sl in enumerate(After.get("slides", []), 1):
                Was = Old.get(Sl.get("id")) if isinstance(Sl, dict) else None
                if Was is None:
                    Splits.append(Sl.get("id", f"slide {No}") if isinstance(Sl, dict) else f"slide {No}")
                    continue
                Moved = kSpecPatcher.MovedOf(Sl)
                for Parts, Text in kSpecPatcher.Leaves(Was):
                    New = kSuggest.Get(Sl, Parts)
                    if isinstance(New, str) and New != Text:
                        Path = kSuggest.PathText(Parts)
                        Patches.append({"slide": No, "id": Sl.get("id", ""), "parts": Parts, "path": Path,
                                        "old": Text, "new": New, "moved": Moved.get(Path), "code": "fit",
                                        "how": "auto-fix"})
            return Patches, Splits
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecPatcher.AutoPatches")
            return [], []
