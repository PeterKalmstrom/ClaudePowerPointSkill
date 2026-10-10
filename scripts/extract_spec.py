"""Turn an existing deck into a best-effort build_deck.py spec, so it can be restyled by rebuilding - any OS.

    python scripts/extract_spec.py old.pptx --out spec.json              # spec + pictures in spec-media/
    python scripts/extract_spec.py old.pptx --out spec.json --no-draft-notes
    python scripts/build_deck.py spec.json --out new.pptx --check        # rebuild in the skill's design
    python scripts/extract_spec.py old.pptx --compare new.pptx           # every word and figure still there?

Each slide is mapped to the pattern that fits what it shows: a chart (kpi_chart when large figures sit beside
it), a table, a picture, the cover, big figures with their labels (big_number / kpi), a decision slide
(statement with the ask), heading + detail pairs (compare, process, a timeline when the headings are dates),
else bullets or a statement. Nothing is dropped: text a pattern has no room for, or that is over a field's
limit, goes to the notes as "MOVED FROM SLIDE: ..." and is listed as a `flag:` line. Slide numbers, footers
and number badges ("1", "!") are left to the builder, which draws its own. Existing notes become the spoken
script verbatim; a slide without notes - or with notes of under 12 words of prose, which are then kept at the
end - gets a DRAFT script written from its own words (marked for review, never with a fact the slide does not
show) unless --no-draft-notes.
The spec is a starting point: read the flags, edit the spec (claim titles, the right pattern), then build.
"""
import argparse
import json
import os
import re
import sys

from pptx import Presentation

from build_deck import LIMITS
from kShared import ToolInputException, ToolReportableException, kRun, kS, kToolException
from _deck_content import DRAFT_MARK, FIGURE_RE, kNotesDraft, kSlideContent
from _rules import SCRIPT_MIN_WORDS, kRules

ORDINAL_RE = re.compile(r"^\(?\d{1,2}[.)]\s+")
NUMBER_RE = re.compile(r"^([+\-−–]?[$€£]?\d[\d.,]*(?:\s*/\s*\d+)?)\s*(.*)$")
DATE_RE = re.compile(r"^(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?(?:\s+\d{2,4})?$|^(?:q[1-4]|h[12])"
                     r"(?:\s+\d{2,4})?$|^(?:19|20)\d{2}$|^(?:week|month|day|phase)\s+\d+$", re.I)
DECISION_RE = re.compile(r"\b(?:decision|approve|approval|the ask|ask of|we ask)\b", re.I)
CHART_TYPES = {"bar": "bar", "column": "column", "line": "line", "pie": "pie", "doughnut": "pie", "area": "line"}
SPLITTERS = (". ", "; ", " - ", " – ", ": ", " (", ", ")


class kSpecMapper:
    """Maps kSlideContent slides onto build_deck patterns; collects flags and the text moved to the notes."""

    def __init__(self, MediaDir, MediaRel, Draft):
        try:
            self.MediaDir, self.MediaRel, self.Draft = MediaDir, MediaRel, Draft
            self.Flags = []
            self.Moved = []
            self.Number = 0
            self.Ids = set()
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.__init__")

    def Flag(self, Message):
        """Record one flag for the current slide."""
        if kS.ErrorMode:
            return None
        try:
            self.Flags.append(f"slide {self.Number}: {Message}")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.Flag")
            return None

    def Move(self, Text, Why):
        """Send slide text to the notes (MOVED FROM SLIDE) and flag it."""
        if kS.ErrorMode:
            return None
        try:
            if Text and Text.strip():
                self.Moved.append(Text.strip())
                self.Flag(f"moved to the notes ({Why}): {Text.strip()[:70]}")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.Move")
            return None

    def Fit(self, Text, Limit, Field):
        """Text within Limit characters: cut at the latest sentence or clause break (else a word break) that
        fits; the rest goes to the notes in full. Nothing is lost."""
        if kS.ErrorMode:
            return ""
        try:
            Text = (Text or "").strip()
            if len(Text) <= Limit:
                return Text
            Cut = -1
            for Sep in SPLITTERS:
                At = Text.rfind(Sep, 0, Limit + 1)
                if At >= Limit * 0.35:
                    Cut = At + (1 if Sep[0] in ".;:," else 0)
                    break
            if Cut < 0:
                Cut = Text.rfind(" ", 0, Limit + 1)
            if Cut <= 0:
                Cut = Limit
            Head = Text[:Cut].rstrip(" ,;:-–(")
            self.Move(Text, f"{Field} is {len(Text)} characters, max {Limit}; the slide keeps '{Head[:40]}'")
            return Head
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.Fit")
            return ""

    @staticmethod
    def Plain(Text):
        """Text without a leading ordinal ('1. ', '2) ') - the builder numbers items itself."""
        if kS.ErrorMode:
            return ""
        try:
            return ORDINAL_RE.sub("", Text or "").strip()
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.Plain")
            return ""

    @staticmethod
    def Median(Items):
        """Median paragraph size of the items (18 when there are none)."""
        if kS.ErrorMode:
            return 18.0
        try:
            Sizes = sorted(X["size"] for X in Items)
            return Sizes[len(Sizes) // 2] if Sizes else 18.0
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.Median")
            return 18.0

    @staticmethod
    def IsFigure(Item, Median):
        """A large short figure: '5.9 MUSD', '41 %', '9 in 10'."""
        if kS.ErrorMode:
            return False
        try:
            return len(Item["text"]) <= 14 and bool(FIGURE_RE.search(Item["text"])) and (
                Item["size"] >= Median * 1.4 or Item["size"] >= 32)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.IsFigure")
            return False

    @staticmethod
    def InColumn(Top, Item, Reach=160.0):
        """True when Item sits under Top in Top's own column (not a full-width line below a row of columns)."""
        if kS.ErrorMode:
            return False
        try:
            if Item["box"] == Top["box"]:
                return Item["seq"] > Top["seq"]
            Width = max(Top["w"], 150.0)
            Centre = Item["l"] + Item["w"] / 2
            return (Item["t"] > Top["t"] and Item["t"] - (Top["t"] + Top["h"]) <= Reach
                    and Top["l"] - 24 <= Centre <= Top["l"] + Width + 24 and Item["w"] <= Width * 2.5)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.InColumn")
            return False

    @staticmethod
    def TopKey(Item):
        """Sort key for items going down a column."""
        if kS.ErrorMode:
            return (0.0, 0)
        try:
            return (Item["t"], Item["seq"])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.TopKey")
            return (0.0, 0)

    def Below(self, Top, Items, Used, Count):
        """Up to Count unused items under Top in its column, nearest first; marks them used."""
        if kS.ErrorMode:
            return []
        try:
            Pool = sorted([X for X in Items if id(X) not in Used and X is not Top and self.InColumn(Top, X)],
                          key=kSpecMapper.TopKey)[:Count]
            for X in Pool:
                Used.add(id(X))
            return Pool
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.Below")
            return []

    def Metrics(self, Figures, Items, Used, Pattern):
        """One metric per figure: value, its label below, a note under that."""
        if kS.ErrorMode:
            return []
        try:
            Lim = LIMITS[Pattern]
            Out = []
            for F in Figures:
                Used.add(id(F))
            for F in Figures:
                Under = self.Below(F, Items, Used, 2)
                Metric = {"value": self.Fit(F["text"], Lim["metrics.value"], "metric value"),
                          "label": self.Fit(Under[0]["text"], Lim["metrics.label"], "metric label") if Under else ""}
                if len(Under) > 1:
                    Metric["note"] = self.Fit(Under[1]["text"], Lim["metrics.note"], "metric note")
                if not Metric["label"]:
                    self.Flag(f"metric '{F['text']}' has no label under it: write one in the spec")
                    Metric["label"] = "(label)"
                Out.append(Metric)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.Metrics")
            return []

    def Leftover(self, Items, Used, Why):
        """Move every item not used by the pattern to the notes."""
        if kS.ErrorMode:
            return None
        try:
            for X in Items:
                if id(X) not in Used:
                    self.Move(X["text"], Why)
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.Leftover")
            return None

    def ChartSlide(self, Info, Slide):
        """chart, or kpi_chart when large figures sit beside the chart."""
        if kS.ErrorMode:
            return None
        try:
            Chart = Info["charts"][0]
            for Extra in Info["charts"][1:]:
                self.Move(kNotesDraft.ChartSentence(Extra), "a second chart: one chart per slide")
            for Table in Info["tables"]:
                self.Move("\n".join(" | ".join(Row) for Row in Table), "a table beside the chart: split the slide")
            Items, Used = Info["items"], set()
            Median = self.Median(Items)
            Figures = [X for X in Items if self.IsFigure(X, Median)][:4]
            Slide["pattern"] = "kpi_chart" if Figures else "chart"
            Kind = CHART_TYPES.get(Chart["type"].split("_")[0], "column")
            if Kind == "column" and not Chart["type"].startswith("column"):
                self.Flag(f"chart type '{Chart['type']}' drawn as a column chart")
            Slide.update({"type": Kind})
            Slide.update(self.ChartSeries(Chart, LIMITS[Slide["pattern"]]["series"][1]))
            if Figures:
                Slide["metrics"] = self.Metrics(Figures, Items, Used, "kpi_chart")
            Rest = [X for X in Items if id(X) not in Used]
            Caption = Rest[0]["text"] if Rest else Chart["title"]
            if Rest:
                Used.add(id(Rest[0]))
            if Caption:
                Slide["caption"] = self.Fit(Caption, LIMITS["kpi_chart"]["caption"], "caption")
            if Rest and Chart["title"]:
                self.Move(Chart["title"], "the chart's own title")
            self.Leftover(Items, Used, "no room beside the chart")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.ChartSlide")
            return None

    def ChartSeries(self, Chart, MaxSeries):
        """{categories, series} for the spec; a one-category chart (bars side by side, one per series) is
        turned so each series becomes a category - the builder needs at least two."""
        if kS.ErrorMode:
            return {}
        try:
            Series = []
            for S in Chart["series"]:
                Values = [V if isinstance(V, (int, float)) else 0 for V in S["values"]]
                if any(not isinstance(V, (int, float)) for V in S["values"]):
                    self.Flag(f"series '{S['name']}' has empty points, written as 0")
                Series.append({"name": S["name"] or Chart["title"] or "Series", "values": Values})
            Cats = list(Chart["categories"])
            if len(Cats) < 2 and len(Series) >= 2:
                self.Flag("one-category chart turned: its series are now the categories")
                Name = Cats[0] if Cats else (Chart["title"] or "Value")
                return {"categories": [S["name"] for S in Series],
                        "series": [{"name": Name, "values": [S["values"][0] if S["values"] else 0 for S in Series]}]}
            for S in Series[MaxSeries:]:
                self.Move(f"{S['name']}: " + ", ".join(f"{C} {V:g}" for C, V in zip(Cats, S["values"])),
                          f"a series past the {MaxSeries}th")
            return {"categories": Cats, "series": Series[:MaxSeries]}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.ChartSeries")
            return {}

    def TableSlide(self, Info, Slide):
        """table: header and up to 8 rows of up to 6 columns; the rest to the notes."""
        if kS.ErrorMode:
            return None
        try:
            Table = Info["tables"][0]
            Width = LIMITS["table"]["header"][1]
            Slide.update({"pattern": "table", "header": Table[0][:Width],
                          "rows": [Row[:Width] for Row in Table[1:LIMITS["table"]["rows"][1] + 1]]})
            for Row in Table:
                if len(Row) > Width:
                    self.Move(" | ".join(Row[Width:]), "a column past the sixth")
            for Row in Table[LIMITS["table"]["rows"][1] + 1:]:
                self.Move(" | ".join(Row), "a row past the eighth")
            for Extra in Info["tables"][1:]:
                self.Move(" | ".join(" / ".join(Row) for Row in Extra), "a second table")
            if len(Table) < 2:
                self.Flag("the table has only a header row")
            self.Leftover(Info["items"], set(), "no room beside the table")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.TableSlide")
            return None

    def ImageSlide(self, Info, Slide, Picture):
        """image: the picture saved next to the spec, a caption from the first line of text."""
        if kS.ErrorMode:
            return None
        try:
            os.makedirs(self.MediaDir, exist_ok=True)
            Name = f"s{self.Number:02d}.{Picture['shape'].image.ext}"
            with open(os.path.join(self.MediaDir, Name), "wb") as File:
                File.write(Picture["shape"].image.blob)
            Slide.update({"pattern": "image", "image": f"{self.MediaRel}/{Name}"})
            if Picture["alt"]:
                Slide["alt"] = Picture["alt"]
            else:
                self.Flag("the picture has no alt text: write 'alt' in the spec")
            Items = Info["items"]
            if Items:
                Slide["caption"] = self.Fit(Items[0]["text"], LIMITS["image"]["caption"], "caption")
            self.Leftover(Items[1:], set(), "no room beside the picture")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.ImageSlide")
            return None

    def CoverSlide(self, Info, Slide):
        """title: the cover, its other lines as the subtitle."""
        if kS.ErrorMode:
            return None
        try:
            Slide["pattern"] = "title"
            Lines = [X["text"] for X in Info["items"]]
            if Lines:
                Slide["subtitle"] = self.Fit(" | ".join(Lines), LIMITS["title"]["subtitle"], "subtitle")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.CoverSlide")
            return None

    def FigureSlide(self, Info, Slide, Figures):
        """big_number for one figure, kpi for two to six."""
        if kS.ErrorMode:
            return None
        try:
            Items, Used = Info["items"], set()
            if len(Figures) == 1:
                F = Figures[0]
                Used.add(id(F))
                Match = NUMBER_RE.match(F["text"])
                Number, Unit = (Match.group(1), Match.group(2)) if Match else (F["text"], "")
                if len(Unit) > LIMITS["big_number"]["unit"]:
                    Number, Unit = F["text"], ""
                Slide.update({"pattern": "big_number", "number": self.Fit(Number, 12, "number")})
                if Unit:
                    Slide["unit"] = Unit
                Under = self.Below(F, Items, Used, 1)
                Rest = Under + [X for X in Items if id(X) not in Used]
                if Rest:
                    Used.add(id(Rest[0]))
                    Slide["caption"] = self.Fit(Rest[0]["text"], LIMITS["big_number"]["caption"], "caption")
                Points = [X for X in Items if id(X) not in Used][:3]
                if Points:
                    Slide["points"] = [self.Fit(X["text"], LIMITS["big_number"]["points.*"], "point") for X in Points]
                    Used.update(id(X) for X in Points)
            else:
                Slide.update({"pattern": "kpi", "metrics": self.Metrics(Figures[:6], Items, Used, "kpi")})
            self.Leftover(Items, Used, "no room beside the figures")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.FigureSlide")
            return None

    def DecisionSlide(self, Info, Slide):
        """statement with the ask as the decision and the next lines as its reasons."""
        if kS.ErrorMode:
            return None
        try:
            Lim = LIMITS["statement"]
            Items = Info["items"]
            Slide.update({"pattern": "statement", "decision": self.Fit(self.Plain(Items[0]["text"]), Lim["decision"],
                                                                       "decision")})
            if len(Items) > 1:
                Slide["points"] = [self.Fit(self.Plain(X["text"]), Lim["points.*"], "point") for X in Items[1:4]]
            for X in Items[4:]:
                self.Move(X["text"], "a statement has up to three points")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.DecisionSlide")
            return None

    def IsHeading(self, Item, Items):
        """A short line over a detail in its own column, set apart by size or weight."""
        if kS.ErrorMode:
            return False
        try:
            if len(Item["text"]) > 40:
                return False
            Under = sorted([X for X in Items if X is not Item and self.InColumn(Item, X, 60.0)], key=kSpecMapper.TopKey)
            if not Under:
                return False
            D = Under[0]
            return Item["size"] >= D["size"] * 1.1 or (Item["bold"] and not D["bold"]) or bool(
                DATE_RE.match(Item["text"]))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.IsHeading")
            return False

    def Pairs(self, Items):
        """[(heading item, [detail items])] for every heading, details being the items under it in its column
        down to the next heading."""
        if kS.ErrorMode:
            return []
        try:
            Heads = [X for X in Items if self.IsHeading(X, Items)]
            Out = []
            for H in Heads:
                Details = []
                for X in sorted([X for X in Items if X not in Heads and self.InColumn(H, X)], key=kSpecMapper.TopKey):
                    Nearest = [O for O in Heads if O is not H and self.InColumn(O, X) and O["t"] > H["t"]]
                    if not Nearest:
                        Details.append(X)
                Out.append((H, Details))
            return [P for P in Out if P[1]]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.Pairs")
            return []

    def DateColumns(self, Items):
        """[(date item, [items under it])] when three to seven dates ('Jan', 'Q1', '2026') each head a column
        of text, else []."""
        if kS.ErrorMode:
            return []
        try:
            Dates = [X for X in Items if DATE_RE.match(X["text"])]
            if not 3 <= len(Dates) <= 7:
                return []
            Used = set(id(X) for X in Dates)
            Out = [(D, self.Below(D, Items, Used, 2)) for D in Dates]
            return Out if all(Under for _, Under in Out) else []
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.DateColumns")
            return []

    def TimelineSlide(self, Info, Slide, Events):
        """timeline: each date with the first line under it; a second line goes to the notes."""
        if kS.ErrorMode:
            return None
        try:
            Lim = LIMITS["timeline"]
            Used = set()
            Out = []
            for D, Under in Events:
                Used.update([id(D)] + [id(X) for X in Under])
                Out.append({"date": self.Fit(D["text"], Lim["events.date"], "date"),
                            "label": self.Fit(Under[0]["text"], Lim["events.label"], "event label")})
                for X in Under[1:]:
                    self.Move(f"{D['text']}: {X['text']}", "an event has one label")
            Slide.update({"pattern": "timeline", "events": Out})
            self.Leftover(Info["items"], Used, "outside the timeline")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.TimelineSlide")
            return None

    def PairSlide(self, Info, Slide, Pairs):
        """timeline (date headings), compare (two or three), process (a row of four or more), else bullets."""
        if kS.ErrorMode:
            return None
        try:
            Used = set(id(H) for H, _ in Pairs) | set(id(D) for _, Ds in Pairs for D in Ds)
            Row = len(set(round(H["t"] / 24) for H, _ in Pairs)) == 1
            if len(Pairs) >= 3 and all(DATE_RE.match(H["text"]) for H, _ in Pairs):
                Lim = LIMITS["timeline"]
                Slide.update({"pattern": "timeline", "events": [
                    {"date": self.Fit(H["text"], Lim["events.date"], "date"),
                     "label": self.Fit(Ds[0]["text"], Lim["events.label"], "event label")} for H, Ds in Pairs[:7]]})
                for H, Ds in Pairs:
                    for D in Ds[1:]:
                        self.Move(f"{H['text']}: {D['text']}", "an event has one label")
            elif len(Pairs) <= 3:
                Lim = LIMITS["compare"]
                Head = 24 if len(Pairs) == 3 else Lim["columns.heading"]
                Slide.update({"pattern": "compare", "columns": [
                    {"heading": self.Fit(self.Plain(H["text"]), Head, "heading"),
                     "points": [self.Fit(D["text"], Lim["columns.points.*"], "point") for D in Ds[:4]]}
                    for H, Ds in Pairs]})
            elif Row and len(Pairs) <= 8:
                Lim = LIMITS["process"]
                Slide.update({"pattern": "process", "steps": [
                    {"label": self.Fit(self.Plain(H["text"]), Lim["steps.label"], "step label"),
                     "detail": self.Fit(" ".join(D["text"] for D in Ds), Lim["steps.detail"], "step detail")}
                    for H, Ds in Pairs]})
            else:
                Slide.update({"pattern": "bullets", "items": [
                    self.Fit(f"{self.Plain(H['text'])}: {' '.join(D['text'] for D in Ds)}", 100, "item")
                    for H, Ds in Pairs[:7]]})
            self.Leftover(Info["items"], Used, "outside the heading + detail layout")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.PairSlide")
            return None

    def TextSlide(self, Info, Slide):
        """statement for one line (or none), bullets for more."""
        if kS.ErrorMode:
            return None
        try:
            Items = Info["items"]
            if len(Items) <= 1:
                Slide["pattern"] = "statement"
                if Items:
                    Slide["support"] = self.Fit(Items[0]["text"], LIMITS["statement"]["support"], "support")
                elif not Info["tables"] and not Info["charts"]:
                    self.Flag("no content found on the slide (shapes without text?): check the render")
                return None
            Slide.update({"pattern": "bullets",
                          "items": [self.Fit(self.Plain(X["text"]), 100, "item") for X in Items[:7]]})
            for X in Items[7:]:
                self.Move(X["text"], "a list has up to seven items")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.TextSlide")
            return None

    def SlideId(self, Title):
        """A unique kebab-case id from the title's first words."""
        if kS.ErrorMode:
            return ""
        try:
            Base = "-".join(re.findall(r"[a-z0-9]+", Title.lower())[:4]) or f"slide-{self.Number}"
            Id, K = Base, 2
            while Id in self.Ids:
                Id, K = f"{Base}-{K}", K + 1
            self.Ids.add(Id)
            return Id
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.SlideId")
            return ""

    def Notes(self, Info, Slide):
        """Existing notes verbatim as the script; else a DRAFT from the slide's own words; moved text after."""
        if kS.ErrorMode:
            return None
        try:
            Notes = {}
            Old = Info["notes_raw"].strip()
            if Old and (kRules.ScriptWords(Old) >= SCRIPT_MIN_WORDS or not self.Draft):
                Notes["say"] = Old
            elif self.Draft:
                Notes["say"] = kNotesDraft.Script(Info) + ([Old] if Old else [])
                Notes["pitfalls"] = [DRAFT_MARK]
                self.Flag(("notes without a script ('" + Old[:30] + "' kept at the end)" if Old else "no notes")
                          + ": a DRAFT script was written from the slide text - review it")
            else:
                self.Flag("no notes: write notes.say")
            if Info["sources"]:
                Notes["sources"] = Info["sources"]
            if self.Moved:
                Notes["facts"] = [f"MOVED FROM SLIDE: {T}" for T in self.Moved]
            if Notes:
                Slide["notes"] = Notes
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.Notes")
            return None

    def Route(self, Info, Slide):
        """Pick the pattern for one slide and fill it."""
        if kS.ErrorMode:
            return None
        try:
            Area = Info["w"] * Info["h"]
            Pictures = sorted([P for P in Info["pictures"] if P["box"][2] * P["box"][3] >= Area * 0.08],
                              key=kSpecMapper.PictureArea, reverse=True)
            Median = self.Median(Info["items"])
            Figures = [X for X in Info["items"] if self.IsFigure(X, Median)]
            if Info["charts"]:
                return self.ChartSlide(Info, Slide)
            if Info["tables"]:
                return self.TableSlide(Info, Slide)
            if Pictures:
                return self.ImageSlide(Info, Slide, Pictures[0])
            if self.Number == 1:
                return self.CoverSlide(Info, Slide)
            if Figures:
                return self.FigureSlide(Info, Slide, Figures)
            if DECISION_RE.search(Info["title"]) and Info["items"]:
                return self.DecisionSlide(Info, Slide)
            Events = self.DateColumns(Info["items"])
            if Events:
                return self.TimelineSlide(Info, Slide, Events)
            Pairs = self.Pairs(Info["items"])
            if len(Pairs) >= 2:
                return self.PairSlide(Info, Slide, Pairs)
            return self.TextSlide(Info, Slide)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.Route")
            return None

    @staticmethod
    def PictureArea(Picture):
        """Sort key: a picture's area."""
        if kS.ErrorMode:
            return 0.0
        try:
            return Picture["box"][2] * Picture["box"][3]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecMapper.PictureArea")
            return 0.0

    def Map(self, Info):
        """One slide dict for the spec."""
        if kS.ErrorMode:
            return None
        try:
            self.Number, self.Moved = Info["number"], []
            if not Info["title"] and Info["items"]:
                First = Info["items"].pop(0)
                Info["title"] = First["text"]
                self.Flag(f"no title found; the first line '{First['text'][:50]}' is the title")
            Slide = {"id": "", "pattern": "", "title": ""}
            self.Route(Info, Slide)
            Limit = LIMITS[Slide["pattern"]].get("title", 70)
            Slide["title"] = self.Fit(Info["title"] or "(title)", Limit, "title")
            Slide["id"] = self.SlideId(Slide["title"])
            if Info["hidden"]:
                self.Flag("hidden in the source deck - delete it from the spec if it should stay out")
            self.Notes(Info, Slide)
            return Slide
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSpecMapper.Map(slide={Info.get('number')})")
            return None


class kExtractSpec:
    """Stateless: a deck file to a spec dict (and its flags)."""

    @staticmethod
    def Extract(Path, SpecPath, Draft=True, Direction=None):
        """(spec, flags) for the deck at Path; pictures go to <spec name>-media/ beside SpecPath."""
        if kS.ErrorMode:
            return None, []
        try:
            Prs = Presentation(Path)
            Stem = os.path.splitext(os.path.basename(SpecPath))[0]
            Mapper = kSpecMapper(os.path.join(os.path.dirname(os.path.abspath(SpecPath)), f"{Stem}-media"),
                                 f"{Stem}-media", Draft)
            Slides = kSlideContent.ReadDeck(Prs)
            Spec = {"$schema": "scripts/spec.schema.json"}
            if Direction:
                Spec["direction"] = Direction
            Footer = next((S["footer"] for S in Slides if S["footer"]), "")
            if Footer:
                Spec["footer"] = Footer[:70]
            Spec["slides"] = [Mapper.Map(S) for S in Slides]
            if kS.ErrorMode:
                return None, []
            return Spec, Mapper.Flags
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kExtractSpec.Extract(file={Path})")
            return None, []


class kTextCompare:
    """Stateless: is every word and figure of a source deck still in a rebuilt one (slides or notes)?"""

    @staticmethod
    def Tokens(Text):
        """Lower-case words and numbers of Text, leading ordinals ('1. ') dropped."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for Line in (Text or "").splitlines():
                Line = ORDINAL_RE.sub("", Line.strip()).lower().replace("−", "-")
                Out += re.findall(r"\d+(?:[.,]\d+)*|[^\W\d_]+", Line)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTextCompare.Tokens")
            return []

    @staticmethod
    def DeckTokens(Path, Furniture):
        """(words, figures) of a deck: titles, text, tables, chart labels and values, notes. Furniture (slide
        numbers, footers, badges) counts only when Furniture is True (the rebuilt side)."""
        if kS.ErrorMode:
            return set(), set()
        try:
            Words = set()
            Figures = set()
            for S in kSlideContent.ReadDeck(Presentation(Path)):
                Texts = [S["title"], S["notes_raw"]] + [X["text"] for X in S["items"]] + S["sources"]
                Texts += [C for Table in S["tables"] for Row in Table for C in Row]
                if Furniture:
                    Texts += S["badges"] + [S["footer"]] + [X["text"] for X in S["raw"]]
                for C in S["charts"]:
                    Texts += C["categories"] + [X["name"] for X in C["series"]] + [C["title"]]
                    Figures.update(f"{V:g}" for X in C["series"] for V in X["values"] if isinstance(V, (int, float)))
                for T in Texts:
                    Words.update(kTextCompare.Tokens(T))
            Figures.update(W for W in Words if W[0].isdigit())
            return Words, Figures
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kTextCompare.DeckTokens(file={Path})")
            return set(), set()

    @staticmethod
    def Compare(Source, Rebuilt):
        """(missing words, missing figures), each sorted."""
        if kS.ErrorMode:
            return [], []
        try:
            SrcWords, SrcFigs = kTextCompare.DeckTokens(Source, False)
            NewWords, NewFigs = kTextCompare.DeckTokens(Rebuilt, True)
            return sorted(SrcWords - NewWords - SrcFigs), sorted(SrcFigs - NewFigs - NewWords)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTextCompare.Compare")
            return [], []


class kExtractSpecApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
            Ap.add_argument("file")
            Ap.add_argument("--out", help="spec .json to write (pictures go to <name>-media/ beside it)")
            Ap.add_argument("--compare", metavar="REBUILT", help="check every word and figure of FILE is in REBUILT")
            Ap.add_argument("--no-draft-notes", action="store_true", help="leave slides without notes unscripted")
            Ap.add_argument("--direction", help="a build_deck direction to write into the spec")
            A = Ap.parse_args()
            if not os.path.isfile(A.file):
                raise ToolReportableException(f"file not found: {A.file}")
            if A.compare:
                return self.CompareRun(A.file, A.compare)
            if not A.out:
                raise ToolInputException("give --out spec.json (or --compare rebuilt.pptx)")
            Spec, Flags = kExtractSpec.Extract(A.file, A.out, not A.no_draft_notes, A.direction)
            if Spec is None:
                return 1
            with open(A.out, "w", encoding="utf-8") as File:
                json.dump(Spec, File, ensure_ascii=False, indent=2)
            for Line in Flags:
                print(f"flag: {Line}")
            Kinds = ", ".join(f"{S['id']}={S['pattern']}" for S in Spec["slides"])
            print(f"\n{len(Spec['slides'])} slides -> {A.out}  ({len(Flags)} flag(s))\n{Kinds}")
            print("next: edit the spec (claim titles, patterns, the flags above), then "
                  f"build_deck.py {A.out} --out new.pptx --check, then extract_spec.py {A.file} --compare new.pptx")
            return 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kExtractSpecApp.Run")
            return 1

    def CompareRun(self, Source, Rebuilt):
        """Print what of Source is missing from Rebuilt; exit 1 when anything is."""
        if kS.ErrorMode:
            return 1
        try:
            if not os.path.isfile(Rebuilt):
                raise ToolReportableException(f"file not found: {Rebuilt}")
            Words, Figures = kTextCompare.Compare(Source, Rebuilt)
            if kS.ErrorMode:
                return 1
            print(f"missing figures ({len(Figures)}): {', '.join(Figures) or '-'}")
            print(f"missing words ({len(Words)}): {', '.join(Words) or '-'}")
            print("COMPARE: all text and figures kept" if not Words and not Figures else "COMPARE: something was lost")
            sys.stdout.flush()
            return 0 if not Words and not Figures else 1
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kExtractSpecApp.CompareRun")
            return 1


if __name__ == "__main__":
    kRun.Main(kExtractSpecApp)
