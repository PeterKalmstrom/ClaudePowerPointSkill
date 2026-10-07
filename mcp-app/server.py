"""PowerPoint Live - an MCP App that drives desktop PowerPoint (Windows, COM) and shows the deck live.

    uvx --with "mcp<2" --with python-pptx --with pillow --with "pywin32; sys_platform == 'win32'" python mcp-app/server.py

Hosts that support MCP Apps (Claude Desktop, claude.ai) show the live view; elsewhere the tools still work.

Tools for the model (the view uses them too):
  powerpoint_open      open a deck (or attach to an open one) and pick the current slide
  powerpoint_run       run Python against the deck: app, prs, slide, goto(n), win32com in scope
  powerpoint_show      make slide n current
  powerpoint_move      move slides (by SlideID) before a slide or to the end of a section
  powerpoint_sections  add, rename or delete sections
  powerpoint_hide      hide or unhide slides
  powerpoint_fix       apply a lint fix to one shape, live
  powerpoint_history   list saved versions; powerpoint_restore puts one back
  powerpoint_resume    re-arm the server after an unexpected error halted it
Tools only the view calls: slide_state, slide_image, deck_outline, slide_thumbs, deck_lint, history_thumb.

Every change made through these tools is preceded by a saved version, so it can be undone from the History view.

Errors follow scripts/kShared.py: expected states (no deck open, file not found, not Windows) come back as tool
errors with a plain message; the first UNEXPECTED error is reported to kS.GlobalErrorHandler and halts the server.
While halted every tool answers "PowerPoint Live halted ..." until a person presses Resume (powerpoint_resume).
"""
import asyncio
import base64
import contextlib
import io
import os
import shutil
import sys
import tempfile
import time
import traceback
from concurrent.futures import ThreadPoolExecutor

from mcp.server.fastmcp import FastMCP

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(HERE), "scripts")
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)

from kShared import ToolInputException, ToolReportableException, kErrorReport, kS, kToolException  # noqa: E402

kErrorReport.HoldForView = True  # a server never exits: each error report waits for the person's Yes/No in the view

URI = "ui://powerpoint-live/slide-view"
MIME = "text/html;profile=mcp-app"
UI = {"ui": {"resourceUri": URI}, "ui/resourceUri": URI}
APP_ONLY = {"ui": {"resourceUri": URI, "visibility": ["app"]}, "ui/resourceUri": URI}
RAMP = [0, 0.35, -0.3, 0.6, -0.5, 0.75]
LINE_CHARTS = {4, 63, 64, 65, 66, 67}
FIXABLE = {"unused_placeholder", "body_below_floor", "text_overflow", "picture_stretched", "a11y_missing_alt_text",
           "chart_default_palette", "chart_legend_steals_plot", "chart_accounting_zero_dash"}
MAX_VERSIONS = 25
NO_TEXT_TYPES = {7, 10, 12, 16}   # msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoMedia
GOTO_VIEWS = {1, 3, 9}            # ppViewSlide, ppViewNotesPage, ppViewNormal: the views that can GotoSlide
SECTION_ACTIONS = ("add", "rename", "delete")

mcp = FastMCP("powerpoint-live")


# ---------------------------------------------------------------- shape maps and change highlights

class kChangeTracker:
    """What each slide looked like when last seen, and which shapes changed since."""

    def __init__(self, Live):
        try:
            self.Live = Live
            self.Seen = {}      # SlideID -> shape map last seen
            self.Changes = {}   # SlideID -> {"stamp", "boxes"}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kChangeTracker.__init__")

    def Clear(self):
        """Forget everything (another deck, or a restore)."""
        if kS.ErrorMode:
            return None
        try:
            self.Seen.clear()
            self.Changes.clear()
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kChangeTracker.Clear")
            return None

    def ShapeText(self, Shape):
        """The text of a shape, or "" for shapes that cannot hold text."""
        if kS.ErrorMode:
            return ""
        try:
            if Shape.Type in NO_TEXT_TYPES or not Shape.HasTextFrame:
                return ""
            return Shape.TextFrame.TextRange.Text if Shape.TextFrame.HasText else ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kChangeTracker.ShapeText")
            return ""

    def ShapeMap(self, Slide):
        """Top-level shapes of a slide: id -> (name, box, text hash)."""
        if kS.ErrorMode:
            return {}
        try:
            Out = {}
            for Shape in Slide.Shapes:
                Text = self.ShapeText(Shape)
                Out[Shape.Id] = {"name": Shape.Name,
                                 "box": [round(Shape.Left, 1), round(Shape.Top, 1), round(Shape.Width, 1),
                                         round(Shape.Height, 1)], "text": hash(Text), "snippet": Text[:80]}
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kChangeTracker.ShapeMap")
            return {}

    def Signature(self, Prs, N):
        """Cheap fingerprint of slide n. Changes when anyone edits the slide."""
        if kS.ErrorMode:
            return "halted"
        try:
            if not N:
                return "empty"
            Slide = Prs.Slides(N)
            Map = self.ShapeMap(Slide)
            return str(hash((Slide.SlideID, bool(Slide.SlideShowTransition.Hidden),
                             tuple((K, tuple(V["box"]), V["text"]) for K, V in Map.items()))))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kChangeTracker.Signature")
            return "halted"

    def Diff(self, Before, After):
        """Boxes of the shapes added, changed or removed between two shape maps."""
        if kS.ErrorMode:
            return []
        try:
            Boxes = []
            for K, V in After.items():
                Old = Before.get(K)
                if Old is None:
                    Boxes.append({"id": K, "name": V["name"], "kind": "added", "box": V["box"]})
                elif Old["box"] != V["box"] or Old["text"] != V["text"]:
                    Boxes.append({"id": K, "name": V["name"], "kind": "changed", "box": V["box"]})
            for K, V in Before.items():
                if K not in After:
                    Boxes.append({"id": K, "name": V["name"], "kind": "removed", "box": V["box"]})
            return Boxes
        except Exception as e:
            kS.GlobalErrorHandler(e, "kChangeTracker.Diff")
            return []

    def NoteChanges(self, Slide, Before=None):
        """Compare a slide with what was last seen (or `Before`) and remember what changed."""
        if kS.ErrorMode:
            return {}
        try:
            After = self.ShapeMap(Slide)
            Prev = Before if Before is not None else self.Seen.get(Slide.SlideID)
            if Prev is not None:
                Boxes = self.Diff(Prev, After)
                if Boxes:
                    self.Changes[Slide.SlideID] = {"stamp": time.time(), "boxes": Boxes}
            self.Seen[Slide.SlideID] = After
            return After
        except Exception as e:
            kS.GlobalErrorHandler(e, "kChangeTracker.NoteChanges")
            return {}

    def NoteNewSlide(self, Prs, Slide):
        """A slide that did not exist before a change: highlight the whole slide."""
        if kS.ErrorMode:
            return None
        try:
            self.Changes[Slide.SlideID] = {"stamp": time.time(), "boxes": [{
                "id": 0, "name": "new slide", "kind": "added",
                "box": [0, 0, Prs.PageSetup.SlideWidth, Prs.PageSetup.SlideHeight]}]}
            self.Seen[Slide.SlideID] = self.ShapeMap(Slide)
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kChangeTracker.NoteNewSlide")
            return None


# ---------------------------------------------------------------- versions (undo)

class kVersionHistory:
    """Copies of the deck saved before every change, newest last."""

    def __init__(self, Live):
        try:
            self.Live = Live
            self.Versions = []   # [{"v", "time", "label", "path", "thumb", "slide", "count"}]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVersionHistory.__init__")

    def Find(self, V):
        """The version entry numbered V, or None."""
        if kS.ErrorMode:
            return None
        try:
            for Entry in self.Versions:
                if Entry["v"] == int(V):
                    return Entry
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVersionHistory.Find")
            return None

    def Snapshot(self, Label):
        """Save a copy of the deck as a version before a change. SaveCopyAs never rebinds the open deck."""
        if kS.ErrorMode:
            return 0
        try:
            Live = self.Live
            Prs = Live.Prs()
            V = (self.Versions[-1]["v"] + 1) if self.Versions else 1
            Ext = os.path.splitext(Prs.FullName)[1] or ".pptx"
            Path = os.path.join(Live.Work, f"v{V:03d}{Ext}")
            Prs.SaveCopyAs(Path)
            N = Live.Clamp(Prs, Live.Current or 1)
            Thumb = Live.Png(Prs.Slides(N), 240) if N else None
            self.Versions.append({"v": V, "time": time.time(), "label": Label, "path": Path, "thumb": Thumb,
                                  "slide": N, "count": Prs.Slides.Count})
            while len(self.Versions) > MAX_VERSIONS:
                Old = self.Versions.pop(0)
                if os.path.exists(Old["path"]):
                    os.remove(Old["path"])
            return V
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVersionHistory.Snapshot")
            return 0

    def List(self):
        """The versions, newest first, without paths or pictures."""
        if kS.ErrorMode:
            return {"versions": []}
        try:
            return {"versions": [{K: H[K] for K in ("v", "time", "label", "slide", "count")}
                                 for H in reversed(self.Versions)]}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVersionHistory.List")
            return {"versions": []}

    def Thumb(self, V):
        """The thumbnail saved with version V."""
        if kS.ErrorMode:
            return {"v": 0, "png": None}
        try:
            Entry = self.Find(V)
            return {"v": int(V), "png": Entry["thumb"] if Entry else None}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVersionHistory.Thumb")
            return {"v": 0, "png": None}

    def Restore(self, V):
        """Replace the deck file with version V and reopen it; the state left behind is saved first."""
        if kS.ErrorMode:
            return None
        try:
            Live = self.Live
            Entry = self.Find(V)
            if Entry is None:
                raise ToolInputException(f"no version {V}")
            Prs = Live.Prs()
            Path = Prs.FullName
            if not os.path.isabs(Path):
                raise ToolReportableException("the deck has never been saved; save it once, then versions can be "
                                              "restored")
            self.Snapshot(f"before restoring v{V}")
            WithWindow = Prs.Windows.Count > 0
            Prs.Close()                  # COM Close discards unsaved changes without a prompt: they are in the version
            shutil.copyfile(Entry["path"], Path)
            New = Live.App().Presentations.Open(Path, 0, 0, -1 if WithWindow else 0)
            Live.PrsName = New.FullName
            Live.Current = Live.Clamp(New, Entry["slide"] or 1)
            Live.Tracker.Seen.clear()
            Live.Version += 1
            Live.SyncWindow(New)
            return {**Live.Summary(New), "restored": int(V)}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVersionHistory.Restore")
            return None


# ---------------------------------------------------------------- changes (handler classes, one per kind)

class kMutation:
    """One change to the deck. kPowerPointLive.Mutate saves a version, calls Apply, then records what changed."""

    def __init__(self, Live, Label):
        try:
            self.Live = Live
            self.Label = Label
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMutation.__init__")

    def Apply(self, Prs):
        """Make the change; return extra keys for the result."""
        if kS.ErrorMode:
            return {}
        try:
            return {}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMutation.Apply")
            return {}


class kRunMutation(kMutation):
    """powerpoint_run: the user's (Claude's) Python against the deck."""

    def __init__(self, Live, Code, Slide, Label):
        try:
            super().__init__(Live, Label or "powerpoint_run")
            self.Code = Code
            self.Slide = Slide
            self.Prs = None
            self.Output = io.StringIO()
            self.Error = None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRunMutation.__init__")

    def Goto(self, N):
        """goto(n) in the user's code: change the current slide."""
        if kS.ErrorMode:
            return None
        try:
            self.Live.Current = self.Live.Clamp(self.Prs, N)
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRunMutation.Goto")
            return None

    def Apply(self, Prs):
        """Run the code with app, prs, slide, goto and win32com in scope; print() output is kept."""
        if kS.ErrorMode:
            return {}
        try:
            Live = self.Live
            self.Prs = Prs
            if self.Slide:
                Live.Current = Live.Clamp(Prs, self.Slide)
            import win32com.client
            Scope = {"app": Live.App(), "prs": Prs, "slide": Prs.Slides(Live.Current) if Live.Current else None,
                     "goto": self.Goto, "win32com": win32com}
            with contextlib.redirect_stdout(self.Output):
                try:
                    exec(compile(self.Code, "<powerpoint_run>", "exec"), Scope)
                except Exception:  # ERROR-SUPPRESSED-JUSTIFIED: the user's code failing is the tool's result, reported back to Claude
                    self.Error = traceback.format_exc(limit=3)
            return {}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRunMutation.Apply")
            return {}


class kMoveMutation(kMutation):
    """powerpoint_move: slides before a slide, to the end of a section, or to the end of the deck."""

    def __init__(self, Live, SlideIds, BeforeId, Section):
        try:
            super().__init__(Live, "move slides")
            self.SlideIds = [int(I) for I in SlideIds]
            self.BeforeId = int(BeforeId or 0)
            self.Section = int(Section or 0)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMoveMutation.__init__")

    def Apply(self, Prs):
        """Move each slide in deck order, so their order is kept; the current slide stays current."""
        if kS.ErrorMode:
            return {}
        try:
            Live = self.Live
            Sorter = Live.Sorter
            Cur = Prs.Slides(Live.Current).SlideID if Live.Current else None
            Ordered = sorted((Sorter.SlideById(Prs, I).SlideIndex, I) for I in self.SlideIds)
            Ids = [I for _, I in Ordered if I != self.BeforeId]
            for Sid in Ids:
                Sorter.MoveOne(Prs, Sorter.SlideById(Prs, Sid), self.BeforeId, self.Section)
            if Cur is not None:
                Live.Current = Sorter.SlideById(Prs, Cur).SlideIndex
            return {"moved": Ids}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMoveMutation.Apply")
            return {}


class kSectionMutation(kMutation):
    """powerpoint_sections: add, rename or delete a section."""

    def __init__(self, Live, Action, Section, Name, BeforeId):
        try:
            super().__init__(Live, f"section {Action}")
            self.Action = Action
            self.Section = int(Section or 0)
            self.Name = Name
            self.BeforeId = int(BeforeId or 0)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSectionMutation.__init__")

    def Apply(self, Prs):
        """Edit the sections; deleting keeps the slides (they join the section before)."""
        if kS.ErrorMode:
            return {}
        try:
            Sp = Prs.SectionProperties
            if self.Action == "add":
                Idx = self.Live.Sorter.SlideById(Prs, self.BeforeId).SlideIndex if self.BeforeId \
                    else Prs.Slides.Count + 1
                if Idx > Prs.Slides.Count:
                    Sp.AddSection(Sp.Count + 1, self.Name or "New section")
                else:
                    Sp.AddBeforeSlide(Idx, self.Name or "New section")
            elif self.Action == "rename":
                Sp.Rename(self.Section, self.Name)
            else:
                Sp.Delete(self.Section, False)
            return {}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSectionMutation.Apply")
            return {}


class kHideMutation(kMutation):
    """powerpoint_hide: hide or unhide slides in the slide show."""

    def __init__(self, Live, SlideIds, Hidden):
        try:
            super().__init__(Live, "hide slides" if Hidden else "unhide slides")
            self.SlideIds = [int(I) for I in SlideIds]
            self.Hidden = Hidden
        except Exception as e:
            kS.GlobalErrorHandler(e, "kHideMutation.__init__")

    def Apply(self, Prs):
        """Set SlideShowTransition.Hidden on each slide."""
        if kS.ErrorMode:
            return {}
        try:
            for Sid in self.SlideIds:
                self.Live.Sorter.SlideById(Prs, Sid).SlideShowTransition.Hidden = -1 if self.Hidden else 0
            return {}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kHideMutation.Apply")
            return {}


class kFixMutation(kMutation):
    """powerpoint_fix: the automatic fix for one lint finding on one shape."""

    def __init__(self, Live, SlideId, ShapeId, Code, Floor):
        try:
            super().__init__(Live, f"fix {Code}")
            self.SlideId = int(SlideId)
            self.ShapeId = int(ShapeId)
            self.Code = Code
            self.Floor = Floor
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFixMutation.__init__")

    def Apply(self, Prs):
        """Find the shape (it may be gone) and fix it; its slide becomes current."""
        if kS.ErrorMode:
            return {}
        try:
            Live = self.Live
            Slide = Live.Sorter.SlideById(Prs, self.SlideId)
            Shape = Live.FindShape(Slide, ShapeId=self.ShapeId)
            if Shape is None:
                raise ToolReportableException("that shape is gone - re-run the lint")
            Done = Live.Lint.FixShape(Shape, self.Code, self.Floor)
            Live.Current = Slide.SlideIndex
            return {"fixed": self.Code, "result": Done}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFixMutation.Apply")
            return {}


# ---------------------------------------------------------------- sorter, outline, thumbnails

class kSorter:
    """Outline, sections, thumbnails and slide moves."""

    def __init__(self, Live):
        try:
            self.Live = Live
            self.Thumbs = {}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSorter.__init__")

    def SlidesById(self, Prs):
        """SlideID -> slide, read by walking the deck (FindBySlideID raises for a deleted slide)."""
        if kS.ErrorMode:
            return {}
        try:
            return {Prs.Slides(I).SlideID: Prs.Slides(I) for I in range(1, Prs.Slides.Count + 1)}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSorter.SlidesById")
            return {}

    def SlideById(self, Prs, SlideId):
        """The slide with this SlideID; a missing one is the caller's mistake."""
        if kS.ErrorMode:
            return None
        try:
            Slide = self.SlidesById(Prs).get(int(SlideId))
            if Slide is None:
                raise ToolInputException(f"no slide with SlideID {SlideId} - read deck_outline again")
            return Slide
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSorter.SlideById")
            return None

    def Title(self, Slide):
        """The text of the slide's title placeholder, or ""."""
        if kS.ErrorMode:
            return ""
        try:
            for Shape in Slide.Shapes:
                if Shape.Type == 14 and Shape.PlaceholderFormat.Type in (1, 3) and Shape.HasTextFrame \
                        and Shape.TextFrame.HasText:  # msoPlaceholder: title, ctrTitle
                    return Shape.TextFrame.TextRange.Text.strip()
            return ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSorter.Title")
            return ""

    def NotesText(self, Slide):
        """The speaker notes, or "" when the slide has no notes body."""
        if kS.ErrorMode:
            return ""
        try:
            if not Slide.HasNotesPage:
                return ""
            Placeholders = Slide.NotesPage.Shapes.Placeholders
            if Placeholders.Count < 2 or not Placeholders(2).HasTextFrame:
                return ""
            return Placeholders(2).TextFrame.TextRange.Text
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSorter.NotesText")
            return ""

    def Sections(self, Prs):
        """The deck's sections: index, name, first slide, slide count."""
        if kS.ErrorMode:
            return []
        try:
            Sp = Prs.SectionProperties
            return [{"index": I, "name": Sp.Name(I), "first": Sp.FirstSlide(I), "count": Sp.SlidesCount(I)}
                    for I in range(1, Sp.Count + 1)]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSorter.Sections")
            return []

    def Outline(self):
        """Sections and slides (id, title, hidden, notes words, change signature)."""
        if kS.ErrorMode:
            return None
        try:
            Live = self.Live
            Prs = Live.Prs()
            Has = Prs.SectionProperties.Count > 0
            Sections = self.Sections(Prs) or [{"index": 0, "name": "", "first": 1, "count": Prs.Slides.Count}]
            Slides = []
            for N in range(1, Prs.Slides.Count + 1):
                Slide = Prs.Slides(N)
                Slides.append({"index": N, "id": Slide.SlideID, "section": Slide.sectionIndex if Has else 0,
                               "title": self.Title(Slide), "hidden": bool(Slide.SlideShowTransition.Hidden),
                               "notes_words": len(self.NotesText(Slide).split()),
                               "signature": Live.Tracker.Signature(Prs, N)})
            for Sec in Sections:
                Sec["slides"] = [X["id"] for X in Slides if X["section"] == Sec["index"]]
            return {**Live.Summary(Prs), "sections": Sections, "slides": Slides, "has_sections": Has}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSorter.Outline")
            return None

    def ThumbsFor(self, Ids, Width):
        """Small PNGs of the given slides, cached by content; deleted slides are left out."""
        if kS.ErrorMode:
            return {"thumbs": []}
        try:
            Live = self.Live
            Prs = Live.Prs()
            ById = self.SlidesById(Prs)
            Out = []
            for Sid in Ids:
                Slide = ById.get(int(Sid))
                if Slide is None:  # deleted since the outline was read: nothing to draw
                    continue
                Sig = Live.Tracker.Signature(Prs, Slide.SlideIndex)
                Key = (Prs.FullName, Sid, Sig, Width)
                if Key not in self.Thumbs:
                    self.Thumbs[Key] = Live.Png(Slide, Width)
                    if len(self.Thumbs) > 400:
                        self.Thumbs.pop(next(iter(self.Thumbs)))
                Out.append({"id": Sid, "signature": Sig, "png": self.Thumbs[Key]})
            return {"thumbs": Out}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSorter.ThumbsFor")
            return {"thumbs": []}

    def MoveOne(self, Prs, Slide, BeforeId, Section):
        """Move one slide before BeforeId, to the end of Section, or to the end of the deck."""
        if kS.ErrorMode:
            return None
        try:
            Has = Prs.SectionProperties.Count > 0
            if BeforeId:
                if int(BeforeId) == Slide.SlideID:
                    return None
                B = self.SlideById(Prs, BeforeId)
                Bi, Bsec = B.SlideIndex, (B.sectionIndex if Has else 0)
                if Has and Prs.SectionProperties.FirstSlide(Bsec) == Bi:
                    Slide.MoveToSectionStart(Bsec)
                else:
                    Slide.MoveTo(Bi - 1 if Slide.SlideIndex < Bi else Bi)
            elif Has and Section:
                Sp = Prs.SectionProperties
                Sec = int(Section)
                if Sp.SlidesCount(Sec) == 0 or (Sp.SlidesCount(Sec) == 1 and Slide.sectionIndex == Sec):
                    Slide.MoveToSectionStart(Sec)
                else:
                    Last = Sp.FirstSlide(Sec) + Sp.SlidesCount(Sec) - 1
                    Slide.MoveTo(Last if Slide.SlideIndex <= Last else Last + 1)
                    if Slide.sectionIndex != Sec:  # landed on the far side of the boundary
                        Slide.MoveTo(Slide.SlideIndex - 1)
            else:
                Slide.MoveTo(Prs.Slides.Count)
            return None
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSorter.MoveOne")
            return None


# ---------------------------------------------------------------- lint + live fixes

class kLintBridge:
    """scripts/lint_deck.py on a copy of the live deck, with shape boxes; and the live fixes."""

    def __init__(self, Live):
        try:
            self.Live = Live
            self.Key = None
            self.Result = None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kLintBridge.__init__")

    def DeckKey(self, Prs):
        """Changes whenever any slide changes."""
        if kS.ErrorMode:
            return None
        try:
            Tracker = self.Live.Tracker
            return (Prs.FullName, Prs.Slides.Count,
                    tuple(Tracker.Signature(Prs, I) for I in range(1, Prs.Slides.Count + 1)))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kLintBridge.DeckKey")
            return None

    def Lint(self):
        """Findings with slide ids and shape boxes, cached until the deck changes."""
        if kS.ErrorMode:
            return None
        try:
            Live = self.Live
            Prs = Live.Prs()
            Key = self.DeckKey(Prs)
            if self.Key == Key:
                return self.Result
            from lint_deck import kLintDeck
            Copy = os.path.join(Live.Work, "lint" + (os.path.splitext(Prs.FullName)[1] or ".pptx"))
            Prs.SaveCopyAs(Copy)
            _, Findings = kLintDeck.Lint(Copy, 18.0, 12)
            Out = []
            for I in Findings.Items:
                Item = {**I, "fixable": I["code"] in FIXABLE and bool(I["shape"]), "slide_id": None, "box": None}
                if I["slide"]:
                    Slide = Prs.Slides(I["slide"])
                    Item["slide_id"] = Slide.SlideID
                    Shape = Live.FindShape(Slide, I["shape"], I.get("shape_id")) if I["shape"] else None
                    if Shape is not None:
                        Item["shape_id"] = Shape.Id
                        Item["box"] = [round(Shape.Left, 1), round(Shape.Top, 1), round(Shape.Width, 1),
                                       round(Shape.Height, 1)]
                Out.append(Item)
            Counts = {}
            for I in Out:
                if I["slide_id"] and I["severity"] in ("error", "warn"):
                    Counts[I["slide_id"]] = Counts.get(I["slide_id"], 0) + 1
            self.Key, self.Result = Key, {"floor_pt": Findings.Floor, "findings": Out, "counts": Counts}
            return self.Result
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kLintBridge.Lint")
            return None

    def FloorPt(self):
        """The body-text floor of the last lint, or 18 pt scaled to the slide width."""
        if kS.ErrorMode:
            return 18.0
        try:
            return (self.Result or {}).get("floor_pt") or 18.0 * max(1.0, self.Live.Prs().PageSetup.SlideWidth / 960)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kLintBridge.FloorPt")
            return 18.0

    def SeriesText(self, Series):
        """'name: cat value, ...' for one chart series, or None when its data cannot be read."""
        if kS.ErrorMode:
            return None
        try:
            try:
                Cats = [str(C) for C in (Series.XValues or [])]
                Vals = list(Series.Values or [])
            except Exception:  # ERROR-SUPPRESSED-JUSTIFIED: some chart kinds (stock, surface, linked data) expose no XValues/Values over COM; the alt text then names only the other series
                return None
            Pairs = ", ".join(f"{C} {V:g}" if isinstance(V, (int, float)) else f"{C} n/a" for C, V in zip(Cats, Vals))
            return f"{Series.Name}: {Pairs}"
        except Exception as e:
            kS.GlobalErrorHandler(e, "kLintBridge.SeriesText")
            return None

    def AltForChart(self, Chart):
        """Alt text for a chart, written from its title and data."""
        if kS.ErrorMode:
            return ""
        try:
            Parts = []
            for I in range(1, Chart.SeriesCollection().Count + 1):
                Text = self.SeriesText(Chart.SeriesCollection(I))
                if Text:
                    Parts.append(Text)
            Title = Chart.ChartTitle.Text + ". " if Chart.HasTitle else ""
            return f"Chart. {Title}" + "; ".join(Parts) + "."
        except Exception as e:
            kS.GlobalErrorHandler(e, "kLintBridge.AltForChart")
            return ""

    def FixShape(self, Shape, Code, Floor):
        """Apply the automatic fix for Code to Shape; returns what was done."""
        if kS.ErrorMode:
            return None
        try:
            if Code not in FIXABLE:
                raise ToolInputException(f"{Code} has no automatic fix - ask Claude")
            if Code == "unused_placeholder":
                Shape.Delete()
                return "deleted the empty placeholder"
            if Code == "body_below_floor":
                N = 0
                for P in Shape.TextFrame.TextRange.Paragraphs():
                    if 0 < P.Font.Size < Floor and len(P.Text.split()) >= 4:
                        P.Font.Size = Floor
                        N += 1
                return f"{N} paragraph(s) raised to {Floor:g} pt"
            if Code == "text_overflow":
                Tf, Tr = Shape.TextFrame, Shape.TextFrame.TextRange
                Room = Shape.Height - Tf.MarginTop - Tf.MarginBottom
                Steps = 0
                while Tr.BoundHeight > Room and Steps < 40:
                    Runs = [R for R in Tr.Runs() if R.Font.Size > Floor]
                    if not Runs:
                        break
                    for R in Runs:
                        R.Font.Size = max(Floor, R.Font.Size - 1)
                    Steps += 1
                Left = " (still too long - cut words)" if Tr.BoundHeight > Room else ""
                return f"text shrunk {Steps} pt to fit its box{Left}"
            if Code == "picture_stretched":
                L, T, W, H = Shape.Left, Shape.Top, Shape.Width, Shape.Height
                Shape.LockAspectRatio = 0
                Shape.ScaleHeight(1, -1)
                Shape.ScaleWidth(1, -1)
                Ratio = Shape.Width / Shape.Height
                Nw, Nh = (W, W / Ratio) if W / H < Ratio else (H * Ratio, H)
                Shape.Width, Shape.Height = Nw, Nh
                Shape.Left, Shape.Top = L + (W - Nw) / 2, T + (H - Nh) / 2
                Shape.LockAspectRatio = -1
                return "resized to its true proportions inside the same box"
            if Code == "a11y_missing_alt_text":
                if Shape.HasChart:
                    Shape.AlternativeText = self.AltForChart(Shape.Chart)
                    return "chart alt text written from its data"
                if Shape.HasTable:
                    Tb = Shape.Table
                    Head = [Tb.Cell(1, C).Shape.TextFrame.TextRange.Text for C in range(1, Tb.Columns.Count + 1)]
                    Shape.AlternativeText = f"Table: {', '.join(Head)}; {Tb.Rows.Count - 1} rows."
                    return "table alt text written from its header"
                raise ToolReportableException("only Claude can describe a picture - ask Claude")
            if not Shape.HasChart:
                raise ToolInputException(f"{Code}: not a chart")
            Chart = Shape.Chart
            if Code == "chart_default_palette":
                for I in range(1, Chart.SeriesCollection().Count + 1):
                    S = Chart.SeriesCollection(I)
                    Fmt = S.Format.Line if Chart.ChartType in LINE_CHARTS else S.Format.Fill
                    Fmt.ForeColor.ObjectThemeColor = 5  # msoThemeColorAccent1
                    Fmt.ForeColor.Brightness = RAMP[(I - 1) % len(RAMP)]
                return "series recoloured as shades of the theme accent"
            if Code == "chart_legend_steals_plot":
                Chart.Legend.Position = -4152  # xlLegendPositionRight
                Chart.Legend.IncludeInLayout = False
                return "legend moved to the right"
            Ax = Chart.Axes(2)  # chart_accounting_zero_dash
            Ax.TickLabels.NumberFormat = "$#,##0" if "$" in Ax.TickLabels.NumberFormat else "#,##0"
            return "Accounting axis format replaced"
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kLintBridge.FixShape")
            return None


# ---------------------------------------------------------------- the live deck (one COM thread)

class kPowerPointLive:
    """Owns the COM thread, the deck binding, the current slide and the change version the view polls."""

    def __init__(self):
        try:
            self.PrsName = None   # FullName of the deck we work on
            self.Current = 1
            self.Version = 0
            self.Work = tempfile.mkdtemp(prefix="pptlive-")
            self.Pool = ThreadPoolExecutor(max_workers=1,
                                           initializer=self.ComInit if sys.platform == "win32" else None)
            self.Tracker = kChangeTracker(self)
            self.History = kVersionHistory(self)
            self.Lint = kLintBridge(self)
            self.Sorter = kSorter(self)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.__init__")

    def ComInit(self):
        """Initialise COM on the pool's one thread."""
        if kS.ErrorMode:
            return None
        try:
            import pythoncom
            pythoncom.CoInitialize()
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.ComInit")
            return None

    # DOCUMENTED EXCEPTION: builds the halt message; it must work exactly when kS.ErrorMode is set
    def HaltedError(self):
        First = kS.FirstError() or {}
        return ToolReportableException(
            f"PowerPoint Live halted after an error in {First.get('location', '?')}: "
            f"{First.get('type', 'Error')}: {First.get('message', '')}. "
            "Press Resume in the view or call powerpoint_resume.")

    # DOCUMENTED EXCEPTION: the one re-arm (a person's Resume); it must work exactly when kS.ErrorMode is set
    def Resume(self):
        First = kS.FirstError()
        kS.Reset()
        return {"resumed": First is not None, "error": First}

    async def OnCom(self, Method, *Args):
        """Run Method on the single COM thread (COM objects must stay on the thread that created them).
        A method that hit the global handler returns its safe default: the caller gets the halt error instead."""
        if kS.ErrorMode:
            raise self.HaltedError()
        try:
            Result = await asyncio.get_running_loop().run_in_executor(self.Pool, Method, *Args)
            if kS.ErrorMode:
                raise self.HaltedError()
            return Result
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.OnCom")
            raise self.HaltedError() from e

    def App(self):
        """The PowerPoint application; on another OS that is an expected state."""
        if kS.ErrorMode:
            return None
        try:
            if sys.platform != "win32":
                raise ToolReportableException("PowerPoint Live needs Windows with desktop PowerPoint")
            import win32com.client
            return win32com.client.Dispatch("PowerPoint.Application")
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.App")
            return None

    def Prs(self):
        """The deck we work on (bound by FullName), or the active deck the first time."""
        if kS.ErrorMode:
            return None
        try:
            App = self.App()
            if self.PrsName:
                for P in App.Presentations:
                    if P.FullName == self.PrsName:
                        return P
                raise ToolReportableException(f"{self.PrsName} is no longer open - call powerpoint_open")
            if App.Presentations.Count == 0:
                raise ToolReportableException("no deck is open - call powerpoint_open with a path")
            P = App.ActivePresentation
            self.PrsName = P.FullName
            return P
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Prs")
            return None

    def Clamp(self, Prs, N):
        """Slide number N kept inside the deck (0 for an empty deck)."""
        if kS.ErrorMode:
            return 0
        try:
            return max(1, min(int(N), Prs.Slides.Count)) if Prs.Slides.Count else 0
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Clamp")
            return 0

    def SyncWindow(self, Prs):
        """Let the PowerPoint window follow the current slide, when its view can (sorter/reading view cannot)."""
        if kS.ErrorMode:
            return None
        try:
            if Prs.Windows.Count and self.Current and Prs.Windows(1).ViewType in GOTO_VIEWS:
                Prs.Windows(1).View.GotoSlide(self.Current)
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.SyncWindow")
            return None

    def Summary(self, Prs):
        """Deck, current slide, count, version and size: the keys every tool result starts with."""
        if kS.ErrorMode:
            return {}
        try:
            return {"deck": os.path.basename(Prs.FullName), "path": Prs.FullName, "slide": self.Current,
                    "count": Prs.Slides.Count, "version": self.Version,
                    "width_pt": Prs.PageSetup.SlideWidth, "height_pt": Prs.PageSetup.SlideHeight}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Summary")
            return {}

    def Png(self, Slide, Width):
        """Base64 PNG of one slide through Slide.Export."""
        if kS.ErrorMode:
            return None
        try:
            Fd, Path = tempfile.mkstemp(suffix=".png", dir=self.Work)
            os.close(Fd)
            try:
                Prs = Slide.Parent
                H = round(Width * Prs.PageSetup.SlideHeight / Prs.PageSetup.SlideWidth)
                Slide.Export(Path, "PNG", int(Width), H)
                with open(Path, "rb") as F:
                    return base64.b64encode(F.read()).decode()
            finally:
                os.remove(Path)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Png")
            return None

    def Walk(self, Shapes):
        """Every shape, group members included."""
        if kS.ErrorMode:
            return
        try:
            for Shape in Shapes:
                yield Shape
                if Shape.Type == 6:  # msoGroup
                    yield from self.Walk(Shape.GroupItems)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Walk")
            return

    def FindShape(self, Slide, Name=None, ShapeId=None):
        """The shape with this id or name, or None."""
        if kS.ErrorMode:
            return None
        try:
            for Shape in self.Walk(Slide.Shapes):
                if (ShapeId and Shape.Id == int(ShapeId)) or (Name and Shape.Name == Name):
                    return Shape
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.FindShape")
            return None

    def Mutate(self, Mutation):
        """Version first, then change, then record what changed on every slide."""
        if kS.ErrorMode:
            return None
        try:
            Prs = self.Prs()
            Tracker = self.Tracker
            self.History.Snapshot(Mutation.Label)
            Before = {Prs.Slides(I).SlideID: Tracker.ShapeMap(Prs.Slides(I)) for I in range(1, Prs.Slides.Count + 1)}
            Result = Mutation.Apply(Prs)
            for I in range(1, Prs.Slides.Count + 1):
                Slide = Prs.Slides(I)
                if Slide.SlideID in Before:
                    Tracker.NoteChanges(Slide, Before[Slide.SlideID])
                else:
                    Tracker.NoteNewSlide(Prs, Slide)
            self.Current = self.Clamp(Prs, self.Current or 1)
            self.Version += 1
            self.SyncWindow(Prs)
            return {**self.Summary(Prs), **(Result or {})}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Mutate")
            return None

    # ------------------------------------------------ COM work (runs on the COM thread)

    def Open(self, Path, Slide):
        """Open or attach to a deck and pick the current slide."""
        if kS.ErrorMode:
            return None
        try:
            App = self.App()
            if Path:
                Full = os.path.abspath(Path)
                Hit = next((P for P in App.Presentations if P.FullName.lower() == Full.lower()), None)
                Hit = Hit or next((P for P in App.Presentations if P.Name.lower() == os.path.basename(Path).lower()),
                                  None)
                if Hit is None:
                    if not os.path.exists(Full):
                        raise ToolInputException(f"not found: {Full}")
                    Hit = App.Presentations.Open(Full)
                if Hit.FullName != self.PrsName:
                    self.Tracker.Clear()
                self.PrsName = Hit.FullName
            Prs = self.Prs()
            self.Current = self.Clamp(Prs, Slide or self.Current)
            self.Version += 1
            self.SyncWindow(Prs)
            return self.Summary(Prs)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Open")
            return None

    def Run(self, Code, Slide, Label):
        """powerpoint_run: the code's print() output and its error text come back with the deck summary."""
        if kS.ErrorMode:
            return None
        try:
            Mutation = kRunMutation(self, Code, Slide, Label)
            Res = self.Mutate(Mutation)
            return {**(Res or {}), "output": Mutation.Output.getvalue()[-8000:], "error": Mutation.Error}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Run")
            return None

    def Show(self, Slide):
        """Make slide n current."""
        if kS.ErrorMode:
            return None
        try:
            Prs = self.Prs()
            self.Current = self.Clamp(Prs, Slide)
            self.Version += 1
            self.SyncWindow(Prs)
            return self.Summary(Prs)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Show")
            return None

    def State(self):
        """Current slide, its shapes, recent changes and a change signature."""
        if kS.ErrorMode:
            return None
        try:
            Prs = self.Prs()
            if self.Current > Prs.Slides.Count:
                self.Current = Prs.Slides.Count
            Out = {**self.Summary(Prs), "signature": self.Tracker.Signature(Prs, self.Current), "shapes": [],
                   "changes": None}
            if self.Current:
                Slide = Prs.Slides(self.Current)
                Shapes = self.Tracker.NoteChanges(Slide)
                Out["shapes"] = [{"id": K, "name": V["name"], "box": V["box"], "text": V["snippet"]}
                                 for K, V in Shapes.items()]
                Out["changes"] = self.Tracker.Changes.get(Slide.SlideID)
                Out["slide_id"] = Slide.SlideID
            return Out
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.State")
            return None

    def AccuratePng(self, Prs, N, Width):
        """Render one slide through Save As JPEG (true embedded fonts), not Slide.Export."""
        if kS.ErrorMode:
            return None
        try:
            Src = os.path.join(self.Work, "acc-src" + (os.path.splitext(Prs.FullName)[1] or ".pptx"))
            Prs.SaveCopyAs(Src)
            Tmp = self.App().Presentations.Open(Src, -1, 0, 0)  # read-only, no window
            Out = os.path.join(self.Work, "acc-out")
            try:
                for I in range(Tmp.Slides.Count, 0, -1):
                    if I != N:
                        Tmp.Slides(I).Delete()
                shutil.rmtree(Out, ignore_errors=True)
                Tmp.SaveCopyAs(Out, 17)  # ppSaveAsJPG -> folder with Slide1.JPG
            finally:
                Tmp.Close()
            Jpg = next(os.path.join(Out, F) for F in os.listdir(Out) if F.lower().endswith(".jpg"))
            from PIL import Image
            Im = Image.open(Jpg)
            Im = Im.resize((int(Width), round(int(Width) * Im.height / Im.width)))
            Buf = io.BytesIO()
            Im.save(Buf, "PNG")
            shutil.rmtree(Out, ignore_errors=True)
            return base64.b64encode(Buf.getvalue()).decode()
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.AccuratePng")
            return None

    def Image(self, Slide, Width, Accurate):
        """PNG of a slide, fast or accurate."""
        if kS.ErrorMode:
            return None
        try:
            Prs = self.Prs()
            N = self.Clamp(Prs, Slide or self.Current)
            if not N:
                return {"slide": 0, "png": None}
            Sl = Prs.Slides(N)
            Png = self.AccuratePng(Prs, N, Width) if Accurate else self.Png(Sl, Width)
            return {"slide": N, "slide_id": Sl.SlideID, "count": Prs.Slides.Count, "png": Png,
                    "accurate": bool(Accurate), "signature": self.Tracker.Signature(Prs, N)}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Image")
            return None

    def Move(self, SlideIds, BeforeId, Section):
        """powerpoint_move."""
        if kS.ErrorMode:
            return None
        try:
            return self.Mutate(kMoveMutation(self, SlideIds, BeforeId, Section))
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Move")
            return None

    def SectionsEdit(self, Action, Section, Name, BeforeId):
        """powerpoint_sections; an unknown action is refused before a version is saved."""
        if kS.ErrorMode:
            return None
        try:
            if Action not in SECTION_ACTIONS:
                raise ToolInputException("action must be add, rename or delete")
            return self.Mutate(kSectionMutation(self, Action, Section, Name, BeforeId))
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.SectionsEdit")
            return None

    def Hide(self, SlideIds, Hidden):
        """powerpoint_hide."""
        if kS.ErrorMode:
            return None
        try:
            return self.Mutate(kHideMutation(self, SlideIds, Hidden))
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Hide")
            return None

    def Fix(self, SlideId, ShapeId, Code):
        """powerpoint_fix; an unknown code is refused before a version is saved."""
        if kS.ErrorMode:
            return None
        try:
            if Code not in FIXABLE:
                raise ToolInputException(f"{Code} has no automatic fix - ask Claude")
            return self.Mutate(kFixMutation(self, SlideId, ShapeId, Code, self.Lint.FloorPt()))
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPowerPointLive.Fix")
            return None


Live = kPowerPointLive()


# ---------------------------------------------------------------- MCP surface

@mcp.resource(URI, name="PowerPoint Live", mime_type=MIME, meta={"ui": {"prefersBorder": True}})
def SlideView() -> str:
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        with open(os.path.join(HERE, "view.html"), encoding="utf-8") as F:
            return F.read()
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.SlideView")
        raise Live.HaltedError() from e


@mcp.tool(name="powerpoint_open", meta=UI)
async def PowerpointOpen(path: str = "", slide: int = 0) -> dict:
    """Open a .pptx in desktop PowerPoint (or attach to it if already open; empty path = the active deck)
    and make `slide` (1-based) the current slide. The live view shows the deck as it changes."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.Open, path, slide)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.PowerpointOpen")
        raise Live.HaltedError() from e


@mcp.tool(name="powerpoint_run", meta=UI)
async def PowerpointRun(code: str, slide: int = 0, label: str = "") -> dict:
    """Run Python against the open deck through COM. In scope: `app`, `prs` (the deck), `slide` (the current slide),
    `goto(n)` (change the current slide - the live view follows), `win32com`. print() output is returned.
    `label` names the change in the History view. A version is saved first, so the user can undo it.
    Changed shapes are highlighted in the live view. Pin to `prs`, never app.ActivePresentation."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.Run, code, slide, label)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.PowerpointRun")
        raise Live.HaltedError() from e


@mcp.tool(name="powerpoint_show", meta=UI)
async def PowerpointShow(slide: int) -> dict:
    """Make slide n (1-based) the current slide in the live view and the PowerPoint window."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.Show, slide)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.PowerpointShow")
        raise Live.HaltedError() from e


@mcp.tool(name="powerpoint_move", meta=UI)
async def PowerpointMove(slide_ids: list[int], before_slide_id: int = 0, section: int = 0) -> dict:
    """Move slides (by SlideID, kept in deck order) to just before `before_slide_id`, or (with `section`, 1-based)
    to the end of that section; neither = to the end of the deck. SlideIDs: prs.Slides(n).SlideID."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.Move, slide_ids, before_slide_id, section)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.PowerpointMove")
        raise Live.HaltedError() from e


@mcp.tool(name="powerpoint_sections", meta=UI)
async def PowerpointSections(action: str, section: int = 0, name: str = "", before_slide_id: int = 0) -> dict:
    """Edit sections. action: "add" (a section named `name` starting at `before_slide_id`, or at the end),
    "rename" (section n to `name`), "delete" (section n; its slides join the section before)."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.SectionsEdit, action, section, name, before_slide_id)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.PowerpointSections")
        raise Live.HaltedError() from e


@mcp.tool(name="powerpoint_hide", meta=UI)
async def PowerpointHide(slide_ids: list[int], hidden: bool = True) -> dict:
    """Hide (or unhide) slides in the slide show, by SlideID."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.Hide, slide_ids, hidden)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.PowerpointHide")
        raise Live.HaltedError() from e


@mcp.tool(name="powerpoint_fix", meta=UI)
async def PowerpointFix(slide_id: int, shape_id: int, code: str) -> dict:
    """Apply the automatic fix for a lint finding to one shape, live: unused_placeholder, body_below_floor,
    text_overflow (shrinks using PowerPoint's own layout, never below the floor), picture_stretched,
    a11y_missing_alt_text (charts and tables), chart_default_palette, chart_legend_steals_plot,
    chart_accounting_zero_dash. A version is saved first."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.Fix, slide_id, shape_id, code)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.PowerpointFix")
        raise Live.HaltedError() from e


@mcp.tool(name="powerpoint_history", meta=UI)
async def PowerpointHistory() -> dict:
    """Versions saved before each change made through these tools (newest first)."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.History.List)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.PowerpointHistory")
        raise Live.HaltedError() from e


@mcp.tool(name="powerpoint_restore", meta=UI)
async def PowerpointRestore(version: int) -> dict:
    """Put a saved version back: the deck file is replaced with it and reopened. The current state is saved
    as a version first, so a restore can itself be undone. Needs a deck that has been saved to disk."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.History.Restore, version)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.PowerpointRestore")
        raise Live.HaltedError() from e


@mcp.tool(name="powerpoint_resume", meta=UI)
async def PowerpointResume() -> dict:
    """Re-arm PowerPoint Live after an unexpected error halted it (the view's Resume button calls this).
    Returns the error that caused the halt, if there was one."""
    if kS.ErrorMode:
        return Live.Resume()
    try:
        return Live.Resume()
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.PowerpointResume")
        raise Live.HaltedError() from e


@mcp.tool(name="error_reports", meta=APP_ONLY)
async def ErrorReports() -> dict:
    """(View only) Error reports waiting for the person's answer to "Do you want to send this error message?".
    Works while halted: the question must be answerable."""
    if kS.ErrorMode:
        return kErrorReport.PendingForView()
    try:
        return kErrorReport.PendingForView()
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.ErrorReports")
        raise Live.HaltedError() from e


@mcp.tool(name="error_report_answer", meta=APP_ONLY)
async def ErrorReportAnswer(send: bool) -> dict:
    """(View only - never the model) The person's answer: send=True sends every waiting report, False discards
    them. Only the view's Yes/No buttons call this."""
    if kS.ErrorMode:
        return kErrorReport.Decide(send)
    try:
        return kErrorReport.Decide(send)
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.ErrorReportAnswer")
        raise Live.HaltedError() from e


@mcp.tool(name="slide_state", meta=APP_ONLY)
async def SlideState() -> dict:
    """(View only) Current slide, its shapes, recent changes and a change signature."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.State)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.SlideState")
        raise Live.HaltedError() from e


@mcp.tool(name="slide_image", meta=APP_ONLY)
async def SlideImage(slide: int = 0, width: int = 1280, accurate: bool = False) -> dict:
    """(View only) PNG of a slide. accurate=True renders through Save As JPEG (true embedded fonts), slower."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.Image, slide, max(320, min(width, 2560)), accurate)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.SlideImage")
        raise Live.HaltedError() from e


@mcp.tool(name="deck_outline", meta=APP_ONLY)
async def DeckOutline() -> dict:
    """(View only) Sections and slides (id, title, hidden, notes words, change signature)."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.Sorter.Outline)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.DeckOutline")
        raise Live.HaltedError() from e


@mcp.tool(name="slide_thumbs", meta=APP_ONLY)
async def SlideThumbs(slide_ids: list[int], width: int = 320) -> dict:
    """(View only) Small PNGs of the given slides, cached by content."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.Sorter.ThumbsFor, slide_ids[:12], max(120, min(width, 640)))
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.SlideThumbs")
        raise Live.HaltedError() from e


@mcp.tool(name="deck_lint", meta=APP_ONLY)
async def DeckLint() -> dict:
    """(View only) lint_deck.py findings with shape boxes, cached until the deck changes."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.Lint.Lint)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.DeckLint")
        raise Live.HaltedError() from e


@mcp.tool(name="history_thumb", meta=APP_ONLY)
async def HistoryThumb(version: int) -> dict:
    """(View only) Thumbnail of the current slide when a version was saved."""
    if kS.ErrorMode:
        raise Live.HaltedError()
    try:
        return await Live.OnCom(Live.History.Thumb, version)
    except kToolException:
        raise
    except Exception as e:
        kS.GlobalErrorHandler(e, "server.HistoryThumb")
        raise Live.HaltedError() from e


if __name__ == "__main__":  # the MCP entry point: FastMCP owns the loop; every tool above carries the pattern
    mcp.run()
