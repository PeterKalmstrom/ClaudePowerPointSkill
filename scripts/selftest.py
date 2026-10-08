"""Self-test for every script in this folder, against a deck it builds itself.

    # any OS: the scripts that need no PowerPoint
    uvx --with python-pptx --with pillow python scripts/selftest.py
    # Windows with PowerPoint installed: everything
    uvx --with python-pptx --with pillow --with pywin32 python scripts/selftest.py --com

The test deck (5 slides, 1440 x 810 pt) has known defects, so each check knows what to expect:
  1 "Decisions"       topic-label title, a 4-word body paragraph at 22 pt   -> body_below_floor, title_is_label
  2 claim title       short body at 24 pt                                    -> OK
  3 claim title       one long word in a narrow box at 60 pt                 -> broken word
  4 claim title       speaker notes                                          -> notes in read_deck
  5 hidden slide                                                             -> [skip]
Exit code 0 = every check passed.
"""
import argparse
import collections
import importlib.util
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

from kShared import kRun, kS, kToolException

HERE = os.path.dirname(os.path.abspath(__file__))
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

# A tiny app for the error-handler checks: Run calls a method that fails, then one that writes a marker file.
HALT_SCRIPT = '''import sys
sys.path.insert(0, {Scripts!r})
from kShared import kRun, kS


class kXApp:
    def Fail(self):
        if kS.ErrorMode:
            return 0
        try:
            Zero = 0
            return 1 / Zero
        except Exception as e:
            kS.GlobalErrorHandler(e, "kXApp.Fail")
            return 0

    def After(self):
        if kS.ErrorMode:
            return
        try:
            with open(sys.argv[1], "w", encoding="utf-8") as File:
                File.write("ran")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kXApp.After")
            return

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            self.Fail()
            self.After()
            return 0
        except Exception as e:
            kS.GlobalErrorHandler(e, "kXApp.Run")
            return 1


kRun.Main(kXApp)
'''

# The same app shape, but Run meets an expected state: a ToolInputException (exit 2, no report).
EXPECTED_SCRIPT = '''import sys
sys.path.insert(0, {Scripts!r})
from kShared import ToolInputException, kRun, kS, kToolException


class kXApp:
    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            raise ToolInputException("selftest: this input is wrong on purpose")
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kXApp.Run")
            return 1


kRun.Main(kXApp)
'''


class kSelfTest:
    """Builds the test decks, runs every script against them and records PASS/FAIL per check."""

    def __init__(self, Com):
        try:
            self.Com = Com
            self.Results = []
            self.Report = []
            self.Tmp = ""
            self.Deck = ""
            self.Photo = ""
            self.Messy = ""
            self.Taste = ""
            self.Sample = ""
            self.Edge = ""
            self.Overflow = ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.__init__")

    def Say(self, Line=""):
        """Print a line and keep it for the report file."""
        if kS.ErrorMode:
            return
        try:
            print(Line)
            self.Report.append(Line)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.Say")
            return

    def Check(self, Name, Ok, Detail=""):
        """Record one check; the detail is shown only when it fails."""
        if kS.ErrorMode:
            return
        try:
            self.Results.append(Ok)
            self.Say(f"{'PASS' if Ok else 'FAIL'}  {Name}" + (f"  ({Detail})" if Detail and not Ok else ""))
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSelfTest.Check({Name})")
            return

    def RunCommand(self, Command):
        """Run a command; returns (exit code, stdout, stderr)."""
        if kS.ErrorMode:
            return -1, "", ""
        try:
            Process = subprocess.run(Command, capture_output=True, text=True, encoding="utf-8")
            return Process.returncode, Process.stdout, Process.stderr
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSelfTest.RunCommand({Command[1:2]})")
            return -1, "", ""

    def RunScript(self, Script, *Args):
        """Run one script of this folder; returns (exit code, stdout + stderr)."""
        if kS.ErrorMode:
            return -1, ""
        try:
            Code, Out, Err = self.RunCommand([sys.executable, os.path.join(HERE, Script), *Args])
            return Code, Out + Err
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSelfTest.RunScript({Script})")
            return -1, ""

    @staticmethod
    def WriteJson(Data, Path):
        """Write a spec file."""
        if kS.ErrorMode:
            return
        try:
            with open(Path, "w", encoding="utf-8") as File:
                json.dump(Data, File)
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSelfTest.WriteJson({Path})")
            return

    @staticmethod
    def ReadText(Path):
        """A text file's contents."""
        if kS.ErrorMode:
            return ""
        try:
            with open(Path, encoding="utf-8") as File:
                return File.read()
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSelfTest.ReadText({Path})")
            return ""

    @staticmethod
    def AddSlide(Deck, Layout, Title):
        """A new slide with its title set."""
        if kS.ErrorMode:
            return None
        try:
            Slide = Deck.slides.add_slide(Layout)
            Slide.shapes.title.text = Title
            return Slide
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.AddSlide")
            return None

    @staticmethod
    def AddText(Slide, Text, Size, Left=100, Top=250, Width=1200, Height=200, Name="Body"):
        """A wrapping text box with one paragraph at the given size."""
        if kS.ErrorMode:
            return None
        try:
            Box = Slide.shapes.add_textbox(Pt(Left), Pt(Top), Pt(Width), Pt(Height))
            Box.name = Name
            Box.text_frame.word_wrap = True
            Box.text_frame.text = Text
            for Run in Box.text_frame.paragraphs[0].runs:
                Run.font.size = Pt(Size)
            return Box
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.AddText")
            return None

    def BuildDeck(self, Path):
        """The 5-slide test deck described at the top of this file."""
        if kS.ErrorMode:
            return
        try:
            Deck = Presentation()
            Deck.slide_width, Deck.slide_height = Pt(1440), Pt(810)
            Layout = Deck.slide_layouts[5]  # Title Only
            self.AddText(self.AddSlide(Deck, Layout, "Decisions"), "Small text below the floor", 22)
            self.AddText(self.AddSlide(Deck, Layout, "The check comes first"), "Plan, check, run", 24)
            self.AddText(self.AddSlide(Deck, Layout, "Long words break lines"), "Internationalization", 60,
                         Width=200, Name="Narrow")
            Slide = self.AddSlide(Deck, Layout, "Notes carry the depth")
            self.AddText(Slide, "Short body", 24)
            Slide.notes_slide.notes_text_frame.text = "Key fact: notes are read by read_deck."
            Hidden = self.AddSlide(Deck, Layout, "Hidden slide")
            Hidden._element.set("show", "0")
            Deck.save(Path)
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSelfTest.BuildDeck({Path})")
            return

    @staticmethod
    def PictureRatio(Picture):
        """Width / height of the picture's image."""
        if kS.ErrorMode:
            return 0
        try:
            Width, Height = Image.open(io.BytesIO(Picture.image.blob)).size
            return Width / Height
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.PictureRatio")
            return 0

    @staticmethod
    def BuildLintEdgeDeck(Path, Photo):
        """Group children in child coordinates, white text on a full-bleed photo, white-on-white table."""
        if kS.ErrorMode:
            return
        try:
            Deck = Presentation()
            Deck.slide_width, Deck.slide_height = Pt(1440), Pt(810)
            Slide = Deck.slides.add_slide(Deck.slide_layouts[5])
            Slide.shapes.title.text = "Groups are measured on the slide"
            Group = Slide.shapes.add_group_shape()
            for Index in range(2):
                Box = Group.shapes.add_textbox(Pt(100 + Index * 450), Pt(300), Pt(400), Pt(100))
                Box.text_frame.text = "Grouped label"
            Xfrm = Group._element.find(f"{P_NS}grpSpPr/{A_NS}xfrm")
            Xfrm.find(f"{A_NS}chOff").set("x", str(Pt(3000)))
            Xfrm.find(f"{A_NS}chOff").set("y", str(Pt(3000)))
            for Box in Group.shapes:  # move children into a far-away child space; the group still sits on the slide
                Box.left, Box.top = Box.left + Pt(2900), Box.top + Pt(2700)
            Slide2 = Deck.slides.add_slide(Deck.slide_layouts[5])
            Slide2.shapes.title.text = "Text on photos is judged by eye"
            Slide2.shapes.add_picture(Photo, 0, 0, Pt(1440), Pt(810))
            Box = Slide2.shapes.add_textbox(Pt(100), Pt(300), Pt(800), Pt(100))
            Box.text_frame.text = "White text over a dark photo"
            Box.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
            Slide3 = Deck.slides.add_slide(Deck.slide_layouts[5])
            Slide3.shapes.title.text = "Table text needs contrast too"
            Table = Slide3.shapes.add_table(2, 2, Pt(100), Pt(300), Pt(800), Pt(200)).table
            for Cell in Table.iter_cells():
                Cell.text = "Unreadable"
                Cell.fill.solid()
                Cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
                Cell.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
            Deck.save(Path)
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSelfTest.BuildLintEdgeDeck({Path})")
            return

    @staticmethod
    def AddBox(Slide, Text, X, Y, W=300, H=80, Size=20):
        """A wrapping text box on the taste slide."""
        if kS.ErrorMode:
            return None
        try:
            Box = Slide.shapes.add_textbox(Pt(X), Pt(Y), Pt(W), Pt(H))
            Box.text_frame.word_wrap = True
            Box.text_frame.text = Text
            for Run in Box.text_frame.paragraphs[0].runs:
                Run.font.size = Pt(Size)
            return Box
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.AddBox")
            return None

    def BuildTasteDeck(self, Path):
        """One slide per group of taste defects, each built deliberately."""
        if kS.ErrorMode:
            return
        try:
            Deck = Presentation()
            Deck.slide_width, Deck.slide_height = Pt(1440), Pt(810)
            Slide = Deck.slides.add_slide(Deck.slide_layouts[5])
            Slide.shapes.title.text = "Taste defects are easy to spot"
            self.AddBox(Slide, "\U0001F4CA Charts", 40, 140)
            self.AddBox(Slide, "Lorem ipsum dolor sit amet, consectetur adipiscing elit.", 40, 240)
            Long = self.AddBox(Slide, "This is a long body sentence that keeps going and going so that it is clearly "
                                      "more than eighty characters.", 380, 140, 600, 120)
            Long.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            Pale = self.AddBox(Slide, "Pale grey text that is hard to read on white", 380, 300, 600, 60)
            Pale.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
            self.AddBox(Slide, "Read the full report in the appendix for the remaining items and...", 380, 380, 600, 60)
            for Index, Slot in enumerate([MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_2, MSO_THEME_COLOR.ACCENT_3,
                                          MSO_THEME_COLOR.ACCENT_4]):
                Shape = Slide.shapes.add_shape(1, Pt(1040 + Index * 90), Pt(140), Pt(80), Pt(80))
                Shape.fill.solid()
                Shape.fill.fore_color.theme_color = Slot
                Shape.shadow.inherit = False
                Effects = etree.SubElement(Shape._element.spPr, f"{A_NS}effectLst")
                etree.SubElement(Effects, f"{A_NS}outerShdw", blurRad="50800")
            for Index, Hex in enumerate(["123456", "654321", "ABCDEF"]):
                Shape = Slide.shapes.add_shape(1, Pt(1040 + Index * 90), Pt(260), Pt(80), Pt(80))
                Shape.fill.solid()
                Shape.fill.fore_color.rgb = RGBColor.from_string(Hex)
            Slide.notes_slide.notes_text_frame.text = "notes"

            ChartSlide = Deck.slides.add_slide(Deck.slide_layouts[5])
            ChartSlide.shapes.title.text = "Revenue rises every quarter"
            Data = CategoryChartData()
            Data.categories = ["East", "West"]
            for Quarter, Values in (("Q1", (1, 2)), ("Q2", (2, 3)), ("Q3", (3, 4))):
                Data.add_series(Quarter, Values)
            Chart = ChartSlide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Pt(100), Pt(150), Pt(800), Pt(500),
                                                Data).chart
            Chart.has_legend = True
            Chart.legend.position = XL_LEGEND_POSITION.TOP
            Chart.legend.include_in_layout = False
            for Series, Hex in zip(Chart.plots[0].series, ["4472C4", "ED7D31", "A5A5A5"]):
                Series.format.fill.solid()
                Series.format.fill.fore_color.rgb = RGBColor.from_string(Hex)
            Chart.value_axis.tick_labels.number_format = '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'
            Chart.value_axis.tick_labels.number_format_is_linked = False
            Deck.save(Path)
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSelfTest.BuildTasteDeck({Path})")
            return

    def Setup(self):
        """The temp folder and the test deck."""
        if kS.ErrorMode:
            return
        try:
            self.Tmp = tempfile.mkdtemp(prefix="pptskill-selftest-")
            self.Deck = os.path.join(self.Tmp, "selftest.pptx")
            self.BuildDeck(self.Deck)
            self.Say(f"test deck: {self.Deck}\n")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.Setup")
            return

    def CheckBasics(self):
        """check_slide_size, cover_crop, backup_snapshot and the rules."""
        if kS.ErrorMode:
            return
        try:
            Code, Out = self.RunScript("check_slide_size.py", self.Deck)
            self.Check("check_slide_size: Full HD deck passes", Code == 0, Out)
            Small = os.path.join(self.Tmp, "small.pptx")
            Presentation().save(Small)
            Code, Out = self.RunScript("check_slide_size.py", Small)
            self.Check("check_slide_size: default 4:3 deck fails", Code == 1, Out)

            self.Photo = os.path.join(self.Tmp, "photo.jpg")
            Image.new("RGB", (1500, 1000), "gray").save(self.Photo)
            Crop = os.path.join(self.Tmp, "photo.crop.jpg")
            Code, Out = self.RunScript("cover_crop.py", self.Photo, "--box", "1440x610", "--out", Crop)
            Width, Height = Image.open(Crop).size if os.path.exists(Crop) else (0, 1)
            self.Check("cover_crop: reports +57% stretch", "+57%" in Out, Out)
            self.Check("cover_crop: output has the box ratio", abs(Width / Height - 1440 / 610) < 0.01,
                       f"{Width}x{Height}")

            Code, Out = self.RunScript("backup_snapshot.py", "--file", self.Deck, "--label", "selftest")
            self.Check("backup_snapshot: copy written", Code == 0 and os.path.exists(Out.strip()), Out)

            sys.path.insert(0, HERE)
            from _rules import kRules
            self.Check("rules: 'Decisions' is a label", kRules.LooksLikeLabel("Decisions"))
            self.Check("rules: 'The check comes first' is a claim", not kRules.LooksLikeLabel("The check comes first"))
            self.Check("rules: a question is not flagged", not kRules.LooksLikeLabel("Why Claude?"))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckBasics")
            return

    def CheckReadDeck(self):
        """read_deck JSON against the test deck."""
        if kS.ErrorMode:
            return
        try:
            Code, Out = self.RunScript("read_deck.py", self.Deck)
            try:
                Data = json.loads(Out)
                self.Check("read_deck: 5 slides with ids and layouts",
                           len(Data["slides"]) == 5 and all(Slide["layout"] for Slide in Data["slides"]))
                self.Check("read_deck: notes read", "Key fact" in Data["slides"][3]["notes"])
                self.Check("read_deck: hidden slide marked", Data["slides"][4]["hidden"])
                self.Check("read_deck: titles and font sizes read", Data["slides"][0]["title"] == "Decisions"
                           and any(22.0 in Shape.get("sizes_pt", []) for Shape in Data["slides"][0]["shapes"]))
            except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: output that is not the expected JSON is a FAIL here
                self.Check("read_deck: valid JSON", False, f"{e}: {Out[:300]}")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckReadDeck")
            return

    def CheckLintDeck(self):
        """lint_deck: cross-platform checks against the test deck, then a messy deck and --fix."""
        if kS.ErrorMode:
            return
        try:
            Code, Out = self.RunScript("lint_deck.py", self.Deck, "--json")
            try:
                Found = {(Item["slide"], Item["code"]) for Item in json.loads(Out.split("\nfixed")[0])["findings"]}
                self.Check("lint_deck: slide 1 body below the floor (27 pt at Full HD)", (1, "body_below_floor") in Found)
                self.Check("lint_deck: slide 1 title is a label", (1, "title_is_label") in Found)
                self.Check("lint_deck: slide 2 has no body/title findings",
                           not {FoundCode for Number, FoundCode in Found if Number == 2}
                           & {"body_below_floor", "title_is_label", "empty_title"})
                self.Check("lint_deck: hidden slide 5 skipped", not any(Number == 5 for Number, _ in Found))
            except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: output that is not the expected JSON is a FAIL here
                self.Check("lint_deck: valid JSON", False, f"{e}: {Out[:300]}")

            self.Messy = os.path.join(self.Tmp, "messy.pptx")
            Deck = Presentation()
            Deck.slide_width, Deck.slide_height = Pt(1440), Pt(810)
            Slide = Deck.slides.add_slide(Deck.slide_layouts[1])  # Title and Content; body left empty
            Slide.shapes.title.text = "Pictures need crops"
            Slide.shapes.add_picture(self.Photo, Pt(100), Pt(300), Pt(600), Pt(200))  # 3:2 photo forced to 3:1
            Deck.save(self.Messy)
            Code, Out = self.RunScript("lint_deck.py", self.Messy)
            self.Check("lint_deck: finds the empty placeholder", "unused_placeholder" in Out, Out)
            self.Check("lint_deck: finds the stretched picture", "picture_stretched" in Out, Out)
            self.Check("lint_deck: treats a file name as missing alt text", "a11y_missing_alt_text" in Out, Out)
            self.Check("lint_deck: exit 1 on errors", Code == 1, Out)
            Fixed = os.path.join(self.Tmp, "messy.fixed.pptx")
            Code, Out = self.RunScript("lint_deck.py", self.Messy, "--fix", "--out", Fixed)
            Code2, Out2 = self.RunScript("lint_deck.py", Fixed)
            self.Check("lint_deck --fix: removes the empty placeholder",
                       os.path.exists(Fixed) and "unused_placeholder" not in Out2, Out2)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckLintDeck")
            return

    def CheckLintTaste(self):
        """lint_deck: taste, contrast and chart checks."""
        if kS.ErrorMode:
            return
        try:
            self.Taste = os.path.join(self.Tmp, "taste.pptx")
            self.BuildTasteDeck(self.Taste)
            Code, Out = self.RunScript("lint_deck.py", self.Taste, "--json")
            try:
                Got = {Item["code"] for Item in json.loads(Out)["findings"]}
                for Want in ("emoji_as_icon", "lorem_ipsum", "centered_long_body", "shadow_overuse",
                             "a11y_low_text_contrast", "off_palette_fill", "accent_overload", "chart_legend_steals_plot",
                             "chart_default_palette", "chart_ordinal_categorical_color", "chart_accounting_zero_dash",
                             "truncated_text"):
                    self.Check(f"lint_deck: finds {Want}", Want in Got, ", ".join(sorted(Got)))
            except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: output that is not the expected JSON is a FAIL here
                self.Check("lint_deck: taste deck JSON", False, f"{e}: {Out[:300]}")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckLintTaste")
            return

    def CheckBuildDeck(self):
        """build_deck: the sample spec builds and lints clean, a bad spec is rejected, directions pass contrast."""
        if kS.ErrorMode:
            return
        try:
            self.Sample = os.path.join(HERE, "..", "examples", "spec", "sample-deck.json")
            Built = os.path.join(self.Tmp, "built.pptx")
            Code, Out = self.RunScript("build_deck.py", self.Sample, "--out", Built)
            self.Check("build_deck: sample spec builds", Code == 0 and os.path.exists(Built), Out)
            Code, Out = self.RunScript("lint_deck.py", Built, "--json")
            try:
                Counts = json.loads(Out)["counts"]
                self.Check("build_deck: output lints clean (no errors or warnings)",
                           not Counts.get("error") and not Counts.get("warn"), Out[-600:])
            except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: output that is not the expected JSON is a FAIL here
                self.Check("build_deck: lint JSON", False, f"{e}: {Out[:300]}")
            Bad = os.path.join(self.Tmp, "bad-spec.json")
            self.WriteJson({"slides": [{"pattern": "kpi", "title": "Too few", "metrics": [{"value": "1", "label": "x"}]},
                                       {"pattern": "bullets", "items": ["no title"]}]}, Bad)
            Code, Out = self.RunScript("build_deck.py", Bad, "--out", os.path.join(self.Tmp, "bad.pptx"))
            self.Check("build_deck: rejects a spec that breaks pattern limits", Code == 2 and "needs 3-6" in Out
                       and "needs a 'title'" in Out, Out)
            sys.path.insert(0, HERE)
            from _rules import kRules
            from build_deck import kDeckDesign
            Directions = json.loads(self.ReadText(os.path.join(HERE, "directions.json")))["directions"]
            Weak = []
            for Direction in Directions:
                Quiet, Muted = kDeckDesign.QuietAndMuted(Direction)
                Ratio = kRules.ContrastRatio
                if (Ratio(Direction["text"], Direction["background"]) < 7
                        or Ratio(Muted, Direction["background"]) < 4.5
                        or Ratio(Muted, Quiet) < 4.5 or Ratio(Direction["text"], Quiet) < 7
                        or Ratio(Direction["accent"], Direction["background"]) < 3
                        or Ratio(Direction["background"], Direction["accent"]) < 4.5):
                    Weak.append(Direction["id"])
            self.Check(f"directions: all {len(Directions)} pass contrast (text 7:1; muted 4.5:1 on background and "
                       "cards; background-on-accent 4.5:1)", not Weak, ", ".join(Weak))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckBuildDeck")
            return

    def CheckRegressions(self):
        """Regressions from the October 2026 code review."""
        if kS.ErrorMode:
            return
        try:
            Tmp = self.Tmp
            self.Edge = Edge = os.path.join(Tmp, "edge")
            os.makedirs(Edge, exist_ok=True)
            Image.new("CMYK", (400, 300), (0, 100, 0, 0)).save(os.path.join(Edge, "cmyk.jpg"))
            Rotated = Image.new("RGB", (600, 300), (200, 30, 30))
            Exif = Rotated.getexif()
            Exif[0x0112] = 6  # stored landscape, shown portrait
            Rotated.save(os.path.join(Edge, "rot.jpg"), exif=Exif)
            EdgeSpec = {"direction": "clean-corporate", "slides": [
                {"pattern": "image", "title": "CMYK photos build", "image": "cmyk.jpg"},
                {"pattern": "image", "title": "Phone photos stay upright", "image": "rot.jpg"},
                {"pattern": "chart", "type": "pie", "title": "Share splits three ways", "categories": ["A", "B", "C"],
                 "series": [{"name": "Share", "values": [1, 2, 3]}]},
                {"pattern": "chart", "type": "line", "title": "Six lines rise", "categories": ["Q1", "Q2", "Q3"],
                 "series": [{"name": f"s{Index}", "values": [1, 2, Index]} for Index in range(6)]},
                {"pattern": "chart", "title": "Gaps in data still build", "categories": ["A", "B"],
                 "series": [{"name": "s", "values": [1, None]}]},
                {"pattern": "big_number", "title": "Numbers as numbers work", "number": 42},
                {"pattern": "bullets", "title": "Loose notes work", "items": ["a"],
                 "notes": {"facts": "one string", "qa": [{"q": "only a question"}]}}]}
            self.WriteJson(EdgeSpec, os.path.join(Edge, "edge.json"))
            EdgeOut = os.path.join(Tmp, "edge.pptx")
            Code, Out = self.RunScript("build_deck.py", os.path.join(Edge, "edge.json"), "--out", EdgeOut, "--lint")
            self.Check("build_deck: edge cases build and lint clean (CMYK, EXIF, pie, 6 lines, None, numbers, notes)",
                       Code == 0 and "0 error(s), 0 warning(s)" in Out, Out[-500:])
            if Code == 0:
                EdgeDeck = Presentation(EdgeOut)
                Picture = [Shape for Shape in EdgeDeck.slides[1].shapes if Shape.shape_type == 13][0]
                self.Check("build_deck: EXIF-rotated photo placed upright",
                           Picture.image.size[0] < Picture.image.size[1] * 2.5
                           and abs(self.PictureRatio(Picture) - Picture.width / Picture.height) < 0.02)
                Chart = [Shape for Shape in EdgeDeck.slides[2].shapes if Shape.has_chart][0].chart
                self.Check("build_deck: pie slices labelled with their category",
                           Chart.plots[0].data_labels.show_category_name)
                Notes = EdgeDeck.slides[6].notes_slide.notes_text_frame.text
                self.Check("build_deck: string 'facts' kept whole in notes", "- one string" in Notes, Notes)
            self.WriteJson({"slides": [
                {"pattern": "table", "title": "Ragged rows fail", "header": ["a", "b"], "rows": [["1", "2", "3"]]},
                {"pattern": "chart", "title": "Unknown highlight fails", "categories": ["A"], "highlight": "Z",
                 "series": [{"name": "s", "values": [1]}]}]}, os.path.join(Edge, "bad.json"))
            Code, Out = self.RunScript("build_deck.py", os.path.join(Edge, "bad.json"), "--out",
                                       os.path.join(Tmp, "x.pptx"))
            self.Check("build_deck: ragged rows and unknown highlight rejected", Code == 2 and "row 1 has 3 cells" in Out
                       and "is not one of the categories" in Out, Out)
            Template = os.path.join(Edge, "tpl.pptx")
            TemplateDeck = Presentation()
            for Layout in TemplateDeck.slide_layouts:
                if Layout.name == "Title Only":
                    Layout.name = "Headline"  # a template whose layouts aren't named the usual way
            TemplateDeck.save(Template)
            self.WriteJson({"template": "tpl.pptx", "slides": [{"pattern": "statement",
                                                                "title": "Templates keep their size"}]},
                           os.path.join(Edge, "tpl.json"))
            Code, Out = self.RunScript("build_deck.py", os.path.join(Edge, "tpl.json"), "--out",
                                       os.path.join(Tmp, "tpl-out.pptx"), "--lint")
            self.Check("build_deck: template without 'Title Only' builds without empty placeholders",
                       "unused_placeholder" not in Out and "slides ->" in Out, Out[-400:])
            LintEdge = os.path.join(Tmp, "lint-edge.pptx")
            self.BuildLintEdgeDeck(LintEdge, self.Photo)
            Code, Out = self.RunScript("lint_deck.py", LintEdge)
            self.Check("lint_deck: grouped shapes measured on the slide (no false off-slide)",
                       "offslide_shape" not in Out, Out)
            self.Check("lint_deck: text on a photo not reported as white on white",
                       "slide 2   a11y_low_text_contrast" not in Out, Out)
            self.Check("lint_deck: table cell contrast checked", "Table cell text" in Out, Out)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckRegressions")
            return

    def CheckTextFit(self):
        """Text fitting: the builder shrinks text to fit and reports what can't; the linter sees the same overflow."""
        if kS.ErrorMode:
            return
        try:
            Tmp, Edge = self.Tmp, self.Edge
            FitSpec = {"slides": [
                {"pattern": "bullets", "title": "Long bullets shrink until they fit", "items": [
                    "Approve the partner budget for the coming two quarters, including regional enablement",
                    "Name an owner per region who reports weekly on pipeline and partner activation",
                    "Agree the Q3 review date and the metrics we will judge the rollout on",
                    "Fund the partner portal and the onboarding content it needs before the first partners sign up",
                    "Hire two partner managers with channel experience and give them a quota from the third quarter",
                    "Publish the playbook internally and walk every regional team through it in a live session",
                    "Revisit pricing for partners so that the margin works for them and for us at volume"]},
                {"pattern": "compare", "title": "Points too long for three cards", "columns": [
                    {"heading": "Option " + Letter, "points": [
                        "Approve the partner budget for two quarters, with regional enablement",
                        "Name an owner per region who reports weekly on pipeline and activation",
                        "Agree the Q3 review date and the metrics we judge the rollout on",
                        "Fund the partner portal and the onboarding content it needs to launch"]}
                    for Letter in "ABC"]}]}
            self.WriteJson(FitSpec, os.path.join(Edge, "fit.json"))
            FitOut = os.path.join(Tmp, "fit.pptx")
            Code, Out = self.RunScript("build_deck.py", os.path.join(Edge, "fit.json"), "--out", FitOut)
            self.Check("build_deck: reports text it cannot fit (exit 3)", Code == 3 and "slide 2" in Out
                       and "does not fit" in Out and "slide 1" not in Out.split("does not fit")[0].split("fit:")[-1],
                       Out)
            FitDeck = Presentation(FitOut)
            Sizes = [Run.font.size.pt for Shape in FitDeck.slides[0].shapes if Shape.name == "Points"
                     for Paragraph in Shape.text_frame.paragraphs for Run in Paragraph.runs]
            self.Check("build_deck: long bullets shrunk, but not below the 27 pt floor",
                       Sizes and 27 <= min(Sizes) < 34, str(Sizes))
            Code, Out = self.RunScript("lint_deck.py", FitOut)
            self.Check("lint_deck: sees the same overflow (text_overflow on slide 2)",
                       "slide 2   text_overflow" in Out, Out)
            self.Overflow = os.path.join(Tmp, "overflow.pptx")
            Deck = Presentation()
            Deck.slide_width, Deck.slide_height = Pt(1440), Pt(810)
            Slide = Deck.slides.add_slide(Deck.slide_layouts[5])
            Slide.shapes.title.text = "Fixed-size boxes overflow"
            Box = Slide.shapes.add_textbox(Pt(100), Pt(200), Pt(500), Pt(80))
            Box.text_frame.word_wrap = True
            Box.text_frame.auto_size = 0  # MSO_AUTO_SIZE.NONE: the box keeps its size
            Box.text_frame.text = "This sentence is far too long for a small fixed box and will spill out below it on the slide"
            Box.text_frame.paragraphs[0].runs[0].font.size = Pt(32)
            Wide = Slide.shapes.add_textbox(Pt(700), Pt(200), Pt(200), Pt(200))
            Wide.text_frame.word_wrap = True
            Wide.text_frame.text = "Internationalization"
            Wide.text_frame.paragraphs[0].runs[0].font.size = Pt(48)
            Deck.save(self.Overflow)
            Code, Out = self.RunScript("lint_deck.py", self.Overflow)
            self.Check("lint_deck: fixed-size box overflow found", "text_overflow" in Out, Out)
            self.Check("lint_deck: word wider than its box found", "word_breaks" in Out, Out)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckTextFit")
            return

    def CheckHarvest(self):
        """harvest_edits: hand edits survive a rebuild."""
        if kS.ErrorMode:
            return
        try:
            Harvest = os.path.join(self.Tmp, "harvest")
            os.makedirs(Harvest, exist_ok=True)
            Spec = {"direction": "clean-corporate", "slides": [
                {"id": "a", "pattern": "statement", "title": "First claim stands"},
                {"id": "b", "pattern": "statement", "title": "Second claim stands"},
                {"id": "c", "pattern": "statement", "title": "Third claim stands"}]}
            self.WriteJson(Spec, os.path.join(Harvest, "spec.json"))
            DeckPath = os.path.join(Harvest, "deck.pptx")
            self.RunScript("build_deck.py", os.path.join(Harvest, "spec.json"), "--out", DeckPath)
            self.RunScript("harvest_edits.py", "manifest", DeckPath)
            Deck = Presentation(DeckPath)
            Deck.save(os.path.join(Harvest, "resaved.pptx"))
            Deck.slides[0].shapes.title.text = "First claim, reworded by a colleague"
            Extra = Deck.slides.add_slide(Deck.slide_layouts[5])
            Extra.shapes.title.text = "A slide a colleague added"
            Extra.shapes.add_picture(self.Photo, Pt(100), Pt(200), Pt(300), Pt(200))
            IdList = Deck.slides._sldIdLst
            Moved = list(IdList)[-1]
            IdList.remove(Moved)
            IdList.insert(1, Moved)  # after slide a
            Gone = list(IdList)[3]   # slide c
            Deck.part.drop_rel(Gone.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"))
            IdList.remove(Gone)
            Deck.save(DeckPath)
            Code, Out = self.RunScript("harvest_edits.py", "harvest", os.path.join(Harvest, "resaved.pptx"),
                                       "--manifest", os.path.join(Harvest, "deck.manifest.json"), "--store",
                                       os.path.join(Harvest, "none"))
            self.Check("harvest_edits: a plain re-save is not an edit", "0 slide(s) harvested, 0 deletion(s)" in Out, Out)
            Code, Out = self.RunScript("harvest_edits.py", "harvest", DeckPath, "--store", os.path.join(Harvest, "store"))
            self.Check("harvest_edits: finds the edit, the added slide and the deletion",
                       "edited   a" in Out and "added" in Out and "deleted  c" in Out, Out)
            Spec["slides"][1]["title"] = "Second claim, updated in the spec"
            self.WriteJson(Spec, os.path.join(Harvest, "spec.json"))
            Rebuilt = os.path.join(Harvest, "rebuilt.pptx")
            self.RunScript("build_deck.py", os.path.join(Harvest, "spec.json"), "--out", Rebuilt)
            Final = os.path.join(Harvest, "final.pptx")
            Code, Out = self.RunScript("harvest_edits.py", "restore", Rebuilt, "--store",
                                       os.path.join(Harvest, "store"), "--out", Final)
            Titles = [Slide.shapes.title.text for Slide in Presentation(Final).slides] if os.path.exists(Final) else []
            self.Check("harvest_edits: restored deck keeps edits, additions, deletions and spec changes",
                       Titles == ["First claim, reworded by a colleague", "A slide a colleague added",
                                  "Second claim, updated in the spec"], str(Titles) + Out)
            Duplicates = ([Part for Part, Count in collections.Counter(zipfile.ZipFile(Final).namelist()).items()
                           if Count > 1] if Titles else ["?"])
            self.Check("harvest_edits: restored file has no duplicate parts", not Duplicates, str(Duplicates))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckHarvest")
            return

    def CheckNotesAndSchema(self):
        """Figures need a source in the notes; the committed schema matches; --plan."""
        if kS.ErrorMode:
            return
        try:
            Tmp, Edge = self.Tmp, self.Edge
            NotesSpec = {"slides": [
                {"pattern": "title", "title": "Revenue up 12 % this year", "notes": "Cover: no source needed."},
                {"pattern": "big_number", "title": "Revenue grew this year", "number": "+12", "unit": "%",
                 "notes": "Strong year."},
                {"pattern": "big_number", "title": "Costs fell this year", "number": "-4", "unit": "%",
                 "notes": {"key_fact": "Costs fell 4 %.", "sources": ["Annual report 2026, p. 12"]}}]}
            self.WriteJson(NotesSpec, os.path.join(Edge, "notes2.json"))
            NotesOut = os.path.join(Tmp, "notes2.pptx")
            self.RunScript("build_deck.py", os.path.join(Edge, "notes2.json"), "--out", NotesOut)
            Code, Out = self.RunScript("lint_deck.py", NotesOut)
            self.Check("lint_deck: figures without a source in the notes flagged; sourced ones and the cover not",
                       "slide 2   figure_without_source" in Out and "slide 3   figure_without_source" not in Out
                       and "slide 1   figure_without_source" not in Out, Out)

            # spec schema: generated from the builder, committed, and the sample validates against it
            Code, Out = self.RunScript("build_deck.py", "--print-schema")
            Committed = self.ReadText(os.path.join(HERE, "spec.schema.json"))
            self.Check("spec schema: committed spec.schema.json matches build_deck --print-schema",
                       Code == 0 and json.loads(Out) == json.loads(Committed),
                       "run: python scripts/build_deck.py --print-schema > scripts/spec.schema.json")
            if importlib.util.find_spec("jsonschema") is not None:
                import jsonschema
                jsonschema.validate(json.loads(self.ReadText(self.Sample)), json.loads(Committed))
                self.Check("spec schema: the sample spec validates", True)
                Typo = {"slides": [{"pattern": "kpi", "title": "Typos are caught", "colour": "red", "metrics": [
                    {"value": "1", "label": "a"}, {"value": "2", "label": "b"}, {"value": "3", "lable": "c"}]}]}
                self.WriteJson(Typo, os.path.join(Edge, "typo.json"))
                Code, Out = self.RunScript("build_deck.py", os.path.join(Edge, "typo.json"), "--out",
                                           os.path.join(Tmp, "t.pptx"))
                self.Check("build_deck: misspelt fields named, nested ones too", Code == 2 and "'colour'" in Out
                           and "'lable' in metrics/2" in Out, Out)
            else:
                self.Say("(skipped schema validation: pip install jsonschema)")
            Code, Out = self.RunScript("build_deck.py", self.Sample, "--plan")
            self.Check("build_deck --plan: prints the story without building",
                       Code == 0 and "Deck plan (19 slides" in Out, Out)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckNotesAndSchema")
            return

    def CheckFixDeck(self):
        """fix_deck: mechanical repairs on the defect fixtures, then lint again."""
        if kS.ErrorMode:
            return
        try:
            Tmp = self.Tmp
            Fixed = os.path.join(Tmp, "taste.fixed.pptx")
            Code, Out = self.RunScript("fix_deck.py", self.Taste, "--out", Fixed)
            for Want in ("palette", "legend", "numfmt", "alt"):
                self.Check(f"fix_deck: '{Want}' fix applied on the taste deck", f"\n{Want} " in "\n" + Out, Out[-600:])
            After = Out.split("What's left")[-1]
            self.Check("fix_deck: chart defects gone after fixing", not any(Name in After for Name in (
                "chart_default_palette", "chart_legend_steals_plot", "chart_accounting_zero_dash")), After)
            Code, Out = self.RunScript("fix_deck.py", self.Messy, "--out", os.path.join(Tmp, "messy.fixed.pptx"))
            After = Out.split("What's left")[-1]
            self.Check("fix_deck: empty placeholder removed and stretched picture un-stretched",
                       "unused_placeholder" not in After and "picture_stretched" not in After, Out[-600:])
            Code, Out = self.RunScript("fix_deck.py", self.Overflow, "--out", os.path.join(Tmp, "overflow.fixed.pptx"),
                                       "--only", "fit")
            self.Check("fix_deck: overflowing text shrunk (not below the floor)",
                       "\nfit " in "\n" + Out and "-> 27 pt" in Out or "-> 2" in Out, Out[-500:])
            Code, Out = self.RunScript("fix_deck.py", self.Taste, "--out", self.Taste)
            self.Check("fix_deck: refuses to overwrite its input", Code == 2, Out)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckFixDeck")
            return

    @staticmethod
    def RichSpec():
        """A spec with one slide of every newer pattern (email, kpi_chart, cost_table, quiz, risks) plus the
        patterns whose rough edges were fixed (process details, a highlighted step, short bullets)."""
        if kS.ErrorMode:
            return {}
        try:
            Notes = {"key_fact": "Self-test slide.", "sources": ["Self-test data"]}
            return {"direction": "clean-corporate", "slides": [
                {"id": "email", "pattern": "email", "title": "This email shows four red flags", "notes": Notes,
                 "from": "IT Desk <help@examp1e-support.co>", "to": "you@example.com",
                 "subject": "URGENT: your mailbox closes today",
                 "body": ["Dear user,", "Verify your password within 2 hours to keep your mailbox:",
                          "https://verify-login.example.co", "IT Department"],
                 "attachment": "Report.html",
                 "callouts": [{"target": "from", "note": "Look-alike sender"},
                              {"target": "subject", "note": "Urgency"},
                              {"target": "body", "line": 2, "note": "Link to a look-alike site"},
                              {"target": "attachment", "note": "Unexpected attachment"}]},
                {"id": "kpichart", "pattern": "kpi_chart", "title": "Revenue grew every quarter", "notes": Notes,
                 "metrics": [{"value": "+44 %", "label": "Q4 vs Q1"}, {"value": "19.8", "label": "MUSD in 2026"}],
                 "highlight_metric": 0, "type": "column", "categories": ["Q1", "Q2", "Q3", "Q4"],
                 "series": [{"name": "Revenue", "values": [4.1, 4.6, 5.2, 5.9]}], "highlight": "Q4"},
                {"id": "cost", "pattern": "cost_table", "title": "The pilot costs 200 kSEK", "notes": Notes,
                 "unit": "kSEK", "note": "The buffer is only spent if needed.",
                 "rows": [{"item": "Overtime buffer", "amount": 120, "detail": "Covers peaks"},
                          {"item": "Tooling", "amount": 30, "detail": "Scheduling tools"},
                          {"item": "Evaluation", "amount": 50, "detail": "Surveys"}]},
                {"id": "quiz", "pattern": "quiz", "title": "Which of these is phishing?", "reveal": "slide",
                 "notes": "Hands up before the answer.", "explain": "A: urgency and a payment request.",
                 "options": [{"text": "A text from 'the CEO' asking for gift cards", "correct": True},
                             {"text": "The HR newsletter from the usual address"},
                             {"text": "A meeting invite you agreed yesterday"}]},
                {"id": "risks", "pattern": "risks", "title": "Three risks could slow next year", "notes": Notes,
                 "highlight": 0, "risks": [
                     {"risk": "Longer sales cycles", "likelihood": "high", "impact": "high",
                      "mitigation": "Add sales engineers to take technical validation off the deal path."},
                     {"risk": "A data-centre outage", "likelihood": "low", "impact": "high",
                      "mitigation": "Report the post-mortem actions at the next review."},
                     {"risk": "Hiring behind plan", "likelihood": "medium", "impact": "medium",
                      "mitigation": "Prioritise the open roles by revenue impact."}]},
                {"id": "steps", "pattern": "process", "title": "Rolling it out takes three steps", "notes": Notes,
                 "highlight": 1, "steps": [
                     {"label": "Pick partners", "detail": "One partner in every region"},
                     {"label": "Train the team", "detail": "Two days of hands-on training"},
                     {"label": "Measure", "detail": "The same metrics as last quarter"}]},
                {"id": "asks", "pattern": "bullets", "title": "Three decisions for today", "notes": Notes,
                 "items": ["Approve the partner budget", "Name an owner per region", "Agree the review date"]}]}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.RichSpec")
            return {}

    @staticmethod
    def Bottom(Slide):
        """The lowest edge (pt) of any shape on a slide other than its title."""
        if kS.ErrorMode:
            return 0.0
        try:
            return max([(Shape.top + Shape.height) / 12700 for Shape in Slide.shapes
                        if not (Shape.is_placeholder and Shape.has_text_frame and Shape == Slide.shapes.title)] or [0])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.Bottom")
            return 0.0

    def CheckPatterns(self):
        """The newer patterns build and lint clean, slides are filled, and the October 2026 rough edges stay fixed."""
        if kS.ErrorMode:
            return
        try:
            Tmp, Edge = self.Tmp, self.Edge
            Spec = self.RichSpec()
            self.WriteJson(Spec, os.path.join(Edge, "rich.json"))
            Rich = os.path.join(Tmp, "rich.pptx")
            Code, Out = self.RunScript("build_deck.py", os.path.join(Edge, "rich.json"), "--out", Rich)
            self.Check("build_deck: email, kpi_chart, cost_table, quiz and risks build (quiz adds its answer slide)",
                       Code == 0 and f"{len(Spec['slides']) + 1} slides" in Out, Out)
            Code, Out = self.RunScript("lint_deck.py", Rich, "--json")
            try:
                Items = json.loads(Out)["findings"]
            except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: output that is not the expected JSON is a FAIL here
                Items = [{"severity": "error", "code": f"no JSON: {e}", "slide": 0}]
            Bad = [f"{I['slide']}:{I['code']}" for I in Items if I["severity"] in ("error", "warn")]
            self.Check("build_deck: the new patterns lint clean (no errors or warnings)", not Bad, ", ".join(Bad))
            Codes = {(I["slide"], I["code"]) for I in Items}
            self.Check("lint_deck: no body_below_floor on process details (labels are no longer 24 pt)",
                       not any(C == "body_below_floor" for _, C in Codes), str(sorted(Codes)))
            self.Check("lint_deck: word budget is per pattern (no word_budget on the email, risks, process or "
                       "cost slides)", not any(C == "word_budget" and N in (1, 3, 6, 7) for N, C in Codes),
                       str(sorted(Codes)))
            Deck = Presentation(Rich)
            Slides = list(Deck.slides)
            Low = [N for N, Slide in enumerate(Slides, 1) if self.Bottom(Slide) < 0.8 * 810]
            self.Check("build_deck: content reaches the lower part of every slide (no empty bottom 40 %)",
                       not Low, f"slides ending above 80 % of the height: {Low}")
            Sizes = [Run.font.size.pt for Shape in Slides[-1].shapes if Shape.name.startswith("Point")
                     and Shape.has_text_frame for Para in Shape.text_frame.paragraphs for Run in Para.runs]
            self.Check("build_deck: a few short bullets become bands with text grown above 34 pt (up to 44 pt)",
                       Sizes and 34 < min(Sizes) <= 44 and len({S.name for S in Slides[-1].shapes} & {
                           "Point1", "Point2", "Point3", "PointBand3"}) == 4, str(Sizes))
            Details = [Run.font.size.pt for Shape in Slides[-2].shapes if Shape.name.startswith("StepDetail")
                       for Para in Shape.text_frame.paragraphs for Run in Para.runs]
            self.Check("build_deck: process details at or above the 27 pt floor", Details and min(Details) >= 27,
                       str(Details))
            Arrow = [Shape for Shape in Slides[-2].shapes if Shape.name == "Step2"]
            self.Check("build_deck: a step's number sits inside its arrow, clear of the notch",
                       Arrow and Arrow[0].text_frame.text == "2"
                       and Arrow[0].text_frame.margin_left >= Pt(96 * 0.32), str(Arrow))
            Table = [Shape for Shape in Slides[2].shapes if Shape.name == "CostTable"]
            Last = Table[0].table.rows[len(Table[0].table.rows) - 1] if Table else None
            self.Check("build_deck: cost_table adds the total row (120 + 30 + 50 = 200)",
                       Last is not None and Last.cells[0].text == "Total" and Last.cells[len(Last.cells) - 1].text == "200",
                       str([C.text for C in Last.cells]) if Last is not None else "no table")
            Names = {Shape.name for Shape in Slides[0].shapes}
            self.Check("build_deck: email has its header, body and a numbered marker per callout",
                       {"EmailFrom", "EmailSubject", "EmailBody1", "EmailAttachment", "Marker4", "Callout4"} <= Names,
                       str(sorted(Names)))
            Answer = Slides[4]
            AnswerNotes = Answer.notes_slide.notes_text_frame.text if Answer.has_notes_slide else ""
            self.Check("build_deck: quiz reveal 'slide' adds an answer slide; the answer is in the notes",
                       Answer.shapes.title.text_frame.text == "Answer: A" and "ANSWER: A:" in AnswerNotes
                       and "ANSWER: A:" in Slides[3].notes_slide.notes_text_frame.text, AnswerNotes)
            Chips = [Shape.text_frame.text for Shape in Slides[5].shapes if "Chip" in Shape.name]
            self.Check("build_deck: risks get a likelihood and an impact chip each", len(Chips) == 6
                       and "HIGH" in Chips and "LOW" in Chips, str(Chips))

            # the spec checks for the newer patterns and the compare heading limit
            Bad = {"slides": [
                {"pattern": "compare", "title": "Three columns", "columns": [
                    {"heading": "A heading of thirty characters", "points": ["x"]}, {"heading": "B", "points": ["y"]},
                    {"heading": "C", "points": ["z"]}]},
                {"pattern": "quiz", "title": "No right answer", "options": [{"text": "a"}, {"text": "b"}]},
                {"pattern": "email", "title": "Bad callouts", "from": "a", "subject": "b", "body": ["c"],
                 "callouts": [{"target": "body", "line": 3, "note": "x"}, {"target": "attachment", "note": "y"}]},
                {"pattern": "cost_table", "title": "Text amount", "rows": [{"item": "a", "amount": "lots"}]},
                {"pattern": "kpi_chart", "title": "Bad metric", "metrics": [{"value": "1", "label": "a"},
                                                                           {"value": "2", "label": "b"}],
                 "highlight_metric": 5, "categories": ["a", "b"], "series": [{"name": "s", "values": [1, 2]}]}]}
            self.WriteJson(Bad, os.path.join(Edge, "bad-rich.json"))
            Code, Out = self.RunScript("build_deck.py", os.path.join(Edge, "bad-rich.json"), "--out",
                                       os.path.join(Tmp, "bad-rich.pptx"))
            for Want, Text in (("3-column compare heading over 24 characters", "with 3 columns the limit is 24"),
                               ("quiz without a correct option", "\"correct\": true"),
                               ("email callout on a missing body line", "line 3 is not a body paragraph"),
                               ("email callout on a missing attachment", "points at the attachment"),
                               ("cost_table amount that is not a number", "needs a numeric 'amount'"),
                               ("kpi_chart highlight_metric out of range", "highlight_metric 5")):
                self.Check(f"build_deck: rejects a {Want} (exit 2)", Code == 2 and Text in Out, Out[-800:])

            # notes: the builder says which slides have none
            self.WriteJson({"slides": [{"pattern": "statement", "title": "A slide without notes"}]},
                           os.path.join(Edge, "nonotes.json"))
            Code, Out = self.RunScript("build_deck.py", os.path.join(Edge, "nonotes.json"), "--out",
                                       os.path.join(Tmp, "nonotes.pptx"))
            self.Check("build_deck: warns about a slide without speaker notes (exit code unchanged)",
                       Code == 0 and "has no speaker notes" in Out, Out)

            # contact sheet from a folder of renders (any OS), fix_deck --in-place keeps a .bak
            Renders = os.path.join(Tmp, "sheet")
            os.makedirs(Renders, exist_ok=True)
            for Number in (1, 2, 10):
                Image.new("RGB", (320, 180), "white").save(os.path.join(Renders, f"s{Number:03d}.png"))
            Code, Out = self.RunScript("contact_sheet.py", "--renders", Renders, "--cols", "2")
            self.Check("contact_sheet --renders: one sheet from a folder of renders (any OS)",
                       Code == 0 and os.path.exists(os.path.join(Renders, "contact.png")), Out)
            InPlace = os.path.join(Tmp, "inplace.pptx")
            shutil.copy(self.Taste, InPlace)
            Code, Out = self.RunScript("fix_deck.py", InPlace, "--in-place")
            self.Check("fix_deck --in-place: fixes the input and keeps the original as .bak",
                       os.path.exists(InPlace + ".bak") and "-> " + InPlace in Out
                       and open(InPlace, "rb").read() != open(InPlace + ".bak", "rb").read(), Out[-400:])
            if shutil.which("soffice") and shutil.which("pdftoppm"):
                Code, Out = self.RunScript("render_lo.py", Rich, "--out", os.path.join(Tmp, "rich-lo"), "--sheet")
                self.Check("render_lo --sheet: renders and a contact sheet", Code == 0 and os.path.exists(
                    os.path.join(Tmp, "rich-lo", "contact.png")), Out)
            else:
                self.Say("(skipped render_lo --sheet: LibreOffice/pdftoppm not installed)")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckPatterns")
            return

    def CheckThemeAndRenders(self):
        """extract_theme, diff_renders on synthetic images, render_lo where LibreOffice is installed."""
        if kS.ErrorMode:
            return
        try:
            Tmp = self.Tmp
            Code, Out = self.RunScript("extract_theme.py", self.Deck)
            try:
                Theme = json.loads(Out)
                self.Check("extract_theme: fonts and accent colours read",
                           bool(Theme["fonts"]["major_latin"]) and len(Theme["colours"]) == 12)
                self.Check("extract_theme: layouts listed", len(Theme["layouts"]) >= 6)
            except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: output that is not the expected JSON is a FAIL here
                self.Check("extract_theme: valid JSON", False, f"{e}: {Out[:300]}")

            RendersA, RendersB = os.path.join(Tmp, "ra"), os.path.join(Tmp, "rb")
            for Folder in (RendersA, RendersB):
                os.makedirs(Folder, exist_ok=True)
                Image.new("RGB", (320, 180), "white").save(os.path.join(Folder, "s001.png"))
            Image.new("RGB", (320, 180), "white").save(os.path.join(RendersA, "s002.png"))
            Image.new("RGB", (320, 180), "black").save(os.path.join(RendersB, "s002.png"))
            Code, Out = self.RunScript("diff_renders.py", RendersA, RendersB)
            self.Check("diff_renders: unchanged slide reported same", "s001.png   same" in Out, Out)
            self.Check("diff_renders: changed slide reported", "s002.png   CHANGED" in Out and Code == 1, Out)

            if shutil.which("soffice") and shutil.which("pdftoppm"):
                Code, Out = self.RunScript("render_lo.py", self.Deck, "--out", os.path.join(Tmp, "lo"))
                Count = len([Word for Word in Out.split() if Word.endswith(".png")])
                self.Check("render_lo: one PNG per visible slide", Code == 0 and Count >= 4, Out)
            else:
                self.Say("(skipped render_lo: LibreOffice/pdftoppm not installed)")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckThemeAndRenders")
            return

    def CheckErrorPattern(self):
        """The pattern audit passes, a reported error halts and exits 1, an expected state exits with its code."""
        if kS.ErrorMode:
            return
        try:
            Code, Out, Err = self.RunCommand([sys.executable, os.path.join(HERE, "..", "tools", "check_kpattern.py")])
            Tail = " | ".join((Out + Err).strip().splitlines()[-5:])
            self.Check("pattern: every function follows the kShared error pattern", Code == 0, Tail)

            HaltScript = os.path.join(self.Tmp, "halt_app.py")
            Marker = os.path.join(self.Tmp, "halt_marker.txt")
            with open(HaltScript, "w", encoding="utf-8") as File:
                File.write(HALT_SCRIPT.format(Scripts=HERE))
            Code, Out, Err = self.RunCommand([sys.executable, HaltScript, Marker])
            self.Check("error handler: a reported error halts and exits 1",
                       Code == 1 and "ERROR in" in Err and "ZeroDivisionError" in Err and not os.path.exists(Marker),
                       f"exit {Code}, marker {'written' if os.path.exists(Marker) else 'absent'}: {Err[-400:]}")

            ExpectedScript = os.path.join(self.Tmp, "expected_app.py")
            with open(ExpectedScript, "w", encoding="utf-8") as File:
                File.write(EXPECTED_SCRIPT.format(Scripts=HERE))
            Code, Out, Err = self.RunCommand([sys.executable, ExpectedScript])
            self.Check("error handler: expected state exits with its code, no report",
                       Code == 2 and "this input is wrong on purpose" in Err and "ERROR in" not in Err,
                       f"exit {Code}: {Err[-400:]}")

            # no terminal + an address: the report is saved, nothing is sent, exit 4 tells Claude to ask the user
            Env = dict(os.environ, KPS_ERROR_REPORT="1", KPS_ERROR_WEBHOOK="http://127.0.0.1:9/never-called")
            Run = subprocess.run([sys.executable, HaltScript, Marker], capture_output=True, text=True,
                                 stdin=subprocess.DEVNULL, env=Env)
            Pending = [Line.split(":", 1)[1].strip() for Line in Run.stderr.splitlines()
                       if Line.startswith("ERROR-REPORT-PENDING:")]
            Saved = bool(Pending) and os.path.isfile(Pending[0])
            self.Check("error report: no terminal -> exit 4, report saved, Claude told to ask, nothing sent",
                       Run.returncode == 4 and Saved and "Do you want to send this error message?" in Run.stderr
                       and "Error report sent" not in Run.stderr, f"exit {Run.returncode}: {Run.stderr[-400:]}")
            if Saved:
                Run = subprocess.run([sys.executable, os.path.join(HERE, "send_error_report.py"), Pending[0], "--no"],
                                     capture_output=True, text=True, stdin=subprocess.DEVNULL, env=Env)
                self.Check("error report: --no deletes it and sends nothing",
                           Run.returncode == 0 and not os.path.exists(Pending[0]), Run.stdout + Run.stderr)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckErrorPattern")
            return

    def CheckCom(self):
        """Windows + PowerPoint: the COM scripts."""
        if kS.ErrorMode:
            return
        try:
            Tmp, Deck = self.Tmp, self.Deck
            if not self.Com:
                self.Say("\n(skipped the PowerPoint scripts; run with --com on Windows)")
                return
            Code, Out = self.RunScript("check_word_breaks.py", "--file", Deck)
            self.Check("check_word_breaks: finds the broken word on slide 3", Code == 1 and "slide 3" in Out, Out)

            Renders = os.path.join(Tmp, "renders")
            Code, Out = self.RunScript("render_slides.py", "--file", Deck, "--out", Renders)
            Count = len([Name for Name in os.listdir(Renders) if Name.endswith(".jpg")]) if os.path.isdir(Renders) else 0
            self.Check("render_slides: one JPEG per slide", Code == 0 and Count == 5, Out)

            Sheet = os.path.join(Tmp, "contact.png")
            Code, Out = self.RunScript("contact_sheet.py", "--file", Deck, "--out", Sheet)
            self.Check("contact_sheet: PNG written", Code == 0 and os.path.exists(Sheet), Out)

            for Mode in ("notes", "handout", "slides"):
                Pdf = os.path.join(Tmp, f"selftest.{Mode}.pdf")
                Code, Out = self.RunScript("export_pdf.py", "--file", Deck, "--mode", Mode, "--out", Pdf)
                Ok = False
                if Code == 0 and os.path.exists(Pdf):
                    with open(Pdf, "rb") as File:
                        Ok = File.read(4) == b"%PDF"
                self.Check(f"export_pdf: {Mode} PDF written", Ok, Out)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckCom")
            return

    def Finish(self):
        """The summary line and the report file. Returns the number of failed checks."""
        if kS.ErrorMode:
            return 1
        try:
            Failed = self.Results.count(False)
            self.Say(f"\n{len(self.Results) - Failed}/{len(self.Results)} passed  -  output in {self.Tmp}")
            Log = os.path.join(self.Tmp, "selftest-report.txt")
            with open(Log, "w", encoding="utf-8") as File:
                File.write("\n".join(self.Report) + "\n")
            print(f"report: {Log}  (paste or attach this file when reporting results)")
            return Failed
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.Finish")
            return 1


class kSelfTestApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Parser = argparse.ArgumentParser()
            Parser.add_argument("--com", action="store_true", help="also run the PowerPoint (COM) scripts; Windows only")
            Args = Parser.parse_args()
            os.environ["KPS_ERROR_REPORT"] = "0"  # the scripts this starts never send error reports
            Test = kSelfTest(Args.com)
            Test.Setup()
            Test.CheckBasics()
            Test.CheckReadDeck()
            Test.CheckLintDeck()
            Test.CheckLintTaste()
            Test.CheckBuildDeck()
            Test.CheckRegressions()
            Test.CheckTextFit()
            Test.CheckHarvest()
            Test.CheckNotesAndSchema()
            Test.CheckFixDeck()
            Test.CheckPatterns()
            Test.CheckThemeAndRenders()
            Test.CheckErrorPattern()
            Test.CheckCom()
            return 1 if Test.Finish() else 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kSelfTestApp)
