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
import glob
import importlib.util
import io
import json
import os
import re
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
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

from kShared import ToolReportableException, kRun, kS, kToolException

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


class kCiFonts:
    """--ci-fonts: measure and substitute with only the fonts the GitHub Actions runner has (DejaVu, Liberation,
    Carlito, Caladea), without touching the system's fonts - a private font folder of links to those files, a
    fontconfig file that knows only it (FONTCONFIG_FILE, for fc-match and LibreOffice) and PPTSKILL_FONT_DIRS (for
    _measure's scan). The scripts this test starts inherit both."""

    FAMILIES = ("DejaVu", "Liberation", "Carlito", "Caladea")

    @staticmethod
    def Setup(Root):
        """Make Root/fonts and Root/fonts.conf and point this process and its children at them; the font count."""
        if kS.ErrorMode:
            return 0
        try:
            Dir = os.path.join(Root, "fonts")
            os.makedirs(Dir, exist_ok=True)
            Seen = set()
            for Base in ("/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts"),
                         os.path.expanduser("~/.local/share/fonts")):
                for Path in sorted(glob.glob(os.path.join(Base, "**", "*.ttf"), recursive=True)):
                    Name = os.path.basename(Path)
                    if Name.startswith(kCiFonts.FAMILIES) and Name not in Seen:
                        Seen.add(Name)
                        shutil.copyfile(Path, os.path.join(Dir, Name))
            if not Seen:
                raise ToolReportableException("--ci-fonts needs the DejaVu, Liberation, Carlito and Caladea fonts "
                                              "installed (fonts-dejavu fonts-liberation2 fonts-crosextra-*)")
            Conf = os.path.join(Root, "fonts.conf")
            with open(Conf, "w", encoding="utf-8") as File:
                File.write('<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "fonts.dtd">\n<fontconfig>\n'
                           f'  <dir>{Dir}</dir>\n  <cachedir>{os.path.join(Root, "cache")}</cachedir>\n'
                           '  <include ignore_missing="yes">/etc/fonts/conf.d</include>\n</fontconfig>\n')
            os.environ["FONTCONFIG_FILE"] = Conf
            os.environ["PPTSKILL_FONT_DIRS"] = Dir
            return len(Seen)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kCiFonts.Setup")
            return 0


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
            self.Check("rules: a time window alone is not a measurable target",
                       not kRules.Measurable("<= last 6 months", "Churn") and not kRules.Measurable("Clearly lower")
                       and not kRules.Measurable("by Q3", "Delivery"))
            self.Check("rules: targets with a unit, comparator or count are measurable",
                       kRules.Measurable("Below 30 %", "Burnout") and kRules.Measurable(">= 95 % of Q4", "Delivery")
                       and kRules.Measurable("4 of 4", "Weekly reviews held")
                       and kRules.Measurable("< 2 days lead time"))
            self.Check("rules: an ask states its cost in money, time or FTE",
                       kRules.StatesCost("Approve a pilot, 200 kSEK") and kRules.StatesCost("2 FTE for 6 months")
                       and not kRules.StatesCost("Approve three sales engineers"))
            self.Check("rules: word count ignores bullets and dashes", kRules.WordCount("- Revenue grew 44 % – fast") == 4)
            from _figures import kFigures
            Facts = {"rev": [4.1, 4.6, 5.2, 5.9]}
            Errs = []
            Text = kFigures.Replace("up {change:rev|abs}, {change:rev}", None, None, Facts, "t", Errs)
            Plain = " ".join(str(Text).replace(" ", " ").split())
            self.Check("figures: {change|abs} drops the sign", Plain == "up 44 %, +44 %" and not Errs,
                       f"{Text} {Errs}")
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
            self.Check("build_deck: rejects a spec that breaks pattern limits", Code == 2 and "needs 2-6" in Out and "big_number" in Out
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
                       Code == 0 and "Deck plan (21 slides" in Out and "SUMMARY / Margins" in Out, Out)
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

    def CheckDesignSystem(self):
        """The visual system: kicker, footer and page numbers, the metrics and next_steps patterns, the decision
        box, balanced titles, and the lint checks for unwanted wraps and title widows."""
        if kS.ErrorMode:
            return
        try:
            Tmp, Edge = self.Tmp, self.Edge
            Notes = {"key_fact": "Self-test slide.", "sources": ["Self-test data"]}
            Spec = {"direction": "clean-corporate", "footer": "Self-test deck · Q4", "slides": [
                {"id": "cover", "pattern": "title", "title": "A self-test deck", "notes": Notes},
                {"id": "part", "pattern": "section", "eyebrow": "Results", "title": "What we found", "notes": Notes},
                {"id": "summary", "pattern": "kpi", "title": "Revenue grew and churn fell", "notes": Notes,
                 "metrics": [{"value": "112 %", "label": "Net retention"},
                             {"value": "\u221233%", "label": "Churn", "trend": [1.8, 1.6, 1.4, 1.2]},
                             {"value": "5.9 M", "label": "Q4 revenue"}],
                 "decision": "Approve three sales engineers, 2.4 MSEK a year"},
                {"id": "success", "pattern": "metrics", "kicker": "Success", "title": "Success has a target",
                 "notes": Notes, "rows": [{"metric": "Burnout", "baseline": "41 %", "target": "Below 30 %",
                                           "owner": "HR", "date": "June"},
                                          {"metric": "Delivery", "baseline": "Q4 average", "target": "95 %"}],
                 "rule": "Stop if delivery drops 10 % for two months."},
                {"id": "next", "pattern": "next_steps", "title": "Three steps take it live", "notes": Notes,
                 "steps": [{"action": "Approve the budget", "owner": "Board", "date": "Today"},
                           {"action": "Hire the team", "owner": "HR", "date": "Q1 2027"},
                           {"action": "Report back", "owner": "VP Sales", "date": "Q2 2027"}]},
                {"id": "widow", "pattern": "bullets", "kicker": "", "notes": Notes,
                 "title": "Enterprise deals now close slower than in any quarter",
                 "items": ["One", "Two"]},
                {"id": "ask", "pattern": "statement", "title": "Approve the pilot today", "notes": Notes,
                 "decision": "Approve a six-month pilot with a 200 kSEK budget", "owner": "CTO", "date": "Today"}]}
            Path = os.path.join(Edge, "design.json")
            self.WriteJson(Spec, Path)
            Deck = os.path.join(Tmp, "design.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck)
            self.Check("build_deck: metrics, next_steps and a decision box build", Code == 0, Out)
            Code, Out = self.RunScript("lint_deck.py", Deck, "--json")
            try:
                Items = json.loads(Out)["findings"]
            except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: output that is not the expected JSON is a FAIL here
                Items = [{"severity": "error", "code": f"no JSON: {e}", "slide": 0}]
            Bad = [f"{I['slide']}:{I['code']}" for I in Items if I["severity"] in ("error", "warn")]
            self.Check("build_deck: the design-system deck lints clean (no wraps, widows, errors or warnings)",
                       not Bad, ", ".join(Bad))
            self.Check("lint_deck: weak_focal_hierarchy does not fire on the builder's own layouts",
                       not any(I["code"] == "weak_focal_hierarchy" for I in Items), str(Items)[:400])
            Slides = list(Presentation(Deck).slides)
            Names = [{Shape.name: Shape for Shape in Slide.shapes} for Slide in Slides]
            self.Check("build_deck: footer and page number on content slides, not on the cover or the divider",
                       "Footer" not in Names[0] and "PageNumber" not in Names[1] and "Footer" in Names[2]
                       and Names[2]["PageNumber"].text_frame.text == "3"
                       and Names[2]["Footer"].text_frame.text == "Self-test deck · Q4",
                       str([sorted(N) for N in Names[:3]]))
            self.Check("build_deck: the kicker defaults to the section's eyebrow, a slide's own kicker wins, "
                       "\"\" turns it off", Names[2].get("Kicker") is not None
                       and Names[2]["Kicker"].text_frame.text == "RESULTS"
                       and Names[3]["Kicker"].text_frame.text == "SUCCESS" and "Kicker" not in Names[5],
                       str([N.get("Kicker").text_frame.text if N.get("Kicker") else None for N in Names]))
            Value = Names[2]["Value1"].text_frame.text
            self.Check("build_deck: number and unit joined by a no-break space ('112\u00a0%')", Value == "112\u00a0%",
                       repr(Value))
            self.Check("build_deck: a tile trend above a decision is legible bars (72 pt or more) or moves to the "
                       "speaker notes - never tiny; the decision box sits under the tiles",
                       {"DecisionBox", "Decision"} <= set(Names[2])
                       and self.TrendLegible(Names[2], self.NotesOf(Slides[2]), "Trend2"), str(sorted(Names[2])))
            Table = Names[3].get("MetricsTable")
            Head = [C.text for C in Table.table.rows[0].cells] if Table is not None else []
            self.Check("build_deck: metrics is a table with baseline, target, owner and date, and a rule below",
                       Head == ["Metric", "Baseline", "Target", "Owner", "By"] and "MetricsRuleText" in Names[3],
                       str(Head))
            self.Check("build_deck: next_steps draws a row per action with its date and owner",
                       {"Action1", "When3", "Owner3", "StepNo3"} <= set(Names[4]), str(sorted(Names[4])))
            sys.path.insert(0, HERE)
            from _measure import kMeasure
            from _theme import kTheme
            Title = Slides[5].shapes.title
            Size = Title.text_frame.paragraphs[0].runs[0].font.size.pt
            Major = kTheme(Presentation(Deck).slide_master).Major  # measured as the builder measures, with the
            Faces = [F for F in (Major, kMeasure.Substitute(Major)) if F]  # fonts this machine has (or a CI's)
            Wraps = [kMeasure.LineWords(Title.text_frame.text, F, Size, Title.width / 12700 - 14.4, True) for F in Faces]
            self.Check("build_deck: a two-line title is balanced (no lone last word) in its font and its substitute",
                       all(len(L) < 2 or len(L[-1]) >= 2 for L in Wraps), str(list(zip(Faces, Wraps))))
            self.Check("build_deck: the closing decision has its owner and date line",
                       {"DecisionBox", "OwnerValue", "ByValue"} <= set(Names[6]), str(sorted(Names[6])))

            # lint: unwanted wraps in single-line roles and title widows on a hand-made deck
            Hand = Presentation()
            Hand.slide_width, Hand.slide_height = Pt(1440), Pt(810)
            Slide = self.AddSlide(Hand, Hand.slide_layouts[5], "Revenue grew in every single quarter of the year")
            Slide.shapes.title.left, Slide.shapes.title.width = Pt(80), Pt(1000)
            Slide.shapes.title.top, Slide.shapes.title.height = Pt(40), Pt(160)
            for Para in Slide.shapes.title.text_frame.paragraphs:
                for Run in Para.runs:
                    Run.font.size = Pt(54)
            self.AddText(Slide, "112 %", 120, Left=100, Top=300, Width=260, Height=170, Name="Value1")
            self.AddText(Slide, "MEDIUM", 26, Left=600, Top=300, Width=90, Height=44, Name="Chip1")
            Bad = os.path.join(Tmp, "wraps.pptx")
            Hand.save(Bad)
            Code, Out = self.RunScript("lint_deck.py", Bad, "--json")
            try:
                Found = [(I["code"], I.get("shape")) for I in json.loads(Out)["findings"]]
            except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: output that is not the expected JSON is a FAIL here
                Found = [(f"no JSON: {e}", None)]
            Codes = [C for C, _ in Found]
            self.Check("lint_deck: unwanted_wrap flags a KPI value and a chip that wrap in their one-line boxes",
                       Codes.count("unwanted_wrap") >= 2, str(Found))
            self.Check("lint_deck: title_widow flags a title that leaves one word on its last line",
                       "title_widow" in Codes, str(Found))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckDesignSystem")
            return

    def LintFindings(self, Path):
        """lint_deck.py --json on Path: a list of (slide, code, shape, severity)."""
        if kS.ErrorMode:
            return []
        try:
            Code, Out = self.RunScript("lint_deck.py", Path, "--json")
            try:
                Items = json.loads(Out)["findings"]
            except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: output that is not the expected JSON is a FAIL here
                return [(0, f"no JSON: {e}: {Out[:200]}", None, "error")]
            return [(I["slide"], I["code"], I.get("shape"), I["severity"]) for I in Items]
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSelfTest.LintFindings({Path})")
            return []

    @staticmethod
    def NotesOf(Slide):
        """A slide's speaker notes as text."""
        if kS.ErrorMode:
            return ""
        try:
            return Slide.notes_slide.notes_text_frame.text if Slide.has_notes_slide else ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.NotesOf")
            return ""

    @staticmethod
    def TrendLegible(Shapes, Notes, Prefix):
        """True when a tile trend is drawn with bars at least 72 pt tall (its tallest bar is the bar area), or
        is not drawn and its TREND line is in the notes - never tiny bars."""
        if kS.ErrorMode:
            return False
        try:
            Bars = [Shapes[N].height for N in Shapes if N.startswith(f"{Prefix}Bar")]
            if Bars:
                return max(Bars) >= Pt(72) - Pt(1)
            return "TREND (" in Notes
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.TrendLegible")
            return False

    @staticmethod
    def MinSize(Shapes, Names):
        """The smallest run size (pt) in the named text shapes; 0 when none has a size."""
        if kS.ErrorMode:
            return 0
        try:
            Sizes = [R.font.size.pt for N in Names if N in Shapes for P in Shapes[N].text_frame.paragraphs
                     for R in P.runs if R.font.size is not None and R.text.strip()]
            return min(Sizes) if Sizes else 0
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.MinSize")
            return 0

    def CheckDataPresentation(self):
        """Data shown as data: a before/after becomes a chart, tile trends carry their figures, the ask has its
        number beside it, a share gets a dot grid, a statement gets its points, vague targets are flagged."""
        if kS.ErrorMode:
            return
        try:
            Tmp, Edge = self.Tmp, self.Edge
            Spec = {"direction": "clean-corporate", "sources": ["Self-test brief"], "slides": [
                {"id": "cover", "pattern": "title", "title": "Northwind: revenue up 44 % in 2026",
                 "notes": {"key_fact": "Cover."}},
                {"id": "churn", "pattern": "kpi", "title": "Churn fell by a third in 2026",
                 "notes": {"key_fact": "Churn 1.8 to 1.2.", "assumptions": ["Start and end are Q1 and Q4."]},
                 "metrics": [{"value": "1.2 %", "label": "Monthly churn", "trend": [1.8, 1.2],
                              "trend_labels": ["Start of 2026", "End of 2026"]},
                             {"value": "112 %", "label": "Net retention"},
                             {"value": "60", "label": "New customers"}]},
                {"id": "summary", "pattern": "kpi", "title": "Every growth metric improved",
                 "notes": {"key_fact": "Summary.", "sources": ["brief"]},
                 "metrics": [{"value": "5.9 M", "label": "Q4 revenue", "trend": [4.1, 4.6, 5.2, 5.9],
                              "trend_labels": ["Q1", "Q4"]},
                             {"value": "112 %", "label": "Net retention"},
                             {"value": "1.2 %", "label": "Churn", "trend": [1.8, 1.2]}],
                 "decision": "Approve three sales engineers, 2.4 MSEK a year"},
                {"id": "share", "pattern": "big_number", "title": "41 % of engineers report burnout",
                 "number": "41", "unit": "%", "caption": "Burnout survey, engineering",
                 "points": ["Burnout drives the retention risk"], "notes": {"key_fact": "41 %."}},
                {"id": "ratio", "pattern": "big_number", "title": "We hired 14 of 20 planned roles",
                 "number": "14/20", "caption": "Six roles are still open.", "notes": {"key_fact": "14/20."}},
                {"id": "what", "pattern": "statement", "title": "Phishing fakes someone you trust",
                 "support": "Its goal: a click, a payment or a password.",
                 "points": ["Email", "Text and phone", "Chat and social"], "notes": {"key_fact": "What."}},
                {"id": "chartless", "pattern": "kpi_chart", "title": "New customers rose to 60 a quarter",
                 "notes": {"key_fact": "From a trend."},
                 "metrics": [{"value": "60", "label": "Q4 new customers", "trend": [38, 41, 52, 60],
                              "trend_labels": ["Q1", "Q2", "Q3", "Q4"]},
                             {"value": "191", "label": "New in 2026"}]},
                {"id": "success", "pattern": "metrics", "title": "We judge the pilot on three numbers",
                 "notes": {"key_fact": "Targets."},
                 "rows": [{"metric": "Burnout", "baseline": "41 %", "target": "Clearly lower"},
                          {"metric": "Delivery", "baseline": "Q4", "target": ">= 95 % of Q4"}]},
                {"id": "ask", "pattern": "statement", "title": "Add sales capacity to keep growing",
                 "notes": {"key_fact": "The ask."}, "decision": "Approve 3 extra sales engineers (2.4 MSEK)",
                 "owner": "Executive team", "date": "Today", "support": "Enterprise deals wait on technical help.",
                 "figure": {"value": "60", "label": "New customers in Q4", "trend": [38, 41, 52, 60],
                            "trend_labels": ["Q1", "Q4"]}},
                {"id": "bars", "pattern": "kpi", "title": "Revenue and churn both moved the right way",
                 "notes": {"key_fact": "Tile trends."}, "trend_chart": False,
                 "metrics": [{"value": "5.9 M", "label": "Q4 revenue", "trend": [4.1, 4.6, 5.2, 5.9],
                              "trend_labels": ["Q1", "Q4"]},
                             {"value": "112 %", "label": "Net retention"},
                             {"value": "1.2 %", "label": "Churn", "trend": [1.8, 1.2]}]}]}
            Path = os.path.join(Edge, "data.json")
            self.WriteJson(Spec, Path)
            Deck = os.path.join(Tmp, "data.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck)
            self.Check("build_deck: the data-presentation deck builds (exit 0)", Code == 0, Out[-600:])
            Found = self.LintFindings(Deck)
            Bad = [f"{Sl}:{C}" for Sl, C, _, Sev in Found if Sev in ("error", "warn") and C != "target_not_measurable"]
            self.Check("build_deck: the data-presentation deck lints clean (apart from the vague target)", not Bad,
                       ", ".join(Bad))
            if not os.path.exists(Deck):
                return
            Slides = list(Presentation(Deck).slides)
            Names = [{Shape.name: Shape for Shape in Slide.shapes} for Slide in Slides]
            Churn = Names[1]
            Cats = list(Churn["Chart"].chart.plots[0].categories) if "Chart" in Churn else []
            self.Check("build_deck: a kpi metric with a before/after trend becomes a labelled chart in its own card",
                       {"ChartCard", "LeadValue", "LeadLabel", "Chart"} <= set(Churn)
                       and Cats == ["Start of 2026", "End of 2026"]
                       and Churn["Chart"].chart.plots[0].has_data_labels, f"{sorted(Churn)} {Cats}")
            Summary = Names[2]
            Moved = [Summary[N].text_frame.text if N in Summary else "" for N in ("Note1", "Note3")]
            Notes2 = self.NotesOf(Slides[2])
            self.Check("build_deck: tile trends above a decision are legible bars or, with no room, the tile's note "
                       "('Q1 4.1 → Q4 5.9') and a TREND line in the notes",
                       self.TrendLegible(Summary, Notes2, "Trend1") and self.TrendLegible(Summary, Notes2, "Trend3")
                       and ("Trend1Bar1" in Summary or Moved == ["Q1 4.1 → Q4 5.9", "1.8 → 1.2"]), str(Moved))
            Summary = Names[9] if len(Names) > 9 else {}
            Figures = [Summary[N].text_frame.text for N in ("Trend1First", "Trend1Last", "Trend3First", "Trend3Last")
                       if N in Summary]
            self.Check("build_deck: tile trends carry their first and last values and periods, at the label floor or "
                       "above", Figures == ["4.1", "5.9", "1.8", "1.2"] and {"Trend1From", "Trend1To"} <= set(Summary)
                       and self.MinSize(Summary, [N for N in Summary if N.startswith(("Trend", "Label"))]) >= 24,
                       f"{Figures} {self.MinSize(Summary, list(Summary))}")
            Dots = [N for N in Names[3] if N.startswith("Share") and N[5:].isdigit()]
            Filled = [N for N in Dots if Names[3][N].fill.fore_color.theme_color == MSO_THEME_COLOR.ACCENT_1]
            self.Check("build_deck: a big number that is a share gets a dot grid (41 % = 41 of 100 dots) and its points",
                       len(Dots) == 100 and len(Filled) == 41 and "Points" in Names[3] and "ShareLabel" in Names[3],
                       f"{len(Dots)} dots, {len(Filled)} filled")
            Ratio = [N for N in Names[4] if N.startswith("Share") and N[5:].isdigit()]
            self.Check("build_deck: '14/20' draws 20 dots", len(Ratio) == 20, str(len(Ratio)))
            self.Check("build_deck: a statement's points become numbered cards beside it",
                       {"StatementPoint1", "StatementPoint3", "PointCard2", "Support"} <= set(Names[5]),
                       str(sorted(Names[5])))
            self.Check("build_deck: kpi_chart without series charts the lead metric's trend",
                       "Chart" in Names[6] and list(Names[6]["Chart"].chart.plots[0].categories) == ["Q1", "Q2", "Q3",
                                                                                                     "Q4"],
                       str(sorted(Names[6])))
            self.Check("build_deck: a metric target that is not measurable is a spec warning, and lint's "
                       "target_not_measurable", "spec warning: slide 8 (metrics): target 'Clearly lower'" in Out
                       and (8, "target_not_measurable") in [(Sl, C) for Sl, C, _, _ in Found]
                       and sum(1 for Sl, C, _, _ in Found if C == "target_not_measurable") == 1, Out[-400:])
            Ask = Names[8]
            Box, Card = Ask.get("DecisionBox"), Ask.get("Card1")
            Beside = Box is not None and Card is not None and Card.left > Box.left + Box.width and \
                abs(Card.top - Box.top) < Pt(2)
            self.Check("build_deck: a decision with a figure: the ask in its box, the number and its trend beside it, "
                       "owner and date inside the box", Beside and {"Value1", "Trend1Last", "OwnerValue", "ByValue"}
                       <= set(Ask) and Ask["OwnerValue"].top < Box.top + Box.height, str(sorted(Ask)))
            Title = Slides[0].shapes.title.text_frame.text
            self.Check("build_deck: number and unit glued in a title too ('44 %')", "44 %" in Title, repr(Title))
            Notes = [self.NotesOf(Sl) for Sl in Slides]
            self.Check("build_deck: notes 'assumptions' become an ASSUMPTIONS: section",
                       "ASSUMPTIONS:\n- Start and end are Q1 and Q4." in Notes[1], Notes[1])
            self.Check("build_deck: deck-level 'sources' fill slides without their own; a slide's own 'sources' "
                       "(even just 'brief') wins; no figure_without_source", "SOURCES:\n- Self-test brief" in Notes[1]
                       and "Self-test brief" not in Notes[2] and "SOURCES:\n- brief" in Notes[2]
                       and "Self-test brief" not in Notes[0]
                       and not any(C == "figure_without_source" for _, C, _, _ in Found), str(Notes[:3]))
            Code, Out = self.RunScript("build_deck.py", Path, "--plan")
            self.Check("build_deck --plan: shows assumptions and the deck's sources",
                       "ASSUMES: Start and end are Q1 and Q4." in Out and "Sources (every slide" in Out, Out[:500])
            Gap = {"slides": [{"id": "t", "pattern": "timeline", "title": "Six months with a mid-point check",
                               "notes": "n", "events": [{"date": "Jan", "label": "Start"}, {"date": "Feb", "label": "Survey"},
                                                        {"date": "Mar", "label": "Review"}, {"date": "May", "label": "Survey"},
                                                        {"date": "Jun", "label": "Report"}]}]}
            GapPath = os.path.join(Edge, "gap.json")
            self.WriteJson(Gap, GapPath)
            Code, Out = self.RunScript("build_deck.py", GapPath, "--plan")
            self.Check("build_deck: a monthly timeline that skips a month is a spec warning (exit unchanged)",
                       Code == 0 and "jump from Mar to May" in Out, Out[-300:])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckDataPresentation")
            return

    def CheckRoughEdges(self):
        """The rough edges of benchmark round 3: kicker on a two-line title, a measured headline check, silent
        overflow in risk cards and the email frame, fix_deck's bare call, tile text below the label floor."""
        if kS.ErrorMode:
            return
        try:
            Tmp, Edge = self.Tmp, self.Edge
            Long = ("Sales engineers join technical evaluations earlier, cut the proof-of-concept time and brief "
                    "every account team")
            Spec = {"direction": "consulting-blue", "slides": [
                {"id": "two", "pattern": "bullets", "kicker": "Quick check", "notes": "n",
                 "title": "A text says your parcel is held: pay two euros at the link. What now?",
                 "items": ["Delete it", "Report it"]},
                {"id": "risks", "pattern": "risks", "title": "Three risks threaten growth", "notes": "n",
                 "risks": [{"risk": "Enterprise sales cycles are lengthening", "likelihood": "high", "impact": "high", "mitigation": Long},
                           {"risk": "Outage", "likelihood": "low", "impact": "high", "mitigation": Long},
                           {"risk": "Hiring behind plan", "likelihood": "high", "impact": "medium",
                            "mitigation": Long}]},
                {"id": "mail", "pattern": "email", "title": "This email shows five red flags", "notes": "n",
                 "from": "IT Support <it-helpdesk@micros0ft-secure.com>", "to": "you@company.example",
                 "subject": "URGENT: Your mailbox will be deleted in 24 hours", "attachment": "Verify.html",
                 "body": ["Dear user,", "Unusual activity was found on your account today. Verify today or your "
                          "mailbox is suspended and every message in it is deleted for good.",
                          "https://micros0ft-secure.com/verify", "Thank you for your cooperation in this matter.",
                          "IT Support Team, Global Service Desk, Building 4"],
                 "callouts": [{"target": "from", "note": "A zero"}, {"target": "subject", "note": "Urgency"}]},
                {"id": "tiles", "pattern": "kpi", "title": "Every growth metric improved", "notes": "n",
                 "metrics": [{"value": "5.9 M", "label": "Quarterly revenue in dollars", "trend": [4.1, 4.6, 5.2, 5.9],
                              "trend_labels": ["Q1", "Q4"], "note": "Up from 4.1 M in the first quarter of 2026"},
                             {"value": "112 %", "label": "Net revenue retention 2026",
                              "note": "Existing customers spend more each quarter"},
                             {"value": "1.2 %", "label": "Monthly churn, all plans", "trend": [1.8, 1.2],
                              "trend_labels": ["Start", "Now"], "note": "Down from 1.8 % at the start of 2026"},
                             {"value": "60", "label": "New customers in Q4 2026", "trend": [38, 41, 52, 60],
                              "note": "Up from 38 in the first quarter of 2026"}],
                 "decision": "Approve three extra sales engineers to keep pace with enterprise demand"}]}
            Path = os.path.join(Edge, "edges.json")
            self.WriteJson(Spec, Path)
            Deck = os.path.join(Tmp, "edges.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck, "--no-auto")
            Found = self.LintFindings(Deck)
            Codes = [(Sl, C) for Sl, C, _, _ in Found]
            Slides = list(Presentation(Deck).slides) if os.path.exists(Deck) else []
            Names = [{Shape.name: Shape for Shape in Slide.shapes} for Slide in Slides]
            Kick, Title = (Names[0].get("Kicker"), Slides[0].shapes.title) if Slides else (None, None)
            self.Check("build_deck: a two-line title leaves room for its kicker (kicker ends above the title box, no "
                       "kicker_title_overlap)", Kick is not None and Kick.top + Kick.height <= Title.top
                       and (1, "kicker_title_overlap") not in Codes and (1, "headline_too_long") not in Codes, str(Codes))
            self.Check("build_deck: overflowing risk mitigations and email body are reported (exit 3, fit: for each)",
                       Code == 3 and "slide 2 (risks)" in Out and "slide 3 (mail)" in Out, Out[-600:])
            self.Check("lint_deck: text_overflow on the email paragraph that runs past the message frame",
                       any(Sl == 3 and C == "text_overflow" and (Sh or "").startswith("EmailBody")
                           for Sl, C, Sh, _ in Found), str(Found))
            Small = self.MinSize(Names[3], [N for N in Names[3] if N.startswith(("Label", "Note", "Trend"))]) \
                if len(Names) > 3 else 0
            Moved = self.NotesOf(Slides[3]) if len(Slides) > 3 else ""
            self.Check("build_deck: tile labels, notes and trend figures stay at the 24 pt label floor; what cannot "
                       "fit is reported or moved to the notes", Small >= 24 and ("label floor" in Out or "TREND (" in Moved),
                       f"{Small} pt; {Out[-300:]}")

            # lint on a hand-made deck: a kicker over a two-line title, a 3-line title, a tiny tile label
            Hand = Presentation()
            Hand.slide_width, Hand.slide_height = Pt(1440), Pt(810)
            Slide = self.AddSlide(Hand, Hand.slide_layouts[5], "Revenue grew in every single quarter of the year "
                                                               "and the team doubled")
            T = Slide.shapes.title
            T.left, T.top, T.width, T.height = Pt(80), Pt(56), Pt(1280), Pt(100)
            T.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
            for Run in T.text_frame.paragraphs[0].runs:
                Run.font.size = Pt(54)
            self.AddText(Slide, "RESULTS", 24, Left=80, Top=40, Width=600, Height=34, Name="Kicker")
            self.AddText(Slide, "Q4 revenue", 18, Left=100, Top=400, Width=300, Height=40, Name="Label1")
            Slide2 = self.AddSlide(Hand, Hand.slide_layouts[5], "Enterprise deals now close more slowly than in any "
                                                                "quarter before, and the pipeline is thinner than "
                                                                "planned at the start of the year")
            for Run in Slide2.shapes.title.text_frame.paragraphs[0].runs:
                Run.font.size = Pt(54)
            Slide2.shapes.title.width = Pt(1280)
            Slide3 = self.AddSlide(Hand, Hand.slide_layouts[5], "Revenue grew in every single quarter of 2026")
            for Run in Slide3.shapes.title.text_frame.paragraphs[0].runs:
                Run.font.size = Pt(54)
            Slide3.shapes.title.width = Pt(1280)
            Bad = os.path.join(Tmp, "edges-hand.pptx")
            Hand.save(Bad)
            Hand = [(Sl, C) for Sl, C, _, _ in self.LintFindings(Bad)]
            self.Check("lint_deck: kicker_title_overlap when a kicker runs into a two-line title",
                       (1, "kicker_title_overlap") in Hand, str(Hand))
            self.Check("lint_deck: tile_text_below_floor for an 18 pt label in a tile", (1, "tile_text_below_floor")
                       in Hand, str(Hand))
            self.Check("lint_deck: headline_too_long measures the title (3 lines: yes; a 45-character one-liner: no)",
                       (2, "headline_too_long") in Hand and (3, "headline_too_long") not in Hand, str(Hand))

            Code, Out = self.RunScript("fix_deck.py")
            Lines = [L for L in Out.strip().splitlines() if L.strip()]
            self.Check("fix_deck: a bare call prints a one-line usage hint and exits 2",
                       Code == 2 and len(Lines) == 1 and all(X in Lines[0] for X in ("--out", "--in-place", "--dry-run")),
                       Out)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckRoughEdges")
            return

    def CheckFigures(self):
        """Computed figures (benchmark round 4: a QBR said 21.8 MUSD for quarters that sum to 19.8): derived
        figures in the text are checked against the deck's data (spec warning: figure, lint figure_mismatch), and
        figure tokens let the builder compute them instead."""
        if kS.ErrorMode:
            return
        try:
            Tmp, Edge = self.Tmp, self.Edge
            Quarters = ["Q1", "Q2", "Q3", "Q4"]
            Revenue = [{"name": "Revenue (MUSD)", "values": [4.1, 4.6, 5.2, 5.9]}]
            Wrong = {"direction": "consulting-blue", "sources": ["the brief"], "slides": [
                {"id": "rev", "pattern": "kpi_chart", "title": "Revenue rose every quarter, to 5.9 MUSD",
                 "notes": {"key_fact": "Revenue.", "facts": ["44 % = 5.9 / 4.1 - 1", "Q4 vs Q1 is 50 % = 5.9 / 4.1 - 1"]},
                 "metrics": [{"value": "+48 %", "label": "Q4 vs Q1", "note": "4.1 to 5.9 MUSD"},
                             {"value": "21.8", "label": "MUSD in 2026", "note": "Sum of four quarters"}],
                 "categories": Quarters, "series": Revenue},
                {"id": "hire", "pattern": "big_number", "title": "Hiring is at 14 of 20 planned roles",
                 "number": "14/20", "caption": "Six roles are still open",
                 "points": ["60 % of the hiring plan filled", "Open roles slow onboarding"], "notes": "n"},
                {"id": "cust", "pattern": "kpi_chart", "title": "New customers per quarter up 58 %", "notes": "n",
                 "metrics": [{"value": "60", "label": "In Q4"},
                             {"value": "191", "label": "In 2026", "note": "38 + 41 + 52 + 60"}],
                 "caption": "On average 47.75 new customers a quarter",
                 "categories": Quarters, "series": [{"name": "New customers", "values": [38, 41, 52, 60]}]}]}
            Path = os.path.join(Edge, "figures-wrong.json")
            self.WriteJson(Wrong, Path)
            Deck = os.path.join(Tmp, "figures-wrong.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck)
            Lines = [L for L in Out.splitlines() if L.startswith("spec warning: figure:")]
            Want = ["'21.8' reads as a total", "19.8", "'+48 %' reads as a change", "= 50 %' does not add up", "'60 %' reads as a share",
                    "70 %"]
            self.Check("build_deck: derived figures that the deck's data contradicts are spec warnings (a sum, a "
                       "change, an equation, a share; exit code unchanged)",
                       Code == 0 and all(W in Out for W in Want) and len(Lines) == 4, "\n".join(Lines) or Out[-600:])
            self.Check("build_deck: correct derived figures are not reported (191 = 38 + 41 + 52 + 60, 44 % = "
                       "5.9 / 4.1 - 1, an average, the title's 58 %)",
                       not any(X in Out for X in ("'191'", "'44 %", "'47.75'", "'58 %'", "'14'")), "\n".join(Lines))
            Found = self.LintFindings(Deck)
            self.Check("lint_deck: figure_mismatch on the built deck, from the chart's data and the tile's text",
                       (1, "figure_mismatch") in [(Sl, C) for Sl, C, _, _ in Found]
                       and any(Sl == 1 and C == "figure_mismatch" and Sh == "Value2" for Sl, C, Sh, _ in Found),
                       str(Found))

            Right = json.loads(json.dumps(Wrong))
            Right["facts"] = {"revenue_q": [4.1, 4.6, 5.2, 5.9], "roles_planned": 20}
            Right["slides"][0]["metrics"] = [{"value": "{change}", "label": "Q4 vs Q1", "note": "4.1 to 5.9 MUSD"},
                                             {"value": "{sum} MUSD", "label": "Revenue in 2026",
                                              "note": "Average {average:revenue_q} a quarter"}]
            Right["slides"][0]["notes"] = {"key_fact": "{sum:revenue_q} MUSD in 2026."}
            Right["slides"][1]["points"] = ["{share} of the {roles_planned} planned roles filled", "Open roles slow "
                                            "onboarding"]
            Path = os.path.join(Edge, "figures-right.json")
            self.WriteJson(Right, Path)
            Deck = os.path.join(Tmp, "figures-right.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck)
            Slides = list(Presentation(Deck).slides) if os.path.exists(Deck) else []
            Names = [{Shape.name: Shape for Shape in Slide.shapes} for Slide in Slides]
            Texts = [Names[0][N].text_frame.text if Slides and N in Names[0] else "" for N in ("LeadValue", "Value1",
                                                                                              "Value2", "Note2")]
            Point = Names[1]["Points"].text_frame.text if len(Names) > 1 and "Points" in Names[1] else ""
            self.Check("build_deck: figure tokens are computed - {change} +44 %, {sum} 19.8, {average:fact} 4.95, "
                       "{share} of a big number 70 %, a scalar fact; no figure warning",
                       Code == 0 and "+44 %" in Texts and "19.8 MUSD" in Texts and "4.95" in Texts[3]
                       and "70 % of the 20 planned roles" in Point.replace(" ", " ")
                       and "19.8 MUSD in 2026." in self.NotesOf(Slides[0]) and "spec warning: figure" not in Out,
                       f"{Texts} {Point!r} {Out[-400:]}")
            Code, Out = self.RunScript("build_deck.py", Path, "--plan")
            self.Check("build_deck --plan: lists the deck's facts", "Fact revenue_q: 4.1, 4.6, 5.2, 5.9 (sum 19.8)"
                       in Out, Out[:400])
            Bad = json.loads(json.dumps(Right))
            Bad["slides"][1]["caption"] = "{sum:nowhere} roles, {change}"
            Path = os.path.join(Edge, "figures-bad.json")
            self.WriteJson(Bad, Path)
            Code, Out = self.RunScript("build_deck.py", Path, "--out", os.path.join(Tmp, "figures-bad.pptx"))
            self.Check("build_deck: a token with nothing to compute from, or an unknown fact, is a spec error (exit 2)",
                       Code == 2 and "'{sum:nowhere}' names no list" in Out and "'{change}' has nothing" in Out,
                       Out[-500:])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckFigures")
            return

    def KickerTitle(self, Family, Size, Width):
        """A title that fits one line in Family but needs two in a wider face this machine has (or "", "")."""
        if kS.ErrorMode:
            return "", ""
        try:
            from _measure import kMeasure
            Words = ("Enterprise deals close faster when sales engineers join the first technical call early and "
                     "stay with every account").split()
            Last = ["it", "now", "fast", "early", "anyway", "together", "throughout", "independently"]
            for Wide in ("dejavu sans", "liberation mono", "dejavu sans mono", "freesans"):
                if Wide not in kMeasure.Index():
                    continue
                for N in range(3, len(Words)):
                    for End in Last:  # one more word of any length: the title's width moves in small steps
                        Text = " ".join(Words[:N] + [End])
                        Own = kMeasure.TextWidth(Text, Family, Size, True)
                        Other = kMeasure.TextWidth(Text, Wide, Size, True)
                        if Own <= Width * 0.94 and Other >= Width * 1.06:
                            return Text, Wide
            return "", ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.KickerTitle")
            return "", ""

    def CheckKickerRenderer(self):
        """Benchmark round 4: a title on one line in PowerPoint wrapped to two in LibreOffice (which draws a
        missing theme font with a wider face) and grew up into its kicker, while lint and the build passed. The
        builder and lint now count the lines in that face too. Built and linted in this process, with the
        substitute pinned, so the check does not depend on which fonts the machine has."""
        if kS.ErrorMode:
            return
        try:
            sys.path.insert(0, HERE)
            from _measure import kMeasure
            from build_deck import kDeckBuilder
            from lint_deck import kLintDeck
            Face = "segoe ui semibold"  # clean-corporate's heading face: not installed on Linux CI
            Title, Wide = self.KickerTitle(Face, 54, 1280 - 15 - 16)
            if not Title or kMeasure.Installed(Face):
                self.Say("(skipped the renderer-substitute kicker check: no suitable fonts on this machine)")
                return
            Spec = {"direction": "clean-corporate", "slides": [
                {"id": "k", "pattern": "bullets", "kicker": "Results", "title": Title, "items": ["One", "Two"],
                 "notes": "n"}]}
            Saved = dict(kMeasure.Substitutes)
            Results = {}
            for Mode, Sub in (("aware", Wide), ("blind", "")):
                kMeasure.Substitutes[Face] = Sub
                Out = os.path.join(self.Tmp, f"kicker-{Mode}.pptx")
                kDeckBuilder(json.loads(json.dumps(Spec))).Build(Out)
                kMeasure.Substitutes[Face] = Wide  # lint as LibreOffice will draw it
                _, Found = kLintDeck.Lint(Out, 18.0, 12)
                Results[Mode] = [I["code"] for I in (Found.Items if Found else [])]
            kMeasure.Substitutes.clear()
            kMeasure.Substitutes.update(Saved)
            self.Check("lint_deck: kicker_title_overlap when LibreOffice's wider substitute face wraps a one-line "
                       "title to two lines", "kicker_title_overlap" in Results["blind"], str(Results))
            self.Check("build_deck: the kicker leaves room for the line count of the substitute face too (no "
                       "kicker_title_overlap)", "kicker_title_overlap" not in Results["aware"]
                       and "title_widow" not in Results["aware"], str(Results))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckKickerRenderer")
            return

    def CheckRound5(self):
        """The rough edges of benchmark round 5: a tile value that wraps, content over the footer, an undated
        timeline's highlight, a decision with its reasons, body roles at the floor, and a closed stdout pipe."""
        if kS.ErrorMode:
            return
        try:
            Tmp, Edge = self.Tmp, self.Edge
            Wide = {"direction": "consulting-blue", "footer": "QBR", "slides": [
                {"id": "kpi", "pattern": "kpi", "title": "Growth accelerated every quarter", "notes": "n",
                 "metrics": [{"value": "19.8 MUSD", "label": "Revenue"}, {"value": "112 %", "label": "Retention"},
                             {"value": "1.2 %", "label": "Churn"}, {"value": "191", "label": "New customers"}],
                 "decision": "Approve 3 extra sales engineers to offset the longer enterprise sales cycle"}]}
            Path = os.path.join(Edge, "round5-wide.json")
            self.WriteJson(Wide, Path)
            Deck = os.path.join(Tmp, "round5-wide.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck, "--no-auto")
            Codes = {C for _, C, _, _ in self.LintFindings(Deck)}
            self.Check("build_deck: a tile value that would wrap in either face is unfit text (exit 3, fit:) and lint "
                       "flags it", Code == 3 and "fit: slide 1 (kpi): value does not fit on one line" in Out
                       and "unwanted_wrap" in Codes, f"exit {Code}, {sorted(Codes)}: {Out[-300:]}")
            Points = ["Burnout drives attrition in every team we asked", "Keeping engineers beats rehiring them",
                      "Six roles still open"]
            Spec = {"direction": "boardroom", "footer": "Four-day week pilot · Board proposal", "slides": [
                {"id": "why", "pattern": "big_number", "title": "41 % of our engineers report burnout", "notes": "n",
                 "number": "41", "unit": "%", "caption": "Engineers reporting burnout in our latest survey",
                 "points": Points},
                {"id": "tl", "pattern": "timeline", "title": "The board decides in July", "notes": "n",
                 "highlight": 6, "events": [{"date": M, "label": "Pulse survey across all engineering"}
                                            for M in ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul")]},
                {"id": "steps", "pattern": "timeline", "title": "Four steps take us to a decision", "notes": "n",
                 "highlight": 3, "events": [{"label": "Survey staff"}, {"label": "Run the pilot"},
                                            {"label": "Measure output"}, {"label": "Board decides"}]},
                {"id": "ask", "pattern": "statement", "title": "Approve the pilot", "notes": "n",
                 "decision": "Approve a six-month four-day-week pilot for engineering", "owner": "The board",
                 "points": ["Burnout is at 41 %", "Costs 200 kSEK, reversible", "Measured against a baseline"]},
                {"id": "what", "pattern": "statement", "title": "Phishing tricks you into helping", "notes": "n",
                 "support": "One click is enough", "points": ["A fake sender", "A link", "A deadline"]}]}
            Path = os.path.join(Edge, "round5.json")
            self.WriteJson(Spec, Path)
            Deck = os.path.join(Tmp, "round5.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck, "--lint")
            Found = self.LintFindings(Deck)
            Bad = [f"{Sl}:{C}" for Sl, C, _, Sev in Found if Sev in ("error", "warn")]
            self.Check("build_deck: a big number with points, a 7-event timeline, stages and a decision with reasons "
                       "build and lint clean (nothing on the footer)", Code == 0 and not Bad,
                       f"exit {Code}, " + ", ".join(Bad) + Out[-400:])
            self.Check("lint_deck: no grid_monotony on undated stages with a highlight",
                       not any(Sl == 3 and C == "grid_monotony" for Sl, C, _, _ in Found), str(Found))
            Slides = list(Presentation(Deck).slides) if os.path.exists(Deck) else []
            if len(Slides) < 5:
                return
            Low = []
            for Slide in Slides:
                for Shape in Slide.shapes:
                    if Shape.name in ("Footer", "PageNumber") or not Shape.has_text_frame:
                        continue
                    if (Shape.top + Shape.height) / 12700 > 756:
                        Low.append(f"{Shape.name} ends at {(Shape.top + Shape.height) / 12700:.0f}")
                    if re.fullmatch(r"Points|StatementPoint\d|Support|Caption|Reason\d", Shape.name):
                        Low += [f"{Shape.name} {R.font.size.pt:g} pt" for P in Shape.text_frame.paragraphs
                                for R in P.runs if R.font.size and R.font.size.pt < 27]
            self.Check("build_deck: no content in the footer band; points, support, captions and reasons at the "
                       "27 pt floor or above, however short", not Low, ", ".join(Low))
            Names = {Shape.name for Shape in Slides[3].shapes}
            self.Check("build_deck: a statement with a decision takes points: the ask, then numbered reason cards",
                       {"DecisionBox", "Reason1", "Reason2", "Reason3", "ReasonNo1"} <= Names, str(sorted(Names)))
            Hand = Presentation()
            Hand.slide_width, Hand.slide_height = Pt(1440), Pt(810)
            Slide = self.AddSlide(Hand, Hand.slide_layouts[5], "Short points are body text too")
            self.AddText(Slide, "Six roles open", 16, Name="Points")
            HandPath = os.path.join(Tmp, "round5-floor.pptx")
            Hand.save(HandPath)
            self.Check("lint_deck: body_below_floor sees short body roles (a 3-word point at 16 pt)",
                       any(C == "body_below_floor" for _, C, _, _ in self.LintFindings(HandPath)))
            Read, Write = os.pipe()
            os.close(Read)  # the reader is gone before the first line, as with `--plan | head -0`
            Process = subprocess.run([sys.executable, os.path.join(HERE, "build_deck.py"), Path, "--plan"],
                                     stdout=Write, stderr=subprocess.PIPE, text=True, encoding="utf-8")
            os.close(Write)
            self.Check("kShared: a closed stdout pipe (`--plan | head`) is a quiet normal end - exit 0, no error "
                       "report", Process.returncode == 0 and "ERROR" not in Process.stderr,
                       f"exit {Process.returncode}: {Process.stderr[-300:]}")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckRound5")
            return

    def CheckRound4(self):
        """The other rough edges of benchmark round 4: a two-metric kpi, the email measured by its real wrapped
        height, undated stages, a kicker on a statement with points, and --slides for the showcase step."""
        if kS.ErrorMode:
            return
        try:
            Tmp, Edge = self.Tmp, self.Edge
            Body = ["Dear user,", "Unusual activity was found on your account today.",
                    "https://micros0ft-secure.com/verify", "Thank you, IT Support"]
            Long = ("Unusual activity was found on your account today. Verify your password within 24 hours or your "
                    "mailbox is suspended and every message in it is deleted for good")
            Spec = {"direction": "teaching-friendly", "slides": [
                {"id": "two", "pattern": "kpi", "title": "Two numbers say it all", "notes": "n",
                 "metrics": [{"value": "112 %", "label": "Net retention"}, {"value": "1.2 %", "label": "Churn"}]},
                {"id": "mail", "pattern": "email", "title": "This email shows two red flags", "notes": "n",
                 "from": "IT Support <it@micros0ft-secure.com>", "subject": "URGENT: mailbox deleted in 24 hours",
                 "body": Body, "callouts": [{"target": "from", "note": "A zero"}]},
                {"id": "after", "pattern": "timeline", "title": "Your report protects everyone else", "notes": "n",
                 "highlight": 1, "events": [{"label": "IT checks the message"}, {"label": "Copies are removed"},
                                            {"label": "The sender is blocked"}, {"label": "You get a thank-you"}]},
                {"id": "what", "pattern": "statement", "kicker": "What it is", "notes": "n",
                 "title": "Phishing is a fake message that wants you to act fast",
                 "support": "It pretends to be someone you trust.",
                 "points": ["Asks you to click a link", "Asks for a password", "Arrives by email or text"]}]}
            Path = os.path.join(Edge, "round4.json")
            self.WriteJson(Spec, Path)
            Deck = os.path.join(Tmp, "round4.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck)
            Found = self.LintFindings(Deck)
            Bad = [f"{Sl}:{C}" for Sl, C, _, Sev in Found if Sev in ("error", "warn")]
            self.Check("build_deck: a two-metric kpi, an email, undated stages and a statement with points build and "
                       "lint clean (no repeated_word)", Code == 0 and not Bad, ", ".join(Bad) + Out[-300:])
            Slides = list(Presentation(Deck).slides) if os.path.exists(Deck) else []
            Names = [{Shape.name: Shape for Shape in Slide.shapes} for Slide in Slides]
            if len(Names) < 4:
                return
            self.Check("build_deck: kpi takes two metrics (two wide tiles)", {"Value1", "Value2"} <= set(Names[0])
                       and "Value3" not in Names[0], str(sorted(Names[0])))
            Stages = [N for N in Names[2] if N.startswith("Stage")]
            self.Check("build_deck: a timeline without dates draws numbered stages, no date labels",
                       len(Stages) == 4 and Names[2]["Stage2"].text_frame.text == "2"
                       and not any(N.startswith("Date") for N in Names[2]), str(sorted(Names[2])))
            Kick, Card = Names[3].get("Kicker"), Names[3].get("PointCard1")
            Clear = Kick is not None and Card is not None and Kick.left + Kick.width <= Card.left
            self.Check("build_deck: a statement with points has its kicker, above the claim and clear of the cards",
                       Clear and Kick.top + Kick.height <= Slides[3].shapes.title.top + Pt(1)
                       and Kick.text_frame.text == "WHAT IT IS", str(sorted(Names[3])))
            sys.path.insert(0, HERE)
            from _measure import kMeasure
            from _theme import kTheme
            Claim = Slides[3].shapes.title
            Size = Claim.text_frame.paragraphs[0].runs[0].font.size.pt
            Major = kTheme(Presentation(Deck).slide_master).Major  # with --ci-fonts the theme font is missing and
            Faces = [F for F in (Major, kMeasure.Substitute(Major)) if F]  # LibreOffice draws DejaVu Sans instead
            Wraps = [kMeasure.LineWords(Claim.text_frame.text, F, Size, Claim.width / 12700 - W, True)
                     for F in Faces for W in (14.4, 9.0)]
            Widows = [f"{Sl}:{C}" for Sl, C, _, Sev in Found if C == "title_widow"]
            self.Check("build_deck: the statement claim has no lone last word in its font or its substitute "
                       f"({', '.join(Faces)}), and lint agrees", all(len(L) < 2 or len(L[-1]) >= 2 for L in Wraps)
                       and not Widows, str(list(zip(Faces, Wraps))) + str(Widows))
            Short = Names[1]["EmailBody2"].height

            Spec["slides"][1]["body"] = Body[:1] + [Long, Long, Long] + Body[2:]
            self.WriteJson(Spec, Path)
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck, "--no-auto")
            Grown = {S.name: S for S in Presentation(Deck).slides[1].shapes}.get("EmailBody2") if os.path.exists(
                Deck) else None
            self.Check("build_deck: an email paragraph's box is as tall as its wrapped text (a long paragraph "
                       "overflows: exit 3; the short one fits)", Code == 3 and "slide 2 (mail)" in Out
                       and Grown is not None and Grown.height > 1.8 * Short, f"{Out[-300:]}")

            Mixed = {"slides": [{"id": "t", "pattern": "timeline", "title": "Mixed dates are a mistake", "notes": "n",
                                 "events": [{"date": "Jan", "label": "Start"}, {"label": "Middle"},
                                            {"date": "Jun", "label": "End"}]},
                                {"id": "one", "pattern": "kpi", "title": "One number", "notes": "n",
                                 "metrics": [{"value": "41 %", "label": "Burnout"}]}]}
            MixedPath = os.path.join(Edge, "mixed.json")
            self.WriteJson(Mixed, MixedPath)
            Code, Out = self.RunScript("build_deck.py", MixedPath, "--out", os.path.join(Tmp, "mixed.pptx"))
            self.Check("build_deck: a timeline with some dates missing, and a one-metric kpi, are spec errors that "
                       "name the fitting pattern", Code == 2 and "give every event a 'date', or none" in Out
                       and "'big_number'" in Out, Out[-500:])

            Show = {"direction": "clean-corporate", "footer": "Showcase", "slides": [
                {"id": "cover", "pattern": "title", "title": "A showcase deck"},
                {"id": "part", "pattern": "section", "eyebrow": "Results", "title": "What we found"},
                {"id": "a", "pattern": "bullets", "title": "The first point is made here", "items": ["One"],
                 "notes": "n"},
                {"id": "b", "pattern": "bullets", "title": "The second point is made here", "items": ["Two"],
                 "notes": "n"},
                {"id": "c", "pattern": "kpi", "title": "Too few metrics but not built", "metrics": [], "notes": "n"}]}
            ShowPath = os.path.join(Edge, "showcase.json")
            self.WriteJson(Show, ShowPath)
            ShowDeck = os.path.join(Tmp, "showcase.pptx")
            Code, Out = self.RunScript("build_deck.py", ShowPath, "--out", ShowDeck, "--slides", "1,3-4")
            Built = list(Presentation(ShowDeck).slides) if os.path.exists(ShowDeck) else []
            Sh = [{Shape.name: Shape for Shape in Slide.shapes} for Slide in Built]
            self.Check("build_deck --slides 1,3-4: builds only those (a mistake on another slide does not stop it), "
                       "with the deck's page numbers and the section's kicker",
                       Code == 0 and len(Built) == 3 and Sh[1]["PageNumber"].text_frame.text == "3"
                       and Sh[2]["PageNumber"].text_frame.text == "4" and Sh[1]["Kicker"].text_frame.text == "RESULTS",
                       Out[-400:])
            Code, Out = self.RunScript("build_deck.py", ShowPath, "--out", ShowDeck, "--slides", "2-9")
            self.Check("build_deck --slides: a range outside the deck exits 2", Code == 2 and "outside the deck" in Out,
                       Out[-300:])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckRound4")
            return

    def CheckAutoAndCheck(self):
        """Speed round: automatic fixes for unfit text (filler words, units, detail to the notes, a split), --no-auto,
        build --check (every problem of every class in one summary, sorted by slide) and --plan's pre-checks."""
        if kS.ErrorMode:
            return
        try:
            from _autofix import kAutoFix
            Tmp, Edge = self.Tmp, self.Edge
            Tidy = kAutoFix.Tidy("We really need to act in order to keep 41 percent of a total of 3 million users")
            self.Check("autofix: filler words dropped, units written short ('in order to' -> 'to', '41 percent' -> "
                       "'41 %', '3 million' -> '3 M')", Tidy == "We need to act to keep 41 % of 3 M users", Tidy)
            Kept, Ok = kAutoFix.Cut("Add three sales engineers to run technical evaluations in parallel with "
                                    "procurement reviews")
            self.Check("autofix: detail is cut at a clause boundary, never mid-phrase", Ok and Kept ==
                       "Add three sales engineers to run technical evaluations in parallel", Kept)
            Risks = [{"risk": "Longer enterprise sales cycle", "likelihood": "high", "impact": "high",
                      "mitigation": "Add three sales engineers to run technical evaluations in parallel with "
                                    "procurement reviews"},
                     {"risk": "Data-centre outage (one in Q3)", "likelihood": "medium", "impact": "high",
                      "mitigation": "Review the Q3 outage and test failover before the big renewals in spring"},
                     {"risk": "Hiring behind plan: 14 of 20", "likelihood": "high", "impact": "medium",
                      "mitigation": "Prioritise the open roles that face customers and use the referral bonus"}]
            Spec = {"direction": "consulting-blue", "footer": "QBR", "slides": [
                {"id": "kpi", "pattern": "kpi", "title": "Growth accelerated every quarter", "notes": "n",
                 "metrics": [{"value": "19.8 MUSD", "label": "Revenue"}, {"value": "112 %", "label": "Retention"},
                             {"value": "1.2 %", "label": "Churn"}, {"value": "191", "label": "New customers"}],
                 "decision": "Approve 3 extra sales engineers to offset the longer enterprise sales cycle"},
                {"id": "risks", "pattern": "risks", "title": "Three risks could slow 2027, sales cycle first",
                 "notes": "n", "risks": Risks}]}
            Path = os.path.join(Edge, "auto.json")
            self.WriteJson(Spec, Path)
            Deck = os.path.join(Tmp, "auto-off.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck, "--no-auto")
            self.Check("build_deck --no-auto: unfit text is only reported (exit 3, no auto: lines)",
                       Code == 3 and "auto:" not in Out and "fit: slide 1" in Out and "fit: slide 2" in Out, Out[-400:])
            Deck = os.path.join(Tmp, "auto.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck)
            Slides = list(Presentation(Deck).slides) if os.path.exists(Deck) else []
            Shapes = [{Sh.name: Sh for Sh in Sl.shapes} for Sl in Slides]
            Notes = self.NotesOf(Slides[1]) if len(Slides) > 1 else ""
            self.Check("build_deck: auto-fixes make the deck fit (exit 0) and print each change as an auto: line "
                       "(slide, field, what changed)", Code == 0 and "auto: slide 1 (kpi) metrics[0].value" in Out
                       and "auto: slide 2 (risks) risks[0].mitigation" in Out, Out[-600:])
            self.Check("build_deck: '19.8 MUSD' becomes '19.8 M' with the currency in the label",
                       len(Shapes) > 1 and Shapes[0]["Value1"].text_frame.text == "19.8\u00a0M"
                       and Shapes[0]["Label1"].text_frame.text == "Revenue (USD)", Out[-300:])
            self.Check("build_deck: detail cut from a card goes to the notes in full, marked MOVED FROM SLIDE:",
                       "MOVED FROM SLIDE (risks[0].mitigation): Add three sales engineers to run technical evaluations "
                       "in parallel with procurement reviews" in Notes and len(Shapes) > 1
                       and "procurement" not in Shapes[1]["Mitigation1"].text_frame.text, Notes[-300:])
            Items = [f"Reason {N} the pilot is worth running, with a short explanation" for N in range(1, 10)]
            Split = {"slides": [{"id": "why", "pattern": "bullets", "title": "Nine reasons to run the pilot",
                                 "notes": "n", "items": Items, "allow_split": True}]}
            Path = os.path.join(Edge, "split.json")
            self.WriteJson(Split, Path)
            Deck = os.path.join(Tmp, "split.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck)
            Slides = list(Presentation(Deck).slides) if os.path.exists(Deck) else []
            self.Check("build_deck: \"allow_split\": true splits 9 bullets over two slides (5 + 4), the second "
                       "'(continued)'", Code == 0 and len(Slides) == 2 and "split into two slides" in Out
                       and Slides[1].shapes.title.text_frame.text.endswith("(continued)"), Out[-300:])
            Split["slides"][0]["allow_split"] = False
            self.WriteJson(Split, Path)
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck)
            self.Check("build_deck: without allow_split a 9-bullet slide stays a spec error (exit 2)", Code == 2, Out[-200:])
            Many = {"direction": "consulting-blue", "footer": "QBR", "slides": [
                {"id": "bare", "pattern": "statement", "title": "Approve the pilot", "points": ["A", "B"]},
                {"id": "what", "pattern": "statement", "title": "Phishing tricks you into helping", "notes": "n",
                 "support": "One click is enough"},
                Spec["slides"][0]]}
            Path = os.path.join(Edge, "check.json")
            self.WriteJson(Many, Path)
            Deck = os.path.join(Tmp, "check.pptx")
            Code, Out = self.RunScript("build_deck.py", Path, "--out", Deck, "--check", "--no-auto")
            Line = [L for L in Out.splitlines() if L.startswith("CHECK-JSON ")]
            Data = json.loads(Line[-1][11:]) if Line else {}
            Found = [(P["slide"], P["code"]) for P in Data.get("problems", [])]
            self.Check("build_deck --check: one summary lists every class - spec warning, missing notes, unfit value "
                       "and the lint finding - sorted by slide, each with an edit, and exits 3",
                       Code == 3 and {(1, "missing_notes"), (2, "spec_warning"), (3, "fit_line"), (3, "unwanted_wrap")}
                       <= set(Found) and Found == sorted(Found, key=self.FirstOf)
                       and all(P.get("edit") for P in Data.get("problems", [])) and "==== check:" in Out,
                       f"exit {Code}: {Found} {Out[-300:]}")
            HasLo = bool(shutil.which("soffice") and shutil.which("pdftoppm"))
            self.Check("build_deck --check: the contact sheet is rendered when LibreOffice is installed",
                       (not HasLo) or (bool(Data.get("sheet")) and os.path.exists(Data.get("sheet", ""))),
                       str(Data.get("sheet")))
            Plan = {"slides": [{"id": "long", "pattern": "statement", "notes": "n", "points": ["A", "B"],
                                "title": "This title runs on and on about many different things at once, so that "
                                         "no slide could ever hold it in two lines at forty points"}]}
            Path = os.path.join(Edge, "plan.json")
            self.WriteJson(Plan, Path)
            Code, Out = self.RunScript("build_deck.py", Path, "--plan")
            self.Check("build_deck --plan: pre-build checks measure the title and count words before building",
                       "plan: slide 1 (long): the title" in Out and "pre-build finding" in Out, Out[-400:])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.CheckAutoAndCheck")
            return

    @staticmethod
    def FirstOf(Pair):
        """Sort key: the first element."""
        if kS.ErrorMode:
            return 0
        try:
            return Pair[0]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTest.FirstOf")
            return 0

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
            Parser.add_argument("--ci-fonts", action="store_true",
                                help="measure with only the CI runner's fonts (DejaVu, Liberation, Carlito, Caladea)")
            Args = Parser.parse_args()
            if Args.ci_fonts:  # before anything imports _measure, which reads PPTSKILL_FONT_DIRS once
                print(f"ci-fonts: {kCiFonts.Setup(tempfile.mkdtemp(prefix='pptskill-cifonts-'))} font files")
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
            Test.CheckDesignSystem()
            Test.CheckDataPresentation()
            Test.CheckRoughEdges()
            Test.CheckFigures()
            Test.CheckKickerRenderer()
            Test.CheckRound4()
            Test.CheckRound5()
            Test.CheckAutoAndCheck()
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
