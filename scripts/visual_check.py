"""Look at rendered slides the way a reviewer does - catches what geometry lint cannot see.

    python scripts/visual_check.py renders/                 # a folder of s001.png .. (render_lo.py, render_slides.py)
    python scripts/visual_check.py renders/ --json          # one JSON document instead of text
    python scripts/visual_check.py renders/ --metrics       # also print the raw measurements per slide (tuning)

Works on pixels only (PIL + numpy), so it judges what the audience sees, whatever made the deck. Per slide:

  empty_area     an empty band across the content area (empty card interiors and pale panels count as empty: a
                 box with nothing in it is still nothing) - warn when it is tall and runs down to the footer
                 (the content stops early), info when shorter or between two elements; a big empty block beside
                 the content - info
  unbalanced     the content mass sits heavily to one side or the top - info
  crowded        almost no whitespace left in the content area - warn
  list_like      many similar one-column text rows with the same left edge: the slide reads like a bullet list - info
  deck_outlier   much emptier or fuller than the deck's median slide - info

Each finding carries a suggested spec edit. Slide 1, section/closing slides on a dark or photo background (in a
deck whose body slides are light - in a dark-themed deck dark slides are judged normally) and slides with almost
no content are treated as statement slides and only checked for crowding. Thresholds were
tuned on the round-9 benchmark decks (evals/benchmark/README.md) against the judge's comments.
Exit code 0 always (advisory); 1 only on a tool error.
"""
import argparse
import glob
import json
import os
import re

import numpy as np
from PIL import Image, ImageFilter

from kShared import ToolInputException, kRun, kS, kToolException

WORK_WIDTH = 640                 # renders are scaled to this width before measuring
GRID_COLS, GRID_ROWS = 64, 36    # occupancy grid (10 px cells at the work width)
SIDE_MARGIN = 2                  # grid columns left out at each side (slide margin, side accent bars)
INK_CONTRAST = 28                # a pixel this far (luminance) from its neighbourhood mean is ink (text, lines, detail)
INK_CELL = 0.012                 # a cell with this share of ink pixels holds content
PALE_DISTANCE = 45               # a flat cell this close to the background colour is empty (background or pale panel)
TITLE_ZONE = 0.30                # the title is looked for in the top 30 %
FOOTER_ZONE = 0.12               # the bottom 12 % is footer/page-number territory
BAND_MIN_WIDTH = 0.8             # an empty band spans at least 80 % of the content width (content stops early,
BAND_WARN = 0.27                 # card bottoms left empty); this tall (share of the slide height) -> warn,
BAND_INFO = 0.16                 # this tall -> info
SIDE_MIN_CELLS = 12              # an empty block beside the content: at least 12 x 12 cells and
SIDE_INFO = 0.40                 # this share of the content area -> info
CROWDED_VOID = 0.08              # less than this share of empty cells in the content area
BALANCE_X, BALANCE_Y = 0.22, 0.25  # centre-of-mass offset (share of the content area) that counts as lopsided
LIST_ROWS = 5                    # this many similar single-column text rows read like a list
OUTLIER_LOW, OUTLIER_HIGH = 0.35, 2.6  # occupancy versus the deck median


class kSlideMeasure:
    """The pixel measurements of one rendered slide: background, ink mask, occupancy grid, content area."""

    def __init__(self, Path):
        try:
            self.Path = Path
            self.Grey = None
            self.Ink = None
            self.Void = None
            self.Mass = None
            self.Background = (255, 255, 255)
            self.Dark = False
            self.Top = 0
            self.Bottom = GRID_ROWS
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideMeasure.__init__")

    def Load(self):
        """Read and scale the render; find background, ink and the grid. False when the image cannot be read."""
        if kS.ErrorMode:
            return False
        try:
            Img = Image.open(self.Path).convert("RGB")
            Img = Img.resize((WORK_WIDTH, round(WORK_WIDTH * Img.height / Img.width)), Image.LANCZOS)
            Rgb = np.asarray(Img, dtype=np.float32)
            self.Grey = np.asarray(Img.convert("L"), dtype=np.float32)
            Blur = np.asarray(Img.convert("L").filter(ImageFilter.BoxBlur(5)), dtype=np.float32)
            self.Ink = np.abs(self.Grey - Blur) > INK_CONTRAST
            self.Background = self.BackgroundOf(Rgb)
            Distance = np.sqrt(((Rgb - np.array(self.Background, dtype=np.float32)) ** 2).sum(axis=2))
            self.Dark = float(np.dot(self.Background, [0.299, 0.587, 0.114])) < 90
            self.BuildGrid(Distance)
            self.FindContentRows()
            return True
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSlideMeasure.Load({self.Path})")
            return False

    @staticmethod
    def BackgroundOf(Rgb):
        """The commonest colour (16-level buckets) of the slide's outer ring - the slide background."""
        if kS.ErrorMode:
            return (255, 255, 255)
        try:
            H, W = Rgb.shape[0], Rgb.shape[1]
            Band = max(2, H // 40)
            Ring = np.concatenate([Rgb[:Band].reshape(-1, 3), Rgb[-Band:].reshape(-1, 3),
                                   Rgb[:, :Band].reshape(-1, 3), Rgb[:, -Band:].reshape(-1, 3)])
            Keys = (Ring // 16).astype(np.int32)
            Packed = Keys[:, 0] * 256 + Keys[:, 1] * 16 + Keys[:, 2]
            Values, Counts = np.unique(Packed, return_counts=True)
            Best = Values[int(np.argmax(Counts))]
            Members = Ring[Packed == Best]
            return tuple(int(V) for V in np.median(Members, axis=0))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideMeasure.BackgroundOf")
            return (255, 255, 255)

    def BuildGrid(self, Distance):
        """Void: cells with no ink that are background or a pale panel. Mass: ink plus strong flat fills."""
        if kS.ErrorMode:
            return
        try:
            H, W = self.Grey.shape
            self.Void = np.zeros((GRID_ROWS, GRID_COLS), dtype=bool)
            self.Mass = np.zeros((GRID_ROWS, GRID_COLS), dtype=np.float32)
            for R in range(GRID_ROWS):
                for C in range(GRID_COLS):
                    Y0, Y1 = R * H // GRID_ROWS, (R + 1) * H // GRID_ROWS
                    X0, X1 = C * W // GRID_COLS, (C + 1) * W // GRID_COLS
                    InkShare = float(self.Ink[Y0:Y1, X0:X1].mean())
                    Fill = float(Distance[Y0:Y1, X0:X1].mean())
                    self.Void[R, C] = InkShare < INK_CELL and Fill < PALE_DISTANCE
                    self.Mass[R, C] = min(1.0, InkShare * 8) if InkShare >= INK_CELL else (
                        0.5 if Fill >= PALE_DISTANCE else 0.0)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideMeasure.BuildGrid")
            return

    def FindContentRows(self):
        """Top: the first grid row under the title block. Bottom: the first row of the footer band."""
        if kS.ErrorMode:
            return
        try:
            HasInk = ~self.Void[:, 2 * SIDE_MARGIN:GRID_COLS - 2 * SIDE_MARGIN].all(axis=1)  # side accent bars are not a title
            Limit = int(GRID_ROWS * TITLE_ZONE)
            Row = 0
            while Row < Limit and not HasInk[Row]:
                Row += 1
            while Row < Limit and HasInk[Row]:
                Row += 1
            self.Top = Row if 0 < Row < Limit else 0
            self.Bottom = GRID_ROWS - max(1, int(round(GRID_ROWS * FOOTER_ZONE)))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideMeasure.FindContentRows")
            return

    def Region(self):
        """The content area of the grid: below the title, above the footer, inside the side margin."""
        if kS.ErrorMode:
            return None
        try:
            return self.Void[self.Top:self.Bottom, SIDE_MARGIN:GRID_COLS - SIDE_MARGIN]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideMeasure.Region")
            return None


class kRectFinder:
    """The largest all-true rectangle in a boolean grid (histogram stack method)."""

    @staticmethod
    def Largest(Grid, MinH, MinW):
        """(area, row, col, height, width) of the largest True rectangle at least MinH x MinW; zeros when none."""
        if kS.ErrorMode:
            return (0, 0, 0, 0, 0)
        try:
            Rows, Cols = Grid.shape
            Heights = [0] * Cols
            Best = (0, 0, 0, 0, 0)
            for R in range(Rows):
                for C in range(Cols):
                    Heights[C] = Heights[C] + 1 if Grid[R, C] else 0
                Best = kRectFinder.BestInRow(Heights, R, MinH, MinW, Best)
            return Best
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRectFinder.Largest")
            return (0, 0, 0, 0, 0)

    @staticmethod
    def BestInRow(Heights, Row, MinH, MinW, Best):
        """Largest rectangle under one row's histogram, kept when bigger than Best."""
        if kS.ErrorMode:
            return Best
        try:
            Stack = []
            for C in range(len(Heights) + 1):
                Height = Heights[C] if C < len(Heights) else 0
                Start = C
                while Stack and Stack[-1][1] >= Height:
                    Start, Top = Stack.pop()
                    Width = C - Start
                    if Top >= MinH and Width >= MinW and Top * Width > Best[0]:
                        Best = (Top * Width, Row - Top + 1, Start, Top, Width)
                Stack.append((Start, Height))
            return Best
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRectFinder.BestInRow")
            return Best


class kSlideJudge:
    """Turns one slide's measurements into metrics and findings."""

    @staticmethod
    def Metrics(M):
        """The numbers every rule reads: occupancy, largest empty rectangle, centre of mass, list rows."""
        if kS.ErrorMode:
            return {}
        try:
            Region = M.Region()
            if Region is None or Region.size == 0:
                return {}
            Band = kRectFinder.Largest(Region, 1, int(np.ceil(Region.shape[1] * BAND_MIN_WIDTH)))
            Side = kRectFinder.Largest(Region, SIDE_MIN_CELLS, SIDE_MIN_CELLS)
            Mass = M.Mass[M.Top:M.Bottom, SIDE_MARGIN:GRID_COLS - SIDE_MARGIN]
            Total = float(Mass.sum())
            Dx, Dy = 0.0, 0.0
            if Total > 0:
                Ys, Xs = np.mgrid[0:Mass.shape[0], 0:Mass.shape[1]]
                Dx = float((Mass * (Xs + 0.5)).sum() / Total / Mass.shape[1] - 0.5)
                Dy = float((Mass * (Ys + 0.5)).sum() / Total / Mass.shape[0] - 0.5)
            return {"occupancy": round(1.0 - float(Region.mean()), 3),
                    "empty_band": round(Band[3] / float(GRID_ROWS), 3),
                    "empty_band_at": {"row": Band[1] + M.Top, "col": Band[2] + SIDE_MARGIN, "rows": Band[3],
                                      "cols": Band[4]},
                    "band_ends_content": bool(Band[3]) and Band[1] + Band[3] >= Region.shape[0],
                    "empty_side": round(Side[0] / float(Region.size), 3),
                    "empty_side_at": {"row": Side[1] + M.Top, "col": Side[2] + SIDE_MARGIN, "rows": Side[3], "cols": Side[4]},
                    "dx": round(Dx, 3), "dy": round(Dy, 3), "list_rows": kTextRows.ListRows(M),
                    "dark": M.Dark, "content_top": M.Top}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideJudge.Metrics")
            return {}

    @staticmethod
    def IsStatement(No, Metrics):
        """Slide 1, dark-background section/closing slides and near-empty slides: no layout judgement."""
        if kS.ErrorMode:
            return True
        try:
            DarkStatement = Metrics.get("dark", False) and not Metrics.get("dark_theme", False)
            return No == 1 or DarkStatement or Metrics.get("occupancy", 0) < 0.06
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideJudge.IsStatement")
            return True

    @staticmethod
    def Where(Rect):
        """'lower half', 'right side' ... for an empty rectangle in grid cells."""
        if kS.ErrorMode:
            return ""
        try:
            Cy = (Rect["row"] + Rect["rows"] / 2.0) / GRID_ROWS
            Cx = (Rect["col"] + Rect["cols"] / 2.0) / GRID_COLS
            Vertical = "top" if Cy < 0.4 else ("bottom" if Cy > 0.6 else "middle")
            Horizontal = "left" if Cx < 0.4 else ("right" if Cx > 0.6 else "centre")
            return f"{Vertical} {Horizontal}" if Horizontal != "centre" else f"{Vertical} band"
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideJudge.Where")
            return ""

    @staticmethod
    def EmptyFinding(No, Metrics):
        """empty_area for a wide empty band (warn when tall) or a big empty block beside the content; else None."""
        if kS.ErrorMode:
            return None
        try:
            Edit = ("fill the space with the message: a bigger figure, a chart or a picture, larger text in fewer or "
                    "shorter cards, or a layout that ends where the content ends")
            Band = Metrics["empty_band"]
            if Band >= BAND_INFO:
                Where = kSlideJudge.Where(Metrics["empty_band_at"])
                # warn only when the content stops early (the band runs down to the footer); a gap between two
                # elements (a quote and its attribution) is a design choice more often than a mistake
                Severity = "warn" if Band >= BAND_WARN and Metrics["band_ends_content"] else "info"
                return kVisualCheck.Finding(No, Severity, "empty_area",
                                            f"empty band across the slide ({Where}, {round(Band * 100)} % of the "
                                            "slide height, empty card interiors included)", Edit)
            if Metrics["empty_side"] >= SIDE_INFO:
                Where = kSlideJudge.Where(Metrics["empty_side_at"])
                return kVisualCheck.Finding(No, "info", "empty_area",
                                            f"empty block beside the content ({Where}, "
                                            f"{round(Metrics['empty_side'] * 100)} % of the content area)", Edit)
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideJudge.EmptyFinding")
            return None

    @staticmethod
    def Findings(No, Metrics):
        """The findings for one slide (deck-level comparison is added by kVisualCheck)."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            if not Metrics:
                return Out
            Statement = kSlideJudge.IsStatement(No, Metrics)
            if Metrics["occupancy"] > 1.0 - CROWDED_VOID:
                Out.append(kVisualCheck.Finding(No, "warn", "crowded",
                           f"almost no whitespace left ({round(Metrics['occupancy'] * 100)} % of the content area "
                           "holds content)", "cut words or items, or split the slide in two"))
            if Statement:
                return Out
            Empty = kSlideJudge.EmptyFinding(No, Metrics)
            if Empty:
                Out.append(Empty)
            elif abs(Metrics["dx"]) > BALANCE_X or Metrics["dy"] < -BALANCE_Y:
                Side = "top" if Metrics["dy"] < -BALANCE_Y else ("left" if Metrics["dx"] < 0 else "right")
                Out.append(kVisualCheck.Finding(No, "info", "unbalanced",
                           f"content mass sits to the {Side} (offset x {Metrics['dx']:+.2f}, y {Metrics['dy']:+.2f})",
                           "use a layout that spreads the content (two columns, a figure beside the text) or "
                           "let the main element span the slide"))
            if Metrics["list_rows"] >= LIST_ROWS:
                Out.append(kVisualCheck.Finding(No, "info", "list_like",
                           f"{Metrics['list_rows']} similar one-column text rows: the slide reads like a bullet list",
                           "show it as a figure instead: cards with icons, a process, a table, or one big statement "
                           "with the list in the notes"))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideJudge.Findings")
            return []


class kTextRows:
    """Text rows found from the ink mask: bands of ink rows, each with its horizontal segments."""

    @staticmethod
    def Bands(Ink, Top, Bottom):
        """(y0, y1) row bands of ink between pixel rows Top and Bottom (gaps of 3+ empty rows split bands)."""
        if kS.ErrorMode:
            return []
        try:
            Profile = Ink[Top:Bottom].sum(axis=1) > max(2, Ink.shape[1] // 200)
            Bands, Start, Gap = [], None, 0
            for Y, On in enumerate(Profile):
                if On:
                    Start = Y if Start is None else Start
                    Gap = 0
                elif Start is not None:
                    Gap += 1
                    if Gap >= 3:
                        Bands.append((Start + Top, Y - Gap + 1 + Top))
                        Start, Gap = None, 0
            if Start is not None:
                Bands.append((Start + Top, Bottom))
            return Bands
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTextRows.Bands")
            return []

    @staticmethod
    def Segments(Ink, Y0, Y1):
        """(x0, x1) horizontal ink segments of one band; gaps wider than 5 % of the width split segments."""
        if kS.ErrorMode:
            return []
        try:
            Columns = np.nonzero(Ink[Y0:Y1].any(axis=0))[0]
            if len(Columns) == 0:
                return []
            MaxGap = Ink.shape[1] * 0.05
            Segs, Start, Prev = [], int(Columns[0]), int(Columns[0])
            for X in Columns[1:]:
                if X - Prev > MaxGap:
                    Segs.append((Start, Prev))
                    Start = int(X)
                Prev = int(X)
            Segs.append((Start, Prev))
            return Segs
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTextRows.Segments")
            return []

    @staticmethod
    def ListRows(M):
        """The longest run of text rows of similar height that are each one short segment with the same left edge."""
        if kS.ErrorMode:
            return 0
        try:
            H, W = M.Ink.shape
            Top, Bottom = M.Top * H // GRID_ROWS, M.Bottom * H // GRID_ROWS
            Rows = []
            for Y0, Y1 in kTextRows.Bands(M.Ink, Top, Bottom):
                Segs = kTextRows.Segments(M.Ink, Y0, Y1)
                if len(Segs) == 1 and 4 <= Y1 - Y0 <= H * 0.12:  # a line, or a line in a bar with a side stripe
                    Rows.append((Segs[0][0], Y1 - Y0))
            return kTextRows.LongestAligned(Rows, W * 0.03)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTextRows.ListRows")
            return 0

    @staticmethod
    def LongestAligned(Rows, Tolerance):
        """Longest run of consecutive (left, height) rows whose left edge and height match the run's first row."""
        if kS.ErrorMode:
            return 0
        try:
            Best, Run, First = 0, 0, None
            for Left, Height in Rows:
                if First is not None and abs(Left - First[0]) <= Tolerance and abs(Height - First[1]) <= First[1] * 0.25:
                    Run += 1
                else:
                    First, Run = (Left, Height), 1
                Best = max(Best, Run)
            return Best
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTextRows.LongestAligned")
            return 0


class kVisualCheck:
    """Every slide of a renders folder: metrics, per-slide findings and the deck-median comparison."""

    @staticmethod
    def Finding(No, Severity, Code, Message, Edit):
        """One finding in the shape --check prints."""
        if kS.ErrorMode:
            return {}
        try:
            return {"slide": No, "severity": Severity, "code": Code, "message": f"slide {No}: {Message}",
                    "edit": f"slides[{No - 1}]: {Edit}"}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVisualCheck.Finding")
            return {}

    @staticmethod
    def SlideNumber(Path):
        """s007.png -> 7 (the last number in the name), else 0."""
        if kS.ErrorMode:
            return 0
        try:
            Found = re.findall(r"(\d+)", os.path.basename(Path))
            return int(Found[-1]) if Found else 0
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVisualCheck.SlideNumber")
            return 0

    @staticmethod
    def Files(Folder):
        """The slide renders in Folder, in slide order (contact sheets and diffs left out)."""
        if kS.ErrorMode:
            return []
        try:
            if not os.path.isdir(Folder):
                raise ToolInputException(f"renders folder not found: {Folder}")
            Files = [F for F in glob.glob(os.path.join(Folder, "*")) if F.lower().endswith((".png", ".jpg", ".jpeg"))
                     and not re.search(r"contact|sheet|diff", os.path.basename(F).lower())]
            return sorted(Files, key=kVisualCheck.SlideNumber)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVisualCheck.Files")
            return []

    @staticmethod
    def Run(Files):
        """{"slides": [{slide, file, metrics}], "findings": [...], "median_occupancy": x} for the renders."""
        if kS.ErrorMode:
            return {"slides": [], "findings": [], "median_occupancy": 0}
        try:
            Slides, Findings = [], []
            for Index, File in enumerate(Files):
                No = kVisualCheck.SlideNumber(File) or Index + 1
                M = kSlideMeasure(File)
                if not M.Load():
                    continue
                Slides.append({"slide": No, "file": File, "metrics": kSlideJudge.Metrics(M)})
            kVisualCheck.MarkDarkTheme(Slides)
            for S in Slides:
                Findings += kSlideJudge.Findings(S["slide"], S["metrics"])
            Median = kVisualCheck.AddOutliers(Slides, Findings)
            Findings.sort(key=kVisualCheck.SortKey)
            return {"slides": Slides, "findings": Findings, "median_occupancy": Median}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVisualCheck.Run")
            return {"slides": [], "findings": [], "median_occupancy": 0}

    @staticmethod
    def MarkDarkTheme(Slides):
        """Sets metrics["dark_theme"] when most slides after slide 1 are dark: then dark is the body style,
        not a section/closing cue, and dark slides get the full layout judgement (a light title slide with
        dark body slides used to leave no content slides at all - median occupancy 0)."""
        if kS.ErrorMode:
            return
        try:
            Body = [S["metrics"] for S in Slides if S["metrics"] and S["slide"] != 1]
            DarkCount = len([Metrics for Metrics in Body if Metrics.get("dark", False)])
            DarkTheme = bool(Body) and DarkCount * 2 > len(Body)
            for S in Slides:
                if S["metrics"]:
                    S["metrics"]["dark_theme"] = DarkTheme
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVisualCheck.MarkDarkTheme")
            return

    @staticmethod
    def AddOutliers(Slides, Findings):
        """deck_outlier findings for content slides far from the median occupancy; returns the median."""
        if kS.ErrorMode:
            return 0
        try:
            Content = [S for S in Slides if S["metrics"] and not kSlideJudge.IsStatement(S["slide"], S["metrics"])]
            if len(Content) < 4:
                return 0
            Median = float(np.median([S["metrics"]["occupancy"] for S in Content]))
            Flagged = {(F["slide"], F["code"]) for F in Findings}
            for S in Content:
                Ratio = S["metrics"]["occupancy"] / Median if Median else 1.0
                if Ratio < OUTLIER_LOW and (S["slide"], "empty_area") not in Flagged:
                    Findings.append(kVisualCheck.Finding(S["slide"], "info", "deck_outlier",
                                    f"much emptier than the deck's other slides ({Ratio:.1f} x the median)",
                                    "give it more substance (a figure, a number, a picture) or merge it "
                                    "with a neighbour"))
                elif Ratio > OUTLIER_HIGH and (S["slide"], "crowded") not in Flagged:
                    Findings.append(kVisualCheck.Finding(S["slide"], "info", "deck_outlier",
                                    f"much fuller than the deck's other slides ({Ratio:.1f} x the median)",
                                    "split it, or move detail to the notes"))
            return round(Median, 3)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVisualCheck.AddOutliers")
            return 0

    @staticmethod
    def SortKey(Finding):
        """Slide, then warnings before info."""
        if kS.ErrorMode:
            return (0, 0)
        try:
            return (Finding["slide"], 0 if Finding["severity"] == "warn" else 1)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVisualCheck.SortKey")
            return (0, 0)

    @staticmethod
    def AddToReport(Files, Report, SpecOf=None):
        """build_deck.py --check: add every finding to its kCheckReport (SpecOf maps a deck slide to a spec slide)."""
        if kS.ErrorMode:
            return 0
        try:
            Result = kVisualCheck.Run(Files)
            if SpecOf is None:  # the deck -> spec slide numbering the build's own findings already carry
                SpecOf = {I["slide"]: I["spec_slide"] for I in Report.Items if I.get("slide") and I.get("spec_slide")}
            for F in Result["findings"]:
                Spec = SpecOf.get(F["slide"], F["slide"]) if SpecOf else F["slide"]
                Report.Add(F["slide"], Spec, F["severity"], "visual_" + F["code"], F["message"])
                if Report.Items:
                    Report.Items[-1]["edit"] = F["edit"].replace(f"slides[{F['slide'] - 1}]", f"slides[{Spec - 1}]")
            return len(Result["findings"])
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVisualCheck.AddToReport")
            return 0


class kVisualCheckApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
            Parser.add_argument("renders", help="a folder of slide renders (s001.png ...)")
            Parser.add_argument("--json", action="store_true", help="print one JSON document")
            Parser.add_argument("--metrics", action="store_true", help="also print the measurements per slide")
            Args = Parser.parse_args()
            Result = kVisualCheck.Run(kVisualCheck.Files(Args.renders))
            if kS.ErrorMode:
                return 1
            if Args.json:
                print(json.dumps(Result, ensure_ascii=False, indent=1))
                return 0
            kVisualCheckApp.PrintText(Result, Args.metrics)
            return 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVisualCheckApp.Run")
            return 1

    @staticmethod
    def PrintText(Result, WithMetrics):
        """The findings as text, optionally preceded by one metrics line per slide."""
        if kS.ErrorMode:
            return
        try:
            if WithMetrics:
                for S in Result["slides"]:
                    M = S["metrics"]
                    print(f"  s{S['slide']:<3} occ {M.get('occupancy', 0):.2f}  band {M.get('empty_band', 0):.2f}  side {M.get('empty_side', 0):.2f}  "
                          f"dx {M.get('dx', 0):+.2f} dy {M.get('dy', 0):+.2f}  list {M.get('list_rows', 0)}  "
                          f"top {M.get('content_top', 0)}{'  dark' if M.get('dark') else ''}")
            Warn = sum(1 for F in Result["findings"] if F["severity"] == "warn")
            print(f"visual check: {len(Result['slides'])} slide(s), {Warn} warning(s), "
                  f"{len(Result['findings']) - Warn} info (median occupancy {Result['median_occupancy']})")
            for F in Result["findings"]:
                print(f"  {F['severity']:<5} {F['code']:<13} {F['message']}")
                print(f"  {'':<5} {'':<13} edit: {F['edit']}")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kVisualCheckApp.PrintText")
            return


if __name__ == "__main__":
    kRun.Main(kVisualCheckApp)
