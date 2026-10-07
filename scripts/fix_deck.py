"""Repair the mechanical defects lint_deck.py finds - any OS, python-pptx. Writes a copy; never the input.

    uvx --with python-pptx --with pillow python scripts/fix_deck.py deck.pptx --out fixed.pptx
    uvx --with python-pptx --with pillow python scripts/fix_deck.py deck.pptx --out fixed.pptx --only floor,alt
    uvx --with python-pptx --with pillow python scripts/fix_deck.py deck.pptx --dry-run

Safe fixes (each one listed as it is made):
  placeholder  delete empty title/body placeholders sitting next to real content
  floor        raise sentence text below the body floor to the floor (27 pt on Full HD)
  fit          shrink text that overflows a fixed-size box, never below the floor
  aspect       un-stretch pictures by shrinking them to their true proportions inside the same box
               (nothing is cut off; use --crop-photos to crop photos to fill the box instead)
  alt          write alt text for charts and tables from their data
  palette      replace Office default chart colours with shades of the theme accent
  legend       move a top/bottom chart legend to the right
  numfmt       replace an Accounting axis format ("$-" for zero) with a plain number format
Then lint_deck.py runs on the result and prints what is left: titles, wording, pictures' alt text, colour
choices - the things that need a person (or Claude looking at the render).
"""
import argparse
import io
import os
import subprocess
import sys

from PIL import Image
from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu, Pt

from kShared import ToolReportableException, kRun, kS, kToolException
from lint_deck import BODY_MIN_WORDS, GROUP_TF, GROWN, NS, STRETCH_TOLERANCE, TEXT_PLACEHOLDERS, kLintDeck
from _measure import kMeasure
from _rules import OFFICE_DEFAULT_SERIES
from _theme import kTheme

HERE = os.path.dirname(os.path.abspath(__file__))
FIXES = ["placeholder", "floor", "fit", "aspect", "alt", "palette", "legend", "numfmt"]
RAMP = [0, 0.35, -0.3, 0.6, -0.5, 0.75]
C_MAP = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart", "a": NS["a"]}


class kFixDeck:
    """Stateless safe repairs over a python-pptx Presentation."""

    @staticmethod
    def ChartAlt(Chart):
        """Alt text for a chart, written from its data."""
        if kS.ErrorMode:
            return None
        try:
            Kind = str(Chart.chart_type).split(".")[-1].split(" ")[0].replace("_", " ").lower()
            Plot = Chart.plots[0]
            Cats = [str(C) for C in Plot.categories]
            Parts = []
            for S in Plot.series:
                Vals = ", ".join(f"{C} {V:g}" if isinstance(V, (int, float)) else f"{C} n/a"
                                 for C, V in zip(Cats, S.values))
                Parts.append(f"{S.name}: {Vals}" if len(Plot.series) > 1 or S.name else Vals)
            Title = Chart.chart_title.text_frame.text.strip() + ". " \
                if Chart.has_title and Chart.chart_title.has_text_frame else ""
            return f"{Kind.capitalize()} chart. {Title}" + "; ".join(Parts) + "."
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFixDeck.ChartAlt")
            return None

    @staticmethod
    def FitSize(Sh, Size, Floor, Family, Bold, Width, Height):
        """Largest size <= Size (never below Floor) at which the shape's text fits Width x Height."""
        if kS.ErrorMode:
            return None
        try:
            Cur = Size
            while Cur > Floor:
                Paras = [(P.text, Family, Cur, Bold, 0) for P in Sh.text_frame.paragraphs]
                Need, Widest, _ = kMeasure.TextHeight(Paras, Width)
                if Need <= Height and Widest <= Width:
                    return Cur
                Cur = max(Floor, Cur - 1)
            return Cur
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFixDeck.FitSize")
            return None

    @staticmethod
    def Fix(Prs, Only, Log, CropPhotos=False):
        """Apply the fixes named in Only; Log(message) is called for each one made."""
        if kS.ErrorMode:
            return None
        try:
            Theme = kTheme(Prs.slide_master)
            Sw = Emu(Prs.slide_width).pt
            Scale = max(1.0, Sw / 960)
            Floor, LabelMax = 18 * Scale, 12 * Scale
            Ins = kLintDeck.Inset
            for N, Slide in enumerate(Prs.slides, 1):
                if Slide._element.get("show") == "0":
                    continue
                GROUP_TF.clear()
                GROWN.clear()
                Shapes = kLintDeck.Walk(Slide.shapes)
                Content = [S for S in Shapes if S.shape_type != MSO_SHAPE_TYPE.GROUP]
                Real = [S for S in Content if not kLintDeck.IsTitle(S) and (
                    S.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART, MSO_SHAPE_TYPE.TABLE)
                    or getattr(S, "has_chart", False) or (S.has_text_frame and S.text_frame.text.strip()))]
                for S in Content:
                    Where = f"slide {N} [{S.name}]"
                    # placeholder
                    if "placeholder" in Only and S.is_placeholder and S.has_text_frame \
                            and not S.text_frame.text.strip() and kLintDeck.PhType(S) in TEXT_PLACEHOLDERS \
                            and not kLintDeck.IsTitle(S) and any(O is not S for O in Real):
                        S._element.getparent().remove(S._element)
                        Log(f"placeholder  {Where}: deleted the empty placeholder")
                        continue
                    # floor
                    if "floor" in Only and S.has_text_frame:
                        Raised = 0
                        for P in S.text_frame.paragraphs:
                            Size = kLintDeck.ParaSize(S, P)
                            if Size and LabelMax < Size < Floor and len(P.text.split()) >= BODY_MIN_WORDS:
                                for R in P.runs:
                                    R.font.size = Pt(Floor)
                                Raised += 1
                        if Raised:
                            Log(f"floor        {Where}: {Raised} paragraph(s) raised to {Floor:g} pt")
                    # fit
                    if "fit" in Only and S.has_text_frame and S.text_frame.text.strip():
                        Bp = S.text_frame._txBody.find("a:bodyPr", NS)
                        FixedBox = Bp is None or (Bp.find("a:spAutoFit", NS) is None
                                                  and Bp.find("a:normAutofit", NS) is None
                                                  and Bp.get("wrap") != "none")
                        if FixedBox:
                            L, T, W, H = kLintDeck.Rect(S)
                            Width = W - Ins(Bp, "lIns", 7.2) - Ins(Bp, "rIns", 7.2)
                            Height = H - Ins(Bp, "tIns", 3.6) - Ins(Bp, "bIns", 3.6)
                            Sizes = [kLintDeck.ParaSize(S, P) or 18.0 for P in S.text_frame.paragraphs if P.text.strip()]
                            Run = next((R for P in S.text_frame.paragraphs for R in P.runs if R.text.strip()), None)
                            Name = Run.font.name if Run is not None else None
                            Family = Theme.Font(Name) if Name else (Theme.Major if kLintDeck.IsTitle(S) else Theme.Minor)
                            Bold = bool(Run.font.bold) if Run is not None and Run.font.bold is not None \
                                else kLintDeck.IsTitle(S)
                            Size = max(Sizes) if Sizes else 18.0
                            Need, Widest, _ = kMeasure.TextHeight([(P.text, Family, kLintDeck.ParaSize(S, P) or 18.0,
                                                                    Bold, 0) for P in S.text_frame.paragraphs], Width)
                            if Need > Height * 1.08 and Width > 0:
                                New = kFixDeck.FitSize(S, Size, Floor if any(len(P.text.split()) >= 4 for P in
                                                                             S.text_frame.paragraphs) else LabelMax,
                                                       Family, Bold, Width, Height)
                                if New < Size:
                                    for P in S.text_frame.paragraphs:
                                        for R in P.runs:
                                            R.font.size = Pt(min(New, (R.font.size.pt if R.font.size else Size)))
                                    Log(f"fit          {Where}: text {Size:g} -> {New:g} pt to fit its box"
                                        + (" (still too long - cut words)" if New <= Floor else ""))
                    # crop
                    if "aspect" in Only and S.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        Stretch = kLintDeck.PictureStretch(S)
                        if abs(Stretch) > STRETCH_TOLERANCE:
                            Iw, Ih = Image.open(io.BytesIO(S.image.blob)).size
                            Box = S.width / S.height
                            VisW = Iw * (1 - S.crop_left - S.crop_right)
                            VisH = Ih * (1 - S.crop_top - S.crop_bottom)
                            Ratio = VisW / VisH
                            if CropPhotos:  # fill the box by cropping - right for photos, wrong for charts/diagrams
                                if Ratio > Box:
                                    Extra = (VisW - VisH * Box) / Iw / 2
                                    S.crop_left, S.crop_right = S.crop_left + Extra, S.crop_right + Extra
                                else:
                                    Extra = (VisH - VisW / Box) / Ih / 2
                                    S.crop_top, S.crop_bottom = S.crop_top + Extra, S.crop_bottom + Extra
                                Log(f"aspect       {Where}: was {Stretch:+.0%} stretched; cropped to fill the box "
                                    "(check the subject is still in frame)")
                            else:            # fit inside the old box at the true ratio, centred: nothing is lost
                                if Ratio < Box:
                                    NewW = int(S.height * Ratio)
                                    S.left, S.width = S.left + (S.width - NewW) // 2, NewW
                                else:
                                    NewH = int(S.width / Ratio)
                                    S.top, S.height = S.top + (S.height - NewH) // 2, NewH
                                Log(f"aspect       {Where}: was {Stretch:+.0%} stretched; resized to its true "
                                    "proportions inside the same box")
                    # alt
                    if "alt" in Only and not kLintDeck.AltText(S) and not kLintDeck.IsDecorative(S):
                        Nv = S._element.find(".//p:cNvPr", NS)
                        if getattr(S, "has_chart", False) and S.has_chart:
                            Nv.set("descr", kFixDeck.ChartAlt(S.chart))
                            Log(f"alt          {Where}: chart alt text written from its data")
                        elif getattr(S, "has_table", False) and S.has_table:
                            Head = [C.text for C in S.table.rows[0].cells]
                            Nv.set("descr", f"Table: {', '.join(Head)}; {len(S.table.rows) - 1} rows.")
                            Log(f"alt          {Where}: table alt text written from its header")
                    # chart fixes
                    if getattr(S, "has_chart", False) and S.has_chart:
                        Ch = S.chart
                        Cx = Ch._chartSpace
                        if "palette" in Only:
                            Sers = list(Ch.plots[0].series) if len(Ch.plots) else []
                            Cols = [Theme.ColourIn(X._element.find("c:spPr/a:solidFill", C_MAP)) for X in Sers[:6]]
                            if sum(1 for C in Cols if C in OFFICE_DEFAULT_SERIES) >= 2:
                                for I, Ser in enumerate(Sers):
                                    Ser.format.fill.solid()
                                    Ser.format.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
                                    Ser.format.fill.fore_color.brightness = RAMP[I % len(RAMP)]
                                Log(f"palette      {Where}: {len(Sers)} series recoloured as shades of the theme accent")
                        if "legend" in Only and Ch.has_legend:
                            Pos = Cx.find(".//c:legend/c:legendPos", C_MAP)
                            if Pos is not None and Pos.get("val") in ("t", "b"):
                                Ch.legend.position = XL_LEGEND_POSITION.RIGHT
                                Ch.legend.include_in_layout = False
                                Log(f"legend       {Where}: legend moved to the right")
                        if "numfmt" in Only:
                            for Fmt in Cx.findall(".//c:valAx/c:numFmt", C_MAP):
                                Code = Fmt.get("formatCode", "")
                                if "_(" in Code or ('"-"' in Code and "*" in Code):
                                    Fmt.set("formatCode", "$#,##0" if "$" in Code else "#,##0")
                                    Fmt.set("sourceLinked", "0")
                                    Log(f"numfmt       {Where}: Accounting axis format replaced")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFixDeck.Fix")
            return None


class kFixDeckApp:
    """Command line: fix, save a copy, then lint the copy."""

    def __init__(self):
        try:
            self._made = []
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFixDeckApp.__init__")

    def Log(self, Message):
        """Record and print one fix."""
        if kS.ErrorMode:
            return None
        try:
            self._made.append(Message)
            print(Message)
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFixDeckApp.Log")
            return None

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
            Ap.add_argument("file")
            Ap.add_argument("--out")
            Ap.add_argument("--only", help=f"comma-separated subset of: {', '.join(FIXES)}")
            Ap.add_argument("--dry-run", action="store_true", help="list the fixes without writing")
            Ap.add_argument("--crop-photos", action="store_true",
                            help="crop stretched pictures to fill their box instead of shrinking them (photos only)")
            A = Ap.parse_args()
            if not A.out and not A.dry_run:
                Ap.error("--out is required (the input is never overwritten), or use --dry-run")
            if A.out and os.path.abspath(A.out) == os.path.abspath(A.file):
                Ap.error("--out must differ from the input")
            Only = set(A.only.split(",")) if A.only else set(FIXES)
            Unknown = Only - set(FIXES)
            if Unknown:
                Ap.error(f"unknown fix(es): {', '.join(sorted(Unknown))}")
            if not os.path.isfile(A.file):
                raise ToolReportableException(f"file not found: {A.file}")
            if hasattr(sys.stdout, "reconfigure"):
                sys.stdout.reconfigure(encoding="utf-8")
            Prs = Presentation(A.file)
            kFixDeck.Fix(Prs, Only, self.Log, A.crop_photos)
            if kS.ErrorMode:
                return 1
            print(f"\n{len(self._made)} fix(es){' (dry run, nothing written)' if A.dry_run else ''}")
            if A.dry_run or not self._made and not A.out:
                return 0
            Prs.save(A.out)
            print(f"-> {A.out}\n\nWhat's left (lint_deck.py):")
            sys.stdout.flush()
            R = subprocess.run([sys.executable, os.path.join(HERE, "lint_deck.py"), A.out])
            return R.returncode
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFixDeckApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kFixDeckApp)
