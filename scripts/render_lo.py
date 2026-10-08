"""Render slides to PNG with LibreOffice - any OS, no PowerPoint.

    python scripts/render_lo.py deck.pptx --out renders/ [--width 1280]
    -> renders/s001.png .. sNNN.png, numbered by SLIDE (hidden slides are skipped, not renumbered)
    python scripts/render_lo.py deck.pptx --out renders/ --sheet [sheet.png] [--cols 4]
    -> the same, plus every slide on one contact sheet (default renders/contact.png) to look at in one go

APPROXIMATE: LibreOffice substitutes fonts and breaks lines differently from PowerPoint. Use it
to catch gross problems - text overflowing its box, overlaps, empty or broken slides, missing
images - not to sign off line breaks or exact spacing (use render_slides.py on Windows for that).
Needs `soffice` (LibreOffice) and `pdftoppm` (poppler-utils) on PATH.
"""
import argparse
import glob
import importlib.util
import os
import pathlib
import shutil
import subprocess
import tempfile

from kShared import ToolReportableException, kRun, kS, kToolException


class kLoRenderer:
    """LibreOffice -> PDF -> pdftoppm -> one PNG per visible slide."""

    @staticmethod
    def VisibleSlides(Path):
        """1-based numbers of the slides LibreOffice will export (it skips hidden ones); None without python-pptx."""
        if kS.ErrorMode:
            return None
        try:
            if importlib.util.find_spec("pptx") is None:
                return None
            from pptx import Presentation
            return [Number for Number, Slide in enumerate(Presentation(Path).slides, 1)
                    if Slide._element.get("show") != "0"]
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kLoRenderer.VisibleSlides(file={Path})")
            return None

    @staticmethod
    def PageNumber(PagePath):
        """The page number in pdftoppm's page-N.png name (sort key)."""
        if kS.ErrorMode:
            return 0
        try:
            return int(PagePath.rsplit("-", 1)[1][:-4])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kLoRenderer.PageNumber")
            return 0

    @staticmethod
    def Render(Path, OutDir, Width=1280):
        """Render Path into OutDir; returns the PNG paths."""
        if kS.ErrorMode:
            return []
        try:
            for Tool in ("soffice", "pdftoppm"):
                if not shutil.which(Tool):
                    raise ToolReportableException(f"{Tool} not found on PATH (install LibreOffice / poppler-utils)")
            os.makedirs(OutDir, exist_ok=True)
            Temp = tempfile.mkdtemp(prefix="lorender-")
            # a private profile dir avoids clashing with a running LibreOffice
            Profile = pathlib.Path(Temp, "profile").as_uri()  # file:///C:/... on Windows, file:///tmp/... elsewhere
            subprocess.run(["soffice", f"-env:UserInstallation={Profile}", "--headless",
                            "--convert-to", "pdf", "--outdir", Temp, os.path.abspath(Path)],
                           check=True, capture_output=True, timeout=300)
            Pdf = os.path.join(Temp, os.path.splitext(os.path.basename(Path))[0] + ".pdf")
            if not os.path.exists(Pdf):
                raise ToolReportableException("LibreOffice produced no PDF")
            subprocess.run(["pdftoppm", "-png", "-scale-to-x", str(Width), "-scale-to-y", "-1", Pdf,
                            os.path.join(Temp, "page")], check=True, timeout=300)
            Pages = sorted(glob.glob(os.path.join(Temp, "page-*.png")), key=kLoRenderer.PageNumber)
            Numbers = kLoRenderer.VisibleSlides(Path) or list(range(1, len(Pages) + 1))
            if len(Numbers) != len(Pages):  # unexpected: fall back to page order rather than mislabel
                Numbers = list(range(1, len(Pages) + 1))
            Files = []
            for Number, Page in zip(Numbers, Pages):
                Destination = os.path.join(OutDir, f"s{Number:03d}.png")
                shutil.move(Page, Destination)
                Files.append(Destination)
            shutil.rmtree(Temp, ignore_errors=True)  # a leftover temp folder is harmless
            return Files
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kLoRenderer.Render(file={Path})")
            return []


class kRenderLoApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Parser = argparse.ArgumentParser()
            Parser.add_argument("file")
            Parser.add_argument("--out", default="renders")
            Parser.add_argument("--width", type=int, default=1280)
            Parser.add_argument("--sheet", nargs="?", const="", default=None,
                                help="also write a contact sheet (default <out>/contact.png)")
            Parser.add_argument("--cols", type=int, default=4, help="contact sheet columns")
            Parser.add_argument("--sheet-width", type=int, default=480, help="thumbnail width on the sheet (px)")
            Args = Parser.parse_args()
            Files = kLoRenderer.Render(Args.file, Args.out, Args.width)
            for File in Files:
                print(File)
            if Args.sheet is not None and Files:
                from contact_sheet import kContactSheet  # Pillow; any OS
                Sheet = kContactSheet.Build(Files, Args.cols, Args.sheet_width,
                                            Args.sheet or os.path.join(Args.out, "contact.png"))
                if Sheet:
                    print(f"sheet: {Sheet}")
            return 1 if kS.ErrorMode else 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRenderLoApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kRenderLoApp)
