"""Check slide size without PowerPoint (python-pptx; works on any OS). Exits 1 on mismatch,
so it doubles as a build gate.

Full HD = 1440 x 810 pt (= 1920 x 1080 px at 96 DPI). Common mistakes it catches:
960 x 540 pt (PowerPoint's default, 720p) and 1920 x 1080 POINTS (2560 x 1440 px).

    uvx --with python-pptx python check_slide_size.py deck.pptx [more.pptx ...] [--expect 1440x810]
"""
import argparse
import os
import re

from pptx import Presentation

from kShared import ToolReportableException, kRun, kS, kToolException

EMU_PER_PT = 12700
SIZE_PATTERN = re.compile(r"^\s*([0-9]*\.?[0-9]+)\s*x\s*([0-9]*\.?[0-9]+)\s*$")


class kSlideSizeCheck:
    """Compares a deck's slide size with the expected one."""

    @staticmethod
    def Check(Path, Want):
        """Print OK/FAIL for one deck; True when it matches."""
        if kS.ErrorMode:
            return False
        try:
            if not os.path.isfile(Path):
                raise ToolReportableException(f"not found: {Path}")
            Deck = Presentation(Path)
            Got = (Deck.slide_width / EMU_PER_PT, Deck.slide_height / EMU_PER_PT)
            Ok = all(abs(G - W) < 0.5 for G, W in zip(Got, Want))
            print(f"{'OK  ' if Ok else 'FAIL'} {Got[0]:g} x {Got[1]:g} pt  {Path}")
            return Ok
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSlideSizeCheck.Check(file={Path})")
            return False


class kCheckSlideSizeApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Parser = argparse.ArgumentParser()
            Parser.add_argument("files", nargs="+")
            Parser.add_argument("--expect", default="1440x810", help="width x height in points")
            Args = Parser.parse_args()
            Size = SIZE_PATTERN.match(Args.expect.lower())
            if not Size:
                raise ToolReportableException(f"--expect must be WIDTHxHEIGHT, got {Args.expect!r}")
            Want = (float(Size.group(1)), float(Size.group(2)))
            Fails = 0
            for File in Args.files:
                Fails += not kSlideSizeCheck.Check(File, Want)
            return 1 if Fails else 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kCheckSlideSizeApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kCheckSlideSizeApp)
