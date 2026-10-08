"""Pass/fail gate for words broken across lines ("constrai / n"), no image needed.

PowerPoint's layout engine uses the embedded font even when Slide.Export does not, so
TextRange.Lines() reports the real breaks. Exit code 1 if any are found.

    uvx --with pywin32 python check_word_breaks.py --file deck.pptx
"""
import argparse

from _ppt import kPpt, kPptSession
from kShared import kRun, kS, kToolException


class kWordBreakCheck:
    """Counts words PowerPoint breaks across two lines."""

    @staticmethod
    def Check(Path):
        """Print every broken word; return how many were found."""
        if kS.ErrorMode:
            return 0
        try:
            Bad = 0
            with kPptSession(Path) as Pres:
                if Pres is None:
                    return 0
                for Slide in Pres.Slides:
                    for Shape in kPpt.ShapeTexts(Slide.Shapes):
                        Range = Shape.TextFrame.TextRange
                        Lines = [Range.Lines(Index, 1).Text for Index in range(1, Range.Lines().Count + 1)]
                        for Current, Next in zip(Lines, Lines[1:]):
                            if Current and Next and Current[-1].isalpha() and Next[0].isalpha():
                                Bad += 1
                                print(f"slide {Slide.SlideIndex} '{Shape.Name}': ...{Current[-12:]} / {Next[:12]}...")
            return Bad
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kWordBreakCheck.Check(file={Path})")
            return 0


class kCheckWordBreaksApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            kPpt.Utf8Stdout()
            Parser = argparse.ArgumentParser()
            Parser.add_argument("--file", required=True)
            Args = Parser.parse_args()
            Bad = kWordBreakCheck.Check(Args.file)
            print(f"{Bad} broken word(s)")
            return 1 if Bad else 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kCheckWordBreaksApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kCheckWordBreaksApp)
