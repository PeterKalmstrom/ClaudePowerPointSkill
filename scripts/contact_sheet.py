"""Every slide as a thumbnail grid in one PNG - shows colour drift, density imbalance,
monotone-anchor stretches and forgotten template slides that per-slide review misses.

    uvx --with pywin32 --with pillow python contact_sheet.py --file deck.pptx [--cols 6] [--slide-width 240]
    -> <deck>_contact.png beside the deck
"""
import argparse
import os
import tempfile

from PIL import Image, ImageDraw

from kShared import ToolReportableException, kRun, kS, kToolException
from render_slides import kSlideRenderer

CAPTION, PAD = 18, 8


class kContactSheet:
    """Lays rendered slides out as a grid."""

    @staticmethod
    def Build(Files, Cols, SlideWidth, Out):
        """Paste the renders into one sheet and save it to Out."""
        if kS.ErrorMode:
            return None
        try:
            if not Files:
                raise ToolReportableException("no slides rendered")
            First = Image.open(Files[0])
            Width = SlideWidth
            Height = round(Width * First.height / First.width)  # aspect ratio from the deck itself
            Rows = -(-len(Files) // Cols)
            Sheet = Image.new("RGB", (Cols * (Width + PAD) + PAD, Rows * (Height + CAPTION + PAD) + PAD), "white")
            Draw = ImageDraw.Draw(Sheet)
            for Index, File in enumerate(Files):
                X = PAD + (Index % Cols) * (Width + PAD)
                Y = PAD + (Index // Cols) * (Height + CAPTION + PAD)
                Sheet.paste(Image.open(File).resize((Width, Height)), (X, Y))
                Draw.text((X, Y + Height + 2), str(Index + 1), fill="black")
            Sheet.save(Out)
            return Out
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kContactSheet.Build")
            return None


class kContactSheetApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Parser = argparse.ArgumentParser()
            Parser.add_argument("--file", required=True)
            Parser.add_argument("--cols", type=int, default=6)
            Parser.add_argument("--slide-width", type=int, default=240)
            Parser.add_argument("--out")
            Args = Parser.parse_args()
            Files = kSlideRenderer.Render(Args.file, tempfile.mkdtemp(prefix="contact-"))
            if kS.ErrorMode:
                return 1
            Out = kContactSheet.Build(Files, Args.cols, Args.slide_width,
                                      Args.out or os.path.splitext(Args.file)[0] + "_contact.png")
            if Out:
                print(Out)
            return 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kContactSheetApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kContactSheetApp)
