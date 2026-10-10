"""Judge pack: every slide of every benchmark deck, large enough to judge without opening single files.

    python evals/benchmark/make_judge_pack.py <BENCH>                     # all decks in <BENCH>/judge
    python evals/benchmark/make_judge_pack.py <BENCH> --only M-postmortem F-postmortem
    python evals/benchmark/make_judge_pack.py <BENCH> --cols 1 --rows 2   # two slides per page, even larger

Reads <BENCH>/judge/<L>-<deck>/s001.jpg ... and <BENCH>/judge/<L>-<deck>-text.txt (made when the decks are
rendered for judging) and writes <BENCH>/judge-pack/<L>-<deck>/:

  p01.png, p02.png ...   pages of 2 x 2 slides at 960 px each (a slide per quarter of a 1940 px page), every
                         slide captioned with its deck label and slide number
  p01.txt, p02.txt ...   the slide text and speaker notes of exactly the slides on that page
  index.md               the deck's slide count and its pages, the checklist a judge ticks off

and <BENCH>/judge-pack/INDEX.md over all decks. Round 10's judges worked mostly from 6-column contact sheets and
opened about ten slides each; a page here shows four slides at a size where 14 pt text, chart labels and
overlaps are readable, so judging every slide costs two or three image reads per deck. Exit 1 on a tool error.
"""
import argparse
import glob
import os
import re
import sys

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "scripts"))
from kShared import ToolInputException, kRun, kS, kToolException  # noqa: E402  (scripts/ is on the path only from here)

SLIDE_WIDTH = 960
PAD = 10
CAPTION = 40


class kJudgeDeck:
    """One deck's renders and text, cut into pages."""

    def __init__(self, Folder):
        try:
            self.Folder = Folder
            self.Name = os.path.basename(Folder.rstrip("\\/"))
            self.Files = []
            self.Texts = {}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kJudgeDeck.__init__")

    @staticmethod
    def SlideNumber(Path):
        """s007.jpg -> 7, else 0."""
        if kS.ErrorMode:
            return 0
        try:
            Found = re.findall(r"(\d+)", os.path.basename(Path))
            return int(Found[-1]) if Found else 0
        except Exception as e:
            kS.GlobalErrorHandler(e, "kJudgeDeck.SlideNumber")
            return 0

    def Load(self):
        """The renders in slide order and the text file split per slide ("## 3. Title" starts slide 3)."""
        if kS.ErrorMode:
            return False
        try:
            Files = [F for F in glob.glob(os.path.join(self.Folder, "*")) if F.lower().endswith((".png", ".jpg", ".jpeg"))]
            self.Files = sorted(Files, key=kJudgeDeck.SlideNumber)
            TextFile = self.Folder.rstrip("\\/") + "-text.txt"
            if os.path.isfile(TextFile):
                with open(TextFile, encoding="utf-8") as Handle:
                    self.Texts = kJudgeDeck.SplitText(Handle.read())
            return bool(self.Files)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kJudgeDeck.Load")
            return False

    @staticmethod
    def SplitText(Text):
        """{slide number: that slide's text and notes} from the "## N. title" sections."""
        if kS.ErrorMode:
            return {}
        try:
            Out = {}
            Parts = re.split(r"(?m)^## (\d+)\.", Text)
            for Index in range(1, len(Parts) - 1, 2):
                Out[int(Parts[Index])] = f"## {Parts[Index]}." + Parts[Index + 1].rstrip() + "\n"
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kJudgeDeck.SplitText")
            return {}

    def WritePages(self, OutFolder, Cols, Rows):
        """p01.png/p01.txt ... and index.md; returns the page count."""
        if kS.ErrorMode:
            return 0
        try:
            os.makedirs(OutFolder, exist_ok=True)
            PerPage = Cols * Rows
            Pages = [self.Files[Start:Start + PerPage] for Start in range(0, len(self.Files), PerPage)]
            Lines = [f"# {self.Name}: {len(self.Files)} slides on {len(Pages)} page(s)", ""]
            for Number, Files in enumerate(Pages, start=1):
                Base = os.path.join(OutFolder, f"p{Number:02d}")
                kJudgePage.Build(self.Name, Files, Cols, Rows).save(Base + ".png")
                Slides = [kJudgeDeck.SlideNumber(F) for F in Files]
                with open(Base + ".txt", "w", encoding="utf-8") as Handle:
                    Handle.write("\n".join(self.Texts.get(No, f"## {No}. (no text found)\n") for No in Slides))
                Lines.append(f"- [ ] p{Number:02d}.png + p{Number:02d}.txt: slides {', '.join(str(No) for No in Slides)}")
            with open(os.path.join(OutFolder, "index.md"), "w", encoding="utf-8") as Handle:
                Handle.write("\n".join(Lines) + "\n")
            return len(Pages)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kJudgeDeck.WritePages")
            return 0


class kJudgePage:
    """One page image: Cols x Rows slides at SLIDE_WIDTH, each captioned."""

    @staticmethod
    def Font():
        """A readable caption font (Arial or DejaVu when present, else Pillow's scalable default)."""
        if kS.ErrorMode:
            return None
        try:
            for Name in ("arialbd.ttf", "arial.ttf", "DejaVuSans-Bold.ttf"):
                try:
                    return ImageFont.truetype(Name, 26)
                except OSError:
                    # ERROR-SUPPRESSED-JUSTIFIED: a missing font is expected; the next one, then the default, is used
                    continue
            return ImageFont.load_default(size=26)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kJudgePage.Font")
            return None

    @staticmethod
    def Build(DeckName, Files, Cols, Rows):
        """The page image for up to Cols x Rows renders."""
        if kS.ErrorMode:
            return Image.new("RGB", (1, 1), "white")
        try:
            First = Image.open(Files[0])
            Height = round(SLIDE_WIDTH * First.height / First.width)
            Page = Image.new("RGB", (Cols * (SLIDE_WIDTH + PAD) + PAD, Rows * (Height + CAPTION + PAD) + PAD),
                             (128, 128, 128))
            Draw = ImageDraw.Draw(Page)
            Font = kJudgePage.Font()
            for Index, File in enumerate(Files):
                X = PAD + (Index % Cols) * (SLIDE_WIDTH + PAD)
                Y = PAD + (Index // Cols) * (Height + CAPTION + PAD)
                Draw.rectangle([X, Y, X + SLIDE_WIDTH - 1, Y + CAPTION - 1], fill="black")
                Draw.text((X + 10, Y + 6), f"{DeckName}  slide {kJudgeDeck.SlideNumber(File)}", fill="white", font=Font)
                Page.paste(Image.open(File).convert("RGB").resize((SLIDE_WIDTH, Height), Image.LANCZOS), (X, Y + CAPTION))
            return Page
        except Exception as e:
            kS.GlobalErrorHandler(e, "kJudgePage.Build")
            return Image.new("RGB", (1, 1), "white")


class kJudgePackApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
            Parser.add_argument("bench", help="the round's folder holding judge/")
            Parser.add_argument("--out", help="default <bench>/judge-pack")
            Parser.add_argument("--only", nargs="*", help="deck folder names (e.g. M-postmortem); default all")
            Parser.add_argument("--cols", type=int, default=2)
            Parser.add_argument("--rows", type=int, default=2)
            Args = Parser.parse_args()
            JudgeFolder = os.path.join(Args.bench, "judge")
            if not os.path.isdir(JudgeFolder):
                raise ToolInputException(f"no judge folder: {JudgeFolder}")
            Out = Args.out or os.path.join(Args.bench, "judge-pack")
            os.makedirs(Out, exist_ok=True)
            Index = ["# Judge pack: open every page of every deck", ""]
            for Folder in sorted(F for F in glob.glob(os.path.join(JudgeFolder, "*")) if os.path.isdir(F)):
                Deck = kJudgeDeck(Folder)
                if (Args.only and Deck.Name not in Args.only) or not Deck.Load():
                    continue
                Pages = Deck.WritePages(os.path.join(Out, Deck.Name), Args.cols, Args.rows)
                Index.append(f"- {Deck.Name}: {len(Deck.Files)} slides, {Pages} page(s) - {Deck.Name}/index.md")
            with open(os.path.join(Out, "INDEX.md"), "w", encoding="utf-8") as Handle:
                Handle.write("\n".join(Index) + "\n")
            print(f"judge pack: {len(Index) - 2} deck(s) -> {Out}")
            return 0 if not kS.ErrorMode else 1
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kJudgePackApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kJudgePackApp)
