"""Read a whole deck into JSON in one go - any OS, no PowerPoint (python-pptx).

    uvx --with python-pptx python scripts/read_deck.py deck.pptx                 # all slides, compact JSON
    uvx --with python-pptx python scripts/read_deck.py deck.pptx --slides 5-20 --pretty --out dump.json
    uvx --with python-pptx python scripts/read_deck.py deck.pptx --text          # just titles, text and notes

Use it before judging or editing a deck: one call gives every slide's id, layout, title, shapes (kind,
name, position in pt, text, font sizes, alt text), chart data, table cells and speaker notes - far faster
than one screenshot round-trip per slide, and enough for cross-slide reasoning (repeated claims,
contradictions, a relationship map, word counts). Positions are slide coordinates, groups included.
"""
import argparse
import json
import os
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

from kShared import ToolReportableException, kRun, kS, kToolException
from lint_deck import GROUP_TF, kLintDeck

KIND = {MSO_SHAPE_TYPE.PICTURE: "picture", MSO_SHAPE_TYPE.CHART: "chart", MSO_SHAPE_TYPE.TABLE: "table",
        MSO_SHAPE_TYPE.GROUP: "group", MSO_SHAPE_TYPE.TEXT_BOX: "text", MSO_SHAPE_TYPE.PLACEHOLDER: "placeholder",
        MSO_SHAPE_TYPE.AUTO_SHAPE: "shape", MSO_SHAPE_TYPE.MEDIA: "media", MSO_SHAPE_TYPE.LINE: "line"}
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


class kReadDeck:
    """Stateless deck reader: one dict (JSON-ready) for a whole deck."""

    @staticmethod
    def ShapeInfo(Sh):
        """One shape as a dict."""
        if kS.ErrorMode:
            return None
        try:
            L, T, W, H = kLintDeck.Rect(Sh)
            Info = {"name": Sh.name, "kind": KIND.get(Sh.shape_type, str(Sh.shape_type).split(".")[-1].lower()),
                    "box_pt": [round(L), round(T), round(W), round(H)]}
            if kLintDeck.IsTitle(Sh):
                Info["title"] = True
            if Sh.has_text_frame and Sh.text_frame.text.strip():
                Info["text"] = Sh.text_frame.text
                Info["sizes_pt"] = sorted({round(S, 1) for P in Sh.text_frame.paragraphs if P.text.strip()
                                           for S in [kLintDeck.ParaSize(Sh, P)] if S})
            if Sh.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART) or getattr(Sh, "has_chart", False):
                Info["alt"] = kLintDeck.AltText(Sh)
            if getattr(Sh, "has_chart", False) and Sh.has_chart:
                Ch = Sh.chart
                Plot = Ch.plots[0] if len(Ch.plots) else None
                Info["chart"] = {"type": str(Ch.chart_type).split(".")[-1].split(" ")[0].lower(),
                                 "categories": [str(C) for C in Plot.categories] if Plot is not None else [],
                                 "series": [{"name": S.name, "values": list(S.values)} for S in Plot.series]
                                 if Plot else []}
            if getattr(Sh, "has_table", False) and Sh.has_table:
                Info["table"] = [[C.text for C in Row.cells] for Row in Sh.table.rows]
            return Info
        except Exception as e:
            kS.GlobalErrorHandler(e, "kReadDeck.ShapeInfo")
            return None

    @staticmethod
    def Read(Path, Lo=None, Hi=None):
        """The whole deck (or slides Lo..Hi, 1-based inclusive) as a dict."""
        if kS.ErrorMode:
            return None
        try:
            Prs = Presentation(Path)
            Out = {"file": Path, "slide_size_pt": [Emu(Prs.slide_width).pt, Emu(Prs.slide_height).pt], "slides": []}
            for N, S in enumerate(Prs.slides, 1):
                if (Lo and N < Lo) or (Hi and N > Hi):
                    continue
                GROUP_TF.clear()
                Shapes = [kReadDeck.ShapeInfo(Sh) for Sh in kLintDeck.Walk(S.shapes)]
                Title = next((X.get("text", "") for X in Shapes if X.get("title")), "")
                Out["slides"].append({
                    "index": N, "id": S._element.find(f"{P_NS}cSld").get("name"),
                    "hidden": S._element.get("show") == "0", "layout": S.slide_layout.name, "title": Title,
                    "words": sum(len(X.get("text", "").split()) for X in Shapes if not X.get("title")),
                    "shapes": Shapes,
                    "notes": S.notes_slide.notes_text_frame.text if S.has_notes_slide else ""})
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kReadDeck.Read(file={Path})")
            return None


class kReadDeckApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
            Ap.add_argument("file")
            Ap.add_argument("--slides", help="range like 5-20 (1-based, inclusive)")
            Ap.add_argument("--text", action="store_true", help="plain text outline instead of JSON")
            Ap.add_argument("--pretty", action="store_true")
            Ap.add_argument("--out")
            A = Ap.parse_args()
            if not os.path.isfile(A.file):
                raise ToolReportableException(f"file not found: {A.file}")
            Lo, Hi = (map(int, A.slides.split("-")) if A.slides else (None, None))
            Data = kReadDeck.Read(A.file, Lo, Hi)
            if Data is None:
                return 1
            if A.text:
                Lines = []
                for S in Data["slides"]:
                    Lines.append(f"## {S['index']}. {S['title'] or '(no title)'}" + ("  [hidden]" if S["hidden"] else ""))
                    Lines += [X["text"] for X in S["shapes"] if X.get("text") and not X.get("title")]
                    if S["notes"]:
                        Lines.append("NOTES: " + S["notes"])
                    Lines.append("")
                Text = "\n".join(Lines)
            else:
                Text = json.dumps(Data, ensure_ascii=False, indent=2 if A.pretty else None, default=str)
            if A.out:
                with open(A.out, "w", encoding="utf-8") as Out:
                    Out.write(Text)
                print(A.out)
            else:
                if hasattr(sys.stdout, "reconfigure"):
                    sys.stdout.reconfigure(encoding="utf-8")
                print(Text)
            return 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kReadDeckApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kReadDeckApp)
