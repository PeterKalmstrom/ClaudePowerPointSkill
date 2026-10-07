"""Pass/fail gate for words broken across lines ("constrai / n"), no image needed.

PowerPoint's layout engine uses the embedded font even when Slide.Export does not, so
TextRange.Lines() reports the real breaks. Exit code 1 if any are found.

    uvx --with pywin32 python check_word_breaks.py --file deck.pptx
"""
import argparse
import sys

from _ppt import open_deck, shape_texts, utf8_stdout

utf8_stdout()
ap = argparse.ArgumentParser()
ap.add_argument("--file", required=True)
a = ap.parse_args()

bad = 0
with open_deck(a.file) as pres:
    for s in pres.Slides:
        for sh in shape_texts(s.Shapes):
            tr = sh.TextFrame.TextRange
            lines = [tr.Lines(i, 1).Text for i in range(1, tr.Lines().Count + 1)]
            for cur, nxt in zip(lines, lines[1:]):
                if cur and nxt and cur[-1].isalpha() and nxt[0].isalpha():
                    bad += 1
                    print(f"slide {s.SlideIndex} '{sh.Name}': ...{cur[-12:]} / {nxt[:12]}...")
print(f"{bad} broken word(s)")
sys.exit(1 if bad else 0)
