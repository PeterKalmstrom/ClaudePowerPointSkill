"""Dump all slide text + speaker notes to JSON in ONE COM session (~6 s for 70 slides,
versus one slide_snapshot round trip per slide).

    uvx --with pywin32 python bulk_read.py --file deck.pptx [--slides 5-20] --out dump.json --pretty
"""
import argparse
import json

from _ppt import is_title, notes_text, open_deck, shape_texts, utf8_stdout

utf8_stdout()
ap = argparse.ArgumentParser()
ap.add_argument("--file", required=True)
ap.add_argument("--slides", help="range like 5-20 (1-based, inclusive)")
ap.add_argument("--out")
ap.add_argument("--pretty", action="store_true")
a = ap.parse_args()

with open_deck(a.file) as pres:
    n = pres.Slides.Count
    lo, hi = (map(int, a.slides.split("-")) if a.slides else (1, n))
    out = []
    for i in range(lo, min(hi, n) + 1):
        s = pres.Slides(i)
        shapes = [{"name": sh.Name, "title": is_title(sh), "text": sh.TextFrame.TextRange.Text}
                  for sh in shape_texts(s.Shapes)]
        out.append({"index": i, "name": s.Name, "hidden": bool(s.SlideShowTransition.Hidden),
                    "layout": s.CustomLayout.Name, "shapes": shapes, "notes": notes_text(s)})

text = json.dumps({"file": a.file, "slides": out}, ensure_ascii=False, indent=2 if a.pretty else None)
if a.out:
    open(a.out, "w", encoding="utf-8").write(text)
else:
    print(text)
