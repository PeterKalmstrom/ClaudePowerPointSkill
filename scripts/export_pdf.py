"""PDF export via ExportAsFixedFormat: speaker handout (slide + notes per page), audience
handout (N-up), or plain slides.

    uvx --with pywin32 python export_pdf.py --file deck.pptx --mode notes
    uvx --with pywin32 python export_pdf.py --file deck.pptx --mode handout --slides-per-page 6
"""
import argparse
import os

from _ppt import norm, open_deck

HANDOUT = {1: 10, 2: 2, 3: 3, 4: 8, 6: 4, 9: 9}  # slides/page -> ppPrintOutputType
ap = argparse.ArgumentParser()
ap.add_argument("--file", required=True)
ap.add_argument("--mode", choices=["slides", "notes", "handout"], default="notes")
ap.add_argument("--slides-per-page", type=int, default=6, choices=sorted(HANDOUT))
ap.add_argument("--out")
a = ap.parse_args()

output_type = {"slides": 1, "notes": 5}.get(a.mode) or HANDOUT[a.slides_per_page]
out = norm(a.out or f"{os.path.splitext(a.file)[0]}.{a.mode}.pdf")
with open_deck(a.file) as pres:
    # Path, ppFixedFormatTypePDF=2, ppFixedFormatIntentPrint=2, FrameSlides, HandoutOrder, OutputType
    pres.ExportAsFixedFormat(out, 2, 2, 0, 1, output_type)
print(out)
