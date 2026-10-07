"""PDF export via ExportAsFixedFormat: speaker handout (slide + notes per page), audience
handout (N-up), or plain slides.

    uvx --with pywin32 python export_pdf.py --file deck.pptx --mode notes
    uvx --with pywin32 python export_pdf.py --file deck.pptx --mode handout --slides-per-page 6

pywin32 trap (runtime-confirmed 2026-10-07): calling ExportAsFixedFormat positionally lets pywin32
fill the optional PrintRange argument itself, and PowerPoint rejects it with "TypeError: The Python
instance can not be converted to a COM object". Pass a real PrintRange object from
PrintOptions.Ranges and name every argument. If that still fails, slides mode falls back to
SaveCopyAs(path, 32) (ppSaveAsPDF); notes and handout modes have no such fallback.
"""
import argparse
import os
import sys

from _ppt import norm, open_deck

HANDOUT = {1: 10, 2: 2, 3: 3, 4: 8, 6: 4, 9: 9}  # slides/page -> ppPrintOutputType
PP_FIXED_FORMAT_PDF, PP_INTENT_PRINT, PP_PRINT_ALL, PP_SAVE_AS_PDF = 2, 2, 1, 32


def export(pres, out, output_type):
    ranges = pres.PrintOptions.Ranges
    ranges.ClearAll()
    print_range = ranges.Add(1, pres.Slides.Count)
    pres.ExportAsFixedFormat(
        Path=out, FixedFormatType=PP_FIXED_FORMAT_PDF, Intent=PP_INTENT_PRINT,
        FrameSlides=0, HandoutOrder=1, OutputType=output_type, PrintHiddenSlides=0,
        PrintRange=print_range, RangeType=PP_PRINT_ALL)


ap = argparse.ArgumentParser()
ap.add_argument("--file", required=True)
ap.add_argument("--mode", choices=["slides", "notes", "handout"], default="notes")
ap.add_argument("--slides-per-page", type=int, default=6, choices=sorted(HANDOUT))
ap.add_argument("--out")
a = ap.parse_args()

output_type = {"slides": 1, "notes": 5}.get(a.mode) or HANDOUT[a.slides_per_page]
out = norm(a.out or f"{os.path.splitext(a.file)[0]}.{a.mode}.pdf")
with open_deck(a.file) as pres:
    try:
        export(pres, out, output_type)
        method = "ExportAsFixedFormat"
    except Exception as e:
        if a.mode != "slides":
            sys.exit(f"ExportAsFixedFormat failed for {a.mode} mode: {e}")
        pres.SaveCopyAs(out, PP_SAVE_AS_PDF)
        method = f"SaveCopyAs fallback (ExportAsFixedFormat failed: {e})"
print(out)
print(f"method: {method}", file=sys.stderr)
