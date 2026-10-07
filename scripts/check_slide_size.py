"""Check slide size without PowerPoint (python-pptx; works on any OS). Exits 1 on mismatch,
so it doubles as a build gate.

Full HD = 1440 x 810 pt (= 1920 x 1080 px at 96 DPI). Common mistakes it catches:
960 x 540 pt (PowerPoint's default, 720p) and 1920 x 1080 POINTS (2560 x 1440 px).

    uvx --with python-pptx python check_slide_size.py deck.pptx [more.pptx ...] [--expect 1440x810]
"""
import argparse
import sys

from pptx import Presentation

EMU_PER_PT = 12700
ap = argparse.ArgumentParser()
ap.add_argument("files", nargs="+")
ap.add_argument("--expect", default="1440x810", help="width x height in points")
a = ap.parse_args()

want = tuple(float(x) for x in a.expect.lower().split("x"))
fails = 0
for f in a.files:
    p = Presentation(f)
    got = (p.slide_width / EMU_PER_PT, p.slide_height / EMU_PER_PT)
    ok = all(abs(g - w) < 0.5 for g, w in zip(got, want))
    fails += not ok
    print(f"{'OK  ' if ok else 'FAIL'} {got[0]:g} x {got[1]:g} pt  {f}")
sys.exit(1 if fails else 0)
