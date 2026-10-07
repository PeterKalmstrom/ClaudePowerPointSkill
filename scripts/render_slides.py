"""Render every slide to JPEG through PowerPoint's own Save As JPEG.

Use this, not Slide.Export: Slide.Export draws a FALLBACK
font for embedded-but-not-installed fonts, which hides real mid-word line breaks.
SaveCopyAs never rebinds the open presentation.

    uvx --with pywin32 python render_slides.py --file deck.pptx --out renders/
    -> renders/s001.jpg .. sNNN.jpg
"""
import argparse
import os
import shutil
import tempfile

from _ppt import norm, open_deck


def render(path, out_dir):
    out_dir = norm(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="pptrender-")
    with open_deck(path) as pres:
        count = pres.Slides.Count
        pres.SaveCopyAs(os.path.join(tmp, "saveas"), 17)  # ppSaveAsJPG -> folder of SlideN.JPG
    src = os.path.join(tmp, "saveas")
    files = []
    for i in range(1, count + 1):
        f = next((n for n in os.listdir(src) if n.lower() in (f"slide{i}.jpg",)), None)
        if f is None:
            raise RuntimeError(f"slide {i} missing from Save As JPEG output")
        dst = os.path.join(out_dir, f"s{i:03d}.jpg")
        shutil.move(os.path.join(src, f), dst)
        files.append(dst)
    shutil.rmtree(tmp, ignore_errors=True)
    assert len(files) == count
    return files


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True)
    ap.add_argument("--out", default="renders")
    a = ap.parse_args()
    for f in render(a.file, a.out):
        print(f)
