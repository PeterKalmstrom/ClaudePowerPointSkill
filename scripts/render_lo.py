"""Render slides to PNG with LibreOffice - any OS, no PowerPoint.

    python scripts/render_lo.py deck.pptx --out renders/ [--width 1280]
    -> renders/s001.png .. sNNN.png

APPROXIMATE: LibreOffice substitutes fonts and breaks lines differently from PowerPoint. Use it
to catch gross problems - text overflowing its box, overlaps, empty or broken slides, missing
images - not to sign off line breaks or exact spacing (use render_slides.py on Windows for that).
Needs `soffice` (LibreOffice) and `pdftoppm` (poppler-utils) on PATH.
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile


def render(path, out_dir, width=1280):
    for tool in ("soffice", "pdftoppm"):
        if not shutil.which(tool):
            sys.exit(f"{tool} not found on PATH (install LibreOffice / poppler-utils)")
    os.makedirs(out_dir, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="lorender-")
    # a private profile dir avoids clashing with a running LibreOffice
    subprocess.run(["soffice", f"-env:UserInstallation=file://{tmp}/profile", "--headless",
                    "--convert-to", "pdf", "--outdir", tmp, os.path.abspath(path)],
                   check=True, capture_output=True, timeout=300)
    pdf = os.path.join(tmp, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    if not os.path.exists(pdf):
        sys.exit("LibreOffice produced no PDF")
    subprocess.run(["pdftoppm", "-png", "-scale-to-x", str(width), "-scale-to-y", "-1", pdf,
                    os.path.join(tmp, "page")], check=True, timeout=300)
    files = []
    for i, page in enumerate(sorted(glob.glob(os.path.join(tmp, "page-*.png"))), 1):
        dst = os.path.join(out_dir, f"s{i:03d}.png")
        shutil.move(page, dst)
        files.append(dst)
    shutil.rmtree(tmp, ignore_errors=True)
    return files


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--out", default="renders")
    ap.add_argument("--width", type=int, default=1280)
    a = ap.parse_args()
    for f in render(a.file, a.out, a.width):
        print(f)
