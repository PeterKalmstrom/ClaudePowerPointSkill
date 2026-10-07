"""Timestamped side copy of a deck before a risky edit. No COM, so it never disturbs an open
PowerPoint session; save in PowerPoint first, because this copies what is on disk.

    python backup_snapshot.py --file deck.pptx --label pre-bulk-fontbump
    -> deck.20261007-143000.pre-bulk-fontbump.pptx (beside the deck)
"""
import argparse
import datetime
import pathlib
import shutil
import sys

ap = argparse.ArgumentParser()
ap.add_argument("--file", required=True)
ap.add_argument("--label", default="backup")
a = ap.parse_args()

src = pathlib.Path(a.file).resolve()
if not src.is_file():
    sys.exit(f"not found: {src}")
ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
dst = src.with_name(f"{src.stem}.{ts}.{a.label}{src.suffix}")
if dst.exists():  # never overwrite an earlier snapshot
    sys.exit(f"refusing to overwrite {dst}")
shutil.copy2(src, dst)
print(dst)
