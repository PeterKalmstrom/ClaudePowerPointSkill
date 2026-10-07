"""Compare two folders of slide renders (before/after an edit) and report what changed.

    uvx --with pillow --with numpy python scripts/diff_renders.py renders-before/ renders-after/ [--heatmaps out/]

Renders are deterministic: the same file renders pixel-identical twice, so ANY difference is a real
change. For each slide it reports the mean pixel difference, the worst cell of an 8 x 8 grid (small
text moves hide in the mean but show in one cell) and a perceptual-hash distance. Slides are paired
by file name (s001.png ...). Use it after a bulk edit to prove which slides changed - and that the
others did not. Exit code 1 if any slide changed beyond --threshold (default: any change).
"""
import argparse
import os
import sys

import numpy as np
from PIL import Image

GRID = 8


def dhash(img, size=16):
    g = np.asarray(img.convert("L").resize((size + 1, size)), dtype=np.int16)
    return (g[:, 1:] > g[:, :-1]).flatten()


def compare(a_path, b_path):
    a = Image.open(a_path).convert("RGB")
    b = Image.open(b_path).convert("RGB")
    resized = a.size != b.size
    if resized:
        b = b.resize(a.size)
    da = np.abs(np.asarray(a, dtype=np.int16) - np.asarray(b, dtype=np.int16)).mean(axis=2) / 255.0
    h, w = da.shape
    cells = [da[r * h // GRID:(r + 1) * h // GRID, c * w // GRID:(c + 1) * w // GRID].mean()
             for r in range(GRID) for c in range(GRID)]
    worst = int(np.argmax(cells))
    return {"mean_pct": round(float(da.mean()) * 100, 3), "worst_cell_pct": round(float(cells[worst]) * 100, 2),
            "worst_cell": f"r{worst // GRID + 1}c{worst % GRID + 1}",
            "hash_distance": int((dhash(a) != dhash(b)).sum()), "size_changed": resized}, da


def heatmap(diff, path):
    img = Image.fromarray(np.uint8(np.clip(diff * 4, 0, 1) * 255))
    img.convert("RGB").save(path)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("before")
    ap.add_argument("after")
    ap.add_argument("--threshold", type=float, default=0.0, help="mean %% difference that counts as changed")
    ap.add_argument("--heatmaps", help="folder for per-slide difference images of changed slides")
    a = ap.parse_args()
    def index(folder):  # pair names case-insensitively (s002.png == S002.PNG)
        return {n.lower(): os.path.join(folder, n) for n in os.listdir(folder) if n.lower().endswith((".png", ".jpg"))}
    before, after = index(a.before), index(a.after)
    names = sorted(set(before) | set(after))
    changed = 0
    for n in names:
        pa, pb = before.get(n), after.get(n)
        if not pa or not pb:
            print(f"{n:<10} {'ADDED' if pb else 'REMOVED'}")
            changed += 1
            continue
        r, diff = compare(pa, pb)
        moved = r["size_changed"] or r["mean_pct"] > a.threshold or (a.threshold == 0 and r["worst_cell_pct"] > 0)
        changed += moved
        size = "  (image size changed)" if r["size_changed"] else ""
        print(f"{n:<10} {'CHANGED' if moved else 'same':<8}{size} mean {r['mean_pct']:.3f}%  worst cell "
              f"{r['worst_cell']} {r['worst_cell_pct']:.2f}%  hash dist {r['hash_distance']}")
        if moved and a.heatmaps:
            os.makedirs(a.heatmaps, exist_ok=True)
            heatmap(diff, os.path.join(a.heatmaps, n.rsplit(".", 1)[0] + ".diff.png"))
    print(f"\n{changed} of {len(names)} slide(s) changed")
    sys.exit(1 if changed else 0)
