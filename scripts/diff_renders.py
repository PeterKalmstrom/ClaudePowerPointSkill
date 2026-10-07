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

import numpy as np
from PIL import Image

from kShared import ToolReportableException, kRun, kS, kToolException

GRID = 8


class kRenderDiff:
    """Pixel comparison of two renders."""

    @staticmethod
    def DHash(Img, Size=16):
        """Difference hash: one bit per horizontal neighbour pair."""
        if kS.ErrorMode:
            return None
        try:
            Grey = np.asarray(Img.convert("L").resize((Size + 1, Size)), dtype=np.int16)
            return (Grey[:, 1:] > Grey[:, :-1]).flatten()
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRenderDiff.DHash")
            return None

    @staticmethod
    def Compare(PathA, PathB):
        """(report dict, per-pixel difference array) for two images."""
        if kS.ErrorMode:
            return None, None
        try:
            ImgA = Image.open(PathA).convert("RGB")
            ImgB = Image.open(PathB).convert("RGB")
            Resized = ImgA.size != ImgB.size
            if Resized:
                ImgB = ImgB.resize(ImgA.size)
            Diff = np.abs(np.asarray(ImgA, dtype=np.int16) - np.asarray(ImgB, dtype=np.int16)).mean(axis=2) / 255.0
            Height, Width = Diff.shape
            Cells = [Diff[Row * Height // GRID:(Row + 1) * Height // GRID,
                          Col * Width // GRID:(Col + 1) * Width // GRID].mean()
                     for Row in range(GRID) for Col in range(GRID)]
            Worst = int(np.argmax(Cells))
            return {"mean_pct": round(float(Diff.mean()) * 100, 3),
                    "worst_cell_pct": round(float(Cells[Worst]) * 100, 2),
                    "worst_cell": f"r{Worst // GRID + 1}c{Worst % GRID + 1}",
                    "hash_distance": int((kRenderDiff.DHash(ImgA) != kRenderDiff.DHash(ImgB)).sum()),
                    "size_changed": Resized}, Diff
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kRenderDiff.Compare({PathA}, {PathB})")
            return None, None

    @staticmethod
    def Heatmap(Diff, Path):
        """Save the difference, amplified 4x, as a grey image."""
        if kS.ErrorMode:
            return None
        try:
            Img = Image.fromarray(np.uint8(np.clip(Diff * 4, 0, 1) * 255))
            Img.convert("RGB").save(Path)
            return Path
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRenderDiff.Heatmap")
            return None

    @staticmethod
    def Index(Folder):
        """Image files by lower-case name, so s002.png pairs with S002.PNG."""
        if kS.ErrorMode:
            return {}
        try:
            if not os.path.isdir(Folder):
                raise ToolReportableException(f"not a folder: {Folder}")
            return {Name.lower(): os.path.join(Folder, Name) for Name in os.listdir(Folder)
                    if Name.lower().endswith((".png", ".jpg"))}
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kRenderDiff.Index({Folder})")
            return {}


class kDiffRendersApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Parser = argparse.ArgumentParser()
            Parser.add_argument("before")
            Parser.add_argument("after")
            Parser.add_argument("--threshold", type=float, default=0.0, help="mean %% difference that counts as changed")
            Parser.add_argument("--heatmaps", help="folder for per-slide difference images of changed slides")
            Args = Parser.parse_args()
            Before, After = kRenderDiff.Index(Args.before), kRenderDiff.Index(Args.after)
            Names = sorted(set(Before) | set(After))
            Changed = 0
            for Name in Names:
                PathA, PathB = Before.get(Name), After.get(Name)
                if not PathA or not PathB:
                    print(f"{Name:<10} {'ADDED' if PathB else 'REMOVED'}")
                    Changed += 1
                    continue
                Report, Diff = kRenderDiff.Compare(PathA, PathB)
                if Report is None:
                    return 1
                Moved = Report["size_changed"] or Report["mean_pct"] > Args.threshold or \
                    (Args.threshold == 0 and Report["worst_cell_pct"] > 0)
                Changed += Moved
                Size = "  (image size changed)" if Report["size_changed"] else ""
                print(f"{Name:<10} {'CHANGED' if Moved else 'same':<8}{Size} mean {Report['mean_pct']:.3f}%  worst cell "
                      f"{Report['worst_cell']} {Report['worst_cell_pct']:.2f}%  hash dist {Report['hash_distance']}")
                if Moved and Args.heatmaps:
                    os.makedirs(Args.heatmaps, exist_ok=True)
                    kRenderDiff.Heatmap(Diff, os.path.join(Args.heatmaps, Name.rsplit(".", 1)[0] + ".diff.png"))
            print(f"\n{Changed} of {len(Names)} slide(s) changed")
            return 1 if Changed else 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDiffRendersApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kDiffRendersApp)
