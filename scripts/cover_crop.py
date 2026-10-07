"""Crop an image to a target box ratio (cover, never stretch) before inserting it.

AddPicture / python-pptx add_picture with BOTH width and height STRETCHES the image. A 3:2
photo in a 1440 x 610 pt box comes out ~57 % too wide - and reads as "a wide crop", not a
defect. Check the distortion, then crop. A centred crop is a default, not a guarantee:
look at the render (faces and subjects get cut).

    uvx --with pillow python cover_crop.py photo.jpg --box 1440x610 [--focus-y 0.4] --out photo.crop.jpg
"""
import argparse

from PIL import Image


def distortion(img_w, img_h, box_w, box_h):
    """Horizontal stretch in % if the image were forced into the box."""
    return ((box_w / box_h) / (img_w / img_h) - 1) * 100


def cover_crop(img, box_w, box_h, focus_x=0.5, focus_y=0.5):
    target = box_w / box_h
    w, h = img.size
    if w / h > target:  # too wide -> crop sides
        nw = round(h * target)
        x = round((w - nw) * focus_x)
        return img.crop((x, 0, x + nw, h))
    nh = round(w / target)  # too tall -> crop top/bottom
    y = round((h - nh) * focus_y)
    return img.crop((0, y, w, y + nh))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("image")
    ap.add_argument("--box", required=True, help="width x height (any unit, only the ratio counts)")
    ap.add_argument("--focus-x", type=float, default=0.5)
    ap.add_argument("--focus-y", type=float, default=0.5)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    bw, bh = (float(x) for x in a.box.lower().split("x"))
    img = Image.open(a.image)
    print(f"stretch if inserted as-is: {distortion(*img.size, bw, bh):+.0f}% horizontal")
    cover_crop(img, bw, bh, a.focus_x, a.focus_y).save(a.out)
    print(a.out)
