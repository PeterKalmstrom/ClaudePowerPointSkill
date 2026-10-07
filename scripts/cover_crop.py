"""Crop an image to a target box ratio (cover, never stretch) before inserting it.

AddPicture / python-pptx add_picture with BOTH width and height STRETCHES the image. A 3:2
photo in a 1440 x 610 pt box comes out ~57 % too wide - and reads as "a wide crop", not a
defect. Check the distortion, then crop. A centred crop is a default, not a guarantee:
look at the render (faces and subjects get cut).

    uvx --with pillow python cover_crop.py photo.jpg --box 1440x610 [--focus-y 0.4] --out photo.crop.jpg
"""
import argparse
import os
import re

from PIL import Image

from kShared import ToolReportableException, kRun, kS, kToolException

BOX_PATTERN = re.compile(r"^\s*([0-9]*\.?[0-9]+)\s*x\s*([0-9]*\.?[0-9]+)\s*$")


class kCoverCrop:
    """Cover-crop geometry."""

    @staticmethod
    def Distortion(ImageWidth, ImageHeight, BoxWidth, BoxHeight):
        """Horizontal stretch in % if the image were forced into the box."""
        if kS.ErrorMode:
            return 0.0
        try:
            return ((BoxWidth / BoxHeight) / (ImageWidth / ImageHeight) - 1) * 100
        except Exception as e:
            kS.GlobalErrorHandler(e, "kCoverCrop.Distortion")
            return 0.0

    @staticmethod
    def Crop(Img, BoxWidth, BoxHeight, FocusX=0.5, FocusY=0.5):
        """The largest crop of Img with the box's ratio, positioned by the focus point."""
        if kS.ErrorMode:
            return None
        try:
            Target = BoxWidth / BoxHeight
            Width, Height = Img.size
            if Width / Height > Target:  # too wide -> crop sides
                NewWidth = round(Height * Target)
                X = round((Width - NewWidth) * FocusX)
                return Img.crop((X, 0, X + NewWidth, Height))
            NewHeight = round(Width / Target)  # too tall -> crop top/bottom
            Y = round((Height - NewHeight) * FocusY)
            return Img.crop((0, Y, Width, Y + NewHeight))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kCoverCrop.Crop")
            return None


class kCoverCropApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Parser = argparse.ArgumentParser()
            Parser.add_argument("image")
            Parser.add_argument("--box", required=True, help="width x height (any unit, only the ratio counts)")
            Parser.add_argument("--focus-x", type=float, default=0.5)
            Parser.add_argument("--focus-y", type=float, default=0.5)
            Parser.add_argument("--out", required=True)
            Args = Parser.parse_args()
            Box = BOX_PATTERN.match(Args.box.lower())
            if not Box or float(Box.group(1)) <= 0 or float(Box.group(2)) <= 0:
                raise ToolReportableException(f"--box must be WIDTHxHEIGHT with positive numbers, got {Args.box!r}")
            if not os.path.isfile(Args.image):
                raise ToolReportableException(f"not found: {Args.image}")
            BoxWidth, BoxHeight = float(Box.group(1)), float(Box.group(2))
            Img = Image.open(Args.image)
            print(f"stretch if inserted as-is: {kCoverCrop.Distortion(*Img.size, BoxWidth, BoxHeight):+.0f}% horizontal")
            Cropped = kCoverCrop.Crop(Img, BoxWidth, BoxHeight, Args.focus_x, Args.focus_y)
            if Cropped is None:
                return 1
            Cropped.save(Args.out)
            print(Args.out)
            return 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kCoverCropApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kCoverCropApp)
