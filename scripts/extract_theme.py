"""Read a deck's or template's theme: colours, fonts, layouts and their placeholders. Any OS.

    uvx --with python-pptx python scripts/extract_theme.py template.potx            # JSON
    uvx --with python-pptx python scripts/extract_theme.py template.pptx --markdown > brand-spec.md

Use it before building from a company template: take colours from the theme slots (accent1-6,
dk1/lt1...) and fonts from the major (headings) / minor (body) pair, and pick slide layouts by
their placeholders instead of drawing boxes on blank slides. A .potx is read like a .pptx.
"""
import argparse
import json
import shutil
import sys
import tempfile

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu
from lxml import etree

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
         "hlink", "folHlink"]


def open_any(path):
    """python-pptx refuses the .potx content type; copy to a .pptx name and patch the type."""
    if not path.lower().endswith((".potx", ".potm")):
        return Presentation(path)
    import zipfile
    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False).name
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(b"presentationml.template.main+xml", b"presentationml.presentation.main+xml")
            dst.writestr(item, data)
    return Presentation(tmp)


def colour(node):
    if node is None:
        return None
    srgb = node.find("a:srgbClr", NS)
    if srgb is not None:
        return srgb.get("val")
    sys_ = node.find("a:sysClr", NS)
    return sys_.get("lastClr") if sys_ is not None else None


def theme_of(master):
    part = master.part.part_related_by(RT.THEME)
    return etree.fromstring(part.blob)


def extract(path):
    prs = open_any(path)
    master = prs.slide_master
    theme = theme_of(master)
    scheme = theme.find(".//a:clrScheme", NS)
    fonts = theme.find(".//a:fontScheme", NS)
    out = {
        "file": path,
        "slide_size_pt": [round(Emu(prs.slide_width).pt, 1), round(Emu(prs.slide_height).pt, 1)],
        "theme_name": theme.get("name"),
        "colour_scheme": scheme.get("name") if scheme is not None else None,
        "colours": {s: colour(scheme.find(f"a:{s}", NS)) for s in SLOTS} if scheme is not None else {},
        "fonts": {
            "major_latin": fonts.find("a:majorFont/a:latin", NS).get("typeface") if fonts is not None else None,
            "minor_latin": fonts.find("a:minorFont/a:latin", NS).get("typeface") if fonts is not None else None,
        },
        "layouts": [],
    }
    for i, layout in enumerate(master.slide_layouts):
        phs = []
        for ph in layout.placeholders:
            pf = ph.placeholder_format
            phs.append({"idx": pf.idx, "type": str(pf.type).split(".")[-1].split(" ")[0], "name": ph.name,
                        "box_pt": [round(Emu(v or 0).pt) for v in (ph.left, ph.top, ph.width, ph.height)]})
        out["layouts"].append({"index": i, "name": layout.name, "placeholders": phs})
    return out


def markdown(t):
    c = t["colours"]
    lines = [f"# Brand spec — {t['theme_name'] or 'theme'}", "",
             f"Extracted from `{t['file']}` by scripts/extract_theme.py. Slide size: "
             f"{t['slide_size_pt'][0]:g} × {t['slide_size_pt'][1]:g} pt.", "",
             "## Fonts", "", f"- Headings (major): **{t['fonts']['major_latin']}**",
             f"- Body (minor): **{t['fonts']['minor_latin']}**", "",
             "## Colours", "", "| Slot | Hex | Use |", "|---|---|---|"]
    use = {"dk1": "main text", "lt1": "main background", "dk2": "secondary dark", "lt2": "secondary light",
           "accent1": "primary accent — the one highlight per slide", "hlink": "links", "folHlink": "visited links"}
    for s in SLOTS:
        lines.append(f"| {s} | `#{c.get(s)}` | {use.get(s, 'chart series / secondary accent')} |")
    lines += ["", "## Layouts", "", "| # | Layout | Placeholders |", "|---|---|---|"]
    for l in t["layouts"]:
        lines.append(f"| {l['index']} | {l['name']} | {', '.join(p['type'] for p in l['placeholders']) or '—'} |")
    lines += ["", "## Fill in by hand", "", "- Logo file and where it goes:", "- Imagery style:",
              "- Voice and tone:", "- Words to avoid:", ""]
    return "\n".join(lines)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--markdown", action="store_true", help="write a brand-spec.md skeleton instead of JSON")
    a = ap.parse_args()
    t = extract(a.file)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print(markdown(t) if a.markdown else json.dumps(t, indent=2, ensure_ascii=False))
