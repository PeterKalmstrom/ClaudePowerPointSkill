"""Resolve DrawingML colours against a deck's theme (python-pptx + lxml; any OS).

Theme colours (schemeClr), their lumMod/lumOff/tint/shade adjustments and the slide background
chain (slide -> layout -> master) are resolved to RRGGBB, so checks can see the colours people
actually see - not only colours set directly on the text. Picture and gradient backgrounds return
None: what is behind the text there is not knowable from the file.
"""
import colorsys

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}
SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
         "hlink", "folHlink"]
DEFAULT_MAP = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
COLOUR_TAGS = ("srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr")
PRESET = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000", "blue": "0000FF",
          "yellow": "FFFF00", "gray": "808080", "grey": "808080"}


def _q(tag):
    return f"{{{A}}}{tag}"


class Theme:
    def __init__(self, master):
        theme_el = etree.fromstring(master.part.part_related_by(RT.THEME).blob)
        scheme = theme_el.find(".//a:clrScheme", NS)
        self.colours = {}
        for slot in SLOTS:
            node = scheme.find(f"a:{slot}", NS) if scheme is not None else None
            if node is not None and len(node):
                c = node[0]
                self.colours[slot] = (c.get("val") if c.tag == _q("srgbClr") else c.get("lastClr") or "000000").upper()
        fonts = theme_el.find(".//a:fontScheme", NS)
        self.major = fonts.find("a:majorFont/a:latin", NS).get("typeface") if fonts is not None else None
        self.minor = fonts.find("a:minorFont/a:latin", NS).get("typeface") if fonts is not None else None
        cmap = master._element.find("p:clrMap", NS)
        self.map = dict(DEFAULT_MAP)
        if cmap is not None:
            self.map.update({k: v for k, v in cmap.attrib.items() if k in DEFAULT_MAP})
        self.master = master

    def font(self, typeface):
        """Map +mj-lt / +mn-lt (theme font references) to the real typeface."""
        if typeface in ("+mj-lt", "+mj-ea", "+mj-cs"):
            return self.major
        if typeface in ("+mn-lt", "+mn-ea", "+mn-cs"):
            return self.minor
        return typeface

    def is_theme_hex(self, hexval):
        return hexval.upper() in self.colours.values()

    def resolve(self, el):
        """RRGGBB for a colour element (srgbClr, schemeClr, ...) with its child adjustments applied."""
        if el is None:
            return None
        tag = etree.QName(el).localname
        if tag == "srgbClr":
            base = el.get("val")
        elif tag == "schemeClr":
            name = el.get("val")
            name = self.map.get(name, name)
            base = self.colours.get(name)
        elif tag == "sysClr":
            base = el.get("lastClr") or ("FFFFFF" if el.get("val") == "window" else "000000")
        elif tag == "prstClr":
            base = PRESET.get(el.get("val"))
        else:
            return None
        if not base:
            return None
        return adjust(base.upper(), el)

    def colour_in(self, parent):
        """Resolve the first colour element directly under `parent` (e.g. an <a:solidFill>)."""
        if parent is None:
            return None
        for child in parent:
            if etree.QName(child).localname in COLOUR_TAGS:
                return self.resolve(child)
        return None

    def background(self, slide):
        """Solid background colour behind a slide (slide -> layout -> master), or None if it is a
        picture/gradient/unknown. White when nothing sets one."""
        for owner in (slide, slide.slide_layout, slide.slide_layout.slide_master):
            bg = owner._element.find("p:cSld/p:bg", NS)
            if bg is None:
                continue
            pr = bg.find("p:bgPr", NS)
            if pr is not None:
                solid = pr.find("a:solidFill", NS)
                return self.colour_in(solid) if solid is not None else None
            ref = bg.find("p:bgRef", NS)
            if ref is not None:
                idx = int(ref.get("idx", "0"))
                # 1001-1003 are theme background fills; a solid one takes the bgRef's colour
                return self.colour_in(ref) if idx in (0, 1001) or idx < 1000 else None
        return self.colours.get(self.map.get("bg1", "lt1"), "FFFFFF")


def _hex_to_rgb(h):
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _rgb_to_hex(r, g, b):
    return "".join(f"{round(max(0, min(1, v)) * 255):02X}" for v in (r, g, b))


def adjust(hexval, el):
    """Apply lumMod/lumOff/tint/shade children (values in 1/1000 of a percent)."""
    r, g, b = _hex_to_rgb(hexval)
    for mod in el:
        name = etree.QName(mod).localname
        val = int(mod.get("val", "100000")) / 100000
        if name in ("lumMod", "lumOff"):
            h, l, s = colorsys.rgb_to_hls(r, g, b)
            l = l * val if name == "lumMod" else l + val
            r, g, b = colorsys.hls_to_rgb(h, max(0, min(1, l)), s)
        elif name == "tint":  # towards white
            r, g, b = (c + (1 - c) * (1 - val) for c in (r, g, b))
        elif name == "shade":  # towards black
            r, g, b = (c * val for c in (r, g, b))
    return _rgb_to_hex(r, g, b)


def saturation(hexval):
    r, g, b = _hex_to_rgb(hexval)
    return colorsys.rgb_to_hsv(r, g, b)[1]


def hue(hexval):
    r, g, b = _hex_to_rgb(hexval)
    return colorsys.rgb_to_hsv(r, g, b)[0]
