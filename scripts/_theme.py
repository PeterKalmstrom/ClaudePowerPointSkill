"""Resolve DrawingML colours against a deck's theme (python-pptx + lxml; any OS).

Theme colours (schemeClr), their lumMod/lumOff/tint/shade adjustments and the slide background
chain (slide -> layout -> master) are resolved to RRGGBB, so checks can see the colours people
actually see - not only colours set directly on the text. Picture and gradient backgrounds return
None: what is behind the text there is not knowable from the file.
"""
import colorsys

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from kShared import kS

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}
SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
         "hlink", "folHlink"]
DEFAULT_MAP = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
COLOUR_TAGS = ("srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr")
PRESET = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000", "blue": "0000FF",
          "yellow": "FFFF00", "gray": "808080", "grey": "808080"}


class kTheme:
    """One slide master's theme: colour scheme, colour map and fonts."""

    def __init__(self, Master):
        try:
            ThemeEl = etree.fromstring(Master.part.part_related_by(RT.THEME).blob)
            Scheme = ThemeEl.find(".//a:clrScheme", NS)
            self._colours = {}
            for Slot in SLOTS:
                Node = Scheme.find(f"a:{Slot}", NS) if Scheme is not None else None
                if Node is not None and len(Node):
                    C = Node[0]
                    self._colours[Slot] = (C.get("val") if C.tag == kTheme.Q("srgbClr") else C.get("lastClr")
                                           or ("FFFFFF" if C.get("val") == "window" else "000000")).upper()
            Fonts = ThemeEl.find(".//a:fontScheme", NS)
            self._major = Fonts.find("a:majorFont/a:latin", NS).get("typeface") if Fonts is not None else None
            self._minor = Fonts.find("a:minorFont/a:latin", NS).get("typeface") if Fonts is not None else None
            CMap = Master._element.find("p:clrMap", NS)
            self._map = dict(DEFAULT_MAP)
            if CMap is not None:
                self._map.update({K: V for K, V in CMap.attrib.items() if K in DEFAULT_MAP})
            self._master = Master
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.__init__")

    @staticmethod
    def Q(Tag):
        """Clark-notation name of a DrawingML tag."""
        if kS.ErrorMode:
            return None
        try:
            return f"{{{A}}}{Tag}"
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.Q")
            return None

    @property
    def Colours(self):
        """Theme slot (dk1, accent1, ...) -> RRGGBB."""
        if kS.ErrorMode:
            return {}
        try:
            return self._colours
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.Colours")
            return {}

    @property
    def Major(self):
        """Heading typeface."""
        if kS.ErrorMode:
            return None
        try:
            return self._major
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.Major")
            return None

    @property
    def Minor(self):
        """Body typeface."""
        if kS.ErrorMode:
            return None
        try:
            return self._minor
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.Minor")
            return None

    @property
    def Map(self):
        """Colour map (bg1/tx1/bg2/tx2 -> slot)."""
        if kS.ErrorMode:
            return {}
        try:
            return self._map
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.Map")
            return {}

    def Font(self, Typeface):
        """Map +mj-lt / +mn-lt (theme font references) to the real typeface."""
        if kS.ErrorMode:
            return None
        try:
            if Typeface in ("+mj-lt", "+mj-ea", "+mj-cs"):
                return self._major
            if Typeface in ("+mn-lt", "+mn-ea", "+mn-cs"):
                return self._minor
            return Typeface
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.Font")
            return None

    def IsThemeHex(self, HexVal):
        """True when HexVal is one of the theme's colours."""
        if kS.ErrorMode:
            return False
        try:
            return HexVal.upper() in self._colours.values()
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.IsThemeHex")
            return False

    def Resolve(self, El):
        """RRGGBB for a colour element (srgbClr, schemeClr, ...) with its child adjustments applied."""
        if kS.ErrorMode:
            return None
        try:
            if El is None:
                return None
            Tag = etree.QName(El).localname
            if Tag == "srgbClr":
                Base = El.get("val")
            elif Tag == "schemeClr":
                Name = El.get("val")
                Name = self._map.get(Name, Name)
                Base = self._colours.get(Name)
            elif Tag == "sysClr":
                Base = El.get("lastClr") or ("FFFFFF" if El.get("val") == "window" else "000000")
            elif Tag == "prstClr":
                Base = PRESET.get(El.get("val"))
            else:
                return None
            if not Base:
                return None
            return kTheme.Adjust(Base.upper(), El)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.Resolve")
            return None

    def ColourIn(self, Parent):
        """Resolve the first colour element directly under `Parent` (e.g. an <a:solidFill>)."""
        if kS.ErrorMode:
            return None
        try:
            if Parent is None:
                return None
            for Child in Parent:
                if etree.QName(Child).localname in COLOUR_TAGS:
                    return self.Resolve(Child)
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.ColourIn")
            return None

    def Background(self, Slide):
        """Solid background colour behind a slide (slide -> layout -> master), or None if it is a
        picture/gradient/unknown. White when nothing sets one."""
        if kS.ErrorMode:
            return None
        try:
            for Owner in (Slide, Slide.slide_layout, Slide.slide_layout.slide_master):
                Bg = Owner._element.find("p:cSld/p:bg", NS)
                if Bg is None:
                    continue
                Pr = Bg.find("p:bgPr", NS)
                if Pr is not None:
                    Solid = Pr.find("a:solidFill", NS)
                    return self.ColourIn(Solid) if Solid is not None else None
                Ref = Bg.find("p:bgRef", NS)
                if Ref is not None:
                    Idx = int(Ref.get("idx", "0"))
                    # 1001-1003 are theme background fills; a solid one takes the bgRef's colour
                    return self.ColourIn(Ref) if Idx in (0, 1001) or Idx < 1000 else None
            return self._colours.get(self._map.get("bg1", "lt1"), "FFFFFF")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.Background")
            return None

    @staticmethod
    def HexToRgb(Hex):
        """RRGGBB -> (r, g, b) in 0..1."""
        if kS.ErrorMode:
            return None
        try:
            return tuple(int(Hex[I:I + 2], 16) / 255 for I in (0, 2, 4))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.HexToRgb")
            return None

    @staticmethod
    def RgbToHex(R, G, B):
        """(r, g, b) in 0..1 -> RRGGBB."""
        if kS.ErrorMode:
            return None
        try:
            return "".join(f"{round(max(0, min(1, V)) * 255):02X}" for V in (R, G, B))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.RgbToHex")
            return None

    @staticmethod
    def Adjust(HexVal, El):
        """Apply lumMod/lumOff/tint/shade children (values in 1/1000 of a percent)."""
        if kS.ErrorMode:
            return None
        try:
            R, G, B = kTheme.HexToRgb(HexVal)
            for Mod in El:
                Name = etree.QName(Mod).localname
                Val = int(Mod.get("val", "100000")) / 100000
                if Name in ("lumMod", "lumOff"):
                    H, L, S = colorsys.rgb_to_hls(R, G, B)
                    L = L * Val if Name == "lumMod" else L + Val
                    R, G, B = colorsys.hls_to_rgb(H, max(0, min(1, L)), S)
                elif Name == "tint":  # towards white
                    R, G, B = (C + (1 - C) * (1 - Val) for C in (R, G, B))
                elif Name == "shade":  # towards black
                    R, G, B = (C * Val for C in (R, G, B))
            return kTheme.RgbToHex(R, G, B)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.Adjust")
            return None

    @staticmethod
    def Saturation(HexVal):
        """HSV saturation 0..1."""
        if kS.ErrorMode:
            return 0.0
        try:
            R, G, B = kTheme.HexToRgb(HexVal)
            return colorsys.rgb_to_hsv(R, G, B)[1]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.Saturation")
            return 0.0

    @staticmethod
    def Hue(HexVal):
        """HSV hue 0..1."""
        if kS.ErrorMode:
            return 0.0
        try:
            R, G, B = kTheme.HexToRgb(HexVal)
            return colorsys.rgb_to_hsv(R, G, B)[0]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTheme.Hue")
            return 0.0
