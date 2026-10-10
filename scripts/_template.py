"""Building into a company template (.potx/.pptx): which of its layouts each pattern uses, whether its colours,
fonts and slide size suit the builder's rules, and a readable report of both (build_deck.py --inspect).

Nothing here knows a project: a second consumer (another builder, an MCP tool) passes a Presentation or a path.
What it could change: ROLE_WORDS (layout-name words per role, add a language) and PATTERN_ROLES (which patterns
open on a cover or divider layout) - both plain module tables.
"""
import os

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Emu

from _measure import kMeasure
from _rules import kRules
from kShared import ToolInputException, kS, kToolException

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS_T = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "p": "http://schemas.openxmlformats.org/presentationml/2006/main", "p14": P14}
TEMPLATE_TYPES = (".potx", ".potm", ".pptx", ".pptm")
TITLE_PH = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
FOOTER_PH = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}
BODY_PH = {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.SUBTITLE}
ROLE_WORDS = [  # checked in this order: "title only" before "title"
    ("title_only", ("title only", "headline only", "heading only", "titel only", "endast rubrik", "rubrik endast")),
    ("section", ("section", "divider", "chapter", "break", "avsnitt", "kapitel", "segment")),
    ("title", ("title slide", "cover", "title page", "opening", "front", "rubrikbild", "titelbild")),
    ("blank", ("blank", "empty", "tom")),
    ("title_content", ("title and content", "content", "text", "body", "innehall"))]
PATTERN_ROLES = {"title": "title", "section": "section"}  # every other pattern is a content slide
FALLBACK = {"title": ("title", "section", "title_only", "title_content", "other"),
            "section": ("section", "title", "title_only", "title_content", "other"),
            "content": ("title_only", "title_content", "other", "section", "title")}
WIDE = 16 / 9


class kTemplateMap:
    """Reads a template's layouts, colours, fonts and size and maps the builder's patterns onto its layouts."""

    @staticmethod
    def CheckPath(Path):
        """Stop with exit 2 when the template path is missing or not a PowerPoint file (the caller's input)."""
        if kS.ErrorMode:
            return
        try:
            if not Path or not os.path.isfile(Path):
                raise ToolInputException(f"template not found: {os.path.abspath(Path or '')}")
            if not Path.lower().endswith(TEMPLATE_TYPES):
                raise ToolInputException(f"template must be .potx, .potm, .pptx or .pptm: {Path}")
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kTemplateMap.CheckPath(path={Path})")
            return

    @staticmethod
    def Open(Path):
        """The template as a Presentation (a .potx/.potm with its content type patched in memory)."""
        if kS.ErrorMode:
            return None
        try:
            kTemplateMap.CheckPath(Path)
            from extract_theme import kThemeReader
            return kThemeReader.OpenAny(Path)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kTemplateMap.Open(path={Path})")
            return None

    @staticmethod
    def DropSlides(Prs):
        """Remove the template's own example slides, and the sections and custom shows that list them (left
        behind, they make PowerPoint offer to repair the file). The masters, layouts, logos and footers stay.
        Returns how many slides were dropped."""
        if kS.ErrorMode:
            return 0
        try:
            Ids = Prs.slides._sldIdLst
            Count = len(Ids)
            for Sld in list(Ids):
                Prs.part.drop_rel(Sld.rId)
                Ids.remove(Sld)
            Root = Prs.part._element
            for Section in Root.iter(f"{{{P14}}}sldIdLst"):
                for Child in list(Section):
                    Section.remove(Child)
            for Shows in Root.findall("p:custShowLst", NS_T):
                Root.remove(Shows)
            return Count
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.DropSlides")
            return 0

    @staticmethod
    def Types(Layout):
        """The placeholder types on a layout."""
        if kS.ErrorMode:
            return []
        try:
            return [Ph.placeholder_format.type for Ph in Layout.placeholders]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.Types")
            return []

    @staticmethod
    def HasTitle(Layout):
        """True when the layout has a title placeholder (every pattern writes its claim into one)."""
        if kS.ErrorMode:
            return False
        try:
            return any(T in TITLE_PH for T in kTemplateMap.Types(Layout))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.HasTitle")
            return False

    @staticmethod
    def RoleByName(Layout):
        """The role the layout's name says (title, section, title_only, title_content, blank), or ""."""
        if kS.ErrorMode:
            return ""
        try:
            Name = " " + Layout.name.lower().replace("-", " ").replace("_", " ").replace("+", " and ") + " "
            for Role, Words in ROLE_WORDS:
                for Word in Words:
                    if f" {Word} " in Name or (len(Word) > 4 and Word in Name):
                        return Role
            return ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.RoleByName")
            return ""

    @staticmethod
    def RoleByShape(Layout, SlideHeight):
        """The role the placeholders suggest, for a layout whose name says nothing."""
        if kS.ErrorMode:
            return "other"
        try:
            Types = kTemplateMap.Types(Layout)
            Others = [T for T in Types if T not in TITLE_PH and T not in FOOTER_PH]
            if not kTemplateMap.HasTitle(Layout):
                return "blank" if not Others else "other"
            if PP_PLACEHOLDER.CENTER_TITLE in Types:
                return "title"
            if not Others:
                return "title_only"
            Bodies = [Ph for Ph in Layout.placeholders if Ph.placeholder_format.type in BODY_PH]
            if len(Others) == 1 and Bodies and (Bodies[0].height or 0) < SlideHeight * 0.25:
                return "section"  # a title and one short text line: a divider
            return "title_content"
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.RoleByShape")
            return "other"

    @staticmethod
    def Roles(Prs):
        """[(layout, role)] for the first master's layouts; a named role needs a title placeholder to count
        (except blank)."""
        if kS.ErrorMode:
            return []
        try:
            Result = []
            for Layout in Prs.slide_layouts:
                Role = kTemplateMap.RoleByName(Layout)
                if Role and Role != "blank" and not kTemplateMap.HasTitle(Layout):
                    Role = ""
                Result.append((Layout, Role or kTemplateMap.RoleByShape(Layout, Prs.slide_height)))
            return Result
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.Roles")
            return []

    @staticmethod
    def Map(Prs):
        """{"title"|"section"|"content": layout}: the first layout of the first role in FALLBACK that has a title
        placeholder. Blank layouts are never used - every pattern writes its claim into the title placeholder."""
        if kS.ErrorMode:
            return {}
        try:
            Roles = kTemplateMap.Roles(Prs)
            Result = {}
            for Group, Order in FALLBACK.items():
                for Wanted in Order:
                    Hit = next((L for L, R in Roles if R == Wanted and kTemplateMap.HasTitle(L)), None)
                    if Hit is not None:
                        Result[Group] = Hit
                        break
            return Result
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.Map")
            return {}

    @staticmethod
    def LayoutFor(Mapping, Pattern, Default):
        """The layout a slide of this pattern is built on: the mapped one for its group, else Default."""
        if kS.ErrorMode:
            return Default
        try:
            if not Mapping:
                return Default
            return Mapping.get(PATTERN_ROLES.get(Pattern, "content")) or Default
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.LayoutFor")
            return Default

    @staticmethod
    def Palette(Prs):
        """The colours the builder paints with, resolved through the master's colour map: text (tx1), background
        (bg1), muted (tx2), quiet/card (bg2), accent (accent1), plus the theme's major/minor fonts."""
        if kS.ErrorMode:
            return {}
        try:
            from extract_theme import kThemeReader
            Master = Prs.slide_master
            Theme = kThemeReader.ThemeOf(Master)
            Slots = kThemeReader.Colours(Theme.find(".//a:clrScheme", NS_T))
            Map = Master._element.find("p:clrMap", NS_T)
            Get = Map.get if Map is not None else {}.get
            Pick = {"text": Get("tx1", "dk1"), "background": Get("bg1", "lt1"), "muted": Get("tx2", "dk2"),
                    "quiet": Get("bg2", "lt2"), "accent": "accent1"}
            Result = {Key: (Slots.get(Slot) or "000000").upper() for Key, Slot in Pick.items()}
            Result.update(kThemeReader.Fonts(Theme.find(".//a:fontScheme", NS_T)))
            return Result
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.Palette")
            return {}

    @staticmethod
    def ContrastWarnings(Pal):
        """One line per colour pair the builder uses that falls below WCAG (4.5:1 text, 3:1 large figures)."""
        if kS.ErrorMode:
            return []
        try:
            Pairs = [("text", "background", 4.5, "body text"), ("muted", "background", 4.5, "captions and footers"),
                     ("text", "quiet", 4.5, "text on cards"), ("muted", "quiet", 4.5, "captions on cards"),
                     ("accent", "background", 3.0, "accent figures and kickers"),
                     ("background", "accent", 4.5, "text on a highlighted (accent) card")]
            Out = []
            for Fg, Bg, Need, Use in Pairs:
                Ratio = kRules.ContrastRatio(Pal[Fg], Pal[Bg]) or 0
                if Ratio < Need:
                    Out.append(f"contrast: {Fg} #{Pal[Fg]} on {Bg} #{Pal[Bg]} is {Ratio:.1f}:1, needs {Need:g}:1 "
                               f"({Use}) - lint will flag these; fix the template's theme colours or build with a "
                               "direction instead")
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.ContrastWarnings")
            return []

    @staticmethod
    def FontWarnings(Pal):
        """One line per theme font that is not installed here (PowerPoint substitutes it; fit is measured with a
        stand-in, so a title that fits here can wrap on a machine that has the real face, or the reverse)."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for Key, Use in (("major_latin", "headings"), ("minor_latin", "body")):
                Face = Pal.get(Key) or ""
                if Face and not Face.startswith("+") and not kMeasure.Installed(Face):
                    Out.append(f"font: '{Face}' ({Use}) is not installed on this machine - PowerPoint will substitute "
                               "it and text fit is measured with a stand-in; install the font or check the render "
                               "on a machine that has it")
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.FontWarnings")
            return []

    @staticmethod
    def SizeWarnings(Prs):
        """A line when the slide is not 16:9: the 1440 x 810 grid is stretched onto it, so shapes keep their
        places but not their proportions (circles become ovals, wide patterns get cramped)."""
        if kS.ErrorMode:
            return []
        try:
            Wd, Ht = Emu(Prs.slide_width).pt, Emu(Prs.slide_height).pt
            if abs(Wd / Ht - WIDE) <= 0.02:
                return []
            return [f"size: the template is {Wd:g} x {Ht:g} pt ({Wd / Ht:.2f}:1, not 16:9) - the 16:9 grid is "
                    "stretched onto it: compare, timeline and kpi rows get cramped and round shapes turn oval; "
                    "use a 16:9 version of the template when there is one"]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.SizeWarnings")
            return []

    @staticmethod
    def MapWarnings(Mapping):
        """A line for each group that fell back to another role's layout, or has none at all."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            if "content" not in Mapping:
                Out.append("layouts: the template has no layout with a title placeholder - it cannot be built on")
            for Group in ("title", "section"):
                if Group in Mapping and Mapping[Group] is Mapping.get("content"):
                    Out.append(f"layouts: no {Group} layout found - {Group} slides use '{Mapping[Group].name}'")
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.MapWarnings")
            return []

    @staticmethod
    def Warnings(Prs):
        """Every template warning: layouts, slide size, contrast, fonts."""
        if kS.ErrorMode:
            return []
        try:
            Pal = kTemplateMap.Palette(Prs)
            return (kTemplateMap.MapWarnings(kTemplateMap.Map(Prs)) + kTemplateMap.SizeWarnings(Prs)
                    + kTemplateMap.ContrastWarnings(Pal) + kTemplateMap.FontWarnings(Pal))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.Warnings")
            return []

    @staticmethod
    def FooterText(Prs):
        """The footer text the template's master carries ("" when none), the deck footer when the spec has none."""
        if kS.ErrorMode:
            return ""
        try:
            for Ph in Prs.slide_master.placeholders:
                if Ph.placeholder_format.type == PP_PLACEHOLDER.FOOTER and Ph.has_text_frame:
                    return Ph.text_frame.text.strip()
            return ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.FooterText")
            return ""

    @staticmethod
    def LayoutLines(Prs, Mapping):
        """The inspect report's layout table: index, name, role, placeholders, what uses it."""
        if kS.ErrorMode:
            return []
        try:
            Uses = {}
            for Group, Layout in Mapping.items():
                Uses.setdefault(id(Layout), []).append(Group)
            Lines = ["Layouts (first master):"]
            for No, (Layout, Role) in enumerate(kTemplateMap.Roles(Prs)):
                Types = ", ".join(str(T).split(".")[-1].split(" ")[0].lower() for T in kTemplateMap.Types(Layout))
                Used = Uses.get(id(Layout))
                Tail = f"  <- {' + '.join(Used)} slides" if Used else ""
                Bg = " [own background]" if Layout._element.find("p:cSld/p:bg", NS_T) is not None else ""
                Lines.append(f"  {No:>2}. {Layout.name:<28} role {Role:<13} [{Types or 'none'}]{Bg}{Tail}")
            return Lines
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTemplateMap.LayoutLines")
            return []

    @staticmethod
    def Inspect(Path):
        """The --inspect report: size, colours, fonts, layouts with their roles, the pattern mapping, the slides
        that will be dropped, and every warning."""
        if kS.ErrorMode:
            return ""
        try:
            Prs = kTemplateMap.Open(Path)
            Pal, Mapping = kTemplateMap.Palette(Prs), kTemplateMap.Map(Prs)
            Wd, Ht = Emu(Prs.slide_width).pt, Emu(Prs.slide_height).pt
            Lines = [f"Template: {os.path.basename(Path)}", f"Slide size: {Wd:g} x {Ht:g} pt ({Wd / Ht:.2f}:1)",
                     f"Masters: {len(Prs.slide_masters)} (the first is used); example slides to drop: "
                     f"{len(Prs.slides)}", "Colours the builder paints with:"]
            for Key in ("text", "background", "muted", "quiet", "accent"):
                Lines.append(f"  {Key:<11} #{Pal[Key]}")
            for Key, Use in (("major_latin", "headings"), ("minor_latin", "body")):
                Face = Pal.get(Key) or "-"
                Lines.append(f"Font, {Use}: {Face}" + ("" if kMeasure.Installed(Face) else " (not installed here)"))
            Foot = kTemplateMap.FooterText(Prs)
            Lines.append(f"Master footer text: {Foot!r}" + (" (used when the spec has no footer)" if Foot else ""))
            Lines += kTemplateMap.LayoutLines(Prs, Mapping)
            Lines.append("Mapping:")
            for Group, Who in (("title", "title"), ("section", "section"), ("content", "every other pattern")):
                Lines.append(f"  {Who:<20} -> " + (Mapping[Group].name if Group in Mapping else "(none)"))
            Found = kTemplateMap.Warnings(Prs)
            Lines += [f"template warning: {W}" for W in Found]
            Lines.append(f"inspect: {len(Found)} warning(s)")
            return "\n".join(Lines)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kTemplateMap.Inspect(path={Path})")
            return ""

    @staticmethod
    def PathWarnings(Path):
        """Warnings for a template file (opened once, read only)."""
        if kS.ErrorMode:
            return []
        try:
            return kTemplateMap.Warnings(kTemplateMap.Open(Path))
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kTemplateMap.PathWarnings(path={Path})")
            return []
