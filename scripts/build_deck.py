"""Build a deck from a JSON (or YAML) spec - any OS, python-pptx only.

    uvx --with python-pptx --with pillow python scripts/build_deck.py spec.json --out deck.pptx
    uvx --with python-pptx --with pillow python scripts/build_deck.py spec.json --out deck.pptx --lint
    python scripts/build_deck.py --list-directions

The spec names a design direction (scripts/directions.json) or a template (.pptx/.potx), then lists
slides by pattern. The builder applies the skill's rules by construction: 1440 x 810 pt, a real title
placeholder on every slide, theme colours and fonts (so the deck re-themes), named shapes, stable slide
ids, cover-cropped pictures, native charts with one highlighted finding, alt text and speaker notes.
Inputs are checked against each pattern's limits first; nothing is written if a limit is broken
(--force builds anyway). The spec format is in reference/BUILDER.md.

Patterns: title, section, statement, big_number, kpi, bullets, compare, process, timeline, quote,
chart, table, image, matrix.
"""
import argparse
import importlib.util
import io
import json
import os
import subprocess
import sys

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Pt

from kShared import ToolReportableException, kRun, kS, kToolException

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from _measure import kMeasure  # noqa: E402  (the skill's own helpers sit beside this script)
from _rules import kRules  # noqa: E402
from _theme import kTheme  # noqa: E402

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}

# Canvas and type scale for a 1440 x 810 pt slide (1920 x 1080 px). Body never below 18 pt.
W, H, M = 1440, 810, 80
TITLE_TOP, TITLE_H = 56, 100
BODY_TOP, BODY_BOTTOM = 196, 730
SIZE = {"title": 54, "section": 84, "statement": 72, "hero": 220, "kpi": 84, "h2": 46, "body": 34,
        "small": 30, "caption": 28, "label": 24}  # 1.5x a 960-pt slide: body 34 ~ 23 pt, floor 27 ~ 18 pt
GAP = 24

TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
FOOTER_TYPES = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}
TEXT, MUTED, ACCENT, BG, QUIET = (MSO_THEME_COLOR.TEXT_1, MSO_THEME_COLOR.TEXT_2, MSO_THEME_COLOR.ACCENT_1,
                                  MSO_THEME_COLOR.BACKGROUND_1, MSO_THEME_COLOR.BACKGROUND_2)

LIMITS = {  # pattern: {field: max characters} plus list-length ranges, checked before building
    "title": {"title": 70, "subtitle": 120},
    "section": {"title": 60, "eyebrow": 30},
    "statement": {"title": 90, "support": 140},
    "big_number": {"title": 70, "number": 12, "unit": 10, "caption": 90},
    "kpi": {"title": 70, "metrics": (3, 6), "metrics.value": 12, "metrics.label": 28},
    "bullets": {"title": 70, "items": (1, 7), "items.*": 100},
    "compare": {"title": 70, "columns": (2, 3), "columns.heading": 30, "columns.points": (1, 4),
                "columns.points.*": 70},
    "process": {"title": 70, "steps": (3, 8), "steps.label": 30, "steps.detail": 60},
    "timeline": {"title": 70, "events": (3, 7), "events.date": 16, "events.label": 40},
    "quote": {"quote": 240, "attribution": 40, "role": 60},
    "chart": {"title": 70, "categories": (2, 24), "series": (1, 6)},
    "table": {"title": 70, "header": (2, 6), "rows": (1, 8)},
    "image": {"title": 70, "caption": 120},
    "matrix": {"title": 70, "quadrants": (4, 4), "quadrants.heading": 30, "quadrants.text": 90,
               "x_axis": 30, "y_axis": 30},
}


# ---------------------------------------------------------------- JSON Schema (generated from LIMITS)

FIELDS = {  # pattern: {field: kind}; kinds: text, int, number, list[text], list[obj:{...}], obj
    "title": {"subtitle": "text"},
    "section": {"eyebrow": "text"},
    "statement": {"support": "text"},
    "big_number": {"number": "text", "unit": "text", "caption": "text"},
    "kpi": {"metrics": {"value": "text", "label": "text"}, "highlight": "int"},
    "bullets": {"items": "list[text]"},
    "compare": {"columns": {"heading": "text", "points": "list[text]"}, "highlight": "int"},
    "process": {"steps": {"label": "text", "detail": "text"}, "highlight": "int"},
    "timeline": {"events": {"date": "text", "label": "text"}, "highlight": "int"},
    "quote": {"quote": "text", "attribution": "text", "role": "text"},
    "chart": {"type": "enum:column,bar,line,pie", "categories": "list[text]",
              "series": {"name": "text", "values": "list[number]"}, "highlight": "int|text",
              "number_format": "text", "caption": "text", "alt": "text"},
    "table": {"header": "list[text]", "rows": "list[list[text]]", "highlight_row": "int", "alt": "text"},
    "image": {"image": "text", "caption": "text", "alt": "text", "focus_x": "number", "focus_y": "number"},
    "matrix": {"quadrants": {"heading": "text", "text": "text"}, "x_axis": "text", "y_axis": "text",
               "highlight": "int"},
}
REQUIRED = {"title": ["title"], "section": ["title"], "statement": ["title"], "quote": ["quote"],
            "big_number": ["title", "number"], "kpi": ["title", "metrics"], "bullets": ["title", "items"],
            "compare": ["title", "columns"], "process": ["title", "steps"], "timeline": ["title", "events"],
            "chart": ["title", "categories", "series"], "table": ["title", "header", "rows"],
            "image": ["title", "image"], "matrix": ["title", "quadrants"]}
_STRS = {"oneOf": [{"type": "string"}, {"type": "array", "items": {"type": "string"}}]}
NOTES_SCHEMA = {"description": "Speaker notes: a string, or key_fact / facts / qa / pitfalls / sources "
                               "(lists may be a single string; a Q&A item may be a string).", "oneOf": [
    {"type": "string"},
    {"type": "object", "additionalProperties": False, "properties": {
        "key_fact": {"type": "string"}, "facts": _STRS, "pitfalls": _STRS, "sources": _STRS,
        "qa": {"oneOf": [{"type": "object"}, {"type": "array", "items": {"oneOf": [
            {"type": "string"},
            {"type": "object", "additionalProperties": False, "required": ["q"],
             "properties": {"q": {"type": "string"}, "a": {"type": "string"}}}]}}]}}}]}

FLOOR = 27        # body text never shrinks below this (1440-pt grid; scaled by F())
LABEL_MIN = 24    # short labels (< 4 words) may go this small
RAMP = [0, 0.35, -0.3, 0.6, -0.5, 0.75]  # shades of the accent for series 1..6
CHART_TYPES = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
               "line": XL_CHART_TYPE.LINE_MARKERS, "pie": XL_CHART_TYPE.PIE}
PATTERNS = {"title": "Title", "section": "Section", "statement": "Statement", "big_number": "BigNumber",
            "kpi": "Kpi", "bullets": "Bullets", "compare": "Compare", "process": "Process",
            "timeline": "Timeline", "quote": "Quote", "chart": "Chart", "table": "Table", "image": "Image",
            "matrix": "Matrix"}  # pattern -> kSlidePatterns method
NUMERIC_KEYS = {"values", "highlight", "highlight_row", "focus_x", "focus_y"}
HIGHLIGHT_ITEMS = {"kpi": "metrics", "compare": "columns", "process": "steps", "timeline": "events",
                   "matrix": "quadrants", "chart": "categories"}
NEEDED_LIST = {"kpi": "metrics", "process": "steps", "timeline": "events", "compare": "columns",
               "matrix": "quadrants", "chart": "series", "table": "rows", "bullets": "items"}


class kSpecSchema:
    """The spec's JSON Schema, generated from FIELDS / LIMITS / REQUIRED."""

    @staticmethod
    def Directions():
        """Every design direction in directions.json."""
        if kS.ErrorMode:
            return []
        try:
            with open(os.path.join(HERE, "directions.json"), encoding="utf-8") as Fh:
                return json.load(Fh)["directions"]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecSchema.Directions")
            return []

    @staticmethod
    def KindSchema(Kind, MaxLen=None):
        """The schema of one field kind."""
        if kS.ErrorMode:
            return {}
        try:
            Text = {"type": ["string", "number"]} if MaxLen is None else {"type": ["string", "number"],
                                                                          "maxLength": MaxLen}
            if isinstance(Kind, dict):
                return {"type": "object", "additionalProperties": False, "properties": {}}
            if Kind == "text":
                return Text
            if Kind == "int":
                return {"type": "integer", "minimum": 0}
            if Kind == "number":
                return {"type": "number"}
            if Kind == "int|text":
                return {"type": ["integer", "string"]}
            if Kind.startswith("enum:"):
                return {"enum": Kind[5:].split(",")}
            if Kind == "list[text]":
                return {"type": "array", "items": Text}
            if Kind == "list[number]":
                return {"type": "array", "items": {"type": ["number", "null"]}}
            if Kind == "list[list[text]]":
                return {"type": "array", "items": {"type": "array", "items": {"type": ["string", "number"]}}}
            raise ValueError(Kind)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecSchema.KindSchema")
            return {}

    @staticmethod
    def SpecSchema():
        """JSON Schema for a build_deck spec, generated from FIELDS / LIMITS / REQUIRED."""
        if kS.ErrorMode:
            return {}
        try:
            Variants = []
            for Pat, Fields in FIELDS.items():
                Lim = LIMITS[Pat]
                Props = {"pattern": {"const": Pat}, "id": {"type": "string", "pattern": "^[A-Za-z0-9_.-]+$",
                                                           "description": "Stable slide id; keep it when the content changes."},
                         "title": {"type": ["string", "number"], "maxLength": Lim.get("title", 90),
                                   "description": "Write it as a claim, not a topic."},
                         "notes": NOTES_SCHEMA}
                for Name, Kind in Fields.items():
                    if isinstance(Kind, dict):
                        Item = {"type": "object", "additionalProperties": False, "properties": {}, "required": []}
                        for Sub, Sk in Kind.items():
                            SubLim = Lim.get(f"{Name}.{Sub}")
                            if isinstance(SubLim, tuple):
                                Item["properties"][Sub] = dict(kSpecSchema.KindSchema(Sk, Lim.get(f"{Name}.{Sub}.*")),
                                                               minItems=SubLim[0], maxItems=SubLim[1])
                            else:
                                Item["properties"][Sub] = kSpecSchema.KindSchema(Sk, SubLim)
                            if Sk in ("text", "list[text]", "list[number]") and Sub not in ("detail", "text"):
                                Item["required"].append(Sub)
                        Schema = {"type": "array", "items": Item}
                    else:
                        Star = Lim.get(f"{Name}.*")
                        Schema = kSpecSchema.KindSchema(Kind, Star if Kind.startswith("list") else Lim.get(Name))
                    Rng = Lim.get(Name)
                    if isinstance(Rng, tuple):
                        Schema["minItems"], Schema["maxItems"] = Rng
                    Props[Name] = Schema
                Variants.append({"type": "object", "additionalProperties": False, "properties": Props,
                                 "required": ["pattern"] + REQUIRED[Pat]})
            return {
                "$schema": "https://json-schema.org/draft/2020-12/schema",
                "$id": "https://github.com/PeterKalmstrom/claude-powerpoint-skill/scripts/spec.schema.json",
                "title": "build_deck.py spec",
                "description": "A deck as data: a design direction or template, then slides by pattern. "
                               "Generated by `build_deck.py --print-schema`; see reference/BUILDER.md.",
                "type": "object", "additionalProperties": False, "required": ["slides"],
                "properties": {
                    "$schema": {"type": "string"},
                    "direction": {"enum": [D["id"] for D in kSpecSchema.Directions()]},
                    "template": {"type": "string", "description": ".pptx or .potx, relative to the spec"},
                    "slides": {"type": "array", "minItems": 1, "items": {"oneOf": Variants}}}}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecSchema.SpecSchema")
            return {}


class kSpecCheck:
    """Spec checks: pattern limits, cross-field rules and (when jsonschema is installed) the schema."""

    @staticmethod
    def Values(Obj, Parts, Path=""):
        """Every (path, value) that a dotted LIMITS key ('columns.points.*') reaches in Obj."""
        if kS.ErrorMode:
            return []
        try:
            Head, Rest = Parts[0], Parts[1:]
            if Head == "*":
                Items = list(enumerate(Obj)) if isinstance(Obj, list) else []
            else:
                Items = [(Head, Obj.get(Head))] if isinstance(Obj, dict) and Head in Obj else []
            Out = []
            for K, V in Items:
                Pth = f"{Path}.{K}" if Path else str(K)
                if not Rest:
                    Out.append((Pth, V))
                elif isinstance(V, list) and Rest[0] != "*":
                    for J, El in enumerate(V):
                        Out += kSpecCheck.Values(El, Rest, f"{Pth}[{J}]")
                else:
                    Out += kSpecCheck.Values(V, Rest, Pth)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecCheck.Values")
            return []

    @staticmethod
    def CheckSpec(Spec):
        """Every pattern-limit and cross-field mistake in the spec, as messages."""
        if kS.ErrorMode:
            return []
        try:
            Errors = []
            for I, Sl in enumerate(Spec.get("slides", []), 1):
                Pat = Sl.get("pattern")
                if Pat not in LIMITS:
                    Errors.append(f"slide {I}: unknown pattern '{Pat}' (one of: {', '.join(LIMITS)})")
                    continue
                if Pat not in ("title", "quote", "section") and not Sl.get("title"):
                    Errors.append(f"slide {I} ({Pat}): needs a 'title' - write it as a claim")
                for Key, Lim in LIMITS[Pat].items():
                    for Pth, Val in kSpecCheck.Values(Sl, Key.split(".")):
                        if isinstance(Lim, tuple):
                            N = len(Val) if isinstance(Val, list) else 0
                            if not Lim[0] <= N <= Lim[1]:
                                Errors.append(f"slide {I} ({Pat}): '{Pth}' has {N} items, needs {Lim[0]}-{Lim[1]}")
                        elif isinstance(Val, str) and len(Val) > Lim:
                            Errors.append(f"slide {I} ({Pat}): '{Pth}' is {len(Val)} characters, max {Lim}: '{Val[:40]}…'")
                if Pat in NEEDED_LIST and NEEDED_LIST[Pat] not in Sl:
                    Errors.append(f"slide {I} ({Pat}): missing '{NEEDED_LIST[Pat]}'")
                if Pat == "image" and not os.path.exists(Sl.get("image", "")):
                    Errors.append(f"slide {I} (image): image file not found: {Sl.get('image')}")
                if Pat == "table":
                    Width = len(Sl.get("header", []))
                    for R, Row in enumerate(Sl.get("rows", [])):
                        if len(Row) != Width:
                            Errors.append(f"slide {I} (table): row {R + 1} has {len(Row)} cells, the header has {Width}")
                    Hr = Sl.get("highlight_row")
                    if Hr is not None and not (isinstance(Hr, int) and 0 <= Hr < len(Sl.get("rows", []))):
                        Errors.append(f"slide {I} (table): highlight_row {Hr!r} is not a row index (0-based)")
                if Pat == "chart":
                    Hi = Sl.get("highlight")
                    if isinstance(Hi, str) and Hi not in Sl.get("categories", []):
                        Errors.append(f"slide {I} (chart): highlight '{Hi}' is not one of the categories")
                    if Sl.get("type", "column") not in CHART_TYPES:
                        Errors.append(f"slide {I} (chart): type '{Sl.get('type')}' (one of: {', '.join(CHART_TYPES)})")
                ItemsKey = HIGHLIGHT_ITEMS.get(Pat)
                Hi = Sl.get("highlight")
                if ItemsKey and isinstance(Hi, int) and not 0 <= Hi < len(Sl.get(ItemsKey, [])):
                    Errors.append(f"slide {I} ({Pat}): highlight {Hi} is out of range (0-based, "
                                  f"{len(Sl.get(ItemsKey, []))} items)")
                if Pat == "chart":
                    Cats = len(Sl.get("categories", []))
                    for S in Sl.get("series", []):
                        if len(S.get("values", [])) != Cats:
                            Errors.append(f"slide {I} (chart): series '{S.get('name')}' has {len(S.get('values', []))} "
                                          f"values for {Cats} categories")
            return Errors
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecCheck.CheckSpec")
            return []

    @staticmethod
    def ErrorPath(Error):
        """'a/b/0' for a jsonschema error's absolute path."""
        if kS.ErrorMode:
            return ""
        try:
            return "/".join(str(Part) for Part in Error.absolute_path)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecCheck.ErrorPath")
            return ""

    @staticmethod
    def SchemaErrors(Spec):
        """Validate against the JSON Schema when the jsonschema package is installed (optional). Each slide is
        checked against its own pattern's schema, so every misspelt or misplaced field is named."""
        if kS.ErrorMode:
            return []
        try:
            if importlib.util.find_spec("jsonschema") is None:
                return []  # optional dependency: without it check_spec alone validates
            import jsonschema
            Schema = kSpecSchema.SpecSchema()
            Variants = {V["properties"]["pattern"]["const"]: V
                        for V in Schema["properties"]["slides"]["items"]["oneOf"]}
            Top = dict(Schema, properties=dict(Schema["properties"], slides={"type": "array", "minItems": 1}))
            Out = [f"{kSpecCheck.ErrorPath(E) or 'spec'}: {E.message}"
                   for E in jsonschema.Draft202012Validator(Top).iter_errors(Spec)]
            for I, Sl in enumerate(Spec.get("slides", []), 1):
                V = Variants.get(Sl.get("pattern")) if isinstance(Sl, dict) else None
                if V is None:
                    continue  # unknown pattern: CheckSpec reports it
                for E in jsonschema.Draft202012Validator(V).iter_errors(Sl):
                    Where = kSpecCheck.ErrorPath(E)
                    if E.validator == "additionalProperties":
                        Extra = sorted(set(E.instance) - set(E.schema.get("properties", {})))
                        Out.append(f"slide {I} ({Sl['pattern']}): unknown field {', '.join(repr(X) for X in Extra)}"
                                   + (f" in {Where}" if Where else "") + " - see reference/BUILDER.md")
                    elif E.validator in ("maxLength", "minItems", "maxItems"):
                        continue  # CheckSpec already reports limits, with friendlier wording
                    else:
                        Out.append(f"slide {I} ({Sl['pattern']}): {Where or 'slide'}: {E.message}")
            return sorted(set(Out))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecCheck.SchemaErrors")
            return []


class kSpecLoader:
    """Reads a spec file (JSON or YAML), normalises text fields and resolves relative paths."""

    @staticmethod
    def Texts(Obj, Key=None):
        """YAML turns 42 and 2024 into numbers; every text field must be a string."""
        if kS.ErrorMode:
            return Obj
        try:
            if Key in NUMERIC_KEYS:
                return Obj
            if isinstance(Obj, dict):
                return {K: kSpecLoader.Texts(V, K) for K, V in Obj.items()}
            if isinstance(Obj, list):
                return [kSpecLoader.Texts(V, Key) for V in Obj]
            if isinstance(Obj, (int, float)) and not isinstance(Obj, bool):
                return f"{Obj:g}" if isinstance(Obj, float) else str(Obj)
            return Obj
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecLoader.Texts")
            return Obj

    @staticmethod
    def Load(Path):
        """The spec at Path with image and template paths made absolute (they are relative to the spec)."""
        if kS.ErrorMode:
            return None
        try:
            with open(Path, encoding="utf-8") as Fh:
                if Path.lower().endswith((".yml", ".yaml")):
                    import yaml  # uvx --with pyyaml
                    Spec = yaml.safe_load(Fh)
                else:
                    Spec = json.load(Fh)
            Spec["slides"] = [kSpecLoader.Texts(Sl) for Sl in Spec.get("slides", [])]
            Base = os.path.dirname(os.path.abspath(Path))
            for Sl in Spec.get("slides", []):
                if Sl.get("image") and not os.path.isabs(Sl["image"]):
                    Sl["image"] = os.path.join(Base, Sl["image"])
            if Spec.get("template") and not os.path.isabs(Spec["template"]):
                Spec["template"] = os.path.join(Base, Spec["template"])
            return Spec
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSpecLoader.Load(path={Path})")
            return None

    @staticmethod
    def PrintPlan(Spec):
        """The deck as a story: one line per slide - the claim, then the key fact from the notes."""
        if kS.ErrorMode:
            return
        try:
            Look = Spec.get("template") or Spec.get("direction", "clean-corporate")
            print(f"Deck plan ({len(Spec['slides'])} slides, look: {Look})\n")
            for N, Sl in enumerate(Spec["slides"], 1):
                Title = Sl.get("title") or Sl.get("quote", "")[:60]
                Notes = Sl.get("notes")
                Key = Notes.get("key_fact") if isinstance(Notes, dict) else (Notes or "").split("\n")[0]
                print(f"{N:>2}. [{Sl.get('pattern')}] {Title}")
                if Key:
                    print(f"      {Key[:110]}")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSpecLoader.PrintPlan")
            return


class kDeckDesign:
    """Design directions and templates: the theme a deck is built on."""

    @staticmethod
    def LoadDirection(Name):
        """The direction called Name; an unknown name stops the build (exit 1)."""
        if kS.ErrorMode:
            return None
        try:
            for D in kSpecSchema.Directions():
                if D["id"] == Name:
                    return D
            raise ToolReportableException(f"unknown direction '{Name}'; see --list-directions")
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckDesign.LoadDirection")
            return None

    @staticmethod
    def Mix(Ca, Cb, T):
        """RRGGBB a fraction T of the way from Ca to Cb."""
        if kS.ErrorMode:
            return "000000"
        try:
            La, Lb = [int(Ca[I:I + 2], 16) for I in (0, 2, 4)], [int(Cb[I:I + 2], 16) for I in (0, 2, 4)]
            return "".join(f"{round(X + (Y - X) * T):02X}" for X, Y in zip(La, Lb))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckDesign.Mix")
            return "000000"

    @staticmethod
    def QuietAndMuted(D):
        """The quiet card colour (a light mix of text into the background) and the muted text colour,
        chosen so muted text stays at 4.5:1 or better on both the background and the quiet card."""
        if kS.ErrorMode:
            return None, None
        try:
            Muted = D["muted"]
            for T in (0.12, 0.10, 0.08, 0.06, 0.05):
                Quiet = kDeckDesign.Mix(D["background"], D["text"], T)
                if kRules.ContrastRatio(Muted, Quiet) >= 4.5:
                    return Quiet, Muted
            Quiet = kDeckDesign.Mix(D["background"], D["text"], 0.06)
            for K in range(1, 11):  # darken (or lighten, on dark themes) the muted colour towards the text colour
                Muted = kDeckDesign.Mix(D["muted"], D["text"], K / 10)
                if kRules.ContrastRatio(Muted, Quiet) >= 4.5:
                    break
            return Quiet, Muted
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckDesign.QuietAndMuted")
            return None, None

    @staticmethod
    def ApplyDirection(Prs, D):
        """Write the direction into the theme so everything below uses theme colours and fonts."""
        if kS.ErrorMode:
            return
        try:
            Part = Prs.slide_master.part.part_related_by(RT.THEME)
            Theme = etree.fromstring(Part.blob)
            Scheme = Theme.find(".//a:clrScheme", NS)
            Quiet, Muted = kDeckDesign.QuietAndMuted(D)
            Mix = kDeckDesign.Mix
            Slots = {"dk1": D["text"], "lt1": D["background"], "dk2": Muted,
                     "lt2": Quiet, "accent1": D["accent"],
                     "accent2": Mix(D["accent"], D["background"], 0.45), "accent3": Muted,
                     "accent4": Mix(D["accent"], D["text"], 0.4), "accent5": Mix(D["muted"], D["background"], 0.4),
                     "accent6": Mix(D["accent"], D["background"], 0.7)}
            for Slot, HexV in Slots.items():
                Node = Scheme.find(f"a:{Slot}", NS)
                for Child in list(Node):
                    Node.remove(Child)
                etree.SubElement(Node, f"{{{A}}}srgbClr", val=HexV)
            Scheme.set("name", D["id"])
            Fonts = Theme.find(".//a:fontScheme", NS)
            Fonts.find("a:majorFont/a:latin", NS).set("typeface", D["heading"])
            Fonts.find("a:minorFont/a:latin", NS).set("typeface", D["body"])
            Fonts.set("name", D["id"])
            Part._blob = etree.tostring(Theme, xml_declaration=True, encoding="UTF-8", standalone=True)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckDesign.ApplyDirection")
            return

    @staticmethod
    def OpenTemplate(Path):
        """A template (.pptx, or .potx/.potm with its content type patched) as a Presentation."""
        if kS.ErrorMode:
            return None
        try:
            if Path.lower().endswith((".potx", ".potm")):
                from extract_theme import kThemeReader
                return kThemeReader.OpenAny(Path)
            return Presentation(Path)
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kDeckDesign.OpenTemplate(path={Path})")
            return None

    @staticmethod
    def TitleOnlyLayout(Prs):
        """'Title Only', else the layout with a title and the fewest other text placeholders."""
        if kS.ErrorMode:
            return None
        try:
            for Layout in Prs.slide_layouts:
                if Layout.name.lower().replace("-", " ").startswith("title only"):
                    return Layout
            Best = None
            for Layout in Prs.slide_layouts:
                Types = [Ph.placeholder_format.type for Ph in Layout.placeholders]
                if any(T in TITLE_TYPES for T in Types):
                    Others = sum(1 for T in Types if T not in TITLE_TYPES and T not in FOOTER_TYPES)
                    if Best is None or Others < Best[0]:
                        Best = (Others, Layout)
            if Best is None:
                raise ToolReportableException("the template has no layout with a title placeholder")
            return Best[1]
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckDesign.TitleOnlyLayout")
            return None


class kSlideText:
    """Stateless text helpers: speaker notes, alt text."""

    @staticmethod
    def Alt(Shape, Text):
        """Set a shape's alt text (descr)."""
        if kS.ErrorMode:
            return
        try:
            Shape._element.find(".//p:cNvPr", NS).set("descr", Text)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideText.Alt")
            return

    @staticmethod
    def QaLine(Q):
        """One Q&A item as notes text."""
        if kS.ErrorMode:
            return ""
        try:
            return f"Q: {Q.get('q', '')}\nA: {Q.get('a', '')}" if isinstance(Q, dict) else f"- {Q}"
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideText.QaLine")
            return ""

    @staticmethod
    def NotesText(Notes):
        """Speaker notes as text: a string as is, or the structured fields in a fixed order."""
        if kS.ErrorMode:
            return ""
        try:
            if not Notes:
                return ""
            if isinstance(Notes, str):
                return Notes
            Out = []
            if Notes.get("key_fact"):
                Out.append(f"KEY FACT: {Notes['key_fact']}")
            for Label, Key in (("FACTS", "facts"), ("PITFALLS", "pitfalls"), ("SOURCES", "sources")):
                Items = Notes.get(Key)
                if Items:
                    Items = [Items] if isinstance(Items, str) else Items
                    Out.append(f"{Label}:\n" + "\n".join(f"- {X}" for X in Items))
            Qa = Notes.get("qa")
            if Qa:
                Qa = [Qa] if isinstance(Qa, dict) else Qa
                Out.append("Q&A:\n" + "\n".join(kSlideText.QaLine(Q) for Q in Qa))
            return "\n\n".join(Out)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideText.NotesText")
            return ""

    @staticmethod
    def ChartAlt(Sl):
        """Alt text that reads a chart's data out."""
        if kS.ErrorMode:
            return ""
        try:
            Parts = []
            for Ser in Sl["series"]:
                Pairs = ", ".join(f"{C} {V:g}" if isinstance(V, (int, float)) else f"{C} n/a"
                                  for C, V in zip(Sl["categories"], Ser["values"]))
                Parts.append(f"{Ser['name']}: {Pairs}")
            return f"{Sl.get('type', 'column').title()} chart. " + "; ".join(Parts) + "."
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideText.ChartAlt")
            return ""


class kDeckBuilder:
    """One deck being built: the presentation, its theme fonts, the grid scale and the fit report."""

    def __init__(self, Spec):
        try:
            self.Spec = Spec
            self.Prs = None
            self.Themed = False   # True when a direction was applied (we own fonts/colours)
            self.Layout = None
            self.Major = self.Minor = None
            self.Problems = []    # text that could not be made to fit
            self.SlideNo, self.SlideId = 0, ""
            self._kx = self._ky = 1.0  # template mode: slide size / 1440 x 810, so the grid follows the template
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.__init__")

    def X(self, V):
        """A horizontal 1440-grid length in points on this deck."""
        if kS.ErrorMode:
            return Pt(0)
        try:
            return Pt(V * self._kx)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.X")
            return Pt(0)

    def Y(self, V):
        """A vertical 1440-grid length in points on this deck."""
        if kS.ErrorMode:
            return Pt(0)
        try:
            return Pt(V * self._ky)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.Y")
            return Pt(0)

    def TypeScale(self):
        """The factor type sizes are scaled by: never below two-thirds, so body text stays >= 18 pt."""
        if kS.ErrorMode:
            return 1.0
        try:
            return max(min(self._kx, self._ky), 2 / 3)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.TypeScale")
            return 1.0

    def F(self, V):
        """A font size on the 1440 grid, scaled to this deck."""
        if kS.ErrorMode:
            return Pt(0)
        try:
            return Pt(round(V * self.TypeScale(), 1))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.F")
            return Pt(0)

    def Fit(self, Lines, Size, Wd, Ht, Heading=False, Bold=False, What="text"):
        """Largest size <= Size at which Lines fit a Wd x Ht box (1440-grid points), never below the
        floor. Records a problem when even the floor doesn't fit."""
        if kS.ErrorMode:
            return Size
        try:
            Sentences = any(len(str(X).split()) >= 4 for X in Lines)
            Smallest = (min(Size, 40) if Heading else (FLOOR if Sentences else LABEL_MIN)) if Size > LABEL_MIN else Size
            Family = self.Major if Heading else self.Minor
            K = self.TypeScale()
            Cur = Size
            while True:
                Paras = [(str(X), Family, Cur * K, Bold or Heading, Cur * K * 0.5) for X in Lines]
                Need, Widest, _ = kMeasure.TextHeight(Paras, (Wd - 1) * self._kx)
                if (Need <= Ht * self._ky * 1.02 and Widest <= Wd * self._kx) or Cur <= Smallest:
                    break
                Cur = max(Smallest, Cur - 2)
            if Need > Ht * self._ky * 1.08 or Widest > Wd * self._kx + 1:
                self.Problems.append(f"slide {self.SlideNo} ({self.SlideId}): {What} does not fit even at {Cur:g} pt "
                                     f"(needs ~{Need / self._ky:.0f} pt of {Ht:.0f}); cut words or split the slide")
            return Cur
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.Fit")
            return Size

    def NewSlide(self, Sl, N):
        """A new title-only slide named after the spec id, with every non-title placeholder removed."""
        if kS.ErrorMode:
            return None
        try:
            S = self.Prs.slides.add_slide(self.Layout)
            S._element.find("p:cSld", NS).set("name", Sl.get("id", f"s{N:02d}"))
            if self.Themed:
                S.background.fill.solid()
                S.background.fill.fore_color.theme_color = BG
            for Ph in list(S.placeholders):
                if Ph.placeholder_format.type not in TITLE_TYPES:
                    Ph._element.getparent().remove(Ph._element)
            if S.shapes.title is None:
                raise ToolReportableException(f"layout '{self.Layout.name}' has no title placeholder; "
                                              "the template needs a layout with one")
            return S
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.NewSlide")
            return None

    def Title(self, S, Text, Size=None, Top=TITLE_TOP, Height=TITLE_H, Align=PP_ALIGN.LEFT, Colour=TEXT):
        """Fill and place the slide's title placeholder."""
        if kS.ErrorMode:
            return None
        try:
            Size = self.Fit([Text], Size or SIZE["title"], W - 2 * M - 15, Height - 8, Heading=True, What="title")
            T = S.shapes.title
            T.left, T.top, T.width, T.height = self.X(M), self.Y(Top), self.X(W - 2 * M), self.Y(Height)
            T.text = Text
            Tf = T.text_frame
            Tf.word_wrap = True
            Tf.vertical_anchor = MSO_ANCHOR.BOTTOM
            for Para in Tf.paragraphs:
                Para.alignment = Align
                for R in Para.runs:
                    R.font.size = self.F(Size or SIZE["title"])
                    R.font.bold = True
                    if self.Themed:
                        R.font.color.theme_color = Colour
                        R.font.name = "+mj-lt"
            return T
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.Title")
            return None

    def Text(self, S, Name, Txt, Lx, Ty, Wd, Ht, Size, Colour=TEXT, Bold=False, Align=PP_ALIGN.LEFT,
             Anchor=MSO_ANCHOR.TOP, Heading=False):
        """A named text box, its size fitted to the box."""
        if kS.ErrorMode:
            return None
        try:
            Lines = Txt if isinstance(Txt, list) else [Txt]
            Size = self.Fit(Lines, Size, Wd, Ht, Heading=Heading, Bold=Bold, What=f"'{Name}'")
            Tb = S.shapes.add_textbox(self.X(Lx), self.Y(Ty), self.X(Wd), self.Y(Ht))
            Tb.name = Name
            Tf = Tb.text_frame
            Tf.word_wrap = True
            Tf.vertical_anchor = Anchor
            Tf.margin_left = Pt(0)
            Tf.margin_right = Pt(0)
            for I, Line in enumerate(Lines):
                Para = Tf.paragraphs[0] if I == 0 else Tf.add_paragraph()
                Para.text = Line
                Para.alignment = Align
                if I:
                    Para.space_before = self.F(Size * 0.5)
                for R in Para.runs:
                    R.font.size, R.font.bold = self.F(Size), Bold
                    R.font.color.theme_color = Colour
                    if self.Themed:
                        R.font.name = "+mj-lt" if Heading else "+mn-lt"
            return Tb
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.Text")
            return None

    def Rect(self, S, Name, Lx, Ty, Wd, Ht, Colour=QUIET, Shape=MSO_SHAPE.RECTANGLE):
        """A named, theme-filled shape with no line or shadow."""
        if kS.ErrorMode:
            return None
        try:
            R = S.shapes.add_shape(Shape, self.X(Lx), self.Y(Ty), self.X(Wd), self.Y(Ht))
            R.name = Name
            R.fill.solid()
            R.fill.fore_color.theme_color = Colour
            R.line.fill.background()
            R.shadow.inherit = False
            return R
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.Rect")
            return None

    def OpenDeck(self):
        """The template with its slides dropped, or a blank deck themed with the direction; sets the scale."""
        if kS.ErrorMode:
            return
        try:
            if self.Spec.get("template"):
                self.Prs = kDeckDesign.OpenTemplate(self.Spec["template"])
                for Sld in list(self.Prs.slides._sldIdLst):  # start from the template's masters, not its slides
                    self.Prs.part.drop_rel(Sld.rId)
                    self.Prs.slides._sldIdLst.remove(Sld)
                self.Themed = False
            else:
                self.Prs = Presentation()
                kDeckDesign.ApplyDirection(self.Prs, kDeckDesign.LoadDirection(self.Spec.get("direction",
                                                                                               "clean-corporate")))
                self.Themed = True
            if self.Themed:
                self.Prs.slide_width, self.Prs.slide_height = Pt(W), Pt(H)
                self._kx = self._ky = 1.0
            else:  # keep the template's size; scale the 1440 x 810 grid onto it
                self._kx, self._ky = self.Prs.slide_width / Pt(W), self.Prs.slide_height / Pt(H)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.OpenDeck")
            return

    def Build(self, Out):
        """Build every slide and save to Out. Returns the slide count (fit problems are in self.Problems)."""
        if kS.ErrorMode:
            return 0
        try:
            self.OpenDeck()
            self.Layout = kDeckDesign.TitleOnlyLayout(self.Prs)
            Th = kTheme(self.Prs.slide_master)
            self.Major, self.Minor = Th.Font("+mj-lt"), Th.Font("+mn-lt")
            Patterns = kSlidePatterns(self)
            for N, Sl in enumerate(self.Spec["slides"], 1):
                self.SlideNo, self.SlideId = N, Sl.get("id", f"s{N:02d}")
                S = self.NewSlide(Sl, N)
                getattr(Patterns, PATTERNS[Sl["pattern"]])(S, Sl)
                Text = kSlideText.NotesText(Sl.get("notes"))
                if Text:
                    S.notes_slide.notes_text_frame.text = Text
            if kS.ErrorMode:
                return 0  # a step was reported and halted: do not write a half-built deck
            self.Prs.save(Out)
            return len(self.Spec["slides"])
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kDeckBuilder.Build(out={Out})")
            return 0


class kSlidePatterns:
    """The slide patterns: one method per pattern, drawing onto a slide through a kDeckBuilder."""

    def __init__(self, Builder):
        try:
            self._b = Builder
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.__init__")

    def Title(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl.get("title", ""), Size=SIZE["section"], Top=260, Height=200)
            if Sl.get("subtitle"):
                B.Text(S, "Subtitle", Sl["subtitle"], M, 480, W - 2 * M, 120, SIZE["h2"], MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Title")
            return

    def Section(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            if Sl.get("eyebrow"):
                B.Text(S, "Eyebrow", Sl["eyebrow"].upper(), M, 250, W - 2 * M, 50, SIZE["label"], ACCENT, Bold=True)
            B.Title(S, Sl["title"], Size=SIZE["section"], Top=300, Height=220)
            B.Rect(S, "AccentRule", M, 540, 160, 8, ACCENT)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Section")
            return

    def Statement(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"], Size=SIZE["statement"], Top=200, Height=300)
            if Sl.get("support"):
                B.Text(S, "Support", Sl["support"], M, 540, W - 2 * M, 120, SIZE["h2"], MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Statement")
            return

    def BigNumber(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Num = Sl["number"] + (f" {Sl['unit']}" if Sl.get("unit") else "")
            B.Text(S, "HeroNumber", Num, M, 240, W - 2 * M, 280, SIZE["hero"], ACCENT, Bold=True, Heading=True)
            if Sl.get("caption"):
                B.Text(S, "Caption", Sl["caption"], M, 560, W - 2 * M, 100, SIZE["h2"], MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.BigNumber")
            return

    def Kpi(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Ms = Sl["metrics"]
            Hi = Sl.get("highlight", 0)
            Wd = (W - 2 * M - GAP * (len(Ms) - 1)) / len(Ms)
            for I, Mt in enumerate(Ms):
                Lx = M + I * (Wd + GAP)
                B.Rect(S, f"Card{I + 1}", Lx, 300, Wd, 300, QUIET if I != Hi else ACCENT)
                Fg = BG if I == Hi else TEXT
                B.Text(S, f"Value{I + 1}", Mt["value"], Lx + 24, 340, Wd - 48, 120, SIZE["kpi"], Fg, Bold=True,
                       Heading=True)
                B.Text(S, f"Label{I + 1}", Mt["label"], Lx + 24, 480, Wd - 48, 90, SIZE["small"],
                       Fg if I == Hi else MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Kpi")
            return

    def Bullets(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            B.Text(S, "Points", ["• " + X for X in Sl["items"]], M, BODY_TOP + 20, W - 2 * M - 300,
                   BODY_BOTTOM - BODY_TOP, SIZE["body"])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Bullets")
            return

    def Compare(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Cols = Sl["columns"]
            Hi = Sl.get("highlight")
            Wd = (W - 2 * M - GAP * 2 * (len(Cols) - 1)) / len(Cols)
            for I, C in enumerate(Cols):
                Lx = M + I * (Wd + 2 * GAP)
                B.Rect(S, f"Rule{I + 1}", Lx, 230, Wd, 8, ACCENT if I == Hi else MUTED)
                B.Text(S, f"Heading{I + 1}", C["heading"], Lx, 260, Wd, 60, SIZE["h2"], TEXT, Bold=True, Heading=True)
                B.Text(S, f"Points{I + 1}", C["points"], Lx, 340, Wd, 360, SIZE["caption"], TEXT)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Compare")
            return

    def Process(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Steps = Sl["steps"]
            N = len(Steps)
            Wd = (W - 2 * M - GAP * (N - 1)) / N
            Hi = Sl.get("highlight")
            for I, St in enumerate(Steps):
                Lx = M + I * (Wd + GAP)
                B.Rect(S, f"Step{I + 1}", Lx, 330, Wd, 120, ACCENT if I == Hi else QUIET,
                       MSO_SHAPE.CHEVRON if N <= 5 else MSO_SHAPE.RECTANGLE)
                B.Text(S, f"StepNo{I + 1}", str(I + 1), Lx + (36 if N <= 5 else 16), 350, Wd - 60, 80, SIZE["h2"],
                       BG if I == Hi else TEXT, Bold=True, Anchor=MSO_ANCHOR.MIDDLE, Heading=True)
                B.Text(S, f"StepLabel{I + 1}", St["label"], Lx, 470, Wd, 70,
                       SIZE["caption"] if N <= 5 else SIZE["label"], TEXT, Bold=True)
                if St.get("detail"):
                    B.Text(S, f"StepDetail{I + 1}", St["detail"], Lx, 545, Wd, 140, SIZE["label"], MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Process")
            return

    def Timeline(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Ev = Sl["events"]
            N = len(Ev)
            Ty = 420
            B.Rect(S, "Rail", M, Ty - 3, W - 2 * M, 6, QUIET)
            Step = (W - 2 * M) / N
            Hi = Sl.get("highlight")
            for I, E in enumerate(Ev):
                Cx = M + Step * I + Step / 2
                B.Rect(S, f"Dot{I + 1}", Cx - 14, Ty - 14, 28, 28, ACCENT if I == Hi or Hi is None else MUTED,
                       MSO_SHAPE.OVAL)
                Top = I % 2 == 0
                B.Text(S, f"Date{I + 1}", E["date"], Cx - Step / 2 + 8, Ty - 165 if Top else Ty + 36, Step - 16, 62,
                       SIZE["h2"] - 2, ACCENT, Bold=True, Align=PP_ALIGN.CENTER, Heading=True)
                B.Text(S, f"Event{I + 1}", E["label"], Cx - Step / 2 + 8, Ty - 100 if Top else Ty + 100, Step - 16, 80,
                       SIZE["label"], TEXT, Align=PP_ALIGN.CENTER)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Timeline")
            return

    def Quote(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            # the title stays a real placeholder (outline, screen readers) but reads as a small label
            T = B.Title(S, Sl.get("title") or "In their words", Size=SIZE["label"], Top=56, Height=50, Colour=MUTED)
            T.text_frame.vertical_anchor = MSO_ANCHOR.TOP
            B.Text(S, "QuoteMark", "\u201C", M - 10, 120, 120, 170, 180, ACCENT, Bold=True, Heading=True)
            B.Text(S, "Quote", Sl["quote"], M + 120, 220, W - 2 * M - 240, 360, SIZE["statement"] - 16, TEXT,
                   Heading=True)
            Who = Sl.get("attribution", "") + (f", {Sl['role']}" if Sl.get("role") else "")
            if Who:
                B.Text(S, "Attribution", "\u2014 " + Who, M + 120, 620, W - 2 * M - 240, 60, SIZE["small"], MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Quote")
            return

    def Chart(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Kind = Sl.get("type", "column")
            Data = CategoryChartData()
            Data.categories = Sl["categories"]
            for Ser in Sl["series"]:
                Data.add_series(Ser["name"], Ser["values"])
            HasSide = bool(Sl.get("caption"))
            Cw = W - 2 * M - (380 if HasSide else 0)
            Gf = S.shapes.add_chart(CHART_TYPES[Kind], B.X(M), B.Y(BODY_TOP), B.X(Cw), B.Y(BODY_BOTTOM - BODY_TOP), Data)
            Gf.name = "Chart"
            Ch = Gf.chart
            Single = len(Sl["series"]) == 1
            Hi = Sl.get("highlight")
            if isinstance(Hi, str):
                Hi = Sl["categories"].index(Hi)  # CheckSpec guarantees it exists
            Ch.has_title = False
            Ch.font.size = B.F(SIZE["label"])
            Ch.font.color.theme_color = MUTED
            # legend: none for a single series; on the right otherwise (never top/bottom - it steals plot height)
            Ch.has_legend = not Single and Kind != "pie"
            if Ch.has_legend:
                Ch.legend.position = XL_LEGEND_POSITION.RIGHT
                Ch.legend.include_in_layout = False
            Plot = Ch.plots[0]
            # data labels: pie, or a single series of bars/columns; never on line charts
            Labels = Kind == "pie" or (Single and Kind in ("column", "bar"))
            Plot.has_data_labels = Labels
            if Labels:
                Dl = Plot.data_labels
                Dl.font.size, Dl.font.bold = B.F(SIZE["label"]), True
                if Sl.get("number_format"):
                    Dl.number_format, Dl.number_format_is_linked = Sl["number_format"], False
                if Kind != "pie":
                    Dl.position = XL_LABEL_POSITION.OUTSIDE_END
                else:  # a pie has no axis or legend: name every slice on the slice
                    Dl.show_category_name, Dl.show_value = True, True
                    Dl.position = XL_LABEL_POSITION.OUTSIDE_END
            if Kind in ("column", "bar"):
                Plot.gap_width = 75
            if Kind != "pie":
                Va = Ch.value_axis
                Va.has_major_gridlines = not Labels
                if Va.has_major_gridlines:
                    Va.major_gridlines.format.line.width = Pt(0.5)
                    Va.major_gridlines.format.line.color.theme_color = QUIET
                Va.visible = not Labels
                Va.has_minor_gridlines = False
                if Sl.get("number_format"):
                    Va.tick_labels.number_format, Va.tick_labels.number_format_is_linked = Sl["number_format"], False
                Ca = Ch.category_axis
                Ca.tick_labels.font.size = B.F(SIZE["label"])
                Ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
                Ca.has_major_gridlines = False
            # colour: one series -> highlight one point in the accent, the rest quiet; several -> accent ramp
            for Si, Ser in enumerate(Plot.series):
                if Kind == "line":
                    Ser.format.line.width = Pt(2.25)
                    Ser.format.line.color.theme_color = ACCENT
                    Ser.format.line.color.brightness = RAMP[Si % len(RAMP)]
                    Ser.smooth = False
                    continue
                if Single:
                    NPts = len(Sl["categories"])
                    for Pi, Point in enumerate(Ser.points):
                        Point.format.fill.solid()
                        Point.format.fill.fore_color.theme_color = ACCENT if (Hi is None or Pi == Hi) else QUIET
                        if Kind == "pie" and Hi is None:  # slices must differ: one hue, light to dark
                            Point.format.fill.fore_color.brightness = round(-0.4 + 0.8 * Pi / max(1, NPts - 1), 2)
                else:
                    Ser.format.fill.solid()
                    Ser.format.fill.fore_color.theme_color = ACCENT
                    Ser.format.fill.fore_color.brightness = RAMP[Si % len(RAMP)]
            kSlideText.Alt(Gf, Sl.get("alt") or kSlideText.ChartAlt(Sl))
            if HasSide:
                B.Text(S, "ChartNote", Sl["caption"], W - M - 340, BODY_TOP + 40, 340, 400, SIZE["small"], MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Chart")
            return

    def Table(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Rows, Cols = len(Sl["rows"]) + 1, len(Sl["header"])
            RowH = 56
            Gf = S.shapes.add_table(Rows, Cols, B.X(M), B.Y(BODY_TOP + 20), B.X(W - 2 * M), B.Y(RowH * Rows))
            Gf.name = "Table"
            Tbl = Gf.table
            Hi = Sl.get("highlight_row")
            for C, Head in enumerate(Sl["header"]):
                Cell = Tbl.cell(0, C)
                Cell.text = str(Head)
                Cell.fill.solid()
                Cell.fill.fore_color.theme_color = TEXT
                for R in Cell.text_frame.paragraphs[0].runs:
                    R.font.size, R.font.bold = B.F(SIZE["label"]), True
                    R.font.color.theme_color = BG
            for Ri, Row in enumerate(Sl["rows"], 1):
                for C, Val in enumerate(Row):
                    Cell = Tbl.cell(Ri, C)
                    Cell.text = str(Val)
                    Cell.fill.solid()
                    Cell.fill.fore_color.theme_color = ACCENT if Hi == Ri - 1 else (QUIET if Ri % 2 == 0 else BG)
                    for R in Cell.text_frame.paragraphs[0].runs:
                        R.font.size = B.F(SIZE["label"])
                        R.font.color.theme_color = BG if Hi == Ri - 1 else TEXT
                    if C > 0 and str(Val).replace(",", "").replace(".", "").replace("%", "").replace("-", "").strip().isdigit():
                        Cell.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
            kSlideText.Alt(Gf, Sl.get("alt") or f"Table: {', '.join(str(X) for X in Sl['header'])}; "
                                                 f"{len(Sl['rows'])} rows.")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Table")
            return

    def Image(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            from PIL import Image, ImageOps
            from cover_crop import kCoverCrop
            B = self._b
            B.Title(S, Sl["title"])
            BoxW, BoxH = W - 2 * M, BODY_BOTTOM - BODY_TOP - (90 if Sl.get("caption") else 0)
            Img = ImageOps.exif_transpose(Image.open(Sl["image"]))  # phone photos: honour the rotation tag
            Img = Img.convert("RGBA" if "A" in Img.getbands() else "RGB")  # CMYK/P/L -> something PNG can hold
            Buf = io.BytesIO()
            Crop = kCoverCrop.Crop(Img, BoxW * B._kx, BoxH * B._ky, Sl.get("focus_x", 0.5), Sl.get("focus_y", 0.5))
            Tw = max(Crop.width, 960)  # resample to the exact box ratio: rounding on small images reads as stretch
            Crop.resize((Tw, round(Tw * BoxH * B._ky / (BoxW * B._kx))), Image.LANCZOS).save(Buf, "PNG")
            Buf.seek(0)
            Pic = S.shapes.add_picture(Buf, B.X(M), B.Y(BODY_TOP), B.X(BoxW), B.Y(BoxH))
            Pic.name = "Photo"
            kSlideText.Alt(Pic, Sl.get("alt") or Sl["title"])
            if Sl.get("caption"):
                B.Text(S, "Caption", Sl["caption"], M, BODY_BOTTOM - 70, BoxW, 70, SIZE["label"], MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSlidePatterns.Image(image={Sl.get('image')})")
            return

    def Matrix(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Q = Sl["quadrants"]
            Hi = Sl.get("highlight")
            Left, Top = M + 70, BODY_TOP + 10
            Wd, Ht = (W - Left - M - GAP) / 2, (BODY_BOTTOM - Top - 50 - GAP) / 2
            for I, Quad in enumerate(Q):
                Lx = Left + (I % 2) * (Wd + GAP)
                Ty = Top + (I // 2) * (Ht + GAP)
                B.Rect(S, f"Quadrant{I + 1}", Lx, Ty, Wd, Ht, ACCENT if I == Hi else QUIET)
                Fg = BG if I == Hi else TEXT
                B.Text(S, f"QHeading{I + 1}", Quad["heading"], Lx + 28, Ty + 24, Wd - 56, 50, SIZE["h2"] - 2, Fg,
                       Bold=True, Heading=True)
                B.Text(S, f"QText{I + 1}", Quad.get("text", ""), Lx + 28, Ty + 88, Wd - 56, Ht - 100, SIZE["caption"], Fg)
            if Sl.get("x_axis"):
                B.Text(S, "XAxis", Sl["x_axis"] + " \u2192", Left, BODY_BOTTOM - 40, W - Left - M, 40, SIZE["label"],
                       MUTED, Align=PP_ALIGN.CENTER)
            if Sl.get("y_axis"):
                Tb = B.Text(S, "YAxis", Sl["y_axis"] + " \u2192", M - 200 + 30, Top + Ht, 400, 40, SIZE["label"],
                            MUTED, Align=PP_ALIGN.CENTER)
                Tb.rotation = -90
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Matrix")
            return


class kBuildDeckApp:
    """Command line: --print-schema, --list-directions, --plan, or build (exit 2 spec errors, 3 unfit text)."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
            Ap.add_argument("spec", nargs="?")
            Ap.add_argument("--out")
            Ap.add_argument("--lint", action="store_true", help="run lint_deck.py on the result")
            Ap.add_argument("--force", action="store_true", help="build even if the spec breaks a pattern limit")
            Ap.add_argument("--list-directions", action="store_true")
            Ap.add_argument("--print-schema", action="store_true", help="print the spec's JSON Schema")
            Ap.add_argument("--plan", action="store_true", help="print the story (titles and key facts) and stop")
            Args = Ap.parse_args()
            if Args.print_schema:
                print(json.dumps(kSpecSchema.SpecSchema(), indent=2, ensure_ascii=False))
                return 0
            if Args.list_directions:
                for D in kSpecSchema.Directions():
                    print(f"{D['id']:<18} {D['tone']:<10} {D['heading']} / {D['body']}  #{D['accent']} on "
                          f"#{D['background']}  — {D['mood']}")
                return 0
            if not Args.spec or not (Args.out or Args.plan):
                Ap.error("spec and --out are required (or --plan)")
            Spec = kSpecLoader.Load(Args.spec)
            if Spec is None:
                return 1
            Errors = kSpecCheck.CheckSpec(Spec) + kSpecCheck.SchemaErrors(Spec)
            if Args.plan:
                kSpecLoader.PrintPlan(Spec)
                for E in Errors:
                    print(f"spec: {E}")
                return 2 if Errors else 0
            for E in Errors:
                print(f"spec: {E}", file=sys.stderr)
            if Errors and not Args.force:
                return 2  # spec mistakes (expected state): listed above, nothing written
            Builder = kDeckBuilder(Spec)
            N = Builder.Build(Args.out)
            if kS.ErrorMode:
                return 1
            print(f"{N} slides -> {Args.out}")
            for Pr in Builder.Problems:
                print(f"fit: {Pr}", file=sys.stderr)
            Code = 0
            if Args.lint:
                Code = subprocess.run([sys.executable, os.path.join(HERE, "lint_deck.py"), Args.out]).returncode
            return 3 if Builder.Problems else Code  # 3: text that does not fit (expected state), listed above
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBuildDeckApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kBuildDeckApp)
