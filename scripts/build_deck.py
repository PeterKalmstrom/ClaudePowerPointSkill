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
chart, table, image, matrix, email, kpi_chart, cost_table, quiz, risks. Every pattern lays its content out over
the whole body area and grows text up to a ceiling per role (never below the 27 pt floor for sentences).
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
GROW = {"body": 44, "point": 40, "detail": 36, "heading": 44, "value": 120, "note": 32}  # ceilings text grows to
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
    "kpi": {"title": 70, "metrics": (3, 6), "metrics.value": 12, "metrics.label": 28, "metrics.note": 50},
    "bullets": {"title": 70, "items": (1, 7), "items.*": 100},
    "compare": {"title": 70, "columns": (2, 3), "columns.heading": 40, "columns.points": (1, 4),
                "columns.points.*": 70},
    "process": {"title": 70, "steps": (3, 8), "steps.label": 30, "steps.detail": 60},
    "timeline": {"title": 70, "events": (3, 7), "events.date": 16, "events.label": 40},
    "quote": {"quote": 240, "attribution": 40, "role": 60},
    "chart": {"title": 70, "categories": (2, 24), "series": (1, 6)},
    "table": {"title": 70, "header": (2, 6), "rows": (1, 8)},
    "image": {"title": 70, "caption": 120},
    "matrix": {"title": 70, "quadrants": (4, 4), "quadrants.heading": 30, "quadrants.text": 90,
               "x_axis": 30, "y_axis": 30},
    "email": {"title": 70, "from": 70, "to": 70, "subject": 90, "body": (1, 6), "body.*": 160, "attachment": 40,
              "callouts": (0, 5), "callouts.note": 70},
    "kpi_chart": {"title": 70, "metrics": (2, 4), "metrics.value": 12, "metrics.label": 28, "metrics.note": 40,
                  "categories": (2, 12), "series": (1, 4), "caption": 90},
    "cost_table": {"title": 70, "rows": (1, 7), "rows.item": 40, "rows.detail": 70, "unit": 12, "total_label": 20,
                   "note": 140},
    "quiz": {"title": 90, "options": (2, 4), "options.text": 80, "explain": 160, "answer_title": 55},
    "risks": {"title": 70, "risks": (2, 4), "risks.risk": 50, "risks.mitigation": 120},
}
COMPARE_HEADING = {2: 40, 3: 24}  # compare: heading characters per column count (3 narrow columns hold less)


# ---------------------------------------------------------------- JSON Schema (generated from LIMITS)

FIELDS = {  # pattern: {field: kind}; kinds: text, int, number, list[text], list[obj:{...}], obj
    "title": {"subtitle": "text"},
    "section": {"eyebrow": "text"},
    "statement": {"support": "text"},
    "big_number": {"number": "text", "unit": "text", "caption": "text"},
    "kpi": {"metrics": {"value": "text", "label": "text", "note": "text"}, "highlight": "int"},
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
    "email": {"from": "text", "to": "text", "subject": "text", "body": "list[text]", "attachment": "text",
              "callouts": {"target": "enum:from,to,subject,body,attachment", "line": "int", "note": "text"}},
    "kpi_chart": {"metrics": {"value": "text", "label": "text", "note": "text"}, "highlight_metric": "int",
                  "type": "enum:column,bar,line,pie", "categories": "list[text]",
                  "series": {"name": "text", "values": "list[number]"}, "highlight": "int|text",
                  "number_format": "text", "caption": "text", "alt": "text"},
    "cost_table": {"rows": {"item": "text", "amount": "number", "detail": "text"}, "unit": "text",
                   "total_label": "text", "note": "text", "highlight": "int"},
    "quiz": {"options": {"text": "text", "correct": "bool"}, "explain": "text", "reveal": "enum:notes,slide",
             "answer_title": "text"},
    "risks": {"risks": {"risk": "text", "likelihood": "enum:low,medium,high", "impact": "enum:low,medium,high",
                        "mitigation": "text"}, "highlight": "int"},
}
OPTIONAL_SUB = {"process.detail", "matrix.text", "kpi.note", "kpi_chart.note", "cost_table.detail",
                "email.line"}  # sub-fields of list items that may be left out
REQUIRED = {"title": ["title"], "section": ["title"], "statement": ["title"], "quote": ["quote"],
            "big_number": ["title", "number"], "kpi": ["title", "metrics"], "bullets": ["title", "items"],
            "compare": ["title", "columns"], "process": ["title", "steps"], "timeline": ["title", "events"],
            "chart": ["title", "categories", "series"], "table": ["title", "header", "rows"],
            "image": ["title", "image"], "matrix": ["title", "quadrants"],
            "email": ["title", "from", "subject", "body"], "kpi_chart": ["title", "metrics", "categories", "series"],
            "cost_table": ["title", "rows"], "quiz": ["title", "options"], "risks": ["title", "risks"]}
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
            "matrix": "Matrix", "email": "Email", "kpi_chart": "KpiChart", "cost_table": "CostTable",
            "quiz": "Quiz", "risks": "Risks"}  # pattern -> kSlidePatterns method
NUMERIC_KEYS = {"values", "highlight", "highlight_row", "highlight_metric", "focus_x", "focus_y", "amount", "line",
                "correct"}
HIGHLIGHT_ITEMS = {"kpi": "metrics", "compare": "columns", "process": "steps", "timeline": "events",
                   "matrix": "quadrants", "chart": "categories", "kpi_chart": "categories", "cost_table": "rows",
                   "risks": "risks"}
NEEDED_LIST = {"kpi": "metrics", "process": "steps", "timeline": "events", "compare": "columns",
               "matrix": "quadrants", "chart": "series", "table": "rows", "bullets": "items", "email": "body",
               "kpi_chart": "series", "cost_table": "rows", "quiz": "options", "risks": "risks"}
LEVELS = {"high": "HIGH", "medium": "MEDIUM", "low": "LOW"}


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
            if Kind == "bool":
                return {"type": "boolean"}
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
                            if Sk in ("text", "list[text]", "list[number]", "number") \
                                    and f"{Pat}.{Sub}" not in OPTIONAL_SUB:
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
                if Pat == "compare":
                    Cols = Sl.get("columns", [])
                    Max = COMPARE_HEADING.get(len(Cols), 40)
                    for C, Col in enumerate(Cols):
                        Head = str(Col.get("heading", "")) if isinstance(Col, dict) else ""
                        if len(Head) > Max:
                            Errors.append(f"slide {I} (compare): 'columns.{C}.heading' is {len(Head)} characters; "
                                          f"with {len(Cols)} columns the limit is {Max}: '{Head[:40]}'")
                if Pat == "quiz" and not any(isinstance(O, dict) and O.get("correct") is True
                                             for O in Sl.get("options", [])):
                    Errors.append(f"slide {I} (quiz): mark at least one option \"correct\": true")
                if Pat == "email":
                    for C, Co in enumerate(Sl.get("callouts", [])):
                        if not isinstance(Co, dict):
                            continue
                        if Co.get("target", "body") == "attachment" and not Sl.get("attachment"):
                            Errors.append(f"slide {I} (email): callout {C + 1} points at the attachment, but there is none")
                        Ln = Co.get("line", 0)
                        if Co.get("target", "body") == "body" and not (isinstance(Ln, int) and
                                                                       0 <= Ln < len(Sl.get("body", []))):
                            Errors.append(f"slide {I} (email): callout {C + 1} line {Ln!r} is not a body paragraph "
                                          f"(0-based, {len(Sl.get('body', []))} paragraphs)")
                if Pat == "kpi_chart":
                    Hm = Sl.get("highlight_metric")
                    if Hm is not None and not (isinstance(Hm, int) and 0 <= Hm < len(Sl.get("metrics", []))):
                        Errors.append(f"slide {I} (kpi_chart): highlight_metric {Hm!r} is not a metric index (0-based)")
                if Pat == "cost_table":
                    for R, Row in enumerate(Sl.get("rows", [])):
                        if not isinstance(Row, dict) or isinstance(Row.get("amount"), bool) \
                                or not isinstance(Row.get("amount"), (int, float)):
                            Errors.append(f"slide {I} (cost_table): row {R + 1} needs a numeric 'amount'")
                if Pat in ("chart", "kpi_chart"):
                    Hi = Sl.get("highlight")
                    if isinstance(Hi, str) and Hi not in Sl.get("categories", []):
                        Errors.append(f"slide {I} ({Pat}): highlight '{Hi}' is not one of the categories")
                    if Sl.get("type", "column") not in CHART_TYPES:
                        Errors.append(f"slide {I} ({Pat}): type '{Sl.get('type')}' (one of: {', '.join(CHART_TYPES)})")
                ItemsKey = HIGHLIGHT_ITEMS.get(Pat)
                Hi = Sl.get("highlight")
                if ItemsKey and isinstance(Hi, int) and not 0 <= Hi < len(Sl.get(ItemsKey, [])):
                    Errors.append(f"slide {I} ({Pat}): highlight {Hi} is out of range (0-based, "
                                  f"{len(Sl.get(ItemsKey, []))} items)")
                if Pat in ("chart", "kpi_chart"):
                    Cats = len(Sl.get("categories", []))
                    for S in Sl.get("series", []):
                        if len(S.get("values", [])) != Cats:
                            Errors.append(f"slide {I} ({Pat}): series '{S.get('name')}' has {len(S.get('values', []))} "
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
                else:
                    print("      (no speaker notes yet)")
                if Sl.get("pattern") == "quiz" and Sl.get("reveal") == "slide":
                    print("      + an answer slide after it")
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
            self.NoNotes = []     # slides the spec gave no speaker notes
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

    def Fit(self, Lines, Size, Wd, Ht, Heading=False, Bold=False, What="text", Spacing=0.5):
        """Largest size <= Size at which Lines fit a Wd x Ht box (1440-grid points), never below the
        floor. Sentences (4+ words) start at the floor at least, so a small role size can't put them below it.
        Records a problem when even the floor doesn't fit."""
        if kS.ErrorMode:
            return Size
        try:
            Sentences = any(len(str(X).split()) >= 4 for X in Lines)
            if Sentences and not Heading:
                Size = max(Size, FLOOR)
            Smallest = (min(Size, 40) if Heading else (FLOOR if Sentences else LABEL_MIN)) if Size > LABEL_MIN else Size
            Family = self.Major if Heading else self.Minor
            K = self.TypeScale()
            Cur = Size
            while True:
                Paras = [(str(X), Family, Cur * K, Bold or Heading, Cur * K * Spacing) for X in Lines]
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

    def FitAll(self, Groups, Size, Wd, Ht, Heading=False, Bold=False, What="text", Spacing=0.5):
        """One size for sibling boxes (cards in a row): the smallest of each group's own fit, so they match."""
        if kS.ErrorMode:
            return Size
        try:
            Sizes = [self.Fit(G if isinstance(G, list) else [G], Size, Wd, Ht, Heading, Bold, What, Spacing)
                     for G in Groups if G]
            return min(Sizes) if Sizes else Size
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.FitAll")
            return Size

    def FitLine(self, Texts, Size, Wd, Heading=True, Bold=True, Smallest=40):
        """Largest size <= Size at which every text in Texts stays on one line Wd wide (big values, numbers)."""
        if kS.ErrorMode:
            return Size
        try:
            K = self.TypeScale()
            Family = self.Major if Heading else self.Minor
            Cur = Size
            while Cur > Smallest and max(kMeasure.TextWidth(str(T), Family, Cur * K, Bold or Heading)
                                         for T in Texts) > Wd * 0.9 * self._kx:  # room for renderers' wider fonts
                Cur -= 2
            return Cur
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.FitLine")
            return Size

    def Need(self, Lines, Size, Wd, Heading=False, Bold=False, Spacing=0.5):
        """Height (1440-grid points) that Lines take at Size in a box Wd wide, with a little room for renderers
        that set lines looser (LibreOffice)."""
        if kS.ErrorMode:
            return 0.0
        try:
            Lines = Lines if isinstance(Lines, list) else [Lines]
            K = self.TypeScale()
            Family = self.Major if Heading else self.Minor
            Paras = [(str(X), Family, Size * K, Bold or Heading, Size * K * Spacing) for X in Lines]
            Height, _, _ = kMeasure.TextHeight(Paras, (Wd - 1) * 0.92 * self._kx)  # slack for other renderers
            return Height / self._ky + Size * 0.3
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.Need")
            return 0.0

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
             Anchor=MSO_ANCHOR.TOP, Heading=False, Spacing=0.5):
        """A named text box, its size fitted to the box (Size is the ceiling; Spacing is the gap before each
        paragraph after the first, as a fraction of the size)."""
        if kS.ErrorMode:
            return None
        try:
            Lines = Txt if isinstance(Txt, list) else [Txt]
            Size = self.Fit(Lines, Size, Wd, Ht, Heading=Heading, Bold=Bold, What=f"'{Name}'", Spacing=Spacing)
            Tb = S.shapes.add_textbox(self.X(Lx), self.Y(Ty), self.X(Wd), self.Y(Ht))
            Tb.name = Name
            Tf = Tb.text_frame
            Tf.word_wrap = True
            Tf.vertical_anchor = Anchor
            Tf.margin_left = Pt(0)
            Tf.margin_right = Pt(0)
            Tf.margin_top = Tf.margin_bottom = Pt(2)
            for I, Line in enumerate(Lines):
                Para = Tf.paragraphs[0] if I == 0 else Tf.add_paragraph()
                Marker = Line[:2] if Line.startswith(("\u2022 ", "\u25A0 ")) else ""
                if Marker:  # a bullet: the glyph in the accent, the words in the text colour
                    Lead = Para.add_run()
                    Lead.text = Marker
                    Lead.font.color.theme_color = ACCENT if Colour == TEXT else Colour
                    Body = Para.add_run()
                    Body.text = Line[2:]
                else:
                    Para.text = Line
                Para.alignment = Align
                if I:
                    Para.space_before = self.F(Size * Spacing)
                for Ri, R in enumerate(Para.runs):
                    R.font.size, R.font.bold = self.F(Size), Bold
                    if not (Marker and Ri == 0):
                        R.font.color.theme_color = Colour
                    if self.Themed:
                        R.font.name = "+mj-lt" if Heading else "+mn-lt"
            return Tb
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.Text")
            return None

    def ShapeText(self, Shape, Text, Size, Colour=BG, Bold=True, Heading=True, Inset=0):
        """Text written into a shape itself (a badge's number, a chip's level), centred both ways."""
        if kS.ErrorMode:
            return None
        try:
            Tf = Shape.text_frame
            Tf.word_wrap = True
            Tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            Tf.margin_top = Tf.margin_bottom = Pt(0)
            Tf.margin_left = Tf.margin_right = self.X(Inset)
            Tf.text = Text
            Para = Tf.paragraphs[0]
            Para.alignment = PP_ALIGN.CENTER
            for R in Para.runs:
                R.font.size, R.font.bold = self.F(Size), Bold
                R.font.color.theme_color = Colour
                if self.Themed:
                    R.font.name = "+mj-lt" if Heading else "+mn-lt"
            return Shape
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.ShapeText")
            return None

    def Rect(self, S, Name, Lx, Ty, Wd, Ht, Colour=QUIET, Shape=MSO_SHAPE.RECTANGLE, Line=None):
        """A named, theme-filled shape with no shadow; no outline unless Line (a theme colour) is given."""
        if kS.ErrorMode:
            return None
        try:
            R = S.shapes.add_shape(Shape, self.X(Lx), self.Y(Ty), self.X(Wd), self.Y(Ht))
            R.name = Name
            R.fill.solid()
            R.fill.fore_color.theme_color = Colour
            if Line is None:
                R.line.fill.background()
            else:
                R.line.color.theme_color = Line
                R.line.width = Pt(1.5)
            R.shadow.inherit = False
            Ref = R._element.find(".//a:effectRef", NS)
            if Ref is not None:  # the shape style's theme effect is a shadow in some renderers: none, flat cards
                Ref.set("idx", "0")
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

    def Notes(self, S, Sl, Extra=""):
        """Write the slide's speaker notes (the spec's, then what the pattern adds); remember slides left without."""
        if kS.ErrorMode:
            return
        try:
            Text = "\n\n".join(X for X in (kSlideText.NotesText(Sl.get("notes")), Extra) if X)
            if Text.strip():
                S.notes_slide.notes_text_frame.text = Text
            if not kSlideText.NotesText(Sl.get("notes")).strip():
                self.NoNotes.append(f"slide {self.SlideNo} ({self.SlideId})")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDeckBuilder.Notes")
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
            N = 0
            for Sl in self.Spec["slides"]:
                N += 1
                self.SlideNo, self.SlideId = N, Sl.get("id", f"s{N:02d}")
                S = self.NewSlide(Sl, N)
                getattr(Patterns, PATTERNS[Sl["pattern"]])(S, Sl)
                self.Notes(S, Sl, kSlidePatterns.ExtraNotes(Sl))
                if Sl["pattern"] == "quiz" and Sl.get("reveal") == "slide":  # the answer on a slide of its own
                    N += 1
                    Answer = dict(Sl, id=f"{self.SlideId}-answer")
                    self.SlideNo, self.SlideId = N, Answer["id"]
                    S = self.NewSlide(Answer, N)
                    Patterns.QuizAnswer(S, Answer)
                    self.Notes(S, Answer, kSlidePatterns.ExtraNotes(Sl))
            if kS.ErrorMode:
                return 0  # a step was reported and halted: do not write a half-built deck
            self.Prs.save(Out)
            return N
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kDeckBuilder.Build(out={Out})")
            return 0


class kSlidePatterns:
    """The slide patterns: one method per pattern, drawing onto a slide through a kDeckBuilder. Each one lays its
    content out over the whole body area (BODY_TOP..BODY_BOTTOM, margin to margin) and grows text up to the
    ceilings in GROW, so a slide with little to say still reads from the back of the room."""

    def __init__(self, Builder):
        try:
            self._b = Builder
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.__init__")

    @staticmethod
    def ExtraNotes(Sl):
        """Notes the pattern adds after the spec's own: an email's callouts, a quiz's answer."""
        if kS.ErrorMode:
            return ""
        try:
            Pat = Sl.get("pattern")
            if Pat == "email" and Sl.get("callouts"):
                return "CALLOUTS:\n" + "\n".join(f"{I}. {C.get('note', '')} ({C.get('target', 'body')})"
                                                 for I, C in enumerate(Sl["callouts"], 1))
            if Pat == "quiz":
                Right = [f"{'ABCD'[I]}: {O.get('text', '')}" for I, O in enumerate(Sl["options"]) if O.get("correct")]
                return "ANSWER: " + "; ".join(Right) + (f"\nWHY: {Sl['explain']}" if Sl.get("explain") else "")
            return ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.ExtraNotes")
            return ""

    def Badge(self, S, Name, Text, Lx, Ty, D, Fill=ACCENT, Fg=BG, Size=28):
        """A numbered or lettered circle: one shape with its text inside, so nothing can drift apart."""
        if kS.ErrorMode:
            return None
        try:
            B = self._b
            Dot = B.Rect(S, Name, Lx, Ty, D, D, Fill, MSO_SHAPE.OVAL)
            B.ShapeText(Dot, Text, Size, Fg)
            return Dot
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Badge")
            return None

    def Title(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Rect(S, "AccentRule", M, 236, 160, 10, ACCENT)
            B.Title(S, Sl.get("title", ""), Size=SIZE["section"], Top=262, Height=230)
            if Sl.get("subtitle"):
                B.Text(S, "Subtitle", Sl["subtitle"], M, 516, W - 2 * M - 160, 130, SIZE["h2"], MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Title")
            return

    def Section(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            if Sl.get("eyebrow"):
                B.Text(S, "Eyebrow", Sl["eyebrow"].upper(), M, 250, W - 2 * M, 50, SIZE["caption"], ACCENT, Bold=True)
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
            B.Rect(S, "AccentRule", M, 196, 160, 10, ACCENT)
            B.Title(S, Sl["title"], Size=SIZE["statement"], Top=220, Height=300)
            if Sl.get("support"):
                B.Text(S, "Support", Sl["support"], M, 550, W - 2 * M - 120, 150, SIZE["h2"], MUTED)
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
            B.Text(S, "HeroNumber", Num, M, 220, W - 2 * M, 300, SIZE["hero"] + 20, ACCENT, Bold=True, Heading=True)
            if Sl.get("caption"):
                B.Text(S, "Caption", Sl["caption"], M, 550, W - 2 * M - 120, 150, SIZE["h2"], TEXT)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.BigNumber")
            return

    def Tiles(self, S, Ms, Hi, Lx, Ty, Wd, Ht, Across, ValueMax, LabelMax):
        """Metric tiles (value, label, optional note) in a row (Across) or a column, filling Wd x Ht; the
        highlighted one in the accent. Values, labels and notes share one size each and line up."""
        if kS.ErrorMode:
            return
        try:
            B = self._b
            N = len(Ms)
            Tw = (Wd - GAP * (N - 1)) / N if Across else Wd
            Th = Ht if Across else (Ht - GAP * (N - 1)) / N
            Pad = 32 if Across else 28
            Iw = Tw - 2 * Pad
            Notes = [str(X.get("note", "")) for X in Ms]
            Vs = B.FitLine([X["value"] for X in Ms], ValueMax, Iw)
            Ls = B.FitAll([str(X["label"]) for X in Ms], LabelMax, Iw, Th * (0.28 if Across else 0.3))
            Ns = B.FitAll([X for X in Notes if X], GROW["note"] - 2, Iw, Th * 0.22) if any(Notes) else 0
            while True:  # the value, label and note must fit the tile together: shrink the largest first
                Hv = max(B.Need(str(X["value"]), Vs, Iw, Heading=True) for X in Ms)
                Hl = max(B.Need(str(X["label"]), Ls, Iw) for X in Ms)
                Hn = max(B.Need(X, Ns, Iw) for X in Notes if X) if Ns else 0
                Block = Hv + 10 + Hl + (14 + Hn if Ns else 0)
                if Block <= Th - Pad * 0.8 or (Vs <= 40 and Ls <= LABEL_MIN + 2):
                    break
                if Vs > 40:
                    Vs -= 4
                else:
                    Ls, Ns = Ls - 2, (max(LABEL_MIN, Ns - 2) if Ns else 0)
            for I, Mt in enumerate(Ms):
                X0 = Lx + (I * (Tw + GAP) if Across else 0)
                Y0 = Ty + (0 if Across else I * (Th + GAP))
                On = I == Hi
                B.Rect(S, f"Card{I + 1}", X0, Y0, Tw, Th, ACCENT if On else QUIET)
                Fg = BG if On else TEXT
                Top = Y0 + max(Pad * 0.6, (Th - Block) / 2)
                B.Text(S, f"Value{I + 1}", str(Mt["value"]), X0 + Pad, Top, Iw, Hv, Vs, Fg, Bold=True, Heading=True)
                B.Text(S, f"Label{I + 1}", str(Mt["label"]), X0 + Pad, Top + Hv + 10, Iw, Hl, Ls, Fg)
                if Notes[I]:
                    B.Text(S, f"Note{I + 1}", Notes[I], X0 + Pad, Top + Hv + 10 + Hl + 14, Iw, Hn, Ns,
                           BG if On else MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Tiles")
            return

    def Kpi(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Top = BODY_TOP + 24
            self.Tiles(S, Sl["metrics"], Sl.get("highlight", 0), M, Top, W - 2 * M, BODY_BOTTOM - Top - 6, True,
                       GROW["value"], 40)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Kpi")
            return

    def Bullets(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            if len(Sl["items"]) <= 4 and all(len(str(X)) <= 80 for X in Sl["items"]):
                self.BulletRows(S, Sl)  # a few short points: one band each, filling the slide
                return
            Items = ["\u2022 " + str(X) for X in Sl["items"]]
            Wd, Ht = W - 2 * M - 120, BODY_BOTTOM - BODY_TOP - 20
            Size = B.Fit(Items, GROW["body"], Wd, Ht)
            Spare = max(0.0, Ht - B.Need(Items, Size, Wd))
            Spacing = 0.5 + (min(1.1, Spare / (len(Items) - 1) / Size * 0.7) if len(Items) > 1 else 0)
            B.Text(S, "Points", Items, M, BODY_TOP + 20, Wd, Ht, Size, Spacing=Spacing)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Bullets")
            return

    def BulletRows(self, S, Sl):
        """Up to four short points as full-width bands with an accent edge, sized to fill the body."""
        if kS.ErrorMode:
            return
        try:
            B = self._b
            Items = [str(X) for X in Sl["items"]]
            N = len(Items)
            Ht = min(150, (BODY_BOTTOM - BODY_TOP - 10 - GAP * (N - 1)) / N)
            Pad = 40
            Tw = W - 2 * M - 14 - 2 * Pad
            Size = B.FitAll(Items, GROW["body"], Tw, Ht - 20)
            for I, Item in enumerate(Items):
                Ty = BODY_TOP + 10 + I * (Ht + GAP)
                B.Rect(S, f"PointBand{I + 1}", M, Ty, W - 2 * M, Ht, QUIET)
                B.Rect(S, f"PointEdge{I + 1}", M, Ty, 14, Ht, ACCENT)
                B.Text(S, f"Point{I + 1}", Item, M + 14 + Pad, Ty + 10, Tw, Ht - 20, Size, TEXT,
                       Anchor=MSO_ANCHOR.MIDDLE)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.BulletRows")
            return

    def Compare(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Cols = Sl["columns"]
            Hi = Sl.get("highlight")
            N = len(Cols)
            Gap = GAP * 1.5
            Wd = (W - 2 * M - Gap * (N - 1)) / N
            Pad = 32
            Iw = Wd - 2 * Pad
            BodyH = BODY_BOTTOM - BODY_TOP - 10
            Hs = B.FitAll([str(C["heading"]) for C in Cols], GROW["heading"] - 4, Iw, 120, Heading=True)
            Hh = max(B.Need(str(C["heading"]), Hs, Iw, Heading=True) for C in Cols)
            Room = BodyH - 12 - 2 * Pad - Hh - 20
            Points = [["• " + str(P) for P in C["points"]] for C in Cols]
            Ps = B.FitAll(Points, GROW["point"], Iw, Room)
            Ph = max(B.Need(P, Ps, Iw) for P in Points)
            Most = max(len(P) for P in Points)
            Spacing = 0.5 + (min(0.9, max(0.0, Room - Ph) / (Most - 1) / Ps * 0.6) if Most > 1 else 0)
            Ph = min(max(B.Need(P, Ps, Iw, Spacing=Spacing) for P in Points), Room)  # overflow stays visible to lint
            CardH = min(BodyH, max(12 + 2 * Pad + Hh + 20 + Ph, BodyH * 0.78))
            Ty = BODY_TOP + 10 + (BodyH - CardH) / 2
            for I, C in enumerate(Cols):
                Lx = M + I * (Wd + Gap)
                B.Rect(S, f"Card{I + 1}", Lx, Ty, Wd, CardH, QUIET)
                B.Rect(S, f"Rule{I + 1}", Lx, Ty, Wd, 12, ACCENT if I == Hi else MUTED)
                B.Text(S, f"Heading{I + 1}", str(C["heading"]), Lx + Pad, Ty + 12 + Pad, Iw, Hh, Hs, TEXT, Bold=True,
                       Heading=True)
                B.Text(S, f"Points{I + 1}", Points[I], Lx + Pad, Ty + 12 + Pad + Hh + 20, Iw, Ph, Ps, TEXT,
                       Spacing=Spacing)
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
            if len(Steps) > 5:
                self.ProcessGrid(S, Sl)
                return
            N = len(Steps)
            Hi = Sl.get("highlight")
            Wd = (W - 2 * M - GAP * (N - 1)) / N
            Pad = 26
            Iw = Wd - 2 * Pad
            ChevH, BodyH = 96, BODY_BOTTOM - BODY_TOP - 10
            Ls = B.FitAll([str(X["label"]) for X in Steps], 38, Iw, 130, Bold=True)
            Lh = max(B.Need(str(X["label"]), Ls, Iw, Bold=True) for X in Steps)
            Details = [str(X.get("detail", "")) for X in Steps]
            Room = BodyH - ChevH - 20 - 2 * Pad - Lh - 14
            Ds = B.FitAll([X for X in Details if X], min(GROW["detail"], max(Ls + 4, FLOOR)), Iw, Room) \
                if any(Details) else 0  # never louder than the label above it
            Dh = max(B.Need(X, Ds, Iw) for X in Details if X) if Ds else 0
            CardH = min(BodyH - ChevH - 20, max(2 * Pad + Lh + (14 + Dh if Ds else 0), BodyH * 0.84 - ChevH - 20))
            Top = BODY_TOP + 10 + (BodyH - ChevH - 20 - CardH) / 2
            Depth = 0.32
            for I, St in enumerate(Steps):
                Lx = M + I * (Wd + GAP)
                On = I == Hi
                Arrow = B.Rect(S, f"Step{I + 1}", Lx, Top, Wd + (GAP * 0.6 if I < N - 1 else 0), ChevH,
                               ACCENT if On else QUIET, MSO_SHAPE.PENTAGON if I == 0 else MSO_SHAPE.CHEVRON)
                Arrow.adjustments[0] = Depth
                # the number sits in the arrow's body, clear of the notch and the point
                B.ShapeText(Arrow, str(I + 1), SIZE["h2"], BG if On else TEXT, Inset=ChevH * Depth + 6)
                Card = Top + ChevH + 20
                B.Rect(S, f"StepCard{I + 1}", Lx, Card, Wd, CardH, QUIET, Line=ACCENT if On else None)
                B.Text(S, f"StepLabel{I + 1}", str(St["label"]), Lx + Pad, Card + Pad, Iw, Lh, Ls, TEXT, Bold=True)
                if Details[I]:
                    B.Text(S, f"StepDetail{I + 1}", Details[I], Lx + Pad, Card + Pad + Lh + 14, Iw, Dh, Ds, TEXT)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Process")
            return

    def ProcessGrid(self, S, Sl):
        """Six to eight steps: numbered cards in two rows, so labels keep a readable size."""
        if kS.ErrorMode:
            return
        try:
            B = self._b
            Steps = Sl["steps"]
            Hi = Sl.get("highlight")
            Cols = -(-len(Steps) // 2)
            Wd = (W - 2 * M - GAP * (Cols - 1)) / Cols
            Ht = (BODY_BOTTOM - BODY_TOP - 10 - GAP) / 2
            Pad, D = 24, 56
            Iw = Wd - 2 * Pad
            Ls = B.FitAll([str(X["label"]) for X in Steps], 34, Iw, 90, Bold=True)
            Lh = max(B.Need(str(X["label"]), Ls, Iw, Bold=True) for X in Steps)
            Details = [str(X.get("detail", "")) for X in Steps]
            Room = Ht - 2 * Pad - D - 14 - Lh - 10
            Ds = B.FitAll([X for X in Details if X], 30, Iw, Room) if any(Details) else 0
            for I, St in enumerate(Steps):
                Lx = M + (I % Cols) * (Wd + GAP)
                Ty = BODY_TOP + 10 + (I // Cols) * (Ht + GAP)
                On = I == Hi
                B.Rect(S, f"Step{I + 1}", Lx, Ty, Wd, Ht, ACCENT if On else QUIET)
                self.Badge(S, f"StepNo{I + 1}", str(I + 1), Lx + Pad, Ty + Pad, D, BG if On else ACCENT,
                           ACCENT if On else BG)
                Fg = BG if On else TEXT
                B.Text(S, f"StepLabel{I + 1}", str(St["label"]), Lx + Pad, Ty + Pad + D + 14, Iw, Lh, Ls, Fg, Bold=True)
                if Details[I]:
                    B.Text(S, f"StepDetail{I + 1}", Details[I], Lx + Pad, Ty + Pad + D + 14 + Lh + 10, Iw, Room, Ds, Fg)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.ProcessGrid")
            return

    def Timeline(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Ev = Sl["events"]
            N = len(Ev)
            Ry = (BODY_TOP + BODY_BOTTOM) / 2 + 10
            Step = (W - 2 * M) / N
            Lw = min(2 * Step - 32, 440) if N > 2 else Step - 32  # neighbours sit on the other side of the rail
            Hi = Sl.get("highlight")
            Centres = [M + Step * I + Step / 2 for I in range(N)]
            Widths = [min(Lw, 2 * (Cx - M), 2 * (W - M - Cx)) for Cx in Centres]  # centred on the dot, on the slide
            Ds = min(B.Fit([str(E["date"])], 54, Widths[I], 70, Heading=True, Bold=True) for I, E in enumerate(Ev))
            Dh = max(B.Need(str(E["date"]), Ds, Widths[I], Heading=True) for I, E in enumerate(Ev))
            Room = Ry - 30 - Dh - 8 - BODY_TOP
            Ls = min(B.Fit([str(E["label"])], 40, Widths[I], Room) for I, E in enumerate(Ev))
            Lh = max(B.Need(str(E["label"]), Ls, Widths[I]) for I, E in enumerate(Ev))
            B.Rect(S, "Rail", M, Ry - 4, W - 2 * M, 8, QUIET)
            for I, E in enumerate(Ev):
                Cx, Wd = Centres[I], Widths[I]
                On = I == Hi or Hi is None
                D = 48 if I == Hi else 32
                B.Rect(S, f"Dot{I + 1}", Cx - D / 2, Ry - D / 2, D, D, ACCENT if On else MUTED, MSO_SHAPE.OVAL)
                Lx = Cx - Wd / 2
                if I % 2 == 0:
                    B.Text(S, f"Date{I + 1}", str(E["date"]), Lx, Ry - 30 - Dh, Wd, Dh, Ds, ACCENT, Bold=True,
                           Align=PP_ALIGN.CENTER, Anchor=MSO_ANCHOR.BOTTOM, Heading=True)
                    B.Text(S, f"Event{I + 1}", str(E["label"]), Lx, Ry - 30 - Dh - 8 - Lh, Wd, Lh, Ls, TEXT,
                           Align=PP_ALIGN.CENTER, Anchor=MSO_ANCHOR.BOTTOM)
                else:
                    B.Text(S, f"Date{I + 1}", str(E["date"]), Lx, Ry + 30, Wd, Dh, Ds, ACCENT, Bold=True,
                           Align=PP_ALIGN.CENTER, Heading=True)
                    B.Text(S, f"Event{I + 1}", str(E["label"]), Lx, Ry + 30 + Dh + 8, Wd, Lh, Ls, TEXT,
                           Align=PP_ALIGN.CENTER)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Timeline")
            return

    def Quote(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            # the title stays a real placeholder (outline, screen readers) but reads as a small label
            T = B.Title(S, Sl.get("title") or "In their words", Size=SIZE["caption"], Top=56, Height=50, Colour=MUTED)
            T.text_frame.vertical_anchor = MSO_ANCHOR.TOP
            B.Text(S, "QuoteMark", "“", M - 10, 120, 120, 170, 180, ACCENT, Bold=True, Heading=True)
            B.Text(S, "Quote", Sl["quote"], M + 120, 220, W - 2 * M - 240, 360, SIZE["statement"] - 12, TEXT,
                   Heading=True)
            Who = Sl.get("attribution", "") + (f", {Sl['role']}" if Sl.get("role") else "")
            if Who:
                B.Text(S, "Attribution", "— " + Who, M + 120, 620, W - 2 * M - 240, 60, SIZE["small"], MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Quote")
            return

    def DrawChart(self, S, Sl, Lx, Ty, Wd, Ht):
        """A native chart of Sl's categories and series in the box: one highlighted finding, the rest quiet."""
        if kS.ErrorMode:
            return None
        try:
            B = self._b
            Kind = Sl.get("type", "column")
            Data = CategoryChartData()
            Data.categories = Sl["categories"]
            for Ser in Sl["series"]:
                Data.add_series(Ser["name"], Ser["values"])
            Gf = S.shapes.add_chart(CHART_TYPES[Kind], B.X(Lx), B.Y(Ty), B.X(Wd), B.Y(Ht), Data)
            Gf.name = "Chart"
            Ch = Gf.chart
            Single = len(Sl["series"]) == 1
            Hi = Sl.get("highlight")
            if isinstance(Hi, str):
                Hi = Sl["categories"].index(Hi)  # CheckSpec guarantees it exists
            Ch.has_title = False
            Ch.font.size = B.F(26)
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
                Dl.font.size, Dl.font.bold = B.F(30), True
                Dl.font.color.theme_color = TEXT
                if Sl.get("number_format"):
                    Dl.number_format, Dl.number_format_is_linked = Sl["number_format"], False
                if Kind != "pie":
                    Dl.position = XL_LABEL_POSITION.OUTSIDE_END
                else:  # a pie has no axis or legend: name every slice on the slice
                    Dl.show_category_name, Dl.show_value = True, True
                    Dl.position = XL_LABEL_POSITION.OUTSIDE_END
            if Kind in ("column", "bar"):
                Plot.gap_width = 60
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
                Ca.tick_labels.font.size = B.F(28)
                if Kind == "bar":  # bars list top-down in the order given, like a table
                    Ca.reverse_order = True
                Ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
                Ca.has_major_gridlines = False
            # colour: one series -> highlight one point in the accent, the rest quiet; several -> accent ramp
            for Si, Ser in enumerate(Plot.series):
                if Kind == "line":
                    Ser.format.line.width = Pt(3)
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
            return Gf
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.DrawChart")
            return None

    def Chart(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            HasSide = bool(Sl.get("caption"))
            Cw = W - 2 * M - (420 if HasSide else 0)
            self.DrawChart(S, Sl, M, BODY_TOP, Cw, BODY_BOTTOM - BODY_TOP)
            if HasSide:
                Px, Pw = W - M - 380, 380
                Size = B.Fit([Sl["caption"]], 38, Pw, 380)
                Ht = B.Need(Sl["caption"], Size, Pw)
                Top = BODY_TOP + max(30, (BODY_BOTTOM - BODY_TOP - Ht - 30) / 2)
                B.Rect(S, "NoteRule", Px, Top, 120, 8, ACCENT)
                B.Text(S, "ChartNote", Sl["caption"], Px, Top + 30, Pw, Ht, Size, TEXT)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Chart")
            return

    def KpiChart(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Tw = 420
            Bottom = BODY_BOTTOM - (70 if Sl.get("caption") else 0)
            self.Tiles(S, Sl["metrics"], Sl.get("highlight_metric"), M, BODY_TOP + 10, Tw, BODY_BOTTOM - BODY_TOP - 10,
                       False, 80, 32)
            Cx = M + Tw + 48
            self.DrawChart(S, Sl, Cx, BODY_TOP, W - M - Cx, Bottom - BODY_TOP)
            if Sl.get("caption"):
                B.Text(S, "ChartNote", Sl["caption"], Cx, Bottom + 10, W - M - Cx, 60, 30, MUTED)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.KpiChart")
            return

    @staticmethod
    def IsNumber(Val):
        """True for a table cell that reads as a number (right-aligned)."""
        if kS.ErrorMode:
            return False
        try:
            return str(Val).replace(",", "").replace(".", "").replace("%", "").replace("-", "").replace(
                " ", "").replace(" ", "").strip().isdigit()
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.IsNumber")
            return False

    def Cell(self, Tbl, R, C, Val, Size, Fill, Fg, Bold=False, Right=False):
        """Fill one table cell: text, size, colours, generous margins, centred vertically."""
        if kS.ErrorMode:
            return
        try:
            B = self._b
            Cell = Tbl.cell(R, C)
            Cell.text = str(Val)
            Cell.fill.solid()
            Cell.fill.fore_color.theme_color = Fill
            Cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            Cell.margin_left = Cell.margin_right = B.X(16)
            Cell.margin_top = Cell.margin_bottom = B.Y(4)
            Para = Cell.text_frame.paragraphs[0]
            if Right:
                Para.alignment = PP_ALIGN.RIGHT
            for Run in Para.runs:
                Run.font.size, Run.font.bold = B.F(Size), Bold
                Run.font.color.theme_color = Fg
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Cell")
            return

    def Table(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Rows, Cols = len(Sl["rows"]) + 1, len(Sl["header"])
            RowH = min(84, (BODY_BOTTOM - BODY_TOP - 20) / Rows)
            Size = 32 if RowH >= 72 else (30 if RowH >= 62 else 27)
            Gf = S.shapes.add_table(Rows, Cols, B.X(M), B.Y(BODY_TOP + 10), B.X(W - 2 * M), B.Y(RowH * Rows))
            Gf.name = "Table"
            Tbl = Gf.table
            Longest = [max([len(str(Sl["header"][C]))] + [len(str(R[C])) for R in Sl["rows"]]) for C in range(Cols)]
            Share = [max(L, 6) for L in Longest]
            for C in range(Cols):
                Tbl.columns[C].width = B.X((W - 2 * M) * Share[C] / sum(Share))
            for R in range(Rows):
                Tbl.rows[R].height = B.Y(RowH)
            Hi = Sl.get("highlight_row")
            for C, Head in enumerate(Sl["header"]):
                self.Cell(Tbl, 0, C, Head, Size, TEXT, BG, Bold=True, Right=C > 0 and self.IsNumber(Sl["rows"][0][C]))
            for Ri, Row in enumerate(Sl["rows"], 1):
                for C, Val in enumerate(Row):
                    On = Hi == Ri - 1
                    self.Cell(Tbl, Ri, C, Val, Size, ACCENT if On else (QUIET if Ri % 2 == 0 else BG),
                              BG if On else TEXT, Bold=On, Right=C > 0 and self.IsNumber(Val))
            kSlideText.Alt(Gf, Sl.get("alt") or f"Table: {', '.join(str(X) for X in Sl['header'])}; "
                                                 f"{len(Sl['rows'])} rows.")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Table")
            return

    @staticmethod
    def Amount(Val, Decimals):
        """A cost as text: thousands grouped with a no-break space, a fixed number of decimals."""
        if kS.ErrorMode:
            return ""
        try:
            return f"{Val:,.{Decimals}f}".replace(",", " ")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Amount")
            return ""

    def CostTable(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Items = Sl["rows"]
            Unit = Sl.get("unit", "")
            Detail = any(R.get("detail") for R in Items)
            Decimals = max(len(f"{R['amount']:g}".split(".")[1]) if "." in f"{R['amount']:g}" else 0 for R in Items)
            Total = sum(R["amount"] for R in Items)
            Tw = 960
            Rows = len(Items) + 2
            RowH = min(104, (BODY_BOTTOM - BODY_TOP - 10) / Rows)
            Size = 32 if RowH >= 72 else (30 if RowH >= 62 else 27)
            if Detail:  # a detail that wraps needs two lines in its row
                Size = min(Size, B.FitAll([str(R.get("detail", "")) for R in Items], Size, Tw - 270 - 190 - 32,
                                          RowH - 8))
            Cols = 3 if Detail else 2
            Ty = BODY_TOP + 10
            Gf = S.shapes.add_table(Rows, Cols, B.X(M), B.Y(Ty), B.X(Tw), B.Y(RowH * Rows))
            Gf.name = "CostTable"
            Tbl = Gf.table
            Widths = [270, Tw - 270 - 190, 190] if Detail else [Tw - 240, 240]
            for C, Wd in enumerate(Widths):
                Tbl.columns[C].width = B.X(Wd)
            for R in range(Rows):
                Tbl.rows[R].height = B.Y(RowH)
            Head = ["Item"] + (["What it pays for"] if Detail else []) + [Unit or "Cost"]
            for C, Text in enumerate(Head):
                self.Cell(Tbl, 0, C, Text, Size, TEXT, BG, Bold=True, Right=C == Cols - 1)
            Hi = Sl.get("highlight")
            for Ri, Row in enumerate(Items, 1):
                On = Hi == Ri - 1
                Fill, Fg = (ACCENT, BG) if On else ((QUIET if Ri % 2 == 0 else BG), TEXT)
                Vals = [Row["item"]] + ([Row.get("detail", "")] if Detail else []) + [self.Amount(Row["amount"],
                                                                                                    Decimals)]
                for C, Val in enumerate(Vals):
                    self.Cell(Tbl, Ri, C, Val, Size, Fill, Fg, Bold=On, Right=C == Cols - 1)
            Last = Rows - 1
            Vals = [Sl.get("total_label", "Total")] + ([""] if Detail else []) + [self.Amount(Total, Decimals)]
            for C, Val in enumerate(Vals):
                self.Cell(Tbl, Last, C, Val, Size, TEXT, BG, Bold=True, Right=C == Cols - 1)
            kSlideText.Alt(Gf, Sl.get("alt") or "Cost table: " + "; ".join(
                f"{R['item']} {self.Amount(R['amount'], Decimals)}" for R in Items)
                + f"; {Sl.get('total_label', 'Total')} {self.Amount(Total, Decimals)} {Unit}.")
            # the total, large, beside the table - the number the decision is about
            Px = M + Tw + 56
            Pw = W - M - Px
            Big = self.Amount(Total, Decimals)
            Bs = B.FitLine([Big], 110, Pw)
            Bh = B.Need(Big, Bs, Pw, Heading=True)
            Note = Sl.get("note", "")
            Ns = B.Fit([Note], 32, Pw, 260) if Note else 0
            Nh = B.Need(Note, Ns, Pw) if Note else 0
            Uh = 52 if Unit else 0
            Block = 40 + 8 + Bh + Uh + (24 + Nh if Note else 0)
            Top = Ty + max(0, (RowH * Rows - Block) / 2)
            B.Text(S, "TotalLabel", Sl.get("total_label", "Total"), Px, Top, Pw, 40, 28, MUTED, Bold=True)
            B.Text(S, "TotalValue", Big, Px, Top + 48, Pw, Bh, Bs, ACCENT, Bold=True, Heading=True)
            if Unit:
                B.Text(S, "TotalUnit", Unit, Px, Top + 48 + Bh, Pw, Uh, 36, ACCENT, Bold=True, Heading=True)
            if Note:
                B.Text(S, "CostNote", Note, Px, Top + 48 + Bh + Uh + 24, Pw, Nh, Ns, TEXT)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.CostTable")
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
                B.Text(S, "Caption", Sl["caption"], M, BODY_BOTTOM - 74, BoxW, 74, SIZE["caption"], MUTED)
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
            Hs = B.FitAll([str(X["heading"]) for X in Q], SIZE["h2"] - 2, Wd - 56, 54, Heading=True)
            Ts = B.FitAll([str(X.get("text", "")) for X in Q if X.get("text")], GROW["detail"], Wd - 56, Ht - 104)
            for I, Quad in enumerate(Q):
                Lx = Left + (I % 2) * (Wd + GAP)
                Ty = Top + (I // 2) * (Ht + GAP)
                B.Rect(S, f"Quadrant{I + 1}", Lx, Ty, Wd, Ht, ACCENT if I == Hi else QUIET)
                Fg = BG if I == Hi else TEXT
                B.Text(S, f"QHeading{I + 1}", Quad["heading"], Lx + 28, Ty + 24, Wd - 56, 54, Hs, Fg,
                       Bold=True, Heading=True)
                if Quad.get("text"):
                    B.Text(S, f"QText{I + 1}", Quad["text"], Lx + 28, Ty + 90, Wd - 56, Ht - 104, Ts, Fg)
            if Sl.get("x_axis"):
                B.Text(S, "XAxis", Sl["x_axis"] + " →", Left, BODY_BOTTOM - 40, W - Left - M, 40, SIZE["label"],
                       MUTED, Align=PP_ALIGN.CENTER)
            if Sl.get("y_axis"):
                Tb = B.Text(S, "YAxis", Sl["y_axis"] + " →", M - 200 + 30, Top + Ht, 400, 40, SIZE["label"],
                            MUTED, Align=PP_ALIGN.CENTER)
                Tb.rotation = -90
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Matrix")
            return

    def Email(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Calls = Sl.get("callouts", [])
            Ew = 900 if Calls else W - 2 * M
            Lx, Ty, Eh = M, BODY_TOP - 16, BODY_BOTTOM - BODY_TOP + 16
            Keys = [(C.get("target", "body"), C.get("line", 0) if C.get("target", "body") == "body" else 0)
                    for C in Calls]
            PerRow = max([Keys.count(K) for K in Keys] or [0])
            Pad, Dm = 28, 46
            Iw = Ew - 2 * Pad - (Dm + 52 * (PerRow - 1) + 16 if PerRow else 0)
            B.Rect(S, "Email", Lx, Ty, Ew, Eh, BG, Line=MUTED)
            B.Rect(S, "EmailBar", Lx, Ty, Ew, 44, QUIET)
            B.Text(S, "EmailBarText", "Inbox · Message", Lx + Pad, Ty + 8, 420, 36, SIZE["label"], MUTED,
                   Anchor=MSO_ANCHOR.MIDDLE)
            Y = Ty + 44 + 14
            RowY = {}
            for Key, Label in (("from", "From:"), ("to", "To:")):
                if Sl.get(Key):
                    Rh = max(42, B.Need(str(Sl[Key]), 27, Iw - 100))
                    B.Text(S, f"Email{Label[:-1]}Label", Label, Lx + Pad, Y, 96, 42, 27, MUTED, Bold=True)
                    B.Text(S, f"Email{Label[:-1]}", str(Sl[Key]), Lx + Pad + 100, Y, Iw - 100, Rh, 27, TEXT)
                    RowY[(Key, 0)] = Y + 21
                    Y += Rh + 4
            Sh = B.Need(str(Sl["subject"]), 28, Iw, Bold=True)
            B.Text(S, "EmailSubject", str(Sl["subject"]), Lx + Pad, Y + 4, Iw, Sh, 28, TEXT, Bold=True)
            RowY[("subject", 0)] = Y + 4 + 22
            Y += 4 + Sh + 8
            Att = Sl.get("attachment")
            if Att:  # where mail programs show it: under the subject
                Cw = kMeasure.TextWidth(f"Attachment: {Att}", B.Minor, 26 * B.TypeScale()) / B._kx + 48
                Chip = B.Rect(S, "EmailAttachment", Lx + Pad, Y, min(Iw, Cw), 48, QUIET, Line=MUTED)
                B.ShapeText(Chip, f"Attachment: {Att}", 26, TEXT, Bold=False, Heading=False, Inset=14)
                RowY[("attachment", 0)] = Y + 24
                Y += 48 + 12
            B.Rect(S, "EmailRule", Lx + Pad, Y, Ew - 2 * Pad, 2, QUIET)
            Y += 16
            Bottom = Ty + Eh - 20
            Body = [str(X) for X in Sl["body"]]
            Size = 30
            while Size > FLOOR and sum(B.Need(X, Size, Iw) for X in Body) + 10 * (len(Body) - 1) > Bottom - Y:
                Size -= 1
            if sum(B.Need(X, Size, Iw) for X in Body) + 10 * (len(Body) - 1) > (Bottom - Y) * 1.05:
                B.Problems.append(f"slide {B.SlideNo} ({B.SlideId}): the email body does not fit even at {Size} pt; "
                                  "shorten its paragraphs or drop one")
            for I, Para in enumerate(Body):
                Ph = B.Need(Para, Size, Iw)
                Tb = B.Text(S, f"EmailBody{I + 1}", Para, Lx + Pad, Y, Iw, Ph, Size, TEXT)
                if Para.lower().startswith(("http://", "https://", "www.")):  # a link looks like one
                    for Run in Tb.text_frame.paragraphs[0].runs:
                        Run.font.underline = True
                        Run.font.color.theme_color = ACCENT
                RowY[("body", I)] = Y + Size * 0.62
                Y += Ph + 10
            Used = {}
            for I, (Key, Line) in enumerate(Keys):
                Cy = RowY.get((Key, Line), RowY.get(("subject", 0)))
                K = Used.get((Key, Line), 0)
                Used[(Key, Line)] = K + 1
                self.Badge(S, f"Marker{I + 1}", str(I + 1), Lx + Ew - Pad - Dm - 52 * K, Cy - Dm / 2, Dm, ACCENT, BG, 24)
            if Calls:
                Px = Lx + Ew + 44
                Pw = W - M - Px
                Tw = Pw - 72
                Notes = [str(C.get("note", "")) for C in Calls]
                Ns = B.FitAll(Notes, 34, Tw, (Eh - 24 * (len(Notes) - 1)) / len(Notes))
                Hs = [max(B.Need(X, Ns, Tw), 54) for X in Notes]
                Gap = min(56, max(16, (Eh - sum(Hs)) / max(1, len(Hs) - 1)))
                Y = Ty + max(0, (Eh - sum(Hs) - Gap * (len(Hs) - 1)) / 2)
                for I, Note in enumerate(Notes):
                    self.Badge(S, f"CalloutNo{I + 1}", str(I + 1), Px, Y, 54, ACCENT, BG, 26)
                    B.Text(S, f"Callout{I + 1}", Note, Px + 72, Y + 2, Tw, Hs[I], Ns, TEXT)
                    Y += Hs[I] + Gap
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Email")
            return

    def Options(self, S, Sl, Answer):
        """The quiz options as lettered cards; on the answer slide the right ones in the accent."""
        if kS.ErrorMode:
            return
        try:
            B = self._b
            Opts = Sl["options"]
            N = len(Opts)
            Explain = Sl.get("explain") if Answer else ""
            Bottom = BODY_BOTTOM - (140 if Explain else 0)
            Grid = N == 4
            Cols = 2 if Grid else 1
            RowsN = 2 if Grid else N
            Wd = (W - 2 * M - GAP * (Cols - 1)) / Cols
            Ht = (Bottom - BODY_TOP - 10 - GAP * (RowsN - 1)) / RowsN
            Pad, D = 28, 64
            Tw = Wd - 2 * Pad - D - 24 - (150 if Answer else 0)
            Ts = B.FitAll([str(O["text"]) for O in Opts], 40, Tw, Ht - 2 * 16)
            for I, O in enumerate(Opts):
                Lx = M + (I % Cols) * (Wd + GAP)
                Ty = BODY_TOP + 10 + (I // Cols) * (Ht + GAP)
                Right = Answer and O.get("correct")
                B.Rect(S, f"Option{I + 1}", Lx, Ty, Wd, Ht, ACCENT if Right else QUIET)
                self.Badge(S, f"Letter{I + 1}", "ABCD"[I], Lx + Pad, Ty + (Ht - D) / 2, D, BG if Right else (
                    MUTED if Answer else ACCENT), ACCENT if Right else BG, 32)
                B.Text(S, f"OptionText{I + 1}", str(O["text"]), Lx + Pad + D + 24, Ty + 12, Tw, Ht - 24, Ts,
                       BG if Right else (MUTED if Answer else TEXT), Anchor=MSO_ANCHOR.MIDDLE)
                if Answer:
                    B.Text(S, f"Verdict{I + 1}", "Correct" if Right else "Not this one", Lx + Wd - Pad - 150, Ty + 12,
                           150, Ht - 24, 24, BG if Right else MUTED, Bold=bool(Right), Align=PP_ALIGN.RIGHT,
                           Anchor=MSO_ANCHOR.MIDDLE)
            if Explain:
                B.Text(S, "Explain", Explain, M, Bottom + 24, W - 2 * M, 116, 34, TEXT)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Options")
            return

    def Quiz(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            self._b.Title(S, Sl["title"])
            self.Options(S, Sl, False)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Quiz")
            return

    def QuizAnswer(self, S, Sl):
        """The slide after a quiz with reveal 'slide': the same options, the right ones marked, the why below."""
        if kS.ErrorMode:
            return
        try:
            Letters = " and ".join("ABCD"[I] for I, O in enumerate(Sl["options"]) if O.get("correct"))
            Default = f"Answer: {Letters}"
            self._b.Title(S, Sl.get("answer_title") or Default)
            self.Options(S, Sl, True)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.QuizAnswer")
            return

    def Risks(self, S, Sl):
        if kS.ErrorMode:
            return
        try:
            B = self._b
            B.Title(S, Sl["title"])
            Rk = Sl["risks"]
            Hi = Sl.get("highlight")
            N = len(Rk)
            Wd = (W - 2 * M - GAP * (N - 1)) / N
            Ty, Ht = BODY_TOP + 10, BODY_BOTTOM - BODY_TOP - 10
            Pad = 28
            Iw = Wd - 2 * Pad
            Hs = B.FitAll([str(R["risk"]) for R in Rk], 40, Iw, 150, Heading=True)
            Hh = max(B.Need(str(R["risk"]), Hs, Iw, Heading=True) for R in Rk)
            ChipsTop = Ty + 14 + Pad + Hh + 16
            MitTop = ChipsTop + 2 * 46 + 22
            Room = Ty + Ht - Pad - MitTop - 40
            Ms = B.FitAll([str(R.get("mitigation", "")) for R in Rk], GROW["detail"], Iw, Room)
            Fills = {"HIGH": (ACCENT, BG), "MEDIUM": (TEXT, BG), "LOW": (BG, TEXT)}
            for I, R in enumerate(Rk):
                Lx = M + I * (Wd + GAP)
                B.Rect(S, f"Risk{I + 1}", Lx, Ty, Wd, Ht, QUIET)
                B.Rect(S, f"RiskRule{I + 1}", Lx, Ty, Wd, 14, ACCENT if I == Hi else MUTED)
                B.Text(S, f"RiskName{I + 1}", str(R["risk"]), Lx + Pad, Ty + 14 + Pad, Iw, Hh, Hs, TEXT, Bold=True,
                       Heading=True)
                for J, (Key, Label) in enumerate((("likelihood", "Likelihood"), ("impact", "Impact"))):
                    Level = LEVELS.get(str(R.get(Key, "medium")).lower(), "MEDIUM")
                    Cy = ChipsTop + J * 46
                    B.Text(S, f"{Label}{I + 1}", Label, Lx + Pad, Cy, Iw - 132, 38, 26, MUTED,
                           Anchor=MSO_ANCHOR.MIDDLE)
                    Chip = B.Rect(S, f"{Label}Chip{I + 1}", Lx + Pad + Iw - 124, Cy, 124, 38, Fills[Level][0],
                                  Line=TEXT if Level == "LOW" else None)
                    B.ShapeText(Chip, Level, 24, Fills[Level][1])
                B.Text(S, f"MitigationLabel{I + 1}", "Mitigation", Lx + Pad, MitTop, Iw, 36, SIZE["label"], MUTED,
                       Bold=True)
                if R.get("mitigation"):
                    B.Text(S, f"Mitigation{I + 1}", str(R["mitigation"]), Lx + Pad, MitTop + 40, Iw, Room, Ms, TEXT)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlidePatterns.Risks")
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
            for Nn in Builder.NoNotes:
                print(f"notes: {Nn} has no speaker notes - write what the speaker says (reference/CONTENT.md)",
                      file=sys.stderr)
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
