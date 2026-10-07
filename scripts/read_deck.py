"""Read a whole deck into JSON in one go - any OS, no PowerPoint (python-pptx).

    uvx --with python-pptx python scripts/read_deck.py deck.pptx                 # all slides, compact JSON
    uvx --with python-pptx python scripts/read_deck.py deck.pptx --slides 5-20 --pretty --out dump.json
    uvx --with python-pptx python scripts/read_deck.py deck.pptx --text          # just titles, text and notes

Use it before judging or editing a deck: one call gives every slide's id, layout, title, shapes (kind,
name, position in pt, text, font sizes, alt text), chart data, table cells and speaker notes - far faster
than one screenshot round-trip per slide, and enough for cross-slide reasoning (repeated claims,
contradictions, a relationship map, word counts). Positions are slide coordinates, groups included.
"""
import argparse
import json
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

from lint_deck import GROUP_TF, alt_text, is_title, para_size, rect, walk

KIND = {MSO_SHAPE_TYPE.PICTURE: "picture", MSO_SHAPE_TYPE.CHART: "chart", MSO_SHAPE_TYPE.TABLE: "table",
        MSO_SHAPE_TYPE.GROUP: "group", MSO_SHAPE_TYPE.TEXT_BOX: "text", MSO_SHAPE_TYPE.PLACEHOLDER: "placeholder",
        MSO_SHAPE_TYPE.AUTO_SHAPE: "shape", MSO_SHAPE_TYPE.MEDIA: "media", MSO_SHAPE_TYPE.LINE: "line"}


def shape_info(sh):
    l, t, w, h = rect(sh)
    info = {"name": sh.name, "kind": KIND.get(sh.shape_type, str(sh.shape_type).split(".")[-1].lower()),
            "box_pt": [round(l), round(t), round(w), round(h)]}
    if is_title(sh):
        info["title"] = True
    if sh.has_text_frame and sh.text_frame.text.strip():
        info["text"] = sh.text_frame.text
        info["sizes_pt"] = sorted({round(s, 1) for p in sh.text_frame.paragraphs if p.text.strip()
                                   for s in [para_size(sh, p)] if s})
    if sh.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART) or getattr(sh, "has_chart", False):
        info["alt"] = alt_text(sh)
    if getattr(sh, "has_chart", False) and sh.has_chart:
        ch = sh.chart
        plot = ch.plots[0] if len(ch.plots) else None
        info["chart"] = {"type": str(ch.chart_type).split(".")[-1].split(" ")[0].lower(),
                         "categories": [str(c) for c in plot.categories] if plot is not None else [],
                         "series": [{"name": s.name, "values": list(s.values)} for s in plot.series] if plot else []}
    if getattr(sh, "has_table", False) and sh.has_table:
        info["table"] = [[c.text for c in row.cells] for row in sh.table.rows]
    return info


def read(path, lo=None, hi=None):
    prs = Presentation(path)
    out = {"file": path, "slide_size_pt": [Emu(prs.slide_width).pt, Emu(prs.slide_height).pt], "slides": []}
    for n, s in enumerate(prs.slides, 1):
        if (lo and n < lo) or (hi and n > hi):
            continue
        GROUP_TF.clear()
        shapes = [shape_info(sh) for sh in walk(s.shapes)]
        title = next((x.get("text", "") for x in shapes if x.get("title")), "")
        out["slides"].append({
            "index": n, "id": s._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}cSld").get("name"),
            "hidden": s._element.get("show") == "0", "layout": s.slide_layout.name, "title": title,
            "words": sum(len(x.get("text", "").split()) for x in shapes if not x.get("title")),
            "shapes": shapes,
            "notes": s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""})
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--slides", help="range like 5-20 (1-based, inclusive)")
    ap.add_argument("--text", action="store_true", help="plain text outline instead of JSON")
    ap.add_argument("--pretty", action="store_true")
    ap.add_argument("--out")
    a = ap.parse_args()
    lo, hi = (map(int, a.slides.split("-")) if a.slides else (None, None))
    data = read(a.file, lo, hi)
    if a.text:
        lines = []
        for s in data["slides"]:
            lines.append(f"## {s['index']}. {s['title'] or '(no title)'}" + ("  [hidden]" if s["hidden"] else ""))
            lines += [x["text"] for x in s["shapes"] if x.get("text") and not x.get("title")]
            if s["notes"]:
                lines.append("NOTES: " + s["notes"])
            lines.append("")
        text = "\n".join(lines)
    else:
        text = json.dumps(data, ensure_ascii=False, indent=2 if a.pretty else None, default=str)
    if a.out:
        open(a.out, "w", encoding="utf-8").write(text)
        print(a.out)
    else:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
        print(text)


if __name__ == "__main__":
    main()
