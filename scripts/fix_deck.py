"""Repair the mechanical defects lint_deck.py finds - any OS, python-pptx. Writes a copy; never the input.

    uvx --with python-pptx --with pillow python scripts/fix_deck.py deck.pptx --out fixed.pptx
    uvx --with python-pptx --with pillow python scripts/fix_deck.py deck.pptx --out fixed.pptx --only floor,alt
    uvx --with python-pptx --with pillow python scripts/fix_deck.py deck.pptx --dry-run

Safe fixes (each one listed as it is made):
  placeholder  delete empty title/body placeholders sitting next to real content
  floor        raise sentence text below the body floor to the floor (27 pt on Full HD)
  fit          shrink text that overflows a fixed-size box, never below the floor
  aspect       un-stretch pictures by shrinking them to their true proportions inside the same box
               (nothing is cut off; use --crop-photos to crop photos to fill the box instead)
  alt          write alt text for charts and tables from their data
  palette      replace Office default chart colours with shades of the theme accent
  legend       move a top/bottom chart legend to the right
  numfmt       replace an Accounting axis format ("$-" for zero) with a plain number format
Then lint_deck.py runs on the result and prints what is left: titles, wording, pictures' alt text, colour
choices - the things that need a person (or Claude looking at the render).
"""
import argparse
import os
import subprocess
import sys

from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu, Pt

import lint_deck as L
from _measure import text_height
from _rules import OFFICE_DEFAULT_SERIES

HERE = os.path.dirname(os.path.abspath(__file__))
FIXES = ["placeholder", "floor", "fit", "aspect", "alt", "palette", "legend", "numfmt"]
RAMP = [0, 0.35, -0.3, 0.6, -0.5, 0.75]


def chart_alt(ch):
    kind = str(ch.chart_type).split(".")[-1].split(" ")[0].replace("_", " ").lower()
    plot = ch.plots[0]
    cats = [str(c) for c in plot.categories]
    parts = []
    for s in plot.series:
        vals = ", ".join(f"{c} {v:g}" if isinstance(v, (int, float)) else f"{c} n/a" for c, v in zip(cats, s.values))
        parts.append(f"{s.name}: {vals}" if len(plot.series) > 1 or s.name else vals)
    title = ch.chart_title.text_frame.text.strip() + ". " if ch.has_title and ch.chart_title.has_text_frame else ""
    return f"{kind.capitalize()} chart. {title}" + "; ".join(parts) + "."


def fit_size(sh, theme, size, floor, family, bold, width, height):
    cur = size
    while cur > floor:
        paras = [(p.text, family, cur, bold, 0) for p in sh.text_frame.paragraphs]
        need, widest, _ = text_height(paras, width)
        if need <= height and widest <= width:
            return cur
        cur = max(floor, cur - 1)
    return cur


def fix(prs, only, log, crop_photos=False):
    theme = L.Theme(prs.slide_master)
    sw = Emu(prs.slide_width).pt
    scale = max(1.0, sw / 960)
    floor, label_max = 18 * scale, 12 * scale
    for n, slide in enumerate(prs.slides, 1):
        if slide._element.get("show") == "0":
            continue
        L.GROUP_TF.clear()
        L.GROWN.clear()
        shapes = list(L.walk(slide.shapes))
        content = [s for s in shapes if s.shape_type != MSO_SHAPE_TYPE.GROUP]
        real = [s for s in content if not L.is_title(s) and (
            s.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART, MSO_SHAPE_TYPE.TABLE)
            or getattr(s, "has_chart", False) or (s.has_text_frame and s.text_frame.text.strip()))]
        for s in content:
            where = f"slide {n} [{s.name}]"
            # placeholder
            if "placeholder" in only and s.is_placeholder and s.has_text_frame and not s.text_frame.text.strip() \
                    and L.ph_type(s) in L.TEXT_PLACEHOLDERS and not L.is_title(s) and any(o is not s for o in real):
                s._element.getparent().remove(s._element)
                log(f"placeholder  {where}: deleted the empty placeholder")
                continue
            # floor
            if "floor" in only and s.has_text_frame:
                raised = 0
                for p in s.text_frame.paragraphs:
                    size = L.para_size(s, p)
                    if size and label_max < size < floor and len(p.text.split()) >= L.BODY_MIN_WORDS:
                        for r in p.runs:
                            r.font.size = Pt(floor)
                        raised += 1
                if raised:
                    log(f"floor        {where}: {raised} paragraph(s) raised to {floor:g} pt")
            # fit
            if "fit" in only and s.has_text_frame and s.text_frame.text.strip():
                bp = s.text_frame._txBody.find("a:bodyPr", L.NS)
                fixed_box = bp is None or (bp.find("a:spAutoFit", L.NS) is None and bp.find("a:normAutofit", L.NS) is None
                                           and bp.get("wrap") != "none")
                if fixed_box:
                    ins = lambda k, d: (int(bp.get(k)) / L.PT if bp is not None and bp.get(k) else d)
                    l, t, w, h = L.rect(s)
                    width, height = w - ins("lIns", 7.2) - ins("rIns", 7.2), h - ins("tIns", 3.6) - ins("bIns", 3.6)
                    sizes = [L.para_size(s, p) or 18.0 for p in s.text_frame.paragraphs if p.text.strip()]
                    run = next((r for p in s.text_frame.paragraphs for r in p.runs if r.text.strip()), None)
                    name = run.font.name if run is not None else None
                    family = theme.font(name) if name else (theme.major if L.is_title(s) else theme.minor)
                    bold = bool(run.font.bold) if run is not None and run.font.bold is not None else L.is_title(s)
                    size = max(sizes) if sizes else 18.0
                    need, widest, _ = text_height([(p.text, family, L.para_size(s, p) or 18.0, bold, 0)
                                                   for p in s.text_frame.paragraphs], width)
                    if need > height * 1.08 and width > 0:
                        new = fit_size(s, theme, size, floor if any(len(p.text.split()) >= 4 for p in
                                                                    s.text_frame.paragraphs) else label_max,
                                       family, bold, width, height)
                        if new < size:
                            for p in s.text_frame.paragraphs:
                                for r in p.runs:
                                    r.font.size = Pt(min(new, (r.font.size.pt if r.font.size else size)))
                            log(f"fit          {where}: text {size:g} -> {new:g} pt to fit its box"
                                + (" (still too long - cut words)" if new <= floor else ""))
            # crop
            if "aspect" in only and s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                stretch = L.picture_stretch(s)
                if abs(stretch) > L.STRETCH_TOLERANCE:
                    from PIL import Image
                    import io
                    iw, ih = Image.open(io.BytesIO(s.image.blob)).size
                    box = s.width / s.height
                    vis_w = iw * (1 - s.crop_left - s.crop_right)
                    vis_h = ih * (1 - s.crop_top - s.crop_bottom)
                    ratio = vis_w / vis_h
                    if crop_photos:  # fill the box by cropping - right for photos, wrong for charts/diagrams
                        if ratio > box:
                            extra = (vis_w - vis_h * box) / iw / 2
                            s.crop_left, s.crop_right = s.crop_left + extra, s.crop_right + extra
                        else:
                            extra = (vis_h - vis_w / box) / ih / 2
                            s.crop_top, s.crop_bottom = s.crop_top + extra, s.crop_bottom + extra
                        log(f"aspect       {where}: was {stretch:+.0%} stretched; cropped to fill the box "
                            "(check the subject is still in frame)")
                    else:            # fit inside the old box at the true ratio, centred: nothing is lost
                        if ratio < box:
                            new_w = int(s.height * ratio)
                            s.left, s.width = s.left + (s.width - new_w) // 2, new_w
                        else:
                            new_h = int(s.width / ratio)
                            s.top, s.height = s.top + (s.height - new_h) // 2, new_h
                        log(f"aspect       {where}: was {stretch:+.0%} stretched; resized to its true proportions "
                            "inside the same box")
            # alt
            if "alt" in only and not L.alt_text(s) and not L.is_decorative(s):
                nv = s._element.find(".//p:cNvPr", L.NS)
                if getattr(s, "has_chart", False) and s.has_chart:
                    nv.set("descr", chart_alt(s.chart))
                    log(f"alt          {where}: chart alt text written from its data")
                elif getattr(s, "has_table", False) and s.has_table:
                    head = [c.text for c in s.table.rows[0].cells]
                    nv.set("descr", f"Table: {', '.join(head)}; {len(s.table.rows) - 1} rows.")
                    log(f"alt          {where}: table alt text written from its header")
            # chart fixes
            if getattr(s, "has_chart", False) and s.has_chart:
                ch = s.chart
                cx = ch._chartSpace
                C = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart", "a": L.NS["a"]}
                if "palette" in only:
                    sers = list(ch.plots[0].series) if len(ch.plots) else []
                    cols = [theme.colour_in(x._element.find("c:spPr/a:solidFill", C)) for x in sers[:6]]
                    if sum(1 for c in cols if c in OFFICE_DEFAULT_SERIES) >= 2:
                        for i, ser in enumerate(sers):
                            ser.format.fill.solid()
                            ser.format.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
                            ser.format.fill.fore_color.brightness = RAMP[i % len(RAMP)]
                        log(f"palette      {where}: {len(sers)} series recoloured as shades of the theme accent")
                if "legend" in only and ch.has_legend:
                    pos = cx.find(".//c:legend/c:legendPos", C)
                    if pos is not None and pos.get("val") in ("t", "b"):
                        ch.legend.position = XL_LEGEND_POSITION.RIGHT
                        ch.legend.include_in_layout = False
                        log(f"legend       {where}: legend moved to the right")
                if "numfmt" in only:
                    for fmt in cx.findall(".//c:valAx/c:numFmt", C):
                        code = fmt.get("formatCode", "")
                        if "_(" in code or ('"-"' in code and "*" in code):
                            fmt.set("formatCode", "$#,##0" if "$" in code else "#,##0")
                            fmt.set("sourceLinked", "0")
                            log(f"numfmt       {where}: Accounting axis format replaced")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--out")
    ap.add_argument("--only", help=f"comma-separated subset of: {', '.join(FIXES)}")
    ap.add_argument("--dry-run", action="store_true", help="list the fixes without writing")
    ap.add_argument("--crop-photos", action="store_true",
                    help="crop stretched pictures to fill their box instead of shrinking them (photos only)")
    a = ap.parse_args()
    if not a.out and not a.dry_run:
        ap.error("--out is required (the input is never overwritten), or use --dry-run")
    if a.out and os.path.abspath(a.out) == os.path.abspath(a.file):
        ap.error("--out must differ from the input")
    only = set(a.only.split(",")) if a.only else set(FIXES)
    unknown = only - set(FIXES)
    if unknown:
        ap.error(f"unknown fix(es): {', '.join(sorted(unknown))}")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    prs = Presentation(a.file)
    made = []
    fix(prs, only, lambda m: (made.append(m), print(m)), a.crop_photos)
    print(f"\n{len(made)} fix(es){' (dry run, nothing written)' if a.dry_run else ''}")
    if a.dry_run or not made and not a.out:
        return
    prs.save(a.out)
    print(f"-> {a.out}\n\nWhat's left (lint_deck.py):")
    r = subprocess.run([sys.executable, os.path.join(HERE, "lint_deck.py"), a.out])
    sys.exit(r.returncode)


if __name__ == "__main__":
    main()
