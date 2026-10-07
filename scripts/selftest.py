"""Self-test for every script in this folder, against a deck it builds itself.

    # any OS: the scripts that need no PowerPoint
    uvx --with python-pptx --with pillow python scripts/selftest.py
    # Windows with PowerPoint installed: everything
    uvx --with python-pptx --with pillow --with pywin32 python scripts/selftest.py --com

The test deck (5 slides, 1440 x 810 pt) has known defects, so each check knows what to expect:
  1 "Decisions"       topic-label title, a 4-word body paragraph at 22 pt   -> ** AUDIT, label?
  2 claim title       short body at 24 pt                                    -> OK
  3 claim title       one long word in a narrow box at 60 pt                 -> broken word
  4 claim title       speaker notes                                          -> notes in bulk_read
  5 hidden slide                                                             -> [skip]
Exit code 0 = every check passed.
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image
from pptx import Presentation
from pptx.util import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
results = []
report = []


def say(line=""):
    print(line)
    report.append(line)


def check(name, ok, detail=""):
    results.append(ok)
    say(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


def run(script, *args):
    p = subprocess.run([sys.executable, os.path.join(HERE, script), *args],
                       capture_output=True, text=True, encoding="utf-8")
    return p.returncode, p.stdout + p.stderr


def build_deck(path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(1440), Pt(810)
    layout = prs.slide_layouts[5]  # Title Only

    def slide(title):
        s = prs.slides.add_slide(layout)
        s.shapes.title.text = title
        return s

    def text(s, txt, size, left=100, top=250, width=1200, height=200, name="Body"):
        tb = s.shapes.add_textbox(Pt(left), Pt(top), Pt(width), Pt(height))
        tb.name = name
        tb.text_frame.word_wrap = True
        tb.text_frame.text = txt
        for r in tb.text_frame.paragraphs[0].runs:
            r.font.size = Pt(size)

    text(slide("Decisions"), "Small text below the floor", 22)
    text(slide("The check comes first"), "Plan, check, run", 24)
    text(slide("Long words break lines"), "Internationalization", 60, width=200, name="Narrow")
    s = slide("Notes carry the depth")
    text(s, "Short body", 24)
    s.notes_slide.notes_text_frame.text = "Key fact: notes are read by bulk_read."
    hidden = slide("Hidden slide")
    hidden._element.set("show", "0")
    prs.save(path)


def build_taste_deck(path):
    """One slide per group of taste defects, each built deliberately."""
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.enum.dml import MSO_THEME_COLOR
    from pptx.enum.text import PP_ALIGN

    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(1440), Pt(810)
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Taste defects are easy to spot"

    def box(txt, x, y, w=300, h=80, size=20):
        tb = s.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
        tb.text_frame.word_wrap = True
        tb.text_frame.text = txt
        for r in tb.text_frame.paragraphs[0].runs:
            r.font.size = Pt(size)
        return tb

    box("\U0001F4CA Charts", 40, 140)
    box("Lorem ipsum dolor sit amet, consectetur adipiscing elit.", 40, 240)
    long = box("This is a long body sentence that keeps going and going so that it is clearly more than eighty characters.",
               380, 140, 600, 120)
    long.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    pale = box("Pale grey text that is hard to read on white", 380, 300, 600, 60)
    pale.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
    box("Read the full report in the appendix for the remaining items and...", 380, 380, 600, 60)
    for i, slot in enumerate([MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_2, MSO_THEME_COLOR.ACCENT_3,
                              MSO_THEME_COLOR.ACCENT_4]):
        r = s.shapes.add_shape(1, Pt(1040 + i * 90), Pt(140), Pt(80), Pt(80))
        r.fill.solid()
        r.fill.fore_color.theme_color = slot
        r.shadow.inherit = False
        sp = r._element.spPr
        from lxml import etree
        eff = etree.SubElement(sp, "{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst")
        etree.SubElement(eff, "{http://schemas.openxmlformats.org/drawingml/2006/main}outerShdw", blurRad="50800")
    for i, hexv in enumerate(["123456", "654321", "ABCDEF"]):
        r = s.shapes.add_shape(1, Pt(1040 + i * 90), Pt(260), Pt(80), Pt(80))
        r.fill.solid()
        r.fill.fore_color.rgb = RGBColor.from_string(hexv)
    s.notes_slide.notes_text_frame.text = "notes"

    c = prs.slides.add_slide(prs.slide_layouts[5])
    c.shapes.title.text = "Revenue rises every quarter"
    data = CategoryChartData()
    data.categories = ["East", "West"]
    for q, v in (("Q1", (1, 2)), ("Q2", (2, 3)), ("Q3", (3, 4))):
        data.add_series(q, v)
    ch = c.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Pt(100), Pt(150), Pt(800), Pt(500), data).chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    for ser, hexv in zip(ch.plots[0].series, ["4472C4", "ED7D31", "A5A5A5"]):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = RGBColor.from_string(hexv)
    ch.value_axis.tick_labels.number_format = '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'
    ch.value_axis.tick_labels.number_format_is_linked = False
    prs.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--com", action="store_true", help="also run the PowerPoint (COM) scripts; Windows only")
    a = ap.parse_args()

    tmp = tempfile.mkdtemp(prefix="pptskill-selftest-")
    deck = os.path.join(tmp, "selftest.pptx")
    build_deck(deck)
    say(f"test deck: {deck}\n")

    # --- any OS
    code, out = run("check_slide_size.py", deck)
    check("check_slide_size: Full HD deck passes", code == 0, out)
    small = os.path.join(tmp, "small.pptx")
    Presentation().save(small)
    code, out = run("check_slide_size.py", small)
    check("check_slide_size: default 4:3 deck fails", code == 1, out)

    photo = os.path.join(tmp, "photo.jpg")
    Image.new("RGB", (1500, 1000), "gray").save(photo)
    crop = os.path.join(tmp, "photo.crop.jpg")
    code, out = run("cover_crop.py", photo, "--box", "1440x610", "--out", crop)
    w, h = Image.open(crop).size if os.path.exists(crop) else (0, 1)
    check("cover_crop: reports +57% stretch", "+57%" in out, out)
    check("cover_crop: output has the box ratio", abs(w / h - 1440 / 610) < 0.01, f"{w}x{h}")

    code, out = run("backup_snapshot.py", "--file", deck, "--label", "selftest")
    check("backup_snapshot: copy written", code == 0 and os.path.exists(out.strip()), out)

    sys.path.insert(0, HERE)
    from _rules import looks_like_label as label
    check("audit_deck: 'Decisions' is a label", label("Decisions"))
    check("audit_deck: 'The check comes first' is a claim", not label("The check comes first"))
    check("audit_deck: a question is not flagged", not label("Why Claude?"))

    # lint_deck: cross-platform checks against the same deck
    code, out = run("lint_deck.py", deck, "--json")
    try:
        found = {(i["slide"], i["code"]) for i in json.loads(out.split("\nfixed")[0])["findings"]}
        check("lint_deck: slide 1 body below the floor (27 pt at Full HD)", (1, "body_below_floor") in found)
        check("lint_deck: slide 1 title is a label", (1, "title_is_label") in found)
        check("lint_deck: slide 2 has no body/title findings",
              not {c for sl, c in found if sl == 2} & {"body_below_floor", "title_is_label", "empty_title"})
        check("lint_deck: hidden slide 5 skipped", not any(sl == 5 for sl, _ in found))
    except Exception as e:
        check("lint_deck: valid JSON", False, f"{e}: {out[:300]}")

    messy = os.path.join(tmp, "messy.pptx")
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(1440), Pt(810)
    s = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content; body left empty
    s.shapes.title.text = "Pictures need crops"
    s.shapes.add_picture(photo, Pt(100), Pt(300), Pt(600), Pt(200))  # 3:2 photo forced to 3:1
    prs.save(messy)
    code, out = run("lint_deck.py", messy)
    check("lint_deck: finds the empty placeholder", "unused_placeholder" in out, out)
    check("lint_deck: finds the stretched picture", "picture_stretched" in out, out)
    check("lint_deck: treats a file name as missing alt text", "a11y_missing_alt_text" in out, out)
    check("lint_deck: exit 1 on errors", code == 1, out)
    fixed = os.path.join(tmp, "messy.fixed.pptx")
    code, out = run("lint_deck.py", messy, "--fix", "--out", fixed)
    code2, out2 = run("lint_deck.py", fixed)
    check("lint_deck --fix: removes the empty placeholder", os.path.exists(fixed) and "unused_placeholder" not in out2, out2)

    # lint_deck: taste, contrast and chart checks
    taste = os.path.join(tmp, "taste.pptx")
    build_taste_deck(taste)
    code, out = run("lint_deck.py", taste, "--json")
    try:
        got = {i["code"] for i in json.loads(out)["findings"]}
        for want in ("emoji_as_icon", "lorem_ipsum", "centered_long_body", "shadow_overuse", "a11y_low_text_contrast",
                     "off_palette_fill", "accent_overload", "chart_legend_steals_plot", "chart_default_palette",
                     "chart_ordinal_categorical_color", "chart_accounting_zero_dash", "truncated_text"):
            check(f"lint_deck: finds {want}", want in got, ", ".join(sorted(got)))
    except Exception as e:
        check("lint_deck: taste deck JSON", False, f"{e}: {out[:300]}")

    # build_deck: the sample spec must build and lint with no errors or warnings
    sample = os.path.join(HERE, "..", "examples", "spec", "sample-deck.json")
    built = os.path.join(tmp, "built.pptx")
    code, out = run("build_deck.py", sample, "--out", built)
    check("build_deck: sample spec builds", code == 0 and os.path.exists(built), out)
    code, out = run("lint_deck.py", built, "--json")
    try:
        counts = json.loads(out)["counts"]
        check("build_deck: output lints clean (no errors or warnings)",
              not counts.get("error") and not counts.get("warn"), out[-600:])
    except Exception as e:
        check("build_deck: lint JSON", False, f"{e}: {out[:300]}")
    bad = os.path.join(tmp, "bad-spec.json")
    json.dump({"slides": [{"pattern": "kpi", "title": "Too few", "metrics": [{"value": "1", "label": "x"}]},
                          {"pattern": "bullets", "items": ["no title"]}]}, open(bad, "w"))
    code, out = run("build_deck.py", bad, "--out", os.path.join(tmp, "bad.pptx"))
    check("build_deck: rejects a spec that breaks pattern limits", code == 2 and "needs 3-6" in out
          and "needs a 'title'" in out, out)
    sys.path.insert(0, HERE)
    from _rules import contrast_ratio
    dirs = json.load(open(os.path.join(HERE, "directions.json"), encoding="utf-8"))["directions"]
    weak = [d["id"] for d in dirs if contrast_ratio(d["text"], d["background"]) < 7
            or contrast_ratio(d["muted"], d["background"]) < 4.5 or contrast_ratio(d["accent"], d["background"]) < 3]
    check(f"directions: all {len(dirs)} pass contrast (text 7:1, muted 4.5:1, accent 3:1)", not weak, ", ".join(weak))

    # extract_theme
    code, out = run("extract_theme.py", deck)
    try:
        t = json.loads(out)
        check("extract_theme: fonts and accent colours read", bool(t["fonts"]["major_latin"]) and len(t["colours"]) == 12)
        check("extract_theme: layouts listed", len(t["layouts"]) >= 6)
    except Exception as e:
        check("extract_theme: valid JSON", False, f"{e}: {out[:300]}")

    # diff_renders on synthetic images
    da, db = os.path.join(tmp, "ra"), os.path.join(tmp, "rb")
    for d in (da, db):
        os.makedirs(d, exist_ok=True)
        Image.new("RGB", (320, 180), "white").save(os.path.join(d, "s001.png"))
    Image.new("RGB", (320, 180), "white").save(os.path.join(da, "s002.png"))
    Image.new("RGB", (320, 180), "black").save(os.path.join(db, "s002.png"))
    code, out = run("diff_renders.py", da, db)
    check("diff_renders: unchanged slide reported same", "s001.png   same" in out, out)
    check("diff_renders: changed slide reported", "s002.png   CHANGED" in out and code == 1, out)

    # render_lo (only where LibreOffice is installed)
    if shutil.which("soffice") and shutil.which("pdftoppm"):
        code, out = run("render_lo.py", deck, "--out", os.path.join(tmp, "lo"))
        n = len([x for x in out.split() if x.endswith(".png")])
        check("render_lo: one PNG per visible slide", code == 0 and n >= 4, out)
    else:
        say("(skipped render_lo: LibreOffice/pdftoppm not installed)")

    # --- Windows + PowerPoint
    if a.com:
        code, out = run("bulk_read.py", "--file", deck)
        try:
            data = json.loads(out)
            check("bulk_read: 5 slides", len(data["slides"]) == 5)
            check("bulk_read: notes read", "Key fact" in data["slides"][3]["notes"])
            check("bulk_read: hidden slide marked", data["slides"][4]["hidden"])
        except Exception as e:
            check("bulk_read: valid JSON", False, f"{e}: {out[:300]}")

        code, out = run("audit_deck.py", "--file", deck)
        rows = {ln.split()[0]: ln for ln in out.splitlines() if ln[:3].strip().isdigit()}
        check("audit_deck: slide 1 flagged", "** AUDIT" in rows.get("1", ""), out)
        check("audit_deck: slide 1 min font 22pt", "22pt" in rows.get("1", ""), out)
        check("audit_deck: slide 1 label? warning", "label?" in rows.get("1", ""), out)
        check("audit_deck: slide 2 OK", " OK " in rows.get("2", "") + " ", out)
        check("audit_deck: hidden slide skipped", "[skip]" in rows.get("5", ""), out)

        code, out = run("check_word_breaks.py", "--file", deck)
        check("check_word_breaks: finds the broken word on slide 3", code == 1 and "slide 3" in out, out)

        renders = os.path.join(tmp, "renders")
        code, out = run("render_slides.py", "--file", deck, "--out", renders)
        n = len([f for f in os.listdir(renders) if f.endswith(".jpg")]) if os.path.isdir(renders) else 0
        check("render_slides: one JPEG per slide", code == 0 and n == 5, out)

        sheet = os.path.join(tmp, "contact.png")
        code, out = run("contact_sheet.py", "--file", deck, "--out", sheet)
        check("contact_sheet: PNG written", code == 0 and os.path.exists(sheet), out)

        for mode in ("notes", "handout", "slides"):
            pdf = os.path.join(tmp, f"selftest.{mode}.pdf")
            code, out = run("export_pdf.py", "--file", deck, "--mode", mode, "--out", pdf)
            ok = code == 0 and os.path.exists(pdf) and open(pdf, "rb").read(4) == b"%PDF"
            check(f"export_pdf: {mode} PDF written", ok, out)
    else:
        say("\n(skipped the PowerPoint scripts; run with --com on Windows)")

    failed = results.count(False)
    say(f"\n{len(results) - failed}/{len(results)} passed  -  output in {tmp}")
    log = os.path.join(tmp, "selftest-report.txt")
    open(log, "w", encoding="utf-8").write("\n".join(report) + "\n")
    print(f"report: {log}  (paste or attach this file when reporting results)")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
