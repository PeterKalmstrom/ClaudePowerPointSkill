"""Self-test for every script in this folder, against a deck it builds itself.

    # any OS: the scripts that need no PowerPoint
    uvx --with python-pptx --with pillow python scripts/selftest.py
    # Windows with PowerPoint installed: everything
    uvx --with python-pptx --with pillow --with pywin32 python scripts/selftest.py --com

The test deck (5 slides, 1440 x 810 pt) has known defects, so each check knows what to expect:
  1 "Decisions"       topic-label title, a 4-word body paragraph at 22 pt   -> body_below_floor, title_is_label
  2 claim title       short body at 24 pt                                    -> OK
  3 claim title       one long word in a narrow box at 60 pt                 -> broken word
  4 claim title       speaker notes                                          -> notes in read_deck
  5 hidden slide                                                             -> [skip]
Exit code 0 = every check passed.
"""
import argparse
import io
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
    s.notes_slide.notes_text_frame.text = "Key fact: notes are read by read_deck."
    hidden = slide("Hidden slide")
    hidden._element.set("show", "0")
    prs.save(path)


def picture_ratio(pic):
    w, h = Image.open(io.BytesIO(pic.image.blob)).size
    return w / h


def build_lint_edge_deck(path, photo):
    """Group children in child coordinates, white text on a full-bleed photo, white-on-white table."""
    from lxml import etree
    from pptx.dml.color import RGBColor
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(1440), Pt(810)
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Groups are measured on the slide"
    g = s.shapes.add_group_shape()
    for i in range(2):
        tb = g.shapes.add_textbox(Pt(100 + i * 450), Pt(300), Pt(400), Pt(100))
        tb.text_frame.text = "Grouped label"
    xfrm = g._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}grpSpPr/"
                           "{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
    xfrm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}chOff").set("x", str(Pt(3000)))
    xfrm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}chOff").set("y", str(Pt(3000)))
    for tb in g.shapes:  # move children into a far-away child space; the group still sits on the slide
        tb.left, tb.top = tb.left + Pt(2900), tb.top + Pt(2700)
    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    s2.shapes.title.text = "Text on photos is judged by eye"
    s2.shapes.add_picture(photo, 0, 0, Pt(1440), Pt(810))
    tb = s2.shapes.add_textbox(Pt(100), Pt(300), Pt(800), Pt(100))
    tb.text_frame.text = "White text over a dark photo"
    tb.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    s3.shapes.title.text = "Table text needs contrast too"
    tbl = s3.shapes.add_table(2, 2, Pt(100), Pt(300), Pt(800), Pt(200)).table
    for cell in tbl.iter_cells():
        cell.text = "Unreadable"
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
        cell.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
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
    from _rules import kRules
    label = kRules.LooksLikeLabel
    check("rules: 'Decisions' is a label", label("Decisions"))
    check("rules: 'The check comes first' is a claim", not label("The check comes first"))
    check("rules: a question is not flagged", not label("Why Claude?"))

    code, out = run("read_deck.py", deck)
    try:
        data = json.loads(out)
        check("read_deck: 5 slides with ids and layouts", len(data["slides"]) == 5 and all(s["layout"] for s in data["slides"]))
        check("read_deck: notes read", "Key fact" in data["slides"][3]["notes"])
        check("read_deck: hidden slide marked", data["slides"][4]["hidden"])
        check("read_deck: titles and font sizes read", data["slides"][0]["title"] == "Decisions"
              and any(22.0 in x.get("sizes_pt", []) for x in data["slides"][0]["shapes"]))
    except Exception as e:
        check("read_deck: valid JSON", False, f"{e}: {out[:300]}")

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
    from _rules import kRules
    dirs = json.load(open(os.path.join(HERE, "directions.json"), encoding="utf-8"))["directions"]
    from build_deck import kDeckDesign
    weak = []
    for d in dirs:
        quiet, muted = kDeckDesign.QuietAndMuted(d)
        if (kRules.ContrastRatio(d["text"], d["background"]) < 7 or kRules.ContrastRatio(muted, d["background"]) < 4.5
                or kRules.ContrastRatio(muted, quiet) < 4.5 or kRules.ContrastRatio(d["text"], quiet) < 7
                or kRules.ContrastRatio(d["accent"], d["background"]) < 3 or kRules.ContrastRatio(d["background"], d["accent"]) < 4.5):
            weak.append(d["id"])
    check(f"directions: all {len(dirs)} pass contrast (text 7:1; muted 4.5:1 on background and cards; "
          "background-on-accent 4.5:1)", not weak, ", ".join(weak))

    # regressions from the October 2026 code review
    edge = os.path.join(tmp, "edge")
    os.makedirs(edge, exist_ok=True)
    Image.new("CMYK", (400, 300), (0, 100, 0, 0)).save(os.path.join(edge, "cmyk.jpg"))
    rot = Image.new("RGB", (600, 300), (200, 30, 30))
    ex = rot.getexif()
    ex[0x0112] = 6  # stored landscape, shown portrait
    rot.save(os.path.join(edge, "rot.jpg"), exif=ex)
    espec = {"direction": "clean-corporate", "slides": [
        {"pattern": "image", "title": "CMYK photos build", "image": "cmyk.jpg"},
        {"pattern": "image", "title": "Phone photos stay upright", "image": "rot.jpg"},
        {"pattern": "chart", "type": "pie", "title": "Share splits three ways", "categories": ["A", "B", "C"],
         "series": [{"name": "Share", "values": [1, 2, 3]}]},
        {"pattern": "chart", "type": "line", "title": "Six lines rise", "categories": ["Q1", "Q2", "Q3"],
         "series": [{"name": f"s{i}", "values": [1, 2, i]} for i in range(6)]},
        {"pattern": "chart", "title": "Gaps in data still build", "categories": ["A", "B"],
         "series": [{"name": "s", "values": [1, None]}]},
        {"pattern": "big_number", "title": "Numbers as numbers work", "number": 42},
        {"pattern": "bullets", "title": "Loose notes work", "items": ["a"],
         "notes": {"facts": "one string", "qa": [{"q": "only a question"}]}}]}
    json.dump(espec, open(os.path.join(edge, "edge.json"), "w"))
    eout = os.path.join(tmp, "edge.pptx")
    code, out = run("build_deck.py", os.path.join(edge, "edge.json"), "--out", eout, "--lint")
    check("build_deck: edge cases build and lint clean (CMYK, EXIF, pie, 6 lines, None, numbers, notes)",
          code == 0 and "0 error(s), 0 warning(s)" in out, out[-500:])
    if code == 0:
        ep = Presentation(eout)
        pic = [sh for sh in ep.slides[1].shapes if sh.shape_type == 13][0]
        check("build_deck: EXIF-rotated photo placed upright", pic.image.size[0] < pic.image.size[1] * 2.5
              and abs(picture_ratio(pic) - pic.width / pic.height) < 0.02)
        pie = ep.slides[2].shapes
        ch = [sh for sh in pie if sh.has_chart][0].chart
        check("build_deck: pie slices labelled with their category", ch.plots[0].data_labels.show_category_name)
        notes = ep.slides[6].notes_slide.notes_text_frame.text
        check("build_deck: string 'facts' kept whole in notes", "- one string" in notes, notes)
    json.dump({"slides": [{"pattern": "table", "title": "Ragged rows fail", "header": ["a", "b"], "rows": [["1", "2", "3"]]},
                          {"pattern": "chart", "title": "Unknown highlight fails", "categories": ["A"], "highlight": "Z",
                           "series": [{"name": "s", "values": [1]}]}]}, open(os.path.join(edge, "bad.json"), "w"))
    code, out = run("build_deck.py", os.path.join(edge, "bad.json"), "--out", os.path.join(tmp, "x.pptx"))
    check("build_deck: ragged rows and unknown highlight rejected", code == 2 and "row 1 has 3 cells" in out
          and "is not one of the categories" in out, out)
    tpl = os.path.join(edge, "tpl.pptx")
    tp = Presentation()
    for layout in tp.slide_layouts:
        if layout.name == "Title Only":
            layout.name = "Headline"  # a template whose layouts aren't named the usual way
    tp.save(tpl)
    json.dump({"template": "tpl.pptx", "slides": [{"pattern": "statement", "title": "Templates keep their size"}]},
              open(os.path.join(edge, "tpl.json"), "w"))
    code, out = run("build_deck.py", os.path.join(edge, "tpl.json"), "--out", os.path.join(tmp, "tpl-out.pptx"), "--lint")
    check("build_deck: template without 'Title Only' builds without empty placeholders",
          "unused_placeholder" not in out and "slides ->" in out, out[-400:])
    lt = os.path.join(tmp, "lint-edge.pptx")
    build_lint_edge_deck(lt, photo)
    code, out = run("lint_deck.py", lt)
    check("lint_deck: grouped shapes measured on the slide (no false off-slide)", "offslide_shape" not in out, out)
    check("lint_deck: text on a photo not reported as white on white", "slide 2   a11y_low_text_contrast" not in out, out)
    check("lint_deck: table cell contrast checked", "Table cell text" in out, out)

    # text fitting: the builder shrinks text to fit and reports what can't; the linter sees the same overflow
    fit_spec = {"slides": [
        {"pattern": "bullets", "title": "Long bullets shrink until they fit", "items": [
            "Approve the partner budget for the coming two quarters, including regional enablement",
            "Name an owner per region who reports weekly on pipeline and partner activation",
            "Agree the Q3 review date and the metrics we will judge the rollout on",
            "Fund the partner portal and the onboarding content it needs before the first partners sign up",
            "Hire two partner managers with channel experience and give them a quota from the third quarter",
            "Publish the playbook internally and walk every regional team through it in a live session",
            "Revisit pricing for partners so that the margin works for them and for us at volume"]},
        {"pattern": "kpi", "title": "Labels too long for six cards", "metrics": [
            {"value": "41.2 M", "label": "Total revenue across regions"}, {"value": "+1.5 pt", "label": "Gross margin change"},
            {"value": "0 %", "label": "Opex change"}, {"value": "12", "label": "New partners"},
            {"value": "3.2 d", "label": "Days to close"}, {"value": "98 %", "label": "Renewals"}]}]}
    json.dump(fit_spec, open(os.path.join(edge, "fit.json"), "w"))
    fout = os.path.join(tmp, "fit.pptx")
    code, out = run("build_deck.py", os.path.join(edge, "fit.json"), "--out", fout)
    check("build_deck: reports text it cannot fit (exit 3)", code == 3 and "slide 2" in out and "does not fit" in out
          and "slide 1" not in out.split("does not fit")[0].split("fit:")[-1], out)
    fp = Presentation(fout)
    sizes = [r.font.size.pt for sh in fp.slides[0].shapes if sh.name == "Points"
             for p in sh.text_frame.paragraphs for r in p.runs]
    check("build_deck: long bullets shrunk, but not below the 27 pt floor", sizes and 27 <= min(sizes) < 34, str(sizes))
    code, out = run("lint_deck.py", fout)
    check("lint_deck: sees the same overflow (text_overflow on slide 2)", "slide 2   text_overflow" in out, out)
    ov = os.path.join(tmp, "overflow.pptx")
    op = Presentation()
    op.slide_width, op.slide_height = Pt(1440), Pt(810)
    sl = op.slides.add_slide(op.slide_layouts[5])
    sl.shapes.title.text = "Fixed-size boxes overflow"
    box = sl.shapes.add_textbox(Pt(100), Pt(200), Pt(500), Pt(80))
    box.text_frame.word_wrap = True
    box.text_frame.auto_size = 0  # MSO_AUTO_SIZE.NONE: the box keeps its size
    box.text_frame.text = "This sentence is far too long for a small fixed box and will spill out below it on the slide"
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(32)
    wide = sl.shapes.add_textbox(Pt(700), Pt(200), Pt(200), Pt(200))
    wide.text_frame.word_wrap = True
    wide.text_frame.text = "Internationalization"
    wide.text_frame.paragraphs[0].runs[0].font.size = Pt(48)
    op.save(ov)
    code, out = run("lint_deck.py", ov)
    check("lint_deck: fixed-size box overflow found", "text_overflow" in out, out)
    check("lint_deck: word wider than its box found", "word_breaks" in out, out)

    # harvest_edits: hand edits survive a rebuild
    hv = os.path.join(tmp, "harvest")
    os.makedirs(hv, exist_ok=True)
    hspec = {"direction": "clean-corporate", "slides": [
        {"id": "a", "pattern": "statement", "title": "First claim stands"},
        {"id": "b", "pattern": "statement", "title": "Second claim stands"},
        {"id": "c", "pattern": "statement", "title": "Third claim stands"}]}
    json.dump(hspec, open(os.path.join(hv, "spec.json"), "w"))
    deck_h = os.path.join(hv, "deck.pptx")
    run("build_deck.py", os.path.join(hv, "spec.json"), "--out", deck_h)
    run("harvest_edits.py", "manifest", deck_h)
    hp = Presentation(deck_h)
    hp.save(os.path.join(hv, "resaved.pptx"))
    hp.slides[0].shapes.title.text = "First claim, reworded by a colleague"
    extra = hp.slides.add_slide(hp.slide_layouts[5])
    extra.shapes.title.text = "A slide a colleague added"
    extra.shapes.add_picture(photo, Pt(100), Pt(200), Pt(300), Pt(200))
    lst = hp.slides._sldIdLst
    moved = list(lst)[-1]
    lst.remove(moved)
    lst.insert(1, moved)  # after slide a
    gone = list(lst)[3]   # slide c
    hp.part.drop_rel(gone.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"))
    lst.remove(gone)
    hp.save(deck_h)
    code, out = run("harvest_edits.py", "harvest", os.path.join(hv, "resaved.pptx"), "--manifest",
                    os.path.join(hv, "deck.manifest.json"), "--store", os.path.join(hv, "none"))
    check("harvest_edits: a plain re-save is not an edit", "0 slide(s) harvested, 0 deletion(s)" in out, out)
    code, out = run("harvest_edits.py", "harvest", deck_h, "--store", os.path.join(hv, "store"))
    check("harvest_edits: finds the edit, the added slide and the deletion",
          "edited   a" in out and "added" in out and "deleted  c" in out, out)
    hspec["slides"][1]["title"] = "Second claim, updated in the spec"
    json.dump(hspec, open(os.path.join(hv, "spec.json"), "w"))
    rebuilt = os.path.join(hv, "rebuilt.pptx")
    run("build_deck.py", os.path.join(hv, "spec.json"), "--out", rebuilt)
    final = os.path.join(hv, "final.pptx")
    code, out = run("harvest_edits.py", "restore", rebuilt, "--store", os.path.join(hv, "store"), "--out", final)
    titles = [sl.shapes.title.text for sl in Presentation(final).slides] if os.path.exists(final) else []
    check("harvest_edits: restored deck keeps edits, additions, deletions and spec changes",
          titles == ["First claim, reworded by a colleague", "A slide a colleague added",
                     "Second claim, updated in the spec"], str(titles) + out)
    import zipfile, collections
    dups = [k for k, v in collections.Counter(zipfile.ZipFile(final).namelist()).items() if v > 1] if titles else ["?"]
    check("harvest_edits: restored file has no duplicate parts", not dups, str(dups))

    # notes: figures on a slide need a source in the notes
    nspec = {"slides": [
        {"pattern": "big_number", "title": "Revenue grew this year", "number": "+12", "unit": "%", "notes": "Strong year."},
        {"pattern": "big_number", "title": "Costs fell this year", "number": "-4", "unit": "%",
         "notes": {"key_fact": "Costs fell 4 %.", "sources": ["Annual report 2026, p. 12"]}}]}
    json.dump(nspec, open(os.path.join(edge, "notes2.json"), "w"))
    nout = os.path.join(tmp, "notes2.pptx")
    run("build_deck.py", os.path.join(edge, "notes2.json"), "--out", nout)
    code, out = run("lint_deck.py", nout)
    check("lint_deck: figures without a source in the notes flagged; sourced ones not",
          "slide 1   figure_without_source" in out and "slide 2   figure_without_source" not in out, out)

    # spec schema: generated from the builder, committed, and the sample validates against it
    code, out = run("build_deck.py", "--print-schema")
    committed = open(os.path.join(HERE, "spec.schema.json"), encoding="utf-8").read()
    check("spec schema: committed spec.schema.json matches build_deck --print-schema",
          code == 0 and json.loads(out) == json.loads(committed), "run: python scripts/build_deck.py --print-schema "
          "> scripts/spec.schema.json")
    try:
        import jsonschema
        jsonschema.validate(json.load(open(sample, encoding="utf-8")), json.loads(committed))
        check("spec schema: the sample spec validates", True)
        typo = {"slides": [{"pattern": "kpi", "title": "Typos are caught", "colour": "red", "metrics": [
            {"value": "1", "label": "a"}, {"value": "2", "label": "b"}, {"value": "3", "lable": "c"}]}]}
        json.dump(typo, open(os.path.join(edge, "typo.json"), "w"))
        code, out = run("build_deck.py", os.path.join(edge, "typo.json"), "--out", os.path.join(tmp, "t.pptx"))
        check("build_deck: misspelt fields named, nested ones too", code == 2 and "'colour'" in out
              and "'lable' in metrics/2" in out, out)
    except ImportError:
        say("(skipped schema validation: pip install jsonschema)")
    code, out = run("build_deck.py", sample, "--plan")
    check("build_deck --plan: prints the story without building", code == 0 and "Deck plan (14 slides" in out, out)

    # fix_deck: mechanical repairs on the defect fixtures, then lint again
    fixed = os.path.join(tmp, "taste.fixed.pptx")
    code, out = run("fix_deck.py", taste, "--out", fixed)
    for want in ("palette", "legend", "numfmt", "alt"):
        check(f"fix_deck: '{want}' fix applied on the taste deck", f"\n{want} " in "\n" + out, out[-600:])
    after = out.split("What's left")[-1]
    check("fix_deck: chart defects gone after fixing", not any(c in after for c in (
        "chart_default_palette", "chart_legend_steals_plot", "chart_accounting_zero_dash")), after)
    code, out = run("fix_deck.py", messy, "--out", os.path.join(tmp, "messy.fixed.pptx"))
    after = out.split("What's left")[-1]
    check("fix_deck: empty placeholder removed and stretched picture un-stretched",
          "unused_placeholder" not in after and "picture_stretched" not in after, out[-600:])
    code, out = run("fix_deck.py", ov, "--out", os.path.join(tmp, "overflow.fixed.pptx"), "--only", "fit")
    check("fix_deck: overflowing text shrunk (not below the floor)", "\nfit " in "\n" + out and "-> 27 pt" in out
          or "-> 2" in out, out[-500:])
    code, out = run("fix_deck.py", taste, "--out", taste)
    check("fix_deck: refuses to overwrite its input", code == 2, out)

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
