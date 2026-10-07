"""Self-test for every script in this folder, against a deck it builds itself.

    # any OS: the scripts that need no PowerPoint
    uvx --with python-pptx --with pillow python scripts/selftest.py
    # Windows with PowerPoint installed: everything
    uvx --with python-pptx --with pillow --with pywin32 python scripts/selftest.py --com

The test deck (5 slides, 1440 x 810 pt) has known defects, so each check knows what to expect:
  1 "Decisions"       topic-label title, a 4-word body paragraph at 14 pt   -> ** AUDIT, label?
  2 claim title       short body at 24 pt                                    -> OK
  3 claim title       one long word in a narrow box at 60 pt                 -> broken word
  4 claim title       speaker notes                                          -> notes in bulk_read
  5 hidden slide                                                             -> [skip]
Exit code 0 = every check passed.
"""
import argparse
import json
import os
import subprocess
import sys
import tempfile

from PIL import Image
from pptx import Presentation
from pptx.util import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
results = []


def check(name, ok, detail=""):
    results.append(ok)
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail and not ok else ""))


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

    text(slide("Decisions"), "Small text below the floor", 14)
    text(slide("The check comes first"), "Plan, check, run", 24)
    text(slide("Long words break lines"), "Internationalization", 60, width=200, name="Narrow")
    s = slide("Notes carry the depth")
    text(s, "Short body", 24)
    s.notes_slide.notes_text_frame.text = "Key fact: notes are read by bulk_read."
    hidden = slide("Hidden slide")
    hidden._element.set("show", "0")
    prs.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--com", action="store_true", help="also run the PowerPoint (COM) scripts; Windows only")
    a = ap.parse_args()

    tmp = tempfile.mkdtemp(prefix="pptskill-selftest-")
    deck = os.path.join(tmp, "selftest.pptx")
    build_deck(deck)
    print(f"test deck: {deck}\n")

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
    import types
    sys.modules.setdefault("_ppt", types.SimpleNamespace(is_title=None, open_deck=None,
                                                         shape_texts=None, utf8_stdout=lambda: None))
    src = open(os.path.join(HERE, "audit_deck.py"), encoding="utf-8").read().split("utf8_stdout()\n")[0]
    ns = {}
    exec(src, ns)
    label = ns["looks_like_label"]
    check("audit_deck: 'Decisions' is a label", label("Decisions"))
    check("audit_deck: 'The check comes first' is a claim", not label("The check comes first"))
    check("audit_deck: a question is not flagged", not label("Why Claude?"))

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
        check("audit_deck: slide 1 min font 14pt", "14pt" in rows.get("1", ""), out)
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
        print("\n(skipped the PowerPoint scripts; run with --com on Windows)")

    failed = results.count(False)
    print(f"\n{len(results) - failed}/{len(results)} passed  -  output in {tmp}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
