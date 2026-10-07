"""One-pass structural audit: visible words per slide (title excluded) and the smallest
BODY font size. Body = a paragraph with 4+ words; shorter runs are labels (12-14 pt allowed).
Sizes are judged against a 960-pt-wide slide: on Full HD (1440 pt) the 18 pt floor becomes 27 pt.

    uvx --with pywin32 python audit_deck.py --file deck.pptx [--budget 12] [--floor 18]

Status: OK | ok-tight (within 2 words of budget) | ** AUDIT | [skip] (hidden slide).
"label?" after the status: the title looks like a topic label, not a claim (1-2 words, no
common verb, not a question). It is a prompt to look - dividers, agenda and Q&A are fine.
An ** AUDIT row is either a defect to fix or an accepted anchor-type exception (gallery,
knowledge graph, quote, chart, reference) - see "Auditing a deck" in SKILL.md.
"""
import argparse

from _ppt import is_title, open_deck, shape_texts, utf8_stdout
from _rules import looks_like_label

utf8_stdout()
ap = argparse.ArgumentParser()
ap.add_argument("--file", required=True)
ap.add_argument("--budget", type=int, default=12)
ap.add_argument("--floor", type=float, default=18)
a = ap.parse_args()

flagged = 0
print(f"{'#':>3}  {'words':>5}  {'min_pt':>6}  {'status':<10}  title")
print("-" * 80)
with open_deck(a.file) as pres:
    # judge sizes against a standard 960-pt-wide slide: on Full HD (1440 pt) the 18 pt floor is 27 pt
    scale = max(1.0, pres.PageSetup.SlideWidth / 960)
    floor = a.floor * scale
    for s in pres.Slides:
        title, full_title, words, min_pt = "", "", 0, None
        for sh in shape_texts(s.Shapes):
            tr = sh.TextFrame.TextRange
            if is_title(sh):
                full_title = tr.Text.replace("\r", " ").strip()
                title = full_title[:45]
                continue
            words += len(tr.Text.split())
            for p in range(1, tr.Paragraphs().Count + 1):
                para = tr.Paragraphs(p)
                if len(para.Text.split()) >= 4:
                    size = para.Font.Size  # mixed sizes report a sentinel; check runs instead
                    sizes = [para.Runs(r).Font.Size for r in range(1, para.Runs().Count + 1)] \
                        if size <= 0 or size > 4000 else [size]
                    min_pt = min([x for x in sizes + [min_pt] if x])
        if s.SlideShowTransition.Hidden:
            status = "[skip]"
        elif words > a.budget or (min_pt is not None and min_pt < floor):
            status, flagged = "** AUDIT", flagged + 1
        elif words > a.budget - 2:
            status = "ok-tight"
        else:
            status = "OK"
        pt = f"{min_pt:g}pt" if min_pt else "-"
        hint = "  label?" if status != "[skip]" and looks_like_label(full_title) else ""
        print(f"{s.SlideIndex:>3}  {words:>5}  {pt:>6}  {status:<10}  {title}{hint}")
    print(f"\n{pres.Slides.Count} slides, {flagged} flagged (body floor {floor:g} pt)")
