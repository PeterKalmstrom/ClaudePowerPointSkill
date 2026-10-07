# Headless PowerPoint and building without COM

*Runs on: headless checks: Windows + PowerPoint; no-COM: any OS.* Part of the `building-powerpoint-decks` skill — start at [`SKILL.md`](../SKILL.md). Paths like `scripts/…` are relative to the skill folder.

## Headless PowerPoint for checks

Scripts that open decks for checking or rendering (audits, renders, comparisons in CI) should not disturb the
user's PowerPoint or leave debris behind:

- **Work on a copy.** Open it with `Presentations.Open(path, ReadOnly=-1, Untitled=-1, WithWindow=0)`; this creates
  no recovery record, and PowerPoint exits by itself 1–3 s after `Quit()`.
- **Don't run alongside the user's PowerPoint** for unattended jobs: refuse if a `POWERPNT.EXE` process exists,
  rather than attaching to it. (The scripts in `scripts/` are for interactive use: they reuse the open deck and
  quit PowerPoint only if they started it.) **Never kill PowerPoint** — wait for it to exit and report it if it
  doesn't.
- **Turn `DisplayAlerts` on when validating a generated file.** With alerts off, PowerPoint **repairs broken files
  silently** (e.g. a non-numeric `<a:off x="abc">`) and the check passes. With alerts on, a damaged file shows the
  "PowerPoint found a problem with content… repair" dialog, which blocks `Presentations.Open` — so a validator needs a
  timeout and treats the prompt as a failure.
- **Rendering is deterministic.** Two exports of the same file are pixel-identical, so in a before/after comparison
  any difference is a real change — there's no noise band to hide in. Mean pixel difference is weak for small text
  moves; a per-region heat map or perceptual hash catches them.
- **Give slides stable ids.** Set each generated slide's `Name` (`<p:cSld name>`) to a permanent id so reports and
  baselines say *which* slide changed, not just its position.
- **Sections** (`p14:sectionLst`) are worth checking too: one section list, every slide in exactly one section,
  unique names, at most 512 sections.

---

## Building .pptx without PowerPoint (no COM)

On Linux, macOS or CI there is no COM. Write the OOXML directly with **python-pptx** (Python) or a library such as
PptxGenJS (Node). Everything about content — anchor types, word budgets, the 18pt floor, notes, no text in images —
still applies. What changes:

- **No renderer.** You can't LOOK without PowerPoint (or LibreOffice, whose layout differs — fine for catching gross
  errors, not for line breaks). Plan a Windows render pass before shipping, or keep text conservative.
- **Set the slide size explicitly:** python-pptx's default template is **720 × 540 pt (4:3)**.
  `prs.slide_width, prs.slide_height = Pt(1440), Pt(810)`.
- **`add_picture` stretches** when given both width and height — crop first (`scripts/cover_crop.py`).
- **Name every shape and slide** (`shape.name`, `<p:cSld name>`), exactly as in COM builds.
- **Validate the package** before handing it over: re-open it with python-pptx, and if possible with the Open XML
  SDK validator. A file that python-pptx reads can still trigger PowerPoint's repair prompt — duplicate zip entries,
  dangling rels and content-type mismatches are the usual causes. Read every zip entry before rewriting any (writing
  then re-reading entries raises "Overlapped entries").
- **Embedding fonts** is possible (`.fntdata` parts) but check the font's licence and `fsType` first — many
  commercial display fonts forbid embedding.
- Copying slides between presentations: see the python-pptx notes under *A generated deck that people also edit*.

---
