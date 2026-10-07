# Presenter prep

*Runs on: Windows + PowerPoint.* Part of the `building-powerpoint-decks` skill — start at [`SKILL.md`](../SKILL.md). Paths like `scripts/…` are relative to the skill folder.

## Presenter prep

Once a deck is content-complete **and audits clean**, and there is a real talk coming up:

1. **Timing markers** — one line at the top of each slide's notes, e.g. `[15 sec]`. Baseline = talk length ÷
   content slides; give openers and summaries 20–25 s, detail slides 10–15 s, votes ~30 s, the close 30 s+.
   Make the script idempotent (skip slides whose notes already start with `[`).
2. **Q&A panic sheet** — a markdown file next to the deck with the 15–25 hardest questions, each with a 2–4
   sentence sourced answer. Build it from the speaker notes: every challengeable claim becomes a question. End with
   a "last-resort exit" line for questions you can't answer.
3. **Printed presenter handout** — slide + notes per page:
   ```bash
   uvx --with pywin32 python scripts/export_pdf.py --file deck.pptx --mode notes
   uvx --with pywin32 python scripts/export_pdf.py --file deck.pptx --mode handout --slides-per-page 6   # audience
   ```
4. **Dress-rehearsal contact sheet** — `scripts/contact_sheet.py`, printed on A3; mark the section breaks.
5. **Rehearsal checklist** — run the deck in slideshow mode, out loud, with a stopwatch; cold-read the panic sheet
   (each answer ≤ 30 s); check projector, mic and clicker in the real room.

---
