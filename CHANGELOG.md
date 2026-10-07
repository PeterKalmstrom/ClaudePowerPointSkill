# Changelog

## v1.0.0 — 2026-10-07

- Renamed `configuring-powerpoint-mcp` → `building-powerpoint-decks`; repo renamed to `claude-powerpoint-skill`.
- Split into a short `SKILL.md` plus `reference/`, with a start-here table and ten core rules.
- Helper scripts (`scripts/`): audit, bulk read, render via Save As JPEG, contact sheet, notes/handout PDF,
  word-break check, slide-size check, cover-crop, backup snapshot, and a self-test (22/22 on Windows).
- CI: script compile, docs link/anchor/size check, self-test, eval format check. Trigger evals.
- New guidance: `mcp<2` startup fix, embedded-font rendering, keeping hand edits, slide size, cropping, deck
  audits with ~40 defect codes, presenter prep, animation, video compression, Veo production lessons, claim
  titles, cognitive load, notes structure, sensitivity labels, building without COM.
- Fixed: `export_pdf.py` on Windows (pywin32 `PrintRange` trap).
- Before/after example with a native, editable chart.

## 2026-05

- First public version (`configuring-powerpoint-mcp`).
