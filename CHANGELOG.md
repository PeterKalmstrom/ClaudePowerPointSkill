# Changelog

## Unreleased

- **`scripts/build_deck.py`**: builds a deck from a JSON/YAML spec — 14 patterns (title, section, statement,
  big number, KPI, bullets, compare, process, timeline, quote, chart, table, image, 2×2 matrix), per-pattern
  limits checked before building, a design direction written into the theme or a company template, native
  charts with one highlighted finding, generated alt text, structured speaker notes, stable slide ids.
  `--lint` lints the result. Sample spec in `examples/spec/`.
- **20 design directions** (`scripts/directions.json`, written for this skill) and **reference/DESIGN.md**
  (Full HD type scale, spacing, chart and table defaults); **reference/BUILDER.md** (spec format).
- **Font sizes scale with slide width:** the 18 pt body floor is 27 pt on Full HD (1440 pt) slides, in
  `lint_deck.py`, `audit_deck.py` and the docs. The same text is two-thirds as big on a wider slide.
- `lint_deck.py` looks behind text boxes for the card or band they sit on when checking contrast, and
  measures rotated shapes as drawn.

- `scripts/lint_deck.py`: lints a deck from the file alone, on any OS — titles, body floor (with
  `--room-depth`), bullets, off-slide shapes, overlaps, empty placeholders (`--fix`), colour count, contrast,
  alt text, stretched pictures, chart palette/title/labels, notes, duplicate titles, font drift. JSON output.
- `scripts/extract_theme.py`: theme colours, fonts, layouts and placeholders from a .pptx or .potx; `--markdown`
  writes a brand-spec skeleton. New *Building from a template* section in LAYOUT.md.
- `scripts/render_lo.py` (LibreOffice renders) and `scripts/diff_renders.py` (before/after diffs with heat maps).
- `lint_deck.py` phase 2b: theme-aware contrast (theme colours, lumMod/tint, slide/layout/master background,
  WCAG large text = 18 pt or 14 pt bold) and 20 more codes — emoji icons, lorem ipsum, truncated text, centred
  long body, wide measure, tiny click targets, shadow overuse, off-palette fills, accent overload, saturated
  gradients, repeated words, weak focal hierarchy, grid monotony, stock imagery, default-font-only, and chart
  legend/ordinal-colour/accounting-format/label-collision checks. `extract_theme.py` reads every master's theme.
- Self-test grows to 34 checks without PowerPoint (36 with `--com`). Rule thresholds partly follow PointClaw.
- Before/after example now uses a real title placeholder and chart alt text (lint-clean).

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
