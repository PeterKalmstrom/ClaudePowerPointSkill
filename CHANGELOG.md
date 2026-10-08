# Changelog

## 2.1.0

- **`build_deck.py` fills the slide.** Every pattern lays its content out over the whole body area: KPI tiles,
  compare and risk cards take the full height, timelines sit on the slide's middle, bullets of up to four short
  items become full-width bands, and sibling boxes share one text size. Text grows up to a ceiling per role
  (bullets 44 pt, card points 40, details 36, KPI values 120 on one line) and nothing of four or more words is set
  below the 27 pt Full HD floor. Fixes from a blind benchmark (October 2026), where the skill's decks were judged
  sparse and small: process/timeline details were 24 pt (`body_below_floor` on specs within the limits); a
  highlighted step's number sat in the chevron's notch (it is now inside the arrow, clear of the notch); `compare`
  headings had no limit (now 40 characters, 24 with three columns, checked before building); content sat in the
  top half of the slide.
- **Five new patterns:** `email` (a mocked message - from, to, subject, body, attachment - with numbered callout
  markers and notes), `kpi_chart` (two to four KPI tiles beside a native chart), `cost_table` (rows with an automatic
  total row, the total large beside it, an optional note), `quiz` (lettered options; the answer in the notes, or on
  an extra answer slide with `"reveal": "slide"`) and `risks` (cards with likelihood and impact chips and the
  mitigation). Schema, `--plan`, limits, docs and self-test cover them; the sample spec shows all 19 patterns.
- **Speaker notes and facts:** the builder lists every slide whose spec has no notes; `CONTENT.md` adds *Every
  slide gets notes; never invent facts beyond the brief* (assumptions marked `Assumption:` in the notes) and *The
  ask states its reasons and its cost*, linked from the core rules and the anti-patterns table.
- **`lint_deck.py`:** `word_budget` is per pattern on decks `build_deck.py` made (bullets 30, compare 32, process
  36, email 130, risks 64 …; `--budget` scales them); `figure_without_source` skips the cover and section dividers;
  `repeated_word` looks at display type only (≥ 45 pt on Full HD), not grown body text.
- **Contact sheets on any OS:** `render_lo.py --sheet` writes the renders plus `contact.png`; `contact_sheet.py
  --renders <folder>` builds a sheet from renders already on disk (`--file` still renders through PowerPoint).
- **`fix_deck.py --in-place`** fixes the input itself after keeping the original as `<deck>.pptx.bak`.

- **PowerPoint Live, round 3:** a **Rehearse** view that drives the real slide show (`powerpoint_slideshow`:
  start/next/previous/black/end/state) with notes, next slide, elapsed and per-slide timers and ahead/behind against
  the timing markers; a **talk-length meter** (markers, else notes or visible words at 130 wpm; per slide and section
  in the sorter, a bar against a remembered target in the Storyline; `deck_outline` gains `visible_words` and the
  estimates); a **Design** gallery previewing the slide in six directions from a copy (`design_previews`) and
  `powerpoint_apply_direction`; a **Theme** panel (`theme_info`, `powerpoint_set_theme`) with editable swatches,
  fonts and contrast pass/fail; **layout variants** Claude makes as tagged hidden slides
  (`powerpoint_layout_variants`), compared in a strip and kept with `powerpoint_choose_variant` or dropped with
  `powerpoint_discard_variants`. Every change saves a version first.

- PowerPoint Live: code run through `powerpoint_run` that closes the deck no longer halts
  the server - the result says `"deck_closed": true`; a Save As is followed to the new file name.
- **One error-handling and code style for all Python** (`scripts/kShared.py`): every method guards on
  `kS.ErrorMode`, wraps its body, reports to `kS.GlobalErrorHandler` and returns a safe default; the first error
  halts the run (PowerPoint Live: Resume). Expected states raise `ToolInputException` / `ToolReportableException`.
  k-prefixed classes, PascalCase, no lambdas or nested functions. `tools/check_kpattern.py` enforces it in the
  self-test and CI. Reported errors can be offered to the support flow (`KPS_ERROR_WEBHOOK`).
- **Error reports work out of the box:** with the person's yes, a report now goes to a public error relay
  (`kErrorReport.EstateUrl()`), which carries no key and cleans and rate-limits what it passes on.
  `KPS_ERROR_WEBHOOK` still overrides the address. A refused report says why (too many, too large, not in the
  expected form), and each field is cut to the length the relay accepts.
- **PowerPoint Live face-lift:** one design system across all four views (tokens, light/dark, icons, chips,
  segmented tabs, switch, empty states), narrow-panel layout, keyboard and screen-reader labels.

- **PowerPoint Live, round 2:** change highlights; lint findings drawn on the slide with live one-click fixes
  (`powerpoint_fix`); before/after slider; point at a shape or area and ask Claude (`ui/message`); a version saved
  before every change with a History view and `powerpoint_restore`; slide sorter multi-select, hide/unhide and
  section add/rename/delete (`powerpoint_sections`, `powerpoint_hide`); a Storyline view of the titles; an
  accurate (Save As) preview toggle.

- **No third-party PowerPoint server any more:** Windows mode drives PowerPoint directly through COM
  (`pywin32`) with the skill's own scripts. `powerpoint-mcp` and every reference to its tools are removed;
  `reference/SETUP.md` is rewritten (install uv, verify with `selftest.py --com`).

- **PowerPoint Live (`mcp-app/`):** the skill's own MCP server for Windows — open, run Python against and show
  slides — with an MCP App view that shows the current slide live as Claude, a script or a person changes it.
  A slide sorter view groups thumbnails by section and rearranges slides by drag and drop. Registered by the plugin.

## 2.0.0

- **`scripts/fix_deck.py`:** repairs what `lint_deck.py` finds mechanically — empty placeholders, text below the
  floor, overflowing boxes, stretched pictures (resized, or cropped with `--crop-photos`), missing chart/table alt
  text, Office default chart colours, top/bottom legends, Accounting axis formats. Writes a copy, then re-lints.
- **Spec schema:** `scripts/spec.schema.json` (generated by `build_deck.py --print-schema`) for editors and
  agents; spec errors now name nested typos (`'lable' in metrics/2`). `--plan` prints the slide plan without
  building.
- **Plugin packaging:** `.claude-plugin/plugin.json` and `marketplace.json` — install with `/plugin`.

- **Text overflow, any OS:** `scripts/_measure.py` wraps text with real font metrics (or metric-compatible twins:
  Carlito for Calibri, Liberation for Arial/Times) — 24 of 24 test boxes matched LibreOffice's line count.
  `lint_deck.py` adds `text_overflow`, `text_shrinks` and `word_breaks`, and uses the grown height of
  "resize to fit" boxes for overlap and off-slide checks. `build_deck.py` shrinks text to fit (never below the
  floor) and exits 3 with a list of anything that still doesn't fit.
- **`scripts/harvest_edits.py`:** manifest → harvest → restore keeps people's edited, added and deleted slides
  (with pictures and charts) across a rebuild; a plain re-save is not an edit.
- **`scripts/read_deck.py`** (any OS) replaces `bulk_read.py`: ids, layouts, shapes with positions and sizes,
  charts, tables, notes; `--text` outline. **`audit_deck.py` is retired** — `lint_deck.py` covers it on any OS.
- `lint_deck.py`: `figure_without_source`; group shapes measured in slide coordinates; text on photos no longer
  misreported; table cell contrast; `--fix` matches by shape id.
- `build_deck.py` fixes: CMYK and EXIF-rotated photos, pie slices labelled and shaded, line series shaded,
  template mode keeps the template's size and drops subtitle placeholders, YAML numbers, ragged tables and unknown
  highlights rejected, loose notes accepted, no temp files; quiet card colour chosen per direction so muted text
  keeps 4.5:1 (12 of 20 directions failed before).
- `render_lo.py` numbers renders by slide (hidden slides skipped, not renumbered); `diff_renders.py` reports
  size changes and pairs names case-insensitively; `extract_theme.py` reads .potx in memory.
- Docs: every size example converted to Full HD (1440 × 810).

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
- Self-test grows to 34 checks without PowerPoint (36 with `--com`). Rule thresholds partly follow the author's PowerPoint add-in.
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
