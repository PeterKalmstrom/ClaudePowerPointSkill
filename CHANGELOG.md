# Changelog

## Unreleased

- **Rough edges from benchmark round 6:**
  - *One word count*: `--plan` and lint count visible words the same way (`kRules.WordCount`: tokens with a letter
    or digit); `--check` counts a built slide's words from its spec, as `--plan` does.
  - *Measurable targets*: a time window alone ("<= last 6 months", "by Q3") is not a target; a number needs a unit,
    a comparator or a count ("4 of 4") and a named metric (`kRules.Measurable`, spec warning and
    `target_not_measurable`).
  - *The ask states its cost*: new lint `ask_without_cost` and a spec warning when a decision box nowhere says
    what it costs (money, time, FTE).
  - *Quiz answer slides*: a spec warning when `answer_title` is missing or only a letter ("Answer: B").
  - *`{change|abs}`*: the change without its sign, for text that already says the direction ("up {change|abs}").
  - *Titles in wide theme fonts*: the builder shrinks a title until it takes no more lines than lint accepts
    (3 for a display title, else 2) - Verdana on Windows made a statement title 4 lines.
  - *Windows: a closed stdout pipe* (`--plan | head`) raises EINVAL there, not BrokenPipeError; it is now a
    quiet normal end on Windows too.
  - Self-test: rule checks for the above (173 checks).

- **Speed round: fewer loops for an agent** (round 5: 205 s / 107k tokens / 26 tool calls for three decks, against
  plain Claude's 165 s / 77k / 10 - the scripts take seconds, the cost was one problem class per build loop):
  - *Automatic fixes* (`scripts/_autofix.py`, `kAutoFix`): when text does not fit, the builder changes the spec in
    memory and rebuilds (up to six rounds): drops filler words, writes units short (`41 percent` → `41 %`,
    `19.8 MUSD` → `19.8 M` with `(USD)` moved into the tile label when it has room), cuts the longest detail text
    of the unfit shape at a clause boundary and puts the whole original in the notes as `MOVED FROM SLIDE:`, and
    splits a `bullets` slide that says `"allow_split": true` (new spec field; also before the build when it has
    more than seven items). Every change prints as `auto: slide N (id) field: ...`; `--no-auto` switches it off.
    Exit 3 only when text still does not fit after the fixes.
  - *One-call check*: `build_deck.py spec.json --out deck.pptx --check` = build + lint (in process) +
    `render_lo --sheet` when LibreOffice is installed + one summary: every problem of every class (spec, fit,
    every lint finding) sorted by slide, with the shape and an `edit:` line naming the spec field, the contact
    sheet's path, and a `CHECK-JSON` line. Limit-only spec errors still build under `--check` so the rest is
    reported in the same pass (exit 2). The builder records problems structured (`kDeckBuilder.Report`, `Issues`,
    `LintItems`) instead of only as text.
  - *`--plan` pre-checks* (`kPlanCheck`): visible words against the pattern's budget (advice) and the title
    measured in the deck's heading font, as `plan:` lines, plus a one-line tally.
  - *Lean `SKILL.md`* (18,655 → about 7,700 bytes, 169 → 84 lines): the New-deck path is a six-step checklist;
    the long prose (new deck in detail, showcase-first, the iteration loop, the anti-patterns table) moved
    verbatim to the new `reference/WORKFLOW.md`; BUILDER.md gains *Automatic fixes and --check*.
  - Benchmark README: a *Speed* section (method and the measured time / tokens / tool calls) and a speed column in
    every round's table.
  - New self-test group `CheckAutoAndCheck` (11 checks); the three detection checks that expect exit 3 run with
    `--no-auto`. Schema regenerated (`allow_split`).

- **Rough edges from benchmark round 5** (thisskill9: each passed builder and lint and only the render showed it):
  - *A tile value wrapping onto its label* ('19.8 MUSD' on a kpi tile): `FitLine` measures a value in the theme
    font AND the face LibreOffice substitutes (`kDeckBuilder.LineWidth`) and reports `fit: … value does not fit on
    one line` (exit 3) when even the smallest size wraps; lint's `unwanted_wrap` measures the substitute too.
  - *A risk mitigation running past its card*: `Fit` measures every box with `kMeasure.LooseHeight` (the taller of
    the theme font and its substitute), and lint's `text_overflow` measures body text the same way.
  - *Content on the footer*: the footer band is reserved - a text box reaching into it is `fit:` (exit 3); a big
    number with points shrinks the number, then the points (to the floor) to end above it; a timeline's labels fit
    the room below the rail as well as above. The builder's layout check now reports every lint **error** of its
    own output (`shape_overlap` etc.), plus `unwanted_wrap` and `body_below_floor`, as unfit text: `--lint` printing
    errors with exit 0 cannot happen any more.
  - *`build_deck.py --plan | head` filed an error report*: a BrokenPipeError on a stdout whose reader has gone
    (`kS.OutputGone`) is the normal end of the output - the handler passes it on unreported, `kRun.Main` /
    `kRun.Finish` point stdout at devnull and exit quietly with the code so far (`kRun.QuietEnd`).
  - `grid_monotony` no longer fires on a built timeline or on a row where one box differs in fill, weight or size
    (the highlight).
  - A `statement` with `decision` takes `points`, as the docs promised: the ask, then the reasons as numbered cards
    in a row (`ReasonCards`).
  - Points, support lines and captions (under a big number, beside a statement, under a decision) never go below
    the 27 pt floor however short (`Fit(..., Floor=)`), and lint's `body_below_floor` checks these roles at any
    length and size.
  - New self-test check group `CheckRound5` (7 checks).

- **No title widow in the substitute face either**: `kDeckBuilder.BalancedWidth` now balances the theme font AND
  the face LibreOffice substitutes for it (`kMeasure.Substitute`, new helper `AnyWidow`), `Title` drops up to 8 pt
  when no width balances both, and lint's `title_widow` measures both faces (naming the substitute). The round-4
  statement claim no longer leaves 'fast' alone in DejaVu Sans; a new self-test check proves it under `--ci-fonts`.

- **Titles measured right under any font set** (the self-test passed locally but failed 6 checks on the CI runner,
  where fontconfig substitutes the wide DejaVu Sans for Aptos instead of Inter): a content title whose own font and
  LibreOffice's substitute wrap differently is no longer narrowed to two lines that overflow its box - every
  candidate size/width is measured against the title's height, and failing all, the title shrinks until the loose
  wrap fits. A widow is checked in the substitute too. A statement claim no longer shrinks below the display size
  while taking three lines. `PPTSKILL_FONT_DIRS` overrides the scanned font folders; `selftest.py --ci-fonts`
  reproduces the runner's fonts; the balanced-title check measures with the deck's theme font and its substitute.

- **Computed figures** (blind benchmark round 4 ranked the skill first, 8.50, but its QBR said 2026 revenue was
  "21.8 MUSD" - "sum of quarters" - for quarters of 4.1 + 4.6 + 5.2 + 5.9 = 19.8; nothing checked computed figures):
  - New `scripts/_figures.py`: every text value is scanned for figures that read as derived - a total (*sum*,
    *total*, *combined*, an addition written out), an average, a change (`+44 %`, *up 44 %*, beside `4.1 to 5.9`),
    a share (`70 % of …` beside `14/20`) and equations (`44 % = 5.9 / 4.1 - 1`) - and each is compared with what the
    deck's chart series, metric trends, cost and table columns and `facts` give. A figure that matches none but is
    close to one (a slip, not another quantity) is a `spec warning: figure: …` (exit code unchanged); `lint_deck.py`
    reports `figure_mismatch` on the built deck from the charts' data and the text (a tile's card is one context).
    A table's *Total* row is checked against its column.
  - Deck-level `facts` (named numbers from the brief) and figure tokens: `{sum}`, `{total}`, `{average}`,
    `{change}`, `{share}`, `{first}`, `{last}`, `{count}` compute from the metric's trend, the slide's chart series
    or cost rows, or the big number's share; `{sum:revenue_q}` from a fact list, `{name}` a single-number fact. A
    token with nothing to compute from is a spec error. `--plan` lists the facts.
  - `CONTENT.md` *Compute derived figures, never by hand*; `BUILDER.md` *Computed figures*; `SKILL.md` start-here
    and anti-pattern rows.
- **Rough edges from round 4:**
  - `kicker_title_overlap` in the LibreOffice render: LibreOffice draws a theme font that is not installed
    (Georgia, Segoe UI on Linux) with a wider substitute, so a title measured as one line took two and grew up into
    its kicker while lint and the build passed. `kMeasure.Substitute` asks fontconfig which face that is;
    `LooseLines` counts lines in the wider of the two. The builder makes both renderers wrap a title alike (up to
    4 pt smaller, or a narrower balanced box) or reserves room for the longer one; lint's kicker check uses the
    same count; a statement's claim box is sized for it too. Checked by rendering 240 two-line titles in 3
    directions with `render_lo.py`: no kicker within 8 pt of a title (13 of 20 collided before in Georgia decks).
  - `kpi` takes 2-6 metrics (two wide tiles), `kpi_chart` 1-4. A count outside a pattern's range is a spec error
    that names the pattern that fits ("one number is a 'big_number' slide …").
  - The email body is measured by each paragraph's real wrapped height (in the wider of the body font and its
    substitute), not a fixed slack per paragraph, so shortening a paragraph clears the overflow report.
  - Tile trends are never tiny: with no room for bars at least 72 pt tall (a `decision` below the tiles), the
    trend becomes the tile's note (`Q1 4.1 → Q4 5.9`) and a `TREND (…)` line in the speaker notes.
  - `timeline` without dates draws undated sequential stages (numbered badges on the rail), so "what happens next"
    needs no invented "Step 1 / Step 2" dates; a timeline with some dates missing is a spec error.
  - A `statement` with `points` has its kicker (above the claim, as wide as the claim's column).
  - `build_deck.py --slides 1,3-4` builds only those spec slides, with the deck's theme, page numbers and section
    kickers, reporting spec mistakes only for them - the showcase-first step from the full spec; documented in
    `SKILL.md`'s unattended rule and `CONTENT.md`.
- Schema regenerated (`facts`, optional timeline `date`, new metric counts); self-test 153 checks.

- **Data shown as data** (blind benchmark round 3 ranked the skill first overall, but its weakest criterion was
  data presentation, 6.3 vs 7.7: "retention slide has no chart, tiny unlabelled bar glyphs, the ask slide is
  text-only, sparse slides, vague targets"):
  - A `kpi` metric with a `trend` (on a slide of up to four metrics, no decision) becomes a native column chart in
    the metric's own card, with data labels and `trend_labels` as categories; a before/after (churn 1.8 → 1.2) is
    a two-value trend. `kpi_chart` without `categories`/`series` charts a metric's trend. `"trend_chart": false`
    opts out.
  - Tile trends are labelled bars: first and last value above, first and last period (`trend_labels`) below, at
    the 24 pt label floor; the trend takes the tile's spare height.
  - `figure` on a closing `statement` with `decision`: the number the ask moves (or its cost) in a tile beside
    the decision box; owner and date move into the box.
  - `big_number`: a share (`41` + `%`, `14/20`) gets a dot grid beside it (`visual`: `dots` / `bar` / `none`);
    `points` add supporting facts under the caption.
  - `statement`: `points` (1–3) become numbered cards beside the claim; without them an accent bar anchors the
    claim and its support, centred on the slide.
  - Spec warnings (exit code unchanged) for a `metrics` target with no number, percent, date or comparison, a
    monthly timeline that skips one month, and a statement or big number with only one line of support;
    `lint_deck.py` adds `target_not_measurable`.
- **Rough edges from round 3:**
  - Kicker and title are one measured block: a two-line title's box grows down and the title shrinks until its
    ink (at the loosest renderer's line height) clears the kicker; lint `kicker_title_overlap`.
    `headline_too_long` now measures the title in its box (more than two lines; three for display titles)
    instead of counting 55 characters. Balanced titles also check the renderer's width (no widow at the edge).
  - Number + unit is glued in titles too (`44 %` on the cover), for `%`, `pt`, `k`/`M`/`B`, `kSEK`, `USD` …
  - No silent overflow: lint `text_overflow` now also fires when a text box that starts on a card or frame runs
    past its bottom (the email body, a risk card); the builder lints its own output and reports every
    `text_overflow`, `kicker_title_overlap` and `tile_text_below_floor` as `fit:` (exit 3).
  - `fix_deck.py` without `--out` / `--in-place` / `--dry-run` (or with no file) prints a one-line usage hint and
    exits 2; the call is documented in `SKILL.md` and `BUILDER.md`.
  - Deck-level `"sources"` fills the notes of every content slide without its own (`["the brief"]` is fine), so
    `figure_without_source` does not fire when the brief is the only source.
  - Tile labels, notes and trend figures never go below the 24 pt label floor (they used to shrink silently);
    what does not fit is reported; lint `tile_text_below_floor`. The decision box is measured once, so an
    executive summary no longer reserves room for a second line it does not need.
  - Notes `assumptions` (written as `ASSUMPTIONS:`, shown by `--plan`); `CONTENT.md` points the no-invented-facts
    rule at it and adds *Success has a number*.
- Schema regenerated; `--plan` shows assumptions, the deck's sources and spec warnings; `BUILDER.md`,
  `DESIGN.md`, `CONTENT.md`, `AUDIT.md`, `SKILL.md`, the sample spec and `sample-deck.png` updated; self-test
  136 checks.
- **A designed visual system in `build_deck.py`** (a blind benchmark ranked the skill first overall but last on
  design: "grey boxes everywhere, no footers or page numbers, sparse closers"). Every content slide now has a
  kicker label above the title (`kicker`, default from the section's eyebrow; `"kickers": false` deck-wide), a
  footer and page number (`footer`, `page_numbers` deck fields; not on title or section slides). Cards share one
  language in all 20 directions: a tint of the theme with an accent edge, the highlight filled with the accent;
  KPI values and targets in the accent. Covers and section dividers are full panels in the text colour.
- **Decisions and closers:** `decision` (+ `decision_label`) on `kpi` (an executive summary), `statement` (the
  designed close, with `owner` and `date`) and the new `next_steps`.
- **New patterns:** `metrics` (metric, baseline, target in the accent, optional owner and date, a stop/go `rule`)
  and `next_steps` (numbered actions with a date and an owner). The sample spec shows all 21 patterns.
- **Trends in tiles:** `trend` on a `kpi` / `kpi_chart` metric draws its series as small bars, so a before/after
  reads as a trend.
- **Wasted space:** short statements are larger (up to 88 pt) and sit in the optical middle; five-step processes
  are numbered full-width rows; step cards are as tall as their text, the text centred.
- **No unwanted wraps:** numbers and units joined by a no-break space (also in body text: `41 %`); values, chips
  and dates measured with a margin for wide glyphs (`kMeasure.SafeWidth`; a Unicode minus no longer wraps);
  risk chips as wide as their longest word; titles and decisions balanced so no lone word sits on the last line
  (`kMeasure.LineWords`). `kMeasure.Wrap` no longer breaks at a no-break space.
- **`lint_deck.py`:** new `unwanted_wrap` (a short single-line role - value, chip, label, number - that wraps in
  its box) and `title_widow` warnings; `weak_focal_hierarchy` no longer fires on the builder's own layouts; the
  footer, page number and kicker do not count toward the word budget; `metrics`, `next_steps` and decision slides
  get their own budgets.
- **`SKILL.md`:** showcase-first says what to do on an unattended run (build the opener and one detail slide,
  look at the renders yourself, proceed, say so in the summary).
- Schema, `--plan` (shows kickers, the footer and decisions), `BUILDER.md`, `DESIGN.md` and the self-test
  (113 checks) cover all of it. Existing specs build unchanged apart from the new look.

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
