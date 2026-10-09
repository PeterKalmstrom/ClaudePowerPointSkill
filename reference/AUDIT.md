# Auditing a deck

*Runs on: Windows + PowerPoint for the scripts; the catalogue is any OS.* Part of the `building-powerpoint-decks` skill — start at [`SKILL.md`](../SKILL.md). Paths like `scripts/…` are relative to the skill folder.

## Lint first — `scripts/lint_deck.py` (any OS)

Run it before anything else: it reads the .pptx itself, so it needs no PowerPoint and works on Linux,
macOS and CI.

```bash
uvx --with python-pptx --with pillow python scripts/lint_deck.py deck.pptx            # table
uvx --with python-pptx --with pillow python scripts/lint_deck.py deck.pptx --json     # for scripts
uvx --with python-pptx --with pillow python scripts/lint_deck.py deck.pptx --room-depth 45
uvx --with python-pptx --with pillow python scripts/lint_deck.py deck.pptx --fix --out fixed.pptx
```

| Code | Severity | Fires when |
|---|---|---|
| `slide_size` | warn (info if 16:9) | Deck isn't 1440 × 810 pt |
| `missing_title` / `empty_title` | warn | No title placeholder / it's empty (screen readers and the outline need one, even hidden) |
| `headline_too_long` / `headline_two_line` | warn | Title takes more than two lines at its size in its box (three for display titles ≥ 40 pt on a 960-pt slide), measured from the font; > 55 characters when its size is unknown / has a hard line break |
| `kicker_title_overlap` | warn | The kicker (small label above the title) ends below the top of the title's ink, measured with the loosest renderer's line height and line count (LibreOffice wraps a missing theme font in a wider substitute) |
| `title_widow` / `unwanted_wrap` | warn | A title leaves one short word alone on its last line / a short single-line role (value, chip, number) wraps in its box |
| `title_is_label` | info | 1–2 words, no verb, not a question — a topic, not a claim |
| `body_below_floor` | warn | A paragraph of 4+ words between the label ceiling and the floor. Both scale with slide width: 12 / 18 pt on a 960-pt slide, **18 / 27 pt on Full HD**; `--room-depth` raises the floor |
| `too_many_bullets` | warn | More than 7 paragraphs in one shape |
| `word_budget` | info | More than 12 visible words — or, on a slide `build_deck.py` made, more than its pattern's budget (bullets 30, compare 32, process 36, email 130, risks 64…; `--budget` scales them). Words are tokens with a letter or digit; in `build_deck.py --check` a built slide is counted from its spec, by the same function as `--plan` — accept for gallery, matrix, chart, quote, reference slides |
| `text_overflow` | warn / error > 1.5× | Text needs more height than its box (estimated from real font metrics, measured against LibreOffice to the line), or a text box that starts on a card or frame (a tile, a risk card, a mocked email) runs past its bottom — as drawn or once grown to its text |
| `tile_text_below_floor` | warn | A label, note or trend figure in a `build_deck.py` metric tile below the label floor (16 pt on a 960-pt slide, 24 pt on Full HD) |
| `figure_mismatch` | warn | A total, average, change or share in the text (or an equation in the notes) that the deck's own chart series and table columns contradict - close but wrong, e.g. "21.8, sum of four quarters" beside 4.1 + 4.6 + 5.2 + 5.9 = 19.8; compute it with a `build_deck.py` token |
| `target_not_measurable` | warn | A target in a `build_deck.py` metrics table without a number that has a unit or a comparator for a named metric ("Clearly lower", "No drop"); a time window alone does not count ("<= last 6 months", "by Q3") |
| `ask_without_cost` | warn | A slide with a decision box that nowhere says what the ask costs - money, time, FTE or the word "cost" ([CONTENT.md](CONTENT.md#the-ask-states-its-reasons-and-its-cost)); `build_deck.py` warns on the spec too |
| `text_shrinks` | info | The box is set to shrink text on overflow — check the shrunk size stays above the floor |
| `word_breaks` | warn | A single word is wider than its box and will break mid-word (on Windows, `check_word_breaks.py` gives the exact answer) |
| `offslide_shape` | warn | A shape sticks out past the slide edge (2 pt tolerance) |
| `shape_overlap` | warn ≥ 4 pt², error ≥ 200 pt² | Two text/picture/chart shapes partly overlap (a box fully inside another is a card, not a defect) |
| `unused_placeholder` | error | Empty title/body placeholder next to real content — `--fix` deletes it |
| `palette_too_many_colours` | warn | More than 5 distinct solid fill colours on a slide |
| `a11y_low_text_contrast` | warn / error < 3:1 | Text below 4.5:1 against its shape fill or the slide background (3:1 for large text: 18 pt, or 14 pt bold). Theme colours and their lumMod/tint adjustments are resolved; text over pictures or gradients is skipped |
| `a11y_missing_alt_text` | warn | Picture, chart or table with no alt text, or just a file name or default name ("Picture 3"), and not marked decorative |
| `picture_stretched` | warn > 3 %, error > 15 % | Shown aspect ratio differs from the (cropped) source image |
| `chart_default_palette` | warn | 2+ of the first 6 series in Office default colours |
| `chart_descriptive_title` | info | Chart title has no finding in it (no *leads, rises, falls, vs, %*…) |
| `chart_redundant_labels` | info | Data labels plus a visible, gridlined value axis |
| `chart_legend_steals_plot` | warn | Legend at top or bottom, taking height from the plot |
| `chart_ordinal_categorical_color` | info | Ordered series (months, quarters, years) in unrelated hues — use one hue, light to dark |
| `chart_accounting_zero_dash` | warn | Accounting number format on the axis (zero shows as `$-`) |
| `chart_label_collision` | info | Data labels on more than one series or more than 12 points — they will likely collide |
| `emoji_as_icon` | warn | A short label (≤ 24 characters) containing an emoji |
| `lorem_ipsum` / `truncated_text` | error / warn | Placeholder Latin left in / text ending in an ellipsis |
| `centered_long_body` | warn | Centred paragraph of 80+ characters |
| `measure_too_wide` | warn | Body text of 90+ characters at ≤ 28 pt with more than ~75 characters per line |
| `tiny_click_target` | warn | A clickable shape under 32 pt in either direction (aim for 44 × 44) |
| `shadow_overuse` | warn | More than 3 shapes with drop shadows on one slide |
| `off_palette_fill` | warn | 3+ shapes with hard-coded colours that aren't theme colours |
| `accent_overload` | warn | 4+ different accent colours on one slide |
| `gradient_high_chroma` | warn | A gradient with a saturated (> 0.45) hard-coded stop |
| `repeated_word` | warn | The same word 3+ times in display (≥ 30 pt on a 960-pt slide, 45 pt on Full HD) type |
| `weak_focal_hierarchy` | info | The two largest text sizes are within 1.08–1.6× of each other — nothing clearly leads |
| `grid_monotony` | info | 4+ identical boxes in a row (not when one differs in fill, weight or size - a highlight - nor on a built timeline) |
| `stock_or_cartoon_image` | info | Picture name, alt text or link points at a stock or generic-illustration site |
| `missing_notes` | info | No speaker notes |
| `figure_without_source` | info | The slide shows numbers (%, currency, a chart or table figures) but the notes name no source (not on the cover or a section divider). `SOURCES: the brief` counts; `build_deck.py` writes a deck-level `sources` into every slide without its own |
| `duplicate_titles` / `mixed_font_families` | warn | Deck-wide: repeated titles / more than 3 fonts set directly on text |
| `default_font_only` | info | Deck-wide: one default face (Calibri, Aptos, Arial, Inter…) for everything |

**What it cannot see exactly:** line breaks and overflow are *estimates* (font metrics, or metric-compatible
stand-ins when the real font isn't installed); it can't judge what is behind text on a picture or gradient, and
it can't tell how the slide *looks*. Those need a render — `render_slides.py` (Windows,
exact) or `render_lo.py` (any OS, approximate) — and your eyes. An `info` finding is a prompt to check,
not a defect. Thresholds follow the skill's rules; the overlap areas match the ones the author's PowerPoint
add-in uses.

## Fixing what lint finds — `scripts/fix_deck.py` (any OS)

Mechanical defects have mechanical fixes. `fix_deck.py` applies only the safe ones, lists each change, writes a
copy — or, with `--in-place`, the input itself after keeping the original as `deck.pptx.bak` — and lints the
result, so what's left is what needs judgement:

```bash
uvx --with python-pptx --with pillow python scripts/fix_deck.py deck.pptx --dry-run            # what it would do
uvx --with python-pptx --with pillow python scripts/fix_deck.py deck.pptx --out fixed.pptx
uvx --with python-pptx --with pillow python scripts/fix_deck.py deck.pptx --out fixed.pptx --only floor,alt
uvx --with python-pptx --with pillow python scripts/fix_deck.py deck.pptx --in-place                # keeps deck.pptx.bak
```

| Fix | Repairs |
|---|---|
| `placeholder` | Deletes empty placeholders next to real content |
| `floor` | Raises sentence text below the body floor to the floor |
| `fit` | Shrinks overflowing text in fixed-size boxes — never below the floor |
| `aspect` | Restores a stretched picture's proportions inside the same box (nothing cut off); `--crop-photos` crops to fill instead |
| `alt` | Writes alt text for charts (from their data) and tables (from their header) |
| `palette` | Recolours Office-default chart series as shades of the theme accent |
| `legend` | Moves a top/bottom legend to the right |
| `numfmt` | Replaces an Accounting axis format (`$-` for zero) |

It never rewrites words, titles or picture descriptions and never changes a design choice — those stay in the
"what's left" list for you (or Claude, looking at the render) to decide. Render the fixed copy and look at it:
raising text to the floor can push a slide into overflow, which the follow-up lint will show.

## Reading a whole deck — `scripts/read_deck.py` (any OS)

For analysing a deck as a whole (critiquing content, building a relationship map, counting words, finding
contradictions), one screenshot per slide is far too slow. `read_deck.py` reads the file once and returns every
slide's id, layout, title, shapes (kind, name, position in pt, text, font sizes, alt text), chart data, table
cells and speaker notes:

```bash
uvx --with python-pptx python scripts/read_deck.py deck.pptx --pretty --out dump.json
uvx --with python-pptx python scripts/read_deck.py deck.pptx --text          # outline: titles, text, notes
```

**When to use it vs a screenshot:** `read_deck.py` for anything about 5+ slides or cross-slide reasoning; a render
or `scripts/render_slides.py` when you need to *see* one slide. Then run `lint_deck.py` (above) for the automatic checks and
`check_word_breaks.py` on Windows for exact line breaks.

---

## Common defects to self-check

Before declaring a slide done, scan against these codes. (Several codes and the two-cadence rule are adapted from [Impeccable](https://impeccable.style/), Apache-2.0 — see [NOTICE](../NOTICE).) Each finding gets a severity:

- **`error`** — broken; must fix before shipping (text off-slide, unreadable contrast)
- **`warn`** — likely problem; should fix
- **`taste`** — works, but reads as machine-made, generic or off-brand. Acceptable to ship **only with the user's
  explicit consent** ("yes, the purple gradient is our brand").

**Two cadences.** Check only the objectively broken (`error`, contrast, overflow, broken images) after every edit;
run the full catalogue once, when the deck is done. Running every taste rule on every edit produces noise that
gets dismissed wholesale.

**Slop patterns**

| Code | What to check |
|---|---|
| `slop_gradient` | Multi-stop gradient outside the brand palette (purple→pink, blue→purple)? Use a solid accent or a 2-stop brand gradient. |
| `slop_emoji_icon` | Emoji used as icons (📊 🚀)? Use a real icon set with a text label. |
| `slop_generic_card` | Rounded card + coloured left border + drop shadow + grey text? Use a top strip, a full-bleed header, or spacing alone. |
| `slop_default_font` | Only Aptos / Calibri / Inter / Arial? Pair with a distinctive display face. |
| `slop_stock_imagery` | Handshakes, team-around-a-laptop, overhead desk shots? Use real photos or none. |
| `slop_svg_cartoon_people` | Flat faceless figures in one accent colour? Use photography, abstract shapes or brand illustration. |
| `slop_lorem_ipsum` | Placeholder text left in? Real copy or one labelled TODO. |
| `slop_drop_shadow_overuse` | Shadow on every surface? Reserve elevation for one or two. |
| `italic_serif_display` | Big italic serif headline for "premium" with no brand reason? Use the brand display face. |
| `nested_cards` | Card inside a card? Flatten; use spacing and a heading. |

**Brand and palette**

| Code | What to check |
|---|---|
| `brand_missing_logo` | Branded deck without the real logo file (or a text imitation of it)? Add the real asset. |
| `brand_off_palette` | Colours not in the brand palette or theme accents? Snap to theme colours. |
| `palette_too_many_colours` | More hues than items, or colours that mean nothing? ≤ 3 hues, tied to meaning. |
| `low_contrast_palette` | Two near-identical accents on one surface? One accent per surface. |
| `gray_on_color` | Grey text on a coloured panel? Use a tint of the panel's hue, or white/ink. |

**Composition and rhythm**

| Code | What to check |
|---|---|
| `monotone_anchor` | Three consecutive slides share the same anchor type? Insert variety or a divider. |
| `density_imbalance` | One slide packed, neighbours sparse? Rebalance. |
| `centered_long_body` | Long body copy (> 2 lines) centred? Left-align body; centre only hero/quote. |
| `inconsistent_spacing` | Mixing 8/12/16pt gaps? Pick one unit. |

**Typography**

| Code | What to check |
|---|---|
| `body_below_floor` | Sentence text below the floor — 27 pt on Full HD (18 pt on a 960-pt slide), more for large rooms? Raise it, or make it a label. |
| `headline_two_line` | Headline wraps? Tighten to ≤ 55 characters or split into headline + sub. |
| `mixed_font_families` | More than two families without a brand reason? Display + body only. |
| `attribution_truncated` | Author/source cut off? Widen, shorten, or give it its own line. |

**Content and charts**

| Code | What to check |
|---|---|
| `redundant_word_repeat` | Same key noun 3+ times in big type? Demote the repeats. |
| `chart_descriptive_title` | Chart title names the data ("Revenue by region, Q1") not the finding? Headline the insight ("East leads Q1, up 8 %"). |
| `chart_contrast_fail` | Bar labels below 3:1 against their bar? Dark labels on light bars, light on dark. |
| `chart_label_collision` (warn) | Data labels overlap? Drop per-bar labels or label only the key bar. |
| `chart_redundant_labels` | Data labels *and* a gridlined value axis? Keep one. |
| `chart_legend_steals_plot` | Legend inside the frame shrinking the bars? Move it to a corner or colour-code the title. |
| `chart_default_palette` | Office default series colours (4472C4, ED7D31, A5A5A5…)? Bind series to theme accents. |
| `chart_ordinal_categorical_color` | Ordered series (months, years) in unrelated hues? Use one hue, light → dark. |
| `chart_accounting_zero_dash` | Axis shows `$-` instead of `$0`? Use Currency or Number format on axes. |

**Accessibility (WCAG 2.1 AA)**

| Code | What to check |
|---|---|
| `a11y_low_text_contrast` | Text below 4.5:1 (3:1 for large text/labels) against its backing? Fix colours. |
| `a11y_color_only_signal` | Meaning carried by colour alone (red = bad)? Add a label, icon or pattern. |
| `a11y_missing_alt_text` | Pictures, charts, icons without alt text? Set `shape.AlternativeText`. |
| `a11y_motion_overload` | Everything animates, or motion > 5 s without control? One or two beats per slide. |
| `a11y_tiny_target` | Clickable shapes in a self-running/kiosk deck under 44 × 44 pt? Enlarge. |

`scripts/lint_deck.py` catches most of these automatically, on any OS; `scripts/check_word_breaks.py` catches broken words. The rest need eyes — see the procedure below.

---

## Auditing a deck (full procedure)

Use for any deck of 5+ slides, after any bulk edit, and when reviewing a deck someone else (or an earlier session) built.

**Step 1 — Structural audit.** Run `scripts/lint_deck.py` (any OS) and, on Windows, `scripts/check_word_breaks.py`.

**Step 2 — Taste pass on the flagged slides.** Render each `** AUDIT` slide and check it against the codes in
*Common defects to self-check*.

**Step 3 — Separate defects from anchor exceptions.** The audit flags anything over 12 words, but several anchor
types legitimately exceed that:

| Anchor type | Word budget | Excess OK? |
|---|---|---|
| Hero stat (one big number) | 5 | No |
| Hero + sub | 10 | No |
| Comparison pair | 12 | No |
| 3–5 card gallery | 25 | Yes — read in sequence |
| Matrix | unlimited | Yes — labels, not body |
| Knowledge graph / cascade | 40+ | Yes — relationships, not text |
| Animated list (sequential reveal) | unlimited | Yes — one item at a time |
| Quote | length of quote | Yes — read as one phrase |
| Chart | 8 + chart labels | Yes — the chart speaks |
| Reference / agenda / rules | 30+ | Yes — the audience studies it |

If the anchor type allows the excess, record it as accepted and move on; otherwise fix it.

**Step 4 — Fix in batches by type:** font bumps (one script raising anything below the floor — 27 pt on Full HD), trim cuts (taglines,
sub-attributions, decorative dashes), layout reflows (edit and re-run `build_slide_NN.py`). Snapshot first, and re-run
the audit after each batch. **An `OK` status only proves word count and font size** — render every fixed slide
before calling it done; the audit cannot see text floating outside its backing or a mis-grouped picture.

**Step 5 — Final visual pass.** Spot-check the cover, the busiest slide, the sparsest slide and two random ones.
For 15+ slides, also look at a contact sheet — it shows colour drift between sections, density imbalance, 5+
similar layouts in a row and forgotten template slides:

```bash
uvx --with pywin32 --with pillow python scripts/contact_sheet.py --file deck.pptx [--cols 6] [--slide-width 240]  # Windows
python scripts/render_lo.py deck.pptx --out renders/ --sheet          # any OS: renders + renders/contact.png
python scripts/contact_sheet.py --renders renders/ --cols 4           # any OS: a sheet from renders on disk
```

**Reporting back:** lead with the numbers (N slides, M flagged, K accepted as anchor exceptions), then a table of
slide # → fix, then the accepted exceptions with their anchor type, then any slides with thin speaker notes.

---
