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
| `headline_too_long` / `headline_two_line` | warn | Title > 55 characters / has a hard line break |
| `title_is_label` | info | 1–2 words, no verb, not a question — a topic, not a claim |
| `body_below_floor` | warn | A paragraph of 4+ words between the label ceiling and the floor. Both scale with slide width: 12 / 18 pt on a 960-pt slide, **18 / 27 pt on Full HD**; `--room-depth` raises the floor |
| `too_many_bullets` | warn | More than 7 paragraphs in one shape |
| `word_budget` | info | More than 12 visible words — accept for gallery, matrix, chart, quote, reference slides |
| `offslide_shape` | warn | A shape sticks out past the slide edge (2 pt tolerance) |
| `shape_overlap` | warn ≥ 4 pt², error ≥ 200 pt² | Two text/picture/chart shapes partly overlap (a box fully inside another is a card, not a defect) |
| `unused_placeholder` | error | Empty title/body placeholder next to real content — `--fix` deletes it |
| `palette_too_many_colours` | warn | More than 5 distinct solid fill colours on a slide |
| `a11y_low_text_contrast` | warn / error < 3:1 | Text below 4.5:1 against its shape fill or the slide background (3:1 for large text: 18 pt, or 14 pt bold). Theme colours and their lumMod/tint adjustments are resolved; text over pictures or gradients is skipped |
| `a11y_missing_alt_text` | warn | Picture or chart with no alt text, or just a file name, and not marked decorative |
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
| `repeated_word` | warn | The same word 3+ times in large (≥ 24 pt) type |
| `weak_focal_hierarchy` | info | The two largest text sizes are within 1.08–1.6× of each other — nothing clearly leads |
| `grid_monotony` | info | 4+ identical boxes in a row |
| `stock_or_cartoon_image` | info | Picture name, alt text or link points at a stock or generic-illustration site |
| `missing_notes` | info | No speaker notes |
| `duplicate_titles` / `mixed_font_families` | warn | Deck-wide: repeated titles / more than 3 fonts set directly on text |
| `default_font_only` | info | Deck-wide: one default face (Calibri, Aptos, Arial, Inter…) for everything |

**What it cannot see:** real line breaks, text overflowing its box, z-order, what is behind text that sits on a
picture or gradient, and anything about how the slide *looks*. Those need a render — `render_slides.py` (Windows,
exact) or `render_lo.py` (any OS, approximate) — and your eyes. An `info` finding is a prompt to check,
not a defect. Thresholds follow the skill's rules; the overlap areas and 55-character title limit
match the ones PointClaw (the author's PowerPoint add-in) uses.

## Bulk reading and the quick audit

### Bulk reading the deck

For analyzing a deck as a whole (critiquing content, building a relationship map, counting words, finding contradictions), `slide_snapshot` is too slow — one round trip per slide. Use [`scripts/bulk_read.py`](../scripts/bulk_read.py): it opens the presentation once and dumps all text + speaker notes to JSON in a single COM session (~6s for 70 slides vs. minutes one slide at a time).

```bash
uvx --with pywin32 python scripts/bulk_read.py --file "C:/path/to/deck.pptx" [--slides 5-20] --out dump.json --pretty
```

**When to use bulk read vs slide_snapshot:**
- **bulk_read.py**: analyzing ≥5 slides, reviewing whole sections, building any cross-slide reasoning
- **slide_snapshot**: editing or inspecting a single slide with visual reference

### Auditing a deck

For a one-shot check across an entire deck, run [`scripts/audit_deck.py`](../scripts/audit_deck.py). It walks every slide and prints a flat report flagging slides over the word budget or under the 18pt body floor:

```bash
uvx --with pywin32 python scripts/audit_deck.py --file "C:/path/to/deck.pptx"
```

Columns: slide #, word count, min body font size, status, title. Statuses: `OK`, `ok-tight` (within budget but close), `** AUDIT` (needs review), `[skip]` (hidden slide).

Run this before declaring any deck done. The full procedure — taste pass, anchor exceptions, batching fixes, reporting — is in **Auditing a deck (full procedure)** below.

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
| `body_below_floor` | Sentence text below 18 pt (24 pt for large rooms)? Raise it, or make it a label. |
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

`scripts/audit_deck.py` catches `body_below_floor` and word-budget problems automatically; `scripts/check_word_breaks.py` catches broken words. The rest need eyes — see the procedure below.

---

## Auditing a deck (full procedure)

Use for any deck of 5+ slides, after any bulk edit, and when reviewing a deck someone else (or an earlier session) built.

**Step 1 — Structural audit.** Run `scripts/audit_deck.py` and `scripts/check_word_breaks.py`.

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

**Step 4 — Fix in batches by type:** font bumps (one script raising anything below 18pt), trim cuts (taglines,
sub-attributions, decorative dashes), layout reflows (edit and re-run `build_slide_NN.py`). Snapshot first, and re-run
the audit after each batch. **An `OK` status only proves word count and font size** — render every fixed slide
before calling it done; the audit cannot see text floating outside its backing or a mis-grouped picture.

**Step 5 — Final visual pass.** Spot-check the cover, the busiest slide, the sparsest slide and two random ones.
For 15+ slides, also look at a contact sheet — it shows colour drift between sections, density imbalance, 5+
similar layouts in a row and forgotten template slides:

```bash
uvx --with pywin32 --with pillow python scripts/contact_sheet.py --file deck.pptx [--cols 6] [--slide-width 240]
```

**Reporting back:** lead with the numbers (N slides, M flagged, K accepted as anchor exceptions), then a table of
slide # → fix, then the accepted exceptions with their anchor type, then any slides with thin speaker notes.

---
