# Layout, text and rendering

*Runs on: mostly any OS; rendering needs Windows + PowerPoint.* Part of the `building-powerpoint-decks` skill — start at [`SKILL.md`](../SKILL.md). Paths like `scripts/…` are relative to the skill folder.

## Slide size and pictures

### Slide size — set it before inserting anything

PowerPoint measures slides in **points, not pixels** (96 DPI): **Full HD 1920 × 1080 px = 1440 × 810 pt.**

Two ways to get it wrong, both common:

1. **Setting 1920 × 1080 *points*.** That is 2560 × 1440 px — a third bigger than intended.
2. **Letting PowerPoint pick.** `Presentations.Add()` defaults to **960 × 540 pt (720p)**, and
   `Slides.InsertFromFile` then **silently scales every inserted slide down** to fit. Nothing errors; the deck ships
   at 720p with every font a third smaller.

```python
pres = app.Presentations.Add()
pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight = 1440, 810   # BEFORE the first InsertFromFile
assert (pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight) == (1440, 810)
```

Keep the size in one shared constant — never hardcode it per build script. Check any deck (no PowerPoint needed;
exits non-zero, so it works as a build gate):

```bash
uvx --with python-pptx python scripts/check_slide_size.py deck.pptx [--expect 1440x810]
```

**Pixel maths for renders:** on a 1920-px-wide render of a 1440-pt slide, 1 pt = 1.333 px; on a 1920-px render of a
960-pt slide, **1 pt = 2 px**. Get this wrong in a comparison script and every text size looks off by a third.

### Pictures stretch — crop to fill, never pass both sizes blindly

`Shapes.AddPicture(..., w, h)` and python-pptx `add_picture(width=, height=)` **stretch** the image to the box. They
do not crop or keep the aspect ratio. A 3:2 photo in a 1440 × 610 pt box comes out **~57 % too wide**, and it reads
as "a wide crop", not as a defect — nobody notices. People and animals in a stretched frame look fatter than they
are, which also misrepresents the photographer's work.

```bash
uvx --with pillow python scripts/cover_crop.py photo.jpg --box 1440x610 --out photo.crop.jpg
# prints: stretch if inserted as-is: +57% horizontal
```

Then insert the cropped file at the box size. **A centred crop is a default, not a guarantee** — heads and subjects
get cut. Use `--focus-y` / `--focus-x` and look at the render.

### Third-party photos — credit in the speaker notes

For any photo you did not make, put the credit (photographer, source, licence, link) in the slide's speaker notes.
Under a CC-BY / CC-BY-NC licence, **cropping is an adaptation and must be indicated** ("cropped from original").
Keep the licence terms with the image file so the next build can rewrite the credit.

---

## Font sizes scale with the slide

Every point size in this skill (the 18 pt body floor, the 12–14 pt label exception, the type ceilings) was
learned on 960 × 540 pt slides. **On Full HD (1440 × 810 pt) multiply by 1.5**: body floor 27 pt, labels
down to 18 pt. The same 18 pt text is two-thirds as big on screen on the wider slide. `lint_deck.py` and
`build_deck.py` scale automatically; the full scale is in [DESIGN.md](DESIGN.md).

## Building from a template

When the user has a company template or brand deck, build **on it**, not next to it:

1. **Read it:** `uvx --with python-pptx python scripts/extract_theme.py template.potx` lists the theme
   colours (dk1/lt1, accent1–6), the heading and body fonts, and every layout with its placeholders.
   `--markdown` writes a `brand-spec.md` skeleton for the user to complete (logo, imagery, voice).
2. **Pick layouts by their placeholders** (Title Only for a hero stat or chart, Two Content for a
   comparison) and fill the placeholders. Don't put text boxes on Blank slides: you lose the outline,
   screen-reader titles and the template's typography.
3. **Use theme colours, not hex values** — `MSO_THEME_COLOR.ACCENT_1` in python-pptx, `ObjectThemeColor`
   in COM — so a re-theme re-skins the deck. Highlight the finding with accent1; everything else neutral.
4. **Fonts:** leave font names unset so text inherits the theme's major (headings) and minor (body)
   fonts. Precedence, highest first: text run → shape → layout placeholder → master text styles → theme.
   A font set on a run beats the template everywhere, which is how off-brand decks happen.
5. **Lint** — `mixed_font_families` and `chart_default_palette` catch the usual drift.

## Text fitting and rendering

### Text wrap — anticipate it, don't trust visual review

**Rule:** PowerPoint TextBoxes word-wrap by default. A headline at 52pt that "should fit" on one line will silently wrap to two lines if it exceeds the container width — and the wrap is invisible until export. Always plan for the worst-case character count and font width.

**Practical ceilings (Aptos Display, bold, single-line; Full HD 1440-pt slide):**

| Container width | Safe size | Approx. character ceiling |
|---|---|---|
| 810 pt (≈ 56 % of the slide) | 54 pt | ~22 chars |
| 810 pt | 48 pt | ~25 chars |
| 810 pt | 42 pt | ~28 chars |
| 1200 pt (≈ 83 %) | 66 pt | ~25 chars |
| 1200 pt | 57 pt | ~28 chars |
| 1200 pt | 48 pt | ~33 chars |

The ceiling depends on the *ratio* of font size to box width, so on a 960-pt slide divide both by 1.5. Aptos
Display is wider than Inter or Calibri; other fonts shift the ceiling. `lint_deck.py` estimates characters per
line (`measure_too_wide`) and the overflow check estimates wrapped height; neither replaces a render.

**Defenses against silent wrap:**

```python
# 1. Disable wrap — overflow is preferred over silent wrap
text.TextFrame.WordWrap = 0  # msoFalse

# 2. Reduce internal margins (default ~7pt each side)
text.TextFrame.MarginLeft = 4
text.TextFrame.MarginRight = 4

# 3. Always render and look before declaring done — Save As JPEG, which uses
#    embedded fonts (slide.Export does not; see the next section)
presentation.SaveCopyAs(out_folder, 17)
```

**4. Size display fonts so the widest *word* fits.** PowerPoint breaks *inside* a word that is wider than its line
("constrai / n"). For a display font in a narrow box, measure the widest word with the font's real metrics (e.g.
Pillow `ImageFont.getlength`) and pick the size from that, not from the character count.

### Embedded fonts: `Slide.Export` renders a FALLBACK — use Save As JPEG

**Symptom:** a deck whose display font is embedded but not installed renders fine in PowerPoint, but PNGs from
`slide.Export(...)` show the display type in an Arial-like fallback. The fallback is narrower, so a real mid-word
break that the audience will see renders clean in the PNG and passes review. `prs.Fonts(name).Embedded == -1` only
proves the font is in the file, not that the renderer used it. The MCP's `slide_snapshot` uses `Slide.Export` too, so
**don't trust its screenshots for embedded fonts.**

**Fix:** render through PowerPoint's own File → Save As → JPEG, which does use the embedded font:

```python
prs.SaveCopyAs(r"C:\temp\out\_saveas", 17)   # ppSaveAsJPG -> Slide1.JPG .. SlideN.JPG in that folder
# rename SlideN.JPG -> sNNN.jpg so they sort; assert count == prs.Slides.Count
```

[`scripts/render_slides.py`](../scripts/render_slides.py) does exactly this. Use `SaveCopyAs`, not `SaveAs`: same render,
and it never rebinds the open presentation. Speed: ~15 s for 148 slides at 1280×720. **Not** a substitute:
`SaveAs(pdf, 32)` — its PDF used Calibri, not the embedded face.

**Line-fit gate without a picture:** PowerPoint's layout engine *does* use the embedded font, so
`TextRange.Lines(i, 1)` reports the real breaks. Flag any line boundary with a letter on both sides:

```bash
uvx --with pywin32 python scripts/check_word_breaks.py --file deck.pptx   # exit 1 if any word is broken
```

**Pattern:** for any headline ≥30pt, export and visually verify that no text wrapped. Don't trust the COM property values — only the rendered PNG tells you what the audience sees.

---

## Word budgets and layout patterns

### On-slide text budget: ~10 words max (excluding title)

**Rule:** For spoken/live presentations, limit visible text on each slide to roughly 10 words beyond the title. If a slide is on screen for ~20 seconds, the audience can spend at most ~2 seconds reading — the remaining ~18 seconds belong to the speaker. More text than that forces the audience to choose between reading and listening.

**Why:**
- Audiences read silently at ~200–250 words/min, so 10 words ≈ 2 seconds
- Dense slides compete with the speaker's voice — attention splits and retention drops
- Presenters fall into "reading their slides" when the deck carries the content
- Visual impact (image, chart, big number) scales with the amount of whitespace around it

**How to apply:**
- Title: one phrase (not counted in the 10)
- Body: one short headline OR a single big number OR one chart — not all three
- All supporting detail, data, sources, and narrative goes into **speaker notes** (notes can be arbitrarily long — load them up)
- If you need more text, make it a new slide
- Exceptions: see anchor-type budgets below

**Pattern:** Pair a strong visual (full-bleed image or chart overlay) with a single punchy headline ≤10 words. Everything else lives in notes.

### Anchor types and word budgets

The 10-word rule is the default for **spoken hero slides**. It does not apply uniformly — different anchor types have different read patterns. Use the budget that matches the anchor:

| Anchor type | Visible word budget (excl. title) | Why |
|---|---|---|
| Hero stat (single big number) | 5 | One thing to absorb |
| Hero + sub (headline + tagline) | 10 | Statement + context |
| Comparison pair (A vs B) | 12 | Two parallel things |
| 3–5 card gallery | ~25 (8/card) | Audience reads one card at a time |
| Vulnerability matrix / table | unlimited | Cells are labels, not body |
| Knowledge graph / cascade diagram | 40+ | Relationships matter more than label content |
| Animated list (sequential reveal) | unlimited | Each item appears alone |
| Quote anchor | length of quote | Single coherent phrase, read as a unit |
| Chart slide | 8 + chart-internal labels | The chart is the message |
| Reference / agenda / rules | 30+ | Audience studies it; not spoken over |

**How to identify the anchor type:** look at what dominates the slide visually. If three icons sit side by side, it's a gallery. If a single 100pt number, it's hero stat. If a network of connected cards, it's a knowledge graph.

**Audit caution:** automated word-count checks will false-positive on gallery / matrix / knowledge-graph slides. Always confirm against the anchor type before "fixing" by stripping content.

### Label exception (smaller sizes allowed for non-body text)

The body floor — **27 pt on Full HD, 18 pt on a 960-pt slide** — applies to **body text**: sentences the audience
reads. It does not apply to **labels**: short identifying tags attached to a visual.

**Labels allowed down to 18–21 pt on Full HD (12–14 pt on a 960-pt slide):**
- Chart axis labels, tick labels, legend labels
- Caption directly under a flag, icon, or photo
- Row/column labels in a matrix or table
- Pill or chip text in a dense gallery
- Source attributions and citations (≥ 18 pt on Full HD; 27 pt preferred)

**Body text — must be at or above the floor (27 pt on Full HD):**
- Headlines, taglines, sentence fragments
- Bullet items the audience is meant to read
- Quotes
- Anything that appears as a sentence

**Test:** if removing the visual the text describes leaves you with a meaningful sentence, it's body — bump it to the floor.

### Hero stat pattern

Recurring recipe for "single big number" slides (`build_deck.py` pattern `big_number`). Sizes for Full HD:

| Element | Size | Notes |
|---|---|---|
| The number | 120–220 pt bold | Accent colour, dominant |
| Unit (next to or below the number) | ≥ 36 pt bold | Same colour |
| Caption | ≥ 27 pt muted | One line of context |
| Source / attribution | in the speaker notes | Not on the slide |

Anti-pattern: number at 150 pt + unit at 21 pt — the unit vanishes next to the number.

### Two-column comparison pattern

For "humans vs AI", "before vs after", "today vs projected" slides:

```
Full HD (1440 x 810 pt); build_deck.py pattern `compare` does this for you
Layout: two equal-width columns side by side, gap 48 pt
Accent rule above the column that is the answer; neutral rule above the other
Column height: 420-510 pt depending on row count
Column structure:
  - Heading: 42-46 pt bold
  - 2-4 points: 28-30 pt (labels, not sentences, may go to 24 pt)
Bottom strip below both columns: one bold 34-36 pt line with the synthesis
```

**Word count:** ≤6 bullets per card × ~3 words each + 1 synthesis line ≈ 35–40 words across the slide. Acceptable as "comparison pair" anchor (each side read as one unit).

### Balanced layout — fill the available space

**Rule:** Content should fill the slide's usable area with visually balanced margins. Don't let content hug the top and leave a large void at the bottom (or any other side). Scale element size and typography proportionally to available space.

**Why:**
- Unbalanced layouts look unfinished — the audience reads "draft" or "amateur" before they read the content
- Top-heavy slides leave the visual center of gravity wrong; the eye keeps looking for something that isn't there
- A slide element sized for a cramped layout looks small and weak when the canvas around it is actually large
- Matching gaps (top/bottom/between) reads as intentional; uneven gaps read as mistakes
- Bigger elements let you use bigger, more impactful typography

**How to apply — sizing:**
1. Measure the usable area: slide height minus title bar (e.g., 810 − 135 = 675 tall content area on Full HD)
2. Count the rows/elements and the gap count (for N rows there are N+1 gaps — one above row 1, one below row N, and N−1 between)
3. Decide gap size (typically 24–36 pt on Full HD) and divide the remaining height among the elements
4. Apply the same formula horizontally for multi-column layouts

**How to apply — typography:**
- When elements grow, grow the font too. A 180-pt-tall row can carry 50 pt text; a 110-pt row only 40 pt.
- Titles: 54–80 pt depending on length
- Body labels on cards: 42–56 pt for impact
- Secondary labels / captions: 21–30 pt
- Body text in content slides: 27–36 pt (the 2-second rule still applies)
- (On a 960-pt slide, divide every number by 1.5.)

**Pattern — three-row, two-column grid (generic example):**
```
Slide: 1440 × 810 (Full HD), title panel 0–135
Content area: y=135 to y=810 = 675 tall
Target: 3 rows × 2 columns of cards, icon + label

Rows: 3 rows of height R, 4 gaps of height G
    3R + 4G = 675
    Pick G = 33, R = 180 → 540 + 132 = 672 (3 pt slack)

Columns: margin M, icon I, gap g, label L, inner gap IG, then repeat
    2M + 2I + 2g + 2L + IG = 1440
    M=30, I=150, g=15, L=495, IG=60 → 60 + 300 + 30 + 990 + 60 = 1440 ✓

Row y positions: 168, 381, 594
Col 1 icon x=30, label x=195
Col 2 icon x=750, label x=915
```

**Sanity-check after any layout change:**
- Top gap above first row ≈ bottom gap below last row (within 15–20 pt)
- Inter-row gaps are equal
- No large void in any quadrant
- Elements proportional to their container (a 180-pt card with 21 pt text looks empty)
- Render the slide and look at it — if you spot a visual imbalance, fix it

**When to break the rule:**
- Deliberate tension or asymmetry (a single big number in one quadrant, whitespace elsewhere for emphasis)
- Reference slides that follow a fixed template across a series
- Full-bleed images where the composition is the layout
