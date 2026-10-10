# Audience, titles, load and speaker notes

*Runs on: any OS.* Part of the `building-powerpoint-decks` skill — start at [`SKILL.md`](../SKILL.md). Paths like `scripts/…` are relative to the skill folder.

## Audience, titles and load

### Pick the audience before the first slide

*Adapted from the personas approach in [Impeccable](https://impeccable.style/) (Apache-2.0).*

When the brief doesn't say who the deck is for, ask — or write down your assumption and show it to the user. One
line per field is enough:

- **Who** — role and what they already know about the topic (beginner / practitioner / expert)
- **Goal** — what they should be able to do or decide after the talk
- **Setting** — live in a room (projector distance), video call (small screen, compression), or read alone later
  (the slides must carry more, the notes become the handout)
- **Accessibility** — low vision, colour blindness, captions needed, non-native language
- **Tolerance** — how dense and how much motion this audience accepts

Every later choice — word budget, font floor (raise it to 24 pt for big rooms), motion, how much goes in notes —
follows from these. "Is 27 pt enough?" has no answer in the abstract; "is 27 pt on a Full HD slide readable from the
back of a 200-seat hall?" does.

### Titles make a claim, not a topic

A topic title ("Decisions", "Cloud costs") gives the audience nothing to agree or disagree with. A claim title
("The check comes first", "East leads Q1, up 8 %") *is* the slide; the body proves it. Titles still count toward
the word budget.

Quick test: a title of **1–2 words with no verb** is almost always a label. Exempt by design: section dividers,
agenda, Q&A and closing slides — and questions, which are fine as titles. Chart titles follow the same rule
(`chart_descriptive_title`). `scripts/lint_deck.py` reports such titles as `title_is_label`; it is a prompt to
look, not a failure.

### Cognitive load limits

Upper bounds — going under is fine, going over is a defect to fix or justify:

- **≤ 7 bullets or ≤ 6 tiles per slide.** Past that the audience stops reading and waits for you to summarise.
- **4–5 items per visual group**; the gap *between* groups ≈ 3× the gap *within* a group, or the grouping
  doesn't read.
- **≤ 3 typefaces per deck** — display + body + optional mono. A fourth reads as an accident.
- **Headline ≤ 55 characters**; a two-line headline is a smell unless line 2 is a deliberate subhead.
- **Line length ≤ 75 characters** in any text block.
- **A chapter break every 7–10 slides** in a long deck — the audience needs a place to exhale.
- **One place for sources** (notes or a consistent footer), never scattered across slides.
- **Lists longer than the limit → reveal one item at a time** (see *Animation*), or split the slide.

**Five-second test:** show the slide for five seconds, hide it, ask what was remembered. If it isn't the title's
claim, the hierarchy is wrong.

---

## Speaker notes — load them up

**Rule:** Speaker notes are unbounded and parallel to the slide. Put **everything that doesn't fit the 10-word visible budget** here. Notes are the deep version of the slide.

**What goes in notes:**
- Numbered sources / citations with author, year, journal
- Anticipated audience questions and your answers
- Full statistical context: confidence intervals, methodology caveats, sample sizes
- Alternative framings the presenter can switch to mid-talk
- "If asked about X, say Y" expansions
- Cross-references to other slides in the deck
- Pacing notes: "spend ~15 sec here", "skip if running long"

Most of this belongs in the short presenter-only reference after the spoken script (see *Write the notes first*
below); the script itself stays 2-5 sentences the presenter says.

**Pattern:** the slide carries the punch; the notes carry the depth. A presenter should be able to deliver a 90-second talk from the slide alone, *and* a 10-minute deep dive from the notes alone, on the same content.

### Every slide gets notes; never invent facts beyond the brief

**Rule:** every slide — the cover and the close included — has speaker notes, and nothing on a slide or in its
notes goes beyond what the brief, the data or a named source says. Anything you add to make the story work (a
rating, a target, a start date, a mitigation, a process step, a time) is listed in the notes' **`assumptions`**
(the builder writes them as `Assumed (confirm before presenting):` in the presenter-only reference after the spoken script;
`--plan` shows them per slide) — so the
presenter knows what to confirm before standing up. When the brief is the only source, say so: `"sources":
["the brief"]` in a slide's notes, or once at deck level (`"sources"` beside `footer`), which fills every content
slide that names none and satisfies `figure_without_source`.

**Why:** in a blind benchmark (October 2026) the judge's notes on all three decks that lost points were the same:
slides with one-word or no notes, and confident details nobody had given ("15 minutes", "ext. 4400", "root cause
fixed", "multi-region failover"). Invented specifics are the fastest way to lose a room that knows the real answer.

**How to apply:** write the notes first (next section). An `assumptions` entry is a short sentence that names
what was assumed ("Likelihood and impact ratings are the team's proposal, not in the brief."); `pitfalls` stay
for what not to say. `build_deck.py` prints `notes: slide N … has no speaker
notes` for every slide the spec left without, and `lint_deck.py` reports `missing_notes`. A worked example
(a training email, a sample price) is labelled as made up in the notes, and placeholder contact details are
written as "give the real number here", never as a plausible fake.

### The ask states its reasons and its cost

**Rule:** the slide that asks for a decision says three things: **what** is asked (a verb and a number: "approve
3 sales engineers"), **why** (two or three reasons from the deck — the risk it answers, the evidence) and **what it
costs or changes** (money, headcount, time, or the impact to expect, and how it will be checked). If the brief gives
no cost, say what the cost consists of and that the figure comes from Finance — do not make one up.

**Why:** an ask with no reasons or cost ("Approve 3 extra sales engineers today") reads as a slogan; the judge
scored it as a thin ask. Decision-makers approve what they can weigh.

**How to apply:** build the ask as a `compare` (*Why now* / *Cost and impact*) or `kpi` slide rather than a bare
`statement`; repeat the decision on the closing slide, with a **`figure`** beside the box — the number the
decision moves (new customers and their trend) or what it costs (`200 kSEK`) — so the close is not text only. Put the expected questions ("What does it cost?", "What if
it fails?") in the notes' Q&A.

### A summary states conclusions; the detail slides carry the charts

**Rule:** chart each data series once. A summary or overview slide states its conclusions in words or KPI figures
("Revenue +44 %", "Churn 1.8 → 1.2 %") with no trend; the detail slide that follows carries the chart.

**Why:** benchmark round 8 marked down a QBR whose summary slide charted revenue and churn and whose slides 3 and 5
then charted the same numbers again - the reader saw the same data twice and the summary added nothing.

**How to apply:** on a summary `kpi`, give metrics a `value`, `label` and `note` but no `trend` when a later slide
charts that series. `build_deck.py` (and `--check` / `--plan`) warns when two slides chart the same numbers.
A `trend_labels` pair such as `["Q1", "Q4"]` over four values is expanded to Q1-Q4; for any other labels give one
per value so every bar is named.

### Success has a number

**Rule:** every success target is measurable — a number, a percent, a date or a comparison with a reference:
"below 30 %", "≥ 95 % of the Q4 level", "≤ the current rate", "by 30 June". Never "Clearly lower", "No drop",
"Improved". The same goes for a timeline: monthly milestones cover every month, or the notes say why one is
skipped; and a worked example (an email with its red flags, a quiz) shows as many cases as the slide promises.

**Why:** in a blind benchmark (round 3) the judge marked targets such as "Clearly lower" and "No drop" as vague,
and a January–June timeline that skipped April as a gap — the board cannot judge a pilot against words.

**How to apply:** if the brief gives no target, propose one and list it in the notes' `assumptions` for the
owner to confirm. `build_deck.py` prints `spec warning: … target … is not measurable` and `… the months jump from
Mar to May`; `lint_deck.py` reports `target_not_measurable` on a metrics table. A statement with nothing but a
support line, or a big number with one caption, also gets a `spec warning`: give it `points` (2–3 short facts,
steps or reasons).

### Compute derived figures, never by hand

**Rule:** a total, an average, a change in percent or a share is computed — by the builder or a script — from the
numbers it is derived from, never added up in your head. In a spec, write `{sum}`, `{average}`, `{change}`,
`{share}` (or `{sum:revenue_q}` against a deck-level `facts` list) and the builder fills in the figure
([BUILDER.md](BUILDER.md#computed-figures)). Put the brief's raw numbers in `facts` and every derived figure
comes from them.

**Why:** in a blind benchmark (round 4) the best-scoring deck said 2026 revenue was "21.8 MUSD" while its own
chart showed quarters of 4.1, 4.6, 5.2 and 5.9 — 19.8. Its notes even said "sum of quarters". One wrong number
costs the whole deck its credibility with a board that can add.

**How to apply:** the builder checks every hand-written figure that reads as derived (next to *sum*, *total*,
*average*, *up 44 %*, `+44 %`, `70 % of …`, an equation in the notes) against the chart series, metric trends,
cost and table columns and `facts` in the same deck, and prints `spec warning: figure: …` when one is close but
wrong; `lint_deck.py` reports `figure_mismatch` on the built deck. The check catches slips — it is not a reason to
compute by hand.

### Write the notes first: a spoken script, then a short reference

Write a slide's notes **before** its visible text — the notes hold the full argument, the slide is the compression.
Every slide's notes have two parts, in this order:

1. **The script (`say`)** — 2-5 natural sentences the presenter actually says. Lead with the slide's point, then
   the why. Work each figure in with its source in the sentence: *"Burnout reached 41 % in the spring survey of all
   40 engineers, according to HR's pulse report."* Write for the ear: short sentences, no labels, no bullet
   fragments, no "KEY FACT:".
2. **The reference (presenter only, kept brief)** — `assumptions` (what you added beyond the brief, to confirm),
   `pitfalls` (what not to say), `sources`, `qa`. The builder writes these after a divider line,
   `--- For the presenter, not to be read out ---`, as one short labelled line each (`Assumed (confirm before presenting):`,
   `Sources:`). Use `facts` only for figures the presenter may need to look up; with `say` they appear as
   `Figures:`.

```json
"notes": {"say": ["Four days a week is a cheap way to cut a 41 % burnout rate, and we can test it in six months.",
                  "HR's spring pulse survey put burnout at 41 %, the highest since we started measuring."],
          "assumptions": ["The pilot covers the 40 engineers only"], "sources": ["HR pulse survey, May 2026"]}
```

**Why:** in benchmark round 8 (October 2026) the judge scored this skill's notes 6.0 against 9.0 a round earlier:
`KEY FACT: 41 % burnout.` / `FACTS: - Burnout survey: 41 %` reads as a form to fill, not words to say. An old
spec with only `key_fact` / `facts` still builds - they become the script's sentences - but a fragment such as
"41 % burnout." is not a script: lint reports `notes_no_script` (fewer than 12 words of prose before the
divider) and `--plan` marks the slide `(no spoken script yet: write notes.say)`. Honesty is unchanged: every
invented or assumed item still goes in `assumptions`, and a made-up example is said to be made up in the script.

**API:**
```python
slide.NotesPage.Shapes(2).TextFrame.TextRange.Text = notes_text
```

(Shape index 2 is the notes placeholder; shape 1 is the slide thumbnail in the notes view.)

### Idempotent notes appending — use the DOI, not author names

**Rule:** When a `notes_*.py` script appends a citation block to existing notes, use the **DOI** (or another guaranteed-unique substring) as the "already present" marker. **Don't** use author-name-and-year strings like `"Smith & Jones (2024)"` — author formatting drifts between the marker text and the actual citation line (`"Smith A., Jones B. (2024)"`, `"Smith et al. 2024"`, etc.), so the marker check fails and the block gets appended again on every re-run.

```python
# BAD — marker doesn't match what's actually in the appended note
CITATION_MARKER = "Smith & Jones (2024)"
APPEND = u"...Primary source: Smith A., Jones B. (2024) Nature Communications, doi:10.1038/..."

# GOOD — DOI is byte-identical between marker and note
CITATION_MARKER = "doi:10.1038/s41467-024-12345-6"
```

DOIs are unique per paper, never abbreviated, and case-stable. For papers without a DOI, fall back to a long quoted phrase (≥ 30 chars) that you literally copied from the appended text.

---

## Showcase-first for multi-slide sections

For any section of **4+ slides** that share a design grammar, build the **opener and one detail slide first**, snapshot both, and ask the user to validate before producing the rest. This is the single highest-ROI rule against rework.

**Procedure:**
1. Build slide 1 (opener) and slide 2 (or last detail) fully — real text, real images, real charts
2. Export both as PNG previews
3. Ask the user: "These two define the design grammar. Approve and I'll batch the rest, or tell me what to change."
4. Wait for explicit sign-off before producing the remaining slides

With the builder, `build_deck.py spec.json --out showcase.pptx --slides 1,3` builds just those two slides (same
theme, page numbers and kickers as in the full deck) from the spec you are still writing; render them with
`render_lo.py showcase.pptx --out renders/ --sheet` and LOOK.

**Why this matters:** iterating slide 1 of 6 is cheap. Iterating slide 5 of 6 after building all six in the wrong grammar is brutal. A single round-trip on two showcase slides saves hours of rebuild.

---
