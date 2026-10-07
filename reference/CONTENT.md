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

**Pattern:** the slide carries the punch; the notes carry the depth. A presenter should be able to deliver a 90-second talk from the slide alone, *and* a 10-minute deep dive from the notes alone, on the same content.

### Write the notes first, in a fixed order

Write a slide's notes **before** its visible text — the notes hold the full argument, the slide is the compression.
Use the same order on every slide so the presenter always knows where to look:

1. **Key fact** — the one sentence the slide exists to land
2. **Facts** — supporting points, numbers and context to browse (no word cap)
3. **Q&A** — likely questions with short answers
4. **Pitfalls** — what is commonly misunderstood, or what not to say
5. **Sources** — citations with DOI or URL

Write facts to browse, not a script to read aloud. A fixed structure also lets a script check notes coverage
(every content slide has a key fact and at least one source) and feeds the Q&A panic sheet in *Presenter prep*.

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

**Why this matters:** iterating slide 1 of 6 is cheap. Iterating slide 5 of 6 after building all six in the wrong grammar is brutal. A single round-trip on two showcase slides saves hours of rebuild.

---
