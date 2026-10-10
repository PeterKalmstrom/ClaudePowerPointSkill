# Benchmark: this skill vs plain Claude vs Anthropic's pptx skill

Three briefs ([briefs.md](briefs.md)): a quarterly business review with real numbers, a 20-minute phishing training,
a board proposal for a four-day-week pilot. Each maker built all three decks, unattended, in its own folder:

- **Plain Claude** - python-pptx, no skill.
- **Anthropic's pptx skill** (`anthropic-skills:pptx`, pptxgenjs).
- **This skill** - spec -> `build_deck.py` -> `lint_deck.py` / `fix_deck.py` -> LibreOffice renders.

All 27 slides per maker were rendered with LibreOffice and scored **blind** by a separate Claude judge that saw
only anonymous labels (keys: `round*-key.json`), on six criteria from 1 to 10: message, visual design, legibility,
layout correctness, data presentation, presenter support (speaker notes).

## Round 10 - five briefs (two holdout), two runs per maker, two blind judges (current)

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall (range)** | Time / tokens / calls per run |
|---|---|---|---|---|---|---|---|---|
| **This skill** | 8.65 | 8.35 | 8.10 | 8.00 | 8.05 | 8.00 | **8.19** (7.97-8.47) | 395 s / 113k / 22; 482 s / 113k / 23 |
| Anthropic pptx skill | 8.00 | 6.95 | 7.00 | 7.00 | 8.00 | 1.70 | 6.45 (6.37-6.47) | 1473 s* / 106k / 25; 373 s / 97k / 17 |
| Plain Claude | 7.45 | 5.60 | 5.90 | 5.90 | 7.45 | 3.95 | 6.04 (5.40-6.80) | 233 s / 95k / 11; 221 s / 133k / 11 |

Each cell is the mean over 2 runs x 2 judges x 5 decks; the range is over the four run/judge overalls.
\* the first pptx-skill run hung on a LibreOffice render and abandoned it, so its time is not representative.

**What changed in the skill since round 9:** icons and pictures, company-template support, a visual check inside
`--check`, ready edits through `--plan` / `--check --apply`, and tools for existing decks.

**Method changes (stronger than rounds 7-9):**
1. **Holdout briefs.** Two new briefs the skill was never tuned on ([briefs-holdout.md](briefs-holdout.md)): B4 a
   product-launch pitch to sales, B5 a blameless project post-mortem for engineering leadership, both with concrete
   numbers. All five briefs were built by every maker run.
2. **Two independent runs per maker** (six unattended agents, all in parallel, each into `<maker>-<run>`), so
   single-run luck (round 8's 3-call plain run) shows up as spread instead of as a result.
3. All 30 decks rendered sequentially in real PowerPoint (`render_slides.py`, no other PowerPoint job running,
   PowerPoint closed afterwards) into slide JPGs, contact sheets and text+notes files, under six fresh random labels
   (`round10-key.json`), one label per maker *run*, so a judge could not tell two runs of one maker were related.
4. **Two independent blind judges** with the same prompt ([prompts/judge.txt](prompts/judge.txt)), each scoring
   all 30 decks; the key and the other judge's file were outside the judge folder.
5. This skill's `lint_deck.py` and `visual_check.py` (on the PowerPoint renders) run on every deck.

**Spread - do the differences exceed noise?**

| Maker | Run 1 | Run 2 | Judge A | Judge B | Original 3 briefs | Holdout 2 briefs |
|---|---|---|---|---|---|---|
| This skill | 8.22 | 8.17 | 8.35 | 8.03 | 8.33 | 7.98 |
| pptx skill | 6.42 | 6.47 | 6.42 | 6.47 | 6.46 | 6.42 |
| Plain Claude | 5.87 | 6.22 | 5.52 | 6.57 | 6.10 | 5.96 |

- **This skill vs the other two: yes.** Its lowest of four run/judge cells (7.97) is 1.5 above the pptx skill's
  highest (6.47). Notes carry much of the gap, but without notes (mean of the other five criteria) it still leads
  8.23 vs 7.39 vs 6.46, and both judges ranked its two runs first and second.
- **pptx skill vs plain Claude: no.** 0.41 apart, inside plain Claude's own 1.4-point range; the judges disagreed
  most on plain Claude (A 5.52, B 6.57).
- **Run-to-run spread was small** (0.05 for this skill and the pptx skill, 0.35 for plain); **judge-to-judge spread
  was larger** (0.32 for this skill, 1.05 for plain). One judge per round, as in rounds 1-9, is the weaker link.
- **Holdout:** this skill scored 0.35 lower on the two unseen briefs (8.33 -> 7.98), the others 0.04-0.14 lower.
  That gap is about the size of the judge spread, so it is a hint, not proof, of some tuning to the original
  briefs; it still led on both holdout decks (launch 7.96, post-mortem 8.00 vs 6.42 and 5.8-6.1).
- **Compared with round 9** (8.55 / 6.78 / 5.78): different briefs, judges and run counts, so the 0.36 drop for
  this skill is not a regression signal; the order and the size of the lead held.

Judges, in short: this skill - claim titles, highlighted charts, dot graphics, a quiz with an answer slide,
decisions with owner and date, and the only notes a presenter could read aloud; assumptions marked rather than
invented (both runs kept the unquantified staffing delay off the post-mortem chart and said so in the notes).
Weaknesses: on the holdout launch deck both runs invented first-30-day targets (one run's "3 deals each by
31 March" is ~75 deals in a month against a 120-customer year target - judge A); short objection answers; one
post-mortem switched from a light cover to a dark body; a four-month-looking pilot timeline. Best decks: fourday
(judge A) and phishing (judge B). pptx skill - clean, consistent, every brief item and arithmetic right (24 MUSD
run rate, 20 % annual-plan saving), but notes almost absent in both runs, small card text, half-empty lower slide
halves, quiz without answers (one run) and an invented 2-week staffing figure (labelled approximate). Plain
Claude - complete and mostly correct, but plain left-aligned slides, tiny chart labels, a pie chart that repeats a
three-row table, notes that repeat the title, chart labels rendered as "6." and one wrong claim ("roughly halves"
annual churn for a one-third drop).

Both judges said they worked from the contact sheets and text files and opened only ~10 single slides each, despite
the prompt asking for every chart/table slide - the same shortcut as round 9's judge. Fine layout defects can be
under-scored for every maker.

**Lint and visual check** (this skill's own tools, not neutral):

| Maker | lint errors / warnings per deck | visual_check warnings per deck (PowerPoint renders) |
|---|---|---|
| This skill | 0 / 0 on all 10 decks (0-2 info) | 0 on all 10 (0-3 info) |
| pptx skill | 0-2 / 18-30 | 0-4 (`empty_area`) |
| Plain Claude | 0-10 / 17-37 | 1-5 (`empty_area`) |

`visual_check` reported `median_occupancy` 0 for both of this skill's post-mortem decks (light cover, dark body)
while every per-slide occupancy was 0.21-0.38 - a reporting bug in the summary value to look at, not a deck defect.

**Speed:** this skill's agent took 22-23 tool calls for five decks (4.5 per deck) against round 9's 16 for three
(5.3 per deck), 395-482 s and 113k tokens per run - slower than the other makers in absolute terms (plain 11 calls,
~225 s), the cost of building each deck through spec, build, check and render.
Raw data: [round10-scores.json](round10-scores.json) (both judges' per-deck scores, comments and the aggregate).

## Round 9 - spoken-script notes, every category label, duplicate-chart warning

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall** | Time / tokens / calls |
|---|---|---|---|---|---|---|---|---|
| **This skill** | 8.3 | 8.3 | 9.0 | 8.0 | 8.7 | 9.0 | **8.55** | 246 s / 100k / 16 |
| Anthropic pptx skill | 8.3 | 8.0 | 6.3 | 7.7 | 8.3 | 2.0 | 6.78 | 175 s / 86k / 9 |
| Plain Claude | 7.7 | 5.7 | 5.7 | 5.7 | 8.0 | 2.0 | 5.78 | 168 s / 76k / 7 |

**What changed in the skill since round 8:** speaker notes are a spoken script (`notes.say`) plus a short
presenter-only reference; every chart category label shows; a warning flags two slides that chart the same data.
The other two makers did not change, but all three were rebuilt fresh, in parallel, under the same conditions.

**Method:** round 8's method repeated - same briefs, one unattended agent per maker run in parallel, slides rendered
in real PowerPoint (`render_slides.py`) into contact sheets, slide text and notes extracted to text, one separate
blind judge with new labels (P/Q/R, `round9-key.json`) on the six criteria. The exact maker and judge prompts are
saved in `_scratch/bench9/prompts/` for reuse. One difference from round 8: the plain and pptx-skill makers reported
they could not find LibreOffice (not on PATH), so neither looked at its own slides; this skill's maker rendered
through its own `render_lo.py` and did. The judge said it worked mostly from the contact sheets plus the text files,
opening only a few single slides.
Lint (this skill's own, not neutral): this skill 0 errors / 0 warnings on all three decks; plain Claude 7-9 errors
per deck; pptx skill 1-2.

Judge, in short: this skill - full notes on every deck that a presenter can deliver, with assumptions marked; the
largest type; charts that highlight the key value; data visuals (41-in-100 dot grid, 14/20 hiring dots); decisions
with owner and date; a separate quiz-answer slide; all figures right. Weaknesses: a few near-bullet slides (phishing
s3, four-day s3) and no cost figure on the QBR ask. Best deck: phishing. pptx skill - the strongest storytelling and
polish (four-day deck with a cost donut summing to 200 kSEK was its best), but notes almost absent, small card text,
awkward title wraps, quiz answers printed on the quiz slide and an unsourced "9 in 10 attacks". Plain Claude - sound
content and correct extra arithmetic (run-rate, 5 kSEK per engineer), but tiny text in mostly empty cards, a KPI
that wraps in its tile, plain design and almost no notes.

**Compared with round 8.** Notes, the round-8 weakness, went from 6.0 to 9.0: the judge called them "full",
"excellent" and "thorough" instead of round 8's "templated KEY FACT/ASSUMPTIONS blocks rather than a talk track". The
round-8 QBR complaints (chart missing Q2/Q3 labels, slide 2 repeating slides 3 and 5) did not come back. Legibility
rose 8.0 -> 9.0; message and data dipped slightly (9.0 -> 8.3, 9.0 -> 8.7), within judge noise. Overall 8.11 -> 8.55,
the highest since round 6. The pptx skill fell 7.44 -> 6.78, almost all from notes (6.3 -> 2.0: this run wrote notes
on two slides per deck); plain Claude rose 5.17 -> 5.78. This skill's agent was faster than in round 8 (328 s / 21
calls -> 246 s / 16).

## Round 8 - round 7's method, makers can see LibreOffice renders

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall** | Time / tokens / calls |
|---|---|---|---|---|---|---|---|---|
| **This skill** | 9.0 | 8.3 | 8.0 | 8.3 | 9.0 | 6.0 | **8.11** | 328 s / 97k / 21 |
| Anthropic pptx skill | 8.0 | 8.0 | 6.0 | 8.0 | 8.3 | 6.3 | 7.44 | 279 s / 94k / 12 |
| Plain Claude | 6.0 | 4.0 | 4.0 | 5.7 | 8.0 | 3.3 | 5.17 | 68 s / 64k / 3 |

**Method:** same as round 7 - the same briefs, one unattended agent per maker run in parallel, slides rendered in real
PowerPoint (`render_slides.py`) into contact sheets, one separate blind judge with new labels (X/Y/Z,
`round8-key.json`). **One difference:** LibreOffice is now installed on the PC, so every maker could render and look
at its own decks (round 7's makers could not). The judge also read the slide text and notes, as in round 7.
Lint (this skill's own, not neutral): this skill 0 errors / 0 warnings on all three decks; plain Claude 7-12 errors
per deck; pptx skill 1-4.

Judge, in short: this skill - claim titles throughout, decisions with owner and date, every figure right (+44 %,
19.8 MUSD, 191 customers, 200 kSEK); the four-day deck (dot chart of 41 in 100, metrics with a stop rule, cost
table) was the best of the nine. Weaknesses: notes read as fill-in blocks (KEY FACT / ASSUMPTIONS / SOURCES) rather
than something to say; the QBR's slide-2 chart lacks Q2/Q3 labels and repeats slides 3 and 5; empty space on two
phishing slides. pptx skill - modern, polished (cards, donuts, timeline), good flow, quiz answers in the notes; but
small body text, notes only on the four-day title slide, an unsourced growth explanation and a churn slip (7 instead
of 6 per 1,000). Plain Claude - arithmetic right, but topic titles, tiny grey text in mostly empty cards, an
overlapping card heading, bullet slides and almost no notes.

**Compared with round 7.** Design, the round-7 weakness, rose from 7.0 to 8.3 and is now the highest of the three
(pptx skill 7.3 -> 8.0, plain 8.0 -> 4.0). Layout rose 7.3 -> 8.3. Notes fell 9.0 -> 6.0: same notes structure,
but this judge marked it down as a template rather than a script - the top fix for next round. Overall 8.22 -> 8.11,
within the half-point judge noise. Plain Claude's swing (6.78 -> 5.17) on a 3-call, 68 s run shows how much
single-run variance the other makers carry. This skill's agent was slower (212 s / 11 calls -> 328 s / 21), likely
because it now could and did loop on renders.

## Round 7 - after the round-6 rough edges, on Windows

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall** | Time / tokens / calls |
|---|---|---|---|---|---|---|---|---|
| **This skill** | 8.7 | 7.0 | 8.7 | 7.3 | 8.7 | 9.0 | **8.22** | 212 s / 100k / 11 |
| Plain Claude | 9.0 | 8.0 | 6.3 | 6.0 | 8.7 | 2.7 | 6.78 | 162 s / 75k / 3 |
| Anthropic pptx skill | 7.7 | 7.3 | 6.7 | 6.3 | 6.7 | 4.0 | 6.45 | 189 s / 84k / 6 |

**Different method from rounds 1-6, so compare within the round, not with round 6:** the makers ran on the Windows
PC without LibreOffice (no maker could look at its own renders), the slides were rendered in real PowerPoint, and the
judge worked mostly from contact sheets. Lint (this skill's own, so not neutral): this skill 0 errors / 0 warnings on
all three decks; plain Claude 7-19 errors per deck; pptx skill 1-3.
Judge, in short: this skill - notes with facts, assumptions and sources on every slide, the largest type, claim
titles, decisions with owner and date, all figures right; plainer design, a plain bullet slide and staggered stage
labels in the phishing deck, small charts on two QBR slides, a pilot timeline that runs into July (flagged as an
assumption in the notes). Plain Claude - the strongest story and most polished look, figures right, but almost no
notes, small crowded text, colliding risk-card headings and a misplaced email callout. pptx skill - clean and
consistent but a wrong figure (+0.6 for +0.7 MUSD), an unsourced statistic, topic titles, the quiz answers on the
quiz slide, a "Thank you" closer and few notes.

## Round 6 - after the speed round

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall** | Time / tokens / calls |
|---|---|---|---|---|---|---|---|---|
| **This skill** | 9.0 | 8.3 | 9.0 | 8.0 | 9.0 | 8.7 | **8.67** | 131 s / 91k / 13 |
| Anthropic pptx skill | 8.0 | 8.0 | 5.0 | 7.3 | 8.3 | 1.7 | 6.39 | 159 s / 90k / 12 |
| Plain Claude | 7.7 | 6.7 | 4.7 | 6.3 | 8.0 | 1.0 | 5.72 | 124 s / 73k / 9 |

All three makers ran fresh, at the same time, so speed is comparable. The speed round cut this skill's agent from
205 s / 107k / 26 calls (round 5) to 131 s / 91k / 13 - now level with the other two - while quality rose.
Judge, in short: this skill - conclusion first, large type, rated risks, a separate quiz answer slide, a stop rule,
0 lint errors and warnings, structured notes on every slide; minor flaws (callout numbers out of order on the
example email). pptx skill - good-looking, strong titles, data right (interpolated churn points labelled), but
9-11 pt body text and notes on few or no slides. Plain Claude - covers each brief but generic, sparse cards, tiny
type, 7 lint errors in one deck, an unsourced statistic and a timeline past the pilot; almost no notes.

## Round 5 - after the computed-figure and font rounds

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall** | Time / tokens / calls |
|---|---|---|---|---|---|---|---|---|
| **This skill** | 8.7 | 7.7 | 9.0 | 9.0 | 8.3 | 8.7 | **8.56** | 205 s / 107k / 26 |
| Plain Claude | 7.7 | 8.0 | 6.7 | 8.0 | 7.7 | 6.0 | 7.33 | not measured |
| Anthropic pptx skill | 7.0 | 7.0 | 5.7 | 7.7 | 7.3 | 2.0 | 6.11 | not measured |

Judge, in short: this skill - claim titles with **all arithmetic correct** (derived figures now come from the
builder's `{sum}`/`{change}`/`{share}` tokens, not hand sums), notes with key fact / facts / assumptions / sources
on every slide, large text, nothing overflowing, decisions with owner and date; weakest on visual flair (the phishing
deck reads plain) and a few uneven stage labels. Plain Claude - the most polished visuals and the best phishing
teaching slide, but invented facts presented as fact and thin notes. pptx skill - right chart types, but almost no
notes and ~10-11 pt body text.

## Round 4 - after the data-presentation round

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall** | Time / tokens / calls |
|---|---|---|---|---|---|---|---|---|
| **This skill** | 8.7 | 8.3 | 9.0 | 8.7 | 7.7 | 8.7 | **8.50** | 189 s / 100k / 18 |
| Plain Claude | 8.0 | 7.7 | 6.0 | 7.7 | 8.0 | 5.0 | 7.06 | not measured |
| Anthropic pptx skill | 7.3 | 7.0 | 6.0 | 7.3 | 7.3 | 1.7 | 6.11 | not measured |

Judge, in short: this skill - claim titles, every deck ends on a decision with owner and deadline, the largest text
and most consistent visual system, structured notes with every invention flagged. **But one factual error**: the QBR
gave 2026 revenue as 21.8 MUSD where the quarters sum to 19.8 (the maker's spec said "sum of quarters" and added
wrong) - the skill does not yet check computed figures. Plain Claude - the best charts and closing structure, but
small grey text, unflagged inventions and thin notes. pptx skill - real charts with axes, but almost no notes, small
text and the quiz answers printed on the slide.

## Round 3 - after the design round

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall** | Time / tokens / calls |
|---|---|---|---|---|---|---|---|---|
| **This skill** | 8.7 | 7.7 | 9.0 | 7.7 | 6.3 | 9.0 | **8.06** | 173 s / 98k / 18 |
| Plain Claude | 8.7 | 8.0 | 7.0 | 8.0 | 7.7 | 6.7 | 7.67 | not measured |
| Anthropic pptx skill | 8.0 | 7.0 | 5.7 | 7.7 | 7.3 | 1.7 | 6.22 | not measured |

Judge, in short: this skill - the most legible decks (big type, clearest contrast), honest notes on every slide
with every assumption flagged, claim titles throughout; weakest on data presentation (a retention slide without a
chart, tiny trend glyphs, a text-only ask) and a few sparse slides. Plain Claude - best story structure and
polished colour-coded design, but the most invented content, none of it flagged. pptx skill - right chart types and
a clean card system, but almost no speaker notes and small body text.

## Round 2 - after the builder quality round

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall** | Time / tokens / calls |
|---|---|---|---|---|---|---|---|---|
| **This skill** | 8.7 | 6.7 | 7.7 | 8.7 | 7.3 | 8.7 | **7.94** | 179 s / 95k / 19 |
| Plain Claude | 8.3 | 8.7 | 8.0 | 8.3 | 7.7 | 5.3 | 7.72 | not measured |
| Anthropic pptx skill | 8.0 | 8.0 | 7.0 | 8.0 | 7.7 | 2.0 | 6.78 | not measured |

Judge, in short: this skill - best speaker notes by far, claim titles, faithful to the brief with assumptions
flagged, clean data slides; but the plainest design. Plain Claude - the most polished visuals and exec storytelling,
but invented facts without flagging them and one-line notes. pptx skill - the most real charts, but almost no
speaker notes and unflagged inventions.

## Round 1 - before (kept for honesty)

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall** | Time / tokens / calls |
|---|---|---|---|---|---|---|---|---|
| Plain Claude | 9.0 | 8.0 | 7.0 | 8.0 | 7.7 | 5.7 | **7.56** | 165 s / 77k / 10 |
| Anthropic pptx skill | 8.0 | 8.0 | 7.0 | 8.0 | 7.7 | 1.7 | 6.72 | 159 s / 92k / 12 |
| This skill | 6.7 | 4.7 | 5.0 | 6.0 | 5.3 | 5.0 | 5.44 | 102 s / 91k / 12 |

Round 1 found sparse slides, small text and generic visuals. The builder was changed in response (slides fill the
space, five new patterns, content rules), and a fresh agent that had not seen the changes built the round-2 decks.
The other two makers' decks are the same in both rounds; their scores moved by at most 0.16, which gives a feel for
the judge's noise.

## Speed

Quality is not the only cost: an agent that loops build → lint → render → edit many times is slow and expensive.
From round 5 on every round records, per maker, the **wall time, total tokens and tool calls** of the unattended
agent that built the three decks (the *Time / tokens / calls* column above).

**Method.** Each maker runs as one unattended agent with the same three briefs ([briefs.md](briefs.md)) and the
same instruction to finish alone. The numbers are the agent run's own totals as the harness reports them when it
ends (`duration`, `total_tokens`, `tool_uses`) - all three decks together, not per deck. One run per maker and
round, so treat differences under about 15 % as noise. Script time is measured separately with `time` on the
final specs: `build_deck.py --plan` 0.4 s, build ~1.8 s, `lint_deck.py` ~1.3 s, `render_lo.py --sheet` ~3.5 s
per deck (round 5) - the scripts are not the bottleneck, the number of loops is.

| Round | Plain Claude | Anthropic pptx skill | This skill |
|---|---|---|---|
| 1 | 165 s / 77k / 10 | 159 s / 92k / 12 | 102 s / 91k / 12 |
| 2 | not measured | not measured | 179 s / 95k / 19 |
| 3 | not measured | not measured | 173 s / 98k / 18 |
| 4 | not measured | not measured | 189 s / 100k / 18 |
| 5 | not measured | not measured | 205 s / 107k / 26 |
| 6 | 124 s / 73k / 9 | 159 s / 90k / 12 | 131 s / 91k / 13 |
| 7 | 162 s / 75k / 3 | 189 s / 84k / 6 | 212 s / 100k / 11 |
| 8 | 68 s / 64k / 3 | 279 s / 94k / 12 | 328 s / 97k / 21 |
| 9 | 168 s / 76k / 7 | 175 s / 86k / 9 | 246 s / 100k / 16 |

**Reading it.** As the skill gained checks (rounds 2-5) its quality rose from 5.44 to 8.56, but the agent needed
more loops: each check found one class of problem per build, so it built, rendered, fixed one thing and built
again. The speed round answers that with one-call `build_deck.py --check` (build + lint + render + one summary of
every problem, sorted by slide, each with its spec edit), automatic fixes for unfit text (`auto:` lines), `--plan`
pre-checks and a checklist-shaped `SKILL.md` (18.7 KB → 7.7 KB). Target for round 6: at or under plain Claude's
tool calls with the round-5 quality.

## Lint (this repo's own checker - home ground, read with care)

| Deck | Plain Claude err / warn | pptx skill err / warn | This skill err / warn |
|---|---|---|---|
| QBR | 11 / 38 | 6 / 23 | 0 / 0 |
| Phishing | 6 / 51 | 7 / 31 | 0 / 0 |
| Four-day week | 13 / 46 | 2 / 27 | 0 / 0 |

The other makers' most common findings were text below the projection floor, overlapping shapes, low-contrast text
and missing alt text.

## Limits

Three briefs, one judge model, LibreOffice renders (fonts substitute), a single run per maker. Treat it as a
signal, not a leaderboard. The other two makers' decks are identical in all three rounds; their overall scores moved by up to 0.56 between
judges, so differences under about half a point are noise. Next for this skill: visual flair on
teaching decks (round 5's main criticism) and the rough edges the round-5 maker reported.
