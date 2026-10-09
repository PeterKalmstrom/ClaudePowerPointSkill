# Benchmark: this skill vs plain Claude vs Anthropic's pptx skill

Three briefs ([briefs.md](briefs.md)): a quarterly business review with real numbers, a 20-minute phishing training,
a board proposal for a four-day-week pilot. Each maker built all three decks, unattended, in its own folder:

- **Plain Claude** - python-pptx, no skill.
- **Anthropic's pptx skill** (`anthropic-skills:pptx`, pptxgenjs).
- **This skill** - spec -> `build_deck.py` -> `lint_deck.py` / `fix_deck.py` -> LibreOffice renders.

All 27 slides per maker were rendered with LibreOffice and scored **blind** by a separate Claude judge that saw
only anonymous labels (keys: `round*-key.json`), on six criteria from 1 to 10: message, visual design, legibility,
layout correctness, data presentation, presenter support (speaker notes).

## Round 6 - after the speed round (current)

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
