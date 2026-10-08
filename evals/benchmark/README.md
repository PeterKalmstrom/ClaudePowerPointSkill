# Benchmark: this skill vs plain Claude vs Anthropic's pptx skill

Three briefs ([briefs.md](briefs.md)): a quarterly business review with real numbers, a 20-minute phishing training,
a board proposal for a four-day-week pilot. Each maker built all three decks, unattended, in its own folder:

- **Plain Claude** - python-pptx, no skill.
- **Anthropic's pptx skill** (`anthropic-skills:pptx`, pptxgenjs).
- **This skill** - spec -> `build_deck.py` -> `lint_deck.py` / `fix_deck.py` -> LibreOffice renders.

All 27 slides per maker were rendered with LibreOffice and scored **blind** by a separate Claude judge that saw
only anonymous labels (keys: `round*-key.json`), on six criteria from 1 to 10: message, visual design, legibility,
layout correctness, data presentation, presenter support (speaker notes).

## Round 2 - after the builder quality round (current)

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall** |
|---|---|---|---|---|---|---|---|
| **This skill** | 8.7 | 6.7 | 7.7 | 8.7 | 7.3 | 8.7 | **7.94** |
| Plain Claude | 8.3 | 8.7 | 8.0 | 8.3 | 7.7 | 5.3 | 7.72 |
| Anthropic pptx skill | 8.0 | 8.0 | 7.0 | 8.0 | 7.7 | 2.0 | 6.78 |

Judge, in short: this skill - best speaker notes by far, claim titles, faithful to the brief with assumptions
flagged, clean data slides; but the plainest design. Plain Claude - the most polished visuals and exec storytelling,
but invented facts without flagging them and one-line notes. pptx skill - the most real charts, but almost no
speaker notes and unflagged inventions.

## Round 1 - before (kept for honesty)

| Maker | Message | Design | Legibility | Layout | Data | Notes | **Overall** |
|---|---|---|---|---|---|---|---|
| Plain Claude | 9.0 | 8.0 | 7.0 | 8.0 | 7.7 | 5.7 | **7.56** |
| Anthropic pptx skill | 8.0 | 8.0 | 7.0 | 8.0 | 7.7 | 1.7 | 6.72 |
| This skill | 6.7 | 4.7 | 5.0 | 6.0 | 5.3 | 5.0 | 5.44 |

Round 1 found sparse slides, small text and generic visuals. The builder was changed in response (slides fill the
space, five new patterns, content rules), and a fresh agent that had not seen the changes built the round-2 decks.
The other two makers' decks are the same in both rounds; their scores moved by at most 0.16, which gives a feel for
the judge's noise.

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
signal, not a leaderboard. Next for this skill: design polish (the judge's main criticism) and the rough edges the
round-2 maker reported (wrapping inside small boxes that lint does not see yet).
