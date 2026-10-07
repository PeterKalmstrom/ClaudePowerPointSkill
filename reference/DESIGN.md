# Design tokens and directions

*Runs on: any OS.* Part of the `building-powerpoint-decks` skill — start at [`SKILL.md`](../SKILL.md).

## Sizes on a Full HD slide

Slides are 1440 × 810 pt; text looks two-thirds as big as on an old 960-pt slide, so every size rule
scales by 1.5. `lint_deck.py` and `audit_deck.py` do this automatically.

| Role | 960-pt slide | 1440-pt slide (Full HD) |
|---|---|---|
| Body floor for sentences | 18 pt | **27 pt** |
| Label / caption ceiling (short labels may go this small) | 12 pt | 18 pt |
| Slide title | 36 pt | 54 pt |
| Section / statement | 48–56 pt | 72–84 pt |
| Hero number | 120–150 pt | 180–220 pt |
| Body text | 22 pt | 34 pt |

Line height: about 1.15 for headings, 1.35 for body. Tracking: tight (−0.5) for big headings, wide (+1.5)
for small uppercase eyebrows.

## Spacing, lines, shadows

- **Spacing steps:** 4 · 8 · 16 · 24 · 36 · 56 pt on a 960-pt slide (× 1.5 on Full HD). Gaps *between*
  groups ≈ 3× the gaps *within* a group.
- **Lines:** hairline 0.5 pt, standard 0.75 pt, emphasis 1.5 pt; gridlines faint (≈ 15 % opacity).
- **Shadows:** at most one or two elevated surfaces per slide — subtle (blur 6, offset 2, ≈ 18 % opacity)
  or elevated (blur 14, offset 6, ≈ 30 %).
- **Accents:** one highlight per slide, at most three accent colours.

## Charts

- Data labels: on for pie/doughnut and single-series bars or columns; off for line, area, scatter and
  multi-series charts.
- Legend: none for one series; otherwise on the right — a top or bottom legend steals plot height.
- Series of one hue (ordered data such as months or years): step the shade from about −40 % to +40 %
  (`tint_ramp()` in `scripts/_rules.py`); beyond −50 % it turns muddy, beyond +60 % it washes out.
- Gridlines thin and light, minor gridlines off; markers only up to ~12 points.
- Title the slide with the finding, not the chart with its dimensions.

## Tables

Header row about 28 pt tall (× 1.5 on Full HD), body rows 24 pt, cell padding 4 / 8 pt, light banding,
numbers right-aligned, one highlighted row at most.

## Motion

Entrance 0.3 s, emphasis 0.4 s, exit 0.25 s, transitions 0.4 s, chart elements staggered by ~0.12 s —
and the whole slide built within ~2 s ([ANIMATION.md](ANIMATION.md)).

## Design directions

Twenty starting points, written for this skill and stored in
[`scripts/directions.json`](../scripts/directions.json). Pass one as `direction` to `build_deck.py`, or use
them as a vocabulary when the user asks for "something bolder" or "more formal". Every direction keeps
text at 7:1 or better, muted text at 4.5:1 and the accent at 3:1 against its background (checked by the
self-test). Pick by audience first ([CONTENT.md](CONTENT.md#pick-the-audience-before-the-first-slide)), then by mood;
with a company template, use its theme instead.

| Direction | Tone | Mood | Fonts (heading / body) | Accent on background | Suits |
|---|---|---|---|---|---|
| `clean-corporate` | safe | calm, trustworthy, no surprises | Segoe UI Semibold / Segoe UI | `#1F5FAD` on `#FFFFFF` | kpi, chart, compare, process |
| `quiet-neutral` | safe | understated, lets the content lead | Calibri / Calibri | `#2F6F5E` on `#FAFAF8` | bullets, table, chart |
| `public-sector` | safe | plain, accessible, high contrast | Arial / Arial | `#005EA5` on `#FFFFFF` | process, bullets, table |
| `consulting-blue` | safe | structured, analytical, answer first | Georgia / Segoe UI | `#0B3D91` on `#FFFFFF` | matrix, chart, compare |
| `teaching-friendly` | safe | warm, clear, step by step | Verdana / Verdana | `#0E7C66` on `#FFFDF7` | process, timeline, bullets |
| `bold-signal` | bold | one loud colour, big numbers | Arial Black / Arial | `#D7261E` on `#FFFFFF` | big_number, statement, kpi |
| `dark-stage` | bold | keynote on a dark stage | Segoe UI Semibold / Segoe UI | `#35C2A0` on `#121417` | statement, big_number, quote |
| `electric-launch` | bold | energetic product launch | Trebuchet MS / Segoe UI | `#5B2EE5` on `#FFFFFF` | image_headline, big_number, compare |
| `sunset-warm` | bold | optimistic, human, warm | Georgia / Calibri | `#C2410C` on `#FFF8F1` | quote, statement, image_headline |
| `data-dense` | bold | dashboards for people who read numbers | Segoe UI Semibold / Segoe UI | `#0F766E` on `#F7F9FA` | kpi, chart, table |
| `editorial-serif` | refined | long-form magazine feature | Georgia / Georgia | `#8B1E3F` on `#FBF9F5` | quote, statement, image_headline |
| `gallery-white` | refined | lots of air, few words | Segoe UI Light / Segoe UI | `#3A3A3A` on `#FFFFFF` | image_headline, statement, section |
| `boardroom` | refined | formal, measured, senior audience | Cambria / Calibri | `#8A6D1D` on `#FDFCF8` | kpi, compare, table |
| `nordic-calm` | refined | soft, natural, considered | Segoe UI Semibold / Segoe UI | `#3E6B5B` on `#F4F1EC` | timeline, process, statement |
| `ink-and-paper` | refined | monochrome with one deep accent | Palatino Linotype / Calibri | `#1E3A5F` on `#FFFFFF` | quote, bullets, compare |
| `brutalist-mono` | contrarian | raw, typewriter, deliberately plain | Consolas / Consolas | `#000000` on `#F2F2EE` | statement, bullets, process |
| `poster-yellow` | contrarian | loud poster, black on yellow | Arial Black / Arial | `#111111` on `#FFD60A` | statement, big_number, section |
| `night-terminal` | contrarian | engineering, console green on black | Consolas / Segoe UI | `#4ADE80` on `#0B0F0C` | process, timeline, big_number |
| `blueprint` | contrarian | technical drawing, white on blue | Segoe UI Semibold / Segoe UI | `#FFD166` on `#0E3A68` | process, matrix, timeline |
| `risograph` | contrarian | zine print, two inks | Georgia / Trebuchet MS | `#D6336C` on `#FFF4E6` | quote, statement, compare |

Tones: **safe** for internal and conservative audiences, **bold** for launches and keynotes, **refined**
for senior or editorial audiences, **contrarian** when the deck must look unlike everyone else's.
Fonts are ones that ship with Windows and Office; on other systems the theme falls back gracefully.
