# Benchmark prompts

The exact prompts used from round 10 on (round 9 used the same text for three briefs and one run per maker), so
every round runs the same way. Replace the placeholders before use:

- `<BENCH>`: the round's working folder (for example `_scratch/bench10`), holding `briefs.md` = [../briefs.md](../briefs.md)
  followed by [../briefs-holdout.md](../briefs-holdout.md) (five briefs).
- `<OUT>`: the maker run's own folder, `<BENCH>/<maker>-<run>` (for example `_scratch/bench10/thisskill-2`).
- `<PYTHON>`: the Python to use (the skill's `.venv`).
- `<SOFFICE>`: the full path to LibreOffice's `soffice` executable. Give the full path: an agent's shell may not
  see a PATH change made after the session started (in round 9 two makers could not find LibreOffice that way).
- Judge only: `<LABELS>` (the random anonymous labels, one per maker run) and `<OUTFILE>` (one file per judge,
  outside `<BENCH>/judge`). Keep the label key outside `<BENCH>/judge` as well.

Files: [maker-plain.txt](maker-plain.txt), [maker-pptxskill.txt](maker-pptxskill.txt),
[maker-thisskill.txt](maker-thisskill.txt) (two independent runs per maker, all six as parallel unattended agents),
[judge.txt](judge.txt) (two independent blind judges after every deck is rendered in PowerPoint into `<BENCH>/judge`).

## Judging from round 11 on: the judge pack

Round 10's judges ([judge-round10.txt](judge-round10.txt), kept for reproducing that round) worked mostly from
6-column contact sheets, opened about ten single slides each and scored one maker almost the same on every deck.
From round 11:

1. After rendering into `<BENCH>/judge`, build the pack:
   `<PYTHON> evals/benchmark/make_judge_pack.py <BENCH>` -> `<BENCH>/judge-pack/<L>-<deck>/` with pages of 2 x 2
   slides at 960 px each (`p01.png` ...), the text and notes of exactly those slides (`p01.txt` ...) and an
   `index.md` listing the pages. `--cols 1 --rows 2` makes even larger pages; `--only M-qbr ...` limits the decks.
2. [judge.txt](judge.txt) points the judge at the pack only, requires every page to be opened, a per-slide note
   (`slides`) and a `defects` list per deck before scoring, anchors each criterion with what a 5, 7 and 9 look
   like, and asks the judge to report how many pages it opened per deck. Extra placeholder: `<DECKS>` (the deck
   names). The output JSON keeps the round-10 shape plus `slides` and `defects`.

Trial (2026-10-10, round-10 postmortem, one judge, makers M = this skill run 1 and F = pptx skill run 2): the
judge opened 2/2 and 3/3 pages (all 17 slides), wrote 6 defects per deck and scored M 9/8/8/8/7/6 (7.67) and
F 7/7/8/6/8/1 (6.17), against round 10's M 8.17 / 7.67 and F 6.5 / 6.5 (F was 8/7/7/7/x/1-2 on all five decks).
The new scores trace to named defects (F: topic titles s4-s8, half-empty s4-s6/s8, no notes; M: unitless
root-cause chart s6, thin notes) and F's criteria are no longer flat.
