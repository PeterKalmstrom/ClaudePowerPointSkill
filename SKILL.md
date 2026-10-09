---
name: building-powerpoint-decks
description: Build, edit, audit and present PowerPoint decks with Claude — via COM automation of desktop PowerPoint on Windows, or python-pptx anywhere. Load when creating or editing a .pptx, auditing a deck, preparing to present one, adding animation, video or AI images to slides, or when PowerPoint automation on Windows is failing.
---

# Building PowerPoint decks

Rules, traps and helper scripts for producing decks that render correctly the first time. Detail lives in
[`reference/`](reference/); helper scripts in [`scripts/`](scripts/README.md). Read this file fully, then open
only the reference file the task needs.

## Start here

| Task | Read, in order |
|---|---|
| **New deck** | The [checklist below](#new-deck--the-checklist); detail in [WORKFLOW](reference/WORKFLOW.md#new-deck-in-detail), patterns in [BUILDER](reference/BUILDER.md) |
| **Edit an existing deck** | Snapshot first · `scripts/read_deck.py` before judging ([AUDIT](reference/AUDIT.md)) · find the open deck by name ([COM](reference/COM.md)) · rebuilding a generated deck? `scripts/harvest_edits.py` first |
| **Review / audit a deck** | `scripts/lint_deck.py` first (any OS) · `scripts/fix_deck.py deck.pptx --out fixed.pptx` (or `--in-place`, or `--dry-run`; one is required) for the mechanical fixes · then [AUDIT](reference/AUDIT.md): taste pass → anchor exceptions → contact sheet |
| **Before a talk** | Audit clean first, then [PRESENTING](reference/PRESENTING.md) |
| **Images, video, Remotion, Veo** | [MEDIA](reference/MEDIA.md), then [ANIMATION](reference/ANIMATION.md) if it moves |
| **PowerPoint automation failing (Windows)** | [SETUP](reference/SETUP.md) → Troubleshooting |
| **No Windows (Linux, macOS, CI)** | [AUTOMATION](reference/AUTOMATION.md) → *Building .pptx without PowerPoint* |
| **Company template / brand** | `scripts/extract_theme.py` → [LAYOUT](reference/LAYOUT.md) → *Building from a template* |
| **Corporate / labelled deck** | [LABELS](reference/LABELS.md) before choosing a toolchain |


## New deck — the checklist

```
1. Spec      content first: audience, claim titles, notes; numbers in `facts` with {sum}/{change}/{share}
2. --plan    python scripts/build_deck.py spec.json --plan          (story + pre-checks; fix `spec:` lines)
3. --check   python scripts/build_deck.py spec.json --out deck.pptx --check
4. Look      open the contact sheet it names - once, all slides together
5. Fix once  make every edit the summary lists (sorted by slide, each with its spec edit), then step 3 again
6. Done      exit 0 and a sheet that looks right; review the `auto:` lines it printed
```

- `--check` = build + lint + `render_lo.py --sheet` (when LibreOffice is installed) + one summary with a
  `CHECK-JSON` line. Exit 0 clean, 2 spec error, 3 text that still does not fit.
- Text that does not fit is fixed first (filler words, short units, detail moved to the notes as
  `MOVED FROM SLIDE:`, a split where `"allow_split": true`); each change prints as `auto:`. `--no-auto` reports only.
- Pick patterns that show the thing itself, give the deck a `footer` and the ask a `decision` with its `figure`
  ([BUILDER](reference/BUILDER.md), [DESIGN](reference/DESIGN.md)). Treat `spec warning:` lines as content to fix.
- **4+ slides in one design:** showcase-first - `--slides 1,3 --check`, look, then the rest; unattended, approve it
  yourself and say so ([WORKFLOW](reference/WORKFLOW.md#showcase-first-for-multi-slide-sections)).
- Three fixes on the same defect without success: escalate - [iteration loop](reference/WORKFLOW.md#iteration-loop--build-render-look-critique-fix).

## Core rules

These apply to every deck. Each links to its full explanation.

1. **Snapshot before anything destructive** — `scripts/backup_snapshot.py`. Undo is not a backup. ([COM](reference/COM.md#snapshot-before-any-risky-bulk-edit))
2. **Never trust `ActivePresentation`** — find the deck by exact name, then substring. ([COM](reference/COM.md#multi-presentation-safety--never-trust-activepresentation))
3. **Rebuild, don't patch** — one idempotent `build_slide_NN.py` per non-trivial slide; name every shape. Three patches = rewrite. ([COM](reference/COM.md#idempotent-build-scripts))
4. **Never overwrite people's edits** — `scripts/harvest_edits.py` before regenerating a deck. ([COM](reference/COM.md#a-generated-deck-that-people-also-edit-in-powerpoint-harvest-before-you-overwrite))
5. **Full HD = 1440 × 810 pt**, set before inserting slides. ([LAYOUT](reference/LAYOUT.md#slide-size--set-it-before-inserting-anything))
6. **About 10 visible words per slide (by anchor type), body ≥ 18 pt on a 960-pt slide — ≥ 27 pt on Full HD (1440 pt)**, claim titles. ([LAYOUT](reference/LAYOUT.md#anchor-types-and-word-budgets), [CONTENT](reference/CONTENT.md#titles-make-a-claim-not-a-topic))
7. **No text baked into images or video** — overlay it in PowerPoint. ([MEDIA](reference/MEDIA.md#generate-images-without-text--overlay-text-in-powerpoint))
8. **Crop pictures, never stretch them.** ([LAYOUT](reference/LAYOUT.md#pictures-stretch--crop-to-fill-never-pass-both-sizes-blindly))
9. **The slide carries the punch, the notes carry the depth** — notes on every slide, written first, fixed order. Never invent facts beyond the brief: put anything you add in the notes' `assumptions` (written as `ASSUMPTIONS:`); when the brief is the only source, say so — deck-level `"sources": ["the brief"]`. An ask states its reasons and its cost or impact. ([CONTENT](reference/CONTENT.md#every-slide-gets-notes-never-invent-facts-beyond-the-brief), [ask](reference/CONTENT.md#the-ask-states-its-reasons-and-its-cost))
10. **Lint, then render and LOOK** before calling anything done — `scripts/lint_deck.py` catches what the file shows (overlaps, stretched pictures, small text, missing alt text) on any OS; the render catches the rest. Render via Save As JPEG on Windows, not `Slide.Export`; `scripts/render_lo.py` elsewhere (approximate). (Loop below; [LAYOUT](reference/LAYOUT.md#embedded-fonts-slideexport-renders-a-fallback--use-save-as-jpeg))

## Error reports - always ask first

A script that exits with code **4** (stderr: `ERROR-REPORT-PENDING: <file>`) hit an unexpected error whose report
is waiting. Ask the user exactly "Do you want to send this error message?" (AskUserQuestion, Yes / No) and show
what would be sent. Only on **Yes** run the `send_error_report.py <file> --yes` command it printed; on **No** run
it with `--no`. PowerPoint Live says `ERROR-REPORT-PENDING` in its error message: ask the same question, then call
`powerpoint_send_error_report` with `send=true` only on Yes (`false` on No). Never send without the user's yes.

## Version-specific facts (as of 2026-10)

These depend on upstream releases. Re-check them when something that used to work breaks:

- Veo model ids `veo-3.1-generate-preview` and `veo-3.1-fast-generate-preview` are confirmed; a `lite` id is not ([MEDIA](reference/MEDIA.md#key-details)).
- `Slide.Export` shows fallback fonts for embedded fonts; `scripts/render_slides.py` does not.
- python-pptx 1.0.2 behaviour is assumed in the slide-copy notes ([COM](reference/COM.md#a-generated-deck-that-people-also-edit-in-powerpoint-harvest-before-you-overwrite)).

## Anti-patterns

The catalogue of silent failures that have shipped broken slides - trusting `ActivePresentation`, the
`HasTextFrame` filter, moving text without its backing, skipping the LOOK step, both sizes on `AddPicture`,
inventing facts, adding up by hand and more - is in [WORKFLOW](reference/WORKFLOW.md#anti-patterns-recurring-com--build-traps).
Skim it before a non-trivial COM patch; add a row there when a new one bites.
