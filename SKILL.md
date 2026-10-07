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
| **New deck** | Audience → claim titles → word budgets ([CONTENT](reference/CONTENT.md), [LAYOUT](reference/LAYOUT.md)) · write a spec (`scripts/spec.schema.json`), show `--plan`, then build with `scripts/build_deck.py` ([BUILDER](reference/BUILDER.md), looks in [DESIGN](reference/DESIGN.md)) · showcase-first (below) · the build loop (below) |
| **Edit an existing deck** | Snapshot first · `scripts/read_deck.py` before judging ([AUDIT](reference/AUDIT.md)) · find the open deck by name ([COM](reference/COM.md)) · rebuilding a generated deck? `scripts/harvest_edits.py` first |
| **Review / audit a deck** | `scripts/lint_deck.py` first (any OS) · `scripts/fix_deck.py` for the mechanical fixes · then [AUDIT](reference/AUDIT.md): taste pass → anchor exceptions → contact sheet |
| **Before a talk** | Audit clean first, then [PRESENTING](reference/PRESENTING.md) |
| **Images, video, Remotion, Veo** | [MEDIA](reference/MEDIA.md), then [ANIMATION](reference/ANIMATION.md) if it moves |
| **PowerPoint automation failing (Windows)** | [SETUP](reference/SETUP.md) → Troubleshooting |
| **No Windows (Linux, macOS, CI)** | [AUTOMATION](reference/AUTOMATION.md) → *Building .pptx without PowerPoint* |
| **Company template / brand** | `scripts/extract_theme.py` → [LAYOUT](reference/LAYOUT.md) → *Building from a template* |
| **Corporate / labelled deck** | [LABELS](reference/LABELS.md) before choosing a toolchain |

## Reference files

| File | Covers | Runs on |
|---|---|---|
| [SETUP.md](reference/SETUP.md) | Windows setup, driving PowerPoint from Python, troubleshooting | Windows + PowerPoint |
| [mcp-app/](mcp-app/README.md) | PowerPoint Live: own MCP server with a live view of the current slide | Windows + PowerPoint |
| [COM.md](reference/COM.md) | Snapshots, multi-deck safety, UTF-8, idempotent builds, keeping hand edits, shape filtering | Windows + PowerPoint |
| [LAYOUT.md](reference/LAYOUT.md) | Slide size, pictures, text wrap, embedded fonts, word budgets, anchor types, layout patterns | Mostly any OS; rendering needs Windows |
| [MEDIA.md](reference/MEDIA.md) | AI images, Remotion, Veo, compression, embedding, combined patterns | Generation any OS; embedding Windows |
| [ANIMATION.md](reference/ANIMATION.md) | When motion earns its place, effect table, native timing traps | Any OS (COM/XML parts marked) |
| [BUILDER.md](reference/BUILDER.md) | Spec-driven deck builder: patterns, limits, chart and type defaults | Any OS |
| [DESIGN.md](reference/DESIGN.md) | Type scale for Full HD, spacing, chart and table defaults, 20 design directions | Any OS |
| [CONTENT.md](reference/CONTENT.md) | Audience, claim titles, cognitive load, speaker notes, showcase-first | Any OS |
| [AUDIT.md](reference/AUDIT.md) | Lint codes, reading a whole deck, defect catalogue with severities, full audit procedure | Scripts Windows; catalogue any OS |
| [PRESENTING.md](reference/PRESENTING.md) | Timing markers, Q&A sheet, notes PDF, rehearsal | Windows + PowerPoint |
| [AUTOMATION.md](reference/AUTOMATION.md) | Headless PowerPoint for checks, building .pptx without COM | Headless Windows; no-COM any OS |
| [LABELS.md](reference/LABELS.md) | Sensitivity labels and what encryption breaks | Any OS |

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
9. **The slide carries the punch, the notes carry the depth** — notes first, fixed order. ([CONTENT](reference/CONTENT.md#write-the-notes-first-in-a-fixed-order))
10. **Lint, then render and LOOK** before calling anything done — `scripts/lint_deck.py` catches what the file shows (overlaps, stretched pictures, small text, missing alt text) on any OS; the render catches the rest. Render via Save As JPEG on Windows, not `Slide.Export`; `scripts/render_lo.py` elsewhere (approximate). (Loop below; [LAYOUT](reference/LAYOUT.md#embedded-fonts-slideexport-renders-a-fallback--use-save-as-jpeg))

## Showcase-first for multi-slide sections

For any section of **4+ slides** that share a design, build the opener and one detail slide first, render both,
and ask the user to approve before producing the rest. Iterating slide 1 of 6 is cheap; rebuilding all six in the
wrong design is not.

## Workflow

The five-step iteration loop that catches silent rendering failures. Use it for every slide edit.

### Iteration loop — build, render, LOOK, critique, fix

**Rule:** every slide edit follows the same five-step loop. Step 3 is the one that gets skipped. Don't skip it.

```
1. Build / edit (idempotent script)
2. Render via presentation.SaveCopyAs(folder, 17)  (Save As JPEG — scripts/render_slides.py)
   slide.Export is fine ONLY when every font in the deck is installed; never for embedded fonts
3. LOOK at the actual rendered image  ←  DO NOT SKIP
4. Lint (scripts/lint_deck.py), then self-critique against the defect catalogue (reference/AUDIT.md)
5. Fix → loop back to step 2,  OR  save → done
```

**Why step 3 matters:** PowerPoint COM properties lie about what you'll see. Rendered PNG is ground truth. Specific silent failures that only show in the render:

- Text wraps to 2 lines despite "fitting" on paper
- Backing rectangle and text desync after one is moved without the other
- Shape stays behind another due to z-order, invisible in render
- Image cache holds the old version after file replacement
- Title placeholder appears empty in COM but renders white-on-white text
- AutoShape rectangles match `HasTextFrame == True`, breaking filter logic

The tool returning `success: true` is *the tool's claim*, not evidence. The PNG is evidence.

**Loop budget:** if 3 fix attempts on the same defect haven't worked, escalate:

- Check if the rule itself is wrong for this anchor type (gallery, knowledge graph, quote — see [Anchor types and word budgets](reference/LAYOUT.md#anchor-types-and-word-budgets))
- Switch from patch to full rebuild — patches drift after the third one
- Ask the user — the expected output may itself be wrong

Never run more than 4 iterations on the same defect without escalating.

**Anti-patterns to recognise:**

- *Source-look-only iteration* — reading COM properties or script values, declaring success without exporting
- *Patch on patch on patch* — each patch tweaks one element without considering layout interactions; backing rectangles drift, z-order shifts, third patch produces shape soup. **Three patches = rewrite the build script.**
- *Tool-output as ground truth* — `{"success": true}` from a script or tool is not the same as a correctly rendered slide
- *Optimistic font sizing* — "36pt should fit" is a guess until the PNG confirms it
- *Caching trust* — replacing an image at the same path doesn't always update the embedded version; re-add the picture explicitly
- *Active context drift* — see [Multi-presentation safety](reference/COM.md#multi-presentation-safety--never-trust-activepresentation); pin to a specific presentation by name, never trust `ActivePresentation`

## Error reports - always ask first

When a script prints `CLAUDE: ask the user "Do you want to send this error message?"`, ask the user exactly that
(AskUserQuestion, Yes / No) and show what would be sent. Only on **Yes** run the `send_error_report.py ... --yes`
command it printed; on **No** run it with `--no`. Never send an error report without the user's yes.

## Version-specific facts (as of 2026-10)

These depend on upstream releases. Re-check them when something that used to work breaks:

- Veo model ids `veo-3.1-generate-preview` and `veo-3.1-fast-generate-preview` are confirmed; a `lite` id is not ([MEDIA](reference/MEDIA.md#key-details)).
- `Slide.Export` shows fallback fonts for embedded fonts; `scripts/render_slides.py` does not.
- python-pptx 1.0.2 behaviour is assumed in the slide-copy notes ([COM](reference/COM.md#a-generated-deck-that-people-also-edit-in-powerpoint-harvest-before-you-overwrite)).

---

## Anti-patterns (recurring COM / build traps)

A consolidated catalog of the silent failures that have actually shipped broken slides. Skim this list whenever you're about to write a non-trivial COM patch — most of these fail without raising an error, so they don't show up in stack traces.

| Anti-pattern | What goes wrong | See |
|---|---|---|
| Trusting `app.ActivePresentation` | Targets the wrong deck if another window has focus | [Multi-presentation safety](reference/COM.md#multi-presentation-safety--never-trust-activepresentation) |
| `if not sh.HasTextFrame:` filter | Silently skips AutoShapes (they have empty text frames) | [Filtering shapes](reference/COM.md#filtering-shapes--hastextframe-is-not-is-this-a-text-shape) |
| Moving text without moving its backing | Text floats outside its card or off-slide | [Move text and its backing together](reference/COM.md#move-text-and-its-backing-together) |
| Patching a 5+ shape slide instead of rebuilding | Z-order drifts, conditional filters miss shapes, layout never quite matches | [Idempotent build scripts](reference/COM.md#idempotent-build-scripts) |
| Skipping the LOOK step (step 3 of the loop) | Code "succeeds" but the rendered slide is wrong | [Iteration loop](#iteration-loop--build-render-look-critique-fix) |
| Re-prompting an AI image after 2 text-baked attempts | Wastes budget; the model is locked into baking text | [Generate images WITHOUT text](reference/MEDIA.md#generate-images-without-text--overlay-text-in-powerpoint) |
| Re-prompting Veo for a physics chain reaction | Diffusion video reliably fails dominoes / cradle / pool break | [Known limits — when not to use Veo](reference/MEDIA.md#known-limits--when-not-to-use-veo) |
| Inline reimplementation of helper logic | Drifts from canonical version, no `--quiet` flag, no error stream propagation | [Wrapping Python helpers via subprocess + uvx](reference/MEDIA.md#wrapping-python-helpers-via-subprocess--uvx) |
| Author-name marker for notes-append idempotency | Marker mismatches actual citation format ("X & Y" vs "X T., Y Q."), block gets appended every re-run | [Idempotent notes appending](reference/CONTENT.md#idempotent-notes-appending--use-the-doi-not-author-names) |
| Missing UTF-8 stdout in scripts that `print()` Unicode | Script crashes *after* `target.Save()` succeeds; operator re-runs and double-applies append-style work | [Force UTF-8 stdout](reference/COM.md#force-utf-8-stdout-in-any-script-that-prints-unicode) |
| Sentinel-text check uses a phrase from the pre-build slide | Idempotency check skips structural step (Duplicate / Insert), then rewrites text — destroys adjacent unrelated slides | [Idempotent build scripts](reference/COM.md#idempotent-build-scripts) (Sentinel rule) |
| `AddPicture` to replace a picture inside a Group | New picture lands as a sibling outside the group; old picture remains; layout breaks | [Swapping a picture inside a Group](reference/MEDIA.md#swapping-a-picture-that-lives-inside-a-group) |
| Hardcoded `OUT = r"C:\Users\<somebody>\..."` in chart scripts | Script does nothing useful on any other machine; PNG fails to update; embedded chart looks stale forever | [Portable OUT paths](reference/MEDIA.md#portable-out-paths-in-chart-scripts) |
| Trusting `Slide.Export` for decks with embedded fonts | Renders a fallback font; real mid-word breaks look clean and pass review | [Embedded fonts](reference/LAYOUT.md#embedded-fonts-slideexport-renders-a-fallback--use-save-as-jpeg) |
| Rebuilding a generated deck over a hand-edited one | A day of human edits silently erased | [Harvest before you overwrite](reference/COM.md#a-generated-deck-that-people-also-edit-in-powerpoint-harvest-before-you-overwrite) |
| `Presentations.Add()` then `InsertFromFile` without setting the size | Whole deck silently scaled to 720p, all fonts a third smaller | [Slide size](reference/LAYOUT.md#slide-size--set-it-before-inserting-anything) |
| `AddPicture` / `add_picture` with both width and height | Photos stretched by up to ~60 %, unnoticed | [Pictures stretch](reference/LAYOUT.md#pictures-stretch--crop-to-fill-never-pass-both-sizes-blindly) |
| Validating a generated .pptx with `DisplayAlerts` off | PowerPoint silently repairs the broken file and the check passes | [Headless PowerPoint](reference/AUTOMATION.md#headless-powerpoint-for-checks) |
| Embedding raw Veo / Remotion MP4s | Deck balloons by hundreds of MB; some installs won't play non-H.264 | [Embedding video](reference/MEDIA.md#embedding-video-remotion-output) |
| Chaining effects with "After Previous" | One slow effect shifts every later one; builds drift | [Native animation traps](reference/ANIMATION.md#native-animation-traps) |
| Rebuilding an encrypted / labelled deck with python-pptx or `Presentations.Add()` | Output carries no sensitivity label — confidential content leaks unlabelled | [Sensitivity labels](reference/LABELS.md#sensitivity-labels) |
| Calling `ExportAsFixedFormat` positionally from pywin32 | `TypeError: The Python instance can not be converted to a COM object` — no PDF | [Presenter prep](reference/PRESENTING.md#presenter-prep) |
| `for p in app.Presentations: if ...: target = p` without `break` | Picks the *last* matching presentation in enumeration order (effectively random when multiple match the substring) | [Multi-presentation safety](reference/COM.md#multi-presentation-safety--never-trust-activepresentation) |

When one of these bites, fix it and **add a row here** if it's a new variant. The signal is: "I lost an hour to a silent failure" → it belongs in this table.
