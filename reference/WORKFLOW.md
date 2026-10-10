# Workflow: the build loop, showcase-first and the anti-patterns

The long form of what `SKILL.md` lists as a checklist. Read it when a check keeps failing, when you work by COM
patch instead of the spec builder, or when you add an anti-pattern.

## New deck in detail

The *New deck* checklist in `SKILL.md`, step by step:

- **Content first.** Audience → claim titles → word budgets ([CONTENT](CONTENT.md), [LAYOUT](LAYOUT.md)).
- **Show the thing itself.** `email`, `kpi_chart`, `cost_table`, `quiz`, `risks`, `metrics`, `next_steps` beat
  bullets; give the deck a `footer` and the ask a `decision` with the number it moves (`figure`); before/after
  numbers get a `trend` (they become a chart), a share a dot grid, a statement its `points`
  ([BUILDER](BUILDER.md), looks in [DESIGN](DESIGN.md)).
- **Never compute by hand.** Put the brief's numbers in `facts` and write `{sum}` / `{change}` / `{share}`
  ([computed figures](BUILDER.md#computed-figures)).
- **`spec warning:` lines are content to fix** - a figure that does not add up, vague targets, a skipped month, a
  sparse slide.
- **`--plan --apply` before building.** It prints the story and runs every pre-check: pattern limits, figures,
  warnings, each slide's words against its budget, each title measured in the deck's heading font, and a dry build
  (fit and layout lint). Each finding it can fix without new facts gets a measured edit, which `--apply` writes into
  the spec ([ready edits](BUILDER.md#ready-edits-and---apply)); the rest are `question:` lines - answer them all in
  one edit.
- **`--check --apply` builds, lints and renders in one call.** Text that does not fit is first fixed automatically
  (`auto:` lines - see [BUILDER](BUILDER.md#automatic-fixes-and---check)) and, with `--apply`, written back to the
  spec; what is left is listed once, every class together, sorted by slide, each with the spec edit that fixes it,
  and the contact sheet's path. Look at the sheet once, make every listed edit in one pass, run `--check` again.
  Clean means done - after a clean `--plan --apply` plus answers, the first `--check` usually is.
- **Review the `auto:` lines.** Each says which field changed and how; detail cut from a card is in the notes
  under `MOVED FROM SLIDE:`. If a cut reads badly, write the short version yourself; `--no-auto` reports only.

## Showcase-first for multi-slide sections

For any section of **4+ slides** that share a design, build the opener and one detail slide first, render both,
and ask the user to approve before producing the rest. Iterating slide 1 of 6 is cheap; rebuilding all six in the
wrong design is not.

**Unattended runs** (no one to ask — a batch job, a benchmark, an agent told to finish alone): still build the
opener and one detail slide first — `build_deck.py spec.json --out showcase.pptx --slides 1,3` builds only those
from the full spec, with the deck's theme and page numbers — render them (`render_lo.py --sheet`) and look at
them yourself against the brief and the design rules; fix what you see, then build the rest. Say in the final summary that the showcase was self-approved, and what you
changed after looking.

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

- Check if the rule itself is wrong for this anchor type (gallery, knowledge graph, quote — see [Anchor types and word budgets](LAYOUT.md#anchor-types-and-word-budgets))
- Switch from patch to full rebuild — patches drift after the third one
- Ask the user — the expected output may itself be wrong

Never run more than 4 iterations on the same defect without escalating.

**Anti-patterns to recognise:**

- *Source-look-only iteration* — reading COM properties or script values, declaring success without exporting
- *Patch on patch on patch* — each patch tweaks one element without considering layout interactions; backing rectangles drift, z-order shifts, third patch produces shape soup. **Three patches = rewrite the build script.**
- *Tool-output as ground truth* — `{"success": true}` from a script or tool is not the same as a correctly rendered slide
- *Optimistic font sizing* — "36pt should fit" is a guess until the PNG confirms it
- *Caching trust* — replacing an image at the same path doesn't always update the embedded version; re-add the picture explicitly
- *Active context drift* — see [Multi-presentation safety](COM.md#multi-presentation-safety--never-trust-activepresentation); pin to a specific presentation by name, never trust `ActivePresentation`

## Anti-patterns (recurring COM / build traps)

A consolidated catalog of the silent failures that have actually shipped broken slides. Skim this list whenever you're about to write a non-trivial COM patch — most of these fail without raising an error, so they don't show up in stack traces.

| Anti-pattern | What goes wrong | See |
|---|---|---|
| Trusting `app.ActivePresentation` | Targets the wrong deck if another window has focus | [Multi-presentation safety](COM.md#multi-presentation-safety--never-trust-activepresentation) |
| `if not sh.HasTextFrame:` filter | Silently skips AutoShapes (they have empty text frames) | [Filtering shapes](COM.md#filtering-shapes--hastextframe-is-not-is-this-a-text-shape) |
| Moving text without moving its backing | Text floats outside its card or off-slide | [Move text and its backing together](COM.md#move-text-and-its-backing-together) |
| Patching a 5+ shape slide instead of rebuilding | Z-order drifts, conditional filters miss shapes, layout never quite matches | [Idempotent build scripts](COM.md#idempotent-build-scripts) |
| Skipping the LOOK step (step 3 of the loop) | Code "succeeds" but the rendered slide is wrong | [Iteration loop](#iteration-loop--build-render-look-critique-fix) |
| Re-prompting an AI image after 2 text-baked attempts | Wastes budget; the model is locked into baking text | [Generate images WITHOUT text](MEDIA.md#generate-images-without-text--overlay-text-in-powerpoint) |
| Re-prompting Veo for a physics chain reaction | Diffusion video reliably fails dominoes / cradle / pool break | [Known limits — when not to use Veo](MEDIA.md#known-limits--when-not-to-use-veo) |
| Inline reimplementation of helper logic | Drifts from canonical version, no `--quiet` flag, no error stream propagation | [Wrapping Python helpers via subprocess + uvx](MEDIA.md#wrapping-python-helpers-via-subprocess--uvx) |
| Author-name marker for notes-append idempotency | Marker mismatches actual citation format ("X & Y" vs "X T., Y Q."), block gets appended every re-run | [Idempotent notes appending](CONTENT.md#idempotent-notes-appending--use-the-doi-not-author-names) |
| Missing UTF-8 stdout in scripts that `print()` Unicode | Script crashes *after* `target.Save()` succeeds; operator re-runs and double-applies append-style work | [Force UTF-8 stdout](COM.md#force-utf-8-stdout-in-any-script-that-prints-unicode) |
| Sentinel-text check uses a phrase from the pre-build slide | Idempotency check skips structural step (Duplicate / Insert), then rewrites text — destroys adjacent unrelated slides | [Idempotent build scripts](COM.md#idempotent-build-scripts) (Sentinel rule) |
| `AddPicture` to replace a picture inside a Group | New picture lands as a sibling outside the group; old picture remains; layout breaks | [Swapping a picture inside a Group](MEDIA.md#swapping-a-picture-that-lives-inside-a-group) |
| Hardcoded `OUT = r"C:\Users\<somebody>\..."` in chart scripts | Script does nothing useful on any other machine; PNG fails to update; embedded chart looks stale forever | [Portable OUT paths](MEDIA.md#portable-out-paths-in-chart-scripts) |
| Trusting `Slide.Export` for decks with embedded fonts | Renders a fallback font; real mid-word breaks look clean and pass review | [Embedded fonts](LAYOUT.md#embedded-fonts-slideexport-renders-a-fallback--use-save-as-jpeg) |
| Rebuilding a generated deck over a hand-edited one | A day of human edits silently erased | [Harvest before you overwrite](COM.md#a-generated-deck-that-people-also-edit-in-powerpoint-harvest-before-you-overwrite) |
| `Presentations.Add()` then `InsertFromFile` without setting the size | Whole deck silently scaled to 720p, all fonts a third smaller | [Slide size](LAYOUT.md#slide-size--set-it-before-inserting-anything) |
| `AddPicture` / `add_picture` with both width and height | Photos stretched by up to ~60 %, unnoticed | [Pictures stretch](LAYOUT.md#pictures-stretch--crop-to-fill-never-pass-both-sizes-blindly) |
| Validating a generated .pptx with `DisplayAlerts` off | PowerPoint silently repairs the broken file and the check passes | [Headless PowerPoint](AUTOMATION.md#headless-powerpoint-for-checks) |
| Embedding raw Veo / Remotion MP4s | Deck balloons by hundreds of MB; some installs won't play non-H.264 | [Embedding video](MEDIA.md#embedding-video-remotion-output) |
| Chaining effects with "After Previous" | One slow effect shifts every later one; builds drift | [Native animation traps](ANIMATION.md#native-animation-traps) |
| Rebuilding an encrypted / labelled deck with python-pptx or `Presentations.Add()` | Output carries no sensitivity label — confidential content leaks unlabelled | [Sensitivity labels](LABELS.md#sensitivity-labels) |
| Calling `ExportAsFixedFormat` positionally from pywin32 | `TypeError: The Python instance can not be converted to a COM object` — no PDF | [Presenter prep](PRESENTING.md#presenter-prep) |
| Inventing specifics the brief never gave (times, extensions, targets, "root cause fixed") | The room knows the real answer; credibility goes, and the notes can't defend it | [Never invent facts](CONTENT.md#every-slide-gets-notes-never-invent-facts-beyond-the-brief) |
| A bare "Approve X" ask slide | Reads as a slogan: no reasons, no cost, nothing to weigh | [The ask states its reasons and its cost](CONTENT.md#the-ask-states-its-reasons-and-its-cost) |
| Two numbers in two tiles ("1.8 %", "1.2 %") for one before/after | Reads as two unrelated facts; the change is never shown | [Patterns and limits](BUILDER.md#patterns-and-limits) (`trend`) |
| Adding up or dividing by hand ("21.8 MUSD, sum of quarters" for 4.1 + 4.6 + 5.2 + 5.9 = 19.8) | One wrong figure and the board doubts every other number | [Compute derived figures](CONTENT.md#compute-derived-figures-never-by-hand) |
| A success target with no number ("Clearly lower", "No drop") | Nobody can say afterwards whether the pilot passed | [Success has a number](CONTENT.md#success-has-a-number) |
| `for p in app.Presentations: if ...: target = p` without `break` | Picks the *last* matching presentation in enumeration order (effectively random when multiple match the substring) | [Multi-presentation safety](COM.md#multi-presentation-safety--never-trust-activepresentation) |

When one of these bites, fix it and **add a row here** if it's a new variant. The signal is: "I lost an hour to a silent failure" → it belongs in this table.
