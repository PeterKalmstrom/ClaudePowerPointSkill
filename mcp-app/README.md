# PowerPoint Live (MCP App)

An MCP server, written for this skill, that drives desktop PowerPoint on Windows through COM — and, in hosts that
support [MCP Apps](https://modelcontextprotocol.io/docs/extensions/apps) (Claude Desktop, claude.ai), shows the
**current slide live** next to the conversation while Claude works on it.

## Install

As part of the plugin it is registered automatically. By hand:

```powershell
claude mcp add --scope user powerpoint-live -- uvx --with "mcp<2" --with python-pptx --with pillow --with "pywin32; sys_platform == 'win32'" python <skill folder>\mcp-app\server.py
```

Needs Windows, desktop PowerPoint and [uv](https://docs.astral.sh/uv/). In a host without MCP Apps support
(e.g. the Claude Code terminal) the tools still work; there is just no live view.

## Tools

| Tool | Who | What |
|---|---|---|
| `powerpoint_open` | Claude | Open a deck (or attach to an open one) and pick the current slide |
| `powerpoint_run` | Claude | Run Python against the deck: `app`, `prs`, `slide`, `goto(n)`, `win32com` in scope; returns `print()` output |
| `powerpoint_show` | Claude, view | Make slide *n* current (the PowerPoint window follows) |
| `powerpoint_move` | Claude, view | Move a slide (by SlideID) before another slide or to the end of a section |
| `deck_outline` | view only | Sections and slides (id, title, hidden, change signature) for the sorter |
| `slide_thumbs` | view only | Small PNGs of up to 12 slides, cached by content |
| `slide_state` | view only | Deck, slide, count and a cheap change signature of the current slide |
| `slide_image` | view only | PNG of a slide |

## Views

- **Slide** — the current slide, large, with previous/next.
  - **Change highlights:** after any change (by Claude, a script or a person) the shapes that were added, changed
    or removed flash green, blue or red-dashed for a few seconds.
  - **Checks:** `lint_deck.py` findings drawn on the slide where they are, with a list below. **Fix** applies the
    automatic fix live (`powerpoint_fix`); **Ask Claude** sends the finding to the conversation.
  - **Before / after:** drag a divider across the slide to compare it with the version before its last change.
  - **Ask Claude:** click a shape or drag an area, type what you want; the request goes into the conversation with
    the slide, SlideID, shape id and box (hosts that can't take messages from a view get it on the clipboard).
  - **Accurate:** render through Save As (true embedded fonts and word breaks) instead of the fast `Slide.Export`.
  - **Design:** the current slide rendered in six design directions (`scripts/directions.json`; **More** pages
    through all of them). Each preview is made from a Save As copy restyled by `build_deck.py`'s
    `kDeckDesign.ApplyDirection` and drawn by a windowless PowerPoint, so the open deck is not touched. Click a
    preview twice to apply that direction (theme colours + heading/body fonts) to the whole deck.
  - **Theme:** the theme colour slots as swatches (click one to change it), the heading and body fonts, and the
    contrast pairs (text 4.5:1, accents as graphics or large text 3:1) with pass/fail, recomputed as you edit.
    **Apply to deck** writes them to every slide master.
  - **Layout variants:** when Claude has made variants of the current slide (`powerpoint_layout_variants`), a
    strip shows the original and each variant; pick one to compare it with the original on a slider, **Use this**
    (click twice) replaces the original, **Discard all** deletes the variants.
- **Slide sorter** — thumbnails grouped by section, with lint counts. Click to select (Ctrl/Shift for more),
  double-click to open, drag one or several onto a slide (goes in front) or a section's empty space (goes to its
  end). Hide/unhide, start a new section at the selection, rename a section (double-click its name), delete one
  (×, keeps its slides). Each slide shows its talk time and each section its total (see below).
- **Storyline** — the titles alone, by section, with title findings (label instead of claim, too long, missing,
  duplicate) and a talk-length estimate from the notes, plus a bar of the talk time by section against the target.
  Click a line to open the slide.
- **Rehearse** — the current slide large, its speaker notes, the next slide, an elapsed timer and a per-slide timer
  against the slide's plan, and an ahead / on time / behind chip. Start, Previous, Next, Black and End drive the
  real slide show in PowerPoint (`powerpoint_slideshow`); ←/→/Space and B work too.

**Talk length.** A slide's time is its timing marker when the notes open with one (`[15 sec]`, `[2 min]`, `[1:30]`
- `reference/PRESENTING.md`), else its notes words at 130 words a minute, else its visible words when it has no
notes. Hidden slides do not count. The target (default 20 min) is set in the sorter or storyline and remembered in
the browser.
- **History** — a version is saved before every change made through PowerPoint Live. Restore puts one back
  (the deck file is replaced and reopened; the state you leave is saved as a version first).

## Tools

| Tool | Who | What |
|---|---|---|
| `powerpoint_open` | Claude | Open a deck (or attach to an open one) and pick the current slide |
| `powerpoint_run` | Claude | Run Python against the deck: `app`, `prs`, `slide`, `goto(n)`, `win32com` in scope; returns `print()` output |
| `powerpoint_show` | Claude, view | Make slide *n* current (the PowerPoint window follows) |
| `powerpoint_move` | Claude, view | Move slides (by SlideID) before a slide or to the end of a section |
| `powerpoint_sections` | Claude, view | Add, rename or delete sections |
| `powerpoint_hide` | Claude, view | Hide or unhide slides |
| `powerpoint_fix` | Claude, view | Apply a lint fix to one shape, live |
| `powerpoint_history` / `powerpoint_restore` | Claude, view | List saved versions / put one back |
| `powerpoint_resume` | Claude, view | Re-arm the server after an unexpected error halted it |
| `powerpoint_send_error_report` | Claude, view | Answer a waiting error report - `send=true` only after the user said yes to "Do you want to send this error message?" |
| `error_reports`, `error_report_answer` | view only | Error reports waiting for the person's Yes/No ("Do you want to send this error message?"); nothing is sent without Yes |
| `powerpoint_slideshow` | Claude, view | `action`: start (from the current slide), next, previous, black (toggle), end, state - the real slide show |
| `powerpoint_apply_direction` | Claude, view | Restyle the whole deck in a design direction (theme colours + fonts, as `build_deck.py` writes them) |
| `powerpoint_set_theme` | Claude, view | Change theme colour slots (`colors`: slot -> RRGGBB) and the major/minor fonts on every slide master |
| `powerpoint_layout_variants` | Claude | How to make layout variants of a slide: Claude duplicates it with `powerpoint_run` as hidden slides right after it, tagged `PPTLIVE-VARIANT-OF=<SlideID>` and `PPTLIVE-VARIANT-NAME` |
| `powerpoint_choose_variant` / `powerpoint_discard_variants` | Claude, view | Replace the original with a variant (the others are deleted) / delete the variants |
| `slide_state`, `slide_image`, `deck_outline`, `slide_thumbs`, `deck_lint`, `history_thumb` | view only | What the views draw (`deck_outline` carries per-slide `visible_words`, `estimate_sec`, `marker_sec`, `planned_sec` and section / deck totals) |
| `design_previews`, `theme_info`, `layout_variants` | view only | Direction previews from a copy; theme slots, fonts and contrast pairs; a slide's variants with thumbnails |

Every theme, direction and variant change saves a version first, so History can undo it. Slide-show actions do
not change the deck and save no version.

Expected states (no deck open, file not found, not on Windows, a shape that is gone) come back as tool errors with
a plain message. The first *unexpected* error is reported to the global error handler (`scripts/kShared.py`) and
halts the server: every tool then answers "PowerPoint Live halted after an error in ..." until a person presses
**Resume** in the view (or Claude calls `powerpoint_resume`).

Live fixes: `unused_placeholder`, `body_below_floor`, `text_overflow` (shrinks using PowerPoint's own text layout,
never below the floor), `picture_stretched`, `a11y_missing_alt_text` (charts and tables; pictures need Claude),
`chart_default_palette`, `chart_legend_steals_plot`, `chart_accounting_zero_dash`.

The view polls a cheap change signature about once a second (every two in the sorter and storyline) and fetches
pictures only when something changed. Pictures go to the view, not to Claude, so the live preview costs no tokens.
The fast preview uses `Slide.Export`, which shows fallback fonts for embedded fonts — tick **Accurate** or use
`scripts/render_slides.py` to check a slide closely. Versions live in a temporary folder for the session.

Manual test of the panel in Claude Desktop: [`docs/powerpoint-live-checklist.md`](../docs/powerpoint-live-checklist.md).
