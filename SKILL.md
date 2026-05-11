---
name: configuring-powerpoint-mcp
description: Configure PowerPoint MCP and the rich-media presentation workflow — uv/uvx prerequisite, claude mcp add, troubleshooting, plus Remotion (animated video), Nanobanana (AI image generation), and Veo 3.1 (AI video generation) integration for professional slides. Load when configuring PowerPoint MCP, creating presentations, embedding video/images/AI-generated video, or diagnosing why PowerPoint tools are missing.
---

# PowerPoint MCP — Setup & Rich Media Workflow

Installs `powerpoint-mcp` so Claude Code can open, read, edit, and create PowerPoint presentations via COM automation on Windows. Also covers integrating **Remotion** (programmatic animated video), **Nanobanana** (AI image generation), and **Veo 3.1** (AI video generation from text prompts) for professional-quality slide media.

## Contents

- Prerequisites
- Installation
- Troubleshooting
- COM patterns & safety
- Workflow
- Available Tools
- Quick Verification
- Nanobanana Integration (AI Images)
- Remotion Integration (Animated Video)
- Veo Integration (AI Video Generation)
- Embedding Media in Slides
- Combined Workflow Patterns
- Speaker notes — load them up
- Showcase-first for multi-slide sections
- Common defects to self-check
- Anti-patterns (recurring COM / build traps)

---

## Prerequisites

- **Windows** — the server uses PowerPoint COM automation; it does not work on macOS or Linux.
- **Microsoft PowerPoint** — must be installed (desktop version, not web-only).
- **uv** (Python package runner) — provides the `uvx` command used to launch the server.

### Install uv

Open a **PowerShell** terminal (not bash):

```powershell
irm https://astral.sh/uv/install.ps1 | iex
```

This installs `uv.exe` and `uvx.exe` to `C:\Users\<USER>\.local\bin\`.

> **Note:** Replace `<USER>` with the actual Windows username throughout this guide (e.g., `C:\Users\alice\.local\bin\uvx.exe`).

Verify:

```powershell
C:\Users\<USER>\.local\bin\uvx.exe --version
```

---

## Installation

### Step 1 — Add the MCP server at user scope with the full path to uvx

```bash
claude mcp add --scope user powerpoint -- "C:\Users\<USER>\.local\bin\uvx.exe" powerpoint-mcp
```

**Critical details:**

- **Use `--scope user`**, not the default project scope. Project-scoped MCP servers are nested inside a project key in `.claude.json` and the VS Code extension may not load them.
- **Use the full absolute path to `uvx.exe`**, not just `uvx`. Claude Code's shell environment does not reliably include `~/.local/bin` in PATH, so a bare `uvx` command will fail silently — the server won't start and no tools will appear.

The command writes to the top-level `mcpServers` block in `C:\Users\<USER>\.claude.json`:

```json
{
  "mcpServers": {
    "powerpoint": {
      "type": "stdio",
      "command": "C:\\Users\\<USER>\\.local\\bin\\uvx.exe",
      "args": ["powerpoint-mcp"],
      "env": {}
    }
  }
}
```

### Step 1b — Pre-download the package

The first launch downloads ~45 packages. Run this once to cache them so the MCP server starts instantly on first real use:

```bash
"C:\Users\<USER>\.local\bin\uvx.exe" powerpoint-mcp --help
```

### Step 1c — Alternative: edit `.claude.json` directly

If the `claude` CLI is not available (e.g., running inside the VS Code extension), add the MCP server by editing `C:\Users\<USER>\.claude.json` directly. Add the `mcpServers` block at the **top level** of the JSON (not nested inside `projects`):

```json
{
  "mcpServers": {
    "powerpoint": {
      "type": "stdio",
      "command": "C:\\Users\\<USER>\\.local\\bin\\uvx.exe",
      "args": ["powerpoint-mcp"],
      "env": {}
    }
  }
}
```

### Step 2 — Restart Claude Code

The MCP server is only discovered at session startup. Close and reopen Claude Code (or the VS Code window) after adding the server.

### Step 3 — Verify connection

After restarting, run:

```bash
claude mcp list
```

Expected output:

```
powerpoint: C:\Users\<USER>\.local\bin\uvx.exe powerpoint-mcp - ✓ Connected
```

If you see `✓ Connected`, the PowerPoint tools are available in the session.

---

## Troubleshooting

### Tools don't appear after restart

1. **Check `claude mcp list`** — if the server isn't listed at all, the config wasn't saved. Re-run the `claude mcp add` command.
2. **Server listed but not connected** — run `uvx.exe` manually to check for errors:
   ```bash
   "C:\Users\<USER>\.local\bin\uvx.exe" powerpoint-mcp
   ```
   The first run downloads ~45 packages; subsequent runs are instant.
3. **Server connected in CLI but tools missing in VS Code** — the server was likely added at project scope instead of user scope. Check `.claude.json`:
   - **Wrong:** `mcpServers` nested inside `"projects"` → `"<path>"` → `"mcpServers"`
   - **Right:** `mcpServers` at the top level of the JSON file
   
   Fix: remove the project-scoped entry and re-add with `--scope user`.

### `uvx` not found

The `uv` installer wasn't run, or PATH wasn't updated. Use the full path `C:\Users\<USER>\.local\bin\uvx.exe` in the MCP config instead of relying on PATH.

### PowerPoint COM errors

- Ensure PowerPoint is installed (not just Office web apps).
- Close any "PowerPoint has stopped working" dialogs.
- If PowerPoint is open with a modal dialog (e.g., save prompt), the COM connection may hang.

### Images get swallowed by content placeholders

When adding a picture with `AddPicture` to a slide that uses a layout with a content placeholder (e.g., "Title and Content"), PowerPoint may absorb the image into the placeholder instead of creating a standalone Picture shape. Deleting the placeholder doesn't help — the layout regenerates it.

**Fix:** Change the slide layout to **"Title Only"** before adding the image:

```python
# In evaluate tool:
title_only = presentation.SlideMaster.CustomLayouts(6)  # "Title Only"
slide.CustomLayout = title_only
# Now AddPicture creates a real Picture shape (type=13)
slide.Shapes.AddPicture(path, False, True, left, top, w, h)
```

This prevents the layout from regenerating content placeholders. The title placeholder is preserved. Use this whenever you need standalone images on slides that originally had content placeholders.

### Safe destructive operations — snapshot first

**Rule:** Before any destructive change to a deck the user cares about (deleting shapes, replacing backgrounds, mass-editing slides, layout changes), write a timestamped side copy with `SaveCopyAs`. PowerPoint's built-in undo covers the current session but cannot rescue you after a save, across sessions, or when COM operations fail silently.

**Why:**
- `ExecuteMso("Undo")` works for most shape edits but does NOT cover media inserts, layout changes, or some COM-driven modifications
- Once you call `presentation.Save()`, the on-disk file is already overwritten — session undo can restore in-memory state but the file has moved
- Closing PowerPoint clears the undo stack forever
- Undo is strictly linear: you cannot undo one change and keep later ones

**How:**

```python
# In evaluate tool — before any risky operation:
import time
src = presentation.FullName
backup = src.replace(".pptx", f".backup-{int(time.time())}.pptx")
presentation.SaveCopyAs(backup)
# Now do the risky thing. If it goes wrong, the backup is on disk.
```

`SaveCopyAs` writes a side copy without affecting the open document — you keep working in the original. If something goes wrong, close the original without saving and reopen the backup.

**When to use:**
- Deleting shapes you cannot easily recreate (AI-generated images, complex diagrams)
- Replacing a slide's layout
- Bulk operations across many slides
- Any time the user says "my big deck" or "the real one" — assume it's irreplaceable
- Before the first destructive operation in a new session on a user's working deck

**When session undo is enough:**
- Text edits on placeholders
- Moving or resizing shapes
- Adding new content (if wrong, just delete it)
- Any reversible tweak within the current session

**Pattern:** Think of `SaveCopyAs` as a git commit and `ExecuteMso("Undo")` as editor undo. Use both.

---

## COM patterns & safety

Foundational rules for any COM script targeting PowerPoint. These are not "things to do when something breaks" — they are constraints to design around from the first line of code. Read this section before writing any non-trivial COM patch or build script.

### Snapshot before any risky bulk edit

**Rule:** Before running any script that touches more than a handful of shapes — bulk font bumps, layout reflows, multi-slide rebuilds — copy the deck to a timestamped backup. COM operations succeed without raising errors when they corrupt the wrong shapes; the backup is your only undo.

The cheapest pattern is a file-copy (no COM), so the user's open PowerPoint session is undisturbed:

```python
import shutil, datetime, pathlib
src = pathlib.Path(r"C:/path/to/deck.pptx")
ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
dst = src.with_name(f"{src.stem}.{ts}.pre-bulk-fontbump{src.suffix}")
shutil.copy2(src, dst)
```

Save in PowerPoint first if there are unwritten changes — the file-copy reads disk, not in-memory state.

When to snapshot:
- Before any bulk script that walks all slides
- Before a deep-rebuild of a large slide (5+ shapes)
- Before final pre-talk edits (label `pre-rehearsal`)
- After a known-good iteration — gives you a "last clean version" rollback

### Multi-presentation safety — never trust `ActivePresentation`

**Rule:** When more than one presentation is open in PowerPoint, `app.ActivePresentation` and the MCP tools that depend on it (`slide_snapshot`, `add_speaker_notes`, etc.) can silently target the wrong file. Other windows (a separate deck, a template, a reviewer's copy) can grab focus and flip the active selection without warning.

**Why this happens:**
- The MCP server resolves the active presentation at the moment each call is made.
- A user clicking another PowerPoint window — or even the OS focus changing — can move the active selection.
- COM operations on the wrong presentation will succeed (no error) and corrupt the wrong deck.

**How to apply:**
Always select the target presentation by name match in COM scripts:

```python
import win32com.client
app = win32com.client.GetActiveObject('PowerPoint.Application')
target = next((p for p in app.Presentations if "MyDeck" in p.Name), None)
if not target:
    raise RuntimeError("Target presentation not open")
slide = target.Slides(34)        # use `target`, never `app.ActivePresentation`
target.Save()
```

When using MCP tools (which target ActivePresentation), close all other PowerPoint windows first, or activate the right window with `target.Windows(1).Activate()` before the call. Verify with `slide_snapshot` after activation.

### Idempotent build scripts

**Pattern:** For every non-trivial slide, write a `build_slide_NN.py` script that clears all shapes and rebuilds from scratch. Save these scripts alongside the deck. They make iteration cheap.

```python
def clear_and_bg(slide, img):
    slide.CustomLayout = pres.Designs(1).SlideMaster.CustomLayouts(6)
    for sh in list(slide.Shapes):
        try:
            sh.Delete()
        except:
            pass
    slide.Shapes.AddPicture(img, 0, -1, 0, 0, 960, 540).Name = "BgImage"
```

Each run produces the same result. To tweak typography or layout, edit the script and re-run — no manual element-by-element fixing in PowerPoint.

**Strongly prefer rebuild over patch.** Once a slide has 5+ shapes and you need a layout change touching more than two elements, rebuild from scratch is almost always cheaper and safer than a patch script. Patches drift — backing rectangles get left in old positions while text moves, shape Z-order silently shifts, conditional filters miss shapes. A rebuild is a single source of truth.

### Filtering shapes — `HasTextFrame` is NOT "is this a text shape"

**Rule:** Don't use `if not sh.HasTextFrame` to mean "this is a rectangle / image / decorative shape". AutoShape rectangles return `HasTextFrame == True` even when they hold no text — they have a text frame, it's just empty. A filter like:

```python
for sh in slide.Shapes:
    if not sh.HasTextFrame:
        # delete or resize the rectangle
```

…will silently skip every AutoShape rectangle on the slide. Layout patches that depend on this filter fail silently.

**Reliable filters:**

```python
# By emptiness — true rectangles vs text shapes
if sh.HasTextFrame and not sh.TextFrame.HasText:
    # AutoShape with no text — likely a backing rect
    ...

# By shape type
# 1 = msoAutoShape, 13 = msoPicture, 17 = msoTextBox, 14 = msoPlaceholder
if sh.Type == 1:
    # AutoShape (rectangle, oval, etc.)
    ...
elif sh.Type == 17:
    # TextBox
    ...

# By name (most reliable when you control the build script)
if sh.Name == "HeroBacking":
    ...
```

**Pattern:** when building a slide programmatically, **name every shape you might need to find later**. `slide.Shapes.AddShape(...).Name = "HeroBacking"`. Then patches use `sh.Name == "HeroBacking"` — survives shape additions, doesn't depend on position or order.

### Move text and its backing together

**Rule:** TextBox and its backing AutoShape are **independent shapes** in COM. Repositioning one does not move the other. Patches that shift text without shifting the backing leave text floating outside its container — sometimes even off-slide.

**Bad:**
```python
# Shift text right; backing stays put → text floats outside its card
for sh in slide.Shapes:
    if sh.HasTextFrame and "vetat" in sh.TextFrame.TextRange.Text:
        sh.Left = 400  # text moves
        sh.Width = 500
# (forgot to move the backing rectangle behind it)
```

**Good — group, then move:**
```python
# Group the backing and text into one Shape, then move/resize as a unit
shape_range = slide.Shapes.Range([backing.Name, text.Name])
group = shape_range.Group()
group.Left = 400
```

**Or — move both explicitly in the same patch:**
```python
backing.Left = 380; backing.Width = 540
text.Left    = 400; text.Width    = 500   # 20px text margin inside the card
```

**Best — rebuild rather than patch.** See "Idempotent build scripts" above. If you find yourself patching text and backing positions, you've outgrown the patch — rewrite the build script.

---

## Workflow

The five-step iteration loop that catches the silent rendering failures the troubleshooting rules above describe. Read this section before starting any slide work, not after something goes wrong.

### Iteration loop — build, render, LOOK, critique, fix

**Rule:** every slide edit follows the same five-step loop. Step 3 is the one that gets skipped. Don't skip it.

```
1. Build / edit (idempotent script)
2. Render to PNG via slide.Export(out, "PNG", 1280, 720)
3. LOOK at the actual rendered PNG  ←  DO NOT SKIP
4. Self-critique against rules (see "Common defects to self-check" section)
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

- Check if the rule itself is wrong for this anchor type (gallery, knowledge graph, quote — see "Anchor types and word budgets")
- Switch from patch to full rebuild — patches drift after the third one
- Ask the user — the expected output may itself be wrong

Never run more than 4 iterations on the same defect without escalating.

**Anti-patterns to recognise:**

- *Source-look-only iteration* — reading COM properties or script values, declaring success without exporting
- *Patch on patch on patch* — each patch tweaks one element without considering layout interactions; backing rectangles drift, z-order shifts, third patch produces shape soup. **Three patches = rewrite the build script.**
- *Tool-output as ground truth* — `{"success": true}` from an MCP tool is not the same as a correctly rendered slide
- *Optimistic font sizing* — "36pt should fit" is a guess until the PNG confirms it
- *Caching trust* — replacing an image at the same path doesn't always update the embedded version; re-add the picture explicitly
- *Active context drift* — see "Multi-presentation safety" above; pin to a specific presentation by name, never trust `ActivePresentation`

### Text wrap — anticipate it, don't trust visual review

**Rule:** PowerPoint TextBoxes word-wrap by default. A headline at 52pt that "should fit" on one line will silently wrap to two lines if it exceeds the container width — and the wrap is invisible until export. Always plan for the worst-case character count and font width.

**Practical ceilings (Aptos Display, bold, single-line):**

| Container width | Safe size | Approx. character ceiling |
|---|---|---|
| 540 px | 36pt | ~22 chars |
| 540 px | 32pt | ~25 chars |
| 540 px | 28pt | ~28 chars |
| 800 px | 44pt | ~25 chars |
| 800 px | 38pt | ~28 chars |
| 800 px | 32pt | ~33 chars |

Aptos Display is wider than Inter or Calibri. Other fonts shift the ceiling.

**Defenses against silent wrap:**

```python
# 1. Disable wrap — overflow is preferred over silent wrap
text.TextFrame.WordWrap = 0  # msoFalse

# 2. Reduce internal margins (default ~7pt each side)
text.TextFrame.MarginLeft = 4
text.TextFrame.MarginRight = 4

# 3. Always export and look at the actual rendering before declaring done
slide.Export(out_path, "PNG", 1280, 720)
```

**Pattern:** for any headline ≥30pt, export and visually verify that no text wrapped. Don't trust the COM property values — only the rendered PNG tells you what the audience sees.

---

## Available Tools

Once connected, these tools become available:

| Tool | Purpose |
|---|---|
| `manage_presentation` | Open, close, create, save, save_as presentations |
| `slide_snapshot` | Capture slide content + annotated screenshot |
| `switch_slide` | Navigate to a specific slide |
| `manage_slide` | Duplicate, delete, or move slides |
| `populate_placeholder` | Write text/HTML/LaTeX/images into placeholders |
| `add_speaker_notes` | Add or replace speaker notes on a slide |
| `add_animation` | Add entrance animations (fade, fly, wipe, zoom) |
| `add_slide_with_layout` | Insert a new slide using a template layout |
| `analyze_template` | Inspect template layouts and placeholders with screenshots |
| `list_templates` | Discover available PowerPoint templates |
| `evaluate` | Execute arbitrary Python code in the PowerPoint COM context |

### Common workflow

```
1. manage_presentation  action="open"  file_path="C:\\path\\to\\deck.pptx"
2. slide_snapshot        slide_number=1    → read content + see screenshot
3. populate_placeholder  placeholder_name="Title 1"  content="New Title"
4. manage_presentation  action="save"
```

### Bulk reading the deck

For analyzing a deck as a whole (critiquing content, building a relationship map, counting words, finding contradictions), `slide_snapshot` is too slow — one round trip per slide. Instead, write a small Python script that opens the presentation once via `win32com.client`, walks every slide, and dumps all text + speaker notes to JSON in a single COM session (~6s for 70 slides vs. minutes one slide at a time).

**When to use bulk read vs slide_snapshot:**
- **Bulk read script**: analyzing ≥5 slides, reviewing whole sections, building any cross-slide reasoning
- **slide_snapshot**: editing or inspecting a single slide with visual reference

### Auditing a deck

For a one-shot quality check across an entire deck (word counts per slide, minimum body font size, taste defects), write a script that walks every slide and emits a flat report flagging slides that violate the per-anchor word budget or the 18pt body floor.

Useful columns: slide #, word count, min body font size, title, status. Statuses like `OK`, `ok-tight` (within budget but close), `** AUDIT` (needs review), `[skip]` (template / cover / vote slides) make it easy to scan.

Run this before declaring any deck done. Treat each `** AUDIT` flag as a defect to either fix or document as an accepted anchor-type exception.

---

## Quick Verification

After setup, test with:

```
Open any .pptx file in the workspace using manage_presentation,
then run slide_snapshot on slide 1.
```

If you get slide content and a screenshot, the setup is complete.

---

## Nanobanana Integration (AI Images)

Nanobanana is an MCP server for AI image generation powered by Gemini models. It produces **raster images only** (PNG/JPG) — no vectors, no animation.

### Wrapping Python helpers via subprocess + uvx

When a project script needs to generate images or videos (e.g., a slide builder that needs a backdrop first), **call a Python helper as a subprocess via `uvx`** rather than importing it. The helpers depend on `google-genai` / `pywin32`, which most projects do not want to add as a hard dependency. `uvx` provisions the dependency per call and tears it down — no virtualenv setup, no project-wide pollution.

**Single-shot pattern** (one image, one video):

```python
import subprocess, sys
HELPER = r"path\to\your\batch_image_gen.py"
sys.exit(subprocess.run([
    "uvx", "--with", "google-genai", "python", HELPER,
    "--prompt", "Cinematic photograph of ... no text, no logos.",
    "--out", r"C:\out\backdrop.png", "--aspect", "16:9",
]).returncode)
```

**Batch pattern** (many jobs from a JSON spec written to a temp file):

```python
import json, os, subprocess, sys, tempfile
HELPER = r"path\to\your\batch_image_gen.py"
spec = [{"prompt": "...", "out": r"C:\images\one.png"}, ...]
with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
    json.dump(spec, f, ensure_ascii=False); spec_path = f.name
try:
    rc = subprocess.run(["uvx", "--with", "google-genai", "python", HELPER,
                         "--spec", spec_path, "--quiet"]).returncode
finally:
    os.unlink(spec_path)
sys.exit(rc)
```

The same pattern works for any Veo / video helper (use `uvx --with google-genai`) and any future helper script. Pass `--quiet` when calling from a larger pipeline so per-job progress doesn't pollute your build script's output (errors still go to stderr).

### Capabilities

- Photo-realistic and stylized image generation
- Image editing and modification
- 4K resolution output
- Google Search grounding for reference accuracy
- Subject consistency across multiple generations

### When to use

- **Background images** for slides (dramatic, atmospheric visuals)
- **Illustrations** that would take too long to source manually
- **Concept art** for topics that need emotional impact
- **Before/after** comparison images

### When NOT to use

- Animatable graphics (use Remotion or native PowerPoint shapes instead)
- Diagrams with precise labels/arrows (text rendering is unreliable)
- Vector graphics or SVGs (not supported)
- Charts or data visualizations (use PowerPoint native charts)

### On-slide text budget: ~10 words max (excluding title)

**Rule:** For spoken/live presentations, limit visible text on each slide to roughly 10 words beyond the title. If a slide is on screen for ~20 seconds, the audience can spend at most ~2 seconds reading — the remaining ~18 seconds belong to the speaker. More text than that forces the audience to choose between reading and listening.

**Why:**
- Audiences read silently at ~200–250 words/min, so 10 words ≈ 2 seconds
- Dense slides compete with the speaker's voice — attention splits and retention drops
- Presenters fall into "reading their slides" when the deck carries the content
- Visual impact (image, chart, big number) scales with the amount of whitespace around it

**How to apply:**
- Title: one phrase (not counted in the 10)
- Body: one short headline OR a single big number OR one chart — not all three
- All supporting detail, data, sources, and narrative goes into **speaker notes** (notes can be arbitrarily long — load them up)
- If you need more text, make it a new slide
- Exceptions: see anchor-type budgets below

**Pattern:** Pair a strong visual (full-bleed image or chart overlay) with a single punchy headline ≤10 words. Everything else lives in notes.

### Anchor types and word budgets

The 10-word rule is the default for **spoken hero slides**. It does not apply uniformly — different anchor types have different read patterns. Use the budget that matches the anchor:

| Anchor type | Visible word budget (excl. title) | Why |
|---|---|---|
| Hero stat (single big number) | 5 | One thing to absorb |
| Hero + sub (headline + tagline) | 10 | Statement + context |
| Comparison pair (A vs B) | 12 | Two parallel things |
| 3–5 card gallery | ~25 (8/card) | Audience reads one card at a time |
| Vulnerability matrix / table | unlimited | Cells are labels, not body |
| Knowledge graph / cascade diagram | 40+ | Relationships matter more than label content |
| Animated list (sequential reveal) | unlimited | Each item appears alone |
| Quote anchor | length of quote | Single coherent phrase, read as a unit |
| Chart slide | 8 + chart-internal labels | The chart is the message |
| Reference / agenda / rules | 30+ | Audience studies it; not spoken over |

**How to identify the anchor type:** look at what dominates the slide visually. If three icons sit side by side, it's a gallery. If a single 100pt number, it's hero stat. If a network of connected cards, it's a knowledge graph.

**Audit caution:** automated word-count checks will false-positive on gallery / matrix / knowledge-graph slides. Always confirm against the anchor type before "fixing" by stripping content.

### Label exception (12–14pt allowed for non-body text)

The 18pt slide floor (see Typography below) applies to **body text** — sentences the audience reads. It does not apply to **labels** — short identifying tags attached to a visual.

**Label allowed at 12–14pt:**
- Chart axis labels, tick labels, legend labels
- Caption directly under a flag, icon, or photo
- Row/column labels in a matrix or table
- Pill or chip text in a dense gallery
- Source attributions and citations (≥12pt; 18pt is preferred)

**Body text — must be ≥18pt:**
- Headlines, taglines, sentence fragments
- Bullet items the audience is meant to read
- Quotes
- Anything that appears as a sentence

**Test:** if removing the visual the text describes leaves you with a meaningful sentence, it's body — bump to 18pt minimum.

### Hero stat pattern

Recurring recipe for "single big number" slides. Use this typography hierarchy:

| Element | Size | Notes |
|---|---|---|
| The number | 80–150pt bold | Accent color, dominant |
| Unit (immediately below number) | ≥24pt bold | Same color or white |
| Caption (below unit) | ≥18pt italic muted | Optional context |
| Source / attribution | ≥18pt italic small | Bottom of card or slide |

Anti-pattern: number at 100pt + unit at 14pt — the unit becomes invisible against the number.

### Two-column comparison pattern

For "humans vs AI", "before vs after", "today vs projected" slides:

```
Layout: two equal-width cards side by side, gap 30px
Left card: cool/warm border (e.g., red) — represents the problem
Right card: warm/cool border (e.g., cyan) — represents the answer
Card height: 280–340px depending on row count
Card structure:
  - Title row: 22pt bold, accent color
  - Subtitle row: 14pt italic muted ("— styrs av:" or "— bryr sig om:")
  - 4–6 bullet rows: 18pt body, white
Bottom strip below both cards: single bold 20–24pt line with the synthesis
```

**Word count:** ≤6 bullets per card × ~3 words each + 1 synthesis line ≈ 35–40 words across the slide. Acceptable as "comparison pair" anchor (each side read as one unit).

### Balanced layout — fill the available space

**Rule:** Content should fill the slide's usable area with visually balanced margins. Don't let content hug the top and leave a large void at the bottom (or any other side). Scale element size and typography proportionally to available space.

**Why:**
- Unbalanced layouts look unfinished — the audience reads "draft" or "amateur" before they read the content
- Top-heavy slides leave the visual center of gravity wrong; the eye keeps looking for something that isn't there
- A slide element sized for a cramped layout looks small and weak when the canvas around it is actually large
- Matching gaps (top/bottom/between) reads as intentional; uneven gaps read as mistakes
- Bigger elements let you use bigger, more impactful typography

**How to apply — sizing:**
1. Measure the usable area: slide height minus title bar (e.g., 540 − 90 = 450 tall content area)
2. Count the rows/elements and the gap count (for N rows there are N+1 gaps — one above row 1, one below row N, and N−1 between)
3. Decide gap size (typically 15–25 points) and divide the remaining height among the elements
4. Apply the same formula horizontally for multi-column layouts

**How to apply — typography:**
- When elements grow, grow the font too. A 120-point-tall row can carry 34pt text; a 75-point row only 26pt.
- Titles: 36–54pt depending on length
- Body labels on cards: 28–38pt for impact
- Secondary labels / captions: 14–20pt
- Body text in content slides: 18–24pt (still 2-second rule applies)

**Pattern — three-row, two-column grid (generic example):**
```
Slide: 960 × 540, title panel 0–90
Content area: y=90 to y=540 = 450 tall
Target: 3 rows × 2 columns of cards, icon + label

Rows: 3 rows of height R, 4 gaps of height G
    3R + 4G = 450
    Pick G = 22, R = 120 → 360 + 88 = 448 (2px slack)

Columns: margin M, icon I, gap g, label L, inner gap IG, then repeat
    2M + 2I + 2g + 2L + IG = 960
    M=20, I=100, g=10, L=330, IG=40 → 40 + 200 + 20 + 660 + 40 = 960 ✓

Row y positions: 112, 254, 396
Col 1 icon x=20, label x=130
Col 2 icon x=500, label x=610
```

**Sanity-check after any layout change:**
- Top gap above first row ≈ bottom gap below last row (within 10-15 points)
- Inter-row gaps are equal
- No large void in any quadrant
- Elements proportional to their container (a 120-tall card with 14pt text looks empty)
- Take a `slide_snapshot` and look at it — if you spot a visual imbalance, fix it

**When to break the rule:**
- Deliberate tension or asymmetry (a single big number in one quadrant, whitespace elsewhere for emphasis)
- Reference slides that follow a fixed template across a series
- Full-bleed images where the composition is the layout

### Generate images WITHOUT text — overlay text in PowerPoint

**Rule:** Always generate images and video without text. Overlay all text in PowerPoint using semi-transparent dark backing rectangles for readability.

**Why:**
- AI image generators (Nanobanana, Gemini) frequently misspell text, render it in the wrong language, or warp it
- Generated text is locked into the raster image — you can't fix typos, change language, or update numbers
- PowerPoint text overlays are easy to edit, translate, restyle, and reposition
- Text overlays render crisply at any scale; embedded text gets blurry when resized
- Same image asset can be reused with different captions across slides

**How:**
1. Prompt with explicit "NO TEXT, NO LABELS, NO WORDS, NO LETTERS, NO NUMBERS anywhere in the image"
2. Add the prompt's negative_prompt: `text, words, letters, numbers, labels, captions, watermark`
3. Use AutoShape (rectangle) with `Fill.Transparency = 0.4` as backing for the text
4. Add `TextBox` shapes positioned over the image with the actual labels in the target language

**Hard rule on retries:** if the generator inserts text after 2 reprompts, give up. Do not iterate a third time hoping for a clean image — accept the spurious text or switch to a different background. Reroll cost adds up fast and the rate of success usually doesn't improve. Safer fallback: regenerate as an abstract texture (no objects = no labels), or fetch a stock image.

**Example overlay pattern:**
```python
# Dark semi-transparent backing
bg = slide.Shapes.AddShape(1, x, y, w, h)  # 1 = msoShapeRectangle
bg.Fill.Solid()
bg.Fill.ForeColor.RGB = 0x000000
bg.Fill.Transparency = 0.4
bg.Line.Visible = False
bg.ZOrder(3)  # send backward (behind text)

# Text on top
txt = slide.Shapes.AddTextbox(1, x, y, w, h)
txt.TextFrame.TextRange.Text = "Swädish text"
txt.TextFrame.TextRange.Font.Size = 16
txt.TextFrame.TextRange.Font.Bold = True
txt.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF  # white in BGR
```

### Workflow

1. **Search first** — check if Nanobanana can find relevant reference imagery
2. **Generate** with a detailed prompt describing the visual — explicitly request NO TEXT
3. **Upload** the generated image to the project folder
4. **Embed** in PowerPoint via `evaluate` → `AddPicture`
5. **Overlay text** using TextBox + transparent backing rectangle

### Image embedding in PowerPoint

```python
# In evaluate tool:
pic = slide.Shapes.AddPicture(
    r"C:\path\to\image.png",
    False,   # LinkToFile
    True,    # SaveWithDocument
    left, top, width, height
)
```

> **Gotcha:** See the "Images get swallowed by content placeholders" section above if the image disappears into a placeholder.

---

## Remotion Integration (Animated Video)

Remotion is a React-based framework for creating programmatic video. Output is MP4 video that embeds directly into PowerPoint slides. This is the recommended tool when native PowerPoint animations are not polished enough.

### When to use Remotion over native PowerPoint animations

| Use case | Tool |
|---|---|
| Simple entrance effects (fade, fly, zoom) | PowerPoint `add_animation` |
| Sequential bullet reveals | PowerPoint `add_animation` with `by_paragraph` |
| Complex diagrams with glowing effects, particles, curves | **Remotion** |
| Network graphs, cascade visualizations, feedback loops | **Remotion** |
| Cinematic transitions, atmospheric builds | **Remotion** |
| Data-driven animations (charts morphing, counters) | **Remotion** |

### Prerequisites

- **Node.js** v18+ — check with `node --version`
- **Google Chrome** installed (used as the rendering engine)

> **PATH issue:** Claude Code's bash shell may not have Node.js in PATH even when installed. Use the full path: `export PATH="/c/Program Files/nodejs:$PATH"` at the start of every bash command.

### Project setup (one-time per animation)

```bash
export PATH="/c/Program Files/nodejs:$PATH"
mkdir -p "path/to/project"
cd "path/to/project"
npm init -y
npm install remotion @remotion/cli @remotion/player react react-dom typescript @types/react @types/react-dom
```

### Required files

Every Remotion project needs three files minimum:

**`src/index.ts`** — Entry point:
```typescript
import { registerRoot } from "remotion";
import { RemotionRoot } from "./Root";
registerRoot(RemotionRoot);
```

**`src/Root.tsx`** — Composition registry:
```typescript
import { Composition } from "remotion";
import { MyAnimation } from "./MyAnimation";

export const RemotionRoot: React.FC = () => (
  <Composition
    id="MyAnimation"
    component={MyAnimation}
    durationInFrames={300}  // 10 seconds at 30fps
    fps={30}
    width={1920}
    height={1080}
  />
);
```

**`src/MyAnimation.tsx`** — The actual animation (React component using Remotion hooks).

### Key Remotion APIs

| API | Purpose |
|---|---|
| `useCurrentFrame()` | Current frame number (0-based) |
| `useVideoConfig()` | Get fps, width, height, duration |
| `interpolate(frame, inputRange, outputRange, options)` | Map frame to any value (position, opacity, scale) |
| `spring({ frame, fps, config })` | Physics-based spring animation |
| `<AbsoluteFill>` | Full-frame container |
| `<Sequence from={60}>` | Delay child content to start at frame 60 |

### Rendering to MP4

```bash
export PATH="/c/Program Files/nodejs:$PATH"
cd "path/to/project"
npx remotion render src/index.ts MyAnimation out/animation.mp4 \
  --browser-executable="/c/Program Files/Google/Chrome/Application/chrome.exe"
```

**Critical:** Always pass `--browser-executable` pointing to the local Chrome installation. Without it, Remotion downloads Chrome Headless Shell (~108 MB) which often fails or times out, especially on cloud-synced folders (Dropbox, OneDrive).

### Rendering performance

- 300 frames (10s at 30fps) renders in ~90 seconds
- Remotion uses 8x concurrency by default
- Output size is typically 2–3 MB for motion graphics

### Animation design patterns

**Hub-and-spoke with growing center:**
```
Center node appears → Group 1 nodes + arrows fade in → arrow returns to center →
center grows → Group 2 appears → center grows more → ... → intensifying finale
```

**Progressive reveal:**
```
Title fades in → elements appear one by one with staggered timing →
connections draw between them → final state holds
```

**Key techniques for impact:**
- `spring()` for organic, bouncy entrances — feels alive
- Glowing SVG filters (`feGaussianBlur` + `feMerge`) for dramatic nodes
- Flowing particles along paths (quadratic bezier interpolation)
- `radial-gradient` backgrounds with vignette for atmosphere
- Vignette overlay that intensifies toward the end for urgency
- Dynamic sizing (nodes/elements that grow over the animation to show escalation)

### Iteration workflow

Remotion renders are fast enough to iterate:

1. Write/modify the `.tsx` animation code
2. Render to MP4 (~90s)
3. Replace the video on the slide via `evaluate`
4. Save the presentation
5. Preview in PowerPoint presentation mode (F5)
6. Repeat

---

## Veo Integration (AI Video Generation)

Google Veo 3.1 generates photorealistic video from text prompts via the Gemini API. This is the missing piece that enables true "prompt to video in a slide" — no stock footage, no manual animation, just describe what you want and get an MP4.

### Prerequisites

- **Google Gemini API key** — same key used by Nanobanana. Get one at [AI Studio](https://aistudio.google.com).
- **google-genai Python SDK** — install via `pip install google-genai` or run via `uvx --with google-genai`.

### Generating video

```python
import time, os
from google import genai

client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])

operation = client.models.generate_videos(
    model="veo-3.1-generate-preview",
    prompt="Your video description here. Be specific about subject, action, setting, lighting, camera angle.",
)

while not operation.done:
    time.sleep(15)
    operation = client.operations.get(operation)

video = operation.response.generated_videos[0]
client.files.download(file=video.video)
video.video.save("output.mp4")
```

### Running via uvx (no virtual environment needed)

If Python/pip is not directly available, use uvx (already installed for Nanobanana):

```bash
GEMINI_API_KEY="your-key" uvx --with google-genai python generate_video.py
```

The API key is in `~/.claude.json` under the Nanobanana MCP config's `env.GEMINI_API_KEY`.

### Key details

- **Model**: `veo-3.1-generate-preview` (best quality), also `veo-3.1-fast` and `veo-3.1-lite` (cheaper)
- **Output**: 8-second 720p MP4 with native audio, typically 5–10 MB
- **Generation time**: ~60 seconds including polling
- **Cost**: Pay-per-second via Gemini API billing. Veo 3.1 Lite is the cheapest option for drafts.

### Prompt tips

Be specific and cinematic:
- Describe subject, action, setting, lighting, camera movement
- "Cinematic, photorealistic" improves quality
- Specify direction of motion: "walks from left to right"
- **Keep prompts to 3–5 sentences.** Long rule lists with numbered "must" requirements actually degrade results — Veo trains on natural cinematic descriptions, not specifications. Anything over ~150 words confuses the model.

### Known limits — when not to use Veo

Diffusion video models reliably fail at certain physics scenarios. Don't waste reprompts on these — go straight to a real-footage clip:

- **Sequential rigid-body chain reactions** (dominoes falling in sequence, pool break, Newton's cradle) — Veo produces plausible-looking but physically wrong cascades. Direction and contact transfer are unreliable.
- **Counting** (exactly N objects, ordered sequences) — frequently miscounts.
- **Text in any language** — see image rule above; same applies to video.
- **Complex multi-step "this then that then that"** beyond ~3 actions.

For these cases: search for stock footage (Pexels free tier, or `yt-dlp` on a Creative Commons clip) and trim with `ffmpeg`. Faster and better than reprompting.

### Image-to-video — when direction matters

When the result depends on a specific starting state (orientation, subject position, scene composition), generate a still image first via the image API, then feed it as the **starting frame** of the video. Anchoring the first frame dramatically reduces variance in direction, framing, and consistency.

### Workflow

1. Write a Python script with the prompt (or inline it)
2. Run via uvx → polls until video is ready (~60s)
3. Embed the MP4 in PowerPoint via `evaluate` → `AddMediaObject2`
4. Done — from text prompt to video playing in a slide

---

## Embedding Media in Slides

### Embedding video (Remotion output)

```python
# In evaluate tool:
s = presentation.Slides(slide_number)

# Remove old video if replacing
for i in range(s.Shapes.Count, 0, -1):
    if s.Shapes(i).Name == "VideoName":
        s.Shapes(i).Delete()

# Add new video — full slide coverage
video = s.Shapes.AddMediaObject2(
    r"C:\path\to\animation.mp4",
    False,  # LinkToFile — False embeds the video in the .pptx
    True,   # SaveWithDocument
    0, 0,   # Left, Top
    960, 540  # Width, Height (full slide at standard size)
)
video.Name = "VideoName"
video.AnimationSettings.PlaySettings.PlayOnEntry = True
video.AnimationSettings.PlaySettings.HideWhileNotPlaying = False
```

### Embedding images (Nanobanana output)

```python
# In evaluate tool:
pic = slide.Shapes.AddPicture(
    r"C:\path\to\image.png",
    False, True,
    left, top, width, height
)
pic.Name = "ImageName"
```

### Hybrid slides (image background + PowerPoint overlays)

Use Nanobanana for a dramatic background image, then add PowerPoint shapes on top for text/labels that need to be editable:

```python
# Background image
bg = slide.Shapes.AddPicture(r"path\to\bg.png", False, True, 0, 0, 960, 540)
bg.ZOrder(1)  # Send to back

# Overlay shapes on top
title = slide.Shapes.AddTextbox(1, 20, 20, 920, 60)
title.TextFrame.TextRange.Text = "Title Over Image"
title.TextFrame.TextRange.Font.Color.RGB = 16777215  # White
```

---

## Combined Workflow Patterns

### Pattern 1: "High-impact animated visualization"

Best for: systemic risks, network effects, process cascades, escalating trends

1. **PowerPoint MCP** → open deck, snapshot existing slides for context
2. **Remotion** → build animated diagram (glowing nodes, flowing arrows, growing elements)
3. **PowerPoint MCP** → embed MP4, add speaker notes, save

### Pattern 2: "Dramatic reveal"

Best for: before/after, impact stories, emotional content

1. **Nanobanana** → generate atmospheric background image
2. **PowerPoint MCP** → embed image as background, add text overlays and animations
3. **PowerPoint MCP** → `add_animation` for progressive text reveals on top

### Pattern 3: "Data + narrative"

Best for: presentations mixing charts with storytelling

1. **PowerPoint MCP** → create slides with native charts/shapes for data
2. **Nanobanana** → generate editorial/emotional images for transition slides
3. **Remotion** → animate the key "aha moment" visualization
4. **PowerPoint MCP** → assemble everything, add speaker notes throughout

### Pattern 4: "Prompt to video in a slide"

Best for: illustrative scenes, metaphors, product demos, any "show don't tell" moment

1. **Veo 3.1** → generate video from a text description (~60 seconds)
2. **PowerPoint MCP** → add slide, embed the MP4, save

This is the simplest and most powerful pattern. One prompt, one API call, one slide.

### Decision guide

```
Need animated diagram/visualization?
  ├─ Simple (3-5 shapes, basic entrances) → PowerPoint native animations
  └─ Complex (glowing, particles, growth, curves) → Remotion

Need a photorealistic video clip?
  ├─ Generic scene Veo can render → Veo 3.1 (~60 seconds)
  └─ Sequential physics (dominoes, collisions, chain reactions) → stock footage,
     not Veo (see "Known limits" above)

Need an image?
  ├─ Photo/illustration → image generator (Nanobanana / Gemini Imagen)
  ├─ Precise diagram with labels → PowerPoint native shapes
  └─ Background atmosphere → image generator

Need text overlays on media?
  └─ Always use PowerPoint shapes on top — never bake text into images/video
      (keeps it editable, avoids rendering artifacts)
```

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

## Common defects to self-check

Before declaring a slide done, scan against these common defects:

| Code | What to check |
|---|---|
| `body_below_floor` | Any sentence text below 18pt? Bump to 18pt or downgrade to a label. |
| `slop_default_font` | Using only Aptos / Calibri / Inter? Pair with a distinctive display face. |
| `monotone_anchor` | Three consecutive slides share the same anchor type? Insert variety. |
| `density_imbalance` | One slide packed, neighbours sparse? Rebalance. |
| `centered_long_body` | Long body copy centered? Left-align body; reserve center for hero/quote. |
| `inconsistent_spacing` | Mixing 8/12/16pt gaps? Pick one unit and stick to it. |
| `redundant_word_repeat` | Same key noun appearing 3+ times in dominant typography? Demote duplicates. |
| `attribution_truncated` | Author/source visibly cut off? Widen container or shorten attribution. |
| `chart_contrast_fail` | Chart bar labels below WCAG 3:1 contrast? Use chart-aware palette. |

Consider scripting these checks as a one-pass audit that walks every slide and prints a flat report — much faster than spot-checking by eye.

---

## Anti-patterns (recurring COM / build traps)

A consolidated catalog of the silent failures that have actually shipped broken slides. Skim this list whenever you're about to write a non-trivial COM patch — most of these fail without raising an error, so they don't show up in stack traces.

| Anti-pattern | What goes wrong | Where to find the rule |
|---|---|---|
| Trusting `app.ActivePresentation` | Targets the wrong deck if another window has focus | "COM patterns & safety → Multi-presentation safety" |
| `if not sh.HasTextFrame:` filter | Silently skips AutoShapes (they have empty text frames) | "COM patterns & safety → Filtering shapes" |
| Moving text without moving its backing | Text floats outside its card or off-slide | "COM patterns & safety → Move text and its backing" |
| Patching a 5+ shape slide instead of rebuilding | Z-order drifts, conditional filters miss shapes, layout never quite matches | "COM patterns & safety → Idempotent build scripts" |
| Skipping the LOOK step (step 3 of the loop) | Code "succeeds" but the rendered slide is wrong | "Workflow → Iteration loop" |
| Re-prompting an AI image after 2 text-baked attempts | Wastes budget; the model is locked into baking text | "Nanobanana → no-text rule" |
| Re-prompting Veo for a physics chain reaction | Diffusion video reliably fails dominoes / cradle / pool break | "Veo → Known limits" |
| Inline reimplementation of helper logic | Drifts from canonical version, no `--quiet` flag, no error stream propagation | "Wrapping Python helpers via subprocess + uvx" |
| Author-name marker for notes-append idempotency | Marker mismatches actual citation format ("X & Y" vs "X T., Y Q."), block gets appended every re-run | "Speaker notes → Idempotent notes appending" |

When one of these bites, fix it and **add a row here** if it's a new variant. The signal is: "I lost an hour to a silent failure" → it belongs in this table.
