# COM patterns and safety

*Runs on: Windows + PowerPoint.* Part of the `building-powerpoint-decks` skill — start at [`SKILL.md`](../SKILL.md). Paths like `scripts/…` are relative to the skill folder.

## COM patterns & safety

Foundational rules for any COM script targeting PowerPoint. These are not "things to do when something breaks" — they are constraints to design around from the first line of code. Read this section before writing any non-trivial COM patch or build script.

### Snapshot before any risky bulk edit

**Rule:** Before running any script that touches more than a handful of shapes — bulk font bumps, layout reflows, multi-slide rebuilds — copy the deck to a timestamped backup. COM operations succeed without raising errors when they corrupt the wrong shapes; the backup is your only undo.

The cheapest pattern is a file-copy (no COM), so the user's open PowerPoint session is undisturbed. Use
[`scripts/backup_snapshot.py`](../scripts/backup_snapshot.py):

```bash
python scripts/backup_snapshot.py --file "C:/path/to/deck.pptx" --label pre-bulk-fontbump
```

Output lands beside the source as `<deck>.YYYYMMDD-HHMMSS.<label>.pptx`. It refuses to overwrite an existing backup,
so re-running it never destroys an earlier snapshot. Save in PowerPoint first if there are unwritten changes — the
file-copy reads disk, not in-memory state.

When to snapshot:
- Before any bulk script that walks all slides
- Before a deep-rebuild of a large slide (5+ shapes)
- Before final pre-talk edits (label `pre-rehearsal`)
- After a known-good iteration — gives you a "last clean version" rollback
- Before the first destructive operation in a new session on a user's working deck — when the user says "my big
  deck" or "the real one", assume it is irreplaceable

**Undo is not a backup.** `ExecuteMso("Undo")` covers most shape edits in the current session, but not media
inserts, layout changes or some COM edits; it is strictly linear; it is gone once PowerPoint closes; and after
`Save()` the file on disk has already changed. Treat the snapshot as a git commit and Undo as editor undo — use
both. Session undo alone is enough only for reversible tweaks: text edits on placeholders, moving or resizing
shapes, adding new content.

**Restoring:** close the original without saving and reopen the backup. (From inside PowerPoint you can also take a
side copy with `presentation.SaveCopyAs(path)`, which leaves the open document untouched — it captures unsaved
changes, which a file copy does not.)

### Multi-presentation safety — never trust `ActivePresentation`

**Rule:** When more than one presentation is open in PowerPoint, `app.ActivePresentation` and any code that depends on it can silently target the wrong file. Other windows (a separate deck, a template, a reviewer's copy) can grab focus and flip the active selection without warning.

**Why this happens:**
- Code using `ActivePresentation` resolves it at the moment each call is made.
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

If some code must use `ActivePresentation`, close all other PowerPoint windows first, or activate the right window with `target.Windows(1).Activate()` before the call.

**Multiple matching presentations open.** `next((p for p in app.Presentations if "MyDeck" in p.Name), None)` returns the *first* match. If both `MyDeck.pptx` and `MyDeck-conflict-...pptx` (or `MyDeck2.pptx`) are open and either could match the substring, prefer **exact name first, then fall back to substring**:

```python
target = next((p for p in app.Presentations if p.Name == "MyDeck.pptx"), None)
if target is None:
    target = next((p for p in app.Presentations if "MyDeck" in p.Name), None)
```

A `for ... if ...: target = p` loop **without `break`** is worse than `next(...)` — it walks the entire list and ends up on the *last* match, which is whichever order PowerPoint enumerates the open windows (effectively random).

### Force UTF-8 stdout in any script that prints Unicode

**Rule:** On Windows, Python's stdout defaults to the legacy cp1252 codec. Any `print()` containing `→`, `↓`, `≈`, non-ASCII diacritics, etc. crashes with `UnicodeEncodeError` — *after* the COM work has already succeeded and (often) saved the deck. The result is a confusing "the script printed something but also raised, did it actually run?" state, and a partial second run if you retry.

```python
# Put at the top of every COM script that has print() with non-ASCII content:
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
```

**Why this matters:** the COM call to `target.Save()` happens before the crashing `print()`, so the deck state is correct, but the operator can't tell from the stack trace. They'll re-run, and an append-style script (e.g. `notes_*.py`) may double-apply. Cheaper to add three lines upfront than to debug duplicated state.

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
    w, h = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight   # never hard-code 960 x 540
    slide.Shapes.AddPicture(img, 0, -1, 0, 0, w, h).Name = "BgImage"
```

Each run produces the same result. To tweak typography or layout, edit the script and re-run — no manual element-by-element fixing in PowerPoint.

**Strongly prefer rebuild over patch.** Once a slide has 5+ shapes and you need a layout change touching more than two elements, rebuild from scratch is almost always cheaper and safer than a patch script. Patches drift — backing rectangles get left in old positions while text moves, shape Z-order silently shifts, conditional filters miss shapes. A rebuild is a single source of truth.

**Sentinel for "already built?" must only exist *after* the build.** A common bug in idempotent expansion/build scripts: the "skip duplication" check searches for a phrase that *also exists in the pre-build content*. The script thinks the work is done, skips the structural step (e.g. `Duplicate()`), then proceeds to rewrite text into the wrong slides — usually destroying adjacent unrelated slides.

```python
# BAD — "We have the solutions" appears on the original opener slide already.
# First run does Duplicate() + rebuild. Second run sees the phrase on the
# pre-build slide, skips Duplicate(), but still rewrites — overwriting the
# NEXT slide with content meant for the duplicate.
if slide_has(target.Slides(start_n), "We have the solutions"):
    skip_duplication = True

# GOOD — sentinel is something the build itself adds and the pre-build slide
# cannot contain. The motto is set by this script; nothing else writes
# "We are lying." into a slide of this section.
if slide_motto(target.Slides(start_n + 1)) == "We are lying.":
    skip_duplication = True
```

**Rule:** the sentinel should be a *post-condition of the build*, not a phrase from the source content. If the same string could plausibly exist before the script runs, it's not a sentinel — it's a coincidence waiting to happen.

### A generated deck that people also edit in PowerPoint: harvest before you overwrite

**Symptom:** a rebuild from source silently erased a day of a client's PowerPoint edits (87 slides). The agent then
made it worse by treating the edits as defects against the slide rules and reverting some of them. **A person's edit
is content: keep it exactly, and report anything that looks accidental.**

**Use the tool** — [`scripts/harvest_edits.py`](../scripts/harvest_edits.py) (any OS) does all of this:

```bash
python scripts/harvest_edits.py manifest deck.pptx                    # after every build
python scripts/harvest_edits.py harvest deck.pptx --store edits/      # before rebuilding
python scripts/build_deck.py spec.json --out new.pptx                 # rebuild
python scripts/harvest_edits.py restore new.pptx --store edits/ --out final.pptx
```

It keeps edited slides XML-exactly (with their pictures and charts), re-inserts added slides after the slide they
followed, keeps deleted slides deleted, and lets everything else take the new build. Read its output: every
decision is listed.

**Working pattern (what the tool does):**
1. Before saving a rebuild, fingerprint the deck on disk against the last build's manifest, and copy every changed,
   added, deleted or moved slide into a frozen store.
2. Key each generated slide by `<p:cSld name="…">` (the slide's `Name` in COM) — set it to a permanent id in the build.
3. After generating, copy the stored slides back in, XML-exactly.
4. Fingerprint the saved file into the manifest for next time.

**Measured facts it depends on:**
- **PowerPoint keeps `<p:cSld name>`** through a text edit, delete, move and save. A **duplicated** slide copies it
  too, so treat the second occurrence as new.
- **What PowerPoint changes on save with no edits:** the `top`/`height` of "resize shape to fit text" boxes
  (`a:spAutoFit`), and **empty `a:r` runs are dropped**. Leave both out of any "was this slide edited?" fingerprint —
  with them excluded, a re-save with no edits changed 0 fingerprints and a real edit changed only its own.
- **Copying a slide into another python-pptx presentation** (python-pptx 1.0.2):
  - Rebuild the rels with the **same rIds** (`part.rels._rels[rId] = _Relationship(...)`).
  - Images go through `package.get_or_add_image_part`; any other part (SmartArt `diagram*`, media) is copied
    recursively with `package.next_partname`.
  - **Clear the cached `shapes` proxy** (`slide.__dict__.pop("shapes")`) after replacing the XML, or reads see the
    old, detached tree.
  - Call **`prs.part.rename_slide_parts(...)` after deleting slides**, or `add_slide` reuses `slide<count+1>.xml`
    and the zip gets duplicate entries.
- A PowerPoint save keeps an **embedded font** even when it is not installed on the machine.

### Normalise paths before `SaveCopyAs` / `SaveAs`

`Presentation.SaveCopyAs(path, …)` fails on a path that mixes `/` and `\` ("PowerPoint can't save ^0 to ^1").
Pass `os.path.normpath(os.path.abspath(path))`.

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
    if sh.HasTextFrame and "intro" in sh.TextFrame.TextRange.Text:
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
