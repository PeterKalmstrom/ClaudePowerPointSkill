# Helper scripts

Referenced from `SKILL.md`. Run them with `uvx` so their dependencies never touch your project.

| Script | Needs | What it does |
|---|---|---|
| `backup_snapshot.py` | Python only | Timestamped side copy of a deck before a risky edit |
| `bulk_read.py` | Windows + PowerPoint, `pywin32` | All slide text + notes to JSON in one COM session |
| `audit_deck.py` | Windows + PowerPoint, `pywin32` | Words per slide, smallest body font, status per slide, topic-label titles |
| `check_word_breaks.py` | Windows + PowerPoint, `pywin32` | Fails if any word is broken across two lines |
| `render_slides.py` | Windows + PowerPoint, `pywin32` | Slides to JPEG via Save As JPEG (renders embedded fonts correctly) |
| `contact_sheet.py` | Windows + PowerPoint, `pywin32`, `pillow` | All slides as one thumbnail grid PNG |
| `export_pdf.py` | Windows + PowerPoint, `pywin32` | Notes-page / handout / slides PDF |
| `check_slide_size.py` | `python-pptx` (any OS) | Fails if a deck is not the expected size (default Full HD) |
| `cover_crop.py` | `pillow` (any OS) | Reports stretch distortion and crops an image to a box ratio |
| `selftest.py` | `python-pptx`, `pillow` (+ `pywin32` with `--com`) | Builds a test deck with known defects and checks every script against it |

`_ppt.py` is the shared COM helper: it reuses a deck already open in PowerPoint (matched by
path, never `ActivePresentation`), otherwise opens it read-only and windowless, and quits
PowerPoint only if it started it.

Run from this folder (or pass the full path) so `_ppt.py` is importable:

```bash
uvx --with pywin32 python scripts/audit_deck.py --file "C:/path/to/deck.pptx"
```

## Self-test

```bash
# any OS: the scripts that need no PowerPoint (CI runs this)
uvx --with python-pptx --with pillow python scripts/selftest.py
# Windows with PowerPoint: every script
uvx --with python-pptx --with pillow --with pywin32 python scripts/selftest.py --com
```

Close other PowerPoint windows first: the COM scripts reuse a running PowerPoint. Exit code 0 means every check
passed; output files are left in a temp folder whose path is printed.
