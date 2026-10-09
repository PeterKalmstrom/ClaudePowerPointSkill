# Helper scripts

Referenced from `SKILL.md`. Run them with `uvx` so their dependencies never touch your project.

| Script | Needs | What it does |
|---|---|---|
| `build_deck.py` | `python-pptx`, `pillow` (any OS) | Builds a deck from a JSON/YAML spec: 21 patterns that fill the slide, design directions or a template, native charts, computed figures; `--check` (build + lint + render + one summary of every problem), `--plan` (story + pre-checks), automatic fixes for unfit text (`auto:` lines; `--no-auto`), `--lint`, `--slides 1,3` |
| `directions.json` | — | The 20 design directions `build_deck.py` and DESIGN.md use |
| `read_deck.py` | `python-pptx` (any OS) | Whole deck to JSON (ids, layouts, shapes with positions and sizes, charts, tables, notes) or a text outline |
| `harvest_edits.py` | `python-pptx` (any OS) | Keeps people's hand edits (edited, added, deleted slides) when a generated deck is rebuilt |
| `fix_deck.py` | `python-pptx`, `pillow` (any OS) | Applies the safe fixes for lint findings to a copy (or `--in-place`, keeping a `.bak`), lists each, re-lints |
| `lint_deck.py` | `python-pptx`, `pillow` (any OS) | Lints a deck for the AUDIT.md defect codes from the file alone; `--json`, `--room-depth`, `--fix --out` |
| `extract_theme.py` | `python-pptx` (any OS) | Theme colours, fonts, layouts and placeholders as JSON, or a `brand-spec.md` skeleton |
| `render_lo.py` | LibreOffice + poppler (any OS) | Approximate slide PNGs without PowerPoint; `--sheet` adds a contact sheet |
| `diff_renders.py` | `pillow`, `numpy` (any OS) | Which slides changed between two render folders, with heat maps |
| `backup_snapshot.py` | Python only | Timestamped side copy of a deck before a risky edit |
| `check_word_breaks.py` | Windows + PowerPoint, `pywin32` | Fails if any word is broken across two lines |
| `render_slides.py` | Windows + PowerPoint, `pywin32` | Slides to JPEG via Save As JPEG (renders embedded fonts correctly) |
| `contact_sheet.py` | `pillow`; `--file` needs Windows + PowerPoint (`pywin32`) | All slides as one thumbnail grid PNG — `--renders <folder>` on any OS |
| `export_pdf.py` | Windows + PowerPoint, `pywin32` | Notes-page / handout / slides PDF |
| `check_slide_size.py` | `python-pptx` (any OS) | Fails if a deck is not the expected size (default Full HD) |
| `cover_crop.py` | `pillow` (any OS) | Reports stretch distortion and crops an image to a box ratio |
| `selftest.py` | `python-pptx`, `pillow` (+ `pywin32` with `--com`) | Builds a test deck with known defects and checks every script against it |

`_theme.py` resolves theme colours and slide backgrounds for the linter. `_measure.py` estimates text wrapping from font metrics. `_figures.py` checks derived figures (totals, averages, changes, shares) against the deck's own data and computes the builder's figure tokens (`{sum}`, `{change}` …). `_rules.py` holds the rule helpers shared by the build and lint scripts (label titles, contrast, overlap, room-depth floors). `_ppt.py` is the shared COM helper: it reuses a deck already open in PowerPoint (matched by
path, never `ActivePresentation`), otherwise opens it read-only and windowless, and quits
PowerPoint only if it started it.

Every script follows the error pattern in `kShared.py`: an unexpected error is reported once (`ERROR in
Class.Method`), halts the run and exits 1; expected states exit 2 (bad input) or 1 with a plain message. See
[Code conventions](../CONTRIBUTING.md#code-conventions); `tools/check_kpattern.py` enforces it.

Run from this folder (or pass the full path) so `_ppt.py` is importable:

```bash
uvx --with python-pptx --with pillow python scripts/lint_deck.py "C:/path/to/deck.pptx"
```

## Self-test

```bash
# any OS: the scripts that need no PowerPoint (CI runs this)
uvx --with python-pptx --with pillow python scripts/selftest.py
# Windows with PowerPoint: every script
uvx --with python-pptx --with pillow --with pywin32 python scripts/selftest.py --com
```

`--ci-fonts` runs the self-test as the GitHub Actions runner measures: only DejaVu, Liberation, Carlito and Caladea,
through a private `FONTCONFIG_FILE` and `PPTSKILL_FONT_DIRS` (the system's fonts are untouched). `PPTSKILL_FONT_DIRS`
(paths joined with `os.pathsep`) replaces the font folders `_measure.py` scans.

Close other PowerPoint windows first: the COM scripts reuse a running PowerPoint. Exit code 0 means every check
passed; output files are left in a temp folder whose path is printed.
