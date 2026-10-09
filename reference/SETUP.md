# Setup and troubleshooting

*Runs on: Windows + PowerPoint.* Part of the `building-powerpoint-decks` skill — start at [`SKILL.md`](../SKILL.md). Paths like `scripts/…` are relative to the skill folder.

Windows mode drives desktop PowerPoint directly through COM automation from Python (`pywin32`). There is no
server to install or register: Claude runs the scripts in `scripts/` (or a short Python snippet) with `uvx`.

## Prerequisites

- **Windows** — COM automation does not exist on macOS or Linux (use the any-OS scripts there).
- **Microsoft PowerPoint** — the desktop version, not web-only.
- **uv** — provides the `uvx` command that runs the scripts with their dependencies.

### Install uv

Open a **PowerShell** terminal:

```powershell
irm https://astral.sh/uv/install.ps1 | iex
```

This installs `uv.exe` and `uvx.exe` to `C:\Users\<USER>\.local\bin\`. If Claude Code can't find `uvx`, use that
full path.

### LibreOffice without administrator rights (optional)

`scripts/render_lo.py` (approximate renders, any OS) needs LibreOffice and `pdftoppm`. Without admin rights, extract
the official MSI instead of installing it - no registry, no elevation:

```powershell
msiexec /a LibreOffice_<version>_Win_x86-64.msi /qn TARGETDIR=C:\Tools\LibreOffice
```

`render_lo.py` finds soffice from `KPS_SOFFICE` (the exe or its folder), then PATH, then
`C:\Tools\LibreOffice\program` and `C:\Program Files\LibreOffice\program`. Prefer a local disk over a synced
folder: the extract is about 19,500 files (~1.5 GB) that a sync client would upload and lock.

## Verify

```powershell
uvx --with python-pptx --with pillow --with numpy --with pywin32 python scripts/selftest.py --com
```

The `--com` checks open PowerPoint, render a slide and close it again. All passing means Windows mode works.

## Talking to PowerPoint from Python

```python
import win32com.client
app = win32com.client.Dispatch("PowerPoint.Application")           # starts PowerPoint, or attaches to it
prs = app.Presentations.Open(r"C:\path\to\deck.pptx", WithWindow=False)
print(prs.Slides.Count)
prs.Save()
prs.Close()
```

Run snippets as `uvx --with pywin32 python snippet.py`. To work on a deck the user already has open, use
`win32com.client.GetActiveObject("PowerPoint.Application")` and pick the presentation by name — see
[COM](COM.md#multi-presentation-safety--never-trust-activepresentation). To *see* a slide, render it with
`scripts/render_slides.py` (not `Slide.Export`).

## Live slide view (optional MCP App)

[`mcp-app/`](../mcp-app/README.md) is this skill's own MCP server: `powerpoint_open`, `powerpoint_run` (Python
against the deck) and `powerpoint_show`. In hosts that support MCP Apps it shows the current slide live while it
changes. The plugin registers it; by hand:

```powershell
claude mcp add --scope user powerpoint-live -- uvx --with "mcp<2" --with python-pptx --with pillow --with "pywin32; sys_platform == 'win32'" python <skill folder>\mcp-app\server.py
```

---

## Troubleshooting

### `uvx` not found

The `uv` installer wasn't run, or PATH wasn't updated. Use the full path `C:\Users\<USER>\.local\bin\uvx.exe`.

### PowerPoint COM errors

- Ensure PowerPoint is installed (not just Office web apps).
- Close any "PowerPoint has stopped working" dialogs.
- If PowerPoint is open with a modal dialog (e.g., save prompt), the COM connection may hang.

### Images get swallowed by content placeholders

When adding a picture with `AddPicture` to a slide that uses a layout with a content placeholder (e.g., "Title and Content"), PowerPoint may absorb the image into the placeholder instead of creating a standalone Picture shape. Deleting the placeholder doesn't help — the layout regenerates it.

**Fix:** Change the slide layout to **"Title Only"** before adding the image:

```python
title_only = presentation.SlideMaster.CustomLayouts(6)  # "Title Only"
slide.CustomLayout = title_only
# Now AddPicture creates a real Picture shape (type=13)
slide.Shapes.AddPicture(path, False, True, left, top, w, h)
```

This prevents the layout from regenerating content placeholders. The title placeholder is preserved. Use this whenever you need standalone images on slides that originally had content placeholders.
