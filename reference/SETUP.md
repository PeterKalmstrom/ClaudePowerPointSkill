# Setup and troubleshooting

*Runs on: Windows + PowerPoint.* Part of the `building-powerpoint-decks` skill — start at [`SKILL.md`](../SKILL.md). Paths like `scripts/…` are relative to the skill folder.

Installs `powerpoint-mcp` so Claude Code can open, read, edit, and create PowerPoint presentations via COM
automation on Windows.

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
claude mcp add --scope user powerpoint -- "C:\Users\<USER>\.local\bin\uvx.exe" --with "mcp<2" powerpoint-mcp
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
      "args": ["--with", "mcp<2", "powerpoint-mcp"],
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
      "args": ["--with", "mcp<2", "powerpoint-mcp"],
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

### `CONNECTION_CLOSED` at startup: `No module named 'mcp.server.fastmcp'`

*As of 2026-10, `powerpoint-mcp` 1.30.0.*

**Symptom:** Claude Code reports `powerpoint (CONNECTION_CLOSED): "Connection closed"` and no PowerPoint tools
appear. Run by hand, the server crashes on import with
`ModuleNotFoundError: No module named 'mcp.server.fastmcp'. This is mcp 2.x, where FastMCP was renamed to MCPServer`.

**Root cause:** `powerpoint-mcp` (1.30.0) does not pin its `mcp` dependency. `uvx` resolves the newest `mcp`, and
`mcp` 2.x renamed `FastMCP` → `MCPServer`, so the server dies before it answers `initialize`. Nothing in your config
changed; a new upstream release broke a setup that used to work.

**Fix:** pin `mcp<2` in the server args:

```json
"powerpoint": {
  "type": "stdio",
  "command": "C:\\Users\\<USER>\\.local\\bin\\uvx.exe",
  "args": ["--with", "mcp<2", "powerpoint-mcp"],
  "env": {}
}
```

or with the CLI: `claude mcp add --scope user powerpoint -- "C:\Users\<USER>\.local\bin\uvx.exe" --with "mcp<2" powerpoint-mcp`.

To confirm before restarting, send `initialize` then `tools/list` over stdin to
`uvx.exe --with "mcp<2" powerpoint-mcp`: a working server returns `serverInfo` "PowerPoint MCP Server" and 11 tools.
Remove the pin once `powerpoint-mcp` itself declares `mcp<2` or supports 2.x.

**Common mistake:** reading `CONNECTION_CLOSED` as a config or PATH problem. Run the exact configured command by
hand first — the traceback names the real cause in one line.

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

## Quick Verification

After setup, test with:

```
Open any .pptx file in the workspace using manage_presentation,
then run slide_snapshot on slide 1.
```

If you get slide content and a screenshot, the setup is complete.

---
