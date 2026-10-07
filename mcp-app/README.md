# PowerPoint Live (MCP App)

An MCP server, written for this skill, that drives desktop PowerPoint on Windows through COM — and, in hosts that
support [MCP Apps](https://modelcontextprotocol.io/docs/extensions/apps) (Claude Desktop, claude.ai), shows the
**current slide live** next to the conversation while Claude works on it.

## Install

As part of the plugin it is registered automatically. By hand:

```powershell
claude mcp add --scope user powerpoint-live -- uvx --with "mcp<2" --with "pywin32; sys_platform == 'win32'" python <skill folder>\mcp-app\server.py
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
- **Slide sorter** — thumbnails of every slide grouped by PowerPoint section (hidden slides dimmed, the current
  one outlined). Drag a thumbnail onto another to move it in front of it, or onto a section's empty space to move
  it to the end of that section (empty sections too). Click a thumbnail to open it in the Slide view.

The view polls `slide_state` about once a second and fetches a new picture only when the signature changes —
so edits by Claude, by a script or by a person in PowerPoint all show up. Images go to the view, not to Claude,
so the live preview costs no tokens.

The preview uses `Slide.Export`, which shows fallback fonts for embedded fonts. To check a slide closely, render
it with `scripts/render_slides.py`.
