# ClaudePowerPointSkill

A Claude Code [skill](https://docs.claude.com/en/docs/claude-code/skills) that teaches Claude how to build production-quality PowerPoint presentations on Windows — slide layout discipline, COM-automation pitfalls, AI image / video generation, and the iteration loop that catches silent rendering failures.

This is **not** an MCP server. It's a body of operational knowledge — patterns, anti-patterns, and a tested workflow — packaged as a single `SKILL.md` that Claude loads when you ask it to work on a deck.

> **Screenshot wanted.** A before/after PNG of a slide built with the skill would help readers decide whether to install. Open an issue if you build something worth showing off.

---

## Why this exists

Vanilla Claude Code can already drive PowerPoint via the `powerpoint-mcp` server — open a deck, snapshot a slide, populate a placeholder, save. That part is one MCP install away.

What it *can't* do out of the box is avoid the silent failures that turn an afternoon of slide work into a debugging session. The recurring pain points:

- **PowerPoint COM lies about success.** A shape edit returns `{"success": true}` while the rendered slide is wrong. Text fits on paper but wraps in the export. The image cache holds the old version after replacement. `app.ActivePresentation` targets the wrong deck because another window grabbed focus.
- **AutoShape rectangles report `HasTextFrame == True`.** Filter logic that means "skip rectangles" silently skips half the slide.
- **Patches drift after the third one.** Backing rectangles get left behind while text moves; z-order shifts; the build script becomes shape soup.
- **AI image generators bake unwanted text into images** — often in the wrong language, often misspelled, always locked into the raster. Two reprompts cost more than starting over with a different background.
- **Diffusion video models reliably fail at rigid-body physics** — dominoes, Newton's cradle, pool break. Veo will produce a beautiful, physically wrong cascade and waste your generation budget if you let it.
- **Optimistic font sizing.** "36pt should fit" is a guess until the PNG confirms it. The COM property values don't tell you what the audience will see.

This skill encodes those lessons. Every rule has a "why" and a "how to apply" line so Claude can judge edge cases rather than pattern-matching blindly, and a consolidated *Anti-patterns* table at the bottom of `SKILL.md` for fast lookup when something feels off.

---

## Example: what working with the skill looks like

A representative session:

> **You:** Add a hero-stat slide showing "3.2°C — projected warming by 2100" to the deck I have open.
>
> **Claude:** *(loads `configuring-powerpoint-mcp` skill)*
> Before I touch the file, I'm going to make a timestamped backup copy — irreversible bulk edits live outside PowerPoint's undo stack.
> *(creates `deck.20260511-1430.pre-hero-stat.pptx`)*
>
> Now matching to "hero stat" anchor type — single big number, ~5-word visible budget excluding title. Typography hierarchy: number at 120pt bold, unit at 28pt bold, caption at 18pt italic. I'll use the "Title Only" layout so the number doesn't get absorbed into a content placeholder.
> *(runs `evaluate` against the open presentation, finds it by name match rather than ActivePresentation, builds the slide, names every shape, exports a PNG)*
>
> Here's the rendered slide — looking at the export, the unit at 28pt reads cleanly against the 120pt number, no text wrap, vertical balance is roughly equal top/bottom. Speaker notes have the IPCC source citation with the DOI. Ready to save?
>
> **You:** ship it.

The skill's contribution is everything Claude did *that you didn't ask for*: the backup, the layout choice, naming the shapes, exporting and inspecting the PNG, putting the citation in notes rather than on the slide. Each one is a rule from `SKILL.md` that prevented a specific past failure.

---

## What this skill gives Claude

Skills are how Claude Code carries domain expertise across sessions. When this skill is loaded, Claude can:

- **Drive PowerPoint via COM** — open / edit / save `.pptx` files programmatically through the `powerpoint-mcp` server.
- **Build slides idempotently** — using rebuild-from-scratch Python scripts rather than fragile patches.
- **Respect typography and word budgets** — anchor-type-aware layouts (hero stat, comparison pair, gallery, knowledge graph, etc.), 18pt body floor, label exception, balanced grids.
- **Generate AI imagery for slides** — via a Gemini-image MCP (e.g. "Nanobanana"), with a strict "no baked-in text" rule.
- **Generate animated video** — via [Remotion](https://www.remotion.dev/) (programmatic React → MP4) for diagrams, cascades, and feedback loops PowerPoint's native animations can't match.
- **Generate AI video from prompts** — via Google [Veo 3.1](https://ai.google.dev/gemini-api/docs/video-generation), with known-limits documented (no rigid-body chain reactions, no counting, no embedded text).
- **Audit and self-critique** — a catalog of taste defects (`body_below_floor`, `monotone_anchor`, `density_imbalance`, etc.) to scan against before declaring a slide done.
- **Avoid COM traps** — never trust `app.ActivePresentation`, `HasTextFrame` is not "is this a text shape", text and its backing shape move independently, etc.

The skill is opinionated. It encodes lessons from shipping real decks: optimistic font sizing, patch-on-patch shape soup, AI image generators that won't stop baking text into images, diffusion video models that lie about physics. Every rule has a "why" and a "how to apply" so Claude can judge edge cases instead of pattern-matching blindly.

---

## Requirements

| Component | Notes |
|---|---|
| **Windows** | The skill is Windows-only. PowerPoint MCP uses COM automation, which has no macOS / Linux equivalent. |
| **Microsoft PowerPoint** | Desktop install, not web-only. The COM bridge needs a real `POWERPNT.EXE` running. |
| **Claude Code** | CLI or VS Code extension. See [docs](https://docs.claude.com/en/docs/claude-code/overview). |
| **uv / uvx** | Python package runner. The PowerPoint MCP launches via `uvx`. Install with `irm https://astral.sh/uv/install.ps1 \| iex`. |

For the rich-media workflow:

| Component | Used by | Notes |
|---|---|---|
| **Node.js v18+** | Remotion | Required to render animated video. |
| **Google Chrome** | Remotion | Used as the rendering engine. Avoid Remotion's default Chrome Headless Shell download — it fails on cloud-synced folders. |
| **Gemini API key** | Nanobanana, Veo | Get one at [aistudio.google.com](https://aistudio.google.com). Same key works for both image and video generation. |
| **google-genai Python SDK** | Veo | Install via `pip install google-genai` or invoke with `uvx --with google-genai`. |

---

## MCP server dependencies

The skill assumes two MCP servers are configured in Claude Code:

### 1. `powerpoint-mcp` *(required)*

The core dependency. Provides the tools Claude uses to manipulate `.pptx` files: `manage_presentation`, `slide_snapshot`, `populate_placeholder`, `add_speaker_notes`, `add_animation`, `evaluate` (arbitrary Python in the COM context), and more.

- **Package**: [powerpoint-mcp on PyPI](https://pypi.org/project/powerpoint-mcp/)
- **Install** (registers with Claude Code at user scope):
  ```bash
  claude mcp add --scope user powerpoint -- "C:\Users\<USER>\.local\bin\uvx.exe" powerpoint-mcp
  ```
- **Verify**: restart Claude Code, then `claude mcp list` should show `powerpoint: ... - ✓ Connected`.
- **Troubleshooting**: see `SKILL.md` → *Troubleshooting* for the full diagnostic tree (project-scope traps, missing PATH, COM modal-dialog hangs, etc.).

### 2. Nanobanana — Gemini image-generation MCP *(optional, for AI imagery)*

"Nanobanana" is the working name for Gemini's flash-image model exposed as an MCP server. There are several community implementations; pick one and register it the same way:

```bash
claude mcp add --scope user nanobanana -- <command for your chosen server>
```

If you don't want an MCP for it, the same Gemini image API can be called directly via the `google-genai` Python SDK — see `SKILL.md` → *Nanobanana Integration* for the subprocess + uvx pattern that calls a Python helper instead of an MCP tool.

### Not MCPs (called directly)

- **Remotion** — CLI tool (`npx remotion render ...`), no MCP wrapper needed.
- **Veo 3.1** — accessed through the `google-genai` Python SDK; the skill shows the polling-loop pattern. There's no MCP layer.

---

## Installing the skill

Claude Code skills live in a `skills/` directory under your Claude config. To install this one:

1. Clone or download this repo.
2. Copy `SKILL.md` into your skills folder. The standard location is:
   ```
   <your-skills-path>/configuring-powerpoint-mcp/SKILL.md
   ```
   where `<your-skills-path>` is wherever you keep your Claude skills (e.g., `~/.claude/skills/`, a shared admin folder, or a per-project `.claude/skills/`).
3. Make sure the `powerpoint-mcp` server is registered (see above).
4. Restart Claude Code so the skill is picked up.

The skill's frontmatter tells Claude when to load it:

```yaml
---
name: configuring-powerpoint-mcp
description: Build production-quality PowerPoint decks on Windows via the powerpoint-mcp server ...
---
```

You don't invoke it explicitly — Claude reads the description and loads the skill when your task matches ("create a slide", "open this deck", "embed a video", "diagnose why PowerPoint tools are missing", etc.).

---

## What's in `SKILL.md`

Top-level sections (each is substantial — the file is ~1100 lines):

| Section | What it covers |
|---|---|
| **Prerequisites** | Windows / PowerPoint / uv versions and verification commands. |
| **Installation** | Three ways to register the MCP (CLI, direct JSON edit, pre-download), plus restart and `claude mcp list` verification. |
| **Troubleshooting** | Tools-not-appearing diagnostic tree, COM error recovery, image-placeholder swallow fix, snapshot-before-destructive-ops rule. |
| **COM patterns & safety** | Snapshot-before-bulk-edit, multi-presentation safety (never trust `ActivePresentation`), idempotent build scripts, the `HasTextFrame` trap, moving text and its backing together. |
| **Workflow** | The five-step iteration loop (build → render → **LOOK** → critique → fix), text-wrap defenses, font-size ceilings per container width. |
| **Available Tools** | Reference table of every MCP tool the PowerPoint server exposes, plus bulk-read and audit-deck strategies. |
| **Nanobanana Integration** | When to use AI-generated images vs. native shapes, the on-slide word budget (~10 words), anchor types with their word budgets (hero stat, comparison pair, gallery, knowledge graph, etc.), label vs. body floor, hero-stat pattern, two-column comparison pattern, balanced-layout sizing math, the "no text in images" rule with overlay pattern. |
| **Remotion Integration** | Project setup, required files, key APIs, render command with `--browser-executable`, animation design patterns (hub-and-spoke, progressive reveal), iteration workflow. |
| **Veo Integration** | Generating video via google-genai, polling pattern, prompt tips, known failure modes (chain reactions, counting, text), image-to-video for direction-sensitive scenes. |
| **Embedding Media in Slides** | COM snippets for `AddMediaObject2` (video) and `AddPicture` (image), plus hybrid image-background + PowerPoint-overlay pattern. |
| **Combined Workflow Patterns** | Four named patterns: animated visualization, dramatic reveal, data + narrative, prompt-to-video-in-a-slide. Plus a decision guide flowchart. |
| **Speaker notes** | What goes in notes (citations, anticipated Q&A, methodology caveats, pacing notes), the slide-vs-notes contract, idempotent notes-append using DOIs as markers. |
| **Showcase-first for multi-slide sections** | The rule that saves the most time: build slide 1 and one detail slide, get sign-off, *then* batch the rest. |
| **Common defects to self-check** | Nine taste-defect codes Claude scans against before declaring a slide done. |
| **Anti-patterns** | Consolidated catalog of recurring COM and build traps that fail silently. Each anti-pattern points back to the rule that prevents it. |

---

## How to work with the skill

The skill assumes a particular workflow: **iterate fast, render every change, look at the rendered PNG, critique against the rules, fix or save**. The single most-violated step is "look at the rendered PNG" — Claude's tool-output (`{"success": true}`) is not the same as a correctly rendered slide. If you find Claude declaring a slide done without showing you the export, push back. The skill warns about this explicitly but the temptation is constant.

Other day-to-day expectations the skill encodes:

- **Name every shape you might need to find later** — patches that find shapes by `Name == "HeroBacking"` survive layout changes; patches that find shapes by position or order break the moment another shape is added.
- **Rebuild over patch** once a slide has 5+ shapes. Three patches in a row = rewrite the build script.
- **Snapshot before risky bulk edits** with a file-copy (no COM). PowerPoint's undo stack does not survive `Save()` or session close.
- **Speaker notes are unbounded** — every citation, every methodology caveat, every "if asked X say Y" goes there. The slide carries the punch; the notes carry the depth.

---

## Project layout

```
ClaudePowerPointSkill/
├── README.md          # This file
├── SKILL.md           # The skill itself — loaded by Claude Code
├── CONTRIBUTING.md    # How to add new rules / anti-patterns
├── LICENSE            # MIT
└── .gitignore
```

No build step, no dependencies in this repo. The skill is pure markdown.

---

## License

[MIT](LICENSE). Use, fork, adapt, and integrate the patterns freely; attribution appreciated but not required.

---

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for the full guide. Short version: the skill grows by accretion — every silent failure that costs an hour of debugging belongs in the *Anti-patterns* table at the bottom of `SKILL.md`, with a linked rule in the relevant section explaining the fix.
