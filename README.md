# ClaudePowerPointSkill

A Claude Code [skill](https://docs.claude.com/en/docs/claude-code/skills) that teaches Claude how to build production-quality PowerPoint presentations on Windows — slide layout discipline, COM-automation pitfalls, AI image / video generation, and the iteration loop that catches silent rendering failures.

This is **not** an MCP server. It's a body of operational knowledge — patterns, anti-patterns, and a tested workflow — packaged as a short `SKILL.md` that Claude loads when you ask it to work on a deck, with detail in `reference/` that it opens only when the task needs it, and helper scripts in `scripts/`.

![Before and after](examples/before-after/after.png)

*A slide built following the skill. See [the before/after example](examples/before-after/README.md) for the same slide built without it.*

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
> **Claude:** *(loads `building-powerpoint-decks` skill)*
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
  claude mcp add --scope user powerpoint -- "C:\Users\<USER>\.local\bin\uvx.exe" --with "mcp<2" powerpoint-mcp
  ```
- **Verify**: restart Claude Code, then `claude mcp list` should show `powerpoint: ... - ✓ Connected`.
- **Troubleshooting**: see `reference/SETUP.md` → *Troubleshooting* for the full diagnostic tree (project-scope traps, missing PATH, COM modal-dialog hangs, etc.).

### 2. Nanobanana — Gemini image-generation MCP *(optional, for AI imagery)*

"Nanobanana" is the working name for Gemini's flash-image model exposed as an MCP server. There are several community implementations; pick one and register it the same way:

```bash
claude mcp add --scope user nanobanana -- <command for your chosen server>
```

If you don't want an MCP for it, the same Gemini image API can be called directly via the `google-genai` Python SDK — see `reference/MEDIA.md` → *Nanobanana Integration* for the subprocess + uvx pattern that calls a Python helper instead of an MCP tool.

### Not MCPs (called directly)

- **Remotion** — CLI tool (`npx remotion render ...`), no MCP wrapper needed.
- **Veo 3.1** — accessed through the `google-genai` Python SDK; the skill shows the polling-loop pattern. There's no MCP layer.

---

## Installing the skill

Claude Code skills live in a `skills/` directory under your Claude config. To install this one:

1. Clone or download this repo.
2. Copy `SKILL.md`, `reference/` and `scripts/` into a folder named after the skill:
   ```
   <your-skills-path>/building-powerpoint-decks/SKILL.md
   <your-skills-path>/building-powerpoint-decks/reference/
   <your-skills-path>/building-powerpoint-decks/scripts/
   ```
   where `<your-skills-path>` is wherever you keep your Claude skills (e.g., `~/.claude/skills/`, a shared admin folder, or a per-project `.claude/skills/`).
3. Make sure the `powerpoint-mcp` server is registered (see above) if you work on Windows.
4. Restart Claude Code so the skill is picked up.

**Upgrading from `configuring-powerpoint-mcp`:** the skill was renamed in October 2026. Delete the old
`configuring-powerpoint-mcp` folder, or both skills will compete for the same tasks.

You don't invoke the skill explicitly — Claude reads its description and loads it when your task matches ("make a deck", "audit this presentation", "embed a video", "PowerPoint tools are missing", etc.).

---

## How the skill is organised

`SKILL.md` is deliberately short (under 500 lines, enforced by CI): a *start here* table, the ten core rules, the
build → render → **LOOK** → critique → fix loop, version-specific facts and the anti-patterns table. Everything
else is in `reference/`, which Claude opens only when the task needs it — so the skill costs little context until
it's actually used.

| File | What it covers | Runs on |
|---|---|---|
| `reference/SETUP.md` | Installing `powerpoint-mcp`, troubleshooting (including the `mcp<2` startup fix), MCP tool list | Windows + PowerPoint |
| `reference/COM.md` | Snapshots, multi-deck safety, UTF-8, idempotent builds, keeping hand edits, shape filtering | Windows + PowerPoint |
| `reference/LAYOUT.md` | Slide size, cropping pictures, text wrap, embedded-font rendering, word budgets, anchor types, layout patterns | Mostly any OS |
| `reference/MEDIA.md` | AI images, Remotion, Veo, video compression, embedding, combined patterns | Generation any OS; embedding Windows |
| `reference/ANIMATION.md` | When motion earns its place, effect table, native timing traps | Any OS |
| `reference/CONTENT.md` | Audience, claim titles, cognitive load, speaker notes, showcase-first | Any OS |
| `reference/AUDIT.md` | Bulk read, ~40 defect codes with severities, the full audit procedure | Scripts Windows; catalogue any OS |
| `reference/PRESENTING.md` | Timing markers, Q&A sheet, notes PDF, rehearsal | Windows + PowerPoint |
| `reference/AUTOMATION.md` | Headless PowerPoint for checks, building .pptx without COM | Mixed |
| `reference/LABELS.md` | Sensitivity labels and what encryption breaks | Any OS |

---

## How to work with the skill

The skill assumes a particular workflow: **iterate fast, render every change, look at the rendered image, critique against the rules, fix or save**. The single most-violated step is "look at the rendered image" — Claude's tool-output (`{"success": true}`) is not the same as a correctly rendered slide. If you find Claude declaring a slide done without showing you the render, push back. The skill warns about this explicitly but the temptation is constant.

Other day-to-day expectations the skill encodes:

- **Name every shape you might need to find later** — patches that find shapes by `Name == "HeroBacking"` survive layout changes; patches that find shapes by position or order break the moment another shape is added.
- **Rebuild over patch** once a slide has 5+ shapes. Three patches in a row = rewrite the build script.
- **Snapshot before risky bulk edits** with a file-copy (no COM). PowerPoint's undo stack does not survive `Save()` or session close.
- **Speaker notes are unbounded** — every citation, every methodology caveat, every "if asked X say Y" goes there. The slide carries the punch; the notes carry the depth.

---

## Testing

| What | How | Where |
|---|---|---|
| Docs: links, anchors, `SKILL.md` size and frontmatter | `python tools/check_docs.py` | CI (GitHub Actions) |
| Scripts that need no PowerPoint | `python scripts/selftest.py` | CI |
| Every script, including PowerPoint/COM ones | `python scripts/selftest.py --com` (see `scripts/README.md`) | Your Windows machine with PowerPoint |
| Does the skill load for the right prompts? | `evals/trigger-evals.json` (see `evals/README.md`) | Claude, with the `skill-creator` skill |

Run the Windows self-test after changing any COM script, and the trigger evals after changing the description.

---

## Project layout

```
ClaudePowerPointSkill/
├── SKILL.md           # Core: start here, core rules, build loop, anti-patterns — loaded by Claude Code
├── reference/         # Detail files SKILL.md links to, opened on demand
├── scripts/           # Helper scripts the skill calls, plus selftest.py
├── examples/          # Before/after example with its build script
├── evals/             # Trigger evals for the skill description
├── tools/             # Repo checks run by CI
├── .github/workflows/ # CI
├── README.md          # This file
├── CONTRIBUTING.md    # How to add new rules / anti-patterns
├── NOTICE             # Third-party credits (Impeccable, Apache-2.0)
└── LICENSE            # MIT
```

---

## Changelog

- **2026-10** — Renamed `configuring-powerpoint-mcp` → `building-powerpoint-decks`. Split into a short
  `SKILL.md` plus `reference/`. Added helper scripts and a self-test, CI, trigger evals and a before/after
  example. New guidance: `mcp<2` startup fix, Save As JPEG rendering for embedded fonts, word-break check,
  keeping hand edits, slide size, cropping, deck audits with ~40 defect codes, presenter prep, animation, video
  compression, Veo production lessons, claim titles, cognitive load, notes structure, sensitivity labels, building
  without COM.
- **2026-05** — First public version (`configuring-powerpoint-mcp`).

---

## License

[MIT](LICENSE). Use, fork, adapt, and integrate the patterns freely; attribution appreciated but not required. Some audit codes and the audience checklist are adapted from [Impeccable](https://impeccable.style/) (Apache-2.0) — see [NOTICE](NOTICE).

---

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for the full guide. Short version: the skill grows by accretion — every silent failure that costs an hour of debugging belongs in the *Anti-patterns* table at the bottom of `SKILL.md`, with a linked rule in the relevant section explaining the fix.
