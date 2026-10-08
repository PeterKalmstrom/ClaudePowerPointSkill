# Enhance PowerPoint with Claude Code

*A tutorial by Peter Kalmström*

Using Claude Code with PowerPoint offers benefits beyond aesthetics. Claude acts as a thinking partner — identifying inconsistencies, validating arguments, and assessing presentation readability and layout quality. Because Claude can read the whole deck at once, it understands the presentation as a whole, not just slide by slide.

## How it works

Claude Code operates using skills: reusable knowledge packages that enable task-specific performance. This tutorial uses a free PowerPoint skill available on GitHub. The skill connects Claude to the PowerPoint MCP (Model Context Protocol) server, which uses COM automation to control a running PowerPoint application on Windows.

Peter Kalmström's demonstration shows Claude enhancing an existing presentation and generating a video embedded in a new slide.

## Prerequisites

- A Microsoft PowerPoint presentation open on your local machine
- Claude Code (requires a Claude Pro or Max subscription)
- Visual Studio Code with the Claude Code extension
- **uv** — a Python package runner that provides the `uvx` command used to launch the MCP server

### Install uv

Open PowerShell and run:

```powershell
irm https://astral.sh/uv/install.ps1 | iex
```

This installs `uvx.exe` to `C:\Users\<USER>\.local\bin\`.

## Setup steps

1. Start a new Claude Code session in VS Code.
2. Copy the skill URL from GitHub and prompt Claude to install the skill.
3. When Claude runs the MCP add command, make sure it uses **`--scope user`** (not the default project scope) and the **full absolute path** to `uvx.exe` — for example:
   ```
   C:\Users\<USER>\.local\bin\uvx.exe
   ```
   Using a bare `uvx` command will fail silently; the full path is required because Claude Code's shell does not reliably include `~/.local/bin` in PATH.
4. Restart VS Code so the MCP server is discovered.
5. Run the following once to pre-download the server's packages (~45 packages, cached after the first run):
   ```powershell
   "C:\Users\<USER>\.local\bin\uvx.exe" powerpoint-mcp --help
   ```
6. Choose to install the MCP (not just use it for the session).
7. Restart VS Code again to activate the MCP.
8. Prompt Claude to create a slide — this confirms the tools are loaded. You can verify with `claude mcp list`; a `✓ Connected` status confirms the setup is complete.
9. Claude creates slides using the theme from the existing open presentation.

## What you can do after setup

Once connected, you can:

- Ask Claude to **analyze your presentation** and receive improvement suggestions covering content, structure, and layout.
- Have Claude **apply edits automatically** — rewriting text, adjusting layouts, adding slides.
- Generate **MP4 videos** using Veo 3.1 (requires a Gemini API key) and have Claude embed them directly into slides.

The Gemini API key is stored in your MCP configuration. A Gemini API key can be obtained from [Google AI Studio](https://aistudio.google.com).
