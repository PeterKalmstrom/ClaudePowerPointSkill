# Benchmark prompts

The exact prompts used from round 9 on, so every round runs the same way. Replace the placeholders before use:

- `<BENCH>`: the round's working folder (for example `_scratch/bench10`), holding a copy of [../briefs.md](../briefs.md).
- `<PYTHON>`: the Python to use (the skill's `.venv`).
- `<SOFFICE>`: the full path to LibreOffice's `soffice` executable. Give the full path: an agent's shell may not
  see a PATH change made after the session started (in round 9 two makers could not find LibreOffice that way).

Files: [maker-plain.txt](maker-plain.txt), [maker-pptxskill.txt](maker-pptxskill.txt),
[maker-thisskill.txt](maker-thisskill.txt) (each maker runs as one unattended agent, all three in parallel),
[judge.txt](judge.txt) (one blind judge after all decks are rendered in PowerPoint into `<BENCH>/judge`).
