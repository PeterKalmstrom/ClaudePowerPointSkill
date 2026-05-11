# Contributing

The skill grows by accretion: every silent failure that costs an hour of debugging belongs in the *Anti-patterns* table at the bottom of `SKILL.md`. The goal is that Claude never re-learns the same lesson twice.

## Reporting a defect

Open an issue with:

- **What you asked Claude to do** — the prompt, ideally verbatim.
- **What Claude did** — the COM call / build script / generated image, whatever the output was.
- **What the rendered slide actually looked like** — the PNG export, not the tool's `{"success": true}` claim.
- **What you expected** — one sentence is enough.

Silent-failure reports are most valuable. If you lost an hour because PowerPoint said the operation succeeded but the rendered slide was wrong, that's the bug we want to capture.

## Submitting a fix

If you hit a defect and figured out the fix, send a PR with:

1. **A new row in the anti-patterns table** at the bottom of `SKILL.md`. Format:
   ```
   | <One-line anti-pattern> | <What goes wrong> | [Section name](#anchor) |
   ```
2. **A linked rule** in the relevant section explaining the fix. Each rule should have:
   - The rule itself, stated as a constraint to design around
   - **Why** — what specifically breaks if you ignore it (ideally a one-sentence incident)
   - **How to apply** — the concrete check or code pattern that prevents recurrence
3. **(Optional)** A minimal repro snippet in the PR description.

## Style

- **Rules over options.** The skill is most useful when it tells Claude what to *do*, not when it enumerates every possible approach. If two patterns work, pick the one you'd want Claude to use by default.
- **State the why.** A rule without a reason becomes cargo-cult behavior; one with a reason can be applied to edge cases. The "why" line is where the value lives.
- **Cite the incident.** "We hit this in November 2025 when the auto-export silently used the cached image" is more useful than "the cache can be stale." Specifics earn trust.
- **Avoid abstraction creep.** If a rule only applies to one specific anchor type or one specific MCP tool, name it. Don't generalize until you've seen the same failure mode three different ways.
- **No nested headings deeper than three.** Skills are read top-to-bottom by an LLM; deep hierarchies hurt retrieval.

## Sanity check before opening a PR

- Does the new rule reduce the chance of a silent failure, or just add an option? Only the former belongs.
- Could Claude infer the rule from the surrounding code? If yes, the rule is redundant.
- Does the new section have a clean anchor link that the anti-patterns table can target? If your heading has a `—` (em-dash) in it, GitHub's slug generator collapses it to a double-hyphen — check that the anti-patterns row's anchor matches.

## License

Contributions are accepted under the same [MIT License](LICENSE) as the rest of the project.
