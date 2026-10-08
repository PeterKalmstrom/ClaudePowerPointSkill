# Contributing

> **How this repository is maintained.** The skill is developed in the author's skills collection and published
> here automatically on every change. Issues and pull requests are welcome: accepted changes are applied at the
> source and arrive here with the next publish (a pull request merged only here would be overwritten).

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
2. **A linked rule** in the right file under `reference/` (see the table in `SKILL.md`), explaining the fix. Link
   it from the anti-patterns row as `reference/FILE.md#anchor`. Each rule should have:
   - The rule itself, stated as a constraint to design around
   - **Why** — what specifically breaks if you ignore it (ideally a one-sentence incident)
   - **How to apply** — the concrete check or code pattern that prevents recurrence
3. **(Optional)** A minimal repro snippet in the PR description.

## Style

- **Rules over options.** The skill is most useful when it tells Claude what to *do*, not when it enumerates every possible approach. If two patterns work, pick the one you'd want Claude to use by default.
- **State the why.** A rule without a reason becomes cargo-cult behavior; one with a reason can be applied to edge cases. The "why" line is where the value lives.
- **Cite the incident.** "We hit this in November 2025 when the auto-export silently used the cached image" is more useful than "the cache can be stale." Specifics earn trust.
- **Avoid abstraction creep.** If a rule only applies to one specific anchor type or one specific COM call, name it. Don't generalize until you've seen the same failure mode three different ways.
- **No nested headings deeper than three.** Skills are read top-to-bottom by an LLM; deep hierarchies hurt retrieval.
- **Keep `SKILL.md` short.** It is loaded every time the skill triggers; CI fails it above 500 lines. New detail
  goes in `reference/`. Only a rule that applies to *every* deck belongs in *Core rules*.
- **Date version-specific facts** ("as of 2026-10") — model ids, package versions, upstream bugs — and list them
  under *Version-specific facts* in `SKILL.md`.
- **Say what it runs on.** Mark COM-only material; Linux, macOS and CI users skip it.

## Code conventions

Every Python file follows the error pattern in `scripts/kShared.py`, the same as the author's C#, TypeScript and
PowerShell code. `tools/check_kpattern.py` enforces it (CI and the self-test run it).

- **Four parts per function.** (1) First statement after the docstring: `if kS.ErrorMode: return <safe default>`
  (not in `__init__`). (2) The whole rest of the body in one `try`. (3) `except kToolException: raise` where
  expected states pass through, then a last `except Exception as e: kS.GlobalErrorHandler(e, "kClass.Method")` -
  a literal location starting with the class and method name. (4) The handler ends in `return <safe default>`
  or `raise`.
- **Expected states are not errors.** A missing file, a spec mistake or no deck open is found by an explicit
  check that raises `ToolInputException` (exit 2) or `ToolReportableException` (exit 1 or a given code) - never
  by catching.
- **The first error halts.** The handler sets `kS.ErrorMode`; every later method returns its default and the run
  exits 1. Only a fresh run or a person's Resume (`kS.Reset()`) re-arms it - never a `catch`.
- **Never hide an error.** A catch that only logs or returns a default needs an inline
  `ERROR-SUPPRESSED-JUSTIFIED: <why>` comment. A function that cannot follow the pattern says why in a
  `DOCUMENTED EXCEPTION: <why>` comment on or above its `def`.
- **Naming and shape.** Classes start with `k` (`kDeckReader`); methods, properties, locals and parameters are
  PascalCase; private fields `_camelCase`. No module-level functions, no lambdas, no nested functions - use a
  named method. A script starts with `kRun.Main(kXxxApp)`.

**Error reports.** A reported error can be sent to the support flow. The address is `KPS_ERROR_WEBHOOK` (or
`scripts/kErrorWebhook.url`, kept out of the repository); without one nothing is sent. **Nothing is ever sent without a
person's yes:** the report is shown and the person is asked "Do you want to send this error message? (yes/no)";
no means nothing is sent. Without a terminal the report is saved and Claude asks the user, sending it with
`scripts/send_error_report.py <file> --yes` only on yes; the script exits **4** and prints
`ERROR-REPORT-PENDING: <file>` so Claude knows. Errors that escape every guarded method (uncaught, other threads,
asyncio) are captured by `kS.InstallUnhandledExceptionCapture()` (installed by `kRun.Main` and PowerPoint Live). PowerPoint Live asks with Yes/No buttons in its view.
`KPS_ERROR_REPORT=0` switches reporting off (the self-test sets it).

## Sanity check before opening a PR

- `python tools/check_docs.py` passes (links, anchors, `SKILL.md` size) and `python scripts/selftest.py` passes.
  CI runs both. If you changed a COM script, also run `python scripts/selftest.py --com` on Windows and say so in
  the PR.
- If you changed the skill's `description`, run the trigger evals (`evals/README.md`) and add a case for what
  prompted the change.

- Does the new rule reduce the chance of a silent failure, or just add an option? Only the former belongs.
- Could Claude infer the rule from the surrounding code? If yes, the rule is redundant.
- Does the new section have a clean anchor link that the anti-patterns table can target? If your heading has a `—` (em-dash) in it, GitHub's slug generator collapses it to a double-hyphen — check that the anti-patterns row's anchor matches.

## License

Contributions are accepted under the same [MIT License](LICENSE) as the rest of the project.
