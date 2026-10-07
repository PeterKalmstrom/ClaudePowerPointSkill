# Trigger evals

`trigger-evals.json` lists prompts that **should** load this skill and prompts that **should not**. They protect
the skill's `name` and `description` — the only part Claude sees before deciding to load it. Re-run them whenever
the description changes.

Format: a JSON list of `{"query": "...", "should_trigger": true|false}` (the format used by Anthropic's
`skill-creator` skill for description optimisation). CI only checks the file is well-formed
(`tools/check_evals.py`); running the evals needs Claude:

- **With `skill-creator`:** ask Claude Code to "run the trigger evals in evals/trigger-evals.json against this
  skill", or to "optimise the description" using them.
- **By hand:** in a fresh session with the skill installed, paste each query and note whether the skill loads.
  Aim for every `true` to load it and no `false` to.

Add a case whenever the skill fails to load when it should, or loads when it shouldn't. Negative cases should
be *near misses* (other Office formats, chart work, HTML pages), not obviously unrelated tasks.
