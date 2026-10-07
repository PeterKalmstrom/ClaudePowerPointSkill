"""Validate evals/trigger-evals.json: a list of {"query": str, "should_trigger": bool},
with both positive and negative cases and no duplicate queries. Any OS, no dependencies."""
import json
import os
import sys

path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "evals", "trigger-evals.json")
items = json.load(open(path, encoding="utf-8"))
problems = []
for i, it in enumerate(items):
    if set(it) != {"query", "should_trigger"} or not isinstance(it["query"], str) \
            or not isinstance(it["should_trigger"], bool) or not it["query"].strip():
        problems.append(f"item {i}: needs exactly a non-empty 'query' string and a 'should_trigger' bool")
queries = [it.get("query") for it in items]
if len(set(queries)) != len(queries):
    problems.append("duplicate queries")
pos = sum(1 for it in items if it.get("should_trigger") is True)
if pos == 0 or pos == len(items):
    problems.append("need both should_trigger true and false cases")
for p in problems:
    print(p)
print(f"{len(items)} evals ({pos} should trigger, {len(items) - pos} should not), {len(problems)} problem(s)")
sys.exit(1 if problems else 0)
