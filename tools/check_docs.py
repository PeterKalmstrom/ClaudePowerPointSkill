"""Repo checks for the skill's markdown (run by CI; any OS, no dependencies).

- every relative link points at a file that exists
- every #anchor matches a heading in its target file (GitHub slug rules)
- SKILL.md frontmatter has a name and a description, and stays under the size budget
Links inside fenced code blocks are ignored.
"""
import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL_MAX_LINES = 500
problems = []


def slug(h):
    h = h.strip().lower()
    h = re.sub(r"[^\w\- ]", "", h)
    return h.replace(" ", "-")


def strip_code(text):
    return re.sub(r"^[ \t]*```.*?^[ \t]*```", "", text, flags=re.M | re.S)


files = [os.path.relpath(f, ROOT) for f in glob.glob(os.path.join(ROOT, "**", "*.md"), recursive=True)
         if "node_modules" not in f]
texts = {f: open(os.path.join(ROOT, f), encoding="utf-8").read() for f in files}
heads = {f: {slug(h) for h in re.findall(r"^#+ (.+)$", strip_code(t), re.M)} for f, t in texts.items()}

for f, t in texts.items():
    for target in re.findall(r"\]\(([^)\s]+)\)", strip_code(t)):
        if re.match(r"^[a-z]+:", target):
            continue
        path, _, anchor = target.partition("#")
        full = os.path.normpath(os.path.join(os.path.dirname(f), path)) if path else f
        if not os.path.exists(os.path.join(ROOT, full)):
            problems.append(f"{f}: link to missing file {target}")
        elif anchor and full in heads and anchor not in heads[full]:
            problems.append(f"{f}: anchor #{anchor} not found in {full}")

skill = texts.get("SKILL.md", "")
m = re.match(r"^---\nname: (\S+)\ndescription: (.+?)\n---\n", skill, re.S)
if not m:
    problems.append("SKILL.md: frontmatter must start with name and description")
elif len(m.group(2)) > 1024:
    problems.append("SKILL.md: description over 1024 characters")
if skill.count("\n") > SKILL_MAX_LINES:
    problems.append(f"SKILL.md: {skill.count(chr(10))} lines, budget {SKILL_MAX_LINES} - move detail to reference/")

for p in problems:
    print(p)
print(f"{len(files)} markdown files checked, {len(problems)} problem(s)")
sys.exit(1 if problems else 0)
