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
sys.path.insert(0, os.path.join(ROOT, "scripts"))
from kShared import kRun, kS, kToolException  # noqa: E402  (the scripts folder is on the path only from here)

SKILL_MAX_LINES = 500


class kDocsCheck:
    """Links, anchors and the SKILL.md budget over every markdown file in the repository."""

    def __init__(self):
        try:
            self.Problems = []
            self.Files = []
            self.Texts = {}
            self.Heads = {}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDocsCheck.__init__")

    @staticmethod
    def Slug(Heading):
        """GitHub's anchor for a heading."""
        if kS.ErrorMode:
            return ""
        try:
            Heading = Heading.strip().lower()
            Heading = re.sub(r"[^\w\- ]", "", Heading)
            return Heading.replace(" ", "-")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDocsCheck.Slug")
            return ""

    @staticmethod
    def StripCode(Text):
        """The text without its fenced code blocks."""
        if kS.ErrorMode:
            return ""
        try:
            return re.sub(r"^[ \t]*```.*?^[ \t]*```", "", Text, flags=re.M | re.S)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDocsCheck.StripCode")
            return ""

    def Load(self):
        """Read every markdown file and collect its heading anchors."""
        if kS.ErrorMode:
            return
        try:
            self.Files = [os.path.relpath(Found, ROOT) for Found in
                          glob.glob(os.path.join(ROOT, "**", "*.md"), recursive=True)
                          if "node_modules" not in Found and "_scratch" not in os.path.relpath(Found, ROOT).split(os.sep)]
            for Name in self.Files:
                with open(os.path.join(ROOT, Name), encoding="utf-8") as File:
                    self.Texts[Name] = File.read()
            for Name, Text in self.Texts.items():
                self.Heads[Name] = {self.Slug(Heading) for Heading in
                                    re.findall(r"^#+ (.+)$", self.StripCode(Text), re.M)}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDocsCheck.Load")
            return

    def CheckLinks(self):
        """Every relative link resolves, and every #anchor names a heading."""
        if kS.ErrorMode:
            return
        try:
            for Name, Text in self.Texts.items():
                for Target in re.findall(r"\]\(([^)\s]+)\)", self.StripCode(Text)):
                    if re.match(r"^[a-z]+:", Target):
                        continue
                    Path, _, Anchor = Target.partition("#")
                    Full = os.path.normpath(os.path.join(os.path.dirname(Name), Path)) if Path else Name
                    if not os.path.exists(os.path.join(ROOT, Full)):
                        self.Problems.append(f"{Name}: link to missing file {Target}")
                    elif Anchor and Full in self.Heads and Anchor not in self.Heads[Full]:
                        self.Problems.append(f"{Name}: anchor #{Anchor} not found in {Full}")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDocsCheck.CheckLinks")
            return

    def CheckSkill(self):
        """SKILL.md frontmatter and size budget."""
        if kS.ErrorMode:
            return
        try:
            Skill = self.Texts.get("SKILL.md", "")
            Match = re.match(r"^---\nname: (\S+)\ndescription: (.+?)\n---\n", Skill, re.S)
            if not Match:
                self.Problems.append("SKILL.md: frontmatter must start with name and description")
            elif len(Match.group(2)) > 1024:
                self.Problems.append("SKILL.md: description over 1024 characters")
            if Skill.count("\n") > SKILL_MAX_LINES:
                self.Problems.append(f"SKILL.md: {Skill.count(chr(10))} lines, budget {SKILL_MAX_LINES} "
                                     "- move detail to reference/")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDocsCheck.CheckSkill")
            return


class kDocsCheckApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Check = kDocsCheck()
            Check.Load()
            Check.CheckLinks()
            Check.CheckSkill()
            for Problem in Check.Problems:
                print(Problem)
            print(f"{len(Check.Files)} markdown files checked, {len(Check.Problems)} problem(s)")
            return 1 if Check.Problems else 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kDocsCheckApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kDocsCheckApp)
