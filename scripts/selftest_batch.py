"""Self-test checks for batch mode (build_deck.py a.json b.json --check --apply, scripts/_batch.py) and the
per-slide `slides:` summary of --check. Run by selftest.py (kSelfTestBatch(Test).Run()); uses its Check / RunScript /
WriteJson and temp folder.
"""
import argparse
import json
import os

from kShared import kS

SAY = {"say": "This slide is part of the batch self-test and says what the presenter would say aloud here."}


class kSelfTestBatch:
    """The checks; Test is the running kSelfTest."""

    def __init__(self, Test):
        try:
            self.Test = Test
            self.Dir = os.path.join(Test.Tmp, "batch")
            os.makedirs(self.Dir, exist_ok=True)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestBatch.__init__")

    @staticmethod
    def Spec(Title):
        """A two-slide deck."""
        if kS.ErrorMode:
            return {}
        try:
            return {"direction": "consulting-blue", "footer": "Batch test", "sources": ["the brief"], "slides": [
                {"id": "cover", "pattern": "title", "title": Title, "subtitle": "Batch self-test", "notes": SAY},
                {"id": "ask", "pattern": "statement", "title": "One call checks every deck", "notes": SAY}]}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestBatch.Spec")
            return {}

    @staticmethod
    def BatchJson(Out):
        """The BATCH-JSON payload, or {}."""
        if kS.ErrorMode:
            return {}
        try:
            for Line in Out.splitlines():
                if Line.startswith("BATCH-JSON "):
                    return json.loads(Line[len("BATCH-JSON "):])
            return {}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestBatch.BatchJson")
            return {}

    def CheckBatch(self):
        """Two good specs and one with a spec error in one --check call: the good decks land beside their specs with
        a slides: table, the bad one is reported, the exit is the worst (2), one BATCH-JSON line covers all three."""
        if kS.ErrorMode:
            return
        try:
            Paths = [os.path.join(self.Dir, f"{Name}.json") for Name in ("one", "two", "bad")]
            self.Test.WriteJson(self.Spec("First batch deck builds"), Paths[0])
            self.Test.WriteJson(self.Spec("Second batch deck builds"), Paths[1])
            self.Test.WriteJson({"slides": [{"pattern": "no_such_pattern", "title": "Broken"}]}, Paths[2])
            Code, Out = self.Test.RunScript("build_deck.py", *Paths, "--check", "--apply")
            Data = self.BatchJson(Out)
            Decks = Data.get("decks", [])
            self.Test.Check("batch: one BATCH-JSON line for all three specs, worst exit code (2)",
                            len(Decks) == 3 and Code == 2, f"code {Code}; {Out[-800:]}")
            Built = [os.path.isfile(os.path.join(self.Dir, f"{N}.pptx")) for N in ("one", "two")]
            self.Test.Check("batch: decks go beside their specs without --out", all(Built), str(Built))
            self.Test.Check("batch: --check prints a slides: table per deck",
                            Out.count("slides:") >= 2 and "One call checks every deck" in Out, Out[-800:])
            Codes = [D.get("code") for D in Decks]
            self.Test.Check("batch: summary lists each deck's own exit code and marks the broken spec FIX",
                            Codes == [0, 0, 2] and "FIX " in Out, f"{Codes} {Out[-600:]}")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestBatch.CheckBatch")
            return

    def CheckHelpers(self):
        """Exit-code ranking and the single-spec default --out."""
        if kS.ErrorMode:
            return
        try:
            from _batch import kBatchCheck
            from build_deck import kBuildDeckApp
            self.Test.Check("batch: worst exit code ranks 4 > 1 > 3 > 2 > 0",
                            kBatchCheck.Worst([0, 2, 3]) == 3 and kBatchCheck.Worst([3, 1]) == 1
                            and kBatchCheck.Worst([1, 4]) == 4 and kBatchCheck.Worst([]) == 0)
            Args = argparse.Namespace(spec=[os.path.join("x", "deck.json")], check=True, out=None)
            Spec = kBuildDeckApp.DefaultOut(Args)
            self.Test.Check("check: without --out the deck goes beside the spec",
                            Spec == os.path.join("x", "deck.json") and Args.out == os.path.join("x", "deck.pptx"),
                            str(Args.out))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestBatch.CheckHelpers")
            return

    def Run(self):
        if kS.ErrorMode:
            return
        try:
            self.CheckHelpers()
            self.CheckBatch()
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestBatch.Run")
            return
