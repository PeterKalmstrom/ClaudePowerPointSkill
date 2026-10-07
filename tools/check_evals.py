"""Validate evals/trigger-evals.json: a list of {"query": str, "should_trigger": bool},
with both positive and negative cases and no duplicate queries. Any OS, no dependencies."""
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "scripts"))
from kShared import kRun, kS, kToolException  # noqa: E402  (the scripts folder is on the path only from here)


class kEvalsCheckApp:
    """Command line: check the trigger evals file."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            with open(os.path.join(ROOT, "evals", "trigger-evals.json"), encoding="utf-8") as File:
                Items = json.load(File)
            Problems = []
            for Index, Item in enumerate(Items):
                if set(Item) != {"query", "should_trigger"} or not isinstance(Item["query"], str) \
                        or not isinstance(Item["should_trigger"], bool) or not Item["query"].strip():
                    Problems.append(f"item {Index}: needs exactly a non-empty 'query' string and a 'should_trigger' bool")
            Queries = [Item.get("query") for Item in Items]
            if len(set(Queries)) != len(Queries):
                Problems.append("duplicate queries")
            Positive = sum(1 for Item in Items if Item.get("should_trigger") is True)
            if Positive == 0 or Positive == len(Items):
                Problems.append("need both should_trigger true and false cases")
            for Problem in Problems:
                print(Problem)
            print(f"{len(Items)} evals ({Positive} should trigger, {len(Items) - Positive} should not), "
                  f"{len(Problems)} problem(s)")
            return 1 if Problems else 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kEvalsCheckApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kEvalsCheckApp)
