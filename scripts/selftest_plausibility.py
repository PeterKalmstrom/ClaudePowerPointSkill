"""Self-test checks for round 10's content findings (_plausibility.py): an invented per-person goal that contradicts
the deck's total target, a goal nobody listed as assumed, and objection answers cut to fragments. Run by selftest.py
(kSelfTestPlausibility(Test).Run()); uses its Check / RunScript / WriteJson and temp folder.
"""
import copy
import os

from kShared import kS

LAUNCH = {"direction": "electric-launch", "footer": "Insight Pro launch, 1 March 2027", "sources": ["the brief"],
          "facts": {"sales_team": 25},
          "slides": [
              {"id": "cover", "pattern": "title", "title": "Insight Pro launches 1 March 2027",
               "notes": {"say": "Insight Pro launches on 1 March, and this is how we sell it."}},
              {"id": "target", "pattern": "kpi", "title": "Our 2027 target: 120 new customers",
               "metrics": [{"value": "120", "label": "New customers in 2027"},
                           {"value": "1.8 MUSD", "label": "New ARR in 2027"}],
               "notes": {"say": "The target for 2027 is 120 new customers and 1.8 million dollars of new ARR."}},
              {"id": "objections", "pattern": "compare", "title": "Three objections, three answers",
               "columns": [{"heading": "No budget", "points": ["15 % off until 30 April"]},
                           {"heading": "Too new", "points": ["10 of 12 beta customers paid after the pilot"]}],
               "notes": {"say": "Expect two objections; answer each with the proof and the next step."}},
              {"id": "first30", "pattern": "next_steps", "title": "Your first 30 days: three moves",
               "steps": [{"action": "List 10 accounts", "owner": "Each salesperson", "date": "Week 1"},
                         {"action": "Book 5 demos", "owner": "Each salesperson", "date": "Week 2"}],
               "decision": "3 deals each by 31 March; takes about 4 days per salesperson", "decision_label": "Goal",
               "notes": {"say": "The goal is three deals each by the end of March.",
                         "assumptions": ["Per-person activity numbers are proposed."]}}]}


class kSelfTestPlausibility:
    """The checks; Test is the running kSelfTest."""

    def __init__(self, Test):
        try:
            self.Test = Test
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestPlausibility.__init__")

    def CheckGoals(self):
        """3 deals each x 25 people by 31 March against 120 a year is implausible; 1 deal each is not; a missing
        headcount asks for one; an unlisted goal number is named, a listed one is not."""
        if kS.ErrorMode:
            return
        try:
            from _plausibility import kTargetCheck
            Out = kTargetCheck.Warnings(LAUNCH)
            self.Test.Check("plausibility: 3 deals each x 25 people in one month vs 120 a year is implausible",
                            any("= 75 in one month" in W and "120" in W for W in Out), str(Out))
            self.Test.Check("plausibility: a per-person goal the notes' assumptions do not name is flagged",
                            any("not in the spec's facts" in W for W in Out), str(Out))
            Fine = copy.deepcopy(LAUNCH)
            Fine["slides"][3]["decision"] = "1 deal each in pipeline by 31 March; about 4 hours a week"
            Fine["slides"][3]["notes"]["assumptions"] = ["The 31 March goal of 1 deal each is proposed."]
            self.Test.Check("plausibility: 1 deal each (25 a month) listed as assumed gives no warning",
                            kTargetCheck.Warnings(Fine) == [], str(kTargetCheck.Warnings(Fine)))
            NoHead = copy.deepcopy(LAUNCH)
            NoHead.pop("facts")
            self.Test.Check("plausibility: a per-person goal beside a total with no headcount asks for the headcount",
                            any("never states the headcount" in W for W in kTargetCheck.Warnings(NoHead)),
                            str(kTargetCheck.Warnings(NoHead)))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestPlausibility.CheckGoals")
            return

    def CheckObjections(self):
        """An objection answer under six words is flagged once per slide, through build_deck's spec warnings."""
        if kS.ErrorMode:
            return
        try:
            Path = os.path.join(self.Test.Tmp, "plausibility.json")
            self.Test.WriteJson(LAUNCH, Path)
            Code, Out = self.Test.RunScript("build_deck.py", Path, "--plan")
            self.Test.Check("build_deck --plan: a terse objection answer is a spec warning (claim, proof, action)",
                            "slide 3 (compare): 1 objection answer(s) under 6 words" in Out, Out[-1500:])
            self.Test.Check("build_deck --plan: the implausible per-person goal reaches the spec warnings",
                            "slide 4: '3 deals each" in Out, Out[-600:])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestPlausibility.CheckObjections")
            return

    def Run(self):
        """All the checks."""
        if kS.ErrorMode:
            return
        try:
            self.CheckGoals()
            self.CheckObjections()
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestPlausibility.Run")
            return
