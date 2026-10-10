"""Self-test checks for the fewer-loops round: ready-to-apply edits (_suggest.py), --plan's dry build, --apply and
the deduplicated --check summary. Run by selftest.py (kSelfTestSuggest(Test).Run()); uses its Check / RunScript
and temp folder.

The fixture is a first draft with the findings that cost the benchmark makers rebuilds: a title too long to fit, a
point over its limit, a bullets list with long items, a quiz answer slide with the default title, an ask without a
cost, a risk card that does not fit, and a figure token that must survive the rewrite.
"""
import json
import os

from kShared import kS

DRAFT = {"direction": "consulting-blue", "footer": "Suggest test", "sources": ["the brief"],
         "facts": {"revenue_q": [4.1, 4.6, 5.2, 5.9]},
         "slides": [
             {"id": "growth", "pattern": "statement",
              "title": "Revenue growth is accelerating strongly in every single quarter and our customers are staying "
                       "with us longer than ever",
              "points": ["Revenue reached {sum:revenue_q} MUSD this year", "Every quarter beat the one before it"],
              "notes": {"say": "Revenue grew every quarter this year, according to the finance figures in the brief."}},
             {"id": "flags", "pattern": "bullets", "title": "Five red flags give most phishing away",
              "items": ["Urgency or threats that push you to act right now before you have a chance to think it over",
                        "Unexpected link or attachment", "Requests for passwords or payment",
                        "Generic greeting, odd tone"],
              "notes": {"say": "Look for these red flags; one is a reason to slow down and two mean report it."}},
             {"id": "hiring", "pattern": "big_number", "title": "Hiring is behind plan at 14 of 20 roles",
              "number": "14/20", "caption": "Roles filled against plan.",
              "points": ["The six open roles are really slowing down delivery of customer projects in every region"],
              "notes": {"say": "We filled 14 of 20 planned roles, which slows delivery, according to the brief."}},
             {"id": "quiz", "pattern": "quiz", "title": "The CEO asks you to buy gift cards urgently. What do you do?",
              "options": [{"text": "Buy them"}, {"text": "Report and verify by phone", "correct": True}],
              "explain": "Verify on a known channel; urgency plus payment is a red flag.", "reveal": "slide",
              "notes": {"say": "Take ten seconds and decide which answer you would pick, then raise your hand."}},
             {"id": "ask", "pattern": "statement", "title": "Approve three extra sales engineers",
              "decision": "Approve three additional sales engineers starting in Q1 2027", "owner": "VP Sales",
              "date": "Q1 2027", "points": ["Shorter sales cycles", "Convert demand"],
              "notes": {"say": "We ask you to approve three sales engineers to shorten our sales cycles."}}]}


class kSelfTestSuggest:
    """The checks; Test is the running kSelfTest."""

    def __init__(self, Test):
        try:
            self.Test = Test
            self.Dir = os.path.join(Test.Tmp, "suggest")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestSuggest.__init__")

    def CheckShorten(self):
        """kShorten removes or abbreviates words only, and keeps the sense-changing cut as a proposal."""
        if kS.ErrorMode:
            return
        try:
            from _suggest import REVIEW, kShorten, kTextFit
            New, Lost, _ = kShorten.Shorten("Sales capacity is very clearly the tightest constraint that we have "
                                            "right now", kTextFit(45))
            self.Test.Check("suggest: filler-only shortening loses no fact and needs no notes copy",
                            New == "Sales capacity is the tightest constraint" and not Lost, repr(New))
            Old = "Use the report button in Outlook, which sends it to the security team automatically"
            New, Lost, How = kShorten.Shorten(Old, kTextFit(60))
            Words = set(Old.replace(",", "").split())
            self.Test.Check("suggest: a cut at a clause boundary only removes words (every word kept was there)",
                            New == "Use the report button in Outlook" and Lost and How != REVIEW
                            and set(New.split()) <= Words, repr(New))
            New, _, How = kShorten.Shorten("A sender address that does not match the company the message claims "
                                           "to come from", kTextFit(60))
            self.Test.Check("suggest: a cut that may change the sense is marked for review (proposed, not applied)",
                            How == REVIEW and New is not None and len(New) <= 60, f"{How}: {New!r}")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestSuggest.CheckShorten")
            return

    def CheckPlan(self, Path):
        """--plan lists measured edits and questions and writes nothing; --plan --apply writes them."""
        if kS.ErrorMode:
            return
        try:
            with open(Path, encoding="utf-8") as Fh:
                Before = Fh.read()
            Code, Out = self.Test.RunScript("build_deck.py", Path, "--plan")
            with open(Path, encoding="utf-8") as Fh:
                Unchanged = Fh.read() == Before
            self.Test.Check("build_deck --plan: each fixable finding gets a ready edit (suggest: old -> new) and "
                            "what needs content stays a question; the spec is not touched",
                            "suggest: slide 1 (growth) title:" in Out and "suggest: slide 2 (flags) items[0]:" in Out
                            and "suggest: slide 3 (hiring) points[0]:" in Out
                            and "suggest: slide 4 (quiz) answer_title:" in Out
                            and "question: slide 5 (ask) needs the ask's cost" in Out and Unchanged, Out[-1500:])
            self.Test.Check("build_deck --plan: the summary line counts the ready edits and the questions",
                            "ready edit(s)" in Out and "1 question(s)" in Out, Out[-400:])
            Code, Out = self.Test.RunScript("build_deck.py", Path, "--plan", "--apply")
            with open(Path, encoding="utf-8") as Fh:
                Spec = json.load(Fh)
            Slides = {Sl["id"]: Sl for Sl in Spec["slides"]}
            Moved = " ".join(Slides["flags"]["notes"].get("moved", []))
            self.Test.Check("build_deck --plan --apply: the edits are written, the original kept in notes.moved, a "
                            ".before-apply copy beside it, figure tokens untouched, and the plan re-run is clean of "
                            "spec errors", "applied:" in Out and "Urgency or threats that push you" in Moved
                            and len(Slides["flags"]["items"][0]) <= 80
                            and os.path.exists(Path.replace(".json", ".before-apply.json"))
                            and "{sum:revenue_q}" in Slides["growth"]["points"][0]
                            and "plan: 0 spec error(s)" in Out and Slides["quiz"].get("answer_title"), Out[-1500:])
            self.Test.Check("build_deck --plan --apply: a quiz answer title comes from the slide's own 'explain', "
                            "never a new sentence", str(Slides["quiz"].get("answer_title", "")).startswith("Verify on"),
                            str(Slides["quiz"].get("answer_title")))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestSuggest.CheckPlan")
            return

    def CheckApplyCheck(self, Path):
        """--check --apply writes the auto-fixes back, so the next build needs none; the summary has no repeats."""
        if kS.ErrorMode:
            return
        try:
            with open(Path, encoding="utf-8") as Fh:
                Spec = json.load(Fh)
            Spec["slides"].append({"id": "risks", "pattern": "risks", "title": "Two risks could slow the rollout",
                                   "notes": {"say": "Two risks, both with a mitigation we can start this month."},
                                   "risks": [{"risk": "Sales cycles lengthen", "likelihood": "high", "impact": "high",
                                              "mitigation": "Run proofs of concept much earlier in the cycle, so that "
                                                            "buyers can see real value before procurement starts"},
                                             {"risk": "Hiring stalls", "likelihood": "medium", "impact": "medium",
                                              "mitigation": "Prioritise customer-facing roles first and use referral "
                                                            "bonuses to widen the pipeline this quarter"}]})
            Spec["slides"].append({"id": "kpi", "pattern": "kpi", "title": "Growth accelerated every quarter",
                                   "notes": {"say": "Four numbers sum up the year, all from the brief's reporting."},
                                   "metrics": [{"value": "19.8 MUSD", "label": "Revenue"},
                                               {"value": "112 %", "label": "Retention"},
                                               {"value": "1.2 %", "label": "Churn"},
                                               {"value": "191", "label": "New customers"}]})
            self.Test.WriteJson(Spec, Path)
            Deck = os.path.join(self.Dir, "draft.pptx")
            Code, Out = self.Test.RunScript("build_deck.py", Path, "--out", Deck, "--check", "--apply")
            Line = [L for L in Out.splitlines() if L.startswith("CHECK-JSON ")]
            Data = json.loads(Line[-1][11:]) if Line else {}
            Codes = [(P["slide"], P["code"]) for P in Data.get("problems", [])]
            self.Test.Check("build_deck --check: no finding is listed twice (a lint finding its spec warning already "
                            "states is dropped)", len(Codes) == len(set(Codes))
                            and (5, "ask_without_cost") not in Codes and "suggestions" in Data, str(Codes))
            Code2, Out2 = self.Test.RunScript("build_deck.py", Path, "--out", Deck, "--check")
            self.Test.Check("build_deck --check --apply: the auto-fixes are written into the spec, so the next build "
                            "has none to make", "auto-fixes: 0" in Out2 and "(auto-fix" in Out and "applied:" in Out,
                            Out[-600:] + " || " + Out2[-300:])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestSuggest.CheckApplyCheck")
            return

    def Run(self):
        """Every check of this module."""
        if kS.ErrorMode:
            return
        try:
            os.makedirs(self.Dir, exist_ok=True)
            Path = os.path.join(self.Dir, "draft.json")
            self.Test.WriteJson(DRAFT, Path)
            self.CheckShorten()
            self.CheckPlan(Path)
            self.CheckApplyCheck(Path)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestSuggest.Run")
            return
