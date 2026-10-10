"""Plausibility of invented figures and of objection answers, checked on the spec before a build.

kTargetCheck: a slide that sets a per-person goal ("3 deals each by 31 March") next to a deck-level total target
("120 new customers in 2027") implies a total - goal x people x periods. Benchmark round 10 marked down a launch deck
whose 3 deals each for a 25-person team meant ~75 deals in one month against 120 for the whole year. The check
computes the implied total and warns when it is wildly off (more than half the annual target in one month, or an
implied year over twice the target), asks for the headcount when the deck gives none, and warns when the goal's
number is in neither the spec's facts nor the slide's notes assumptions (an invented target nobody flagged).

kObjectionCheck: an objection / FAQ slide (compare columns or table rows) whose answers are fragments under six
words ("49 USD vs 59-79") reads as terse; each answer states the claim, one proof point and the action, the full
answer goes in the say notes. Both are wired into build_deck.kSpecCheck.Warnings.
"""
import re

from kShared import kS

WORD_NUMBERS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9,
                "ten": 10, "eleven": 11, "twelve": 12, "a": 1, "an": 1}
MONTHS = ("january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
          "november", "december")
MONTH_RE = r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|" \
           r"oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
COUNT = r"(\d+(?:\.\d+)?|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve)"
UNIT = r"((?:new\s+|paid\s+|signed\s+|closed\s+)?(?:deals?|customers?|clients?|accounts?|contracts?|sales|logos?|" \
       r"orders?|subscriptions?|wins?|signups?|sign-ups?))"
PERSON = r"(?:person|head|salesperson|sales\s?person|seller|rep|representative|account\s+executive|AE|member|" \
         r"employee|agent|engineer|analyst)"
EACH_RE = re.compile(COUNT + r"\s+" + UNIT + r"(?:\s+in\s+pipeline)?\s+(?:each|apiece|per\s+" + PERSON + r")\b", re.I)
PEOPLE_RE = re.compile(r"(?<![\w.])(\d+)[\s-]+(?:person\s+)?(?:sales\s?people|salespeople|sales\s+reps?|sellers|reps|"
                       r"representatives|account\s+executives|people|persons|employees|staff|engineers|analysts|"
                       r"team\s+members|agents|members)\b", re.I)
HEADCOUNT_FACT_RE = re.compile(r"team|head|people|seller|rep|sales|staff|employee|members", re.I)
TOTAL_RE = re.compile(r"(?<![\w.])(\d{1,3}(?:[,\s]\d{3})+|\d+)\s+" + UNIT, re.I)
YEAR_CUE_RE = re.compile(r"\b(?:20\d\d|annual|annually|per\s+year|a\s+year|this\s+year|full[- ]year|target|goal)\b",
                         re.I)
BY_DATE_RE = re.compile(r"\bby\s+(?:the\s+end\s+of\s+|end\s+of\s+|\d{1,2}\s+)?(" + MONTH_RE + r")\b", re.I)
MONTH_NAME_RE = re.compile(r"\b(" + MONTH_RE + r")\b", re.I)
OBJECTION_RE = re.compile(r"\b(?:objections?|faqs?|pushback|push-back|questions?\s+(?:you(?:'ll|\s+will)\s+hear|"
                          r"customers\s+ask)|you(?:'ll|\s+will)\s+hear|concerns?)\b", re.I)
MIN_ANSWER_WORDS = 6


class kTargetCheck:
    """Per-person goals checked against the deck's total target, headcount and listed assumptions."""

    @staticmethod
    def Value(Word):
        """The number a count token stands for ('3', '2.5', 'three'); None when it is not one."""
        if kS.ErrorMode:
            return None
        try:
            W = str(Word).strip().lower()
            if W in WORD_NUMBERS:
                return float(WORD_NUMBERS[W])
            W = W.replace(",", "").replace(" ", "")
            return float(W) if re.fullmatch(r"\d+(?:\.\d+)?", W) else None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.Value")
            return None

    @staticmethod
    def SlideTexts(Sl):
        """Every visible text of a slide (notes left out), as one list."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            kTargetCheck.Walk({K: V for K, V in Sl.items() if K not in ("notes", "id", "pattern", "image")}, Out)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.SlideTexts")
            return []

    @staticmethod
    def Walk(Obj, Out):
        """Strings under Obj appended to Out; a metric tile is joined into one text (value + label + note)."""
        if kS.ErrorMode:
            return Out
        try:
            if isinstance(Obj, str):
                Out.append(Obj)
            elif isinstance(Obj, dict):
                if "value" in Obj and "label" in Obj:
                    Out.append(" ".join(str(Obj.get(K, "")) for K in ("value", "label", "note") if Obj.get(K)))
                    return Out
                for V in Obj.values():
                    kTargetCheck.Walk(V, Out)
            elif isinstance(Obj, list):
                for V in Obj:
                    kTargetCheck.Walk(V, Out)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.Walk")
            return Out

    @staticmethod
    def AllText(Spec):
        """All text of the spec, notes and footer included, as one string (for headcount and start month)."""
        if kS.ErrorMode:
            return ""
        try:
            Out = []
            kTargetCheck.Walk(Spec.get("footer", ""), Out)
            for Sl in Spec.get("slides", []):
                if isinstance(Sl, dict):
                    kTargetCheck.Walk(Sl, Out)
            return "\n".join(Out)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.AllText")
            return ""

    @staticmethod
    def Headcount(Spec):
        """The team size the deck states ('25-person sales team', '25 sellers', a facts entry such as
        team_size), or None."""
        if kS.ErrorMode:
            return None
        try:
            Facts = Spec.get("facts") if isinstance(Spec.get("facts"), dict) else {}
            for Name, Val in Facts.items():
                if HEADCOUNT_FACT_RE.search(str(Name)) and isinstance(Val, (int, float)) and not isinstance(Val, bool):
                    return float(Val)
            Found = [float(M.group(1)) for M in PEOPLE_RE.finditer(kTargetCheck.AllText(Spec)) if int(M.group(1)) > 1]
            return max(Found) if Found else None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.Headcount")
            return None

    @staticmethod
    def Family(Unit):
        """The unit's family: deals, customers, contracts and wins all count closed business."""
        if kS.ErrorMode:
            return ""
        try:
            U = re.sub(r"^(?:new|paid|signed|closed)\s+", "", str(Unit).lower())
            return "business" if U[:4] in ("deal", "cust", "clie", "acco", "cont", "sale", "logo", "orde", "subs", "win",
                                           "wins", "sign") else U
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.Family")
            return ""

    @staticmethod
    def Totals(Spec):
        """The deck's total targets: [(number, family, slide)] from texts that name a year or say target/goal and do
        not set a per-person goal ('Our 2027 target: 120 new customers', a tile '120 · New customers in 2027')."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for I, Sl in enumerate(Spec.get("slides", []), 1):
                if not isinstance(Sl, dict):
                    continue
                for Text in kTargetCheck.SlideTexts(Sl):
                    if EACH_RE.search(Text) or not YEAR_CUE_RE.search(Text):
                        continue
                    for M in TOTAL_RE.finditer(Text):
                        N = kTargetCheck.Value(M.group(1))
                        if N is not None and N >= 10 and not 1900 <= N <= 2100:
                            Out.append((N, kTargetCheck.Family(M.group(2)), I))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.Totals")
            return []

    @staticmethod
    def StartMonth(Spec):
        """The month the deck starts from: the first month named on the cover or in the footer, else None."""
        if kS.ErrorMode:
            return None
        try:
            Slides = Spec.get("slides", [])
            Cover = kTargetCheck.SlideTexts(Slides[0]) if Slides and isinstance(Slides[0], dict) else []
            for Text in Cover + [str(Spec.get("footer", ""))]:
                M = MONTH_NAME_RE.search(Text)
                if M and not re.match(r"may\b", M.group(1), re.I):  # 'may' is a verb more often than a month
                    return kTargetCheck.MonthIndex(M.group(1))
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.StartMonth")
            return None

    @staticmethod
    def MonthIndex(Name):
        """0-11 for a month name or its abbreviation."""
        if kS.ErrorMode:
            return None
        try:
            Key = str(Name).lower()[:3]
            for J, M in enumerate(MONTHS):
                if M.startswith(Key):
                    return J
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.MonthIndex")
            return None

    @staticmethod
    def Months(Text, Title, Start):
        """How many months a per-person goal covers: a 'by <month>' deadline from the deck's start month, 'per
        week/month/quarter', '30 days' / 'first month' in the text or the slide title; 12 when nothing says."""
        if kS.ErrorMode:
            return 12.0
        try:
            Both = f"{Text} {Title}"
            if re.search(r"\b(?:per|a|each)\s+week\b|\bweekly\b", Text, re.I):
                return 0.25
            if re.search(r"\b(?:per|a|each)\s+month\b|\bmonthly\b", Text, re.I):
                return 1.0
            if re.search(r"\b(?:per|a|each)\s+quarter\b|\bquarterly\b", Text, re.I):
                return 3.0
            By = BY_DATE_RE.search(Text)
            if By and Start is not None:
                return float((kTargetCheck.MonthIndex(By.group(1)) - Start) % 12 + 1)
            if re.search(r"\b(?:30|thirty)\s+days\b|\bfirst\s+month\b", Both, re.I):
                return 1.0
            if re.search(r"\b(?:90|ninety)\s+days\b|\bfirst\s+quarter\b", Both, re.I):
                return 3.0
            return 12.0
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.Months")
            return 12.0

    @staticmethod
    def Goals(Spec):
        """Per-person goals: [(slide, text, number token, count, unit, months, recurring rate)]. A rate ('per month')
        repeats over the year; a deadline goal ('by 31 March', 'first 30 days') happens once."""
        if kS.ErrorMode:
            return []
        try:
            Start, Out = kTargetCheck.StartMonth(Spec), []
            for I, Sl in enumerate(Spec.get("slides", []), 1):
                if not isinstance(Sl, dict):
                    continue
                Title = str(Sl.get("title", ""))
                for Text in kTargetCheck.SlideTexts(Sl):
                    for M in EACH_RE.finditer(Text):
                        N = kTargetCheck.Value(M.group(1))
                        Clause = re.split(r"[;\n]", Text[M.start():])[0]  # the period that belongs to this goal
                        Rate = bool(re.search(r"\b(?:per|a|each)\s+(?:week|month|quarter)\b|\b(?:weekly|monthly|"
                                              r"quarterly)\b", Clause, re.I))
                        if N is not None:
                            Out.append((I, Text, M.group(1), N, M.group(2), kTargetCheck.Months(Clause, Title, Start),
                                        Rate))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.Goals")
            return []

    @staticmethod
    def Listed(Spec, Sl, Token, Unit, Text):
        """True when a goal's number is a spec fact or the slide's notes assumptions name it (its number, its
        deadline or its unit)."""
        if kS.ErrorMode:
            return True
        try:
            N = kTargetCheck.Value(Token)
            Facts = Spec.get("facts") if isinstance(Spec.get("facts"), dict) else {}
            for Val in Facts.values():
                Vals = Val if isinstance(Val, list) else [Val]
                if any(isinstance(V, (int, float)) and not isinstance(V, bool) and float(V) == N for V in Vals):
                    return True
            Notes = Sl.get("notes") if isinstance(Sl.get("notes"), dict) else {}
            Assumed = " ".join(str(X) for X in Notes.get("assumptions") or []).lower()
            if not Assumed:
                return False
            Words = {str(Token).lower(), f"{N:g}"} | {W for W, V in WORD_NUMBERS.items() if V == N and len(W) > 2}
            if any(re.search(r"(?<![\w.])" + re.escape(W) + r"(?![\w.])", Assumed) for W in Words):
                return True
            By = BY_DATE_RE.search(Text)
            Stem = re.sub(r"^(?:new|paid|signed|closed)\s+", "", Unit.lower()).rstrip("s")
            return bool((By and By.group(1).lower()[:3] in Assumed) or re.search(r"\b" + re.escape(Stem), Assumed))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.Listed")
            return True

    @staticmethod
    def Judge(Goal, People, Totals):
        """The message for one goal against the totals, or '' when it is plausible."""
        if kS.ErrorMode:
            return ""
        try:
            I, Text, Token, N, Unit, Months, Rate = Goal
            Same = [T for T in Totals if T[1] == kTargetCheck.Family(Unit)]
            if not Same:
                return ""
            Total, _, TotalSlide = max(Same)
            Head = f"slide {I}: '{Text}' sets {Token} {Unit} per person"
            if People is None:
                return (f"{Head} beside a total target of {Total:g} (slide {TotalSlide}) but the deck never states "
                        "the headcount - say how many people (e.g. '25 sellers') so the implied total can be "
                        "checked, and confirm goal x people fits the target")
            Implied = N * People
            Year = Implied * 12.0 / Months if Rate or Months >= 12 else Implied
            if (Months <= 1 and Implied > Total / 2) or Year > 2 * Total or (Months < 12 and Implied > Total):
                Span = "in one month" if Months <= 1 else f"in {Months:g} months"
                Pace = f" (~{Year:g} a year at that rate)" if Rate else ""
                return (f"{Head}: x {People:g} people = {Implied:g} {Span}{Pace} against the total "
                        f"target of {Total:g} (slide {TotalSlide}) - implausible; derive the per-person goal from the "
                        f"target (e.g. {Total:g} / {People:g} = {Total / People:.1f} each for the year) and give its "
                        "basis in the notes' assumptions")
            return ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.Judge")
            return ""

    @staticmethod
    def Warnings(Spec):
        """The goal warnings of a spec, as messages."""
        if kS.ErrorMode:
            return []
        try:
            Goals = kTargetCheck.Goals(Spec)
            if not Goals:
                return []
            People, Totals, Out = kTargetCheck.Headcount(Spec), kTargetCheck.Totals(Spec), []
            Slides = Spec.get("slides", [])
            for Goal in Goals:
                Msg = kTargetCheck.Judge(Goal, People, Totals)
                if Msg:
                    Out.append(Msg)
                I, Text, Token, _, Unit, _, _ = Goal
                if not kTargetCheck.Listed(Spec, Slides[I - 1], Token, Unit, Text):
                    Out.append(f"slide {I}: the goal '{Token} {Unit}' per person is not in the spec's facts and the "
                               "slide's notes 'assumptions' do not name it - an invented target; list it there "
                               "('3 deals each is proposed, not in the brief') or derive it from a stated target")
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTargetCheck.Warnings")
            return []


class kObjectionCheck:
    """Objection / FAQ slides whose answers are cut to fragments."""

    @staticmethod
    def Answers(Sl):
        """(objection, answer) pairs of an objection slide: a compare column's heading and each point, or a table
        row's first cell and its other cells joined."""
        if kS.ErrorMode:
            return []
        try:
            Pat, Out = Sl.get("pattern"), []
            if Pat == "compare":
                for C in Sl.get("columns", []):
                    if isinstance(C, dict):
                        Out += [(str(C.get("heading", "")), str(P)) for P in C.get("points", [])]
            elif Pat == "table":
                for R in Sl.get("rows", []):
                    if isinstance(R, list) and len(R) >= 2:
                        Out.append((str(R[0]), " ".join(str(X) for X in R[1:])))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kObjectionCheck.Answers")
            return []

    @staticmethod
    def IsObjection(Sl):
        """True when a slide's title (or a table's header) says it answers objections, FAQs or concerns."""
        if kS.ErrorMode:
            return False
        try:
            Text = " ".join([str(Sl.get("title", "")), str(Sl.get("kicker", ""))] +
                            [str(X) for X in Sl.get("header") or []])
            return bool(OBJECTION_RE.search(Text))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kObjectionCheck.IsObjection")
            return False

    @staticmethod
    def Warnings(Spec):
        """One message per objection slide with answers under MIN_ANSWER_WORDS words."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for I, Sl in enumerate(Spec.get("slides", []), 1):
                if not isinstance(Sl, dict) or not kObjectionCheck.IsObjection(Sl):
                    continue
                Short = [A for _, A in kObjectionCheck.Answers(Sl)  # words only: a lone '%' or '-' is not one
                         if len(re.findall(r"\w[\w'.,/%-]*", A)) < MIN_ANSWER_WORDS]
                if Short:
                    Out.append(f"slide {I} ({Sl.get('pattern')}): {len(Short)} objection answer(s) under "
                               f"{MIN_ANSWER_WORDS} words ('{Short[0]}') read as fragments - give each answer the "
                               "claim, one proof point and the action ('Saves 4.5 h a week per analyst: book a "
                               "demo'), the full answer in the say notes (reference/CONTENT.md)")
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kObjectionCheck.Warnings")
            return []
