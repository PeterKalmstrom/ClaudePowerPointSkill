"""Computed figures: check the numbers a deck derives (totals, averages, changes, shares) against the data in it.

A slide that says "21.8 MUSD in 2026 - sum of four quarters" next to a chart of 4.1, 4.6, 5.2 and 5.9 is wrong
(the sum is 19.8), and nothing else in the skill notices. kFigures reads every text value, finds the numbers that
read as derived - a total ("sum", "total", "38 + 41 + 52 + 60"), an average, a change ("+44 %", "up 44 %"), a share
("70 % of ..." beside "14/20") or an equation ("44 % = 5.9 / 4.1 - 1") - and compares each with what the numbers
elsewhere in the same deck give: chart series, metric trends, cost and table columns and the spec's "facts". A
figure that matches nothing but is close to one of them (a slip, not a different quantity) is reported.

It also computes figures for the spec: {sum}, {total}, {average}, {change}, {share}, {first}, {last} and {count} in
any text become the number, from the metric's trend, the slide's chart series or cost rows, the big number's share,
or a named deck fact ({sum:revenue_q}); {name} is a scalar deck fact. Used by build_deck.py (spec warnings, tokens)
and lint_deck.py (figure_mismatch on a built deck).
"""
import ast
import re

from kShared import kS

NBSP = " "
MINUS = "−"
NUM_RE = re.compile(r"(?<![\w.,])([+\-−]?)((?:\d{1,3}(?:[,   ]\d{3})+|\d+)(?:[.,]\d+)?)(?!\w)"
                    r"(\s?%)?")
EXPR_RE = re.compile(r"(?<![\w.,])\d+(?:[.,]\d+)?(?:\s*[-+*/×÷−]\s*\d+(?:[.,]\d+)?)+(?![\w.,]\d)")
PAIR_RE = re.compile(r"(?<![\w.,])(\d+(?:[.,]\d+)?)\s?%?\s*(?:to|→|->|–)\s*(\d+(?:[.,]\d+)?)(?!\w)")
FRACTION_RE = re.compile(r"(?<![\w.,])(\d+)\s*(?:/|of|out of)\s*(\d+)(?![\w.,]\d)")
CLAUSE_RE = re.compile(r"[;\n•]|(?<!\d)[.,](?!\d)|[.,](?=\s)|\s—\s|\s-\s(?!\d)")
TOTAL_CUE = re.compile(r"\b(?:sum|summed|total|totals|combined|altogether|cumulative|full[- ]year|year[- ]to[- ]date|"
                       r"ytd)\b", re.I)
AVERAGE_CUE = re.compile(r"\b(?:average|averages|avg|mean|on average)\b", re.I)
UP_WORDS = ("up", "grew", "grow", "grows", "growing", "growth", "rose", "rise", "rises", "risen", "increase",
            "increased", "increases", "gained", "gain", "gains", "higher", "more", "jumped", "climbed", "improved",
            "uplift")
DOWN_WORDS = ("down", "fell", "fall", "falls", "fallen", "drop", "dropped", "drops", "decline", "declined",
              "declines", "decrease", "decreased", "lower", "less", "fewer", "cut", "cuts", "shrank", "lost",
              "reduced", "reduction")
FILLER = ("by", "about", "nearly", "almost", "roughly", "around", "some", "over", "~", "just", "only")
TOKEN_RE = re.compile(r"\{(sum|total|average|change|share|first|last|count)(?::([A-Za-z_][\w.-]*))?\}|"
                      r"\{([A-Za-z_][\w.-]*)\}")
NEAR = {"total": 0.35, "average": 0.35, "change": 0.12, "share": 0.35}  # 'close but wrong': a slip, not another figure


class kFigures:
    """Derived-figure checks and figure tokens. Stateless."""

    # ------------------------------------------------------------------ numbers

    @staticmethod
    def Numbers(Text):
        """Every number in Text: [{value, places, percent, signed, start, end, text}] (no-break spaces and commas
        as thousands separators, a comma or point as the decimal mark)."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for Mt in NUM_RE.finditer(str(Text)):
                Raw = Mt.group(2)
                Grouped = re.search(r"\d[,   ]\d{3}(?!\d)", Raw) and not re.fullmatch(r"\d+[.,]\d+", Raw)
                Clean = re.sub(r"[,   ]", "", Raw) if Grouped else Raw.replace(",", ".")
                Value = float(Clean)
                Places = len(Clean.split(".")[1]) if "." in Clean else 0
                Sign = Mt.group(1)
                if Sign in ("-", MINUS):
                    Value = -Value
                Out.append({"value": Value, "places": Places, "percent": bool(Mt.group(3)), "signed": bool(Sign),
                            "start": Mt.start(), "end": Mt.end(), "text": Mt.group(0).strip()})
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Numbers")
            return []

    @staticmethod
    def IsYear(Num):
        """True for a number that is a year (2026), not a quantity."""
        if kS.ErrorMode:
            return False
        try:
            return (not Num["percent"] and not Num["signed"] and Num["places"] == 0
                    and 1900 <= Num["value"] <= 2100)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.IsYear")
            return False

    @staticmethod
    def Places(Values):
        """The most decimals any value in the list is written with."""
        if kS.ErrorMode:
            return 0
        try:
            return max([len(f"{V:g}".split(".")[1]) if "." in f"{V:g}" else 0 for V in Values] or [0])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Places")
            return 0

    @staticmethod
    def Format(Value, Places):
        """A number as a slide shows it: fixed decimals, thousands grouped with a no-break space from 10 000."""
        if kS.ErrorMode:
            return ""
        try:
            Text = f"{Value:,.{Places}f}" if abs(Value) >= 10000 else f"{Value:.{Places}f}"
            return Text.replace(",", NBSP).replace("-", MINUS)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Format")
            return ""

    @staticmethod
    def Evaluate(Expr):
        """The value of a plain arithmetic expression ('5.9 / 4.1 - 1', '38 + 41'), or None."""
        if kS.ErrorMode:
            return None
        try:
            Text = str(Expr).replace("×", "*").replace("÷", "/").replace(MINUS, "-").replace(",", ".")
            if not re.fullmatch(r"\s*[-+]?\d+(?:\.\d+)?(?:\s*[-+*/]\s*\d+(?:\.\d+)?)*\s*", Text):
                return None  # not plain arithmetic: not a figure to check
            return kFigures.Node(ast.parse(Text, mode="eval").body)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Evaluate")
            return None

    @staticmethod
    def Node(Nd):
        """One node of an arithmetic expression: numbers, + - * / and a leading sign only (None otherwise)."""
        if kS.ErrorMode:
            return None
        try:
            if isinstance(Nd, ast.Constant) and isinstance(Nd.value, (int, float)):
                return float(Nd.value)
            if isinstance(Nd, ast.UnaryOp) and isinstance(Nd.op, (ast.USub, ast.UAdd)):
                Val = kFigures.Node(Nd.operand)
                return None if Val is None else (-Val if isinstance(Nd.op, ast.USub) else Val)
            if isinstance(Nd, ast.BinOp) and isinstance(Nd.op, (ast.Add, ast.Sub, ast.Mult, ast.Div)):
                Left, Right = kFigures.Node(Nd.left), kFigures.Node(Nd.right)
                if Left is None or Right is None or (isinstance(Nd.op, ast.Div) and Right == 0):
                    return None
                if isinstance(Nd.op, ast.Add):
                    return Left + Right
                if isinstance(Nd.op, ast.Sub):
                    return Left - Right
                if isinstance(Nd.op, ast.Mult):
                    return Left * Right
                return Left / Right
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Node")
            return None

    # ------------------------------------------------------------------ what the data gives

    @staticmethod
    def Candidates(Kind, Series):
        """[(value, basis)] that a derived figure of Kind can be, from every series: its sum, its mean, or its
        change from the first to the last value in percent."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for Sr in Series:
                Vals = [float(V) for V in Sr["values"] if isinstance(V, (int, float)) and not isinstance(V, bool)]
                if len(Vals) < 2:
                    continue
                Name = f"'{Sr['name']}'" + (f" (slide {Sr['slide']})" if Sr.get("slide") else " (deck facts)")
                Shown = " + ".join(f"{V:g}" for V in Vals) if len(Vals) <= 6 else f"{len(Vals)} values"
                if Kind == "total":
                    Out.append((sum(Vals), f"the sum of {Name}, {Shown}, is"))
                elif Kind == "average":
                    Out.append((sum(Vals) / len(Vals), f"the average of {Name} is"))
                elif Kind == "change" and Vals[0]:
                    Out.append(((Vals[-1] / Vals[0] - 1) * 100, f"{Name} from {Vals[0]:g} to {Vals[-1]:g} is"))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Candidates")
            return []

    @staticmethod
    def Matches(Claim, Value):
        """True when the written figure is Value as rounded (or cut) to the decimals it is written with."""
        if kS.ErrorMode:
            return True
        try:
            return abs(Claim["value"] - Value) <= 10 ** -Claim["places"] * 1.0001 + 1e-9
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Matches")
            return True

    @staticmethod
    def Judge(Claim, Kind, Local, Others):
        """None when the claim matches a value it can be; else the message for the nearest one. Local values (an
        explicit '38 + 41 + 52 + 60', a 'from 4.1 to 5.9' beside it) must match when there are any; the deck's
        series count only when one is close (a slip), so an unrelated figure is not reported."""
        if kS.ErrorMode:
            return None
        try:
            Pool = Local + Others
            if any(kFigures.Matches(Claim, V) for V, _ in Pool):
                return None
            Near = [(abs(Claim["value"] - V) / max(abs(V), 1e-9), V, B) for V, B in Others
                    if abs(Claim["value"] - V) <= NEAR[Kind] * abs(V) and (V >= 0) == (Claim["value"] >= 0)]
            Pick = None
            if Local:
                Pick = min(((abs(Claim["value"] - V), V, B) for V, B in Local), key=kFigures.First)
            elif Near:
                Pick = min(Near, key=kFigures.First)
            if Pick is None:
                return None
            _, V, Basis = Pick
            Places = max(Claim["places"], 1 if abs(V - round(V)) > 1e-9 else 0)
            Unit = " %" if Kind in ("change", "share") else ""
            Sign = "+" if Kind == "change" and V > 0 else ""
            Word = {"total": "a total", "average": "an average", "change": "a change", "share": "a share"}[Kind]
            return (f"'{Claim['text']}' reads as {Word}, but {Basis} {Sign}{kFigures.Format(V, Places)}{Unit}; "
                    "compute it (a {" + ("sum" if Kind == "total" else Kind) + "} token) or fix the figure")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Judge")
            return None

    @staticmethod
    def First(Item):
        """Sort key: the first element."""
        if kS.ErrorMode:
            return 0
        try:
            return Item[0]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.First")
            return 0

    # ------------------------------------------------------------------ claims in one piece of text

    @staticmethod
    def ChangeSign(Text, Num):
        """+1 / -1 when a percent reads as a change (a sign, or a change word right before or after it), else 0."""
        if kS.ErrorMode:
            return 0
        try:
            if Num["signed"]:
                return -1 if Num["value"] < 0 else 1
            Before = re.findall(r"[\w~]+", Text[:Num["start"]].lower())
            while Before and Before[-1] in FILLER:
                Before.pop()
            After = re.findall(r"\w+", Text[Num["end"]:].lower())[:1]
            for Word in (Before[-1:] + After):
                if Word in UP_WORDS:
                    return 1
                if Word in DOWN_WORDS:
                    return -1
            return 0
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.ChangeSign")
            return 0

    @staticmethod
    def Clauses(Text):
        """Text split into clauses (sentences, list items, parts between semicolons and commas)."""
        if kS.ErrorMode:
            return []
        try:
            return [C for C in CLAUSE_RE.split(str(Text)) if C and C.strip()]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Clauses")
            return []

    @staticmethod
    def Spans(Text, Pattern):
        """(start, end) of every match of Pattern in Text."""
        if kS.ErrorMode:
            return []
        try:
            return [(Mt.start(), Mt.end()) for Mt in Pattern.finditer(Text)]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Spans")
            return []

    @staticmethod
    def Inside(Num, Spans):
        """True when the number lies inside one of the spans."""
        if kS.ErrorMode:
            return False
        try:
            return any(A <= Num["start"] and Num["end"] <= B + 2 for A, B in Spans)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Inside")
            return False

    @staticmethod
    def Equations(Text):
        """Messages for equations in Text whose two sides differ: '44 % = 5.9 / 4.1 - 1', '38 + 41 = 79'."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for Mt in EXPR_RE.finditer(Text):
                Value = kFigures.Evaluate(Mt.group(0))
                if Value is None:
                    continue
                After = re.match(r"\s*=\s*([+\-−]?\d[\d., ]*\s?%?)", Text[Mt.end():])
                Before = re.search(r"([+\-−]?\d[\d., ]*\s?%?)\s*=\s*$", Text[:Mt.start()])
                Side = (After.group(1) if After else None) or (Before.group(1) if Before else None)
                Nums = kFigures.Numbers(Side) if Side else []
                if not Nums:
                    continue
                Claim = Nums[0]
                Shown = Value * 100 if Claim["percent"] and abs(Value) < 10 and "%" not in Mt.group(0) else Value
                if not kFigures.Matches(Claim, Shown):
                    Places = max(Claim["places"], 1)
                    Out.append(f"'{Mt.group(0).strip()} = {Claim['text']}' does not add up: the left side is "
                               f"{kFigures.Format(Shown, Places)}{' %' if Claim['percent'] else ''}")
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Equations")
            return []

    @staticmethod
    def Sums(Text):
        """[(value, basis)] for every plain addition written out in Text ('38 + 41 + 52 + 60')."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for Mt in EXPR_RE.finditer(Text):
                Expr = Mt.group(0)
                if re.search(r"[-*/×÷−]", Expr):
                    continue
                Value = kFigures.Evaluate(Expr)
                if Value is not None:
                    Out.append((Value, f"{Expr.strip()} is"))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Sums")
            return []

    @staticmethod
    def Pairs(Text):
        """[(change in percent, basis)] for every 'from A to B' / 'A to B' / 'A -> B' written in Text."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for Mt in PAIR_RE.finditer(Text):
                A, B = float(Mt.group(1).replace(",", ".")), float(Mt.group(2).replace(",", "."))
                if A:
                    Out.append(((B / A - 1) * 100, f"{Mt.group(1)} to {Mt.group(2)} is"))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Pairs")
            return []

    @staticmethod
    def Fractions(Text):
        """[(share in percent, basis)] for every 'A of B' / 'A/B' (A <= B, B not a year) in Text."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for Mt in FRACTION_RE.finditer(str(Text)):
                A, B = int(Mt.group(1)), int(Mt.group(2))
                if 0 < B and A <= B and not 1900 <= B <= 2100:
                    Out.append((100 * A / B, f"{A} of {B} is"))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Fractions")
            return []

    @staticmethod
    def CheckOne(Ctx, Series, Shares):
        """Messages for one context: {text, primary (the figure being stated, or None: every number in a clause
        with a cue), sums (explicit totals it must match)}."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            Text = str(Ctx["text"])
            Primary = Ctx.get("primary")
            Parts = [(Text, str(Primary))] if Primary is not None else [(C, C) for C in kFigures.Clauses(Text)]
            if Primary is None:
                Out += kFigures.Equations(Text)
            for Cue, Said in Parts:
                Sums = kFigures.Sums(Cue) + [(V, "the column adds up to") for V in Ctx.get("sums", [])]
                Pairs = kFigures.Pairs(Cue)
                Skip = kFigures.Spans(Said, EXPR_RE) + kFigures.Spans(Said, PAIR_RE) + kFigures.Spans(Said, FRACTION_RE)
                Total, Average = bool(TOTAL_CUE.search(Cue)), bool(AVERAGE_CUE.search(Cue))
                for Num in kFigures.Numbers(Said):
                    if kFigures.IsYear(Num) or kFigures.Inside(Num, Skip):
                        continue
                    Msg = None
                    if Num["percent"]:
                        Sign = kFigures.ChangeSign(Said, Num) if Primary is None else (
                            (-1 if Num["value"] < 0 else 1) if Num["signed"] else 0)
                        if Sign:
                            Claim = dict(Num, value=abs(Num["value"]) * Sign)
                            Msg = kFigures.Judge(Claim, "change", Pairs, kFigures.Candidates("change", Series))
                        elif re.match(r"\s*of\b", Said[Num["end"]:]) and Shares:
                            Msg = kFigures.Judge(Num, "share", [], Shares)
                    elif Sums or Total:
                        Msg = kFigures.Judge(Num, "total", Sums, kFigures.Candidates("total", Series))
                    elif Average:
                        Msg = kFigures.Judge(Num, "average", [], kFigures.Candidates("average", Series))
                    if Msg:
                        Out.append(Msg)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.CheckOne")
            return []

    @staticmethod
    def Check(Contexts, Series):
        """[{slide, where, message}] for every derived figure in Contexts that the deck's Series contradict.
        Shares are judged against the fractions on the same slide."""
        if kS.ErrorMode:
            return []
        try:
            Shares = {}
            for Ctx in Contexts:
                Shares.setdefault(Ctx["slide"], []).extend(kFigures.Fractions(Ctx["text"]))
            Out, Seen = [], set()
            for Ctx in Contexts:
                for Msg in kFigures.CheckOne(Ctx, Series, Shares.get(Ctx["slide"], [])):
                    Key = (Ctx["slide"], Msg)
                    if Key not in Seen:
                        Seen.add(Key)
                        Out.append({"slide": Ctx["slide"], "where": Ctx.get("where", ""), "message": Msg,
                                    "shape": Ctx.get("shape")})
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Check")
            return []

    # ------------------------------------------------------------------ a spec

    @staticmethod
    def Facts(Spec):
        """The spec's deck-level facts: {name: number or [numbers]} (anything else is left out)."""
        if kS.ErrorMode:
            return {}
        try:
            Facts = Spec.get("facts") if isinstance(Spec.get("facts"), dict) else {}
            Out = {}
            for Name, Val in Facts.items():
                if isinstance(Val, (int, float)) and not isinstance(Val, bool):
                    Out[str(Name)] = Val
                elif isinstance(Val, list) and all(isinstance(V, (int, float)) and not isinstance(V, bool)
                                                   for V in Val):
                    Out[str(Name)] = Val
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Facts")
            return {}

    @staticmethod
    def TableColumns(Sl, I):
        """(series, total-row contexts) of a table slide: each numeric column is a series; a row whose first cell
        says total/sum is not data but a claim that its cells add up the column."""
        if kS.ErrorMode:
            return [], []
        try:
            Header = [str(X) for X in Sl.get("header") or []]
            Rows = [R for R in Sl.get("rows") or [] if isinstance(R, list)]
            Data = [R for R in Rows if not (R and TOTAL_CUE.search(str(R[0])))]
            Totals = [R for R in Rows if R and TOTAL_CUE.search(str(R[0]))]
            Series, Ctx = [], []
            for C in range(1, len(Header)):
                Vals = []
                for R in Data:
                    Nums = kFigures.Numbers(R[C]) if C < len(R) else []
                    if len(Nums) != 1:
                        Vals = []
                        break
                    Vals.append(Nums[0]["value"])
                if len(Vals) >= 2:
                    Series.append({"slide": I, "name": Header[C], "values": Vals})
                    for R in Totals:
                        if C < len(R):
                            Ctx.append({"slide": I, "where": f"table total '{Header[C]}'", "text": str(R[C]),
                                        "primary": str(R[C]), "sums": [sum(Vals)]})
            return Series, Ctx
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.TableColumns")
            return [], []

    @staticmethod
    def Strings(Obj, Path, Out):
        """Every (path, text) under Obj, appended to Out (lists and objects walked)."""
        if kS.ErrorMode:
            return Out
        try:
            if isinstance(Obj, str):
                Out.append((Path, Obj))
            elif isinstance(Obj, dict):
                for K, V in Obj.items():
                    kFigures.Strings(V, f"{Path}.{K}" if Path else str(K), Out)
            elif isinstance(Obj, list):
                for J, V in enumerate(Obj):
                    kFigures.Strings(V, f"{Path}[{J}]", Out)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Strings")
            return Out

    @staticmethod
    def SpecData(Spec):
        """(contexts, series) of a spec: every text value, metric tiles as one context each (the value is the
        figure, its label and note say what it is), and every series of numbers in the deck."""
        if kS.ErrorMode:
            return [], []
        try:
            Contexts, Series = [], []
            for Name, Val in kFigures.Facts(Spec).items():
                if isinstance(Val, list):
                    Series.append({"slide": 0, "name": Name, "values": Val})
            for I, Sl in enumerate(Spec.get("slides", []), 1):
                if not isinstance(Sl, dict):
                    continue
                Pat = Sl.get("pattern")
                for Sr in Sl.get("series") or []:
                    if isinstance(Sr, dict):
                        Series.append({"slide": I, "name": str(Sr.get("name", "series")), "values": Sr.get("values") or []})
                if Pat == "cost_table":
                    Series.append({"slide": I, "name": "the cost rows", "values": [R.get("amount") for R in Sl.get("rows", [])
                                                                                   if isinstance(R, dict)]})
                if Pat == "table":
                    Cols, Totals = kFigures.TableColumns(Sl, I)
                    Series += Cols
                    Contexts += Totals
                Tiles = [(f"metrics[{J}]", Mt) for J, Mt in enumerate(Sl.get("metrics") or []) if isinstance(Mt, dict)]
                if isinstance(Sl.get("figure"), dict):
                    Tiles.append(("figure", Sl["figure"]))
                for Where, Mt in Tiles:
                    if Mt.get("trend"):
                        Series.append({"slide": I, "name": str(Mt.get("label", Where)), "values": Mt["trend"]})
                    Text = " · ".join(str(Mt.get(K, "")) for K in ("value", "label", "note") if Mt.get(K))
                    Contexts.append({"slide": I, "where": f"{Pat} {Where}", "text": Text,
                                     "primary": str(Mt.get("value", ""))})
                if Pat == "big_number":
                    Num = f"{Sl.get('number', '')}{(' ' + str(Sl['unit'])) if Sl.get('unit') else ''}"
                    Contexts.append({"slide": I, "where": "big_number number", "primary": Num,
                                     "text": f"{Num} · {Sl.get('caption', '')}"})
                Skip = ("metrics", "figure", "series", "categories", "rows", "header", "id", "pattern", "image",
                        "number", "unit", "highlight", "type", "number_format", "kicker", "direction")
                for Key, Val in Sl.items():
                    if Key in Skip:
                        continue
                    for Path, Text in kFigures.Strings(Val, Key, []):
                        Contexts.append({"slide": I, "where": f"{Pat} {Path}", "text": Text, "primary": None})
                if Pat == "table":
                    for R, Row in enumerate(Sl.get("rows") or []):
                        Contexts.append({"slide": I, "where": f"table rows[{R}]", "primary": None,
                                         "text": " · ".join(str(X) for X in Row) if isinstance(Row, list) else ""})
            return Contexts, Series
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.SpecData")
            return [], []

    @staticmethod
    def SpecWarnings(Spec):
        """'figure: slide N (where): ...' for every derived figure in the spec that its own numbers contradict."""
        if kS.ErrorMode:
            return []
        try:
            Contexts, Series = kFigures.SpecData(Spec)
            return [f"figure: slide {F['slide']} ({F['where']}): {F['message']}"
                    for F in kFigures.Check(Contexts, Series)]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.SpecWarnings")
            return []

    # ------------------------------------------------------------------ tokens: figures the builder computes

    @staticmethod
    def SlideSource(Sl):
        """The numbers a token on this slide computes from: the first chart series, the cost rows, or None."""
        if kS.ErrorMode:
            return None
        try:
            Ser = [S for S in Sl.get("series") or [] if isinstance(S, dict)]
            if Ser:
                return [V for V in Ser[0].get("values") or [] if isinstance(V, (int, float))]
            if Sl.get("pattern") == "cost_table":
                return [R.get("amount") for R in Sl.get("rows", []) if isinstance(R, dict)
                        and isinstance(R.get("amount"), (int, float))]
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.SlideSource")
            return None

    @staticmethod
    def Compute(Op, Values, Share):
        """The text a token stands for: Op over Values (sum/total, average, change, first, last, count), or the
        big number's share (Share = (part, whole)); None when there is nothing to compute it from."""
        if kS.ErrorMode:
            return None
        try:
            if Op == "share":
                return f"{round(100 * Share[0] / Share[1])}{NBSP}%" if Share and Share[1] else None
            Vals = [float(V) for V in Values or [] if isinstance(V, (int, float)) and not isinstance(V, bool)]
            if not Vals:
                return None
            Places = kFigures.Places(Vals)
            if Op in ("sum", "total"):
                return kFigures.Format(sum(Vals), Places)
            if Op == "average":
                Mean = sum(Vals) / len(Vals)  # one decimal more than the data when it does not come out even
                return kFigures.Format(Mean, Places if abs(Mean - round(Mean, Places)) < 1e-9 else Places + 1)
            if Op == "change":
                if len(Vals) < 2 or not Vals[0]:
                    return None
                Pct = round((Vals[-1] / Vals[0] - 1) * 100)
                return f"{'+' if Pct > 0 else (MINUS if Pct < 0 else '')}{abs(Pct)}{NBSP}%"
            if Op == "first":
                return kFigures.Format(Vals[0], Places)
            if Op == "last":
                return kFigures.Format(Vals[-1], Places)
            if Op == "count":
                return str(len(Vals))
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Compute")
            return None

    @staticmethod
    def Replace(Text, Source, Share, Facts, Where, Errors):
        """Text with every figure token replaced; a token with nothing to compute from is an error."""
        if kS.ErrorMode:
            return Text
        try:
            Out, Last = [], 0
            for Mt in TOKEN_RE.finditer(Text):
                Op, Name, Scalar = Mt.group(1), Mt.group(2), Mt.group(3)
                if Scalar is not None:
                    Val = Facts.get(Scalar)
                    if not isinstance(Val, (int, float)):
                        continue  # not a figure token: braces in ordinary text stay as written
                    Value = kFigures.Format(Val, kFigures.Places([Val]))
                elif Name is not None:
                    if not isinstance(Facts.get(Name), list):
                        Errors.append(f"{Where}: '{Mt.group(0)}' names no list in the deck's 'facts' "
                                      f"({', '.join(sorted(Facts)) or 'none given'})")
                        continue
                    Value = kFigures.Compute(Op, Facts[Name], None)
                else:
                    Value = kFigures.Compute(Op, Source, Share)
                if Value is None:
                    Errors.append(f"{Where}: '{Mt.group(0)}' has nothing to compute from - give the metric a "
                                  "'trend', the slide a 'series' (or cost rows), a big number that is a share for "
                                  "{share}, or name a deck fact ('{sum:revenue_q}')")
                    continue
                Out.append(Text[Last:Mt.start()] + Value)
                Last = Mt.end()
            return "".join(Out) + Text[Last:]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Replace")
            return Text

    @staticmethod
    def Fill(Obj, Source, Share, Facts, Where, Errors):
        """Obj with tokens replaced in every string under it."""
        if kS.ErrorMode:
            return Obj
        try:
            if isinstance(Obj, str):
                return kFigures.Replace(Obj, Source, Share, Facts, Where, Errors) if "{" in Obj else Obj
            if isinstance(Obj, list):
                return [kFigures.Fill(V, Source, Share, Facts, Where, Errors) for V in Obj]
            if isinstance(Obj, dict):
                return {K: kFigures.Fill(V, Source, Share, Facts, Where, Errors) for K, V in Obj.items()}
            return Obj
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Fill")
            return Obj

    @staticmethod
    def Resolve(Spec, ShareOf):
        """Replace the figure tokens in every slide of Spec (in place). ShareOf(slide) gives a big number's
        (part, whole). Returns the errors (a token with nothing to compute from)."""
        if kS.ErrorMode:
            return []
        try:
            Errors = []
            Facts = kFigures.Facts(Spec)
            for I, Sl in enumerate(Spec.get("slides", []), 1):
                if not isinstance(Sl, dict):
                    continue
                Pat = Sl.get("pattern")
                Source, Share = kFigures.SlideSource(Sl), ShareOf(Sl) if Pat == "big_number" else None
                for Key in list(Sl):
                    Where = f"slide {I} ({Pat})"
                    if Key == "metrics" and isinstance(Sl[Key], list):
                        Sl[Key] = [kFigures.Fill(Mt, (Mt.get("trend") if isinstance(Mt, dict) and Mt.get("trend")
                                                      else Source), Share, Facts, f"{Where} metrics[{J}]", Errors)
                                   for J, Mt in enumerate(Sl[Key])]
                    elif Key == "figure" and isinstance(Sl[Key], dict):
                        Sl[Key] = kFigures.Fill(Sl[Key], Sl[Key].get("trend") or Source, Share, Facts,
                                                f"{Where} figure", Errors)
                    else:
                        Sl[Key] = kFigures.Fill(Sl[Key], Source, Share, Facts, Where, Errors)
            return Errors
        except Exception as e:
            kS.GlobalErrorHandler(e, "kFigures.Resolve")
            return []
