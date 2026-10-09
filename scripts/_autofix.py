"""Safe, deterministic fixes build_deck.py applies to a spec slide whose text does not fit (no dependencies).

Each pass, per slide with a problem, the first step that changes something:
  1 tidy    drop filler words ("very", "really", "in order to" -> "to", "that is,"...), write numbers' units
            short ("4 percent" -> "4 %", "3 million" -> "3 M"), and a tile value "19.8 MUSD" -> "19.8 M" when
            its label can carry the currency ("Revenue (USD)")
  2 move    cut the longest detail text of the shape that does not fit (a mitigation, a step's detail, a
            caption, points...) at a clause boundary and put the whole original into the speaker notes,
            marked "MOVED FROM SLIDE:"
  3 split   a bullets slide with "allow_split": true becomes two slides (the second titled "... (continued)")
Every change is returned as a line for an `auto:` report; build_deck.py --no-auto switches all of it off.
"""
import copy
import re

from kShared import kS

FILLERS = [  # (pattern, replacement); matched case-insensitively on word boundaries
    (r"\bin order to\b", "to"), (r"\bdue to the fact that\b", "because"), (r"\bat this point in time\b", "now"),
    (r"\bin the event that\b", "if"), (r"\bfor the purpose of\b", "for"), (r"\b(?:is|are) able to\b", "can"),
    (r"\bthat is,\s*", ""), (r"\ba total of\s+", ""), (r"\b(?:very|really|actually|basically|essentially|quite|"
                                                       r"simply|literally|truly)\s+", ""),
    (r"(\d)\s*(?:per ?cent|percent)\b", "\\1 %"), (r"(\d)\s*million\b", "\\1 M"), (r"(\d)\s*thousand\b", "\\1 k"),
]
SKIP_KEYS = {"id", "pattern", "notes", "image", "type", "reveal", "visual", "likelihood", "impact", "target",
             "number_format", "alt", "_moved", "categories", "series", "header", "rows", "value", "number", "unit",
             "date", "owner", "baseline", "from", "to", "attachment", "kicker", "eyebrow"}
MOVABLE = ["mitigation", "detail", "note", "caption", "support", "explain", "text", "points", "items", "body",
           "rule", "subtitle"]  # secondary text: its detail can go to the notes; never titles, values or actions
SHAPE_FIELDS = {"mitigation": "mitigation", "stepdetail": "detail", "points": "points", "point": "points",
                "statementpoint": "points", "reason": "points", "caption": "caption", "support": "support",
                "explain": "explain", "emailbody": "body", "email": "body", "qtext": "text", "optiontext": "text",
                "note": "note", "costnote": "note", "leadnote": "note", "chartnote": "caption",
                "metricsruletext": "rule", "subtitle": "subtitle", "bullets": "items", "body": "items"}
BREAKS = [r"(?<=[.!?])\s+", r";\s+", r"\s+[-–—]\s+", r":\s+", r",\s+",
          r"\s+(?=(?:with|before|after|by|using|so that|because|while|via|through|which|and)\b)"]
UNIT_VALUE = re.compile(r"^\s*([\d.,]+)\s*([kM])(USD|EUR|SEK|GBP)\s*$")
VALUE_CODES = {"unwanted_wrap"}


class kAutoFix:
    """Stateless spec fixes for text that does not fit; each returns what it changed, for the `auto:` lines."""

    @staticmethod
    def Tidy(Text):
        """Text without filler words, with short units; the first letter keeps its case."""
        if kS.ErrorMode:
            return Text
        try:
            New = str(Text)
            for Pattern, Repl in FILLERS:
                New = re.sub(Pattern, Repl, New, flags=re.I)
            New = re.sub(r"\s{2,}", " ", New).strip()
            if New and str(Text)[:1].isupper() and New[:1].islower():
                New = New[0].upper() + New[1:]
            return New
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.Tidy")
            return Text

    @staticmethod
    def TidyAll(Obj, Path, Changes):
        """Tidy every visible string under Obj (a slide, or part of one) in place; record each change."""
        if kS.ErrorMode:
            return Obj
        try:
            if isinstance(Obj, dict):
                for Key in list(Obj):
                    if Key not in SKIP_KEYS:
                        Obj[Key] = kAutoFix.TidyAll(Obj[Key], f"{Path}.{Key}" if Path else Key, Changes)
                return Obj
            if isinstance(Obj, list):
                return [kAutoFix.TidyAll(X, f"{Path}[{I}]", Changes) for I, X in enumerate(Obj)]
            if isinstance(Obj, str):
                New = kAutoFix.Tidy(Obj)
                if New != Obj:
                    Changes.append((Path, f"dropped filler: '{kAutoFix.Short(Obj)}' -> '{kAutoFix.Short(New)}'"))
                return New
            return Obj
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.TidyAll")
            return Obj

    @staticmethod
    def Short(Text, Max=60):
        """Text cut to Max characters for a report line."""
        if kS.ErrorMode:
            return ""
        try:
            Text = str(Text)
            return Text if len(Text) <= Max else Text[:Max - 1] + "…"
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.Short")
            return ""

    @staticmethod
    def CompactUnits(Sl, Changes):
        """A tile value '19.8 MUSD' -> '19.8 M' with the currency moved into its label ('Revenue (USD)'), only
        where the label has room (28 characters) and does not already name another unit."""
        if kS.ErrorMode:
            return
        try:
            Items = [(f"metrics[{I}]", Mt) for I, Mt in enumerate(Sl.get("metrics") or []) if isinstance(Mt, dict)]
            if isinstance(Sl.get("figure"), dict):
                Items.append(("figure", Sl["figure"]))
            for Path, Mt in Items:
                Hit = UNIT_VALUE.match(str(Mt.get("value", "")))
                if not Hit:
                    continue
                Label = str(Mt.get("label") or "")
                Cur = Hit.group(3)
                if Cur in Label:
                    NewLabel = Label
                elif "(" in Label or len(Label) + len(Cur) + 3 > 28:
                    continue  # the deck's unit context would be lost: leave it to the author
                else:
                    NewLabel = f"{Label} ({Cur})" if Label else Cur
                Old = Mt["value"]
                Mt["value"], Mt["label"] = f"{Hit.group(1)} {Hit.group(2)}", NewLabel
                Changes.append((f"{Path}.value", f"'{Old}' -> '{Mt['value']}', label '{Label}' -> '{NewLabel}'"))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.CompactUnits")
            return

    @staticmethod
    def Cut(Text):
        """(kept, True) when Text can be cut at a clause boundary to at most 75 % of its length, keeping at least
        three words; (Text, False) otherwise. Stronger boundaries (a sentence end) win over weaker ones."""
        if kS.ErrorMode:
            return Text, False
        try:
            Text = str(Text)
            for Pattern in BREAKS:
                Best = None
                for Hit in re.finditer(Pattern, Text):
                    Kept = Text[:Hit.start()].rstrip(" ,;:-–—")
                    if len(Kept) <= len(Text) * 0.75 and len(Kept.split()) >= 3:
                        Best = Kept  # the last boundary that is short enough keeps the most
                if Best:
                    return Best, True
            return Text, False
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.Cut")
            return Text, False

    @staticmethod
    def Places(Sl, Field):
        """Every (path, container, key) holding a string of Field on the slide: a top-level text, a list of
        texts, or the field of each object in a list (risks[].mitigation)."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for Key, Val in Sl.items():
                if Key in SKIP_KEYS:
                    continue
                if Key == Field and isinstance(Val, str):
                    Out.append((Key, Sl, Key))
                elif Key == Field and isinstance(Val, list):
                    Out += [(f"{Key}[{I}]", Val, I) for I, X in enumerate(Val) if isinstance(X, str)]
                elif isinstance(Val, list):
                    Out += [(f"{Key}[{I}].{Field}", X, Field) for I, X in enumerate(Val)
                            if isinstance(X, dict) and isinstance(X.get(Field), str)]
                elif isinstance(Val, dict) and isinstance(Val.get(Field), str):
                    Out.append((f"{Key}.{Field}", Val, Field))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.Places")
            return []

    @staticmethod
    def PlaceLength(Place):
        """Sort key: the length of the text a place holds."""
        if kS.ErrorMode:
            return 0
        try:
            return len(str(Place[1][Place[2]]))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.PlaceLength")
            return 0

    @staticmethod
    def MoveDetail(Sl, Field, Changes):
        """Cut the longest text of Field at a clause boundary, and its siblings nearly as long (cards in a row share
        one size); each whole original goes to the notes. True when something moved."""
        if kS.ErrorMode:
            return False
        try:
            Places = sorted(kAutoFix.Places(Sl, Field), key=kAutoFix.PlaceLength, reverse=True)
            Longest = kAutoFix.PlaceLength(Places[0]) if Places else 0
            Moved = False
            for Path, Box, Key in Places:
                Old = str(Box[Key])
                if Moved and len(Old) < Longest * 0.8:
                    break
                Kept, Ok = kAutoFix.Cut(Old)
                if not Ok:
                    continue
                Box[Key] = Kept
                Sl.setdefault("_moved", []).append(f"MOVED FROM SLIDE ({Path}): {Old}")
                Changes.append((Path, f"cut to '{kAutoFix.Short(Kept)}'; the full text is in the notes "
                                      "(MOVED FROM SLIDE:)"))
                Moved = True
            return Moved
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.MoveDetail")
            return False

    @staticmethod
    def FieldOf(Shape):
        """The spec field a builder shape name stands for ('Mitigation2' -> 'mitigation'), or ''."""
        if kS.ErrorMode:
            return ""
        try:
            return SHAPE_FIELDS.get(re.sub(r"\d+$", "", str(Shape or "")).lower(), "")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.FieldOf")
            return ""

    @staticmethod
    def Split(Spec, Index, Changes):
        """Split the bullets slide at Index (0-based) into two when it says "allow_split": true. True when split."""
        if kS.ErrorMode:
            return False
        try:
            Sl = Spec["slides"][Index]
            Items = Sl.get("items") or []
            if Sl.get("pattern") != "bullets" or not Sl.get("allow_split") or len(Items) < 2:
                return False
            Half = (len(Items) + 1) // 2
            Second = copy.deepcopy(Sl)
            Sl["items"], Second["items"] = Items[:Half], Items[Half:]
            Sl["allow_split"] = Second["allow_split"] = False
            Second["id"] = f"{Sl.get('id', f's{Index + 1:02d}')}-2"
            Second["title"] = f"{str(Sl.get('title', '')).replace(' (continued)', '')} (continued)"
            Second["notes"] = "Continued from the previous slide; its notes carry the detail."
            Second.pop("_moved", None)
            Spec["slides"].insert(Index + 1, Second)
            Changes.append(("items", f"split into two slides ({Half} + {len(Items) - Half} bullets); the second is "
                                     f"'{Second['id']}'"))
            return True
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.Split")
            return False

    @staticmethod
    def PreSplit(Spec, Max):
        """Before the build: every bullets slide with "allow_split": true and more than Max items is split until
        each part holds at most Max. Returns the changes, as Apply does."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for Index in range(len(Spec.get("slides", [])) - 1, -1, -1):
                Sl = Spec["slides"][Index]
                if Sl.get("pattern") != "bullets" or not Sl.get("allow_split") or len(Sl.get("items") or []) <= Max:
                    continue
                Changes = []
                Parts = [Index]
                while Parts:
                    At = Parts.pop()
                    if len(Spec["slides"][At].get("items") or []) > Max:
                        Spec["slides"][At]["allow_split"] = True
                        if kAutoFix.Split(Spec, At, Changes):
                            Parts += [At, At + 1]
                Out += [(Index + 1, Sl.get("id", f"s{Index + 1:02d}"), Path, What) for Path, What in Changes]
            return sorted(Out, key=kAutoFix.ChangeKey)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.PreSplit")
            return []

    @staticmethod
    def FixSlide(Spec, Index, Issues):
        """The first fix that changes something on spec slide Index (0-based): tidy, compact units, move detail,
        split. Returns [(path, what changed)]."""
        if kS.ErrorMode:
            return []
        try:
            Sl = Spec["slides"][Index]
            Changes = []
            kAutoFix.TidyAll(Sl, "", Changes)
            kAutoFix.CompactUnits(Sl, Changes)
            if Changes:
                return Changes
            if all(I.get("code") in VALUE_CODES or "on one line" in I.get("message", "") for I in Issues):
                return []  # a value that wraps: only its unit can change, and that was tried
            Fields = []
            for I in Issues:
                Fd = kAutoFix.FieldOf(I.get("shape"))
                if Fd and Fd not in Fields:
                    Fields.append(Fd)
            Fields += [Fd for Fd in MOVABLE if Fd not in Fields]
            for Fd in Fields:
                if kAutoFix.MoveDetail(Sl, Fd, Changes):
                    return Changes
            kAutoFix.Split(Spec, Index, Changes)
            return Changes
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.FixSlide")
            return []

    @staticmethod
    def Apply(Spec, Issues, AllowSplit=True):
        """One pass over every spec slide with an issue (from the last slide up, so a split keeps the earlier
        indexes). Returns [(spec slide number, slide id, path, what changed)] - empty when nothing could change."""
        if kS.ErrorMode:
            return []
        try:
            BySlide = {}
            for I in Issues:
                if I.get("spec"):
                    BySlide.setdefault(I["spec"], []).append(I)
            Out = []
            for No in sorted(BySlide, reverse=True):
                Sl = Spec["slides"][No - 1]
                if not AllowSplit:
                    Sl["allow_split"] = False
                for Path, What in kAutoFix.FixSlide(Spec, No - 1, BySlide[No]):
                    Out.append((No, Sl.get("id", f"s{No:02d}"), Path, What))
            return sorted(Out, key=kAutoFix.ChangeKey)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.Apply")
            return []

    @staticmethod
    def ChangeKey(Change):
        """Sort key of a change: its slide number."""
        if kS.ErrorMode:
            return 0
        try:
            return Change[0]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAutoFix.ChangeKey")
            return 0
