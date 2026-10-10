"""Batch mode for build_deck.py: several specs in one call.

    build_deck.py a.json b.json c.json --check --apply [--out DIR]

Each spec runs as its own build_deck.py process (same flags, its deck beside the spec or in DIR), its full output is
printed under a `==== spec: ... ====` header, and the call ends with one summary table, one combined contact sheet
(one row per deck) and a BATCH-JSON line. The exit code is the worst of the decks' codes (4 > 1 > 3 > 2 > 0).

Why: an unattended maker building five decks spent one call per deck per step (plan, check, look) - the batch
turns that into one check call and one look for all of them (benchmark round 10, evals/benchmark/README.md).
"""
import json
import os
import subprocess
import sys

from kShared import kS, kToolException

HERE = os.path.dirname(os.path.abspath(__file__))
WORST = {4: 5, 1: 4, 3: 3, 2: 2, 0: 0}  # exit-code severity: error report > crash > unfit text > spec error > ok
ROW_SLIDE_WIDTH = 300  # px per slide on the combined sheet
PARALLEL = 3  # decks built side by side


class kBatchCheck:
    """Plan, check and apply several specs in one call, with one summary."""

    @staticmethod
    def OutFor(Spec, OutDir):
        """The deck path for one spec: <OutDir>/<spec name>.pptx, or beside the spec when no folder is given."""
        if kS.ErrorMode:
            return ""
        try:
            Name = os.path.splitext(os.path.basename(Spec))[0] + ".pptx"
            return os.path.join(OutDir, Name) if OutDir else os.path.join(os.path.dirname(os.path.abspath(Spec)),
                                                                           Name)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBatchCheck.OutFor")
            return ""

    @staticmethod
    def Command(Args, Spec, Out):
        """The single-spec build_deck.py command line that carries the batch's flags."""
        if kS.ErrorMode:
            return []
        try:
            Cmd = [sys.executable, os.path.join(HERE, "build_deck.py"), Spec]
            if Out:
                Cmd += ["--out", Out]
            for Flag, On in (("--check", Args.check), ("--plan", Args.plan), ("--apply", Args.apply),
                             ("--no-auto", Args.no_auto), ("--force", Args.force), ("--lint", Args.lint)):
                if On:
                    Cmd.append(Flag)
            if Args.template:
                Cmd += ["--template", Args.template]
            return Cmd
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBatchCheck.Command")
            return []

    @staticmethod
    def ParseJson(Text):
        """The CHECK-JSON payload in one deck's output, or {} when there is none (plan only, spec errors)."""
        if kS.ErrorMode:
            return {}
        try:
            for Line in reversed(Text.splitlines()):
                if Line.startswith("CHECK-JSON "):
                    return json.loads(Line[len("CHECK-JSON "):])
            return {}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBatchCheck.ParseJson")
            return {}

    @staticmethod
    def Start(Args, Spec, Out):
        """Start one spec's build_deck.py process (each LibreOffice render has its own profile, so decks can run
        side by side)."""
        if kS.ErrorMode:
            return None
        try:
            Env = dict(os.environ, PYTHONIOENCODING="utf-8")
            return subprocess.Popen(kBatchCheck.Command(Args, Spec, Out), stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace", env=Env)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBatchCheck.Start")
            return None

    @staticmethod
    def Finish(Args, Spec, Out, Proc):
        """Wait for one spec, print its output (without the CHECK-JSON line; BATCH-JSON carries it) under a
        header, return its result row."""
        if kS.ErrorMode:
            return {"spec": Spec, "code": 1}
        try:
            if Proc is None:
                return {"spec": Spec, "code": 1}
            Stdout, Stderr = Proc.communicate()
            Text = (Stdout or "") + (Stderr or "")
            Shown = [L for L in Text.rstrip().splitlines() if not L.startswith("CHECK-JSON ")]
            print(f"\n==================== spec: {Spec} ====================")
            print("\n".join(Shown), flush=True)
            return {"spec": Spec, "deck": Out if Args.check else "", "code": Proc.returncode,
                    "json": kBatchCheck.ParseJson(Stdout or ""), "questions": Text.count("question:")}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBatchCheck.Finish")
            return {"spec": Spec, "code": 1}

    @staticmethod
    def SlideFiles(Sheet):
        """The slide renders beside one deck's contact sheet, in slide order."""
        if kS.ErrorMode:
            return []
        try:
            if not Sheet:
                return []
            from contact_sheet import kContactSheet
            return kContactSheet.FromFolder(os.path.dirname(Sheet))
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBatchCheck.SlideFiles")
            return []

    @staticmethod
    def AllSheet(Results, Out):
        """One sheet for every deck: one row per deck, its slides left to right, the deck name above. Returns the
        path, or "" when nothing was rendered."""
        if kS.ErrorMode:
            return ""
        try:
            from PIL import Image, ImageDraw
            Rows = [(os.path.basename(R["deck"]), kBatchCheck.SlideFiles(R["json"].get("sheet", "")))
                    for R in Results if R.get("json")]
            Rows = [R for R in Rows if R[1]]
            if not Rows:
                return ""
            First = Image.open(Rows[0][1][0])
            Wd, Pad, Cap = ROW_SLIDE_WIDTH, 8, 18
            Ht = round(Wd * First.height / First.width)
            Cols = max(len(Files) for _, Files in Rows)
            Sheet = Image.new("RGB", (Cols * (Wd + Pad) + Pad, len(Rows) * (Ht + Cap + Pad) + Pad), "white")
            Draw = ImageDraw.Draw(Sheet)
            for RowNo, (Name, Files) in enumerate(Rows):
                Y = Pad + RowNo * (Ht + Cap + Pad)
                Draw.text((Pad, Y), Name, fill="black")
                for ColNo, File in enumerate(Files):
                    Sheet.paste(Image.open(File).convert("RGB").resize((Wd, Ht)), (Pad + ColNo * (Wd + Pad), Y + Cap))
            Sheet.save(Out)
            return Out
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBatchCheck.AllSheet")
            return ""

    @staticmethod
    def Worst(Codes):
        """The most severe exit code."""
        if kS.ErrorMode:
            return 1
        try:
            return max(Codes, key=kBatchCheck.Severity) if Codes else 0
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBatchCheck.Worst")
            return 1

    @staticmethod
    def Severity(Code):
        """Sort key for an exit code (unknown codes rank as a crash)."""
        if kS.ErrorMode:
            return 4
        try:
            return WORST.get(Code, 4)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBatchCheck.Severity")
            return 4

    @staticmethod
    def Summary(Results, AllSheet):
        """One line per deck, the combined sheet, and the BATCH-JSON line."""
        if kS.ErrorMode:
            return
        try:
            print("\n==================== batch summary ====================")
            for R in Results:
                Counts = (R.get("json") or {}).get("counts") or {}
                Slides = (R.get("json") or {}).get("slides", "-")
                State = "clean" if R["code"] == 0 and not Counts.get("error") and not R.get("questions") else "FIX"
                print(f"  {State:<5} exit {R['code']}  {Slides!s:>2} slides  {Counts.get('error', 0)} err  "
                      f"{Counts.get('warn', 0)} warn  {R.get('questions', 0)} question(s)  {R['spec']}")
            if AllSheet:
                print(f"all-sheet: {AllSheet}  <- one look at every deck (one row each); per-deck sheets above")
            Rows = [{"spec": R["spec"], "deck": R.get("deck", ""), "code": R["code"],
                     "counts": (R.get("json") or {}).get("counts", {}), "questions": R.get("questions", 0),
                     "sheet": (R.get("json") or {}).get("sheet", "")} for R in Results]
            print("BATCH-JSON " + json.dumps({"decks": Rows, "all_sheet": AllSheet}, ensure_ascii=False))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBatchCheck.Summary")
            return

    @staticmethod
    def Run(Args):
        """Every spec in Args.spec (a list), then the summary. Returns the worst exit code."""
        if kS.ErrorMode:
            return 1
        try:
            OutDir = Args.out if Args.out else ""
            if OutDir:
                os.makedirs(OutDir, exist_ok=True)
            Results = []
            Jobs = [(Spec, kBatchCheck.OutFor(Spec, OutDir) if Args.check or not Args.plan else "")
                    for Spec in Args.spec]
            for First in range(0, len(Jobs), PARALLEL):
                Group = Jobs[First:First + PARALLEL]
                Procs = [kBatchCheck.Start(Args, Spec, Out) for Spec, Out in Group]
                for (Spec, Out), Proc in zip(Group, Procs):
                    Results.append(kBatchCheck.Finish(Args, Spec, Out, Proc))
            Sheet = ""
            if Args.check:
                Folder = OutDir or os.path.dirname(os.path.abspath(Args.spec[0]))
                Sheet = kBatchCheck.AllSheet(Results, os.path.join(Folder, "contact-all.png"))
            kBatchCheck.Summary(Results, Sheet)
            return kBatchCheck.Worst([R["code"] for R in Results])
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBatchCheck.Run")
            return 1
