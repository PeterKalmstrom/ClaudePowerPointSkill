"""Improve a deck in place and keep its design - any OS, python-pptx. For decks that must keep their look; to
restyle one in the skill's design instead, use extract_spec.py and rebuild.

    python scripts/improve_deck.py deck.pptx --out improved.pptx
    python scripts/improve_deck.py deck.pptx --in-place            # keeps deck.pptx.bak first
    python scripts/improve_deck.py deck.pptx --dry-run             # list what it would do
    python scripts/improve_deck.py deck.pptx --out improved.pptx --no-fix --report findings.txt

Three steps, each listed as it is made:
  1 fix     fix_deck.py's safe fixes (placeholder, floor, fit, aspect, alt, palette, legend, numfmt)
  2 notes   a slide without notes gets a DRAFT spoken script written from its own title, text, chart and
            table figures; notes without a script (under 12 words of prose: 'Churn 1.8 to 1.2.') get the
            draft in front and the original kept below the divider. Every draft is marked for review and says
            nothing the slide does not say - check it and add the figures' sources before presenting.
  3 report  lint_deck.py on the result: what still needs a person (titles, wording, crowding, alt text).
Nothing else is moved, restyled or reworded.
"""
import argparse
import os
import shutil
import subprocess
import sys

from pptx import Presentation

from kShared import ToolInputException, ToolReportableException, kRun, kS, kToolException
from fix_deck import FIXES, kFixDeck
from _deck_content import DRAFT_MARK, kNotesDraft, kSlideContent
from _rules import NOTES_DIVIDER, SCRIPT_MIN_WORDS, kRules

HERE = os.path.dirname(os.path.abspath(__file__))


class kNotesRepair:
    """Stateless: write or repair speaker notes from each slide's own content."""

    @staticmethod
    def Draft(Info):
        """A draft script (one string) of at least SCRIPT_MIN_WORDS words, or '' when the slide has too little
        text to say anything without inventing it."""
        if kS.ErrorMode:
            return ""
        try:
            Script = " ".join(kNotesDraft.Script(Info))
            return Script if kRules.ScriptWords(Script) >= SCRIPT_MIN_WORDS or (
                Script and Info["number"] == 1) else ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kNotesRepair.Draft")
            return ""

    @staticmethod
    def NewNotes(Old, Draft):
        """The notes text: the draft, the divider, the review mark, then the original notes (kept in full)."""
        if kS.ErrorMode:
            return ""
        try:
            Parts = [Draft, "", NOTES_DIVIDER, DRAFT_MARK]
            if Old.strip():
                Parts += ["Original notes:", Old.strip()]
            return "\n".join(Parts)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kNotesRepair.NewNotes")
            return ""

    @staticmethod
    def Apply(Prs, Log, DryRun=False):
        """Draft notes where they are missing or have no script; Log(message) for each. Returns the slide numbers
        that got a draft."""
        if kS.ErrorMode:
            return []
        try:
            Drafted = []
            for Info, Slide in zip(kSlideContent.ReadDeck(Prs), Prs.slides):
                if Info["hidden"]:
                    continue
                Old = Info["notes_raw"]
                Missing = not Old.strip()
                if not Missing and kRules.ScriptWords(Old) >= SCRIPT_MIN_WORDS:
                    continue
                Draft = kNotesRepair.Draft(Info)
                N = Info["number"]
                if not Draft:
                    Log(f"notes        slide {N}: too little text on the slide to draft a script - write it by hand")
                    continue
                if not DryRun:
                    Slide.notes_slide.notes_text_frame.text = kNotesRepair.NewNotes(Old, Draft)
                Drafted.append(N)
                Log(f"notes        slide {N}: {'DRAFT script written' if Missing else 'DRAFT script added above the old notes'}"
                    f" ({len(Draft.split())} words) - review it")
            return Drafted
        except Exception as e:
            kS.GlobalErrorHandler(e, "kNotesRepair.Apply")
            return []


class kImproveDeckApp:
    """Command line: fix, draft notes, save, lint."""

    def __init__(self):
        try:
            self._made = []
        except Exception as e:
            kS.GlobalErrorHandler(e, "kImproveDeckApp.__init__")

    def Log(self, Message):
        """Record and print one change."""
        if kS.ErrorMode:
            return None
        try:
            self._made.append(Message)
            print(Message)
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kImproveDeckApp.Log")
            return None

    @staticmethod
    def Arguments():
        """Parsed command line."""
        if kS.ErrorMode:
            return None
        try:
            Ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
            Ap.add_argument("file")
            Ap.add_argument("--out")
            Ap.add_argument("--in-place", action="store_true", help="improve the input; the original is kept as .bak")
            Ap.add_argument("--dry-run", action="store_true", help="list the changes without writing")
            Ap.add_argument("--no-fix", action="store_true", help="skip fix_deck.py's safe fixes")
            Ap.add_argument("--no-notes", action="store_true", help="leave the speaker notes alone")
            Ap.add_argument("--report", help="also write the lint findings to this text file")
            return Ap.parse_args()
        except Exception as e:
            kS.GlobalErrorHandler(e, "kImproveDeckApp.Arguments")
            return None

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            A = self.Arguments()
            if A is None:
                return 1
            if sum(1 for X in (A.out, A.in_place, A.dry_run) if X) != 1:
                raise ToolInputException("give exactly one of --out improved.pptx, --in-place or --dry-run")
            if not os.path.isfile(A.file):
                raise ToolReportableException(f"file not found: {A.file}")
            if A.out and os.path.abspath(A.out) == os.path.abspath(A.file):
                raise ToolInputException("--out must differ from the input (use --in-place to keep a .bak)")
            Prs = Presentation(A.file)
            if not A.no_fix:
                kFixDeck.Fix(Prs, set(FIXES), self.Log)
            Drafted = [] if A.no_notes else kNotesRepair.Apply(Prs, self.Log, A.dry_run)
            if kS.ErrorMode:
                return 1
            print(f"\n{len(self._made)} change(s){' (dry run, nothing written)' if A.dry_run else ''}"
                  + (f"; DRAFT notes to review on slide(s) {', '.join(str(N) for N in Drafted)}" if Drafted else ""))
            if A.dry_run:
                return 0
            return self.SaveAndLint(Prs, A)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kImproveDeckApp.Run")
            return 1

    def SaveAndLint(self, Prs, A):
        """Save (the original kept as .bak for --in-place), then lint the result; returns lint's exit code."""
        if kS.ErrorMode:
            return 1
        try:
            Out = A.out
            if A.in_place:
                shutil.copy2(A.file, A.file + ".bak")
                print(f"original kept as {A.file}.bak")
                Out = A.file
            Prs.save(Out)
            print(f"-> {Out}\n\nWhat's left for a person (lint_deck.py):")
            sys.stdout.flush()
            R = subprocess.run([sys.executable, os.path.join(HERE, "lint_deck.py"), Out], capture_output=True,
                               text=True, encoding="utf-8")
            print(R.stdout + R.stderr)
            if A.report:
                with open(A.report, "w", encoding="utf-8") as File:
                    File.write("\n".join(self._made) + "\n\n" + R.stdout)
                print(f"report: {A.report}")
            return R.returncode
        except Exception as e:
            kS.GlobalErrorHandler(e, "kImproveDeckApp.SaveAndLint")
            return 1


if __name__ == "__main__":
    kRun.Main(kImproveDeckApp)
