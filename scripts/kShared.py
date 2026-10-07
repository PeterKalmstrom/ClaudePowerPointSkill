"""Shared error handling for every Python file in this skill (any OS, standard library only).

The rules, the same as the C#, TypeScript and PowerShell code this skill's author writes:

* An error is never hidden. It is either HANDLED (corrected, so the operation still succeeds) or REPORTED to
  kS.GlobalErrorHandler. A catch that only logs, returns an empty result or swallows is not allowed.
* The first unexpected error halts the run: the handler sets kS.ErrorMode and every guarded method returns its
  safe default from then on. Only a person (a Resume button) or a fresh run re-arms it - kS.Reset().
* Every method has the same four parts:

      def Load(self, Path):
          if kS.ErrorMode:                                    # 1 guard - first statement, outside the try
              return None                                     #   safe default
          try:
              ...                                             # 2 the whole body
              return Result
          except kToolException:
              raise                                           #   expected states pass through to the entry point
          except Exception as e:
              kS.GlobalErrorHandler(e, "kDeckReader.Load")    # 3 report, location = "Class.Method"
              return None                                     # 4 safe default

* EXPECTED states (file not found, a spec mistake, no deck open) are found by explicit precondition checks that
  raise ToolInputException (exit 2) or ToolReportableException (exit 1 or a given code) - never by catching.
* A deliberate non-reporting catch carries an inline "ERROR-SUPPRESSED-JUSTIFIED: <why>" comment.

tools/check_kpattern.py audits every method for the four parts; the self-test and CI run it.
"""
import datetime
import os
import sys
import traceback


class kToolException(Exception):
    """Base of the expected-state exceptions. They carry their exit code and a message for a person."""

    def __init__(self, Message, ExitCode=1):
        super().__init__(Message)
        self.ExitCode = ExitCode


class ToolInputException(kToolException):
    """The caller's input is wrong (a bad argument, a spec mistake, a missing file). Exit code 2."""

    def __init__(self, Message, ExitCode=2):
        super().__init__(Message, ExitCode)


class ToolReportableException(kToolException):
    """An expected state of the world stops the tool (no deck open, text that does not fit). Exit code 1 or given."""

    def __init__(self, Message, ExitCode=1):
        super().__init__(Message, ExitCode)


class kS:
    """The global error handler and the halt latch. Infrastructure: exempt from the guard, never raises."""

    ErrorMode = False
    Errors = []          # every report, in order; the first one is the cause
    LogPath = os.environ.get("PPTSKILL_ERROR_LOG", "")

    @staticmethod
    def GlobalErrorHandler(Error, Location, AdditionalMessage=""):
        """Report an unexpected error once, loudly, then halt every guarded method. Never raises."""
        try:
            Halted = kS.ErrorMode
            Record = {"time": datetime.datetime.now().isoformat(timespec="seconds"), "location": Location,
                      "type": type(Error).__name__, "message": str(Error), "while_halted": Halted,
                      "detail": AdditionalMessage}
            kS.Errors.append(Record)
            kS.ErrorMode = True
            if Halted:
                sys.stderr.write(f"[while halted] ERROR in {Location}: {type(Error).__name__}: {Error}\n")
            else:
                Stack = "".join(traceback.format_exception(type(Error), Error, Error.__traceback__))
                sys.stderr.write(f"\nERROR in {Location}: {type(Error).__name__}: {Error}\n"
                                 + (f"{AdditionalMessage}\n" if AdditionalMessage else "") + Stack
                                 + "Halted: later steps are skipped. Fix the cause and run again.\n")
            if kS.LogPath:
                with open(kS.LogPath, "a", encoding="utf-8") as Log:
                    Log.write(repr(Record) + "\n")
        except Exception as Inner:  # ERROR-SUPPRESSED-JUSTIFIED: the handler must never raise or recurse
            sys.stderr.write(f"error handler failed: {Inner!r} while reporting {Location}\n")

    @staticmethod
    def Reset():
        """Re-arm the handler. Only for a fresh run, a person's Resume, or test setup - never in a catch."""
        kS.ErrorMode = False
        kS.Errors = []

    @staticmethod
    def FirstError():
        """The report that halted the run, or None."""
        return kS.Errors[0] if kS.Errors else None


class kRun:
    """The one way a script starts and exits. Infrastructure: exempt from the guard."""

    @staticmethod
    def Main(AppClass):
        """Run AppClass().Run() (which returns an exit code) and exit: the code, 1 if the handler fired,
        or the code of an expected-state exception after printing its message."""
        kS.Reset()
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
        Code = 0
        try:
            Code = AppClass().Run() or 0
        except kToolException as e:
            sys.stderr.write(f"{e}\n")
            Code = e.ExitCode
        except Exception as e:
            kS.GlobalErrorHandler(e, "kRun.Main")
        kRun.Finish(Code)

    @staticmethod
    def Finish(Code):
        """Exit with 1 when the handler fired, otherwise with Code."""
        sys.stdout.flush()
        sys.exit(1 if kS.ErrorMode else Code)
