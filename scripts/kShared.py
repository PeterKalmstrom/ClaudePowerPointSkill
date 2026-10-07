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
* Reported errors can be sent to the support flow (Power Automate), the same payload the PowerShell kPSErrorReport
  sends. The address is KPS_ERROR_WEBHOOK, else the first line of scripts/kErrorWebhook.url (kept out of the
  public repository); without one nothing is sent. A person is asked first (y/N) when there is a terminal;
  KPS_ERROR_AUTOSEND=1 sends without asking (agreed in advance), KPS_ERROR_REPORT=0 switches it off.

tools/check_kpattern.py audits every method for the four parts; the self-test and CI run it.
"""
import datetime
import json
import os
import platform
import sys
import traceback
import urllib.error
import urllib.request
import uuid


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


class kErrorDetail:
    """One reported error, captured for the console and for the support flow. Infrastructure: never raises."""

    def __init__(self, Error, Location, AdditionalMessage=""):
        self.WhenIso = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        self.Location = Location
        self.CorrelationId = uuid.uuid4().hex[:12]
        self.Message = str(Error)
        self.ExceptionType = type(Error).__module__ + "." + type(Error).__name__
        self.Doing = AdditionalMessage
        self.ScriptName, self.LineNumber, self.LineText, self.StackTrace = "", 0, "", ""
        try:
            Frames = traceback.extract_tb(Error.__traceback__) if Error.__traceback__ else []
            if Frames:
                self.ScriptName, self.LineNumber, self.LineText = Frames[-1].filename, Frames[-1].lineno or 0, \
                    (Frames[-1].line or "").strip()
            self.StackTrace = "".join(traceback.format_exception(type(Error), Error, Error.__traceback__))
        except Exception:  # ERROR-SUPPRESSED-JUSTIFIED: a detail that cannot be captured is left empty, the report goes on
            pass

    def Summary(self):
        Lines = [f"when      : {self.WhenIso}   ref {self.CorrelationId}", f"where     : {self.Location}",
                 f"what      : {self.Message}", f"type      : {self.ExceptionType}"]
        if self.ScriptName:
            Lines.append(f"script    : {self.ScriptName} line {self.LineNumber}")
        if self.LineText:
            Lines.append(f"statement : {self.LineText}")
        if self.Doing:
            Lines.append(f"doing     : {self.Doing}")
        Lines.append(f"product   : {kErrorReport.ProductLabel()}")
        Lines.append(f"python    : {platform.python_version()}   {platform.platform()}")
        return Lines

    def ToPayload(self):
        """The fields the support flow expects - identical to kPSErrorDetail.ToPayload."""
        return {"ProductName": kErrorReport.ProductLabel(), "ProductVersion": kErrorReport.ProductVersion,
                "FunctionName": self.Location, "ErrorFileName": os.path.basename(self.ScriptName),
                "ErrorNumber": int(self.LineNumber or 0), "ErrorDescription": self.Message,
                "Comments": "\n".join(self.Summary() + ([f"full path : {self.ScriptName}"] if self.ScriptName else [])
                                      + (["Stack:", self.StackTrace] if self.StackTrace else [])),
                "ClientName": kErrorReport.Component, "ClientEmail": "",
                "OSVersion": f"Python {platform.python_version()} {platform.system()} {platform.release()}"}


class kErrorReport:
    """Offers reported errors to the support flow once per run (Power Automate HTTP trigger). Infrastructure."""

    Product = "building-powerpoint-decks"
    ProductVersion = ""
    Component = os.path.basename(sys.argv[0]) if sys.argv and sys.argv[0] else "python"
    TimeoutSeconds = 20
    SendImmediately = False     # long-running servers: offer each error as it happens, not at exit
    Pending = []
    Offered = False
    LastStatus = ""

    @staticmethod
    def ProductLabel():
        Label = f"{kErrorReport.Product} ({kErrorReport.Component})"
        return f"{Label} {kErrorReport.ProductVersion}".strip()

    @staticmethod
    def EffectiveUrl():
        """KPS_ERROR_WEBHOOK, else the first line of kErrorWebhook.url beside this file, else ""."""
        try:
            if os.environ.get("KPS_ERROR_WEBHOOK"):
                return os.environ["KPS_ERROR_WEBHOOK"].strip()
            Path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kErrorWebhook.url")
            if os.path.isfile(Path):
                with open(Path, encoding="utf-8") as File:
                    return (File.readline() or "").strip()
            return ""
        except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: no readable address means "not configured", said below
            kErrorReport.LastStatus = f"the reporting address could not be read: {e}"
            return ""

    @staticmethod
    def Version():
        """The plugin version, read once from .claude-plugin/plugin.json when it is there."""
        try:
            Path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".claude-plugin",
                                "plugin.json")
            if os.path.isfile(Path):
                with open(Path, encoding="utf-8") as File:
                    return json.load(File).get("version", "")
            return ""
        except Exception:  # ERROR-SUPPRESSED-JUSTIFIED: the version is a label on the report, not needed to send it
            return ""

    @staticmethod
    def AskConsent():
        if os.environ.get("KPS_ERROR_AUTOSEND") == "1":
            return True
        if not (sys.stdin and sys.stdin.isatty() and sys.stderr.isatty()):
            kErrorReport.LastStatus = "no one to ask (not run from a terminal); set KPS_ERROR_AUTOSEND=1 to send anyway"
            return False
        sys.stderr.write("Send this error report to the support team? [y/N] ")
        sys.stderr.flush()
        return (sys.stdin.readline() or "").strip().lower() in ("y", "yes")

    @staticmethod
    def OfferOnce():
        """Show what would be sent, ask, send. Returns the number sent. Never raises."""
        Sent = 0
        try:
            if kErrorReport.Offered or not kErrorReport.Pending:
                return 0
            kErrorReport.Offered = True
            if os.environ.get("KPS_ERROR_REPORT") == "0":
                return 0
            Url = kErrorReport.EffectiveUrl()
            if not Url:
                sys.stderr.write("Not offering to send this report: no reporting address is configured "
                                 "(set KPS_ERROR_WEBHOOK to enable it).\n")
                return 0
            sys.stderr.write("The error above can be sent to the support team. This would be sent:\n")
            for Detail in kErrorReport.Pending:
                sys.stderr.write("".join(f"    {Line}\n" for Line in Detail.Summary()) + "\n")
            if not kErrorReport.AskConsent():
                sys.stderr.write(f"Not sent: {kErrorReport.LastStatus or 'declined'}.\n")
                return 0
            for Detail in kErrorReport.Pending:
                if kErrorReport.Send(Url, Detail):
                    Sent += 1
            if Sent:
                sys.stderr.write(f"Error report sent. Quote reference {kErrorReport.Pending[0].CorrelationId} "
                                 "to support.\n")
            return Sent
        except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: the offer failing must not hide the error it reports
            sys.stderr.write(f"the error report offer failed: {e!r}\n")
            return Sent

    @staticmethod
    def Send(Url, Detail):
        try:
            Body = json.dumps(Detail.ToPayload()).encode("utf-8")
            Request = urllib.request.Request(Url, data=Body, method="POST",
                                             headers={"Content-Type": "application/json"})
            with urllib.request.urlopen(Request, timeout=kErrorReport.TimeoutSeconds) as Response:
                Status = Response.status
            kErrorReport.LastStatus = f"sent, HTTP {Status}"
            return 200 <= Status < 300
        except urllib.error.HTTPError as e:  # ERROR-SUPPRESSED-JUSTIFIED: said on stderr; the error itself still stands
            kErrorReport.LastStatus = f"the reporting address answered HTTP {e.code}"
        except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: said on stderr; the error itself still stands
            kErrorReport.LastStatus = f"could not reach the reporting address: {e}"
        sys.stderr.write(f"The report was not sent ({kErrorReport.LastStatus}). The error above still stands.\n")
        return False

    @staticmethod
    def Reset():
        kErrorReport.Pending = []
        kErrorReport.Offered = False


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
            kErrorReport.Pending.append(kErrorDetail(Error, Location, AdditionalMessage))
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
            if kErrorReport.SendImmediately:
                kErrorReport.OfferOnce()
                kErrorReport.Reset()
        except Exception as Inner:  # ERROR-SUPPRESSED-JUSTIFIED: the handler must never raise or recurse
            sys.stderr.write(f"error handler failed: {Inner!r} while reporting {Location}\n")

    @staticmethod
    def Reset():
        """Re-arm the handler. Only for a fresh run, a person's Resume, or test setup - never in a catch."""
        kS.ErrorMode = False
        kS.Errors = []
        kErrorReport.Reset()

    @staticmethod
    def FirstError():
        """The report that halted the run, or None."""
        return kS.Errors[0] if kS.Errors else None


kErrorReport.ProductVersion = kErrorReport.Version()


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
        """Offer any reported error to the support flow, then exit with 1 when the handler fired, otherwise Code."""
        sys.stdout.flush()
        kErrorReport.OfferOnce()
        sys.exit(1 if kS.ErrorMode else Code)
