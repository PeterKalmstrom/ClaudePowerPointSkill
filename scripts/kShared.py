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
  public repository); without one nothing is sent. NOTHING is ever sent without a person's yes: the report is
  shown and the person is asked "Do you want to send this error message? (yes/no)"; no means nothing is sent.
  Without a terminal (Claude ran the script) the report is saved and Claude must ask the user the same question;
  only on yes is it sent with scripts/send_error_report.py --yes. The script then exits 4 and prints
  "ERROR-REPORT-PENDING: <file>", so Claude knows. PowerPoint Live asks in its view and tells Claude in its
  error message. Errors that escape every guarded method (uncaught, other threads, asyncio) are captured too:
  kS.InstallUnhandledExceptionCapture().
  KPS_ERROR_REPORT=0 switches reporting off.

tools/check_kpattern.py audits every method for the four parts; the self-test and CI run it.
"""
import datetime
import json
import os
import platform
import sys
import tempfile
import threading
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
    HoldForView = False         # PowerPoint Live: reports wait for the person's Yes/No (view, or Claude asks)
    SavedPath = ""              # set when a report was saved for Claude to ask about (exit code 4)
    Question = "Do you want to send this error message?"
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
    def CanAsk():
        return bool(sys.stdin and sys.stdin.isatty() and sys.stderr.isatty())

    @staticmethod
    def AskConsent():
        """Ask the person at the terminal. True only for yes/y; anything else is no."""
        sys.stderr.write(f"{kErrorReport.Question} (yes/no) ")
        sys.stderr.flush()
        return (sys.stdin.readline() or "").strip().lower() in ("y", "yes")

    @staticmethod
    def SaveForLater():
        """Write the pending reports to a file, so Claude can ask the user and send them only on yes."""
        Path = os.path.join(tempfile.gettempdir(), f"pptskill-error-{kErrorReport.Pending[0].CorrelationId}.json")
        with open(Path, "w", encoding="utf-8") as File:
            json.dump({"question": kErrorReport.Question,
                       "summary": [Detail.Summary() for Detail in kErrorReport.Pending],
                       "payloads": [Detail.ToPayload() for Detail in kErrorReport.Pending]}, File, indent=2)
        return Path

    @staticmethod
    def SendAll(Payloads):
        """Send payloads (the person said yes). Returns the number sent."""
        Url = kErrorReport.EffectiveUrl()
        if not Url:
            kErrorReport.LastStatus = "no reporting address is configured (set KPS_ERROR_WEBHOOK)"
            return 0
        Sent = 0
        for Payload in Payloads:
            if kErrorReport.Send(Url, Payload):
                Sent += 1
        return Sent

    @staticmethod
    def PendingForView():
        """What the view shows with its Yes/No question."""
        return {"question": kErrorReport.Question, "configured": bool(kErrorReport.EffectiveUrl()),
                "reports": [{"ref": Detail.CorrelationId, "summary": Detail.Summary()}
                            for Detail in kErrorReport.Pending]}

    @staticmethod
    def Decide(SendIt):
        """The person answered the view's question: yes sends every pending report, no discards them."""
        Payloads = [Detail.ToPayload() for Detail in kErrorReport.Pending]
        Refs = [Detail.CorrelationId for Detail in kErrorReport.Pending]
        kErrorReport.Pending = []
        if not SendIt or os.environ.get("KPS_ERROR_REPORT") == "0":
            return {"sent": 0, "refs": Refs, "status": "not sent"}
        Sent = kErrorReport.SendAll(Payloads)
        return {"sent": Sent, "refs": Refs, "status": kErrorReport.LastStatus}

    @staticmethod
    def OfferOnce():
        """Show what would be sent, ask, send. Returns the number sent. Never raises."""
        Sent = 0
        try:
            if kErrorReport.Offered or not kErrorReport.Pending or kErrorReport.HoldForView:
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
            if not kErrorReport.CanAsk():
                Saved = kErrorReport.SaveForLater()
                kErrorReport.SavedPath = Saved
                Here = os.path.dirname(os.path.abspath(__file__))
                sys.stderr.write(
                    f"ERROR-REPORT-PENDING: {Saved}\n"
                    "Not sent: there is no one at a terminal to ask.\n"
                    f"CLAUDE: ask the user \"{kErrorReport.Question}\" (yes/no). Only if they answer yes, run:\n"
                    f"    python \"{os.path.join(Here, 'send_error_report.py')}\" \"{Saved}\" --yes\n"
                    "If they answer no, do not send it.\n")
                return 0
            if not kErrorReport.AskConsent():
                sys.stderr.write("Not sent.\n")
                return 0
            Sent = kErrorReport.SendAll([Detail.ToPayload() for Detail in kErrorReport.Pending])
            if Sent:
                sys.stderr.write(f"Error report sent. Quote reference {kErrorReport.Pending[0].CorrelationId} "
                                 "to support.\n")
            return Sent
        except Exception as e:  # ERROR-SUPPRESSED-JUSTIFIED: the offer failing must not hide the error it reports
            sys.stderr.write(f"the error report offer failed: {e!r}\n")
            return Sent

    @staticmethod
    def Send(Url, Payload):
        try:
            Body = json.dumps(Payload).encode("utf-8")
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
        kErrorReport.SavedPath = ""


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
        except Exception as Inner:  # ERROR-SUPPRESSED-JUSTIFIED: the handler must never raise or recurse
            sys.stderr.write(f"error handler failed: {Inner!r} while reporting {Location}\n")

    @staticmethod
    def Reset():
        """Re-arm the handler. Only for a fresh run, a person's Resume, or test setup - never in a catch."""
        kS.ErrorMode = False
        kS.Errors = []
        if not kErrorReport.HoldForView:  # undecided reports wait for the person's Yes/No even after Resume
            kErrorReport.Reset()

    @staticmethod
    def InstallUnhandledExceptionCapture():
        """Route errors that escape every guarded method - uncaught, in other threads, in asyncio - to the handler,
        so no error that can occur goes unreported. Idempotent."""
        try:
            if getattr(kS, "_captureInstalled", False):
                return
            kS._captureInstalled = True
            sys.excepthook = kS.OnUnhandled
            threading.excepthook = kS.OnUnhandledThread
            sys.unraisablehook = kS.OnUnraisable
        except Exception as Inner:  # ERROR-SUPPRESSED-JUSTIFIED: capture setup failing must not stop the program
            sys.stderr.write(f"could not install unhandled-error capture: {Inner!r}\n")

    @staticmethod
    def OnUnhandled(ExcType, Error, Tb):
        if issubclass(ExcType, (KeyboardInterrupt, SystemExit)):
            sys.__excepthook__(ExcType, Error, Tb)
            return
        kS.GlobalErrorHandler(Error, "[unhandled] main thread")
        kRun.Finish(1)

    @staticmethod
    def OnUnhandledThread(Args):
        if Args.exc_type is SystemExit:
            return
        kS.GlobalErrorHandler(Args.exc_value, f"[unhandled] thread {getattr(Args.thread, 'name', '?')}")

    @staticmethod
    def OnUnraisable(Unraisable):
        kS.GlobalErrorHandler(Unraisable.exc_value, f"[unhandled] {Unraisable.err_msg or 'unraisable'}")

    @staticmethod
    def OnAsyncioError(Loop, Context):
        """For loop.set_exception_handler: an exception in a task nobody awaited."""
        Error = Context.get("exception") or RuntimeError(Context.get("message", "asyncio error"))
        kS.GlobalErrorHandler(Error, "[unhandled] asyncio")

    @staticmethod
    def FirstError():
        """The report that halted the run, or None."""
        return kS.Errors[0] if kS.Errors else None


kErrorReport.ProductVersion = kErrorReport.Version()


class kRun:
    """The one way a script starts and exits. Infrastructure: exempt from the guard.

    Exit codes: 0 ok; 1 an unexpected error (reported; halted); 2 bad input (ToolInputException); 3 text that does
    not fit (build_deck); 4 an unexpected error whose report is waiting for the user's yes - stderr has a line
    "ERROR-REPORT-PENDING: <file>" and Claude must ask "Do you want to send this error message?"."""

    EXIT_REPORT_PENDING = 4

    @staticmethod
    def Main(AppClass):
        """Run AppClass().Run() (which returns an exit code) and exit: the code, 1 if the handler fired,
        or the code of an expected-state exception after printing its message."""
        kS.Reset()
        kErrorReport.Reset()
        kS.InstallUnhandledExceptionCapture()
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
        """Offer any reported error to the support flow, then exit: EXIT_REPORT_PENDING (4) when an error report
        was saved for Claude to ask the user about, 1 when the handler fired, otherwise Code."""
        sys.stdout.flush()
        kErrorReport.OfferOnce()
        sys.stderr.flush()
        if kErrorReport.SavedPath:
            sys.exit(kRun.EXIT_REPORT_PENDING)
        sys.exit(1 if kS.ErrorMode else Code)
