"""Shared COM helpers for the scripts in this folder (Windows + PowerPoint only).

Opens a deck safely: if the deck is already open in PowerPoint it is reused (and left open);
otherwise it is opened read-only and windowless, then closed again. PowerPoint is quit
only if this script started it.

    with kPptSession(Path) as Pres:
        ...
"""
import io
import os
import sys

from kShared import kS, kToolException

MSO_GROUP = 6
MSO_PLACEHOLDER = 14
PP_PLACEHOLDER_TITLE, PP_PLACEHOLDER_BODY, PP_PLACEHOLDER_CENTER_TITLE = 1, 2, 3


class kPpt:
    """Stateless COM helpers."""

    @staticmethod
    def Utf8Stdout():
        """cp1252 stdout crashes on non-ASCII print() AFTER the COM work succeeded."""
        if kS.ErrorMode:
            return None
        try:
            if hasattr(sys.stdout, "reconfigure"):
                sys.stdout.reconfigure(encoding="utf-8")
            elif hasattr(sys.stdout, "buffer"):
                sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPpt.Utf8Stdout")
            return None

    @staticmethod
    def Norm(Path):
        """PowerPoint rejects paths that mix / and \\ (SaveCopyAs: "can't save ^0 to ^1")."""
        if kS.ErrorMode:
            return ""
        try:
            return os.path.normpath(os.path.abspath(Path))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPpt.Norm")
            return ""

    @staticmethod
    def ShapeTexts(Shapes):
        """Every shape with text, descending into groups."""
        if kS.ErrorMode:
            return []
        try:
            Result = []
            for Shape in Shapes:
                if Shape.Type == MSO_GROUP:
                    Result += kPpt.ShapeTexts(Shape.GroupItems)
                elif Shape.HasTextFrame and Shape.TextFrame.HasText:
                    Result.append(Shape)
            return Result
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPpt.ShapeTexts")
            return []

    @staticmethod
    def IsTitle(Shape):
        """True for a title or centre-title placeholder."""
        if kS.ErrorMode:
            return False
        try:
            return Shape.Type == MSO_PLACEHOLDER and \
                Shape.PlaceholderFormat.Type in (PP_PLACEHOLDER_TITLE, PP_PLACEHOLDER_CENTER_TITLE)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPpt.IsTitle")
            return False

    @staticmethod
    def NotesText(Slide):
        """The text of the slide's notes body placeholder, or "" when there is none."""
        if kS.ErrorMode:
            return ""
        try:
            if not Slide.HasNotesPage:
                return ""
            for Shape in Slide.NotesPage.Shapes:
                if Shape.Type == MSO_PLACEHOLDER and Shape.PlaceholderFormat.Type == PP_PLACEHOLDER_BODY:
                    return Shape.TextFrame.TextRange.Text
            return ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPpt.NotesText")
            return ""


class kPptSession:
    """Context manager: the presentation at Path, reused when already open, else opened and closed again."""

    def __init__(self, Path, ReadOnly=True):
        try:
            self._path = kPpt.Norm(Path)
            self._readOnly = ReadOnly
            self._app = None
            self._pres = None
            self._started = False
            self._openedHere = False
            self._comInitialized = False
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPptSession.__init__")

    def __enter__(self):
        """Attach to (or start) PowerPoint and return the presentation, or None when halted."""
        if kS.ErrorMode:
            return None
        try:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            self._comInitialized = True
            try:
                self._app = win32com.client.GetActiveObject("PowerPoint.Application")
            except pythoncom.com_error:  # ERROR-SUPPRESSED-JUSTIFIED: "not running" is answered by starting it
                self._app = win32com.client.Dispatch("PowerPoint.Application")
                self._started = True
            # Never trust ActivePresentation: match by full path.
            for Pres in self._app.Presentations:
                if self._pres is None and kPpt.Norm(Pres.FullName).lower() == self._path.lower():
                    self._pres = Pres
            self._openedHere = self._pres is None
            if self._openedHere:
                # ReadOnly, Untitled=False, WithWindow=False
                self._pres = self._app.Presentations.Open(self._path, -1 if self._readOnly else 0, 0, 0)
            return self._pres
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kPptSession.__enter__(file={self._path})")
            return None

    # DOCUMENTED EXCEPTION: cleanup must run even while halted; only __exit__ calls it, and it reports itself
    def ReleaseCom(self):
        """Close what this session opened, quit PowerPoint if it was started here, CoUninitialize."""
        try:
            if self._openedHere and self._pres is not None:
                self._pres.Close()
            self._pres = None
            if self._started and self._app is not None and self._app.Presentations.Count == 0:
                self._app.Quit()
            self._app = None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPptSession.ReleaseCom")
        finally:
            if self._comInitialized:
                import pythoncom
                pythoncom.CoUninitialize()
                self._comInitialized = False

    def __exit__(self, ExcType, ExcValue, Traceback):
        """Always release; never suppress an exception from the with-body."""
        if kS.ErrorMode:
            self.ReleaseCom()
            return False
        try:
            self.ReleaseCom()
            return False
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kPptSession.__exit__")
            return False
