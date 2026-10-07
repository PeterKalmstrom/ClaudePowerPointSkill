"""Shared COM helpers for the scripts in this folder (Windows + PowerPoint only).

Opens a deck safely: if the deck is already open in PowerPoint it is reused (and left open);
otherwise it is opened read-only and windowless, then closed again. PowerPoint is quit
only if this script started it.
"""
import contextlib
import io
import os
import sys


def utf8_stdout():
    """cp1252 stdout crashes on non-ASCII print() AFTER the COM work succeeded."""
    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")


def norm(path):
    """PowerPoint rejects paths that mix / and \\ (SaveCopyAs: "can't save ^0 to ^1")."""
    return os.path.normpath(os.path.abspath(path))


@contextlib.contextmanager
def open_deck(path, read_only=True):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    path = norm(path)
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
        started = False
    except Exception:
        app = win32com.client.Dispatch("PowerPoint.Application")
        started = True

    # Never trust ActivePresentation: match by full path, then by exact name.
    pres = next((p for p in app.Presentations if norm(p.FullName).lower() == path.lower()), None)
    opened_here = pres is None
    if opened_here:
        # ReadOnly, Untitled=False, WithWindow=False
        pres = app.Presentations.Open(path, -1 if read_only else 0, 0, 0)
    try:
        yield pres
    finally:
        if opened_here:
            pres.Close()
        if started and app.Presentations.Count == 0:
            app.Quit()
        pythoncom.CoUninitialize()


def shape_texts(shapes):
    """Yield every shape with text, descending into groups."""
    for sh in shapes:
        if sh.Type == 6:  # msoGroup
            yield from shape_texts(sh.GroupItems)
        elif sh.HasTextFrame and sh.TextFrame.HasText:
            yield sh


def is_title(sh):
    try:
        return sh.Type == 14 and sh.PlaceholderFormat.Type in (1, 3)  # title, centre title
    except Exception:
        return False


def notes_text(slide):
    try:
        for sh in slide.NotesPage.Shapes:
            if sh.Type == 14 and sh.PlaceholderFormat.Type == 2:  # ppPlaceholderBody
                return sh.TextFrame.TextRange.Text
    except Exception:
        pass
    return ""
