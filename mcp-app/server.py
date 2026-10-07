"""PowerPoint Live - an MCP App that drives desktop PowerPoint (Windows, COM) and shows the current slide live.

    uvx --with "mcp<2" --with "pywin32; sys_platform == 'win32'" python mcp-app/server.py

Register with Claude Code (or any MCP client; hosts that support MCP Apps - e.g. Claude Desktop and claude.ai -
also show the live slide view):

    claude mcp add --scope user powerpoint-live -- uvx --with "mcp<2" --with "pywin32; sys_platform == 'win32'" python <skill>/mcp-app/server.py

Tools for the model:
  powerpoint_open   open a deck (or attach to an open one by name) and show a slide
  powerpoint_run    run Python against the deck: `app`, `prs`, `slide` (current), `goto(n)`, `win32com` in scope
  powerpoint_show   make slide n the current slide
  powerpoint_move   move a slide (before another slide, or to the end of a section)
Tools only the view calls: slide_state (cheap change signature), slide_image (PNG of a slide), deck_outline
(sections and slides for the slide sorter) and slide_thumbs (small PNGs, cached by content).
The view polls slide_state about once a second and re-renders when the slide changes - whether Claude, a
script or a person changed it.
"""
import asyncio
import base64
import contextlib
import io
import os
import sys
import tempfile
import threading
import traceback
from concurrent.futures import ThreadPoolExecutor

from mcp.server.fastmcp import FastMCP

HERE = os.path.dirname(os.path.abspath(__file__))
URI = "ui://powerpoint-live/slide-view"
MIME = "text/html;profile=mcp-app"
UI = {"ui": {"resourceUri": URI}, "ui/resourceUri": URI}
APP_ONLY = {"ui": {"resourceUri": URI, "visibility": ["app"]}, "ui/resourceUri": URI}

mcp = FastMCP("powerpoint-live")


# ---------------------------------------------------------------- one COM thread

def _com_init():
    import pythoncom
    pythoncom.CoInitialize()


_pool = ThreadPoolExecutor(max_workers=1, initializer=_com_init if sys.platform == "win32" else None)


async def on_com(fn, *args):
    """Run fn on the single COM thread (COM objects must stay on the thread that created them)."""
    return await asyncio.get_running_loop().run_in_executor(_pool, fn, *args)


class State:
    prs_name = None   # FullName of the deck we work on
    current = 1
    version = 0


S = State()
_lock = threading.Lock()


def _app():
    if sys.platform != "win32":
        raise RuntimeError("PowerPoint Live needs Windows with desktop PowerPoint")
    import win32com.client
    return win32com.client.Dispatch("PowerPoint.Application")


def _prs():
    app = _app()
    if S.prs_name:
        for p in app.Presentations:
            if p.FullName == S.prs_name:
                return p
        raise RuntimeError(f"{S.prs_name} is no longer open - call powerpoint_open")
    if app.Presentations.Count == 0:
        raise RuntimeError("no deck is open - call powerpoint_open with a path")
    p = app.ActivePresentation
    S.prs_name = p.FullName
    return p


def _clamp(prs, n):
    return max(1, min(int(n), prs.Slides.Count)) if prs.Slides.Count else 0


def _sync_window(prs):
    """Let the PowerPoint window follow the current slide, when there is one."""
    try:
        if prs.Windows.Count and S.current:
            prs.Windows(1).View.GotoSlide(S.current)
    except Exception:  # a window in slide sorter / reading view cannot GotoSlide; the preview still updates
        pass


def _summary(prs):
    return {"deck": os.path.basename(prs.FullName), "path": prs.FullName, "slide": S.current,
            "count": prs.Slides.Count, "version": S.version}


def _signature(prs, n):
    """Cheap fingerprint of slide n: shapes, positions, text lengths. Changes when anyone edits the slide."""
    if not n:
        return "empty"
    sl = prs.Slides(n)
    parts = [str(sl.SlideID), str(sl.Shapes.Count)]
    for sh in sl.Shapes:
        t = ""
        if sh.HasTextFrame and sh.TextFrame.HasText:
            t = sh.TextFrame.TextRange.Text
        parts.append(f"{sh.Id}:{sh.Left:.0f},{sh.Top:.0f},{sh.Width:.0f},{sh.Height:.0f}:{hash(t)}")
    return str(hash("|".join(parts)))


# ---------------------------------------------------------------- COM work (runs on the COM thread)

def _open(path, slide):
    app = _app()
    if path:
        full = os.path.abspath(path)
        hit = next((p for p in app.Presentations if p.FullName.lower() == full.lower()), None)
        hit = hit or next((p for p in app.Presentations if p.Name.lower() == os.path.basename(path).lower()), None)
        if hit is None:
            if not os.path.exists(full):
                raise RuntimeError(f"not found: {full}")
            hit = app.Presentations.Open(full)
        S.prs_name = hit.FullName
    prs = _prs()
    S.current = _clamp(prs, slide or S.current)
    S.version += 1
    _sync_window(prs)
    return _summary(prs)


def _run(code, slide):
    prs = _prs()
    if slide:
        S.current = _clamp(prs, slide)
    out = io.StringIO()

    def goto(n):
        S.current = _clamp(prs, n)

    import win32com.client
    scope = {"app": _app(), "prs": prs, "slide": prs.Slides(S.current) if S.current else None,
             "goto": goto, "win32com": win32com}
    error = None
    with contextlib.redirect_stdout(out):
        try:
            exec(compile(code, "<powerpoint_run>", "exec"), scope)
        except Exception:
            error = traceback.format_exc(limit=3)
    S.current = _clamp(prs, S.current or 1)
    S.version += 1
    _sync_window(prs)
    return {**_summary(prs), "output": out.getvalue()[-8000:], "error": error}


def _show(slide):
    prs = _prs()
    S.current = _clamp(prs, slide)
    S.version += 1
    _sync_window(prs)
    return _summary(prs)


def _state():
    prs = _prs()
    if S.current > prs.Slides.Count:
        S.current = prs.Slides.Count
    return {**_summary(prs), "signature": _signature(prs, S.current)}


def _image(slide, width):
    prs = _prs()
    n = _clamp(prs, slide or S.current)
    if not n:
        return {"slide": 0, "png": None}
    w = int(width)
    h = round(w * prs.PageSetup.SlideHeight / prs.PageSetup.SlideWidth)
    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        prs.Slides(n).Export(path, "PNG", w, h)
        with open(path, "rb") as f:
            png = base64.b64encode(f.read()).decode()
    finally:
        os.remove(path)
    return {"slide": n, "count": prs.Slides.Count, "width": w, "height": h, "png": png,
            "signature": _signature(prs, n)}


def _title(sl):
    for sh in sl.Shapes:
        try:
            if sh.Type == 14 and sh.PlaceholderFormat.Type in (1, 3) and sh.TextFrame.HasText:  # title, ctrTitle
                return sh.TextFrame.TextRange.Text.strip()
        except Exception:  # a placeholder without a usable PlaceholderFormat: skip it, look at the next shape
            continue
    return ""


def _sections(prs):
    sp = prs.SectionProperties
    out = []
    for i in range(1, sp.Count + 1):
        out.append({"index": i, "name": sp.Name(i), "first": sp.FirstSlide(i), "count": sp.SlidesCount(i)})
    return out


def _outline():
    prs = _prs()
    sections = _sections(prs) or [{"index": 0, "name": "", "first": 1, "count": prs.Slides.Count}]
    slides = []
    for n in range(1, prs.Slides.Count + 1):
        sl = prs.Slides(n)
        slides.append({"index": n, "id": sl.SlideID, "section": sl.sectionIndex if prs.SectionProperties.Count else 0,
                       "title": _title(sl), "hidden": bool(sl.SlideShowTransition.Hidden),
                       "signature": _signature(prs, n)})
    for sec in sections:
        sec["slides"] = [x["id"] for x in slides if x["section"] == sec["index"]]
    return {**_summary(prs), "sections": sections, "slides": slides}


_thumbs = {}


def _thumbs_for(ids, width):
    prs = _prs()
    out = []
    for sid in ids:
        try:
            sl = prs.Slides.FindBySlideID(int(sid))
        except Exception:  # the slide was deleted since the outline was read: nothing to draw
            continue
        sig = _signature(prs, sl.SlideIndex)
        key = (prs.FullName, sid, sig, width)
        if key not in _thumbs:
            h = round(width * prs.PageSetup.SlideHeight / prs.PageSetup.SlideWidth)
            fd, path = tempfile.mkstemp(suffix=".png")
            os.close(fd)
            try:
                sl.Export(path, "PNG", width, h)
                with open(path, "rb") as f:
                    _thumbs[key] = base64.b64encode(f.read()).decode()
            finally:
                os.remove(path)
            if len(_thumbs) > 400:
                _thumbs.pop(next(iter(_thumbs)))
        out.append({"id": sid, "signature": sig, "png": _thumbs[key]})
    return {"thumbs": out}


def _move(slide_id, before_id, section):
    prs = _prs()
    sl = prs.Slides.FindBySlideID(int(slide_id))
    cur_id = prs.Slides(S.current).SlideID if S.current else None
    has_sections = prs.SectionProperties.Count > 0
    if before_id:
        if int(before_id) == int(slide_id):
            return {**_summary(prs), "moved": False}
        b = prs.Slides.FindBySlideID(int(before_id))
        bi, bsec = b.SlideIndex, (b.sectionIndex if has_sections else 0)
        starts_section = has_sections and prs.SectionProperties.FirstSlide(bsec) == bi
        if starts_section:
            sl.MoveToSectionStart(bsec)
        else:
            sl.MoveTo(bi - 1 if sl.SlideIndex < bi else bi)
    elif has_sections and section:
        sp = prs.SectionProperties
        sec = int(section)
        if sp.SlidesCount(sec) == 0 or (sp.SlidesCount(sec) == 1 and sl.sectionIndex == sec):
            sl.MoveToSectionStart(sec)
        else:
            last = sp.FirstSlide(sec) + sp.SlidesCount(sec) - 1
            sl.MoveTo(last if sl.SlideIndex <= last else last + 1)
            if sl.sectionIndex != sec:  # landed on the far side of the boundary
                sl.MoveTo(sl.SlideIndex - 1)
    else:
        sl.MoveTo(prs.Slides.Count)
    if cur_id is not None:
        S.current = prs.Slides.FindBySlideID(cur_id).SlideIndex
    S.version += 1
    return {**_summary(prs), "moved": True, "now_at": sl.SlideIndex,
            "section": sl.sectionIndex if has_sections else 0}


# ---------------------------------------------------------------- MCP surface

@mcp.resource(URI, name="PowerPoint Live slide view", mime_type=MIME,
              meta={"ui": {"prefersBorder": True}})
def slide_view() -> str:
    with open(os.path.join(HERE, "view.html"), encoding="utf-8") as f:
        return f.read()


@mcp.tool(meta=UI)
async def powerpoint_open(path: str = "", slide: int = 0) -> dict:
    """Open a .pptx in desktop PowerPoint (or attach to it if already open; empty path = the active deck)
    and make `slide` (1-based) the current slide. The live view shows the current slide as it changes."""
    return await on_com(_open, path, slide)


@mcp.tool(meta=UI)
async def powerpoint_run(code: str, slide: int = 0) -> dict:
    """Run Python against the open deck through COM. In scope: `app` (PowerPoint.Application), `prs` (the deck),
    `slide` (the current slide object), `goto(n)` (change the current slide - the live view follows), `win32com`.
    print() output is returned. Pass `slide` to make that slide current first. Pin to `prs`, never
    app.ActivePresentation. To check the result closely, render with scripts/render_slides.py."""
    return await on_com(_run, code, slide)


@mcp.tool(meta=UI)
async def powerpoint_show(slide: int) -> dict:
    """Make slide n (1-based) the current slide in the live view and the PowerPoint window."""
    return await on_com(_show, slide)


@mcp.tool(meta=UI)
async def powerpoint_move(slide_id: int, before_slide_id: int = 0, section: int = 0) -> dict:
    """Rearrange slides. Move the slide with SlideID `slide_id` so it sits just before `before_slide_id`, or (with
    `section`, 1-based) at the end of that section; neither = to the end of the deck. SlideIDs are stable across
    moves - get them from prs.Slides(n).SlideID. The slide sorter view uses this for drag and drop."""
    return await on_com(_move, slide_id, before_slide_id, section)


@mcp.tool(meta=APP_ONLY)
async def deck_outline() -> dict:
    """(View only) Sections and slides (id, index, title, hidden, change signature) for the slide sorter."""
    return await on_com(_outline)


@mcp.tool(meta=APP_ONLY)
async def slide_thumbs(slide_ids: list[int], width: int = 320) -> dict:
    """(View only) Small PNGs of the given slides, cached by content."""
    return await on_com(_thumbs_for, slide_ids[:12], max(120, min(width, 640)))


@mcp.tool(meta=APP_ONLY)
async def slide_state() -> dict:
    """(View only) Current deck, slide, count and a change signature of the current slide."""
    return await on_com(_state)


@mcp.tool(meta=APP_ONLY)
async def slide_image(slide: int = 0, width: int = 1280) -> dict:
    """(View only) PNG of a slide (default: the current one), base64."""
    return await on_com(_image, slide, max(320, min(width, 2560)))


if __name__ == "__main__":
    mcp.run()
