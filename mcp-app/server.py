"""PowerPoint Live - an MCP App that drives desktop PowerPoint (Windows, COM) and shows the deck live.

    uvx --with "mcp<2" --with python-pptx --with pillow --with "pywin32; sys_platform == 'win32'" python mcp-app/server.py

Hosts that support MCP Apps (Claude Desktop, claude.ai) show the live view; elsewhere the tools still work.

Tools for the model (the view uses them too):
  powerpoint_open      open a deck (or attach to an open one) and pick the current slide
  powerpoint_run       run Python against the deck: app, prs, slide, goto(n), win32com in scope
  powerpoint_show      make slide n current
  powerpoint_move      move slides (by SlideID) before a slide or to the end of a section
  powerpoint_sections  add, rename or delete sections
  powerpoint_hide      hide or unhide slides
  powerpoint_fix       apply a lint fix to one shape, live
  powerpoint_history   list saved versions; powerpoint_restore puts one back
Tools only the view calls: slide_state, slide_image, deck_outline, slide_thumbs, deck_lint, history_thumb.

Every change made through these tools is preceded by a saved version, so it can be undone from the History view.
"""
import asyncio
import base64
import contextlib
import io
import os
import shutil
import sys
import tempfile
import time
import traceback
from concurrent.futures import ThreadPoolExecutor

from mcp.server.fastmcp import FastMCP

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(HERE), "scripts")
URI = "ui://powerpoint-live/slide-view"
MIME = "text/html;profile=mcp-app"
UI = {"ui": {"resourceUri": URI}, "ui/resourceUri": URI}
APP_ONLY = {"ui": {"resourceUri": URI, "visibility": ["app"]}, "ui/resourceUri": URI}
RAMP = [0, 0.35, -0.3, 0.6, -0.5, 0.75]
LINE_CHARTS = {4, 63, 64, 65, 66, 67}
FIXABLE = {"unused_placeholder", "body_below_floor", "text_overflow", "picture_stretched", "a11y_missing_alt_text",
           "chart_default_palette", "chart_legend_steals_plot", "chart_accounting_zero_dash"}
MAX_VERSIONS = 25

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
SEEN = {}       # SlideID -> shape map last seen (for change highlights)
CHANGES = {}    # SlideID -> {"stamp", "boxes"}
HISTORY = []    # [{"v", "time", "label", "path", "thumb", "slide"}]
THUMBS = {}
LINT = {"key": None, "result": None}
WORK = tempfile.mkdtemp(prefix="pptlive-")


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
            "count": prs.Slides.Count, "version": S.version,
            "width_pt": prs.PageSetup.SlideWidth, "height_pt": prs.PageSetup.SlideHeight}


def _text(sh):
    try:
        if sh.HasTextFrame and sh.TextFrame.HasText:
            return sh.TextFrame.TextRange.Text
    except Exception:  # some shape kinds (OLE, media) raise on HasTextFrame: they simply have no text
        pass
    return ""


def _shape_map(sl):
    """Top-level shapes of a slide: id -> (name, box, text hash)."""
    out = {}
    for sh in sl.Shapes:
        t = _text(sh)
        out[sh.Id] = {"name": sh.Name, "box": [round(sh.Left, 1), round(sh.Top, 1), round(sh.Width, 1),
                                                round(sh.Height, 1)], "text": hash(t), "snippet": t[:80]}
    return out


def _signature(prs, n):
    """Cheap fingerprint of slide n. Changes when anyone edits the slide."""
    if not n:
        return "empty"
    sl = prs.Slides(n)
    m = _shape_map(sl)
    return str(hash((sl.SlideID, bool(sl.SlideShowTransition.Hidden),
                     tuple((k, tuple(v["box"]), v["text"]) for k, v in m.items()))))


def _diff(before, after):
    boxes = []
    for k, v in after.items():
        old = before.get(k)
        if old is None:
            boxes.append({"id": k, "name": v["name"], "kind": "added", "box": v["box"]})
        elif old["box"] != v["box"] or old["text"] != v["text"]:
            boxes.append({"id": k, "name": v["name"], "kind": "changed", "box": v["box"]})
    for k, v in before.items():
        if k not in after:
            boxes.append({"id": k, "name": v["name"], "kind": "removed", "box": v["box"]})
    return boxes


def _note_changes(sl, before=None):
    """Compare a slide with what was last seen (or `before`) and remember what changed."""
    after = _shape_map(sl)
    prev = before if before is not None else SEEN.get(sl.SlideID)
    if prev is not None:
        boxes = _diff(prev, after)
        if boxes:
            CHANGES[sl.SlideID] = {"stamp": time.time(), "boxes": boxes}
    SEEN[sl.SlideID] = after
    return after


def _export(sl, width, path):
    prs = sl.Parent
    h = round(width * prs.PageSetup.SlideHeight / prs.PageSetup.SlideWidth)
    sl.Export(path, "PNG", int(width), h)
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def _png(sl, width):
    fd, path = tempfile.mkstemp(suffix=".png", dir=WORK)
    os.close(fd)
    try:
        return _export(sl, width, path)
    finally:
        os.remove(path)


def _walk(shapes):
    for sh in shapes:
        yield sh
        if sh.Type == 6:  # msoGroup
            yield from _walk(sh.GroupItems)


def _find_shape(sl, name=None, shape_id=None):
    for sh in _walk(sl.Shapes):
        if (shape_id and sh.Id == int(shape_id)) or (name and sh.Name == name):
            return sh
    return None


# ---------------------------------------------------------------- versions (undo)

def _snapshot(label):
    """Save a copy of the deck as a version before a change. SaveCopyAs never rebinds the open deck."""
    prs = _prs()
    v = (HISTORY[-1]["v"] + 1) if HISTORY else 1
    ext = os.path.splitext(prs.FullName)[1] or ".pptx"
    path = os.path.join(WORK, f"v{v:03d}{ext}")
    prs.SaveCopyAs(path)
    n = _clamp(prs, S.current or 1)
    thumb = _png(prs.Slides(n), 240) if n else None
    HISTORY.append({"v": v, "time": time.time(), "label": label, "path": path, "thumb": thumb, "slide": n,
                    "count": prs.Slides.Count})
    while len(HISTORY) > MAX_VERSIONS:
        old = HISTORY.pop(0)
        with contextlib.suppress(OSError):
            os.remove(old["path"])
    return v


def _history():
    return {"versions": [{k: h[k] for k in ("v", "time", "label", "slide", "count")} for h in reversed(HISTORY)]}


def _history_thumb(v):
    h = next((h for h in HISTORY if h["v"] == int(v)), None)
    return {"v": int(v), "png": h["thumb"] if h else None}


def _restore(v):
    h = next((h for h in HISTORY if h["v"] == int(v)), None)
    if h is None:
        raise RuntimeError(f"no version {v}")
    prs = _prs()
    path = prs.FullName
    if not os.path.isabs(path):
        raise RuntimeError("the deck has never been saved; save it once, then versions can be restored")
    _snapshot(f"before restoring v{v}")
    with_window = prs.Windows.Count > 0
    prs.Close()                      # COM Close discards unsaved changes without a prompt: they are in the version
    shutil.copyfile(h["path"], path)
    new = _app().Presentations.Open(path, 0, 0, -1 if with_window else 0)
    S.prs_name = new.FullName
    S.current = _clamp(new, h["slide"] or 1)
    SEEN.clear()
    S.version += 1
    _sync_window(new)
    return {**_summary(new), "restored": int(v)}


def _mutate(label, fn):
    """Version first, then change, then record what changed on every slide."""
    prs = _prs()
    _snapshot(label)
    before = {prs.Slides(i).SlideID: _shape_map(prs.Slides(i)) for i in range(1, prs.Slides.Count + 1)}
    result = fn(prs)
    for i in range(1, prs.Slides.Count + 1):
        sl = prs.Slides(i)
        if sl.SlideID in before:
            _note_changes(sl, before[sl.SlideID])
        else:
            CHANGES[sl.SlideID] = {"stamp": time.time(), "boxes": [{"id": 0, "name": "new slide", "kind": "added",
                                                                    "box": [0, 0, prs.PageSetup.SlideWidth,
                                                                            prs.PageSetup.SlideHeight]}]}
            SEEN[sl.SlideID] = _shape_map(sl)
    S.current = _clamp(prs, S.current or 1)
    S.version += 1
    _sync_window(prs)
    return {**_summary(prs), **(result or {})}


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
        if hit.FullName != S.prs_name:
            SEEN.clear()
            CHANGES.clear()
        S.prs_name = hit.FullName
    prs = _prs()
    S.current = _clamp(prs, slide or S.current)
    S.version += 1
    _sync_window(prs)
    return _summary(prs)


def _run(code, slide, label):
    out = io.StringIO()
    box = {}

    def body(prs):
        if slide:
            S.current = _clamp(prs, slide)

        def goto(n):
            S.current = _clamp(prs, n)

        import win32com.client
        scope = {"app": _app(), "prs": prs, "slide": prs.Slides(S.current) if S.current else None,
                 "goto": goto, "win32com": win32com}
        with contextlib.redirect_stdout(out):
            try:
                exec(compile(code, "<powerpoint_run>", "exec"), scope)
            except Exception:
                box["error"] = traceback.format_exc(limit=3)
        return {}

    res = _mutate(label or "powerpoint_run", body)
    return {**res, "output": out.getvalue()[-8000:], "error": box.get("error")}


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
    out = {**_summary(prs), "signature": _signature(prs, S.current), "shapes": [], "changes": None}
    if S.current:
        sl = prs.Slides(S.current)
        shapes = _note_changes(sl)
        out["shapes"] = [{"id": k, "name": v["name"], "box": v["box"], "text": v["snippet"]} for k, v in shapes.items()]
        out["changes"] = CHANGES.get(sl.SlideID)
        out["slide_id"] = sl.SlideID
    return out


def _accurate_png(prs, n, width):
    """Render one slide through Save As JPEG (true embedded fonts), not Slide.Export."""
    src = os.path.join(WORK, "acc-src" + (os.path.splitext(prs.FullName)[1] or ".pptx"))
    prs.SaveCopyAs(src)
    tmp = _app().Presentations.Open(src, -1, 0, 0)  # read-only, no window
    try:
        for i in range(tmp.Slides.Count, 0, -1):
            if i != n:
                tmp.Slides(i).Delete()
        out = os.path.join(WORK, "acc-out")
        shutil.rmtree(out, ignore_errors=True)
        tmp.SaveCopyAs(out, 17)  # ppSaveAsJPG -> folder with Slide1.JPG
    finally:
        tmp.Close()
    jpg = next(os.path.join(out, f) for f in os.listdir(out) if f.lower().endswith(".jpg"))
    from PIL import Image
    im = Image.open(jpg)
    im = im.resize((int(width), round(int(width) * im.height / im.width)))
    buf = io.BytesIO()
    im.save(buf, "PNG")
    shutil.rmtree(out, ignore_errors=True)
    return base64.b64encode(buf.getvalue()).decode()


def _image(slide, width, accurate):
    prs = _prs()
    n = _clamp(prs, slide or S.current)
    if not n:
        return {"slide": 0, "png": None}
    sl = prs.Slides(n)
    png = _accurate_png(prs, n, width) if accurate else _png(sl, width)
    return {"slide": n, "slide_id": sl.SlideID, "count": prs.Slides.Count, "png": png, "accurate": bool(accurate),
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
    return [{"index": i, "name": sp.Name(i), "first": sp.FirstSlide(i), "count": sp.SlidesCount(i)}
            for i in range(1, sp.Count + 1)]


def _outline():
    prs = _prs()
    has = prs.SectionProperties.Count > 0
    sections = _sections(prs) or [{"index": 0, "name": "", "first": 1, "count": prs.Slides.Count}]
    slides = []
    for n in range(1, prs.Slides.Count + 1):
        sl = prs.Slides(n)
        notes = ""
        with contextlib.suppress(Exception):  # a slide without a notes page has no placeholder 2
            notes = sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        slides.append({"index": n, "id": sl.SlideID, "section": sl.sectionIndex if has else 0, "title": _title(sl),
                       "hidden": bool(sl.SlideShowTransition.Hidden), "notes_words": len(notes.split()),
                       "signature": _signature(prs, n)})
    for sec in sections:
        sec["slides"] = [x["id"] for x in slides if x["section"] == sec["index"]]
    return {**_summary(prs), "sections": sections, "slides": slides, "has_sections": has}


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
        if key not in THUMBS:
            THUMBS[key] = _png(sl, width)
            if len(THUMBS) > 400:
                THUMBS.pop(next(iter(THUMBS)))
        out.append({"id": sid, "signature": sig, "png": THUMBS[key]})
    return {"thumbs": out}


def _move_one(prs, sl, before_id, section):
    has = prs.SectionProperties.Count > 0
    if before_id:
        if int(before_id) == sl.SlideID:
            return
        b = prs.Slides.FindBySlideID(int(before_id))
        bi, bsec = b.SlideIndex, (b.sectionIndex if has else 0)
        if has and prs.SectionProperties.FirstSlide(bsec) == bi:
            sl.MoveToSectionStart(bsec)
        else:
            sl.MoveTo(bi - 1 if sl.SlideIndex < bi else bi)
    elif has and section:
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


def _move(slide_ids, before_id, section):
    def body(prs):
        cur = prs.Slides(S.current).SlideID if S.current else None
        ids = sorted((int(i) for i in slide_ids), key=lambda i: prs.Slides.FindBySlideID(i).SlideIndex)
        ids = [i for i in ids if i != int(before_id or 0)]
        for sid in ids:  # in deck order, each lands before the target / at the section end: order is kept
            _move_one(prs, prs.Slides.FindBySlideID(sid), before_id, section)
        if cur is not None:
            S.current = prs.Slides.FindBySlideID(cur).SlideIndex
        return {"moved": ids}
    return _mutate("move slides", body)


def _sections_edit(action, section, name, before_id):
    def body(prs):
        sp = prs.SectionProperties
        if action == "add":
            idx = prs.Slides.FindBySlideID(int(before_id)).SlideIndex if before_id else prs.Slides.Count + 1
            if idx > prs.Slides.Count:
                sp.AddSection(sp.Count + 1, name or "New section")
            else:
                sp.AddBeforeSlide(idx, name or "New section")
        elif action == "rename":
            sp.Rename(int(section), name)
        elif action == "delete":
            sp.Delete(int(section), False)  # keep the slides; they join the section before
        else:
            raise RuntimeError("action must be add, rename or delete")
        return {}
    return _mutate(f"section {action}", body)


def _hide(slide_ids, hidden):
    def body(prs):
        for sid in slide_ids:
            prs.Slides.FindBySlideID(int(sid)).SlideShowTransition.Hidden = -1 if hidden else 0
        return {}
    return _mutate("hide slides" if hidden else "unhide slides", body)


# ---------------------------------------------------------------- lint + live fixes

def _deck_key(prs):
    return (prs.FullName, prs.Slides.Count,
            tuple(_signature(prs, i) for i in range(1, prs.Slides.Count + 1)))


def _lint():
    prs = _prs()
    key = _deck_key(prs)
    if LINT["key"] == key:
        return LINT["result"]
    if SCRIPTS not in sys.path:
        sys.path.insert(0, SCRIPTS)
    import lint_deck
    copy = os.path.join(WORK, "lint" + (os.path.splitext(prs.FullName)[1] or ".pptx"))
    prs.SaveCopyAs(copy)
    _, f = lint_deck.lint(copy, 18.0, 12)
    out = []
    for i in f.items:
        item = {**i, "fixable": i["code"] in FIXABLE and bool(i["shape"]), "slide_id": None, "box": None}
        if i["slide"]:
            sl = prs.Slides(i["slide"])
            item["slide_id"] = sl.SlideID
            sh = _find_shape(sl, i["shape"], i.get("shape_id")) if i["shape"] else None
            if sh is not None:
                item["shape_id"] = sh.Id
                item["box"] = [round(sh.Left, 1), round(sh.Top, 1), round(sh.Width, 1), round(sh.Height, 1)]
        out.append(item)
    counts = {}
    for i in out:
        if i["slide_id"] and i["severity"] in ("error", "warn"):
            counts[i["slide_id"]] = counts.get(i["slide_id"], 0) + 1
    LINT["key"], LINT["result"] = key, {"floor_pt": f.floor, "findings": out, "counts": counts}
    return LINT["result"]


def _alt_for_chart(ch):
    parts = []
    with contextlib.suppress(Exception):
        for i in range(1, ch.SeriesCollection().Count + 1):
            s = ch.SeriesCollection(i)
            cats = [str(c) for c in (s.XValues or [])]
            vals = list(s.Values or [])
            pairs = ", ".join(f"{c} {v:g}" if isinstance(v, (int, float)) else f"{c} n/a" for c, v in zip(cats, vals))
            parts.append(f"{s.Name}: {pairs}")
    title = ""
    with contextlib.suppress(Exception):
        if ch.HasTitle:
            title = ch.ChartTitle.Text + ". "
    return f"Chart. {title}" + "; ".join(parts) + "."


def _fix_shape(prs, sh, code, floor):
    if code == "unused_placeholder":
        sh.Delete()
        return "deleted the empty placeholder"
    if code == "body_below_floor":
        n = 0
        for p in sh.TextFrame.TextRange.Paragraphs():
            if 0 < p.Font.Size < floor and len(p.Text.split()) >= 4:
                p.Font.Size = floor
                n += 1
        return f"{n} paragraph(s) raised to {floor:g} pt"
    if code == "text_overflow":
        tf, tr = sh.TextFrame, sh.TextFrame.TextRange
        room = sh.Height - tf.MarginTop - tf.MarginBottom
        steps = 0
        while tr.BoundHeight > room and steps < 40:
            runs = [r for r in tr.Runs() if r.Font.Size > floor]
            if not runs:
                break
            for r in runs:
                r.Font.Size = max(floor, r.Font.Size - 1)
            steps += 1
        left = " (still too long - cut words)" if tr.BoundHeight > room else ""
        return f"text shrunk {steps} pt to fit its box{left}"
    if code == "picture_stretched":
        l, t, w, h = sh.Left, sh.Top, sh.Width, sh.Height
        sh.LockAspectRatio = 0
        sh.ScaleHeight(1, -1)
        sh.ScaleWidth(1, -1)
        ratio = sh.Width / sh.Height
        nw, nh = (w, w / ratio) if w / h < ratio else (h * ratio, h)
        sh.Width, sh.Height = nw, nh
        sh.Left, sh.Top = l + (w - nw) / 2, t + (h - nh) / 2
        sh.LockAspectRatio = -1
        return "resized to its true proportions inside the same box"
    if code == "a11y_missing_alt_text":
        if sh.HasChart:
            sh.AlternativeText = _alt_for_chart(sh.Chart)
            return "chart alt text written from its data"
        if sh.HasTable:
            tb = sh.Table
            head = [tb.Cell(1, c).Shape.TextFrame.TextRange.Text for c in range(1, tb.Columns.Count + 1)]
            sh.AlternativeText = f"Table: {', '.join(head)}; {tb.Rows.Count - 1} rows."
            return "table alt text written from its header"
        raise RuntimeError("only Claude can describe a picture - ask Claude")
    ch = sh.Chart if sh.HasChart else None
    if ch is None:
        raise RuntimeError(f"{code}: not a chart")
    if code == "chart_default_palette":
        for i in range(1, ch.SeriesCollection().Count + 1):
            s = ch.SeriesCollection(i)
            fmt = s.Format.Line if ch.ChartType in LINE_CHARTS else s.Format.Fill
            fmt.ForeColor.ObjectThemeColor = 5  # msoThemeColorAccent1
            fmt.ForeColor.Brightness = RAMP[(i - 1) % len(RAMP)]
        return "series recoloured as shades of the theme accent"
    if code == "chart_legend_steals_plot":
        ch.Legend.Position = -4152  # xlLegendPositionRight
        ch.Legend.IncludeInLayout = False
        return "legend moved to the right"
    if code == "chart_accounting_zero_dash":
        ax = ch.Axes(2)
        ax.TickLabels.NumberFormat = "$#,##0" if "$" in ax.TickLabels.NumberFormat else "#,##0"
        return "Accounting axis format replaced"
    raise RuntimeError(f"{code} has no automatic fix - ask Claude")


def _fix(slide_id, shape_id, code):
    floor = (LINT["result"] or {}).get("floor_pt") or 18.0 * max(1.0, _prs().PageSetup.SlideWidth / 960)
    box = {}

    def body(prs):
        sl = prs.Slides.FindBySlideID(int(slide_id))
        sh = _find_shape(sl, shape_id=shape_id)
        if sh is None:
            raise RuntimeError("that shape is gone - re-run the lint")
        box["done"] = _fix_shape(prs, sh, code, floor)
        S.current = sl.SlideIndex
        return {"fixed": code, "result": box["done"]}
    return _mutate(f"fix {code}", body)


# ---------------------------------------------------------------- MCP surface

@mcp.resource(URI, name="PowerPoint Live", mime_type=MIME, meta={"ui": {"prefersBorder": True}})
def slide_view() -> str:
    with open(os.path.join(HERE, "view.html"), encoding="utf-8") as f:
        return f.read()


@mcp.tool(meta=UI)
async def powerpoint_open(path: str = "", slide: int = 0) -> dict:
    """Open a .pptx in desktop PowerPoint (or attach to it if already open; empty path = the active deck)
    and make `slide` (1-based) the current slide. The live view shows the deck as it changes."""
    return await on_com(_open, path, slide)


@mcp.tool(meta=UI)
async def powerpoint_run(code: str, slide: int = 0, label: str = "") -> dict:
    """Run Python against the open deck through COM. In scope: `app`, `prs` (the deck), `slide` (the current slide),
    `goto(n)` (change the current slide - the live view follows), `win32com`. print() output is returned.
    `label` names the change in the History view. A version is saved first, so the user can undo it.
    Changed shapes are highlighted in the live view. Pin to `prs`, never app.ActivePresentation."""
    return await on_com(_run, code, slide, label)


@mcp.tool(meta=UI)
async def powerpoint_show(slide: int) -> dict:
    """Make slide n (1-based) the current slide in the live view and the PowerPoint window."""
    return await on_com(_show, slide)


@mcp.tool(meta=UI)
async def powerpoint_move(slide_ids: list[int], before_slide_id: int = 0, section: int = 0) -> dict:
    """Move slides (by SlideID, kept in deck order) to just before `before_slide_id`, or (with `section`, 1-based)
    to the end of that section; neither = to the end of the deck. SlideIDs: prs.Slides(n).SlideID."""
    return await on_com(_move, slide_ids, before_slide_id, section)


@mcp.tool(meta=UI)
async def powerpoint_sections(action: str, section: int = 0, name: str = "", before_slide_id: int = 0) -> dict:
    """Edit sections. action: "add" (a section named `name` starting at `before_slide_id`, or at the end),
    "rename" (section n to `name`), "delete" (section n; its slides join the section before)."""
    return await on_com(_sections_edit, action, section, name, before_slide_id)


@mcp.tool(meta=UI)
async def powerpoint_hide(slide_ids: list[int], hidden: bool = True) -> dict:
    """Hide (or unhide) slides in the slide show, by SlideID."""
    return await on_com(_hide, slide_ids, hidden)


@mcp.tool(meta=UI)
async def powerpoint_fix(slide_id: int, shape_id: int, code: str) -> dict:
    """Apply the automatic fix for a lint finding to one shape, live: unused_placeholder, body_below_floor,
    text_overflow (shrinks using PowerPoint's own layout, never below the floor), picture_stretched,
    a11y_missing_alt_text (charts and tables), chart_default_palette, chart_legend_steals_plot,
    chart_accounting_zero_dash. A version is saved first."""
    return await on_com(_fix, slide_id, shape_id, code)


@mcp.tool(meta=UI)
async def powerpoint_history() -> dict:
    """Versions saved before each change made through these tools (newest first)."""
    return await on_com(_history)


@mcp.tool(meta=UI)
async def powerpoint_restore(version: int) -> dict:
    """Put a saved version back: the deck file is replaced with it and reopened. The current state is saved
    as a version first, so a restore can itself be undone. Needs a deck that has been saved to disk."""
    return await on_com(_restore, version)


@mcp.tool(meta=APP_ONLY)
async def slide_state() -> dict:
    """(View only) Current slide, its shapes, recent changes and a change signature."""
    return await on_com(_state)


@mcp.tool(meta=APP_ONLY)
async def slide_image(slide: int = 0, width: int = 1280, accurate: bool = False) -> dict:
    """(View only) PNG of a slide. accurate=True renders through Save As JPEG (true embedded fonts), slower."""
    return await on_com(_image, slide, max(320, min(width, 2560)), accurate)


@mcp.tool(meta=APP_ONLY)
async def deck_outline() -> dict:
    """(View only) Sections and slides (id, title, hidden, notes words, change signature)."""
    return await on_com(_outline)


@mcp.tool(meta=APP_ONLY)
async def slide_thumbs(slide_ids: list[int], width: int = 320) -> dict:
    """(View only) Small PNGs of the given slides, cached by content."""
    return await on_com(_thumbs_for, slide_ids[:12], max(120, min(width, 640)))


@mcp.tool(meta=APP_ONLY)
async def deck_lint() -> dict:
    """(View only) lint_deck.py findings with shape boxes, cached until the deck changes."""
    return await on_com(_lint)


@mcp.tool(meta=APP_ONLY)
async def history_thumb(version: int) -> dict:
    """(View only) Thumbnail of the current slide when a version was saved."""
    return await on_com(_history_thumb, version)


if __name__ == "__main__":
    mcp.run()
