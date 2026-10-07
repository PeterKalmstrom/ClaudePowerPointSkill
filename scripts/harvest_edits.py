"""Keep people's PowerPoint edits when a generated deck is rebuilt. Any OS (python-pptx).

A deck built from a spec gets hand-edited - a colleague rewords a slide, adds one, deletes one. The next
rebuild from the spec would silently erase all of that. This script stops that:

    # 1. after every build, record what the builder produced
    python scripts/harvest_edits.py manifest deck.pptx                      -> deck.manifest.json
    # 2. before rebuilding, collect every slide people changed, added or deleted since then
    python scripts/harvest_edits.py harvest deck.pptx --store edits/        -> edits/edits.pptx + edits.json
    # 3. rebuild (build_deck.py ... --out new.pptx), then put the edits back
    python scripts/harvest_edits.py restore new.pptx --store edits/ --out final.pptx   (+ final.manifest.json)

Slides are matched by their stable id (`<p:cSld name>`, the spec's `id`). An edited slide is copied back
XML-exactly with its pictures and charts, so the person's wording and layout survive; a slide the person
added is inserted after the slide it followed; a slide they deleted stays deleted. A person's edit is
content - nothing here judges it. `restore` prints every decision, so review the list.

Saving in PowerPoint changes a few things without anyone editing: the height and top of "resize shape
to fit text" boxes, and empty text runs. The fingerprint ignores those, so a plain re-save is not an edit.
"""
import argparse
import copy
import hashlib
import io
import json
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import PartFactory, _Relationship

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P}


# ---------------------------------------------------------------- fingerprints

def slide_name(slide, n):
    return slide._element.find("p:cSld", NS).get("name") or f"#{n}"


def notes_text(slide):
    return slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""


def fingerprint(slide):
    """Hash of the slide's content, ignoring what a plain PowerPoint save changes."""
    el = copy.deepcopy(slide._element)
    for r in el.iter(f"{{{A}}}r"):  # empty runs are dropped on save
        t = r.find(f"{{{A}}}t")
        if t is None or not (t.text or ""):
            r.getparent().remove(r)
    for sp in el.iter(f"{{{P}}}sp"):  # autofit boxes get a new height/top on save
        if sp.find(f".//{{{A}}}spAutoFit") is not None:
            xfrm = sp.find(f"{{{P}}}spPr/{{{A}}}xfrm")
            if xfrm is not None:
                for tag, attr in (("off", "y"), ("ext", "cy")):
                    node = xfrm.find(f"{{{A}}}{tag}")
                    if node is not None:
                        node.set(attr, "0")
    for node in el.iter():  # revision ids differ between saves
        for k in [k for k in node.attrib if k.endswith("}rsid") or k.startswith("rsid")]:
            del node.attrib[k]
    blob = etree.tostring(el, method="c14n") + notes_text(slide).encode("utf-8")
    # pictures: their bytes, not just their rIds
    for rel in slide.part.rels.values():
        if not rel.is_external and rel.reltype == RT.IMAGE:
            blob += hashlib.sha1(rel.target_part.blob).digest()
    return hashlib.sha1(blob).hexdigest()


def deck_manifest(path):
    prs = Presentation(path)
    slides = [{"name": slide_name(s, n), "fingerprint": fingerprint(s)} for n, s in enumerate(prs.slides, 1)]
    return {"deck": os.path.basename(path), "slides": slides}


# ---------------------------------------------------------------- slide copying

def copy_part(part, package, done):
    """Copy a part (chart, embedded workbook, diagram ...) and everything it relates to."""
    if part in done:
        return done[part]
    ext = os.path.splitext(part.partname)[1]
    base = os.path.dirname(part.partname)
    stem = "".join(c for c in os.path.basename(part.partname)[:-len(ext)] if not c.isdigit())
    partname = package.next_partname(f"{base}/{stem}%d{ext}")
    new = PartFactory(partname, part.content_type, package, part.blob)
    done[part] = new
    for rId, rel in part.rels.items():
        target = rel.target_ref if rel.is_external else copy_part(rel.target_part, package, done)
        new.rels._rels[rId] = _Relationship(new.rels._base_uri, rId, rel.reltype,
                                            "External" if rel.is_external else "Internal", target)
    new.rels.__dict__.pop("_rels_by_reltype", None)
    return new


def copy_slide(src, dst_prs, position=None):
    """Append a copy of `src` (from another presentation) to dst_prs, keeping its rIds."""
    layout = next((l for l in dst_prs.slide_layouts if l.name == src.slide_layout.name), dst_prs.slide_layouts[0])
    new = dst_prs.slides.add_slide(layout)
    for ph in list(new.placeholders):  # the layout's empty placeholders; the copy brings its own shapes
        ph._element.getparent().remove(ph._element)
    package, done = dst_prs.part.package, {}
    for rId, rel in src.part.rels.items():
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        if rel.is_external:
            target = rel.target_ref
        elif rel.reltype == RT.IMAGE:
            target, _ = new.part.get_or_add_image_part(io.BytesIO(rel.target_part.blob))
        else:
            target = copy_part(rel.target_part, package, done)
        new.part.rels._rels[rId] = _Relationship(new.part.rels._base_uri, rId, rel.reltype,
                                                 "External" if rel.is_external else "Internal", target)
    new.part.rels.__dict__.pop("_rels_by_reltype", None)
    # replace the whole slide element content (cSld incl. name, transitions, timing) XML-exactly
    el = new._element
    for child in list(el):
        el.remove(child)
    for child in src._element:
        el.append(copy.deepcopy(child))
    new.__dict__.pop("shapes", None)  # drop python-pptx's cached shape proxy for the old tree
    if src.has_notes_slide and notes_text(src):
        new.notes_slide.notes_text_frame.text = notes_text(src)
    if position is not None:
        move_slide(dst_prs, len(dst_prs.slides) - 1, position)
    return new


def move_slide(prs, old, new):
    lst = prs.slides._sldIdLst
    item = list(lst)[old]
    lst.remove(item)
    lst.insert(new, item)


def delete_slide(prs, index):
    lst = prs.slides._sldIdLst
    item = list(lst)[index]
    prs.part.drop_rel(item.get(f"{{{R}}}id"))
    lst.remove(item)
    # renumber slide parts, or the next add_slide reuses slideN.xml and the zip gets duplicate entries
    prs.part.rename_slide_parts([sld.get(f"{{{R}}}id") for sld in lst])


# ---------------------------------------------------------------- commands

def cmd_manifest(deck, out=None):
    out = out or os.path.splitext(deck)[0] + ".manifest.json"
    json.dump(deck_manifest(deck), open(out, "w", encoding="utf-8"), indent=2)
    print(f"manifest: {out}")
    return out


def cmd_harvest(deck, manifest, store):
    base = {s["name"]: s["fingerprint"] for s in json.load(open(manifest, encoding="utf-8"))["slides"]}
    base_order = [s["name"] for s in json.load(open(manifest, encoding="utf-8"))["slides"]]
    prs = Presentation(deck)
    os.makedirs(store, exist_ok=True)
    keep = Presentation(deck)  # the store is the deck minus unchanged slides: layouts and media come along
    record, seen, prev = [], set(), None
    keep_idx = []
    for n, s in enumerate(prs.slides, 1):
        name = slide_name(s, n)
        dup = name in seen
        seen.add(name)
        if dup or name not in base:
            status = "added"
            name = f"{name}~{n}" if dup else name
        elif fingerprint(s) != base[name]:
            status = "edited"
        else:
            status = None
        if status == "added" and name.startswith("#"):
            name = f"added-{n}"  # a slide made in PowerPoint has no id yet: give it a permanent one
        if status:
            record.append({"name": name, "status": status, "after": prev, "store_index": len(keep_idx)})
            keep_idx.append(n - 1)
            if status == "added":  # store the id with the slide so the next manifest knows it
                keep.slides[n - 1]._element.find("p:cSld", NS).set("name", name)
        prev = name
    for name in base_order:
        if name not in seen:
            record.append({"name": name, "status": "deleted"})
    for i in reversed(range(len(keep.slides))):
        if i not in keep_idx:
            delete_slide(keep, i)
    keep.save(os.path.join(store, "edits.pptx"))
    json.dump({"source": os.path.basename(deck), "slides": record},
              open(os.path.join(store, "edits.json"), "w", encoding="utf-8"), indent=2)
    for r in record:
        print(f"{r['status']:<8} {r['name']}")
    print(f"{sum(1 for r in record if r['status'] != 'deleted')} slide(s) harvested, "
          f"{sum(1 for r in record if r['status'] == 'deleted')} deletion(s) recorded -> {store}")


def cmd_restore(new_deck, store, out):
    meta = json.load(open(os.path.join(store, "edits.json"), encoding="utf-8"))["slides"]
    edits = Presentation(os.path.join(store, "edits.pptx"))
    prs = Presentation(new_deck)
    names = lambda: [slide_name(s, n) for n, s in enumerate(prs.slides, 1)]
    for r in meta:
        if r["status"] == "deleted" and r["name"] in names():
            delete_slide(prs, names().index(r["name"]))
            print(f"kept deleted   {r['name']}")
        elif r["status"] == "deleted":
            print(f"(gone anyway)  {r['name']}")
    for r in meta:
        if r["status"] == "edited":
            src = edits.slides[r["store_index"]]
            if r["name"] in names():
                pos = names().index(r["name"])
                delete_slide(prs, pos)
                copy_slide(src, prs, pos)
                print(f"kept edit      {r['name']}")
            else:
                print(f"!! not in the new build, re-added at the end: {r['name']}")
                copy_slide(src, prs)
    for r in meta:
        if r["status"] == "added":
            src = edits.slides[r["store_index"]]
            pos = names().index(r["after"]) + 1 if r["after"] in names() else len(prs.slides)
            copy_slide(src, prs, pos)
            print(f"kept added     {r['name']} (after {r['after'] or 'start'})")
    prs.save(out)
    cmd_manifest(out)
    print(f"-> {out}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    m = sub.add_parser("manifest", help="record a freshly built deck")
    m.add_argument("deck")
    m.add_argument("--out")
    h = sub.add_parser("harvest", help="collect slides changed since the manifest")
    h.add_argument("deck")
    h.add_argument("--manifest", help="default: <deck>.manifest.json")
    h.add_argument("--store", required=True)
    r = sub.add_parser("restore", help="put harvested edits into a rebuilt deck")
    r.add_argument("deck")
    r.add_argument("--store", required=True)
    r.add_argument("--out", required=True)
    a = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if a.cmd == "manifest":
        cmd_manifest(a.deck, a.out)
    elif a.cmd == "harvest":
        cmd_harvest(a.deck, a.manifest or os.path.splitext(a.deck)[0] + ".manifest.json", a.store)
    else:
        if os.path.abspath(a.out) == os.path.abspath(a.deck):
            sys.exit("--out must differ from the rebuilt deck")
        cmd_restore(a.deck, a.store, a.out)


if __name__ == "__main__":
    main()
