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

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import PartFactory, _Relationship

from kShared import ToolReportableException, kRun, kS, kToolException

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P}


class kSlideFingerprint:
    """Slide identity and content hashes (python-pptx)."""

    @staticmethod
    def Name(Slide, Number):
        """The slide's stable id (<p:cSld name>), or #<number> when it has none."""
        if kS.ErrorMode:
            return ""
        try:
            return Slide._element.find("p:cSld", NS).get("name") or f"#{Number}"
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideFingerprint.Name")
            return ""

    @staticmethod
    def Names(Deck):
        """Every slide's id, in order."""
        if kS.ErrorMode:
            return []
        try:
            return [kSlideFingerprint.Name(Slide, Number) for Number, Slide in enumerate(Deck.slides, 1)]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideFingerprint.Names")
            return []

    @staticmethod
    def NotesText(Slide):
        """The speaker notes, or ""."""
        if kS.ErrorMode:
            return ""
        try:
            return Slide.notes_slide.notes_text_frame.text if Slide.has_notes_slide else ""
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideFingerprint.NotesText")
            return ""

    @staticmethod
    def Fingerprint(Slide):
        """Hash of the slide's content, ignoring what a plain PowerPoint save changes."""
        if kS.ErrorMode:
            return ""
        try:
            Element = copy.deepcopy(Slide._element)
            for Run in Element.iter(f"{{{A}}}r"):  # empty runs are dropped on save
                Text = Run.find(f"{{{A}}}t")
                if Text is None or not (Text.text or ""):
                    Run.getparent().remove(Run)
            for Shape in Element.iter(f"{{{P}}}sp"):  # autofit boxes get a new height/top on save
                if Shape.find(f".//{{{A}}}spAutoFit") is not None:
                    Xfrm = Shape.find(f"{{{P}}}spPr/{{{A}}}xfrm")
                    if Xfrm is not None:
                        for Tag, Attribute in (("off", "y"), ("ext", "cy")):
                            Node = Xfrm.find(f"{{{A}}}{Tag}")
                            if Node is not None:
                                Node.set(Attribute, "0")
            for Node in Element.iter():  # revision ids differ between saves
                for Key in [Key for Key in Node.attrib if Key.endswith("}rsid") or Key.startswith("rsid")]:
                    del Node.attrib[Key]
            Blob = etree.tostring(Element, method="c14n") + kSlideFingerprint.NotesText(Slide).encode("utf-8")
            # pictures: their bytes, not just their rIds
            for Rel in Slide.part.rels.values():
                if not Rel.is_external and Rel.reltype == RT.IMAGE:
                    Blob += hashlib.sha1(Rel.target_part.blob).digest()
            return hashlib.sha1(Blob).hexdigest()
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideFingerprint.Fingerprint")
            return ""

    @staticmethod
    def DeckManifest(Path):
        """{"deck": file name, "slides": [{name, fingerprint}]}."""
        if kS.ErrorMode:
            return None
        try:
            Deck = Presentation(Path)
            Slides = [{"name": kSlideFingerprint.Name(Slide, Number), "fingerprint": kSlideFingerprint.Fingerprint(Slide)}
                      for Number, Slide in enumerate(Deck.slides, 1)]
            return {"deck": os.path.basename(Path), "slides": Slides}
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSlideFingerprint.DeckManifest(file={Path})")
            return None


class kSlideCopier:
    """Copies, moves and deletes slides between python-pptx presentations."""

    @staticmethod
    def CopyPart(Part, Package, Done):
        """Copy a part (chart, embedded workbook, diagram ...) and everything it relates to."""
        if kS.ErrorMode:
            return None
        try:
            if Part in Done:
                return Done[Part]
            Extension = os.path.splitext(Part.partname)[1]
            Base = os.path.dirname(Part.partname)
            Stem = "".join(C for C in os.path.basename(Part.partname)[:-len(Extension)] if not C.isdigit())
            PartName = Package.next_partname(f"{Base}/{Stem}%d{Extension}")
            New = PartFactory(PartName, Part.content_type, Package, Part.blob)
            Done[Part] = New
            for RId, Rel in Part.rels.items():
                Target = Rel.target_ref if Rel.is_external else kSlideCopier.CopyPart(Rel.target_part, Package, Done)
                New.rels._rels[RId] = _Relationship(New.rels._base_uri, RId, Rel.reltype,
                                                    "External" if Rel.is_external else "Internal", Target)
            New.rels.__dict__.pop("_rels_by_reltype", None)
            return New
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSlideCopier.CopyPart({Part.partname})")
            return None

    @staticmethod
    def CopySlide(Source, Deck, Position=None):
        """Append a copy of Source (from another presentation) to Deck, keeping its rIds."""
        if kS.ErrorMode:
            return None
        try:
            Layout = next((Candidate for Candidate in Deck.slide_layouts
                           if Candidate.name == Source.slide_layout.name), Deck.slide_layouts[0])
            New = Deck.slides.add_slide(Layout)
            for Placeholder in list(New.placeholders):  # the layout's empty placeholders; the copy brings its own
                Placeholder._element.getparent().remove(Placeholder._element)
            Package, Done = Deck.part.package, {}
            for RId, Rel in Source.part.rels.items():
                if Rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
                    continue
                if Rel.is_external:
                    Target = Rel.target_ref
                elif Rel.reltype == RT.IMAGE:
                    Target, _ = New.part.get_or_add_image_part(io.BytesIO(Rel.target_part.blob))
                else:
                    Target = kSlideCopier.CopyPart(Rel.target_part, Package, Done)
                New.part.rels._rels[RId] = _Relationship(New.part.rels._base_uri, RId, Rel.reltype,
                                                         "External" if Rel.is_external else "Internal", Target)
            New.part.rels.__dict__.pop("_rels_by_reltype", None)
            # replace the whole slide element content (cSld incl. name, transitions, timing) XML-exactly
            Element = New._element
            for Child in list(Element):
                Element.remove(Child)
            for Child in Source._element:
                Element.append(copy.deepcopy(Child))
            New.__dict__.pop("shapes", None)  # drop python-pptx's cached shape proxy for the old tree
            Notes = kSlideFingerprint.NotesText(Source)
            if Source.has_notes_slide and Notes:
                New.notes_slide.notes_text_frame.text = Notes
            if Position is not None:
                kSlideCopier.MoveSlide(Deck, len(Deck.slides) - 1, Position)
            return New
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideCopier.CopySlide")
            return None

    @staticmethod
    def MoveSlide(Deck, Old, New):
        """Move the slide at index Old to index New."""
        if kS.ErrorMode:
            return None
        try:
            List = Deck.slides._sldIdLst
            Item = list(List)[Old]
            List.remove(Item)
            List.insert(New, Item)
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideCopier.MoveSlide")
            return None

    @staticmethod
    def DeleteSlide(Deck, Index):
        """Remove the slide at Index and renumber the slide parts."""
        if kS.ErrorMode:
            return None
        try:
            List = Deck.slides._sldIdLst
            Item = list(List)[Index]
            Deck.part.drop_rel(Item.get(f"{{{R}}}id"))
            List.remove(Item)
            # renumber slide parts, or the next add_slide reuses slideN.xml and the zip gets duplicate entries
            Deck.part.rename_slide_parts([Sld.get(f"{{{R}}}id") for Sld in List])
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideCopier.DeleteSlide")
            return None


class kHarvestEdits:
    """The three commands: manifest, harvest, restore."""

    @staticmethod
    def RequireFile(Path):
        """Expected state: an input file that does not exist."""
        if kS.ErrorMode:
            return None
        try:
            if not os.path.isfile(Path):
                raise ToolReportableException(f"not found: {Path}")
            return None
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kHarvestEdits.RequireFile")
            return None

    @staticmethod
    def Manifest(Deck, Out=None):
        """Write <deck>.manifest.json (or Out). Returns its path."""
        if kS.ErrorMode:
            return None
        try:
            kHarvestEdits.RequireFile(Deck)
            Out = Out or os.path.splitext(Deck)[0] + ".manifest.json"
            Manifest = kSlideFingerprint.DeckManifest(Deck)
            if Manifest is None:
                return None
            with open(Out, "w", encoding="utf-8") as File:
                json.dump(Manifest, File, indent=2)
            print(f"manifest: {Out}")
            return Out
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kHarvestEdits.Manifest(file={Deck})")
            return None

    @staticmethod
    def Harvest(Deck, ManifestPath, Store):
        """Store every slide changed, added or deleted since the manifest."""
        if kS.ErrorMode:
            return None
        try:
            kHarvestEdits.RequireFile(Deck)
            kHarvestEdits.RequireFile(ManifestPath)
            with open(ManifestPath, encoding="utf-8") as File:
                BaseSlides = json.load(File)["slides"]
            Base = {Slide["name"]: Slide["fingerprint"] for Slide in BaseSlides}
            BaseOrder = [Slide["name"] for Slide in BaseSlides]
            Current = Presentation(Deck)
            os.makedirs(Store, exist_ok=True)
            Keep = Presentation(Deck)  # the store is the deck minus unchanged slides: layouts and media come along
            Record, Seen, Previous = [], set(), None
            KeepIndexes = []
            for Number, Slide in enumerate(Current.slides, 1):
                Name = kSlideFingerprint.Name(Slide, Number)
                Duplicate = Name in Seen
                Seen.add(Name)
                if Duplicate or Name not in Base:
                    Status = "added"
                    Name = f"{Name}~{Number}" if Duplicate else Name
                elif kSlideFingerprint.Fingerprint(Slide) != Base[Name]:
                    Status = "edited"
                else:
                    Status = None
                if Status == "added" and Name.startswith("#"):
                    Name = f"added-{Number}"  # a slide made in PowerPoint has no id yet: give it a permanent one
                if Status:
                    Record.append({"name": Name, "status": Status, "after": Previous, "store_index": len(KeepIndexes)})
                    KeepIndexes.append(Number - 1)
                    if Status == "added":  # store the id with the slide so the next manifest knows it
                        Keep.slides[Number - 1]._element.find("p:cSld", NS).set("name", Name)
                Previous = Name
            for Name in BaseOrder:
                if Name not in Seen:
                    Record.append({"name": Name, "status": "deleted"})
            for Index in reversed(range(len(Keep.slides))):
                if Index not in KeepIndexes:
                    kSlideCopier.DeleteSlide(Keep, Index)
            if kS.ErrorMode:
                return None
            Keep.save(os.path.join(Store, "edits.pptx"))
            with open(os.path.join(Store, "edits.json"), "w", encoding="utf-8") as File:
                json.dump({"source": os.path.basename(Deck), "slides": Record}, File, indent=2)
            for Entry in Record:
                print(f"{Entry['status']:<8} {Entry['name']}")
            print(f"{sum(1 for Entry in Record if Entry['status'] != 'deleted')} slide(s) harvested, "
                  f"{sum(1 for Entry in Record if Entry['status'] == 'deleted')} deletion(s) recorded -> {Store}")
            return Record
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kHarvestEdits.Harvest(file={Deck})")
            return None

    @staticmethod
    def Restore(NewDeck, Store, Out):
        """Put the stored edits into a rebuilt deck and save it as Out (+ its manifest)."""
        if kS.ErrorMode:
            return None
        try:
            if os.path.abspath(Out) == os.path.abspath(NewDeck):
                raise ToolReportableException("--out must differ from the rebuilt deck")
            kHarvestEdits.RequireFile(NewDeck)
            kHarvestEdits.RequireFile(os.path.join(Store, "edits.json"))
            kHarvestEdits.RequireFile(os.path.join(Store, "edits.pptx"))
            with open(os.path.join(Store, "edits.json"), encoding="utf-8") as File:
                Meta = json.load(File)["slides"]
            Edits = Presentation(os.path.join(Store, "edits.pptx"))
            Deck = Presentation(NewDeck)
            for Entry in Meta:
                if Entry["status"] == "deleted" and Entry["name"] in kSlideFingerprint.Names(Deck):
                    kSlideCopier.DeleteSlide(Deck, kSlideFingerprint.Names(Deck).index(Entry["name"]))
                    print(f"kept deleted   {Entry['name']}")
                elif Entry["status"] == "deleted":
                    print(f"(gone anyway)  {Entry['name']}")
            for Entry in Meta:
                if Entry["status"] == "edited":
                    Source = Edits.slides[Entry["store_index"]]
                    if Entry["name"] in kSlideFingerprint.Names(Deck):
                        Position = kSlideFingerprint.Names(Deck).index(Entry["name"])
                        kSlideCopier.DeleteSlide(Deck, Position)
                        kSlideCopier.CopySlide(Source, Deck, Position)
                        print(f"kept edit      {Entry['name']}")
                    else:
                        print(f"!! not in the new build, re-added at the end: {Entry['name']}")
                        kSlideCopier.CopySlide(Source, Deck)
            for Entry in Meta:
                if Entry["status"] == "added":
                    Source = Edits.slides[Entry["store_index"]]
                    Names = kSlideFingerprint.Names(Deck)
                    Position = Names.index(Entry["after"]) + 1 if Entry["after"] in Names else len(Deck.slides)
                    kSlideCopier.CopySlide(Source, Deck, Position)
                    print(f"kept added     {Entry['name']} (after {Entry['after'] or 'start'})")
            if kS.ErrorMode:
                return None
            Deck.save(Out)
            kHarvestEdits.Manifest(Out)
            print(f"-> {Out}")
            return Out
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kHarvestEdits.Restore(file={NewDeck})")
            return None


class kHarvestEditsApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
            Commands = Parser.add_subparsers(dest="cmd", required=True)
            Manifest = Commands.add_parser("manifest", help="record a freshly built deck")
            Manifest.add_argument("deck")
            Manifest.add_argument("--out")
            Harvest = Commands.add_parser("harvest", help="collect slides changed since the manifest")
            Harvest.add_argument("deck")
            Harvest.add_argument("--manifest", help="default: <deck>.manifest.json")
            Harvest.add_argument("--store", required=True)
            Restore = Commands.add_parser("restore", help="put harvested edits into a rebuilt deck")
            Restore.add_argument("deck")
            Restore.add_argument("--store", required=True)
            Restore.add_argument("--out", required=True)
            Args = Parser.parse_args()
            if Args.cmd == "manifest":
                kHarvestEdits.Manifest(Args.deck, Args.out)
            elif Args.cmd == "harvest":
                kHarvestEdits.Harvest(Args.deck, Args.manifest or os.path.splitext(Args.deck)[0] + ".manifest.json",
                                      Args.store)
            else:
                kHarvestEdits.Restore(Args.deck, Args.store, Args.out)
            return 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kHarvestEditsApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kHarvestEditsApp)
