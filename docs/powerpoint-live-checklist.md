# PowerPoint Live panel - manual checklist (items 3.1-3.16)

For the owner, in **Claude Desktop on Windows** with desktop PowerPoint. These are the items a script cannot click;
the server behind them is already covered by `selftest.py --com` and the live smoke tests. About 20 minutes.

**Before you start**

- Close PowerPoint. Build the sample deck:
  `uvx --with python-pptx --with pillow python scripts/build_deck.py examples/spec/sample-deck.json --out out/sample-deck.pptx`
- In Claude Desktop ask: *"Open out/sample-deck.pptx with PowerPoint Live and show slide 2."*
- Tick each box. On a fail, note **item, slide, what you did, what you saw**, and a screenshot if it is visual.

| # | Do this | Expected | Pass |
|---|---|---|---|
| 3.1 | Ask Claude: *"Open the deck with powerpoint_open and show slide 2"* | Panel shows slide 2 within ~3 s; the PowerPoint window follows to the same slide | [ ] |
| 3.2 | Ask Claude to change the title on the current slide (powerpoint_run) | Changed shape flashes **blue** (added green, removed red-dashed); History gets a version | [ ] |
| 3.3 | Drag the **Before / after** slider across the slide | Left shows the slide before its last change, right the current one | [ ] |
| 3.4 | Open **Checks** (add a long paragraph to get a finding); press **Fix** on a fixable one | Lint boxes drawn on the slide and listed with counts; Fix changes the slide live in PowerPoint and the finding disappears | [ ] |
| 3.5 | **Ask Claude** (point and ask): click a shape, type a request, send | Message appears in the chat with slide number, SlideID, shape id and box (or goes to the clipboard, with a note saying so) | [ ] |
| 3.6 | Tick **Accurate** | Slide re-renders via Save As (slower); fonts and line breaks match PowerPoint exactly | [ ] |
| 3.7 | **Slide sorter**: drag a slide onto another; drag one into another section; multi-select drag | Order and sections change in PowerPoint too | [ ] |
| 3.8 | **Sorter**: hide/unhide a slide, new section, rename a section (double-click), delete a section (x) | Done in PowerPoint; deleting a section keeps its slides | [ ] |
| 3.9 | **Storyline** tab | Titles listed by section; title flags (label/too long/missing/duplicate); talk-length bar vs target | [ ] |
| 3.10 | **Rehearse** tab: Start, Next, Previous, Black (B), End | Drives the real slide show in PowerPoint; timers run; ahead / on time / behind chip changes | [ ] |
| 3.11 | **Design**: open the gallery, preview the 6 directions, apply one | Previews render without touching the deck; apply restyles every slide (colours + fonts); History has a version | [ ] |
| 3.12 | **Theme**: change Accent 1 and the heading font, **Apply** | Colours and fonts change in PowerPoint (check the colour is not channel-swapped, e.g. red shown as blue); contrast list updates | [ ] |
| 3.13 | Ask Claude for 2 layout variants of a slide (powerpoint_layout_variants), compare, **Use this** | Variants appear as hidden slides; the chosen one replaces the original in place (also when the slide starts a section) | [ ] |
| 3.14 | **History**: Restore an earlier version | Deck file replaced and reopened; the state you left appears as a new version (restore is undoable) | [ ] |
| 3.15 | Close the deck from PowerPoint while the panel is open; then reopen it and Save As under a new name | No halt; the panel says the deck is closed, then follows the new name | [ ] |
| 3.16 | Close PowerPoint completely while a design preview or an Accurate render is running (an error inside powerpoint_run code does not count - that is returned as the code's own result) | Either a plain message, or a halt with **Resume** and the question *"Do you want to send this error message?"* - answer **No** | [ ] |

## Report back

Paste this into the chat (or save it as `test-results/powerpoint/<date>-<pc>/live-panel.md`):

```text
PowerPoint Live panel - <date>, Claude Desktop <version>, PowerPoint <version>
Passed: 3.x, 3.y, ...
Failed: 3.z - slide N - did: ... - saw: ... - repeats: yes/no - screenshot: <file>
Not tested: 3.w - why
Anything odd (slowness, flicker, a dialog in PowerPoint): ...
```
