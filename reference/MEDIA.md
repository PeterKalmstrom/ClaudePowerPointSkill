# Images, video and media

*Runs on: generation: any OS; embedding: Windows + PowerPoint.* Part of the `building-powerpoint-decks` skill — start at [`SKILL.md`](../SKILL.md). Paths like `scripts/…` are relative to the skill folder.

## Nanobanana Integration (AI Images)

Nanobanana is an MCP server for AI image generation powered by Gemini models. It produces **raster images only** (PNG/JPG) — no vectors, no animation.

### Wrapping Python helpers via subprocess + uvx

When a project script needs to generate images or videos (e.g., a slide builder that needs a backdrop first), **call a Python helper as a subprocess via `uvx`** rather than importing it (see [`scripts/`](../scripts/README.md) for the deck helpers). The helpers depend on `google-genai` / `pywin32`, which most projects do not want to add as a hard dependency. `uvx` provisions the dependency per call and tears it down — no virtualenv setup, no project-wide pollution.

**Single-shot pattern** (one image, one video):

```python
import subprocess, sys
HELPER = r"path\to\your\batch_image_gen.py"
sys.exit(subprocess.run([
    "uvx", "--with", "google-genai", "python", HELPER,
    "--prompt", "Cinematic photograph of ... no text, no logos.",
    "--out", r"C:\out\backdrop.png", "--aspect", "16:9",
]).returncode)
```

**Batch pattern** (many jobs from a JSON spec written to a temp file):

```python
import json, os, subprocess, sys, tempfile
HELPER = r"path\to\your\batch_image_gen.py"
spec = [{"prompt": "...", "out": r"C:\images\one.png"}, ...]
with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
    json.dump(spec, f, ensure_ascii=False); spec_path = f.name
try:
    rc = subprocess.run(["uvx", "--with", "google-genai", "python", HELPER,
                         "--spec", spec_path, "--quiet"]).returncode
finally:
    os.unlink(spec_path)
sys.exit(rc)
```

The same pattern works for any Veo / video helper (use `uvx --with google-genai`) and any future helper script. Pass `--quiet` when calling from a larger pipeline so per-job progress doesn't pollute your build script's output (errors still go to stderr).

### Capabilities

- Photo-realistic and stylized image generation
- Image editing and modification
- 4K resolution output
- Google Search grounding for reference accuracy
- Subject consistency across multiple generations

### When to use

- **Background images** for slides (dramatic, atmospheric visuals)
- **Illustrations** that would take too long to source manually
- **Concept art** for topics that need emotional impact
- **Before/after** comparison images

### When NOT to use

- Animatable graphics (use Remotion or native PowerPoint shapes instead)
- Diagrams with precise labels/arrows (text rendering is unreliable)
- Vector graphics or SVGs (not supported)
- Charts or data visualizations (use PowerPoint native charts)

### Generate images WITHOUT text — overlay text in PowerPoint

**Rule:** Always generate images and video without text. Overlay all text in PowerPoint using semi-transparent dark backing rectangles for readability.

**Why:**
- AI image generators (Nanobanana, Gemini) frequently misspell text, render it in the wrong language, or warp it
- Generated text is locked into the raster image — you can't fix typos, change language, or update numbers
- PowerPoint text overlays are easy to edit, translate, restyle, and reposition
- Text overlays render crisply at any scale; embedded text gets blurry when resized
- Same image asset can be reused with different captions across slides

**How:**
1. Prompt with explicit "NO TEXT, NO LABELS, NO WORDS, NO LETTERS, NO NUMBERS anywhere in the image"
2. Add the prompt's negative_prompt: `text, words, letters, numbers, labels, captions, watermark`
3. Use AutoShape (rectangle) with `Fill.Transparency = 0.4` as backing for the text
4. Add `TextBox` shapes positioned over the image with the actual labels in the target language

**Hard rule on retries:** if the generator inserts text after 2 reprompts, give up. Do not iterate a third time hoping for a clean image — accept the spurious text or switch to a different background. Reroll cost adds up fast and the rate of success usually doesn't improve. Safer fallback: regenerate as an abstract texture (no objects = no labels), or fetch a stock image.

**Example overlay pattern:**
```python
# Dark semi-transparent backing
bg = slide.Shapes.AddShape(1, x, y, w, h)  # 1 = msoShapeRectangle
bg.Fill.Solid()
bg.Fill.ForeColor.RGB = 0x000000
bg.Fill.Transparency = 0.4
bg.Line.Visible = False
bg.ZOrder(3)  # send backward (behind text)

# Text on top
txt = slide.Shapes.AddTextbox(1, x, y, w, h)
txt.TextFrame.TextRange.Text = "Caption text"
txt.TextFrame.TextRange.Font.Size = 16
txt.TextFrame.TextRange.Font.Bold = True
txt.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF  # white in BGR
```

### Workflow

1. **Search first** — check if Nanobanana can find relevant reference imagery
2. **Generate** with a detailed prompt describing the visual — explicitly request NO TEXT
3. **Upload** the generated image to the project folder
4. **Embed** in PowerPoint via COM → `AddPicture`
5. **Overlay text** using TextBox + transparent backing rectangle

### Image embedding in PowerPoint

```python
# COM (pywin32):
pic = slide.Shapes.AddPicture(
    r"C:\path\to\image.png",
    False,   # LinkToFile
    True,    # SaveWithDocument
    left, top, width, height
)
```

> **Gotcha:** See [Images get swallowed by content placeholders](SETUP.md#images-get-swallowed-by-content-placeholders) in SETUP.md if the image disappears into a placeholder.

### Swapping a picture that lives inside a Group

**Rule:** When the picture you want to replace is a child of a Group (e.g. a chart wrapped with a rounded-rectangle backdrop), you can't just `AddPicture` over it — and Group children don't have a clean "replace source" API. The pattern is **delete the group, rebuild both shapes from scratch, then regroup by name**.

```python
# 1. Capture the bounds before deleting
for sh in slide.Shapes:
    if sh.Name == "Group 1":
        gx, gy, gw, gh = sh.Left, sh.Top, sh.Width, sh.Height
        for i in range(1, sh.GroupItems.Count + 1):
            child = sh.GroupItems(i)
            if child.Name == "ChartOverlay":
                px, py, pw, ph = child.Left, child.Top, child.Width, child.Height
        sh.Delete()
        break

# 2. Rebuild backdrop + picture at the captured coords
backdrop = slide.Shapes.AddShape(msoShapeRoundedRectangle, gx, gy, gw, gh)
backdrop.Name = "Rounded Rectangle 31"
# ...style backdrop...
pic = slide.Shapes.AddPicture(new_png_path, 0, -1, px, py, pw, ph)
pic.Name = "ChartOverlay"

# 3. Regroup so the slide structure matches the original
slide.Shapes.Range([backdrop.Name, pic.Name]).Group().Name = "Group 1"
```

A regenerated PNG with the same filename won't update an *embedded* picture — `AddPicture(SaveWithDocument=True)` copies bytes into the deck at insertion time. To pick up the new chart you must replace the shape.

### Portable OUT paths in chart scripts

**Rule:** Chart scripts that emit PNGs alongside themselves should derive the output directory from `__file__`, not hardcode `C:\Users\<somebody>\...`. Hardcoded paths break the moment the project moves between machines, between users, or between drives (cloud-synced repos are common offenders).

```python
# BAD — fails on every other machine, often silently if the dir exists
OUT = r"C:\Users\alice\Dropbox\my-project\images"

# GOOD — chart drops next to the script, wherever the repo lives
import os
OUT = os.path.dirname(os.path.abspath(__file__))
# ...later...
plt.savefig(os.path.join(OUT, "amoc-chart.png"), ...)
```

Same principle applies to `DECK_PATH` resolution in build scripts — use `os.path.join(os.path.dirname(__file__), "..", "Deck.pptx")` over absolute paths.

---

## Remotion Integration (Animated Video)

Remotion is a React-based framework for creating programmatic video. Output is MP4 video that embeds directly into PowerPoint slides. This is the recommended tool when native PowerPoint animations are not polished enough.

### When to use Remotion over native PowerPoint animations

| Use case | Tool |
|---|---|
| Simple entrance effects (fade, fly, zoom) | PowerPoint native animation ([ANIMATION](ANIMATION.md)) |
| Sequential bullet reveals | PowerPoint native animation, by paragraph ([ANIMATION](ANIMATION.md)) |
| Complex diagrams with glowing effects, particles, curves | **Remotion** |
| Network graphs, cascade visualizations, feedback loops | **Remotion** |
| Cinematic transitions, atmospheric builds | **Remotion** |
| Data-driven animations (charts morphing, counters) | **Remotion** |

### Prerequisites

- **Node.js** v18+ — check with `node --version`
- **Google Chrome** installed (used as the rendering engine)

> **PATH issue:** Claude Code's bash shell may not have Node.js in PATH even when installed. Use the full path: `export PATH="/c/Program Files/nodejs:$PATH"` at the start of every bash command.

### Project setup (one-time per animation)

```bash
export PATH="/c/Program Files/nodejs:$PATH"
mkdir -p "path/to/project"
cd "path/to/project"
npm init -y
npm install remotion @remotion/cli @remotion/player react react-dom typescript @types/react @types/react-dom
```

### Required files

Every Remotion project needs three files minimum:

**`src/index.ts`** — Entry point:
```typescript
import { registerRoot } from "remotion";
import { RemotionRoot } from "./Root";
registerRoot(RemotionRoot);
```

**`src/Root.tsx`** — Composition registry:
```typescript
import { Composition } from "remotion";
import { MyAnimation } from "./MyAnimation";

export const RemotionRoot: React.FC = () => (
  <Composition
    id="MyAnimation"
    component={MyAnimation}
    durationInFrames={300}  // 10 seconds at 30fps
    fps={30}
    width={1920}
    height={1080}
  />
);
```

**`src/MyAnimation.tsx`** — The actual animation (React component using Remotion hooks).

### Key Remotion APIs

| API | Purpose |
|---|---|
| `useCurrentFrame()` | Current frame number (0-based) |
| `useVideoConfig()` | Get fps, width, height, duration |
| `interpolate(frame, inputRange, outputRange, options)` | Map frame to any value (position, opacity, scale) |
| `spring({ frame, fps, config })` | Physics-based spring animation |
| `<AbsoluteFill>` | Full-frame container |
| `<Sequence from={60}>` | Delay child content to start at frame 60 |

### Rendering to MP4

```bash
export PATH="/c/Program Files/nodejs:$PATH"
cd "path/to/project"
npx remotion render src/index.ts MyAnimation out/animation.mp4 \
  --browser-executable="/c/Program Files/Google/Chrome/Application/chrome.exe"
```

**Critical:** Always pass `--browser-executable` pointing to the local Chrome installation. Without it, Remotion downloads Chrome Headless Shell (~108 MB) which often fails or times out, especially on cloud-synced folders (Dropbox, OneDrive).

### Rendering performance

- 300 frames (10s at 30fps) renders in ~90 seconds
- Remotion uses 8x concurrency by default
- Output size is typically 2–3 MB for motion graphics

### Animation design patterns

**Hub-and-spoke with growing center:**
```
Center node appears → Group 1 nodes + arrows fade in → arrow returns to center →
center grows → Group 2 appears → center grows more → ... → intensifying finale
```

**Progressive reveal:**
```
Title fades in → elements appear one by one with staggered timing →
connections draw between them → final state holds
```

**Key techniques for impact:**
- `spring()` for organic, bouncy entrances — feels alive
- Glowing SVG filters (`feGaussianBlur` + `feMerge`) for dramatic nodes
- Flowing particles along paths (quadratic bezier interpolation)
- `radial-gradient` backgrounds with vignette for atmosphere
- Vignette overlay that intensifies toward the end for urgency
- Dynamic sizing (nodes/elements that grow over the animation to show escalation)

### Iteration workflow

Remotion renders are fast enough to iterate:

1. Write/modify the `.tsx` animation code
2. Render to MP4 (~90s)
3. Replace the video on the slide via COM
4. Save the presentation
5. Preview in PowerPoint presentation mode (F5)
6. Repeat

---

## Veo Integration (AI Video Generation)

Google Veo 3.1 generates photorealistic video from text prompts via the Gemini API. This is the missing piece that enables true "prompt to video in a slide" — no stock footage, no manual animation, just describe what you want and get an MP4.

### Prerequisites

- **Google Gemini API key** — same key used by Nanobanana. Get one at [AI Studio](https://aistudio.google.com).
- **google-genai Python SDK** — install via `pip install google-genai` or run via `uvx --with google-genai`.

### Generating video

```python
import time, os
from google import genai

client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])

operation = client.models.generate_videos(
    model="veo-3.1-generate-preview",
    prompt="Your video description here. Be specific about subject, action, setting, lighting, camera angle.",
)

while not operation.done:
    time.sleep(15)
    operation = client.operations.get(operation)

video = operation.response.generated_videos[0]
client.files.download(file=video.video)
video.video.save("output.mp4")
```

### Running via uvx (no virtual environment needed)

If Python/pip is not directly available, use uvx (already installed for Nanobanana):

```bash
GEMINI_API_KEY="your-key" uvx --with google-genai python generate_video.py
```

The API key is in `~/.claude.json` under the Nanobanana MCP config's `env.GEMINI_API_KEY`.

### Key details

- **Model** (as of 2026-10): `veo-3.1-generate-preview` (best quality) and `veo-3.1-fast-generate-preview` (cheaper) are confirmed model ids; a `lite` id is **unverified** — list models before relying on it
- **Output**: 8-second MP4, 720p by default or 1080p on request; audio is always added (strip it, e.g. `ffmpeg -an`, if the slide should be silent), typically 5–15 MB
- **Generation time**: ~60 seconds including polling
- **Cost**: Pay-per-second via Gemini API billing. Use the fast model for drafts.

### Prompt tips

Be specific and cinematic:
- Describe subject, action, setting, lighting, camera movement
- "Cinematic, photorealistic" improves quality
- Specify direction of motion: "walks from left to right"
- **Keep prompts to 3–5 sentences.** Long rule lists with numbered "must" requirements actually degrade results — Veo trains on natural cinematic descriptions, not specifications. Anything over ~150 words confuses the model.

### Settings and prompt lessons that held up in production

- **1080p, muted.** `veo-3.1-fast-generate-preview` at 1080p gives 1920×1080, 24 fps, 8 s, H.264, ~7–15 MB. The
  Developer API always adds audio (`generate_audio` is rejected), so strip it afterwards:
  `ffmpeg -i in.mp4 -an -c:v copy out.mp4`.
- **Colours as words, never hex.** A hex code in the prompt (`#0078D4`) came back printed on the object.
- **Fake lettering is the default.** Machines and screens grow made-up words even with "no text" in the prompt.
  Add a negative prompt (`"text, letters, logos, captions"`) *and* say what the surface looks like: "a completely
  blank, unmarked surface" fixed about 3 of 5 retries. The standard model was cleaner than fast on the hardest clips.
- **Characters drift.** Keep a *character sheet* — one fixed description per recurring character — and paste it
  identically into every prompt. Plan retries anyway: diffusion adds extra characters or swaps one for another.
- **Content filter.** A blocked clip returns no video (`op.response` is `None`) rather than an error. Check for it,
  reword the action ("taps the glass" was blocked; "presses a button" passed) and retry.
- **API keys:** if both `GOOGLE_API_KEY` and `GEMINI_API_KEY` are set, the SDK uses `GOOGLE_API_KEY`.
- **Folders:** keep Veo originals in `videos/master/` (superseded takes in `v1/`, `v2/`) and compressed copies in
  `videos/final/`. Only `final/` goes into the deck.

### Known limits — when not to use Veo

Diffusion video models reliably fail at certain physics scenarios. Don't waste reprompts on these — go straight to a real-footage clip:

- **Sequential rigid-body chain reactions** (dominoes falling in sequence, pool break, Newton's cradle) — Veo produces plausible-looking but physically wrong cascades. Direction and contact transfer are unreliable.
- **Counting** (exactly N objects, ordered sequences) — frequently miscounts.
- **Text in any language** — see image rule above; same applies to video.
- **Complex multi-step "this then that then that"** beyond ~3 actions.

For these cases: search for stock footage (Pexels free tier, or `yt-dlp` on a Creative Commons clip) and trim with `ffmpeg`. Faster and better than reprompting.

### Image-to-video — when direction matters

When the result depends on a specific starting state (orientation, subject position, scene composition), generate a still image first via the image API, then feed it as the **starting frame** of the video. Anchoring the first frame dramatically reduces variance in direction, framing, and consistency.

### Workflow

1. Write a Python script with the prompt (or inline it)
2. Run via uvx → polls until video is ready (~60s)
3. Embed the MP4 in PowerPoint via COM → `AddMediaObject2`
4. Done — from text prompt to video playing in a slide

---

## Embedding Media in Slides

### Embedding video (Remotion output)

**Compress first — PowerPoint wants H.264, and embedded video is most of a deck's size.** Veo and Remotion output
runs at ~7 Mbps. Re-encode before embedding:

```bash
ffmpeg -i in.mp4 -c:v libx264 -preset slow -crf 23 -pix_fmt yuv420p -movflags +faststart -an out.mp4
```

That took 20 clips from 202 MB to 122 MB (~40 %) with no visible loss. Measured on one 7.1 MB clip (SSIM vs source,
1.0 = identical):

| Setting | Size | SSIM |
|---|---|---|
| x264 crf 20 | 7.0 MB | .990 |
| x264 crf 22 | 5.2 MB | .987 |
| x264 crf 24 | 3.9 MB | .979 |
| x264 crf 28 | 2.0 MB | .969 |
| x265 crf 26 | 2.8 MB | .975 |
| AV1 crf 32 | 2.4 MB | .979 |
| VP9 crf 34 | 3.3 MB | .949 |

x265, AV1 and VP9 are smaller but not safe in every PowerPoint install; stay on x264 (`yuv420p`). Drop `-an` if the
clip needs its sound.

```python
# COM (pywin32):
s = presentation.Slides(slide_number)

# Remove old video if replacing
for i in range(s.Shapes.Count, 0, -1):
    if s.Shapes(i).Name == "VideoName":
        s.Shapes(i).Delete()

# Add new video — full slide coverage
video = s.Shapes.AddMediaObject2(
    r"C:\path\to\animation.mp4",
    False,  # LinkToFile — False embeds the video in the .pptx
    True,   # SaveWithDocument
    0, 0,   # Left, Top
    presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight   # full slide, whatever the size
)
video.Name = "VideoName"
video.AnimationSettings.PlaySettings.PlayOnEntry = True
video.AnimationSettings.PlaySettings.HideWhileNotPlaying = False
```

### Embedding images (Nanobanana output)

```python
# COM (pywin32):
pic = slide.Shapes.AddPicture(
    r"C:\path\to\image.png",
    False, True,
    left, top, width, height
)
pic.Name = "ImageName"
```

### Hybrid slides (image background + PowerPoint overlays)

Use Nanobanana for a dramatic background image, then add PowerPoint shapes on top for text/labels that need to be editable:

```python
# Background image
bg = slide.Shapes.AddPicture(r"path\to\bg.png", False, True, 0, 0,
                              presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight)
bg.ZOrder(1)  # Send to back

# Overlay shapes on top
sw = presentation.PageSetup.SlideWidth                     # 1440 on Full HD
title = slide.Shapes.AddTextbox(1, 80, 56, sw - 160, 100)   # Full HD margins: 80 pt sides, title at 56 pt
title.TextFrame.TextRange.Text = "Title Over Image"
title.TextFrame.TextRange.Font.Color.RGB = 16777215  # White - check contrast on the busiest part of the photo
```

---

## Combined Workflow Patterns

### Pattern 1: "High-impact animated visualization"

Best for: systemic risks, network effects, process cascades, escalating trends

1. **PowerPoint (COM)** → open deck, render existing slides for context
2. **Remotion** → build animated diagram (glowing nodes, flowing arrows, growing elements)
3. **PowerPoint (COM)** → embed MP4, add speaker notes, save

### Pattern 2: "Dramatic reveal"

Best for: before/after, impact stories, emotional content

1. **Nanobanana** → generate atmospheric background image
2. **PowerPoint (COM)** → embed image as background, add text overlays and animations
3. **PowerPoint (COM)** → entrance animations for progressive text reveals on top

### Pattern 3: "Data + narrative"

Best for: presentations mixing charts with storytelling

1. **PowerPoint (COM)** → create slides with native charts/shapes for data
2. **Nanobanana** → generate editorial/emotional images for transition slides
3. **Remotion** → animate the key "aha moment" visualization
4. **PowerPoint (COM)** → assemble everything, add speaker notes throughout

### Pattern 4: "Prompt to video in a slide"

Best for: illustrative scenes, metaphors, product demos, any "show don't tell" moment

1. **Veo 3.1** → generate video from a text description (~60 seconds)
2. **PowerPoint (COM)** → add slide, embed the MP4, save

This is the simplest and most powerful pattern. One prompt, one API call, one slide.

### Decision guide

```
Need animated diagram/visualization?
  ├─ Simple (3-5 shapes, basic entrances) → PowerPoint native animations
  └─ Complex (glowing, particles, growth, curves) → Remotion

Need a photorealistic video clip?
  ├─ Generic scene Veo can render → Veo 3.1 (~60 seconds)
  └─ Sequential physics (dominoes, collisions, chain reactions) → stock footage,
     not Veo (see "Known limits" above)

Need an image?
  ├─ Photo/illustration → image generator (Nanobanana / Gemini Imagen)
  ├─ Precise diagram with labels → PowerPoint native shapes
  └─ Background atmosphere → image generator

Need text overlays on media?
  └─ Always use PowerPoint shapes on top — never bake text into images/video
      (keeps it editable, avoids rendering artifacts)
```

---
