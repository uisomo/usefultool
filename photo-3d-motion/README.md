# Photo → 3D Motion Graphics

Turn a single photo into a pseudo-3D scene, place **text / images / shapes inside it**
(with correct perspective and occlusion — text can float *behind* people), animate the
camera, and export a video. Everything runs locally in your browser; no server, no API keys.

```
Photo → Depth Anything V2 (in-browser) → depth map → displaced Three.js mesh
      → 3D layers (text / PNG / shapes) → camera animation → WebM/MP4 export
```

## Run it

The app is two static files, but browsers block ES modules over `file://`, so serve the
folder with any static server:

```bash
cd photo-3d-motion
python3 -m http.server 8000
# then open http://localhost:8000
```

(or `npx serve`, VS Code Live Server, etc.)

**Requirements:** a modern Chromium browser (Chrome/Edge) is best — it has WebGPU, so
depth estimation takes ~1–2 s. Other browsers fall back to WASM (slower but works).
Three.js is vendored in `vendor/`, so the editor itself works offline; internet access
is only needed the first time you use **Generate Depth (AI)** to fetch the model
(~50 MB, cached by the browser afterwards).

## Workflow

1. **Load Photo** (or drag & drop one onto the viewport).
2. **Generate Depth (AI)** — runs Depth Anything V2 (small) locally via transformers.js.
   - Alternatively **Upload Depth Map** if you have one from another tool
     (white = near, black = far).
   - **View Depth** toggles a preview of the depth map on the mesh.
3. Tune **3D Depth Strength** and **Depth Smoothing** in the Scene section.
4. Add layers:
   - **+ Text** — canvas-rendered text with font, weight, fill and outline colors
     (works with Japanese text and emoji).
   - **+ Image** — any PNG/JPG (transparent PNGs work great for logos/stickers).
   - **+ Shape** — rounded rectangle, circle, arrow, line, star.
5. Position layers:
   - Drag them directly in the viewport.
   - The **Depth (Z)** slider moves a layer into/out of the scene. Because the photo is
     a real displaced mesh, a layer placed at a Z behind the foreground subject is
     automatically hidden by it — that's the "text behind the person" effect.
   - **Snap onto surface** places the layer just in front of whatever part of the photo
     is behind it (arrows pinned to walls, labels on objects).
   - **Push behind subject** places it just *behind* that surface.
   - Give each layer an entrance animation (fade / rise / pop / slide) and an
     appear time on the timeline.
6. Pick a **Camera Move** (orbit, dolly-in, 3D Ken Burns, vertigo/dolly-zoom, floaty
   handheld) plus motion amount and duration. Scrub the timeline or press Play.
   You can also freely orbit the preview camera by dragging empty space.
7. **Export Video** — records one full playthrough via `MediaRecorder`.
   Chrome usually saves WebM (VP9); Safari saves MP4. Convert WebM → MP4 with:

   ```bash
   ffmpeg -i photo-3d-motion.webm -c:v libx264 -pix_fmt yuv420p out.mp4
   ```

8. **Save / Open Project** stores everything (photo, depth map, layers, camera settings)
   in a single JSON file so you can continue later.

## How it works

- **Depth estimation** — [`onnx-community/depth-anything-v2-small`](https://huggingface.co/onnx-community/depth-anything-v2-small)
  through [transformers.js](https://github.com/huggingface/transformers.js), WebGPU when
  available, WASM otherwise. The model outputs inverse depth (bright = near).
- **3D scene** — Three.js `PlaneGeometry` (256×256 segments) with the depth map as a
  `displacementMap`. The material uses the emissive-map trick
  (`color: black, emissive: white, emissiveMap: photo`) so the photo renders unlit but
  still supports displacement.
- **Occlusion** — layers are ordinary meshes in the same scene, so the GPU depth buffer
  handles hiding them behind displaced foreground geometry. No masking or matting needed.
- **Text/shapes** — rendered to 2D canvases and used as textures on transparent planes.
  This supports any system font, CJK text, and emoji with zero extra dependencies.
- **Camera moves** — parametric pose functions over normalized time, including a true
  dolly-zoom (FOV compensated against distance).
- **Export** — `canvas.captureStream(30)` + `MediaRecorder` at ~14 Mbps.

## Known limitations

- Strong camera moves reveal "stretching" at depth edges (the mesh has no inpainting
  behind the foreground). Lower the motion amount / depth strength, or increase depth
  smoothing. Adding layer-based inpainting (e.g. LaMa) is the natural next upgrade.
- Text and shapes are flat cards placed in 3D (Canva/CapCut style), not extruded meshes.
- Export duration is recorded in real time (a 6 s video takes 6 s to record).
