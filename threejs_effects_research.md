# three.js Visual Effects Research — TouchDesigner-like Idea Visualizer

Goal: a browser-based app where **live speech → keywords → visual presentation**, with a
"TouchDesigner-ish" atmosphere (glow, trails, particles, audio-reactive, procedural motion).

All example names below were verified against the live three.js `dev` source
(`examples/files.json`). View any of them at:

```
https://threejs.org/examples/#<example_name>
# e.g. https://threejs.org/examples/#webgl_postprocessing_unreal_bloom
```

> **Docs ≠ Examples.** `threejs.org/docs` is the API reference (how a class works).
> `threejs.org/examples` is the visual gallery (runnable demos). For the look you want,
> browse **examples**; use docs to learn the class once a demo hooks you.
>
> **Important caveat:** examples are *code demos*, not click-to-apply presets. They show
> what is possible and give copy-able source. For an actual "effect menu," use the
> `postprocessing` library (below) and treat these examples as the inspiration board.

---

## 1. Postprocessing / cinematic effects (the core of the "TD look")

These run as a full-screen pass after the scene renders (via `EffectComposer`). This is
where 80% of the "atmosphere" comes from.

| Example | What it does | Relevance to your app |
|---|---|---|
| `webgl_postprocessing_unreal_bloom` | Soft HDR glow on bright pixels | **Top priority.** The single biggest "TD glow" win. |
| `webgl_postprocessing_unreal_bloom_selective` | Bloom only on tagged objects | Glow keyword text but not the background. |
| `webgl_postprocessing_afterimage` | Motion trails / ghosting | **High value.** Makes moving keywords feel alive/fluid. |
| `webgl_postprocessing_glitch` | Digital glitch / datamosh | Punctuate transitions, "signal" moments. |
| `webgl_postprocessing_godrays` | Volumetric light shafts | Cinematic depth behind floating text. |
| `webgl_postprocessing_rgb_halftone` | Comic/print halftone dots | Stylized graphic-design look. |
| `webgl_postprocessing_pixel` | Pixelation / mosaic | Lo-fi / retro mode. |
| `webgl_postprocessing_outline` | Edge outline on selection | Highlight the *current* keyword. |
| `webgl_postprocessing_sobel` | Edge-detection (wireframe-ish) | Abstract "data" rendering. |
| `webgl_postprocessing_ssao` (also `gtao`, `sao`) | Ambient-occlusion contact shadows | Grounds 3D depth (subtle, optional). |
| `webgl_postprocessing_transition` | Cross-fade between two scenes | **Useful** for moving between keyword "slides." |
| `webgl_postprocessing_3dlut` | LUT color grading | One-knob mood/palette control — strong cheap win. |
| `webgl_postprocessing_dof` (`dof2`) | Depth of field / bokeh | Focus on foreground keyword, blur the rest. |
| `webgl_postprocessing_procedural` | Generate textures via shader | Procedural backgrounds with no asset files. |

Also present (quality/AA, lower glamour but useful): `fxaa`, `smaa`, `ssaa`, `taa`,
`ssr` (screen-space reflections), `masking`, `advanced`, `backgrounds`.

---

## 2. Particles / motion / "data-feeling" effects

Closest to "live data becomes motion" — the heart of an idea-visualizer.

| Example | What it does | Relevance |
|---|---|---|
| `webgl_points_waves` | Grid of points rippling in waves | Calm ambient background that can react to audio amplitude. |
| `webgl_points_sprites` / `webgl_points_billboards` | Textured sprite point clouds | Glowing dots / soft particles. |
| `webgl_points_dynamic` | Points updated per-frame on CPU | Map keywords → spawn/move particles live. |
| `webgl_buffergeometry_custom_attributes_particles` | Per-particle size/color via shader attributes | **Key technique** for thousands of independently-styled particles. |
| `webgl_gpgpu_birds` | Flocking/boids on the GPU | **Showcase effect.** Organic swarm; map keywords → flock behavior. |
| `webgl_gpgpu_water` | GPU height-field water sim | Reactive liquid surface. |
| `webgl_gpgpu_protoplanet` | GPU N-body particle gravity | Particles clustering into "ideas." |

GPGPU = compute simulated on the GPU (positions stored in textures). Scales to
100k+ particles at 60fps, which CPU loops cannot.

---

## 3. Shader / procedural effects (the "alive surface" feel)

| Example | What it does | Relevance |
|---|---|---|
| `webgl_shader` | Minimal custom fragment shader | The starting template for any custom look. |
| `webgl_shader_lava` | Animated noise-distorted lava | Flowing energetic background. |
| `webgl_shaders_ocean` | Realistic animated ocean | Atmospheric reflective surface. |
| `webgl_shaders_sky` | Physical sky / atmosphere | Gradient sky / mood lighting. |
| `webgl_volume_cloud` | Raymarched volumetric clouds | Dreamy depth (heavier GPU cost). |
| `webgl_volume_perlin` | 3D Perlin-noise volume | Procedural fog/nebula — very "TD." |
| `webgl_volume_instancing` | Many volumes via instancing | Scaled volumetric fields. |

---

## 4. Text / presentation / audio-reactive

Directly serves **speech → keywords → on-screen words**.

| Example | What it does | Relevance |
|---|---|---|
| `webgl_geometry_text` | Extruded 3D text from a font | Render recognized keywords as 3D words. |
| `webgl_geometry_text_shapes` | Filled text as 2D shapes | Flat graphic text. |
| `webgl_geometry_text_stroke` | Outlined/stroked text | Neon-outline keyword style. |
| `css3d_periodictable` | DOM elements in 3D space | **Great pattern** for a grid/wall of keyword cards. |
| `css3d_sprites` | DOM sprites in 3D | Crisp HTML text labels in 3D (sharper than mesh text). |
| `webaudio_visualizer` | FFT spectrum → geometry | **Core** for audio-reactive visuals driven by the mic. |

> For high-quality dynamic text consider **troika-three-text** (SDF text, instant updates,
> no font pre-baking) — better than `TextGeometry` when words change constantly from speech.

---

## 5. Recommended first browsing order

1. `webgl_postprocessing_unreal_bloom` — confirm the glow is what you want
2. `webgl_postprocessing_afterimage` — trails
3. `webgl_postprocessing_glitch` — accents/transitions
4. `webgl_points_waves` — ambient reactive field
5. `webgl_gpgpu_birds` — the "wow" swarm
6. `webgl_shader_lava` — flowing procedural energy
7. `webgl_volume_perlin` — nebula/fog atmosphere
8. `webaudio_visualizer` — audio reactivity end-to-end

If these eight feel right, the browser version can absolutely hit a TouchDesigner-ish mood.

---

## 6. Don't hand-wire passes — use a library

three.js examples wire `EffectComposer` + individual passes manually. For an app, prefer:

- **`postprocessing`** (vanilla three.js) — npm `postprocessing` by pmndrs. Merges effects
  into fewer passes (faster), named effects: `BloomEffect`, `GlitchEffect`, `GodRaysEffect`,
  `PixelationEffect`, `OutlineEffect`, `DepthOfFieldEffect`, `LUT3DEffect`, etc.
- **`@react-three/postprocessing`** — same, declarative, if you use React Three Fiber
  (`@react-three/fiber`). Recommended if you want a component/"effect menu" structure.

Pattern: build the scene + particles in three.js / R3F, then stack named effects from the
`postprocessing` library, using the official examples above only as a reference for the math.

---

## 7. Suggested architecture for the idea-visualizer

```
Mic ─► Web Speech API (or Whisper) ─► transcript
        │
        ├─► keyword/NLP extraction ─► state store (current + recent keywords)
        │                                   │
        │                                   ├─► troika-three-text  (the words)
        │                                   ├─► GPGPU particles    (one cloud per keyword)
        │                                   └─► css3d cards        (keyword "wall")
        │
        └─► AnalyserNode (FFT)  ─► audio-reactive uniforms (amplitude/beat)
                                          │
                                          ▼
   three.js / R3F scene ─► postprocessing stack (Bloom → Afterimage → LUT → Glitch on transition)
```

Build order to de-risk: (1) bloom + a particle field reacting to mic FFT, (2) feed in
real speech keywords as troika text, (3) layer transitions/glitch, (4) tune with a 3D LUT.

---

## 8. Reality check vs TouchDesigner

- **Achievable in browser:** bloom, trails, glitch, godrays, LUT grading, DOF, GPGPU
  particles/flocking, procedural/volumetric shaders, FFT audio reactivity. This covers
  most of the recognizable "TD atmosphere."
- **Harder in browser:** TD's node-graph live-patching workflow, heavy volumetrics at high
  res, and some compositing — possible but more hand-coding. The *visual output* is reachable;
  the *authoring ergonomics* of TD are not replicated by three.js itself.

---

### Sources
- three.js examples gallery — https://threejs.org/examples/
- example names verified against `dev` `examples/files.json` —
  https://raw.githubusercontent.com/mrdoob/three.js/dev/examples/files.json
- unreal bloom example — https://threejs.org/examples/webgl_postprocessing_unreal_bloom.html
- godrays example — https://threejs.org/examples/webgl_postprocessing_godrays.html
- `postprocessing` library — https://github.com/pmndrs/postprocessing
- `@react-three/postprocessing` — https://github.com/pmndrs/react-postprocessing
- troika-three-text — https://github.com/protectwise/troika
