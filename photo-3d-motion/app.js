// Photo → 3D Motion Graphics
// Pipeline: photo → monocular depth (Depth Anything V2, in-browser) → displaced
// Three.js mesh → text/image/shape layers placed in 3D (with real occlusion via
// the depth buffer) → animated camera → MediaRecorder video export.

import * as THREE from 'three';
import { OrbitControls } from 'three/addons/controls/OrbitControls.js';

// ---------------------------------------------------------------------------
// Constants & state
// ---------------------------------------------------------------------------

const PLANE_H = 2;               // world height of the photo plane
const BASE_FOV = 50;
const CAM_Z = 2.35;              // base camera distance
const BASE_LAYER_H = { text: 0.32, image: 0.8, shape: 0.5 };
const ENTRANCE_DUR = 0.7;

const state = {
  photoDataURL: null,
  photoW: 0, photoH: 0,
  hasDepth: false,
  showDepth: false,
  depthScale: 1.0,
  depthBlur: 2,
  bg: '#05070a',
  camPreset: 'orbit',
  camAmount: 0.55,
  duration: 6,
  layers: [],
  selectedId: null,
  t: 0,
  playing: false,
  exporting: false,
};

let layerSeq = 1;

// ---------------------------------------------------------------------------
// Three.js setup
// ---------------------------------------------------------------------------

const stage = document.getElementById('stage');
const renderer = new THREE.WebGLRenderer({ antialias: true, preserveDrawingBuffer: true });
renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
stage.appendChild(renderer.domElement);

const scene = new THREE.Scene();
scene.background = new THREE.Color(state.bg);

const camera = new THREE.PerspectiveCamera(BASE_FOV, 1, 0.05, 50);
camera.position.set(0, 0, CAM_Z);

const controls = new OrbitControls(camera, renderer.domElement);
controls.target.set(0, 0, 0);
controls.enableDamping = false;
controls.minDistance = 0.4;
controls.maxDistance = 10;

new ResizeObserver(() => {
  const w = stage.clientWidth, h = stage.clientHeight;
  if (!w || !h) return;
  renderer.setSize(w, h);
  camera.aspect = w / h;
  camera.updateProjectionMatrix();
}).observe(stage);

// Photo mesh --------------------------------------------------------------

let photoMesh = null;
let photoTex = null;
let depthTex = null;
let depthCanvas = null;          // blurred working canvas (bound to depthTex)
let depthCanvasOriginal = null;  // unblurred source

function makeFlatDepthCanvas() {
  const c = document.createElement('canvas');
  c.width = c.height = 4;
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#808080';
  ctx.fillRect(0, 0, 4, 4);
  return c;
}

function rebuildDepthWorkingCanvas() {
  if (!depthCanvasOriginal) return;
  if (!depthCanvas) depthCanvas = document.createElement('canvas');
  depthCanvas.width = depthCanvasOriginal.width;
  depthCanvas.height = depthCanvasOriginal.height;
  const ctx = depthCanvas.getContext('2d');
  ctx.filter = state.depthBlur > 0 ? `blur(${state.depthBlur}px)` : 'none';
  ctx.drawImage(depthCanvasOriginal, 0, 0);
  if (depthTex) {
    depthTex.needsUpdate = true;
  }
}

function buildPhotoMesh(image) {
  if (photoMesh) {
    scene.remove(photoMesh);
    photoMesh.geometry.dispose();
    photoMesh.material.dispose();
  }
  photoTex = new THREE.Texture(image);
  photoTex.colorSpace = THREE.SRGBColorSpace;
  photoTex.needsUpdate = true;

  if (!depthCanvasOriginal) depthCanvasOriginal = makeFlatDepthCanvas();
  rebuildDepthWorkingCanvas();
  depthTex = new THREE.CanvasTexture(depthCanvas);

  const aspect = image.width / image.height;
  const geo = new THREE.PlaneGeometry(PLANE_H * aspect, PLANE_H, 256, 256);
  // MeshStandardMaterial supports displacementMap; emissive trick = unlit photo.
  const mat = new THREE.MeshStandardMaterial({
    color: 0x000000,
    emissive: 0xffffff,
    emissiveMap: photoTex,
    displacementMap: depthTex,
    displacementScale: state.depthScale,
    displacementBias: -state.depthScale * 0.45,
    roughness: 1,
  });
  photoMesh = new THREE.Mesh(geo, mat);
  scene.add(photoMesh);
  updateDepthView();
}

function updateDisplacement() {
  if (!photoMesh) return;
  photoMesh.material.displacementScale = state.depthScale;
  photoMesh.material.displacementBias = -state.depthScale * 0.45;
}

function updateDepthView() {
  if (!photoMesh) return;
  photoMesh.material.emissiveMap = state.showDepth ? depthTex : photoTex;
  photoMesh.material.needsUpdate = true;
}

// ---------------------------------------------------------------------------
// Depth estimation (Depth Anything V2 via transformers.js, runs in-browser)
// ---------------------------------------------------------------------------

let depthPipeline = null;

async function loadDepthPipeline(progressCb) {
  if (depthPipeline) return depthPipeline;
  const { pipeline } = await import(
    'https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.3.1'
  );
  const model = 'onnx-community/depth-anything-v2-small';
  const opts = { progress_callback: progressCb };
  try {
    if (!navigator.gpu) throw new Error('no webgpu');
    depthPipeline = await pipeline('depth-estimation', model, { ...opts, device: 'webgpu' });
  } catch (e) {
    console.warn('WebGPU unavailable, falling back to WASM:', e);
    depthPipeline = await pipeline('depth-estimation', model, opts);
  }
  return depthPipeline;
}

function rawDepthToCanvas(raw) {
  const { width, height, data, channels } = raw;
  const c = document.createElement('canvas');
  c.width = width; c.height = height;
  const ctx = c.getContext('2d');
  const id = ctx.createImageData(width, height);
  for (let i = 0; i < width * height; i++) {
    const v = data[i * channels];
    id.data[i * 4] = v; id.data[i * 4 + 1] = v; id.data[i * 4 + 2] = v;
    id.data[i * 4 + 3] = 255;
  }
  ctx.putImageData(id, 0, 0);
  return c;
}

async function generateDepth() {
  if (!state.photoDataURL) return;
  const btn = document.getElementById('btnGenDepth');
  btn.disabled = true;
  try {
    showStatus('Loading depth model (first run downloads ~50 MB, then cached)…', 0);
    const pipe = await loadDepthPipeline((p) => {
      if (p.status === 'progress' && p.total) {
        showStatus(`Downloading model: ${p.file}`, (p.loaded / p.total) * 100);
      }
    });
    showStatus('Estimating depth…', null);
    const out = await pipe(state.photoDataURL);
    depthCanvasOriginal = rawDepthToCanvas(out.depth);
    rebuildDepthWorkingCanvas();
    if (depthTex) depthTex.dispose();
    depthTex = new THREE.CanvasTexture(depthCanvas);
    if (photoMesh) {
      photoMesh.material.displacementMap = depthTex;
      photoMesh.material.needsUpdate = true;
    }
    state.hasDepth = true;
    document.getElementById('btnShowDepth').disabled = false;
    updateDepthView();
    hideStatus();
  } catch (e) {
    console.error(e);
    showStatus(
      'Depth model failed to load (offline? unsupported browser?). ' +
      'You can still use “Upload Depth Map” with an external depth image.', null);
    setTimeout(hideStatus, 8000);
  } finally {
    btn.disabled = false;
  }
}

// ---------------------------------------------------------------------------
// Layer textures (canvas-based: text, shapes, images)
// ---------------------------------------------------------------------------

function makeTextCanvas(l) {
  const fontPx = 128;
  const pad = fontPx * 0.3;
  const c = document.createElement('canvas');
  const ctx = c.getContext('2d');
  const font = `${l.weight} ${fontPx}px ${l.font}`;
  ctx.font = font;
  const lines = String(l.text || ' ').split('\\n');
  const w = Math.max(...lines.map((s) => ctx.measureText(s).width), fontPx * 0.5);
  c.width = Math.ceil(w + pad * 2);
  c.height = Math.ceil(fontPx * 1.25 * lines.length + pad * 2);
  ctx.font = font;
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';
  ctx.lineJoin = 'round';
  lines.forEach((line, i) => {
    const y = pad + fontPx * 1.25 * (i + 0.5);
    ctx.strokeStyle = l.outline;
    ctx.lineWidth = fontPx * 0.09;
    ctx.strokeText(line, c.width / 2, y);
    ctx.fillStyle = l.color;
    ctx.fillText(line, c.width / 2, y);
  });
  return c;
}

function makeShapeCanvas(l) {
  const S = 512;
  const c = document.createElement('canvas');
  c.width = S; c.height = S;
  const ctx = c.getContext('2d');
  ctx.fillStyle = l.color;
  ctx.strokeStyle = l.outline;
  ctx.lineWidth = S * 0.04;
  ctx.lineJoin = 'round';
  const m = S * 0.08;
  switch (l.shape) {
    case 'rect': {
      const r = S * 0.09;
      ctx.beginPath();
      ctx.roundRect(m, S * 0.28, S - m * 2, S * 0.44, r);
      ctx.fill(); ctx.stroke();
      break;
    }
    case 'circle':
      ctx.beginPath();
      ctx.arc(S / 2, S / 2, S / 2 - m, 0, Math.PI * 2);
      ctx.fill(); ctx.stroke();
      break;
    case 'arrow': {
      const y = S / 2, h = S * 0.11, head = S * 0.24;
      ctx.beginPath();
      ctx.moveTo(m, y - h);
      ctx.lineTo(S - m - head, y - h);
      ctx.lineTo(S - m - head, y - h * 2.2);
      ctx.lineTo(S - m, y);
      ctx.lineTo(S - m - head, y + h * 2.2);
      ctx.lineTo(S - m - head, y + h);
      ctx.lineTo(m, y + h);
      ctx.closePath();
      ctx.fill(); ctx.stroke();
      break;
    }
    case 'line':
      ctx.fillRect(m, S / 2 - S * 0.03, S - m * 2, S * 0.06);
      break;
    case 'star': {
      const cx = S / 2, cy = S / 2, R = S / 2 - m, r = R * 0.45;
      ctx.beginPath();
      for (let i = 0; i < 10; i++) {
        const rad = i % 2 === 0 ? R : r;
        const a = (i / 10) * Math.PI * 2 - Math.PI / 2;
        ctx[i === 0 ? 'moveTo' : 'lineTo'](cx + Math.cos(a) * rad, cy + Math.sin(a) * rad);
      }
      ctx.closePath();
      ctx.fill(); ctx.stroke();
      break;
    }
  }
  return c;
}

function rebuildLayerTexture(l) {
  const apply = (source, aspect) => {
    const tex = source instanceof HTMLCanvasElement
      ? new THREE.CanvasTexture(source)
      : new THREE.Texture(source);
    tex.colorSpace = THREE.SRGBColorSpace;
    tex.anisotropy = 4;
    tex.needsUpdate = true;
    if (l.mesh.material.map) l.mesh.material.map.dispose();
    l.mesh.material.map = tex;
    l.mesh.material.needsUpdate = true;
    l.aspect = aspect;
  };
  if (l.type === 'text') {
    const c = makeTextCanvas(l);
    apply(c, c.width / c.height);
  } else if (l.type === 'shape') {
    const c = makeShapeCanvas(l);
    apply(c, 1);
  } else if (l.type === 'image' && l.imageDataURL) {
    const img = new Image();
    img.onload = () => apply(img, img.width / img.height);
    img.src = l.imageDataURL;
  }
}

// ---------------------------------------------------------------------------
// Layer management
// ---------------------------------------------------------------------------

function createLayer(type, extra = {}) {
  const l = {
    id: layerSeq++,
    type,
    name: extra.name || (type === 'text' ? (extra.text || 'Text') : type),
    text: 'Hello!',
    font: 'sans-serif',
    weight: '700',
    shape: 'rect',
    color: '#ffffff',
    outline: '#000000',
    scale: 1,
    x: 0, y: 0, z: 0.9,
    rot: 0,
    opacity: 1,
    anim: 'fade',
    appear: 0,
    aspect: 1,
    imageDataURL: null,
    ...extra,
  };
  const geo = new THREE.PlaneGeometry(1, 1);
  const mat = new THREE.MeshBasicMaterial({
    transparent: true,
    depthWrite: false,
    side: THREE.DoubleSide,
  });
  l.mesh = new THREE.Mesh(geo, mat);
  l.mesh.userData.layerId = l.id;
  scene.add(l.mesh);
  rebuildLayerTexture(l);
  state.layers.push(l);
  selectLayer(l.id);
  renderLayerList();
  return l;
}

function deleteLayer(id) {
  const i = state.layers.findIndex((l) => l.id === id);
  if (i < 0) return;
  const l = state.layers[i];
  scene.remove(l.mesh);
  l.mesh.geometry.dispose();
  if (l.mesh.material.map) l.mesh.material.map.dispose();
  l.mesh.material.dispose();
  state.layers.splice(i, 1);
  if (state.selectedId === id) state.selectedId = null;
  renderLayerList();
  renderProps();
}

function duplicateLayer(id) {
  const l = state.layers.find((x) => x.id === id);
  if (!l) return;
  const { mesh, id: _id, ...rest } = l;
  createLayer(l.type, { ...rest, x: l.x + 0.15, y: l.y - 0.1, name: l.name + ' copy' });
}

function selectedLayer() {
  return state.layers.find((l) => l.id === state.selectedId) || null;
}

function selectLayer(id) {
  state.selectedId = id;
  renderLayerList();
  renderProps();
}

// ---------------------------------------------------------------------------
// Per-frame layer transforms (entrance animations)
// ---------------------------------------------------------------------------

const easeOutCubic = (p) => 1 - Math.pow(1 - p, 3);
const easeInOutSine = (p) => -(Math.cos(Math.PI * p) - 1) / 2;
const easeOutBack = (p) => {
  const c1 = 1.70158, c3 = c1 + 1;
  return 1 + c3 * Math.pow(p - 1, 3) + c1 * Math.pow(p - 1, 2);
};

function updateLayersAtTime(t) {
  for (const l of state.layers) {
    let alpha = 1, ox = 0, oy = 0, pop = 1;
    if (l.anim !== 'none' || t < l.appear) {
      const p = Math.min(Math.max((t - l.appear) / ENTRANCE_DUR, 0), 1);
      const e = easeOutCubic(p);
      if (t < l.appear) alpha = 0;
      else switch (l.anim) {
        case 'fade': alpha = e; break;
        case 'rise': alpha = e; oy = -(1 - e) * 0.3; break;
        case 'pop': alpha = Math.min(1, p * 3); pop = 0.3 + 0.7 * easeOutBack(p); break;
        case 'slide': alpha = e; ox = -(1 - e) * 0.8; break;
      }
    }
    const h = BASE_LAYER_H[l.type] * l.scale * pop;
    l.mesh.position.set(l.x + ox, l.y + oy, l.z);
    l.mesh.rotation.z = (l.rot * Math.PI) / 180;
    l.mesh.scale.set(h * l.aspect, h, 1);
    l.mesh.material.opacity = l.opacity * alpha;
    l.mesh.visible = l.mesh.material.opacity > 0.001;
  }
}

// ---------------------------------------------------------------------------
// Camera presets
// ---------------------------------------------------------------------------

function cameraPose(preset, u, m) {
  const p = { x: 0, y: 0, z: CAM_Z, fov: BASE_FOV, lx: 0, ly: 0, lz: 0 };
  const s = easeInOutSine(u);
  switch (preset) {
    case 'orbit': {
      const a = (s * 2 - 1) * 0.5 * m;
      p.x = Math.sin(a) * CAM_Z;
      p.z = Math.cos(a) * CAM_Z;
      p.y = 0.12 * m * Math.sin(u * Math.PI);
      p.lz = 0.2;
      break;
    }
    case 'dolly':
      p.z = CAM_Z - 1.1 * m * s;
      p.x = 0.15 * m * Math.sin(u * Math.PI);
      break;
    case 'kenburns':
      p.z = CAM_Z - 0.6 * m * u;
      p.x = (s * 2 - 1) * 0.25 * m;
      p.y = (s * 2 - 1) * -0.1 * m;
      p.lx = p.x * 0.5; p.ly = p.y * 0.5;
      break;
    case 'vertigo': {
      p.z = CAM_Z + 1.5 * m * s;
      const k = Math.tan((BASE_FOV / 2) * (Math.PI / 180)) * CAM_Z;
      p.fov = (2 * Math.atan(k / p.z)) * (180 / Math.PI);
      break;
    }
    case 'floaty':
      p.x = 0.14 * m * Math.sin(u * Math.PI * 4);
      p.y = 0.09 * m * Math.sin(u * Math.PI * 6 + 1.3);
      p.z = CAM_Z - 0.15 * m * Math.sin(u * Math.PI * 2);
      p.lx = 0.05 * m * Math.sin(u * Math.PI * 3);
      break;
    case 'still':
    default:
      break;
  }
  return p;
}

function applyCameraAtTime(t) {
  const u = state.duration > 0 ? Math.min(t / state.duration, 1) : 0;
  const p = cameraPose(state.camPreset, u, state.camAmount);
  camera.position.set(p.x, p.y, p.z);
  camera.fov = p.fov;
  camera.updateProjectionMatrix();
  camera.lookAt(p.lx, p.ly, p.lz);
}

// ---------------------------------------------------------------------------
// Playback / render loop
// ---------------------------------------------------------------------------

let playStartMs = 0;
let recorder = null;

function setTime(t, driveCamera = true) {
  state.t = Math.min(Math.max(t, 0), state.duration);
  document.getElementById('scrub').value = state.t;
  document.getElementById('timeDisplay').textContent =
    `${state.t.toFixed(1)} / ${state.duration.toFixed(1)}s`;
  if (driveCamera) applyCameraAtTime(state.t);
}

function startPlayback() {
  state.playing = true;
  controls.enabled = false;
  playStartMs = performance.now() - state.t * 1000;
  if (state.t >= state.duration - 0.01) playStartMs = performance.now();
  document.getElementById('btnPlay').textContent = '⏸ Pause';
}

function stopPlayback() {
  state.playing = false;
  controls.enabled = true;
  document.getElementById('btnPlay').textContent = '▶ Play';
  if (recorder && recorder.state === 'recording') recorder.stop();
}

function animate() {
  requestAnimationFrame(animate);
  if (state.playing) {
    const t = (performance.now() - playStartMs) / 1000;
    if (t >= state.duration) {
      setTime(state.duration);
      stopPlayback();
    } else {
      setTime(t);
    }
  }
  updateLayersAtTime(state.t);
  renderer.render(scene, camera);
}
animate();

// ---------------------------------------------------------------------------
// Video export
// ---------------------------------------------------------------------------

async function exportVideo() {
  if (!photoMesh || state.exporting) return;
  state.exporting = true;
  const btn = document.getElementById('btnExport');
  btn.disabled = true;
  btn.textContent = '● Recording…';

  const types = [
    'video/mp4;codecs=avc1.42E01E',
    'video/webm;codecs=vp9',
    'video/webm',
  ];
  const mime = types.find((t) => MediaRecorder.isTypeSupported(t)) || '';
  const ext = mime.startsWith('video/mp4') ? 'mp4' : 'webm';
  const stream = renderer.domElement.captureStream(30);
  const chunks = [];
  recorder = new MediaRecorder(stream, {
    mimeType: mime || undefined,
    videoBitsPerSecond: 14_000_000,
  });
  recorder.ondataavailable = (e) => e.data.size && chunks.push(e.data);
  recorder.onstop = () => {
    const blob = new Blob(chunks, { type: mime || 'video/webm' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `photo-3d-motion.${ext}`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 10_000);
    recorder = null;
    state.exporting = false;
    btn.disabled = false;
    btn.textContent = '⬤ Export Video';
    showStatus(`Saved photo-3d-motion.${ext}` +
      (ext === 'webm' ? ' (convert to MP4 with ffmpeg if needed)' : ''), null);
    setTimeout(hideStatus, 5000);
  };

  setTime(0);
  recorder.start();
  startPlayback();
}

// ---------------------------------------------------------------------------
// Pointer interaction: select & drag layers in the viewport
// ---------------------------------------------------------------------------

const raycaster = new THREE.Raycaster();
const pointerNDC = new THREE.Vector2();
let dragging = null;
const dragPlane = new THREE.Plane();
const dragPoint = new THREE.Vector3();
const dragOffset = new THREE.Vector3();

function updatePointer(e) {
  const r = renderer.domElement.getBoundingClientRect();
  pointerNDC.x = ((e.clientX - r.left) / r.width) * 2 - 1;
  pointerNDC.y = -((e.clientY - r.top) / r.height) * 2 + 1;
  raycaster.setFromCamera(pointerNDC, camera);
}

renderer.domElement.addEventListener('pointerdown', (e) => {
  if (state.playing || state.exporting) return;
  updatePointer(e);
  const meshes = state.layers.filter((l) => l.mesh.visible).map((l) => l.mesh);
  const hits = raycaster.intersectObjects(meshes, false);
  if (hits.length) {
    const l = state.layers.find((x) => x.id === hits[0].object.userData.layerId);
    selectLayer(l.id);
    dragging = l;
    controls.enabled = false;
    const camDir = new THREE.Vector3();
    camera.getWorldDirection(camDir);
    dragPlane.setFromNormalAndCoplanarPoint(camDir, l.mesh.position);
    raycaster.ray.intersectPlane(dragPlane, dragPoint);
    dragOffset.copy(l.mesh.position).sub(dragPoint);
    renderer.domElement.setPointerCapture(e.pointerId);
  }
}, true);

renderer.domElement.addEventListener('pointermove', (e) => {
  if (!dragging) return;
  updatePointer(e);
  if (raycaster.ray.intersectPlane(dragPlane, dragPoint)) {
    dragging.x = dragPoint.x + dragOffset.x;
    dragging.y = dragPoint.y + dragOffset.y;
  }
});

window.addEventListener('pointerup', () => {
  if (dragging) {
    dragging = null;
    if (!state.playing) controls.enabled = true;
  }
});

function raycastLayerToSurface(l) {
  if (!photoMesh) return null;
  const dir = l.mesh.position.clone().sub(camera.position).normalize();
  raycaster.set(camera.position, dir);
  const hits = raycaster.intersectObject(photoMesh, false);
  return hits.length ? hits[0].point : null;
}

// ---------------------------------------------------------------------------
// UI: status, layer list, props panel
// ---------------------------------------------------------------------------

function showStatus(text, pct) {
  const el = document.getElementById('status');
  el.classList.add('visible');
  document.getElementById('statusText').textContent = text;
  document.getElementById('statusBar').style.width =
    pct == null ? '0%' : `${pct.toFixed(0)}%`;
}
function hideStatus() {
  document.getElementById('status').classList.remove('visible');
}

function renderLayerList() {
  const list = document.getElementById('layerList');
  list.innerHTML = '';
  for (const l of [...state.layers].reverse()) {
    const item = document.createElement('div');
    item.className = 'layer-item' + (l.id === state.selectedId ? ' selected' : '');
    const name = document.createElement('span');
    name.className = 'lname';
    name.textContent = l.type === 'text' ? l.text : l.name;
    const type = document.createElement('span');
    type.className = 'ltype';
    type.textContent = l.type;
    const del = document.createElement('button');
    del.className = 'del';
    del.textContent = '✕';
    del.onclick = (e) => { e.stopPropagation(); deleteLayer(l.id); };
    item.append(name, type, del);
    item.onclick = () => selectLayer(l.id);
    list.appendChild(item);
  }
}

const $ = (id) => document.getElementById(id);

function renderProps() {
  const l = selectedLayer();
  const panel = $('props');
  panel.classList.toggle('visible', !!l);
  if (!l) return;
  $('propText').style.display = l.type === 'text' ? '' : 'none';
  $('propShape').style.display = l.type === 'shape' ? '' : 'none';
  $('pText').value = l.text;
  $('pFont').value = l.font;
  $('pWeight').value = l.weight;
  $('pShape').value = l.shape;
  $('pColor').value = l.color;
  $('pOutline').value = l.outline;
  $('pScale').value = l.scale;
  $('pZ').value = l.z;
  $('pRot').value = l.rot;
  $('pOpacity').value = l.opacity;
  $('pAnim').value = l.anim;
  $('pAppear').value = l.appear;
  $('pScaleV').textContent = `(${(+l.scale).toFixed(2)})`;
  $('pZV').textContent = `(${(+l.z).toFixed(2)})`;
  $('pRotV').textContent = `(${l.rot}°)`;
  $('pOpacityV').textContent = `(${(+l.opacity).toFixed(2)})`;
}

function bindLayerProp(id, prop, { rebuild = false, num = false } = {}) {
  $(id).addEventListener('input', () => {
    const l = selectedLayer();
    if (!l) return;
    l[prop] = num ? parseFloat($(id).value) || 0 : $(id).value;
    if (rebuild) rebuildLayerTexture(l);
    if (prop === 'text') renderLayerList();
    renderProps();
  });
}

bindLayerProp('pText', 'text', { rebuild: true });
bindLayerProp('pFont', 'font', { rebuild: true });
bindLayerProp('pWeight', 'weight', { rebuild: true });
bindLayerProp('pShape', 'shape', { rebuild: true });
bindLayerProp('pColor', 'color', { rebuild: true });
bindLayerProp('pOutline', 'outline', { rebuild: true });
bindLayerProp('pScale', 'scale', { num: true });
bindLayerProp('pZ', 'z', { num: true });
bindLayerProp('pRot', 'rot', { num: true });
bindLayerProp('pOpacity', 'opacity', { num: true });
bindLayerProp('pAnim', 'anim');
bindLayerProp('pAppear', 'appear', { num: true });

$('btnSnap').onclick = () => {
  const l = selectedLayer();
  if (!l) return;
  const hit = raycastLayerToSurface(l);
  if (hit) { l.z = hit.z + 0.04; renderProps(); }
};
$('btnBehind').onclick = () => {
  const l = selectedLayer();
  if (!l) return;
  const hit = raycastLayerToSurface(l);
  if (hit) { l.z = hit.z - 0.18; renderProps(); }
};
$('btnDup').onclick = () => state.selectedId && duplicateLayer(state.selectedId);
$('btnDelete').onclick = () => state.selectedId && deleteLayer(state.selectedId);

// Scene controls -----------------------------------------------------------

$('depthScale').addEventListener('input', () => {
  state.depthScale = parseFloat($('depthScale').value);
  $('depthScaleV').textContent = `(${state.depthScale.toFixed(2)})`;
  updateDisplacement();
});
$('depthBlur').addEventListener('input', () => {
  state.depthBlur = parseInt($('depthBlur').value, 10);
  $('depthBlurV').textContent = `(${state.depthBlur}px)`;
  rebuildDepthWorkingCanvas();
});
$('bgColor').addEventListener('input', () => {
  state.bg = $('bgColor').value;
  scene.background = new THREE.Color(state.bg);
});
$('camPreset').addEventListener('change', () => {
  state.camPreset = $('camPreset').value;
  applyCameraAtTime(state.t);
});
$('camAmount').addEventListener('input', () => {
  state.camAmount = parseFloat($('camAmount').value);
  $('camAmountV').textContent = `(${state.camAmount.toFixed(2)})`;
  applyCameraAtTime(state.t);
});
$('duration').addEventListener('input', () => {
  state.duration = parseFloat($('duration').value);
  $('durationV').textContent = `(${state.duration.toFixed(1)}s)`;
  $('scrub').max = state.duration;
  setTime(Math.min(state.t, state.duration));
});
$('scrub').max = state.duration;
$('scrub').addEventListener('input', () => {
  if (state.exporting) return;
  stopPlayback();
  setTime(parseFloat($('scrub').value));
});
$('btnPlay').onclick = () => {
  if (state.playing) stopPlayback();
  else { if (state.t >= state.duration - 0.01) setTime(0); startPlayback(); }
};
$('btnExport').onclick = exportVideo;

// Top bar ------------------------------------------------------------------

$('btnLoadPhoto').onclick = () => $('filePhoto').click();
$('btnLoadDepth').onclick = () => $('fileDepth').click();
$('btnGenDepth').onclick = generateDepth;
$('btnShowDepth').onclick = () => {
  state.showDepth = !state.showDepth;
  $('btnShowDepth').textContent = state.showDepth ? 'View Photo' : 'View Depth';
  updateDepthView();
};
$('btnAddText').onclick = () => createLayer('text', { text: 'Hello!', name: 'Text' });
$('btnAddShape').onclick = () => createLayer('shape', { name: 'Shape', color: '#5b8cff' });
$('btnAddImage').onclick = () => $('fileImageLayer').click();

function fileToDataURL(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = rej;
    r.readAsDataURL(file);
  });
}

function loadImageFromURL(url) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => res(img);
    img.onerror = rej;
    img.src = url;
  });
}

async function loadPhoto(file) {
  const dataURL = await fileToDataURL(file);
  await loadPhotoFromDataURL(dataURL, true);
}

async function loadPhotoFromDataURL(dataURL, resetDepth) {
  const img = await loadImageFromURL(dataURL);
  state.photoDataURL = dataURL;
  state.photoW = img.width;
  state.photoH = img.height;
  if (resetDepth) {
    depthCanvasOriginal = null;
    state.hasDepth = false;
    state.showDepth = false;
    $('btnShowDepth').disabled = true;
    $('btnShowDepth').textContent = 'View Depth';
  }
  buildPhotoMesh(img);
  document.getElementById('dropHint').style.display = 'none';
  for (const id of ['btnGenDepth', 'btnLoadDepth', 'btnAddText', 'btnAddImage',
    'btnAddShape', 'btnExport', 'btnPlay', 'btnSave']) {
    $(id).disabled = false;
  }
  setTime(0);
}

$('filePhoto').addEventListener('change', async (e) => {
  if (e.target.files[0]) await loadPhoto(e.target.files[0]);
  e.target.value = '';
});

$('fileDepth').addEventListener('change', async (e) => {
  const f = e.target.files[0];
  e.target.value = '';
  if (!f) return;
  const img = await loadImageFromURL(await fileToDataURL(f));
  const c = document.createElement('canvas');
  c.width = img.width; c.height = img.height;
  c.getContext('2d').drawImage(img, 0, 0);
  depthCanvasOriginal = c;
  rebuildDepthWorkingCanvas();
  if (depthTex) depthTex.dispose();
  depthTex = new THREE.CanvasTexture(depthCanvas);
  if (photoMesh) {
    photoMesh.material.displacementMap = depthTex;
    photoMesh.material.needsUpdate = true;
  }
  state.hasDepth = true;
  $('btnShowDepth').disabled = false;
  updateDepthView();
});

$('fileImageLayer').addEventListener('change', async (e) => {
  const f = e.target.files[0];
  e.target.value = '';
  if (!f) return;
  const dataURL = await fileToDataURL(f);
  createLayer('image', { imageDataURL: dataURL, name: f.name });
});

// Drag & drop photo onto stage
stage.addEventListener('dragover', (e) => e.preventDefault());
stage.addEventListener('drop', async (e) => {
  e.preventDefault();
  const f = e.dataTransfer.files[0];
  if (f && f.type.startsWith('image/')) {
    if (!state.photoDataURL) await loadPhoto(f);
    else createLayer('image', { imageDataURL: await fileToDataURL(f), name: f.name });
  }
});

// ---------------------------------------------------------------------------
// Project save / load
// ---------------------------------------------------------------------------

$('btnSave').onclick = () => {
  const project = {
    version: 1,
    photo: state.photoDataURL,
    depth: state.hasDepth && depthCanvasOriginal ? depthCanvasOriginal.toDataURL() : null,
    settings: {
      depthScale: state.depthScale,
      depthBlur: state.depthBlur,
      bg: state.bg,
      camPreset: state.camPreset,
      camAmount: state.camAmount,
      duration: state.duration,
    },
    layers: state.layers.map(({ mesh, id, ...rest }) => rest),
  };
  const blob = new Blob([JSON.stringify(project)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'photo-3d-project.json';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10_000);
};

$('btnOpen').onclick = () => $('fileProject').click();
$('fileProject').addEventListener('change', async (e) => {
  const f = e.target.files[0];
  e.target.value = '';
  if (!f) return;
  try {
    const project = JSON.parse(await f.text());
    // clear existing layers
    for (const l of [...state.layers]) deleteLayer(l.id);
    if (project.depth) {
      const img = await loadImageFromURL(project.depth);
      const c = document.createElement('canvas');
      c.width = img.width; c.height = img.height;
      c.getContext('2d').drawImage(img, 0, 0);
      depthCanvasOriginal = c;
      state.hasDepth = true;
    } else {
      depthCanvasOriginal = null;
      state.hasDepth = false;
    }
    const s = project.settings || {};
    state.depthScale = s.depthScale ?? 1;
    state.depthBlur = s.depthBlur ?? 2;
    state.bg = s.bg ?? '#05070a';
    state.camPreset = s.camPreset ?? 'orbit';
    state.camAmount = s.camAmount ?? 0.55;
    state.duration = s.duration ?? 6;
    $('depthScale').value = state.depthScale;
    $('depthBlur').value = state.depthBlur;
    $('bgColor').value = state.bg;
    $('camPreset').value = state.camPreset;
    $('camAmount').value = state.camAmount;
    $('duration').value = state.duration;
    $('scrub').max = state.duration;
    scene.background = new THREE.Color(state.bg);
    await loadPhotoFromDataURL(project.photo, false);
    if (state.hasDepth) $('btnShowDepth').disabled = false;
    for (const ld of project.layers || []) createLayer(ld.type, ld);
    state.selectedId = null;
    renderLayerList();
    renderProps();
  } catch (err) {
    console.error(err);
    showStatus('Could not open project file.', null);
    setTimeout(hideStatus, 4000);
  }
});

// Initial UI value labels
$('depthScaleV').textContent = `(${state.depthScale.toFixed(2)})`;
$('depthBlurV').textContent = `(${state.depthBlur}px)`;
$('camAmountV').textContent = `(${state.camAmount.toFixed(2)})`;
$('durationV').textContent = `(${state.duration.toFixed(1)}s)`;
setTime(0, false);
