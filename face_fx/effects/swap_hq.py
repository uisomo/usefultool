"""高品質フェイススワップ（InsightFace inswapper / Deep-Live-Cam方式）.

Deep-Live-Cam と同じく InsightFace の検出 + inswapper_128.onnx を使う。
onnxruntime(-gpu) と insightface が必要で、モデル(inswapper_128.onnx, ~300MB)を
別途入手して AppConfig.hq_model_path に指定する。GPUがあれば CUDA/CoreML を使う。

依存が無い環境でも face_fx 全体は動くよう、import は遅延・任意にしている。

複数人対応:
  InsightFace の顔検出は1フレームにつき1回だけ実行してキャッシュする
  (検出が重いため)。各 ID の HQSwapEffect は「自分のトラッキング顔に最も近い
  検出顔」を1つだけ選び、自分の source 顔へ差し替える。同じ検出顔を複数IDが
  奪わないよう、フレーム単位で使用済みインデックスを管理する。
  → ID 0 は顔A、ID 1 は顔B…と人別に別人化できる。
"""
from __future__ import annotations

from typing import List, Optional

import numpy as np

from .base import Effect


class HQSwapEngine:
    """InsightFace のアプリ + inswapper を1つだけ持ち、全IDで共有する."""

    def __init__(self, cfg):
        self.cfg = cfg
        self._app = None
        self._swapper = None
        self._ready = False
        # フレーム単位の検出キャッシュ
        self._cache_frame_id = None
        self._cache_targets: List = []
        self._cache_centers: Optional[np.ndarray] = None
        self._cache_used: set = set()

    def ensure_loaded(self):
        if self._ready:
            return True
        try:
            import insightface
            from insightface.app import FaceAnalysis
        except ImportError as e:  # pragma: no cover - 任意依存
            raise ImportError(
                "高品質スワップには insightface / onnxruntime が必要です。\n"
                "  pip install insightface onnxruntime  (GPUなら onnxruntime-gpu)\n"
                "さらに inswapper_128.onnx を入手して --hq-model で指定してください。"
            ) from e

        providers = (
            ["CUDAExecutionProvider", "CoreMLExecutionProvider", "CPUExecutionProvider"]
            if self.cfg.hq_use_gpu else ["CPUExecutionProvider"]
        )
        self._app = FaceAnalysis(name="buffalo_l", providers=providers)
        self._app.prepare(ctx_id=0, det_size=(640, 640))
        if not self.cfg.hq_model_path:
            raise ValueError("hq_model_path(inswapper_128.onnx) が未設定です。")
        self._swapper = insightface.model_zoo.get_model(
            self.cfg.hq_model_path, providers=providers
        )
        self._ready = True
        return True

    def detect_source(self, image_path):
        """source 顔画像から InsightFace の Face 埋め込みを取得."""
        import cv2
        self.ensure_loaded()
        img = cv2.imread(image_path)
        if img is None:
            raise FileNotFoundError(f"swap_face 画像を読めません: {image_path}")
        faces = self._app.get(img)
        if not faces:
            raise ValueError(f"swap_face 画像から顔を検出できません: {image_path}")
        return max(faces, key=lambda f: (f.bbox[2] - f.bbox[0]) * (f.bbox[3] - f.bbox[1]))

    def make_effect(self, ecfg) -> "HQSwapEffect":
        return HQSwapEffect(ecfg, self)

    # --- フレーム単位の検出キャッシュ ---
    def _ensure_targets(self, frame):
        fid = id(frame)
        if fid == self._cache_frame_id:
            return
        self.ensure_loaded()
        targets = self._app.get(frame)
        centers = np.array(
            [[(f.bbox[0] + f.bbox[2]) / 2.0, (f.bbox[1] + f.bbox[3]) / 2.0] for f in targets],
            dtype=np.float32,
        ) if targets else np.empty((0, 2), np.float32)
        self._cache_frame_id = fid
        self._cache_targets = targets
        self._cache_centers = centers
        self._cache_used = set()

    def swap_nearest(self, frame, mp_centroid, src_face, blend):
        """mp_centroid(トラッキング顔の重心)に最も近い検出顔を1つだけ差し替える."""
        import cv2
        self._ensure_targets(frame)
        if not self._cache_targets:
            return frame

        c = np.asarray(mp_centroid, np.float32)
        d = np.linalg.norm(self._cache_centers - c, axis=1)
        # 使用済みは除外
        order = np.argsort(d)
        ti = None
        for idx in order:
            if int(idx) not in self._cache_used:
                ti = int(idx)
                break
        if ti is None:
            return frame
        # 近すぎない(別人/誤検出)を弾く: 画面対角の25%以内のみ採用
        diag = float(np.hypot(frame.shape[1], frame.shape[0]))
        if d[ti] > 0.25 * diag:
            return frame
        self._cache_used.add(ti)

        target = self._cache_targets[ti]
        swapped = self._swapper.get(frame, target, src_face, paste_back=True)
        if self.cfg.hq_enhance:
            swapped = self._maybe_enhance(swapped, target)

        # blend は対象顔の bbox 内だけで適用（フレーム全体を混ぜない）
        if blend >= 1.0:
            return swapped
        x1, y1, x2, y2 = [int(v) for v in target.bbox]
        x1, y1 = max(0, x1), max(0, y1)
        x2 = min(frame.shape[1], x2); y2 = min(frame.shape[0], y2)
        if x2 > x1 and y2 > y1:
            roi_new = swapped[y1:y2, x1:x2]
            roi_old = frame[y1:y2, x1:x2]
            swapped[y1:y2, x1:x2] = cv2.addWeighted(roi_new, blend, roi_old, 1 - blend, 0)
        return swapped

    @staticmethod
    def _maybe_enhance(img, target):  # GFPGAN等を入れる拡張ポイント
        return img


class HQSwapEffect(Effect):
    """1つの EffectConfig(=ID)に対応。自分の source 顔を保持し、
    そのIDのトラッキング顔だけを差し替える（人別に別人化できる）。
    """

    def __init__(self, ecfg, engine: HQSwapEngine):
        super().__init__(ecfg)
        self.engine = engine
        self._src_face = None

    def _ensure_src(self):
        if self._src_face is None and self.cfg.swap_face_path:
            self._src_face = self.engine.detect_source(self.cfg.swap_face_path)

    def apply(self, frame, face):
        self._ensure_src()
        if self._src_face is None:
            return frame
        return self.engine.swap_nearest(
            frame, face.centroid, self._src_face, float(self.cfg.swap_blend)
        )
