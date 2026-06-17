"""高品質フェイススワップ（InsightFace inswapper / Deep-Live-Cam方式）.

Deep-Live-Cam と同じく InsightFace の検出 + inswapper_128.onnx を使う。
onnxruntime(-gpu) と insightface が必要で、モデル(inswapper_128.onnx, ~300MB)を
別途入手して AppConfig.hq_model_path に指定する。GPUがあれば CUDA/CoreML を使う。

依存が無い環境でも face_fx 全体は動くよう、import は遅延・任意にしている。
この HQ エンジンは複数IDで共有し、各 EffectConfig ごとに source 顔を保持する
軽量ラッパ(Effect)を make_effect() で生成する。
"""
from __future__ import annotations

from typing import Optional

import numpy as np

from .base import Effect


class HQSwapEngine:
    """InsightFace のアプリ + inswapper を1つだけ持ち、全IDで共有する."""

    def __init__(self, cfg):
        self.cfg = cfg
        self._app = None
        self._swapper = None
        self._ready = False

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

    def swap_in_frame(self, frame, src_face, blend):
        """frame 全体から顔を検出して src_face に差し替える(原寸)。"""
        import cv2
        self.ensure_loaded()
        targets = self._app.get(frame)
        out = frame
        for tf in targets:
            out = self._swapper.get(out, tf, src_face, paste_back=True)
        if self.cfg.hq_enhance:
            out = self._maybe_enhance(out)
        if blend >= 1.0:
            return out
        return cv2.addWeighted(out, blend, frame, 1 - blend, 0)

    @staticmethod
    def _maybe_enhance(img):  # GFPGAN等を入れる拡張ポイント
        return img


class HQSwapEffect(Effect):
    """1つの EffectConfig(=ID)に対応。source 顔を保持して HQ エンジンで差し替える。

    注: InsightFace は frame 全体で顔検出するため、per-face ではなく
    「最初の1回だけ」frame 全体を処理する。pipeline 側で同一フレームに対する
    重複実行を避けるためフレームidで一度だけ走るようにしている。
    """

    def __init__(self, ecfg, engine: HQSwapEngine):
        super().__init__(ecfg)
        self.engine = engine
        self._src_face = None
        self._last_frame_id = None

    def _ensure_src(self):
        if self._src_face is None and self.cfg.swap_face_path:
            self._src_face = self.engine.detect_source(self.cfg.swap_face_path)

    def apply(self, frame, face):
        # フレーム全体を処理する方式なので、同じ frame オブジェクトには一度だけ適用
        fid = id(frame)
        if fid == self._last_frame_id:
            return frame
        self._last_frame_id = fid
        self._ensure_src()
        if self._src_face is None:
            return frame
        return self.engine.swap_in_frame(frame, self._src_face, float(self.cfg.swap_blend))
