"""人物セグメンテーション（背景除去）.

MediaPipe Selfie Segmentation で「人 vs 背景」マスクを軽量に得る。
ただし Selfie Segmentation は“画面内の全員”を人として残すため、
`keep_only_tracked=True` のときは、トラッキング中の顔を含む連結成分だけを
残し、それ以外の人（＝トラッキング対象外の他人）も背景として消す。

検出は proc_scale で縮小して行い、マスクは原寸へ拡大してから合成する。
"""
from __future__ import annotations

from typing import List

import numpy as np

try:
    import mediapipe as mp
except ImportError:
    mp = None


class PersonSegmenter:
    def __init__(self, cfg):
        if mp is None:
            raise ImportError(
                "mediapipe が見つかりません。`pip install -r face_fx/requirements.txt` を実行してください。"
            )
        self.cfg = cfg
        # model_selection=1: 一般用（全身寄り・少し高精度）
        self.seg = mp.solutions.selfie_segmentation.SelfieSegmentation(model_selection=1)

    def mask(self, frame_bgr: np.ndarray, faces: "List") -> np.ndarray:
        """0..1 の float32 マスク(原寸)を返す。1=残す(人), 0=消す(背景)."""
        import cv2

        h, w = frame_bgr.shape[:2]
        scale = self.cfg.proc_scale
        if scale != 1.0:
            small = cv2.resize(frame_bgr, (max(1, int(w * scale)), max(1, int(h * scale))))
        else:
            small = frame_bgr

        rgb = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
        rgb.flags.writeable = False
        res = self.seg.process(rgb)
        prob = res.segmentation_mask  # (sh, sw) float32 0..1
        binary = (prob > 0.5).astype(np.uint8)

        if self.cfg.keep_only_tracked and faces:
            binary = self._keep_face_components(binary, faces, w, h)

        mask = cv2.resize(binary.astype(np.float32), (w, h), interpolation=cv2.INTER_LINEAR)

        # 境界を羽化して合成を自然にする
        f = self.cfg.mask_feather
        if f and f >= 3:
            f = f if f % 2 == 1 else f + 1
            mask = cv2.GaussianBlur(mask, (f, f), 0)
        return np.clip(mask, 0.0, 1.0)

    def _keep_face_components(self, binary, faces, full_w, full_h):
        """トラッキング中の顔重心を含む連結成分のみ残す."""
        import cv2

        sh, sw = binary.shape[:2]
        num, labels = cv2.connectedComponents(binary)
        keep = np.zeros_like(binary)
        sx, sy = sw / full_w, sh / full_h
        kept_labels = set()
        for f in faces:
            cx, cy = f.centroid
            px = int(np.clip(cx * sx, 0, sw - 1))
            py = int(np.clip(cy * sy, 0, sh - 1))
            lab = labels[py, px]
            # 顔の中心が背景(0)に落ちた場合は近傍を探索
            if lab == 0:
                lab = self._nearest_label(labels, px, py)
            if lab > 0:
                kept_labels.add(lab)
        for lab in kept_labels:
            keep[labels == lab] = 1
        return keep

    @staticmethod
    def _nearest_label(labels, px, py, radius=6):
        sh, sw = labels.shape
        for r in range(1, radius + 1):
            y0, y1 = max(0, py - r), min(sh, py + r + 1)
            x0, x1 = max(0, px - r), min(sw, px + r + 1)
            patch = labels[y0:y1, x0:x1]
            nz = patch[patch > 0]
            if nz.size:
                vals, counts = np.unique(nz, return_counts=True)
                return int(vals[counts.argmax()])
        return 0

    def close(self):
        self.seg.close()
