"""顔を“若干”変えるエフェクト（肌スムージング・明るさ）.

別人化まではせず、顔領域だけを軽く整える。FaceMesh の凸包内のみ加工するので
背景や髪は影響を受けない。バイラテラルで毛穴/ノイズを抑えつつ輪郭は保つ。
GPU不要で軽い。MediaPipe FaceStylizer を足せばスタイル変換も可能（任意拡張）。
"""
from __future__ import annotations

import numpy as np

from .base import Effect

# FaceMesh の顔輪郭インデックス（顔オーバル）
_FACE_OVAL = [
    10, 338, 297, 332, 284, 251, 389, 356, 454, 323, 361, 288, 397, 365, 379,
    378, 400, 377, 152, 148, 176, 149, 150, 136, 172, 58, 132, 93, 234, 127,
    162, 21, 54, 103, 67, 109,
]


class AlterEffect(Effect):
    def apply(self, frame, face):
        import cv2

        n = len(face.landmarks)
        oval = [i for i in _FACE_OVAL if i < n]
        if len(oval) < 3:
            return frame
        pts = face.landmarks[oval].astype(np.int32)
        hull = cv2.convexHull(pts)

        x, y, w, h = cv2.boundingRect(hull)
        x, y = max(0, x), max(0, y)
        w = min(frame.shape[1] - x, w); h = min(frame.shape[0] - y, h)
        if w <= 0 or h <= 0:
            return frame

        roi = frame[y:y + h, x:x + w]
        smooth = float(np.clip(self.cfg.alter_smooth, 0.0, 1.0))
        out = roi
        if smooth > 0:
            d = int(5 + smooth * 10)
            sigma = 20 + smooth * 60
            blurred = cv2.bilateralFilter(roi, d, sigma, sigma)
            out = cv2.addWeighted(blurred, smooth, roi, 1 - smooth, 0)

        if abs(self.cfg.alter_brightness) > 0.5:
            out = cv2.convertScaleAbs(out, alpha=1.0, beta=float(self.cfg.alter_brightness))

        mask = np.zeros((h, w), np.uint8)
        cv2.fillConvexPoly(mask, hull - [x, y], 255)
        mask = cv2.GaussianBlur(mask, (15, 15), 0)
        m = (mask[:, :, None].astype(np.float32)) / 255.0
        frame[y:y + h, x:x + w] = (out * m + roi * (1 - m)).astype(np.uint8)
        return frame
