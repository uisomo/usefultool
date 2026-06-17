"""オーバーレイ（PNGステッカー/マスク）を顔に合わせて貼る軽量エフェクト.

2つのアンカーランドマーク（既定は左右の目尻）から顔の幅・傾き・中心を求め、
RGBA画像を回転・拡大して合成する。表情でランドマークが動くと追従する。
GPU不要で非常に軽い。眼鏡・仮面・キャラ顔など何でも貼れる。
"""
from __future__ import annotations

import numpy as np

from .base import Effect


class OverlayEffect(Effect):
    def __init__(self, cfg):
        super().__init__(cfg)
        self._png = None
        if cfg.overlay_path:
            import cv2
            img = cv2.imread(cfg.overlay_path, cv2.IMREAD_UNCHANGED)
            if img is None:
                raise FileNotFoundError(f"overlay 画像を読めません: {cfg.overlay_path}")
            if img.shape[2] == 3:  # アルファが無ければ不透明扱い
                alpha = np.full(img.shape[:2] + (1,), 255, np.uint8)
                img = np.concatenate([img, alpha], axis=2)
            self._png = img

    def apply(self, frame, face):
        if self._png is None:
            return frame
        import cv2

        a, b = self.cfg.overlay_anchor
        n = len(face.landmarks)
        if a >= n or b >= n:
            return frame
        pa = face.landmarks[a]
        pb = face.landmarks[b]

        center = (pa + pb) / 2.0
        dx, dy = (pb - pa)
        face_w = float(np.hypot(dx, dy))
        if face_w < 1:
            return frame
        angle = np.degrees(np.arctan2(dy, dx))

        target_w = face_w * self.cfg.overlay_scale
        oh, ow = self._png.shape[:2]
        s = target_w / ow
        target_h = oh * s

        # 縦オフセット（顔幅基準）
        center = center + np.array([0.0, self.cfg.overlay_y_offset * face_w], np.float32)

        # 拡大→回転した一枚絵を作り、中心合わせで貼る
        resized = cv2.resize(self._png, (max(1, int(target_w)), max(1, int(target_h))))
        rot = self._rotate_rgba(resized, -angle)
        self._alpha_paste(frame, rot, center)
        return frame

    @staticmethod
    def _rotate_rgba(img, angle_deg):
        import cv2
        h, w = img.shape[:2]
        diag = int(np.ceil(np.hypot(h, w)))
        canvas = np.zeros((diag, diag, 4), np.uint8)
        oy, ox = (diag - h) // 2, (diag - w) // 2
        canvas[oy:oy + h, ox:ox + w] = img
        M = cv2.getRotationMatrix2D((diag / 2, diag / 2), angle_deg, 1.0)
        return cv2.warpAffine(canvas, M, (diag, diag), flags=cv2.INTER_LINEAR,
                              borderMode=cv2.BORDER_CONSTANT, borderValue=(0, 0, 0, 0))

    @staticmethod
    def _alpha_paste(frame, rgba, center):
        h, w = rgba.shape[:2]
        cx, cy = center
        x0 = int(round(cx - w / 2)); y0 = int(round(cy - h / 2))
        x1, y1 = x0 + w, y0 + h

        fx0, fy0 = max(0, x0), max(0, y0)
        fx1, fy1 = min(frame.shape[1], x1), min(frame.shape[0], y1)
        if fx0 >= fx1 or fy0 >= fy1:
            return
        ox0, oy0 = fx0 - x0, fy0 - y0
        ox1, oy1 = ox0 + (fx1 - fx0), oy0 + (fy1 - fy0)

        sub = rgba[oy0:oy1, ox0:ox1]
        alpha = (sub[:, :, 3:4].astype(np.float32)) / 255.0
        roi = frame[fy0:fy1, fx0:fx1].astype(np.float32)
        blended = sub[:, :, :3].astype(np.float32) * alpha + roi * (1 - alpha)
        frame[fy0:fy1, fx0:fx1] = blended.astype(np.uint8)
