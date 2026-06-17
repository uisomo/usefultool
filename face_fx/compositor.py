"""背景処理と最終合成（原寸）.

person_mask(0..1, 1=残す) を使って、トラッキング中の人以外の領域を
blur / 単色(クロマキー) / 画像 に差し替える。合成はすべて原寸で行うため、
iPhone等の高画質をそのまま保てる（検出だけ縮小して軽量化している）。
"""
from __future__ import annotations

from typing import Optional

import numpy as np


class Compositor:
    def __init__(self, cfg):
        self.cfg = cfg
        self._bg_image = None
        self._bg_image_size = None

    def _background_plate(self, frame):
        import cv2
        mode = self.cfg.background_mode
        h, w = frame.shape[:2]
        if mode == "blur":
            k = self.cfg.background_blur
            k = k if k % 2 == 1 else k + 1
            return cv2.GaussianBlur(frame, (k, k), 0)
        if mode in ("color", "transparent"):
            color = self.cfg.background_color  # BGR
            plate = np.empty_like(frame)
            plate[:] = color
            return plate
        if mode == "image":
            if self._bg_image is None or self._bg_image_size != (w, h):
                if not self.cfg.background_image:
                    raise ValueError("background_mode='image' だが background_image 未指定")
                img = cv2.imread(self.cfg.background_image)
                if img is None:
                    raise FileNotFoundError(f"背景画像を読めません: {self.cfg.background_image}")
                self._bg_image = cv2.resize(img, (w, h))
                self._bg_image_size = (w, h)
            return self._bg_image
        return None  # "keep"

    def composite(self, frame, person_mask):
        """frame: 加工済みBGR(原寸), person_mask: (H,W) float 0..1."""
        if self.cfg.background_mode == "keep" or person_mask is None:
            return frame
        plate = self._background_plate(frame)
        if plate is None:
            return frame
        m = person_mask[:, :, None]
        out = frame.astype(np.float32) * m + plate.astype(np.float32) * (1 - m)
        return out.astype(np.uint8)
