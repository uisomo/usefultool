"""映像入力（カメラ / 動画ファイル / iPhone）.

iPhone を高画質ソースとして使う方法（macOS）:
- Continuity Camera: iPhone が通常のカメラとして OS に現れる → source をその index に
- もしくは OBS Virtual Camera / EpocCam / Camo を経由して仮想カメラ index を指定
capture_width/height に 1920x1080 や 3840x2160 を設定すると、対応端末では
その解像度のまま取り込み、合成も原寸で行うため画質を落とさない。
"""
from __future__ import annotations

import cv2


class VideoSource:
    def __init__(self, source, width=0, height=0, fps=0):
        self.cap = cv2.VideoCapture(source)
        if not self.cap.isOpened():
            raise RuntimeError(f"映像ソースを開けませんでした: {source!r}")
        # 高解像度を要求（端末が対応していれば反映される）
        if width:
            self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, width)
        if height:
            self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, height)
        if fps:
            self.cap.set(cv2.CAP_PROP_FPS, fps)

    @property
    def size(self):
        w = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        h = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        return w, h

    @property
    def fps(self):
        f = self.cap.get(cv2.CAP_PROP_FPS)
        return f if f and f > 0 else 30.0

    def read(self):
        ok, frame = self.cap.read()
        return frame if ok else None

    def __iter__(self):
        while True:
            frame = self.read()
            if frame is None:
                break
            yield frame

    def release(self):
        self.cap.release()
