"""出力先（プレビュー窓 / mp4ファイル / 仮想カメラ）.

仮想カメラ(pyvirtualcam)に流すと Zoom / OBS / Google Meet 等で
「face_fx カメラ」として選択でき、Camo→加工→各アプリ の構成が完成する。
"""
from __future__ import annotations

from typing import List

import numpy as np


class OutputSink:
    def write(self, frame_bgr: np.ndarray):
        ...

    def close(self):
        ...


class WindowSink(OutputSink):
    def __init__(self, name="face_fx"):
        import cv2
        self.cv2 = cv2
        self.name = name

    def write(self, frame_bgr):
        self.cv2.imshow(self.name, frame_bgr)

    def close(self):
        self.cv2.destroyAllWindows()


class FileSink(OutputSink):
    def __init__(self, path, size, fps):
        import cv2
        fourcc = cv2.VideoWriter_fourcc(*"mp4v")
        self.writer = cv2.VideoWriter(path, fourcc, fps, size)
        if not self.writer.isOpened():
            raise RuntimeError(f"出力ファイルを開けません: {path}")

    def write(self, frame_bgr):
        self.writer.write(frame_bgr)

    def close(self):
        self.writer.release()


class VirtualCamSink(OutputSink):
    def __init__(self, size, fps, backend=None):
        try:
            import pyvirtualcam
        except ImportError as e:
            raise ImportError(
                "仮想カメラ出力には pyvirtualcam が必要です: pip install pyvirtualcam\n"
                "(macOS は OBS Virtual Camera, Windows は OBS / Unity Capture が必要)"
            ) from e
        self._pvc = pyvirtualcam
        w, h = size
        kwargs = dict(width=w, height=h, fps=int(round(fps)),
                      fmt=pyvirtualcam.PixelFormat.BGR)
        if backend:
            kwargs["backend"] = backend
        self.cam = pyvirtualcam.Camera(**kwargs)

    def write(self, frame_bgr):
        self.cam.send(frame_bgr)
        self.cam.sleep_until_next_frame()

    def close(self):
        self.cam.close()


def build_sinks(cfg, size, fps) -> List[OutputSink]:
    sinks: List[OutputSink] = []
    for o in cfg.outputs:
        if o == "window":
            sinks.append(WindowSink())
        elif o == "file":
            if not cfg.output_path:
                raise ValueError("outputs に 'file' があるが output_path 未指定")
            sinks.append(FileSink(cfg.output_path, size, fps))
        elif o == "virtualcam":
            sinks.append(VirtualCamSink(size, fps, cfg.virtualcam_backend))
        else:
            raise ValueError(f"未知の出力先: {o}")
    return sinks
