"""パイプライン全体の制御.

capture → track(複数人/安定ID) → segment(人だけ残す) → 顔エフェクト(ID別)
→ 背景合成 → 出力(窓/ファイル/仮想カメラ)。
ライブ中のキー操作でエフェクト種別・スワップエンジン・背景を切り替えられる。
"""
from __future__ import annotations

import time
from typing import Dict, Optional

import numpy as np

from .capture import VideoSource
from .compositor import Compositor
from .config import AppConfig, EffectConfig
from .effects import build_effect
from .effects.swap_hq import HQSwapEngine
from .output import build_sinks
from .segmenter import PersonSegmenter
from .tracker import FaceTracker


class Pipeline:
    def __init__(self, cfg: AppConfig):
        self.cfg = cfg
        self.tracker = FaceTracker(cfg)
        self.segmenter = PersonSegmenter(cfg)
        self.compositor = Compositor(cfg)
        self.hq_engine = HQSwapEngine(cfg)  # 遅延ロード（使うまで重い依存を読まない）
        self._effects: Dict[int, object] = {}
        self._effect_sig: Dict[int, tuple] = {}

    # --- エフェクト解決（ID別 + ランタイム切替に追従） ---
    def _effect_for(self, face_id: int):
        ecfg = self.cfg.effects.get(face_id, self.cfg.default_effect)
        sig = (ecfg.type, ecfg.swap_engine, ecfg.overlay_path, ecfg.swap_face_path)
        if self._effect_sig.get(face_id) != sig:
            self._effects[face_id] = build_effect(ecfg, hq_engine=self.hq_engine)
            self._effect_sig[face_id] = sig
        return self._effects[face_id]

    def _draw_debug(self, frame, faces):
        import cv2
        for f in faces:
            x, y, w, h = f.bbox
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
            cv2.putText(frame, f"ID {f.id}", (x, max(0, y - 8)),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
        hud = (f"bg={self.cfg.background_mode}  fx={self.cfg.default_effect.type}"
               f"  swap={self.cfg.default_effect.swap_engine}  faces={len(faces)}")
        cv2.putText(frame, hud, (10, 28), cv2.FONT_HERSHEY_SIMPLEX,
                    0.7, (0, 220, 255), 2)

    def process_frame(self, frame):
        import cv2
        if self.cfg.mirror:
            frame = cv2.flip(frame, 1)

        faces = self.tracker.process(frame)
        mask = self.segmenter.mask(frame, faces) if self.cfg.background_mode != "keep" else None

        for f in faces:
            eff = self._effect_for(f.id)
            frame = eff.apply(frame, f)

        frame = self.compositor.composite(frame, mask)
        if self.cfg.draw_debug:
            self._draw_debug(frame, faces)
        return frame

    # --- ライブ キー操作 ---
    def _handle_key(self, key) -> bool:
        """戻り値 False で終了."""
        if key in (27, ord('q')):
            return False
        if key == ord('b'):
            order = ["keep", "blur", "color", "image"]
            i = order.index(self.cfg.background_mode) if self.cfg.background_mode in order else 0
            self.cfg.background_mode = order[(i + 1) % len(order)]
        elif key == ord('e'):
            order = ["none", "overlay", "swap", "alter"]
            de = self.cfg.default_effect
            i = order.index(de.type) if de.type in order else 0
            de.type = order[(i + 1) % len(order)]
            self._effect_sig.clear()
        elif key == ord('s'):
            de = self.cfg.default_effect
            de.swap_engine = "hq" if de.swap_engine == "mesh" else "mesh"
            self._effect_sig.clear()
        elif key == ord('d'):
            self.cfg.draw_debug = not self.cfg.draw_debug
        return True

    def run(self):
        src = VideoSource(self.cfg.source_value, self.cfg.capture_width,
                          self.cfg.capture_height, self.cfg.capture_fps)
        size = src.size
        fps = src.fps
        sinks = build_sinks(self.cfg, size, fps)
        has_window = any(s.__class__.__name__ == "WindowSink" for s in sinks)

        print(f"[face_fx] {size[0]}x{size[1]} @ {fps:.0f}fps  outputs={self.cfg.outputs}")
        print("[face_fx] keys: q=quit  b=背景  e=エフェクト  s=swap切替(mesh/hq)  d=debug")

        t0, n = time.time(), 0
        try:
            for frame in src:
                out = self.process_frame(frame)
                for s in sinks:
                    s.write(out)
                n += 1
                if has_window:
                    import cv2
                    key = cv2.waitKey(1) & 0xFF
                    if key != 255 and not self._handle_key(key):
                        break
                if n % 60 == 0:
                    dt = time.time() - t0
                    print(f"[face_fx] {n/dt:5.1f} fps", end="\r")
        finally:
            for s in sinks:
                s.close()
            self.close()

    def close(self):
        self.tracker.close()
        self.segmenter.close()
