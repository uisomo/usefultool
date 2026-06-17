"""エフェクトの共通インターフェース."""
from __future__ import annotations

import numpy as np


class Effect:
    """1人の顔に対して適用するエフェクトの基底クラス.

    apply は frame(BGR, 原寸) を **その場で / もしくは新規配列で** 加工して返す。
    face には id / landmarks(原寸ピクセル) / bbox が入る。
    """

    def __init__(self, cfg):
        self.cfg = cfg

    def apply(self, frame: np.ndarray, face) -> np.ndarray:
        return frame


class NoneEffect(Effect):
    def apply(self, frame, face):
        return frame
