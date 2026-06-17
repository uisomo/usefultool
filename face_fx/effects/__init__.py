"""顔エフェクト群."""
from .base import Effect, NoneEffect
from .overlay import OverlayEffect
from .mesh_warp import MeshWarpSwap
from .alter import AlterEffect

__all__ = ["Effect", "NoneEffect", "OverlayEffect", "MeshWarpSwap", "AlterEffect", "build_effect"]


def build_effect(ecfg, hq_engine=None):
    """EffectConfig からエフェクトインスタンスを生成.

    hq_engine: 高品質スワップ用の共有エンジン（あれば swap_engine="hq" で使用）。
    """
    t = ecfg.type
    if t == "overlay":
        return OverlayEffect(ecfg)
    if t == "swap":
        if ecfg.swap_engine == "hq" and hq_engine is not None:
            return hq_engine.make_effect(ecfg)
        return MeshWarpSwap(ecfg)
    if t == "alter":
        return AlterEffect(ecfg)
    return NoneEffect(ecfg)
