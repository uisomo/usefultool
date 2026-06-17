"""face_fx — 軽量リアルタイム顔トラッキング & エフェクトパイプライン.

主な機能:
- 複数人の顔を個別にトラッキング（安定ID付与）
- 顔メッシュに合わせたオーバーレイ / 別人フェイススワップ / 微調整
- 表情を保持（ライブのメッシュ形状をそのまま使用）
- トラッキング対象“以外”（背景・他人）を消す背景除去
- iPhone 等の高解像度入力をそのまま合成（検出は縮小、合成は原寸）
"""

from .config import AppConfig, EffectConfig
from .pipeline import Pipeline

__all__ = ["AppConfig", "EffectConfig", "Pipeline"]
__version__ = "0.1.0"
