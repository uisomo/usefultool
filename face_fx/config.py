"""設定データクラス群."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple


@dataclass
class EffectConfig:
    """1人(=1つの安定ID)に割り当てるエフェクト設定.

    type:
        - "none"     : 何もしない（素通し）
        - "overlay"  : PNG(RGBA)を顔ランドックに合わせて貼る
        - "swap"     : 別人の顔画像を顔メッシュにワープして差し替え
        - "alter"    : 肌スムージング等で“若干”顔を変える
    """
    type: str = "none"
    # swap のとき使うエンジン: "mesh"(軽量) / "hq"(Deep-Live-Cam系・高品質)
    swap_engine: str = "mesh"
    # overlay 用: 貼り付けるRGBA画像パス
    overlay_path: Optional[str] = None
    # overlay 用: アンカーにするランドマークindex（左右の基準点）と縦位置基準
    overlay_anchor: Tuple[int, int] = (33, 263)  # 左右の目尻（width/回転の基準）
    overlay_scale: float = 2.4   # 顔幅に対する倍率
    overlay_y_offset: float = 0.0  # 顔幅に対する縦オフセット（+で下）
    # swap 用: 差し替え元の顔画像パス
    swap_face_path: Optional[str] = None
    swap_blend: float = 0.85     # 0..1 差し替えの不透明度（seamlessClone後の合成）
    # alter 用
    alter_smooth: float = 0.6    # 0..1 肌スムージング強度
    alter_brightness: float = 0.0  # -50..50 明るさ調整


@dataclass
class AppConfig:
    # 入力ソース: 整数=カメラindex, それ以外=動画ファイルパス
    source: str = "0"
    # 取り込み解像度（iPhone等の高画質を活かす。0なら端末既定）
    capture_width: int = 1920
    capture_height: int = 1080
    capture_fps: int = 30

    # 検出処理を行う縮小スケール（軽量化の肝。合成は常に原寸）
    proc_scale: float = 0.5
    # 何フレームに1回 検出するか（1=毎フレーム）。トラッカで間を補間
    detect_every: int = 1

    # 顔トラッキング
    max_faces: int = 4
    refine_landmarks: bool = True  # 虹彩等の精緻化（少し重い）
    min_detection_confidence: float = 0.5
    min_tracking_confidence: float = 0.5

    # 背景処理
    # "keep" : 何もしない / "blur" / "color" / "image" / "transparent"
    background_mode: str = "blur"
    background_color: Tuple[int, int, int] = (0, 177, 64)  # クロマキー緑(BGR)
    background_image: Optional[str] = None
    background_blur: int = 35  # ぼかしカーネル(奇数)
    # トラッキングしている人だけ残す（他人も背景として消す）
    keep_only_tracked: bool = True
    mask_feather: int = 9      # マスク境界のぼかし(奇数, 羽化)

    # ID -> エフェクト 割り当て。未指定IDには default_effect を適用
    effects: Dict[int, EffectConfig] = field(default_factory=dict)
    default_effect: EffectConfig = field(default_factory=EffectConfig)

    # 出力先（複数指定可）: "window" / "file" / "virtualcam"
    outputs: List[str] = field(default_factory=lambda: ["window"])
    output_path: Optional[str] = None  # outputs に "file" を含む場合の mp4 パス
    virtualcam_backend: Optional[str] = None  # pyvirtualcam backend(任意): "obs" 等
    mirror: bool = True  # 自撮りミラー表示
    draw_debug: bool = False  # ランドマーク/ID描画

    # 高品質スワップ(Deep-Live-Cam / InsightFace inswapper)用
    hq_model_path: Optional[str] = None   # inswapper_128.onnx のパス
    hq_use_gpu: bool = True                # CUDA/CoreML を試す
    hq_enhance: bool = False               # GFPGAN等で後段強調（重い・任意）

    @property
    def source_value(self):
        """source を camera index(int) か path(str) に解決."""
        s = self.source.strip()
        return int(s) if s.isdigit() else s
