"""face_fx CLI.

例:
  # iPhone(Camo の仮想カメラ index=1) を加工して、仮想カメラへ出力
  python -m face_fx --source 1 --width 1920 --height 1080 \
      --outputs virtualcam window --bg blur

  # ステッカーを全員に貼る
  python -m face_fx --source 0 --effect overlay --overlay assets/glasses.png

  # 軽量メッシュワープで別人化（GPU不要）
  python -m face_fx --source 0 --effect swap --swap-engine mesh \
      --swap-face assets/target.jpg --bg color

  # 高品質スワップ（要 insightface + inswapper_128.onnx）
  python -m face_fx --source 0 --effect swap --swap-engine hq \
      --swap-face assets/target.jpg --hq-model models/inswapper_128.onnx

  # 録画済み動画を加工して mp4 出力
  python -m face_fx --source input.mp4 --outputs file --out out.mp4 --bg image \
      --bg-image bg.jpg
"""
from __future__ import annotations

import argparse

from .config import AppConfig, EffectConfig
from .pipeline import Pipeline


def parse_args(argv=None):
    p = argparse.ArgumentParser("face_fx", description="軽量リアルタイム顔トラッキング/エフェクト/背景除去")
    p.add_argument("--source", default="0", help="カメラindex か 動画ファイルパス")
    p.add_argument("--width", type=int, default=1920)
    p.add_argument("--height", type=int, default=1080)
    p.add_argument("--fps", type=int, default=30)
    p.add_argument("--proc-scale", type=float, default=0.5, help="検出の縮小率(軽量化)")
    p.add_argument("--max-faces", type=int, default=4)
    p.add_argument("--no-mirror", action="store_true")

    # 出力
    p.add_argument("--outputs", nargs="+", default=["window"],
                   choices=["window", "file", "virtualcam"])
    p.add_argument("--out", dest="output_path", default=None, help="file 出力先 mp4")
    p.add_argument("--vcam-backend", default=None)

    # 背景
    p.add_argument("--bg", dest="background_mode", default="blur",
                   choices=["keep", "blur", "color", "image", "transparent"])
    p.add_argument("--bg-color", nargs=3, type=int, default=[0, 177, 64], help="BGR")
    p.add_argument("--bg-image", default=None)
    p.add_argument("--keep-all", action="store_true", help="他人も残す(連結成分フィルタ無効)")

    # エフェクト（全員へ既定適用）
    p.add_argument("--effect", default="none", choices=["none", "overlay", "swap", "alter"])
    p.add_argument("--overlay", default=None, help="RGBA PNG パス")
    p.add_argument("--swap-engine", default="mesh", choices=["mesh", "hq"])
    p.add_argument("--swap-face", default=None, help="差し替え元の顔画像")
    p.add_argument("--swap-blend", type=float, default=0.85)
    p.add_argument("--alter-smooth", type=float, default=0.6)

    # HQ
    p.add_argument("--hq-model", default=None, help="inswapper_128.onnx パス")
    p.add_argument("--no-gpu", action="store_true")

    p.add_argument("--debug", action="store_true")
    return p.parse_args(argv)


def build_config(a) -> AppConfig:
    default_effect = EffectConfig(
        type=a.effect,
        swap_engine=a.swap_engine,
        overlay_path=a.overlay,
        swap_face_path=a.swap_face,
        swap_blend=a.swap_blend,
        alter_smooth=a.alter_smooth,
    )
    return AppConfig(
        source=str(a.source),
        capture_width=a.width,
        capture_height=a.height,
        capture_fps=a.fps,
        proc_scale=a.proc_scale,
        max_faces=a.max_faces,
        mirror=not a.no_mirror,
        outputs=a.outputs,
        output_path=a.output_path,
        virtualcam_backend=a.vcam_backend,
        background_mode=a.background_mode,
        background_color=tuple(a.bg_color),
        background_image=a.bg_image,
        keep_only_tracked=not a.keep_all,
        default_effect=default_effect,
        hq_model_path=a.hq_model,
        hq_use_gpu=not a.no_gpu,
        draw_debug=a.debug,
    )


def main(argv=None):
    a = parse_args(argv)
    cfg = build_config(a)
    Pipeline(cfg).run()


if __name__ == "__main__":
    main()
