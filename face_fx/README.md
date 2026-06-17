# face_fx — 軽量リアルタイム顔トラッキング / エフェクト / 背景除去

iPhone 等の高画質映像を入力に、**複数人を個別トラッキング**して
顔に合わせた**マスク/オーバーレイ**・**別人化（フェイススワップ）**・**微調整**を行い、
**トラッキング中の人以外（背景・他人）を消す**システム。
土台は Google の **MediaPipe**（軽量・GPU不要・~30fps）で組み、
別人化を最高品質にしたい部分だけ **Deep-Live-Cam 方式（InsightFace inswapper）** を
キーで切替えて使える。

## 特長（要件対応表）

| 要件 | 実装 |
|---|---|
| 人別に認識してその人の顔にマスク | MediaPipe FaceMesh + 重心トラッカで**安定ID**付与 → ID別エフェクト |
| 軽い | 検出は `proc_scale`(既定0.5)で縮小、**合成は原寸** → 速くて高画質 |
| iPhone高画質をそのまま | `--width/--height` で高解像度取り込み、原寸合成で劣化なし |
| 表情を保つ | ライブのメッシュ形状にワープ → 口の開閉・向きが追従 |
| 別人に変える | `swap`（mesh=軽量 / hq=高品質、`s`キーで切替）。**mesh/hq とも人別に別の顔へ差し替え可** |
| 若干顔を変える | `alter`（肌スムージング・明るさ） |
| オーバーレイ | `overlay`（RGBA PNGを目位置に合わせ回転・拡大して貼る） |
| 背景除去（人以外を消す） | MediaPipe Selfie Segmentation + **連結成分で他人も除去** |
| ライブ出力 | プレビュー窓 / mp4 / **仮想カメラ(Zoom/OBS/Meet)** |

## セットアップ

```bash
python3 -m pip install -r face_fx/requirements.txt   # Python 3.10 推奨
# 高品質スワップを使う場合のみ:
#   pip install insightface onnxruntime   # GPUなら onnxruntime-gpu
#   inswapper_128.onnx を入手して --hq-model で指定
```

### iPhone を高画質入力にする（Camo 推奨）
1. iPhone と Mac/PC に **Camo** を入れる（Continuity Camera / EpocCam でも可）
2. Camo が「仮想カメラ」として OS に現れる → そのカメラ index を `--source` に指定
3. `--width 1920 --height 1080`（対応端末なら 3840x2160 も）で高画質取り込み

### 仮想カメラ出力
`--outputs virtualcam` で「face_fx カメラ」として Zoom / OBS / Meet から選べる。
macOS は **OBS Virtual Camera**、Windows は OBS 等の仮想カメラドライバが必要。

## 使い方

```bash
# iPhone(Camo=index1) を加工 → 仮想カメラ＋プレビュー、背景ぼかし
python -m face_fx --source 1 --width 1920 --height 1080 \
    --outputs virtualcam window --bg blur

# 全員に眼鏡ステッカー
python -m face_fx --source 0 --effect overlay --overlay assets/glasses.png

# 軽量メッシュワープで別人化（GPU不要・表情保持）
python -m face_fx --source 0 --effect swap --swap-engine mesh \
    --swap-face assets/target.jpg --bg color

# 高品質スワップ（要 insightface + inswapper_128.onnx）
python -m face_fx --source 0 --effect swap --swap-engine hq \
    --swap-face assets/target.jpg --hq-model models/inswapper_128.onnx

# 録画済み動画を加工して mp4 出力（背景を画像に差し替え）
python -m face_fx --source input.mp4 --outputs file --out out.mp4 \
    --bg image --bg-image assets/bg.jpg
```

## ライブ操作キー（プレビュー窓フォーカス時）

| キー | 動作 |
|---|---|
| `q` / `Esc` | 終了 |
| `b` | 背景モード切替 keep→blur→color→image |
| `e` | エフェクト切替 none→overlay→swap→alter |
| `s` | スワップエンジン切替 **mesh ⇄ hq** |
| `d` | デバッグ表示（ID・FPS） |

## 構成

```
face_fx/
  app.py / __main__.py   CLI
  config.py              設定(AppConfig / EffectConfig)
  capture.py             映像入力(カメラ/動画/iPhone-Camo)
  tracker.py             FaceMesh + 安定IDトラッカ(複数人)
  segmenter.py           人セグメンテーション(他人も除去)
  effects/
    overlay.py           ステッカー/マスク貼り付け
    mesh_warp.py         軽量フェイススワップ(GPU不要)
    alter.py             顔の微調整(スムージング)
    swap_hq.py           高品質スワップ(InsightFace/Deep-Live-Cam)
  compositor.py          背景合成(原寸)
  output.py              窓/mp4/仮想カメラ
  pipeline.py            全体制御 + キー操作
```

## パフォーマンスのコツ
- `--proc-scale 0.4` 程度まで下げると検出が軽くなる（合成画質は不変）。
- `--max-faces` を実際の人数に合わせると速い。
- 高品質 `hq` は GPU 推奨（NVIDIA=CUDA / Apple Silicon=CoreML）。CPUだと重い。
- mesh スワップ・オーバーレイ・背景除去は CPU だけで実用速度。

## ライセンス/注意
顔差し替え機能は、本人の同意・適法な用途でのみ使用すること。
MediaPipe(Apache-2.0)、InsightFace/inswapper はモデルの利用規約に従う。
