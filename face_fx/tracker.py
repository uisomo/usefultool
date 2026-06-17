"""複数人の顔トラッキング（MediaPipe FaceMesh）＋ 安定ID付与.

MediaPipe FaceMesh は 1フレームに最大 max_faces 人を検出し、478点(refine時)の
ランドマークを返す。検出順は毎フレーム入れ替わり得るので、ここで重心ベースの
簡易トラッカを噛ませて「同じ人 = 同じID」を維持する（=人別エフェクト割当の土台）。

軽量化のため検出は proc_scale で縮小した画像に対して行い、ランドマーク座標は
原寸に戻して返す（合成は常に原寸 → 画質を落とさない）。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional

import numpy as np

try:
    import mediapipe as mp
except ImportError:  # 実行時に分かりやすく失敗させる
    mp = None


@dataclass
class Face:
    id: int
    landmarks: np.ndarray          # (N,2) float32 原寸ピクセル座標
    bbox: tuple                    # (x, y, w, h) 原寸ピクセル
    centroid: np.ndarray = field(default=None)  # (2,)

    def __post_init__(self):
        if self.centroid is None:
            self.centroid = self.landmarks.mean(axis=0)


class _CentroidTracker:
    """重心の最近傍マッチングで安定IDを割り当てる軽量トラッカ."""

    def __init__(self, max_distance_ratio: float = 0.15, max_missing: int = 8):
        self.next_id = 0
        self.tracks = {}      # id -> centroid
        self.missing = {}     # id -> 連続未検出フレーム数
        self.max_distance_ratio = max_distance_ratio
        self.max_missing = max_missing

    def update(self, centroids: List[np.ndarray], frame_diag: float) -> List[int]:
        max_dist = self.max_distance_ratio * frame_diag
        assigned = [None] * len(centroids)

        # 既存トラックを近い順に割当
        unmatched_tracks = set(self.tracks.keys())
        pairs = []
        for i, c in enumerate(centroids):
            for tid, tc in self.tracks.items():
                pairs.append((np.linalg.norm(c - tc), i, tid))
        pairs.sort(key=lambda p: p[0])

        used_det = set()
        for dist, i, tid in pairs:
            if i in used_det or tid not in unmatched_tracks:
                continue
            if dist > max_dist:
                continue
            assigned[i] = tid
            self.tracks[tid] = centroids[i]
            self.missing[tid] = 0
            used_det.add(i)
            unmatched_tracks.discard(tid)

        # 未割当の検出 → 新規ID
        for i, c in enumerate(centroids):
            if assigned[i] is None:
                tid = self.next_id
                self.next_id += 1
                self.tracks[tid] = c
                self.missing[tid] = 0
                assigned[i] = tid

        # 未検出トラックの寿命管理
        for tid in list(unmatched_tracks):
            self.missing[tid] = self.missing.get(tid, 0) + 1
            if self.missing[tid] > self.max_missing:
                self.tracks.pop(tid, None)
                self.missing.pop(tid, None)

        return assigned


class FaceTracker:
    def __init__(self, cfg):
        if mp is None:
            raise ImportError(
                "mediapipe が見つかりません。`pip install -r face_fx/requirements.txt` を実行してください。"
            )
        self.cfg = cfg
        self.mesh = mp.solutions.face_mesh.FaceMesh(
            static_image_mode=False,
            max_num_faces=cfg.max_faces,
            refine_landmarks=cfg.refine_landmarks,
            min_detection_confidence=cfg.min_detection_confidence,
            min_tracking_confidence=cfg.min_tracking_confidence,
        )
        self.ctracker = _CentroidTracker()

    def process(self, frame_bgr: np.ndarray) -> List[Face]:
        import cv2  # 遅延import（capture側で必須なので存在前提）

        h, w = frame_bgr.shape[:2]
        scale = self.cfg.proc_scale
        if scale != 1.0:
            small = cv2.resize(frame_bgr, (max(1, int(w * scale)), max(1, int(h * scale))))
        else:
            small = frame_bgr
        rgb = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
        rgb.flags.writeable = False
        res = self.mesh.process(rgb)

        faces_lm = []
        centroids = []
        if res.multi_face_landmarks:
            for fl in res.multi_face_landmarks:
                pts = np.array([[lm.x * w, lm.y * h] for lm in fl.landmark], dtype=np.float32)
                faces_lm.append(pts)
                centroids.append(pts.mean(axis=0))

        diag = float(np.hypot(w, h))
        ids = self.ctracker.update(centroids, diag) if centroids else []

        faces = []
        for pts, fid in zip(faces_lm, ids):
            x0, y0 = pts.min(axis=0)
            x1, y1 = pts.max(axis=0)
            bbox = (int(x0), int(y0), int(x1 - x0), int(y1 - y0))
            faces.append(Face(id=fid, landmarks=pts, bbox=bbox))
        return faces

    def close(self):
        self.mesh.close()
