"""軽量フェイススワップ（メッシュワープ方式・GPU不要）.

差し替え元の顔画像から FaceMesh で 468点を一度だけ求め、Delaunay 三角形分割を
作る。各フレームでは、ライブの顔ランドマーク(=表情そのまま)へ三角形単位で
アフィンワープし、seamlessClone で自然に合成する。
表情・口の開閉・向きはライブ側の形状に従うため“表情を保ったまま別人化”できる。

品質は中程度（テクスチャは元画像由来、陰影は付かない）。
最高品質が必要なら swap_engine="hq"（Deep-Live-Cam系）を使う。
"""
from __future__ import annotations

from typing import List, Optional, Tuple

import numpy as np

from .base import Effect

# FaceMesh 標準の頂点数（refine_landmarks の虹彩点を除く）
_BASE_POINTS = 468


class MeshWarpSwap(Effect):
    def __init__(self, cfg):
        super().__init__(cfg)
        self._src_img = None       # (H,W,3)
        self._src_pts = None       # (468,2)
        self._triangles = None     # List[(i,j,k)]
        if cfg.swap_face_path:
            self._load_source(cfg.swap_face_path)

    def _load_source(self, path):
        import cv2
        import mediapipe as mp

        img = cv2.imread(path)
        if img is None:
            raise FileNotFoundError(f"swap_face 画像を読めません: {path}")
        h, w = img.shape[:2]
        with mp.solutions.face_mesh.FaceMesh(
            static_image_mode=True, max_num_faces=1, refine_landmarks=False,
            min_detection_confidence=0.5,
        ) as fm:
            res = fm.process(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
        if not res.multi_face_landmarks:
            raise ValueError(f"swap_face 画像から顔を検出できません: {path}")
        lm = res.multi_face_landmarks[0].landmark
        pts = np.array([[p.x * w, p.y * h] for p in lm[:_BASE_POINTS]], dtype=np.float32)
        self._src_img = img
        self._src_pts = pts
        self._triangles = self._delaunay_triangles(pts, (w, h))

    @staticmethod
    def _delaunay_triangles(points, size) -> List[Tuple[int, int, int]]:
        import cv2
        w, h = size
        rect = (0, 0, w, h)
        subdiv = cv2.Subdiv2D(rect)
        for p in points:
            subdiv.insert((float(np.clip(p[0], 0, w - 1)), float(np.clip(p[1], 0, h - 1))))
        tri_list = subdiv.getTriangleList()

        # 頂点座標 → index への対応（最近傍）
        def idx_of(pt):
            d = np.hypot(points[:, 0] - pt[0], points[:, 1] - pt[1])
            j = int(d.argmin())
            return j if d[j] < 2.0 else -1

        triangles = []
        for t in tri_list:
            ia = idx_of((t[0], t[1]))
            ib = idx_of((t[2], t[3]))
            ic = idx_of((t[4], t[5]))
            if min(ia, ib, ic) >= 0:
                triangles.append((ia, ib, ic))
        return triangles

    def apply(self, frame, face):
        if self._src_img is None or self._triangles is None:
            return frame
        import cv2

        dst_pts = face.landmarks[:_BASE_POINTS]
        if len(dst_pts) < _BASE_POINTS:
            return frame

        warped = np.zeros_like(frame)
        for (i, j, k) in self._triangles:
            self._warp_triangle(self._src_img, warped,
                                self._src_pts[[i, j, k]], dst_pts[[i, j, k]])

        # 合成領域（凸包）でマスクを作り seamlessClone
        hull = cv2.convexHull(dst_pts.astype(np.int32))
        mask = np.zeros(frame.shape[:2], np.uint8)
        cv2.fillConvexPoly(mask, hull, 255)
        x, y, w, h = cv2.boundingRect(hull)
        if w == 0 or h == 0:
            return frame
        center = (x + w // 2, y + h // 2)
        try:
            out = cv2.seamlessClone(warped, frame, mask, center, cv2.NORMAL_CLONE)
        except cv2.error:
            # seamlessClone が境界条件で失敗する場合は単純アルファ合成
            out = frame.copy()
            m3 = (mask[:, :, None] > 0)
            out = np.where(m3, warped, out)

        blend = float(self.cfg.swap_blend)
        if blend >= 1.0:
            return out
        return cv2.addWeighted(out, blend, frame, 1 - blend, 0)

    @staticmethod
    def _warp_triangle(src_img, dst_img, t_src, t_dst):
        import cv2
        r1 = cv2.boundingRect(t_src.astype(np.float32))
        r2 = cv2.boundingRect(t_dst.astype(np.float32))
        x1, y1, w1, h1 = r1
        x2, y2, w2, h2 = r2
        if w1 <= 0 or h1 <= 0 or w2 <= 0 or h2 <= 0:
            return

        t1 = t_src - [x1, y1]
        t2 = t_dst - [x2, y2]

        src_crop = src_img[y1:y1 + h1, x1:x1 + w1]
        if src_crop.size == 0:
            return
        M = cv2.getAffineTransform(t1.astype(np.float32), t2.astype(np.float32))
        warped = cv2.warpAffine(src_crop, M, (w2, h2), flags=cv2.INTER_LINEAR,
                                borderMode=cv2.BORDER_REFLECT_101)

        mask = np.zeros((h2, w2), np.uint8)
        cv2.fillConvexPoly(mask, t2.astype(np.int32), 255)

        roi = dst_img[y2:y2 + h2, x2:x2 + w2]
        if roi.shape[:2] != warped.shape[:2]:
            return
        m3 = mask[:, :, None] > 0
        roi[:] = np.where(m3, warped, roi)
