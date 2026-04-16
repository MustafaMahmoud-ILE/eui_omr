#!/usr/bin/env python3
"""
calibrate.py — Interactive Calibration Tool for eui_omr
========================================================

Stand-alone script that:
 1.  Loads ``assets/template.jpg``.
 2.  Detects the four corner orientation markers via contour analysis.
 3.  Applies ``cv2.warpPerspective`` to flatten the sheet to a fixed
     A4 resolution (1654 x 2339 px, 200 DPI equivalent).
 4.  Opens the warped image in an interactive OpenCV window and lets the user draw bounding-box ROIs.
 5.  Writes the ROI coordinates to ``config.json``.

Author : EUI OMR Contributors
License: MIT
"""

from __future__ import annotations

import json
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Final

import cv2
import numpy as np
from numpy.typing import NDArray

from src.utils.paths import get_resource_path

# --- Constants ---
A4_WIDTH: Final[int] = 1654
A4_HEIGHT: Final[int] = 2339

MIN_MARKER_AREA_RATIO: Final[float] = 0.0005
MAX_MARKER_AREA_RATIO: Final[float] = 0.02
RECT_ASPECT_LO: Final[float] = 1.4
RECT_ASPECT_HI: Final[float] = 3.0
SQR_ASPECT_LO: Final[float] = 0.6
SQR_ASPECT_HI: Final[float] = 1.4

COLORS: Final[list[tuple[int, int, int]]] = [
    (0, 165, 255), (0, 255, 255), (255, 0, 255),
    (0, 255, 0), (255, 165, 0), (255, 0, 0),
]

ROI_LABELS: Final[list[str]] = [
    "student_name", "version", "student_id",
    "questions_col1", "questions_col2", "questions_col3",
]

ROI_DISPLAY_NAMES: Final[list[str]] = [
    "Student Name Signature", "Version Section", "Student ID Section",
    "Question Column 1 (Q1-20)", "Question Column 2 (Q21-40)", "Question Column 3 (Q41-60)",
]

@dataclass
class BBox:
    x: int; y: int; w: int; h: int

@dataclass
class CalibrationConfig:
    a4_width: int = A4_WIDTH
    a4_height: int = A4_HEIGHT
    student_name: BBox | None = None
    version: BBox | None = None
    student_id: BBox | None = None
    questions_col1: BBox | None = None
    questions_col2: BBox | None = None
    questions_col3: BBox | None = None

    def to_dict(self) -> dict:
        d = {"a4_width": self.a4_width, "a4_height": self.a4_height}
        for label in ROI_LABELS:
            bbox = getattr(self, label)
            d[label] = asdict(bbox) if bbox is not None else None
        return d

@dataclass
class DrawingState:
    roi_index: int = 0
    drawing: bool = False
    start_x: int = 0; start_y: int = 0
    end_x: int = 0; end_y: int = 0
    rois: list[BBox | None] = field(default_factory=lambda: [None] * len(ROI_LABELS))
    scale: float = 1.0

# --- Marker Detection Helpers ---

def _preprocess_for_markers(image: NDArray[np.uint8]) -> NDArray[np.uint8]:
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    thresh = cv2.adaptiveThreshold(
        blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV, blockSize=31, C=15
    )
    return thresh

def _contour_centroid(cnt: NDArray) -> tuple[float, float]:
    m = cv2.moments(cnt)
    if m["m00"] == 0:
        x, y, w, h = cv2.boundingRect(cnt)
        return (x + w / 2.0, y + h / 2.0)
    return (m["m10"] / m["m00"], m["m01"] / m["m00"])

def _contour_aspect_ratio(cnt: NDArray) -> float:
    (_, _), (side_a, side_b), _ = cv2.minAreaRect(cnt)
    if side_a == 0 or side_b == 0: return 0.0
    return max(side_a, side_b) / min(side_a, side_b)

_Candidate = tuple[NDArray, float, float]

def _find_marker_candidates(thresh: NDArray[np.uint8], image_area: int) -> list[_Candidate]:
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    candidates: list[_Candidate] = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if not (image_area * MIN_MARKER_AREA_RATIO < area < image_area * MAX_MARKER_AREA_RATIO): continue
        peri = cv2.arcLength(cnt, True)
        approx = cv2.approxPolyDP(cnt, 0.04 * peri, True)
        if not (4 <= len(approx) <= 8): continue
        hull = cv2.convexHull(cnt)
        hull_area = cv2.contourArea(hull)
        if hull_area == 0 or (area / hull_area) < 0.75: continue
        aspect = _contour_aspect_ratio(cnt)
        candidates.append((cnt, aspect, area))
    return candidates

def detect_corners(image: NDArray[np.uint8]) -> NDArray[np.float32]:
    h, w = image.shape[:2]
    image_area = h * w
    thresh = _preprocess_for_markers(image)
    candidates = _find_marker_candidates(thresh, image_area)
    
    if len(candidates) < 4:
        raise RuntimeError(f"Found only {len(candidates)} markers. Need 4.")

    # Sort candidates by area to get the most prominent ones
    candidates.sort(key=lambda x: x[2], reverse=True)
    top_4 = candidates[:4]
    
    # Extract points
    top_4_sorted_by_aspect = sorted(top_4, key=lambda x: x[1], reverse=True)
    anchor_cand = top_4_sorted_by_aspect[0]
    square_cands = top_4_sorted_by_aspect[1:]
    
    anchor_pt = _contour_centroid(anchor_cand[0])
    square_pts = [_contour_centroid(s[0]) for s in square_cands]
    all_pts = [anchor_pt] + square_pts
    
    # --- Geometric Validation ---
    # Ensure points are distributed well. Calculate area of the quad formed by these points.
    pts_arr = np.array(all_pts, dtype=np.float32)
    quad_hull = cv2.convexHull(pts_arr.reshape(-1, 1, 2))
    quad_area = cv2.contourArea(quad_hull)
    
    # A valid OMR sheet should have markers spanning most of the page.
    # We expect the quad area to be > 30% of the image area.
    if quad_area < (image_area * 0.3):
        raise RuntimeError("Detected markers do not span the page. Possible false positive or blank page.")

    # 2. Determine Orientation based on Anchor position
    xs = [p[0] for p in all_pts]; ys = [p[1] for p in all_pts]
    center_x = sum(xs) / 4.0; center_y = sum(ys) / 4.0
    
    # 3. Map points to TL, TR, BR, BL correctly
    # If anchor is Top-Left (x < center_x, y < center_y) -> Normal
    # If anchor is Bottom-Right (x > center_x, y > center_y) -> 180 Rotated
    # We can use pure geometric sorting relative to the center to identify corners
    # then adjust based on where the Anchor (Rectangle) landed.
    
    # Initial corner sort (standard order)
    # TL: min(x+y), BR: max(x+y), TR: max(x-y), BL: min(x-y)
    def get_corners(pts):
        pts = np.array(pts)
        tl = pts[np.argmin(pts.sum(axis=1))]
        br = pts[np.argmax(pts.sum(axis=1))]
        diff = np.diff(pts, axis=1)
        tr = pts[np.argmin(diff)]
        bl = pts[np.argmax(diff)]
        return tl, tr, br, bl

    tl, tr, br, bl = get_corners(all_pts)
    
    # Rotation Check: If the anchor (rectangle) is NOT at the TL, we rotate!
    is_anchor_tl = (abs(anchor_pt[0] - tl[0]) < 10 and abs(anchor_pt[1] - tl[1]) < 10)
    is_anchor_br = (abs(anchor_pt[0] - br[0]) < 10 and abs(anchor_pt[1] - br[1]) < 10)
    
    if is_anchor_br:
        # 180-degree rotation detected! Swap TL/BR and TR/BL
        return np.array([br, bl, tl, tr], dtype=np.float32)
    
    # Default (or other rotations not yet handled)
    return np.array([tl, tr, br, bl], dtype=np.float32)

def warp_to_a4(image: NDArray[np.uint8], src_pts: NDArray[np.float32], width: int = A4_WIDTH, height: int = A4_HEIGHT) -> NDArray[np.uint8]:
    # True physical centroids in cm on A4 (21 x 29.7)
    ppcm_x = width / 21.0
    ppcm_y = height / 29.7
    tl_cm = (1.2 + 0.7, 1.0 + 0.35)
    tr_cm = (21.0 - 1.2 - 0.35, 1.0 + 0.35)
    br_cm = (21.0 - 1.2 - 0.35, 29.7 - 1.0 - 0.35)
    bl_cm = (1.2 + 0.35, 29.7 - 1.0 - 0.35)
    
    dst_pts = np.array([
        [tl_cm[0] * ppcm_x, tl_cm[1] * ppcm_y],
        [tr_cm[0] * ppcm_x, tr_cm[1] * ppcm_y],
        [br_cm[0] * ppcm_x, br_cm[1] * ppcm_y],
        [bl_cm[0] * ppcm_x, bl_cm[1] * ppcm_y]
    ], dtype=np.float32)

    matrix = cv2.getPerspectiveTransform(src_pts, dst_pts)
    return cv2.warpPerspective(image, matrix, (width, height))

# --- Interactive Logic ---

def _mouse_callback(event, x, y, flags, state: DrawingState):
    if event == cv2.EVENT_LBUTTONDOWN:
        state.drawing = True
        state.start_x, state.start_y = x, y
        state.end_x, state.end_y = x, y
    elif event == cv2.EVENT_MOUSEMOVE and state.drawing:
        state.end_x, state.end_y = x, y
    elif event == cv2.EVENT_LBUTTONUP:
        state.drawing = False
        fx1 = int(round(min(state.start_x, state.end_x) * state.scale))
        fy1 = int(round(min(state.start_y, state.end_y) * state.scale))
        fx2 = int(round(max(state.start_x, state.end_x) * state.scale))
        fy2 = int(round(max(state.start_y, state.end_y) * state.scale))
        if (fx2 - fx1) > 5 and (fy2 - fy1) > 5:
            state.rois[state.roi_index] = BBox(x=fx1, y=fy1, w=fx2 - fx1, h=fy2 - fy1)
            state.roi_index += 1

def _render_overlay(base: NDArray[np.uint8], state: DrawingState) -> NDArray[np.uint8]:
    canvas = base.copy()
    inv_scale = 1.0 / state.scale
    for i, bbox in enumerate(state.rois):
        if bbox is None: continue
        dx, dy, dw, dh = int(bbox.x * inv_scale), int(bbox.y * inv_scale), int(bbox.w * inv_scale), int(bbox.h * inv_scale)
        color = COLORS[i % len(COLORS)]
        cv2.rectangle(canvas, (dx, dy), (dx + dw, dy + dh), color, 2)
        cv2.putText(canvas, ROI_DISPLAY_NAMES[i], (dx + 4, dy + 18), cv2.FONT_HERSHEY_SIMPLEX, 0.5, color, 1)
    if state.drawing and state.roi_index < len(ROI_LABELS):
        color = COLORS[state.roi_index % len(COLORS)]
        cv2.rectangle(canvas, (state.start_x, state.start_y), (state.end_x, state.end_y), color, 1)
    return canvas

def main(path_str: str | None = None):
    template_path = Path(path_str) if path_str else Path("assets/template.jpg")
    if not template_path.exists():
        sys.exit(f"Template not found: {template_path}")
        
    if template_path.suffix.lower() == ".pdf":
        import fitz
        doc = fitz.open(str(template_path))
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)[:, :, ::-1].copy()
    else:
        image = cv2.imread(str(template_path))

    corners = detect_corners(image)
    warped = warp_to_a4(image, corners)
    
    scale = 800.0 / A4_HEIGHT
    disp_img = cv2.resize(warped, (int(A4_WIDTH * scale), 800))
    state = DrawingState(scale=1.0 / scale)
    
    win_name = "Calibration Tool"
    cv2.namedWindow(win_name)
    cv2.setMouseCallback(win_name, _mouse_callback, state)
    
    while True:
        canvas = _render_overlay(disp_img, state)
        cv2.imshow(win_name, canvas)
        key = cv2.waitKey(1) & 0xFF
        if key == ord("q"): break
        if key == ord("s") and state.roi_index >= len(ROI_LABELS):
            cfg = CalibrationConfig(*state.rois)
            path = get_resource_path("config.json")
            with open(path, "w") as f: json.dump(cfg.to_dict(), f, indent=2)
            print("Config saved.")
            break
    cv2.destroyAllWindows()

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
