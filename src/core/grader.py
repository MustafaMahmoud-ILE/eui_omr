"""
grader.py — Core Grading Pipeline for eui_omr

Takes a raw scanned image, uses `calibrate.py` functions to detect markers and warp it,
then slices the pre-calibrated Regions of Interest (`config.json`) into grids to extract bubble answers.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import cv2
import numpy as np
from numpy.typing import NDArray

# Reuse our excellent warping pipeline
from src.core.calibrate import detect_corners, warp_to_a4, BBox
from src.models.schemas import GradingResult


class OMRGrader:
    def __init__(self, config_path: Path | str, sensitivity: int = 75):
        self.config_path = Path(config_path)
        with open(self.config_path, "r", encoding="utf-8") as fh:
            cfg = json.load(fh)
            
        self.a4_width = cfg["a4_width"]
        self.a4_height = cfg["a4_height"]
        self.config = cfg
        
        self.CHOICES = ["A", "B", "C", "D", "E", "F"]
        self.VERSIONS = ["A", "B", "C", "D", "E", "F"]
        # Sensitivity 100 = 0.0 threshold (extreme)
        # Sensitivity 75 = 0.25 threshold (recommended)
        # Sensitivity 0 = 1.0 threshold (impossible)
        self.FILL_THRESHOLD = max(0.01, min(0.99, (100 - sensitivity) / 100.0))

    def _get_bbox(self, key: str) -> BBox | None:
        val = self.config.get(key)
        if val is None:
            return None
        return BBox(x=val["x"], y=val["y"], w=val["w"], h=val["h"])

    def _get_roi_image(self, image: NDArray[np.uint8], bbox: BBox) -> NDArray[np.uint8]:
        return image[bbox.y:bbox.y + bbox.h, bbox.x:bbox.x + bbox.w]

    def _get_grid_cells(self, roi: NDArray[np.uint8], rows: int, cols: int) -> list[list[float]]:
        h, w = roi.shape[:2]
        cell_h = h / rows
        cell_w = w / cols
        
        fill_percentages = []
        for r in range(rows):
            row_fills = []
            for c in range(cols):
                y_start = int(r * cell_h)
                y_end = int((r + 1) * cell_h)
                x_start = int(c * cell_w)
                x_end = int((c + 1) * cell_w)
                
                # Accuracy Fix: Leave a small 3-pixel margin to ignore bubble borders
                # This is more robust than a percentage for various bubble sizes.
                m = 3
                cell = roi[y_start + m : y_end - m, x_start + m : x_end - m]
                
                if cell.size == 0: # Fallback for tiny ROIs
                    cell = roi[y_start:y_end, x_start:x_end]
                    
                filled_pixels = cv2.countNonZero(cell)
                total_pixels = cell.shape[0] * cell.shape[1]
                ratio = filled_pixels / total_pixels if total_pixels > 0 else 0.0
                row_fills.append(ratio)
            fill_percentages.append(row_fills)
        return fill_percentages

    def process_image(self, image_path: Path | str, page_num: int = 1) -> GradingResult | None:
        image = cv2.imread(str(image_path))
        if image is None:
            raise FileNotFoundError(f"Could not read image at {image_path}")
        return self.process_array(image, page_num)

    def process_array(self, image: NDArray[np.uint8], page_num: int, expected_questions: int = 60, save_dir: Path | None = None) -> GradingResult | None:
        try:
            corners = detect_corners(image)
        except RuntimeError:
            return None # Skip blank/upside down pages with no corners
            
        warped = warp_to_a4(image, corners, self.a4_width, self.a4_height)
        
        # We need the color warped image for showing UI crops natively
        color_warped = warped.copy()
        
        gray = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)
        thresh = cv2.adaptiveThreshold(
            gray, 255, 
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
            cv2.THRESH_BINARY_INV, 
            blockSize=51, C=15
        )
        kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
        thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)

        # Blank page filter (Are there barely any pencil marks/print left?)
        density = cv2.countNonZero(thresh) / (self.a4_width * self.a4_height)
        if density < 0.005:
            return None # Safely skip blank side of duplex paper

        # Extract textual fields
        student_id, id_err, id_roi = self._extract_student_id(thresh, color_warped)
        version, ver_err, ver_roi = self._extract_version(thresh, color_warped)
        answers, q_errors, q_crops = self._extract_answers(thresh, color_warped, expected_questions)
        
        # Get signature crop if defined
        sig_bbox = self._get_bbox("student_name")
        sig_crop = self._get_roi_image(color_warped, sig_bbox) if sig_bbox else None

        result = GradingResult(
            page_number=page_num,
            student_id=student_id,
            version=version,
            answers=answers,
            id_error=id_err,
            version_error=ver_err,
            question_errors=q_errors,
            _id_crop=id_roi,
            _signature_crop=sig_crop,
            _version_crop=ver_roi,
            _question_crops=q_crops
        )

        # --- DISK CACHING LOGIC ---
        if save_dir and save_dir.is_dir():
            # Save ID crop
            if id_roi is not None:
                p = save_dir / f"p{page_num}_id.jpg"
                cv2.imwrite(str(p), id_roi, [cv2.IMWRITE_JPEG_QUALITY, 85])
                result.id_crop_path = p.name
            
            # Save Version crop
            if ver_roi is not None:
                p = save_dir / f"p{page_num}_ver.jpg"
                cv2.imwrite(str(p), ver_roi, [cv2.IMWRITE_JPEG_QUALITY, 85])
                result.version_crop_path = p.name
                
            # Save Signature crop
            if sig_crop is not None:
                p = save_dir / f"p{page_num}_sig.jpg"
                cv2.imwrite(str(p), sig_crop, [cv2.IMWRITE_JPEG_QUALITY, 80])
                result.signature_crop_path = p.name
                
            # Save Question crops
            for q_num, crop in q_crops.items():
                p = save_dir / f"p{page_num}_q{q_num}.jpg"
                cv2.imwrite(str(p), crop, [cv2.IMWRITE_JPEG_QUALITY, 85])
                result.question_crop_paths[q_num] = p.name

            # CRITICAL: Clear in-memory crops after save to release RAM
            result._id_crop = None
            result._version_crop = None
            result._signature_crop = None
            result._question_crops = {}

        return result

    def extract_crops_only(self, image: NDArray[np.uint8], expected_questions: int = 60) -> dict:
        """Lightweight extraction of just the visual ROIs (crops) for UI review."""
        try:
            from .calibrate import detect_corners, warp_to_a4
            corners = detect_corners(image)
            warped = warp_to_a4(image, corners, self.a4_width, self.a4_height)
            color_warped = warped.copy()
            
            gray = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)
            thresh = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                cv2.THRESH_BINARY_INV, blockSize=51, C=15
            )
            
            _, _, id_roi = self._extract_student_id(thresh, color_warped)
            _, _, ver_roi = self._extract_version(thresh, color_warped)
            _, _, q_crops = self._extract_answers(thresh, color_warped, expected_questions)
            
            sig_bbox = self._get_bbox("student_name")
            sig_crop = self._get_roi_image(color_warped, sig_bbox) if sig_bbox else None
            
            return {
                "id_crop": id_roi,
                "version_crop": ver_roi,
                "signature_crop": sig_crop,
                "question_crops": q_crops
            }
        except Exception:
            return {}

    def optimize_sensitivity(self, images: list[NDArray[np.uint8]], expected_questions: int = 60) -> int:
        """Find the optimal sensitivity level by minimizing errors across multiple sample pages."""
        processed_samples = []
        
        # 1. Pre-process all sample images once (Warping + Thresh)
        from .calibrate import detect_corners, warp_to_a4
        for img in images:
            try:
                corners = detect_corners(img)
                warped = warp_to_a4(img, corners, self.a4_width, self.a4_height)
                color_warped = warped.copy()
                
                gray = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)
                thresh = cv2.adaptiveThreshold(
                    gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                    cv2.THRESH_BINARY_INV, blockSize=51, C=15
                )
                kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
                thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)
                
                processed_samples.append((thresh, color_warped))
            except Exception:
                continue # Skip unreadable samples
                
        if not processed_samples:
            return 75 # Fallback to default
            
        # 2. Iterate through sensitivities and aggregate scores
        original_threshold = self.FILL_THRESHOLD
        best_sens = 75
        min_total_score = float('inf')
        
        for sens in range(45, 101, 5):
            self.FILL_THRESHOLD = max(0.01, min(0.99, (100 - sens) / 100.0))
            
            total_score_for_this_sens = 0.0
            
            for thresh, color_warped in processed_samples:
                sid, _, _ = self._extract_student_id(thresh, color_warped)
                ver, _, _ = self._extract_version(thresh, color_warped)
                ans, _, _ = self._extract_answers(thresh, color_warped, expected_questions)
                
                num_conflicts = sid.count("*") + (1 if ver == "*" else 0)
                num_blanks = sid.count("?") + (1 if ver == "?" else 0)
                
                for q in range(1, expected_questions + 1):
                    q_ans = ans.get(q, [])
                    if len(q_ans) == 0:
                        num_blanks += 1
                    elif len(q_ans) > 1:
                        num_conflicts += 1
                
                # Combine into a single score for this page
                total_score_for_this_sens += (num_conflicts * 2.5) + (num_blanks * 1.0)
                
            if total_score_for_this_sens < min_total_score:
                min_total_score = total_score_for_this_sens
                best_sens = sens
            elif total_score_for_this_sens == min_total_score:
                if abs(sens - 75) < abs(best_sens - 75):
                    best_sens = sens
                    
        self.FILL_THRESHOLD = original_threshold
        return best_sens

    def _extract_student_id(self, thresh: NDArray, color_img: NDArray) -> tuple[str, bool, NDArray | None]:
        bbox = self._get_bbox("student_id")
        if not bbox: return ("", True, None)
        
        roi = self._get_roi_image(thresh, bbox)
        color_roi = self._get_roi_image(color_img, bbox)
        grid = self._get_grid_cells(roi, rows=10, cols=8)
        
        id_str = ""
        has_error = False
        for col_idx in range(8):
            filled_rows = []
            for row_idx in range(10):
                if grid[row_idx][col_idx] > self.FILL_THRESHOLD:
                    filled_rows.append(row_idx)
            
            if len(filled_rows) == 1:
                id_str += str(filled_rows[0])
            elif len(filled_rows) == 0:
                id_str += "?" 
                has_error = True
            else:
                id_str += "*" 
                has_error = True
                
        return id_str, has_error, color_roi
        
    def _extract_version(self, thresh: NDArray, color_img: NDArray) -> tuple[str, bool, NDArray | None]:
        bbox = self._get_bbox("version")
        if not bbox: return ("", True, None)
        
        roi = self._get_roi_image(thresh, bbox)
        color_roi = self._get_roi_image(color_img, bbox)
        grid = self._get_grid_cells(roi, rows=6, cols=1)
        
        filled_rows = []
        for row_idx in range(6):
            if grid[row_idx][0] > self.FILL_THRESHOLD:
                filled_rows.append(row_idx)
                
        if len(filled_rows) == 1:
            return self.VERSIONS[filled_rows[0]], False, color_roi
        elif len(filled_rows) == 0:
            return "?", True, color_roi
        else:
            return "*", True, color_roi

    def _extract_answers(self, thresh: NDArray, color_img: NDArray, expected_questions: int) -> tuple[dict[int, list[str]], list[int], dict[int, NDArray]]:
        answers = {}
        errors = []
        crops = {}
        
        col_names = ["questions_col1", "questions_col2", "questions_col3"]
        for col_offset, col_name in enumerate(col_names):
            bbox = self._get_bbox(col_name)
            if not bbox: continue
            
            roi = self._get_roi_image(thresh, bbox)
            grid = self._get_grid_cells(roi, rows=20, cols=6)
            
            row_h_float = bbox.h / 20.0
            
            for row_idx in range(20):
                question_num = (col_offset * 20) + row_idx + 1
                
                # Ignore questions beyond the instructor's configured exam size
                if question_num > expected_questions:
                    continue
                
                filled_cols = []
                for col_idx in range(6):
                    if grid[row_idx][col_idx] > self.FILL_THRESHOLD:
                        filled_cols.append(col_idx)
                
                # Assign chosen letters
                chosen = [self.CHOICES[c] for c in filled_cols]
                answers[question_num] = chosen
                
                # Flag if blank or if multiple are chosen (for manual review override)
                if len(filled_cols) != 1:
                    errors.append(question_num)
                    
                    # Extract a tiny horizontal crop just for this question row
                    # Use float arithmetic to prevent cumulative shift!
                    y_start = bbox.y + int(row_idx * row_h_float)
                    y_end = bbox.y + int((row_idx + 1) * row_h_float)
                    crops[question_num] = color_img[y_start:y_end, bbox.x:bbox.x + bbox.w].copy()

        return answers, errors, crops


if __name__ == "__main__":
    import pprint
    
    if len(sys.argv) < 2:
        print("Usage: python -m src.core.grader path/to/filled_scan.jpg")
        sys.exit(1)
        
    project_root = Path(__file__).resolve().parent.parent.parent
    config_p = project_root / "config.json"
    
    if not config_p.is_file():
        print("Error: config.json not found! Run the calibrate.py tool first.")
        sys.exit(1)
        
    grader = OMRGrader(config_p)
    try:
        target_path = Path(sys.argv[1])
        res = grader.process_image(target_path)
        
        print(f"\n[{target_path.name}] === GRADING EXPORT ===")
        print(f"Student ID : {res.student_id}")
        print(f"Version    : {res.version}")
        print("Answers:")
        pprint.pprint(res.answers, width=80)
        
    except Exception as e:
        print(f"Extraction failed: {e}")
