import sys
import os
from pathlib import Path

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import fitz
import cv2
import numpy as np
from src.core.calibrate import detect_corners, warp_to_a4

def debug_page_20():
    pdf_path = "test_data_gen/mock_exam_20pages.pdf"
    if not Path(pdf_path).exists():
        print("PDF not found.")
        return

    doc = fitz.open(pdf_path)
    # Page 20 is index 19
    page = doc.load_page(19)
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    if pix.n == 3: img = img[:, :, ::-1].copy()
    doc.close()

    cv2.imwrite("test_data_gen/debug_page_20.jpg", img)
    print("Saved page 20 to test_data_gen/debug_page_20.jpg")

    try:
        corners = detect_corners(img)
        print("Detected corners:", corners)
        
        # Visualize
        for i, pt in enumerate(corners):
            cv2.circle(img, (int(pt[0]), int(pt[1])), 20, (0, 0, 255), -1)
            cv2.putText(img, str(i), (int(pt[0]), int(pt[1])), cv2.FONT_HERSHEY_SIMPLEX, 2, (0, 255, 0), 5)
        
        cv2.imwrite("test_data_gen/debug_page_20_corners.jpg", img)
        
        # Try warp
        warped = warp_to_a4(img, corners)
        cv2.imwrite("test_data_gen/debug_page_20_warped.jpg", warped)
        print("Warped image saved.")
        
    except Exception as e:
        print("Error during corner detection/warp:", e)

if __name__ == "__main__":
    debug_page_20()
