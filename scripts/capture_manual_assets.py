import sys
import os
from pathlib import Path
from PySide6.QtWidgets import QApplication, QTableWidgetItem
from PySide6.QtGui import QFont, QFontDatabase, QColor
from PySide6.QtCore import Qt

# Force offscreen rendering
os.environ["QT_QPA_PLATFORM"] = "offscreen"

# Add project root to path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.ui.main_window import MainWindow, ReviewModal
from src.models.schemas import GradingResult

def capture_all():
    app = QApplication(sys.argv)
    
    # 1. Load Arial for clean headless text
    font_path = "C:/Windows/Fonts/arial.ttf"
    font_id = QFontDatabase.addApplicationFont(font_path)
    if font_id != -1:
        family = QFontDatabase.applicationFontFamilies(font_id)[0]
        app.setFont(QFont(family, 10))
    
    # Ensure config exists
    config_path = PROJECT_ROOT / "config.json"
    if not config_path.exists():
        with open(config_path, "w") as f: f.write("{}")
        
    window = MainWindow(str(config_path))
    window.resize(1280, 800)
    window.show()
    
    output_dir = PROJECT_ROOT / "user_guide"
    output_dir.mkdir(exist_ok=True)
    # Create Dummy Crops for Review Modal using REAL image data
    crops_dir = PROJECT_ROOT / "crops"
    crops_dir.mkdir(exist_ok=True)
    import cv2
    import numpy as np
    
    # Load actual test scan to extract realistic looking crops
    real_scan_path = str(PROJECT_ROOT / "assets" / "test_scan_1.jpg")
    real_img = cv2.imread(real_scan_path)
    
    if real_img is not None:
        # Extract generic sections that look like bubbles
        # ID section crop (approximate top-right)
        real_id = real_img[100:400, 800:1100] if real_img.shape[0] > 400 else real_img
        # Version crop (approximate top-left)
        real_ver = real_img[100:200, 100:400] if real_img.shape[0] > 200 else real_img 
        # Question crop (approximate middle)
        real_q = real_img[500:600, 200:700] if real_img.shape[0] > 600 else real_img
        
        cv2.imwrite(str(crops_dir / "dummy_id.png"), real_id)
        cv2.imwrite(str(crops_dir / "dummy_ver.png"), real_ver)
        cv2.imwrite(str(crops_dir / "dummy_q.png"), real_q)
    else:
        # Fallback if image missing
        dummy_img = np.zeros((100, 300, 3), dtype=np.uint8) + 128
        cv2.imwrite(str(crops_dir / "dummy_id.png"), dummy_img)
        cv2.imwrite(str(crops_dir / "dummy_ver.png"), dummy_img)
        cv2.imwrite(str(crops_dir / "dummy_q.png"), dummy_img)
    
    # --- SCENE 1: Home/Setup ---
    window.stack.setCurrentIndex(0)
    window.grab().save(str(output_dir / "real_setup.png"))
    print("Captured real_setup.png")
    
    # --- SCENE 2: Dashboard with Mock Data ---
    mock_results = [
        GradingResult(
            page_number=1, student_id="2023001", version="A",
            answers={1: ["A"], 2: ["B"], 3: ["C"], 4: ["D"], 5: ["A"]},
            id_error=False, version_error=False, question_errors=[],
            id_crop_path="dummy_id.png"
        ),
        GradingResult(
            page_number=2, student_id="202300*", version="B",
            answers={1: ["A"], 2: ["*"], 3: ["C"], 4: ["D"], 5: ["?"]},
            id_error=True, version_error=False, question_errors=[2, 5],
            id_crop_path="dummy_id.png",
            question_crop_paths={2: "dummy_q.png", 5: "dummy_q.png"}
        ),
        GradingResult(
            page_number=3, student_id="2023003", version="?",
            answers={1: ["A"], 2: ["B"], 3: ["C"], 4: ["D"], 5: ["A"]},
            id_error=False, version_error=True, question_errors=[],
            version_crop_path="dummy_ver.png"
        )
    ]
    
    # Manually populate the table
    window.stack.setCurrentIndex(2) # Show dashboard
    table = window.table # Corrected attribute name
    table.setRowCount(0)
    for res in mock_results:
        row = table.rowCount()
        table.insertRow(row)
        
        id_item = QTableWidgetItem(res.student_id)
        if res.id_error: id_item.setForeground(QColor("#ff4d4d"))
        table.setItem(row, 0, id_item)
        
        ver_item = QTableWidgetItem(res.version)
        if res.version_error: ver_item.setForeground(QColor("#ff4d4d"))
        table.setItem(row, 1, ver_item)
        
        table.setItem(row, 2, QTableWidgetItem("85%"))
        
        err_text = []
        if res.id_error: err_text.append("ID")
        if res.version_error: err_text.append("Ver")
        if res.question_errors: err_text.append(f"Q:{len(res.question_errors)}")
        table.setItem(row, 3, QTableWidgetItem(", ".join(err_text) if err_text else "Clean"))

    window.grab().save(str(output_dir / "real_dashboard.png"))
    print("Captured real_dashboard.png")
    
    # --- SCENE 3: Review Modal ---
    # Path trick: ReviewModal looks for Path(pdf_path).parent.parent / "crops"
    # So if we provide "D:/eui_omr/dummy/test.pdf", it looks in "D:/eui_omr/crops"
    dummy_pdf_path = str(PROJECT_ROOT / "dummy_folder" / "mock.pdf")
    modal = ReviewModal(window, mock_results[1], dummy_pdf_path, 5, str(config_path))
    modal.resize(1000, 700)
    modal.show()
    modal.grab().save(str(output_dir / "real_review.png"))
    print("Captured real_review.png")
    
    app.quit()

if __name__ == "__main__":
    capture_all()
