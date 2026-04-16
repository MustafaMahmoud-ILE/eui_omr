import sys
import os
import time
from pathlib import Path
import numpy as np
import cv2
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTimer, Qt

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.ui.main_window import MainWindow
from src.core.project import ProjectManager
from src.models.schemas import GradingResult

def capture_uis():
    app = QApplication(sys.argv)
    
    # Ensure assets/screenshots exists
    screenshot_dir = PROJECT_ROOT / "assets" / "screenshots"
    screenshot_dir.mkdir(parents=True, exist_ok=True)
    
    window = MainWindow()
    window.show()
    
    # 1. Capture Welcome Screen
    window.stack.setCurrentIndex(0)
    QApplication.processEvents()
    window.grab().save(str(screenshot_dir / "01_welcome.png"))
    print("Captured Welcome Screen")
    
    # 2. Setup a Mock Project for better screenshots
    temp_proj = PROJECT_ROOT / "temp_screenshot_project"
    temp_proj.mkdir(exist_ok=True)
    pm = ProjectManager(temp_proj)
    pm.create_project("Academic Finals 2026")
    window.pm = pm
    window._transition_to_setup()
    
    # 3. Capture Setup View
    QApplication.processEvents()
    window.grab().save(str(screenshot_dir / "02_setup.png"))
    print("Captured Setup View")
    
    # 4. Mock some results for the Dashboard
    mock_results = []
    for i in range(1, 11):
        res = GradingResult(
            page_number=i,
            student_id=f"202300{i}",
            version=f"{chr(64+ (i % 3 + 1))}",
            answers={q: ["A"] for q in range(1, 61)},
            id_error=False,
            version_error=False,
            question_errors=[]
        )
        # Add some errors to show the 'Needs Review' state
        if i == 3:
            res.id_error = True
            res.student_id = "2023*03"
        if i == 7:
            res.question_errors = [12, 45]
            
        mock_results.append(res)
        
    window.results_data = mock_results
    window.stack.setCurrentIndex(2)
    window._refresh_table = lambda: None # Avoid actual refresh logic issues
    for res in mock_results:
        window._update_table_row(res)
    window._update_stats()
    
    QApplication.processEvents()
    window.grab().save(str(screenshot_dir / "03_dashboard.png"))
    print("Captured Dashboard View")
    
    # 5. Capture a 'Review' Modal (Manual intervention simulation)
    # We can't easily capture a modal in the same script without a second event loop,
    # but we can try to grab the window while the modal is meant to be up.
    # For now, these 3 are the most important.
    
    print(f"\nAll screenshots saved to {screenshot_dir}")
    # Cleanup temp project
    import shutil
    try:
        # shutil.rmtree(temp_proj) # Keep it for a bit if debugging
        pass
    except:
        pass
        
    window.close()
    sys.exit(0)

if __name__ == "__main__":
    # We use a timer to run the capture once the event loop starts
    QTimer.singleShot(500, capture_uis)
    capture_uis()
