import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont, QFontDatabase

# Force offscreen rendering
os.environ["QT_QPA_PLATFORM"] = "offscreen"

# Add project root to path
sys.path.insert(0, os.getcwd())

from src.ui.main_window import MainWindow

def take_snapshots():
    app = QApplication(sys.argv)
    
    # Try to set a standard font
    app.setFont(QFont("Arial", 10))
    print(f"Font set to: {app.font().family()}")
    
    # We need a dummy config path
    config_path = "config.json"
    if not os.path.exists(config_path):
        with open(config_path, "w") as f: f.write("{}")
        
    window = MainWindow(config_path)
    window.resize(1280, 800)
    window.show()
    
    # Snapshot 1: Setup View
    window.stack.setCurrentIndex(0)
    window.grab().save("user_manual/screenshot_setup.png")
    print("Saved setup screenshot")
    
    # Snapshot 2: Dashboard (Simulate having results)
    window.stack.setCurrentIndex(2)
    window.grab().save("user_manual/screenshot_dashboard.png")
    print("Saved dashboard screenshot")
    
    app.quit()

if __name__ == "__main__":
    if not os.path.exists("user_manual"):
        os.makedirs("user_manual")
    take_snapshots()
