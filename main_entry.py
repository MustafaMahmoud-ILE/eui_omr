#!/usr/bin/env python3
"""Standalone Entry Point for the EUI OMR Engine."""

import sys
import os
from pathlib import Path

# Add project root to sys.path is no longer strictly needed if main is at root,
# but it's good practice for frozen environments.

try:
    from PySide6.QtCore import Qt, QSharedMemory
    from PySide6.QtWidgets import QApplication, QMessageBox
    from src.ui.main_window import MainWindow
    from src.utils.paths import get_config_path
except ImportError as e:
    print(f"Error loading modules: {e}")
    sys.exit(1)

def main():
    app = QApplication(sys.argv)
    
    # Enforce Single-Instance Application
    shared_mem = QSharedMemory("EUI_OMR_Engine_Unique_Instance_Lock")
    if shared_mem.attach():
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Already Running")
        msg.setText("The EUI OMR Engine is already open.")
        msg.setInformativeText("Please check your taskbar. Running multiple instances is blocked to prevent data corruption.")
        msg.exec()
        sys.exit(0)
    shared_mem.create(1)
    
    # In One-Dir mode, we check for config.json next to the EXE
    config_path = get_config_path("config.json")

    window = MainWindow(config_path)
    window.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()
