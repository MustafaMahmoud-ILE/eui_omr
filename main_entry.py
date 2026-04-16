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

import traceback

def global_exception_handler(exctype, value, tb):
    """Global hook to catch any unhandled exceptions and show a user-friendly dialog."""
    error_msg = "".join(traceback.format_exception(exctype, value, tb))
    print(error_msg, file=sys.stderr)
    
    # Try to show a GUI message box if QApplication exists
    if QApplication.instance():
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setWindowTitle("Application Error")
        msg.setText("A critical error has occurred and the application must close.")
        msg.setDetailedText(error_msg)
        msg.setInformativeText("Please take a screenshot of this error and report it to academic support.")
        msg.exec()
    sys.exit(1)

def main():
    # Install the global exception hook
    sys.excepthook = global_exception_handler
    
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

    try:
        window = MainWindow(config_path)
        window.show()
        sys.exit(app.exec())
    except Exception:
        # Catch errors during window initialization as well
        global_exception_handler(*sys.exc_info())

if __name__ == "__main__":
    main()
