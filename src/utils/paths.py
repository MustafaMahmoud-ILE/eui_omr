import sys
import os
from pathlib import Path

def get_resource_path(relative_path: str) -> str:
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Development mode
        base_path = os.getcwd()

    return os.path.join(base_path, relative_path)

def get_config_path(filename: str = "config.json") -> str:
    """ Returns the path to a config file that should live beside the EXE/Script, not inside the bundle. """
    if hasattr(sys, 'frozen'):
        # We are running as an EXE
        base_dir = os.path.dirname(sys.executable)
    else:
        # We are running as a script
        base_dir = os.getcwd()
        
    return str(Path(base_dir) / filename)
