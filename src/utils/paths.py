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
    """ 
    Priority Path Logic:
    1. Beside the EXE (User's external config)
    2. Internal Bundle (Default fallback)
    3. Current Working Directory (Dev mode)
    """
    # 1. Path beside the EXE (External override)
    if hasattr(sys, 'frozen'):
        beside_exe = Path(sys.executable).parent / filename
        if beside_exe.exists():
            return str(beside_exe)
            
    # 2. Path inside the PyInstaller bundle (Bundled default)
    try:
        internal_bundle = Path(sys._MEIPASS) / filename
        if internal_bundle.exists():
            return str(internal_bundle)
    except Exception:
        pass
        
    # 3. Fallback to CWD
    return str(Path.cwd() / filename)
