import sys
from pathlib import Path


def resource_path(relative_path: str) -> str:
    """
    Resolve paths for bundled execution (PyInstaller onedir/onefile) or source.
    """
    base_path = getattr(sys, "_MEIPASS", None) or Path(__file__).resolve().parent
    return str(Path(base_path) / relative_path)
