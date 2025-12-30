# app/resource_path.py
from __future__ import annotations

import sys
from pathlib import Path


def resource_path(relative_path: str) -> Path:
    """
    PyInstaller EXE içinde: sys._MEIPASS altındaki gömülü dosyaları bulur.
    Normal python çalışmada: proje köküne göre çözer.

    relative_path örn: "app/layouts/isko11.json"
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    else:
        # Bu dosya: app/resource_path.py
        # app/.. => proje kökü
        base = Path(__file__).resolve().parents[1]
    return base / relative_path
