# -*- coding: utf-8 -*-
"""
Core constants and shared helpers for SIS Hub.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import List

APP_NAME = "SIS Hub"
APP_ID = "sishub"
APP_VERSION = "0.1.0"

# UI
ANCHO = 1024
ALTO = 640
PADDING = 10

# Dark theme palette (default).
COLORS = {
    "bg": "#0f131a",
    "panel": "#171e29",
    "panel_soft": "#1f2937",
    "accent": "#3b82f6",
    "accent_fg": "#ffffff",
    "fg": "#e5e7eb",
    "muted": "#94a3b8",
    "danger": "#f87171",
    "ok": "#34d399",
    "border": "#2b3442",
    "focus": "#60a5fa",
    "input_bg": "#0b1220",
    "input_fg": "#f8fafc",
    "list_select_bg": "#2563eb",
    "list_select_fg": "#ffffff",
}


def user_data_dir() -> Path:
    """Return the user data directory, creating it if necessary."""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("TMP") or os.getcwd()
    path = Path(base) / "SISHub"
    path.mkdir(parents=True, exist_ok=True)
    return path


def app_base_dir() -> Path:
    """
    Main root of the application.
    - Development: project folder.
    - Frozen (PyInstaller): directory where the executable lives.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def resource_roots() -> List[Path]:
    """
    Candidate locations for extra resources (plugins, assets, docs).
    Prioritises the external shared folder, then the PyInstaller temp extract.
    """
    candidates: list[Path] = []
    base = app_base_dir()
    candidates.append(base)

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass))

    package_root = Path(__file__).resolve().parents[1]
    candidates.append(package_root)

    unique: list[Path] = []
    seen = set()
    for candidate in candidates:
        try:
            key = candidate.resolve(strict=False)
        except Exception:
            key = candidate.absolute()
        if key in seen:
            continue
        seen.add(key)
        unique.append(candidate)
    return unique
