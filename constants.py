# -*- coding: utf-8 -*-
"""
Created on Thu Oct 16 13:10:16 2025

@author: bbruna
"""

from __future__ import annotations

APP_NAME = "SIS Hub"
APP_ID = "sishub"
APP_VERSION = "0.1.0"

# UI
ANCHO = 1024
ALTO = 640
PADDING = 10

# Paleta simple y de alto contraste
COLORS = {
    "bg": "#2a2b2a",          # air superiority blue
    "panel": "#706C61",      # gray-900
    "accent": "#2563eb",      # blue-600
    "accent_fg": "#ffffff",
    "fg": "#e5e7eb",          # gray-200
    "muted": "#9ca3af",       # gray-400
    "danger": "#ef4444",      # red-500
    "ok": "#10b981",          # green-500
}

# Carpeta de datos del usuario (logs, cache)
import os
from pathlib import Path

def user_data_dir() -> Path:
    # Windows: %LOCALAPPDATA%\SISHub
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("TMP") or os.getcwd()
    path = Path(base) / "SISHub"
    path.mkdir(parents=True, exist_ok=True)
    return path
