# -*- coding: utf-8 -*-
"""
Utilities for configuring the application logger.
"""

from __future__ import annotations

import logging
import os
from pathlib import Path

from .constants import APP_ID

LOG_FILENAME = "sishub_log.txt"


def _desktop_candidates() -> list[Path]:
    """Possible Desktop locations for the current user."""
    home = Path.home()
    candidates = [home / "Desktop"]

    userprofile = os.environ.get("USERPROFILE")
    if userprofile:
        candidates.append(Path(userprofile) / "Desktop")

    onedrive = os.environ.get("OneDrive")
    if onedrive:
        candidates.append(Path(onedrive) / "Desktop")

    unique: list[Path] = []
    seen: set[Path] = set()
    for candidate in candidates:
        try:
            key = candidate.resolve()
        except Exception:
            key = candidate.absolute()
        if key in seen:
            continue
        seen.add(key)
        unique.append(candidate)
    return unique


def _resolve_log_path() -> Path:
    for desktop in _desktop_candidates():
        if desktop.exists():
            return desktop / LOG_FILENAME
    # Fall back to the default Desktop location even if it does not exist,
    # so the error message mentions a concrete path.
    return Path.home() / "Desktop" / LOG_FILENAME


def setup_logging() -> logging.Logger:
    log_path = _resolve_log_path()
    if not log_path.exists():
        raise FileNotFoundError(
            f"No se encontro el archivo de log requerido: '{log_path}'. "
            "Crea un archivo de texto vacio con ese nombre en tu Escritorio "
            "y vuelve a abrir la aplicacion."
        )

    logger = logging.getLogger(APP_ID)
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger

    handler = logging.FileHandler(log_path, encoding="utf-8")
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(fmt)
    logger.addHandler(handler)

    console = logging.StreamHandler()
    console.setFormatter(fmt)
    logger.addHandler(console)

    logger.info("Logging iniciado %s", log_path)
    return logger
