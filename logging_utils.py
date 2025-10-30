# -*- coding: utf-8 -*-
"""
Created on Thu Oct 16 13:08:14 2025

@author: bbruna
"""

from __future__ import annotations
import logging
import logging.handlers
from pathlib import Path
from .constants import user_data_dir, APP_ID

def setup_logging() -> logging.Logger:
    log_dir = user_data_dir() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"{APP_ID}.log"

    logger = logging.getLogger(APP_ID)
    logger.setLevel(logging.INFO)

    # Avoid duplicate handlers if called twice
    if logger.handlers:
        return logger

    handler = logging.handlers.RotatingFileHandler(
        log_path, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
    )
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(fmt)
    logger.addHandler(handler)

    # Also echo minimal logs to console (useful in dev)
    console = logging.StreamHandler()
    console.setFormatter(fmt)
    logger.addHandler(console)

    logger.info("Logging iniciado %s", log_path)
    return logger
