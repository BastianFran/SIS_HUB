# -*- coding: utf-8 -*-
"""
Created on Thu Oct 16 13:09:25 2025

@author: bbruna
"""

from __future__ import annotations

import importlib.util
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .constants import COLORS, APP_VERSION
from .logging_utils import setup_logging

@dataclass
class PluginContext:
    root: Path
    data_dir: Path
    colors: dict
    app_version: str
    logger_name: str

    def logger(self):
        import logging
        return logging.getLogger(self.logger_name)

def load_entrypoint(plugin_dir: Path, entrypoint: str) -> Callable:
    module_name, func_name = entrypoint.split(":")
    module_path = plugin_dir / "src" / f"{module_name}.py"
    if not module_path.exists():
        raise FileNotFoundError(f"Entrypoint no encontrado: {module_path}")
 
    unique_name = f"sishub_plugin_{plugin_dir.name}_{module_name}"
 
    # >>> AÑADIR ESTO <<<
    src_path = plugin_dir / "src"
    added = False
    if str(src_path) not in sys.path:
        sys.path.insert(0, str(src_path))
        added = True
    # <<<
 
    spec = importlib.util.spec_from_file_location(unique_name, module_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"No se pudo cargar el módulo: {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[unique_name] = module
 
    try:
        spec.loader.exec_module(module)  # type: ignore[attr-defined]
    finally:
        # >>> AÑADIR ESTO <<<
        if added:
            try:
                sys.path.remove(str(src_path))
            except ValueError:
                pass
        # <<<
 
    func = getattr(module, func_name, None)
    if func is None:
        raise AttributeError(f"Función '{func_name}' no encontrada en {module_path}")
    return func

def build_context(app_root: Path, logger_name: str) -> PluginContext:
    from .constants import user_data_dir
    return PluginContext(
        root=app_root,
        data_dir=user_data_dir(),
        colors=COLORS,
        app_version=APP_VERSION,
        logger_name=logger_name,
    )
