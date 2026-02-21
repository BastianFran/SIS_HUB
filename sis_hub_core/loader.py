# -*- coding: utf-8 -*-
"""
Created on Thu Oct 16 13:09:25 2025

@author: bbruna
"""

from __future__ import annotations

import hashlib
import importlib
import importlib.util
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .constants import APP_VERSION, COLORS


def _normalized_path(path: Path) -> Path:
    try:
        return path.resolve()
    except Exception:
        return Path(str(path))


def _same_path(left: str | Path, right: Path) -> bool:
    try:
        return _normalized_path(Path(left)) == _normalized_path(right)
    except Exception:
        return str(left) == str(right)


def _module_namespace(plugin_dir: Path) -> str:
    key = str(_normalized_path(plugin_dir)).encode("utf-8", "ignore")
    digest = hashlib.sha1(key).hexdigest()[:10]
    plugin_name = "".join(ch if ch.isalnum() else "_" for ch in plugin_dir.name)
    return f"sishub_plugin_{plugin_name}_{digest}"


def activate_plugin_environment(src_path: Path) -> None:
    """
    Keep plugin src path available while the plugin view is open.
    This avoids delayed-import failures in callbacks and worker threads.
    """
    src_path = _normalized_path(src_path)
    if not src_path.exists():
        raise FileNotFoundError(f"No existe el directorio src del plugin: {src_path}")

    for item in sys.path:
        if _same_path(item, src_path):
            return
    sys.path.insert(0, str(src_path))


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
    entrypoint = str(entrypoint).strip()
    if ":" not in entrypoint:
        raise ValueError(
            f"Entrypoint invalido '{entrypoint}'. Formato esperado: 'modulo:funcion'"
        )

    module_name, func_name = entrypoint.split(":", 1)
    module_name = module_name.strip()
    func_name = func_name.strip()
    if not module_name or not func_name:
        raise ValueError(
            f"Entrypoint invalido '{entrypoint}'. Formato esperado: 'modulo:funcion'"
        )

    module_rel = Path(*module_name.split(".")).with_suffix(".py")
    module_path = plugin_dir / "src" / module_rel
    if not module_path.exists():
        raise FileNotFoundError(f"Entrypoint no encontrado: {module_path}")

    unique_name = f"{_module_namespace(plugin_dir)}_{module_name.replace('.', '_')}"
    spec = importlib.util.spec_from_file_location(unique_name, module_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"No se pudo cargar el modulo: {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[unique_name] = module
    try:
        spec.loader.exec_module(module)  # type: ignore[attr-defined]
    except Exception:
        sys.modules.pop(unique_name, None)
        raise

    func = getattr(module, func_name, None)
    if func is None:
        raise AttributeError(f"Funcion '{func_name}' no encontrada en {module_path}")
    if not callable(func):
        raise TypeError(
            f"Entrypoint '{entrypoint}' invalido: '{func_name}' no es una funcion callable"
        )
    return func


def cleanup_plugin_environment(src_path: Path) -> None:
    """
    Clean plugin import environment:
    - removes src_path from sys.path
    - removes modules loaded from src_path from sys.modules
    This avoids collisions between plugins (application, utils, etc.).
    """
    src_path = _normalized_path(src_path)

    # 1) Remove src_path from sys.path (all matches)
    new_sys_path = []
    for p in sys.path:
        if _same_path(p, src_path):
            continue
        new_sys_path.append(p)
    sys.path[:] = new_sys_path

    # 2) Remove modules loaded from that src_path
    to_delete = []
    for name, mod in list(sys.modules.items()):
        if mod is None:
            continue

        mod_file = getattr(mod, "__file__", None)
        if not mod_file:
            continue

        try:
            mod_path = Path(mod_file).resolve()
        except Exception:
            continue

        if src_path == mod_path or src_path in mod_path.parents:
            to_delete.append(name)

    for name in to_delete:
        sys.modules.pop(name, None)

    importlib.invalidate_caches()


def build_context(app_root: Path, logger_name: str) -> PluginContext:
    from .constants import user_data_dir

    return PluginContext(
        root=app_root,
        data_dir=user_data_dir(),
        colors=COLORS,
        app_version=APP_VERSION,
        logger_name=logger_name,
    )
