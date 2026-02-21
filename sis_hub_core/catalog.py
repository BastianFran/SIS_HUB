# -*- coding: utf-8 -*-
"""
Plugin discovery utilities.
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .constants import resource_roots


@dataclass(frozen=True)
class PluginInfo:
    id: str
    name: str
    category: str
    description: str
    entrypoint: str          # "module:function"
    path: Path               # plugin root path (contains plugin.json and src/)
    requires: list           # e.g., ["outlook", "network"]
    enabled: bool = True


def _as_bool(value: Any, default: bool = True) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"1", "true", "yes", "si"}:
            return True
        if normalized in {"0", "false", "no"}:
            return False
    return default if value is None else bool(value)


def _as_requires(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    text = str(value).strip()
    return [text] if text else []


def _validate_manifest(data: dict[str, Any], manifest_path: Path) -> PluginInfo:
    plugin_id = str(data.get("id", "")).strip()
    if not plugin_id:
        raise ValueError("Campo requerido 'id' vacio")

    entrypoint = str(data.get("entrypoint", "")).strip()
    if not entrypoint:
        raise ValueError("Campo requerido 'entrypoint' vacio")
    if ":" not in entrypoint:
        raise ValueError("Entrypoint invalido, se espera 'modulo:funcion'")

    plugin_root = manifest_path.parent
    src_dir = plugin_root / "src"
    if not src_dir.exists():
        raise ValueError(f"No existe carpeta src para el plugin ({src_dir})")

    return PluginInfo(
        id=plugin_id,
        name=str(data.get("name", plugin_id)).strip() or plugin_id,
        category=str(data.get("category", "otros")).strip() or "otros",
        description=str(data.get("description", "")).strip(),
        entrypoint=entrypoint,
        path=plugin_root,
        requires=_as_requires(data.get("requires", [])),
        enabled=_as_bool(data.get("enabled", True), default=True),
    )


def discover_plugins(root: Path) -> list[PluginInfo]:
    """
    Scan every potential plugin directory (external or embedded) and build the
    resulting catalog. When duplicated IDs are found the first location wins,
    favouring the external folder next to the executable.
    """
    logger = logging.getLogger("sishub.catalog")
    search_dirs: list[Path] = []
    seen_dirs: set[Path] = set()

    def add_dir(path: Path) -> None:
        try:
            key = path.resolve(strict=False)
        except Exception:
            key = path.absolute()
        if key in seen_dirs:
            return
        seen_dirs.add(key)
        search_dirs.append(path)

    add_dir(root / "plugins")
    for base in resource_roots():
        add_dir(base / "plugins")

    logger.debug(
        "Buscando plugins en: %s",
        " | ".join(str(d) for d in search_dirs),
    )

    results: list[PluginInfo] = []
    seen_ids: set[str] = set()

    for plugins_dir in search_dirs:
        if not plugins_dir.exists():
            continue
        manifests = sorted(
            plugins_dir.glob("**/plugin.json"),
            key=lambda p: str(p).lower(),
        )
        for plugin_json in manifests:
            try:
                data = json.loads(plugin_json.read_text(encoding="utf-8"))
                if not isinstance(data, dict):
                    raise ValueError("El manifiesto debe contener un objeto JSON")

                entry = _validate_manifest(data, plugin_json)
                plugin_id = entry.id
                if plugin_id in seen_ids:
                    logger.debug(
                        "Plugin duplicado '%s' ignorado en %s",
                        plugin_id,
                        plugin_json.parent,
                    )
                    continue
                results.append(entry)
                seen_ids.add(plugin_id)
            except Exception as exc:
                logger.warning("Manifiesto de plugin invalido: %s (%s)", plugin_json, exc)
    results.sort(key=lambda x: (x.category.lower(), x.name.lower()))
    return results


def group_by_category(plugins: list[PluginInfo]) -> dict[str, list[PluginInfo]]:
    cats: dict[str, list[PluginInfo]] = {}
    for plugin in plugins:
        if plugin.enabled:
            cats.setdefault(plugin.category, []).append(plugin)
    for category in cats:
        cats[category].sort(key=lambda x: x.name.lower())
    return cats
