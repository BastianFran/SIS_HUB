# -*- coding: utf-8 -*-
"""
Created on Thu Oct 16 13:10:48 2025

@author: bbruna
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict

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

def discover_plugins(root: Path) -> list[PluginInfo]:
    """
    Escanea root/plugins/**/plugin.json y construye el catalogo.
    """
    plugins_dir = root / "plugins"
    results: list[PluginInfo] = []
    for plugin_json in plugins_dir.glob("**/plugin.json"):
        try:
            data = json.loads(plugin_json.read_text(encoding="utf-8"))
            entry = PluginInfo(
                id=data["id"],
                name=data.get("name", data["id"]),
                category=data.get("category", "otros"),
                description=data.get("description", ""),
                entrypoint=data["entrypoint"],
                path=plugin_json.parent,
                requires=data.get("requires", []),
                enabled=data.get("enabled", True),
            )
            results.append(entry)
        except Exception as exc:
            print(f"[WARN] Manifiesto de plugin invalido: {plugin_json}: {exc}")
    results.sort(key=lambda x: (x.category.lower(), x.name.lower()))
    return results

def group_by_category(plugins: list[PluginInfo]) -> dict[str, list[PluginInfo]]:
    cats: dict[str, list[PluginInfo]] = {}
    for p in plugins:
        if p.enabled:
            cats.setdefault(p.category, []).append(p)
    for k in cats:
        cats[k].sort(key=lambda x: x.name.lower())
    return cats
