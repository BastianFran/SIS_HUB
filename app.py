# -*- coding: utf-8 -*-
"""
Created on Thu Oct 16 13:11:58 2025

@author: bbruna
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from pathlib import Path

from .constants import APP_NAME, ANCHO, ALTO, COLORS
from .logging_utils import setup_logging
from .catalog import discover_plugins, group_by_category, PluginInfo
from .loader import load_entrypoint, build_context
from .router import Router

def run_app() -> None:
    root = tk.Tk()
    root.title(APP_NAME)
    root.geometry(f"{ANCHO}x{ALTO}")
    root.configure(bg=COLORS["bg"])

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("TFrame", background=COLORS["bg"])
    style.configure("Panel.TFrame", background=COLORS["panel"])
    style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["fg"])
    style.configure("Panel.TLabel", background=COLORS["panel"], foreground=COLORS["fg"])
    style.configure("TButton", padding=6)

    logger = setup_logging()
    app_root = Path(__file__).resolve().parents[1]

    # Layout principal: topbar + content
    topbar = ttk.Frame(root, style="Panel.TFrame")
    topbar.pack(side="top", fill="x")
    content = ttk.Frame(root, style="TFrame")
    content.pack(side="top", fill="both", expand=True)

    router = Router(content)

    # Topbar controls
    
    def update_back_button():
        if router.current() is home:
            btn_back.state(["disabled"])
        else:
            btn_back.state(["!disabled"])
    
    def go_back():
        router.pop()
        update_back_button()
        
    def go_home():
        router.home()
        update_back_button()
        
    
    btn_back = ttk.Button(topbar, text="Atrás", command = go_back)
    btn_home = ttk.Button(topbar, text="Inicio", command=go_home)
    btn_back.pack(side="left", padx=6, pady=6)
    btn_home.pack(side="left", padx=6, pady=6)
    
    
    # Home view
    home = HomeView(content, app_root, logger, router,
                    on_nav_change = update_back_button)
    router.push(home)
    update_back_button()
    
    root.mainloop()

class HomeView(ttk.Frame):
    def __init__(self, parent, app_root: Path, logger, router: Router,
                 on_nav_change = None):
        
        super().__init__(parent, style="TFrame")
        self.app_root = app_root
        self.logger = logger
        self.router = router
        self.on_nav_change = on_nav_change

        title = ttk.Label(self, text="SIS Hub", font=("Segoe UI", 16, "bold"))
        title.pack(padx=10, pady=(10, 0), anchor="w")

        # Search + Refresh
        ctl = ttk.Frame(self, style="TFrame")
        ctl.pack(fill="x", padx=10, pady=10)
        ttk.Label(ctl, text="Buscar:").pack(side="left", padx=(0, 6))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(ctl, textvariable=self.search_var, width=40)
        self.search_entry.pack(side="left")
        ttk.Button(ctl, text="Actualizar", command=self.refresh).pack(side="left", padx=6)

        # Lists: categories left, plugins right
        body = ttk.Frame(self, style="TFrame")
        body.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.categories_list = tk.Listbox(body, height=20)
        self.categories_list.pack(side="left", fill="y")
        self.categories_list.bind("<<ListboxSelect>>", lambda e: self.render_plugins())

        self.plugins_container = ttk.Frame(body, style="TFrame")
        self.plugins_container.pack(side="left", fill="both", expand=True, padx=10)

        self.plugins = discover_plugins(self.app_root)
        self.refresh()

    def refresh(self):
        self.plugins = discover_plugins(self.app_root)
        self._cats = group_by_category(self.plugins)
        self.categories_list.delete(0, tk.END)
        for cat in sorted(self._cats.keys()):
            self.categories_list.insert(tk.END, cat)
        if self.categories_list.size() > 0:
            self.categories_list.selection_set(0)
        self.render_plugins()

    def render_plugins(self):
        for w in self.plugins_container.winfo_children():
            w.destroy()

        search = (self.search_var.get() or "").strip().lower()
        cur_idx = self.categories_list.curselection()
        if not cur_idx:
            return
        cat = self.categories_list.get(cur_idx[0])
        items = self._cats.get(cat, [])

        # Filter by search term
        if search:
            items = [p for p in items if search in p.name.lower() or search in p.description.lower()]

        if not items:
            ttk.Label(self.plugins_container, text="Sin plugins disponibles en esta categoría.",
                      style="TLabel").pack(anchor="w", pady=10)
            return

        for p in items:
            self._render_plugin_card(p)

    def _render_plugin_card(self, plugin: PluginInfo):
        card = ttk.Frame(self.plugins_container, style="Panel.TFrame")
        card.pack(fill="x", pady=6)

        header = ttk.Frame(card, style="Panel.TFrame")
        header.pack(fill="x", padx=10, pady=10)

        name_lbl = ttk.Label(header, text=plugin.name, style="Panel.TLabel",
                             font=("Segoe UI", 12, "bold"))
        name_lbl.pack(side="left")

        open_btn = ttk.Button(header, text="Abrir", command=lambda p=plugin: self.open_plugin(p))
        open_btn.pack(side="right")

        desc = plugin.description or "(Sin descripción)"
        ttk.Label(card, text=desc, style="Panel.TLabel").pack(anchor="w", padx=10, pady=(0,10))

        if plugin.requires:
            req = ", ".join(plugin.requires)
            ttk.Label(card, text=f"Requiere: {req}", style="Panel.TLabel").pack(anchor="w", padx=10, pady=(0,10))

    def open_plugin(self, plugin: PluginInfo):
        try:
            entry = load_entrypoint(plugin.path, plugin.entrypoint)
            ctx = build_context(self.app_root, logger_name=f"{plugin.id}")
            frame = ttk.Frame(self.router.parent, style="TFrame")
            child = entry(frame, ctx)  # plugin must mount UI inside 'frame'
            frame.pack_propagate(True)
            self.router.push(frame)
            if self.on_nav_change:
                self.on_nav_change()
        except Exception as exc:
            self._show_error(str(exc))

    def _show_error(self, message: str):
        dlg = tk.Toplevel(self)
        dlg.title("Error")
        ttk.Label(dlg, text="No se pudo abrir el plugin:", style="TLabel").pack(padx=10, pady=(10, 0))
        ttk.Label(dlg, text=message, style="TLabel", foreground="#ef4444").pack(padx=10, pady=(0, 10))
        ttk.Button(dlg, text="Cerrar", command=dlg.destroy).pack(pady=(0,10))
