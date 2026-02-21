# -*- coding: utf-8 -*-
"""
Main Tkinter application bootstrap for SIS Hub.
"""

from __future__ import annotations

import importlib.util
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path

from .catalog import PluginInfo, discover_plugins, group_by_category
from .constants import ALTO, ANCHO, APP_NAME, APP_VERSION, COLORS, app_base_dir
from .loader import (
    activate_plugin_environment,
    build_context,
    cleanup_plugin_environment,
    load_entrypoint,
)
from .logging_utils import setup_logging
from .router import Router

MODULE_PACKAGE_HINTS = {
    "win32com": "pywin32",
    "pil": "pillow",
    "fitz": "pymupdf",
    "cv2": "opencv-python",
    "sklearn": "scikit-learn",
    "yaml": "pyyaml",
    "dateutil": "python-dateutil",
    "cx_oracle": "cx-Oracle",
    "pyodbc": "pyodbc",
}


def _configure_styles(style: ttk.Style) -> None:
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure("TFrame", background=COLORS["bg"])
    style.configure("Panel.TFrame", background=COLORS["panel"])
    style.configure(
        "Topbar.TFrame",
        background=COLORS["panel"],
        borderwidth=1,
        relief="solid",
    )
    style.configure(
        "Card.TFrame",
        background=COLORS["panel"],
        borderwidth=1,
        relief="solid",
    )
    style.configure(
        "TLabel",
        background=COLORS["bg"],
        foreground=COLORS["fg"],
        font=("Segoe UI", 10),
    )
    style.configure(
        "Panel.TLabel",
        background=COLORS["panel"],
        foreground=COLORS["fg"],
        font=("Segoe UI", 10),
    )
    style.configure(
        "TopbarTitle.TLabel",
        background=COLORS["panel"],
        foreground=COLORS["fg"],
        font=("Segoe UI", 13, "bold"),
    )
    style.configure(
        "TopbarSub.TLabel",
        background=COLORS["panel"],
        foreground=COLORS["muted"],
        font=("Segoe UI", 9),
    )
    style.configure(
        "Section.TLabel",
        background=COLORS["panel"],
        foreground=COLORS["fg"],
        font=("Segoe UI", 10, "bold"),
    )
    style.configure(
        "CardTitle.TLabel",
        background=COLORS["panel"],
        foreground=COLORS["fg"],
        font=("Segoe UI", 11, "bold"),
    )
    style.configure(
        "CardBody.TLabel",
        background=COLORS["panel"],
        foreground=COLORS["fg"],
        font=("Segoe UI", 10),
    )
    style.configure(
        "Muted.TLabel",
        background=COLORS["bg"],
        foreground=COLORS["muted"],
        font=("Segoe UI", 9),
    )
    style.configure("TLabelframe", background=COLORS["panel"], bordercolor=COLORS["border"])
    style.configure(
        "TLabelframe.Label",
        background=COLORS["panel"],
        foreground=COLORS["fg"],
        font=("Segoe UI", 10, "bold"),
    )
    style.configure(
        "Badge.TLabel",
        background=COLORS["panel_soft"],
        foreground=COLORS["muted"],
        font=("Segoe UI", 9, "bold"),
        padding=(8, 2),
    )
    style.configure(
        "TButton",
        padding=6,
        background=COLORS["panel"],
        foreground=COLORS["fg"],
        borderwidth=1,
    )
    style.map(
        "TButton",
        background=[
            ("disabled", COLORS["panel_soft"]),
            ("active", COLORS["panel_soft"]),
            ("pressed", COLORS["panel_soft"]),
        ],
        foreground=[("disabled", COLORS["muted"]), ("!disabled", COLORS["fg"])],
    )
    style.configure(
        "Primary.TButton",
        padding=6,
        background=COLORS["accent"],
        foreground=COLORS["accent_fg"],
        borderwidth=0,
    )
    style.map(
        "Primary.TButton",
        background=[
            ("disabled", COLORS["panel_soft"]),
            ("active", COLORS["focus"]),
            ("pressed", COLORS["focus"]),
        ],
        foreground=[
            ("disabled", COLORS["muted"]),
            ("!disabled", COLORS["accent_fg"]),
        ],
    )
    style.configure(
        "Secondary.TButton",
        padding=6,
        background=COLORS["panel"],
        foreground=COLORS["fg"],
    )
    style.configure(
        "TEntry",
        fieldbackground=COLORS["input_bg"],
        foreground=COLORS["input_fg"],
        insertcolor=COLORS["input_fg"],
        borderwidth=1,
    )
    style.map(
        "TEntry",
        fieldbackground=[("disabled", COLORS["panel_soft"]), ("!disabled", COLORS["input_bg"])],
        foreground=[("disabled", COLORS["muted"]), ("!disabled", COLORS["input_fg"])],
    )
    style.configure(
        "TCombobox",
        fieldbackground=COLORS["input_bg"],
        foreground=COLORS["input_fg"],
        borderwidth=1,
    )
    style.map(
        "TCombobox",
        fieldbackground=[
            ("readonly", COLORS["input_bg"]),
            ("disabled", COLORS["panel_soft"]),
        ],
        foreground=[
            ("readonly", COLORS["input_fg"]),
            ("disabled", COLORS["muted"]),
        ],
        selectbackground=[("readonly", COLORS["accent"])],
        selectforeground=[("readonly", COLORS["accent_fg"])],
    )
    style.configure(
        "Vertical.TScrollbar",
        troughcolor=COLORS["bg"],
        background=COLORS["panel_soft"],
        arrowcolor=COLORS["muted"],
        bordercolor=COLORS["border"],
    )


def run_app() -> None:
    root = tk.Tk()
    root.title(f"{APP_NAME} - v{APP_VERSION}")
    root.geometry(f"{ANCHO}x{ALTO}")
    root.minsize(900, 560)
    root.configure(bg=COLORS["bg"])
    root.option_add("*Font", "{Segoe UI} 10")
    root.option_add("*Listbox.Background", COLORS["panel"])
    root.option_add("*Listbox.Foreground", COLORS["fg"])
    root.option_add("*Listbox.SelectBackground", COLORS["list_select_bg"])
    root.option_add("*Listbox.SelectForeground", COLORS["list_select_fg"])

    style = ttk.Style(root)
    _configure_styles(style)

    try:
        logger = setup_logging()
    except FileNotFoundError as exc:
        messagebox.showerror(
            "Configurar log",
            f"{exc}\n\nCrea el archivo y vuelve a abrir SIS Hub.",
        )
        root.destroy()
        return

    app_root = app_base_dir()
    logger.info("Directorio base de la aplicacion: %s", app_root)

    topbar = ttk.Frame(root, style="Topbar.TFrame")
    topbar.pack(side="top", fill="x", padx=12, pady=(12, 0))

    branding = ttk.Frame(topbar, style="Topbar.TFrame")
    branding.pack(side="left", fill="x", expand=True, padx=10, pady=10)
    ttk.Label(branding, text=APP_NAME, style="TopbarTitle.TLabel").pack(
        side="left", anchor="w"
    )
    ttk.Label(branding, text=f"v{APP_VERSION}", style="TopbarSub.TLabel").pack(
        side="left", padx=(8, 0), pady=(2, 0)
    )

    nav = ttk.Frame(topbar, style="Topbar.TFrame")
    nav.pack(side="right", padx=10, pady=10)

    content = ttk.Frame(root, style="TFrame")
    content.pack(side="top", fill="both", expand=True, padx=12, pady=(10, 12))

    router = Router(content)

    def update_back_button() -> None:
        if router.current() is home:
            btn_back.state(["disabled"])
        else:
            btn_back.state(["!disabled"])

    def go_back() -> None:
        router.pop()
        update_back_button()

    def go_home() -> None:
        router.home()
        update_back_button()

    btn_back = ttk.Button(nav, text="Atras", command=go_back, style="Secondary.TButton")
    btn_home = ttk.Button(
        nav, text="Inicio", command=go_home, style="Secondary.TButton"
    )
    btn_back.pack(side="left", padx=(0, 6))
    btn_home.pack(side="left")

    home = HomeView(content, app_root, logger, router, on_nav_change=update_back_button)
    router.push(home)
    update_back_button()

    root.mainloop()


class HomeView(ttk.Frame):
    def __init__(
        self,
        parent,
        app_root: Path,
        logger,
        router: Router,
        on_nav_change=None,
    ):
        super().__init__(parent, style="TFrame")
        self.app_root = app_root
        self.logger = logger
        self.router = router
        self.on_nav_change = on_nav_change

        self.plugins: list[PluginInfo] = []
        self._cats: dict[str, list[PluginInfo]] = {}
        self.status_var = tk.StringVar(value="")
        self._open_buttons: list[ttk.Button] = []
        self._opening_plugin = False

        title_row = ttk.Frame(self, style="TFrame")
        title_row.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(title_row, text="Catalogo de plugins", font=("Segoe UI", 15, "bold")).pack(
            side="left", anchor="w"
        )
        ttk.Label(title_row, text="SIS HUB", style="Badge.TLabel").pack(
            side="left", padx=(10, 0), pady=(1, 0)
        )
        ttk.Label(
            title_row,
            text="Abre proyectos desde una sola interfaz.",
            style="Muted.TLabel",
        ).pack(side="left", padx=(10, 0), pady=(4, 0))

        controls = ttk.Frame(self, style="TFrame")
        controls.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Label(controls, text="Buscar:").pack(side="left", padx=(0, 6))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.render_plugins())
        self.search_entry = ttk.Entry(controls, textvariable=self.search_var, width=46)
        self.search_entry.pack(side="left")
        self.btn_refresh = ttk.Button(controls, text="Actualizar", command=self.refresh)
        self.btn_refresh.pack(side="left", padx=8)

        body = ttk.Frame(self, style="TFrame")
        body.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        left_panel = ttk.Frame(body, style="Panel.TFrame")
        left_panel.pack(side="left", fill="y")
        ttk.Label(left_panel, text="Categorias", style="Section.TLabel").pack(
            anchor="w", padx=10, pady=(10, 6)
        )

        categories_wrap = ttk.Frame(left_panel, style="Panel.TFrame")
        categories_wrap.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.categories_list = tk.Listbox(
            categories_wrap,
            height=18,
            activestyle="none",
            exportselection=False,
            relief="flat",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["focus"],
            bg=COLORS["panel"],
            fg=COLORS["fg"],
            selectbackground=COLORS["list_select_bg"],
            selectforeground=COLORS["list_select_fg"],
            font=("Segoe UI", 10),
        )
        self.categories_list.pack(side="left", fill="y")
        self.categories_list.bind("<<ListboxSelect>>", lambda _: self.render_plugins())

        categories_scroll = ttk.Scrollbar(
            categories_wrap, orient="vertical", command=self.categories_list.yview
        )
        categories_scroll.pack(side="left", fill="y")
        self.categories_list.configure(yscrollcommand=categories_scroll.set)

        right_panel = ttk.Frame(body, style="TFrame")
        right_panel.pack(side="left", fill="both", expand=True, padx=(12, 0))

        self.plugins_canvas = tk.Canvas(
            right_panel,
            bg=COLORS["bg"],
            highlightthickness=0,
            borderwidth=0,
        )
        self.plugins_canvas.pack(side="left", fill="both", expand=True)

        plugins_scroll = ttk.Scrollbar(
            right_panel, orient="vertical", command=self.plugins_canvas.yview
        )
        plugins_scroll.pack(side="right", fill="y")
        self.plugins_canvas.configure(yscrollcommand=plugins_scroll.set)

        self.plugins_container = ttk.Frame(self.plugins_canvas, style="TFrame")
        self._plugins_window = self.plugins_canvas.create_window(
            (0, 0),
            window=self.plugins_container,
            anchor="nw",
        )
        self.plugins_container.bind(
            "<Configure>", lambda _evt: self._on_plugins_container_configure()
        )
        self.plugins_canvas.bind(
            "<Configure>", lambda evt: self.plugins_canvas.itemconfigure(self._plugins_window, width=evt.width)
        )
        self._bind_mousewheel()

        ttk.Label(self, textvariable=self.status_var, style="Muted.TLabel").pack(
            anchor="w", padx=10, pady=(0, 6)
        )

        self.refresh()

    def _on_plugins_container_configure(self) -> None:
        bbox = self.plugins_canvas.bbox("all")
        if bbox:
            self.plugins_canvas.configure(scrollregion=bbox)

    def _bind_mousewheel(self) -> None:
        self.plugins_canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.plugins_canvas.bind_all("<Button-4>", self._on_mousewheel, add="+")
        self.plugins_canvas.bind_all("<Button-5>", self._on_mousewheel, add="+")

    def _on_mousewheel(self, event) -> None:
        if not self.winfo_exists():
            return
        px, py = self.winfo_pointerxy()
        x0 = self.plugins_canvas.winfo_rootx()
        y0 = self.plugins_canvas.winfo_rooty()
        x1 = x0 + self.plugins_canvas.winfo_width()
        y1 = y0 + self.plugins_canvas.winfo_height()
        if not (x0 <= px <= x1 and y0 <= py <= y1):
            return
        if event.num == 4:
            delta = -1
        elif event.num == 5:
            delta = 1
        else:
            delta = -1 * int(event.delta / 120) if event.delta else 0
        if delta != 0:
            self.plugins_canvas.yview_scroll(delta, "units")

    def _selected_category(self) -> str | None:
        cur_idx = self.categories_list.curselection()
        if not cur_idx:
            return None
        return self.categories_list.get(cur_idx[0])

    def _set_opening_state(self, opening: bool) -> None:
        self._opening_plugin = opening
        if opening:
            self.btn_refresh.state(["disabled"])
            self.search_entry.state(["disabled"])
            self.categories_list.configure(state="disabled")
        else:
            self.btn_refresh.state(["!disabled"])
            self.search_entry.state(["!disabled"])
            self.categories_list.configure(state="normal")
        for btn in self._open_buttons:
            if opening:
                btn.state(["disabled"])
            else:
                btn.state(["!disabled"])

    def refresh(self) -> None:
        if self._opening_plugin:
            return
        previous_category = self._selected_category()
        try:
            self.plugins = discover_plugins(self.app_root)
        except Exception as exc:
            self.logger.exception("Error al descubrir plugins")
            self._show_error(f"Error al leer catalogo de plugins:\n{exc}")
            return
        self._cats = group_by_category(self.plugins)

        categories = sorted(self._cats.keys(), key=str.lower)
        self.categories_list.delete(0, tk.END)
        for cat in categories:
            self.categories_list.insert(tk.END, cat)

        if categories:
            select_index = 0
            if previous_category in categories:
                select_index = categories.index(previous_category)
            self.categories_list.selection_clear(0, tk.END)
            self.categories_list.selection_set(select_index)
            self.categories_list.activate(select_index)

        self.render_plugins()

    def render_plugins(self) -> None:
        self._open_buttons.clear()
        for widget in self.plugins_container.winfo_children():
            widget.destroy()

        cat = self._selected_category()
        if not cat:
            ttk.Label(
                self.plugins_container,
                text="No hay categorias disponibles.",
                style="Muted.TLabel",
            ).pack(anchor="w", pady=8)
            self.status_var.set("0 plugins disponibles.")
            self._on_plugins_container_configure()
            return

        items = list(self._cats.get(cat, []))
        total_in_category = len(items)

        search = (self.search_var.get() or "").strip().lower()
        if search:
            items = [
                p
                for p in items
                if search in p.name.lower() or search in p.description.lower()
            ]

        if not items:
            ttk.Label(
                self.plugins_container,
                text="Sin plugins para este filtro.",
                style="Muted.TLabel",
            ).pack(anchor="w", pady=8)
        else:
            for plugin in items:
                self._render_plugin_card(plugin)

        total_enabled = sum(len(cat_items) for cat_items in self._cats.values())
        if search:
            self.status_var.set(
                f"{len(items)} resultado(s) en '{cat}' (de {total_in_category}). Total: {total_enabled}."
            )
        else:
            self.status_var.set(f"{len(items)} plugin(s) en '{cat}'. Total: {total_enabled}.")
        self._on_plugins_container_configure()

    def _render_plugin_card(self, plugin: PluginInfo) -> None:
        card = ttk.Frame(self.plugins_container, style="Card.TFrame")
        card.pack(fill="x", pady=6)

        header = ttk.Frame(card, style="Panel.TFrame")
        header.pack(fill="x", padx=12, pady=(10, 4))

        ttk.Label(header, text=plugin.name, style="CardTitle.TLabel").pack(side="left")
        open_btn = ttk.Button(
            header,
            text="Abrir",
            style="Primary.TButton",
            command=lambda p=plugin: self.open_plugin(p),
        )
        open_btn.pack(side="right")
        self._open_buttons.append(open_btn)
        if self._opening_plugin:
            open_btn.state(["disabled"])

        desc = plugin.description or "(Sin descripcion)"
        ttk.Label(
            card,
            text=desc,
            style="CardBody.TLabel",
            wraplength=760,
            justify="left",
        ).pack(anchor="w", padx=12, pady=(0, 8))

        if plugin.requires:
            req = ", ".join(plugin.requires)
            ttk.Label(
                card,
                text=f"Requiere: {req}",
                style="Muted.TLabel",
            ).pack(anchor="w", padx=12, pady=(0, 10))
        else:
            ttk.Label(
                card,
                text="Requiere: solo entorno base del Hub.",
                style="Muted.TLabel",
            ).pack(anchor="w", padx=12, pady=(0, 10))

    def open_plugin(self, plugin: PluginInfo) -> None:
        if self._opening_plugin:
            return

        src_path = plugin.path / "src"
        frame: ttk.Frame | None = None
        env_active = False
        opened = False
        self.logger.info("Abriendo plugin '%s' en %s", plugin.id, plugin.path)
        self._set_opening_state(True)
        self.status_var.set(f"Abriendo '{plugin.name}'...")

        try:
            missing = self._find_missing_declared_dependencies(plugin)
            if missing:
                self.logger.warning(
                    "Dependencias declaradas no detectadas en '%s': %s",
                    plugin.id,
                    ", ".join(missing),
                )

            activate_plugin_environment(src_path)
            env_active = True

            entry = load_entrypoint(plugin.path, plugin.entrypoint)
            ctx = build_context(self.app_root, logger_name=f"sishub.plugin.{plugin.id}")
            frame = ttk.Frame(self.router.parent, style="TFrame")
            frame._sishub_cleanup = lambda sp=src_path: cleanup_plugin_environment(sp)

            entry(frame, ctx)  # plugin mounts UI inside frame
            self._ensure_plugin_layout(frame)

            frame.pack_propagate(True)
            self.router.push(frame)
            opened = True
            if self.on_nav_change:
                self.on_nav_change()
            self.logger.info("Plugin abierto correctamente: %s", plugin.id)
            self.status_var.set(f"Plugin abierto: {plugin.name}")
        except ModuleNotFoundError as exc:
            self.logger.exception("Dependencia faltante en plugin '%s'", plugin.id)
            self._show_error(self._missing_dependency_message(plugin, exc))
            self.status_var.set("No se pudo abrir el plugin por dependencias faltantes.")
        except Exception as exc:
            self.logger.exception("Error al abrir plugin '%s'", plugin.id)
            self._show_error(str(exc))
            self.status_var.set("No se pudo abrir el plugin.")
        finally:
            if env_active and not opened:
                cleanup_plugin_environment(src_path)
            if frame is not None and not opened:
                try:
                    frame.destroy()
                except Exception:
                    pass
            self._set_opening_state(False)

    def _ensure_plugin_layout(self, frame: ttk.Frame) -> None:
        children = frame.winfo_children()
        if not children:
            ttk.Label(
                frame,
                text="El plugin no monto widgets en su contenedor.",
                style="Muted.TLabel",
            ).pack(anchor="w", padx=12, pady=12)
            return
        if all(not child.winfo_manager() for child in children):
            # Fallback for plugins that return a root frame but forget to pack/grid it.
            try:
                children[0].pack(fill="both", expand=True)
            except Exception:
                pass

    def _find_missing_declared_dependencies(self, plugin: PluginInfo) -> list[str]:
        missing: list[str] = []
        for requirement in plugin.requires:
            module_name = self._requirement_to_module(requirement)
            if not module_name:
                continue
            try:
                found = importlib.util.find_spec(module_name)
            except Exception:
                found = None
            if found is None:
                missing.append(module_name)
        return sorted(set(missing))

    def _requirement_to_module(self, requirement: str) -> str | None:
        text = (requirement or "").strip()
        if not text:
            return None

        base = text.split("(", 1)[0].strip()
        token = base.split()[0].strip().lower().replace("-", "_")
        package_to_module = {
            "pywin32": "win32com",
            "pillow": "PIL",
            "pymupdf": "fitz",
            "python_dateutil": "dateutil",
            "cx_oracle": "cx_Oracle",
        }
        return package_to_module.get(token, token)

    def _missing_dependency_message(
        self, plugin: PluginInfo, exc: ModuleNotFoundError
    ) -> str:
        module_name = exc.name or "dependencia_desconocida"
        package_name = MODULE_PACKAGE_HINTS.get(module_name.lower(), module_name)
        lines = [
            f"No se pudo abrir '{plugin.name}'.",
            "",
            "Falta una dependencia de Python:",
            f"  - modulo: {module_name}",
            "",
            "Instala la dependencia en tu entorno activo y vuelve a intentar:",
            f"  pip install {package_name}",
        ]
        if plugin.requires:
            lines.extend(["", "Dependencias declaradas por el plugin:", f"  {', '.join(plugin.requires)}"])
        return "\n".join(lines)

    def _show_error(self, message: str) -> None:
        text = (message or "").strip()
        if len(text) > 1800:
            text = text[:1800] + "\n...\n(Detalle recortado. Revisa el log para ver el error completo.)"
        messagebox.showerror("No se pudo abrir el plugin", text, parent=self)
