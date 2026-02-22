# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 10:10:10 2025

@author: BBRUNA
"""

from __future__ import annotations

import importlib.util
import threading
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .master_store import MasterWorkbookStore, SearchResult
from .models import DependencyStatus, RuntimeOptions
from .pipeline import execute_customer_flow


def _is_module_available(name: str) -> bool:
    try:
        return importlib.util.find_spec(name) is not None
    except Exception:
        return False


def _dependency_checks() -> list[DependencyStatus]:
    return [
        DependencyStatus("pandas", "pandas", _is_module_available("pandas"), "Maestro / reportes"),
        DependencyStatus("openpyxl", "openpyxl", _is_module_available("openpyxl"), "Lectura Excel maestro"),
        DependencyStatus("requests", "requests", _is_module_available("requests"), "Descarga SOAP"),
        DependencyStatus("fitz", "pymupdf", _is_module_available("fitz"), "Resumen PDF y mezcla PDF"),
        DependencyStatus("keyring", "keyring", _is_module_available("keyring"), "Credenciales Windows"),
        DependencyStatus("cx_Oracle", "cx_Oracle", _is_module_available("cx_Oracle"), "Consultas Oracle"),
        DependencyStatus("win32com", "pywin32", _is_module_available("win32com"), "Correo Outlook"),
    ]


def _default_downloads_output() -> Path:
    home = Path.home()
    downloads = home / "Downloads"
    base = downloads if downloads.exists() else home
    target = base / "Facturacion Institucionales"
    target.mkdir(parents=True, exist_ok=True)
    return target


class FacturacionInstitucionalesUI:
    def __init__(self, parent, context):
        self.context = context
        self.log = context.logger()
        self.root = ttk.Frame(parent, style="TFrame")
        self.root.pack(fill="both", expand=True, padx=12, pady=12)

        plugin_root = Path(__file__).resolve().parents[2]
        self.master_path = plugin_root / "Facturacion_Institucionales_Maestro.xlsx"
        default_output = _default_downloads_output()

        self.output_dir_var = tk.StringVar(value=str(default_output))
        self.rut_var = tk.StringVar(value="")
        self.selected_client_var = tk.StringVar(value="-")
        self.group_var = tk.StringVar(value="-")
        self.source_var = tk.StringVar(value="-")
        self.status_var = tk.StringVar(value="Listo.")

        self.store = MasterWorkbookStore(self.master_path)
        self.selected_result: SearchResult | None = None
        self.running = False
        self._buttons: list[ttk.Button] = []

        self._build()
        self._reload_master(silent=True)

    def _build(self) -> None:
        ttk.Label(
            self.root,
            text="Facturacion Institucionales",
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            self.root,
            text="Flujo operativo con workbook maestro: facturas, XML, resumen, reportes y correo manual.",
        ).pack(anchor="w", pady=(0, 10))

        self._build_config_box()
        self._build_search_box()
        self._build_output_box()

    def _build_config_box(self) -> None:
        frame = ttk.LabelFrame(self.root, text="Configuracion")
        frame.pack(fill="x", pady=(0, 8))

        fixed_master = ttk.Frame(frame)
        fixed_master.pack(fill="x", padx=8, pady=(6, 4))
        ttk.Label(fixed_master, text="Workbook maestro fijo:").pack(side="left", padx=(0, 6))
        ttk.Label(
            fixed_master,
            text=str(self.master_path),
            style="Muted.TLabel",
        ).pack(side="left", fill="x", expand=True)

        self._path_row(frame, "Carpeta salida:", self.output_dir_var, self._pick_output)

        actions = ttk.Frame(frame)
        actions.pack(fill="x", padx=8, pady=(0, 8))
        btn_deps = ttk.Button(actions, text="Validar dependencias", command=self._show_dependencies)
        btn_deps.pack(side="left")
        btn_open_out = ttk.Button(actions, text="Abrir salida", command=self._open_output_folder)
        btn_open_out.pack(side="left", padx=6)
        self._buttons.extend([btn_deps, btn_open_out])

        self.counts_label = ttk.Label(frame, text="", style="Muted.TLabel")
        self.counts_label.pack(anchor="w", padx=8, pady=(0, 6))

    def _build_search_box(self) -> None:
        frame = ttk.LabelFrame(self.root, text="Busqueda y ejecucion")
        frame.pack(fill="both", expand=True)

        row = ttk.Frame(frame)
        row.pack(fill="x", padx=8, pady=(8, 6))
        ttk.Label(row, text="RUT:").pack(side="left", padx=(0, 6))
        rut_entry = ttk.Entry(row, textvariable=self.rut_var, width=24)
        rut_entry.pack(side="left")
        rut_entry.bind("<Return>", lambda _e: self._search_customer())
        btn_search = ttk.Button(row, text="Buscar", style="Primary.TButton", command=self._search_customer)
        btn_search.pack(side="left", padx=6)
        btn_run_all = ttk.Button(row, text="Ejecutar todo", style="Primary.TButton", command=self._run_all_actions)
        btn_run_all.pack(side="left", padx=6)
        btn_run = ttk.Button(row, text="Ejecutar seleccion", command=self._run_flow)
        btn_run.pack(side="left", padx=6)
        btn_mail_only = ttk.Button(row, text="Solo correo", command=self._run_mail_only)
        btn_mail_only.pack(side="left", padx=6)
        self._buttons.extend([btn_search, btn_run_all, btn_run, btn_mail_only])

        info_grid = ttk.Frame(frame)
        info_grid.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(info_grid, text="Cliente:").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        ttk.Entry(info_grid, textvariable=self.selected_client_var, state="readonly").grid(
            row=0, column=1, sticky="ew", pady=2
        )
        ttk.Label(info_grid, text="Grupo:").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=2)
        ttk.Entry(info_grid, textvariable=self.group_var, state="readonly").grid(
            row=1, column=1, sticky="ew", pady=2
        )
        ttk.Label(info_grid, text="Origen:").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=2)
        ttk.Entry(info_grid, textvariable=self.source_var, state="readonly").grid(
            row=2, column=1, sticky="ew", pady=2
        )
        info_grid.columnconfigure(1, weight=1)

        panes = ttk.Frame(frame)
        panes.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        left = ttk.LabelFrame(panes, text="Acciones detectadas")
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))
        ttk.Label(
            left,
            text="Selecciona una o varias acciones. Si no seleccionas ninguna, se ejecutan todas.",
            style="Muted.TLabel",
        ).pack(anchor="w", padx=6, pady=(6, 0))
        actions_toolbar = ttk.Frame(left)
        actions_toolbar.pack(fill="x", padx=6, pady=(4, 2))
        ttk.Button(actions_toolbar, text="Seleccionar todo", command=self._select_all_actions).pack(side="left")
        ttk.Button(actions_toolbar, text="Limpiar seleccion", command=self._clear_actions_selection).pack(
            side="left", padx=6
        )
        self.actions_selection_var = tk.StringVar(value="Sin acciones cargadas.")
        ttk.Label(actions_toolbar, textvariable=self.actions_selection_var, style="Muted.TLabel").pack(
            side="left", padx=(8, 0)
        )
        self.actions_list = tk.Listbox(left, height=8, exportselection=False, selectmode="extended")
        self.actions_list.pack(fill="both", expand=True, padx=6, pady=6)
        self.actions_list.bind("<<ListboxSelect>>", self._on_actions_selection_changed)

        right = ttk.LabelFrame(panes, text="Notas y validaciones")
        right.pack(side="left", fill="both", expand=True)
        self.notices_box = tk.Text(right, height=8, wrap="word")
        self.notices_box.pack(fill="both", expand=True, padx=6, pady=6)

        progress_row = ttk.Frame(frame)
        progress_row.pack(fill="x", padx=8, pady=(0, 6))
        self.progress = ttk.Progressbar(progress_row, mode="indeterminate")
        self.progress.pack(fill="x", expand=True)

        ttk.Label(frame, textvariable=self.status_var, style="Muted.TLabel").pack(
            anchor="w", padx=8, pady=(0, 6)
        )

    def _build_output_box(self) -> None:
        frame = ttk.LabelFrame(self.root, text="Salida del proceso")
        frame.pack(fill="both", pady=(8, 0))
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Button(toolbar, text="Limpiar log", command=self._clear_log).pack(side="left")
        self.log_box = tk.Text(frame, height=10, wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _path_row(self, parent, label: str, var: tk.StringVar, browse_command) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=4)
        ttk.Label(row, text=label).pack(side="left", padx=(0, 6))
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Examinar...", command=browse_command).pack(side="left", padx=6)

    def _pick_output(self) -> None:
        path = filedialog.askdirectory(title="Selecciona carpeta de salida")
        if path:
            self.output_dir_var.set(path)

    def _append_log(self, text: str) -> None:
        def _write():
            self.log_box.insert("end", text + "\n")
            self.log_box.see("end")

        self.root.after(0, _write)

    def _clear_log(self) -> None:
        self.log_box.delete("1.0", "end")

    def _on_actions_selection_changed(self, _event=None) -> None:
        total = self.actions_list.size()
        selected = len(self.actions_list.curselection())
        if total <= 0:
            self.actions_selection_var.set("Sin acciones cargadas.")
            return
        if selected <= 0:
            self.actions_selection_var.set(f"Sin seleccion ({total} disponibles, se ejecutan todas).")
            return
        self.actions_selection_var.set(f"Seleccionadas: {selected} de {total}.")

    def _select_all_actions(self) -> None:
        if self.actions_list.size() <= 0:
            return
        self.actions_list.selection_clear(0, "end")
        self.actions_list.selection_set(0, "end")
        self._on_actions_selection_changed()

    def _clear_actions_selection(self) -> None:
        self.actions_list.selection_clear(0, "end")
        self._on_actions_selection_changed()

    def _find_action_index(self, action_name: str) -> int | None:
        for idx in range(self.actions_list.size()):
            if str(self.actions_list.get(idx)) == action_name:
                return idx
        return None

    def _select_only_actions(self, action_names: list[str]) -> bool:
        self.actions_list.selection_clear(0, "end")
        found_any = False
        for action_name in action_names:
            idx = self._find_action_index(action_name)
            if idx is None:
                continue
            self.actions_list.selection_set(idx)
            found_any = True
        self._on_actions_selection_changed()
        return found_any

    def _run_all_actions(self) -> None:
        self._select_all_actions()
        self._run_flow()

    def _run_mail_only(self) -> None:
        if self.selected_result is None:
            self._search_customer()
            if self.selected_result is None:
                return
        ok = self._select_only_actions(["ENVIO_MANUAL_OUTLOOK_DISPLAY"])
        if not ok:
            messagebox.showwarning(
                "Facturacion Institucionales",
                "Este cliente no tiene accion de correo disponible.",
            )
            return
        self._run_flow()

    def _set_busy(self, busy: bool) -> None:
        self.running = busy
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()
        for btn in self._buttons:
            if busy:
                btn.state(["disabled"])
            else:
                btn.state(["!disabled"])

    def _reload_master(self, silent: bool = False) -> None:
        try:
            path = self.master_path
            self.store = MasterWorkbookStore(path)
            self.store.reload()
            if self.store.group_members_count > 0:
                counts_text = (
                    f"Maestro cargado | clientes: {self.store.clients_count} | "
                    f"grupos: {self.store.group_members_count} | total: {self.store.customer_count}"
                )
            else:
                counts_text = f"Maestro cargado | registros: {self.store.customer_count}"
            self.counts_label.configure(text=counts_text)
            self.status_var.set("Maestro cargado.")
            self._append_log(f"[OK] Maestro cargado: {path}")
        except Exception as exc:
            self.status_var.set("Error cargando maestro.")
            self.counts_label.configure(text="")
            self._append_log(f"[ERROR] Maestro: {exc}")
            if not silent:
                messagebox.showerror("Facturacion Institucionales", str(exc))

    def _search_customer(self) -> None:
        rut = (self.rut_var.get() or "").strip()
        if not rut:
            messagebox.showwarning("Facturacion Institucionales", "Debes ingresar un RUT.")
            return
        try:
            if not self.store.loaded:
                self._reload_master(silent=False)
            result = self.store.find_customer(rut)
        except Exception as exc:
            self._append_log(f"[ERROR] Busqueda: {exc}")
            messagebox.showerror("Facturacion Institucionales", str(exc))
            return

        self.actions_list.delete(0, "end")
        self.notices_box.delete("1.0", "end")
        self.selected_result = result
        self._on_actions_selection_changed()

        if result is None:
            self.selected_client_var.set("No encontrado")
            self.group_var.set("-")
            self.source_var.set("-")
            self.status_var.set("RUT no encontrado.")
            self._append_log(f"[WARN] RUT no encontrado: {rut}")
            messagebox.showwarning("Facturacion Institucionales", "RUT no encontrado en el maestro.")
            return

        self.selected_client_var.set(result.profile.display_name)
        self.group_var.set(result.profile.group_name or "-")
        self.source_var.set(result.profile.source_sheet)

        if result.actions:
            for item in result.actions:
                self.actions_list.insert("end", item)
            self._select_all_actions()
            if result.profile.tomy_delivery_enabled:
                for idx, action in enumerate(result.actions):
                    if action == "ENVIO_MANUAL_OUTLOOK_DISPLAY":
                        self.actions_list.selection_clear(idx)
                self._on_actions_selection_changed()
        else:
            self.actions_list.insert("end", "(sin acciones automaticas)")
            self._on_actions_selection_changed()

        for notice in result.notices:
            self.notices_box.insert("end", f"- {notice}\n")
        if not result.notices:
            self.notices_box.insert("end", "Sin observaciones.\n")

        self.status_var.set("Cliente encontrado.")
        self._append_log(
            f"[OK] Cliente {result.profile.display_name} | Acciones: {', '.join(result.actions) or 'ninguna'}"
        )

    def _show_dependencies(self) -> None:
        checks = _dependency_checks()
        lines = []
        missing_required = []
        for item in checks:
            if item.available:
                lines.append(f"[OK] {item.module_name} ({item.required_for})")
            else:
                lines.append(f"[MISSING] {item.module_name} -> pip install {item.package_name} ({item.required_for})")
                if item.module_name in {"pandas", "openpyxl", "requests", "keyring"}:
                    missing_required.append(item.module_name)
            self._append_log(lines[-1])

        message = "\n".join(lines)
        if missing_required:
            messagebox.showwarning("Dependencias", "Faltan dependencias base:\n\n" + message)
        else:
            messagebox.showinfo("Dependencias", message)

    def _open_output_folder(self) -> None:
        out_dir = Path((self.output_dir_var.get() or "").strip())
        out_dir.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(out_dir))  # type: ignore[attr-defined]
        except Exception as exc:
            self._append_log(f"[WARN] No se pudo abrir carpeta: {exc}")
            messagebox.showwarning("Carpeta salida", str(out_dir))

    def _runtime_options(self) -> RuntimeOptions:
        master = self.master_path
        output = Path((self.output_dir_var.get() or "").strip())
        if not str(output).strip():
            output = _default_downloads_output()
            self.output_dir_var.set(str(output))
        output.mkdir(parents=True, exist_ok=True)
        return RuntimeOptions(master_workbook=master, output_base_dir=output)

    def _run_flow(self) -> None:
        if self.running:
            return
        if self.selected_result is None:
            self._search_customer()
            if self.selected_result is None:
                return

        profile = self.selected_result.profile
        if not self.selected_result.actions and not self.selected_result.notices:
            messagebox.showwarning(
                "Facturacion Institucionales",
                "El cliente no tiene acciones configuradas ni avisos. Revisa el maestro.",
            )
            return

        selected_indices = list(self.actions_list.curselection())
        selected_actions = (
            [str(self.actions_list.get(i)) for i in selected_indices]
            if selected_indices
            else list(self.selected_result.actions)
        )
        selected_actions = [a for a in selected_actions if not a.startswith("(")]
        if not selected_actions and self.selected_result.actions:
            selected_actions = list(self.selected_result.actions)

        if profile.tomy_delivery_enabled and "ENVIO_MANUAL_OUTLOOK_DISPLAY" in selected_actions:
            confirm = messagebox.askyesno(
                "Confirmar correo manual",
                (
                    "Atencion: este cliente esta configurado para envio por TOMY.\n\n"
                    "El flujo normal es entregar la carpeta al RPA.\n"
                    "Deseas abrir el correo manual de todas formas?"
                ),
            )
            if not confirm:
                selected_actions = [a for a in selected_actions if a != "ENVIO_MANUAL_OUTLOOK_DISPLAY"]
                if not selected_actions:
                    self.status_var.set("Ejecucion cancelada por usuario.")
                    self._append_log("[INFO] Correo manual en cliente TOMY cancelado por usuario.")
                    return

        try:
            options = self._runtime_options()
        except Exception as exc:
            messagebox.showerror("Facturacion Institucionales", str(exc))
            return

        self._set_busy(True)
        self.status_var.set("Ejecutando flujo...")
        self._append_log(f"[INICIO] Flujo para {profile.display_name}")
        self._append_log(f"[INFO] Acciones seleccionadas UI: {', '.join(selected_actions) or 'ninguna'}")

        def worker() -> None:
            try:
                report = execute_customer_flow(
                    store=self.store,
                    profile=profile,
                    options=options,
                    logger=self.log,
                    requested_actions=selected_actions,
                    log_callback=self._append_log,
                )
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._on_error(e))
            else:
                self.root.after(0, lambda r=report: self._on_success(r))
            finally:
                self.root.after(0, lambda: self._set_busy(False))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, report) -> None:
        summary_lines = [
            f"Cliente: {report.customer.display_name}",
            f"Carpeta: {report.output_dir}",
            f"Fecha referencia: {report.reference_date or '-'}",
            f"Archivos generados: {len(report.generated_files)}",
        ]
        if report.warnings:
            summary_lines.append("")
            summary_lines.append("Advertencias:")
            summary_lines.extend([f"- {w}" for w in report.warnings])
            for warning in report.warnings:
                self._append_log(f"[WARN] {warning}")
        if report.info:
            for info in report.info:
                self._append_log(f"[INFO] {info}")

        self.status_var.set("Flujo completado.")
        self._append_log("[OK] Flujo completado")
        messagebox.showinfo("Facturacion Institucionales", "\n".join(summary_lines))

    def _on_error(self, exc: Exception) -> None:
        self.status_var.set("Flujo con error.")
        self._append_log(f"[ERROR] {exc}")
        self.log.exception("Error en plugin Facturacion Institucionales")
        messagebox.showerror("Facturacion Institucionales", str(exc))


def crear_interfaz_facturacion(parent, context):
    ui = FacturacionInstitucionalesUI(parent, context)
    return ui.root
