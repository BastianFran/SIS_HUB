# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 10:42:11 2025

@author: BBRUNA

Interfaz del plugin Facturas BCH para SIS Hub.
"""

import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk


def _load_runner():
    """
    Import diferido para permitir abrir la UI aunque falten dependencias
    pesadas (selenium, cx_Oracle, etc.). El error se muestra al ejecutar.
    """
    try:
        from .application.descarga_archivos import ejecutar as run_facturas

        return run_facturas
    except ImportError:
        try:
            from application.descarga_archivos import ejecutar as run_facturas

            return run_facturas
        except ImportError:
            from descarga_archivos import ejecutar as run_facturas

            return run_facturas


def crear_interfaz(parent, context):
    """
    Monta la UI dentro de `parent` y la devuelve.
    """
    root = ttk.Frame(parent)
    root.pack(fill="both", expand=True, padx=12, pady=12)

    colors = getattr(context, "colors", {}) if context else {}
    color_text = colors.get("colorLetras", "#EAEAEA")
    color_ok = "#6EE7B7"
    color_info = "#A5B4FC"

    ttk.Label(
        root,
        text="Facturas BCH - Generar y (opcional) levantar correos",
        font=("Segoe UI", 14, "bold"),
    ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))

    ttk.Label(
        root,
        text="Fecha (dd-mm-aaaa):",
        foreground=color_text,
    ).grid(row=1, column=0, sticky="w")

    fecha_var = tk.StringVar(value=datetime.now().strftime("%d-%m-%Y"))
    fecha_entry = ttk.Entry(root, textvariable=fecha_var, width=14)
    fecha_entry.grid(row=1, column=1, sticky="w", padx=(8, 0))

    levantar_var = tk.BooleanVar(value=True)
    levantar_chk = ttk.Checkbutton(
        root,
        text="Levantar correos (Outlook)",
        variable=levantar_var,
    )
    levantar_chk.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

    helper_var = tk.StringVar(
        value="Genera archivos y, si corresponde, prepara correos en Outlook (Display)."
    )
    helper_lbl = ttk.Label(root, textvariable=helper_var, foreground=color_info)
    helper_lbl.grid(row=3, column=0, columnspan=4, sticky="w", pady=(8, 0))

    status_var = tk.StringVar(value="")
    status_lbl = ttk.Label(root, textvariable=status_var, foreground=color_ok)
    status_lbl.grid(row=5, column=0, columnspan=4, sticky="w", pady=(10, 0))

    ejecutar_btn = ttk.Button(root, text="Ejecutar")
    ejecutar_btn.grid(row=4, column=0, pady=(10, 0), sticky="w")

    def validar_fecha(texto):
        try:
            datetime.strptime(texto, "%d-%m-%Y")
            return True
        except Exception:
            return False

    def ui_call(fn):
        if root.winfo_exists():
            root.after(0, fn)

    def set_status(text):
        ui_call(lambda: status_var.set(text))

    def set_controls_enabled(enabled):
        def _apply():
            state = "normal" if enabled else "disabled"
            fecha_entry.configure(state=state)
            levantar_chk.configure(state=state)
            ejecutar_btn.configure(state=state)

        ui_call(_apply)

    def show_info(title, text):
        ui_call(lambda: messagebox.showinfo(title, text, parent=root))

    def show_warning(title, text):
        ui_call(lambda: messagebox.showwarning(title, text, parent=root))

    def show_error(title, text):
        ui_call(lambda: messagebox.showerror(title, text, parent=root))

    def worker(fecha_text, levantar_correo):
        try:
            set_status("Procesando...")
            run_facturas = _load_runner()
            run_facturas(
                fecha_str=fecha_text,
                levantar_correo=levantar_correo,
            )
            show_info("Exito", "Proceso finalizado.")
        except Exception as exc:
            msg = str(exc).strip() or exc.__class__.__name__
            if "CoInitialize" in msg:
                msg = (
                    "Fallo al inicializar Outlook (COM) en el hilo de trabajo.\n"
                    "Actualiza el plugin o ejecuta nuevamente.\n\n"
                    f"Detalle: {msg}"
                )
            show_error("Error", msg)
        finally:
            set_status("")
            set_controls_enabled(True)

    def ejecutar_handler():
        fecha_text = fecha_var.get().strip()
        if not validar_fecha(fecha_text):
            show_warning("Fecha invalida", "Usa el formato dd-mm-aaaa.")
            return

        set_controls_enabled(False)
        thread = threading.Thread(
            target=worker,
            args=(fecha_text, bool(levantar_var.get())),
            daemon=True,
            name="correo_mapa_bch_worker",
        )
        thread.start()

    ejecutar_btn.configure(command=ejecutar_handler)

    for col in range(4):
        root.grid_columnconfigure(col, weight=0)
    root.grid_columnconfigure(3, weight=1)

    return root
