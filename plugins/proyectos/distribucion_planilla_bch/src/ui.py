# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 10:29:33 2025

@author: BBRUNA
"""

from __future__ import annotations

import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def _load_runtime():
    try:
        from .distribucion_engine import (
            crear_excel_con_facturas_reemplazadas,
            ejecutar_proceso,
            leer_excel_entrada,
        )
        from .pdf_facturas_bch import extraer_facturas_desde_pdfs
    except ImportError:
        from distribucion_engine import (
            crear_excel_con_facturas_reemplazadas,
            ejecutar_proceso,
            leer_excel_entrada,
        )
        from pdf_facturas_bch import extraer_facturas_desde_pdfs

    return {
        "crear_excel_con_facturas_reemplazadas": crear_excel_con_facturas_reemplazadas,
        "ejecutar_proceso": ejecutar_proceso,
        "leer_excel_entrada": leer_excel_entrada,
        "extraer_facturas_desde_pdfs": extraer_facturas_desde_pdfs,
    }


def crear_interfaz(parent, context):
    root = ttk.Frame(parent)
    root.pack(fill="both", expand=True, padx=12, pady=12)

    downloads_dir = Path.home() / "Downloads"
    default_input = downloads_dir / "input.xlsx"
    default_output = downloads_dir

    state = {
        "pdf_paths": [],
        "facturas_df": None,
        "generated_input_path": None,
        "busy": False,
    }

    def ui_call(fn):
        if root.winfo_exists():
            root.after(0, fn)

    def set_busy(flag: bool):
        state["busy"] = bool(flag)

        def _apply():
            widget_state = "disabled" if flag else "normal"
            for w in (
                btn_input,
                btn_output,
                btn_add_pdf,
                btn_remove_pdf,
                btn_clear_pdf,
                btn_validar,
                btn_extraer_pdf,
                btn_guardar_excel_pdf,
                btn_ejecutar,
                btn_extraer_ejecutar,
            ):
                try:
                    w.configure(state=widget_state)
                except Exception:
                    pass
            # Entradas
            for entry in (entry_input, entry_output):
                try:
                    entry.configure(state=widget_state)
                except Exception:
                    pass
            if flag:
                status_var.set("Procesando...")
            else:
                if status_var.get() == "Procesando...":
                    status_var.set("Listo.")

        ui_call(_apply)

    def append_log(text: str):
        def _write():
            txt_log.configure(state="normal")
            txt_log.insert("end", text.rstrip() + "\n")
            txt_log.see("end")
            txt_log.configure(state="disabled")
        ui_call(_write)

    def clear_log():
        def _clear():
            txt_log.configure(state="normal")
            txt_log.delete("1.0", "end")
            txt_log.configure(state="disabled")
        ui_call(_clear)

    def set_status(text: str):
        ui_call(lambda: status_var.set(text))

    def show_info(title: str, text: str):
        ui_call(lambda: messagebox.showinfo(title, text, parent=root))

    def show_warning(title: str, text: str):
        ui_call(lambda: messagebox.showwarning(title, text, parent=root))

    def show_error(title: str, text: str):
        ui_call(lambda: messagebox.showerror(title, text, parent=root))

    def refresh_pdf_list():
        list_pdfs.delete(0, "end")
        for p in state["pdf_paths"]:
            list_pdfs.insert("end", p)
        resumen = f"PDFs cargados: {len(state['pdf_paths'])}"
        if state["facturas_df"] is not None:
            df = state["facturas_df"]
            folios = df["folio"].nunique() if "folio" in df.columns and not df.empty else 0
            resumen += f" | Facturas extraidas: {len(df)} filas / {folios} folios"
        pdf_summary_var.set(resumen)

    def seleccionar_input():
        path = filedialog.askopenfilename(
            parent=root,
            title="Seleccionar Excel de entrada",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")],
        )
        if path:
            input_var.set(path)

    def seleccionar_output():
        path = filedialog.askdirectory(
            parent=root,
            title="Seleccionar carpeta de salida",
            initialdir=str(Path(output_var.get()).parent if output_var.get() else default_output),
        )
        if path:
            output_var.set(path)

    def agregar_pdfs():
        paths = filedialog.askopenfilenames(
            parent=root,
            title="Seleccionar uno o mas PDFs de facturas",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
        )
        if not paths:
            return
        existentes = set(state["pdf_paths"])
        for p in paths:
            if p not in existentes:
                state["pdf_paths"].append(p)
                existentes.add(p)
        refresh_pdf_list()

    def quitar_pdf():
        sel = list(list_pdfs.curselection())
        if not sel:
            return
        for idx in reversed(sel):
            del state["pdf_paths"][int(idx)]
        refresh_pdf_list()

    def limpiar_pdfs():
        state["pdf_paths"] = []
        state["facturas_df"] = None
        state["generated_input_path"] = None
        refresh_pdf_list()
        append_log("[INFO] Lista de PDFs limpiada.")

    def validar_rutas_base() -> tuple[Path, Path]:
        ruta_input = Path(input_var.get().strip()).expanduser()
        carpeta_salida = Path(output_var.get().strip()).expanduser()

        if not str(ruta_input):
            raise ValueError("Debes indicar la ruta del Excel de entrada.")
        if ruta_input.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            raise ValueError("El archivo de entrada debe ser un Excel (.xlsx/.xlsm/.xls).")
        if not ruta_input.exists():
            raise FileNotFoundError(f"No existe el Excel de entrada: {ruta_input}")
        carpeta_salida.mkdir(parents=True, exist_ok=True)
        return ruta_input, carpeta_salida

    def run_in_thread(fn, *, clear_console=False):
        if state["busy"]:
            return
        if clear_console:
            clear_log()
        set_busy(True)

        def _runner():
            try:
                fn()
            except Exception as exc:  # noqa: BLE001
                detalle = str(exc).strip() or exc.__class__.__name__
                append_log(f"[ERROR] {detalle}")
                append_log(traceback.format_exc())
                show_error("Error", detalle)
                set_status("Error.")
            finally:
                set_busy(False)

        threading.Thread(target=_runner, daemon=True, name="dist_planilla_worker").start()

    def accion_validar_excel():
        def _work():
            set_status("Validando Excel...")
            ruta_input, _ = validar_rutas_base()
            rt = _load_runtime()
            asignaciones, lineas = rt["leer_excel_entrada"](ruta_input)
            append_log(f"[OK] Excel valido: {ruta_input}")
            append_log(f"[INFO] Asignaciones: {len(asignaciones)}")
            append_log(f"[INFO] Lineas facturas: {len(lineas)}")
            append_log(f"[INFO] Folios: {len({l.folio for l in lineas})}")
            set_status("Excel validado.")
            show_info("Validacion", "Excel validado correctamente.")

        run_in_thread(_work, clear_console=False)

    def accion_extraer_pdfs():
        def _work():
            set_status("Extrayendo facturas desde PDF...")
            if not state["pdf_paths"]:
                raise ValueError("Debes agregar uno o mas PDFs.")
            rt = _load_runtime()
            df = rt["extraer_facturas_desde_pdfs"]([Path(p) for p in state["pdf_paths"]])
            state["facturas_df"] = df
            state["generated_input_path"] = None
            append_log("[OK] Extraccion PDF completada.")
            append_log(f"[INFO] Filas facturas: {len(df)}")
            append_log(f"[INFO] Folios detectados: {df['folio'].nunique() if not df.empty else 0}")
            append_log(f"[INFO] Archivos origen: {df['_origen_pdf'].nunique() if not df.empty else 0}")
            if not df.empty:
                try:
                    cant_total = (
                        df["cantidad"]
                        .astype(str)
                        .str.replace(".", "", regex=False)
                        .str.split(",", n=1)
                        .str[0]
                        .astype(int)
                        .sum()
                    )
                    neto_total = (
                        df["monto_neto"]
                        .astype(str)
                        .str.replace(".", "", regex=False)
                        .str.split(",", n=1)
                        .str[0]
                        .astype(int)
                        .sum()
                    )
                    append_log(f"[INFO] Cantidad total extraida: {cant_total}")
                    append_log(f"[INFO] Neto total extraido: {neto_total}")
                except Exception:
                    append_log("[WARN] No se pudo calcular resumen numerico de la extraccion.")
            ui_call(refresh_pdf_list)
            set_status("Facturas extraidas desde PDF.")
            show_info(
                "Extraccion PDF",
                f"Se extrajeron {len(df)} filas de facturas en {df['folio'].nunique()} folios.",
            )

        run_in_thread(_work, clear_console=False)

    def _guardar_excel_con_facturas_pdf_impl() -> Path:
        ruta_input, carpeta_salida = validar_rutas_base()
        if state["facturas_df"] is None:
            raise ValueError("Primero extrae facturas desde uno o mas PDFs.")
        rt = _load_runtime()
        ruta_generada = rt["crear_excel_con_facturas_reemplazadas"](
            ruta_excel_base=ruta_input,
            carpeta_salida=carpeta_salida,
            df_facturas=state["facturas_df"],
            nombre_archivo=f"{ruta_input.stem}_facturas_pdf.xlsx",
        )
        state["generated_input_path"] = ruta_generada
        return ruta_generada

    def accion_guardar_excel_pdf():
        def _work():
            set_status("Guardando Excel con hoja facturas desde PDF...")
            ruta_generada = _guardar_excel_con_facturas_pdf_impl()
            ui_call(lambda: input_var.set(str(ruta_generada)))
            append_log(f"[OK] Excel generado con facturas PDF: {ruta_generada}")
            set_status("Excel con facturas PDF generado.")
            show_info(
                "Excel generado",
                "Se reemplazo la hoja 'facturas' y se actualizo la ruta de entrada en la UI.",
            )

        run_in_thread(_work, clear_console=False)

    def _ejecutar_distribucion_impl(ruta_excel_real: Path, carpeta_salida: Path):
        rt = _load_runtime()
        resultado = rt["ejecutar_proceso"](ruta_excel_entrada=ruta_excel_real, carpeta_salida=carpeta_salida)
        append_log("[OK] Distribucion completada.")
        append_log(f"[INFO] Planilla: {resultado.ruta_planilla}")
        append_log(f"[INFO] Archivos de carga: {resultado.carpeta_archivos_carga}")
        append_log(f"[INFO] Asignaciones: {resultado.cantidad_asignaciones}")
        append_log(f"[INFO] Lineas facturas: {resultado.cantidad_lineas_factura}")
        append_log(f"[INFO] Folios: {resultado.cantidad_folios}")
        return resultado

    def accion_ejecutar_distribucion():
        def _work():
            set_status("Ejecutando distribucion...")
            ruta_input, carpeta_salida = validar_rutas_base()
            resultado = _ejecutar_distribucion_impl(ruta_input, carpeta_salida)
            set_status("Distribucion completada.")
            show_info(
                "Proceso finalizado",
                f"Planilla generada:\n{resultado.ruta_planilla}\n\n"
                f"Archivos de carga:\n{resultado.carpeta_archivos_carga}",
            )

        run_in_thread(_work, clear_console=True)

    def accion_extraer_y_ejecutar():
        def _work():
            set_status("Extrayendo PDFs + ejecutando distribucion...")
            ruta_input, carpeta_salida = validar_rutas_base()
            if not state["pdf_paths"]:
                raise ValueError("Debes agregar uno o mas PDFs para esta accion.")
            rt = _load_runtime()
            df = rt["extraer_facturas_desde_pdfs"]([Path(p) for p in state["pdf_paths"]])
            state["facturas_df"] = df
            ui_call(refresh_pdf_list)
            append_log(f"[OK] Facturas extraidas desde PDF: {len(df)} filas")

            ruta_generada = rt["crear_excel_con_facturas_reemplazadas"](
                ruta_excel_base=ruta_input,
                carpeta_salida=carpeta_salida,
                df_facturas=df,
                nombre_archivo=f"{ruta_input.stem}_facturas_pdf.xlsx",
            )
            state["generated_input_path"] = ruta_generada
            ui_call(lambda: input_var.set(str(ruta_generada)))
            append_log(f"[OK] Excel intermedio generado: {ruta_generada}")

            resultado = _ejecutar_distribucion_impl(ruta_generada, carpeta_salida)
            set_status("Extraccion + distribucion completadas.")
            show_info(
                "Proceso finalizado",
                f"Excel usado:\n{ruta_generada}\n\n"
                f"Planilla generada:\n{resultado.ruta_planilla}",
            )

        run_in_thread(_work, clear_console=True)

    ttk.Label(
        root,
        text="Distribucion de Planilla (BCH)",
        font=("Segoe UI", 14, "bold"),
    ).pack(anchor="w")
    ttk.Label(
        root,
        text=(
            "Usa un Excel con hojas 'asignaciones' y 'facturas'. "
            "Opcionalmente completa la hoja 'facturas' extrayendo lineas desde uno o mas PDFs BCH."
        ),
    ).pack(anchor="w", pady=(2, 10))

    frm_paths = ttk.LabelFrame(root, text="Entradas y salida")
    frm_paths.pack(fill="x", pady=(0, 8))

    ttk.Label(frm_paths, text="Excel de entrada:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
    input_var = tk.StringVar(value=str(default_input))
    entry_input = ttk.Entry(frm_paths, textvariable=input_var)
    entry_input.grid(row=0, column=1, sticky="ew", padx=6, pady=6)
    btn_input = ttk.Button(frm_paths, text="Buscar...", command=seleccionar_input)
    btn_input.grid(row=0, column=2, sticky="w", padx=6, pady=6)

    ttk.Label(frm_paths, text="Carpeta salida:").grid(row=1, column=0, sticky="w", padx=6, pady=6)
    output_var = tk.StringVar(value=str(default_output))
    entry_output = ttk.Entry(frm_paths, textvariable=output_var)
    entry_output.grid(row=1, column=1, sticky="ew", padx=6, pady=6)
    btn_output = ttk.Button(frm_paths, text="Elegir...", command=seleccionar_output)
    btn_output.grid(row=1, column=2, sticky="w", padx=6, pady=6)

    frm_paths.grid_columnconfigure(1, weight=1)

    frm_pdf = ttk.LabelFrame(root, text="PDFs BCH para completar hoja 'facturas' (opcional)")
    frm_pdf.pack(fill="both", expand=False, pady=(0, 8))

    list_pdfs = tk.Listbox(frm_pdf, height=5, selectmode="extended")
    list_pdfs.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=6, pady=6)
    scroll_pdfs = ttk.Scrollbar(frm_pdf, orient="vertical", command=list_pdfs.yview)
    scroll_pdfs.grid(row=0, column=3, sticky="ns", pady=6)
    list_pdfs.configure(yscrollcommand=scroll_pdfs.set)

    btn_add_pdf = ttk.Button(frm_pdf, text="Agregar PDF(s)...", command=agregar_pdfs)
    btn_add_pdf.grid(row=1, column=0, sticky="w", padx=6, pady=(0, 6))
    btn_remove_pdf = ttk.Button(frm_pdf, text="Quitar seleccion", command=quitar_pdf)
    btn_remove_pdf.grid(row=1, column=1, sticky="w", padx=6, pady=(0, 6))
    btn_clear_pdf = ttk.Button(frm_pdf, text="Limpiar PDFs", command=limpiar_pdfs)
    btn_clear_pdf.grid(row=1, column=2, sticky="w", padx=6, pady=(0, 6))

    pdf_summary_var = tk.StringVar(value="PDFs cargados: 0")
    ttk.Label(frm_pdf, textvariable=pdf_summary_var).grid(
        row=2, column=0, columnspan=4, sticky="w", padx=6, pady=(0, 6)
    )
    frm_pdf.grid_columnconfigure(0, weight=1)
    frm_pdf.grid_rowconfigure(0, weight=1)

    frm_actions = ttk.LabelFrame(root, text="Acciones")
    frm_actions.pack(fill="x", pady=(0, 8))

    btn_validar = ttk.Button(frm_actions, text="Validar Excel", command=accion_validar_excel)
    btn_validar.grid(row=0, column=0, sticky="w", padx=6, pady=6)

    btn_extraer_pdf = ttk.Button(frm_actions, text="Extraer facturas desde PDF(s)", command=accion_extraer_pdfs)
    btn_extraer_pdf.grid(row=0, column=1, sticky="w", padx=6, pady=6)

    btn_guardar_excel_pdf = ttk.Button(
        frm_actions,
        text="Guardar Excel con facturas PDF",
        command=accion_guardar_excel_pdf,
    )
    btn_guardar_excel_pdf.grid(row=0, column=2, sticky="w", padx=6, pady=6)

    btn_ejecutar = ttk.Button(frm_actions, text="Ejecutar distribucion", command=accion_ejecutar_distribucion)
    btn_ejecutar.grid(row=1, column=0, sticky="w", padx=6, pady=6)

    btn_extraer_ejecutar = ttk.Button(
        frm_actions,
        text="Extraer PDF + Ejecutar",
        command=accion_extraer_y_ejecutar,
    )
    btn_extraer_ejecutar.grid(row=1, column=1, sticky="w", padx=6, pady=6)

    ttk.Label(
        frm_actions,
        text=(
            "Flujo recomendado: 1) Agregar PDF(s) 2) Extraer facturas 3) Guardar Excel con facturas PDF "
            "4) Ejecutar distribucion."
        ),
    ).grid(row=2, column=0, columnspan=4, sticky="w", padx=6, pady=(0, 6))

    frm_log = ttk.LabelFrame(root, text="Log")
    frm_log.pack(fill="both", expand=True)
    txt_log = tk.Text(frm_log, height=12, wrap="word")
    txt_log.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
    txt_log.configure(state="disabled")
    scroll_log = ttk.Scrollbar(frm_log, orient="vertical", command=txt_log.yview)
    scroll_log.pack(side="right", fill="y", padx=6, pady=6)
    txt_log.configure(yscrollcommand=scroll_log.set)

    footer = ttk.Frame(root)
    footer.pack(fill="x", pady=(8, 0))
    status_var = tk.StringVar(value="Listo.")
    ttk.Label(footer, textvariable=status_var).pack(side="left")
    ttk.Button(footer, text="Limpiar log", command=clear_log).pack(side="right")

    append_log("[INFO] Plugin cargado.")
    append_log(f"[INFO] Excel por defecto: {default_input}")
    append_log(f"[INFO] Carpeta de salida por defecto: {default_output}")
    append_log("[INFO] Puedes usar el Excel directamente o completar 'facturas' desde PDFs BCH.")
    refresh_pdf_list()

    return root
