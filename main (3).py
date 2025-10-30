from __future__ import annotations

import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from utils.constantes import DEFAULT_SHEET, DEFAULT_EXPORT_FORMAT, VALID_EXPORT_FORMATS
from application.funciones import leer_excel, concatenar, guardar


def crear_interfaz(parent, context):
    """Interfaz del plugin: seleccionar 2 Excel, concatenar y exportar con progreso."""
    root = ttk.Frame(parent)
    root.pack(fill="both", expand=True, padx=10, pady=10)

    ttk.Label(root, text="Concatenar Excel (Demo)", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 8))
    ttk.Label(root, text="Cargar dos archivos Excel, concatenar y exportar (XLSX o CSV).").pack(anchor="w")

    file1_var = tk.StringVar()
    file2_var = tk.StringVar()
    sheet1_var = tk.StringVar(value="" if DEFAULT_SHEET is None else str(DEFAULT_SHEET))
    sheet2_var = tk.StringVar(value="" if DEFAULT_SHEET is None else str(DEFAULT_SHEET))
    out_dir_var = tk.StringVar(value=str(context.data_dir / "exports"))
    out_name_var = tk.StringVar(value="concat_result")
    out_fmt_var = tk.StringVar(value=DEFAULT_EXPORT_FORMAT)

    status_var = tk.StringVar(value="Listo.")
    running = tk.BooleanVar(value=False)
    cancel_flag = tk.BooleanVar(value=False)

    def browse_file(target_var: tk.StringVar):
        path = filedialog.askopenfilename(
            title="Selecciona un Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        if path:
            target_var.set(path)

    def browse_dir():
        path = filedialog.askdirectory(title="Selecciona carpeta de salida")
        if path:
            out_dir_var.set(path)

    form = ttk.Frame(root); form.pack(fill="x", pady=(10, 6))

    r1 = ttk.Frame(form); r1.pack(fill="x", pady=4)
    ttk.Label(r1, text="Archivo 1:").pack(side="left", padx=(0, 6))
    ttk.Entry(r1, textvariable=file1_var, width=60).pack(side="left", fill="x", expand=True)
    ttk.Button(r1, text="Buscar...", command=lambda: browse_file(file1_var)).pack(side="left", padx=6)
    ttk.Label(r1, text="Hoja (opcional):").pack(side="left", padx=(12, 6))
    ttk.Entry(r1, textvariable=sheet1_var, width=12).pack(side="left")

    r2 = ttk.Frame(form); r2.pack(fill="x", pady=4)
    ttk.Label(r2, text="Archivo 2:").pack(side="left", padx=(0, 6))
    ttk.Entry(r2, textvariable=file2_var, width=60).pack(side="left", fill="x", expand=True)
    ttk.Button(r2, text="Buscar...", command=lambda: browse_file(file2_var)).pack(side="left", padx=6)
    ttk.Label(r2, text="Hoja (opcional):").pack(side="left", padx=(12, 6))
    ttk.Entry(r2, textvariable=sheet2_var, width=12).pack(side="left")

    r3 = ttk.Frame(form); r3.pack(fill="x", pady=4)
    ttk.Label(r3, text="Carpeta salida:").pack(side="left", padx=(0, 6))
    ttk.Entry(r3, textvariable=out_dir_var, width=60).pack(side="left", fill="x", expand=True)
    ttk.Button(r3, text="Examinar...", command=browse_dir).pack(side="left", padx=6)

    r4 = ttk.Frame(form); r4.pack(fill="x", pady=4)
    ttk.Label(r4, text="Nombre archivo:").pack(side="left", padx=(0, 6))
    ttk.Entry(r4, textvariable=out_name_var, width=30).pack(side="left")
    ttk.Label(r4, text="Formato:").pack(side="left", padx=(12, 6))
    fmt_combo = ttk.Combobox(r4, textvariable=out_fmt_var, values=sorted(list(VALID_EXPORT_FORMATS)), width=8, state="readonly")
    fmt_combo.pack(side="left")
    fmt_combo.set(out_fmt_var.get())

    btns = ttk.Frame(root); btns.pack(fill="x", pady=(8, 6))
    btn_run = ttk.Button(btns, text="Concatenar", command=lambda: start())
    btn_run.pack(side="left")
    btn_cancel = ttk.Button(btns, text="Cancelar", command=lambda: cancel_flag.set(True))
    btn_cancel.pack(side="left", padx=6)
    btn_open = ttk.Button(btns, text="Abrir carpeta", command=lambda: open_folder(out_dir_var.get()), state="disabled")
    btn_open.pack(side="left", padx=6)

    progress = ttk.Progressbar(root, mode="determinate", maximum=100, length=420)
    progress.pack(pady=(4,2))
    ttk.Label(root, textvariable=status_var).pack(anchor="w")

    log = tk.Text(root, height=10, wrap="none")
    log.pack(fill="both", expand=True, pady=(6, 0))

    def set_status(text: str):
        root.after(0, lambda: status_var.set(text))

    def set_progress(val: int):
        root.after(0, lambda: progress.configure(value=val))

    def append_log(text: str):
        def _do():
            log.insert("end", text + "\n")
            log.see("end")
        root.after(0, _do)

    def open_folder(path: str):
        p = Path(path)
        try:
            p.mkdir(parents=True, exist_ok=True)
            import os
            if os.name == "nt":
                import subprocess
                subprocess.Popen(["explorer", str(p)])
            else:
                import webbrowser
                webbrowser.open(p.as_uri())
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta:\n{exc}")

    def validate_inputs():
        f1 = Path(file1_var.get().strip())
        f2 = Path(file2_var.get().strip())
        if not f1.exists() or not f2.exists():
            raise FileNotFoundError("Debes seleccionar 2 archivos Excel válidos.")
        s1 = sheet1_var.get().strip()
        s2 = sheet2_var.get().strip()
        s1 = int(s1) if s1.isdigit() else (s1 if s1 else None)
        s2 = int(s2) if s2.isdigit() else (s2 if s2 else None)
        out_dir = Path(out_dir_var.get().strip())
        out_name = out_name_var.get().strip() or "concat_result"
        out_fmt = out_fmt_var.get().strip().lower()
        if out_fmt not in VALID_EXPORT_FORMATS:
            raise ValueError(f"Formato no soportado: {out_fmt}")
        return f1, f2, s1, s2, out_dir, out_name, out_fmt

    def run_task():
        logger = context.logger()
        try:
            f1, f2, s1, s2, out_dir, out_name, out_fmt = validate_inputs()
        except Exception as e:
            append_log(f"[ERROR] {e}")
            set_status("Error de validación.")
            running.set(False)
            return

        try:
            cancel_flag.set(False)
            set_progress(0)
            set_status("Leyendo archivo 1..."); append_log(f"Leyendo: {f1} (hoja={s1})")
            df1 = leer_excel(f1, s1)
            if cancel_flag.get(): set_status("Cancelado."); running.set(False); return
            set_progress(25)

            set_status("Leyendo archivo 2..."); append_log(f"Leyendo: {f2} (hoja={s2})")
            df2 = leer_excel(f2, s2)
            if cancel_flag.get(): set_status("Cancelado."); running.set(False); return
            set_progress(55)

            set_status("Concatenando...")
            df = concatenar(df1, df2, ignore_index=True)
            append_log(f"Filas resultado: {len(df)}")
            if cancel_flag.get(): set_status("Cancelado."); running.set(False); return
            set_progress(75)

            set_status("Guardando...")
            out_file = guardar(df, out_dir, out_name, fmt=out_fmt)
            set_progress(100)
            set_status(f"Listo: {out_file.name}")
            append_log(f"Guardado en: {out_file}")
            root.after(0, lambda: btn_open.configure(state="normal"))
            logger.info("Exportado: %s", out_file)
        except Exception as exc:
            set_status("Error en la tarea.")
            append_log(f"[ERROR] {exc}")
            logger.exception("excel_concat_demo")
        finally:
            running.set(False)

    def start():
        if running.get():
            return
        running.set(True)
        root.after(0, lambda: [progress.configure(value=0),
                               status_var.set("Iniciando..."),
                               btn_open.configure(state="disabled"),
                               log.delete("1.0", "end")])
        t = threading.Thread(target=run_task, daemon=True)
        t.start()

    def on_destroy(_evt=None):
        cancel_flag.set(True)

    root.bind("<Destroy>", on_destroy)
    return root
