# src/interfaz.py
from __future__ import annotations
import logging, threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
 
def crear_interfaz(parent, context):
    # ---------- raíz del plugin ----------
    root = ttk.Frame(parent)           # usa ttk consistente con el Hub
    root.pack(fill="both", expand=True, padx=12, pady=12)
 
    log = logging.getLogger(f"{context.logger_name}.mi_plugin")
    log.info("Cargando plugin 'mi_plugin'")
 
    # ---------- título y descripción ----------
    ttk.Label(root, text="Mi Plugin de Ejemplo",
              font=("Segoe UI", 14, "bold")).pack(anchor="w")
    ttk.Label(root, text="Demostración de controles: combobox, radios, check, archivos y progreso."
             ).pack(anchor="w", pady=(0,10))
 
    # ---------- sección parámetros ----------
    frm = ttk.LabelFrame(root, text="Parámetros")
    frm.pack(fill="x", pady=6)
 
    # Entrada de texto
    ttk.Label(frm, text="Nombre:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
    nombre = tk.StringVar(value="")
    ttk.Entry(frm, textvariable=nombre, width=28).grid(row=0, column=1, sticky="w", padx=6, pady=6)
 
    # Combobox (lista desplegable)
    ttk.Label(frm, text="Formato:").grid(row=1, column=0, sticky="w", padx=6, pady=6)
    formatos = ("xlsx", "csv")
    fmt = tk.StringVar(value="xlsx")
    ttk.Combobox(frm, textvariable=fmt, values=formatos, state="readonly", width=10
                ).grid(row=1, column=1, sticky="w", padx=6, pady=6)
 
    # Radiobuttons
    ttk.Label(frm, text="Modo:").grid(row=2, column=0, sticky="w", padx=6, pady=6)
    modo = tk.StringVar(value="rapido")
    rb = ttk.Frame(frm); rb.grid(row=2, column=1, sticky="w", padx=6, pady=6)
    ttk.Radiobutton(rb, text="Rápido", value="rapido", variable=modo).pack(side="left")
    ttk.Radiobutton(rb, text="Seguro", value="seguro", variable=modo).pack(side="left", padx=10)
 
    # Checkbutton
    guardar_log = tk.BooleanVar(value=True)
    ttk.Checkbutton(frm, text="Guardar log del proceso", variable=guardar_log
                   ).grid(row=3, column=1, sticky="w", padx=6, pady=6)
 
    # Ajuste grid
    frm.columnconfigure(1, weight=1)
 
    # ---------- selección de archivo ----------
    files_frame = ttk.LabelFrame(root, text="Archivo de entrada")
    files_frame.pack(fill="x", pady=6)
 
    ruta_in = tk.StringVar(value="")
    e = ttk.Entry(files_frame, textvariable=ruta_in)
    e.pack(side="left", fill="x", expand=True, padx=6, pady=6)
 
    def elegir_archivo():
        path = filedialog.askopenfilename(
            title="Selecciona un archivo",
            filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("Todos", "*.*")]
        )
        if path:
            ruta_in.set(path)
    ttk.Button(files_frame, text="Buscar...", command=elegir_archivo).pack(side="left", padx=6, pady=6)
 
    # ---------- progreso y estado ----------
    progress = ttk.Progressbar(root, mode="determinate", maximum=100, length=380)
    progress.pack(pady=(8,4))
    status = tk.StringVar(value="Listo.")
    ttk.Label(root, textvariable=status).pack(anchor="w")
 
    # ---------- acciones ----------
    btns = ttk.Frame(root); btns.pack(pady=8)
    btn_ejecutar = ttk.Button(btns, text="Ejecutar")
    btn_ejecutar.pack(side="left", padx=4)
    ttk.Button(btns, text="Limpiar", command=lambda: [ruta_in.set(""), nombre.set("")]
              ).pack(side="left", padx=4)
 
    # ---------- lógica de tarea en thread ----------
    cancel = tk.BooleanVar(value=False)
    running = tk.BooleanVar(value=False)
 
    def tarea_larga():
        try:
            # Validaciones básicas
            if not ruta_in.get():
                raise ValueError("Debes seleccionar un archivo.")
            if nombre.get().strip() == "":
                raise ValueError("Debes indicar un nombre.")
 
            # Simulación de pasos con progreso
            pasos = [("Leyendo archivo", 20), ("Procesando", 70), ("Guardando", 100)]
            for msg, pct in pasos:
                if cancel.get(): return
                status.set(msg)
                progress.configure(value=pct)
                root.after(10)  # cede UI
                # aquí llamarías a tus funciones reales (leer, procesar, guardar)
 
            if guardar_log.get():
                (context.data_dir / "mi_plugin.log").write_text("OK\n", encoding="utf-8")
 
            status.set("Completado.")
            messagebox.showinfo("Éxito", "Proceso terminado.")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            status.set("Error.")
        finally:
            running.set(False)
            btn_ejecutar.configure(state="normal")
            cancel.set(False)
 
    def ejecutar():
        if running.get():
            return
        # reset UI
        progress.configure(value=0)
        status.set("Iniciando...")
        btn_ejecutar.configure(state="disabled")
        running.set(True)
        threading.Thread(target=tarea_larga, daemon=True).start()
 
    btn_ejecutar.configure(command=ejecutar)
 
    # ---------- limpieza al cerrar ----------
    def on_destroy(_evt=None):
        cancel.set(True)
    root.bind("<Destroy>", on_destroy)
 
    return root