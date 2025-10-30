from __future__ import annotations
import threading
import time
import tkinter as tk
from tkinter import ttk

def crear_interfaz(parent, context):
    """
    Ejemplo de tarea larga fuera del hilo de UI con cancelación y progreso.
    """
    root = ttk.Frame(parent, style="TFrame")
    root.pack(fill="both", expand=True)

    ttk.Label(root, text="Demo de tarea con progreso", font=("Segoe UI", 14, "bold")).pack(pady=10)

    progress = ttk.Progressbar(root, mode="determinate", maximum=100, length=360)
    progress.pack(pady=10)
    status = ttk.Label(root, text="Listo para iniciar")
    status.pack()

    running = tk.BooleanVar(value=False)
    cancel_flag = tk.BooleanVar(value=False)

    def worker():
        logger = context.logger()
        logger.info("Demo task started")
        for i in range(101):
            if cancel_flag.get():
                logger.info("Demo task canceled at %d%%", i)
                _set_status(f"Cancelado en {i}%")
                _set_progress(i)
                running.set(False)
                return
                # Avoid tight loop if cancelled late
            time.sleep(0.03)  # simula trabajo
            _set_progress(i)
            _set_status(f"Progreso: {i}%")
        logger.info("Demo task finished")
        running.set(False)

    def _set_progress(val: int):
        # Safe update from worker thread using 'after'
        root.after(0, lambda: progress.configure(value=val))

    def _set_status(text: str):
        root.after(0, lambda: status.configure(text=text))

    def start():
        if running.get():
            return
        running.set(True)
        cancel_flag.set(False)
        status.configure(text="Iniciando...")
        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def cancel():
        if running.get():
            cancel_flag.set(True)

    btns = ttk.Frame(root, style="TFrame")
    btns.pack(pady=10)
    ttk.Button(btns, text="Iniciar", command=start).pack(side="left", padx=5)
    ttk.Button(btns, text="Cancelar", command=cancel).pack(side="left", padx=5)

    def on_destroy(_evt=None):
        # Señal de cancelación y pequeña espera (opcional) para threads
        cancel_flag.set(True)

    root.bind("<Destroy>", on_destroy)
    return root
