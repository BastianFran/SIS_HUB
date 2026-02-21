from __future__ import annotations
import tkinter as tk
from tkinter import ttk

def crear_interfaz(parent, context):
    """
    Interfaz mínima: muestra título, un texto y un botón que cuenta clics.
    Debe montar su UI DENTRO de 'parent' y retornar (opcional) su frame raíz.
    """
    root = ttk.Frame(parent, style="TFrame")
    root.pack(fill="both", expand=True)

    title = ttk.Label(root, text="Demo UI mínima", font=("Segoe UI", 14, "bold"))
    title.pack(pady=10)

    info = ttk.Label(root, text=f"Versión del Hub: {context.app_version}")
    info.pack(pady=(0,10))

    clicks = tk.IntVar(value=0)
    counter = ttk.Label(root, text="Clics: 0")
    counter.pack(pady=5)

    def inc():
        clicks.set(clicks.get() + 1)
        counter.configure(text=f"Clics: {clicks.get()}")

    ttk.Button(root, text="Sumar 1", command=inc).pack(pady=5)

    # Limpieza opcional: si tuvieras timers o threads, los cierras aquí
    def on_destroy(_evt=None):
        # ejemplo: cancelar after/threads si existieran
        pass
    root.bind("<Destroy>", on_destroy)

    return root
