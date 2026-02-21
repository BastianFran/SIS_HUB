from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk


def crear_interfaz(parent, context):
    """
    Demo de navegacion interna en un plugin usando varios frames.
    Mantiene estado simple y evita crear otra ventana Tk.
    """
    root = ttk.Frame(parent, style="TFrame")
    root.pack(fill="both", expand=True, padx=12, pady=12)

    log = context.logger()
    log.info("demo_navigation_frames: inicializando")

    state = {
        "nombre": tk.StringVar(value=""),
        "correo": tk.StringVar(value=""),
        "area": tk.StringVar(value="Operaciones"),
    }

    ttk.Label(root, text="Demo de Navegacion Interna", font=("Segoe UI", 14, "bold")).pack(
        anchor="w", pady=(0, 4)
    )
    ttk.Label(
        root,
        text="Flujo simple en 3 pasos: Bienvenida -> Datos -> Resumen.",
    ).pack(anchor="w", pady=(0, 10))

    step_var = tk.StringVar(value="Paso 1 de 3")
    ttk.Label(root, textvariable=step_var, style="Muted.TLabel").pack(anchor="w", pady=(0, 8))

    content = ttk.Frame(root, style="Panel.TFrame")
    content.pack(fill="both", expand=True)

    nav = ttk.Frame(root, style="TFrame")
    nav.pack(fill="x", pady=(10, 0))

    pages: dict[str, ttk.Frame] = {}
    current_page = {"name": None}

    def show_page(page_name: str) -> None:
        old_name = current_page["name"]
        if old_name and old_name in pages:
            pages[old_name].pack_forget()

        pages[page_name].pack(fill="both", expand=True, padx=10, pady=10)
        current_page["name"] = page_name

        if page_name == "welcome":
            step_var.set("Paso 1 de 3")
            btn_back.state(["disabled"])
            btn_next.state(["!disabled"])
            btn_next.configure(text="Siguiente")
            btn_finish.state(["disabled"])
        elif page_name == "form":
            step_var.set("Paso 2 de 3")
            btn_back.state(["!disabled"])
            btn_next.state(["!disabled"])
            btn_next.configure(text="Siguiente")
            btn_finish.state(["disabled"])
        else:
            step_var.set("Paso 3 de 3")
            btn_back.state(["!disabled"])
            btn_next.state(["disabled"])
            btn_finish.state(["!disabled"])
            lbl_resume.configure(
                text=(
                    f"Nombre: {state['nombre'].get().strip() or '-'}\n"
                    f"Correo: {state['correo'].get().strip() or '-'}\n"
                    f"Area: {state['area'].get().strip() or '-'}"
                )
            )

    def validate_form() -> bool:
        nombre = state["nombre"].get().strip()
        correo = state["correo"].get().strip()

        if not nombre:
            messagebox.showwarning("Dato requerido", "Debes ingresar un nombre.")
            return False
        if "@" not in correo or "." not in correo:
            messagebox.showwarning("Dato requerido", "Debes ingresar un correo valido.")
            return False
        return True

    page_welcome = ttk.Frame(content, style="Panel.TFrame")
    pages["welcome"] = page_welcome
    ttk.Label(
        page_welcome,
        text="Bienvenido al plugin de ejemplo.",
        style="Panel.TLabel",
        font=("Segoe UI", 11, "bold"),
    ).pack(anchor="w", pady=(6, 8))
    ttk.Label(
        page_welcome,
        text=(
            "Este plugin muestra una forma simple y robusta de crear\n"
            "navegacion interna sin abrir nuevas ventanas."
        ),
        style="Panel.TLabel",
        justify="left",
    ).pack(anchor="w")

    page_form = ttk.Frame(content, style="Panel.TFrame")
    pages["form"] = page_form
    form = ttk.LabelFrame(page_form, text="Datos del usuario")
    form.pack(fill="x", pady=(4, 0))
    ttk.Label(form, text="Nombre:").grid(row=0, column=0, sticky="w", padx=8, pady=8)
    ttk.Entry(form, textvariable=state["nombre"], width=40).grid(
        row=0, column=1, sticky="ew", padx=8, pady=8
    )
    ttk.Label(form, text="Correo:").grid(row=1, column=0, sticky="w", padx=8, pady=8)
    ttk.Entry(form, textvariable=state["correo"], width=40).grid(
        row=1, column=1, sticky="ew", padx=8, pady=8
    )
    ttk.Label(form, text="Area:").grid(row=2, column=0, sticky="w", padx=8, pady=8)
    area_combo = ttk.Combobox(
        form,
        textvariable=state["area"],
        values=["Operaciones", "Riesgo", "Finanzas", "Soporte", "TI"],
        state="readonly",
        width=37,
    )
    area_combo.grid(row=2, column=1, sticky="w", padx=8, pady=8)
    form.columnconfigure(1, weight=1)

    page_summary = ttk.Frame(content, style="Panel.TFrame")
    pages["summary"] = page_summary
    ttk.Label(
        page_summary,
        text="Resumen del flujo",
        style="Panel.TLabel",
        font=("Segoe UI", 11, "bold"),
    ).pack(anchor="w", pady=(6, 8))
    lbl_resume = ttk.Label(page_summary, text="", style="Panel.TLabel", justify="left")
    lbl_resume.pack(anchor="w")

    btn_back = ttk.Button(nav, text="Atras")
    btn_back.pack(side="left")
    btn_next = ttk.Button(nav, text="Siguiente")
    btn_next.pack(side="left", padx=6)
    btn_finish = ttk.Button(nav, text="Finalizar", style="Primary.TButton")
    btn_finish.pack(side="left", padx=6)

    def on_back() -> None:
        page = current_page["name"]
        if page == "form":
            show_page("welcome")
        elif page == "summary":
            show_page("form")

    def on_next() -> None:
        page = current_page["name"]
        if page == "welcome":
            show_page("form")
            return
        if page == "form":
            if not validate_form():
                return
            show_page("summary")

    def on_finish() -> None:
        log.info(
            "demo_navigation_frames: finalizado nombre=%s correo=%s area=%s",
            state["nombre"].get().strip(),
            state["correo"].get().strip(),
            state["area"].get().strip(),
        )
        messagebox.showinfo("Completado", "Flujo finalizado correctamente.")
        show_page("welcome")

    btn_back.configure(command=on_back)
    btn_next.configure(command=on_next)
    btn_finish.configure(command=on_finish)

    show_page("welcome")
    return root
