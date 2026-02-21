# -*- coding: utf-8 -*-
"""
Created on Fri Jun 27 12:04:52 2025

@author: CSANTOS
"""

import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import pandas as pd
import os
from application import buscar_cliente as bc
from application import call_archivos as call
# =============================================================================
from utils.constantes import RUTA, IMAGEN_FONDO, USUARIOS_SOP_INT
# =============================================================================
df = pd.read_excel(RUTA)
def iniciar_interfaz():
    ''' Flujo '''
    archivo_labels = []
    btn_ejecutar = None

    def buscar_cliente():
        nonlocal archivo_labels, btn_ejecutar
        rut = entry_id.get().strip()
        cliente, archivos_disponibles = bc.buscar_cliente_por_rut(rut, df)

        for widget in archivo_labels:
            widget.destroy()
        archivo_labels = []

        if btn_ejecutar:
            btn_ejecutar.destroy()
            btn_ejecutar = None

        if cliente:
            entry_resultado.config(state="normal")
            entry_resultado.delete(0, tk.END)
            entry_resultado.insert(0, cliente)
            entry_resultado.config(state="readonly")

            y_pos = 230
            for archivo in archivos_disponibles:
                entry_archivo = tk.Entry(ventana, font=("Helvetica", 12), width=60,
                                          readonlybackground="white", fg="black")
                entry_archivo.insert(0, f"{archivo.capitalize()}: Sí")
                entry_archivo.config(state="readonly")
                canvas.create_window(100, y_pos, window=entry_archivo, anchor="w")
                archivo_labels.append(entry_archivo)
                y_pos += 35

            if archivos_disponibles:
                def ejecutar_y_alertar():
                    resultado = call.obtener_archivos(rut, archivos_disponibles)
                    if resultado:
                        # entry_resultado.config(state="normal")
                        # entry_resultado.delete(0, tk.END)
                        # entry_resultado.insert(0, "Proceso completado")
                        # entry_resultado.config(state="readonly")
                        messagebox.showinfo("Finalizado",
                                            "✅ El procedimiento ha sido completado exitosamente.")

                btn_ejecutar = tk.Button(ventana, text="Obtener Archivos Cliente",
                                          command=ejecutar_y_alertar)
                canvas.create_window(325, 150, window=btn_ejecutar, anchor="w")

        else:
            entry_resultado.config(state="normal")
            entry_resultado.delete(0, tk.END)
            entry_resultado.insert(0, "ID no encontrado")
            entry_resultado.config(state="readonly")
            messagebox.showwarning("Aviso", "El ID ingresado no existe.")

    # Crear ventana principal
    ventana = tk.Tk()
    ventana.title("Interfaz Renta Variable Institucionales")
    ventana.geometry("800x450")

    canvas = tk.Canvas(ventana, width=800, height=450)
    canvas.pack(fill="both", expand=True)

    imagen = Image.open(IMAGEN_FONDO)
    imagen = imagen.resize((800, 450), Image.ANTIALIAS)
    imagen_fondo = ImageTk.PhotoImage(imagen)
    canvas.create_image(0, 0, image=imagen_fondo, anchor="nw")

    usuario_id = os.getlogin()
    usuario_nombre = USUARIOS_SOP_INT.get(usuario_id, "Usuario")
    canvas.create_text(400, 50, text=f"Bienvenido, {usuario_nombre}.",
                        font=("Helvetica", 16), fill="white", anchor="center")

    canvas.create_text(100, 150, text="Rut:", font=("Helvetica", 14), fill="white", anchor="e")
    entry_id = tk.Entry(ventana)
    canvas.create_window(100, 150, window=entry_id, anchor="w")

    btn_buscar = tk.Button(ventana, text="Buscar", command=buscar_cliente)
    canvas.create_window(250, 150, window=btn_buscar, anchor="w")

    canvas.create_text(100, 190, text="Cliente:", font=("Helvetica", 14), fill="white", anchor="e")
    entry_resultado = tk.Entry(ventana, font=("Helvetica", 12), width=60,
                                state="readonly", readonlybackground="white", fg="black")
    canvas.create_window(100, 190, window=entry_resultado, anchor="w")

    canvas.create_text(100, 230, text="Archivos:", font=("Helvetica", 14), fill="white", anchor="e")

    ventana.mainloop()
    
# import tkinter as tk
# from tkinter import messagebox
# from PIL import Image, ImageTk
# from tkinter import ttk 
# import pandas as pd
# import os
# from application import buscar_cliente as bc
# from application import call_archivos as call
# from utils.constantes import RUTA, IMAGEN_FONDO, USUARIOS_SOP_INT

# df = pd.read_excel(RUTA)

# def iniciar_interfaz():
#     ''' Flujo principal '''
#     archivo_labels = []
#     btn_ejecutar = None

#     def limpiar_canvas():
#         canvas.delete("all")
#         for widget in ventana.winfo_children():
#             if widget != canvas:
#                 widget.destroy()

#     def vista_irvi_grupos():
#         limpiar_canvas()
    
#         imagen = Image.open(IMAGEN_FONDO)
#         imagen = imagen.resize((800, 450), Image.ANTIALIAS)
#         imagen_fondo = ImageTk.PhotoImage(imagen)
#         canvas.image = imagen_fondo
#         canvas.create_image(0, 0, image=imagen_fondo, anchor="nw")
    
#         usuario_id = os.getlogin()
#         usuario_nombre = USUARIOS_SOP_INT.get(usuario_id, "Usuario")
#         canvas.create_text(400, 50, text=f"Bienvenido, {usuario_nombre}.",
#                            font=("Helvetica", 16), fill="white", anchor="center")
    
#         btn_volver = tk.Button(ventana, text="IRVI Orginal", command=render_interfaz_completa)
#         canvas.create_window(20, 20, window=btn_volver, anchor="nw")
    
#         # 🆕 Etiqueta "Grupo" en posición de texto "Rut"
#         canvas.create_text(100, 150, text="Grupo:", font=("Helvetica", 14), fill="white", anchor="e")
    
#         # 🆕 Leer Excel y obtener lista de grupos únicos
#         grupos_unicos = sorted(set(df['GRUPO'].dropna().astype(str)))
    
#         # 🆕 Crear Combobox con valores únicos
#         combo_grupos = ttk.Combobox(ventana, values=grupos_unicos, state="readonly", width=30)
#         combo_grupos.set("Seleccione un grupo")  # Placeholder inicial
#         canvas.create_window(100, 150, window=combo_grupos, anchor="w")


#     def render_interfaz_completa():
#         nonlocal archivo_labels, btn_ejecutar
#         archivo_labels = []
#         btn_ejecutar = None
#         limpiar_canvas()

#         imagen = Image.open(IMAGEN_FONDO)
#         imagen = imagen.resize((800, 450), Image.ANTIALIAS)
#         imagen_fondo = ImageTk.PhotoImage(imagen)
#         canvas.image = imagen_fondo
#         canvas.create_image(0, 0, image=imagen_fondo, anchor="nw")

#         usuario_id = os.getlogin()
#         usuario_nombre = USUARIOS_SOP_INT.get(usuario_id, "Usuario")
#         canvas.create_text(400, 50, text=f"Bienvenido, {usuario_nombre}.",
#                            font=("Helvetica", 16), fill="white", anchor="center")

#         btn_irvi_grupos = tk.Button(ventana, text="IRVI Grupos", command=vista_irvi_grupos)
#         canvas.create_window(20, 20, window=btn_irvi_grupos, anchor="nw")

#         canvas.create_text(100, 150, text="Rut:", font=("Helvetica", 14), fill="white", anchor="e")
#         entry_id = tk.Entry(ventana)
#         canvas.create_window(100, 150, window=entry_id, anchor="w")

#         def buscar_cliente():
#             nonlocal archivo_labels, btn_ejecutar
#             rut = entry_id.get().strip()
#             cliente, archivos_disponibles = bc.buscar_cliente_por_rut(rut, df)

#             for widget in archivo_labels:
#                 widget.destroy()
#             archivo_labels = []

#             if btn_ejecutar:
#                 btn_ejecutar.destroy()
#                 btn_ejecutar = None

#             if cliente:
#                 entry_resultado.config(state="normal")
#                 entry_resultado.delete(0, tk.END)
#                 entry_resultado.insert(0, cliente)
#                 entry_resultado.config(state="readonly")

#                 y_pos = 230
#                 for archivo in archivos_disponibles:
#                     entry_archivo = tk.Entry(ventana, font=("Helvetica", 12), width=60,
#                                              readonlybackground="white", fg="black")
#                     entry_archivo.insert(0, f"{archivo.capitalize()}: Sí")
#                     entry_archivo.config(state="readonly")
#                     canvas.create_window(100, y_pos, window=entry_archivo, anchor="w")
#                     archivo_labels.append(entry_archivo)
#                     y_pos += 35

#                 if archivos_disponibles:
#                     def ejecutar_y_alertar():
#                         resultado = call.obtener_archivos(rut, archivos_disponibles)
#                         if resultado:
#                             messagebox.showinfo("Finalizado",
#                                                 "✅ El procedimiento ha sido completado exitosamente.")

#                     btn_ejecutar = tk.Button(ventana, text="Obtener Archivos Cliente",
#                                              command=ejecutar_y_alertar)
#                     canvas.create_window(325, 150, window=btn_ejecutar, anchor="w")

#             else:
#                 entry_resultado.config(state="normal")
#                 entry_resultado.delete(0, tk.END)
#                 entry_resultado.insert(0, "ID no encontrado")
#                 entry_resultado.config(state="readonly")
#                 messagebox.showwarning("Aviso", "El ID ingresado no existe.")

#         btn_buscar = tk.Button(ventana, text="Buscar", command=buscar_cliente)
#         canvas.create_window(250, 150, window=btn_buscar, anchor="w")

#         canvas.create_text(100, 190, text="Cliente:", font=("Helvetica", 14), fill="white", anchor="e")
#         entry_resultado = tk.Entry(ventana, font=("Helvetica", 12), width=60,
#                                    state="readonly", readonlybackground="white", fg="black")
#         canvas.create_window(100, 190, window=entry_resultado, anchor="w")

#         canvas.create_text(100, 230, text="Archivos:", font=("Helvetica", 14), fill="white", anchor="e")

#     ventana = tk.Tk()
#     ventana.title("Interfaz Renta Variable Institucionales")
#     ventana.geometry("800x450")

#     canvas = tk.Canvas(ventana, width=800, height=450)
#     canvas.pack(fill="both", expand=True)

#     render_interfaz_completa()
#     ventana.mainloop()

