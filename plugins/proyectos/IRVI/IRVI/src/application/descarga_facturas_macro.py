# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 13:48:35 2025

@author: CSANTOS
"""
import xlwings as xw
import time

MACRO_FAC = r"C:\Users\csantos\OneDrive - Banchile\Escritorio\Proyectos\IRVI\Descargar Facturas3.1.xlsm"

def procesar_excel_con_macro(df_resultado, nombre_macro, ruta_descarga):
    if df_resultado.empty:
        print("⚠️ No hay datos para procesar en Excel.")
        return

    print(f"\n📥 Abriendo Excel y ejecutando macro: {nombre_macro}...")

    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(MACRO_FAC)

    macro_limpiar = wb.app.macro("Hoja1.limpiar")
    macro_limpiar()
    time.sleep(1)

    sheet = wb.sheets[0]
    for cell in range(len(df_resultado)):
        sheet.range(cell + 6, 1).value = df_resultado['RUT'].iloc[cell]
        sheet.range(cell + 6, 2).value = df_resultado['SUB_RUT'].iloc[cell]
        sheet.range(cell + 6, 3).value = df_resultado['FOLIO'].iloc[cell]
        sheet.range(cell + 6, 4).value = 33
        sheet.range(cell + 6, 5).value = ruta_descarga

    print(f"\n📄 Ejecutando macro para {nombre_macro}...")
    time.sleep(1)
    macro = wb.app.macro(f"Hoja1.{nombre_macro}")
    macro()
    time.sleep(1)

    wb.close()
    app.quit()
    print("✅ Proceso completado.")