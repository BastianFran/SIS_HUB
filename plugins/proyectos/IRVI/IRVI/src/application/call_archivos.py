# -*- coding: utf-8 -*-
"""
Created on Thu Jun 26 12:20:27 2025

@author: CSANTOS
"""
# =============================================================================
#
# =============================================================================
import pandas as pd
import time
from datetime import datetime, timedelta
from application import crear_carpeta_cliente as ccc
from application import limpiar_carpeta_cliente as lcc
from application import conectar_credenciales as bbdd
from application import query_folios as qryff
from application import cliente_dte as cd
from application import combinar_pdfs as cpdf
from application import combinar_xmls as cxml
from application import query_resumen as qryr
from application import resumen_pdf as rpdf
from application import generar_excel as gexl
from application import correo_manual_interfaz as cmi
# =============================================================================
from utils.constantes import DESCARGA, RUTA_ARCHIVOS, IMAGEN_LOGO
# =============================================================================
RUTA = r"J:\MOEQUITY\Family Office\Scripts\csantos\IRVI\DATOS CLIENTES.xlsx"
df = pd.read_excel(RUTA)

def obtener_archivos(id_cliente, archivos_disponibles):
    print(f"\n Obteniendo archivos para cliente {id_cliente}...")

    try:
        cursor, connection, _, _ = bbdd.conectar_bbdd()
    except Exception as e:
        print(f"Error de conexión: {e}")
        return

    # Asegurar que el RUT y el DataFrame estén en formato string
    id_cliente = str(id_cliente).strip()
    df['RUT'] = df['RUT'].astype(str)

    try:
        nombre_cliente = df[df['RUT'] == id_cliente]['CARPETA'].values[0]
        rut_comp = df[df['RUT'] == id_cliente]['RUT DV'].values[0]
    except IndexError:
        print(" No se encontró el nombre del cliente en el DataFrame.")
        return

    # Crear y limpiar carpeta del cliente solo una vez
    ruta_descarga = ccc.crear_carpeta_cliente(nombre_cliente, DESCARGA)
    lcc.limpiar_carpeta_cliente(ruta_descarga)

    # Consultar datos desde la base
    fecha_hoy = datetime.today().strftime('%d-%m-%Y')
    # OPCIONAL MODO EJEMPLO FECHA ANTERIOR.
    fecha_ayer = (datetime.today() - timedelta(days=1)).strftime('%d-%m-%Y')
    # Ejecutar fecha hoy
    df_resultado = qryff.qry_folios_fecha(connection, id_cliente, fecha_hoy)

    if df_resultado.empty:
        df_resultado = qryff.qry_folios_fecha(connection, id_cliente, fecha_ayer)

    if df_resultado.empty:
        print(" No se encontraron datos para procesar.")
        return

    # Procesar según tipo de archivo
    for archivo in archivos_disponibles:
        print(f" Procesando {archivo} para {id_cliente}")
        if archivo == "FACTURA":
            # procesar_excel_con_macro(df_resultado, "descargarPDF", ruta_descarga)
            FORMATO = 'pdf'
            folios_por_rut = {nombre_cliente:
                              df_resultado['FOLIO'].tolist()}
            time.sleep(30)
            cd.descargar_facturas(FORMATO, RUTA_ARCHIVOS, folios_por_rut)

        elif archivo == "XML":
            # procesar_excel_con_macro(df_resultado, "descargarXML", ruta_descarga)
            FORMATO = 'xml'
            folios_por_rut = {nombre_cliente:
                              df_resultado['FOLIO'].tolist()}
            cd.descargar_facturas(FORMATO, RUTA_ARCHIVOS, folios_por_rut)
    # Rescatar clave cliente
    df_resultado['CLAVE'] = df_resultado['RUT'].str[-4:]
    clave = df_resultado['CLAVE'].iloc[0]

    cpdf.combinar_pdfs(ruta_descarga, clave)

    cxml.comprimir_xml(ruta_descarga)
    #  Procesar RESUMEN si está disponible.
    if "RESUMEN" in archivos_disponibles:
        df_resumen = qryr.qry_resumen_cliente(
            connection, id_cliente, fecha_hoy)
        if df_resumen.empty:
            df_resumen = qryr.qry_resumen_cliente(
                connection, id_cliente, fecha_ayer)

        if df_resumen.empty:
            print(" No se encontraron datos de resumen.")
        else:
            print(" Datos de resumen obtenidos correctamente.")
            rpdf.generar_documento_resumen(df_resumen,
                                           rut_cliente=rut_comp,
                                           nombre_cliente=nombre_cliente,
                                           fecha_operacion=fecha_hoy,
                                           ruta_guardado=ruta_descarga,
                                           imagen_logo=IMAGEN_LOGO)
    if "EXCEL" in archivos_disponibles:
        print("se ejecutará consulta excel")
        df_cliente = df[df['RUT'].astype(str) == id_cliente]
        consulta = df_cliente['CONSULTA QRY'].iloc[0]
        df_qry = gexl.qry_consulta_cliente(
            connection, id_cliente, consulta, fecha_hoy)
        df_qry.to_excel(
            rf"{RUTA_ARCHIVOS}\{nombre_cliente}\{nombre_cliente}.xlsx", index=False)
        # print("FIN EJECUCIÓN")
        
    if "CORREO MANUAL" in archivos_disponibles:
        # crear variable asunto
        if id_cliente == '969662507':
            asunto = f'Facturas Banchile CDB BTG Renta Variable Local // {fecha_hoy} '
        else:
            asunto = f'Op. RV Banchile // {nombre_cliente} {fecha_hoy}'
        destinatarios = df[df['RUT'] == id_cliente]['DESTINATARIOS'].values[0]
        cmi.correo_manual(destinatarios,ruta_descarga, asunto)

    return True
