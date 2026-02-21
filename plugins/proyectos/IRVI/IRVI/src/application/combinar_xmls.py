# -*- coding: utf-8 -*-
"""
Created on Thu Jun 26 08:07:13 2025

@author: CSANTOS
"""
import os
import zipfile
import logging

def comprimir_xml(carpeta: str) -> None:
    """
    Comprime todos los archivos XML en un solo archivo ZIP.
    Elimina los archivos XML originales después de crear el ZIP.
    No crea el archivo ZIP si no hay archivos XML.
    """
    if not os.path.exists(carpeta):
        mensaje = f"❌ La carpeta '{carpeta}' no existe."
        print(mensaje)
        logging.error(mensaje)
        return

    archivos_xml = [archivo for archivo in os.listdir(carpeta) if archivo.lower().endswith('.xml')]

    if not archivos_xml:
        mensaje = "No se encontraron archivos XML para comprimir."
        print(mensaje)
        logging.warning(mensaje)
        return

    ruta_zip = os.path.join(carpeta, 'XML_MASIVO.zip')

    try:
        with zipfile.ZipFile(ruta_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for archivo in archivos_xml:
                ruta_archivo = os.path.join(carpeta, archivo)
                zipf.write(ruta_archivo, archivo)
                mensaje = f"Añadido al ZIP: {archivo}"
                print(mensaje)
                logging.info(mensaje)
    except (OSError, zipfile.BadZipFile) as error:
        mensaje_error = f"❌ Error al crear el archivo ZIP: {error}"
        print(mensaje_error)
        logging.error(mensaje_error)
        return

    # Eliminar archivos XML originales
    for archivo in archivos_xml:
        ruta_archivo = os.path.join(carpeta, archivo)
        try:
            os.remove(ruta_archivo)
            mensaje = f"Archivo eliminado: {archivo}"
            print(mensaje)
            logging.info(mensaje)
        except OSError as error:
            mensaje_error = f"Error al eliminar {archivo}: {error}"
            print(mensaje_error)
            logging.error(mensaje_error)

    mensaje_final = "Archivos XML comprimidos correctamente."
    print(mensaje_final)
    logging.info(mensaje_final)
