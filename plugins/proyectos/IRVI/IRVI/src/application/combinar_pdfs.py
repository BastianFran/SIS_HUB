# -*- coding: utf-8 -*-
"""
Created on Wed Jun 25 13:52:25 2025

@author: CSANTOS
"""
import os
import logging
from pypdf import PdfReader, PdfWriter

def combinar_pdfs(carpeta: str, password: str = "0000") -> None:
    """
    Combina todos los archivos PDF en la carpeta del cliente.
    Desencripta los archivos PDF con la contraseña proporcionada si es necesario.
    Elimina los archivos PDF originales después de combinarlos.
    """
    if not os.path.exists(carpeta):
        mensaje = f"❌ La carpeta '{carpeta}' no existe."
        print(mensaje)
        logging.error(mensaje)
        return

    pdf_writer = PdfWriter()
    archivos_procesados = []

    for archivo in os.listdir(carpeta):
        if archivo.lower().endswith('.pdf'):
            ruta_archivo = os.path.join(carpeta, archivo)
            try:
                pdf_reader = PdfReader(ruta_archivo)
                if pdf_reader.is_encrypted:
                    if not pdf_reader.decrypt(password):
                        mensaje = f"❌ No se pudo desencriptar: {archivo}"
                        print(mensaje)
                        logging.warning(mensaje)
                        continue
                for pagina in pdf_reader.pages:
                    pdf_writer.add_page(pagina)
                archivos_procesados.append(ruta_archivo)
                mensaje = f"Añadido: {archivo}"
                print(mensaje)
                logging.info(mensaje)
            except (OSError, ValueError) as error:
                mensaje_error = f"⚠️ Error con {archivo}: {error}"
                print(mensaje_error)
                logging.error(mensaje_error)

    if not pdf_writer.pages:
        mensaje = "⚠️ No se combinaron archivos."
        print(mensaje)
        logging.warning(mensaje)
        return

    ruta_guardado = os.path.join(carpeta, 'pdf-masivo.pdf')
    try:
        with open(ruta_guardado, "wb") as archivo_salida:
            pdf_writer.write(archivo_salida)
        mensaje = f"PDF combinado guardado como: {ruta_guardado}"
        print(mensaje)
        logging.info(mensaje)
    except OSError as error:
        mensaje_error = f"❌ Error al guardar el PDF combinado: {error}"
        print(mensaje_error)
        logging.error(mensaje_error)
        return

    # Eliminar archivos originales
    for archivo in archivos_procesados:
        try:
            os.remove(archivo)
            mensaje = f"Archivo eliminado: {archivo}"
            print(mensaje)
            logging.info(mensaje)
        except OSError as error:
            mensaje_error = f"⚠️ Error al eliminar {archivo}: {error}"
            print(mensaje_error)
            logging.error(mensaje_error)
