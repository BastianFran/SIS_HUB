'''
Limpiar carpeta para no pisar información.
'''

import os
import shutil
import logging

def limpiar_carpeta_cliente(ruta_carpeta: str) -> None:
    """
    Elimina todos los archivos y subcarpetas dentro de la carpeta del cliente.

    Parámetros:
    - ruta_carpeta (str): Ruta absoluta de la carpeta a limpiar.

    Retorna:
    - None
    """
    if not os.path.exists(ruta_carpeta):
        mensaje = f"La carpeta no existe: {ruta_carpeta}"
        print(mensaje)
        logging.warning(mensaje)
        return

    for nombre in os.listdir(ruta_carpeta):
        ruta_elemento = os.path.join(ruta_carpeta, nombre)
        try:
            if os.path.isfile(ruta_elemento) or os.path.islink(ruta_elemento):
                os.unlink(ruta_elemento)
            elif os.path.isdir(ruta_elemento):
                shutil.rmtree(ruta_elemento)
        except (PermissionError, FileNotFoundError, OSError) as error:
            mensaje_error = f"Error al eliminar '{ruta_elemento}': {error}"
            print(mensaje_error)
            logging.error(mensaje_error)

    mensaje_final = f"Carpeta limpiada: {ruta_carpeta}"
    print(mensaje_final)
    logging.info(mensaje_final)
