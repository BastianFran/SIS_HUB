# -*- coding: utf-8 -*-
"""
Crear carpeta si no existe.
"""
import os
import logging
from typing import Optional

def crear_carpeta_cliente(nombre_cliente: str, carpeta_base: str) -> Optional[str]:
    """
    Crea una carpeta con el nombre del cliente dentro de la carpeta base.
    
    Parámetros:
    - nombre_cliente (str): Nombre del cliente que se usará como nombre de la carpeta.
    - carpeta_base (str): Ruta base donde se creará la carpeta del cliente.

    Retorna:
    - str | None: Ruta completa de la carpeta creada o existente. Retorna None si ocurre un error.
    """
    ruta_carpeta = os.path.join(carpeta_base, nombre_cliente)

    try:
        os.makedirs(ruta_carpeta, exist_ok=True)
        mensaje = f"Carpeta lista: {ruta_carpeta}"
        print(mensaje)
        logging.info(mensaje)
        return ruta_carpeta
    except OSError as error:
        mensaje_error = f"Error al crear la carpeta '{ruta_carpeta}': {error}"
        print(mensaje_error)
        logging.error(mensaje_error)
        return None
