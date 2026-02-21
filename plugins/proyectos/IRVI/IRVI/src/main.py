# -*- coding: utf-8 -*-
"""
@author: csantos
"""
import os
import logging
from datetime import datetime

# Configurar logging una sola vez
fecha_actual = datetime.now().strftime("%d-%m-%Y")
RUTA_LOGS = r"J:\MOEQUITY\Family Office\Scripts\csantos\IRVI\logs"
os.makedirs(RUTA_LOGS, exist_ok=True)
archivo_log = os.path.join(RUTA_LOGS, f"{fecha_actual}.log")

logging.basicConfig(
    filename=archivo_log,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

if __name__ == '__main__':
    from application import interfaz as itfz
    itfz.iniciar_interfaz()


