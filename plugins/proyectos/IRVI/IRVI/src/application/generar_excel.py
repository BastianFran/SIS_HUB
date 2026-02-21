# -*- coding: utf-8 -*-
"""
Created on Fri Jun 27 13:13:30 2025

@author: CSANTOS
"""
import logging
import pandas as pd
def qry_consulta_cliente(connection, rut_completo, consulta, fecha):
    """
    Extrae y ejecuta una consulta SQL para un cliente específico en una fecha dada.

    Parámetros:
        connection (sqlalchemy.engine.base.Connection): Conexión activa a la base de datos.
        rut_completo (str): RUT completo del cliente.
        consulta (str): Consulta SQL parametrizada.
        fecha (str): Fecha para filtrar los datos (formato YYYY-MM-DD).

    Retorna:
        pd.DataFrame: Resultados de la consulta o DataFrame vacío si ocurre un error.
    """
    try:
        df_excel = pd.read_sql(consulta, connection, params=[fecha])
        logging.info("Consulta ejecutada para RUT %s en la fecha %s", rut_completo, fecha)

        if df_excel.empty:
            logging.info("No se encontraron registros de resumen para RUT %s.", rut_completo)
        else:
            logging.info("Se encontraron %d registros de resumen para RUT %s.",
                         len(df_excel), rut_completo)

        return df_excel

    except Exception as error:
        logging.exception("Error al ejecutar la consulta para RUT %s: %s", rut_completo, error)
        return pd.DataFrame()
