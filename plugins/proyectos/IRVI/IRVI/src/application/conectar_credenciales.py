'''
Conectar a credenciales STANDBY
'''
import os
import sys
import logging
import cx_Oracle

# Agregar ruta personalizada antes de importar módulos propios
sys.path.append(os.path.abspath(r"J:\MOEQUITY\Family Office\Scripts\csantos"))
from accesos import Credentials

def conectar_bbdd():
    '''Conectar Credenciales para querys'''
    try:
        credenciales = Credentials()
        result = credenciales.get_generic_credential('OperacionesBD')
        usuario = result.username
        passw = result.password

        dsn = cx_Oracle.makedsn("standby-bd.bantcent.cl", 1531, service_name="ORA9")
        connection = cx_Oracle.connect(user=usuario, password=passw, dsn=dsn, encoding="UTF-8")
        cursor = connection.cursor()

        mensaje = f"Conexión exitosa a la base de datos como '{usuario}'"
        print(mensaje)
        logging.info(mensaje)

        return cursor, connection, usuario, passw

    except cx_Oracle.DatabaseError as db_error:
        mensaje_error = f"Error de base de datos: {db_error}"
        print(mensaje_error)
        logging.error(mensaje_error)
        return None, None, None, None

    except Exception as error:
        mensaje_error = f"Error inesperado al conectar a la base de datos: {error}"
        print(mensaje_error)
        logging.error(mensaje_error)
        return None, None, None, None
