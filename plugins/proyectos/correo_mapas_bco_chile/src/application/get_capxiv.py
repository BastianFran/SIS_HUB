# -*- coding: utf-8 -*-
"""
Created on Tue Jan 20 10:51:36 2026

@author: BBRUNA
"""

import pandas as pd
import keyring
import cx_Oracle
from datetime import datetime
import re

def normalizar_fecha_ddmmaaaa(fecha_str: str) -> str:
    """
    Convierte una fecha en formato variable (p. ej. 'dd/mm/aaaa', 'dd-mm-aaaa',
    'd/m/aaaa', 'd-m-aaaa', con espacios o separadores mixtos) al formato 'ddmmaaaa'.

    Parámetros
    ----------
    fecha_str : str
        Cadena de fecha a normalizar.

    Retorna
    -------
    str
        Fecha en formato 'ddmmaaaa'.

    Errores
    -------
    ValueError
        Si la fecha no cumple el patrón día/mes/año (año de 4 dígitos) o es inválida.
    """
    if not isinstance(fecha_str, str):
        raise ValueError("fecha_str debe ser str")

    texto = fecha_str.strip()
    if not texto:
        raise ValueError("fecha_str está vacío")

    # Unificar separadores a '/'
    texto = re.sub(r"[.\-\s_]+", "/", texto)

    # Validación simple de estructura d/m/aaaa (admite 1-2 dígitos en d y m, 4 en aaaa)
    partes = texto.split("/")
    if len(partes) != 3 or not re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", texto):
        raise ValueError("Formato no reconocido. Use día/mes/año con año de 4 dígitos")

    try:
        fecha = datetime.strptime(texto, "%d/%m/%Y")
    except ValueError as error:
        # Fecha imposible (p. ej., 31/02/2025)
        raise ValueError(f"Fecha inválida: {fecha_str}") from error

    return f"{fecha.day:02d}{fecha.month:02d}{fecha.year:04d}"


def obtener_credenciales(nombre_credencial: str) -> tuple[str, str]:
    """
    Obtiene usuario y clave desde el almacén de credenciales genéricas de Windows.

    Parámetros:
        nombre_credencial (str): Nombre de la credencial (ej. 'OperacionesBD').

    Retorna:
        tuple[str, str]: (usuario, clave).

    Errores:
        Lanza ValueError si no se encuentra la credencial.
    """
    print(f"Obteniendo credenciales para: {nombre_credencial}")
    try:
        cred = keyring.get_credential(nombre_credencial, None)
        if cred:
            print("Credenciales obtenidas correctamente.")
            return cred.username, cred.password
        raise ValueError(f"No se encontraron credenciales para: {nombre_credencial}")
    except Exception as error:
        raise RuntimeError(f"Error al obtener credenciales: {error}") from error


def conectar_bbdd(nombre_credencial: str, host: str, puerto: int, service_name: str) -> tuple[cx_Oracle.Cursor, cx_Oracle.Connection]:
    """
    Establece conexión con la base de datos Oracle usando credenciales seguras.

    Parámetros:
        nombre_credencial (str): Nombre de la credencial en Windows Credential Manager.
        host (str): Host del servidor Oracle.
        puerto (int): Puerto de conexión.
        service_name (str): Nombre del servicio Oracle.

    Retorna:
        tuple: Cursor y conexión a la base de datos.
    """
    usuario, clave = obtener_credenciales(nombre_credencial)
    print("Iniciando conexión a la base de datos...")
    dsn = cx_Oracle.makedsn(host, puerto, service_name=service_name)

    try:
        conexion = cx_Oracle.connect(user=usuario, password=clave, dsn=dsn, encoding="UTF-8")
        cursor = conexion.cursor()
        print("Conexión establecida correctamente.")
        return cursor, conexion
    except cx_Oracle.Error as error:
        raise RuntimeError(f"Error al conectar a la base de datos: {error}") from error
    finally:
        # Eliminar referencias sensibles
        del usuario
        del clave



def ejecutar_consulta(conexion: cx_Oracle.Connection, sql: str, parametros: tuple | None = None) -> pd.DataFrame:
    """
    Ejecuta una consulta SQL (con o sin parámetros) y retorna los resultados en un DataFrame.
    """
    print("Ejecutando consulta SQL...")
    try:
        if parametros is None:
            df = pd.read_sql(sql, conexion)
        else:
            df = pd.read_sql(sql, conexion, params=parametros)

        if df.empty:
            print("No se encontraron registros.")
        else:
            print(f"Consulta ejecutada correctamente. Filas: {len(df)}")
            print(df.head(10).to_string(index=False))

        return df
    except Exception as error:
        print(f"Error al ejecutar la consulta: {error}")
        return pd.DataFrame()


def ejecutar(fecha_str):
    """
    Flujo completo:
    1. Obtener credenciales seguras desde Windows Credential Manager.
    2. Conectar a Oracle.
    3. Ejecutar consulta parametrizada.
    4. Cerrar recursos y limpiar credenciales.
    """
    
    fecha = normalizar_fecha_ddmmaaaa(fecha_str)
    
    # Parámetros de conexión
    nombre_credencial = "OperacionesBD"
    host = "standby-bd.bantcent.cl"
    puerto = 1531
    service_name = "ORA9"

    # SQL parametrizado
    sql = """
    SELECT
    det.rut_cliente              rut_cliente,
    det.nombre_cliente           nombre_cliente,
    det.fecha_operacion          fecha_operacion,
    det.fecha_vencimiento        fecha_vencimiento,
    det.fol_fac                  fol_fac,
    det.instrumento              instrumento,
    det.tipo_operacion           tipo_operacion,
    det.cantidad_acciones        cantidad_acciones,
    det.monto_$                  monto_$,
    upper(substr(det.nombre_cliente, 1, 2)
          || substr(fac.num_fac, - 3)) clave_dcv,
    det.sub_rut                  sub_rut,
    det.fec_liq_req              fecha_liquidacion
FROM
    (
        SELECT
            fac.tg_sru_tg_cli_tg_per_rut_per rut_cli,
            fac.tg_sru_sru_sru               sru_cli,
            obn.tip_ope_ord                  tip_ope,
            obn.ta_ser_nom_ser               nom_ser,
            MIN(fac.fol_fac)                 num_fac
        FROM
            ta_obn obn,
            ta_asg asg,
            tg_fac fac
        WHERE
                fac.fec_vct_fac = TO_DATE(:p_fec_ini, 'ddmmyyyy')
            AND asg.fec_asg BETWEEN cg_fnc_gbl.fg_res_dia_hab_date(TO_DATE(:p_fec_ini, 'ddmmyyyy'),
                                                                   2) AND cg_fnc_gbl.fg_res_dia_hab_date(TO_DATE(:p_fec_ini, 'ddmmyyyy'
                                                                   ),
                                                                                                         1)
            AND decode(trunc(TO_NUMBER(fac.tg_sru_tg_cli_tg_per_rut_per) / 1000000),
                       47,
                       'SI',
                       59,
                       'SI',
                       'NO') = 'SI'
            AND fac.num_fac = asg.ta_fac_num_fac
            AND asg.ta_obn_num_ord = obn.num_ord
        GROUP BY
            fac.tg_sru_tg_cli_tg_per_rut_per,
            fac.tg_sru_sru_sru,
            obn.tip_ope_ord,
            obn.ta_ser_nom_ser
    ) fac,
    (
        SELECT
            fac.tg_sru_tg_cli_tg_per_rut_per                        rut_cli,
            cg_fnc_gbl.fg_rut_fmt(fac.tg_sru_tg_cli_tg_per_rut_per) rut_cliente,
            cg_fnc_gbl.fg_nom_per(fac.tg_sru_tg_cli_tg_per_rut_per) nombre_cliente,
            trunc(asg.fec_asg)                                      fecha_operacion,
            fac.fec_vct_fac                                         fecha_vencimiento,
            fac.fol_fac,
            obn.ta_ser_nom_ser                                      instrumento,
            obn.tip_ope_ord                                         tipo_operacion,
            SUM(asg.cnt_asg)                                        cantidad_acciones,
            fac.mnt_tot_fac                                         monto_$,
            fac.tg_sru_sru_sru                                      sub_rut,
            liq.fec_liq_req
        FROM
            ta_obn     obn,
            ta_asg     asg,
            tg_fac     fac,
            tl_ord_liq liq
        WHERE
                fac.fec_vct_fac = TO_DATE(:p_fec_ini, 'ddmmyyyy')
            AND asg.fec_asg BETWEEN cg_fnc_gbl.fg_res_dia_hab_date(TO_DATE(:p_fec_ini, 'ddmmyyyy'),
                                                                   2) AND cg_fnc_gbl.fg_res_dia_hab_date(TO_DATE(:p_fec_ini, 'ddmmyyyy'
                                                                   ),
                                                                                                         1)
            AND decode(trunc(TO_NUMBER(fac.tg_sru_tg_cli_tg_per_rut_per) / 1000000),
                       47,
                       'SI',
                       59,
                       'SI',
                       'NO') = 'SI'
            AND fac.num_fac = asg.ta_fac_num_fac
            AND asg.ta_obn_num_ord = obn.num_ord
            AND fac.num_ord_liq = liq.num_ord
        GROUP BY
            fac.tg_sru_tg_cli_tg_per_rut_per,
            cg_fnc_gbl.fg_nom_per(fac.tg_sru_tg_cli_tg_per_rut_per),
            trunc(asg.fec_asg),
            fac.fec_vct_fac,
            fac.fol_fac,
            obn.ta_ser_nom_ser,
            obn.tip_ope_ord,
            fac.mnt_tot_fac,
            fac.tg_sru_sru_sru,
            liq.fec_liq_req
    ) det
WHERE
        det.rut_cli = fac.rut_cli
    AND det.sub_rut = fac.sru_cli
    AND det.tipo_operacion = fac.tip_ope
    AND det.instrumento = fac.nom_ser
ORDER BY
    10
    """
    parametros = {"p_fec_ini": fecha}
    # parametros = ("valor1", "valor2")
    # parametros = None  # clave: no enviar tupla vacía si no quieres usar params

    # Conexión segura
    cursor, conexion = None, None
    try:
        cursor, conexion = conectar_bbdd(nombre_credencial, host, puerto, service_name)

        # Ejecutar consulta
        df = ejecutar_consulta(conexion, sql, parametros)
        return df
        # Procesar resultados
        if not df.empty:
            print("Primeras filas:")
            print(df.head(5).to_string(index=False))

    except Exception as error:
        print(f"Error en el flujo: {error}")
    finally:
        # Cierre seguro
        if cursor:
            cursor.close()
            print("Cursor cerrado.")
        if conexion:
            conexion.close()
            print("Conexión cerrada.")

if __name__ == "__main__":
    fecha_str = "20-01-2026"
    df = ejecutar(fecha_str)
