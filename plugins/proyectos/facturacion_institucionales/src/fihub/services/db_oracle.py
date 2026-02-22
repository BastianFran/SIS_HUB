# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 09:35:05 2025

@author: BBRUNA
"""

from __future__ import annotations

import logging
import os
from contextlib import contextmanager
from typing import Iterator

import pandas as pd


LOGGER = logging.getLogger(__name__)


DEFAULT_CREDENTIAL_NAME = os.getenv("FIHUB_ORACLE_CREDENTIAL", "OperacionesBD")
DEFAULT_ORACLE_HOST = os.getenv("FIHUB_ORACLE_HOST", "standby-bd.bantcent.cl")
DEFAULT_ORACLE_PORT = int(os.getenv("FIHUB_ORACLE_PORT", "1531"))
DEFAULT_ORACLE_SERVICE = os.getenv("FIHUB_ORACLE_SERVICE", "ORA9")


FOLIOS_SQL = """
SELECT  TG_SRU_TG_CLI_TG_PER_RUT_PER AS Rut,
        TG_SRU_SRU_SRU AS Sub_Rut,
        TO_CHAR(FEC_FAC, 'DD-MM-YYYY') AS Fecha_Operacion,
        FOL_FAC AS Folio
FROM    ta_fac
WHERE   FEC_FAC = TO_DATE(:1, 'DD-MM-YYYY')
        AND FOL_FAC IS NOT NULL
        AND TG_SRU_TG_CLI_TG_PER_RUT_PER = :2
        AND USR_IMP_FAC IN ('JCHANDIA','JSEPULVE','MCARVACH1','BBRUNA','CSANTOS')
"""

SUMMARY_SQL = """
SELECT
    sru.cnp_sru Fondo,
    A.tip_ope_ord Operacion,
    A.ta_ser_nom_ser Instrumento,
    SUM(b.cnt_asg) Cantidad,
    SUM(ROUND(b.cnt_asg * b.pre_asg)) Monto
FROM
    ta_ult_trs c,
    ta_obn A,
    ta_asg b,
    tg_sru sru
WHERE
    b.fec_asg >= TO_DATE(:1, 'DD-MM-YYYY')
    AND a.tg_SRU_TG_CLI_TG_per_rut_per = :2
    AND b.ta_obn_num_ord = A.num_ord
    AND b.ta_ult_trs_cor_trs = c.cor_trs
    AND SRU.TG_CLI_TG_PER_RUT_PER = a.TG_SRU_TG_CLI_TG_PER_RUT_PER
    AND SRU.SRU_SRU = a.TG_SRU_SRU_SRU
GROUP BY
    sru.cnp_sru,
    A.tip_ope_ord,
    A.ta_ser_nom_ser
"""


def _load_keyring_credential(credential_name: str) -> tuple[str, str]:
    try:
        import keyring  # type: ignore
    except Exception as exc:
        raise RuntimeError("Falta dependencia 'keyring' para leer credenciales de Windows.") from exc

    try:
        cred = keyring.get_credential(credential_name, None)
    except Exception as exc:
        raise RuntimeError(f"Error al leer credencial '{credential_name}' en keyring: {exc}") from exc

    if cred is None or not getattr(cred, "username", None) or not getattr(cred, "password", None):
        raise RuntimeError(f"No se encontro la credencial '{credential_name}' en Windows Credential Manager.")
    return str(cred.username), str(cred.password)


@contextmanager
def oracle_session() -> Iterator[object]:
    try:
        import cx_Oracle  # type: ignore
    except Exception as exc:
        raise RuntimeError("Falta dependencia 'cx_Oracle' para consultas Oracle.") from exc

    user, password = _load_keyring_credential(DEFAULT_CREDENTIAL_NAME)
    dsn = cx_Oracle.makedsn(
        DEFAULT_ORACLE_HOST,
        DEFAULT_ORACLE_PORT,
        service_name=DEFAULT_ORACLE_SERVICE,
    )
    connection = None
    try:
        connection = cx_Oracle.connect(
            user=user,
            password=password,
            dsn=dsn,
            encoding="UTF-8",
        )
        yield connection
    except Exception as exc:
        raise RuntimeError(f"Error al conectar a Oracle: {exc}") from exc
    finally:
        try:
            del user
            del password
        except Exception:
            pass
        if connection is not None:
            try:
                connection.close()
            except Exception:
                pass


def _rut_without_dv(rut_norm: str) -> str:
    rut_norm = (rut_norm or "").strip().upper()
    return rut_norm[:-1] if len(rut_norm) > 1 else rut_norm


def query_folios(connection, rut_norm: str, date_label: str) -> pd.DataFrame:
    try:
        return pd.read_sql(FOLIOS_SQL, connection, params=(date_label, _rut_without_dv(rut_norm)))
    except Exception:
        return pd.DataFrame()


def query_summary(connection, rut_norm: str, date_label: str) -> pd.DataFrame:
    try:
        return pd.read_sql(SUMMARY_SQL, connection, params=(date_label, _rut_without_dv(rut_norm)))
    except Exception:
        return pd.DataFrame()


def query_custom(connection, sql_text: str, date_label: str) -> pd.DataFrame:
    if not (sql_text or "").strip():
        return pd.DataFrame()
    try:
        return pd.read_sql(sql_text, connection, params=[date_label])
    except Exception as exc:
        LOGGER.exception("Error ejecutando query personalizada con fecha %s: %s", date_label, exc)
        return pd.DataFrame()
