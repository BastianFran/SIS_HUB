# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 10:31:13 2025

@author: BBRUNA
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd


PLUGIN_ROOT = Path(__file__).resolve().parents[1]
MASTER_PATH = PLUGIN_ROOT / "Facturacion_Institucionales_Maestro.xlsx"


def _text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def _flag(value: Any) -> bool:
    return _text(value).upper() in {"Y", "SI", "YES", "TRUE", "1"}


def _opt_float(value: Any) -> float | None:
    t = _text(value)
    if not t:
        return None
    try:
        return float(t)
    except Exception:
        return None


def _opt_int(value: Any) -> int | None:
    t = _text(value)
    if not t:
        return None
    try:
        return int(float(t))
    except Exception:
        return None


def _factura_mode(pdf_flag: Any, xml_flag: Any) -> str:
    pdf = _flag(pdf_flag)
    xml = _flag(xml_flag)
    if pdf and xml:
        return "PDF_XML"
    if pdf:
        return "PDF"
    return "NO"


def _delivery_mode_from_row(row: dict[str, Any]) -> str:
    if _flag(row.get("correo_manual_habilitado")):
        return "MANUAL"
    has_outputs = any(
        [
            _flag(row.get("factura_pdf")),
            _flag(row.get("factura_xml")),
            _flag(row.get("resumen_pdf")),
            _text(row.get("reporte_proceso_estado")).upper() in {"QUERY", "PENDIENTE_QUERY", "PENDIENTE_REVISAR"},
            _text(row.get("archivo_plano_estado")).upper() in {"QUERY", "PENDIENTE_QUERY", "PENDIENTE_REVISAR"},
        ]
    )
    return "TOMY" if has_outputs else "NINGUNO"


def _build_clients_sheet(df_clients: pd.DataFrame, df_groups: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for source_name, source_df in [("CLIENTE", df_clients), ("GRUPO", df_groups)]:
        for _, row in source_df.iterrows():
            r = {str(k): row[k] for k in source_df.columns}
            rows.append(
                {
                    "id_registro": _text(r.get("record_id")),
                    "tipo_registro": source_name,
                    "activo": "Y" if _flag(r.get("activo") or "Y") else "N",
                    "prioridad": _opt_int(r.get("prioridad")),
                    "rut": _text(r.get("rut_origen")) or _text(r.get("rut_norm")),
                    "rut_sin_dv": _text(r.get("rut_sin_dv")),
                    "rut_dv": _text(r.get("rut_dv")),
                    "cliente": _text(r.get("cliente_nombre")),
                    "grupo": _text(r.get("grupo_nombre")),
                    "carpeta": _text(r.get("carpeta_salida")) or _text(r.get("cliente_nombre")),
                    "factura": _factura_mode(r.get("factura_pdf"), r.get("factura_xml")),
                    "resumen_pdf": "Y" if _flag(r.get("resumen_pdf")) else "N",
                    "reporte_proceso": _text(r.get("reporte_proceso_estado")).upper() or "NO",
                    "reporte_query_id": _text(r.get("reporte_proceso_query_id")),
                    "reporte_emisor_columna": _text(r.get("reporte_proceso_emisor_columna")),
                    "reporte_fila_inicio": _opt_int(r.get("reporte_proceso_fila_inicio")),
                    "archivo_plano_oms": _text(r.get("archivo_plano_estado")).upper() or "NO",
                    "archivo_plano_query_id": _text(r.get("archivo_plano_query_id")),
                    "modo_envio": _delivery_mode_from_row(r),
                    "grupo_correo": _text(r.get("correo_grupo")) or _text(r.get("grupo_nombre")),
                    "clave_archivos": _text(r.get("clave_archivos_efectiva")),
                    "tamano_max_mb": _opt_float(r.get("tamano_max_mb_efectivo")),
                    "observaciones": _text(r.get("observaciones")),
                }
            )

    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(by=["tipo_registro", "prioridad", "cliente"], na_position="last")
    return out


def _build_queries_sheet(df_queries: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df_queries.iterrows():
        r = {str(k): row[k] for k in df_queries.columns}
        rows.append(
            {
                "query_id": _text(r.get("query_id")),
                "tipo": _text(r.get("query_kind")),
                "estado": _text(r.get("query_status")),
                "sql_text": _text(r.get("sql_text")),
                "emisor_columna": _text(r.get("emisor_columna")),
                "fila_inicio": _opt_int(r.get("fila_inicio")),
                "source_sheet": _text(r.get("source_sheet")),
                "source_record_id": _text(r.get("source_record_id")),
                "rut": _text(r.get("rut_norm")),
                "cliente": _text(r.get("cliente_nombre")),
                "grupo": _text(r.get("grupo_nombre")),
            }
        )
    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(by=["source_sheet", "source_record_id", "tipo"], na_position="last")
    return out


def _build_groups_sheet(df_groups_mail: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df_groups_mail.iterrows():
        r = {str(k): row[k] for k in df_groups_mail.columns}
        group_name = _text(r.get("grupo_nombre"))
        if not group_name:
            continue
        rows.append(
            {
                "grupo": group_name,
                "destinatarios": _text(r.get("emails_origen")),
                "clave_archivos": _text(r.get("clave_archivos_grupo")),
                "tamano_max_mb": _opt_float(r.get("tamano_max_mb_grupo")),
                "modo_envio_default": "MANUAL" if _flag(r.get("correo_manual_habilitado")) else "TOMY",
                "observaciones": _text(r.get("nota_migracion")),
            }
        )
    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(by=["grupo"], na_position="last")
    return out


def _build_readme() -> pd.DataFrame:
    rows = [
        ("GENERAL", "uso", "Este es el maestro simplificado editable del plugin Facturacion Institucionales."),
        ("CLIENTES", "factura", "Valores permitidos: NO, PDF, PDF_XML."),
        ("CLIENTES", "reporte_proceso", "Valores: NO, PENDIENTE_QUERY, QUERY, PENDIENTE_REVISAR."),
        ("CLIENTES", "archivo_plano_oms", "Valores: NO, PENDIENTE_QUERY, QUERY, PENDIENTE_REVISAR."),
        ("CLIENTES", "modo_envio", "Valores: MANUAL, TOMY, NINGUNO."),
        ("CLIENTES", "grupo_correo", "Debe coincidir con hoja GRUPOS_CORREO si modo_envio = MANUAL."),
        ("QUERIES", "sql_text", "Se usa :1 para fecha en formato DD-MM-YYYY."),
        ("GRUPOS_CORREO", "destinatarios", "Se pueden separar con coma o punto y coma."),
    ]
    return pd.DataFrame(rows, columns=["seccion", "campo", "detalle"])


def _build_rules() -> pd.DataFrame:
    rows = [
        ("Y/No legacy", "Se transformo a valores mas simples para edicion manual."),
        ("MANUAL", "Abre Outlook Display con adjuntos y destinatarios de GRUPOS_CORREO."),
        ("TOMY", "Solo deja carpeta preparada para el RPA."),
        ("carpeta", "El nombre de carpeta de salida se toma desde esta columna."),
    ]
    return pd.DataFrame(rows, columns=["concepto", "descripcion"])


def _build_meta(df_clients: pd.DataFrame, df_queries: pd.DataFrame, df_groups: pd.DataFrame, df_val: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        [
            ("schema", "simple_v1"),
            ("clientes_rows", len(df_clients)),
            ("queries_rows", len(df_queries)),
            ("grupos_correo_rows", len(df_groups)),
            ("validacion_rows", len(df_val)),
        ],
        columns=["key", "value"],
    )


def simplify_master(path: Path = MASTER_PATH) -> Path:
    if not path.exists():
        raise FileNotFoundError(f"No existe maestro: {path}")

    xl = pd.ExcelFile(path)
    if {"CLIENTES", "GRUPOS_CORREO", "QUERIES"} <= set(xl.sheet_names) and "GRUPOS_MIEMBROS" not in set(xl.sheet_names):
        return path
    needed = {"CLIENTES", "GRUPOS_MIEMBROS", "QUERIES", "CORREOS_GRUPO"}
    if not (needed <= set(xl.sheet_names)):
        raise ValueError("El archivo no tiene el schema tecnico esperado para simplificar.")

    df_clients = pd.read_excel(path, sheet_name="CLIENTES")
    df_group_members = pd.read_excel(path, sheet_name="GRUPOS_MIEMBROS")
    df_queries = pd.read_excel(path, sheet_name="QUERIES")
    df_group_mail = pd.read_excel(path, sheet_name="CORREOS_GRUPO")
    df_validation = pd.read_excel(path, sheet_name="VALIDACION") if "VALIDACION" in xl.sheet_names else pd.DataFrame()

    simple_clients = _build_clients_sheet(df_clients, df_group_members)
    simple_queries = _build_queries_sheet(df_queries)
    simple_groups = _build_groups_sheet(df_group_mail)
    readme = _build_readme()
    rules = _build_rules()
    meta = _build_meta(simple_clients, simple_queries, simple_groups, df_validation)

    if df_validation.empty:
        df_validation = pd.DataFrame(columns=["severity", "sheet", "record_key", "field", "detail"])

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        readme.to_excel(writer, sheet_name="README", index=False)
        meta.to_excel(writer, sheet_name="META", index=False)
        simple_clients.to_excel(writer, sheet_name="CLIENTES", index=False)
        simple_groups.to_excel(writer, sheet_name="GRUPOS_CORREO", index=False)
        simple_queries.to_excel(writer, sheet_name="QUERIES", index=False)
        rules.to_excel(writer, sheet_name="REGLAS", index=False)
        df_validation.to_excel(writer, sheet_name="VALIDACION", index=False)

    return path


if __name__ == "__main__":
    out = simplify_master()
    print(f"Maestro simplificado: {out}")
