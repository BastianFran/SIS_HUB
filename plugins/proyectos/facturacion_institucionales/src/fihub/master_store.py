# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 09:07:01 2025

@author: BBRUNA
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from .models import CustomerProfile, QueryDefinition


RUT_CLEAN_RE = re.compile(r"[^0-9K]")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-']+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def _flag(value: Any) -> bool:
    return _clean_text(value).upper() in {"Y", "SI", "YES", "TRUE", "1"}


def _normalize_rut(value: Any) -> str:
    return RUT_CLEAN_RE.sub("", _clean_text(value).upper())


def _optional_float(value: Any) -> float | None:
    text = _clean_text(value)
    if not text:
        return None
    try:
        return float(text)
    except Exception:
        return None


def _optional_int(value: Any) -> int | None:
    text = _clean_text(value)
    if not text:
        return None
    try:
        return int(float(text))
    except Exception:
        return None


def _delivery_mode(value: Any, *, fallback_manual_flag: bool = False) -> str:
    token = _clean_text(value).upper()
    if token in {"MANUAL", "TOMY", "NINGUNO"}:
        return token
    if fallback_manual_flag:
        return "MANUAL"
    return "NINGUNO"


def _factura_flags_from_simple(value: Any) -> tuple[bool, bool]:
    token = _clean_text(value).upper()
    if token in {"PDF_XML", "XML"}:
        return True, True
    if token in {"PDF", "Y"}:
        return True, False
    return False, False


def _extract_emails_from_text(raw: str) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for match in EMAIL_RE.finditer(raw or ""):
        email = match.group(0).strip(" ,;")
        key = email.casefold()
        if key in seen:
            continue
        seen.add(key)
        result.append(email)
    return result


@dataclass(frozen=True)
class SearchResult:
    profile: CustomerProfile
    actions: list[str]
    notices: list[str]


class MasterWorkbookStore:
    def __init__(self, workbook_path: Path):
        self.workbook_path = workbook_path
        self._customers_by_rut: dict[str, list[CustomerProfile]] = {}
        self._queries: dict[str, QueryDefinition] = {}
        self._emails_by_group: dict[str, list[str]] = {}
        self._loaded = False
        self._clients_count = 0
        self._group_members_count = 0

    @property
    def loaded(self) -> bool:
        return self._loaded

    @property
    def customer_count(self) -> int:
        return self._clients_count + self._group_members_count

    @property
    def clients_count(self) -> int:
        return self._clients_count

    @property
    def group_members_count(self) -> int:
        return self._group_members_count

    def reload(self) -> None:
        if not self.workbook_path.exists():
            raise FileNotFoundError(f"No existe workbook maestro: {self.workbook_path}")

        workbook = pd.ExcelFile(self.workbook_path)
        sheets = set(workbook.sheet_names)

        self._customers_by_rut = {}

        # Maestro simple (nuevo)
        if {"CLIENTES", "QUERIES", "GRUPOS_CORREO"} <= sheets:
            df_clients = pd.read_excel(self.workbook_path, sheet_name="CLIENTES")
            df_queries = pd.read_excel(self.workbook_path, sheet_name="QUERIES")
            df_groups = pd.read_excel(self.workbook_path, sheet_name="GRUPOS_CORREO")

            self._queries = self._load_queries_simple(df_queries)
            self._emails_by_group = self._load_group_emails_simple(df_groups)

            total = self._load_customers_simple(df_clients)
            self._clients_count = total
            self._group_members_count = 0
            self._loaded = True
            return

        # Maestro anterior (compatibilidad temporal)
        required = {"CLIENTES", "GRUPOS_MIEMBROS", "QUERIES", "CORREOS_EMAILS"}
        missing = required - sheets
        if missing:
            raise ValueError(f"Faltan hojas requeridas en el maestro: {sorted(missing)}")

        df_clients = pd.read_excel(self.workbook_path, sheet_name="CLIENTES")
        df_group_members = pd.read_excel(self.workbook_path, sheet_name="GRUPOS_MIEMBROS")
        df_queries = pd.read_excel(self.workbook_path, sheet_name="QUERIES")
        df_emails = pd.read_excel(self.workbook_path, sheet_name="CORREOS_EMAILS")

        self._queries = self._load_queries_legacy(df_queries)
        self._emails_by_group = self._load_group_emails_legacy(df_emails)

        self._clients_count = self._load_customers_legacy(df_clients, "CLIENTE")
        self._group_members_count = self._load_customers_legacy(df_group_members, "GRUPO")
        self._loaded = True

    def _load_customers_simple(self, df: pd.DataFrame) -> int:
        count = 0
        for _, row in df.iterrows():
            row_map = {str(k): row[k] for k in df.columns}
            rut_norm = _normalize_rut(row_map.get("rut") or row_map.get("rut_norm"))
            if not rut_norm:
                continue

            factura_pdf, factura_xml = _factura_flags_from_simple(row_map.get("factura"))
            delivery_mode = _delivery_mode(row_map.get("modo_envio"))

            profile = CustomerProfile(
                record_id=_clean_text(row_map.get("id_registro") or row_map.get("record_id")),
                source_sheet=_clean_text(row_map.get("tipo_registro")) or "CLIENTE",
                rut_norm=rut_norm,
                rut_sin_dv=_clean_text(row_map.get("rut_sin_dv")),
                rut_dv=_clean_text(row_map.get("rut_dv")),
                display_name=_clean_text(row_map.get("cliente")) or _clean_text(row_map.get("cliente_nombre")) or rut_norm,
                output_folder_name=(
                    _clean_text(row_map.get("carpeta"))
                    or _clean_text(row_map.get("carpeta_salida"))
                    or _clean_text(row_map.get("cliente"))
                    or rut_norm
                ),
                group_name=_clean_text(row_map.get("grupo")) or _clean_text(row_map.get("grupo_nombre")),
                active=_flag(row_map.get("activo") or "Y"),
                factura_pdf=factura_pdf,
                factura_xml=factura_xml,
                resumen_pdf=_flag(row_map.get("resumen_pdf")),
                report_status=_clean_text(row_map.get("reporte_proceso")).upper() or "NO",
                report_query_id=_clean_text(row_map.get("reporte_query_id")),
                flat_status=_clean_text(row_map.get("archivo_plano_oms")).upper() or "NO",
                flat_query_id=_clean_text(row_map.get("archivo_plano_query_id")),
                delivery_mode=delivery_mode,
                mail_group=_clean_text(row_map.get("grupo_correo")) or _clean_text(row_map.get("grupo")),
                decryption_key=_clean_text(row_map.get("clave_archivos")),
                max_mail_mb=_optional_float(row_map.get("tamano_max_mb")),
                notes=_clean_text(row_map.get("observaciones")),
                raw_row={k: ("" if pd.isna(v) else v) for k, v in row_map.items()},
            )
            self._customers_by_rut.setdefault(rut_norm, []).append(profile)
            count += 1
        return count

    def _load_customers_legacy(self, df: pd.DataFrame, source_kind: str) -> int:
        count = 0
        for _, row in df.iterrows():
            row_map = {str(k): row[k] for k in df.columns}
            rut_norm = _normalize_rut(row_map.get("rut_norm") or row_map.get("rut_origen"))
            if not rut_norm:
                continue

            display_name = _clean_text(row_map.get("cliente_nombre"))
            folder_name = _clean_text(row_map.get("carpeta_salida")) or display_name or rut_norm
            group_name = _clean_text(row_map.get("grupo_nombre"))
            decryption_key = _clean_text(row_map.get("clave_archivos_efectiva"))

            manual_flag = _flag(row_map.get("correo_manual_habilitado"))
            inferred_mode = "MANUAL" if manual_flag else "TOMY"

            profile = CustomerProfile(
                record_id=_clean_text(row_map.get("record_id")),
                source_sheet=source_kind,
                rut_norm=rut_norm,
                rut_sin_dv=_clean_text(row_map.get("rut_sin_dv")),
                rut_dv=_clean_text(row_map.get("rut_dv")),
                display_name=display_name or folder_name,
                output_folder_name=folder_name,
                group_name=group_name,
                active=_flag(row_map.get("activo") or "Y"),
                factura_pdf=_flag(row_map.get("factura_pdf")),
                factura_xml=_flag(row_map.get("factura_xml")),
                resumen_pdf=_flag(row_map.get("resumen_pdf")),
                report_status=_clean_text(row_map.get("reporte_proceso_estado")).upper() or "NO",
                report_query_id=_clean_text(row_map.get("reporte_proceso_query_id")),
                flat_status=_clean_text(row_map.get("archivo_plano_estado")).upper() or "NO",
                flat_query_id=_clean_text(row_map.get("archivo_plano_query_id")),
                delivery_mode=inferred_mode,
                mail_group=_clean_text(row_map.get("correo_grupo")) or group_name,
                decryption_key=decryption_key,
                max_mail_mb=_optional_float(row_map.get("tamano_max_mb_efectivo")),
                notes=_clean_text(row_map.get("observaciones")),
                raw_row={k: ("" if pd.isna(v) else v) for k, v in row_map.items()},
            )
            self._customers_by_rut.setdefault(rut_norm, []).append(profile)
            count += 1
        return count

    def _load_queries_simple(self, df: pd.DataFrame) -> dict[str, QueryDefinition]:
        queries: dict[str, QueryDefinition] = {}
        for _, row in df.iterrows():
            query_id = _clean_text(row.get("query_id"))
            if not query_id:
                continue
            queries[query_id] = QueryDefinition(
                query_id=query_id,
                query_kind=_clean_text(row.get("tipo") or row.get("query_kind")).upper(),
                query_status=_clean_text(row.get("estado") or row.get("query_status")).upper() or "QUERY",
                sql_text=_clean_text(row.get("sql_text")),
                start_row=_optional_int(row.get("fila_inicio")),
                emitter_column=_clean_text(row.get("emisor_columna")),
            )
        return queries

    def _load_queries_legacy(self, df: pd.DataFrame) -> dict[str, QueryDefinition]:
        queries: dict[str, QueryDefinition] = {}
        for _, row in df.iterrows():
            query_id = _clean_text(row.get("query_id"))
            if not query_id:
                continue
            queries[query_id] = QueryDefinition(
                query_id=query_id,
                query_kind=_clean_text(row.get("query_kind")).upper(),
                query_status=_clean_text(row.get("query_status")).upper() or "QUERY",
                sql_text=_clean_text(row.get("sql_text")),
                start_row=_optional_int(row.get("fila_inicio")),
                emitter_column=_clean_text(row.get("emisor_columna")),
            )
        return queries

    def _load_group_emails_simple(self, df: pd.DataFrame) -> dict[str, list[str]]:
        groups: dict[str, list[str]] = {}
        for _, row in df.iterrows():
            group_name = _clean_text(row.get("grupo"))
            if not group_name:
                continue
            emails = _extract_emails_from_text(_clean_text(row.get("destinatarios")))
            if not emails:
                continue
            key = group_name.casefold()
            groups.setdefault(key, [])
            seen = {x.casefold() for x in groups[key]}
            for email in emails:
                if email.casefold() in seen:
                    continue
                groups[key].append(email)
                seen.add(email.casefold())
        return groups

    def _load_group_emails_legacy(self, df: pd.DataFrame) -> dict[str, list[str]]:
        groups: dict[str, list[str]] = {}
        for _, row in df.iterrows():
            group_name = _clean_text(row.get("grupo_nombre"))
            if not group_name:
                continue
            email = _clean_text(row.get("email"))
            if not email:
                continue
            key = group_name.casefold()
            groups.setdefault(key, [])
            seen = {x.casefold() for x in groups[key]}
            if email.casefold() not in seen:
                groups[key].append(email)
        return groups

    def get_query(self, query_id: str) -> QueryDefinition | None:
        return self._queries.get((query_id or "").strip())

    def get_group_emails(self, group_name: str) -> list[str]:
        return list(self._emails_by_group.get((group_name or "").casefold(), []))

    def find_customer(self, rut_input: str) -> SearchResult | None:
        if not self._loaded:
            self.reload()

        rut_norm = _normalize_rut(rut_input)
        candidates = self._customers_by_rut.get(rut_norm, [])
        if not candidates:
            return None

        active_candidates = [p for p in candidates if p.active]
        chosen = active_candidates[0] if active_candidates else candidates[0]
        return SearchResult(
            profile=chosen,
            actions=self.describe_actions(chosen),
            notices=self.describe_notices(chosen),
        )

    def describe_actions(self, profile: CustomerProfile) -> list[str]:
        actions: list[str] = []
        if profile.factura_pdf:
            actions.append("FACTURA_PDF")
        if profile.factura_xml:
            actions.append("FACTURA_XML")
        if profile.resumen_pdf:
            actions.append("RESUMEN_PDF")
        if profile.report_status == "QUERY" and profile.report_query_id:
            actions.append("REPORTE_PROCESO_DATOS")
        if profile.flat_status == "QUERY" and profile.flat_query_id:
            actions.append("ARCHIVO_PLANO_OMS")
        if profile.manual_mail_enabled:
            actions.append("ENVIO_MANUAL_OUTLOOK_DISPLAY")
        elif profile.tomy_delivery_enabled:
            actions.append("ENTREGA_TOMY_CARPETA")
            actions.append("ENVIO_MANUAL_OUTLOOK_DISPLAY")
        return actions

    def describe_notices(self, profile: CustomerProfile) -> list[str]:
        notices: list[str] = []
        if profile.report_status == "PENDIENTE_QUERY":
            notices.append("El cliente requiere reporte Excel de proceso de datos, pero la query aun no esta definida.")
        elif profile.report_status == "PENDIENTE_REVISAR":
            notices.append("El campo de reporte Excel tiene un valor no estandar; revisar el maestro.")

        if profile.flat_status == "PENDIENTE_QUERY":
            notices.append("El cliente requiere archivo plano OMS, pero la query aun no esta definida.")
        elif profile.flat_status == "PENDIENTE_REVISAR":
            notices.append("El campo de archivo plano tiene un valor no estandar; revisar el maestro.")

        if profile.manual_mail_enabled:
            notices.append("Modo de envio: MANUAL (abre Outlook en Display con adjuntos y destinatarios del grupo).")
        elif profile.tomy_delivery_enabled:
            notices.append("Modo de envio: TOMY (solo prepara carpeta de salida para el RPA).")
            notices.append("Puedes forzar correo manual desde la UI; el sistema pedira confirmacion.")
            if not self.get_group_emails(profile.mail_group or profile.group_name):
                notices.append("Advertencia: este grupo no tiene destinatarios cargados para correo manual.")
        elif profile.mail_group:
            notices.append("Sin modo de envio definido. Revisa columna 'modo_envio' en el maestro.")

        if profile.notes:
            notices.append(f"Observacion: {profile.notes}")
        return notices
