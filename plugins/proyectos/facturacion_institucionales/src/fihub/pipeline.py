# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 09:21:03 2025

@author: BBRUNA
"""

from __future__ import annotations

import logging
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable

import pandas as pd

from .master_store import MasterWorkbookStore
from .models import CustomerProfile, ExecutionReport, RuntimeOptions, format_reference_date
from .services.db_oracle import oracle_session, query_custom, query_folios, query_summary
from .services.dte_fetcher import download_documents
from .services.file_bundle import (
    ensure_empty_folder,
    merge_pdfs_in_place,
    write_excel,
    write_flat_file,
    zip_xmls_in_place,
)
from .services.mailer import create_outlook_draft
from .services.reporting import build_summary_pdf


LOGGER = logging.getLogger(__name__)
INVALID_NAME_RE = re.compile(r'[<>:"/\\\\|?*]+')


def _safe_name(value: str, fallback: str) -> str:
    cleaned = INVALID_NAME_RE.sub("_", (value or "").strip()).strip(" .")
    return cleaned or fallback


def _log(cb: Callable[[str], None] | None, message: str) -> None:
    if cb:
        cb(message)


def _rut_display(profile: CustomerProfile) -> str:
    if profile.rut_sin_dv and profile.rut_dv:
        return f"{profile.rut_sin_dv}-{profile.rut_dv}"
    return profile.rut_norm


def _pick_reference_date_and_folios(connection, profile: CustomerProfile) -> tuple[pd.DataFrame, str]:
    today_label = format_reference_date(date.today())
    yesterday_label = format_reference_date(date.today() - timedelta(days=1))
    folios_today = query_folios(connection, profile.rut_norm, today_label)
    if not folios_today.empty:
        return folios_today, today_label
    return query_folios(connection, profile.rut_norm, yesterday_label), yesterday_label


def _sort_summary_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    rename_map = {col: str(col).upper() for col in df.columns}
    df = df.rename(columns=rename_map)
    if "OPERACION" in df.columns:
        return df.sort_values(by="OPERACION")
    return df


def _minus_one_day(date_label: str) -> str:
    try:
        parsed = datetime.strptime(date_label, "%d-%m-%Y").date()
        return format_reference_date(parsed - timedelta(days=1))
    except Exception:
        return format_reference_date(date.today() - timedelta(days=1))


def _query_from_store(
    store: MasterWorkbookStore,
    query_id: str,
    kind_label: str,
    report: ExecutionReport,
) -> str | None:
    if not query_id:
        report.add_warning(f"{kind_label}: no existe query_id configurado.")
        return None
    query_def = store.get_query(query_id)
    if query_def is None:
        report.add_warning(f"{kind_label}: query_id '{query_id}' no encontrada en hoja QUERIES.")
        return None
    if query_def.query_status not in {"QUERY", "PENDIENTE_REVISAR"}:
        report.add_warning(f"{kind_label}: estado de query '{query_def.query_status}' no ejecutable.")
        return None
    return query_def.sql_text


def execute_customer_flow(
    *,
    store: MasterWorkbookStore,
    profile: CustomerProfile,
    options: RuntimeOptions,
    logger,
    requested_actions: list[str] | None = None,
    log_callback: Callable[[str], None] | None = None,
) -> ExecutionReport:
    output_folder = options.output_base_dir / _safe_name(profile.output_folder_name, profile.rut_norm)
    result = ExecutionReport(customer=profile, output_dir=output_folder)

    available_actions = store.describe_actions(profile)
    available_set = set(available_actions)
    if requested_actions:
        selected_actions = [a for a in requested_actions if a in available_set]
    else:
        selected_actions = list(available_actions)
    if not selected_actions:
        selected_actions = list(available_actions)
    selected_set = set(selected_actions)

    generation_actions = {
        "FACTURA_PDF",
        "FACTURA_XML",
        "RESUMEN_PDF",
        "REPORTE_PROCESO_DATOS",
        "ARCHIVO_PLANO_OMS",
    }
    will_generate_files = any(action in generation_actions for action in selected_set)
    if will_generate_files:
        ensure_empty_folder(output_folder)
    else:
        output_folder.mkdir(parents=True, exist_ok=True)

    _log(log_callback, f"[INICIO] Cliente {profile.display_name} ({_rut_display(profile)})")
    _log(log_callback, f"[INFO] Carpeta de salida: {output_folder}")
    _log(log_callback, f"[INFO] Acciones solicitadas: {', '.join(selected_actions) or 'ninguna'}")

    wants_factura_pdf = "FACTURA_PDF" in selected_set and profile.factura_pdf
    wants_factura_xml = "FACTURA_XML" in selected_set and profile.factura_xml
    wants_resumen = "RESUMEN_PDF" in selected_set and profile.resumen_pdf
    wants_report = "REPORTE_PROCESO_DATOS" in selected_set
    wants_flat = "ARCHIVO_PLANO_OMS" in selected_set
    wants_manual_mail = "ENVIO_MANUAL_OUTLOOK_DISPLAY" in selected_set
    wants_tomy = "ENTREGA_TOMY_CARPETA" in selected_set
    mail_only_mode = wants_manual_mail and not any(
        [wants_factura_pdf, wants_factura_xml, wants_resumen, wants_report, wants_flat, wants_tomy]
    )

    db_required = any([wants_factura_pdf, wants_factura_xml, wants_resumen, wants_report, wants_flat])
    reference_date = format_reference_date(date.today())
    result.reference_date = reference_date
    folios_df = pd.DataFrame()

    if db_required:
        try:
            with oracle_session() as connection:
                try:
                    folios_df, reference_date = _pick_reference_date_and_folios(connection, profile)
                    result.reference_date = reference_date
                    _log(log_callback, f"[INFO] Fecha de referencia: {reference_date}")
                except Exception as exc:
                    result.add_warning(f"No fue posible obtener fecha/folios de referencia: {exc}")
                    _log(log_callback, f"[WARN] Fecha/folios: {exc}")
                    folios_df = pd.DataFrame()
                    reference_date = format_reference_date(date.today())
                    result.reference_date = reference_date

                need_folios = wants_factura_pdf or wants_factura_xml
                folios = [str(v).strip() for v in folios_df.get("Folio", [])] if not folios_df.empty else []
                if folios:
                    _log(log_callback, f"[INFO] Folios encontrados: {len(folios)}")
                elif need_folios:
                    result.add_warning("No se encontraron folios para hoy/ayer; se omiten descargas de factura/XML.")
                    _log(log_callback, "[WARN] Sin folios para descargas de factura/XML")

                if wants_factura_pdf or wants_factura_xml:
                    try:
                        downloaded = download_documents(
                            output_folder=output_folder,
                            folios=folios,
                            include_pdf=wants_factura_pdf,
                            include_xml=wants_factura_xml,
                            log_callback=log_callback,
                        )
                        for path in downloaded:
                            result.add_file(path)
                    except Exception as exc:
                        result.add_warning(f"Error en descarga de documentos: {exc}")
                        _log(log_callback, f"[WARN] Descarga documentos: {exc}")

                    if wants_factura_pdf:
                        try:
                            merged = merge_pdfs_in_place(output_folder, password=profile.decryption_key or None)
                            if merged is not None:
                                result.add_file(merged)
                                _log(log_callback, f"[OK] PDF masivo: {merged.name}")
                            else:
                                result.add_warning("No fue posible crear el PDF masivo (sin PDFs utilizables).")
                        except Exception as exc:
                            result.add_warning(f"Error combinando PDFs: {exc}")
                            _log(log_callback, f"[WARN] Combinar PDFs: {exc}")

                    if wants_factura_xml:
                        try:
                            zipped = zip_xmls_in_place(output_folder)
                            if zipped is not None:
                                result.add_file(zipped)
                                _log(log_callback, f"[OK] ZIP XML: {zipped.name}")
                            else:
                                result.add_warning("No se genero ZIP XML porque no habia XMLs.")
                        except Exception as exc:
                            result.add_warning(f"Error comprimiendo XML: {exc}")
                            _log(log_callback, f"[WARN] Comprimir XML: {exc}")

                if wants_resumen:
                    try:
                        summary_df = query_summary(connection, profile.rut_norm, reference_date)
                        if summary_df.empty:
                            alt_date = _minus_one_day(reference_date)
                            summary_df = query_summary(connection, profile.rut_norm, alt_date)
                            if not summary_df.empty:
                                result.reference_date = alt_date
                                reference_date = alt_date

                        if summary_df.empty:
                            result.add_warning("No se encontraron datos para el resumen de transacciones.")
                        else:
                            summary_df = _sort_summary_df(summary_df)
                            summary_path = (
                                output_folder
                                / f"ResumenTransacciones_{_safe_name(profile.output_folder_name, profile.rut_norm)}.pdf"
                            )
                            build_summary_pdf(
                                summary_df,
                                rut_display=_rut_display(profile),
                                customer_name=profile.display_name,
                                reference_date=reference_date,
                                output_path=summary_path,
                            )
                            result.add_file(summary_path)
                            _log(log_callback, f"[OK] Resumen PDF: {summary_path.name}")
                    except Exception as exc:
                        result.add_warning(f"Error generando resumen PDF: {exc}")
                        _log(log_callback, f"[WARN] Resumen PDF: {exc}")

                if wants_report:
                    if profile.report_status == "QUERY":
                        try:
                            sql_text = _query_from_store(store, profile.report_query_id, "Reporte proceso de datos", result)
                            if sql_text:
                                report_df = query_custom(connection, sql_text, reference_date)
                                report_xlsx = (
                                    output_folder
                                    / f"{_safe_name(profile.output_folder_name, profile.rut_norm)}_proceso_datos.xlsx"
                                )
                                write_excel(report_df, report_xlsx)
                                result.add_file(report_xlsx)
                                _log(log_callback, f"[OK] Reporte Excel: {report_xlsx.name} ({len(report_df)} filas)")
                        except Exception as exc:
                            result.add_warning(f"Error generando reporte Excel: {exc}")
                            _log(log_callback, f"[WARN] Reporte Excel: {exc}")
                    elif profile.report_status == "PENDIENTE_QUERY":
                        result.add_warning(
                            "Este cliente requiere reporte Excel de proceso de datos, pero la query aun no esta cargada en el maestro."
                        )
                    elif profile.report_status == "PENDIENTE_REVISAR":
                        result.add_warning("Campo de reporte de proceso de datos requiere revision manual.")

                if wants_flat:
                    if profile.flat_status == "QUERY":
                        try:
                            sql_text = _query_from_store(store, profile.flat_query_id, "Archivo plano", result)
                            if sql_text:
                                flat_df = query_custom(connection, sql_text, reference_date)
                                flat_path = output_folder / f"{_safe_name(profile.output_folder_name, profile.rut_norm)}_oms.txt"
                                write_flat_file(flat_df, flat_path)
                                result.add_file(flat_path)
                                _log(log_callback, f"[OK] Archivo plano: {flat_path.name} ({len(flat_df)} filas)")
                        except Exception as exc:
                            result.add_warning(f"Error generando archivo plano: {exc}")
                            _log(log_callback, f"[WARN] Archivo plano: {exc}")
                    elif profile.flat_status == "PENDIENTE_QUERY":
                        result.add_warning(
                            "Este cliente requiere archivo plano OMS, pero la query aun no esta cargada en el maestro."
                        )
                    elif profile.flat_status == "PENDIENTE_REVISAR":
                        result.add_warning("Campo de archivo plano requiere revision manual.")
        except Exception as exc:
            result.add_warning(f"No fue posible conectar o consultar Oracle: {exc}")
            _log(log_callback, f"[WARN] Oracle: {exc}")

    delivery_mode = (profile.delivery_mode or "").strip().upper()
    if wants_manual_mail:
        recipients = store.get_group_emails(profile.mail_group or profile.group_name)
        if not recipients:
            result.add_warning("No hay destinatarios cargados para el grupo del correo manual.")
        else:
            if delivery_mode == "TOMY":
                result.add_warning(
                    "Correo manual forzado en cliente configurado para TOMY (continuando por solicitud del usuario)."
                )
                _log(log_callback, "[WARN] Correo manual forzado en cliente TOMY")
            elif delivery_mode not in {"MANUAL", "TOMY"}:
                result.add_warning("Correo manual solicitado en cliente sin modo MANUAL/TOMY definido.")

            if not any(output_folder.iterdir()):
                if mail_only_mode:
                    result.add_info("Correo manual solicitado sin descarga previa: se abrira correo sin adjuntos.")
                    _log(log_callback, "[INFO] Correo manual sin adjuntos (sin descargas seleccionadas)")
                else:
                    result.add_warning("La carpeta de salida no tiene archivos para adjuntar en el correo manual.")
            try:
                if profile.rut_norm == "969662507":
                    subject = (
                        "Facturas Banchile CDB BTG Renta Variable Local // "
                        f"{result.reference_date or format_reference_date(date.today())}"
                    )
                else:
                    subject = (
                        f"Op. RV Banchile // {profile.display_name} "
                        f"{result.reference_date or format_reference_date(date.today())}"
                    )
                create_outlook_draft(
                    to_recipients=recipients,
                    subject=subject,
                    body="Buenas tardes estimados, adjunto archivos del dia.",
                    attachments_dir=output_folder,
                )
                result.add_info("Se abrio borrador de correo manual en Outlook (Display) con adjuntos.")
                _log(log_callback, "[OK] Borrador de correo abierto en Outlook")
            except Exception as exc:
                result.add_warning(f"No se pudo abrir el correo manual en Outlook: {exc}")
                _log(log_callback, f"[WARN] Correo manual: {exc}")

    if wants_tomy:
        if delivery_mode != "TOMY":
            result.add_warning("Se solicito entrega TOMY, pero el cliente no esta configurado en modo TOMY.")
        else:
            result.add_info("Modo de envio TOMY: se deja la carpeta lista para el RPA con los archivos generados.")
            _log(log_callback, "[INFO] Modo TOMY: carpeta preparada para envio automatico")

    if profile.notes:
        result.add_info(f"Observacion de maestro: {profile.notes}")

    return result
