# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 09:14:02 2025

@author: BBRUNA
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path


@dataclass(frozen=True)
class QueryDefinition:
    query_id: str
    query_kind: str
    query_status: str
    sql_text: str
    start_row: int | None = None
    emitter_column: str = ""


@dataclass(frozen=True)
class CustomerProfile:
    record_id: str
    source_sheet: str
    rut_norm: str
    rut_sin_dv: str
    rut_dv: str
    display_name: str
    output_folder_name: str
    group_name: str
    active: bool
    factura_pdf: bool
    factura_xml: bool
    resumen_pdf: bool
    report_status: str
    report_query_id: str
    flat_status: str
    flat_query_id: str
    delivery_mode: str
    mail_group: str
    decryption_key: str
    max_mail_mb: float | None
    notes: str
    raw_row: dict

    @property
    def has_any_download(self) -> bool:
        return self.factura_pdf or self.factura_xml

    @property
    def manual_mail_enabled(self) -> bool:
        return (self.delivery_mode or "").strip().upper() == "MANUAL"

    @property
    def tomy_delivery_enabled(self) -> bool:
        return (self.delivery_mode or "").strip().upper() == "TOMY"


@dataclass(frozen=True)
class RuntimeOptions:
    master_workbook: Path
    output_base_dir: Path


@dataclass
class ExecutionReport:
    customer: CustomerProfile
    output_dir: Path
    reference_date: str = ""
    generated_files: list[Path] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    info: list[str] = field(default_factory=list)

    def add_file(self, path: Path) -> None:
        self.generated_files.append(path)

    def add_warning(self, message: str) -> None:
        self.warnings.append(message)

    def add_info(self, message: str) -> None:
        self.info.append(message)


@dataclass(frozen=True)
class DependencyStatus:
    module_name: str
    package_name: str
    available: bool
    required_for: str


def format_reference_date(value: date) -> str:
    return value.strftime("%d-%m-%Y")
