# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 10:03:09 2025

@author: BBRUNA
"""

from __future__ import annotations

from pathlib import Path

import fitz  # type: ignore
import pandas as pd


def _format_line(values: list[str], widths: list[int]) -> str:
    parts: list[str] = []
    for index, value in enumerate(values):
        text = (value or "").replace("\n", " ").strip()
        width = widths[index]
        if len(text) > width:
            text = text[: max(0, width - 1)] + "…"
        parts.append(text.ljust(width))
    return " | ".join(parts)


def build_summary_pdf(
    df: pd.DataFrame,
    *,
    rut_display: str,
    customer_name: str,
    reference_date: str,
    output_path: Path,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    table = df.copy()
    table.columns = [str(col).upper() for col in table.columns]

    ordered = [c for c in ["FONDO", "OPERACION", "INSTRUMENTO", "CANTIDAD", "MONTO"] if c in table.columns]
    if ordered:
        table = table[ordered]

    # widths tuned for A4 landscape monospace
    widths_map = {
        "FONDO": 18,
        "OPERACION": 12,
        "INSTRUMENTO": 34,
        "CANTIDAD": 14,
        "MONTO": 14,
    }
    widths = [widths_map.get(str(c).upper(), 18) for c in table.columns]

    doc = fitz.open()
    font_name = "cour"
    font_size = 10
    line_height = 13
    margin_x = 28
    margin_y = 28
    usable_bottom = 565

    def new_page():
        page = doc.new_page(width=842, height=595)  # A4 horizontal approx in points
        return page

    page = new_page()
    y = margin_y
    page.insert_text((margin_x, y), "Resumen de Transacciones", fontname="helv", fontsize=16)
    y += 22
    page.insert_text((margin_x, y), f"Cliente: {customer_name}", fontname="helv", fontsize=10)
    y += 14
    page.insert_text((margin_x, y), f"RUT: {rut_display}", fontname="helv", fontsize=10)
    y += 14
    page.insert_text((margin_x, y), f"Fecha operacion: {reference_date}", fontname="helv", fontsize=10)
    y += 22

    header = _format_line([str(c) for c in table.columns], widths)
    separator = "-" * min(140, len(header))
    page.insert_text((margin_x, y), header, fontname=font_name, fontsize=font_size)
    y += line_height
    page.insert_text((margin_x, y), separator, fontname=font_name, fontsize=font_size)
    y += line_height

    if table.empty:
        page.insert_text((margin_x, y), "Sin registros para el resumen.", fontname="helv", fontsize=10)
    else:
        for _, row in table.iterrows():
            if y > usable_bottom:
                page = new_page()
                y = margin_y
                page.insert_text((margin_x, y), header, fontname=font_name, fontsize=font_size)
                y += line_height
                page.insert_text((margin_x, y), separator, fontname=font_name, fontsize=font_size)
                y += line_height
            values = [str("" if pd.isna(v) else v) for v in row.tolist()]
            page.insert_text((margin_x, y), _format_line(values, widths), fontname=font_name, fontsize=font_size)
            y += line_height

    doc.save(output_path)
    doc.close()
    return output_path
