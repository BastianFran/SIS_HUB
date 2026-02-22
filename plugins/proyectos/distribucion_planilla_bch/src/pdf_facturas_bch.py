# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 09:27:18 2025

@author: BBRUNA
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Sequence

import fitz  # PyMuPDF
import pandas as pd


_RE_FOLIO = re.compile(r"N[º°o]\s*([0-9]{5,})", re.IGNORECASE)
_RE_TIPO = re.compile(r"\b(Compra|Venta)\s+De\s+Renta\s+Variable\b", re.IGNORECASE)
_RE_NUM_TOKEN = re.compile(r"^[0-9][0-9\.,]*$")
_RE_ONLY_DIGITS = re.compile(r"^\d+$")


@dataclass(frozen=True)
class FacturaFilaExtraida:
    folio: int
    cantidad: str
    precio_unitario: str
    monto_neto: str
    tipo_factura: str
    origen_pdf: str
    pagina_pdf: int


def _normalizar_lineas(texto: str) -> List[str]:
    return [ln.strip() for ln in texto.splitlines() if ln and ln.strip()]


def _buscar_indice(lineas: Sequence[str], patron: str) -> int:
    p = patron.lower()
    for idx, ln in enumerate(lineas):
        if p in ln.lower():
            return idx
    return -1


def _extraer_folio(texto: str) -> int:
    m = _RE_FOLIO.search(texto)
    if not m:
        raise ValueError("No se encontro numero de factura (folio) en la pagina.")
    return int(m.group(1))


def _extraer_tipo(texto: str) -> str:
    m = _RE_TIPO.search(texto)
    if not m:
        return "Desconocido"
    tipo = m.group(1).strip().lower()
    return "Compra" if tipo == "compra" else "Venta"


def _es_token_numerico(linea: str) -> bool:
    if not _RE_NUM_TOKEN.match(linea):
        return False
    # Evita tomar lineas tipo fechas con guiones (ya filtradas por regex) o vacias.
    return True


def _parece_cantidad(token: str) -> bool:
    return "," in token


def _parece_precio(token: str) -> bool:
    return "," in token


def _parece_monto(token: str) -> bool:
    return _RE_ONLY_DIGITS.match(token.replace(".", "")) is not None


def _extraer_tokens_tabla(lineas: Sequence[str]) -> List[str]:
    """
    Extrae tokens numericos del bloque de detalle.

    Regla BCH observada:
    - El detalle aparece despues de la linea "Acciones Con Presencia Bursatil..."
    - Antes de "Monto Operacion"
    - Los numericos vienen como tripletas:
      cantidad / precio unitario / monto (compra o venta)
    """
    idx_inicio = _buscar_indice(lineas, "Acciones Con Presencia Burs")
    idx_fin = _buscar_indice(lineas, "Monto Operaci")

    if idx_inicio < 0:
        raise ValueError("No se encontro el inicio del detalle (bloque de acciones).")
    if idx_fin < 0:
        raise ValueError("No se encontro el fin del detalle (Monto Operacion).")
    if idx_fin <= idx_inicio:
        raise ValueError("Bloque de detalle invalido en la pagina.")

    candidatos = []
    for ln in lineas[idx_inicio + 1 : idx_fin]:
        if _es_token_numerico(ln):
            candidatos.append(ln)

    if not candidatos:
        raise ValueError("No se encontraron lineas numericas en el detalle de la factura.")

    return candidatos


def _tripletas_desde_tokens(tokens: Sequence[str]) -> List[tuple[str, str, str]]:
    """
    Convierte tokens numericos a tripletas cantidad/precio/monto.

    Se usa agrupacion secuencial y se valida el patron de cada tripleta.
    """
    if len(tokens) % 3 != 0:
        raise ValueError(
            f"El detalle numerico no forma tripletas exactas (tokens={len(tokens)})."
        )

    tripletas: List[tuple[str, str, str]] = []
    for i in range(0, len(tokens), 3):
        cantidad, precio, monto = tokens[i], tokens[i + 1], tokens[i + 2]
        if not (_parece_cantidad(cantidad) and _parece_precio(precio) and _parece_monto(monto)):
            raise ValueError(
                "No se pudo interpretar una tripleta del detalle "
                f"(cantidad={cantidad!r}, precio={precio!r}, monto={monto!r})."
            )
        tripletas.append((cantidad, precio, monto))
    return tripletas


def extraer_facturas_desde_pdf(ruta_pdf: Path) -> List[FacturaFilaExtraida]:
    ruta_pdf = Path(ruta_pdf)
    if not ruta_pdf.exists():
        raise FileNotFoundError(f"No existe el PDF: {ruta_pdf}")

    filas: List[FacturaFilaExtraida] = []
    doc = fitz.open(ruta_pdf)
    try:
        for pagina_idx in range(doc.page_count):
            texto = doc[pagina_idx].get_text("text")
            lineas = _normalizar_lineas(texto)
            if not lineas:
                continue

            try:
                folio = _extraer_folio(texto)
                tipo = _extraer_tipo(texto)
                tokens = _extraer_tokens_tabla(lineas)
                tripletas = _tripletas_desde_tokens(tokens)
            except Exception as exc:
                raise ValueError(
                    f"Error leyendo {ruta_pdf.name} pagina {pagina_idx + 1}: {exc}"
                ) from exc

            for cantidad, precio, monto in tripletas:
                filas.append(
                    FacturaFilaExtraida(
                        folio=folio,
                        cantidad=cantidad,
                        precio_unitario=precio,
                        monto_neto=monto,
                        tipo_factura=tipo,
                        origen_pdf=ruta_pdf.name,
                        pagina_pdf=pagina_idx + 1,
                    )
                )
    finally:
        doc.close()

    if not filas:
        raise ValueError(f"No se extrajeron lineas de factura desde {ruta_pdf.name}.")

    return filas


def extraer_facturas_desde_pdfs(rutas_pdf: Iterable[Path | str]) -> pd.DataFrame:
    filas: List[FacturaFilaExtraida] = []
    rutas = [Path(p) for p in rutas_pdf]
    if not rutas:
        raise ValueError("Debes seleccionar al menos un PDF.")

    for ruta in rutas:
        filas.extend(extraer_facturas_desde_pdf(ruta))

    df = pd.DataFrame(
        [
            {
                "folio": int(f.folio),
                "cantidad": f.cantidad,
                "precio_unitario": f.precio_unitario,
                "monto_neto": f.monto_neto,
                "_tipo_factura": f.tipo_factura,
                "_origen_pdf": f.origen_pdf,
                "_pagina_pdf": int(f.pagina_pdf),
            }
            for f in filas
        ]
    )

    # Orden estable por folio / origen / pagina / fila.
    if not df.empty:
        df = df.sort_values(
            by=["folio", "_origen_pdf", "_pagina_pdf"],
            kind="stable",
        ).reset_index(drop=True)
    return df

