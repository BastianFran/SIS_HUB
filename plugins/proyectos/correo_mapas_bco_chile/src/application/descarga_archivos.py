# -*- coding: utf-8 -*-
"""
Created on Tue Jan 20 11:12:36 2026

@author: BBRUNA
# -*- coding: utf-8 -*-

Flujo completo: descarga DTEs, genera Mapa, empaqueta por lotes (<=14MB)
y levanta correos Outlook en modo display con 2 ZIP por correo
(PDFs y XMLs), adjuntando el Mapa solo en el correo 1 de N.

Python 3.10.9 | Spyder
"""

from __future__ import annotations

import os
import math
import shutil
import tempfile
import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import getpass
import pandas as pd
import fitz  # PyMuPDF

# --- Modulos del plugin (soporta carga desde SIS Hub y ejecucion directa) ---
try:
    from .get_citi import ejecutar as get_citi
    from .get_capxiv import ejecutar as get_capxiv
    from .cliente_dte import (
        crear_sesion,
        descargar_por_soap_online,
        _safe_filename,
        DownloadResult,
    )
except ImportError:
    from get_citi import ejecutar as get_citi
    from get_capxiv import ejecutar as get_capxiv
    from cliente_dte import (
        crear_sesion,
        descargar_por_soap_online,
        _safe_filename,
        DownloadResult,
    )

# --------------------------------------------------------
# ---------------------- CONFIG --------------------------
# --------------------------------------------------------

# Parámetros operativos
MAX_EMAIL_MB = 14
OVERHEAD_EMAIL_BYTES = 100 * 1024  # holgura para headers/cuerpo
SUBDIR_ZIPS = "email_zips"
INCLUIR_ESTRUCTURA_RUT_EN_ZIP = True  # guardar como RUT/archivo en el ZIP
USAR_MAS_RECIENTE = False  # usar siempre la carpeta mapas_bch_* más reciente o False para la de la ejecución
ADJUNTAR_MAPA_EN_PRIMER_CORREO = True

# Correo: destinatarios, asunto, cuerpo
RECIPIENTS_TO = (
    "Alejandro Chaves Perez <achaves@bancochile.cl>; "
    "Daniel Angel Gomez Rodriguez <dagomez@bancochile.cl>"
)
RECIPIENTS_CC = (
    "Carlos Alberto Pino Camano <capinoc@banchile.cl>; "
    "Soporte Institucionales <SoporteInstitucionales@banchile.cl>; "
    "MiddleOffice.GSS <MiddleOffice.GSS2@banchile.cl>; "
    "FacturasBCH@banchile.cl"
)
ASUNTO_TEMPLATE = "Facturas {fecha_str} / Custodio Banco de Chile {i} de {N}"

CUERPO_HTML_TEMPLATE = """\
<p>Buenas tardes,</p>
<p>Estimados junto con saludar, adjunto facturas correspondientes a <strong>{fecha_str}</strong>.</p>
<p>Saludos,</p>
<div style="font-size: 10px; line-height: 1.2; margin-top: 8px;">
  <strong>Equipo Soporte Institucional</strong><br/>
  Gerencia de Operaciones RV / RF<br/>
  Banchile Inversiones<br/>
  Huérfanos 670, Piso 3<br/>
  Santiago - Chile<br/>
  <a href="https://www.banchileinversiones.cl" target="_blank">www.banchileinversiones.cl</a>
</div>
"""

# Logging básico
logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s",
)


# --------------------------------------------------------
# ----------------- UTILIDADES GENERALES ----------------
# --------------------------------------------------------

def crear_carpeta_unica(base_dir: str, prefijo: str = "mapas_bch") -> str:
    """
    Crea una carpeta única {prefijo}_YYYYMMDD_HHMMSS en base_dir.
    """
    base = Path(base_dir).expanduser().resolve()
    base.mkdir(parents=True, exist_ok=True)

    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre = f"{prefijo}_{sello}"
    ruta = base / nombre

    try:
        ruta.mkdir(exist_ok=False)
    except FileExistsError:
        sello_ms = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        ruta = base / f"{prefijo}_{sello_ms}"
        ruta.mkdir(exist_ok=False)

    return str(ruta)


def find_latest_mapas_dir(downloads_dir: Path) -> Optional[Path]:
    """
    Busca la carpeta mapas_bch_* más reciente en Descargas.
    """
    if not downloads_dir.exists():
        return None

    candidatos = sorted(
        [p for p in downloads_dir.iterdir()
         if p.is_dir() and p.name.startswith("mapas_bch_")],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidatos[0] if candidatos else None


def _rut_sin_dv(rut: str) -> str:
    """
    Parte numérica del RUT sin DV, sin puntos ni guion.
    """
    texto = (rut or "").strip().replace(".", "").replace(" ", "")
    if "-" in texto:
        texto = texto.split("-")[0]
    return "".join(ch for ch in texto if ch.isdigit())


def clave_pdf_desde_rut(rut: str) -> str:
    """
    Últimos 4 dígitos del RUT sin DV.
    """
    base = _rut_sin_dv(rut)
    return base[-4:] if base else ""


def desproteger_pdf(ruta_pdf: str, clave: str) -> None:
    """
    Abre PDF, autentica (si requiere) y guarda sin cifrado.
    """
    ruta = Path(ruta_pdf)
    if not ruta.is_file():
        raise FileNotFoundError(f"No existe el archivo: {ruta}")

    doc = fitz.open(ruta_pdf)
    try:
        if doc.needs_pass:
            if not clave or not doc.authenticate(clave):
                raise RuntimeError("Clave PDF incorrecta o no derivada.")
        ruta_tmp = str(ruta) + ".tmp"
        doc.save(
            ruta_tmp,
            encryption=fitz.PDF_ENCRYPT_NONE,
            garbage=4,
            deflate=True,
        )
    finally:
        doc.close()

    os.replace(ruta_tmp, ruta_pdf)


def _carpeta_rut(base_dir: str, rut: str) -> str:
    """
    Crea/retorna la carpeta por RUT (sanitizada).
    """
    rut_sanitizado = _safe_filename(rut) or "RUT_DESCONOCIDO"
    carpeta = Path(base_dir) / rut_sanitizado
    carpeta.mkdir(parents=True, exist_ok=True)
    return str(carpeta)


def _guardar_bytes_por_rut(
    base_dir: str,
    rut: str,
    folio: str,
    tipo_documento: str,
    formato: str,
    contenido: bytes,
) -> str:
    """
    Guarda bytes bajo carpeta del RUT con patrón DTE{tipo}_F{folio}.{ext}.
    Evita colisiones con sufijos incrementales.
    """
    carpeta = Path(_carpeta_rut(base_dir, rut))
    ext = "pdf" if formato.lower() == "pdf" else "xml"
    base_nombre = _safe_filename(f"DTE{tipo_documento}_F{folio}") or "archivo"
    ruta = carpeta / f"{base_nombre}.{ext}"

    if ruta.exists():
        i = 1
        while True:
            candidato = carpeta / f"{base_nombre}_{i}.{ext}"
            if not candidato.exists():
                ruta = candidato
                break
            i += 1

    with open(ruta, "wb") as f:
        f.write(contenido)
    return str(ruta)


def _descargar_y_guardar_uno(
    sesion,
    rut: str,
    folio: str,
    tipo_documento: str,
    formato: str,
    directorio_salida: str,
) -> str:
    """
    Descarga un DTE por SOAP 'online' y lo guarda por carpeta de RUT.
    Si es PDF, lo desprotege con clave derivada del RUT.
    """
    data = descargar_por_soap_online(
        sesion=sesion,
        folio=folio,
        tipo_documento=tipo_documento,
        formato=formato,
    )
    ruta = _guardar_bytes_por_rut(
        base_dir=directorio_salida,
        rut=rut,
        folio=folio,
        tipo_documento=tipo_documento,
        formato=formato,
        contenido=data,
    )
    if formato.lower() == "pdf":
        clave = clave_pdf_desde_rut(rut)
        desproteger_pdf(ruta_pdf=ruta, clave=clave)
    return ruta


def filtrar_capxiv_por_clientes(
    df_capxiv: pd.DataFrame,
    df_citi: pd.DataFrame,
) -> pd.DataFrame:
    """
    Filtra df_capxiv conservando solo RUT presentes en df_citi.
    Soporta nombres de columna variantes:
      - df_citi: 'Rut' o 'RUT'
      - df_capxiv: 'RUT_CLIENTE' o 'Rut Cliente'
    """
    col_rut_citi = "Rut" if "Rut" in df_citi.columns else "RUT"
    if col_rut_citi not in df_citi.columns:
        raise KeyError("df_citi requiere columna 'Rut' o 'RUT'.")

    if "RUT_CLIENTE" in df_capxiv.columns:
        col_rut_cap = "RUT_CLIENTE"
    elif "Rut Cliente" in df_capxiv.columns:
        col_rut_cap = "Rut Cliente"
    else:
        raise KeyError(
            "df_capxiv requiere 'RUT_CLIENTE' o 'Rut Cliente'."
        )

    ruts_validos = df_citi[col_rut_citi].dropna().astype(str).unique()
    return df_capxiv[df_capxiv[col_rut_cap].astype(str).isin(ruts_validos)]


def construir_dict_folios(df_filtrado: pd.DataFrame) -> Dict[str, List[str]]:
    """
    Construye diccionario {rut: [folios]}.
    Soporta 'RUT_CLIENTE' o 'Rut Cliente', 'FOL_FAC' o 'Folio'.
    """
    if "RUT_CLIENTE" in df_filtrado.columns:
        col_rut = "RUT_CLIENTE"
    elif "Rut Cliente" in df_filtrado.columns:
        col_rut = "Rut Cliente"
    else:
        raise KeyError(
            "df_filtrado requiere 'RUT_CLIENTE' o 'Rut Cliente'."
        )

    if "FOL_FAC" in df_filtrado.columns:
        col_folio = "FOL_FAC"
    elif "Folio" in df_filtrado.columns:
        col_folio = "Folio"
    else:
        raise KeyError("df_filtrado requiere 'FOL_FAC' o 'Folio'.")

    df = df_filtrado.loc[
        df_filtrado[col_rut].notna() & df_filtrado[col_folio].notna(),
        [col_rut, col_folio],
    ].copy()
    df[col_rut] = df[col_rut].astype(str).str.strip()
    df[col_folio] = df[col_folio].astype(str).str.strip()

    folios_por_rut: Dict[str, List[str]] = {}
    for rut, sub in df.groupby(col_rut):
        folios = [f for f in sub[col_folio].tolist() if f]
        if folios:
            folios_por_rut[rut] = folios
    return folios_por_rut


def descargar_facturas(
    formato: str,
    directorio: str,
    folios_por_rut: Dict[str, List[str]],
    tipo_documento: str,
) -> List[DownloadResult]:
    """
    Descarga todas las facturas (pdf|xml) por SOAP 'online',
    guardando por carpeta del RUT y quitando clave a PDFs.
    """
    formato = (formato or "").lower().strip()
    if formato not in {"pdf", "xml"}:
        raise ValueError("formato debe ser 'pdf' o 'xml'")

    sesion = crear_sesion()
    resultados: List[DownloadResult] = []

    for rut, folios in folios_por_rut.items():
        for folio in folios:
            folio_str = str(folio).strip()
            try:
                ruta = _descargar_y_guardar_uno(
                    sesion=sesion,
                    rut=rut,
                    folio=folio_str,
                    tipo_documento=tipo_documento,
                    formato=formato,
                    directorio_salida=directorio,
                )
                resultados.append(
                    DownloadResult(
                        folio=folio_str,
                        ok=True,
                        message=f"OK {formato.upper()}",
                        filepath=ruta,
                    )
                )
            except Exception as exc:  # noqa: BLE001 - logeamos detalle
                resultados.append(
                    DownloadResult(
                        folio=folio_str,
                        ok=False,
                        message=f"{formato.upper()}: {exc}",
                        filepath=None,
                    )
                )
    return resultados


def main_descarga_facturas(
    tipo_documento: str,
    df_filtrado: pd.DataFrame,
    directorio_salida: str,
) -> List[DownloadResult]:
    """
    Orquesta la construcción del diccionario RUT->[folios] y
    ejecuta la descarga en PDF y XML vía SOAP 'online', quitando
    contraseña a los PDF.
    """
    folios_por_rut = construir_dict_folios(df_filtrado)

    resultados: List[DownloadResult] = []

    # PDF
    resultados.extend(
        descargar_facturas(
            formato="pdf",
            directorio=directorio_salida,
            folios_por_rut=folios_por_rut,
            tipo_documento=tipo_documento,
        )
    )

    # XML
    resultados.extend(
        descargar_facturas(
            formato="xml",
            directorio=directorio_salida,
            folios_por_rut=folios_por_rut,
            tipo_documento=tipo_documento,
        )
    )

    return resultados


# --------------------------------------------------------
# ------------------- MAPA (EXCEL) ----------------------
# --------------------------------------------------------

def crear_mapa_excel_plano(
    df_capxiv: pd.DataFrame,
    df_citi: pd.DataFrame,
    ruta_carpeta: Path,
    fecha_str: str,
) -> Path:
    """
    Crea el 'Mapa' plano filtrando df_capxiv por RUT vigentes en df_citi.
    Usa columnas: df_citi['RUT'] y df_capxiv['Rut Cliente'] si existen,
    de lo contrario intenta variantes compatibles.
    """
    # Selección flexible de columnas
    col_rut_citi = "RUT" if "RUT" in df_citi.columns else "Rut"
    if col_rut_citi not in df_citi.columns:
        raise KeyError("df_citi requiere columna 'RUT' o 'Rut'.")

    col_rut_cap = "Rut Cliente" if "Rut Cliente" in df_capxiv.columns else "RUT_CLIENTE"
    if col_rut_cap not in df_capxiv.columns:
        raise KeyError("df_capxiv requiere 'Rut Cliente' o 'RUT_CLIENTE'.")

    ruts_unicos = df_citi[col_rut_citi].dropna().astype(str).drop_duplicates()
    df_filtrado = df_capxiv[df_capxiv[col_rut_cap].astype(str).isin(ruts_unicos)]

    ruta_mapa = ruta_carpeta / f"Mapa_Facturas_{fecha_str}.xlsx"
    # Plano, sin formato extra
    df_filtrado.to_excel(ruta_mapa, index=False)
    logging.info("Mapa creado: %s", ruta_mapa.name)
    return ruta_mapa


# --------------------------------------------------------
# -------- EMPAQUETADO Y PLANIFICACIÓN DE LOTES ----------
# --------------------------------------------------------

@dataclass
class FileItem:
    path: Path
    size: int
    arcname: str  # cómo se verá dentro del zip


@dataclass
class LotPlan:
    pdf_items: List[FileItem]
    xml_items: List[FileItem]
    incluye_mapa: bool


def base64_encoded_size(num_bytes: int) -> int:
    """Tamaño tras codificación Base64."""
    if num_bytes <= 0:
        return 0
    return 4 * math.ceil(num_bytes / 3)


def discover_files_for_zips(ruta_base: Path) -> Tuple[List[FileItem], List[FileItem]]:
    """
    Busca PDFs y XMLs recursivamente, omitiendo:
      - la carpeta SUBDIR_ZIPS
      - archivos .zip
      - el propio Mapa .xlsx
    """
    pdfs: List[FileItem] = []
    xmls: List[FileItem] = []

    excluir_dir = ruta_base / SUBDIR_ZIPS

    for path in ruta_base.rglob("*"):
        if path.is_dir():
            continue
        if excluir_dir in path.parents:
            continue
        if path.suffix.lower() in {".zip"}:
            continue
        if path.name.startswith("Mapa_Facturas_") and path.suffix.lower() == ".xlsx":
            # El mapa no va dentro de ZIPs
            continue
        if path.suffix.lower() not in {".pdf", ".xml"}:
            continue

        # Construir arcname (opcional: mantener RUT/archivo)
        if INCLUIR_ESTRUCTURA_RUT_EN_ZIP:
            # Intentar detectar carpeta RUT (hijo directo de ruta_base)
            try:
                rel = path.relative_to(ruta_base)
            except ValueError:
                rel = Path(path.name)
            # Mantener rel completo (con RUT/archivo) si corresponde
            arcname = str(rel)
        else:
            arcname = path.name

        item = FileItem(path=path, size=path.stat().st_size, arcname=arcname)
        if path.suffix.lower() == ".pdf":
            pdfs.append(item)
        else:
            xmls.append(item)

    # Ordenamos por tamaño descendente (heurística útil)
    pdfs.sort(key=lambda x: x.size, reverse=True)
    xmls.sort(key=lambda x: x.size, reverse=True)
    logging.info("Detectados PDFs: %d | XMLs: %d", len(pdfs), len(xmls))
    return pdfs, xmls


def build_zip_temp(items: List[FileItem]) -> Tuple[Path, int]:
    """
    Construye un ZIP temporal con 'items' y retorna (ruta, tamaño).
    """
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    tmp_path = Path(tmp.name)
    tmp.close()
    with tempfile.TemporaryDirectory() as _:
        import zipfile
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for it in items:
                zf.write(it.path, arcname=it.arcname)
    size = tmp_path.stat().st_size
    return tmp_path, size


def plan_lotes_por_tamano(
    pdfs: List[FileItem],
    xmls: List[FileItem],
    mapa_path: Optional[Path],
    max_email_mb: int = MAX_EMAIL_MB,
    overhead_bytes: int = OVERHEAD_EMAIL_BYTES,
) -> List[LotPlan]:
    """
    Planifica lotes asegurando que:
    base64(pdf_zip) + base64(xml_zip) + (base64(mapa si 1er lote)) + overhead <= max_email_mb.

    Estrategia:
      - Intercalado 'greedy' probando siempre el archivo más grande disponible
        entre PDF y XML.
      - Tras cada intento, se reconstruyen zips temporales para medir tamaño
        real y validar la cota (aceptar/revertir).
      - Cuando nada más cabe, se cierra el lote y se continúa.
    """
    max_email_bytes = max_email_mb * 1024 * 1024
    mapa_size = mapa_path.stat().st_size if (mapa_path and mapa_path.exists()) else 0
    mapa_b64 = base64_encoded_size(mapa_size)

    rem_pdfs = pdfs.copy()
    rem_xmls = xmls.copy()

    lotes: List[LotPlan] = []
    is_first = True

    while rem_pdfs or rem_xmls:
        inc_pdf: List[FileItem] = []
        inc_xml: List[FileItem] = []

        # Intentamos llenar el lote
        while True:
            # Elegir candidato mayor entre el siguiente PDF y XML
            next_pdf = rem_pdfs[0] if rem_pdfs else None
            next_xml = rem_xmls[0] if rem_xmls else None

            # No hay nada para intentar
            if not next_pdf and not next_xml:
                break

            # Orden de preferencia por tamaño
            intentos: List[Tuple[str, Optional[FileItem]]] = []
            if next_pdf and next_xml:
                if next_pdf.size >= next_xml.size:
                    intentos = [("pdf", next_pdf), ("xml", next_xml)]
                else:
                    intentos = [("xml", next_xml), ("pdf", next_pdf)]
            elif next_pdf:
                intentos = [("pdf", next_pdf)]
            else:
                intentos = [("xml", next_xml)]

            agregado = False
            for tipo, candidato in intentos:
                if candidato is None:
                    continue

                # Probar agregar
                if tipo == "pdf":
                    inc_pdf_try = inc_pdf + [candidato]
                    tmp_pdf_path, tmp_pdf_size = build_zip_temp(inc_pdf_try)
                    tmp_xml_path, tmp_xml_size = build_zip_temp(inc_xml)
                else:
                    inc_xml_try = inc_xml + [candidato]
                    tmp_pdf_path, tmp_pdf_size = build_zip_temp(inc_pdf)
                    tmp_xml_path, tmp_xml_size = build_zip_temp(inc_xml_try)

                try:
                    total_encoded = (
                        base64_encoded_size(tmp_pdf_size)
                        + base64_encoded_size(tmp_xml_size)
                        + (mapa_b64 if is_first and ADJUNTAR_MAPA_EN_PRIMER_CORREO else 0)
                        + overhead_bytes
                    )
                    if total_encoded <= max_email_bytes:
                        # Aceptamos el candidato
                        if tipo == "pdf":
                            inc_pdf.append(candidato)
                            rem_pdfs.pop(0)
                        else:
                            inc_xml.append(candidato)
                            rem_xmls.pop(0)
                        agregado = True
                        # Seguimos intentando agregar más
                        # (nos quedamos con los zips temporales generados para
                        # esta evaluación, pero los eliminamos al final del try)
                    else:
                        # No cabe; probamos el otro tipo (si queda)
                        agregado = False
                finally:
                    # Limpiar zips temporales
                    try:
                        tmp_pdf_path.unlink(missing_ok=True)
                    except Exception:  # noqa: BLE001
                        pass
                    try:
                        tmp_xml_path.unlink(missing_ok=True)
                    except Exception:  # noqa: BLE001
                        pass

                if agregado:
                    break  # salimos del for intentos, buscar siguiente candidato

            if not agregado:
                # Ningún candidato actual cabe; cerramos lote
                break

        # Registrar lote
        if not inc_pdf and not inc_xml:
            # Caso borde: si no se pudo meter nada (p.ej. un único archivo gigante)
            # Evitamos loop infinito: intentamos meter 1 archivo del tipo que haya.
            # Si ni así cabe, abortamos con log de advertencia.
            if rem_pdfs:
                solo = [rem_pdfs.pop(0)]
                logging.warning(
                    "Archivo PDF demasiado grande para combinar; se intentará solo."
                )
                lotes.append(LotPlan(pdf_items=solo, xml_items=[], incluye_mapa=is_first))
                is_first = False
                continue
            if rem_xmls:
                solo = [rem_xmls.pop(0)]
                logging.warning(
                    "Archivo XML demasiado grande para combinar; se intentará solo."
                )
                lotes.append(LotPlan(pdf_items=[], xml_items=solo, incluye_mapa=is_first))
                is_first = False
                continue
            # Nada que hacer
            break

        lotes.append(LotPlan(pdf_items=inc_pdf, xml_items=inc_xml, incluye_mapa=is_first))
        is_first = False

    logging.info("Total de lotes planificados: %d", len(lotes))
    return lotes


def crear_zips_de_lote(
    lote: LotPlan,
    carpeta_zips: Path,
    fecha_str: str,
    indice: int,
    total: int,
) -> Tuple[Optional[Path], Optional[Path]]:
    """
    Crea los zips definitivos del lote (omitiendo vacíos).
    Retorna rutas (zip_pdf, zip_xml), cada una puede ser None si ese tipo está vacío.
    """
    carpeta_zips.mkdir(parents=True, exist_ok=True)

    zip_pdf_path = None
    zip_xml_path = None

    import zipfile

    if lote.pdf_items:
        nombre_pdf = f"PDFs_Facturas_{fecha_str}_{indice}de{total}.zip"
        zip_pdf_path = carpeta_zips / nombre_pdf
        with zipfile.ZipFile(zip_pdf_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for it in lote.pdf_items:
                zf.write(it.path, arcname=it.arcname)

    if lote.xml_items:
        nombre_xml = f"XMLs_Facturas_{fecha_str}_{indice}de{total}.zip"
        zip_xml_path = carpeta_zips / nombre_xml
        with zipfile.ZipFile(zip_xml_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for it in lote.xml_items:
                zf.write(it.path, arcname=it.arcname)

    return zip_pdf_path, zip_xml_path


def levantar_correos_outlook(
    lotes: List[LotPlan],
    carpeta_zips: Path,
    fecha_str: str,
    mapa_path: Optional[Path],
) -> None:
    """
    Crea correos en Outlook (modo display), adjuntando 2 zips por correo
    (omitiendo los vacios) y el mapa solo en el 1er correo si corresponde.
    """
    try:
        import win32com.client as win32
        import pythoncom  # type: ignore
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "No se pudo importar pywin32/pythoncom para Outlook. Instala pywin32."
        ) from exc

    # pywin32 + Outlook via COM requiere inicializacion COM por hilo.
    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch("Outlook.Application")
        total = len(lotes)

        for i, lote in enumerate(lotes, start=1):
            asunto = ASUNTO_TEMPLATE.format(fecha_str=fecha_str, i=i, N=total)
            cuerpo_html = CUERPO_HTML_TEMPLATE.format(fecha_str=fecha_str)

            zip_pdf, zip_xml = crear_zips_de_lote(
                lote=lote,
                carpeta_zips=carpeta_zips,
                fecha_str=fecha_str,
                indice=i,
                total=total,
            )

            if zip_pdf is None and zip_xml is None and not (lote.incluye_mapa and mapa_path):
                logging.warning("Lote %d vacio; correo omitido.", i)
                continue

            mail = outlook.CreateItem(0)
            mail.To = RECIPIENTS_TO
            mail.CC = RECIPIENTS_CC
            mail.Subject = asunto
            mail.HTMLBody = cuerpo_html

            if zip_pdf:
                mail.Attachments.Add(str(zip_pdf))
            if zip_xml:
                mail.Attachments.Add(str(zip_xml))
            if lote.incluye_mapa and ADJUNTAR_MAPA_EN_PRIMER_CORREO and mapa_path:
                mail.Attachments.Add(str(mapa_path))

            mail.Display(False)
            logging.info("Correo %d/%d preparado: %s", i, total, asunto)
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def empaquetar_y_preparar_correos(
    ruta_trabajo: Path,
    fecha_str: str,
    levantar_correo: bool = False,
) -> None:
    """
    Descubre archivos, planifica lotes por tamaño, genera ZIPs por lote
    (definitivos al crear cada correo) y levanta correos si se solicita.
    """
    carpeta_zips = ruta_trabajo / SUBDIR_ZIPS
    carpeta_zips.mkdir(parents=True, exist_ok=True)

    # Descubrir PDFs y XMLs
    pdfs, xmls = discover_files_for_zips(ruta_trabajo)

    if not pdfs and not xmls:
        logging.warning("No se encontraron PDFs ni XMLs para empaquetar en %s.",
                        ruta_trabajo)
        return

    # Detectar Mapa (opcional)
    mapa_path = next(
        iter(ruta_trabajo.glob(f"Mapa_Facturas_{fecha_str}.xlsx")),
        None,
    )

    # Planificar lotes
    lotes = plan_lotes_por_tamano(
        pdfs=pdfs,
        xmls=xmls,
        mapa_path=mapa_path if ADJUNTAR_MAPA_EN_PRIMER_CORREO else None,
        max_email_mb=MAX_EMAIL_MB,
        overhead_bytes=OVERHEAD_EMAIL_BYTES,
    )

    if levantar_correo:
        levantar_correos_outlook(
            lotes=lotes,
            carpeta_zips=carpeta_zips,
            fecha_str=fecha_str,
            mapa_path=mapa_path if ADJUNTAR_MAPA_EN_PRIMER_CORREO else None,
        )
    else:
        # Si no levantamos correos, al menos generamos los ZIPs con nombres finales
        total = len(lotes)
        for i, lote in enumerate(lotes, start=1):
            crear_zips_de_lote(
                lote=lote,
                carpeta_zips=carpeta_zips,
                fecha_str=fecha_str,
                indice=i,
                total=total,
            )
        logging.info("ZIPs generados en: %s", carpeta_zips)


# --------------------------------------------------------
# -------------------- FLUJO PRINCIPAL -------------------
# --------------------------------------------------------

def ejecutar(fecha_str: str, levantar_correo: bool = False) -> None:
    """
    Flujo completo:
      1) Crea carpeta única en Descargas.
      2) Obtiene df_capxiv, df_citi.
      3) Filtra y descarga (PDF/XML), desprotege PDF.
      4) Crea Mapa (Excel) plano.
      5) Empaqueta y (opcional) levanta correos por lotes.
      6) Si USAR_MAS_RECIENTE=True, asegura usar la carpeta mapas_bch_* más nueva.
    """
    TIPO_DOCUMENTO_FACTURA = "33"

    usuario = getpass.getuser()
    ruta_descargas = Path(r"C:\Users") / usuario / "Downloads"
    ruta_descargas = ruta_descargas.resolve()

    # 1) Crear carpeta destino (la más reciente normalmente será esta)
    ruta_creada = Path(crear_carpeta_unica(str(ruta_descargas), prefijo="mapas_bch"))
    logging.info("Carpeta de trabajo creada: %s", ruta_creada)

    # 2) Obtener data
    logging.info("Obteniendo datos capxiv/citi...")
    df_capxiv = get_capxiv(fecha_str)
    df_citi = get_citi(fecha_str)

    # 3) Filtrar y descargar
    logging.info("Filtrando y descargando DTEs...")
    df_filtrado_descarga = filtrar_capxiv_por_clientes(df_capxiv, df_citi)
    _ = main_descarga_facturas(
        tipo_documento=TIPO_DOCUMENTO_FACTURA,
        df_filtrado=df_filtrado_descarga,
        directorio_salida=str(ruta_creada),
    )

    # 4) Crear Mapa (Excel) plano
    logging.info("Creando Mapa (Excel) plano...")
    try:
        crear_mapa_excel_plano(
            df_capxiv=df_capxiv,
            df_citi=df_citi,
            ruta_carpeta=ruta_creada,
            fecha_str=fecha_str,
        )
    except Exception as exc:  # noqa: BLE001
        logging.warning("No se pudo crear el Mapa (continuando): %s", exc)

    # 5) Seleccionar carpeta de trabajo (más reciente si se solicita)
    ruta_trabajo: Path
    if USAR_MAS_RECIENTE:
        ultima = find_latest_mapas_dir(ruta_descargas)
        if ultima:
            ruta_trabajo = ultima
            if ruta_trabajo != ruta_creada:
                logging.info("Usando carpeta más reciente: %s", ruta_trabajo)
        else:
            ruta_trabajo = ruta_creada
    else:
        ruta_trabajo = ruta_creada

    # 6) Empaquetar y levantar correos
    logging.info("Empaquetando y preparando correos...")
    empaquetar_y_preparar_correos(
        ruta_trabajo=ruta_trabajo,
        fecha_str=fecha_str,
        levantar_correo=levantar_correo,
    )

    logging.info("Proceso finalizado.")


def main():
    # --------- Parámetros de ejecución ----------
    LEVANTAR_CORREO = True  # True = levantar correos (Display), False = solo zips
    FECHA_STR = "04-02-2026"

    ejecutar(fecha_str=FECHA_STR, levantar_correo=LEVANTAR_CORREO)


if __name__ == "__main__":
    main()

# if __name__ == "__main__":
#     # --------- Parámetros de ejecución ----------
#     LEVANTAR_CORREO = True  # True = levantar correos (Display), False = solo zips
#     FECHA_STR = "03-02-2026"

#     ejecutar(fecha_str=FECHA_STR, levantar_correo=LEVANTAR_CORREO)
