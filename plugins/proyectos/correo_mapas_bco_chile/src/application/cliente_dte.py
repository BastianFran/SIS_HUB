# -*- coding: utf-8 -*-
"""
Cliente DTE — descarga por folio (PDF/XML).

Soporta:
- SOAP OnlineRecovery (PDF/XML)
- SOAP MisDocumentos (solo PDF)
- HTTP /consultadte/obtenerDTEs (PDF/XML)

Requisito de red:
- Los endpoints internos (HTTP/SOAP) requieren estar en la red corporativa correspondiente.
"""

from __future__ import annotations

import base64
import html
import os
import time
from dataclasses import dataclass
from typing import Optional, Tuple, Dict, Any, List

import requests
import xml.etree.ElementTree as ET


HTTP_BASE_URL = "http://appweb.banchile.cl/consultadte/obtenerDTEs"
SOAP_ONLINE_URL = (
    "http://asp403r.paperless.cl:80/"
    "axis2/services/Online.OnlineHttpSoap11Endpoint/"
)
SOAP_MIS_URL = (
    "http://fsw.bantcent.cl/cl.banchile/"
    "mis-documentos-srv/1.0.0/soap/misDocumentos"
)

# Encabezados para peticiones SOAP
HEADERS_SOAP = {"Content-Type": "text/xml; charset=utf-8"}

# Emisor por defecto visto en flujos internos (leer.txt)
DEFAULT_EMISOR = "96571220-8"

# Credenciales incrustadas en el SOAP OnlineRecovery (del script original)
ONLINE_ARGS0 = "96571220"
ONLINE_ARGS1 = "6782E92C-CFEB-4F74-8E7E-717480FD3A44"
ONLINE_ARGS2 = "hcyjPd7G"


@dataclass
class DownloadResult:
    folio: str
    ok: bool
    message: str
    filepath: Optional[str] = None


def crear_sesion() -> requests.Session:
    """Crea y devuelve una sesión HTTP reutilizable."""
    return requests.Session()


def _build_online_envelope(folio: str, tipo_documento: str, formato: str) -> str:
    """
    Arma el XML para OnlineRecovery.

    En la macro: args5 = 2 para PDF; args5 = 1 para XML
    """
    valor_args5 = "2" if formato.lower() == "pdf" else "1"

    return f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
                      xmlns:web="http://webservices.online.webapp.paperless.cl">
  <soapenv:Header/>
  <soapenv:Body>
    <web:OnlineRecovery>
      <web:args0>{ONLINE_ARGS0}</web:args0>
      <web:args1>{ONLINE_ARGS1}</web:args1>
      <web:args2>{ONLINE_ARGS2}</web:args2>
      <web:args3>{tipo_documento}</web:args3>
      <web:args4>{folio}</web:args4>
      <web:args5>{valor_args5}</web:args5>
    </web:OnlineRecovery>
  </soapenv:Body>
</soapenv:Envelope>"""


def _build_mis_envelope(folio: str, tipo_documento: str) -> str:
    """Arma el XML para el servicio misDocumentos (retorna base64 PDF)."""
    return f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
                      xmlns:urn="urn:cl.banchile.eis:mis-documentos-srv:1.0">
  <soapenv:Header/>
  <soapenv:Body>
    <urn:docPdfInSrv>
      <requestBody>
        <idDocumento>{folio}</idDocumento>
        <tipoDocumento>{tipo_documento}</tipoDocumento>
        <origenDocumento>DTE</origenDocumento>
      </requestBody>
    </urn:docPdfInSrv>
  </soapenv:Body>
</soapenv:Envelope>"""


def descargar_por_soap_online(
    sesion: requests.Session, folio: str, tipo_documento: str, formato: str
) -> bytes:
    """
    Llama a OnlineRecovery y devuelve los bytes del archivo (PDF/XML).
    """
    envelope = _build_online_envelope(folio, tipo_documento, formato)
    resp = sesion.post(SOAP_ONLINE_URL, headers=HEADERS_SOAP, data=envelope, timeout=60)
    resp.raise_for_status()

    root = ET.fromstring(resp.content)
    ns = {"ns": "http://webservices.online.webapp.paperless.cl"}
    nodo = root.find(".//ns:return", ns)
    if nodo is None or not nodo.text:
        raise RuntimeError("No se encontró <ns:return> en respuesta SOAP (OnlineRecovery)")

    xml_interno = html.unescape(nodo.text.strip())
    interno = ET.fromstring(xml_interno)
    url_archivo = interno.findtext(".//Mensaje")
    if not url_archivo:
        raise RuntimeError("No se encontró <Mensaje> en el XML interno (OnlineRecovery)")

    file_resp = sesion.get(url_archivo.strip(), timeout=60)
    file_resp.raise_for_status()
    return file_resp.content


def descargar_por_soap_mis(
    sesion: requests.Session, folio: str, tipo_documento: str
) -> bytes:
    """
    Llama a misDocumentos y devuelve los bytes decodificados del PDF.
    """
    envelope = _build_mis_envelope(folio, tipo_documento)
    resp = sesion.post(SOAP_MIS_URL, headers=HEADERS_SOAP, data=envelope, timeout=60)
    resp.raise_for_status()

    root = ET.fromstring(resp.content)
    code = root.findtext(".//responseCodeService")
    if code != "000":
        raise RuntimeError(f"misDocumentos devolvió código {code}")

    b64 = root.findtext(".//documentoStr")
    if not b64:
        raise RuntimeError("No se encontró <documentoStr> en respuesta SOAP (misDocumentos)")

    return base64.b64decode(b64)


def descargar_por_http(
    sesion: requests.Session,
    folio: str,
    tipo_documento: str,
    formato: str,
    emisor: str = DEFAULT_EMISOR,
    cliente: Optional[str] = None,
    fecha_emision: Optional[str] = None,
) -> bytes:
    """
    Descarga vía endpoint HTTP /consultadte/obtenerDTEs.
    Params (según leer.txt):
      folio, tipo, emisor, cliente, fechaEmision, formato, fileName, avoidCaching
    """
    params: Dict[str, Any] = {
        "folio": folio,
        "tipo": str(tipo_documento),
        "emisor": emisor,
        "formato": formato.lower(),
        "fileName": folio,
        "avoidCaching": int(time.time() * 1000),
    }
    if cliente:
        params["cliente"] = cliente
    if fecha_emision:
        params["fechaEmision"] = fecha_emision

    resp = sesion.get(HTTP_BASE_URL, params=params, timeout=60)
    resp.raise_for_status()
    return resp.content


def _safe_filename(base_name: str) -> str:
    # simple sanitization
    return "".join(ch for ch in base_name if ch.isalnum() or ch in ("-", "_", ".", " ")).strip() or "archivo"


def guardar_archivo_plano(destino: str, folio: str, tipo_documento: str, contenido: bytes, formato: str) -> str:
    """
    Guarda el archivo en una carpeta plana (sin subcarpetas por cliente).
    No cifra ni pone clave.
    """
    os.makedirs(destino, exist_ok=True)
    ext = "pdf" if formato.lower() == "pdf" else "xml"
    base = _safe_filename(f"DTE{tipo_documento}_F{folio}")
    ruta = os.path.join(destino, f"{base}.{ext}")

    # evitar sobreescritura
    if os.path.exists(ruta):
        i = 1
        while True:
            candidato = os.path.join(destino, f"{base}_{i}.{ext}")
            if not os.path.exists(candidato):
                ruta = candidato
                break
            i += 1

    with open(ruta, "wb") as f:
        f.write(contenido)
    return ruta


def descargar_lote(
    folios: List[str],
    destino: str,
    formato: str,
    tipo_documento: str,
    metodo: str,
    soap_backend: str = "online",
    cliente: Optional[str] = None,
    fecha_emision: Optional[str] = None,
    emisor: str = DEFAULT_EMISOR,
    sesion: Optional[requests.Session] = None,
) -> List[DownloadResult]:
    """
    Descarga una lista de folios. Un fallo no interrumpe el resto.
    Retorna lista de resultados con ok/err y mensajes.
    """
    metodo = (metodo or "").lower().strip()
    soap_backend = (soap_backend or "online").lower().strip()
    formato = (formato or "").lower().strip()

    if formato not in {"pdf", "xml"}:
        raise ValueError("formato debe ser 'pdf' o 'xml'")

    if metodo not in {"soap", "http"}:
        raise ValueError("metodo debe ser 'soap' o 'http'")

    if sesion is None:
        sesion = crear_sesion()

    results: List[DownloadResult] = []
    for folio in folios:
        try:
            if metodo == "http":
                data = descargar_por_http(
                    sesion,
                    folio=folio,
                    tipo_documento=tipo_documento,
                    formato=formato,
                    emisor=emisor,
                    cliente=cliente,
                    fecha_emision=fecha_emision,
                )
            else:
                # SOAP
                if soap_backend == "mis":
                    if formato != "pdf":
                        raise RuntimeError("MisDocumentos solo permite PDF (no XML).")
                    data = descargar_por_soap_mis(sesion, folio=folio, tipo_documento=tipo_documento)
                else:
                    data = descargar_por_soap_online(
                        sesion, folio=folio, tipo_documento=tipo_documento, formato=formato
                    )

            ruta = guardar_archivo_plano(destino, folio, tipo_documento, data, formato)
            results.append(DownloadResult(folio=folio, ok=True, message="OK", filepath=ruta))
        except Exception as e:
            results.append(DownloadResult(folio=folio, ok=False, message=str(e), filepath=None))
    return results
