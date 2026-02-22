# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 09:42:06 2025

@author: BBRUNA
"""

from __future__ import annotations

import base64
import html
import logging
import xml.etree.ElementTree as ET
from pathlib import Path

import requests


LOGGER = logging.getLogger(__name__)
DOCUMENT_TYPE_DTE = 33
ONLINE_URL = "http://asp403r.paperless.cl:80/axis2/services/Online.OnlineHttpSoap11Endpoint/"
MISDOC_URL = "http://fsw.bantcent.cl/cl.banchile/mis-documentos-srv/1.0.0/soap/misDocumentos"
SOAP_HEADERS = {"Content-Type": "text/xml; charset=utf-8"}

# Credenciales del servicio heredado del proceso de negocio.
ONLINE_ARGS0 = "96571220"
ONLINE_ARGS1 = "6782E92C-CFEB-4F74-8E7E-717480FD3A44"
ONLINE_ARGS2 = "hcyjPd7G"


def _online_envelope(folio: str, fmt: str) -> str:
    arg5 = "2" if fmt.lower() == "pdf" else "1"
    return f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
xmlns:web="http://webservices.online.webapp.paperless.cl">
  <soapenv:Header/>
  <soapenv:Body>
    <web:OnlineRecovery>
      <web:args0>{ONLINE_ARGS0}</web:args0>
      <web:args1>{ONLINE_ARGS1}</web:args1>
      <web:args2>{ONLINE_ARGS2}</web:args2>
      <web:args3>{DOCUMENT_TYPE_DTE}</web:args3>
      <web:args4>{folio}</web:args4>
      <web:args5>{arg5}</web:args5>
    </web:OnlineRecovery>
  </soapenv:Body>
</soapenv:Envelope>"""


def _misdoc_envelope(folio: str) -> str:
    return f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
xmlns:urn="urn:cl.banchile.eis:mis-documentos-srv:1.0">
  <soapenv:Header/>
  <soapenv:Body>
    <urn:docPdfInSrv>
      <requestBody>
        <idDocumento>{folio}</idDocumento>
        <tipoDocumento>{DOCUMENT_TYPE_DTE}</tipoDocumento>
        <origenDocumento>DTE</origenDocumento>
      </requestBody>
    </urn:docPdfInSrv>
  </soapenv:Body>
</soapenv:Envelope>"""


def _download_online(session: requests.Session, folio: str, fmt: str) -> bytes:
    response = session.post(
        ONLINE_URL,
        headers=SOAP_HEADERS,
        data=_online_envelope(folio, fmt),
        timeout=(5, 20),
    )
    response.raise_for_status()

    xml_root = ET.fromstring(response.content)
    namespaces = {"ns": "http://webservices.online.webapp.paperless.cl"}
    result_node = xml_root.find(".//ns:return", namespaces)
    if result_node is None or not result_node.text:
        raise RuntimeError("Respuesta SOAP sin nodo de retorno util.")

    inner_xml = ET.fromstring(html.unescape(result_node.text.strip()))
    message_url = inner_xml.findtext(".//Mensaje")
    if not message_url:
        raise RuntimeError("Respuesta SOAP sin URL de documento.")

    file_response = session.get(message_url.strip(), timeout=(5, 20))
    file_response.raise_for_status()
    return file_response.content


def _download_misdoc_pdf(session: requests.Session, folio: str) -> bytes:
    response = session.post(
        MISDOC_URL,
        headers=SOAP_HEADERS,
        data=_misdoc_envelope(folio),
        timeout=(5, 20),
    )
    response.raise_for_status()
    xml_root = ET.fromstring(response.content)
    code = xml_root.findtext(".//responseCodeService")
    if code != "000":
        raise RuntimeError(f"Servicio misDocumentos devolvio codigo {code}")
    payload_b64 = xml_root.findtext(".//documentoStr")
    if not payload_b64:
        raise RuntimeError("Respuesta de misDocumentos sin documentoStr")
    return base64.b64decode(payload_b64)


def download_documents(
    *,
    output_folder: Path,
    folios: list[str],
    include_pdf: bool,
    include_xml: bool,
    log_callback=None,
) -> list[Path]:
    generated: list[Path] = []
    if not folios or not (include_pdf or include_xml):
        return generated

    output_folder.mkdir(parents=True, exist_ok=True)
    session = requests.Session()

    for folio in folios:
        folio_text = str(folio).strip()
        if not folio_text:
            continue

        if include_pdf:
            try:
                content = _download_online(session, folio_text, "pdf")
            except Exception:
                try:
                    content = _download_misdoc_pdf(session, folio_text)
                except Exception as exc:
                    LOGGER.warning("No se pudo descargar PDF folio=%s: %s", folio_text, exc)
                    if log_callback:
                        log_callback(f"[WARN] PDF folio {folio_text}: {exc}")
                else:
                    path = output_folder / f"{folio_text}.pdf"
                    path.write_bytes(content)
                    generated.append(path)
                    if log_callback:
                        log_callback(f"[OK] PDF folio {folio_text}")
            else:
                path = output_folder / f"{folio_text}.pdf"
                path.write_bytes(content)
                generated.append(path)
                if log_callback:
                    log_callback(f"[OK] PDF folio {folio_text}")

        if include_xml:
            try:
                content = _download_online(session, folio_text, "xml")
            except Exception as exc:
                LOGGER.warning("No se pudo descargar XML folio=%s: %s", folio_text, exc)
                if log_callback:
                    log_callback(f"[WARN] XML folio {folio_text}: {exc}")
                continue
            path = output_folder / f"{folio_text}.xml"
            path.write_bytes(content)
            generated.append(path)
            if log_callback:
                log_callback(f"[OK] XML folio {folio_text}")

    session.close()
    return generated
