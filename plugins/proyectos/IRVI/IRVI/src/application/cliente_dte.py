# -*- coding: utf-8 -*-
"""
Created on Fri Jun 13 11:08:06 2025

@author: BBRUNA

Contiene las funciones para descargar y
guardar facturas vía los servicios SOAP.
"""
import os
import html
import base64
import time
import logging
import requests
import xml.etree.ElementTree as ET
from typing import Dict, List

TIPO_DOCUMENTO_DTE = 33

logger = logging.getLogger(__name__)


def crear_sesion() -> requests.Session:
    """
    Crea y devuelve una sesión HTTP reutilizable.
    """
    logger.debug("Creando nueva sesión HTTP")
    session = requests.Session()
    return session


def construir_envelope_online(folio: str, formato: str) -> str:
    """
    Arma el XML para OnlineRecovery.
    Args:
      folio  -- número de factura
      formato -- "pdf" (default) o "xml"
    """
    logger.debug("Construyendo envelope online para folio=%s, formato=%s",
                 folio, formato)
    ARGS0 = "96571220"
    ARGS1 = "6782E92C-CFEB-4F74-8E7E-717480FD3A44"
    ARGS2 = "hcyjPd7G"
    valor_args5 = "2" if formato.lower() == "pdf" else "1"

    envelope = f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
                      xmlns:web="http://webservices.online.webapp.paperless.cl">
  <soapenv:Header/>
  <soapenv:Body>
    <web:OnlineRecovery>
      <web:args0>{ARGS0}</web:args0>
      <web:args1>{ARGS1}</web:args1>
      <web:args2>{ARGS2}</web:args2>
      <web:args3>{TIPO_DOCUMENTO_DTE}</web:args3>
      <web:args4>{folio}</web:args4>
      <web:args5>{valor_args5}</web:args5>
    </web:OnlineRecovery>
  </soapenv:Body>
</soapenv:Envelope>"""
    return envelope


def construir_envelope_mis(folio: str) -> str:
    """
    Arma el XML para el servicio misDocumentos.
    """
    logger.debug("Construyendo envelope misDocumentos para folio=%s", folio)
    envelope = f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
                      xmlns:urn="urn:cl.banchile.eis:mis-documentos-srv:1.0">
  <soapenv:Header/>
  <soapenv:Body>
    <urn:docPdfInSrv>
      <requestBody>
        <idDocumento>{folio}</idDocumento>
        <tipoDocumento>{TIPO_DOCUMENTO_DTE}</tipoDocumento>
        <origenDocumento>DTE</origenDocumento>
      </requestBody>
    </urn:docPdfInSrv>
  </soapenv:Body>
</soapenv:Envelope>"""
    return envelope


def descargar_factura_online(
    sesion: requests.Session, folio: str, formato: str
) -> bytes:
    """
    Llama a OnlineRecovery y devuelve los bytes del PDF o XML.
    """
    URL_ONLINE = (
        "http://asp403r.paperless.cl:80/"
        "axis2/services/Online.OnlineHttpSoap11Endpoint/"
    )
    HEADERS_SOAP = {"Content-Type": "text/xml; charset=utf-8"}
    logger.info("Descargando factura ONLINE folio=%s formato=%s", folio, formato)

    start = time.perf_counter()
    envelope = construir_envelope_online(folio, formato)
    try:
        resp = sesion.post(URL_ONLINE, headers=HEADERS_SOAP, data=envelope, timeout=(5, 15))
        resp.raise_for_status()
        logger.debug("POST OnlineRecovery status=%s tiempo=%.2fs",
                     resp.status_code, time.perf_counter() - start)

        root = ET.fromstring(resp.content)
        ns = {"ns": "http://webservices.online.webapp.paperless.cl"}
        nodo = root.find(".//ns:return", ns)
        if nodo is None or not nodo.text:
            raise RuntimeError("No se encontró <ns:return> en respuesta SOAP")

        xml_interno = html.unescape(nodo.text.strip())
        interno = ET.fromstring(xml_interno)
        url_pdf = interno.findtext(".//Mensaje")
        if not url_pdf:
            raise RuntimeError("No se encontró <Mensaje> en el XML interno")

        # Obtener PDF
        start_get = time.perf_counter()
        pdf_resp = sesion.get(url_pdf.strip(), timeout=(5, 15))
        pdf_resp.raise_for_status()
        logger.debug("GET PDF status=%s tiempo=%.2fs",
                     pdf_resp.status_code, time.perf_counter() - start_get)

        total_time = time.perf_counter() - start
        logger.info("Factura ONLINE folio=%s recuperada en %.2fs", folio, total_time)
        return pdf_resp.content

    except Exception as e:
        logger.error("Error al descargar factura ONLINE folio=%s: %s", folio, e,
                     exc_info=True)
        raise


def descargar_factura_mis(
    sesion: requests.Session, folio: str
) -> bytes:
    """
    Llama a misDocumentos y devuelve los bytes decodificados del PDF.
    """
    URL_MIS = (
        "http://fsw.bantcent.cl/cl.banchile/"
        "mis-documentos-srv/1.0.0/soap/misDocumentos"
    )
    HEADERS_SOAP = {"Content-Type": "text/xml; charset=utf-8"}
    logger.info("Descargando factura MIS folio=%s", folio)

    start = time.perf_counter()
    try:
        envelope = construir_envelope_mis(folio)
        resp = sesion.post(URL_MIS, headers=HEADERS_SOAP, data=envelope, timeout=(5, 15))
        resp.raise_for_status()
        logger.debug("POST misDocumentos status=%s tiempo=%.2fs",
                     resp.status_code, time.perf_counter() - start)

        root = ET.fromstring(resp.content)
        code = root.findtext(".//responseCodeService")
        if code != "000":
            raise RuntimeError(f"misDocumentos devolvió código {code}")

        b64 = root.findtext(".//documentoStr")
        if not b64:
            raise RuntimeError("No se encontró <documentoStr> en respuesta SOAP")

        contenido = base64.b64decode(b64)
        total_time = time.perf_counter() - start
        logger.info("Factura MIS folio=%s decodificada en %.2fs", folio, total_time)
        return contenido

    except Exception as e:
        logger.error("Error al descargar factura MIS folio=%s: %s", folio, e,
                     exc_info=True)
        raise


def guardar_archivo(
    carpeta: str,
    folio: str,
    contenido: bytes,
    formato: str = "pdf"
) -> None:
    """
    Guarda el contenido en un archivo y registra éxito o error.
    """
    os.makedirs(carpeta, exist_ok=True)
    ext = "pdf" if formato.lower() == "pdf" else "xml"
    ruta = os.path.join(carpeta, f"{folio}.{ext}")
    try:
        with open(ruta, "wb") as f:
            f.write(contenido)
        logger.info("Archivo guardado: %s", ruta)
    except Exception as e:
        logger.error("No se pudo guardar %s: %s", ruta, e, exc_info=True)
        raise


def descargar_facturas(
    formato: str,
    directorio: str,
    folios_por_rut: Dict[str, List[str]],
) -> None:
    """
    Orquesta la descarga de facturas para cada RUT y folio.
    """
    logger.info("Inicia descarga de facturas formato=%s", formato)
    servicio = "online"
    sesion = crear_sesion()

    for rut, lista_folios in folios_por_rut.items():
        carpeta_rut = os.path.join(directorio, rut)
        for folio in lista_folios:
            try:
                if servicio == "online":
                    datos = descargar_factura_online(sesion, folio, formato)
                else:
                    datos = descargar_factura_mis(sesion, folio)
                guardar_archivo(carpeta_rut, folio, datos, formato)
            except Exception:
                # El error ya fue logueado en la función específica
                continue

    logger.info("Finaliza descarga de facturas")
