# -*- coding: utf-8 -*-
"""
Created on Wed Jun 18 14:05:47 2025

@author: CSANTOS
"""

from application.cliente_dte import descargar_facturas

###############################################################################    
    
directorio_salida = r"C:\Users\csantos\Downloads"
    
    # Diccionario con datos simples
folios_por_rut = {'77750920': ['10618562','10618563']}
    
    # Tipo de documento DTE (33) 33 es FACTURA ELECTRONICA que puede ser PDF o xML
TIPO_DOCUMENTO_DTE = "33"
    
    # Cambiar a "xml" o "pdf" segun deseado
FORMATO = 'pdf'
    
###############################################################################    
descargar_facturas(
        FORMATO,
        directorio_salida,
        folios_por_rut 
    )

# FORMATO = 'pdf'
# directorio_salida = r"C:\Users\csantos\Downloads"
# SERVICIO_PREDETERMINADO = "online"

# descargar_facturas(
#     FORMATO,
#     directorio_salida,
#     folios_por_rut,
#     servicio=SERVICIO_PREDETERMINADO,
    
# )