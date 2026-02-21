# -*- coding: utf-8 -*-
"""
Created on Fri Jun 27 12:25:52 2025

@author: CSANTOS
"""

def buscar_cliente_por_rut(rut, df):
    ''' Validar cliente en data. 
        Extrae nombre cliente.
        Verifica e indica los archivos que lleva.'''
    # Comparar intput con base clientes.
    resultado = df[df['RUT'].astype(str) == str(rut)]
    if resultado.empty:
        return None, []
    # Rescatar name.
    cliente = resultado.iloc[0]['CLIENTE']
    archivos = []
    # Rescatar archivos.
    for campo in ['FACTURA', 'XML', 'RESUMEN', 'TXT', 'EXCEL','CORREO MANUAL']:
        if str(resultado.iloc[0][campo]).strip().lower() == 'si':
            archivos.append(campo)
    return cliente, archivos
