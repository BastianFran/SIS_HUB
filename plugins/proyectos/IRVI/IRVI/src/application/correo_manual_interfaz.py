# -*- coding: utf-8 -*-
"""
Created on Thu Jul 17 13:51:16 2025

@author: CSANTOS
"""
import os
import win32com.client

def correo_manual(destinatarios,carpeta,asunto):
    ''' '''
    # Crear instancia de Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # Adjuntar todos los archivos de la carpeta
    for archivo in os.listdir(carpeta):
        ruta_completa = os.path.join(carpeta, archivo)
        if os.path.isfile(ruta_completa):  # Asegura que sea un archivo
            mail.Attachments.Add(ruta_completa)
    # Opcional: configurar destinatario, asunto, etc.
    mail.To = destinatarios
    mail.CC = 'SoporteInstitucionales@banchile.cl; TradingRVInstitucional@banchile.cl>; InstitutionalSales@banchile.cl'
    mail.Subject = asunto
    mail.Body = "Buenas tardes estimados, adjunto archivos del día."
    mail.Display()









