# -*- coding: utf-8 -*-
"""
Automatiza el inicio de sesión y la descarga de reportes desde CitiVelocity.

Creado el 9 de septiembre de 2025
Autor: BBRUNA
"""

import os
import time
import glob
import keyring
import shutil
import datetime
import pandas as pd
from typing import Optional, Tuple
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from datetime import date


def fecha_formato_mdyyyy_flexible(fecha_str: str) -> str:
    """Convierte fechas 'dd-mm-aaaa', 'dd/mm/aaaa' o 'dd.mm.aaaa' a 'm/d/aaaa'."""
    formatos = ["%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y"]
    for fmt in formatos:
        try:
            fecha = datetime.datetime.strptime(fecha_str.strip(), fmt).date()
            return f"{fecha.month}/{fecha.day}/{fecha.year}"
        except ValueError:
            continue
    raise ValueError("Formato no soportado. Use dd-mm-aaaa, dd/mm/aaaa o dd.mm.aaaa")


def obtener_credenciales(nombre: str) -> Tuple[Optional[str], Optional[str]]:
    """Obtiene usuario y contraseña desde el administrador de credenciales."""
    try:
        credencial = keyring.get_credential(nombre, None)
        if credencial:
            print("Credenciales obtenidas correctamente desde keyring.")
            return credencial.username, credencial.password
        print("No se encontraron credenciales en keyring.")
        return None, None
    except Exception as error:
        print(f"Error al obtener las credenciales: {error}")
        return None, None


def limpiar_carpeta_descargas(carpeta: str) -> None:
    """Elimina todos los archivos de la carpeta de descargas antes de iniciar una nueva descarga."""
    if not os.path.exists(carpeta):
        os.makedirs(carpeta)
        print(f"Carpeta de descargas creada: {carpeta}")
        return

    for archivo in os.listdir(carpeta):
        ruta = os.path.join(carpeta, archivo)
        try:
            if os.path.isfile(ruta):
                os.remove(ruta)
                print(f"Archivo eliminado: {ruta}")
        except PermissionError:
            print(f"No se pudo eliminar (posiblemente abierto): {ruta}")


def configurar_navegador(ruta_driver: str, carpeta_descargas: str):
    """Configura el navegador Chrome para descargas automáticas."""
    opciones = webdriver.ChromeOptions()
    preferencias = {"download.default_directory": carpeta_descargas}
    opciones.add_experimental_option("prefs", preferencias)
    servicio = Service(ruta_driver)
    navegador = webdriver.Chrome(service=servicio, options=opciones)
    espera = WebDriverWait(navegador, 30)
    print("Navegador Chrome configurado y lanzado.")
    return navegador, espera


def realizar_login_si_es_necesario(navegador, espera, usuario: str, contraseña: str) -> None:
    """Detecta si hay formulario de login en la página del reporte y lo completa si es necesario."""
    try:
        campo_usuario = espera.until(EC.presence_of_element_located((By.ID, "userPwd")))
        campo_usuario.send_keys(usuario)
        esperar_y_clickear(espera, By.ID, "btncontinue")
        esperar_y_clickear(espera, By.ID, "btnAltrnateLgn")
        esperar_y_clickear(espera, By.ID, "icgUpassword")
        esperar_y_escribir(espera, By.ID, "icgUpassword", contraseña)
        esperar_y_clickear(espera, By.ID, "btnSignin")
        print("Login en CitiVelocity realizado correctamente.")
    except Exception:
        print("No se detectó formulario de login. Se asume sesión activa.")


def esperar_y_clickear(espera, tipo: By, identificador: str) -> None:
    elemento = espera.until(EC.element_to_be_clickable((tipo, identificador)))
    elemento.click()


def esperar_y_escribir(espera, tipo: By, identificador: str, texto: str) -> None:
    elemento = espera.until(EC.element_to_be_clickable((tipo, identificador)))
    elemento.send_keys(texto)


def clickear_boton_descarga(fecha, navegador, espera):
    # Entrar al iframe Main
    espera.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "Main")))

    # ============================
    # BLOQUE TRADE DATE (usa fecha)
    # ============================
    bloque_trade = espera.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//div[@id='calendarinput'][.//label[normalize-space()='Trade Date:']]"
        ))
    )
    
    
    select = Select(navegador.find_element("id", "ddlTradeStatus"))
    select.select_by_visible_text("Prematched")


    trade_from = bloque_trade.find_element(By.XPATH, ".//input[@placeholder='From']")
    trade_to   = bloque_trade.find_element(By.XPATH, ".//input[@placeholder='To']")
    
    
    trade_from.click()
    
    trade_from.click()
    trade_from.send_keys(Keys.CONTROL, "a")
    trade_from.send_keys(Keys.DELETE)
    trade_from.send_keys("")
    # navegador.execute_script("arguments[0].value = arguments[1];", trade_from, '')
    
    trade_to.click()
    trade_to.click()
    trade_to.send_keys(Keys.CONTROL, "a")
    trade_to.send_keys(Keys.DELETE)
    trade_to.send_keys("")
    navegador.execute_script("arguments[0].value = arguments[1];", trade_to, '')

    # ============================
    # BLOQUE CSD (usa '')
    # ============================
    bloque_csd = espera.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//div[@id='calendarinput'][.//label[normalize-space()='CSD:']]"
        ))
    )

    csd_from = bloque_csd.find_element(By.XPATH, ".//input[@placeholder='From']")
    csd_to   = bloque_csd.find_element(By.XPATH, ".//input[@placeholder='To']")
    
    
    csd_from.click()
    csd_from.send_keys(Keys.CONTROL, "a")
    csd_from.send_keys(Keys.DELETE)
    csd_from.send_keys(fecha)
    # navegador.execute_script("arguments[0].value = arguments[1];", csd_from, "1/8/2026")
   
    
    csd_to.click()
    csd_to.click()
    csd_to.send_keys(Keys.CONTROL, "a")
    csd_to.send_keys(Keys.DELETE)
    csd_to.send_keys(fecha)
    
    # navegador.execute_script("arguments[0].value = arguments[1];", csd_to, "1/8/2026")
    
    
    boton = espera.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="filter_accordion_parent"]/div/div/div[2]/div/div[2]/a[2]')
        )
    )
    navegador.execute_script("arguments[0].click();", boton)
    

    # ============================
    # BOTÓN DOWNLOAD REPORT
    # ============================
    boton = espera.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[normalize-space()='Download Report']")
        )
    )
    navegador.execute_script("arguments[0].click();", boton)

    navegador.switch_to.default_content()



def esperar_descarga(carpeta: str, tiempo_maximo: int = 60) -> str:
    """Espera hasta que el archivo Excel se descargue completamente."""
    segundos = 0
    while segundos < tiempo_maximo:
        archivos = glob.glob(os.path.join(carpeta, "*.xls*"))
        if archivos:
            archivo_reciente = max(archivos, key=os.path.getctime)
            if not archivo_reciente.endswith(".crdownload"):
                print(f"Archivo descargado: {archivo_reciente}")
                return archivo_reciente
        time.sleep(1)
        segundos += 1
    print("Timeout: no se detectó archivo descargado en el tiempo esperado.")
    raise TimeoutError("No se detectó archivo descargado en el tiempo esperado.")


def mover_y_renombrar_archivo(ruta_archivo: str, carpeta_destino: str) -> str:
    """Mueve el archivo descargado a la carpeta destino con un nombre único."""
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    extension = os.path.splitext(ruta_archivo)[1]
    nuevo_nombre = f"reporte_citi_{timestamp}{extension}"
    nueva_ruta = os.path.join(carpeta_destino, nuevo_nombre)

    shutil.move(ruta_archivo, nueva_ruta)
    print(f"Archivo movido y renombrado: {nueva_ruta}")
    return nueva_ruta


def procesar_descarga(carpeta: str, carpeta_destino: str) -> pd.DataFrame:
    """Espera la descarga, carga el archivo a un DataFrame y lo mueve/renombra."""
    ruta_archivo = esperar_descarga(carpeta)
    df = pd.read_excel(ruta_archivo, engine="openpyxl")
    
    COLUMNAS_FINAL = [
        "Fecha Ejecución",
        "Nemotecnico",
        "Operación",
        "Rut",
        "Cantidad",
        "Monto",
        "Fecha Operacion",
        "Fecha Liquidacion",
        "Cliente",
        "Observaciones",
    ]
    
    mapeo = {
        "LOCAL": "Nemotecnico",
        "TYPE": "Operación",
        "RUT": "Rut",
        "QUANTITY": "Cantidad",
        "AMOUNT": "Monto",
        "TRADE DATE": "Fecha Operacion",
        "CSD": "Fecha Liquidacion",
        "ACCOUNT NAME": "Cliente",
    }
    df = df[list(mapeo)].rename(columns=mapeo)

    hoy_str = datetime.date.today().strftime("%d-%m-%Y")
    df["Fecha Ejecución"] = hoy_str
    df["Observaciones"] = ""

    df["Cantidad"] = df["Cantidad"].astype(str).str.replace(",", "", regex=False).astype(float)
    df["Monto"] = df["Monto"].astype(str).str.replace(",", "", regex=False).astype(float)

    df["Operación"] = df["Operación"].str.upper()

    # df = formatear_fechas(df, "Fecha Operacion", "Fecha Liquidacion", dayfirst=False)

    ruta_final = mover_y_renombrar_archivo(ruta_archivo, carpeta_destino)
    print(f"Archivo procesado y guardado en: {ruta_final}")

    return df[COLUMNAS_FINAL]


def ejecutar(fecha_str) -> pd.DataFrame:
    """Función principal que orquesta el proceso completo y retorna el DataFrame."""
    
    fecha = fecha_formato_mdyyyy_flexible(fecha_str)
    USUARIO = os.getlogin()
    NOMBRE_CREDENCIAL_CITI = "citivelocity"
    URL_REPORTE_CITI = "https://www.citivelocity.com/cv2/go/SMI_NEXUS_BROKER_INTERFACE"
    RUTA_DRIVER = rf"C:\Users\{USUARIO}\OneDrive - Banchile\Soporte Institucionales\Utilidades\webdrivers\chrome\chromedriver_142.0.7444.175.exe"
    CARPETA_BASE = rf"C:\Users\{USUARIO}\OneDrive - Banchile\Soporte Institucionales\Desarrollos\en_desarrollo\mapas_facturas_bch"
    CARPETA_DESCARGAS_CITI = os.path.join(CARPETA_BASE, "download", "citi_velocity")
    # === Columnas finales comunes ===

    print("Iniciando proceso de descarga CitiVelocity...")
    usuario, contraseña = obtener_credenciales(NOMBRE_CREDENCIAL_CITI)
    
    if not usuario or not contraseña:
        print("Credenciales no encontradas en keyring.")
        raise ValueError("Credenciales no encontradas. Verifica el administrador de credenciales.")

    limpiar_carpeta_descargas(CARPETA_DESCARGAS_CITI)
    navegador, espera = configurar_navegador(RUTA_DRIVER, CARPETA_DESCARGAS_CITI)

    try:
        navegador.get(URL_REPORTE_CITI)
        realizar_login_si_es_necesario(navegador, espera, usuario, contraseña)
        clickear_boton_descarga(fecha, navegador, espera)
        datos = procesar_descarga(CARPETA_DESCARGAS_CITI, CARPETA_DESCARGAS_CITI)
        print("Proceso CitiVelocity finalizado con éxito.")
        return datos
    finally:
        navegador.quit()
        print("Navegador cerrado correctamente.")

# if __name__ == "__main__":
#     fecha_str = "20-01-2026"
#     df_citi = ejecutar(fecha_str)
