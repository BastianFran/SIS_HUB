# -*- coding: utf-8 -*-
"""
Bootstrap de SIS Hub con imports guardian para que PyInstaller pueda incluir
dependencias usadas por distintos plugins.
"""

from __future__ import annotations

from sis_hub_core.app import run_app

# =========================
# IMPORTS NUCLEO (SIS Hub)
# =========================
import tkinter as tk  # noqa: F401
from tkinter import filedialog, messagebox, ttk  # noqa: F401
from tkinter.scrolledtext import ScrolledText  # noqa: F401

# =========================
# IMPORTS PARA PLUGINS
# =========================
OPTIONAL_IMPORT_ERRORS: dict[str, str] = {}


def _record_optional_import_error(module_name: str, exc: Exception) -> None:
    OPTIONAL_IMPORT_ERRORS[module_name] = f"{type(exc).__name__}: {exc}"


try:
    import pandas as pd  # noqa: F401
except Exception as exc:
    _record_optional_import_error("pandas", exc)

try:
    import numpy as np  # noqa: F401
except Exception as exc:
    _record_optional_import_error("numpy", exc)

try:
    import pulp  # noqa: F401
except Exception as exc:
    _record_optional_import_error("pulp", exc)

try:
    import openpyxl  # noqa: F401
    from openpyxl import Workbook  # noqa: F401
    from openpyxl.utils.dataframe import dataframe_to_rows  # noqa: F401
    from openpyxl.worksheet.table import Table  # noqa: F401
except Exception as exc:
    _record_optional_import_error("openpyxl", exc)

try:
    import xlsxwriter  # noqa: F401
except Exception as exc:
    _record_optional_import_error("xlsxwriter", exc)

try:
    from PIL import Image, ImageTk  # noqa: F401
except Exception as exc:
    _record_optional_import_error("PIL", exc)

try:
    import fitz  # noqa: F401  # PyMuPDF
except Exception as exc:
    _record_optional_import_error("fitz", exc)

try:
    import win32com.client as win32  # noqa: F401
    from win32com.client import Dispatch  # noqa: F401
except Exception as exc:
    _record_optional_import_error("win32com", exc)

try:
    import pyodbc  # noqa: F401
except Exception as exc:
    _record_optional_import_error("pyodbc", exc)

try:
    import requests  # noqa: F401
except Exception as exc:
    _record_optional_import_error("requests", exc)

try:
    from lxml import etree  # noqa: F401
except Exception as exc:
    _record_optional_import_error("lxml", exc)

try:
    import holidays  # noqa: F401
except Exception as exc:
    _record_optional_import_error("holidays", exc)

try:
    import keyring  # noqa: F401
except Exception as exc:
    _record_optional_import_error("keyring", exc)

try:
    import cx_Oracle  # noqa: F401
except Exception as exc:
    _record_optional_import_error("cx_Oracle", exc)

try:
    from selenium import webdriver  # noqa: F401
    from selenium.webdriver.chrome.service import Service  # noqa: F401
    from selenium.webdriver.common.by import By  # noqa: F401
    from selenium.webdriver.support import expected_conditions as EC  # noqa: F401
    from selenium.webdriver.support.ui import WebDriverWait  # noqa: F401
except Exception as exc:
    _record_optional_import_error("selenium", exc)

# stdlib commonly used by plugins
import base64  # noqa: F401
import calendar  # noqa: F401
import os  # noqa: F401
import sys  # noqa: F401
import time  # noqa: F401
from concurrent.futures import ThreadPoolExecutor, as_completed  # noqa: F401
from datetime import datetime  # noqa: F401
from typing import Optional, Sequence  # noqa: F401


if __name__ == "__main__":
    if OPTIONAL_IMPORT_ERRORS:
        print("Advertencia: dependencias opcionales no disponibles para algunos plugins:")
        for module_name, error in sorted(OPTIONAL_IMPORT_ERRORS.items()):
            print(f"  - {module_name}: {error}")
    run_app()
