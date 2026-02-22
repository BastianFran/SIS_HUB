# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 09:56:08 2025

@author: BBRUNA
"""

from __future__ import annotations

from pathlib import Path


DEFAULT_CC = (
    "SoporteInstitucionales@banchile.cl;"
    "TradingRVInstitucional@banchile.cl;"
    "InstitutionalSales@banchile.cl"
)


def create_outlook_draft(
    *,
    to_recipients: list[str],
    subject: str,
    body: str,
    attachments_dir: Path,
    cc_recipients: str = DEFAULT_CC,
) -> None:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError("Falta dependencia pywin32 para generar correo en Outlook.") from exc

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = ";".join([x for x in to_recipients if x.strip()])
    mail.CC = cc_recipients
    mail.Subject = subject
    mail.Body = body

    if attachments_dir.exists():
        for file_path in sorted(attachments_dir.iterdir()):
            if file_path.is_file():
                mail.Attachments.Add(str(file_path))

    mail.Display()
