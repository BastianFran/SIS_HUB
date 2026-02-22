# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 09:49:07 2025

@author: BBRUNA
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

import fitz  # type: ignore
import pandas as pd


def ensure_empty_folder(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    for child in path.iterdir():
        if child.is_dir():
            shutil.rmtree(child, ignore_errors=True)
        else:
            try:
                child.unlink()
            except Exception:
                pass
    return path


def merge_pdfs_in_place(folder: Path, password: str | None = None, output_name: str = "pdf-masivo.pdf") -> Path | None:
    pdf_files = sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"])
    if not pdf_files:
        return None

    output_pdf = folder / output_name
    merged = fitz.open()
    processed: list[Path] = []

    try:
        for pdf_path in pdf_files:
            if pdf_path == output_pdf:
                continue
            with fitz.open(pdf_path) as part:
                if part.needs_pass:
                    if not password or not part.authenticate(password):
                        continue
                merged.insert_pdf(part)
            processed.append(pdf_path)

        if merged.page_count == 0:
            return None
        merged.save(output_pdf)
    finally:
        merged.close()

    for pdf_path in processed:
        try:
            pdf_path.unlink()
        except Exception:
            pass
    return output_pdf


def zip_xmls_in_place(folder: Path, output_name: str = "xml-masivo.zip") -> Path | None:
    xml_files = sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".xml"])
    if not xml_files:
        return None

    zip_path = folder / output_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for xml_file in xml_files:
            zf.write(xml_file, arcname=xml_file.name)

    for xml_file in xml_files:
        try:
            xml_file.unlink()
        except Exception:
            pass
    return zip_path


def write_excel(df: pd.DataFrame, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    return output_path


def write_flat_file(df: pd.DataFrame, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, sep="|", index=False, encoding="utf-8-sig")
    return output_path
