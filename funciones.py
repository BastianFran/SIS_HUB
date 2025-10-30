"""Funciones de negocio: lectura, concatenación y guardado de DataFrames."""
from __future__ import annotations

from pathlib import Path
from typing import Optional
import pandas as pd


def leer_excel(path: Path | str, sheet: Optional[str | int] = None) -> pd.DataFrame:
    """Lee un archivo Excel y retorna un DataFrame.
    Si `sheet` es None, usa la primera hoja.
    """
    path = Path(path)
    if sheet in ("", None):
        sheet = 0
    df = pd.read_excel(path, sheet_name=sheet)
    if isinstance(df, dict):
        df = next(iter(df.values()))
    return df


def concatenar(df1: pd.DataFrame, df2: pd.DataFrame, ignore_index: bool = True) -> pd.DataFrame:
    """Concatena dos DataFrames verticalmente."""
    return pd.concat([df1, df2], ignore_index=ignore_index)


def guardar(df: pd.DataFrame, output_dir: Path | str, filename_stem: str, fmt: str = "xlsx") -> Path:
    """Guarda el DataFrame en `output_dir/filename_stem.fmt`.
    Soporta 'xlsx' (requiere openpyxl) y 'csv'. Retorna la ruta final.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    fmt = (fmt or "xlsx").lower()
    if fmt == "xlsx":
        out = output_dir / f"{filename_stem}.xlsx"
        df.to_excel(out, index=False)
    elif fmt == "csv":
        out = output_dir / f"{filename_stem}.csv"
        df.to_csv(out, index=False, encoding="utf-8-sig")
    else:
        raise ValueError(f"Formato no soportado: {fmt!r}")
    return out
