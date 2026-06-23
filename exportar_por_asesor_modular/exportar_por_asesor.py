"""
Exportar un Excel en múltiples archivos, agrupando por la columna "Asesor".
Cada asesor recibe un archivo Excel con únicamente sus filas.

Uso básico:
    python exportar_por_asesor.py "PROCESOS DE VENTA  (1)(1).xlsx"

Uso con carpeta de salida:
    python exportar_por_asesor.py "PROCESOS DE VENTA  (1)(1).xlsx" --output-dir "exports_asesores"

Uso indicando hoja:
    python exportar_por_asesor.py "PROCESOS DE VENTA  (1)(1).xlsx" --sheet "Export"
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Iterable

import pandas as pd


DEFAULT_SHEET = "Export"
DEFAULT_GROUP_COL = "Asesor"


def clean_filename(value: object, max_len: int = 80) -> str:
    """Convierte el nombre del asesor en un nombre de archivo válido para Windows."""
    text = "Sin_Asesor" if pd.isna(value) else str(value).strip()
    if not text:
        text = "Sin_Asesor"

    text = re.sub(r"[\\/:*?\"<>|]+", "_", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:max_len]


def normalize_column_name(col: object) -> str:
    """Normaliza nombres de columna para encontrar columnas aunque tengan espacios extra."""
    return re.sub(r"\s+", " ", str(col).strip()).lower()


def find_column(df: pd.DataFrame, expected_col: str) -> str:
    """Busca una columna por nombre, tolerando mayúsculas, minúsculas y espacios."""
    expected_norm = normalize_column_name(expected_col)

    for col in df.columns:
        if normalize_column_name(col) == expected_norm:
            return col

    available = ", ".join(map(str, df.columns))
    raise ValueError(
        f"No encontré la columna '{expected_col}'. Columnas disponibles: {available}"
    )


def read_source_excel(input_path: Path, sheet_name: str) -> pd.DataFrame:
    """Lee el Excel fuente y devuelve un DataFrame."""
    if not input_path.exists():
        raise FileNotFoundError(f"No existe el archivo: {input_path}")

    return pd.read_excel(input_path, sheet_name=sheet_name)


def autofit_excel_columns(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    """Ajusta ancho de columnas de forma simple para que el Excel sea legible."""
    worksheet = writer.sheets[sheet_name]

    for idx, col in enumerate(df.columns):
        series = df[col].astype(str).fillna("")
        max_len = max([len(str(col)), *series.map(len).head(1000).tolist()])
        width = min(max(max_len + 2, 10), 45)
        worksheet.set_column(idx, idx, width)


def export_group_to_excel(
    group_df: pd.DataFrame,
    output_path: Path,
    sheet_name: str = "Data",
) -> None:
    """Exporta un grupo filtrado a un archivo Excel individual."""
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        group_df.to_excel(writer, index=False, sheet_name=sheet_name)
        autofit_excel_columns(writer, sheet_name, group_df)

        worksheet = writer.sheets[sheet_name]
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(group_df), len(group_df.columns) - 1)


def build_summary(records: Iterable[dict], output_dir: Path) -> None:
    """Crea un resumen de archivos exportados por asesor."""
    summary_df = pd.DataFrame(records).sort_values("filas", ascending=False)
    summary_path = output_dir / "00_resumen_exportacion_por_asesor.xlsx"

    with pd.ExcelWriter(summary_path, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Resumen")
        autofit_excel_columns(writer, "Resumen", summary_df)
        writer.sheets["Resumen"].freeze_panes(1, 0)


def export_by_advisor(
    input_path: Path,
    output_dir: Path,
    sheet_name: str = DEFAULT_SHEET,
    group_col: str = DEFAULT_GROUP_COL,
) -> None:
    """Función principal: agrupa por asesor y exporta un Excel por cada asesor."""
    output_dir.mkdir(parents=True, exist_ok=True)

    df = read_source_excel(input_path, sheet_name=sheet_name)
    asesor_col = find_column(df, group_col)

    # Normaliza solo para agrupar; no altera el valor visual de tus columnas originales.
    df[asesor_col] = df[asesor_col].fillna("Sin Asesor").astype(str).str.strip()
    df.loc[df[asesor_col].eq(""), asesor_col] = "Sin Asesor"

    summary_records = []

    for asesor, group_df in df.groupby(asesor_col, dropna=False, sort=True):
        safe_name = clean_filename(asesor)
        output_path = output_dir / f"{safe_name}.xlsx"

        # Copia para evitar warnings y mantener solo filas del asesor.
        group_df = group_df.copy()
        export_group_to_excel(group_df, output_path)

        summary_records.append(
            {
                "asesor": asesor,
                "filas": len(group_df),
                "archivo": str(output_path.name),
            }
        )

        print(f"OK: {asesor} -> {output_path.name} ({len(group_df)} filas)")

    build_summary(summary_records, output_dir)
    print(f"\nListo. Archivos exportados en: {output_dir.resolve()}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Exporta un Excel en archivos separados por Asesor."
    )
    parser.add_argument("input_file", help="Ruta del Excel fuente")
    parser.add_argument(
        "--output-dir",
        default="exports_por_asesor",
        help="Carpeta de salida. Default: exports_por_asesor",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET,
        help=f"Nombre de la hoja a leer. Default: {DEFAULT_SHEET}",
    )
    parser.add_argument(
        "--group-col",
        default=DEFAULT_GROUP_COL,
        help=f"Columna para agrupar. Default: {DEFAULT_GROUP_COL}",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    export_by_advisor(
        input_path=Path(args.input_file),
        output_dir=Path(args.output_dir),
        sheet_name=args.sheet,
        group_col=args.group_col,
    )


if __name__ == "__main__":
    main()
