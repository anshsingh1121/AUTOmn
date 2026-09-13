"""
Data Loader
Reads Excel workbooks for Master Tracker, ServiceNow export, and IC Lookup.
Uses pandas for data manipulation and openpyxl for workbook structure.
"""

import os
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


class DataLoadError(Exception):
    """Raised when a data source cannot be loaded."""
    pass


def load_excel_to_dataframe(
    file_path: str,
    sheet_name: str,
    required_columns: Optional[List[str]] = None
) -> pd.DataFrame:
    """
    Load an Excel sheet into a pandas DataFrame.

    Args:
        file_path: Path to the Excel file.
        sheet_name: Name of the sheet to read.
        required_columns: Optional list of columns to validate after loading.

    Returns:
        pandas DataFrame.

    Raises:
        DataLoadError: If the file or sheet cannot be loaded.
    """
    if not os.path.isfile(file_path):
        raise DataLoadError(f"File not found: {file_path}")

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    except ValueError as e:
        raise DataLoadError(
            f"Sheet '{sheet_name}' not found in '{file_path}'. "
            f"Available sheets can be checked with openpyxl. Error: {e}"
        )
    except Exception as e:
        raise DataLoadError(f"Error reading '{file_path}' sheet '{sheet_name}': {e}")

    return df


def load_master_tracker(config) -> Tuple[pd.DataFrame, Any]:
    """
    Load the Master Tracker workbook.

    Returns:
        Tuple of (DataFrame, openpyxl Workbook) — the workbook is kept
        for formula preservation during output writing.
    """
    path = config.master_tracker_path
    sheet = config.master_sheet

    if not os.path.isfile(path):
        raise DataLoadError(f"Master Tracker file not found: {path}")

    # Load as DataFrame for data operations
    df = load_excel_to_dataframe(path, sheet)

    # Also load the openpyxl workbook for structure/formula preservation
    try:
        wb = load_workbook(path, data_only=False)  # data_only=False preserves formulas
    except Exception as e:
        raise DataLoadError(f"Cannot open Master Tracker workbook: {e}")

    return df, wb


def load_servicenow(config) -> pd.DataFrame:
    """
    Load the ServiceNow monthly export.

    Returns:
        pandas DataFrame.
    """
    path = config.servicenow_input_path
    sheet = config.servicenow_sheet
    df = load_excel_to_dataframe(path, sheet)

    # AUTO-CLEAN: Remove empty rows or Excel pivot summary rows (e.g. "Grand Total") at the bottom
    if "Number" in df.columns:
        df = df.dropna(subset=["Number"])
        df = df[df["Number"].astype(str).str.startswith("INC")]

    return df


def load_ic_lookup(config) -> pd.DataFrame:
    """
    Load the IC Lookup table.

    Returns:
        pandas DataFrame with IC-to-Region mappings.
    """
    path = config.ic_lookup_path
    sheet = config.ic_lookup_sheet
    return load_excel_to_dataframe(path, sheet)


def detect_formula_columns(
    ws: Worksheet,
    header_row: int = 1
) -> Dict[str, List[int]]:
    """
    Detect columns containing Excel formulas in a worksheet.

    Args:
        ws: openpyxl Worksheet object.
        header_row: Row number containing headers (1-indexed).

    Returns:
        Dict mapping column header names to list of row numbers with formulas.
    """
    formula_columns = {}

    # Get headers
    headers = {}
    for col_idx, cell in enumerate(ws[header_row], 1):
        if cell.value is not None:
            headers[col_idx] = str(cell.value)

    # Scan data rows for formulas
    for row_idx in range(header_row + 1, ws.max_row + 1):
        for col_idx, header_name in headers.items():
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                if header_name not in formula_columns:
                    formula_columns[header_name] = []
                formula_columns[header_name].append(row_idx)

    return formula_columns


def get_sheet_names(file_path: str) -> List[str]:
    """Return list of sheet names in a workbook."""
    if not os.path.isfile(file_path):
        raise DataLoadError(f"File not found: {file_path}")
    try:
        wb = load_workbook(file_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    except Exception as e:
        raise DataLoadError(f"Cannot read workbook: {e}")
