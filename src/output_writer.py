"""
Output Writer
Writes the reconciled data to Excel with formula preservation.
Uses openpyxl for workbook manipulation to preserve formulas, formatting, and structure.
"""

import os
import shutil
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet


class OutputWriteError(Exception):
    """Raised when output writing fails."""
    pass


def create_output_path(output_dir: str, prefix: str = "Master_Updated") -> str:
    """
    Create a timestamped output file path.

    Args:
        output_dir: Output directory.
        prefix: Filename prefix.

    Returns:
        Full path to the output file.
    """
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{prefix}_{timestamp}.xlsx"
    return os.path.join(output_dir, filename)


def create_archive_copy(source_path: str, archive_dir: str) -> Optional[str]:
    """
    Create an archive copy of a file before modification.

    Args:
        source_path: Path to the original file.
        archive_dir: Archive directory.

    Returns:
        Path to the archive copy, or None if source doesn't exist.
    """
    if not os.path.isfile(source_path):
        return None

    os.makedirs(archive_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    basename = os.path.splitext(os.path.basename(source_path))[0]
    ext = os.path.splitext(source_path)[1]
    archive_name = f"{basename}_before_{timestamp}{ext}"
    archive_path = os.path.join(archive_dir, archive_name)

    shutil.copy2(source_path, archive_path)
    return archive_path


def write_output_workbook(
    output_df: pd.DataFrame,
    output_path: str,
    source_workbook_path: Optional[str] = None,
    sheet_name: str = "FCB",
    formula_columns: Optional[Dict[str, str]] = None
) -> str:
    """
    Write the output DataFrame to an Excel workbook.

    Strategy:
        1. If a source workbook exists, copy it as the base (preserving
           other sheets, pivots, formulas in non-target sheets).
        2. Clear the target sheet data rows (preserve header).
        3. Write updated data.
        4. Re-apply known formula patterns for formula columns.

    Args:
        output_df: The reconciled output DataFrame.
        output_path: Path to write the output file.
        source_workbook_path: Path to the original workbook (for structure preservation).
        sheet_name: Name of the target sheet.
        formula_columns: Dict of column_name -> formula_template for formula columns.
                        Formula templates use structured references like '=[@Resolved]-[@Created]'.

    Returns:
        Path to the written output file.
    """
    if source_workbook_path and os.path.isfile(source_workbook_path):
        # Copy source workbook to preserve other sheets, pivots, formatting
        shutil.copy2(source_workbook_path, output_path)
        wb = load_workbook(output_path)

        if sheet_name not in wb.sheetnames:
            raise OutputWriteError(
                f"Sheet '{sheet_name}' not found in workbook. "
                f"Available sheets: {wb.sheetnames}"
            )

        ws = wb[sheet_name]

        # Detect existing formula patterns before clearing
        detected_formulas = _detect_formula_patterns(ws)

        # Merge with explicitly configured formulas
        if formula_columns:
            detected_formulas.update(formula_columns)

        # Clear data rows (preserve header in row 1)
        _clear_data_rows(ws)

        # Write data rows
        _write_dataframe_to_sheet(ws, output_df, detected_formulas)

    else:
        # Create new workbook
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name

        detected_formulas = formula_columns or {}
        _write_dataframe_to_sheet(ws, output_df, detected_formulas, write_header=True)

    # Save
    try:
        wb.save(output_path)
    except Exception as e:
        raise OutputWriteError(f"Cannot save output workbook: {e}")

    return output_path


def _detect_formula_patterns(ws: Worksheet) -> Dict[str, str]:
    """
    Detect formula patterns from data rows of a worksheet.
    Scans all rows (not just row 2) since formulas may only appear
    in rows where dependent values exist.
    Returns dict of column_header -> formula_string (first formula found).
    """
    patterns = {}
    if ws.max_row < 2:
        return patterns

    # Get headers from row 1
    headers = {}
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        if cell.value is not None:
            headers[col_idx] = str(cell.value)

    # Scan all data rows for formulas
    for row_idx in range(2, ws.max_row + 1):
        for col_idx, header in headers.items():
            if header in patterns:
                continue  # Already found a formula for this column
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                patterns[header] = cell.value

    return patterns


def _clear_data_rows(ws: Worksheet):
    """Clear all data rows below the header (row 1)."""
    if ws.max_row < 2:
        return

    for row_idx in range(ws.max_row, 1, -1):
        ws.delete_rows(row_idx)


def _write_dataframe_to_sheet(
    ws: Worksheet,
    df: pd.DataFrame,
    formula_columns: Dict[str, str],
    write_header: bool = False
):
    """
    Write DataFrame data to worksheet rows.

    Args:
        ws: Target worksheet.
        df: DataFrame to write.
        formula_columns: Dict of column_name -> formula_template.
        write_header: If True, write column names in row 1.
    """
    columns = list(df.columns)

    if write_header:
        for col_idx, col_name in enumerate(columns, 1):
            ws.cell(row=1, column=col_idx, value=col_name)

    start_row = 2  # Data starts in row 2 (after header)

    for row_offset, (_, row_data) in enumerate(df.iterrows()):
        row_num = start_row + row_offset
        for col_idx, col_name in enumerate(columns, 1):
            if col_name in formula_columns:
                # Write formula instead of static value
                ws.cell(row=row_num, column=col_idx, value=formula_columns[col_name])
            else:
                value = row_data[col_name]
                # Convert pandas NaN/NaT to None for openpyxl
                if pd.isna(value):
                    ws.cell(row=row_num, column=col_idx, value=None)
                else:
                    ws.cell(row=row_num, column=col_idx, value=value)


def write_reconciliation_report_excel(
    report_data: Dict[str, Any],
    output_path: str
) -> str:
    """
    Write the reconciliation report to an Excel file.

    Args:
        report_data: Dict with report sections.
        output_path: Path to write the report.

    Returns:
        Path to the written report.
    """
    wb = Workbook()

    # Summary sheet
    ws_summary = wb.active
    ws_summary.title = "Summary"

    row = 1
    ws_summary.cell(row=row, column=1, value="FCB Incident Tracker — Reconciliation Report")
    row += 1
    ws_summary.cell(row=row, column=1, value=f"Generated: {datetime.now().isoformat()}")
    row += 2

    summary = report_data.get("summary", {})
    for key, value in summary.items():
        ws_summary.cell(row=row, column=1, value=key)
        ws_summary.cell(row=row, column=2, value=value)
        row += 1

    # Changes sheet
    if "changes" in report_data and report_data["changes"]:
        ws_changes = wb.create_sheet("Changes")
        ws_changes.cell(row=1, column=1, value="Incident Number")
        ws_changes.cell(row=1, column=2, value="Field Changes")

        row = 2
        for number, changes in report_data["changes"].items():
            for change in changes:
                ws_changes.cell(row=row, column=1, value=number)
                ws_changes.cell(row=row, column=2, value=change)
                row += 1

    # New incidents sheet
    if "new_incidents" in report_data and report_data["new_incidents"]:
        ws_new = wb.create_sheet("New Incidents")
        ws_new.cell(row=1, column=1, value="Incident Number")
        for i, number in enumerate(report_data["new_incidents"], 2):
            ws_new.cell(row=i, column=1, value=number)

    # IC Lookup Misses sheet
    if "ic_misses" in report_data and report_data["ic_misses"]:
        ws_ic = wb.create_sheet("IC Lookup Misses")
        ws_ic.cell(row=1, column=1, value="Details")
        for i, miss in enumerate(report_data["ic_misses"], 2):
            ws_ic.cell(row=i, column=1, value=miss)

    wb.save(output_path)
    return output_path
