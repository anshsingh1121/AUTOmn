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
    Create a timestamped output file path inside a Year/Month folder structure.

    Args:
        output_dir: Base output directory.
        prefix: Filename prefix.

    Returns:
        Full path to the output file.
    """
    now = datetime.now()
    nested_dir = os.path.join(output_dir, now.strftime("%Y"), now.strftime("%m_%B"))
    os.makedirs(nested_dir, exist_ok=True)
    
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    filename = f"{prefix}_{timestamp}.xlsx"
    return os.path.join(nested_dir, filename)


def create_archive_copy(source_path: str, archive_dir: str) -> Optional[str]:
    """
    Create an archive copy of a file inside a Year/Month folder structure before modification.

    Args:
        source_path: Path to the original file.
        archive_dir: Base archive directory.

    Returns:
        Path to the archive copy, or None if source doesn't exist.
    """
    if not os.path.isfile(source_path):
        return None

    now = datetime.now()
    nested_dir = os.path.join(archive_dir, now.strftime("%Y"), now.strftime("%m_%B"))
    os.makedirs(nested_dir, exist_ok=True)
    
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    basename = os.path.splitext(os.path.basename(source_path))[0]
    ext = os.path.splitext(source_path)[1]
    archive_name = f"{basename}_before_{timestamp}{ext}"
    archive_path = os.path.join(nested_dir, archive_name)

    shutil.copy2(source_path, archive_path)
    return archive_path


def write_output_workbook(
    output_df: pd.DataFrame,
    output_path: str,
    source_workbook_path: Optional[str] = None,
    target_sheets: Optional[List[str]] = None,
    formula_columns: Optional[Dict[str, str]] = None
) -> str:
    """
    Write the output DataFrame to an Excel workbook.

    Strategy:
        1. If a source workbook exists, copy it as the base (preserving
           other sheets, pivots, formulas in non-target sheets).
        2. Iterate over target sheets (e.g., Consolidated, FCB).
        3. Clear the target sheet data rows (preserve header).
        4. Write updated data mapped to existing headers.
        5. Refresh Pivot Tables.

    Args:
        output_df: The reconciled output DataFrame.
        output_path: Path to write the output file.
        source_workbook_path: Path to the original workbook.
        target_sheets: List of sheet names to update.
        formula_columns: Dict of column_name -> formula_template for formula columns.

    Returns:
        Path to the written output file.
    """
    target_sheets = target_sheets or ["FCB"]

    if source_workbook_path and os.path.isfile(source_workbook_path):
        # Copy source workbook to preserve other sheets, pivots, formatting
        shutil.copy2(source_workbook_path, output_path)
        wb = load_workbook(output_path)

        for sheet_name in target_sheets:
            if sheet_name not in wb.sheetnames:
                print(f"Warning: Sheet '{sheet_name}' not found in workbook, skipping.")
                continue

            ws = wb[sheet_name]

            # Detect existing formula patterns before clearing
            detected_formulas = _detect_formula_patterns(ws)

            # Force remove Duration from detected formulas so our Python calculation is used
            if "Duration" in detected_formulas:
                del detected_formulas["Duration"]

            # Merge with explicitly configured formulas
            if formula_columns:
                for k, v in formula_columns.items():
                    if k != "Duration":
                        detected_formulas[k] = v

            # Clear data rows (preserve header in row 1)
            _clear_data_rows(ws)

            # Write data rows
            _write_dataframe_to_sheet(ws, output_df, detected_formulas)

    else:
        # Create new workbook
        wb = Workbook()
        
        for i, sheet_name in enumerate(target_sheets):
            if i == 0:
                ws = wb.active
                ws.title = sheet_name
            else:
                ws = wb.create_sheet(sheet_name)

            detected_formulas = formula_columns or {}
            _write_dataframe_to_sheet(ws, output_df, detected_formulas, write_header=True)

    # Save
    try:
        wb.save(output_path)
    except Exception as e:
        raise OutputWriteError(f"Cannot save output workbook: {e}")

    # Use win32com to perform a hard physical refresh of all pivot tables, bypassing Protected View issues
    _hard_refresh_pivots_com(output_path)

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
    Write DataFrame data to worksheet rows, mapping to the sheet's existing headers.
    """
    df_columns = set(df.columns)

    if write_header:
        columns_to_write = list(df.columns)
        for col_idx, col_name in enumerate(columns_to_write, 1):
            ws.cell(row=1, column=col_idx, value=col_name)
        header_map = {col_idx: col_name for col_idx, col_name in enumerate(columns_to_write, 1)}
    else:
        # Read existing headers from row 1
        header_map = {}
        for col_idx in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=1, column=col_idx).value
            if cell_val is not None:
                header_map[col_idx] = str(cell_val)

    start_row = 2  # Data starts in row 2 (after header)

    for row_offset, (_, row_data) in enumerate(df.iterrows()):
        row_num = start_row + row_offset
        for col_idx, col_name in header_map.items():
            # If the sheet's column doesn't exist in our DataFrame, skip writing to it
            if col_name not in df_columns:
                continue
                
            if col_name in formula_columns:
                ws.cell(row=row_num, column=col_idx, value=formula_columns[col_name])
            else:
                value = row_data[col_name]
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

def _hard_refresh_pivots_com(file_path: str):
    """
    Uses native Windows COM (pywin32) to invisibly open Excel, force a complete RefreshAll, 
    and save the file. This guarantees Pivot Tables update and bypasses Protected View bugs.
    """
    try:
        import win32com.client
        import os
        
        abs_path = os.path.abspath(file_path)
        
        # Initialize COM
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        
        try:
            wb = excel.Workbooks.Open(abs_path)
            
            # Step 1: Expand ListObjects (Excel Tables) dynamically
            for ws in wb.Worksheets:
                try:
                    last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row  # xlUp
                    last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column  # xlToLeft
                    
                    if last_row > 1 and last_col > 0:
                        for obj in ws.ListObjects:
                            new_range = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col))
                            obj.Resize(new_range)
                except Exception:
                    pass
            
            # Step 2: Expand PivotCaches manually if they don't use a ListObject
            for pc in wb.PivotCaches():
                try:
                    src = pc.SourceData
                    if isinstance(src, str) and "!" in src:
                        sheet_name = src.split("!")[0].replace("'", "")
                        ws = wb.Worksheets(sheet_name)
                        last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row
                        last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column
                        if last_row > 1 and last_col > 0:
                            new_src = f"'{sheet_name}'!R1C1:R{last_row}C{last_col}"
                            pc.SourceData = new_src
                except Exception:
                    pass
            
            wb.RefreshAll()
            excel.CalculateUntilAsyncQueriesDone()
            wb.Save()
            wb.Close(True)
        finally:
            excel.Quit()
            del excel
    except ImportError:
        print("Notice: win32com not installed. Pivot tables will require manual refresh.")
    except Exception as e:
        print(f"Notice: Background Pivot refresh encountered an error: {e}")
