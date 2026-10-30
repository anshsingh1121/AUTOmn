"""
Output Writer
Writes the reconciled data to Excel.

Strategy:
- Main output workbook:
  1. Primary: Uses win32com (COM automation) to write data into a copy
     of the original master tracker. This preserves ALL Excel internals (Pivot Tables,
     Charts, Macros, Data Models, Conditional Formatting).
  2. Fallback: If COM is unavailable (e.g., Office not installed, pywin32 missing,
     or security policies blocking COM automation), falls back to openpyxl
     with keep_vba=True to preserve macros (.xlsm), all other sheets, and structure
     without corruption.
- Reconciliation report: Uses openpyxl to create a simple NEW workbook (no pivots
  or charts to corrupt).
"""

import gc
import os
import shutil
from datetime import datetime
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd


class OutputWriteError(Exception):
    """Exception raised for errors during output writing."""
    pass


# ============================================================
# PATH & ARCHIVE HELPERS
# ============================================================

def create_output_path(
    output_dir: str,
    prefix: str = "Master_Updated",
    source_path: Optional[str] = None,
) -> str:
    """
    Create a timestamped output file path inside a Year/Month folder structure.

    Args:
        output_dir: Base output directory.
        prefix: Filename prefix.
        source_path: If provided, the output file extension is copied from this
                     file (e.g. .xlsm stays .xlsm). Otherwise defaults to .xlsx.
    """
    now = datetime.now()
    year_folder = now.strftime("%Y")
    month_folder = now.strftime("%m_%B")

    target_dir = os.path.join(output_dir, year_folder, month_folder)
    os.makedirs(target_dir, exist_ok=True)

    timestamp = now.strftime("%Y%m%d_%H%M%S")

    # Preserve source extension (.xlsm, .xlsx, etc.)
    ext = ".xlsx"
    if source_path:
        _, src_ext = os.path.splitext(source_path)
        if src_ext:
            ext = src_ext

    return os.path.join(target_dir, f"{prefix}_{timestamp}{ext}")


def create_archive_copy(source_path: str, archive_dir: str) -> str:
    """Create a backup copy of the original master tracker."""
    if not source_path or not os.path.isfile(source_path):
        return ""

    now = datetime.now()
    year_folder = now.strftime("%Y")
    month_folder = now.strftime("%m_%B")

    target_dir = os.path.join(archive_dir, year_folder, month_folder)
    os.makedirs(target_dir, exist_ok=True)

    timestamp = now.strftime("%Y%m%d_%H%M%S")
    filename = os.path.basename(source_path)
    name, ext = os.path.splitext(filename)

    archive_path = os.path.join(target_dir, f"{name}_before_{timestamp}{ext}")
    shutil.copy2(source_path, archive_path)
    return archive_path


# ============================================================
# COM TYPE CONVERSION HELPER
# ============================================================

def _to_com_safe(val: Any) -> Any:
    """
    Convert a Python / numpy / pandas value to a type that the Windows COM
    bridge can marshal into an Excel cell without error.
    """
    if val is None:
        return ""

    if isinstance(val, np.integer):
        return int(val)

    if isinstance(val, np.floating):
        if np.isnan(val):
            return ""
        return float(val)

    if isinstance(val, np.bool_):
        return bool(val)

    if isinstance(val, (pd.Timestamp, datetime)):
        if pd.isna(val):
            return ""
        delta = pd.Timestamp(val) - pd.Timestamp("1899-12-30")
        return float(delta.total_seconds() / 86400.0)

    if isinstance(val, float) and val != val:
        return ""

    if isinstance(val, str):
        if val.strip() == "" or val.strip().lower() == "nan":
            return ""
        return val

    return val


# ============================================================
# OPENPYXL FALLBACK WRITER (Preserves .xlsm, all sheets, no corruption)
# ============================================================

def write_output_workbook_openpyxl(
    output_df: pd.DataFrame,
    output_path: str,
    source_workbook_path: str,
    target_sheets: Optional[List[str]] = None,
) -> str:
    """
    Safe fallback when Excel COM is unavailable or disabled.
    Uses openpyxl with keep_vba=True to preserve macros (.xlsm), all other sheets,
    and styles without destroying the workbook structure.
    """
    import openpyxl

    is_xlsm = str(source_workbook_path).lower().endswith(".xlsm") or str(output_path).lower().endswith(".xlsm")
    target_sheets = target_sheets or ["Consolidated"]

    try:
        wb = openpyxl.load_workbook(source_workbook_path, keep_vba=is_xlsm, data_only=False)
    except Exception as e:
        raise OutputWriteError(f"Openpyxl cannot load source workbook '{source_workbook_path}': {e}")

    df_write = output_df.copy()

    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            print(f"  Warning (openpyxl): Sheet '{sheet_name}' not found in workbook, skipping.")
            continue

        ws = wb[sheet_name]

        # Read headers from row 1
        headers = []
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=1, column=c).value
            headers.append(str(val).strip() if val is not None else "")

        # Detect new columns that are not in headers
        headers_set = set(headers)
        new_cols = [c for c in df_write.columns if c not in headers_set and c != ""]
        if new_cols:
            start_col = len(headers) + 1
            for i, col_name in enumerate(new_cols):
                ws.cell(row=1, column=start_col + i, value=col_name)
            headers.extend(new_cols)

        # Clear existing data rows (row 2 down)
        if ws.max_row >= 2:
            ws.delete_rows(2, ws.max_row - 1)

        # Write data rows
        for _, row in df_write.iterrows():
            row_vals = []
            for h in headers:
                if h in df_write.columns:
                    val = row[h]
                    if pd.isna(val):
                        val = ""
                    elif isinstance(val, (pd.Timestamp, datetime)):
                        val = val.to_pydatetime() if hasattr(val, "to_pydatetime") else val
                    row_vals.append(val)
                else:
                    row_vals.append("")
            ws.append(row_vals)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    wb.close()
    return output_path


# ============================================================
# INTERNAL COM WORKER
# ============================================================

def _com_populate_workbook(
    excel,
    abs_output: str,
    df_write: pd.DataFrame,
    target_sheets: List[str],
    formula_columns: Optional[Dict[str, str]],
):
    """
    Subroutine isolated in its own scope so that all COM sheet, range, and table
    proxies go out of scope and are collected before excel.Quit() is called.
    """
    wb = excel.Workbooks.Open(abs_output, UpdateLinks=0)
    try:
        for sheet_name in target_sheets:
            try:
                ws = wb.Worksheets(sheet_name)
            except Exception:
                print(f"  Warning: Sheet '{sheet_name}' not found in workbook, skipping.")
                continue

            last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column  # xlToLeft
            if last_col < 1:
                continue

            headers: List[str] = []
            for c in range(1, last_col + 1):
                val = ws.Cells(1, c).Value
                headers.append(str(val).strip() if val else "")

            # Detect formulas from row 2
            existing_formulas: Dict[int, str] = {}
            old_last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp
            if old_last_row >= 2:
                for c in range(1, last_col + 1):
                    cell = ws.Cells(2, c)
                    if cell.HasFormula:
                        h = headers[c - 1]
                        if h.lower() != "duration":
                            existing_formulas[c] = cell.FormulaR1C1

            if formula_columns:
                for c_idx, h in enumerate(headers, start=1):
                    if h in formula_columns and h.lower() != "duration":
                        existing_formulas[c_idx] = formula_columns[h]

            # Clear old data rows (row 2 -> end)
            if old_last_row >= 2:
                ws.Range(
                    ws.Cells(2, 1),
                    ws.Cells(old_last_row, last_col),
                ).ClearContents()

            num_records = len(df_write)
            if num_records == 0:
                continue

            headers_set = set(headers)
            new_cols = [c for c in df_write.columns if c not in headers_set and c != ""]
            if new_cols:
                for i, col_name in enumerate(new_cols):
                    ws.Cells(1, last_col + 1 + i).Value = col_name
                headers.extend(new_cols)
                new_last_col = last_col + len(new_cols)
            else:
                new_last_col = last_col

            write_data: List[List[Any]] = []
            for _, row in df_write.iterrows():
                row_data: List[Any] = []
                for c_idx, h in enumerate(headers, start=1):
                    if c_idx in existing_formulas:
                        row_data.append("")
                    elif h in df_write.columns:
                        row_data.append(_to_com_safe(row[h]))
                    else:
                        row_data.append("")
                write_data.append(row_data)

            target_range = ws.Range(
                ws.Cells(2, 1),
                ws.Cells(1 + num_records, new_last_col),
            )
            target_range.Value = write_data

            # Apply preserved formulas
            for c_idx, f_str in existing_formulas.items():
                f_range = ws.Range(
                    ws.Cells(2, c_idx),
                    ws.Cells(1 + num_records, c_idx)
                )
                f_range.FormulaR1C1 = f_str

            # Format date columns
            date_only_columns = {"Created", "Resolved"}
            datetime_columns = {"Actual Incident Resolve", "Actual Incident Start"}
            for c_idx, h in enumerate(headers, start=1):
                try:
                    col_range = ws.Range(
                        ws.Cells(2, c_idx),
                        ws.Cells(1 + num_records, c_idx)
                    )
                    if h in datetime_columns:
                        col_range.NumberFormat = "yyyy-mm-dd hh:mm:ss"
                    elif h in date_only_columns:
                        col_range.NumberFormat = "yyyy-mm-dd"
                except Exception:
                    pass

            try:
                ws.Columns.AutoFit()
            except Exception:
                pass

            # Resize ListObjects (Excel Tables)
            try:
                for i in range(1, ws.ListObjects.Count + 1):
                    obj = ws.ListObjects(i)
                    try:
                        original_style = obj.TableStyle
                    except Exception:
                        original_style = None
                    try:
                        original_banding = obj.ShowTableStyleRowStripes
                    except Exception:
                        original_banding = True

                    tbl_width = obj.Range.Columns.Count
                    tbl_start_col = obj.Range.Column
                    tbl_start_row = obj.Range.Row

                    new_range = ws.Range(
                        ws.Cells(tbl_start_row, tbl_start_col),
                        ws.Cells(tbl_start_row + num_records, tbl_start_col + tbl_width - 1),
                    )
                    obj.Resize(new_range)

                    try:
                        if original_style:
                            obj.TableStyle = original_style
                        obj.ShowTableStyleRowStripes = original_banding
                    except Exception:
                        pass
            except Exception as e:
                print(f"  Warning: Could not resize table: {e}")

            print(f"  Sheet '{sheet_name}': {num_records} records written.")

        # Update Pivots Sheet
        try:
            from src.pivot_updater import update_pivots_sheet
            update_pivots_sheet(wb, df_write)
        except Exception as e:
            print(f"  Warning: Pivots sheet update failed: {e}")

        wb.Save()
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass


# ============================================================
# MAIN OUTPUT WORKBOOK (COM with OpenPyXL Fallback)
# ============================================================

def write_output_workbook(
    output_df: pd.DataFrame,
    output_path: str,
    source_workbook_path: Optional[str] = None,
    target_sheets: Optional[List[str]] = None,
    formula_columns: Optional[Dict[str, str]] = None,
) -> str:
    """
    Write the output DataFrame into an Excel workbook.
    Attempts native Windows COM for complete pivot/chart/table preservation.
    If COM fails or is unavailable, cleanly falls back to openpyxl with
    keep_vba=True, preventing any corruption.
    """
    if not source_workbook_path or not os.path.isfile(source_workbook_path):
        raise OutputWriteError(
            f"A source workbook template is required. Path: {source_workbook_path}"
        )

    target_sheets = target_sheets or ["Consolidated"]

    # 1. First copy pristine template -> output path
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    try:
        shutil.copy2(source_workbook_path, output_path)
    except Exception as e:
        raise OutputWriteError(f"Cannot copy template to output path: {e}")

    abs_output = os.path.abspath(output_path)

    # 2. Attempt COM-based write
    com_available = False
    try:
        import win32com.client
        import pythoncom
        com_available = True
    except ImportError:
        com_available = False

    if not com_available:
        print("  Notice: pywin32 not available. Using openpyxl fallback.")
        return write_output_workbook_openpyxl(
            output_df=output_df,
            output_path=abs_output,
            source_workbook_path=source_workbook_path,
            target_sheets=target_sheets,
        )

    # Prepare DataFrame for writing
    df_write = output_df.copy()
    date_columns = ["Created", "Resolved", "Actual Incident Resolve"]
    for col in df_write.columns:
        if col in date_columns or pd.api.types.is_datetime64_any_dtype(df_write[col]):
            df_write[col] = pd.to_datetime(df_write[col], errors='coerce')
            df_write[col] = df_write[col].apply(
                lambda x: "" if pd.isna(x) else x.to_pydatetime()
            )
    df_write = df_write.fillna("")

    pythoncom.CoInitialize()
    excel = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 1  # msoAutomationSecurityLow

        # Run population in isolated scope so COM proxies are destroyed
        _com_populate_workbook(
            excel=excel,
            abs_output=abs_output,
            df_write=df_write,
            target_sheets=target_sheets,
            formula_columns=formula_columns,
        )
        print("  Output workbook saved (Consolidated, FCB, Pivots updated).")

    except Exception as e:
        print(f"  COM write encountered an issue: {e}. Falling back to openpyxl...")
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            excel = None

        return write_output_workbook_openpyxl(
            output_df=output_df,
            output_path=abs_output,
            source_workbook_path=source_workbook_path,
            target_sheets=target_sheets,
        )

    finally:
        # Proper, clean COM teardown in correct sequence
        gc.collect()

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            del excel
            excel = None

        gc.collect()

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    return output_path


# ============================================================
# RECONCILIATION REPORT (openpyxl — safe for NEW files)
# ============================================================

def write_reconciliation_report_excel(
    report_data: Dict[str, Any],
    output_path: str,
) -> str:
    """
    Write a reconciliation report to a brand-new Excel workbook.
    """
    from openpyxl import Workbook

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb = Workbook()

    # --- Summary sheet ---
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.cell(row=1, column=1, value="Metric")
    ws_summary.cell(row=1, column=2, value="Value")
    for i, (key, value) in enumerate(
        report_data.get("summary", {}).items(), 2
    ):
        ws_summary.cell(row=i, column=1, value=str(key))
        ws_summary.cell(row=i, column=2, value=str(value))

    # --- New Incidents sheet ---
    new_incidents = report_data.get("new_incidents", [])
    if new_incidents:
        ws_new = wb.create_sheet("New Incidents")
        ws_new.cell(row=1, column=1, value="Incident Number")
        for i, num in enumerate(new_incidents, 2):
            ws_new.cell(row=i, column=1, value=num)

    # --- Changes sheet ---
    changes = report_data.get("changes", {})
    if changes:
        ws_changes = wb.create_sheet("Changes")
        ws_changes.cell(row=1, column=1, value="Incident Number")
        ws_changes.cell(row=1, column=2, value="Change Details")
        row_idx = 2
        for number, change_list in changes.items():
            for change in change_list:
                ws_changes.cell(row=row_idx, column=1, value=str(number))
                ws_changes.cell(row=row_idx, column=2, value=str(change))
                row_idx += 1

    # --- IC Lookup Misses sheet ---
    ic_misses = report_data.get("ic_misses", [])
    if ic_misses:
        ws_ic = wb.create_sheet("IC Lookup Misses")
        ws_ic.cell(row=1, column=1, value="Details")
        for i, miss in enumerate(ic_misses, 2):
            ws_ic.cell(row=i, column=1, value=str(miss))

    wb.save(output_path)
    wb.close()
    return output_path
