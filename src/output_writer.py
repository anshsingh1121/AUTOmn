"""
Output Writer
Writes the reconciled data to Excel using pure native Windows COM.
By using win32com instead of openpyxl, we completely bypass openpyxl's infamous 
bug of silently stripping Pivot Charts and breaking Pivot Caches upon save.
"""

import os
import shutil
from datetime import datetime
from typing import Dict, Any, List, Optional
import pandas as pd

class OutputWriteError(Exception):
    """Exception raised for errors during output writing."""
    pass

def create_output_path(output_dir: str, prefix: str = "Master_Updated") -> str:
    """Create a timestamped output file path inside a Year/Month folder structure."""
    now = datetime.now()
    year_folder = now.strftime("%Y")
    month_folder = now.strftime("%m_%B")
    
    target_dir = os.path.join(output_dir, year_folder, month_folder)
    os.makedirs(target_dir, exist_ok=True)
    
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    return os.path.join(target_dir, f"{prefix}_{timestamp}.xlsx")

def create_archive_copy(source_path: str, archive_dir: str) -> str:
    """Create a backup copy of the original master tracker inside a Year/Month folder."""
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

def write_output_workbook(
    output_df: pd.DataFrame,
    output_path: str,
    source_workbook_path: Optional[str] = None,
    target_sheets: Optional[List[str]] = None,
    formula_columns: Optional[Dict[str, str]] = None
) -> str:
    """
    Write the output DataFrame to an Excel workbook using native COM.
    This guarantees Pivot Tables, Data Models, and Charts are never destroyed.
    """
    if not source_workbook_path:
        raise OutputWriteError("A source workbook template is absolutely required for COM writing.")
        
    target_sheets = target_sheets or ["FCB"]
    
    # Clean up DataFrame for COM transfer
    df_write = output_df.copy()
    for col in df_write.columns:
        if pd.api.types.is_datetime64_any_dtype(df_write[col]):
            # Convert to pure datetime objects, replace NaT with empty string
            df_write[col] = df_write[col].apply(lambda x: "" if pd.isna(x) else x.to_pydatetime())
    df_write = df_write.fillna("")
    
    # 1. Copy the pristine template directly to the output path
    try:
        shutil.copy2(source_workbook_path, output_path)
    except Exception as e:
        raise OutputWriteError(f"Cannot copy template to output path: {e}")
        
    abs_output = os.path.abspath(output_path)
    
    # 2. Use win32com to natively open, edit, and refresh the file
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        
        try:
            wb = excel.Workbooks.Open(abs_output)
            
            for sheet_name in target_sheets:
                try:
                    ws = wb.Worksheets(sheet_name)
                except Exception:
                    continue # Sheet doesn't exist, skip
                
                # Determine boundaries
                last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column # xlToLeft
                if last_col < 1:
                    continue
                    
                # Read headers
                headers = []
                for c in range(1, last_col + 1):
                    val = ws.Cells(1, c).Value
                    headers.append(str(val).strip() if val else "")
                
                # Detect existing formulas in row 2
                existing_formulas = {}
                last_row = ws.Cells(ws.Rows.Count, "A").End(-4162).Row # xlUp
                if last_row >= 2:
                    for c in range(1, last_col + 1):
                        cell = ws.Cells(2, c)
                        if cell.HasFormula:
                            header = headers[c-1]
                            # Let Python calculation override if it's Duration
                            if header != "Duration":
                                existing_formulas[c] = cell.FormulaR1C1
                
                # Override with explicit formula config if any
                if formula_columns:
                    for c_idx, h in enumerate(headers, start=1):
                        if h in formula_columns and h != "Duration":
                            existing_formulas[c_idx] = formula_columns[h]
                
                # Clear existing data rows (from row 2 down)
                if last_row >= 2:
                    ws.Range(ws.Cells(2, 1), ws.Cells(last_row, last_col)).ClearContents()
                
                # Build a 2D array of the new data
                num_records = len(df_write)
                if num_records == 0:
                    continue
                    
                write_data = []
                for _, row in df_write.iterrows():
                    row_data = []
                    for c_idx, h in enumerate(headers, start=1):
                        if c_idx in existing_formulas:
                            row_data.append(existing_formulas[c_idx])
                        elif h in df_write.columns:
                            row_data.append(row[h])
                        else:
                            row_data.append("")
                    write_data.append(row_data)
                
                # Write massive 2D array to Excel via COM (extremely fast)
                start_cell = ws.Cells(2, 1)
                end_cell = ws.Cells(1 + num_records, last_col)
                ws.Range(start_cell, end_cell).FormulaR1C1 = write_data
                
                # Expand any ListObjects (Excel Tables) safely
                for obj in ws.ListObjects:
                    try:
                        new_range = ws.Range(ws.Cells(1, 1), ws.Cells(1 + num_records, last_col))
                        obj.Resize(new_range)
                    except Exception:
                        pass
            
            # Step 3: Expand PivotCaches dynamically
            for pc in wb.PivotCaches():
                try:
                    src = pc.SourceData
                    if isinstance(src, str) and "!" in src:
                        p_sheet = src.split("!")[0].replace("'", "")
                        p_ws = wb.Worksheets(p_sheet)
                        p_last_row = p_ws.Cells(p_ws.Rows.Count, "A").End(-4162).Row
                        p_last_col = p_ws.Cells(1, p_ws.Columns.Count).End(-4159).Column
                        if p_last_row > 1 and p_last_col > 0:
                            new_src = f"'{p_sheet}'!R1C1:R{p_last_row}C{p_last_col}"
                            pc.SourceData = new_src
                except Exception:
                    pass
            
            # Refresh everything
            wb.RefreshAll()
            excel.CalculateUntilAsyncQueriesDone()
            wb.Save()
            wb.Close(True)
            
        finally:
            excel.Quit()
            del excel
            
    except ImportError:
        raise OutputWriteError("win32com is required to write pivot-safe workbooks. Please pip install pywin32.")
    except Exception as e:
        raise OutputWriteError(f"COM output write failed: {e}")
        
    return output_path
