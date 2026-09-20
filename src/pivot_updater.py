"""
Pivot Updater
Directly writes calculated summary tables and a chart into the 'Pivots' sheet
using Python (pandas) for calculations and COM for cell writing.

This completely bypasses Excel's native Pivot Table engine, which is blocked
by corporate security policies (Trust Center yellow banner).

Layout (matches the user's original Pivots sheet screenshot exactly):

    Table 1 — Priority × Month (all months)
        Row 3:  "Count of Number"  |  "Column Labels"
        Row 4:  "Row Labels"       |  1-Critical  |  2-High  |  …  |  Grand Total
        Row 5+: Jan … Sep …
        Last:   Grand Total

    Table 2 — Region × Month (last 3 months only)
        Starts 8 rows below Table 1's Grand Total
        Same header structure, columns = India / US / Grand Total

    Chart — "IMT Region wise Incident Volume"
        Stacked Column chart sourced from Table 2's data range
"""

import pandas as pd
from typing import Any, Dict, List, Optional

# Chronological 3-letter month abbreviations
MONTHS_ORDERED = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']


def update_pivots_sheet(wb, df: pd.DataFrame):
    """
    Calculate summary tables from the data and write them directly
    into the 'Pivots' sheet, replicating the exact layout of the
    original Pivot Tables.

    Args:
        wb: COM Workbook object (already open via win32com)
        df: The output DataFrame with all incident records
    """
    # --- Get or create the Pivots sheet ---
    ws = _get_or_create_sheet(wb, "Pivots")

    # --- Clear everything: cells, shapes (charts), pivot tables ---
    _clear_sheet(ws)

    # --- Prepare the data ---
    df_calc = df.copy()
    df_calc['Created'] = pd.to_datetime(df_calc['Created'], errors='coerce')
    df_calc = df_calc.dropna(subset=['Created'])
    df_calc = df_calc[df_calc['Number'].notna() & (df_calc['Number'] != '')]

    df_calc['Month'] = df_calc['Created'].dt.strftime('%b')
    df_calc['Month'] = pd.Categorical(
        df_calc['Month'], categories=MONTHS_ORDERED, ordered=True
    )

    # --- TABLE 1: Priority × Month (all months with data) ---
    t1_end_row = _write_priority_table(ws, df_calc)

    # --- TABLE 2: Region × Last 3 Months ---
    t2_start_row = t1_end_row + 8   # 8-row gap (matches screenshot spacing)
    t2_data_info = _write_region_table(ws, df_calc, t2_start_row)

    # --- CHART: IMT Region wise Incident Volume ---
    if t2_data_info:
        _create_region_chart(ws, t2_data_info)

    # --- Auto-fit columns for clean display ---
    try:
        ws.Columns.AutoFit()
    except Exception:
        pass

    print("  Pivots sheet updated with Python-calculated summaries and chart.")


# ============================================================
# INTERNAL HELPERS
# ============================================================

def _get_or_create_sheet(wb, name: str):
    """Return existing sheet or create a new one."""
    try:
        return wb.Worksheets(name)
    except Exception:
        ws = wb.Worksheets.Add()
        ws.Name = name
        return ws


def _clear_sheet(ws):
    """Remove all content, shapes, and pivot tables from the sheet."""
    try:
        ws.Cells.Clear()
    except Exception:
        pass

    # Delete shapes in reverse to avoid index shifts
    try:
        for i in range(ws.Shapes.Count, 0, -1):
            ws.Shapes(i).Delete()
    except Exception:
        pass

    # Delete any surviving pivot tables
    try:
        for i in range(ws.PivotTables().Count, 0, -1):
            ws.PivotTables()(i).TableRange2.Clear()
    except Exception:
        pass


def _write_priority_table(ws, df_calc: pd.DataFrame) -> int:
    """
    Write Table 1: Count of incidents by Month × Priority.
    Starts at row 3, matching the screenshot exactly.

    Returns the row number of the Grand Total row.
    """
    ct = pd.crosstab(df_calc['Month'], df_calc['Priority'], dropna=False)
    # Remove months that have zero incidents
    ct = ct.loc[(ct != 0).any(axis=1)]

    if ct.empty:
        ws.Cells(3, 1).Value = "Count of Number"
        ws.Cells(3, 2).Value = "Column Labels"
        ws.Cells(4, 1).Value = "No data"
        return 5

    # Sort priority columns naturally (1-Critical < 2-High < 3-Medium …)
    ct = ct.reindex(sorted(ct.columns), axis=1)

    # Grand Total column and row
    ct['Grand Total'] = ct.sum(axis=1)
    gt_series = ct.sum(axis=0)
    gt_series.name = 'Grand Total'
    ct = pd.concat([ct, gt_series.to_frame().T])

    priority_cols = list(ct.columns)
    num_data_cols = len(priority_cols)

    # --- Row 3: top header ---
    ws.Cells(3, 1).Value = "Count of Number"
    ws.Cells(3, 2).Value = "Column Labels"

    # --- Row 4: sub-headers (priority names) ---
    ws.Cells(4, 1).Value = "Row Labels"
    for i, col_name in enumerate(priority_cols):
        ws.Cells(4, 2 + i).Value = str(col_name)

    # --- Data rows (row 5 onward) ---
    row_idx = 5
    for month_label, row_data in ct.iterrows():
        ws.Cells(row_idx, 1).Value = str(month_label)
        for i, col_name in enumerate(priority_cols):
            val = int(row_data[col_name])
            # Zeros → blank (matches original pivot behaviour)
            ws.Cells(row_idx, 2 + i).Value = val if val != 0 else ""
        row_idx += 1

    last_row = row_idx - 1          # Grand Total row
    last_col = 1 + num_data_cols    # Last column

    # --- Formatting ---
    _format_header_rows(ws, 3, 4, last_col)
    _format_grand_total_row(ws, last_row, last_col)

    return last_row


def _write_region_table(ws, df_calc: pd.DataFrame, start_row: int):
    """
    Write Table 2: Count of incidents by Month × Region (last 3 months only).

    Returns a dict describing the table coordinates (for chart creation),
    or None if no region data is available.
    """
    region_col = 'Region' if 'Region' in df_calc.columns else None
    if region_col is None:
        return None

    # Only rows with a non-blank Region
    df_region = df_calc[df_calc[region_col].notna() & (df_calc[region_col] != '')]
    if df_region.empty:
        return None

    ct = pd.crosstab(df_region['Month'], df_region[region_col], dropna=False)
    has_data = (ct != 0).any(axis=1)
    if not has_data.any():
        return None

    # Find the latest month that has data (e.g. 'Jun')
    last_valid_month = has_data[::-1].idxmax()
    end_idx = MONTHS_ORDERED.index(last_valid_month)
    # Grab EXACTLY the last 3 months leading up to it (e.g. Apr, May, Jun)
    start_idx = max(0, end_idx - 2)
    ct = ct.iloc[start_idx : end_idx + 1]

    # Sort region columns alphabetically (India < US)
    ct = ct.reindex(sorted(ct.columns), axis=1)

    # Grand Total column + row
    ct['Grand Total'] = ct.sum(axis=1)
    gt_series = ct.sum(axis=0)
    gt_series.name = 'Grand Total'
    ct = pd.concat([ct, gt_series.to_frame().T])

    region_cols = list(ct.columns)
    num_data_cols = len(region_cols)

    # --- Headers ---
    ws.Cells(start_row, 1).Value = "Count of Number"
    ws.Cells(start_row, 2).Value = "Column Labels"

    ws.Cells(start_row + 1, 1).Value = "Row Labels"
    for i, col_name in enumerate(region_cols):
        ws.Cells(start_row + 1, 2 + i).Value = str(col_name)

    # --- Data rows ---
    row_idx = start_row + 2
    data_start_row = row_idx
    for month_label, row_data in ct.iterrows():
        ws.Cells(row_idx, 1).Value = str(month_label)
        for i, col_name in enumerate(region_cols):
            val = int(row_data[col_name])
            ws.Cells(row_idx, 2 + i).Value = val if val != 0 else ""
        row_idx += 1

    last_row = row_idx - 1
    last_col = 1 + num_data_cols

    # --- Formatting ---
    _format_header_rows(ws, start_row, start_row + 1, last_col)
    _format_grand_total_row(ws, last_row, last_col)

    return {
        'header_row': start_row + 1,       # Row with "India" / "US" headers
        'data_start_row': data_start_row,   # First month data row
        'data_end_row': last_row - 1,       # Last month data row (before Grand Total)
        'last_col': last_col - 1,           # Last region column (before Grand Total)
        'start_row': start_row,             # Table 2 top-left
    }


def _create_region_chart(ws, data_info: dict):
    """
    Create a Stacked Column Chart for Region data,
    matching the screenshot's "IMT Region wise Incident Volume" chart.
    """
    try:
        header_row = data_info['header_row']
        data_end   = data_info['data_end_row']
        last_col   = data_info['last_col']
        start_row  = data_info['start_row']

        # Source range: header row → last data row, columns A → last region column
        # (excludes Grand Total row and Grand Total column)
        src_range = ws.Range(
            ws.Cells(header_row, 1),
            ws.Cells(data_end, last_col)
        )

        # xlColumnStacked = 52, Style 201
        chart_obj = ws.Shapes.AddChart2(201, 52)
        chart = chart_obj.Chart
        chart.SetSourceData(src_range)
        chart.HasTitle = True
        chart.ChartTitle.Text = "IMT Region wise Incident Volume"

        # Add data labels to each bar segment (matches screenshot)
        try:
            for i in range(1, chart.SeriesCollection().Count + 1):
                series = chart.SeriesCollection(i)
                series.HasDataLabels = True
                series.DataLabels().ShowValue = True
        except Exception:
            pass

        # Position the chart next to Table 2
        chart_obj.Left = ws.Cells(start_row, 5).Left
        chart_obj.Top  = ws.Cells(start_row, 5).Top
        chart_obj.Width  = 400
        chart_obj.Height = 250

    except Exception as e:
        print(f"  Warning: Could not create chart: {e}")


# ============================================================
# FORMATTING HELPERS
# ============================================================

def _format_header_rows(ws, row1: int, row2: int, last_col: int):
    """Apply light-blue header formatting to two header rows."""
    try:
        rng = ws.Range(ws.Cells(row1, 1), ws.Cells(row2, last_col))
        rng.Interior.ThemeColor = 5            # xlThemeColorAccent1
        rng.Interior.TintAndShade = 0.6        # Light tint → light blue
        rng.Font.Bold = True
    except Exception:
        pass


def _format_grand_total_row(ws, row: int, last_col: int):
    """Apply light-blue bold formatting to the Grand Total row."""
    try:
        rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, last_col))
        rng.Interior.ThemeColor = 5
        rng.Interior.TintAndShade = 0.6
        rng.Font.Bold = True
    except Exception:
        pass
