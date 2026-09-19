import pandas as pd
import win32com.client

def build_static_dashboard(wb, df: pd.DataFrame):
    """
    Completely abandons Excel's Pivot Table engine.
    Calculates the exact incident volumes natively in Python (Pandas),
    and writes them directly into the 'Pivots' sheet as static, beautifully formatted tables.
    Also draws a native Excel chart linked to the static Python data.
    """
    try:
        ws = wb.Worksheets("Pivots")
    except Exception:
        ws = wb.Worksheets.Add()
        ws.Name = "Pivots"
        
    # 1. Purge everything on the sheet (broken pivots and old charts)
    try:
        ws.Cells.Clear()
        for shape in ws.Shapes:
            shape.Delete()
    except Exception:
        pass

    # 2. Process Data natively in Python
    df_calc = df.copy()
    # Force pure datetimes, bypassing the string conversion in the write loop
    df_calc['Created'] = pd.to_datetime(df_calc['Created'], errors='coerce')
    df_calc = df_calc.dropna(subset=['Created', 'Number'])
    
    months_ordered = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    df_calc['Month'] = pd.Categorical(df_calc['Created'].dt.strftime('%b'), categories=months_ordered, ordered=True)
    
    # ---------------------------------------------------------
    # TABLE 1: Priority Summary
    # ---------------------------------------------------------
    pt1 = pd.crosstab(df_calc['Month'], df_calc['Priority'], dropna=False)
    # Drop rows that are completely zero
    pt1 = pt1.loc[(pt1 != 0).any(axis=1)]
    if not pt1.empty:
        pt1.loc['Grand Total'] = pt1.sum(axis=0)
        pt1['Grand Total'] = pt1.sum(axis=1)

    # Write Table 1 to A3
    ws.Cells(3, 1).Value = "Count of Number"
    ws.Cells(3, 2).Value = "Column Labels"
    
    col_idx = 2
    for col_name in pt1.columns:
        ws.Cells(4, col_idx).Value = str(col_name)
        col_idx += 1
        
    row_idx = 5
    for idx, row in pt1.iterrows():
        ws.Cells(row_idx, 1).Value = str(idx)
        c_idx = 2
        for val in row:
            ws.Cells(row_idx, c_idx).Value = int(val)
            c_idx += 1
        row_idx += 1
        
    # Format Table 1
    t1_end_row = row_idx - 1
    t1_end_col = col_idx - 1
    # Light blue header background: ColorIndex 37
    try:
        ws.Range(ws.Cells(3, 1), ws.Cells(4, t1_end_col)).Interior.ThemeColor = 5 # xlThemeColorAccent1
        ws.Range(ws.Cells(3, 1), ws.Cells(4, t1_end_col)).Interior.TintAndShade = 0.6
        ws.Range(ws.Cells(3, 1), ws.Cells(4, t1_end_col)).Font.Bold = True
        
        ws.Range(ws.Cells(t1_end_row, 1), ws.Cells(t1_end_row, t1_end_col)).Interior.ThemeColor = 5
        ws.Range(ws.Cells(t1_end_row, 1), ws.Cells(t1_end_row, t1_end_col)).Interior.TintAndShade = 0.6
        ws.Range(ws.Cells(t1_end_row, 1), ws.Cells(t1_end_row, t1_end_col)).Font.Bold = True
    except:
        pass

    # ---------------------------------------------------------
    # TABLE 2: Region Summary
    # ---------------------------------------------------------
    # In earlier configs this was IC or Region. Checking df_calc
    region_col = 'Region' if 'Region' in df_calc.columns else 'IC'
    
    pt2 = pd.crosstab(df_calc['Month'], df_calc[region_col], dropna=False)
    # Only keep last 3 months for the chart? The original pivot had Jun, Jul, Aug.
    # We will keep all non-zero months
    pt2 = pt2.loc[(pt2 != 0).any(axis=1)]
    if not pt2.empty:
        pt2.loc['Grand Total'] = pt2.sum(axis=0)
        pt2['Grand Total'] = pt2.sum(axis=1)

    t2_start_row = t1_end_row + 8
    ws.Cells(t2_start_row, 1).Value = "Count of Number"
    ws.Cells(t2_start_row, 2).Value = "Column Labels"
    
    col_idx = 2
    for col_name in pt2.columns:
        if col_name != "Grand Total":
            ws.Cells(t2_start_row + 1, col_idx).Value = str(col_name)
        else:
            ws.Cells(t2_start_row + 1, col_idx).Value = "Grand Total"
        col_idx += 1
        
    row_idx = t2_start_row + 2
    for idx, row in pt2.iterrows():
        ws.Cells(row_idx, 1).Value = str(idx)
        c_idx = 2
        for val in row:
            ws.Cells(row_idx, c_idx).Value = int(val)
            c_idx += 1
        row_idx += 1

    t2_end_row = row_idx - 1
    t2_end_col = col_idx - 1
    
    try:
        ws.Range(ws.Cells(t2_start_row, 1), ws.Cells(t2_start_row + 1, t2_end_col)).Interior.ThemeColor = 5 
        ws.Range(ws.Cells(t2_start_row, 1), ws.Cells(t2_start_row + 1, t2_end_col)).Interior.TintAndShade = 0.6
        ws.Range(ws.Cells(t2_start_row, 1), ws.Cells(t2_start_row + 1, t2_end_col)).Font.Bold = True
        
        ws.Range(ws.Cells(t2_end_row, 1), ws.Cells(t2_end_row, t2_end_col)).Interior.ThemeColor = 5
        ws.Range(ws.Cells(t2_end_row, 1), ws.Cells(t2_end_row, t2_end_col)).Interior.TintAndShade = 0.6
        ws.Range(ws.Cells(t2_end_row, 1), ws.Cells(t2_end_row, t2_end_col)).Font.Bold = True
    except:
        pass

    # ---------------------------------------------------------
    # CHART: IMT Region wise Incident Volume
    # ---------------------------------------------------------
    try:
        # Create a Stacked Column Chart
        # 52 = xlColumnStacked
        chart_obj = ws.Shapes.AddChart2(201, 52)
        chart = chart_obj.Chart
        chart.HasTitle = True
        chart.ChartTitle.Text = "IMT Region wise Incident Volume"
        
        # Source data: Exclude the Grand Total row and column
        chart_src_range = ws.Range(
            ws.Cells(t2_start_row + 1, 1), 
            ws.Cells(t2_end_row - 1, t2_end_col - 1)
        )
        chart.SetSourceData(chart_src_range)
        
        # Position the chart right next to Table 2 (Column E)
        chart_obj.Left = ws.Cells(t2_start_row, 5).Left
        chart_obj.Top = ws.Cells(t2_start_row, 5).Top
        chart_obj.Width = 400
        chart_obj.Height = 250
    except Exception as e:
        print(f"  Warning: Could not build static chart: {e}")
