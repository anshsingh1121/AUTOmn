import os
import glob
import win32com.client

def get_latest_output():
    out_dir = r"c:\Users\Ansh Singh\OneDrive\Desktop\Automation\FCB_Incident_Tracker_Automation\output"
    files = glob.glob(os.path.join(out_dir, "*.xlsm"))
    if not files: 
        return None
    return max(files, key=os.path.getctime)

def force_update_pivots():
    file_path = get_latest_output()
    if not file_path:
        print("Could not find any .xlsm file in the output/ directory.")
        return

    print(f"Opening {os.path.basename(file_path)} in Excel...")
    
    # Force Excel to be visible so you can see exactly what it is doing
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True  
        excel.DisplayAlerts = True # This will show any hidden errors Excel was suppressing
        
        wb = excel.Workbooks.Open(os.path.abspath(file_path))
        
        print("Refreshing all Pivot Caches...")
        for pc in wb.PivotCaches():
            pc.BackgroundQuery = False
            pc.MissingItemsLimit = 0
            try:
                pc.Refresh()
                print(" - Cache Refreshed")
            except Exception as e:
                print(f" - Error refreshing cache: {e}")
            
        print("Clearing filters and updating Pivot Tables...")
        for ws in wb.Worksheets:
            for pt in ws.PivotTables():
                try:
                    pt.ClearAllFilters()
                    pt.Update()
                    print(f" - Pivot Table '{pt.Name}' updated on sheet '{ws.Name}'")
                except Exception as e:
                    print(f" - Error updating pivot table: {e}")
                
        print("Recalculating formulas...")
        try:
            excel.CalculateFullRebuild()
        except:
            pass
        
        print("\n========================================================")
        print("DONE! Please look at the Excel window.")
        print("If the Pivot Table updated successfully, you can manually save and close.")
        print("========================================================")
        
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    force_update_pivots()
