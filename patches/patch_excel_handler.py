# =====================================================================
# PATCH FOR: excel_handler.py
# 1. At the top of excel_handler.py, ensure datetime is imported:
#    from datetime import datetime
# 2. At the bottom of excel_handler.py, replace the save_output function
#    with the following functions:
# =====================================================================
from datetime import datetime
import os
import sys

def generate_report_filename(
    start_date: str | None = None,
    end_date: str | None = None,
    change_numbers: list[str] | None = None,
    prefix: str = "QAR_Audit_Report"
) -> str:
    """Generate a standardized corporate audit filename with professional date-time frame nomenclature.

    Nomenclature Patterns:
      - Single CR:  QAR_Audit_Report_CHG6892954_Run_YYYYMMDD_HHMMSS.xlsx
      - Batch CRs:  QAR_Audit_Report_Batch_12CRs_Run_YYYYMMDD_HHMMSS.xlsx
      - Date Range: QAR_Audit_Report_DateRange_YYYYMMDD_to_YYYYMMDD_Run_YYYYMMDD_HHMMSS.xlsx
      - Fallback:   QAR_Audit_Report_Run_YYYYMMDD_HHMMSS.xlsx
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if start_date and end_date:
        s_clean = start_date.replace("-", "").replace("/", "")
        e_clean = end_date.replace("-", "").replace("/", "")
        return f"{prefix}_DateRange_{s_clean}_to_{e_clean}_Run_{timestamp}.xlsx"
    elif change_numbers:
        valid_crs = [str(c).strip().upper() for c in change_numbers if str(c).strip()]
        if len(valid_crs) == 1:
            return f"{prefix}_{valid_crs[0]}_Run_{timestamp}.xlsx"
        elif len(valid_crs) > 1:
            return f"{prefix}_Batch_{len(valid_crs)}CRs_Run_{timestamp}.xlsx"
    return f"{prefix}_Run_{timestamp}.xlsx"


def get_default_report_dir() -> str:
    """Resolve report output directory to ensure it is created next to the executable when packaged."""
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.getcwd()
    report_dir = os.path.join(base_dir, "report")
    os.makedirs(report_dir, exist_ok=True)
    return report_dir


def save_output(workbook, output_path: str | None = None) -> str:
    """Save the populated workbook with professional nomenclature and safe directory resolution."""
    report_dir = get_default_report_dir()

    if output_path:
        if os.path.isabs(output_path):
            path = output_path
        else:
            base_name = os.path.basename(output_path)
            path = os.path.join(report_dir, base_name)
    else:
        filename = generate_report_filename()
        path = os.path.join(report_dir, filename)

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    print(f"Saving audit report to: {path}")
    workbook.save(path)
    return path
