# Codebase Modifications & Packaging Guide

This guide details the **minimal codebase updates** applied to achieve:
1. **Interactive Authentication on Startup:** Prompts the user for ServiceNow Instance URL, Username, and Password (masked) as soon as the packaged executable is launched.
2. **Professional Date-Time Frame Nomenclature:** Generates unique, professional, timestamped audit report filenames (e.g., `QAR_Audit_Report_CHG6892954_Run_20261102_221350.xlsx` or `QAR_Audit_Report_DateRange_20261001_to_20261031_Run_20261102_221350.xlsx`), preventing accidental overwriting and standardizing audit records.
3. **Enterprise Input Safety Constraints:**
   - **Change Requests Limit:** Strictly enforces a maximum of **100 Change Requests** per audit execution (from console or `.txt` input).
   - **Date Range Limit:** Strictly enforces that audit date ranges **do not exceed 1 month** (maximum 31 days).
4. **Robust Windows `--onefile` Packaging:** Standalone, rock-solid compilation via updated `.spec` and `.bat` files.

---

## 1. Modifications in `constants.py`

### Location:
Near line 65 and lines 115–130 in [`constants.py`](file:///c:/Users/Ansh%20Singh/Downloads/CHANGE_MANAGEMENT_AUDIT_AUTOMATION_RECONSTRUCTED/constants.py).

### Change Summary:
- Added `MAX_CHANGE_NUMBERS = 100` and `MAX_DATE_RANGE_DAYS = 31`.
- Enhanced `update_credentials()` to support non-empty string overrides.
- Added `has_credentials() -> bool` helper.

### Code Diff:

```python
# ==========================================
# PART A: Input Constraints (Near line 65)
# ==========================================
CHANGE_NUMBER_PATTERN = re.compile(r"CHG\d+", re.IGNORECASE)

# Corporate Audit Input Constraints
MAX_CHANGE_NUMBERS = 100      # Maximum allowed Change Requests per audit run
MAX_DATE_RANGE_DAYS = 31      # Maximum allowed date range span (1 calendar month)


# ==========================================
# PART B: Credentials (Near line 115)
# ==========================================
def update_credentials(username: str, password: str, instance_url: str | None = None, verify_ssl: bool | str | None = None):
    """Dynamically update runtime credentials in memory."""
    global USERNAME, PASSWORD, INSTANCE_URL, BASE_URL, ATT_FILE_URL, VERIFY_SSL
    if username is not None and username != "":
        USERNAME = username
    if password is not None and password != "":
        PASSWORD = password
    if instance_url:
        INSTANCE_URL = instance_url.rstrip("/")
        BASE_URL = f"{INSTANCE_URL}/api/now/table/"
        ATT_FILE_URL = f"{INSTANCE_URL}/api/now/attachment/"
    if verify_ssl is not None:
        VERIFY_SSL = verify_ssl


def has_credentials() -> bool:
    """Return True if both username and password are populated in memory."""
    return bool(USERNAME and PASSWORD)
```

---

## 2. Modifications in `excel_handler.py`

### Locations:
1. Top import block (add `from datetime import datetime`).
2. Bottom of file: replace `save_output` and add `generate_report_filename` & `get_default_report_dir`.

### Change Summary:
- Implements standardized audit report nomenclature with timestamps and scope (Single CR, Batch CRs, Date Range).
- Guarantees report directory (`report/`) is placed directly next to the executable in frozen `--onefile` mode.

### Code Diff:

```python
from __future__ import annotations
import os
import sys
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill
import constants

...

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
```

---

## 3. Modifications in `populate_cr_audit.py`

### Locations:
1. Lines 18–35: Add `safe_getpass()` before `pause_for_user()`.
2. Lines 100–125: Update `load_change_numbers()` (max 100) and `validate_date_range()` (max 31 days).
3. Lines 130–140: Update `report_filename()`.
4. Lines 180–210: Update `create_change_wise_report()` to use `excel_handler.generate_report_filename()`.
5. Lines 300–320: Update `run_date_range_audit()` to use `excel_handler.generate_report_filename()`.
6. Lines 360–480: Update `ensure_credentials()`, `interactive_prompt()`, and `main()`.

### Code Diffs:

#### Part A: `safe_getpass()`
```python
def safe_getpass(prompt: str = "Enter ServiceNow Password (masked): ") -> str:
    """Prompt for password securely using getpass on TTY or fallback to readline if piped."""
    try:
        if not sys.stdin.isatty():
            sys.stdout.write(prompt)
            sys.stdout.flush()
            line = sys.stdin.readline()
            return line.strip()
        import getpass
        return getpass.getpass(prompt).strip()
    except Exception:
        try:
            return input(prompt).strip()
        except (KeyboardInterrupt, EOFError):
            return ""
```

#### Part B: Constraints in `load_change_numbers()` and `validate_date_range()`
```python
def load_change_numbers(change_inputs):
    """Load validated Change numbers from CLI values or text files."""
    change_numbers, invalid_values = [], []
    for change_input in change_inputs or []:
        if os.path.isfile(change_input):
            try:
                with open(change_input, encoding="utf-8-sig") as change_file:
                    content = change_file.read()
            except OSError as exc:
                raise RuntimeError(f"Failed to read '{change_input}': {exc}") from exc
        else:
            content = change_input
        for raw_value in re.split(r"[\s,;]+", content):
            value = raw_value.strip().upper()
            if not value:
                continue
            if not constants.CHANGE_NUMBER_PATTERN.fullmatch(value):
                invalid_values.append(value)
                continue
            change_numbers.append(value)
    if invalid_values:
        print("Ignored invalid Change-number values:")
        for value in invalid_values:
            print(f"  {value}")
    unique_changes = list(dict.fromkeys(change_numbers))

    max_allowed = getattr(constants, "MAX_CHANGE_NUMBERS", 100)
    if len(unique_changes) > max_allowed:
        raise ValueError(
            f"Input exceeds maximum allowed limit of {max_allowed} Change Requests. "
            f"Received {len(unique_changes)} unique change numbers. "
            f"Please reduce the input to at most {max_allowed} changes."
        )
    return unique_changes


def validate_date_range(start_date, end_date):
    try:
        start_datetime = datetime.strptime(start_date, "%Y-%m-%d")
        end_datetime = datetime.strptime(end_date, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError("Dates must use YYYY-MM-DD format (e.g. 2026-10-01)") from exc
    if start_datetime > end_datetime:
        raise ValueError("--start-date must not be after --end-date.")

    max_days = getattr(constants, "MAX_DATE_RANGE_DAYS", 31)
    delta_days = (end_datetime - start_datetime).days
    if delta_days > max_days:
        raise ValueError(
            f"Date range cannot exceed 1 month (maximum {max_days} days). "
            f"Provided range spans {delta_days} days ({start_date} to {end_date}). "
            f"Please specify a date range within 1 month."
        )
    return start_datetime, end_datetime
```

#### Part C: `interactive_prompt()` with Live Input Validation Loops
```python
        if choice == "1":
            inp = input(f"\nEnter Change Number(s) or path to text file (max {constants.MAX_CHANGE_NUMBERS}): ").strip()
            if not inp:
                print("No input provided. Returning to main menu.")
                continue
            try:
                candidate_numbers = load_change_numbers(inp.split())
                if not candidate_numbers:
                    print("No valid Change Request numbers found in input. Returning to main menu.")
                    continue
            except (RuntimeError, ValueError) as err:
                print(f"\n[INPUT ERROR] {err}\n")
                continue
            return argparse.Namespace(
                change_nums=candidate_numbers,
                start_date=None,
                end_date=None,
                percentage=20,
                template="QAR Report.xlsx",
                tester="Tester",
                output=None,
            )
        elif choice == "2":
            print(f"\nEnter Date Range (maximum window: 1 month / {constants.MAX_DATE_RANGE_DAYS} days):")
            start_date = input("Enter Start Date (YYYY-MM-DD): ").strip()
            end_date = input("Enter End Date (YYYY-MM-DD): ").strip()
            if not start_date or not end_date:
                print("Start and End dates are required. Returning to main menu.")
                continue
            try:
                validate_date_range(start_date, end_date)
            except ValueError as val_err:
                print(f"\n[INPUT ERROR] {val_err}\n")
                continue
            pct_input = input("Enter sample percentage [default: 20]: ").strip()
            percentage = int(pct_input) if pct_input.isdigit() else 20
            tester = input("Enter Tester Name [default: Tester]: ").strip() or "Tester"
            return argparse.Namespace(
                change_nums=None,
                start_date=start_date,
                end_date=end_date,
                percentage=percentage,
                template="QAR Report.xlsx",
                tester=tester,
                output=None,
            )
```

---

## 4. Packaging Files (`.bat` and `.spec`)

1. **[`CMAA.spec`](file:///c:/Users/Ansh%20Singh/Downloads/CHANGE_MANAGEMENT_AUDIT_AUTOMATION_RECONSTRUCTED/CMAA.spec)**
   - Configured with `upx=False` for zero decompression delay and corporate AV heuristic safety.
   - Embeds `getpass` and all required modules.
   - Embeds master template `QAR Report.xlsx`.
2. **[`build_cmaa_exe.bat`](file:///c:/Users/Ansh%20Singh/Downloads/CHANGE_MANAGEMENT_AUDIT_AUTOMATION_RECONSTRUCTED/build_cmaa_exe.bat)**
   - Automates syntax check, clean build, and post-build verification.
   - Highlights enterprise input constraints (max 100 CRs, max 1 month window).
   - Deploys sample `changes.txt` and `.env.example` into `dist\`.
