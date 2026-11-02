# Codebase Modifications & Packaging Guide

This guide details the **minimal codebase updates** applied to achieve:
1. **Interactive Authentication on Startup:** Prompts the user for ServiceNow Instance URL, Username, and Password (masked) as soon as the packaged executable is launched.
2. **Professional Date-Time Frame Nomenclature:** Generates unique, professional, timestamped audit report filenames (e.g., `QAR_Audit_Report_CHG6892954_Run_20261102_221350.xlsx` or `QAR_Audit_Report_DateRange_20261001_to_20261031_Run_20261102_221350.xlsx`), preventing accidental overwriting and standardizing audit records.
3. **Robust Windows `--onefile` Packaging:** Standalone, rock-solid compilation via updated `.spec` and `.bat` files.

---

## 1. Modifications in `constants.py`

### Location:
Around lines 113–127 in [`constants.py`](file:///c:/Users/Ansh%20Singh/Downloads/CHANGE_MANAGEMENT_AUDIT_AUTOMATION_RECONSTRUCTED/constants.py).

### Change Summary:
- Allow non-empty string overrides for credentials.
- Add `has_credentials()` helper function to quickly verify if credentials are set in memory.

### Code Diff:

```python
# ==========================================
# BEFORE:
# ==========================================
def update_credentials(username: str, password: str, instance_url: str | None = None, verify_ssl: bool | str | None = None):
    """Dynamically update runtime credentials in memory."""
    global USERNAME, PASSWORD, INSTANCE_URL, BASE_URL, ATT_FILE_URL, VERIFY_SSL
    if username:
        USERNAME = username
    if password:
        PASSWORD = password
    if instance_url:
        INSTANCE_URL = instance_url.rstrip("/")
        BASE_URL = f"{INSTANCE_URL}/api/now/table/"
        ATT_FILE_URL = f"{INSTANCE_URL}/api/now/attachment/"
    if verify_ssl is not None:
        VERIFY_SSL = verify_ssl

# ==========================================
# AFTER:
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
# ==========================================
# BEFORE:
# ==========================================
from __future__ import annotations
import os
import sys
import openpyxl
from openpyxl.styles import PatternFill
import constants

...

def save_output(workbook, output_path):
    os.makedirs("report", exist_ok=True)
    path = output_path if os.path.isabs(output_path) else os.path.join("report", output_path)
    print(f"Saving output to: {path}")
    workbook.save(path)
    return path

# ==========================================
# AFTER:
# ==========================================
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
1. Lines 18–30: Add `safe_getpass()` before `pause_for_user()`.
2. Lines 104–110: Update `report_filename()`.
3. Lines 170–180: Update `create_change_wise_report()`.
4. Lines 285–300: Update `run_date_range_audit()`.
5. Lines 335–455: Update `ensure_credentials()` and `main()`.

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

#### Part B: `report_filename()`
```python
# BEFORE:
def report_filename(month, change_wise=False):
    year = datetime.today().strftime("%y")
    if change_wise:
        return f"FCB {month}-QAR Change Wise.xlsx"
    return f"FCB {month}'{year} - QAR Report.xlsx"

# AFTER:
def report_filename(month=None, change_wise=False, start_date=None, end_date=None, change_numbers=None):
    """Generate professional corporate report filename with date-time frame nomenclature."""
    return excel_handler.generate_report_filename(
        start_date=start_date,
        end_date=end_date,
        change_numbers=change_numbers
    )
```

#### Part C: `create_change_wise_report()`
```python
# BEFORE:
    filename = report_filename(month, change_wise=True)
    existing_report = os.path.join("report", filename)
    template_path = existing_report if os.path.exists(existing_report) else resolve_template_path(args.template)

# AFTER:
    change_numbers = [cr.get("number") for cr in all_crs if cr.get("number")]
    filename = args.output or excel_handler.generate_report_filename(change_numbers=change_numbers)
    template_path = resolve_template_path(args.template)
```

#### Part D: `run_date_range_audit()`
```python
# BEFORE:
    filename = report_filename(month)
    output_path = args.output or filename
    existing_report = os.path.join("report", filename)
    template_path = existing_report if os.path.exists(existing_report) else resolve_template_path(args.template)

# AFTER:
    filename = args.output or excel_handler.generate_report_filename(start_date=start_date, end_date=end_date)
    template_path = resolve_template_path(args.template)
```

#### Part E: `ensure_credentials()` and `main()`
```python
def ensure_credentials(is_interactive: bool = True) -> bool:
    """Ensure valid ServiceNow credentials exist; in interactive mode, prompt user at startup."""
    if not is_interactive:
        if constants.has_credentials():
            return True
        print("\n[ERROR] Missing ServiceNow credentials.")
        print("Please configure SN_USERNAME and SN_PASSWORD in your environment or in a local .env file.")
        return False

    print("\n" + "=" * 65)
    print("      ServiceNow Corporate Authentication")
    print("=" * 65)
    print(f"Target Instance URL: {constants.INSTANCE_URL}")
    try:
        new_url = input("Enter Instance URL (or press Enter to keep default): ").strip()
    except (KeyboardInterrupt, EOFError):
        return False
    if new_url:
        constants.update_credentials("", "", instance_url=new_url)

    default_user = constants.USERNAME
    user_prompt = f"Enter ServiceNow Username [{default_user}]: " if default_user else "Enter ServiceNow Username: "
    try:
        username = input(user_prompt).strip()
    except (KeyboardInterrupt, EOFError):
        return False
    if not username and default_user:
        username = default_user
    if not username:
        print("[ERROR] Username cannot be empty.")
        return False

    password = safe_getpass("Enter ServiceNow Password (masked): ")
    if not password:
        print("[ERROR] Password cannot be empty.")
        return False

    constants.update_credentials(username=username, password=password)
    print(f"[OK] Authenticated user '{username}' for this audit session.")

    try:
        save_env = input("Save credentials to local .env for future automatic runs? [y/N]: ").strip().lower()
        if save_env in {"y", "yes"}:
            target_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.getcwd()
            env_path = os.path.join(target_dir, ".env")
            with open(env_path, "w", encoding="utf-8") as f:
                f.write(f"# CMAA Enterprise Configuration\n")
                f.write(f"SN_INSTANCE_URL={constants.INSTANCE_URL}\n")
                f.write(f"SN_USERNAME={username}\n")
                f.write(f"SN_PASSWORD={password}\n")
                f.write(f"SN_VERIFY_SSL={str(constants.VERIFY_SSL).lower()}\n")
            print(f"[OK] Saved credentials to: {env_path}")
    except (KeyboardInterrupt, EOFError):
        pass
    except Exception as exc:
        print(f"[WARNING] Could not save .env: {exc}")

    return True


def main():
    is_interactive = (len(sys.argv) == 1)

    if is_interactive:
        print("=" * 65)
        print("      Change Management Audit Automation (CMAA) - Standalone")
        print("=" * 65)
        if not ensure_credentials(is_interactive=True):
            pause_for_user("\nPress Enter to exit...")
            sys.exit(0)
        args = interactive_prompt()
        if not args:
            pause_for_user("\nPress Enter to exit...")
            sys.exit(0)
    else:
        args = parse_arguments()
        if not ensure_credentials(is_interactive=False):
            sys.exit(1)
```

---

## 4. Packaging Files Updated

1. **[`CMAA.spec`](file:///c:/Users/Ansh%20Singh/Downloads/CHANGE_MANAGEMENT_AUDIT_AUTOMATION_RECONSTRUCTED/CMAA.spec)**
   - Added `'getpass'` to `hiddenimports`.
   - Set `upx=False` to prevent corporate antivirus false positives and avoid decompression latency.
2. **[`build_cmaa_exe.bat`](file:///c:/Users/Ansh%20Singh/Downloads/CHANGE_MANAGEMENT_AUDIT_AUTOMATION_RECONSTRUCTED/build_cmaa_exe.bat)**
   - Copies `changes.txt` and `.env.example` into `dist\` alongside `CMAA_Audit.exe`.
   - Performs automated 5-step build with clean exit.
