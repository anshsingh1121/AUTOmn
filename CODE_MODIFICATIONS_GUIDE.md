# Corporate Workstation Manual Modification Guide (Copy-Paste Reference)

> **Context:** This guide is specifically prepared for modifying your corporate codebase when `git pull` is restricted on your corporate system.  
> Follow these exact steps to update your files manually using **Find & Replace** in your code editor (VS Code, Notepad++, PyCharm, etc.), then build your standalone application using `build_cmaa_exe.bat`.

---

## File 1: `constants.py`

### Step 1.1: Add Input Constraints
* **Location:** Around line 65, right after `CHANGE_NUMBER_PATTERN`.
* **Find this:**
```python
CHANGE_NUMBER_PATTERN = re.compile(r"CHG\d+", re.IGNORECASE)
```
* **Replace with:**
```python
CHANGE_NUMBER_PATTERN = re.compile(r"CHG\d+", re.IGNORECASE)

# Corporate Audit Input Constraints
MAX_CHANGE_NUMBERS = 100      # Maximum allowed Change Requests per audit run
MAX_DATE_RANGE_DAYS = 31      # Maximum allowed date range span (1 calendar month)
```

---

### Step 1.2: Update `update_credentials()` & Add `has_credentials()`
* **Location:** Around lines 115–128.
* **Find this:**
```python
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
```
* **Replace with:**
```python
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

## File 2: `excel_handler.py`

### Step 2.1: Add `datetime` import at top of file
* **Location:** Lines 1–7.
* **Find this:**
```python
from __future__ import annotations
import os
import sys
import openpyxl
from openpyxl.styles import PatternFill
import constants
```
* **Replace with:**
```python
from __future__ import annotations
import os
import sys
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill
import constants
```

---

### Step 2.2: Replace `save_output` at bottom of file
* **Location:** Very end of `excel_handler.py` (around line 137).
* **Find this:**
```python
def save_output(workbook, output_path):
    os.makedirs("report", exist_ok=True)
    path = output_path if os.path.isabs(output_path) else os.path.join("report", output_path)
    print(f"Saving output to: {path}")
    workbook.save(path)
    return path
```
* **Replace with:**
```python
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


def save_output(workbook, output_path: str | None = None, filename: str | None = None) -> str:
    """Save the populated workbook with professional nomenclature and safe directory resolution."""
    report_dir = get_default_report_dir()

    target_name = output_path or filename
    if target_name:
        if os.path.isabs(target_name):
            path = target_name
        else:
            base_name = os.path.basename(target_name)
            path = os.path.join(report_dir, base_name)
    else:
        path = os.path.join(report_dir, generate_report_filename())

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    print(f"Saving audit report to: {path}")
    workbook.save(path)
    return path
```

---

## File 3: `populate_cr_audit.py`

### Step 3.1: Add `shlex` & `api_logger` imports and `safe_getpass()`
* **Location:** Top of file and around line 18.
* **Add imports at top of file:**
```python
import os
import re
import shlex
import sys
import api_logger
```
* **Find this:**
```python
def pause_for_user(prompt: str = "\nPress Enter to exit..."):
```
* **Replace with:**
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


def pause_for_user(prompt: str = "\nPress Enter to exit..."):
```

---

### Step 3.2: Add External File Resolver & Update `load_change_numbers()`
* **Location:** Around lines 83–115 (replace original `load_change_numbers`).
* **Find this:**
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
    return list(dict.fromkeys(change_numbers))
```
* **Replace with:**
```python
def resolve_file_path(path_str: str) -> str | None:
    """Resolve an input file path whether absolute, relative, home-expanded, or next to exe."""
    if not path_str or not isinstance(path_str, str):
        return None
    raw = path_str.strip().strip("\"'").strip()
    if not raw:
        return None
    expanded = os.path.expanduser(os.path.expandvars(raw))

    # 1. Direct or absolute path check (handles paths anywhere on disk)
    if os.path.isfile(expanded):
        return os.path.abspath(expanded)
    if os.path.isfile(os.path.abspath(raw)):
        return os.path.abspath(raw)

    # 2. Check relative to executable directory (if running as compiled exe)
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        candidate = os.path.join(exe_dir, raw)
        if os.path.isfile(candidate):
            return os.path.abspath(candidate)

    # 3. Check relative to script directory
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidate = os.path.join(script_dir, raw)
        if os.path.isfile(candidate):
            return os.path.abspath(candidate)
    except Exception:
        pass

    return None


def parse_interactive_change_input(inp: str) -> list[str]:
    """Parse user input from interactive console prompt supporting raw paths with spaces, quotes, or CRs."""
    raw = inp.strip()
    if not raw:
        return []
    clean_raw = raw.strip("\"'").strip()
    # Check if the entire raw input is an existing file path (even if unquoted with spaces)
    if resolve_file_path(clean_raw):
        return [clean_raw]

    # Try shlex split to respect quoted tokens with spaces
    try:
        tokens = shlex.split(raw, posix=False)
        return [t.strip("\"'").strip() for t in tokens if t.strip("\"'").strip()]
    except Exception:
        # Fallback to whitespace/comma split
        return [t.strip("\"'").strip() for t in re.split(r"[\s,;]+", raw) if t.strip("\"'").strip()]


def load_change_numbers(change_inputs):
    """Load validated Change numbers from CLI values or text files (anywhere on disk)."""
    if not change_inputs:
        return []

    # Check if all inputs joined together represent a single unquoted path with spaces
    if len(change_inputs) > 1:
        joined_path = " ".join(change_inputs).strip().strip("\"'").strip()
        resolved_joined = resolve_file_path(joined_path)
        if resolved_joined:
            change_inputs = [resolved_joined]

    change_numbers, invalid_values = [], []
    for change_input in change_inputs:
        resolved_file = resolve_file_path(change_input)
        if resolved_file:
            print(f"[OK] Loading Change Requests from external file: {resolved_file}")
            try:
                with open(resolved_file, encoding="utf-8-sig") as change_file:
                    content = change_file.read()
            except OSError as exc:
                raise RuntimeError(f"Failed to read '{resolved_file}': {exc}") from exc
        else:
            clean_token = change_input.strip().strip("\"'").strip()
            # If the user specified a path that doesn't exist, provide an immediate clear error
            if clean_token.lower().endswith((".txt", ".csv", ".tsv", ".log", ".text")) or ("\\" in clean_token or "/" in clean_token):
                raise ValueError(f"Could not locate text file: '{clean_token}'. Please verify the path exists.")
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

---

### Step 3.3: Update `validate_date_range()` (1 Month / 31-Day Limit)
* **Location:** Around lines 110–120.
* **Find this:**
```python
def validate_date_range(start_date, end_date):
    try:
        start_datetime = datetime.strptime(start_date, "%Y-%m-%d")
        end_datetime = datetime.strptime(end_date, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError("Dates must use YYYY-MM-DD format (e.g. 2026-10-01)") from exc
    if start_datetime > end_datetime:
        raise ValueError("--start-date must not be after --end-date")
    return start_datetime, end_datetime
```
* **Replace with:**
```python
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

---

### Step 3.4: Update `report_filename()`
* **Location:** Around lines 120–126.
* **Find this:**
```python
def report_filename(month, change_wise=False):
    year = datetime.today().strftime("%y")
    if change_wise:
        return f"FCB {month}-QAR Change Wise.xlsx"
    return f"FCB {month}'{year} - QAR Report.xlsx"
```
* **Replace with:**
```python
def report_filename(month=None, change_wise=False, start_date=None, end_date=None, change_numbers=None):
    """Generate professional corporate report filename with date-time frame nomenclature."""
    return excel_handler.generate_report_filename(
        start_date=start_date,
        end_date=end_date,
        change_numbers=change_numbers
    )
```

---

### Step 3.4B: Update `save_workbook()` (User Audit Logging)
* **Location:** Around line 187 (or right before `populate_all_basic_data`).
* **Find this:**
```python
def save_workbook(workbook, output_path):
    try:
        saved_path = excel_handler.save_output(workbook, output_path)
    except Exception as exc:
        print(f"[ERROR] Failed saving output file: {exc}")
        return None
    print(f"\n[SUCCESS] Audit report created successfully:")
    print(f"  --> {os.path.abspath(saved_path)}")
    return saved_path
```
* **Replace with:**
```python
def save_workbook(workbook, output_path):
    try:
        saved_path = excel_handler.save_output(workbook, output_path)
    except Exception as exc:
        print(f"[ERROR] Failed saving output file: {exc}")
        if api_logger:
            api_logger.record_audit_trail(
                action="SAVE_FAILED",
                servicenow_user=getattr(constants, "USERNAME", ""),
                details=f"Save failed: {exc}",
                status="FAILED",
            )
        return None
    print(f"\n[SUCCESS] Audit report created successfully:")
    print(f"  --> {os.path.abspath(saved_path)}")
    if api_logger:
        api_logger.record_audit_trail(
            action="AUDIT_COMPLETED",
            servicenow_user=getattr(constants, "USERNAME", ""),
            details="Report created successfully",
            status="SUCCESS",
            output_report=saved_path,
        )
        if getattr(constants.api_logger, "_HIT_COUNTER", 0) > 0:
            print(f"API Audit Trail ({constants.api_logger._HIT_COUNTER} hits logged): {constants.api_logger.get_summary_path()}")
        print(f"User Audit Trail (Workstation & Network Trace): {api_logger.get_audit_trail_path()}")
    return saved_path
```

---

### Step 3.5: Update `create_change_wise_report()`
* **Location:** Around lines 170–176.
* **Find this:**
```python
def create_change_wise_report(args, all_crs, tester, month):
    sc_col = sc_template_col = nc_col = ec_col = exp_col = 3
    filename = report_filename(month, change_wise=True)
    existing_report = os.path.join("report", filename)
    template_path = existing_report if os.path.exists(existing_report) else resolve_template_path(args.template)
```
* **Replace with:**
```python
def create_change_wise_report(args, all_crs, tester, month):
    sc_col = sc_template_col = nc_col = ec_col = exp_col = 3
    change_numbers = [cr.get("number") for cr in all_crs if cr.get("number")]
    filename = args.output or excel_handler.generate_report_filename(change_numbers=change_numbers)
    template_path = resolve_template_path(args.template)
```
* **And at the bottom of `create_change_wise_report`:**
* **Find this:**
```python
    saved = save_workbook(workbook, args.output or filename)
```
* **Replace with:**
```python
    saved = save_workbook(workbook, filename)
```

---

### Step 3.6: Update `run_date_range_audit()`
* **Location:** Around lines 290–300.
* **Find this:**
```python
    filename = report_filename(month)
    output_path = args.output or filename
    existing_report = os.path.join("report", filename)
    template_path = existing_report if os.path.exists(existing_report) else resolve_template_path(args.template)
```
* **Replace with:**
```python
    filename = args.output or excel_handler.generate_report_filename(start_date=start_date, end_date=end_date)
    template_path = resolve_template_path(args.template)
```
* **And at the bottom of `run_date_range_audit`:**
* **Find this:**
```python
    saved = save_workbook(workbook, output_path)
```
* **Replace with:**
```python
    saved = save_workbook(workbook, filename)
```

---

### Step 3.7: Update `ensure_credentials()`, `interactive_prompt()`, and `main()`
* **Location:** From line 335 to the end of `populate_cr_audit.py`.
* **Find from `def ensure_credentials(` all the way to `if __name__ == "__main__":`:**
* **Replace that entire lower section with:**
```python
def ensure_credentials(is_interactive: bool = True) -> bool:
    """Ensure valid ServiceNow credentials exist; in interactive mode, prompt user at startup."""
    if not is_interactive:
        if constants.has_credentials():
            if api_logger:
                api_logger.record_audit_trail(
                    action="USER_AUTHENTICATED",
                    servicenow_user=constants.USERNAME,
                    details=f"Target: {constants.INSTANCE_URL} (CLI/ENV)",
                    status="SUCCESS",
                )
            return True
        print("\n[ERROR] Missing ServiceNow credentials.")
        print("Please configure SN_USERNAME and SN_PASSWORD in your environment or in a local .env file.")
        if api_logger:
            api_logger.record_audit_trail(
                action="AUTH_FAILED",
                details="Missing credentials in non-interactive mode",
                status="FAILED",
            )
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
    if api_logger:
        api_logger.record_audit_trail(
            action="USER_AUTHENTICATED",
            servicenow_user=username,
            details=f"Target: {constants.INSTANCE_URL}",
            status="SUCCESS",
        )

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


def interactive_prompt():
    """Interactive fallback when executed without CLI arguments (e.g., double-clicked executable)."""
    while True:
        print("\n" + "=" * 65)
        print("      Change Management Audit Automation (CMAA) - Standalone")
        print("=" * 65)
        print("\nSelect an Audit Execution Mode:")
        print("  [1] Audit specific Change Request(s) (e.g. CHG6892954 or changes.txt)")
        print("  [2] Audit by Date Range with Sampling (e.g. 2026-10-01 to 2026-10-31)")
        print("  [3] Exit\n")
        try:
            choice = input("Enter choice [1/2/3]: ").strip()
        except (KeyboardInterrupt, EOFError):
            return None

        if choice == "1":
            inp = input(f"\nEnter Change Number(s) or path to text file (max {constants.MAX_CHANGE_NUMBERS}): ").strip()
            if not inp:
                print("No input provided. Returning to main menu.")
                continue
            try:
                candidate_tokens = parse_interactive_change_input(inp)
                candidate_numbers = load_change_numbers(candidate_tokens)
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
        elif choice == "3":
            print("\nExiting application. Goodbye.")
            return None
        else:
            print("\n[!] Invalid selection. Please enter 1, 2, or 3.")


def main():
    is_interactive = (len(sys.argv) == 1)
    if api_logger:
        api_logger.record_audit_trail(
            action="APPLICATION_LAUNCH",
            details=f"Execution Mode: {'Interactive' if is_interactive else 'CLI'}",
            status="STARTED",
        )

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

    try:
        change_numbers = load_change_numbers(args.change_nums)
    except (RuntimeError, ValueError) as exc:
        print(f"\n[INPUT ERROR] {exc}")
        if api_logger:
            api_logger.record_audit_trail(
                action="INPUT_VALIDATION_ERROR",
                servicenow_user=getattr(constants, "USERNAME", ""),
                details=str(exc),
                status="FAILED",
            )
        if is_interactive:
            pause_for_user("\nPress Enter to exit...")
        sys.exit(1)

    try:
        if change_numbers:
            run_change_number_audit(args, change_numbers)
        else:
            if not args.start_date or not args.end_date:
                print("\n[ERROR] Please provide either --change-nums or both --start-date and --end-date.")
                if is_interactive:
                    pause_for_user("\nPress Enter to exit...")
                sys.exit(2)
            validate_date_range(args.start_date, args.end_date)
            run_date_range_audit(args, args.start_date, args.end_date)
    except ValueError as val_err:
        print(f"\n[INPUT ERROR] {val_err}")
        if api_logger:
            api_logger.record_audit_trail(
                action="INPUT_VALIDATION_ERROR",
                servicenow_user=getattr(constants, "USERNAME", ""),
                details=str(val_err),
                status="FAILED",
            )
    except Exception as exc:
        print(f"\n[ERROR] Execution halted: {exc}")
        if api_logger:
            api_logger.record_audit_trail(
                action="EXECUTION_HALTED",
                servicenow_user=getattr(constants, "USERNAME", ""),
                details=str(exc),
                status="FAILED",
            )
    finally:
        if is_interactive:
            pause_for_user("\nExecution completed. Press Enter to exit...")
```

---

## File 3B: `api_logger.py` (Workstation & User Audit Logging)

To record complete user identity (Windows logon, corporate domain, computer hostname, local IP, hardware MAC address, and timestamp) in `audit_logs/USER_AUDIT_TRAIL.csv`:
* **Location:** Top of `api_logger.py` (replace from line 6 to `_get_session_dir()`).
* **Replace header imports & setup with:**
```python
from __future__ import annotations
import csv
import getpass
import json
import os
import platform
import re
import socket
import sys
import time
import urllib.parse
import uuid
from datetime import datetime
from pathlib import Path
import requests

def get_base_dir() -> Path:
    """Resolve base directory next to the executable when frozen, or current working directory."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path.cwd()

def get_log_root() -> Path:
    root = get_base_dir() / "api_logs"
    root.mkdir(parents=True, exist_ok=True)
    return root

def get_audit_dir() -> Path:
    audit_dir = get_base_dir() / "audit_logs"
    audit_dir.mkdir(parents=True, exist_ok=True)
    return audit_dir

def get_system_identity() -> dict:
    """Extract complete corporate workstation, network, and user identifiers."""
    try:
        os_user = os.getlogin()
    except Exception:
        os_user = getpass.getuser()

    user_domain = os.environ.get("USERDOMAIN", "LOCAL")
    user_dns = os.environ.get("USERDNSDOMAIN", "")
    full_network_id = f"{user_domain}\\{os_user}" if user_domain else os_user

    hostname = socket.gethostname()
    try:
        fqdn = socket.getfqdn()
    except Exception:
        fqdn = hostname

    try:
        local_ip = socket.gethostbyname(hostname)
    except Exception:
        local_ip = "127.0.0.1"

    raw_mac = uuid.getnode()
    mac_addr = ":".join([f"{(raw_mac >> ele) & 0xff:02x}" for ele in range(0, 12, 2)][::-1])

    return {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "os_user": os_user,
        "user_domain": user_domain,
        "full_network_id": full_network_id,
        "user_dns_domain": user_dns,
        "computer_name": hostname,
        "fqdn": fqdn,
        "ip_address": local_ip,
        "mac_address": mac_addr,
        "os_platform": platform.platform(),
    }

def record_audit_trail(action: str, servicenow_user: str = "", details: str = "", status: str = "INFO", output_report: str = ""):
    """Append structured corporate user audit record to master CSV audit log."""
    try:
        audit_dir = get_audit_dir()
        csv_path = audit_dir / "USER_AUDIT_TRAIL.csv"
        ident = get_system_identity()
        timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        headers = [
            "Timestamp", "Action", "Status", "Network_ID", "OS_User", "Computer_Name",
            "IP_Address", "MAC_Address", "ServiceNow_User", "Details", "Output_Report", "Platform"
        ]

        file_exists = csv_path.exists()
        with open(csv_path, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            if not file_exists:
                writer.writeheader()
            writer.writerow({
                "Timestamp": timestamp_str,
                "Action": action,
                "Status": status,
                "Network_ID": ident["full_network_id"],
                "OS_User": ident["os_user"],
                "Computer_Name": ident["computer_name"],
                "IP_Address": ident["ip_address"],
                "MAC_Address": ident["mac_address"],
                "ServiceNow_User": servicenow_user or "N/A",
                "Details": details,
                "Output_Report": output_report or "",
                "Platform": ident["os_platform"],
            })
    except Exception as exc:
        print(f"[api_logger] Warning: Could not record audit trail: {exc}")
```

*(Tip: You can also simply copy the entire `full_source_files/api_logger.py`)*

---

## File 4: `CMAA.spec`
Make sure `upx=False` and `'getpass'` is in `hiddenimports`:
```python
    hiddenimports=[
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.styles.fills',
        'openpyxl.styles.colors',
        'openpyxl.reader.excel',
        'openpyxl.writer.excel',
        'et_xmlfile',
        'requests',
        'urllib3',
        'certifi',
        'rapidfuzz',
        'thefuzz',
        'extract_msg',
        'constants',
        'excel_handler',
        'questions_data',
        'sc_sheet_data',
        'servicenow_api',
        'api_logger',
        'dotenv',
        'getpass',
        'shlex',
    ],
```
And inside `exe = EXE(`:
```python
    strip=False,
    upx=False,
    upx_exclude=[],
```

---

## File 5: `build_cmaa_exe.bat`
Your complete, working batch script to build `--onefile`:
```bat
@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo =====================================================================
echo   CMAA Executable Builder (--onefile Standalone Package)
echo   Production-Ready Corporate Build System
echo =====================================================================
echo.

rem 1. Verify Python availability
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in system PATH.
    echo Please install Python 3.9+ or run within your corporate virtual environment.
    echo.
    pause
    exit /b 1
)

rem 2. Check/Install PyInstaller and core libraries
echo [STEP 1/5] Checking required build dependencies...
if exist "requirements-cmaa.txt" (
    python -m pip install -r requirements-cmaa.txt >nul 2>&1
) else if exist "requirements.txt" (
    python -m pip install -r requirements.txt >nul 2>&1
    python -m pip install pyinstaller >nul 2>&1
) else (
    python -m pip install pyinstaller openpyxl requests certifi urllib3 rapidfuzz python-dotenv >nul 2>&1
)
if %ERRORLEVEL% NEQ 0 (
    echo [WARNING] Pip install returned non-zero - possibly offline or air-gapped network.
    echo Proceeding with currently installed packages in Python environment...
)

rem 3. Clean prior build cache and all stale bytecode
echo.
echo [STEP 2/5] Cleaning previous build cache and stale bytecode...
if exist "build\CMAA_Audit" rmdir /s /q "build\CMAA_Audit"
if exist "dist\CMAA_Audit.exe" del /f /q "dist\CMAA_Audit.exe"
if exist "__pycache__" rmdir /s /q "__pycache__"

rem 4. Validate codebase integrity before building
echo.
echo [STEP 3/5] Validating codebase syntax...
python -m py_compile populate_cr_audit.py constants.py excel_handler.py sc_sheet_data.py servicenow_api.py api_logger.py questions_data.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [FATAL ERROR] Python syntax validation failed!
    echo A syntax error was detected in your codebase files.
    echo Please fix the errors reported above before building the executable.
    echo.
    pause
    exit /b 1
)
echo [OK] All codebase files compiled cleanly.

rem 5. Verify master template is present
if not exist "QAR Report.xlsx" (
    echo.
    echo [ERROR] Master template 'QAR Report.xlsx' not found in current directory!
    echo Please ensure 'QAR Report.xlsx' is present before packaging.
    echo.
    pause
    exit /b 1
)

rem 6. Build standalone executable
echo.
echo [STEP 4/5] Compiling standalone --onefile executable using CMAA.spec...
python -m PyInstaller CMAA.spec --clean --noconfirm

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [FATAL ERROR] PyInstaller build failed. Review output log above.
    echo.
    pause
    exit /b 1
)

rem 7. Deploy companion corporate test files
if exist "changes.txt" copy /y "changes.txt" "dist\changes.txt" >nul 2>&1
(
    echo # CMAA Corporate ServiceNow Configuration
    echo # Optional: Uncomment to pre-fill instance URL or defaults
    echo # SN_INSTANCE_URL=https://your-instance.service-now.com
    echo # SN_USERNAME=your_username
    echo # SN_VERIFY_SSL=false
) > "dist\.env.example"

rem 8. Run post-build smoke verification test
echo.
echo [STEP 5/5] Performing post-build verification on compiled binary...
if exist "dist\CMAA_Audit.exe" (
    dist\CMAA_Audit.exe --help >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        echo [OK] Standalone binary self-test passed!
    ) else (
        echo [WARNING] Binary returned non-zero exit on --help. Check antivirus or permissions.
    )
) else (
    echo [ERROR] Expected executable dist\CMAA_Audit.exe not found!
    echo.
    pause
    exit /b 1
)

echo.
echo =====================================================================
echo [SUCCESS] CMAA Standalone Executable Successfully Built!
echo.
echo Executable File: dist\CMAA_Audit.exe
echo Master Template: Bundled inside binary + hot-override enabled.
echo Authentication : Interactive prompt on launch (Username + Password masked).
echo Input Limits   : Max 100 CRs per run / Max 1-month date range window.
echo Output Reports : Saved to 'report\' with professional date-time nomenclature.
echo =====================================================================
echo.
pause
```

---

## File 6: `.env.example` (Optional Corporate Configuration)
If you wish to pre-configure your corporate credentials or instance URL without being prompted every time, create a file named `.env` in the same folder as `CMAA_Audit.exe` (or `populate_cr_audit.py`) with this syntax:

```ini
# =====================================================================
# Change Management Audit Automation (CMAA) - Environment Configuration
# =====================================================================
# Target ServiceNow Instance URL (HTTPS required)
SN_INSTANCE_URL=https://your-company.service-now.com

# ServiceNow Corporate Account Username
SN_USERNAME=your.username

# ServiceNow Corporate Password or API Token
SN_PASSWORD=your_password_here

# SSL Certificate Verification Setting:
# - 'false' : Skip SSL certificate verification (useful for internal corporate proxies)
# - 'true'  : Standard public SSL certificate verification (strict)
# - Path    : Full path to your corporate custom Root CA certificate bundle (.crt or .pem file)
#             Example: C:\Certs\CorporateRootCA.crt
SN_VERIFY_SSL=false
```
