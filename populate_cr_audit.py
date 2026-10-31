"""Command-line and interactive entry point for Change Management Audit Automation (CMAA)."""
from __future__ import annotations
import argparse
import os
import re
import sys
from datetime import datetime
import requests
from urllib3.exceptions import InsecureRequestWarning
import constants
import excel_handler
import servicenow_api

# SSL warnings suppression (corporate CA verification can be enabled via SN_VERIFY_SSL)
requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)


def pause_for_user(prompt: str = "\nPress Enter to exit..."):
    """Safely pause if running in interactive mode, avoiding EOFError if stdin is piped."""
    try:
        input(prompt)
    except (EOFError, KeyboardInterrupt):
        pass
    except Exception:
        pass


def _global_exception_hook(exc_type, exc_value, exc_tb):
    """Ensure the console window NEVER closes prematurely on fatal exceptions."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return
    print("\n" + "!" * 72)
    print(" [FATAL ERROR] An unexpected exception halted the application:")
    print("!" * 72)
    import traceback
    traceback.print_exception(exc_type, exc_value, exc_tb)
    print("!" * 72)
    print(" The terminal window has been kept open so you can inspect the error above.")
    print("!" * 72)
    pause_for_user()


sys.excepthook = _global_exception_hook


def resolve_template_path(template_name: str | None = None) -> str:
    """Resolve master template path across direct filesystem and PyInstaller bundles."""
    return excel_handler.resolve_template_path(template_name)


def parse_arguments():
    parser = argparse.ArgumentParser(description="Populate Change Management audit results in the Excel template")
    parser.add_argument("--start-date", help="Start date in YYYY-MM-DD format")
    parser.add_argument("--end-date", help="End date in YYYY-MM-DD format")
    parser.add_argument("--percentage", type=int, default=20,
                        help="Percentage of date-range population to sample (default: 20)")
    parser.add_argument("--template", default="QAR Report.xlsx", help="Path to the Excel template")
    parser.add_argument("--tester", default="Tester", help="Tester name")
    parser.add_argument("--change-nums", nargs="+", metavar="CHANGE_OR_FILE",
                        help="Change numbers or a path to a text file")
    parser.add_argument("--output", help="Optional output filename/path")
    return parser.parse_args()


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


def validate_date_range(start_date, end_date):
    try:
        start_datetime = datetime.strptime(start_date, "%Y-%m-%d")
        end_datetime = datetime.strptime(end_date, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError("Dates must use YYYY-MM-DD format (e.g. 2026-10-01)") from exc
    if start_datetime > end_datetime:
        raise ValueError("--start-date must not be after --end-date")
    return start_datetime, end_datetime


def report_filename(month, change_wise=False):
    year = datetime.today().strftime("%y")
    if change_wise:
        return f"FCB {month}-QAR Change Wise.xlsx"
    return f"FCB {month}'{year} - QAR Report.xlsx"


def get_reference_display_value(value):
    """Extract a ServiceNow reference display value or string value."""
    if isinstance(value, dict):
        return value.get("display_value")
    if isinstance(value, str):
        return value
    return None


def print_missing_change_numbers(requested_numbers, returned_crs):
    returned_numbers = {(cr.get("number") or "").strip().upper() for cr in returned_crs if cr.get("number")}
    missing_numbers = [n for n in requested_numbers if n not in returned_numbers]
    print(f"Requested: {len(requested_numbers)}")
    print(f"Returned : {len(returned_numbers)}")
    if missing_numbers:
        print("\nMissing CRs:")
        for number in missing_numbers:
            print(f"  {number}")
    return missing_numbers


def load_workbook(template_path):
    resolved_path = resolve_template_path(template_path)
    try:
        return excel_handler.load_template(resolved_path)
    except FileNotFoundError:
        print("\n" + "=" * 65)
        print(" [ERROR] Excel Master Template Missing!")
        print("=" * 65)
        print(f" Could not locate the master template file: '{template_path}'")
        print(f" Looked in paths including: '{resolved_path}'")
        print(" Please place 'QAR Report.xlsx' in the application folder.")
        print("=" * 65)
        return None
    except Exception as exc:
        print(f"[ERROR] Failed to load Excel template '{resolved_path}': {exc}")
        return None


def save_workbook(workbook, output_path):
    try:
        saved_path = excel_handler.save_output(workbook, output_path)
    except Exception as exc:
        print(f"[ERROR] Failed saving output file: {exc}")
        return None
    print(f"\n[SUCCESS] Audit report created successfully:")
    print(f"  --> {os.path.abspath(saved_path)}")
    if getattr(constants, "api_logger", None) and getattr(constants.api_logger, "_HIT_COUNTER", 0) > 0:
        print(f"API Audit Trail ({constants.api_logger._HIT_COUNTER} hits logged): {constants.api_logger.get_summary_path()}")
    return saved_path


def populate_all_basic_data(workbook, basic_data):
    for sheet_name in ["SC Result Sheet", "SC Template Result Sheet", "NC Result Sheet", "EC Result Sheet", "Exp Result Sheet"]:
        excel_handler.populate_basic_data(workbook, sheet_name, basic_data)


def create_change_wise_report(args, all_crs, tester, month):
    sc_col = sc_template_col = nc_col = ec_col = exp_col = 3
    filename = report_filename(month, change_wise=True)
    existing_report = os.path.join("report", filename)
    template_path = existing_report if os.path.exists(existing_report) else resolve_template_path(args.template)
    workbook = load_workbook(template_path)
    if workbook is None:
        return False
    basic_data = servicenow_api.get_basic_data(tester, month, len(all_crs), len(all_crs))
    populate_all_basic_data(workbook, basic_data)
    remaining = len(all_crs)
    for cr in all_crs:
        change_number = cr.get("number")
        change_type = (cr.get("type") or "").casefold()
        print(f"\nProcessing {change_number} ({change_type})")
        print(f"Changes remaining: {remaining}")
        remaining -= 1
        audit_data = servicenow_api.get_audit_data_single(cr)
        if change_type == constants.STANDARD.casefold():
            excel_handler.populate_sheet_single(workbook, "SC Result Sheet", audit_data, sc_col)
            sc_col += 1
            standard_template_data = servicenow_api.get_audit_data_single(cr, sheet_name="SC Template Result Sheet")
            excel_handler.populate_sheet_single(workbook, "SC Template Result Sheet", standard_template_data, sc_template_col)
            sc_template_col += 1
        elif change_type == constants.NORMAL.casefold():
            excel_handler.populate_sheet_single(workbook, "NC Result Sheet", audit_data, nc_col)
            nc_col += 1
        elif change_type == constants.EMERGENCY.casefold():
            excel_handler.populate_sheet_single(workbook, "EC Result Sheet", audit_data, ec_col)
            ec_col += 1
        elif change_type == constants.EXPEDITED.casefold():
            excel_handler.populate_sheet_single(workbook, "Exp Result Sheet", audit_data, exp_col)
            exp_col += 1
        else:
            print(f"Skipping unsupported Change type: {cr.get('type')} ({change_number})")
    saved = save_workbook(workbook, args.output or filename)
    return saved is not None


def get_standard_template_proposals(all_crs, today):
    unique_proposals = {}
    for cr in all_crs:
        if (cr.get("type") or "").casefold() != constants.STANDARD.casefold():
            continue
        proposal_name = get_reference_display_value(cr.get("std_change_producer_version"))
        if proposal_name and proposal_name not in unique_proposals:
            unique_proposals[proposal_name] = cr
    selected = []
    for cr in unique_proposals.values():
        proposal = servicenow_api.get_proposals(cr)
        last_review = proposal.get("last_review_date")
        if not last_review:
            continue
        try:
            review_date = datetime.strptime(last_review, "%Y-%m-%d")
        except ValueError:
            print(f"Invalid Standard Proposal review date for {cr.get('number')}: {last_review}")
            continue
        months = (today.year - review_date.year) * 12 + today.month - review_date.month
        if months > 10:
            selected.append(cr)
    return selected


def process_monthly_sheet(args, workbook, sheet_name, change_type, all_crs, tester, month, today):
    if sheet_name.casefold() == "SC Template Result Sheet".casefold():
        selected_crs = get_standard_template_proposals(all_crs, today)
    else:
        selected_crs = servicenow_api.random_select(all_crs, args.percentage)
    print(f"Randomly selected {len(selected_crs)} {change_type.upper()} CRs ({args.percentage}%)")
    if not selected_crs:
        print(f"No {change_type} CRs selected. Skipping {sheet_name}.")
        return
    basic_data = servicenow_api.get_basic_data(tester, month, len(all_crs), len(selected_crs))
    audit_data = servicenow_api.get_audit_data(selected_crs, sheet_name=sheet_name)
    excel_handler.populate_basic_data(workbook, sheet_name, basic_data)
    excel_handler.populate_sheet(workbook, sheet_name, audit_data)


def handle_service_error(exc: Exception):
    """Format clear, actionable error diagnosis for corporate networks."""
    if isinstance(exc, requests.exceptions.HTTPError):
        code = getattr(exc.response, "status_code", None)
        if code == 401:
            print("\n" + "=" * 65)
            print(" [ERROR 401] ServiceNow Authentication Failed")
            print("=" * 65)
            print(" The supplied username or password was rejected by ServiceNow.")
            print(" Please verify your corporate credentials or password rotation.")
            print("=" * 65)
            return
        elif code == 403:
            print("\n" + "=" * 65)
            print(" [ERROR 403] ServiceNow Permission Denied")
            print("=" * 65)
            print(" Your account does not have read access to the requested table.")
            print("=" * 65)
            return
        elif code == 404:
            print("\n" + "=" * 65)
            print(" [ERROR 404] ServiceNow Resource Not Found")
            print("=" * 65)
            print(f" Verify your SN_INSTANCE_URL: {constants.INSTANCE_URL}")
            print("=" * 65)
            return
    elif isinstance(exc, (requests.exceptions.ConnectionError, requests.exceptions.Timeout)):
        print("\n" + "=" * 65)
        print(" [CONNECTION ERROR] Unable to Reach ServiceNow")
        print("=" * 65)
        print(f" Target Instance: {constants.INSTANCE_URL}")
        print(" Please check:")
        print("  1. Are you connected to the corporate VPN or office network?")
        print("  2. Is your corporate HTTP/HTTPS proxy configured if required?")
        print("  3. Does your workstation DNS resolve the instance domain?")
        print("=" * 65)
        return
    print(f"\n[ERROR] ServiceNow API failure: {exc}")


def run_date_range_audit(args, start_date, end_date):
    start_datetime, end_datetime = validate_date_range(start_date, end_date)
    if not 1 <= args.percentage <= 100:
        print("Error: --percentage must be between 1 and 100.")
        return False
    today = datetime.today()
    month = start_datetime.strftime("%B")
    print(f"Date Range: {start_date} to {end_date}")
    filename = report_filename(month)
    output_path = args.output or filename
    existing_report = os.path.join("report", filename)
    template_path = existing_report if os.path.exists(existing_report) else resolve_template_path(args.template)
    workbook = load_workbook(template_path)
    if workbook is None:
        return False
    for sheet_name, change_type in constants.CHANGE_RESULTS.items():
        try:
            all_crs = servicenow_api.fetch_crs_from_servicenow(
                start_date=start_date, end_date=end_date, change_type=change_type.casefold().replace(" result sheet", ""))
        except Exception as exc:
            handle_service_error(exc)
            return False
        print(f"Fetched {len(all_crs)} {change_type} CRs")
        if not all_crs:
            print(f"No {change_type} CRs found. Skipping {sheet_name}.")
            continue
        process_monthly_sheet(args, workbook, sheet_name, change_type, all_crs, args.tester, month, today)
    saved = save_workbook(workbook, output_path)
    return saved is not None


def run_change_number_audit(args, change_numbers):
    print(f"Requested Change numbers: {len(change_numbers)}")
    try:
        all_crs = servicenow_api.fetch_crs_from_servicenow(change_numbers=change_numbers)
    except Exception as exc:
        handle_service_error(exc)
        return False
    print_missing_change_numbers(change_numbers, all_crs)
    if not all_crs:
        print("No matching Change Requests returned from ServiceNow.")
        return False
    return create_change_wise_report(args, all_crs, args.tester, datetime.today().strftime("%B"))


def ensure_credentials(is_interactive: bool = True) -> bool:
    """Ensure valid ServiceNow credentials exist; prompt securely if missing in interactive mode."""
    if constants.USERNAME and constants.PASSWORD:
        return True

    if not is_interactive:
        print("[ERROR] Missing ServiceNow credentials.")
        print("Please configure SN_USERNAME and SN_PASSWORD in your environment or in a local .env file.")
        return False

    print("\n" + "=" * 65)
    print("      ServiceNow Corporate Authentication Setup")
    print("=" * 65)
    print(f"Default Instance URL: {constants.INSTANCE_URL}")
    new_url = input("Enter Instance URL (or press Enter to keep default): ").strip()
    if new_url:
        constants.update_credentials("", "", instance_url=new_url)

    import getpass
    username = input("Enter ServiceNow Username: ").strip()
    if not username:
        print("Username cannot be empty.")
        return False
    password = getpass.getpass("Enter ServiceNow Password (masked): ").strip()
    if not password:
        print("Password cannot be empty.")
        return False

    constants.update_credentials(username=username, password=password)

    save_env = input("\nSave credentials to local .env for future automatic runs? [y/N]: ").strip().lower()
    if save_env in {"y", "yes"}:
        try:
            target_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.getcwd()
            env_path = os.path.join(target_dir, ".env")
            with open(env_path, "w", encoding="utf-8") as f:
                f.write(f"# CMAA Enterprise Configuration\n")
                f.write(f"SN_INSTANCE_URL={constants.INSTANCE_URL}\n")
                f.write(f"SN_USERNAME={username}\n")
                f.write(f"SN_PASSWORD={password}\n")
                f.write(f"SN_VERIFY_SSL={str(constants.VERIFY_SSL).lower()}\n")
            print(f"[OK] Saved credentials to: {env_path}")
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
            inp = input("\nEnter Change Number(s) or path to text file: ").strip()
            if not inp:
                print("No input provided. Returning to main menu.")
                continue
            return argparse.Namespace(
                change_nums=inp.split(),
                start_date=None,
                end_date=None,
                percentage=20,
                template="QAR Report.xlsx",
                tester="Tester",
                output=None,
            )
        elif choice == "2":
            start_date = input("Enter Start Date (YYYY-MM-DD): ").strip()
            end_date = input("Enter End Date (YYYY-MM-DD): ").strip()
            if not start_date or not end_date:
                print("Start and End dates are required. Returning to main menu.")
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
    is_interactive = False
    if len(sys.argv) == 1:
        is_interactive = True
        args = interactive_prompt()
        if not args:
            if is_interactive:
                pause_for_user("\nPress Enter to exit...")
            sys.exit(0)
    else:
        args = parse_arguments()

    if not ensure_credentials(is_interactive=is_interactive):
        if is_interactive:
            pause_for_user("\nPress Enter to exit...")
        sys.exit(1)

    try:
        change_numbers = load_change_numbers(args.change_nums)
    except RuntimeError as exc:
        print(f"\n[ERROR] {exc}")
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
    except Exception as exc:
        print(f"\n[ERROR] Execution halted: {exc}")
    finally:
        if is_interactive:
            pause_for_user("\nExecution completed. Press Enter to exit...")


if __name__ == "__main__":
    main()
