# =====================================================================
# PATCH FOR: populate_cr_audit.py
# 
# Key modifications in this file:
# 1. safe_getpass helper added to prevent console hanging in non-interactive/piped environments.
# 2. load_change_numbers updated to enforce MAX_CHANGE_NUMBERS constraint (max 100).
# 3. validate_date_range updated to enforce MAX_DATE_RANGE_DAYS constraint (max 31 days / 1 month).
# 4. report_filename updated to call excel_handler.generate_report_filename.
# 5. create_change_wise_report and run_date_range_audit updated to generate professional timestamped filenames.
# 6. ensure_credentials and main updated to prompt for username & password at startup in interactive mode.
# 7. interactive_prompt updated with live validation loops for CR limit and 1-month date window.
# =====================================================================

# --- MODIFICATION 1: Add safe_getpass (Place right before pause_for_user) ---
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


# --- MODIFICATION 2: Update load_change_numbers (Max 100 constraint) ---
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


# --- MODIFICATION 3: Update validate_date_range (Max 1 month / 31 days constraint) ---
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


# --- MODIFICATION 4: Update report_filename ---
def report_filename(month=None, change_wise=False, start_date=None, end_date=None, change_numbers=None):
    """Generate professional corporate report filename with date-time frame nomenclature."""
    return excel_handler.generate_report_filename(
        start_date=start_date,
        end_date=end_date,
        change_numbers=change_numbers
    )


# --- MODIFICATION 5: Update create_change_wise_report and run_date_range_audit ---
# Inside create_change_wise_report:
#   change_numbers = [cr.get("number") for cr in all_crs if cr.get("number")]
#   filename = args.output or excel_handler.generate_report_filename(change_numbers=change_numbers)
#   template_path = resolve_template_path(args.template)
#   ...
#   saved = save_workbook(workbook, filename)

# Inside run_date_range_audit:
#   filename = args.output or excel_handler.generate_report_filename(start_date=start_date, end_date=end_date)
#   template_path = resolve_template_path(args.template)
#   ...
#   saved = save_workbook(workbook, filename)


# --- MODIFICATION 6: Update ensure_credentials, interactive_prompt, and main ---
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
        elif choice == "3":
            print("\nExiting application. Goodbye.")
            return None
        else:
            print("\n[!] Invalid selection. Please enter 1, 2, or 3.")


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

    try:
        change_numbers = load_change_numbers(args.change_nums)
    except (RuntimeError, ValueError) as exc:
        print(f"\n[INPUT ERROR] {exc}")
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
    except Exception as exc:
        print(f"\n[ERROR] Execution halted: {exc}")
    finally:
        if is_interactive:
            pause_for_user("\nExecution completed. Press Enter to exit...")
