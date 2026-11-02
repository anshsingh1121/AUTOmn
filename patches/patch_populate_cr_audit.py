# =====================================================================
# PATCH FOR: populate_cr_audit.py
# 
# Key modifications in this file:
# 1. safe_getpass helper added to prevent console hanging in non-interactive/piped environments.
# 2. report_filename updated to call excel_handler.generate_report_filename.
# 3. create_change_wise_report and run_date_range_audit updated to generate professional timestamped filenames.
# 4. ensure_credentials and main updated to prompt for username & password at startup in interactive mode.
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


# --- MODIFICATION 2: Update report_filename ---
def report_filename(month=None, change_wise=False, start_date=None, end_date=None, change_numbers=None):
    """Generate professional corporate report filename with date-time frame nomenclature."""
    return excel_handler.generate_report_filename(
        start_date=start_date,
        end_date=end_date,
        change_numbers=change_numbers
    )


# --- MODIFICATION 3: Update create_change_wise_report ---
# (Inside create_change_wise_report, replace the top section where filename and template_path are resolved)
# BEFORE:
#   filename = report_filename(month, change_wise=True)
#   existing_report = os.path.join("report", filename)
#   template_path = existing_report if os.path.exists(existing_report) else resolve_template_path(args.template)
# AFTER:
#   change_numbers = [cr.get("number") for cr in all_crs if cr.get("number")]
#   filename = args.output or excel_handler.generate_report_filename(change_numbers=change_numbers)
#   template_path = resolve_template_path(args.template)
#   ...
#   saved = save_workbook(workbook, filename)


# --- MODIFICATION 4: Update run_date_range_audit ---
# (Inside run_date_range_audit, replace the top section where filename is resolved)
# BEFORE:
#   filename = report_filename(month)
#   output_path = args.output or filename
#   existing_report = os.path.join("report", filename)
#   template_path = existing_report if os.path.exists(existing_report) else resolve_template_path(args.template)
# AFTER:
#   filename = args.output or excel_handler.generate_report_filename(start_date=start_date, end_date=end_date)
#   template_path = resolve_template_path(args.template)
#   ...
#   saved = save_workbook(workbook, filename)


# --- MODIFICATION 5: Update ensure_credentials and main ---
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
