def interactive_prompt():
    """Interactive fallback when executed without CLI arguments (e.g., double-clicked executable)."""
    print("=" * 65)
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
            print("No input provided.")
            return None
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
    return None


def main():
    is_interactive = False
    if len(sys.argv) == 1:
        args = interactive_prompt()
        if not args:
            sys.exit(0)
        is_interactive = True
    else:
        args = parse_arguments()

    try:
        change_numbers = load_change_numbers(
            args.change_nums
        )
    except RuntimeError as exc:
        print(f"Error: {exc}")
        if is_interactive:
            input("\nPress Enter to exit...")
        sys.exit(1)

    try:
        if change_numbers:
            run_change_number_audit(
                args=args,
                change_numbers=change_numbers,
            )
            return

        if not args.start_date or not args.end_date:
            print(
                "Error: provide either --change-nums "
                "or both --start-date and --end-date."
            )
            if is_interactive:
                input("\nPress Enter to exit...")
            sys.exit(2)

        try:
            validate_date_range(
                args.start_date,
                args.end_date,
            )
        except ValueError as exc:
            print(f"Error: {exc}")
            if is_interactive:
                input("\nPress Enter to exit...")
            sys.exit(2)

        run_date_range_audit(
            args=args,
            start_date=args.start_date,
            end_date=args.end_date,
        )
    finally:
        if is_interactive:
            input("\nAudit complete. Press Enter to exit...")


if __name__ == "__main__":
    main()
