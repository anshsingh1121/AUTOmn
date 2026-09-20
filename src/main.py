"""
FCB Monthly Incident Tracker Automation — Main Orchestrator

Pipeline:
    INPUT → VALIDATION → NORMALIZATION → RECONCILIATION →
    BUSINESS RULES → OUTPUT → VALIDATION → REPORTING → AUDIT

Usage:
    python -m src.main                           # Dry run with default config
    python -m src.main --config path/to/config.json
    python -m src.main --dry-run                 # Explicit dry run
    python -m src.main --production              # Production mode (writes output)
"""

import argparse
import os
import sys
from datetime import datetime

# Add project root to path for imports
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from src.config_loader import Config, ConfigError, load_config
from src.data_loader import DataLoadError, load_excel_to_dataframe, load_master_tracker, load_servicenow, load_ic_lookup
from src.validator import (
    ValidationResult,
    validate_required_columns,
    validate_number_column,
    validate_ic_lookup,
    validate_output,
)
from src.normalizer import normalize_whitespace
from src.business_rules import build_ic_lookup_map
from src.reconciler import reconcile, reconciliation_to_dataframe
from src.output_writer import (
    create_archive_copy,
    create_output_path,
    write_output_workbook,
    write_reconciliation_report_excel,
)
from src.report_generator import (
    generate_reconciliation_summary,
    format_console_report,
)
from src.audit import create_audit_record, save_audit_log


def parse_args():
    parser = argparse.ArgumentParser(
        description="FCB Monthly Incident Tracker Automation"
    )
    parser.add_argument(
        "--config",
        default=os.path.join(project_root, "config", "config.json"),
        help="Path to configuration file (default: config/config.json)"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=False,
        help="Dry run: validate and report without writing output"
    )
    parser.add_argument(
        "--production",
        action="store_true",
        default=False,
        help="Production mode: write output after successful validation"
    )
    return parser.parse_args()


def run_pipeline(config: Config) -> int:
    """
    Execute the full automation pipeline.

    Returns:
        0 on success, 1 on validation failure, 2 on error.
    """
    print("=" * 60)
    print("FCB MONTHLY INCIDENT TRACKER AUTOMATION")
    print(f"Mode: {config.mode} | Dry Run: {config.dry_run}")
    print(f"Timestamp: {datetime.now().isoformat()}")
    print("=" * 60)
    print()

    # ========================================
    # STAGE 1: INPUT
    # ========================================
    print("[1/8] Loading input files...")

    try:
        master_df, master_wb = load_master_tracker(config)
        print(f"  Master Tracker: {len(master_df)} records loaded")
    except DataLoadError as e:
        print(f"  ✗ FATAL: Cannot load Master Tracker: {e}")
        return 2

    try:
        servicenow_df = load_servicenow(config)
        print(f"  ServiceNow:     {len(servicenow_df)} records loaded")
    except DataLoadError as e:
        print(f"  ✗ FATAL: Cannot load ServiceNow export: {e}")
        return 2

    try:
        ic_lookup_df = load_ic_lookup(config)
        print(f"  IC Lookup:      {len(ic_lookup_df)} entries loaded")
    except DataLoadError as e:
        print(f"  ⚠ WARNING: Cannot load IC Lookup: {e}")
        print("  Continuing without IC Lookup (Region will be blank)")
        ic_lookup_df = None

    print()

    # ========================================
    # STAGE 2: INPUT VALIDATION
    # ========================================
    print("[2/8] Validating input data...")

    input_validation = ValidationResult()

    # Validate ServiceNow columns
    sn_col_result = validate_required_columns(
        servicenow_df, config.required_servicenow_columns, "ServiceNow"
    )
    input_validation.merge(sn_col_result)

    # Validate Master columns
    master_col_result = validate_required_columns(
        master_df, config.required_master_columns, "Master Tracker"
    )
    input_validation.merge(master_col_result)

    # Validate Number column — ServiceNow
    sn_num_result = validate_number_column(
        servicenow_df, "Number", "ServiceNow",
        block_on_duplicates=config.block_on_duplicate_numbers,
        block_on_blanks=config.block_on_blank_numbers
    )
    input_validation.merge(sn_num_result)

    # Validate Number column — Master
    master_num_result = validate_number_column(
        master_df, "Number", "Master Tracker",
        block_on_duplicates=config.block_on_duplicate_numbers,
        block_on_blanks=config.block_on_blank_numbers
    )
    input_validation.merge(master_num_result)

    # Validate IC Lookup
    ic_lookup_warnings = []
    if ic_lookup_df is not None:
        region_cfg = config.region_rule
        ic_result = validate_ic_lookup(
            ic_lookup_df,
            region_cfg.get("lookup_ic_column", "IC"),
            region_cfg.get("lookup_region_column", "Region")
        )
        input_validation.merge(ic_result)

    print(input_validation.summary())
    print()

    if not input_validation.is_valid:
        print("✗ INPUT VALIDATION FAILED — aborting pipeline.")
        return 1

    # ========================================
    # STAGE 3: NORMALIZATION & LOOKUP BUILD
    # ========================================
    print("[3/8] Building IC lookup map...")

    ic_lookup_map = {}
    if ic_lookup_df is not None:
        region_cfg = config.region_rule
        ic_lookup_map, lookup_warnings = build_ic_lookup_map(
            ic_lookup_df,
            ic_column=region_cfg.get("lookup_ic_column", "IC"),
            region_column=region_cfg.get("lookup_region_column", "Region"),
            case_sensitive=region_cfg.get("case_sensitive", False)
        )
        for w in lookup_warnings:
            input_validation.add_warning(w)
        print(f"  IC Lookup map: {len(ic_lookup_map)} entries")
    else:
        print("  IC Lookup: not available (Region will be blank)")
    print()

    # ========================================
    # STAGE 4: RECONCILIATION + BUSINESS RULES
    # ========================================
    print("[4/8] Reconciling incidents...")

    master_before_count = len(master_df)

    recon_result = reconcile(
        master_df=master_df,
        servicenow_df=servicenow_df,
        ic_lookup=ic_lookup_map,
        config=config
    )

    summary = recon_result.summary()
    print(f"  Matched (updated): {summary['updated']}")
    print(f"  New records:       {summary['new']}")
    print(f"  Historical:        {summary['historical_retained']}")
    print(f"  Total output:      {summary['total_output']}")
    print(f"  IC lookup misses:  {summary['ic_lookup_misses']}")
    print()

    # ========================================
    # STAGE 5: BUILD OUTPUT DATAFRAME
    # ========================================
    print("[5/8] Building output dataset...")

    master_columns = list(master_df.columns)
    output_df = reconciliation_to_dataframe(recon_result, master_columns)
    print(f"  Output DataFrame: {len(output_df)} records, {len(output_df.columns)} columns")
    print()

    # ========================================
    # STAGE 6: OUTPUT VALIDATION
    # ========================================
    print("[6/8] Validating output...")

    output_validation = validate_output(output_df, master_df, "Number")
    print(output_validation.summary())
    print()

    if not output_validation.is_valid:
        print("✗ OUTPUT VALIDATION FAILED — output will not be written.")
        return 1

    # ========================================
    # STAGE 7: WRITE OUTPUT (if not dry run)
    # ========================================
    output_path = None
    archive_path = None
    has_updates = bool(recon_result.new_records or recon_result.change_details)

    if not has_updates:
        print("[7/8] No new incidents or changes detected. Skipping output write.")
    elif config.dry_run:
        print("[7/8] DRY RUN — skipping output write.")
        print("  (Re-run with --production to write output)")
    else:
        print("[7/8] Writing output...")

        # Archive the original master
        archive_path = create_archive_copy(
            config.master_tracker_path, config.archive_directory
        )
        if archive_path:
            print(f"  Archive created: {os.path.basename(archive_path)}")

        # Write output workbook
        output_path = create_output_path(
            config.output_directory,
            source_path=config.master_tracker_path,
        )
        write_output_workbook(
            output_df=output_df,
            output_path=output_path,
            source_workbook_path=config.master_tracker_path,
            target_sheets=config.target_sheets,
            formula_columns=None  # Auto-detected from source workbook
        )
        print(f"  Output written: {os.path.basename(output_path)}")

    print()

    # ========================================
    # STAGE 8: REPORTING & AUDIT
    # ========================================
    print("[8/8] Generating reports and audit trail...")

    # Generate report
    report = generate_reconciliation_summary(
        master_before_count=master_before_count,
        servicenow_count=len(servicenow_df),
        recon=recon_result,
        input_validation=input_validation,
        output_validation=output_validation
    )

    # Print console report
    print()
    print(format_console_report(report))
    print()

    # Write Excel reconciliation report
    if has_updates:
        report_path = create_output_path(config.output_directory, prefix="Reconciliation_Report")
        write_reconciliation_report_excel(report, report_path)
        print(f"  Report generated: {os.path.basename(report_path)}")
    else:
        print("  Skipped Excel report generation (no changes).")

    # Write audit log
    if has_updates:
        audit_record = create_audit_record(
            config=config,
            master_before_count=master_before_count,
            servicenow_count=len(servicenow_df),
            recon_summary=summary,
            validation_passed=(input_validation.is_valid and output_validation.is_valid),
            report_summary=report["summary"],
            output_path=output_path,
            archive_path=archive_path,
            errors=report.get("errors", []),
            warnings=report.get("warnings", [])
        )
        audit_path = save_audit_log(audit_record, config.output_directory)
        print(f"  Audit log: {os.path.basename(audit_path)}")
    else:
        print("  Skipped Audit log generation (no changes).")

    print()
    print("=" * 60)
    final_status = report["summary"].get("Status", "UNKNOWN")
    if final_status == "PASSED" and not has_updates:
        final_status = "UP TO DATE (NO CHANGES)"
    print(f"PIPELINE COMPLETE — STATUS: {final_status}")
    print("=" * 60)

    return 0 if "PASSED" in final_status or "UP TO DATE" in final_status else 1


def main():
    args = parse_args()

    # Load config
    try:
        config = load_config(args.config)
    except ConfigError as e:
        print(f"✗ CONFIGURATION ERROR: {e}")
        sys.exit(2)

    # Override dry_run from command line
    if args.production:
        config._data["dry_run"] = False
        config._data["mode"] = "production"
    elif args.dry_run:
        config._data["dry_run"] = True

    # GUI File Picker for non-tech users
    from src.gui_picker import prompt_for_file
    import os

    # 1. Ask for Master Tracker if not found
    raw_master = config._data.get("paths", {}).get("master_tracker", "")
    if not raw_master or raw_master == "ASK" or not os.path.exists(config.master_tracker_path):
        print("Waiting for user to select the Master Tracker file...")
        selected = prompt_for_file(
            "Select the Master Tracker Excel File",
            [("Excel Macro-Enabled", "*.xlsm"), ("Excel Workbook", "*.xlsx"), ("All Files", "*.*")]
        )
        if not selected:
            print("✗ Operation cancelled. No Master Tracker selected.")
            sys.exit(1)
        config._data["paths"]["master_tracker"] = selected
        print(f"Master Tracker: {selected}")

    # 2. Ask for ServiceNow Export if not found
    raw_sn = config._data.get("paths", {}).get("servicenow_input", "")
    if not raw_sn or raw_sn == "ASK" or not os.path.exists(config.servicenow_input_path):
        print("Waiting for user to select the ServiceNow Export file...")
        selected = prompt_for_file(
            "Select the downloaded ServiceNow Export File",
            [("Excel Workbook", "*.xlsx"), ("CSV Document", "*.csv"), ("All Files", "*.*")]
        )
        if not selected:
            print("✗ Operation cancelled. No ServiceNow Export selected.")
            sys.exit(1)
        config._data["paths"]["servicenow_input"] = selected
        print(f"ServiceNow Export: {selected}")

    print()

    # Run pipeline
    exit_code = run_pipeline(config)

    # Auto-open the output folder for the user (Windows only)
    if exit_code == 0 and os.name == 'nt':
        try:
            os.startfile(os.path.abspath(config.output_directory))
        except Exception:
            pass

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
