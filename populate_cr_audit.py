"""Command-line entry point for Change Management audit automation."""
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

# Retained from the source screenshots, but SSL verification can be enabled via .env.
requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)


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
        raise ValueError("Dates must use YYYY-MM-DD format") from exc
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
    try:
        return excel_handler.load_template(template_path)
    except FileNotFoundError:
        print(f"Error: Template file '{template_path}' not found.")
        sys.exit(1)


def save_workbook(workbook, output_path):
    try:
        saved_path = excel_handler.save_output(workbook, output_path)
    except Exception as exc:
        print(f"Error saving output file: {exc}")
        sys.exit(1)
    print(f"Successfully created: {saved_path}")
    if hasattr(constants, "api_logger") and constants.api_logger._HIT_COUNTER > 0:
        print(f"API Audit Trail ({constants.api_logger._HIT_COUNTER} hits logged): {constants.api_logger.get_summary_path()}")


def populate_all_basic_data(workbook, basic_data):
    for sheet_name in ["SC Result Sheet", "SC Template Result Sheet", "NC Result Sheet", "EC Result Sheet", "Exp Result Sheet"]:
        excel_handler.populate_basic_data(workbook, sheet_name, basic_data)


def create_change_wise_report(args, all_crs, tester, month):
    sc_col = sc_template_col = nc_col = ec_col = exp_col = 3
    filename = report_filename(month, change_wise=True)
    existing_report = os.path.join("report", filename)
    template_path = existing_report if os.path.exists(existing_report) else args.template
    workbook = load_workbook(template_path)
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
    save_workbook(workbook, args.output or filename)


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
        try:
            from servicenow_api import get_proposals
            proposal_result = get_proposals(cr)
        except (requests.RequestException, ValueError):
            selected.append(cr)
            continue
        last_review = proposal_result.get("u_last_reviewed")
        if not last_review:
            selected.append(cr)
            continue
        try:
            review_date = datetime.strptime(str(last_review), "%m/%d/%Y")
        except (TypeError, ValueError):
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


def run_date_range_audit(args, start_date, end_date):
    start_datetime, end_datetime = validate_date_range(start_date, end_date)
    if not 1 <= args.percentage <= 100:
        print("Error: --percentage must be between 1 and 100.")
        sys.exit(2)
    today = datetime.today()
    month = start_datetime.strftime("%B")
    print(f"Date Range: {start_date} to {end_date}")
    filename = report_filename(month)
    output_path = args.output or filename
    existing_report = os.path.join("report", filename)
    template_path = existing_report if os.path.exists(existing_report) else args.template
    workbook = load_workbook(template_path)
    for sheet_name, change_type in constants.CHANGE_RESULTS.items():
        try:
            all_crs = servicenow_api.fetch_crs_from_servicenow(
                start_date=start_date, end_date=end_date, change_type=change_type.casefold().replace(" result sheet", ""))
        except (requests.RequestException, ValueError) as exc:
            print(f"Error fetching {change_type} CRs: {exc}")
            sys.exit(1)
        print(f"Fetched {len(all_crs)} {change_type} CRs")
        if not all_crs:
            print(f"No {change_type} CRs found. Skipping {sheet_name}.")
            continue
        process_monthly_sheet(args, workbook, sheet_name, change_type, all_crs, args.tester, month, today)
    save_workbook(workbook, output_path)


def run_change_number_audit(args, change_numbers):
    print(f"Requested Change numbers: {len(change_numbers)}")
    try:
        all_crs = servicenow_api.fetch_crs_from_servicenow(change_numbers=change_numbers)
    except (requests.RequestException, ValueError) as exc:
        print(f"Error fetching Change Requests: {exc}")
        sys.exit(1)
    print_missing_change_numbers(change_numbers, all_crs)
    if not all_crs:
        print("No matching Change Requests returned.")
        sys.exit(0)
    create_change_wise_report(args, all_crs, args.tester, datetime.today().strftime("%B"))


def main():
    args = parse_arguments()
    try:
        change_numbers = load_change_numbers(args.change_nums)
    except RuntimeError as exc:
        print(f"Error: {exc}")
        sys.exit(1)
    if change_numbers:
        run_change_number_audit(args, change_numbers)
        return
    if not args.start_date or not args.end_date:
        print("Error: provide either --change-nums or both --start-date and --end-date.")
        sys.exit(2)
    try:
        validate_date_range(args.start_date, args.end_date)
    except ValueError as exc:
        print(f"Error: {exc}")
        sys.exit(2)
    run_date_range_audit(args, args.start_date, args.end_date)


if __name__ == "__main__":
    main()
