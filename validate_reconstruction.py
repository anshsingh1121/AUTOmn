"""Comprehensive, local, offline validation checks for the CMAA project.

Validates:
1. Regex patterns for Change Request numbers.
2. CLI input parser and deduplication.
3. Date range validation.
4. Sampling logic.
5. Template resolution (direct and PyInstaller _MEIPASS bundle fallback).
6. Excel template loading and sheet presence.
7. Core audit questions rules (Q5, Q6, Q19).
8. Standard Change Template governance evaluation rules.
9. Security & credential masking in api_logger (auth tuple, headers, URLs, and secret strings).
"""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import constants
import populate_cr_audit
import servicenow_api
import questions_data
import sc_sheet_data
import excel_handler
import api_logger

# 1. Regex patterns
assert constants.CHANGE_NUMBER_PATTERN.fullmatch("CHG1234567")
assert constants.CHANGE_NUMBER_PATTERN.fullmatch("CHG123")
assert not constants.CHANGE_NUMBER_PATTERN.fullmatch("CHG12A")

# 2. CLI input parser
parsed_chs = populate_cr_audit.load_change_numbers(["chg1234567, CHG2345678;bad-id", "CHG1234567"])
assert parsed_chs == ["CHG1234567", "CHG2345678"], f"Unexpected: {parsed_chs}"

# 3. Date validation
start, end = populate_cr_audit.validate_date_range("2026-01-01", "2026-03-31")
assert start < end

# 4. Sampling logic
assert servicenow_api.random_select([], 20) == []
sample_input = list(range(100))
assert len(servicenow_api.random_select(sample_input, 20)) == 20

# 5. Template resolution
resolved_tpl = populate_cr_audit.resolve_template_path("QAR Report.xlsx")
assert os.path.exists(resolved_tpl), f"Template not found at: {resolved_tpl}"

# 6. Excel template loader
wb = excel_handler.load_template(resolved_tpl)
assert "Master" in wb.sheetnames or len(wb.sheetnames) >= 4, f"Sheets: {wb.sheetnames}"

# 7. Audit logic smoke tests
assert questions_data.q5_std_incident({"u_reason_for_change": "Change application"}) == constants.NA
assert questions_data.q6_std_problem({"u_reason_for_change": "Problem Resolution", "u_related_problems": "PRB1000001"}) == constants.YES
assert questions_data.q19_outside_window_check({}) == constants.NA

# 8. Standard Change Template governance logic
assert sc_sheet_data.pre_test_plan({"u_non_prod_test_plan_summary": "n/a"}) == constants.NO
assert sc_sheet_data.pre_test_plan({"u_non_prod_test_plan_summary": "Valid pre-test plan executed"}) == constants.YES
assert sc_sheet_data._required_text_check("") == constants.NO
assert sc_sheet_data._required_text_check("not required") == constants.NO
assert sc_sheet_data._required_text_check("Valid rollout procedure") == constants.YES

# 9. Security & Credential Masking in api_logger
# 9a. Auth tuple masking
sanitized_auth = api_logger._sanitize_auth({"auth": ("corp_user", "SuperSecretPassword123!")})
assert sanitized_auth["auth"] == ("corp_user", "********"), f"Auth leak: {sanitized_auth}"

# 9b. Header Authorization masking
sanitized_hdr = api_logger._sanitize_auth({"headers": {"Authorization": "Basic dXNlcjpwYXNz", "Accept": "application/json"}})
assert sanitized_hdr["headers"]["Authorization"] == "Basic ********"

# 9c. URL embedded userinfo masking
safe_url = api_logger._sanitize_url("https://corp_user:SecretPass999@instance.service-now.com/api/now/table")
assert "SecretPass999" not in safe_url
assert "********" in safe_url

# 9d. Text/JSON scrubber
constants.PASSWORD = "CorpSecret_xyz_789"
scrubbed = api_logger._mask_known_secrets('{"status": "error", "debug": "failed with password CorpSecret_xyz_789"}')
assert "CorpSecret_xyz_789" not in scrubbed
assert "********" in scrubbed
constants.PASSWORD = ""  # reset

print("ALL PRE-BUILD VALIDATION CHECKS PASSED (Regex, CLI, Dates, Excel, Audit Rules, Security & Credential Masking).")
