"""Local, no-network validation checks for the reconstructed project."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
import constants
import populate_cr_audit
import servicenow_api
import questions_data

assert constants.CHANGE_NUMBER_PATTERN.fullmatch("CHG1234567")
assert constants.CHANGE_NUMBER_PATTERN.fullmatch("CHG123")
assert not constants.CHANGE_NUMBER_PATTERN.fullmatch("CHG12A")
assert populate_cr_audit.load_change_numbers(["chg1234567, CHG2345678;bad-id", "CHG1234567"]) == ["CHG1234567", "CHG2345678"]
start, end = populate_cr_audit.validate_date_range("2026-01-01", "2026-03-31")
assert start < end
assert servicenow_api.random_select([], 20) == []
assert questions_data.q5_std_incident({"u_reason_for_change": "Change application"}) == constants.NA
assert questions_data.q6_std_problem({"u_reason_for_change": "Problem Resolution", "u_related_problems": "PRB1000001"}) == constants.YES
assert questions_data.q19_outside_window_check({}) == constants.NA
print("PASS: import checks, change-number parser, date validation, selection edge case, and local audit-rule smoke tests.")
