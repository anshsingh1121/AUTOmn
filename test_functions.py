"""Manual smoke-test harness adapted from the screenshots.

This script requires valid ServiceNow credentials and is intentionally not run
as part of a local syntax-only validation.
"""
import requests
from urllib3.exceptions import InsecureRequestWarning
import constants
from questions_data import (
    q17_post_validation_check, q9_approver_check, q15_approval_timing,
    q14_std_approval, q3_nc_exp_business_project_patching,
    q2_nc_exp_sdlc_ritm, q4_reason_vs_sdlc, q1_std_testing_completed,
    q1_std_testing_completed_new, q17_post_validation_check_new,
)

requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)


def get_changes_by_numbers(change_numbers):
    query_params = {
        "sysparm_query": "numberIN" + ",".join(change_numbers),
        "sysparm_display_value": "true",
    }
    response = requests.get(
        constants.BASE_URL + constants.CHANGE_REQUEST_TABLE,
        auth=(constants.USERNAME, constants.PASSWORD),
        headers=constants.headers, params=query_params,
        verify=constants.VERIFY_SSL, timeout=30,
    )
    response.raise_for_status()
    results = response.json().get("result", [])
    changes_by_number = {change.get("number"): change for change in results if change.get("number")}
    return [changes_by_number.get(number) for number in change_numbers]


if __name__ == "__main__":
    # Replace/extend with approved test CR numbers before running.
    change_numbers = ["CHG6892954", "CHG6957477", "CHG6738045", "CHG6908979"]
    changes = get_changes_by_numbers(change_numbers)
    for index, (change_number, change) in enumerate(zip(change_numbers, changes), start=1):
        if change is None:
            result = "NOT FOUND"
        else:
            result = q17_post_validation_check_new(change)
        print(f"{index}, {change_number}: {result}")
