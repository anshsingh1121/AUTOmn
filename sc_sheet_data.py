"""Checks specific to standard-change template data."""
from __future__ import annotations
from datetime import datetime
import requests
import constants


def _text(value):
    if isinstance(value, dict):
        return str(value.get("display_value") or "")
    return str(value or "")


def is_retired(change):
    ci = change.get("cmdb_ci") or {}
    link = ci.get("link") if isinstance(ci, dict) else None
    if not link:
        return constants.NA
    response = requests.get(link, auth=(constants.USERNAME, constants.PASSWORD),
                            headers=constants.headers, params={"sysparm_display_value": "true"},
                            verify=constants.VERIFY_SSL, timeout=30)
    response.raise_for_status()
    result = response.json().get("result", {})
    operational_status = _text(result.get("operational_status"))
    if not operational_status:
        return constants.NA
    return constants.YES if operational_status.casefold() == "retired" else constants.NO


def _required_text_check(value):
    text = _text(value)
    if not text.strip():
        return constants.NO
    folded = text.casefold()
    if any(keyword.casefold() in folded for keyword in constants.FAIL_KEYWORDS):
        return constants.NO
    return constants.YES


def pre_test_plan(cr_template):
    summary = _text(cr_template.get("u_non_prod_test_plan_summary"))
    if not summary.strip():
        return constants.NO
    if any(keyword.casefold() in summary.casefold()
           for keyword in constants.STD_TEST_FAIL_KEYWORDS):
        return constants.NO
    return constants.YES


def deployment_plan(cr_template):
    return _required_text_check(cr_template.get("implementation_plan"))


def backout_plan(cr_template):
    return _required_text_check(cr_template.get("backout_plan"))


def verification_plan(cr_template):
    return _required_text_check(cr_template.get("u_post_prod_deployment_test_plan_summary"))


def _parse_date(value):
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(str(value)[:19], fmt)
        except ValueError:
            pass
    return None


def fetch_cr_close_codes(change_number):
    params = {"sysparm_query": f"number={change_number}", "sysparm_display_value": "true"}
    response = requests.get(constants.BASE_URL + constants.CHANGE_REQUEST_TABLE,
                            auth=(constants.USERNAME, constants.PASSWORD),
                            headers=constants.headers, params=params,
                            verify=constants.VERIFY_SSL, timeout=30)
    response.raise_for_status()
    rows = response.json().get("result", [])
    if rows and str(rows[0].get("state", "")).casefold() == "closed":
        return rows[0].get("close_code")
    return constants.NA


def sample_changes(proposal_result):
    """Check prior changes referenced by a standard proposal.

    The screenshot showed a cutoff date of 2025-09-08, a minimum of one
    previous change for Modify proposals, and four for New proposals.
    """
    proposal_type = _text(proposal_result.get("proposal_type"))
    change_requests = _text(proposal_result.get("change_requests"))
    created_date = _parse_date(proposal_result.get("sys_created_on"))
    cutoff_date = datetime(2025, 9, 8)
    if created_date is not None and created_date < cutoff_date:
        return constants.NA
    changes = [c.strip() for c in change_requests.split(",") if c.strip()]
    if proposal_type.startswith("Modify") and len(changes) >= 1:
        required = 1
    elif proposal_type.startswith("New") and len(changes) >= 4:
        required = 4
    else:
        return constants.NO
    codes = {change: fetch_cr_close_codes(change) for change in changes[:max(required, len(changes))]}
    return constants.YES if all(str(code or "").casefold() == "successful" for code in codes.values()) else constants.NO
