"""ServiceNow data-access layer used by the audit workflow."""
from __future__ import annotations
import math
import random
from datetime import datetime
import requests
import constants
import questions_data
import sc_sheet_data


def _get(url, *, params=None):
    response = requests.get(url, auth=(constants.USERNAME, constants.PASSWORD),
                            headers=constants.headers, params=params,
                            verify=constants.VERIFY_SSL, timeout=30)
    response.raise_for_status()
    return response.json()


def fetch_crs_from_servicenow(start_date=None, end_date=None, change_type=None, change_numbers=None):
    """Fetch CRs either by exact numbers or by date range and type."""
    if change_numbers:
        normalized_numbers = [str(n).strip().upper() for n in change_numbers if str(n).strip()]
        if not normalized_numbers:
            return []
        query = "numberIN" + ",".join(normalized_numbers)
        print(f"Fetching {len(normalized_numbers)} requested Change Requests")
    else:
        if not start_date or not end_date or not change_type:
            raise ValueError("start_date, end_date, and change_type are required for date-range mode")
        query = (
            f"type={change_type}"
            f"^end_dateBETWEENjavascript:gs.dateGenerate('{start_date}','00:00:00')"
            f"@javascript:gs.dateGenerate('{end_date}','23:59:59')"
            "^stateIN-2,-1,0,3^production_system=true"
        )
        print(f"Fetching {change_type} CRs from ServiceNow for period: {start_date} to {end_date}")
    data = _get(constants.BASE_URL + constants.CHANGE_REQUEST_TABLE,
                params={"sysparm_query": query, "sysparm_display_value": "true"})
    return data.get("result", [])


def random_select(crs, percentage):
    if not crs:
        return []
    count = max(1, int(math.ceil(len(crs) * (percentage / 100.0))))
    return random.sample(crs, min(count, len(crs)))


def get_proposals(change):
    template_version = change.get("std_change_producer_version") or {}
    link = template_version.get("link") if isinstance(template_version, dict) else None
    if not link:
        return {}
    template_result = _get(link, params={"sysparm_display_value": "true"}).get("result", {})
    proposal_link = (template_result.get("std_change_proposal") or {}).get("link")
    if not proposal_link:
        return {}
    return _get(proposal_link, params={"sysparm_display_value": "true"}).get("result", {})


def get_basic_data(tester, cur_month, full_len, sample_len):
    return [{"test_month": str(cur_month), "sample_month": str(cur_month),
             "tester": tester, "population": full_len, "samples_tested": sample_len}]


def get_audit_data(changes, sheet_name):
    audit_data = []
    for change in changes:
        print(f"Fetching audit data for {change.get('type')} CR {change.get('number')}")
        common_data = get_common_data(change)
        merged = common_data | get_questions_data(change)
        if sheet_name == "SC Template Result Sheet":
            merged = merged | get_std_questions_data(change)
        audit_data.append(merged)
    return audit_data


def get_audit_data_single(change, sheet_name=None):
    common_data = get_common_data(change)
    merged = common_data | get_questions_data(change)
    if sheet_name == "SC Template Result Sheet":
        merged = merged | get_std_questions_data(change)
    return merged


def get_all_related_incidents(change):
    change_no = change.get("number")
    related_inc = []
    raw = change.get("u_related_incidents")
    if isinstance(raw, str):
        related_inc.extend([x.strip() for x in raw.split(",") if x.strip()])
    elif isinstance(raw, list):
        related_inc.extend([_display(x) for x in raw])
    query_params = {"sysparm_query": f"caused_by.number={change_no}",
                    "sysparm_display_value": "true", "sysparm_fields": "number"}
    response = _get(constants.BASE_URL + constants.INCIDENT_TABLE, params=query_params)
    for incident in response.get("result", []):
        related_inc.append(incident.get("number"))
    sys_id = change.get("sys_id")
    if sys_id:
        query_params = {"sysparm_query": f"rfcISNOTEMPTY^rfc={sys_id}",
                        "sysparm_display_value": "true", "sysparm_fields": "number"}
        response = _get(constants.BASE_URL + constants.INCIDENT_TABLE, params=query_params)
        related_inc.extend([x.get("number") for x in response.get("result", [])])
    related_inc = sorted(set(filter(None, related_inc)))
    return ", ".join(related_inc) if related_inc else constants.NA


def get_all_related_problems(change):
    change_no = change.get("number")
    related_prb = []
    raw = change.get("u_related_problems")
    if isinstance(raw, str):
        related_prb.extend([x.strip() for x in raw.split(",") if x.strip()])
    elif isinstance(raw, list):
        related_prb.extend([_display(x) for x in raw])
    params = {"sysparm_query": f"rfc.number={change_no}",
              "sysparm_display_value": "true", "sysparm_fields": "number"}
    response = _get(constants.BASE_URL + constants.PROBLEM_TABLE, params=params)
    related_prb.extend([x.get("number") for x in response.get("result", [])])
    related_prb = sorted(set(filter(None, related_prb)))
    return ", ".join(related_prb) if related_prb else constants.NA


def get_common_data(change):
    return {
        "cr_number": change.get("number"), "test_date": datetime.today(),
        "change_type": _display(change.get("type")), "category": _display(change.get("category")),
        "closure_status": change.get("close_code") or questions_data.get_closure_status(change),
        "related_inc": get_all_related_incidents(change),
        "related_prb": get_all_related_problems(change), "jira_issue": "",
    }


def get_questions_data(change):
    return {
        "q1_testing_completed": questions_data.q1_std_testing_completed(change),
        "q1_testing_completed_new": questions_data.q1_std_testing_completed_new(change),
        "q2_nc_exp_sdlc_ritm": questions_data.q2_nc_exp_sdlc_ritm(change),
        "q3_nc_exp_business_project_patching": questions_data.q3_nc_exp_business_project_patching(change),
        "q4_reason_vs_sdlc": questions_data.q4_reason_vs_sdlc(change),
        "q5_inc_assoc": questions_data.q5_std_incident(change),
        "q6_prb_assoc": questions_data.q6_std_problem(change),
        "q7_std_template": questions_data.q7_std_template(change),
        "q8_differentiation": questions_data.q8_differentiation(change),
        "q9_approver_check": questions_data.q9_approver_check(change),
        "q10_emergency_approval": questions_data.q10_emergency_approval_check(change),
        "q11_fcb_expedited": questions_data.q11_expedited_fcb_approval_check(change),
        "q12_normal_low": questions_data.q12_normal_fcb_lowrisk_approval_check(change),
        "q13_normal_high": questions_data.q13_normal_fcb_highrisk_approval_check(change),
        "q14_std_approval": questions_data.q14_std_approval(change),
        "q15_approval_timing": questions_data.q15_approval_timing(change),
        "q16_major_incident": questions_data.q16_major_incident_check(change),
        "q17_post_validation": questions_data.q17_post_validation_check(change),
        "q17_post_validation_new": questions_data.q17_post_validation_check_new(change),
        "q18_issues_doc": questions_data.q18_success_issues_check(change),
        "q19_outside_window": questions_data.q19_outside_window_check(change),
    }


def get_std_questions_data(change):
    proposal_result = get_proposals(change)
    proposal = proposal_result.get("number")
    template_blob = proposal_result.get("template_value") or ""
    cr_template = {}
    for element in str(template_blob).split("*"):
        pair = element.split("=", 1)
        if len(pair) > 1:
            cr_template[pair[0]] = pair[1]
            if pair[0] == "u_reason_for_change" and not pair[1]:
                print(f"{proposal} — Reason for change not found.")
    version = change.get("std_change_producer_version") or {}
    return {
        "cr_number": change.get("number"), "test_date": datetime.today(),
        "template_ver": _display(version), "std_proposal": proposal,
        "category": _display(change.get("category")),
        "closure_status": questions_data.get_closure_status(change),
        "related_inc": get_all_related_incidents(change), "related_prb": get_all_related_problems(change),
        "jira_issue": "", "is_retired": sc_sheet_data.is_retired(change),
        "pre_test_plan": sc_sheet_data.pre_test_plan(cr_template),
        "deployment_plan": sc_sheet_data.deployment_plan(cr_template),
        "backout_plan": sc_sheet_data.backout_plan(cr_template),
        "verification_plan": sc_sheet_data.verification_plan(cr_template),
        "sample_changes": sc_sheet_data.sample_changes(proposal_result),
    }


def _display(value):
    if isinstance(value, dict):
        return value.get("display_value") or value.get("value") or ""
    return value or ""
