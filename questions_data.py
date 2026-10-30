"""Audit-question logic recovered from screenshots of the original module.

The original file was long (roughly 1,400 visible editor lines). This file
recreates the visible functions and decision paths in a compact, auditable
implementation. A few checks that depended on unreadable lines are marked as
best-effort; see RECONSTRUCTION_REPORT.md and reconstruction_evidence/.
"""
from __future__ import annotations
import os
import re
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
import requests
import constants

try:
    from thefuzz import fuzz
except Exception:  # The screenshots also show RapidFuzz installed.
    from rapidfuzz import fuzz

try:
    import extract_msg
except Exception:
    extract_msg = None


def _display(value):
    if isinstance(value, dict):
        return str(value.get("display_value") or value.get("value") or "")
    return str(value or "")


def _truth(value):
    return _display(value).strip().casefold() in {"true", "yes", "1", "y"}


def _get(url, params=None):
    response = requests.get(url, auth=(constants.USERNAME, constants.PASSWORD),
                            headers=constants.headers, params=params,
                            verify=constants.VERIFY_SSL, timeout=30)
    response.raise_for_status()
    return response.json()


def _query(table, query, fields=None):
    params = {"sysparm_query": query, "sysparm_display_value": "true"}
    if fields:
        params["sysparm_fields"] = fields
    return _get(constants.BASE_URL + table, params=params).get("result", [])


def get_closure_status(change):
    if _display(change.get("state")).casefold() == "closed":
        return change.get("close_code") or constants.NA
    return ""


def _normalize_name(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").casefold()).strip()


def _has_fail_keyword(value):
    folded = str(value or "").casefold()
    return any(key.casefold() in folded for key in constants.FATL_KEYWORDS)


def _attachments(change):
    sys_id = change.get("sys_id")
    if not sys_id:
        return []
    rows = _query(constants.CHANGE_ATTACH_TABLE, f"table_sys_id={sys_id}",
                  fields="sys_id,file_name,content_type,size_bytes,table_name")
    return rows


def _matching_attachment(change, keywords):
    for attachment in _attachments(change):
        name = _normalize_name(attachment.get("file_name"))
        for keyword in keywords:
            if _normalize_name(keyword) in name:
                return attachment
    return None


def q1_std_testing_completed(change):
    """Q1: Was testing completed/documented and were artifacts attached?"""
    not_testable = change.get("u_change_cannot_be_pre_tested")
    if _truth(not_testable):
        if _display(change.get("u_why_can_this_not_be_pre_tested")).strip():
            return constants.NON_TESTABLE
        return constants.NO
    test_summary = change.get("u_non_prod_test_plan_summary")
    test_results = change.get("u_non_prod_test_plan_results")
    if not _display(test_summary).strip() or not _display(test_results).strip():
        return constants.NO
    # Screenshots showed filename normalization (punctuation/underscores to spaces)
    # and matching against PRE_PROD_KEYWORDS in sys_attachment.
    if _matching_attachment(change, constants.PRE_PROD_KEYWORDS):
        return constants.YES
    return constants.NO


def q1_std_testing_completed_new(change):
    """Q1 alternate/revised check, retained as a separate output column."""
    # The new variant was visible in the screenshots, but not every condition was
    # legible. Keep the same prerequisites and accept a clearly named test artifact.
    not_testable = change.get("u_change_cannot_be_pre_tested")
    if _truth(not_testable):
        return constants.NON_TESTABLE if _display(change.get("u_why_can_this_not_be_pre_tested")).strip() else constants.NO
    if not _display(change.get("u_non_prod_test_plan_summary")).strip():
        return constants.NO
    if not _display(change.get("u_non_prod_test_plan_results")).strip():
        return constants.NO
    names = [_normalize_name(a.get("file_name")) for a in _attachments(change)]
    evidence_words = ["test result", "test results", "test evidence", "test cycle report", "non prod"]
    return constants.YES if any(any(w in name for w in evidence_words) for name in names) else constants.NO


def q2_nc_exp_sdlc_ritm(change):
    """Q2: For an untestable normal/expedited CR, is an SDLC SII/RITM identified?"""
    ctype = _display(change.get("type")).casefold()
    if ctype not in {"normal", "expedited"}:
        return constants.NA
    if not _truth(change.get("u_change_cannot_be_pre_tested")):
        return constants.NA
    sdlc_type = _display(change.get("u_sdlc_initiative_type")).casefold()
    not_applicable = {"dr exercise - initiative not applicable", "maintenance - initiative not applicable"}
    if sdlc_type in not_applicable:
        return constants.NA
    summary = _display(change.get("u_non_prod_test_plan_summary"))
    if not summary.strip() or _has_fail_keyword(summary):
        return constants.NO
    if "sii" not in summary.casefold() and "ritm" not in summary.casefold():
        return constants.NO
    return constants.YES


def q3_nc_exp_business_project_patching(change):
    """Q3: Validate business-project/patching reason for normal/expedited changes."""
    ctype = _display(change.get("type")).casefold()
    if ctype not in {"normal", "expedited"}:
        return constants.NA
    reason = _display(change.get("u_reason_for_change")).casefold()
    short_description = _display(change.get("short_description")).casefold()
    if not reason:
        return constants.NO
    if any(word in reason for word in ("business project", "patching", "project")):
        return constants.YES if short_description.strip() else constants.NO
    return constants.NA


def q4_reason_vs_sdlc(change):
    """Q4: Check reason-for-change alignment with SDLC initiative type."""
    ctype = _display(change.get("type")).casefold()
    if ctype not in {"normal", "expedited"}:
        return constants.NA
    reason = _display(change.get("u_reason_for_change")).casefold()
    sdlc_type = _display(change.get("u_sdlc_initiative_type")).casefold()
    if reason == "maintenance":
        return constants.YES if sdlc_type == "maintenance - initiative not applicable" else constants.MM
    if reason == "dr exercise":
        return constants.YES if sdlc_type == "dr exercise - initiative not applicable" else constants.DM
    return constants.NA


def q5_std_incident(change):
    """Q5: If the reason is incident resolution, is an incident associated?"""
    reason = _display(change.get("u_reason_for_change"))
    if reason and reason.casefold() != "incident resolution":
        return constants.NA
    return constants.YES if _display(change.get("u_related_incidents")).strip() else constants.NO


def q6_std_problem(change):
    """Q6: If the reason is problem resolution, is a problem associated?"""
    reason = _display(change.get("u_reason_for_change"))
    if reason and reason.casefold() != "problem resolution":
        return constants.NA
    return constants.YES if _display(change.get("u_related_problems")).strip() else constants.NO


def q7_std_template(change):
    """Q7: Is the standard change template applicable to the change?"""
    if _display(change.get("type")).casefold() != constants.STANDARD.casefold():
        return constants.NA
    template = change.get("std_change_producer_version") or {}
    template_name = _display(template)
    if not template_name:
        return constants.NO
    if any(template_name.casefold().startswith(t.casefold()) or t.casefold() in template_name.casefold()
           for t in constants.STANDARD_CH_APPROVAL_EXCEPTION_TEMPLATES):
        return constants.NA
    proposal = _get_proposal_template_value(change)
    short_description = _display(change.get("short_description")).casefold()
    proposal_short_description = proposal.get("short_description", "").casefold()
    if proposal_short_description:
        return constants.YES if short_description == proposal_short_description else constants.NO
    return constants.YES if template_name else constants.NO


def _get_proposal_template_value(change):
    """Fetch the template proposal record; fail closed when unavailable."""
    try:
        from servicenow_api import get_proposals
        proposal = get_proposals(change)
        blob = proposal.get("template_value") or ""
        values = {}
        for element in str(blob).split("*"):
            pair = element.split("=", 1)
            if len(pair) == 2:
                values[pair[0]] = pair[1]
        return values
    except (requests.RequestException, ValueError, TypeError):
        return {}


def q8_differentiation(change):
    """Q8: Can the change be differentiated from others using the same template?"""
    if _display(change.get("type")).casefold() != constants.STANDARD.casefold():
        return constants.NA
    template = change.get("std_change_producer_version") or {}
    template_name = _display(template)
    description = _display(change.get("short_description"))
    if not description or not template_name:
        return constants.NO
    similarity = fuzz.token_sort_ratio(description, template_name)
    return constants.YES if similarity >= 40 else constants.NO


def get_actual_change_approvers(change):
    number = _display(change.get("number"))
    results = _query(constants.SYS_APPROVAL_TABLE,
                     f"sysapproval.number={number}^state=approved",
                     fields="approver,sys_created_on,sys_updated_on,state")
    return [_display(row.get("approver")) for row in results if _display(row.get("approver"))]


def _approval_rows(change):
    return _query(constants.SYS_APPROVAL_TABLE,
                  f"sysapproval.number={_display(change.get('number'))}^state=approved",
                  fields="approver,sys_created_on,sys_updated_on,state")


def _group_record(group_ref):
    link = group_ref.get("link") if isinstance(group_ref, dict) else None
    if link:
        return _get(link, {"sysparm_display_value": "true"}).get("result", {})
    name = _display(group_ref)
    if name:
        rows = _query(constants.SYS_USER_GROUP_TABLE, f"name={name}", fields="sys_id,name,u_group_slt,parent")
        return rows[0] if rows else {}
    return {}


def _group_members(group_sys_id):
    if not group_sys_id:
        return []
    rows = _query(constants.SYS_USER_GROUP_MEM_TABLE, f"group={group_sys_id}", fields="user")
    return [_display(row.get("user")) for row in rows if _display(row.get("user"))]


def _delegates_for(user_names):
    delegates = set()
    for name in user_names:
        if not name:
            continue
        rows = _query(constants.SYS_USER_DELEGATE_TABLE, f"user.name={name}", fields="delegate,starts,ends")
        delegates.update(_display(row.get("delegate")) for row in rows if _display(row.get("delegate")))
    return delegates


def q9_approver_check(change):
    """Q9: Change initiator/requester and assignee should not be the sole approver."""
    approvals = get_actual_change_approvers(change)
    if not approvals:
        return constants.NO
    requester = _display(change.get("requested_by"))
    assignee = _display(change.get("assigned_to"))
    excluded = {x.casefold() for x in (requester, assignee) if x}
    distinct_approvers = {x.casefold() for x in approvals if x}
    if distinct_approvers.isdisjoint(excluded):
        return constants.YES
    # Screenshots mention checking multiple approvers other than requester/assignee.
    return constants.YES if distinct_approvers - excluded else constants.NO


def q10_emergency_approval_check(change):
    """Q10: Emergency CR must have SLT approval and an approval-group member."""
    if _display(change.get("type")).casefold() != constants.EMERGENCY.casefold():
        return constants.NA
    group = _group_record(change.get("assignment_group") or {})
    group_id = group.get("sys_id")
    slt_name = _display(group.get("u_group_slt"))
    approvals = {x.casefold() for x in get_actual_change_approvers(change)}
    members = {x.casefold() for x in _group_members(group_id)}
    slt_approved = bool(slt_name) and slt_name.casefold() in approvals
    member_approved = bool(members & approvals)
    return constants.YES if slt_approved and member_approved else constants.MA


def q11_expedited_fcb_approval_check(change):
    """Q11: Expedited CR approval check (ITCM, CI owner, assignment-group SLT)."""
    if _display(change.get("type")).casefold() != constants.EXPEDITED.casefold():
        return constants.NA
    approvals = {x.casefold() for x in get_actual_change_approvers(change)}
    assignment_group = _group_record(change.get("assignment_group") or {})
    slt_name = _display(assignment_group.get("u_group_slt"))
    ci_name = _display(change.get("cmdb_ci"))
    ci_owner = ""
    ci = change.get("cmdb_ci") or {}
    link = ci.get("link") if isinstance(ci, dict) else None
    if link:
        ci_result = _get(link, {"sysparm_display_value": "true"}).get("result", {})
        ci_owner = _display(ci_result.get("owned_by"))
    itcm_members = set()
    for row in _query(constants.SYS_USER_GROUP_TABLE, "name=Change Management Approval", fields="sys_id"):
        itcm_members.update(_group_members(row.get("sys_id")))
    checks = [slt_name.casefold() in approvals if slt_name else False,
              ci_owner.casefold() in approvals if ci_owner else False,
              bool({x.casefold() for x in itcm_members} & approvals)]
    return constants.YES if all(checks) else constants.MA


def _risk(change):
    return _display(change.get("risk")).casefold()


def q12_normal_fcb_lowrisk_approval_check(change):
    if _display(change.get("type")).casefold() != constants.NORMAL.casefold() or _risk(change) != "low":
        return constants.NA
    approvals = {x.casefold() for x in get_actual_change_approvers(change)}
    if not approvals:
        return constants.MA
    group = _group_record(change.get("assignment_group") or {})
    members = {x.casefold() for x in _group_members(group.get("sys_id"))}
    return constants.YES if approvals & members else constants.MA


def q13_normal_fcb_highrisk_approval_check(change):
    if _display(change.get("type")).casefold() != constants.NORMAL.casefold() or _risk(change) not in {"high", "moderate"}:
        return constants.NA
    approvals = {x.casefold() for x in get_actual_change_approvers(change)}
    if not approvals:
        return constants.MA
    group = _group_record(change.get("assignment_group") or {})
    slt_name = _display(group.get("u_group_slt"))
    required = {slt_name.casefold()} if slt_name else set()
    requested_by = _display(change.get("requested_by"))
    assigned_to = _display(change.get("assigned_to"))
    required -= {x.casefold() for x in (requested_by, assigned_to) if x}
    return constants.YES if required and required <= approvals else constants.MA


def q14_std_approval(change):
    if _display(change.get("type")).casefold() != constants.STANDARD.casefold():
        return constants.NA
    template_name = _display(change.get("std_change_producer_version"))
    if any(template_name.casefold().startswith(x.casefold()) for x in constants.STANDARD_CH_APPROVAL_EXCEPTION_TEMPLATES):
        return constants.NA
    approvals = {x.casefold() for x in get_actual_change_approvers(change)}
    if not approvals:
        return constants.MA
    assignment_group = _group_record(change.get("assignment_group") or {})
    members = {x.casefold() for x in _group_members(assignment_group.get("sys_id"))}
    requester = _display(change.get("requested_by")).casefold()
    assignee = _display(change.get("assigned_to")).casefold()
    valid = (approvals & members) - {requester, assignee}
    return constants.YES if valid else constants.MA


def _parse_datetime(value):
    text = _display(value)
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y"):
        try:
            return datetime.strptime(text[:19], fmt)
        except ValueError:
            continue
    return None


def q15_approval_timing(change):
    """Q15: Actual start should not precede the last approval."""
    rows = _approval_rows(change)
    if not rows:
        return constants.MA
    start = _parse_datetime(change.get("actual_start") or change.get("work_start") or change.get("start_date"))
    approvals = [_parse_datetime(r.get("sys_updated_on") or r.get("sys_created_on")) for r in rows]
    approvals = [x for x in approvals if x is not None]
    if not start or not approvals:
        return constants.NA
    last_approval = max(approvals)
    return constants.YES if start >= last_approval else constants.NO


def q16_major_incident_check(change):
    """Q16: A major incident associated with the change must be traceable."""
    sys_id = change.get("sys_id")
    if not sys_id:
        return constants.NA
    rows = _query(constants.INCIDENT_TABLE, f"caused_by={sys_id}^major_incident=true", fields="number,major_incident")
    if rows:
        return constants.YES
    return constants.NO


def q17_post_validation_check(change):
    """Q17: Search for a post-production validation artifact by attachment name."""
    names = [_normalize_name(x.get("file_name")) for x in _attachments(change)]
    if any(_normalize_name(k) in name for name in names for k in constants.POST_PROD_KEYWORDS):
        return constants.YES
    # The original logic used a KWNF code if no matching keyword was found.
    return constants.KWNF


def q17_post_validation_check_new(change):
    """Q17 revised: inspect the post-implementation review artefact."""
    attachments = _attachments(change)
    for item in attachments:
        name = _normalize_name(item.get("file_name"))
        if "implementation review" in name or ("post" in name and "validation" in name):
            return constants.YES
    # Some screenshots showed attachment download and Outlook .msg parsing.
    # If a review message exists but cannot be parsed, avoid claiming pass.
    return constants.KWNF


def q18_success_issues_check(change):
    """Q18: Documentation for implementation issues is available when needed."""
    state = _display(change.get("state")).casefold()
    if state != "closed":
        return constants.NA
    close_code = _display(change.get("close_code")).casefold()
    if close_code == "successful":
        return constants.NA
    return constants.YES if _matching_attachment(change, constants.ISSUE_DOCUMENTATION_KEYWORDS) else constants.NO


def q19_outside_window_check(change):
    """Q19: Compare actual execution timestamps with the approved change window."""
    planned_start = _parse_datetime(change.get("start_date"))
    planned_end = _parse_datetime(change.get("end_date"))
    actual_start = _parse_datetime(change.get("actual_start") or change.get("work_start"))
    actual_end = _parse_datetime(change.get("actual_end") or change.get("work_end"))
    if not planned_start or not planned_end or not actual_start:
        return constants.NA
    if actual_start < planned_start or actual_start > planned_end:
        return constants.NO
    if actual_end and (actual_end < planned_start or actual_end > planned_end):
        return constants.NO
    return constants.YES


def download_attachment(attachment_sys_id, save_path):
    """Download an attachment using the ServiceNow attachment API."""
    url = f"{constants.ATT_FILE_URL}{attachment_sys_id}/file"
    response = requests.get(url, auth=(constants.USERNAME, constants.PASSWORD),
                            headers=constants.headers, verify=constants.VERIFY_SSL, timeout=60)
    response.raise_for_status()
    Path(save_path).write_bytes(response.content)
    return save_path


def analyze_dashboard_status(change):
    """A tiny utility retained because it was referenced in the screenshots."""
    return {"number": change.get("number"), "state": _display(change.get("state")),
            "close_code": _display(change.get("close_code"))}
