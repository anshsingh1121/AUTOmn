# Change Management Audit Automation — Complete Codebase Documentation

> **Document purpose:** Complete, simple-to-follow technical documentation of the codebase reconstructed from the supplied ZIP containing 93 screenshots of the source code and workbook.
>
> **Source fidelity:** This document describes the implementation visible in the screenshots. Where a screenshot was partially obscured or OCR was uncertain, the behavior is explicitly marked instead of inventing code that was not visible.

---

## 1. Executive Summary

This project is a **ServiceNow Change Management Audit Automation** tool.

Its job is to:

```text
Take ServiceNow Change Requests
        ↓
Fetch the required ServiceNow data
        ↓
Select which Changes must be audited
        ↓
Run predefined audit/control rules
        ↓
Inspect ServiceNow attachments when evidence is required
        ↓
Use local OCR for dashboard/image evidence
        ↓
Produce YES / NO / NA / ERROR audit results
        ↓
Populate a predefined Excel QAR report
```

The important point is:

> **This is a deterministic audit/rules engine, not a machine-learning prediction system.**

The code mainly uses:

- ServiceNow REST APIs
- Python business rules
- `openpyxl` for Excel
- `RapidFuzz`/`thefuzz` for one text-similarity check
- OpenCV + Tesseract OCR for image evidence
- local file/report handling

There is no visible trained ML model, model training pipeline, neural network, or learned classifier in the supplied screenshots.

---

# 2. Project Mental Model

The easiest way to remember the whole project is:

```text
                         ┌─────────────────────────────┐
                         │   populate_cr_audit.py      │
                         │      MAIN ORCHESTRATOR      │
                         └──────────────┬──────────────┘
                                        │
                         ┌──────────────▼──────────────┐
                         │      servicenow_api.py      │
                         │      GET DATA FROM SNOW     │
                         └──────────────┬──────────────┘
                                        │
                 ┌──────────────────────┴──────────────────────┐
                 │                                             │
       ┌─────────▼─────────┐                         ┌──────────▼──────────┐
       │ questions_data.py │                         │  sc_sheet_data.py   │
       │  Q1 → Q19 RULES  │                         │ STANDARD-TEMPLATE   │
       │   AUDIT ENGINE    │                         │      RULES           │
       └─────────┬─────────┘                         └──────────┬──────────┘
                 │                                             │
                 └──────────────────┬──────────────────────────┘
                                    │
                         ┌──────────▼──────────┐
                         │      search.py      │
                         │ OCR / IMAGE EVIDENCE│
                         └──────────┬──────────┘
                                    │
                         ┌──────────▼──────────┐
                         │    excel_handler.py │
                         │    REPORT WRITER    │
                         └──────────┬──────────┘
                                    │
                                    ▼
                              QAR_report.xlsx
```

### One-line explanation of each core file

| File | Simple meaning |
|---|---|
| `populate_cr_audit.py` | Decides **what to audit and in what order** |
| `servicenow_api.py` | Decides **how to get ServiceNow data** |
| `questions_data.py` | Decides **how Q1–Q19 are evaluated** |
| `sc_sheet_data.py` | Decides **how Standard Change Template checks are evaluated** |
| `search.py` | Converts image evidence into text/status using OCR |
| `excel_handler.py` | Puts all results into the QAR Excel template |
| `constants.py` | Stores shared API/table/keyword/configuration values |
| `test_functions.py` | Manually tests selected audit functions |
| `QAR_report.xlsx` | Final report template |
| `changes.txt` | Input list of Change Request numbers |
| `tesseract/` | Bundled local Tesseract OCR executable/data |
| `report/` | Generated report/evidence output area |

---

# 3. Repository / Folder Structure

The visible project contains roughly:

```text
project-root/
│
├── constants.py
├── excel_handler.py
├── populate_cr_audit.py
├── questions_data.py
├── sc_sheet_data.py
├── search.py
├── servicenow_api.py
├── test_functions.py
├── changes.txt
├── QAR_report.xlsx
│
├── report/
│   └── generated Excel reports / downloaded evidence
│
├── tesseract/
│   ├── tesseract.exe
│   └── tessdata/
│
├── .idea/
├── .venv/
└── __pycache__/
```

The screenshots also show the standard Python development artifacts such as `.idea`, `.venv`, and `__pycache__`.

---

# 4. Core Technologies

## 4.1 Python

The entire logic is implemented in Python.

## 4.2 ServiceNow REST API

The project communicates directly with ServiceNow tables using HTTP `GET` requests.

## 4.3 OpenPyXL

Used to open the predefined QAR workbook, write values into specific cells, and save the resulting report.

## 4.4 RapidFuzz / thefuzz

Used by Q8 to perform a token-sort fuzzy similarity comparison for Standard Changes.

Visible import:

```python
from thefuzz import fuzz
```

## 4.5 OpenCV

Used for image preprocessing before OCR.

## 4.6 Tesseract OCR

A local Tesseract installation is bundled under the project directory.

This is important because the code is not using a cloud vision API for the screenshot/dashboard evidence step.

## 4.7 Requests

Used to call ServiceNow and download attachments.

---

# 5. Main Functional Goal

The system automates a **Change Management Quality Assurance / audit process**.

Instead of manually opening each Change Request and checking every control, the tool:

1. Finds the relevant Change Requests.
2. Fetches the records and related data.
3. Determines which audit questions apply.
4. Evaluates each question.
5. Pulls attachment evidence when necessary.
6. Runs OCR on certain image evidence.
7. Writes the final results into the required workbook.

The report is structured around Change Management audit sheets such as:

- Normal Change result sheet
- Standard Change result sheet
- Emergency Change result sheet
- Expedited Change result sheet
- Standard Change Template result sheet
- Revision Control sheet

---

# 6. Input Modes

There are **two main ways to run the project**.

## 6.1 Explicit Change-number mode

You directly provide the Change Request numbers.

Examples:

```bash
python populate_cr_audit.py --change-nums CHG1234567 CHG1234568
```

Comma-separated values are supported:

```bash
python populate_cr_audit.py --change-nums CHG1234567,CHG1234568
```

A text file is also supported:

```bash
python populate_cr_audit.py --change-nums changes.txt
```

The code therefore supports:

```text
CHG number
CHG number CHG number CHG number
CHG number,CHG number,CHG number
text file containing CHG numbers
```

### What happens in explicit mode

```text
User provides CHG numbers
        ↓
Load + normalize + validate numbers
        ↓
Fetch those records from ServiceNow
        ↓
Audit each returned CR
        ↓
Route the CR to its appropriate Excel sheet
        ↓
Save change-wise report
```

---

## 6.2 Date-range / percentage mode

Instead of naming the CRs manually, the user provides a date range and a sample percentage.

Example:

```bash
python populate_cr_audit.py \
    --start-date 2026-01-01 \
    --end-date 2026-01-31 \
    --percentage 20
```

The visible parser defines:

- `--start-date`
- `--end-date`
- `--percentage` (default `20`)
- `--template` (default visible as `QAR Report.xlsx` in the parser)
- `--tester` (default `Tester`)
- `--change-nums`
- `--output`

### Date-range flow

```text
Start date + end date
        ↓
Validate dates
        ↓
For each configured change type
        ↓
Fetch population from ServiceNow
        ↓
Select audit sample
        ↓
Run audit checks
        ↓
Populate that month's sheet
        ↓
Save report
```

---

# 7. Command-line Argument Details

## `--start-date`

Expected format:

```text
YYYY-MM-DD
```

Example:

```text
2026-09-01
```

## `--end-date`

Expected format:

```text
YYYY-MM-DD
```

## `--percentage`

Percentage of the date-range population to randomly select for audit.

Default visible in the parser:

```text
20
```

The date-range execution path validates that the value is between `1` and `100` inclusive.

## `--template`

Path/name of the Excel template.

The parser visibly defaults to:

```text
QAR Report.xlsx
```

## `--tester`

Name placed into the report's tester field.

Default:

```text
Tester
```

## `--change-nums`

One or more Change Request numbers, or a text file containing them.

The parser uses:

```python
nargs="+"
```

and a custom normalization function.

## `--output`

Requested output filename/path.

The visible Excel writer ultimately places the saved workbook under the `report/` area.

---

# 8. `populate_cr_audit.py` — MAIN ORCHESTRATOR

This is the **brain of the application**.

It does not itself contain all the audit logic. Instead, it coordinates other modules.

Its responsibility is:

```text
INPUT
  ↓
Validate
  ↓
Fetch CRs
  ↓
Choose sample
  ↓
Call audit engine
  ↓
Call Excel writer
  ↓
SAVE
```

---

## 8.1 `parse_arguments()`

Creates the CLI parser.

Visible arguments:

```text
--start-date
--end-date
--percentage
--template
--tester
--change-nums
--output
```

---

## 8.2 `load_change_numbers(change_inputs)`

Purpose:

> Convert whatever the user supplied into a clean list of valid Change Request numbers.

Supported input:

```text
Direct CR numbers
OR
Existing text file
```

### File handling

When an input is an existing file, it is read using:

```python
encoding="utf-8-sig"
```

### Splitting

The contents are split with whitespace/comma/semicolon separators using a regular expression equivalent to:

```python
r"[\s,;]+"
```

### Normalization

Each value is:

```text
stripped
→ uppercased
→ regex validated
```

The visible validation regex is:

```python
re.compile(r"CHG\d+", re.IGNORECASE)
```

### Invalid values

Invalid entries are reported as ignored.

### Duplicate handling

Duplicates are removed while preserving order using the equivalent of:

```python
list(dict.fromkeys(change_numbers))
```

### Example

Input:

```text
chg123, CHG124; CHG123
hello
```

becomes approximately:

```text
CHG123
CHG124
```

with `hello` rejected.

---

## 8.3 `validate_date_range(start_date, end_date)`

Parses dates using:

```python
%Y-%m-%d
```

Two validations are performed:

1. Dates must use the expected format.
2. Start date cannot be after end date.

Visible error messages include:

```text
Dates must use YYYY-MM-DD format
start_date must not be after end_date
```

---

## 8.4 `report_filename(month, change_wise=False)`

Generates a standard report filename based on the current year.

Visible patterns:

```text
Normal/monthly report:
FCB {month} {year} - QAR Report.xlsx
```

and:

```text
Change-wise report:
FCB {month}-QAR Change Wise.xlsx
```

---

## 8.5 `get_reference_display_value(value)`

ServiceNow reference fields may appear as dictionaries.

This helper normalizes values into a human-readable display value.

Conceptually:

```python
if value is dict:
    return value.get("display_value")
elif value is str:
    return value
else:
    return None
```

---

## 8.6 `print_missing_change_numbers(requested_numbers, all_crs)`

Compares what the user asked for with what ServiceNow actually returned.

Purpose:

```text
Requested CRs
     vs
Returned CRs
```

Missing CRs are printed so that a user knows some requested records were not returned.

---

# 9. Explicit Change-number Execution Path

The function `run_change_number_audit(args, change_numbers)` handles explicit CR input.

Flow:

```text
change_numbers
      ↓
fetch_crs_from_servicenow(change_numbers=...)
      ↓
print missing numbers
      ↓
if nothing returned → exit cleanly
      ↓
get current month
      ↓
create_change_wise_report(...)
```

API failures are caught as `requests.RequestException` / `ValueError` and reported before terminating that path.

---

# 10. `create_change_wise_report(...)`

This path is used when the user gives specific Change numbers.

The visible flow is effectively:

```text
For each Change Request:
    determine change type
    ↓
    collect audit data
    ↓
    write into the corresponding sheet
```

The screenshot shows special handling for Standard Changes:

```text
Standard CR
    ↓
SC Result Sheet
    +
SC Template Result Sheet
```

Other change types are routed to their dedicated result sheets.

The exact internal column-increment values are implementation details tied to the workbook layout and are intentionally not treated as business logic in this documentation.

---

# 11. Date-range Execution Path

The main date-based function is `run_date_range_audit(args, start_date, end_date)`.

### Validation

It first validates the date range and percentage.

### Month handling

The report month is derived from the start date.

### Existing report behavior

The visible code checks for an already existing report under `report/` using the generated filename.

Conceptually:

```text
Existing generated report?
      │
      ├── YES → use it as workbook base
      │
      └── NO  → use supplied template
```

This means rerunning a report can reuse an existing workbook rather than always starting from a blank template.

### Change-type loop

The code iterates through the configured change-result mapping from `constants.CHANGE_RESULTS`.

The visible mapping associates result sheets with types approximately as:

```text
SC Result Sheet          → standard
SC Template Result Sheet → standard
NC Result Sheet          → normal
EC Result Sheet          → emergency
Exp Result Sheet         → expedited
```

Standard appears twice because the Standard Template sheet is a specialized audit view.

---

# 12. `process_monthly_sheet(...)`

This is the central per-sheet processing function.

For normal result sheets:

```text
Population
   ↓
Random sample
```

For the Standard Template sheet:

```text
Population
   ↓
Find unique Standard proposals
   ↓
Check proposal review age
   ↓
Select proposals requiring audit
```

Then:

```text
Get basic report metadata
        ↓
Get audit data for selected CRs
        ↓
Populate basic data in Excel
        ↓
Populate CR columns in Excel
```

---

# 13. Sample Selection Logic

## `random_select(crs, percentage)`

Purpose:

> Select a random subset of the population for date-range auditing.

Conceptually:

```python
count = ceil(len(crs) * percentage / 100)
count = at least one when a non-empty population exists
selected = random.sample(crs, min(count, len(crs)))
```

The visible source uses a `max(...)` lower bound and `random.sample(...)`.

### Important implementation note

The screenshot/OCR around this line is visually ambiguous between `max(1, ...)` and `max(2, ...)`. The stronger visual/code reconstruction used in review indicates a one-sample minimum in the observed implementation, but this exact line should be verified against the live source before treating the minimum as contractual behavior.

Also, the function uses `math.ceil(...)`; the visible import list in the screenshots does not clearly show `import math`. This should be verified in the live file because a missing import would produce a runtime error.

---

# 14. Standard Change Template Selection Logic

Standard Template auditing is different from ordinary random sampling.

The function visible in `populate_cr_audit.py` is:

```text
get_standard_template_proposals(all_crs, today)
```

### Step 1 — only Standard changes

CRs whose `type` is not `standard` are skipped.

### Step 2 — identify proposal

The code obtains the Standard Change Proposal / Template version reference using the CR's:

```text
std_change_producer_version
```

and extracts the display value.

### Step 3 — deduplicate

A dictionary keyed by proposal name is used so that multiple CRs associated with the same proposal are collapsed to one representative CR.

### Step 4 — fetch proposal record

For each unique proposal, `servicenow_api.get_proposals(cr)` is called.

### Step 5 — read last review date

The proposal's:

```text
u_last_reviewed
```

is read.

### Step 6 — missing review date

If the review date is missing, the proposal is selected.

### Step 7 — invalid review date

An invalid review date is printed and skipped.

### Step 8 — calculate review age

The code calculates month difference approximately as:

```text
(year difference × 12) + month difference
```

and selects proposals once that value exceeds the configured age cutoff visible in the source.

### Important screenshot discrepancy

One screenshot OCR rendered the cutoff as `> 18` while another reconstruction of the same code path showed `> 10`. Because the supplied artifact is screenshot-only and this line is business-critical, **verify the exact cutoff in the live `populate_cr_audit.py` before relying on it operationally**.

The broader business rule is clear:

> **Standard Change Templates with no review date or sufficiently old review date are selected for audit.**

---

# 15. `servicenow_api.py` — ServiceNow Data Layer

This file is the gateway between Python and ServiceNow.

Its role is:

```text
Python audit engine
       ↕
ServiceNow REST API
```

Visible imports include:

```python
import random
from datetime import datetime
import requests
import questions_data
import sc_sheet_data
import constants
```

---

# 16. ServiceNow Tables Used

The constants screenshot exposes these table constants:

| Constant | ServiceNow table |
|---|---|
| `CHANGE_REQUEST_TABLE` | `change_request` |
| `SYS_APPROVAL_TABLE` | `sysapproval_approver` |
| `CHANGE_ATTACH_TABLE` | `sys_attachment` |
| `SYS_USER_GROUP_TABLE` | `sys_user_group` |
| `SYS_USER_GROUP_MEM_TABLE` | `sys_user_grmember` |
| `CMDB_CI` | `cmdb_ci` |
| `CMDB_CI_SERV` | `cmdb_ci_server` |
| `INCIDENT_TABLE` | `incident` |
| `PROBLEM_TABLE` | `problem` |
| `CHANGE_TASK` | `change_task` |
| `SYS_USER_DELEGATE_TABLE` | `sys_user_delegate` |

These tables let the application reconstruct the context needed for an audit rather than depending only on the main Change Request record.

---

# 17. `fetch_crs_from_servicenow(...)`

This is the primary Change Request retrieval function.

It supports two modes.

## 17.1 Explicit numbers

When `change_numbers` is supplied, the code builds a query using ServiceNow's `numberIN` style filtering.

Conceptually:

```text
numberINCHG123,CHG124,CHG125
```

## 17.2 Date-range retrieval

When dates are supplied, the function requires:

```text
start_date
end_date
change_type
```

The query visibly includes:

```text
type = <change_type>
```

and a ServiceNow `BETWEEN` date expression based on the supplied start/end dates.

It also includes a production-system condition:

```text
production_system=true
```

The screenshots show a state filter using a set of ServiceNow state codes. The exact whitespace/formatting is less important than the business purpose:

> **The date-range population is restricted to the configured Change types, date interval, production changes, and selected states.**

---

# 18. HTTP Request Pattern

The code generally calls ServiceNow like this:

```python
requests.get(
    constants.BASE_URL + constants.<TABLE>,
    auth=(constants.USERNAME, constants.PASSWORD),
    headers=constants.headers,
    params=query_params,
    verify=False,
    timeout=30,
)
```

Then:

```python
response.raise_for_status()
data = response.json()
return data.get("result", [])
```

This pattern appears throughout the codebase.

---

# 19. `get_proposals(change)`

Used for Standard Change Template logic.

Conceptually:

```text
CR
 ↓
std_change_producer_version
 ↓
follow ServiceNow reference link
 ↓
obtain linked proposal/template
 ↓
return proposal result dictionary
```

The returned proposal data is then used by Q7/Q8 and Standard Template checks.

---

# 20. `get_basic_data(...)`

Creates report-level metadata.

Visible output structure:

```python
{
    "test_month": str(cur_month),
    "sample_month": str(cur_month),
    "tester": tester,
    "population": full_len,
    "samples_tested": sample_len,
}
```

This data is written into the Excel workbook's metadata area.

---

# 21. `get_audit_data(changes, sheet_name)`

This is the primary aggregator for one or more CRs.

For each Change Request:

```text
common_data
    +
question_results
    +
Standard Template data when applicable
    ↓
merged record
```

The visible merging logic is essentially:

```python
common_data = get_common_data(change)
questions_data_result = get_questions_data(change)
merged = common_data | questions_data_result
```

For the Standard Template sheet:

```python
std_data = get_std_questions_data(change)
merged = merged | std_data
```

The merged object is then appended to the final `audit_data` list.

---

# 22. `get_audit_data_single(...)`

A one-record version of the same aggregation idea.

This is used by the explicit Change-number / change-wise path.

Conceptually:

```text
One CR
 ↓
common data
 ↓
Q1-Q19
 ↓
optional Standard Template data
 ↓
return one audit dictionary
```

---

# 23. Common Change Data

The visible `get_common_data(change)` result contains fields aligned with the Excel report mapping.

Core fields visible include:

```text
number
 test_date
 change_type
 category
 closure_status
 related_inc
 related_prb
 jira_issue
```

### `test_date`

The visible implementation uses the current Python datetime.

### `closure_status`

The report uses the Change's close code when the Change is closed.

---

# 24. Related Incident Collection

## `get_all_related_incidents(change)`

This function does more than simply read one field.

It combines several sources.

### Source 1 — related incidents field

The Change's `u_related_incidents` field is read.

When it is a string, the code splits it by commas and trims entries.

### Source 2 — incidents caused by the Change

A ServiceNow Incident query is made using the Change number.

### Source 3 — incidents fixed by the Change

Another query uses the Change relationship/reference information.

The purpose is to catch incidents linked through different ServiceNow relationship paths.

### Final cleanup

The result is:

```text
deduplicated
→ empty values removed
→ sorted
→ joined into a comma-separated string
```

If nothing is found, the result is the configured `NA` value.

---

# 25. Related Problem Collection

## `get_all_related_problems(change)`

The same concept is used for Problems.

Data comes from:

1. the Change's related-problems field
2. a ServiceNow Problem query using the Change/RFC relationship

Results are deduplicated and returned as a comma-separated string, or `NA` when none exist.

---

# 26. Question Result Dictionary

`get_questions_data(change)` maps each business rule function to the workbook field name.

Visible mapping:

```python
{
    "q1_testing_completed": questions_data.q1_std_testing_completed(change),
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
    "q12_normal_lowrisk": questions_data.q12_normal_fcb_lowrisk_approval_check(change),
    "q13_normal_high": questions_data.q13_normal_modhigh_risk_approval_check(change),
    "q14_std_approval": questions_data.q14_std_approval(change),
    "q15_approval_timing": questions_data.q15_approval_timing(change),
    "q16_major_incident": questions_data.q16_major_incident_check(change),
    "q17_post_validation": questions_data.q17_post_validation_check(change),
    "q18_issues_doc": questions_data.q18_success_issues_check(change),
    "q19_outside_window": questions_data.q19_outside_window_check(change),
}
```

This dictionary is the bridge between the business rules and Excel.

---

# 27. Audit Result Values

The codebase clearly uses result values such as:

```text
YES
NO
NA
ERROR
```

There is also a special result visible for Q1:

```text
NON_TESTABLE
```

This is not a generic error; it represents a deliberate business condition where the Change cannot be pre-tested and an explanation is available.

---

# 28. `questions_data.py` — Audit Rules Engine

This is the **main business-rule module**.

It contains the actual logic that answers the audit questions.

The questions are not inferred by an ML model.

Instead, they follow patterns like:

```text
read field
↓
check condition
↓
query related table if necessary
↓
possibly inspect attachment
↓
return YES / NO / NA / ERROR
```

---

# 29. Helper: `get_closure_status(change)`

Visible logic:

```python
if change.get("state").casefold() == "closed":
    return change.get("close_code")
else:
    return None
```

So the helper converts:

```text
State = Closed
```

into:

```text
Close Code
```

and otherwise returns no closure status.

---

# 30. Q1 — Testing Completed and Required Artifacts

Function:

```text
q1_std_testing_completed(change)
```

Business intent:

> Was testing completed/documented and were the required artifacts attached?

### Logic

The function starts with a negative/default evidence flag.

It checks whether the Change cannot be pre-tested:

```text
u_change_cannot_be_pre_tested
```

If the field says `true`, the code expects a reason:

```text
u_why_can_this_not_be_pre_tested
```

If the reason exists, the function returns the special `NON_TESTABLE` value.

Otherwise the normal path checks required non-production test-plan information.

Visible fields include:

```text
u_non_prod_test_plan_summary
u_non_prod_test_plan_results
```

Then the function queries the `sys_attachment` table for attachments belonging to the Change.

Attachment filenames are inspected for evidence-like terms such as:

```text
test
non prod
```

A qualifying filename causes the evidence flag to become true.

### Mental model

```text
Can it be pre-tested?
      │
      ├── NO + reason present → NON_TESTABLE
      │
      └── YES
           ↓
      Test summary exists?
      ↓
      Test results exists?
      ↓
      Test/non-prod attachment exists?
      ↓
      YES / NO
```

---

# 31. Q2 — Non-testable Normal/Expedited: SDLC SII/RITM

Function:

```text
q2_nc_exp_sdlc_ritm(change)
```

Applies only to:

```text
Normal
Expedited
```

Other change types return `NA`.

It also skips the check for visible SDLC initiative types such as:

```text
DR Routine without Code
Maintenance - Initiative Not Applicable
```

The actual check is relevant when the Change is marked as unable to be pre-tested.

It then inspects:

```text
u_non_prod_test_plan_summary
```

and looks for an SDLC identifier such as:

```text
RITM
SII
```

Failure keywords are also checked.

### Result concept

```text
Not applicable → NA
Not non-testable → NA
Non-testable + required identifier present → YES
Non-testable + required identifier absent/invalid → NO
```

---

# 32. Q3 — Normal/Expedited Reason for Change

Function:

```text
q3_nc_exp_business_project_patching(change)
```

Applies only to:

```text
Normal
Expedited
```

It reads:

```text
u_reason_for_change
```

The visible code marks the check as `NO` when the reason contains terms associated with:

```text
business request
project
patch
```

Otherwise it returns `YES`.

This is a straightforward keyword/business-category rule.

---

# 33. Q4 — Reason of Change vs SDLC Initiative Type

Function:

```text
q4_reason_vs_sdlc(change)
```

Applies only to:

```text
Normal
Expedited
```

The logic explicitly handles:

### Reason = Maintenance

Must match an SDLC type equivalent to:

```text
Maintenance - Initiative Not Applicable
```

If the mapping matches:

```text
YES
```

otherwise:

```text
NO
```

### Reason = DR Exercise

Must match:

```text
DR Routine without Code
```

Again:

```text
match → YES
mismatch → NO
```

### Other reasons

Visible implementation defaults to `YES`.

---

# 34. Q5 — Incident Association

Function:

```text
q5_std_incident(change)
```

Business question:

> If Reason for Change is `Incident Resolution`, is an Incident associated?

### Logic

```text
Reason != Incident Resolution
        ↓
       NA
```

Otherwise:

```text
u_related_incidents populated?
        ↓
YES / NO
```

---

# 35. Q6 — Problem Association

Function:

```text
q6_std_problem(change)
```

Business question:

> If Reason for Change is `Problem Resolution`, is a Problem associated?

Logic:

```text
Reason != Problem Resolution
        ↓
       NA
```

Otherwise:

```text
u_related_problems populated?
        ↓
YES / NO
```

---

# 36. Q7 — Standard Change Template Applicability

Function:

```text
q7_std_template(change)
```

Applies only to:

```text
Standard
```

Other types return `NA`.

### Logic

1. Read the Change `short_description`.
2. Call `servicenow_api.get_proposals(change)`.
3. Read the proposal `template_value`.
4. Parse pipe-separated key/value entries.
5. Search for the proposal's `short_description`.
6. Compare the Change Short Description with the Proposal Short Description.
7. If the CR's Standard Change Producer reference exists, it can also satisfy the check.

### Simplified logic

```text
Not Standard → NA

Standard
   ↓
Fetch Standard proposal
   ↓
Compare short descriptions
   ↓
If appropriate Standard template reference exists
        → YES
Else
        → NO
```

---

# 37. Q8 — Standard Change Differentiation

Function:

```text
q8_differentiation(change)
```

Applies only to Standard Changes.

The visible code uses:

```python
fuzz.token_sort_ratio(...)
```

It compares the Change short description against the Standard Change reference/template display value.

The visible threshold is:

```text
>= 40 → YES
< 40  → NO
```

### Important semantic note

The function name/question suggests a “differentiation” concept, but the visible implementation returns `YES` when similarity is high enough.

Therefore the documentation should distinguish:

```text
Implementation behavior:
High similarity ≥ threshold → YES
```

from:

```text
Potential business interpretation:
The question appears to ask whether the Change is sufficiently differentiated.
```

This is an item worth validating with the business owner.

---

# 38. Q9 — Requester/Assignee Must Not Be Sole Approver

Function:

```text
q9_approver_check(change)
```

Business intent:

> Requested By and Assigned To should not be the only approver(s).

### Special Standard Change exceptions

For Standard Changes, the code checks:

```text
STANDARD_CH_APPROVAL_EXCEPTION_TEMPLATES
```

If the Standard Change Template version starts with one of the configured exception names, the result is:

```text
NA
```

### Normal path

The code extracts:

```text
Requested By
Assigned To
```

Then queries:

```text
sysapproval_approver
```

for approved approvals related to the Change.

### Result logic

```text
No approvals → NO
```

If either the requester or assignee is not the only approved approver, the test passes.

A visible rule is effectively:

```text
Someone other than requester/assignee approved
    OR
More than one approver exists
        ↓
       YES
```

Otherwise:

```text
NO
```

---

# 39. Q10 — Emergency Change Approval

Function:

```text
q10_emergency_approval_check(change)
```

Applies only to Emergency Changes.

### Required concepts

The visible code checks for approval involving:

```text
Assignment Group approval structure
+
SLT approval
```

### Flow

```text
Emergency CR
    ↓
Get Assignment Group
    ↓
Get Assignment Group details
    ↓
Find child Approval Group
    ↓
Get members of Approval Group
    ↓
Get actual approved CR approvers
    ↓
Check approval-group member approval
    ↓
Check SLT approval
    ↓
Both satisfied?
     ├── YES → YES
     └── NO  → NO
```

---

# 40. Q11 — Expedited Change Approval

Function:

```text
q11_expedited_fcb_approval_check(change)
```

Applies only to Expedited Changes.

The question is visibly centered on:

```text
ITCM
CI Owner
Assignment Group SLT
```

### ITCM approval

The code gets members of the Change Management Approval group and checks whether an approved CR approver belongs to that group.

### CI Owner approval

The code fetches the CI from CMDB, then looks at ownership information such as:

```text
owned_by
```

It then checks whether a CI owner approved the Change.

### Assignment Group / SLT approval

It gets the Assignment Group, follows its approval-group relationship, obtains group members, and checks approved approvers.

### Final condition

The visible implementation requires all required approval flags to be true.

Conceptually:

```text
ITCM approved
AND
CI Owner approved
AND
Assignment Group / SLT approval confirmed
        ↓
       YES
```

Otherwise:

```text
NO
```

---

# 41. Q12 — Normal Low-risk Approval

Function:

```text
q12_normal_fcb_lowrisk_approval_check(change)
```

Applies only to Normal Changes with Low Risk.

The visible comment describes a required approval chain involving concepts such as:

```text
ITCM
CI Owner
Assignment Group approval
```

and mentions an FCB/LSVB variation.

### Visible implementation flow

```text
Normal + Low Risk
       ↓
Get actual approved approvers
       ↓
Check Change Management Approval (ITCM)
       ↓
Find Assignment Group / child approval group
       ↓
Check assignment-group approval
       ↓
Get CI owner
       ↓
Check CI owner approval
       ↓
All required flags true?
```

Result:

```text
YES / NO
```

### Delegates

The approval logic also considers delegated approval relationships in this area of the codebase.

---

# 42. Q13 — Normal Moderate/High-risk Approval

Function:

```text
q13_normal_modhigh_risk_approval_check(change)
```

Applies only when:

```text
Change type = Normal
AND
Risk = Moderate or High
```

### CAB approval

The code queries the CAB approval group members.

### ITCM approval

The code queries the Change Management Approval group members.

### Final logic

```text
At least one CAB member approved
AND
At least one ITCM member approved
        ↓
       YES
```

Otherwise:

```text
NO
```

---

# 43. Q14 — Standard Approval

Function:

```text
q14_std_approval(change)
```

Purpose:

> Validate Standard Change approval requirements.

The screenshots confirm that this function exists and is wired into `get_questions_data()` and `test_functions.py`.

The visible code also contains Standard Change approval-exception handling based on:

```text
STANDARD_CH_APPROVAL_EXCEPTION_TEMPLATES
```

However, some of the exact Q14 source lines are obscured in the screenshot sequence.

Therefore the safest source-faithful description is:

```text
Standard Change
    ↓
Evaluate Standard approval rule
    ↓
Configured exception templates can result in NA
    ↓
Otherwise return the corresponding approval result
```

**Do not use this section as a substitute for the live Q14 function source when changing the exact business rule.**

---

# 44. Q15 — Approval Timing

Function:

```text
q15_approval_timing(change)
```

Business intent:

> Were the required approvals completed before the relevant Change start time, with a special exception process for cases such as Emergency Change approval evidence?

### Main timing comparison

The visible code:

1. Queries approved `sysapproval_approver` records.
2. Reads the latest/last approval record in the returned list.
3. Extracts `sys_updated_on`.
4. Parses Change start time.
5. Parses approval time.
6. Compares them.

Visible comparison concept:

```text
Actual/Work start >= Last approval time
        ↓
       YES
```

Otherwise it enters an evidence/exception path.

### Attachment exception process

It queries `sys_attachment` and looks for evidence filenames containing concepts such as:

```text
.msg
approval
```

For the relevant message file, it:

```text
downloads the attachment
        ↓
reads the Outlook/message metadata
        ↓
gets sent time and sender
```

The code uses message-parsing packages, preferring `extract_msg` and falling back to another parser when necessary.

### SLT / delegate validation

The code also fetches:

```text
Assignment Group
        ↓
its SLT reference
        ↓
SLT sys_id/link
        ↓
possible delegate records
```

The exception may be accepted when the approval message was sent by the SLT approver or an allowed delegate and the message timestamp satisfies the relevant time requirement.

### Result

```text
Timing requirement satisfied → YES
Valid exception evidence → YES
Requirement not satisfied → NO
Unexpected failure → ERROR
```

---

# 45. Q16 — Major Incident Association

Function:

```text
q16_major_incident_check(change)
```

Purpose:

> Check whether a Major Incident was associated with / caused by the Change.

The code queries the Incident table using the Change relationship and checks:

```text
major_incident_state = accepted
```

### Result

```text
Accepted Major Incident found → YES
No matching record           → NO
Invalid/missing CR context   → NA
```

---

# 46. Q17 — Post-production Validation

Function:

```text
q17_post_validation_check(change)
```

This is one of the most technically interesting controls because it combines:

```text
ServiceNow fields
+
ServiceNow attachments
+
image filtering
+
OCR
```

### Applicability

A closed Change is required for the meaningful post-validation check.

### Required fields

The visible logic requires both:

```text
u_post_prod_deployment_test_plan_results
u_post_prod_deployment_test_plan_summary
```

Missing required information causes a negative result.

### Attachment search

The function queries `sys_attachment` and checks attachment filenames against configured post-production keywords such as:

```text
post
deployment
post implementation
prod
```

Only image extensions are considered.

Configured image extensions visible in `constants.py` are:

```python
("png", "jpg", "jpeg")
```

### OCR pipeline

For each usable image:

```text
Download attachment
      ↓
Temporary/local image
      ↓
analyze_dashboard_status(...)
      ↓
SUCCESS / FAILURE / UNKNOWN
```

### Final behavior

Conceptually:

```text
Recognized SUCCESS → YES
Recognized FAILURE → NO
Artifact found but status unclear → UNKNOWN
No usable artifact → NO
```

The exact code precedence is inherited from `search.py`, described later.

---

# 47. Q18 — Successful with Issues Documentation

Function:

```text
q18_success_issues_check(change)
```

Business intent:

> When a Change closes as `Successful with Issues`, is the issue fully documented and supported by the required post-implementation review evidence?

### Applicability

If the Change is not closed:

```text
NA
```

### Closure code

The special path activates for:

```text
Successful with Issues
```

### Attachment check

The code searches Change attachments for a specific post-implementation review message naming pattern containing the Change number.

The visible pattern is approximately:

```text
RE_FCB [CHG<number>] - Post Implementation Review.msg
```

### Result

```text
Matching evidence found → YES
Not found               → NO
Otherwise / not applicable → NA
```

---

# 48. Q19 — Change Window Timing

Function:

```text
q19_outside_window_check(change)
```

Business concept:

> Check whether actual Change execution occurred inside the planned Change window.

### Applicability

Closed Change required.

If the Change is not closed:

```text
NA
```

### Required timestamps

The code reads:

```text
start_date
end_date
work_start
work_end
```

If required values are missing, it returns `NA`.

### Parsing

The implementation uses a generic date parser for these timestamps.

### Tolerance

A visible tolerance is:

```python
15 minutes
```

### Important semantic note

The visible code returns:

```text
YES when the actual execution is within the tolerated planned window
NO when it falls outside
```

even though the function/question name is `outside_window_check`.

This is another important business-semantic review item: **the output polarity appears to represent “within window” rather than “outside window.”**

---

# 49. `questions_data.py` — Approval Helper Functions

Several helper functions support the approval controls.

---

## 49.1 `get_change_management_approvers()`

Queries the `sys_user_grmember` table for the group:

```text
Change Management Approval
```

It returns the user display names of that group's members.

Purpose:

```text
Define who counts as ITCM / Change Management approval
```

---

## 49.2 `get_actual_change_approvers(change)`

Queries:

```text
sysapproval_approver
```

using the Change number and:

```text
state = approved
```

It extracts the approver display values.

Purpose:

```text
Get actual people who approved THIS Change
```

The approval question functions then compare actual approvers to the expected group membership.

---

## 49.3 Delegate handling

The code also queries:

```text
sys_user_delegate
```

This is used in approval exception/representation logic.

The function extracts delegate display values for a given user.

---

# 50. Approval Logic Pattern Used Throughout the Project

Approval controls follow a common architecture:

```text
Expected approval role/group
        ↓
Find group / members
        ↓
Find actual approved CR approvers
        ↓
Compare identities
        ↓
YES / NO
```

This is more reliable than simply checking a single “approval = approved” field because it tries to verify **who approved**.

---

# 51. `download_attachment(...)`

Attachment evidence is downloaded from ServiceNow.

The visible implementation uses the attachment URL base plus the attachment `sys_id` and `/file`.

Conceptually:

```text
attachment sys_id
      ↓
ServiceNow attachment endpoint
      ↓
GET binary content
      ↓
write to local save path
```

The implementation streams the response in chunks.

---

# 52. Outlook / MSG Evidence Parsing

`get_sent_time_app_name_from_msg(file_path)` parses `.msg` evidence.

The visible implementation prefers:

```text
extract_msg
```

and has a fallback path using another parser.

It extracts approximately:

```text
message date / sent time
sender / app name
```

This is used mainly in the Q15 approval-timing exception workflow.

---

# 53. CI / Sys-ID Helper

A helper named conceptually like:

```text
get_sys_id(name_field, name_value)
```

queries a ServiceNow CMDB/server table to resolve a sys_id from a field/value pair.

This supports the approval/group/reference logic.

---

# 54. `sc_sheet_data.py` — Standard Change Template Rules

This module contains checks specific to the `SC Template Result Sheet`.

Visible imports include:

```python
import constants
import requests
from datetime import datetime
```

The standard-template side checks the underlying proposal/template and its associated sample changes.

---

# 55. Standard Template Result Fields

The Excel mapping shows the following Standard Template fields:

```text
cr_number
 test_date
 template_ver
 std_proposal
 category
 closure_status
 related_inc
 jira_issue
 is_retired
 pre_test_plan
 deployment_plan
 backout_plan
 verification_plan
 sample_changes
```

These correspond to the rows in `STD_Q_ROWS`.

---

# 56. `convert_to_datetime(string)`

Converts a ServiceNow-style date/time string into a Python `datetime` value.

The screenshot is slightly blurry around the exact year token in the format string, so the business behavior should be treated as:

> Parse the Standard Change proposal's date/time field into a Python datetime for comparison.

---

# 57. `is_retired(change)`

Checks whether the Configuration Item is retired.

Flow:

```text
CR
 ↓
CMDB CI reference
 ↓
ServiceNow CI lookup
 ↓
read operational_status
 ↓
```

Result:

```text
operational_status = retired → YES
not retired                → NO
missing status             → NA
```

---

# 58. `pre_test_plan(cr_template)`

Checks the presence/quality of the Standard Template's non-production test-plan information.

Visible field:

```text
u_non_prod_test_plan_summary
```

If absent:

```text
NO
```

If the text contains configured failure keywords:

```text
NO
```

Otherwise:

```text
YES
```

There is visible commented-out logic related to non-testability and pass keywords, but the active implementation is the simpler summary-presence + failure-keyword rule.

---

# 59. `deployment_plan(cr_template)`

Reads:

```text
implementation_plan
```

Logic:

```text
Missing → NO
Contains failure keyword → NO
Otherwise → YES
```

---

# 60. `backout_plan(cr_template)`

Reads:

```text
backout_plan
```

Logic:

```text
Missing → NO
Contains failure keyword → NO
Otherwise → YES
```

---

# 61. `verification_plan(cr_template)`

Reads:

```text
u_post_prod_deployment_test_plan_summary
```

Logic:

```text
Missing → NO
Contains failure keyword → NO
Otherwise → YES
```

---

# 62. `fetch_cr_close_codes(change_no)`

Looks up a Change Request by its Change number and returns its close code if the Change is closed.

Otherwise the visible implementation returns `NA`.

This helper is used by Standard Template sample-change validation.

---

# 63. `sample_changes(e_proposal_result)`

This is the key Standard Change Template sample-control function.

It uses proposal information such as:

```text
proposal_type
change_requests
sys_created_on
```

The `change_requests` field is split into individual Change numbers.

### Post-cutoff/new proposal behavior

The visible code contains a fixed cutoff date around:

```text
2025-09-08
```

For proposals created on/after that date, the code requires at least four associated Change Requests.

All sampled/associated changes are then checked for a `Successful` close code.

### Modify proposals

For visible `Modify` proposals:

```text
At least one change is needed
AND
all checked changes must be Successful
```

### New proposals

For visible `New` proposals:

```text
At least four changes are needed
AND
all checked changes must be Successful
```

### Overall result

Conceptually:

```text
No associated changes → NO
Too few required samples → NO
Any unsuccessful sample → NO
All required samples successful → YES
```

### Important maintenance note

The fixed cutoff date means this rule is **time-sensitive but hard-coded** rather than automatically moving with the current date.

---

# 64. `get_std_questions_data(change)`

Located in `servicenow_api.py`, this function assembles all Standard Template-specific outputs.

Flow:

```text
Get proposal
    ↓
Extract proposal number
    ↓
Extract template_value
    ↓
Parse template_value into key/value pairs
    ↓
Build cr_template dictionary
    ↓
Call sc_sheet_data functions
    ↓
Return Standard Template audit fields
```

The parser handles pipe-separated entries and splits each entry around `=`.

A visible diagnostic checks whether a `reason_for_change` item is present in the parsed template data.

---

# 65. `search.py` — OCR / Dashboard Evidence Engine

This file is responsible for turning image evidence into a simple machine-readable status.

The flow is:

```text
Image file
   ↓
OpenCV preprocessing
   ↓
Tesseract OCR
   ↓
Normalize text
   ↓
Search failure keywords
   ↓
Search success keywords
   ↓
SUCCESS / FAILURE / UNKNOWN
```

---

# 66. Local Tesseract Setup

The code calculates the project directory using `Path` and points Tesseract to a local executable.

Conceptually:

```python
BASE_DIR = Path(__file__).resolve().parent
TESSERACT_DIR = BASE_DIR / "tesseract"
TESSDATA_DIR = TESSERACT_DIR / "tessdata"
```

The Tesseract command path is configured to the bundled `tesseract.exe`.

The code also updates the process environment so Tesseract can locate its dependencies and language data.

---

# 67. Tesseract Configuration

Visible configuration:

```text
--oem 3 --psm 6
```

The OCR engine therefore uses a standard layout-oriented page segmentation configuration.

---

# 68. OCR Success Keywords

The visible `SUCCESS_PATTERNS` list contains terms such as:

```text
success
successful
completed
complete
succeeded
healthy
available
active
resolved
passed
up to date
```

These are examples of the configured positive-status vocabulary.

---

# 69. OCR Failure Keywords

The visible `FAILURE_PATTERNS` contains many operationally negative/in-progress words, including examples such as:

```text
failed
failure
aborted
terminated
critical
timeout
stopped
denied
unhealthy
running
executing
pending
starting
provisioning
deploying
processing
in progress
progress
```

The exact full list should be read from the live `search.py` if it must be treated as a contractual keyword dictionary.

---

# 70. `preprocess_image(image_path)`

The image preprocessing steps visible in the screenshots are approximately:

```text
Check file exists
      ↓
Read with cv2.imread
      ↓
Convert to grayscale
      ↓
Upscale image
      ↓
Apply thresholding / Otsu
      ↓
Optionally invert image based on mean intensity
      ↓
Return cleaned image
```

### Why preprocessing exists

OCR is more reliable when screenshots contain:

- readable contrast
- larger text
- binary foreground/background separation

The code is therefore preparing screenshots before giving them to Tesseract.

---

# 71. `normalize_text(text)`

The visible implementation conceptually does:

```python
re.sub(r"\s+", " ", text.lower()).strip()
```

So OCR text is:

```text
lowercased
↓
whitespace normalized
↓
trimmed
```

before keyword matching.

---

# 72. `extract_text(image)`

Passes the processed image into:

```text
pytesseract.image_to_string(...)
```

using the configured Tesseract settings.

Output:

```text
plain OCR text
```

---

# 73. `analyze_dashboard_status(image_path)`

Returns a structured status object.

Core logic:

```text
Preprocess image
      ↓
OCR
      ↓
Normalize
      ↓
Check failure patterns FIRST
      ↓
If no failure → check success patterns
      ↓
If neither → UNKNOWN
```

### Critical precedence rule

**Failure matching happens before success matching.**

Therefore, if a screenshot contains both:

```text
successful
```

and a configured failure/in-progress word such as:

```text
running
```

then the failure branch can win.

This is an important behavior when interpreting Q17.

---

# 74. `constants.py` — Shared Configuration

The codebase centralizes many shared values in `constants.py`.

This avoids repeating:

- table names
- keywords
- result constants
- API settings
- report mappings

---

# 75. Date Format

Visible constant:

```python
DATE_FORMAT = "%m/%d/%Y %I:%M:%S %p"
```

This is used for one family of ServiceNow timestamps, especially approval timing.

---

# 76. Change Number Pattern

Visible regex:

```python
CHANGE_NUMBER_PATTERN = re.compile(r"CHG\d+", re.IGNORECASE)
```

This is used for input validation.

---

# 77. Result Constants / Sheet Mapping

Visible mappings include:

```python
SCT = "Standard Change Template"
```

and a `CHANGE_RESULTS` mapping corresponding to:

```text
NC Result Sheet          → normal
SC Result Sheet          → standard
EC Result Sheet          → emergency
Exp Result Sheet         → expedited
SC Template Result Sheet → standard
```

The duplicated `standard` is deliberate because the project produces both the Standard Change result view and the Standard Change Template audit view.

---

# 78. Keyword Families in `constants.py`

Visible keyword groups include:

### `STD_TEST_FAIL_KEYWORDS`

Examples visible:

```text
n/a
not required
not needed
does not exist
```

### `FAIL_KEYWORDS`

Contains the same type of negative/non-applicable phrases and is reused by several template/control checks.

### `PASS_KEYWORDS`

Visible examples include:

```text
applies to production
deployment plan
back out
verification
```

### `NON_PROD_KEYWORDS`

Visible examples include:

```text
test
non prod evidence
non prod artifact
```

### `POST_PROD_KEYWORDS`

Visible examples include:

```text
post
deployment
post implementation
prod
```

### `IMAGE_EXTENSIONS`

Visible value:

```python
("png", "jpg", "jpeg")
```

---

# 79. Standard Change Approval Exception Templates

`constants.py` contains a configured list named conceptually:

```text
STANDARD_CH_APPROVAL_EXCEPTION_TEMPLATES
```

The visible screenshots show examples including items such as:

```text
Cyber Threat Proxy Change
Cyber Threat Black Hole Route
XSOAR Content Change
```

There are additional configured entries in the source.

These templates are used to bypass certain Standard approval controls by returning `NA` rather than attempting the normal approval test.

**Do not manually duplicate the list in business logic; keep it centralized in `constants.py`.**

---

# 80. API Credentials and URLs

The constants screenshot shows configuration variables conceptually including:

```text
BASE_URL
USERNAME
PASSWORD
headers
INSTANCE_URL
ATT_FILE_URL
```

The documentation deliberately does not reproduce secret values.

Operationally, these constants control:

```text
ServiceNow base API URL
authentication
HTTP headers
attachment URL
instance/reference URL
```

---

# 81. `excel_handler.py` — Excel Reporting Engine

This module is deliberately thin.

It does not decide whether a Change passed a control.

Instead, its job is:

```text
Take already-computed data
        ↓
Put values into exact template rows/columns
        ↓
Save workbook
```

It uses:

```python
import openpyxl
import os
```

---

# 82. Excel Row Mapping — General Change Data

Visible `GENERAL_DATA_ROWS` mapping:

| Row | Data key |
|---:|---|
| 10 | `cr_number` |
| 11 | `test_date` |
| 12 | `change_type` |
| 13 | `category` |
| 14 | `closure_status` |
| 15 | `related_inc` |
| 16 | `related_prb` |
| 17 | `jira_issue` |

This establishes the top/basic data area of the result sheet.

---

# 83. Excel Row Mapping — Basic Report Data

Visible `BASIC_DATA_ROWS`:

| Row | Data key |
|---:|---|
| 2 | `test_month` |
| 3 | `sample_month` |
| 4 | `tester` |
| 5 | `population` |
| 6 | `samples_tested` |

These values go into a fixed metadata column rather than one column per Change.

---

# 84. Excel Row Mapping — Q1 to Q19

Visible `QUESTION_ROWS`:

| Row | Key | Control |
|---:|---|---|
| 18 | `q1_testing_completed` | Q1 |
| 19 | `q2_nc_exp_sdlc_ritm` | Q2 |
| 20 | `q3_nc_exp_business_project_patching` | Q3 |
| 21 | `q4_reason_vs_sdlc` | Q4 |
| 22 | `q5_inc_assoc` | Q5 |
| 23 | `q6_prb_assoc` | Q6 |
| 24 | `q7_std_template` | Q7 |
| 25 | `q8_differentiation` | Q8 |
| 26 | `q9_approver_check` | Q9 |
| 27 | `q10_emergency_approval` | Q10 |
| 28 | `q11_fcb_expedited` | Q11 |
| 29 | `q12_normal_lowrisk` | Q12 |
| 30 | `q13_normal_high` | Q13 |
| 31 | `q14_std_approval` | Q14 |
| 32 | `q15_approval_timing` | Q15 |
| 33 | `q16_major_incident` | Q16 |
| 34 | `q17_post_validation` | Q17 |
| 35 | `q18_issues_doc` | Q18 |
| 36 | `q19_outside_window` | Q19 |

This row map is one of the most important integration contracts in the project.

---

# 85. Excel Row Mapping — Standard Change Template

Visible `STD_Q_ROWS`:

| Row | Data key |
|---:|---|
| 10 | `cr_number` |
| 11 | `test_date` |
| 12 | `template_ver` |
| 13 | `std_proposal` |
| 14 | `category` |
| 15 | `closure_status` |
| 16 | `related_inc` |
| 17 | `jira_issue` |
| 18 | `is_retired` |
| 19 | `pre_test_plan` |
| 20 | `deployment_plan` |
| 21 | `backout_plan` |
| 22 | `verification_plan` |
| 23 | `sample_changes` |

---

# 86. `load_template(template_path)`

Behavior:

```text
Check file exists
    ↓
If not → FileNotFoundError
    ↓
Open with openpyxl
```

Visible message:

```text
Template file not found at: <path>
```

---

# 87. `write_rows(ws, row_mapping, cr_data, col_idx)`

This is the generic cell writer.

Algorithm:

```text
For every row/key mapping:
    get value from cr_data
    if value is neither None nor empty string:
        write value into ws.cell(row, col_idx)
```

This means empty values do not overwrite the template unnecessarily.

---

# 88. `populate_basic_data(...)`

Writes:

```text
Test Month
Sample Month
Tester
Population
Samples Tested
```

into the fixed metadata column for the chosen sheet.

---

# 89. `populate_cr_column(...)`

Chooses the correct row mapping depending on sheet name.

Logic:

```text
SC Template Result Sheet?
       │
       ├── YES → use STD_Q_ROWS
       │
       └── NO  → use GENERAL_DATA_ROWS + QUESTION_ROWS
```

This is the exact bridge between an audit dictionary and the workbook layout.

---

# 90. `populate_sheet(...)`

Validates that the requested worksheet exists.

Then starts Change columns from the visible starting index of:

```text
3
```

Conceptually:

```text
Column A/B = template/meta structure
Column C onward = Change 1, Change 2, Change 3, ...
```

For every Change:

```text
calculate column
↓
populate_cr_column(...)
```

---

# 91. `populate_sheet_single(...)`

A lower-level variant that allows the caller to specify the starting column.

This is useful for the change-wise report path because the orchestrator manages the current column position for each sheet.

---

# 92. `save_output(workbook, output_path)`

The visible implementation:

1. Prints the output target.
2. Ensures the `report` directory exists.
3. Saves the workbook under `report/`.

Conceptually:

```python
os.makedirs("report", exist_ok=True)
workbook.save(f"report/{output_path}")
```

### Important implication

The `--output` argument is not an entirely independent filesystem destination in the visible implementation; it is combined with the `report/` directory in the save function.

This should be verified if absolute paths are expected.

---

# 93. Workbook Structure — `QAR_report.xlsx`

The visible workbook contains these sheets:

```text
Exp Result Sheet
EC Result Sheet
SC Result Sheet
Revision Control
NC Result Sheet
SC Template Result Sheet
```

The workbook is not generated from scratch.

It is a **pre-designed report template** that Python fills.

---

# 94. Normal / Standard / Emergency / Expedited Sheets

The general result sheets contain a structured table with:

```text
Test metadata
Change details
Audit control rows
Totals / summary
```

Visible report areas include:

```text
Change Number
Test Date
Change Type
Category
Closure Status
Related INC #
Jira Issue

Q1
Q2
...
Q19

Total Controls Passed
Total Failed
Error Rate
Pass/Fail summary
```

The actual workbook also contains labels, notes, formatting, and print configuration.

---

# 95. Standard Change Template Sheet

The Standard Template sheet is specialized.

It contains fields such as:

```text
Sample
Test Date
Standard Change Template Version
Standard Change Proposal
Category
Closure Status
Related INC #
Jira Issue
Is CI Retired?
Pre-Test Plan
Deployment Plan
Backout Plan
Verification Plan
Sample Changes
```

This sheet is populated from proposal/template data rather than only the normal Q1–Q19 question set.

---

# 96. Revision Control Sheet

The workbook includes a `Revision Control` tab.

The screenshots show historical workbook/process revisions across 2024–2026.

Visible history includes changes such as:

- increasing the capacity for larger audit populations/samples
- increasing Standard Template sample capacity
- changes around combined testing/artifact controls
- additional Pre-Implementation questions around Reason for Change

This sheet is a documentation/history area of the Excel deliverable rather than Python execution logic.

---

# 97. Report Population vs Sample

The workbook distinguishes:

```text
Population
Samples Tested
```

The population is the total number of Change Requests returned for the relevant query/type.

The sample count is the number actually selected for testing.

For example:

```text
Population = 100
Percentage = 20
Sample tested ≈ 20
```

The exact final count is integer-rounded by the random sampling function.

---

# 98. Full End-to-End Flow

A single Change Request can move through the system like this:

```text
                           USER INPUT
                               │
                    CHG1234567 / date range
                               │
                               ▼
                  populate_cr_audit.py
                               │
                    validate / normalize
                               │
                               ▼
                 servicenow_api.py
                               │
                  fetch Change Request
                               │
          ┌────────────────────┼────────────────────┐
          │                    │                    │
          ▼                    ▼                    ▼
    Change fields       Related records       Attachments
          │                    │                    │
          │          ┌─────────┴─────────┐          │
          │          ▼                   ▼          │
          │       INC / PRB          Approvals      │
          │                              │          │
          └──────────────┬───────────────┘          │
                         ▼                          │
                 questions_data.py                 │
                         │                          │
              Q1 → Q19 deterministic rules         │
                         │                          │
                         ├──────────────────────────┘
                         │
                         ▼
                   Evidence path
                         │
               image / MSG attachment
                         │
              ┌──────────┴──────────┐
              ▼                     ▼
          `search.py`         MSG parser path
              │                     │
            OCR                sent time/sender
              │                     │
              └──────────┬──────────┘
                         ▼
                  final audit dict
                         │
                         ▼
                  excel_handler.py
                         │
               fill correct cells
                         │
                         ▼
                   QAR workbook
                         │
                         ▼
                   report/*.xlsx
```

---

# 99. Standard Change Template End-to-End Flow

The Standard Template path is different:

```text
Date-range population
        ↓
Standard Change CRs
        ↓
Get std_change_producer_version
        ↓
Deduplicate proposals
        ↓
Fetch proposal
        ↓
Check last review date
        ↓
Select old/unreviewed proposals
        ↓
For selected proposal/CR:
    ├── CI retired?
    ├── Pre-test plan?
    ├── Deployment plan?
    ├── Backout plan?
    ├── Verification plan?
    └── Sample Changes valid?
        ↓
Populate SC Template Result Sheet
```

---

# 100. Data Flow — What a CR Dictionary Looks Like

The exact runtime object is assembled from multiple sources.

A representative shape is:

```python
{
    "number": "CHG...",
    "test_date": datetime(...),
    "change_type": "normal/standard/emergency/expedited",
    "category": "...",
    "closure_status": "Successful / ...",
    "related_inc": "INC...",
    "related_prb": "PRB...",
    "jira_issue": "...",

    "q1_testing_completed": "YES/NO/NA/NON_TESTABLE",
    "q2_nc_exp_sdlc_ritm": "YES/NO/NA",
    "q3_nc_exp_business_project_patching": "YES/NO/NA",
    "q4_reason_vs_sdlc": "YES/NO/NA",
    "q5_inc_assoc": "YES/NO/NA",
    "q6_prb_assoc": "YES/NO/NA",
    "q7_std_template": "YES/NO/NA",
    "q8_differentiation": "YES/NO/NA",
    "q9_approver_check": "YES/NO/NA",
    "q10_emergency_approval": "YES/NO/NA",
    "q11_fcb_expedited": "YES/NO/NA",
    "q12_normal_lowrisk": "YES/NO/NA",
    "q13_normal_high": "YES/NO/NA",
    "q14_std_approval": "YES/NO/NA",
    "q15_approval_timing": "YES/NO/NA/ERROR",
    "q16_major_incident": "YES/NO/NA",
    "q17_post_validation": "YES/NO/NA/UNKNOWN",
    "q18_issues_doc": "YES/NO/NA",
    "q19_outside_window": "YES/NO/NA",
}
```

For the Standard Template sheet, additional keys are merged in, such as:

```text
supplier/template version
proposal number
is_retired
pre_test_plan
deployment_plan
backout_plan
verification_plan
sample_changes
```

---

# 101. Important Separation of Concerns

The project has a reasonable conceptual separation:

```text
populate_cr_audit.py
    = workflow / orchestration

servicenow_api.py
    = external data access + aggregation

questions_data.py
    = business audit decisions

sc_sheet_data.py
    = Standard Template business decisions

search.py
    = image evidence interpretation

excel_handler.py
    = output formatting / workbook writing

constants.py
    = shared configuration
```

This makes the project easier to reason about because the same question logic can theoretically be reused outside of the Excel UI.

---

# 102. What Happens When a Question Is Not Applicable?

Many questions are scoped to certain Change types.

Examples:

```text
Q2 → Normal / Expedited
Q3 → Normal / Expedited
Q4 → Normal / Expedited
Q7 → Standard
Q8 → Standard
Q10 → Emergency
Q11 → Expedited
Q12 → Normal + Low Risk
Q13 → Normal + Moderate/High Risk
```

For other types, the function generally returns:

```text
NA
```

This is essential because a `NA` is not a failure.

---

# 103. Error Handling Philosophy

The code uses several levels of handling.

## Input errors

Examples:

```text
Invalid date
Invalid date order
Invalid CHG number
Missing required arguments
Missing template
```

These are reported to the user and may terminate execution.

## ServiceNow API errors

Common pattern:

```python
response.raise_for_status()
```

and outer functions catch:

```text
requests.RequestException
ValueError
```

## Audit-rule exceptions

Some question functions return:

```text
ERROR
```

when an unexpected processing exception occurs, particularly in more complex evidence/timing logic.

---

# 104. Where the Business Rules Actually Live

When someone asks:

> “Where do I change the audit rule?”

The answer is usually:

### Q1–Q19

Go to:

```text
questions_data.py
```

### Standard Template checks

Go to:

```text
sc_sheet_data.py
```

### Keyword definitions

Go to:

```text
constants.py
```

### ServiceNow field/table retrieval

Go to:

```text
servicenow_api.py
```

### Workbook row/cell mapping

Go to:

```text
excel_handler.py
```

### Overall workflow / CLI behavior

Go to:

```text
populate_cr_audit.py
```

---

# 105. Where to Change a ServiceNow Table

If ServiceNow changes a table or endpoint:

```text
constants.py
```

holds the table names.

Then the actual query usually lives in:

```text
servicenow_api.py
```

or `questions_data.py` / `sc_sheet_data.py` when the query is tightly coupled to one control.

---

# 106. Where to Change an Excel Cell

Do not start in the audit question.

Start with:

```text
excel_handler.py
```

For general rows:

```text
GENERAL_DATA_ROWS
QUESTION_ROWS
```

For Standard Template rows:

```text
STD_Q_ROWS
```

Then verify that the workbook still has the same sheet names and row layout.

---

# 107. Where to Change Sampling

Sampling behavior lives primarily in:

```text
populate_cr_audit.py
```

and:

```text
servicenow_api.py
```

Specifically:

```text
process_monthly_sheet(...)
random_select(...)
```

Standard Template sampling is separately controlled through:

```text
get_standard_template_proposals(...)
sample_changes(...)
```

---

# 108. Where to Change OCR Behavior

Go to:

```text
search.py
```

and then:

```text
constants.py
```

for the keyword dictionaries.

The two layers are:

```text
search.py
    = image processing + OCR + matching order

constants.py
    = keyword vocabulary
```

---

# 109. Where to Change Approval Rules

Most approval-control logic lives in:

```text
questions_data.py
```

with ServiceNow membership/reference data queried directly from:

```text
sysapproval_approver
sys_user_group
sys_user_grmember
sys_user_delegate
cmdb_ci
```

---

# 110. Important Implementation Observations

These are code-review observations, not claims that the business process is wrong.

## 110.1 `verify=False`

Many ServiceNow requests visibly use:

```python
verify=False
```

This disables TLS certificate verification.

For enterprise production use, certificate verification should be reviewed carefully.

---

## 110.2 Credentials in constants

The code exposes variables for:

```text
USERNAME
PASSWORD
```

The actual secret values are not reproduced here.

Secrets should not be committed into source control.

---

## 110.3 `math.ceil` import

`random_select()` visibly uses `math.ceil(...)`.

The screenshot's visible import list does not clearly include `math`.

Therefore verify that the live source includes:

```python
import math
```

or equivalent before execution.

---

## 110.4 Standard Template review cutoff

The screenshot evidence is inconsistent on the literal month cutoff shown by OCR.

One rendering suggests one value and another suggests another.

The important observed business behavior is:

```text
No review date or review age exceeds the configured cutoff
→ select for audit
```

Verify the exact numeric cutoff in the live file before making operational assumptions.

---

## 110.5 Hard-coded Standard sample cutoff date

`sc_sheet_data.py` visibly contains a fixed cutoff around:

```text
2025-09-08
```

This means the rule will eventually become stale unless that date is intentionally fixed by policy.

---

## 110.6 Q8 threshold / meaning

The visible implementation is:

```text
token_sort_ratio >= 40 → YES
```

The semantic question appears to concern differentiation.

This should be business-validated because high similarity normally means less differentiation.

---

## 110.7 Q19 polarity

The function is named:

```text
q19_outside_window_check
```

but visible logic produces `YES` for execution inside the allowed window plus tolerance.

This should be validated against the actual desired report semantics.

---

## 110.8 OCR failure precedence

`search.py` checks failure keywords before success keywords.

This matters because words such as:

```text
running
in progress
pending
```

can produce `FAILURE` even when another success term is also visible.

---

## 110.9 Fixed report directory

The Excel save layer is tied to:

```text
report/
```

This is simple but reduces flexibility for deployment.

---

## 110.10 Existing report reuse

Date-range execution can reuse an existing generated report workbook if one already exists at the generated location.

This should be understood before rerunning the process because report contents may be based on an existing workbook state.

---

# 111. Security Considerations

Because this is an enterprise ServiceNow integration, review at least:

```text
1. Credential storage
2. TLS verification (`verify=False`)
3. Attachment data downloaded to disk
4. Report directory permissions
5. Handling of potentially sensitive Incident/Problem/Change data
6. Temporary image and MSG cleanup
7. Logging of API responses / identifiers
```

The screenshots show enterprise/internal data patterns, so report files should be treated as potentially sensitive.

---

# 112. Performance Characteristics

The current architecture is mostly sequential.

A large audit can involve many API calls per Change Request.

For one CR, depending on its type, the tool may query:

```text
Change Request
Related Incidents
Related Problems
Approvals
Assignment Group
Group members
Delegates
CMDB CI
Attachments
Standard Proposal
Incident Major Incident state
```

Therefore the number of HTTP requests can become large as sample size grows.

The visible API timeouts are generally around:

```text
30 seconds
```

with some screenshot paths appearing shorter.

---

# 113. Why the Project Can Be Called “Automation”

It automates three manual activities that would otherwise be repetitive:

```text
1. Data collection
2. Audit decision checking
3. Report preparation
```

The code does not remove the business rules; it **encodes them**.

---

# 114. Difference Between the Three “Decision” Layers

It is useful to distinguish these:

### Layer 1 — Applicability

```text
Does Q12 apply to this CR?
```

Answer may be:

```text
NA
```

### Layer 2 — Rule evaluation

```text
If it applies, did the control pass?
```

Answer:

```text
YES / NO
```

### Layer 3 — Evidence interpretation

For some controls, the code first has to understand an attachment:

```text
OCR / MSG parsing
```

Then that evidence contributes to the final YES/NO.

---

# 115. Simplified Example — Normal Change

Suppose a Normal Change has:

```text
Risk = Low
Reason = Maintenance
Testing = documented
Required evidence = attached
```

The flow may be:

```text
Normal Change
   ↓
Q1 → testing/evidence rule
Q2 → only if non-testable
Q3 → reason keyword rule
Q4 → Maintenance ↔ SDLC alignment
Q5 → only if Incident Resolution
Q6 → only if Problem Resolution
Q7/Q8 → not applicable to normal → NA
Q9 → approver rule
Q10 → emergency only → NA
Q11 → expedited only → NA
Q12 → low-risk normal approval rule
Q13 → moderate/high only → NA
Q14 → standard-only → typically NA
Q15 → approval timing
Q16 → major incident
Q17 → post-production validation
Q18 → only Successful with Issues
Q19 → execution window
```

So a CR does **not** necessarily execute all rules in a meaningful way. Many return `NA` by design.

---

# 116. Simplified Example — Standard Change

A Standard Change follows a larger path:

```text
Standard CR
   ↓
Q1
Q5/Q6 when relevant
Q7 Standard Template
Q8 Template/description similarity
Q9 Approver structure
Q10/Q11/Q12/Q13 → mostly not applicable by type/risk
Q14 Standard approval
Q15 approval timing
Q16 major incident
Q17 post validation
Q18 successful-with-issues evidence
Q19 execution window
```

And the Standard Template sheet additionally checks:

```text
CI retired?
Pre-test plan?
Deployment plan?
Backout plan?
Verification plan?
Sample changes successful?
```

---

# 117. Simplified Example — Evidence-heavy Q17

```text
Closed Change
      ↓
Post-production summary present?
      ↓
Post-production results present?
      ↓
Attachment exists?
      ↓
Filename looks like post/deployment/prod evidence?
      ↓
Download image
      ↓
OpenCV preprocessing
      ↓
Tesseract OCR
      ↓
Normalize text
      ↓
Failure keywords?
      ├── YES → FAILURE
      └── NO
            ↓
       Success keywords?
            ├── YES → SUCCESS
            └── NO  → UNKNOWN
```

---

# 118. Simplified Example — Approval Timing

```text
CR
 ↓
Approved records
 ↓
Get latest/last approval timestamp
 ↓
Get Change start timestamp
 ↓
Compare
 ↓
Within required timing?
    ├── YES → YES
    └── NO
         ↓
     Look for approval MSG evidence
         ↓
     Parse sender + sent time
         ↓
     Check SLT/delegate relationship
         ↓
     Valid exception?
       ├── YES → YES
       └── NO  → NO
```

---

# 119. Main Call Graph

A simplified call graph is:

```text
main()
│
├── parse_arguments()
├── load_change_numbers()
│
├── run_change_number_audit()
│   ├── fetch_crs_from_servicenow()
│   └── create_change_wise_report()
│       ├── get_audit_data_single()
│       │   ├── get_common_data()
│       │   ├── get_questions_data()
│       │   └── get_std_questions_data() [Standard Template]
│       └── excel_handler.populate_sheet_single()
│
└── run_date_range_audit()
    ├── validate_date_range()
    ├── load_workbook()
    └── process_monthly_sheet()
        ├── random_select()
        │
        └── get_standard_template_proposals()
              [for SC Template sheet]
        │
        ├── get_basic_data()
        ├── get_audit_data()
        │   ├── get_common_data()
        │   ├── get_questions_data()
        │   └── get_std_questions_data()
        └── excel_handler.populate_sheet()
```

---

# 120. Main Dependency Graph

```text
populate_cr_audit.py
        │
        ├──────────────→ servicenow_api.py
        │                      │
        │                      ├──→ questions_data.py
        │                      │       └──→ search.py
        │                      │
        │                      └──→ sc_sheet_data.py
        │
        └──────────────→ excel_handler.py

constants.py
   ↑       ↑        ↑        ↑
   │       │        │        │
 API     Rules    OCR     Excel-related shared config
```

The code is therefore not a collection of isolated scripts; the modules form a clear dependency chain.

---

# 121. What Is NOT in This Codebase

Based on the supplied screenshots, the project does not visibly contain:

```text
- ML training
- model registry
- model inference
- embeddings
- neural networks
- vector database
- RAG
- LLM calls
- cloud OCR API
- browser automation
- ServiceNow UI automation
```

The architecture is primarily:

```text
REST API + business rules + OCR + Excel
```

---

# 122. Operational Runbook

## Step 1 — Prepare Python environment

Install the Python dependencies visible/required by the imports, including the core packages:

```text
requests
openpyxl
thefuzz / rapidfuzz
opencv-python
numpy
pytesseract
```

plus the dependency used by the MSG parser path.

The exact production `requirements.txt` was not present in the screenshot set, so dependency versions should be taken from the live environment rather than inferred here.

## Step 2 — Configure ServiceNow

Ensure the values referenced by `constants.py` are correct:

```text
BASE_URL
USERNAME
PASSWORD
headers
ATT_FILE_URL
```

## Step 3 — Ensure Tesseract exists

The screenshots show a local `tesseract/` directory with the executable/data.

## Step 4 — Ensure Excel template exists

Use the QAR template workbook with the expected sheet names.

## Step 5 — Run either explicit or date-range mode

Explicit:

```bash
python populate_cr_audit.py --change-nums CHG1234567 CHG1234568
```

Date range:

```bash
python populate_cr_audit.py --start-date 2026-09-01 --end-date 2026-09-30 --percentage 20
```

## Step 6 — Check output

Look in:

```text
report/
```

for the generated Excel workbook and any downloaded evidence files.

---

# 123. Troubleshooting Guide

## Problem: Template not found

Check:

```text
--template path
working directory
actual QAR workbook filename
```

## Problem: No CRs returned

Check:

```text
Change number
Date range
Change type
ServiceNow credentials
production_system filter
state filter
```

## Problem: Authentication/API failure

Check:

```text
BASE_URL
USERNAME
PASSWORD
headers
network access
TLS configuration
```

## Problem: OCR returns UNKNOWN

Check:

```text
image quality
Tesseract installation
tessdata
keyword dictionaries
attachment filename filtering
```

## Problem: Q17 gives unexpected FAILURE

Check whether the OCR text contains one of the failure-pattern words before the success word.

Because failure patterns are checked first, this can determine the final result.

## Problem: Standard Template check returns NO unexpectedly

Check:

```text
proposal type
associated Change count
close codes
fixed cutoff date
proposal review date
```

## Problem: Standard approval check does not behave as expected

Check:

```text
standard template exception list
group membership
actual approver records
approval state
CI owner
SLT/delegate data
```

---

# 124. Testing Strategy Present in the Codebase

`test_functions.py` is a lightweight/manual testing utility.

It visibly imports selected functions such as:

```text
q19_outside_window_check
q1_std_testing_completed
q4_reason_vs_sdlc
q12_normal_fcb_lowrisk_approval_check
q14_std_approval
```

It also has a helper to retrieve a Change by number directly from ServiceNow.

The intended workflow is:

```text
Enter one CR number
        ↓
Fetch CR
        ↓
Run selected question(s)
        ↓
Inspect printed result
```

This is useful for debugging a business rule without generating the entire workbook.

---

# 125. Recommended Testing Layers

For maintainability, the visible architecture naturally supports three levels of testing.

## Level 1 — Unit tests

Test individual functions with mocked Change dictionaries:

```text
q3
q4
q5
q6
q8
etc.
```

## Level 2 — ServiceNow integration tests

Use controlled test records or a test instance to verify:

```text
queries
relationships
approvers
attachments
```

## Level 3 — End-to-end report tests

Run the full pipeline and verify:

```text
correct sheet
correct Change column
correct question rows
correct summary counts
correct saved filename
```

---

# 126. Business Logic Summary Table

| Control | Main purpose | Key data |
|---|---|---|
| Q1 | Testing + required evidence | Test fields + attachments |
| Q2 | Non-testable CR SDLC evidence | Non-prod plan + SII/RITM |
| Q3 | Reason for Change keyword/business-rule validation | Reason field |
| Q4 | Reason vs SDLC alignment | Reason + SDLC type |
| Q5 | Incident association | Related INC |
| Q6 | Problem association | Related PRB |
| Q7 | Standard template applicability | Proposal/template data |
| Q8 | Standard template/description similarity | Short descriptions + RapidFuzz |
| Q9 | Requester/assignee not sole approver | Approvals + users |
| Q10 | Emergency approvals | Assignment group + SLT + approvals |
| Q11 | Expedited approvals | ITCM + CI owner + Assignment Group/SLT |
| Q12 | Normal low-risk approvals | ITCM + CI owner + Assignment Group |
| Q13 | Normal moderate/high approvals | CAB + ITCM |
| Q14 | Standard approval | Standard approval workflow |
| Q15 | Approval timing | Approval timestamp + Change start + evidence |
| Q16 | Major Incident relation | Incident table |
| Q17 | Post-production validation | Fields + image attachment + OCR |
| Q18 | Successful-with-Issues documentation | Closure code + P.I.R. MSG evidence |
| Q19 | Change execution window | Planned vs actual timestamps + 15-min tolerance |

---

# 127. Standard Template Rule Summary

| Check | Main source | Logic summary |
|---|---|---|
| CI retired | CMDB CI | Retired → YES, otherwise NO |
| Pre-test plan | Proposal/template fields | Required text + failure-keyword check |
| Deployment plan | `implementation_plan` | Required + no failure keywords |
| Backout plan | `backout_plan` | Required + no failure keywords |
| Verification plan | Post-prod summary | Required + no failure keywords |
| Sample changes | Proposal associated CRs | Required count + all successful |

---

# 128. The Project in Very Simple Words

Imagine a human auditor doing this manually:

```text
Open ServiceNow Change
        ↓
Look at testing
        ↓
Look at incidents/problems
        ↓
Look at approvals
        ↓
Look at CI owner
        ↓
Look at change dates
        ↓
Open attachments
        ↓
Read screenshots/emails
        ↓
Decide YES/NO/NA
        ↓
Type everything into Excel
```

This project simply automates that sequence:

```text
ServiceNow API
      ↓
Python rules
      ↓
Evidence/OCR
      ↓
Excel
```

---

# 129. The Most Important Files to Learn First

For a developer joining the project, the recommended reading order is:

```text
1. populate_cr_audit.py
2. servicenow_api.py
3. questions_data.py
4. sc_sheet_data.py
5. constants.py
6. search.py
7. excel_handler.py
8. QAR_report.xlsx
9. test_functions.py
```

Why:

```text
populate_cr_audit.py
    tells you the journey

servicenow_api.py
    tells you where the data comes from

questions_data.py
    tells you why a control becomes YES/NO

sc_sheet_data.py
    tells you the Standard Template rules

constants.py
    tells you the shared names/keywords

search.py
    tells you how image evidence is interpreted

excel_handler.py
    tells you where results go

QAR_report.xlsx
    shows the final business-facing output

test_functions.py
    shows how individual controls can be tested
```

---

# 130. Quick Reference — If You Need to Modify Something

| Requirement | Primary file |
|---|---|
| Add/change audit question | `questions_data.py` |
| Change Standard Template check | `sc_sheet_data.py` |
| Change ServiceNow table/query | `servicenow_api.py` |
| Change keyword | `constants.py` |
| Change OCR matching behavior | `search.py` |
| Change Excel row mapping | `excel_handler.py` |
| Change CLI | `populate_cr_audit.py` |
| Change sample selection | `populate_cr_audit.py` / `servicenow_api.py` |
| Change Standard proposal review selection | `populate_cr_audit.py` |
| Change output filename/path | `populate_cr_audit.py` / `excel_handler.py` |
| Debug one specific Q | `test_functions.py` |
| Change report layout | `QAR_report.xlsx` |

---

# 131. Final Architecture Summary

The codebase can be compressed into six logical layers:

```text
LAYER 1 — INPUT
CLI / CHG numbers / date range

        ↓

LAYER 2 — DATA ACQUISITION
ServiceNow REST API

        ↓

LAYER 3 — AUDIT LOGIC
Q1–Q19 + Standard Template rules

        ↓

LAYER 4 — EVIDENCE INTERPRETATION
Attachments + OCR + MSG parsing

        ↓

LAYER 5 — RESULT AGGREGATION
One dictionary per audited Change

        ↓

LAYER 6 — REPORTING
OpenPyXL → QAR_report.xlsx
```

The overall business pipeline is therefore:

```text
                 SERVICE NOW
                      │
                      ▼
              CHANGE REQUESTS
                      │
                      ▼
             SAMPLE / SELECTION
                      │
                      ▼
          ┌───────────────────────┐
          │   AUDIT RULE ENGINE   │
          │                       │
          │ Q1  Q2  Q3 ... Q19   │
          └───────────┬───────────┘
                      │
           ┌──────────┴──────────┐
           │                     │
           ▼                     ▼
     STRUCTURED DATA         EVIDENCE
                               │
                    ┌──────────┴──────────┐
                    ▼                     ▼
                   MSG                  IMAGE
                    │                     │
                 Parser                  OCR
                    │                     │
                    └──────────┬──────────┘
                               ▼
                         FINAL RESULTS
                               │
                               ▼
                         EXCEL QAR
                               │
                               ▼
                           REPORT/
```

---

# 132. Final Mental Model

Remember these three questions:

### 1. Where does the data come from?

```text
servicenow_api.py
```

### 2. How does the program decide pass/fail?

```text
questions_data.py
sc_sheet_data.py
```

### 3. Where does the result go?

```text
excel_handler.py → QAR_report.xlsx
```

And the complete controller around those three is:

```text
populate_cr_audit.py
```

So the whole project is:

> **ServiceNow data → deterministic audit rules → evidence/OCR where required → Excel QAR report.**

---

# 133. Source-Fidelity Notes

This documentation was reconstructed from the supplied screenshot codebase rather than the original source files.

Therefore:

1. The **architecture and main function names/relationships** are strongly supported by the screenshots.
2. The **row mappings and visible constants** are documented from the readable source.
3. A small number of exact literals are marked for verification where the screenshot/OCR was ambiguous.
4. Secret credentials and environment-specific URLs are intentionally not reproduced.
5. Q14 is documented at the level supported by the visible screenshots because some exact implementation lines were obscured.
6. The exact numeric Standard Template review-age cutoff should be verified in the live source because the screenshot evidence/OCR was inconsistent.

---

# 134. One-Page Cheat Sheet

```text
PROJECT:
ServiceNow Change Management Audit Automation

MAIN:
populate_cr_audit.py

DATA SOURCE:
ServiceNow REST APIs

MAIN DATA MODULE:
servicenow_api.py

AUDIT ENGINE:
questions_data.py

STANDARD TEMPLATE ENGINE:
sc_sheet_data.py

OCR:
search.py + local Tesseract

EXCEL:
excel_handler.py + QAR_report.xlsx

INPUT:
CHG numbers OR date range + sample percentage

OUTPUT:
report/*.xlsx

AUDIT RESULT:
YES / NO / NA / ERROR

MAIN CONTROLS:
Q1–Q19

SPECIAL EVIDENCE:
ServiceNow attachments, screenshots, MSG approval emails

SPECIAL TECHNIQUES:
REST API
RapidFuzz token similarity
OpenCV preprocessing
Tesseract OCR
OpenPyXL workbook population

CORE FLOW:
INPUT
 ↓
FETCH
 ↓
SELECT
 ↓
AUDIT
 ↓
EVIDENCE
 ↓
AGGREGATE
 ↓
EXCEL
 ↓
REPORT
```

---

## End of Documentation

---

# 135. Screenshot-to-Module Coverage Map

The 93 supplied screenshots were read as one continuous codebase snapshot. The visible sequence maps approximately as follows:

| Screenshot range | Main artifact visible |
|---|---|
| 1 | Project/file tree |
| 2 | `constants.py` |
| 3 | `changes.txt` |
| 4–7 | `excel_handler.py` |
| 8–26 | `populate_cr_audit.py` |
| 27–41 | `QAR_report.xlsx` / workbook structure and revision controls |
| 42–74 | `questions_data.py` |
| 75–79 | `sc_sheet_data.py` / Standard Template logic |
| 80–83 | `search.py` / OCR logic |
| 84–92 | `servicenow_api.py` |
| 93 | `test_functions.py` |

This map is useful when manually checking the original screenshots against this documentation.

---

# 136. Module-by-Module Function Inventory

## `populate_cr_audit.py`

Visible/strongly supported functions and logic:

```text
parse_arguments()
load_change_numbers(...)
validate_date_range(...)
report_filename(...)
get_reference_display_value(...)
print_missing_change_numbers(...)
load_workbook(...)
populate_all_basic_data(...)
create_change_wise_report(...)
run_change_number_audit(...)
get_standard_template_proposals(...)
process_monthly_sheet(...)
run_date_range_audit(...)
main()
```

## `servicenow_api.py`

Visible/strongly supported functions and logic:

```text
fetch_crs_from_servicenow(...)
random_select(...)
get_proposals(...)
get_basic_data(...)
get_audit_data(...)
get_audit_data_single(...)
get_all_related_incidents(...)
get_all_related_problems(...)
get_common_data(...)
get_questions_data(...)
get_std_questions_data(...)
```

The file also contains the individual ServiceNow GET logic used by those aggregators.

## `questions_data.py`

Visible/strongly supported audit functions:

```text
get_closure_status(...)
q1_std_testing_completed(...)
q2_nc_exp_sdlc_ritm(...)
q3_nc_exp_business_project_patching(...)
q4_reason_vs_sdlc(...)
q5_std_incident(...)
q6_std_problem(...)
q7_std_template(...)
q8_differentiation(...)
q9_approver_check(...)
q10_emergency_approval_check(...)
q11_expedited_fcb_approval_check(...)
q12_normal_fcb_lowrisk_approval_check(...)
q13_normal_modhigh_risk_approval_check(...)
q14_std_approval(...)
q15_approval_timing(...)
q16_major_incident_check(...)
q17_post_validation_check(...)
q18_success_issues_check(...)
q19_outside_window_check(...)
```

Supporting helpers visible in the screenshots include logic for:

```text
get_change_management_approvers()
get_actual_change_approvers(...)
download_attachment(...)
get_sent_time_app_name_from_msg(...)
get_delegates(...)
get_sys_id(...)
```

## `sc_sheet_data.py`

Visible functions:

```text
convert_to_datetime(...)
is_retired(...)
pre_test_plan(...)
deployment_plan(...)
backout_plan(...)
verification_plan(...)
fetch_cr_close_codes(...)
sample_changes(...)
```

## `search.py`

Visible functions:

```text
preprocess_image(...)
normalize_text(...)
extract_text(...)
analyze_dashboard_status(...)
```

## `excel_handler.py`

Visible functions:

```text
load_template(...)
write_rows(...)
populate_basic_data(...)
populate_cr_column(...)
populate_sheet(...)
populate_sheet_single(...)
save_output(...)
```

## `test_functions.py`

Visible helper/functionality:

```text
get_change_by_number(...)
manual execution of selected q* functions
```

---

# 137. Exact High-Level Control Routing

The questions are not all universally applicable.

A useful routing table is:

```text
                    NORMAL   STANDARD   EMERGENCY   EXPEDITED
Q1 Testing            ✔         ✔          ✔           ✔
Q2 Non-testable       ✔         -          -           ✔
Q3 Reason             ✔         -          -           ✔
Q4 Reason/SDLC        ✔         -          -           ✔
Q5 Incident           based on reason for change
Q6 Problem            based on reason for change
Q7 Std Template       -         ✔          -           -
Q8 Differentiation    -         ✔          -           -
Q9 Approver           ✔         ✔*         ✔           ✔
Q10 Emergency         -         -          ✔           -
Q11 Expedited         -         -          -           ✔
Q12 Low Risk          ✔**       -          -           -
Q13 Mod/High Risk     ✔***      -          -           -
Q14 Std Approval      -         ✔          -           -
Q15 Approval Timing   ✔         ✔          ✔           ✔
Q16 Major Incident    ✔         ✔          ✔           ✔
Q17 Post Validation   ✔         ✔          ✔           ✔
Q18 Issues Doc        conditional on closure code
Q19 Window Timing     ✔         ✔          ✔           ✔
```

Where:

```text
* Standard-specific exception templates can make Q9 NA.
** Q12 is specifically Low Risk Normal.
*** Q13 is specifically Moderate/High Risk Normal.
```

This table is a mental model, not a substitute for the individual function conditions.

---

# 138. Control-Engine Pattern by Data Type

The project uses five main categories of evidence.

## A. Direct Change fields

Examples:

```text
number
type
category
risk
state
close_code
reason_for_change
SDLC initiative type
planned dates
actual/work dates
testing fields
post-production fields
```

## B. Reference objects

Examples:

```text
assignment_group
assigned_to
requested_by
cmdb_ci
std_change_producer_version
```

These often arrive from ServiceNow as dictionaries containing:

```text
display_value
value
link
```

## C. Related records

Examples:

```text
incidents
problems
approvals
group memberships
delegates
```

## D. File attachments

Examples:

```text
PNG/JPG/JPEG evidence
MSG approval emails
```

## E. Parsed content

Examples:

```text
OCR text
approval email sender
approval email sent timestamp
```

The final audit value is produced by combining some of these categories according to each question.

---

# 139. ServiceNow Relationship Philosophy

A notable design feature is that the code does not assume that all useful audit evidence is stored directly on the Change Request.

It follows relationships:

```text
Change
 ├──→ Incident
 ├──→ Problem
 ├──→ Approval
 ├──→ Assignment Group
 │      └──→ Approval Group
 │              └──→ Members
 ├──→ CI
 │      └──→ Owner
 ├──→ Standard Change Proposal
 └──→ Attachments
```

This is why the API layer contains many small queries.

---

# 140. Why Reference-field `display_value` Matters

ServiceNow reference fields commonly contain both an internal identifier and a human-readable value.

The code therefore frequently requests:

```text
sysparm_display_value=true
```

and then uses:

```text
display_value
```

for comparisons and reporting.

Example concept:

```text
assignment_group.value
        = internal sys_id

assignment_group.display_value
        = human-readable group name
```

The audit logic normally needs the latter for business comparisons while sometimes following `link` URLs or using `value`/sys_id for subsequent API calls.

---

# 141. Why the Project Has Both Names and Sys-IDs

The code uses three forms of identity depending on the operation:

```text
Display name
→ compare with human-readable approver/group name

Sys-ID / value
→ exact ServiceNow record identity

Reference link
→ retrieve the referenced object
```

This is visible in the approval, group, CI, proposal, and delegate logic.

---

# 142. Why Approval Checks Can Be Complex

Suppose the business says:

```text
“ITCM must approve.”
```

The code cannot simply look for a text field saying “ITCM”.

Instead it effectively does:

```text
Find ITCM group
      ↓
Find its members
      ↓
Find actual approved CR approvers
      ↓
Compare lists
```

That pattern is repeated for several approval roles.

This explains why the approval code is much larger than a simple `if` statement.

---

# 143. Why Attachments Are Downloaded Instead of Just Named

Some controls need to verify the **content** of evidence rather than only its existence.

For example Q17 needs to determine whether an attached dashboard screenshot indicates a successful or failed system status.

Therefore:

```text
Attachment metadata
      ↓
filename looks relevant
      ↓
download binary file
      ↓
process file content
```

Similarly Q15 may need the content/metadata of a `.msg` email so that sender and sent timestamp can be evaluated.

---

# 144. The Two Evidence Parsers

The project has two separate evidence-processing paths.

### Image evidence

```text
OpenCV
→ Tesseract
→ keyword status
```

### Email/MSG evidence

```text
MSG parser
→ sender
→ sent time
```

They solve different problems:

```text
Image parser:
“What does the dashboard screenshot say?”

MSG parser:
“Who sent the approval email and when?”
```

---

# 145. Report Generation Is Template-Driven

The workbook is effectively part of the application's API.

The Python code expects:

```text
specific sheet names
specific row positions
specific column structure
```

Therefore changing the Excel layout without changing `excel_handler.py` can silently break the report.

The safest model is:

```text
QAR_report.xlsx
        ↕
excel_handler.py
```

Treat them as a paired component.

---

# 146. Data Contract Between Audit Layer and Excel Layer

The audit engine produces keys such as:

```text
q1_testing_completed
q2_nc_exp_sdlc_ritm
...
q19_outside_window
```

The Excel layer knows exactly which row each key belongs to.

Therefore:

```text
Question function name/result key
           ↓
Dictionary key
           ↓
Excel row mapping
           ↓
Workbook cell
```

If a key is renamed in one module but not in `excel_handler.py`, that control will no longer populate correctly.

This is one of the most important maintenance contracts in the codebase.

---

# 147. Minimal Developer Understanding

A developer does not need to memorize every API query to understand the application.

They need to memorize:

```text
1. populate_cr_audit.py = orchestration
2. servicenow_api.py = data retrieval/aggregation
3. questions_data.py = Q1–Q19
4. sc_sheet_data.py = Standard Template checks
5. search.py = OCR evidence
6. excel_handler.py = Excel mapping
7. constants.py = shared config
```

After that, tracing any bug becomes much easier.

---

# 148. Debugging by Symptom

```text
Wrong CR population
    → servicenow_api.py / query filters

Wrong sample size
    → random_select() / process_monthly_sheet()

Wrong Standard Template selection
    → get_standard_template_proposals()

Wrong YES/NO for a control
    → corresponding q* function

Wrong approval result
    → questions_data approval helper + ServiceNow group queries

Wrong OCR result
    → search.py + constants keyword lists

Wrong Standard Template result
    → sc_sheet_data.py

Correct result but wrong Excel cell
    → excel_handler.py / row maps

Correct workbook but wrong sheet
    → populate_cr_audit.py routing

Wrong output location/name
    → report_filename() / save_output()
```

---

# 149. Recommended Future Refactoring Areas

These are engineering improvement ideas derived from the visible architecture, not requirements that the current system necessarily needs.

## Configuration

Move credentials and environment URLs out of source constants into secure environment/secret management.

## HTTP client wrapper

Create one reusable ServiceNow client so that authentication, headers, timeout, TLS verification, retries, and logging are not repeated.

## Caching

Cache repeated group/CI/user/proposal lookups during one audit run.

## Rules as data

Some keyword-based rules could be represented in configuration rather than hard-coded functions.

## Unit tests

Add automated tests for all Q1–Q19 with controlled input dictionaries and mocked API responses.

## Typed models

Replace loosely structured dictionaries with dataclasses/Pydantic models if the project grows.

## Logging

Replace scattered `print(...)` calls with structured logging.

## OCR confidence

Record OCR confidence/recognized text to make evidence decisions more explainable.

## Temporary-file cleanup

Explicitly clean downloaded attachments after report generation when retention is not required.

---

# 150. Final “Explain It in an Interview” Version

A concise technical explanation of the project is:

> The project is a Python-based ServiceNow Change Management audit automation system. It supports both direct Change-number auditing and date-range sampling. It fetches Change Requests and related Incidents, Problems, approvals, CI information, Standard Change proposals, group membership, delegates, and attachments through ServiceNow REST APIs. The audit layer applies 19 deterministic business controls plus separate Standard Change Template checks. Certain controls use attachment evidence: image evidence is processed locally with OpenCV and Tesseract OCR, while approval email evidence can be parsed from MSG files for sender and timestamp validation. The resulting audit dictionary is mapped into a predefined QAR Excel workbook using OpenPyXL, producing change-specific result sheets and a Standard Change Template audit sheet.

---

# 151. Final One-Liner

```text
ServiceNow Change data
        → selection
        → rule-based audit
        → attachment evidence / OCR
        → YES/NO/NA/ERROR
        → QAR Excel report
```
