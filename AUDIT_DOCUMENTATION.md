# Change Management Audit Automation (CMAA) — Complete Documentation

A simple, concise, and to-the-point guide to the codebase architecture, file-by-file breakdown, and detailed explanation of all 19 audit questions.

---

## 1. What is CMAA? (High-Level Overview)

**CMAA (Change Management Audit Automation)** is an automated quality assurance reporting tool. 
- **Goal:** Automates the auditing of IT Change Requests (CRs) in **ServiceNow** against ITIL and organizational compliance policies.
- **How it works:**
  1. Pulls Change Requests from ServiceNow via Table REST API.
  2. Runs 19 automated audit compliance questions (Q1 to Q19) + Standard Change Template checks.
  3. Formats and writes the results into a multi-tab Excel workbook (`QAR Report.xlsx`).
  4. Automatically applies color-coding for audit exceptions and mismatches.

---

## 2. File-by-File Breakdown

### `populate_cr_audit.py` (CLI & Main Workflow Orchestrator)
- **Role:** Entry point for executing audits.
- **Key Functions:**
  - `parse_arguments()`: Parses CLI arguments (`--start-date`, `--end-date`, `--percentage`, `--change-nums`, `--template`, `--tester`, `--output`).
  - `load_change_numbers()`: Validates CR IDs against regex (`CHG\d+`) from command-line strings or text files (e.g., `changes.txt`).
  - `run_change_number_audit()`: **Mode 1** — Audits specific CRs and places each into its respective sheet (Normal, Standard, Emergency, Expedited).
  - `run_date_range_audit()`: **Mode 2** — Queries all closed production changes in a date range, randomly samples them (default 20%), audits them, and outputs a monthly QA report.
  - `report_filename()`: Generates report names (e.g., `FCB October'26 - QAR Report.xlsx` or `FCB October-QAR Change Wise.xlsx`).

### `servicenow_api.py` (ServiceNow API Client & Data Assembly)
- **Role:** Handles HTTP communication with ServiceNow REST APIs and compiles audit payloads.
- **Key Functions:**
  - `_get()`: Sends authenticated HTTP GET requests with basic auth, timeout, and SSL verification settings.
  - `fetch_crs_from_servicenow()`: Queries `change_request` table using sysparm queries (by list of CR numbers or date range + state filter `stateIN-2,-1,0,3^production_system=true`).
  - `random_select()`: Selects a percentage-based random sample from a list of CRs.
  - `get_proposals()`: Navigates standard change producer version links to fetch the underlying proposal record.
  - `get_all_related_incidents()` / `get_all_related_problems()`: Traverses `u_related_incidents`, `u_related_problems`, and queries `incident` / `problem` tables for linked records.
  - `get_audit_data()` / `get_audit_data_single()`: Merges common CR metadata with question results (Q1–Q19) and standard template checks.

### `questions_data.py` (Core Compliance & Audit Engine)
- **Role:** Implements the core logic for all 19 audit questions (Q1–Q19) and approval validation rules.
- **Key Functions:**
  - `q1_std_testing_completed` & `q1_std_testing_completed_new`: Pre-implementation testing check.
  - `q2_nc_exp_sdlc_ritm`: SDLC initiative / RITM tracking for untestable changes.
  - `q3_nc_exp_business_project_patching`: Reason for change validation.
  - `q4_reason_vs_sdlc`: Reason vs. SDLC initiative alignment check.
  - `q5_std_incident` & `q6_std_problem`: Incident/Problem association checks.
  - `q7_std_template` & `q8_differentiation`: Standard change template validity and fuzzy differentiation.
  - `q9_approver_check`: Separation of duties (requester/assignee cannot be sole approver).
  - `q10` to `q14`: Change-type specific approval rules (Emergency, Expedited, Normal Low, Normal High/Moderate, Standard).
  - `q15_approval_timing`: Approvals completed prior to implementation start.
  - `q16_major_incident`: Association of major incidents caused by the change.
  - `q17_post_validation_check` & `q17_post_validation_check_new`: Post-deployment validation artifact verification.
  - `q18_issues_doc`: Documentation for unsuccessful changes or changes with issues.
  - `q19_outside_window`: Execution within scheduled change window.
  - `download_attachment()`: Utility to download attachments via ServiceNow Attachment API.

### `sc_sheet_data.py` (Standard Change Template Evaluation)
- **Role:** Specific checks for the `SC Template Result Sheet`.
- **Key Functions:**
  - `is_retired()`: Checks if the target CMDB Configuration Item (CI) is marked as "Retired".
  - `pre_test_plan()`, `deployment_plan()`, `backout_plan()`, `verification_plan()`: Inspects the standard change template definition text for missing plans or placeholder text ("n/a", "not required", etc.).
  - `sample_changes()`: Verifies that proposals have required historical successful changes (1 for Modify, 4 for New proposals created after Sep 8, 2025).

### `excel_handler.py` (Excel Reporting & Formatting)
- **Role:** Reads `QAR Report.xlsx` template, maps audit data into cells, applies formatting/highlighting, and saves output.
- **Key Mappings:**
  - `BASIC_DATA_ROWS` (rows 3–7): Header metadata (month, tester, population, sample count).
  - `GENERAL_DATA_ROWS` (rows 10–17): CR number, date, change type, category, close code, related incident/problem/jira.
  - `QUESTION_ROWS` (rows 18–38): Q1 to Q19 results.
  - `STD_Q_ROWS` (rows 10–23): Standard change template evaluation rows.
- **Color Fills:**
  - `KWNF` (Keyword Not Found) -> Light Yellow (`#FFF2CC`)
  - `MM` (Maintenance Mismatch) -> Light Red (`#F4CCCC`)
  - `DM` (DR Mismatch) -> Light Blue (`#DDEBF7`)
  - `MA` (Missing Approval) -> Light Purple (`#E4DFEC`)

### `constants.py` (Central Configuration & Lookup Tables)
- **Role:** Holds environment config, Table names, status literals, keyword lists, and regex patterns.
- **Key Items:**
  - ServiceNow tables: `change_request`, `sysapproval_approver`, `sys_attachment`, `sys_user_group`, `sys_user_grmember`, `sys_user_delegate`, `cmdb_ci`, `incident`, `problem`.
  - Output values: `Yes`, `No`, `N/A`, `Error`, `not testable`, `Keyword not found` (`KWNF`), `maintenance mismatch` (`MM`), `dr mismatch` (`DM`), `Missing Approval` (`MA`).
  - Lists of fail/pass/pre-prod/post-prod keywords.
  - Approval exception template names.

### `test_functions.py` & `validate_reconstruction.py`
- `validate_reconstruction.py`: Fast, offline syntax and unit assertions (regex validation, date ranges, input parsing).
- `test_functions.py`: Live smoke-test harness against a ServiceNow instance.

---

## 3. Detailed Understanding of All 19 Questions

Below is the simple, to-the-point explanation of every question (Q1 through Q19):

| Question # | Question Title & Intent | Applicable Types | Pass Criteria (`Yes`) | Fail / Exception Criteria |
|---|---|---|---|---|
| **Q1** | **Testing Completed & Documented**<br>Verifies pre-implementation testing was documented and artifacts attached. | All | Summary and results fields are filled AND an attachment with test keywords (e.g. `test results`, `non prod`) is present. | - `not testable`: if flagged untestable with valid reason.<br>- `No`: if fields empty or test attachment missing. |
| **Q1 (New)** | **Testing Completed (Alternate/Revised)**<br>Stricter attachment filename inspection. | All | Summary and results present AND attachment contains specific evidence terms (`test result`, `test evidence`, `test cycle report`). | `No` if missing. |
| **Q2** | **Untestable CR SDLC Tracking**<br>If a change cannot be tested, is an SDLC SII or RITM number provided? | Normal, Expedited | If untestable: `u_non_prod_test_plan_summary` contains `"sii"` or `"ritm"`. | - `N/A`: if testable or type not Normal/Expedited or DR/Maintenance.<br>- `No`: if missing. |
| **Q3** | **Reason for Change Active & Accurate**<br>Validates business project or patching reason. | Normal, Expedited | If reason is `business project` or `patching`, `short_description` must be non-empty. | - `N/A`: other reasons or types.<br>- `No`: reason blank or short description missing. |
| **Q4** | **Reason vs. SDLC Initiative Alignment**<br>Ensures change reason matches SDLC category. | Normal, Expedited | - Maintenance reason matches `maintenance - initiative not applicable`.<br>- DR Exercise reason matches `dr exercise - initiative not applicable`. | - `MM` (Maintenance Mismatch): reason is maintenance but SDLC is not.<br>- `DM` (DR Mismatch): reason is DR but SDLC is not.<br>- `N/A`: all other reasons. |
| **Q5** | **Incident Association**<br>If change was caused by an incident, is the INC linked? | All | If reason is `Incident Resolution`, `u_related_incidents` is populated. | - `N/A`: reason is not incident resolution.<br>- `No`: related incident field is empty. |
| **Q6** | **Problem Association**<br>If change was for problem resolution, is the PRB linked? | All | If reason is `Problem Resolution`, `u_related_problems` is populated. | - `N/A`: reason is not problem resolution.<br>- `No`: related problem field is empty. |
| **Q7** | **Standard Change Template Applicability**<br>Confirms standard change uses an active, matching template. | Standard | Template exists and CR short description matches proposal template short description. | - `N/A`: non-standard changes or 4 whitelist exception templates.<br>- `No`: template missing or description mismatch. |
| **Q8** | **Standard Change Differentiation**<br>Ensures CR short description differentiates itself from generic template name. | Standard | Fuzzy token sort similarity between CR short description and template name is $\ge 40\%$. | - `N/A`: non-standard changes.<br>- `No`: description empty or similarity $< 40\%$. |
| **Q9** | **Separation of Duties (Approver Check)**<br>Initiator and Assignee must not be the sole approver. | All | Approver list contains at least one person who is NEITHER `requested_by` NOR `assigned_to`. | `No` if sole approver is requester/assignee or no approvals exist. |
| **Q10** | **Emergency Change Approvals**<br>Requires both Assignment Group Approver and Group SLT. | Emergency | Actual approvers include both Assignment Group SLT (`u_group_slt`) AND a group member. | - `N/A`: non-emergency changes.<br>- `MA` (Missing Approval): missing SLT or group member approval. |
| **Q11** | **Expedited Change Approvals**<br>Requires 3-way signoff. | Expedited | Approved by: 1) ITCM (`Change Management Approval`), 2) Assignment Group SLT, and 3) CI Owner (`owned_by`). | - `N/A`: non-expedited changes.<br>- `MA` (Missing Approval): any of the 3 approvals missing. |
| **Q12** | **Normal Change (Low Risk) Approvals**<br>Requires peer/group signoff. | Normal (Low Risk) | Approvals include at least one member of the Assignment Group. | - `N/A`: non-normal or risk $\ne$ low.<br>- `MA`: no group member approved. |
| **Q13** | **Normal Change (Moderate/High Risk) Approvals**<br>Requires leadership & CAB signoff. | Normal (Moderate / High Risk) | Approvals include Assignment Group SLT (excluding requester and assignee). | - `N/A`: low risk or non-normal.<br>- `MA`: SLT approval missing. |
| **Q14** | **Standard Change Peer Approval**<br>Peer signoff within assignment group. | Standard | Approvals include a group member other than requester and assignee. | - `N/A`: non-standard or exception templates.<br>- `MA`: missing peer approval. |
| **Q15** | **Approval Timing vs. Implementation Start**<br>Were approvals completed BEFORE work started? | All | `actual_start` (or `work_start`) $\ge$ timestamp of the final approval. | - `No`: implementation started before approvals finished.<br>- `MA`: no approvals found.<br>- `N/A`: missing timestamps. |
| **Q16** | **Major Incident Association**<br>If change caused a Major Incident, is it linked? | All | Incident table query confirms an incident with `caused_by = CR.sys_id` AND `major_incident = true`. | - `Yes` if linked.<br>- `No` if no record found.<br>- `N/A` if CR sys_id missing. |
| **Q17** | **Post-Deployment Validation Evidence**<br>Attachment verifying post-deployment health check. | All | Attachment filename contains post-prod keywords (`post prod`, `prod deployment`, `post implementation`, `prod validation`). | `KWNF` (Keyword Not Found): no matching attachment keyword found. |
| **Q17 (New)** | **Post-Deployment Validation (Revised)**<br>Stricter validation terms inspection. | All | Attachment contains `implementation review` or both `post` and `validation`. | `KWNF`: keyword not found. |
| **Q18** | **Issues Documentation (Unsuccessful Changes)**<br>Were issues/backouts documented if CR failed? | Closed Changes | If close code is NOT `successful`, an attachment with issue keywords (`issue`, `incident`, `defect`, `problem`) is attached. | - `N/A`: state not closed or close code was `successful`.<br>- `No`: failed change without issue attachment. |
| **Q19** | **Execution Within Approved Window**<br>Did implementation stay inside planned start/end? | All | `actual_start` and `actual_end` fall completely within `planned_start` and `planned_end`. | - `No`: started early, started late, or ran past planned end.<br>- `N/A`: missing planned or actual timestamps. |

---

## 4. Special Audit Codes & Cell Fills

| Code | Meaning | Cell Background Color | Trigger Condition |
|---|---|---|---|
| `KWNF` | Keyword Not Found | **Light Yellow** (`#FFF2CC`) | Q17 post-deployment validation attachment keyword missing |
| `MM` | Maintenance Mismatch | **Light Red** (`#F4CCCC`) | Q4 Reason is Maintenance but SDLC initiative is not |
| `DM` | DR Mismatch | **Light Blue** (`#DDEBF7`) | Q4 Reason is DR Exercise but SDLC initiative is not |
| `MA` | Missing Approval | **Light Purple** (`#E4DFEC`) | Q10–Q15 required approver role/group was not found in approvals |
| `not testable`| Not Testable | Normal | Q1 change cannot be pre-tested with documented reason |

---

## 5. How to Run the Tool

### Setup
```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
# Fill SN_INSTANCE_URL, SN_USERNAME, SN_PASSWORD in .env
```

### Run Mode 1: Specific Change Numbers
```bash
python populate_cr_audit.py --template "QAR Report.xlsx" --change-nums CHG6892954 CHG6957477
# Or using a file:
python populate_cr_audit.py --template "QAR Report.xlsx" --change-nums changes.txt
```

### Run Mode 2: Date-Range Monthly Audit Sampling
```bash
python populate_cr_audit.py --template "QAR Report.xlsx" --start-date 2026-01-01 --end-date 2026-03-31 --percentage 20 --tester "Auditor Name"
```

All generated workbooks are saved automatically in the `report/` folder.
