# Change Management Audit Automation (CMAA)

Automated audit & quality assurance tool for ITIL Change Requests in ServiceNow.

CMAA queries ServiceNow Change Requests via REST API, evaluates 19 compliance & control questions, and populates the results into a standardized Excel report (`QAR Report.xlsx`).

---

## Quick Navigation

- **[Detailed Audit & Questions Documentation](AUDIT_DOCUMENTATION.md)** — Complete file-by-file architecture and in-depth breakdown of all 19 questions (Q1–Q19).
- **[Reconstruction Report](RECONSTRUCTION_REPORT.md)** — Historical context and reconstruction notes from source artifacts.

---

## Project Structure

| File | Purpose |
|---|---|
| [`populate_cr_audit.py`](populate_cr_audit.py) | Main CLI runner for change-wise and date-range monthly audits |
| [`servicenow_api.py`](servicenow_api.py) | ServiceNow REST API communication and audit data extraction |
| [`questions_data.py`](questions_data.py) | Compliance logic for all 19 audit questions (Q1–Q19) |
| [`sc_sheet_data.py`](sc_sheet_data.py) | Standard Change Template governance checks |
| [`excel_handler.py`](excel_handler.py) | Excel workbook reading, population, styling, and saving |
| [`constants.py`](constants.py) | Shared constants, table names, status codes, and keywords |
| [`QAR Report.xlsx`](QAR%20Report.xlsx) | Excel audit template containing 5 result sheets + Revision Control |
| [`validate_reconstruction.py`](validate_reconstruction.py) | Local offline unit checks and rule assertions |
| [`test_functions.py`](test_functions.py) | Manual smoke-test script for live ServiceNow connections |
| [`changes.txt`](changes.txt) | Sample change numbers list for testing |
| [`requirements.txt`](requirements.txt) | Python dependencies |

---

## 19 Audit Questions Summary

1. **Q1 / Q1 (New)**: Pre-implementation testing completed, documented, and evidence attached.
2. **Q2**: Untestable Normal/Expedited changes have SDLC SII / RITM reference.
3. **Q3**: Active & accurate reason for change (business project / patching).
4. **Q4**: Reason for change aligns with SDLC initiative (Maintenance / DR checks).
5. **Q5**: Incident associated if reason is "Incident Resolution".
6. **Q6**: Problem associated if reason is "Problem Resolution".
7. **Q7**: Standard Change template matches change proposal.
8. **Q8**: Standard Change short description differentiates from generic template ($\ge 40\%$ fuzzy match).
9. **Q9**: Separation of duties: Requester/Assignee is not the sole approver.
10. **Q10**: Emergency Change approved by Assignment Group Approver AND SLT.
11. **Q11**: Expedited Change approved by ITCM, Assignment Group SLT, and CI Owner.
12. **Q12**: Normal Low Risk Change approved by Assignment Group member.
13. **Q13**: Normal High/Moderate Risk Change approved by Assignment Group SLT.
14. **Q14**: Standard Change approved by Assignment Group peer (non-initiator).
15. **Q15**: Approvals completed prior to implementation start.
16. **Q16**: Major Incident linked if caused by the change.
17. **Q17 / Q17 (New)**: Post-deployment validation evidence attached.
18. **Q18**: Issue/backout documentation attached for unsuccessful changes.
19. **Q19**: Implementation executed within scheduled change window.

---

## Quickstart

### 1. Installation
```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Configure Environment
Copy `.env.example` to `.env` and fill in credentials:
```bash
SN_INSTANCE_URL=https://your-instance.service-now.com
SN_USERNAME=your_username
SN_PASSWORD=your_password
SN_VERIFY_SSL=false
```

### 3. Run Audit

**By Change Numbers:**
```bash
python populate_cr_audit.py --template "QAR Report.xlsx" --change-nums CHG6892954 CHG6957477
# Or from file:
python populate_cr_audit.py --template "QAR Report.xlsx" --change-nums changes.txt
```

**By Date Range (Random Sampling):**
```bash
python populate_cr_audit.py --template "QAR Report.xlsx" --start-date 2026-01-01 --end-date 2026-03-31 --percentage 20 --tester "Auditor Name"
```

Outputs are automatically saved in the `report/` folder.
