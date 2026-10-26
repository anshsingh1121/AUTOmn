# ServiceNow Change Management Audit Automation — Technical Q&A

This document provides concrete, line-by-line technical answers to the Change Management Audit Automation (CMAA) architectural, API, and validation questions.

---

### **1. What is the ServiceNow URL?**

In [`constants.py`](constants.py#L17-L22), the ServiceNow instance URL and base endpoints are configured via environment variables:

- **Instance URL:**
  ```python
  INSTANCE_URL = os.getenv("SN_INSTANCE_URL", "https://your-instance.service-now.com").rstrip("/")
  ```
  *(Default: `https://<your-instance>.service-now.com`)*

- **Table REST API Base URL:**
  ```python
  BASE_URL = f"{INSTANCE_URL}/api/now/table/"
  ```
  *(Full endpoint: `https://<your-instance>.service-now.com/api/now/table/`)*

- **Attachment API Base URL:**
  ```python
  ATT_FILE_URL = f"{INSTANCE_URL}/api/now/attachment/"
  ```
  *(Full endpoint: `https://<your-instance>.service-now.com/api/now/attachment/`)*

---

### **2. After retrieving the change request details, are you getting only 20% of the changes or all change requests within the specified date range?**

- **From ServiceNow (The API call):** You retrieve **ALL (100%)** of the Change Requests within the specified date range that match the filter criteria (`type=<change_type>`, `stateIN-2,-1,0,3`, and `production_system=true`).
- **In Python Memory (Local sampling):** Once the complete population is loaded into memory, [`servicenow_api.py:random_select()`](servicenow_api.py#L43-L48) calculates the sample size:
  $$\text{Sample Size} = \max\left(1, \left\lceil \text{len(population)} \times \left(\frac{20}{100}\right) \right\rceil\right)$$
  and uses Python's built-in `random.sample()` to select the **20% subset**.
- **Reason:** The audit workbook requires both the **Total Population** (100%) and the **Samples Tested** (20%) to be recorded in the Excel header metadata (rows 6 and 7 of each sheet).

---

### **3. What is the URL used to retrieve the 20% of change request details?**

There is **no separate ServiceNow API URL to query 20% directly**. ServiceNow does not support random percentage sampling in sysparm queries.

The API URL fetches **100% of the population for the date range**, and Python samples the 20% locally:

```http
GET https://<your-instance>.service-now.com/api/now/table/change_request?sysparm_query=type=<change_type>^end_dateBETWEENjavascript:gs.dateGenerate('<start_date>','00:00:00')@javascript:gs.dateGenerate('<end_date>','23:59:59')^stateIN-2,-1,0,3^production_system=true&sysparm_display_value=true
```

*(For manual change-number audits via `--change-nums`, the URL queries exact CR numbers:)*
```http
GET https://<your-instance>.service-now.com/api/now/table/change_request?sysparm_query=numberIN<CHG1>,<CHG2>,...&sysparm_display_value=true
```

---

### **4. Are you calling the ServiceNow API for each change request or for each question? What is the final API URL being used?**

- **Change Request Payload:** Retrieved in **one batch call** per change type for all matching CRs upfront.
- **Audit Questions:** The API is called **per Change Request on-demand**, **NOT 19 separate times for 19 questions**.
  - Common helper functions cache and share results across questions:
    - `sys_attachment` is queried **once per CR** and shared across **Q1**, **Q17**, and **Q18**.
    - `sysapproval_approver` is queried **once per CR** and shared across **Q9**, **Q10**, **Q11**, **Q12**, **Q13**, **Q14**, and **Q15**.
    - Questions that do not apply to the change type/risk level return `N/A` immediately without making unnecessary HTTP requests.
- **Final API Endpoints Used:**
  - Change Requests: `.../api/now/table/change_request`
  - Approvals: `.../api/now/table/sysapproval_approver`
  - Attachment Metadata: `.../api/now/table/sys_attachment`
  - Attachment Binary Download: `.../api/now/attachment/{sys_id}/file`
  - Incidents: `.../api/now/table/incident`
  - Problems: `.../api/now/table/problem`
  - User Groups: `.../api/now/table/sys_user_group`
  - Group Memberships: `.../api/now/table/sys_user_grmember`
  - Configuration Items: `.../api/now/table/cmdb_ci/{sys_id}`

---

### **5. What other associated data are you retrieving from different tables (for example, approvals, attachments, activity logs, work notes, comments, etc.)?**

The automation queries **8 associated ServiceNow tables**:

| Table Name | Associated Data Retrieved | Used in Questions |
|---|---|---|
| **`sysapproval_approver`** | Approver names, approval state (`approved`), created date, approval timestamp (`sys_updated_on`). | Q9, Q10, Q11, Q12, Q13, Q14, Q15 |
| **`sys_attachment`** | Attachment metadata (`file_name`, `content_type`, `size_bytes`) and binary payload via Attachment API. | Q1, Q17, Q18, download_attachment() |
| **`incident`** | Linked incidents (`caused_by.number`, `rfc`), and Major Incident flag (`major_incident=true`). | Header metadata, Q5, Q16 |
| **`problem`** | Linked problem tickets (`rfc.number`). | Header metadata, Q6 |
| **`sys_user_group`** | Assignment group sys_id, name, and Senior Leadership Team reference (`u_group_slt`). | Q10, Q11, Q12, Q13, Q14 |
| **`sys_user_grmember`** | Group members of Assignment Groups and "Change Management Approval" (ITCM). | Q10, Q11, Q12, Q14 |
| **`sys_user_delegate`** | Active user delegates during approval windows (`delegate`, `starts`, `ends`). | Approval delegation resolution |
| **`cmdb_ci`** | Configuration Item owner (`owned_by`) and operational status (`operational_status`). | Q11, sc_sheet_data.is_retired() |
| **`std_change_producer_version` & `std_change_proposal`** | Standard change template versions, template definition blobs (`template_value`), and review dates (`u_last_reviewed`). | Q7, Q8, sc_sheet_data |

*(Note: Activity logs, work notes, and comments are **not** stored or queried in separate tables; text fields like `short_description`, `u_non_prod_test_plan_summary`, `u_non_prod_test_plan_results`, and `close_code` come directly from the main `change_request` table).*

---

### **6. What is the API URL used to fetch each type of data from ServiceNow?**

```http
# 1. CHANGE REQUESTS (Date Range Mode)
GET /api/now/table/change_request?sysparm_query=type=<TYPE>^end_dateBETWEENjavascript:gs.dateGenerate('<START>','00:00:00')@javascript:gs.dateGenerate('<END>','23:59:59')^stateIN-2,-1,0,3^production_system=true&sysparm_display_value=true

# 2. APPROVAL RECORDS
GET /api/now/table/sysapproval_approver?sysparm_query=sysapproval.number=<CHG_NUM>^state=approved&sysparm_fields=approver,sys_created_on,sys_updated_on,state&sysparm_display_value=true

# 3. ATTACHMENT METADATA
GET /api/now/table/sys_attachment?sysparm_query=table_sys_id=<CR_SYS_ID>&sysparm_fields=sys_id,file_name,content_type,size_bytes,table_name&sysparm_display_value=true

# 4. ATTACHMENT FILE DOWNLOAD
GET /api/now/attachment/<ATTACHMENT_SYS_ID>/file

# 5. RELATED INCIDENTS
GET /api/now/table/incident?sysparm_query=caused_by.number=<CHG_NUM>&sysparm_fields=number&sysparm_display_value=true
GET /api/now/table/incident?sysparm_query=rfcISNOTEMPTY^rfc=<CR_SYS_ID>&sysparm_fields=number&sysparm_display_value=true

# 6. MAJOR INCIDENTS
GET /api/now/table/incident?sysparm_query=caused_by=<CR_SYS_ID>^major_incident=true&sysparm_fields=number,major_incident&sysparm_display_value=true

# 7. RELATED PROBLEMS
GET /api/now/table/problem?sysparm_query=rfc.number=<CHG_NUM>&sysparm_fields=number&sysparm_display_value=true

# 8. USER GROUPS & SLT
GET /api/now/table/sys_user_group?sysparm_query=name=<GROUP_NAME>&sysparm_fields=sys_id,name,u_group_slt,parent&sysparm_display_value=true

# 9. GROUP MEMBERS
GET /api/now/table/sys_user_grmember?sysparm_query=group=<GROUP_SYS_ID>&sysparm_fields=user&sysparm_display_value=true

# 10. USER DELEGATES
GET /api/now/table/sys_user_delegate?sysparm_query=user.name=<USER_NAME>&sysparm_fields=delegate,starts,ends&sysparm_display_value=true

# 11. CMDB CONFIGURATION ITEM
GET /api/now/table/cmdb_ci/<CI_SYS_ID>?sysparm_display_value=true
```

---

### **7. What is the sequence of steps followed to complete the validation?**

The validation executes through an 8-step pipeline:

```
[CLI Inputs & Date Validation]
             │
             ▼
[Environment & API Auth Setup]
             │
             ▼
[Fetch 100% Population per Change Type]
             │
             ▼
[Local Random Sampling (20%)]
             │
             ▼
[Load 'QAR Report.xlsx' & Write Header (Rows 3-7)]
             │
             ▼
[Audit Each Sampled CR (Common Data + Q1-Q19)]
             │
             ▼
[Populate Columns C, D, E... & Apply Exception Colors]
             │
             ▼
[Save Report to report/FCB <Month>'<YY> - QAR Report.xlsx]
```

1. **CLI Input Validation (`populate_cr_audit.py`):** Validates `--start-date` and `--end-date` (`YYYY-MM-DD`), sample percentage (default: 20%), tester name, and verifies that the Excel template exists.
2. **Environment & Credential Setup (`constants.py`):** Loads `.env` file credentials (`SN_INSTANCE_URL`, `SN_USERNAME`, `SN_PASSWORD`) and configures basic HTTP authentication with SSL settings.
3. **Fetch Total Population (`servicenow_api.py:fetch_crs_from_servicenow`):** Queries ServiceNow for all closed production changes in the date range for each change type (`normal`, `standard`, `emergency`, `expedited`).
4. **Sample Selection (`servicenow_api.py:random_select`):** Calculates sample count and draws a 20% random sample via `random.sample()`. *(Standard templates filter proposals older than 10 months)*.
5. **Template Loading & Header Setup (`excel_handler.py:populate_basic_data`):** Opens `QAR Report.xlsx` and writes Test Month, Tester Name, Total Population count, and Samples Tested count into Column B (rows 3–7) across all result sheets.
6. **Per-CR Audit Evaluation (`questions_data.py`):**
   - Retrieves common CR metadata and linked Incidents/Problems.
   - Evaluates Q1–Q19: scans attachments for test keywords, checks SDLC/RITM numbers, matches reasons for change, validates separation of duties, confirms SLT/ITCM/CI Owner approvals, verifies approval timestamps precede actual start times, and checks that implementation was within the approved change window.
7. **Excel Cell Population & Color Formatting (`excel_handler.py:populate_sheet`):** Writes values into Column C, D, E... (rows 10–38) and applies fill colors:
   - `KWNF` (Keyword Not Found) $\rightarrow$ **Light Yellow** (`#FFF2CC`)
   - `MM` (Maintenance Mismatch) $\rightarrow$ **Light Red** (`#F4CCCC`)
   - `DM` (DR Mismatch) $\rightarrow$ **Light Blue** (`#DDEBF7`)
   - `MA` (Missing Approval) $\rightarrow$ **Light Purple** (`#E4DFEC`)
8. **Save Workbook (`excel_handler.py:save_output`):** Writes the final report to disk under `report/FCB <Month>'<YY> - QAR Report.xlsx`.
