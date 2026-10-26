# ServiceNow API Calling Architecture & Complete Hit URLs

This document provides a concrete, comprehensive breakdown of every API call made during the Change Management Audit Automation (CMAA) process, including the **exact full HTTP Hit URLs**, query parameters, request counts, and an explanation of why nearly half of the 19 questions make **ZERO API calls**.

---

## 1. Executive Summary: Total API Call Formula

$$\text{Total API Hits} = \mathbf{N_{\text{batch}}} + \left(\mathbf{N_{\text{sampled\_CRs}}} \times \mathbf{\sim 7}\right)$$

1. **Initial Population Fetch ($\mathbf{N_{\text{batch}}}$):**
   - **Date-Range Mode:** Exactly **4 batch calls** (1 call per change type: Normal, Standard, Emergency, Expedited).
   - **Change-Numbers Mode:** Exactly **1 batch call** (`numberIN...`) for all change numbers.
2. **Per-Sample Validation:** Approximately **6 to 8 API hit calls** per sampled Change Request.
3. **Per-Question Execution:** The API is **NOT** called 19 times for 19 questions. Out of 19 questions, **8 questions make 0 API calls** because they evaluate in-memory fields or run local Python logic.

---

## 2. Why Nearly Half the Questions Make ZERO API Calls

When the initial batch call runs, ServiceNow returns the **complete dictionary of Change Request fields** for every ticket into Python memory. 

The following questions evaluate data already present in that memory payload, resulting in **0 external network requests**:

| Question | Evaluation Logic | Why 0 API Hits? |
|---|---|---|
| **Q2 (SDLC/RITM for Untestable)** | Checks `u_change_cannot_be_pre_tested` flag and scans text inside `u_non_prod_test_plan_summary` for `"sii"` or `"ritm"`. | Both fields already exist in the CR payload. |
| **Q3 (Reason for Change Active)** | Checks if `u_reason_for_change` has "project"/"patching" and verifies `short_description` is filled. | Both fields already exist in the CR payload. |
| **Q4 (Reason vs. SDLC Alignment)** | Compares `u_reason_for_change` against `u_sdlc_initiative_type`. | Both dropdown strings already exist in memory. |
| **Q5 (Incident Resolution Check)** | Checks if `u_related_incidents` text field is populated. | Field is part of the CR payload. |
| **Q6 (Problem Resolution Check)** | Checks if `u_related_problems` text field is populated. | Field is part of the CR payload. |
| **Q8 (Standard Differentiation)** | Calculates fuzzy string similarity (`fuzz.token_sort_ratio`) between `short_description` and template name. | Evaluated 100% locally in Python memory. |
| **Q18 (Issue Doc for Successful CRs)**| If `close_code == "successful"`, returns `N/A` immediately. | Only makes an API call if the CR failed/backed out. |
| **Q19 (Change Window Check)** | Compares `actual_start` and `actual_end` against `start_date` and `end_date`. | All 4 timestamps are already in the CR payload. |

---

## 3. The Complete Hit URLs Catalog

All requests are HTTP `GET` calls with the following headers:
```http
Accept: application/json
Content-Type: application/json
Authorization: Basic <base64_encoded_credentials>
```

---

### Hit URL 1: Initial Population Batch Fetch (Date-Range Mode)
- **When Triggered:** Once per change type at the start of the audit run (4 calls total).
- **ServiceNow Table:** `change_request`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/change_request?sysparm_query=type=<CHANGE_TYPE>^end_dateBETWEENjavascript:gs.dateGenerate('<START_DATE>','00:00:00')@javascript:gs.dateGenerate('<END_DATE>','23:59:59')^stateIN-2,-1,0,3^production_system=true&sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/change_request?sysparm_query=type=normal^end_dateBETWEENjavascript:gs.dateGenerate('2026-10-01','00:00:00')@javascript:gs.dateGenerate('2026-10-05','23:59:59')^stateIN-2,-1,0,3^production_system=true&sysparm_display_value=true
```

---

### Hit URL 2: Initial Population Fetch (Change-Numbers Mode)
- **When Triggered:** Exactly once if running with `--change-nums` (1 call total).
- **ServiceNow Table:** `change_request`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/change_request?sysparm_query=numberIN<CHG1>,<CHG2>,<CHG3>...&sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/change_request?sysparm_query=numberINCHG6892954,CHG6957477,CHG6738045&sysparm_display_value=true
```

---

### Hit URL 3: Related Incidents Caused by Change
- **When Triggered:** For each sampled CR during `get_common_data()`.
- **ServiceNow Table:** `incident`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/incident?sysparm_query=caused_by.number=<CHANGE_NUMBER>&sysparm_fields=number&sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/incident?sysparm_query=caused_by.number=CHG6892954&sysparm_fields=number&sysparm_display_value=true
```

---

### Hit URL 4: Related Incidents Linked by RFC
- **When Triggered:** For each sampled CR during `get_common_data()`.
- **ServiceNow Table:** `incident`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/incident?sysparm_query=rfcISNOTEMPTY^rfc=<CR_SYS_ID>&sysparm_fields=number&sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/incident?sysparm_query=rfcISNOTEMPTY^rfc=a9e8d7c6b5a41234567890abcdef1234&sysparm_fields=number&sysparm_display_value=true
```

---

### Hit URL 5: Related Problems Linked by RFC
- **When Triggered:** For each sampled CR during `get_common_data()`.
- **ServiceNow Table:** `problem`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/problem?sysparm_query=rfc.number=<CHANGE_NUMBER>&sysparm_fields=number&sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/problem?sysparm_query=rfc.number=CHG6892954&sysparm_fields=number&sysparm_display_value=true
```

---

### Hit URL 6: Attachment Metadata (Q1, Q17, Q18)
- **When Triggered:** For each sampled CR (shared by Q1 pre-test, Q17 post-validation, and Q18 issue checks).
- **ServiceNow Table:** `sys_attachment`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/sys_attachment?sysparm_query=table_sys_id=<CR_SYS_ID>&sysparm_fields=sys_id,file_name,content_type,size_bytes,table_name&sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/sys_attachment?sysparm_query=table_sys_id=a9e8d7c6b5a41234567890abcdef1234&sysparm_fields=sys_id,file_name,content_type,size_bytes,table_name&sysparm_display_value=true
```

---

### Hit URL 7: Attachment Binary Download
- **When Triggered:** On-demand when inspecting inside file contents (`download_attachment()`).
- **ServiceNow Endpoint:** Attachment API
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/attachment/<ATTACHMENT_SYS_ID>/file
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/attachment/b1c2d3e4f567890123456789abcdef01/file
```

---

### Hit URL 8: Approved Records & Timestamps (Q9, Q10–Q15)
- **When Triggered:** For each sampled CR to fetch actual approvers and approval timestamps.
- **ServiceNow Table:** `sysapproval_approver`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/sysapproval_approver?sysparm_query=sysapproval.number=<CHANGE_NUMBER>^state=approved&sysparm_fields=approver,sys_created_on,sys_updated_on,state&sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/sysapproval_approver?sysparm_query=sysapproval.number=CHG6892954^state=approved&sysparm_fields=approver,sys_created_on,sys_updated_on,state&sysparm_display_value=true
```

---

### Hit URL 9: Assignment Group & SLT Lookup (Q10, Q11, Q12, Q13, Q14)
- **When Triggered:** To retrieve group metadata and Senior Leadership Team (`u_group_slt`).
- **ServiceNow Table:** `sys_user_group`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/sys_user_group?sysparm_query=name=<GROUP_NAME>&sysparm_fields=sys_id,name,u_group_slt,parent&sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/sys_user_group?sysparm_query=name=Database%20Operations&sysparm_fields=sys_id,name,u_group_slt,parent&sysparm_display_value=true
```

---

### Hit URL 10: Group Membership Lookup (Q10, Q11, Q12, Q14)
- **When Triggered:** To retrieve all members of an Assignment Group or ITCM group.
- **ServiceNow Table:** `sys_user_grmember`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/sys_user_grmember?sysparm_query=group=<GROUP_SYS_ID>&sysparm_fields=user&sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/sys_user_grmember?sysparm_query=group=d4e5f6a7b8901234567890abcdef1234&sysparm_fields=user&sysparm_display_value=true
```

---

### Hit URL 11: User Delegation Lookup
- **When Triggered:** If checking whether an approval was performed by an authorized delegate.
- **ServiceNow Table:** `sys_user_delegate`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/sys_user_delegate?sysparm_query=user.name=<USER_NAME>&sysparm_fields=delegate,starts,ends&sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/sys_user_delegate?sysparm_query=user.name=John%20Doe&sysparm_fields=delegate,starts,ends&sysparm_display_value=true
```

---

### Hit URL 12: Major Incident Outage Check (Q16)
- **When Triggered:** For each sampled CR to verify if an outage was flagged as a Major Incident.
- **ServiceNow Table:** `incident`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/incident?sysparm_query=caused_by=<CR_SYS_ID>^major_incident=true&sysparm_fields=number,major_incident&sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/incident?sysparm_query=caused_by=a9e8d7c6b5a41234567890abcdef1234^major_incident=true&sysparm_fields=number,major_incident&sysparm_display_value=true
```

---

### Hit URL 13: Configuration Item (CI) Details & Operational Status (Q11 & SC Sheet)
- **When Triggered:** To fetch CI Owner (`owned_by`) or check if CI is retired (`operational_status`).
- **ServiceNow Table:** `cmdb_ci`
- **Full Hit URL:**
```http
GET https://<instance>.service-now.com/api/now/table/cmdb_ci/<CI_SYS_ID>?sysparm_display_value=true
```
- **Example Real Hit:**
```text
https://company.service-now.com/api/now/table/cmdb_ci/e5f6a7b8c901234567890abcdef12345?sysparm_display_value=true
```

---

### Hit URL 14: Standard Change Proposal & Template Links (Q7 & SC Sheet)
- **When Triggered:** Standard Changes only, following the template version and proposal links.
- **ServiceNow Tables:** `std_change_producer_version` & `std_change_proposal`
- **Full Hit URLs:**
```http
# Step 1: Follow template version link
GET https://<instance>.service-now.com/api/now/table/std_change_producer_version/<VERSION_SYS_ID>?sysparm_display_value=true

# Step 2: Follow proposal link
GET https://<instance>.service-now.com/api/now/table/std_change_proposal/<PROPOSAL_SYS_ID>?sysparm_display_value=true
```

---

## 4. Complete Audit Call Sequence Summary

```
[Start Audit]
   │
   ├─► HIT 1: GET /change_request (Batch 100% Population per change type)
   │
   └─► [Python Memory: 20% Random Sample Selection]
         │
         ▼  (For Each Sampled Change Request):
         ├─► HIT 3: GET /incident (caused_by.number = CHG)
         ├─► HIT 4: GET /incident (rfc = sys_id)
         ├─► HIT 5: GET /problem (rfc.number = CHG)
         ├─► HIT 6: GET /sys_attachment (table_sys_id = sys_id)
         ├─► HIT 8: GET /sysapproval_approver (sysapproval.number = CHG)
         ├─► HIT 9: GET /sys_user_group (name = Assignment Group)
         ├─► HIT 10: GET /sys_user_grmember (group = sys_id)
         ├─► HIT 12: GET /incident (caused_by = sys_id ^ major_incident = true)
         └─► (Q2, Q3, Q4, Q5, Q6, Q8, Q18, Q19: ZERO API HITS - evaluated in Python memory)
```
