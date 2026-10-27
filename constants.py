"""Shared configuration and business-rule constants.

Credentials are intentionally read from environment variables instead of being
embedded in source. The screenshots showed a ServiceNow service account and a
password; those values are not copied into this reconstructed package.
"""
from __future__ import annotations
import os
import re

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import api_logger
except ImportError:
    api_logger = None

INSTANCE_URL = os.getenv("SN_INSTANCE_URL", "https://your-instance.service-now.com").rstrip("/")
BASE_URL = f"{INSTANCE_URL}/api/now/table/"
ATT_FILE_URL = f"{INSTANCE_URL}/api/now/attachment/"
USERNAME = os.getenv("SN_USERNAME", "")
PASSWORD = os.getenv("SN_PASSWORD", "")
VERIFY_SSL = os.getenv("SN_VERIFY_SSL", "false").casefold() in {"1", "true", "yes"}

headers = {"Content-Type": "application/json", "Accept": "application/json"}

YES = "Yes"
NO = "No"
NA = "N/A"
ERROR = "Error"
NON_TESTABLE = "not testable"
KWNF = "Keyword not found"       # Q17
MM = "maintenance mismatch"      # Q4
DM = "dr mismatch"                # Q4
MA = "Missing Approval"           # approval questions

NORMAL = "normal"
STANDARD = "Standard"
EMERGENCY = "Emergency"
EXPEDITED = "Expedited"
UNKNOWN = "Unknown"
DATE_FORMAT = "%m/%d/%Y %I:%M:%S %p"
SCT = "Standard Change Template"

# The source screenshot uses CHG followed by one or more digits, not a fixed
# number of digits. Validate any organization-specific length constraint separately.
CHANGE_NUMBER_PATTERN = re.compile(r"CHG\d+", re.IGNORECASE)

# ServiceNow table names visible in the supplied constants.py screenshots.
CHANGE_REQUEST_TABLE = "change_request"
SYS_APPROVAL_TABLE = "sysapproval_approver"
CHANGE_ATTACH_TABLE = "sys_attachment"
SYS_USER_GROUP_TABLE = "sys_user_group"
SYS_USER_GROUP_MEM_TABLE = "sys_user_grmember"
SYS_USER_DELEGATE_TABLE = "sys_user_delegate"
CMDB_CI = "cmdb_ci"
CMDB_CI_SERV = "cmdb_ci_server"
INCIDENT_TABLE = "incident"
PROBLEM_TABLE = "problem"
CHANGE_TASK = "change_task"

CHANGE_RESULTS = {
    "NC Result Sheet": "normal",
    "SC Result Sheet": "standard",
    "EC Result Sheet": "emergency",
    "Exp Result Sheet": "expedited",
    "SC Template Result Sheet": "standard",
}

# Terms visible in the source screenshots; comparisons should be case-insensitive.
STD_TEST_FAIL_KEYWORDS = ["n/a", "not required", "not needed", "to prod", "not exist"]
# Compatibility alias: some partially recovered modules referenced this misspelling.
FATL_KEYWORDS = STD_TEST_FAIL_KEYWORDS
FAIL_KEYWORDS = ["n/a", "not required", "not needed", "does not exist"]
PASS_KEYWORDS = ["applies to production", "deployment plan", "back out", "verification"]
NON_PROD_KEYWORDS = ["test", "non prod evidence", "non prod artifact"]
PRE_PROD_KEYWORDS = [
    "non prod", "non prod evidence", "test results", "test result evidence",
    "test results evidence", "test cycle report", "test report",
]
POST_PROD_KEYWORDS = [
    "post prod", "prod deployment", "post deployment", "proddeployment",
    "post implementation", "prod validation", "prodvalidation",
]
DEPLOYMENT_KEYWORDS = PASS_KEYWORDS
ISSUE_DOCUMENTATION_KEYWORDS = ["issue", "incident", "problem", "defect", "implementation review"]
STANDARD_CH_APPROVAL_EXCEPTION_TEMPLATES = [
    "Cyber Threat Proxy Change", "Cyber Threat Black Hole Route",
    "XSOAR Content Change", "Linux AIX Server Decommission",
]
