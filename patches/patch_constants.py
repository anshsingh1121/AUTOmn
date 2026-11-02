# =====================================================================
# PATCH FOR: constants.py
# 
# 1. Input Constraints: Add near CHANGE_NUMBER_PATTERN (around line 65):
# ---------------------------------------------------------------------
CHANGE_NUMBER_PATTERN = re.compile(r"CHG\d+", re.IGNORECASE)

# Corporate Audit Input Constraints
MAX_CHANGE_NUMBERS = 100      # Maximum allowed Change Requests per audit run
MAX_DATE_RANGE_DAYS = 31      # Maximum allowed date range span (1 calendar month)


# 2. Credential Updates: Replace update_credentials function (lines 115-130):
# ---------------------------------------------------------------------
def update_credentials(username: str, password: str, instance_url: str | None = None, verify_ssl: bool | str | None = None):
    """Dynamically update runtime credentials in memory."""
    global USERNAME, PASSWORD, INSTANCE_URL, BASE_URL, ATT_FILE_URL, VERIFY_SSL
    if username is not None and username != "":
        USERNAME = username
    if password is not None and password != "":
        PASSWORD = password
    if instance_url:
        INSTANCE_URL = instance_url.rstrip("/")
        BASE_URL = f"{INSTANCE_URL}/api/now/table/"
        ATT_FILE_URL = f"{INSTANCE_URL}/api/now/attachment/"
    if verify_ssl is not None:
        VERIFY_SSL = verify_ssl


def has_credentials() -> bool:
    """Return True if both username and password are populated in memory."""
    return bool(USERNAME and PASSWORD)
