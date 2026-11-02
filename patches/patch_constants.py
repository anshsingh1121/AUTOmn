# =====================================================================
# PATCH FOR: constants.py
# Location: Replace the update_credentials function (lines 113-126)
# =====================================================================

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
