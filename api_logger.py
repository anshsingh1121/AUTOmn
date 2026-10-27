"""Automated, structured API payload and hit logger for ServiceNow requests.

Intercepts all HTTP GET requests made via the requests library, logs request details,
masks sensitive credentials, and saves full, pretty-printed JSON payloads to disk.
"""
from __future__ import annotations
import json
import os
import re
import time
import urllib.parse
from datetime import datetime
from pathlib import Path
import requests

# Root directory for storing all API hit payloads and summaries
LOG_ROOT = Path("api_logs")
_CURRENT_SESSION_DIR = None
_HIT_COUNTER = 0
_ORIGINAL_REQUESTS_GET = requests.get
_INSTALLED = False


def _get_session_dir() -> Path:
    global _CURRENT_SESSION_DIR
    if _CURRENT_SESSION_DIR is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        _CURRENT_SESSION_DIR = LOG_ROOT / f"session_{timestamp}"
        _CURRENT_SESSION_DIR.mkdir(parents=True, exist_ok=True)
        # Initialize the markdown summary dashboard
        summary_file = _CURRENT_SESSION_DIR / "00_API_HITS_SUMMARY.md"
        if not summary_file.exists():
            summary_file.write_text(
                "# ServiceNow API Hits & Payloads Summary\n\n"
                f"**Session Started:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
                "| Hit # | Timestamp | Table / Resource | Status | Duration | Records | Payload File |\n"
                "|:---:|:---:|---|:---:|:---:|:---:|---|\n",
                encoding="utf-8",
            )
    return _CURRENT_SESSION_DIR


def _sanitize_auth(kwargs: dict) -> dict:
    """Mask password from auth tuple or headers."""
    sanitized = {}
    auth = kwargs.get("auth")
    if auth and isinstance(auth, (tuple, list)) and len(auth) >= 2:
        sanitized["auth"] = (auth[0], "********")
    headers = kwargs.get("headers")
    if headers and isinstance(headers, dict):
        clean_headers = dict(headers)
        if "Authorization" in clean_headers:
            clean_headers["Authorization"] = "Basic ********"
        sanitized["headers"] = clean_headers
    return sanitized


def _extract_table_or_endpoint(url: str) -> str:
    """Extract a friendly label for the table or endpoint from the URL."""
    parsed = urllib.parse.urlparse(url)
    path = parsed.path
    if "/api/now/table/" in path:
        return path.split("/api/now/table/")[-1].strip("/")
    if "/api/now/attachment/" in path:
        return "sys_attachment (file)"
    return path.split("/")[-1] or "api_endpoint"


def _log_hit(url: str, args, kwargs, response: requests.Response, elapsed_ms: float):
    global _HIT_COUNTER
    _HIT_COUNTER += 1
    hit_num = _HIT_COUNTER
    session_dir = _get_session_dir()

    table_name = _extract_table_or_endpoint(url)
    safe_table_name = re.sub(r"[^a-zA-Z0-9_\-]", "_", table_name)
    payload_filename = f"hit_{hit_num:03d}_{safe_table_name}.json"
    payload_path = session_dir / payload_filename

    # Extract query params
    params = kwargs.get("params") or {}
    if not params and "?" in url:
        parsed_url = urllib.parse.urlparse(url)
        params = dict(urllib.parse.parse_qsl(parsed_url.query))

    # Parse JSON payload if applicable
    is_json = "application/json" in response.headers.get("Content-Type", "").casefold()
    records_count = 0
    payload_content = None

    if is_json:
        try:
            payload_content = response.json()
            if isinstance(payload_content, dict):
                result = payload_content.get("result")
                if isinstance(result, list):
                    records_count = len(result)
                elif result is not None:
                    records_count = 1
        except Exception:
            payload_content = response.text[:2000]
    else:
        # Binary or text stream (e.g. attachment file download)
        content_size = len(response.content) if response.content else 0
        records_count = 1 if content_size > 0 else 0
        payload_content = f"[Binary attachment content: {content_size} bytes, Content-Type: {response.headers.get('Content-Type')}]"

    # Assemble structured hit log
    log_record = {
        "hit_number": hit_num,
        "timestamp": datetime.now().isoformat(),
        "method": "GET",
        "url": url,
        "decoded_url": urllib.parse.unquote(url),
        "target_resource": table_name,
        "query_parameters": params,
        "request_options": _sanitize_auth(kwargs),
        "response": {
            "status_code": response.status_code,
            "reason": response.reason,
            "elapsed_ms": round(elapsed_ms, 2),
            "records_returned": records_count,
            "content_type": response.headers.get("Content-Type", ""),
            "payload": payload_content,
        },
    }

    # Write formatted, pretty-printed JSON file
    try:
        with open(payload_path, "w", encoding="utf-8") as f:
            json.dump(log_record, f, indent=2, ensure_ascii=False)
    except Exception as exc:
        print(f"[api_logger] Warning: Failed to save payload for hit {hit_num}: {exc}")

    # Append row to markdown summary table
    summary_file = session_dir / "00_API_HITS_SUMMARY.md"
    status_str = f"{response.status_code} {response.reason}"
    duration_str = f"{round(elapsed_ms, 1)}ms"
    row = (
        f"| {hit_num:03d} | {datetime.now().strftime('%H:%M:%S')} | "
        f"`{table_name}` | {status_str} | {duration_str} | "
        f"{records_count} | [`{payload_filename}`]({payload_filename}) |\n"
    )
    try:
        with open(summary_file, "a", encoding="utf-8") as f:
            f.write(row)
    except Exception:
        pass


def logged_get(url, *args, **kwargs):
    """Wrapper intercepting requests.get to record payloads."""
    start_time = time.perf_counter()
    response = _ORIGINAL_REQUESTS_GET(url, *args, **kwargs)
    elapsed_ms = (time.perf_counter() - start_time) * 1000.0

    try:
        _log_hit(url, args, kwargs, response, elapsed_ms)
    except Exception as exc:
        print(f"[api_logger] Warning: Failed to log API hit: {exc}")

    return response


def install_api_logging():
    """Install the requests.get interceptor."""
    global _INSTALLED
    if not _INSTALLED:
        requests.get = logged_get
        _INSTALLED = True


def get_current_session_dir() -> Path:
    return _get_session_dir()


def get_summary_path() -> Path:
    return _get_session_dir() / "00_API_HITS_SUMMARY.md"


# Auto-install upon import if enabled via environment (default: enabled)
if os.getenv("SN_LOG_PAYLOADS", "true").casefold() in {"1", "true", "yes"}:
    install_api_logging()
