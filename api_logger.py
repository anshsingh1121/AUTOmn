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


def _sanitize_url(url: str) -> str:
    """Mask password from URL if embedded in userinfo (e.g. http://user:pass@host)."""
    try:
        parsed = urllib.parse.urlsplit(url)
        if parsed.password:
            netloc = f"{parsed.username}:********@{parsed.hostname}"
            if parsed.port:
                netloc += f":{parsed.port}"
            return urllib.parse.urlunsplit((parsed.scheme, netloc, parsed.path, parsed.query, parsed.fragment))
    except Exception:
        pass
    return url


def _sanitize_auth(kwargs: dict) -> dict:
    """Mask password from auth tuple or headers."""
    sanitized = {}
    auth = kwargs.get("auth")
    if auth and isinstance(auth, (tuple, list)) and len(auth) >= 2:
        sanitized["auth"] = (str(auth[0]), "********")
    headers = kwargs.get("headers")
    if headers and isinstance(headers, dict):
        clean_headers = dict(headers)
        for h_key in list(clean_headers.keys()):
            if h_key.casefold() in {"authorization", "proxy-authorization"}:
                clean_headers[h_key] = "Basic ********"
        sanitized["headers"] = clean_headers
    return sanitized


def _mask_known_secrets(text: str) -> str:
    """Scrub known sensitive values such as constants.PASSWORD from text/JSON output."""
    if not text:
        return text
    try:
        import constants
        pwd = getattr(constants, "PASSWORD", None)
        if pwd and isinstance(pwd, str) and len(pwd) >= 3 and pwd in text:
            text = text.replace(pwd, "********")
    except Exception:
        pass
    return text


def _extract_table_or_endpoint(url: str) -> str:
    """Extract a friendly label for the table or endpoint from the URL."""
    parsed = urllib.parse.urlparse(url)
    path = parsed.path
    if "/api/now/table/" in path:
        return path.split("/api/now/table/")[-1].strip("/")
    if "/api/now/attachment/" in path:
        return "sys_attachment (file)"
    return path.split("/")[-1] or "api_endpoint"


def _extract_params(url: str, args, kwargs) -> dict:
    """Extract query parameters safely from kwargs, args, or the URL itself."""
    params = kwargs.get("params")
    if not params and args and len(args) > 0 and isinstance(args[0], dict):
        params = args[0]
    if not params and "?" in url:
        parsed_url = urllib.parse.urlparse(url)
        params = dict(urllib.parse.parse_qsl(parsed_url.query))
    return params or {}


def _log_hit(url: str, args, kwargs, response: requests.Response, elapsed_ms: float):
    global _HIT_COUNTER
    _HIT_COUNTER += 1
    hit_num = _HIT_COUNTER
    session_dir = _get_session_dir()

    table_name = _extract_table_or_endpoint(url)
    safe_table_name = re.sub(r"[^a-zA-Z0-9_\-]", "_", table_name)
    payload_filename = f"hit_{hit_num:03d}_{safe_table_name}.json"
    payload_path = session_dir / payload_filename

    params = _extract_params(url, args, kwargs)

    # Determine if response is JSON or binary/text
    content_type = response.headers.get("Content-Type", "") if hasattr(response, "headers") else ""
    is_json = "application/json" in content_type.casefold()
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
            payload_content = response.text[:2000] if hasattr(response, "text") else ""
    else:
        content_size = len(response.content) if hasattr(response, "content") and response.content else 0
        records_count = 1 if content_size > 0 else 0
        payload_content = f"[Binary/Stream attachment content: {content_size} bytes, Content-Type: {content_type}]"

    safe_url = _sanitize_url(url)
    log_record = {
        "hit_number": hit_num,
        "timestamp": datetime.now().isoformat(),
        "method": "GET",
        "url": safe_url,
        "decoded_url": _sanitize_url(urllib.parse.unquote(url)),
        "target_resource": table_name,
        "query_parameters": params,
        "request_options": _sanitize_auth(kwargs),
        "response": {
            "status_code": getattr(response, "status_code", None),
            "reason": getattr(response, "reason", ""),
            "elapsed_ms": round(elapsed_ms, 2),
            "records_returned": records_count,
            "content_type": content_type,
            "payload": payload_content,
        },
    }

    try:
        raw_json = json.dumps(log_record, indent=2, ensure_ascii=False, default=str)
        scrubbed_json = _mask_known_secrets(raw_json)
        with open(payload_path, "w", encoding="utf-8") as f:
            f.write(scrubbed_json)
    except Exception as exc:
        print(f"[api_logger] Warning: Failed to save payload for hit {hit_num}: {exc}")

    summary_file = session_dir / "00_API_HITS_SUMMARY.md"
    status_str = f"{getattr(response, 'status_code', 'ERR')} {getattr(response, 'reason', '')}".strip()
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


def _log_error_hit(url: str, args, kwargs, error: Exception, elapsed_ms: float):
    """Log an API hit that raised an exception (e.g. timeout, connection error)."""
    global _HIT_COUNTER
    _HIT_COUNTER += 1
    hit_num = _HIT_COUNTER
    session_dir = _get_session_dir()

    table_name = _extract_table_or_endpoint(url)
    safe_table_name = re.sub(r"[^a-zA-Z0-9_\-]", "_", table_name)
    payload_filename = f"hit_{hit_num:03d}_{safe_table_name}_ERROR.json"
    payload_path = session_dir / payload_filename

    params = _extract_params(url, args, kwargs)
    safe_url = _sanitize_url(url)

    log_record = {
        "hit_number": hit_num,
        "timestamp": datetime.now().isoformat(),
        "method": "GET",
        "url": safe_url,
        "decoded_url": _sanitize_url(urllib.parse.unquote(url)),
        "target_resource": table_name,
        "query_parameters": params,
        "request_options": _sanitize_auth(kwargs),
        "error": {
            "type": type(error).__name__,
            "message": str(error),
            "elapsed_ms": round(elapsed_ms, 2),
        },
    }

    try:
        raw_json = json.dumps(log_record, indent=2, ensure_ascii=False, default=str)
        scrubbed_json = _mask_known_secrets(raw_json)
        with open(payload_path, "w", encoding="utf-8") as f:
            f.write(scrubbed_json)
    except Exception:
        pass

    summary_file = session_dir / "00_API_HITS_SUMMARY.md"
    status_str = f"FAILED ({type(error).__name__})"
    duration_str = f"{round(elapsed_ms, 1)}ms"
    row = (
        f"| {hit_num:03d} | {datetime.now().strftime('%H:%M:%S')} | "
        f"`{table_name}` | **{status_str}** | {duration_str} | "
        f"0 | [`{payload_filename}`]({payload_filename}) |\n"
    )
    try:
        with open(summary_file, "a", encoding="utf-8") as f:
            f.write(row)
    except Exception:
        pass


def logged_get(url, *args, **kwargs):
    """Wrapper intercepting requests.get to record payloads and errors."""
    start_time = time.perf_counter()
    try:
        response = _ORIGINAL_REQUESTS_GET(url, *args, **kwargs)
        elapsed_ms = (time.perf_counter() - start_time) * 1000.0
        try:
            _log_hit(url, args, kwargs, response, elapsed_ms)
        except Exception as log_exc:
            print(f"[api_logger] Warning: Failed to log API hit: {log_exc}")
        return response
    except Exception as req_error:
        elapsed_ms = (time.perf_counter() - start_time) * 1000.0
        try:
            _log_error_hit(url, args, kwargs, req_error, elapsed_ms)
        except Exception:
            pass
        raise req_error


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


# Automatically install logger upon import
install_api_logging()
