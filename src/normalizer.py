"""
Data Normalizer
Priority normalization, whitespace trimming, and data cleaning utilities.
"""

import re
from typing import Any, Dict, List, Optional, Union


def normalize_whitespace(value: Any) -> Optional[str]:
    """
    Normalize whitespace in a string value.
    Returns None for blank/None values.
    """
    if value is None:
        return None
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return None
    return s


def extract_priority_class(priority_value: Any, config: Optional[Dict] = None) -> Optional[int]:
    """
    Extract the numeric priority class from a priority value.

    Handles formats like:
        "3 - Moderate" -> 3
        "1 - Critical" -> 1
        "3"            -> 3
        3              -> 3

    Args:
        priority_value: Raw priority value from source data.
        config: Optional priority normalization config with 'separator' and 'patterns'.

    Returns:
        Integer priority class, or None if unparseable.
    """
    if priority_value is None:
        return None

    val = str(priority_value).strip()
    if val == "" or val.lower() == "nan":
        return None

    # Try direct integer parse
    try:
        return int(val)
    except (ValueError, TypeError):
        pass

    # Try extracting leading number (e.g., "3 - Moderate")
    separator = " - "
    if config and "separator" in config:
        separator = config["separator"]

    if separator in val:
        parts = val.split(separator, 1)
        try:
            return int(parts[0].strip())
        except (ValueError, TypeError):
            pass

    # Try regex: find first sequence of digits
    match = re.match(r"(\d+)", val)
    if match:
        return int(match.group(1))

    return None


def is_priority_3(priority_value: Any, config: Optional[Dict] = None) -> bool:
    """
    Check if a priority value corresponds to Priority 3.

    Args:
        priority_value: Raw priority value.
        config: Optional priority normalization config.

    Returns:
        True if priority class is 3.
    """
    priority_class = extract_priority_class(priority_value, config)
    target_class = 3
    if config and "priority_3_class" in config:
        target_class = config["priority_3_class"]
    return priority_class == target_class


def normalize_for_matching(value: Any, case_sensitive: bool = False) -> Optional[str]:
    """
    Normalize a value for matching/comparison purposes.
    Strips whitespace, optionally lowercases.

    Returns None for blank values.
    """
    normalized = normalize_whitespace(value)
    if normalized is None:
        return None
    if not case_sensitive:
        return normalized.lower()
    return normalized


def is_blank(value: Any) -> bool:
    """Check if a value is None, empty, or whitespace-only."""
    if value is None:
        return True
    s = str(value).strip()
    return s == "" or s.lower() == "nan"


def contains_substring(value: Any, substring: str, case_sensitive: bool = False) -> bool:
    """
    Check if a value contains a substring.

    Args:
        value: The string to search within.
        substring: The substring to search for.
        case_sensitive: Whether matching is case-sensitive.

    Returns:
        True if the substring is found.
    """
    if is_blank(value) or is_blank(substring):
        return False
    val = str(value).strip()
    sub = str(substring).strip()
    if not case_sensitive:
        val = val.lower()
        sub = sub.lower()
    return sub in val
