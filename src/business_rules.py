"""
Business Rules Engine
Centralized, deterministic business rules for the FCB Incident Tracker.

Rules implemented:
    1. IC Determination (3-tier decision tree)
    2. Region Lookup
    3. Bank/SVB Determination
    4. Priority Normalization
    5. Configurable Field Mapping with non-blank overwrite protection
"""

from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from .normalizer import (
    contains_substring,
    extract_priority_class,
    is_blank,
    is_priority_3,
    normalize_for_matching,
    normalize_whitespace,
)


# ============================================================
# IC DETERMINATION RULE
# ============================================================
# Decision tree:
#
#              IC populated?
#                /      \
#              YES       NO
#               |         |
#          use IC     Priority = 3?
#                      /       \
#                    YES        NO
#                     |          |
#              Proposed By      blank
# ============================================================

def determine_ic(
    servicenow_ic: Any,
    priority: Any,
    proposed_by: Any,
    priority_config: Optional[Dict] = None
) -> Optional[str]:
    """
    Apply the IC determination business rule.

    Args:
        servicenow_ic: IC value from ServiceNow.
        priority: Priority value from ServiceNow.
        proposed_by: Proposed By value from ServiceNow.
        priority_config: Optional priority normalization config.

    Returns:
        Determined IC value (string or None).
    """
    # Rule 1: If ServiceNow IC contains a value, use it
    ic_val = normalize_whitespace(servicenow_ic)
    if ic_val is not None:
        return ic_val

    # Rule 2: If IC is blank AND Priority = 3, use Proposed By
    if is_priority_3(priority, priority_config):
        pb_val = normalize_whitespace(proposed_by)
        return pb_val  # May be None if Proposed By is also blank

    # Rule 3: If IC is blank AND Priority != 3, IC is blank
    return None


# ============================================================
# REGION LOOKUP RULE
# ============================================================

def build_ic_lookup_map(
    lookup_df: pd.DataFrame,
    ic_column: str = "IC",
    region_column: str = "Region",
    case_sensitive: bool = False
) -> Tuple[Dict[str, str], List[str]]:
    """
    Build a normalized IC-to-Region lookup dictionary.

    Args:
        lookup_df: IC Lookup DataFrame.
        ic_column: Column name for IC keys.
        region_column: Column name for Region values.
        case_sensitive: Whether matching is case-sensitive.

    Returns:
        Tuple of (lookup_dict, warnings_list).
    """
    lookup = {}
    warnings = []

    for idx, row in lookup_df.iterrows():
        ic_raw = row.get(ic_column)
        region_raw = row.get(region_column)

        ic_normalized = normalize_for_matching(ic_raw, case_sensitive)
        if ic_normalized is None:
            warnings.append(f"IC Lookup row {idx}: blank IC key — skipped.")
            continue

        region_val = normalize_whitespace(region_raw)

        if ic_normalized in lookup:
            warnings.append(
                f"IC Lookup: duplicate key '{ic_raw}' "
                f"(normalized: '{ic_normalized}'). "
                f"Previous region: '{lookup[ic_normalized]}', "
                f"new region: '{region_val}'. Keeping first."
            )
            continue

        lookup[ic_normalized] = region_val

    return lookup, warnings


def determine_region(
    ic_value: Any,
    ic_lookup: Dict[str, str],
    case_sensitive: bool = False
) -> Tuple[Optional[str], bool]:
    """
    Determine Region from IC value using the IC Lookup.

    Args:
        ic_value: The IC value to look up.
        ic_lookup: Normalized IC-to-Region dictionary.
        case_sensitive: Whether matching is case-sensitive.

    Returns:
        Tuple of (region_value, was_matched).
    """
    normalized_ic = normalize_for_matching(ic_value, case_sensitive)
    if normalized_ic is None:
        return None, False

    if normalized_ic in ic_lookup:
        return ic_lookup[normalized_ic], True

    return None, False


# ============================================================
# BANK DETERMINATION RULE
# ============================================================
# IF Assignment Group contains "SVB" (case-insensitive):
#     Bank = "SVB"
# ELSE:
#     Bank = blank
# ============================================================

def determine_bank(
    assignment_group: Any,
    svb_keyword: str = "SVB",
    svb_value: str = "SVB",
    default_value: str = "",
    case_sensitive: bool = False
) -> str:
    """
    Determine Bank value based on Assignment Group.

    Args:
        assignment_group: Assignment Group value.
        svb_keyword: Keyword to search for (default "SVB").
        svb_value: Value to assign if keyword found (default "SVB").
        default_value: Value if keyword not found (default "").
        case_sensitive: Whether matching is case-sensitive.

    Returns:
        Bank value string.
    """
    if contains_substring(assignment_group, svb_keyword, case_sensitive):
        return svb_value
    return default_value


# ============================================================
# FIELD MAPPING ENGINE
# ============================================================

def apply_field_mapping(
    master_record: Dict[str, Any],
    servicenow_record: Dict[str, Any],
    field_mapping: Dict[str, Dict[str, Any]]
) -> Tuple[Dict[str, Any], List[str]]:
    """
    Apply configurable field mapping from ServiceNow to Master record.

    Respects the non-blank overwrite rule:
    - Do NOT overwrite fields merely because a corresponding ServiceNow field is blank.
    - Only overwrite if the source field has a non-blank value OR overwrite_if_blank is True.

    Args:
        master_record: Existing Master Tracker record (dict).
        servicenow_record: ServiceNow record (dict).
        field_mapping: Field mapping configuration.

    Returns:
        Tuple of (updated_record, list_of_changes).
        Each change is a string like "Field: old_value -> new_value".
    """
    updated = dict(master_record)
    changes = []

    for mapping_name, mapping_config in field_mapping.items():
        source_col = mapping_config.get("source_column")
        target_col = mapping_config.get("target_column")
        overwrite_if_blank = mapping_config.get("overwrite_if_blank", False)
        is_key = mapping_config.get("is_key", False)

        if source_col is None or target_col is None:
            continue

        # Skip the key column (Number) — it doesn't change
        if is_key:
            continue

        source_val = servicenow_record.get(source_col)
        old_val = updated.get(target_col)

        # Non-blank overwrite rule
        if is_blank(source_val) and not overwrite_if_blank:
            continue

        new_val = normalize_whitespace(source_val)

        # Only record a change if the value actually changed
        old_normalized = normalize_whitespace(old_val)
        if old_normalized != new_val:
            changes.append(f"{target_col}: '{old_normalized}' -> '{new_val}'")
            updated[target_col] = new_val if new_val is not None else ""

    return updated, changes


def create_new_record(
    servicenow_record: Dict[str, Any],
    field_mapping: Dict[str, Dict[str, Any]],
    master_columns: List[str]
) -> Dict[str, Any]:
    """
    Create a new Master Tracker record from a ServiceNow record.

    Args:
        servicenow_record: The ServiceNow record.
        field_mapping: Field mapping configuration.
        master_columns: List of all Master Tracker column names.

    Returns:
        Dict representing the new Master Tracker record.
    """
    record = {col: "" for col in master_columns}

    for mapping_name, mapping_config in field_mapping.items():
        source_col = mapping_config.get("source_column")
        target_col = mapping_config.get("target_column")

        if source_col is None or target_col is None:
            continue

        source_val = servicenow_record.get(source_col)
        val = normalize_whitespace(source_val)
        if val is not None and target_col in record:
            record[target_col] = val

    return record


def apply_all_business_rules(
    record: Dict[str, Any],
    servicenow_record: Dict[str, Any],
    ic_lookup: Dict[str, str],
    config
) -> Dict[str, Any]:
    """
    Apply all business rules (IC, Region, Bank) to a record.

    Args:
        record: The Master Tracker record (will be modified).
        servicenow_record: The corresponding ServiceNow record.
        ic_lookup: Normalized IC-to-Region lookup dictionary.
        config: Config object.

    Returns:
        Updated record dict.
    """
    ic_config = config.ic_rule
    region_config = config.region_rule
    bank_config = config.bank_rule

    # IC Determination
    ic_value = determine_ic(
        servicenow_ic=servicenow_record.get(ic_config.get("ic_source_column", "IC")),
        priority=servicenow_record.get(ic_config.get("priority_column", "Priority")),
        proposed_by=servicenow_record.get(ic_config.get("proposed_by_column", "Proposed By")),
        priority_config=config.priority_normalization
    )
    record[ic_config.get("target_column", "IC")] = ic_value if ic_value else ""

    # Region Determination
    region_value, matched = determine_region(
        ic_value=ic_value,
        ic_lookup=ic_lookup,
        case_sensitive=region_config.get("case_sensitive", False)
    )
    record[region_config.get("target_column", "Region")] = region_value if region_value else ""

    # Bank Determination
    bank_value = determine_bank(
        assignment_group=record.get(bank_config.get("source_column", "Assignment group"), ""),
        svb_keyword=bank_config.get("svb_keyword", "SVB"),
        svb_value=bank_config.get("svb_value", "SVB"),
        default_value=bank_config.get("default_value", ""),
        case_sensitive=bank_config.get("case_sensitive", False)
    )
    record[bank_config.get("target_column", "Bank")] = bank_value

    # Duration Calculation using Actual Incident Start and Actual Incident Resolve
    try:
        start_val = servicenow_record.get("Actual Incident Start")
        resolve_val = servicenow_record.get("Actual Incident Resolve")
        if start_val and resolve_val and pd.notna(start_val) and pd.notna(resolve_val):
            start_dt = pd.to_datetime(start_val)
            resolve_dt = pd.to_datetime(resolve_val)
            duration_td = resolve_dt - start_dt
            # Calculate total days (float) for native Excel duration formatting
            # 1.0 in Excel = 1 Day. So difference in seconds / 86400 allows native formatting.
            record["Duration"] = duration_td.total_seconds() / 86400.0
        else:
            record["Duration"] = ""
    except Exception:
        record["Duration"] = ""

    return record
