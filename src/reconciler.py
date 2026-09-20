"""
Reconciler
Number-based incident matching between ServiceNow and Master Tracker.
Categorizes records as MATCHED (update), NEW (add), or HISTORICAL (retain).
"""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple

import pandas as pd

from .business_rules import (
    apply_all_business_rules,
    apply_field_mapping,
    create_new_record,
)
from .normalizer import normalize_whitespace

def is_equivalent(val1: Any, val2: Any) -> bool:
    """Safely compare two values ignoring type coercion artifacts like '3.0' vs '3'."""
    str1 = (normalize_whitespace(val1) or "").strip().lower()
    str2 = (normalize_whitespace(val2) or "").strip().lower()
    
    if str1 == str2:
        return True
        
    # Try numeric comparison (handles '3.0' == '3')
    try:
        if float(str1) == float(str2):
            return True
    except (ValueError, TypeError):
        pass
        
    # Try date comparison (handles '2024-09-01 00:00:00' == '2024-09-01')
    try:
        if pd.to_datetime(str1) == pd.to_datetime(str2):
            return True
    except Exception:
        pass
        
    return False


@dataclass
class ReconciliationResult:
    """Container for reconciliation results."""
    updated_records: List[Dict[str, Any]] = field(default_factory=list)
    new_records: List[Dict[str, Any]] = field(default_factory=list)
    historical_records: List[Dict[str, Any]] = field(default_factory=list)
    change_details: Dict[str, List[str]] = field(default_factory=dict)
    ic_lookup_misses: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def total_output_records(self) -> int:
        return (
            len(self.updated_records)
            + len(self.new_records)
            + len(self.historical_records)
        )

    def summary(self) -> Dict[str, int]:
        return {
            "updated": len(self.updated_records),
            "new": len(self.new_records),
            "historical_retained": len(self.historical_records),
            "total_output": self.total_output_records,
            "ic_lookup_misses": len(self.ic_lookup_misses),
            "changes_recorded": len(self.change_details),
            "warnings": len(self.warnings),
        }


def reconcile(
    master_df: pd.DataFrame,
    servicenow_df: pd.DataFrame,
    ic_lookup: Dict[str, str],
    config
) -> ReconciliationResult:
    """
    Reconcile ServiceNow data against the Master Tracker.

    Algorithm:
        1. Index master records by normalized Number.
        2. For each ServiceNow record:
           a. If Number exists in master → UPDATE (apply field mapping + business rules).
           b. If Number does not exist → ADD NEW (create record + apply business rules).
        3. For each master record not matched by ServiceNow → RETAIN (historical).

    Args:
        master_df: Master Tracker DataFrame.
        servicenow_df: ServiceNow export DataFrame.
        ic_lookup: Normalized IC-to-Region lookup dictionary.
        config: Config object.

    Returns:
        ReconciliationResult with categorized records.
    """
    result = ReconciliationResult()
    field_mapping = config.field_mapping
    master_columns = list(master_df.columns)

    # Build master index by Number
    master_index = {}
    for idx, row in master_df.iterrows():
        number = normalize_whitespace(row.get("Number"))
        if number is not None:
            master_index[number] = dict(row)

    # Track which master records have been matched
    matched_master_numbers: Set[str] = set()

    # Process ServiceNow records
    for idx, sn_row in servicenow_df.iterrows():
        sn_record = dict(sn_row)
        sn_number = normalize_whitespace(sn_record.get("Number"))

        if sn_number is None:
            result.warnings.append(f"ServiceNow row {idx}: blank Number — skipped.")
            continue

        if sn_number in master_index:
            # MATCHED — update existing record
            matched_master_numbers.add(sn_number)
            master_record = master_index[sn_number]

            # Apply field mapping with non-blank overwrite protection
            updated_record, _ = apply_field_mapping(
                master_record, sn_record, field_mapping
            )

            # Apply business rules (IC, Region, Bank)
            updated_record = apply_all_business_rules(
                updated_record, sn_record, ic_lookup, config
            )

            # Explicitly detect all changes (including business rule modifications)
            changes = []
            for col, new_val in updated_record.items():
                old_val = master_record.get(col, "")
                if not is_equivalent(old_val, new_val):
                    str_old = normalize_whitespace(old_val) or ""
                    str_new = normalize_whitespace(new_val) or ""
                    changes.append(f"{col}: '{str_old}' -> '{str_new}'")

            # Track IC lookup misses
            ic_val = normalize_whitespace(updated_record.get("IC", ""))
            if ic_val and ic_val.lower() not in ic_lookup:
                # Check with original casing normalization
                from .normalizer import normalize_for_matching
                normalized_ic = normalize_for_matching(ic_val, case_sensitive=False)
                if normalized_ic and normalized_ic not in ic_lookup:
                    result.ic_lookup_misses.append(
                        f"{sn_number}: IC='{ic_val}' not found in lookup"
                    )

            result.updated_records.append(updated_record)
            if changes:
                result.change_details[sn_number] = changes
        else:
            # NEW — add new record
            new_record = create_new_record(sn_record, field_mapping, master_columns)

            # Apply business rules (IC, Region, Bank)
            new_record = apply_all_business_rules(
                new_record, sn_record, ic_lookup, config
            )

            # Track IC lookup misses
            ic_val = normalize_whitespace(new_record.get("IC", ""))
            if ic_val:
                from .normalizer import normalize_for_matching
                normalized_ic = normalize_for_matching(ic_val, case_sensitive=False)
                if normalized_ic and normalized_ic not in ic_lookup:
                    result.ic_lookup_misses.append(
                        f"{sn_number}: IC='{ic_val}' not found in lookup"
                    )

            result.new_records.append(new_record)

    # HISTORICAL — retain master records not matched by ServiceNow
    for number, master_record in master_index.items():
        if number not in matched_master_numbers:
            result.historical_records.append(master_record)

    return result


def reconciliation_to_dataframe(
    recon: ReconciliationResult,
    master_columns: List[str]
) -> pd.DataFrame:
    """
    Convert reconciliation result to a single output DataFrame.

    Order: updated records, then historical records, then new records
    (preserves existing order as much as possible, appends new at end).

    Args:
        recon: ReconciliationResult.
        master_columns: Ordered list of column names.

    Returns:
        Output DataFrame with all records.
    """
    all_records = (
        recon.updated_records
        + recon.historical_records
        + recon.new_records
    )

    if not all_records:
        return pd.DataFrame(columns=master_columns)

    df = pd.DataFrame(all_records)

    # Ensure all master columns exist in output
    for col in master_columns:
        if col not in df.columns:
            df[col] = ""

    # Reorder columns to match master
    extra_cols = [c for c in df.columns if c not in master_columns]
    df = df[master_columns + extra_cols]

    return df
