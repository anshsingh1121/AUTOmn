"""
Input and Output Validator
Validates data integrity at each stage of the pipeline.
"""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set

import pandas as pd

from .normalizer import normalize_whitespace


@dataclass
class ValidationResult:
    """Container for validation results with blocking errors and warnings."""
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        """True if no blocking errors."""
        return len(self.errors) == 0

    def add_error(self, message: str):
        self.errors.append(message)

    def add_warning(self, message: str):
        self.warnings.append(message)

    def merge(self, other: "ValidationResult"):
        """Merge another ValidationResult into this one."""
        self.errors.extend(other.errors)
        self.warnings.extend(other.warnings)

    def summary(self) -> str:
        lines = []
        lines.append(f"Validation: {'PASSED' if self.is_valid else 'FAILED'}")
        lines.append(f"  Errors:   {len(self.errors)}")
        lines.append(f"  Warnings: {len(self.warnings)}")
        if self.errors:
            lines.append("  --- Errors ---")
            for e in self.errors:
                lines.append(f"    ✗ {e}")
        if self.warnings:
            lines.append("  --- Warnings ---")
            for w in self.warnings:
                lines.append(f"    ⚠ {w}")
        return "\n".join(lines)


def validate_required_columns(
    df: pd.DataFrame,
    required_columns: List[str],
    source_name: str
) -> ValidationResult:
    """
    Validate that all required columns are present in a DataFrame.

    Args:
        df: The DataFrame to validate.
        required_columns: List of required column names.
        source_name: Human-readable source name (e.g., 'ServiceNow', 'Master Tracker').

    Returns:
        ValidationResult with errors for any missing columns.
    """
    result = ValidationResult()
    actual_columns = set(df.columns)

    for col in required_columns:
        if col not in actual_columns:
            result.add_error(f"[{source_name}] Missing required column: '{col}'")

    # Check for duplicate headers
    if len(df.columns) != len(set(df.columns)):
        seen = set()
        duplicates = set()
        for col in df.columns:
            if col in seen:
                duplicates.add(col)
            seen.add(col)
        result.add_error(f"[{source_name}] Duplicate column headers detected: {duplicates}")

    # Warn about unexpected columns (schema drift detection)
    expected = set(required_columns)
    extra = actual_columns - expected
    if extra:
        result.add_warning(
            f"[{source_name}] Additional columns detected (not in schema): "
            f"{sorted(extra)}"
        )

    return result


def validate_number_column(
    df: pd.DataFrame,
    number_column: str,
    source_name: str,
    block_on_duplicates: bool = True,
    block_on_blanks: bool = True
) -> ValidationResult:
    """
    Validate the Number (incident key) column for blanks and duplicates.

    Args:
        df: The DataFrame to validate.
        number_column: Name of the Number column.
        source_name: Human-readable source name.
        block_on_duplicates: If True, duplicate Numbers produce errors.
        block_on_blanks: If True, blank Numbers produce errors.

    Returns:
        ValidationResult.
    """
    result = ValidationResult()

    if number_column not in df.columns:
        result.add_error(f"[{source_name}] Number column '{number_column}' not found.")
        return result

    numbers = df[number_column]

    # Check for blanks
    blank_mask = numbers.apply(lambda x: normalize_whitespace(x) is None)
    blank_count = blank_mask.sum()
    if blank_count > 0:
        msg = f"[{source_name}] {blank_count} blank Number value(s) detected."
        if block_on_blanks:
            result.add_error(msg)
        else:
            result.add_warning(msg)

    # Check for duplicates (among non-blank values)
    non_blank = numbers[~blank_mask]
    normalized = non_blank.apply(lambda x: normalize_whitespace(x))
    duplicates = normalized[normalized.duplicated(keep=False)]
    if len(duplicates) > 0:
        dup_values = sorted(set(duplicates.values))
        msg = (
            f"[{source_name}] Duplicate Number value(s) detected: "
            f"{dup_values}"
        )
        if block_on_duplicates:
            result.add_error(msg)
        else:
            result.add_warning(msg)

    return result


def validate_ic_lookup(
    df: pd.DataFrame,
    ic_column: str,
    region_column: str,
    source_name: str = "IC Lookup"
) -> ValidationResult:
    """
    Validate the IC Lookup table.

    Checks for:
        - Required columns
        - Blank IC keys
        - Duplicate IC keys

    Returns:
        ValidationResult.
    """
    result = ValidationResult()

    if ic_column not in df.columns:
        result.add_error(f"[{source_name}] Missing IC column: '{ic_column}'")
    if region_column not in df.columns:
        result.add_error(f"[{source_name}] Missing Region column: '{region_column}'")
    if not result.is_valid:
        return result

    # Check for blank IC keys
    blank_mask = df[ic_column].apply(lambda x: normalize_whitespace(x) is None)
    blank_count = blank_mask.sum()
    if blank_count > 0:
        result.add_error(f"[{source_name}] {blank_count} blank IC key(s) detected.")

    # Check for duplicate IC keys (after normalization)
    non_blank = df[~blank_mask]
    normalized = non_blank[ic_column].apply(
        lambda x: normalize_whitespace(x).lower() if normalize_whitespace(x) else None
    )
    duplicates = normalized[normalized.duplicated(keep=False)]
    if len(duplicates) > 0:
        dup_values = sorted(set(duplicates.values))
        result.add_error(
            f"[{source_name}] Duplicate IC key(s) detected: {dup_values}"
        )

    return result


def validate_output(
    output_df: pd.DataFrame,
    master_before_df: pd.DataFrame,
    number_column: str
) -> ValidationResult:
    """
    Validate the output DataFrame against business invariants.

    Invariants:
        1. All historical master Numbers must be present in output.
        2. No duplicate Numbers in output.
        3. Output count >= Master before count (can only grow or stay same).

    Args:
        output_df: The merged/output DataFrame.
        master_before_df: The original Master Tracker DataFrame.
        number_column: Name of the Number column.

    Returns:
        ValidationResult.
    """
    result = ValidationResult()

    if number_column not in output_df.columns:
        result.add_error(f"[Output] Missing Number column: '{number_column}'")
        return result
    if number_column not in master_before_df.columns:
        result.add_error(f"[Master Before] Missing Number column: '{number_column}'")
        return result

    # Invariant 1: Historical retention
    master_numbers = set(
        master_before_df[number_column]
        .apply(lambda x: normalize_whitespace(x))
        .dropna()
    )
    output_numbers = set(
        output_df[number_column]
        .apply(lambda x: normalize_whitespace(x))
        .dropna()
    )

    missing = master_numbers - output_numbers
    if missing:
        result.add_error(
            f"[Output] HISTORICAL DATA LOSS: {len(missing)} master record(s) "
            f"missing from output: {sorted(missing)}"
        )

    # Invariant 2: No duplicate Numbers in output
    output_nums = output_df[number_column].apply(lambda x: normalize_whitespace(x)).dropna()
    dup = output_nums[output_nums.duplicated(keep=False)]
    if len(dup) > 0:
        dup_values = sorted(set(dup.values))
        result.add_error(
            f"[Output] Duplicate Number(s) in output: {dup_values}"
        )

    # Invariant 3: Count must not decrease
    if len(output_df) < len(master_before_df):
        result.add_error(
            f"[Output] Record count decreased: "
            f"master before={len(master_before_df)}, output={len(output_df)}"
        )

    return result
