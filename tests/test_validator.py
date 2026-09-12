"""
Tests for the Validator module.
Tests column validation, Number uniqueness, IC lookup validation, and output invariants.

Covers test cases 11-15, 19 from requirements.
"""

import pytest
import pandas as pd

from src.validator import (
    ValidationResult,
    validate_required_columns,
    validate_number_column,
    validate_ic_lookup,
    validate_output,
)


class TestValidationResult:
    def test_valid_when_no_errors(self):
        result = ValidationResult()
        assert result.is_valid is True

    def test_invalid_when_errors(self):
        result = ValidationResult()
        result.add_error("Test error")
        assert result.is_valid is False

    def test_valid_with_warnings_only(self):
        result = ValidationResult()
        result.add_warning("Test warning")
        assert result.is_valid is True

    def test_merge(self):
        r1 = ValidationResult()
        r1.add_error("err1")
        r2 = ValidationResult()
        r2.add_warning("warn1")
        r1.merge(r2)
        assert len(r1.errors) == 1
        assert len(r1.warnings) == 1


class TestValidateRequiredColumns:
    def test_all_present(self):
        df = pd.DataFrame(columns=["Number", "State", "Priority"])
        result = validate_required_columns(df, ["Number", "State"], "Test")
        assert result.is_valid is True

    def test_missing_column(self):
        """TEST 14: Missing required column -> error."""
        df = pd.DataFrame(columns=["Number", "State"])
        result = validate_required_columns(df, ["Number", "State", "Priority"], "Test")
        assert result.is_valid is False
        assert any("Priority" in e for e in result.errors)

    def test_extra_columns_warning(self):
        """Schema drift: extra columns produce a warning, not error."""
        df = pd.DataFrame(columns=["Number", "State", "ExtraCol"])
        result = validate_required_columns(df, ["Number", "State"], "Test")
        assert result.is_valid is True
        assert len(result.warnings) > 0

    def test_duplicate_headers(self):
        df = pd.DataFrame([[1, 2, 3]], columns=["Number", "State", "Number"])
        result = validate_required_columns(df, ["Number", "State"], "Test")
        assert result.is_valid is False
        assert any("Duplicate" in e for e in result.errors)


class TestValidateNumberColumn:
    def test_valid_numbers(self):
        df = pd.DataFrame({"Number": ["INC001", "INC002", "INC003"]})
        result = validate_number_column(df, "Number", "Test")
        assert result.is_valid is True

    def test_duplicate_numbers(self):
        """TEST 11/12: Duplicate Number -> error."""
        df = pd.DataFrame({"Number": ["INC001", "INC001", "INC003"]})
        result = validate_number_column(df, "Number", "Test")
        assert result.is_valid is False
        assert any("Duplicate" in e for e in result.errors)

    def test_blank_numbers(self):
        """TEST 13: Blank Number -> error."""
        df = pd.DataFrame({"Number": ["INC001", "", "INC003"]})
        result = validate_number_column(df, "Number", "Test")
        assert result.is_valid is False
        assert any("blank" in e.lower() for e in result.errors)

    def test_none_numbers(self):
        """TEST 13 variant: None Number -> error."""
        df = pd.DataFrame({"Number": ["INC001", None, "INC003"]})
        result = validate_number_column(df, "Number", "Test")
        assert result.is_valid is False

    def test_missing_column(self):
        df = pd.DataFrame({"State": ["Open"]})
        result = validate_number_column(df, "Number", "Test")
        assert result.is_valid is False


class TestValidateICLookup:
    def test_valid_lookup(self):
        df = pd.DataFrame({"IC": ["Alice", "Bob"], "Region": ["US", "India"]})
        result = validate_ic_lookup(df, "IC", "Region")
        assert result.is_valid is True

    def test_missing_ic_column(self):
        """TEST 15: Missing lookup -> error."""
        df = pd.DataFrame({"Name": ["Alice"], "Region": ["US"]})
        result = validate_ic_lookup(df, "IC", "Region")
        assert result.is_valid is False

    def test_blank_ic_key(self):
        df = pd.DataFrame({"IC": ["", "Bob"], "Region": ["US", "India"]})
        result = validate_ic_lookup(df, "IC", "Region")
        assert result.is_valid is False
        assert any("blank" in e.lower() for e in result.errors)

    def test_duplicate_ic_key(self):
        df = pd.DataFrame({"IC": ["Alice", "Alice"], "Region": ["US", "India"]})
        result = validate_ic_lookup(df, "IC", "Region")
        assert result.is_valid is False
        assert any("Duplicate" in e for e in result.errors)


class TestValidateOutput:
    def test_valid_output(self):
        master = pd.DataFrame({"Number": ["INC001", "INC002"]})
        output = pd.DataFrame({"Number": ["INC001", "INC002", "INC003"]})
        result = validate_output(output, master, "Number")
        assert result.is_valid is True

    def test_historical_data_loss(self):
        """TEST 19: Historical record count must not decrease."""
        master = pd.DataFrame({"Number": ["INC001", "INC002", "INC003"]})
        output = pd.DataFrame({"Number": ["INC001", "INC002"]})
        result = validate_output(output, master, "Number")
        assert result.is_valid is False
        assert any("HISTORICAL DATA LOSS" in e for e in result.errors)

    def test_duplicate_in_output(self):
        master = pd.DataFrame({"Number": ["INC001"]})
        output = pd.DataFrame({"Number": ["INC001", "INC001"]})
        result = validate_output(output, master, "Number")
        assert result.is_valid is False
        assert any("Duplicate" in e for e in result.errors)

    def test_count_decrease(self):
        master = pd.DataFrame({"Number": ["INC001", "INC002", "INC003"]})
        output = pd.DataFrame({"Number": ["INC004"]})
        result = validate_output(output, master, "Number")
        assert result.is_valid is False
