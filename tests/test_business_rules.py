"""
Tests for the Business Rules engine.
Tests IC determination, Region lookup, Bank/SVB rule, and field mapping.

Covers test cases 4-10 from requirements.
"""

import pytest
import pandas as pd

from src.business_rules import (
    determine_ic,
    build_ic_lookup_map,
    determine_region,
    determine_bank,
    apply_field_mapping,
    create_new_record,
)


class TestDetermineIC:
    """Tests for IC determination business rule (Test Cases 4-6)."""

    def test_ic_populated_uses_ic(self):
        """TEST 4: ServiceNow IC populated -> use IC."""
        result = determine_ic(
            servicenow_ic="Alice Johnson",
            priority="1 - Critical",
            proposed_by="Wendy Thomas"
        )
        assert result == "Alice Johnson"

    def test_ic_blank_priority_3_uses_proposed_by(self):
        """TEST 5: IC blank + Priority 3 -> use Proposed By."""
        result = determine_ic(
            servicenow_ic="",
            priority="3 - Moderate",
            proposed_by="Sam Parker"
        )
        assert result == "Sam Parker"

    def test_ic_blank_priority_3_int_uses_proposed_by(self):
        """TEST 5 variant: IC blank + Priority = 3 (integer)."""
        result = determine_ic(
            servicenow_ic=None,
            priority=3,
            proposed_by="Sam Parker"
        )
        assert result == "Sam Parker"

    def test_ic_blank_priority_not_3_blank(self):
        """TEST 6: IC blank + Priority != 3 -> IC blank."""
        result = determine_ic(
            servicenow_ic="",
            priority="2 - High",
            proposed_by="Sam Parker"
        )
        assert result is None

    def test_ic_blank_priority_1_blank(self):
        """TEST 6 variant: IC blank + Priority 1 -> IC blank."""
        result = determine_ic(
            servicenow_ic=None,
            priority="1 - Critical",
            proposed_by="Wendy Thomas"
        )
        assert result is None

    def test_ic_whitespace_only_treated_as_blank(self):
        """IC with only whitespace treated as blank."""
        result = determine_ic(
            servicenow_ic="   ",
            priority="3 - Moderate",
            proposed_by="Sam Parker"
        )
        assert result == "Sam Parker"

    def test_ic_populated_with_whitespace(self):
        """IC with leading/trailing whitespace is trimmed."""
        result = determine_ic(
            servicenow_ic="  Alice Johnson  ",
            priority="1 - Critical",
            proposed_by="Wendy Thomas"
        )
        assert result == "Alice Johnson"


class TestRegionLookup:
    """Tests for Region determination (Test Cases 7-8)."""

    @pytest.fixture
    def ic_lookup_map(self, sample_ic_lookup_df):
        lookup, _ = build_ic_lookup_map(sample_ic_lookup_df)
        return lookup

    def test_ic_found_in_lookup(self, ic_lookup_map):
        """TEST 7: IC exists in lookup -> Region populated."""
        region, matched = determine_region("Alice Johnson", ic_lookup_map)
        assert region == "US"
        assert matched is True

    def test_ic_found_india(self, ic_lookup_map):
        """TEST 7 variant: IC maps to India."""
        region, matched = determine_region("Carol Davis", ic_lookup_map)
        assert region == "India"
        assert matched is True

    def test_ic_not_in_lookup(self, ic_lookup_map):
        """TEST 8: IC not in lookup -> Region blank."""
        region, matched = determine_region("Unknown Person", ic_lookup_map)
        assert region is None
        assert matched is False

    def test_ic_blank(self, ic_lookup_map):
        """Blank IC -> Region blank."""
        region, matched = determine_region(None, ic_lookup_map)
        assert region is None
        assert matched is False

    def test_ic_case_insensitive(self, ic_lookup_map):
        """IC lookup is case-insensitive by default."""
        region, matched = determine_region("alice johnson", ic_lookup_map)
        assert region == "US"
        assert matched is True


class TestBuildICLookupMap:
    """Tests for IC Lookup map building."""

    def test_duplicate_key_warning(self):
        """Duplicate IC keys produce a warning."""
        df = pd.DataFrame([
            {"IC": "Alice Johnson", "Region": "US"},
            {"IC": "Alice Johnson", "Region": "India"},
        ])
        lookup, warnings = build_ic_lookup_map(df)
        assert len(lookup) == 1  # First occurrence kept
        assert lookup["alice johnson"] == "US"
        assert any("duplicate" in w.lower() for w in warnings)

    def test_blank_key_warning(self):
        """Blank IC keys produce a warning."""
        df = pd.DataFrame([
            {"IC": "", "Region": "US"},
            {"IC": "Alice Johnson", "Region": "India"},
        ])
        lookup, warnings = build_ic_lookup_map(df)
        assert len(lookup) == 1
        assert any("blank" in w.lower() for w in warnings)


class TestDetermineBank:
    """Tests for Bank/SVB determination (Test Cases 9-10)."""

    def test_svb_in_group(self):
        """TEST 9: Assignment Group contains SVB -> Bank = SVB."""
        assert determine_bank("Tier 3 Application Support SVB") == "SVB"

    def test_svb_at_start(self):
        """TEST 9 variant: SVB at start of group name."""
        assert determine_bank("SVB Go Support") == "SVB"

    def test_svb_standalone(self):
        """TEST 9 variant: Assignment Group is exactly SVB."""
        assert determine_bank("SVB") == "SVB"

    def test_svb_case_insensitive(self):
        """SVB matching is case-insensitive."""
        assert determine_bank("tier 3 application support svb") == "SVB"

    def test_no_svb(self):
        """TEST 10: No SVB -> Bank blank."""
        assert determine_bank("Tier 3 Application Support FCB") == ""

    def test_blank_group(self):
        """Blank Assignment Group -> Bank blank."""
        assert determine_bank(None) == ""
        assert determine_bank("") == ""


class TestFieldMapping:
    """Tests for field mapping with non-blank overwrite protection."""

    def test_update_non_blank_fields(self):
        """Fields with values are updated."""
        master = {"Number": "INC001", "State": "Open", "Assignee": "Old Person"}
        sn = {"Number": "INC001", "State": "Resolved", "Assigned To": "New Person"}
        mapping = {
            "State": {"source_column": "State", "target_column": "State", "overwrite_if_blank": False},
            "Assignee": {"source_column": "Assigned To", "target_column": "Assignee", "overwrite_if_blank": False},
        }
        updated, changes = apply_field_mapping(master, sn, mapping)
        assert updated["State"] == "Resolved"
        assert updated["Assignee"] == "New Person"
        assert len(changes) == 2

    def test_blank_source_does_not_overwrite(self):
        """Blank source does NOT overwrite existing value (non-blank overwrite protection)."""
        master = {"Number": "INC001", "Assignee": "Old Person"}
        sn = {"Number": "INC001", "Assigned To": ""}
        mapping = {
            "Assignee": {"source_column": "Assigned To", "target_column": "Assignee", "overwrite_if_blank": False},
        }
        updated, changes = apply_field_mapping(master, sn, mapping)
        assert updated["Assignee"] == "Old Person"
        assert len(changes) == 0

    def test_key_field_skipped(self):
        """Key fields (Number) are not modified."""
        master = {"Number": "INC001"}
        sn = {"Number": "INC001"}
        mapping = {
            "Number": {"source_column": "Number", "target_column": "Number", "is_key": True, "overwrite_if_blank": False},
        }
        updated, changes = apply_field_mapping(master, sn, mapping)
        assert updated["Number"] == "INC001"
        assert len(changes) == 0
