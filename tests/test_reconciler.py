"""
Tests for the Reconciler module.
Tests incident matching, new/historical categorization, and idempotency.

Covers test cases 1-3, 19-20 from requirements.
"""

import pytest
import pandas as pd

from src.reconciler import reconcile, reconciliation_to_dataframe, ReconciliationResult
from src.business_rules import build_ic_lookup_map


class TestReconcile:
    """Tests for the core reconciliation logic."""

    @pytest.fixture
    def ic_lookup(self, sample_ic_lookup_df):
        lookup, _ = build_ic_lookup_map(sample_ic_lookup_df)
        return lookup

    def test_existing_incident_updated(self, sample_master_df, sample_servicenow_df, ic_lookup, config):
        """TEST 1: Existing incident matches -> record updated."""
        result = reconcile(sample_master_df, sample_servicenow_df, ic_lookup, config)

        assert len(result.updated_records) >= 1
        # INC0010001 should be updated
        updated_numbers = [r["Number"] for r in result.updated_records]
        assert "INC0010001" in updated_numbers

        # Check that State was updated for INC0010001
        for r in result.updated_records:
            if r["Number"] == "INC0010001":
                assert r["State"] == "Resolved"

    def test_new_incident_added(self, sample_master_df, sample_servicenow_df, ic_lookup, config):
        """TEST 2: New incident -> record added."""
        result = reconcile(sample_master_df, sample_servicenow_df, ic_lookup, config)

        assert len(result.new_records) >= 1
        new_numbers = [r["Number"] for r in result.new_records]
        assert "INC0010004" in new_numbers

    def test_historical_incident_retained(self, sample_master_df, sample_servicenow_df, ic_lookup, config):
        """TEST 3: Historical incident absent from ServiceNow -> retained unchanged."""
        result = reconcile(sample_master_df, sample_servicenow_df, ic_lookup, config)

        assert len(result.historical_records) >= 1
        historical_numbers = [r["Number"] for r in result.historical_records]
        # INC0010002 is in master but not matched by SN (since SN only has INC0010001 and INC0010004)
        assert "INC0010002" in historical_numbers or "INC0010003" in historical_numbers

    def test_historical_records_unchanged(self, sample_master_df, sample_servicenow_df, ic_lookup, config):
        """TEST 3 extended: Historical records keep original values."""
        result = reconcile(sample_master_df, sample_servicenow_df, ic_lookup, config)

        for r in result.historical_records:
            if r["Number"] == "INC0010003":
                assert r["State"] == "Resolved"
                assert r["Short description"] == "Historical resolved incident"

    def test_all_master_numbers_in_output(self, sample_master_df, sample_servicenow_df, ic_lookup, config):
        """TEST 19: All historical master Numbers present in output."""
        result = reconcile(sample_master_df, sample_servicenow_df, ic_lookup, config)

        master_numbers = set(sample_master_df["Number"].values)
        output_numbers = set(
            [r["Number"] for r in result.updated_records]
            + [r["Number"] for r in result.historical_records]
        )

        # All master numbers must be in output
        assert master_numbers.issubset(output_numbers)

    def test_output_count_not_less_than_master(self, sample_master_df, sample_servicenow_df, ic_lookup, config):
        """TEST 19: Output record count >= master record count."""
        result = reconcile(sample_master_df, sample_servicenow_df, ic_lookup, config)
        assert result.total_output_records >= len(sample_master_df)

    def test_change_details_recorded(self, sample_master_df, sample_servicenow_df, ic_lookup, config):
        """Change details are recorded for updated records."""
        result = reconcile(sample_master_df, sample_servicenow_df, ic_lookup, config)

        # INC0010001 had State changed from "In Progress" to "Resolved"
        assert "INC0010001" in result.change_details
        changes = result.change_details["INC0010001"]
        assert any("State" in c for c in changes)


class TestIdempotency:
    """TEST 20: Idempotency — same input processed twice produces same output."""

    @pytest.fixture
    def ic_lookup(self, sample_ic_lookup_df):
        lookup, _ = build_ic_lookup_map(sample_ic_lookup_df)
        return lookup

    def test_idempotent_reconciliation(self, sample_master_df, sample_servicenow_df, ic_lookup, config):
        """Running reconciliation twice with same SN produces equivalent output."""
        # Run 1
        result1 = reconcile(sample_master_df, sample_servicenow_df, ic_lookup, config)
        master_columns = list(sample_master_df.columns)
        output1_df = reconciliation_to_dataframe(result1, master_columns)

        # Run 2: use output1 as the new master
        result2 = reconcile(output1_df, sample_servicenow_df, ic_lookup, config)
        output2_df = reconciliation_to_dataframe(result2, master_columns)

        # Verify same record count
        assert len(output2_df) == len(output1_df)

        # Verify same Numbers (no duplicates introduced)
        nums1 = sorted(output1_df["Number"].tolist())
        nums2 = sorted(output2_df["Number"].tolist())
        assert nums1 == nums2

        # Verify no new records in run 2
        assert len(result2.new_records) == 0


class TestReconciliationToDataframe:
    def test_output_has_all_columns(self, sample_master_df, sample_servicenow_df, sample_ic_lookup_df, config):
        ic_lookup, _ = build_ic_lookup_map(sample_ic_lookup_df)
        result = reconcile(sample_master_df, sample_servicenow_df, ic_lookup, config)
        master_columns = list(sample_master_df.columns)
        output_df = reconciliation_to_dataframe(result, master_columns)

        for col in master_columns:
            assert col in output_df.columns
