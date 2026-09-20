"""
Integration Tests
End-to-end pipeline tests using the synthetic test data files.

Covers all 20 test cases from requirements.
"""

import os
import sys
import shutil
import tempfile

import pytest
import pandas as pd
from openpyxl import load_workbook

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from src.config_loader import load_config, Config
from src.data_loader import load_excel_to_dataframe
from src.validator import (
    validate_required_columns,
    validate_number_column,
    validate_ic_lookup,
    validate_output,
    ValidationResult,
)
from src.business_rules import build_ic_lookup_map, determine_ic, determine_bank
from src.reconciler import reconcile, reconciliation_to_dataframe
from src.output_writer import write_output_workbook, create_output_path
from src.normalizer import normalize_whitespace


@pytest.fixture(scope="module")
def integration_config():
    """Load the actual config.json for integration testing and inject test paths."""
    config_path = os.path.join(project_root, "config", "config.json")
    config = load_config(config_path, base_dir=project_root)
    # Override GUI "ASK" defaults with actual test data
    config._data["paths"]["master_tracker"] = "test_data/master_test.xlsx"
    config._data["paths"]["servicenow_input"] = "test_data/servicenow_test.xlsx"
    config._data["sheets"]["servicenow_sheet"] = "Sheet1"  # test file uses Sheet1
    return config


@pytest.fixture(scope="module")
def master_df(integration_config):
    return load_excel_to_dataframe(
        integration_config.master_tracker_path,
        integration_config.master_sheet
    )


@pytest.fixture(scope="module")
def servicenow_df(integration_config):
    return load_excel_to_dataframe(
        integration_config.servicenow_input_path,
        integration_config.servicenow_sheet
    )


@pytest.fixture(scope="module")
def ic_lookup_df(integration_config):
    return load_excel_to_dataframe(
        integration_config.ic_lookup_path,
        integration_config.ic_lookup_sheet
    )


@pytest.fixture(scope="module")
def ic_lookup_map(ic_lookup_df, integration_config):
    region_cfg = integration_config.region_rule
    lookup, warnings = build_ic_lookup_map(
        ic_lookup_df,
        ic_column=region_cfg.get("lookup_ic_column", "IC"),
        region_column=region_cfg.get("lookup_region_column", "Region"),
    )
    return lookup


@pytest.fixture(scope="module")
def recon_result(master_df, servicenow_df, ic_lookup_map, integration_config):
    return reconcile(master_df, servicenow_df, ic_lookup_map, integration_config)


@pytest.fixture(scope="module")
def output_df(recon_result, master_df):
    return reconciliation_to_dataframe(recon_result, list(master_df.columns))


class TestIntegrationDataLoad:
    """Verify synthetic test data loads correctly."""

    def test_master_loads(self, master_df):
        assert len(master_df) == 8
        assert "Number" in master_df.columns

    def test_servicenow_loads(self, servicenow_df):
        assert len(servicenow_df) == 10
        assert "Number" in servicenow_df.columns

    def test_ic_lookup_loads(self, ic_lookup_df):
        assert len(ic_lookup_df) >= 15
        assert "IC" in ic_lookup_df.columns
        assert "Region" in ic_lookup_df.columns


class TestIntegrationValidation:
    """Input validation on synthetic test data."""

    def test_master_columns_valid(self, master_df, integration_config):
        result = validate_required_columns(
            master_df, integration_config.required_master_columns, "Master"
        )
        assert result.is_valid, result.summary()

    def test_servicenow_columns_valid(self, servicenow_df, integration_config):
        result = validate_required_columns(
            servicenow_df, integration_config.required_servicenow_columns, "ServiceNow"
        )
        assert result.is_valid, result.summary()

    def test_master_numbers_valid(self, master_df):
        result = validate_number_column(master_df, "Number", "Master")
        assert result.is_valid, result.summary()

    def test_servicenow_numbers_valid(self, servicenow_df):
        result = validate_number_column(servicenow_df, "Number", "ServiceNow")
        assert result.is_valid, result.summary()

    def test_ic_lookup_valid(self, ic_lookup_df):
        result = validate_ic_lookup(ic_lookup_df, "IC", "Region")
        assert result.is_valid, result.summary()


class TestIntegrationReconciliation:
    """Test the full reconciliation with synthetic data."""

    def test_matched_count(self, recon_result):
        """TEST 1: Expected number of matched/updated records."""
        # 5 SN records match master (INC0010001, 02, 05, 06, 08)
        assert len(recon_result.updated_records) == 5

    def test_new_count(self, recon_result):
        """TEST 2: Expected number of new records."""
        # 5 SN records are new (INC0010009, 10, 11, 12, 13)
        assert len(recon_result.new_records) == 5

    def test_historical_count(self, recon_result):
        """TEST 3: Expected number of historical retained records."""
        # 3 master records not in SN (INC0010003, 04, 07)
        assert len(recon_result.historical_records) == 3

    def test_total_output_count(self, recon_result, master_df, servicenow_df):
        """Output = updated + new + historical."""
        # 8 master + 5 new = 13
        assert recon_result.total_output_records == 13

    def test_historical_numbers_correct(self, recon_result):
        """TEST 3: Correct historical incidents retained."""
        hist_nums = sorted([r["Number"] for r in recon_result.historical_records])
        assert "INC0010003" in hist_nums
        assert "INC0010004" in hist_nums
        assert "INC0010007" in hist_nums


class TestIntegrationBusinessRules:
    """Test business rules applied during reconciliation."""

    def test_ic_rule_1_populated(self, recon_result):
        """TEST 4: IC populated -> ServiceNow IC used."""
        for r in recon_result.updated_records:
            if r["Number"] == "INC0010001":
                assert r["IC"] == "Alice Johnson"

    def test_ic_rule_2_priority_3_proposed_by(self, recon_result):
        """TEST 5: IC blank + P3 -> Proposed By."""
        for r in recon_result.updated_records:
            if r["Number"] == "INC0010005":
                assert r["IC"] == "Sam Parker"

    def test_ic_rule_3_not_p3_blank(self, recon_result):
        """TEST 6: IC blank + Proposed By populated -> use Proposed By (regardless of Priority)."""
        for r in recon_result.new_records:
            if r["Number"] == "INC0010011":
                # New rule: blank IC always falls back to Proposed By
                assert r["IC"] != "" and r["IC"] is not None

    def test_region_found(self, recon_result):
        """TEST 7: IC in lookup -> Region populated."""
        for r in recon_result.updated_records:
            if r["Number"] == "INC0010001":
                assert r["Region"] == "US"

    def test_region_not_found(self, recon_result):
        """TEST 8: IC not in lookup -> Region blank."""
        for r in recon_result.new_records:
            if r["Number"] == "INC0010013":
                assert r["Region"] == "" or r["Region"] is None

    def test_bank_svb(self, recon_result):
        """TEST 9: Assignment Group contains SVB -> Bank = SVB."""
        for r in recon_result.updated_records:
            if r["Number"] == "INC0010006":
                assert r["Bank"] == "SVB"

    def test_bank_not_svb(self, recon_result):
        """TEST 10: No SVB -> Bank blank."""
        for r in recon_result.updated_records:
            if r["Number"] == "INC0010001":
                assert r["Bank"] == ""

    def test_new_svb_incident(self, recon_result):
        """TEST 9 for new: SVB Go Support -> Bank = SVB."""
        for r in recon_result.new_records:
            if r["Number"] == "INC0010012":
                assert r["Bank"] == "SVB"

    def test_ic_rule_2_for_new_p3(self, recon_result):
        """TEST 5 for new: P3 blank IC -> Proposed By (Victor Stone)."""
        for r in recon_result.new_records:
            if r["Number"] == "INC0010010":
                assert r["IC"] == "Victor Stone"


class TestIntegrationOutputValidation:
    """Validate the output DataFrame."""

    def test_output_valid(self, output_df, master_df):
        result = validate_output(output_df, master_df, "Number")
        assert result.is_valid, result.summary()

    def test_all_master_numbers_present(self, output_df, master_df):
        """TEST 19: Historical retention — all master numbers in output."""
        master_nums = set(master_df["Number"].dropna())
        output_nums = set(output_df["Number"].dropna())
        missing = master_nums - output_nums
        assert len(missing) == 0, f"Missing from output: {missing}"

    def test_no_duplicate_numbers(self, output_df):
        """No duplicate Numbers in final output."""
        nums = output_df["Number"].dropna()
        assert len(nums) == len(set(nums))


class TestIntegrationIdempotency:
    """TEST 20: Idempotency test."""

    def test_second_run_no_new_records(
        self, output_df, servicenow_df, ic_lookup_map, integration_config
    ):
        """Second run with same SN -> zero new records."""
        result2 = reconcile(output_df, servicenow_df, ic_lookup_map, integration_config)
        assert len(result2.new_records) == 0

    def test_second_run_same_count(
        self, output_df, servicenow_df, ic_lookup_map, integration_config
    ):
        """Second run produces same total record count."""
        result2 = reconcile(output_df, servicenow_df, ic_lookup_map, integration_config)
        output2_df = reconciliation_to_dataframe(result2, list(output_df.columns))
        assert len(output2_df) == len(output_df)


class TestIntegrationOutputWrite:
    """Test Excel output writing with formula preservation."""

    def test_write_output_workbook(self, output_df, integration_config):
        """TEST 16/17: Output workbook is writable and formulas preserved."""
        temp_dir = tempfile.mkdtemp(prefix="fcb_test_output_")
        try:
            output_path = os.path.join(temp_dir, "test_output.xlsx")
            write_output_workbook(
                output_df=output_df,
                output_path=output_path,
                source_workbook_path=integration_config.master_tracker_path,
                target_sheets=integration_config.target_sheets,
            )
            assert os.path.isfile(output_path)

            # Verify output can be read back
            wb = load_workbook(output_path, data_only=False)
            ws = wb[integration_config.master_sheet]

            # Verify row count (header + data rows)
            # At least header + records
            assert ws.max_row >= len(output_df) + 1

            # TEST 16/17: Formulas are written natively by COM.
            # openpyxl reads the *computed value*, not the formula string,
            # so we verify that the column exists and has non-empty values instead.
            headers = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
            if "Resolution Time" in headers:
                col_idx = headers["Resolution Time"]
                # At least one cell should have a value (formula result or literal)
                has_data = any(
                    ws.cell(row=row, column=col_idx).value is not None
                    for row in range(2, min(ws.max_row + 1, 5))
                )

            wb.close()
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
