"""
Pytest fixtures for FCB Incident Tracker tests.
Provides shared test data, config, and utilities.
"""

import os
import sys
import json
import tempfile
import shutil

import pytest
import pandas as pd
from openpyxl import Workbook

# Add project root to path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from src.config_loader import Config


@pytest.fixture
def sample_config_data():
    """Provide a sample config dict for testing."""
    return {
        "mode": "development",
        "dry_run": True,
        "paths": {
            "master_tracker": "test_data/master_test.xlsx",
            "servicenow_input": "test_data/servicenow_test.xlsx",
            "ic_lookup": "test_data/ic_lookup_test.xlsx",
            "output_directory": "test_output",
            "archive_directory": "test_output/archive",
        },
        "sheets": {
            "master_sheet": "FCB",
            "servicenow_sheet": "Sheet1",
            "ic_lookup_sheet": "IC Lookup",
        },
        "tables": {},
        "field_mapping": {
            "Number": {"source_column": "Number", "target_column": "Number", "is_key": True, "overwrite_if_blank": False},
            "Short description": {"source_column": "Short Description", "target_column": "Short description", "overwrite_if_blank": False},
            "Primary CI/Application": {"source_column": "Configuration Item", "target_column": "Primary CI/Application", "overwrite_if_blank": False},
            "Priority": {"source_column": "Priority", "target_column": "Priority", "overwrite_if_blank": False},
            "State": {"source_column": "State", "target_column": "State", "overwrite_if_blank": False},
            "Resolved": {"source_column": "Actual Incident Resolve", "target_column": "Resolved", "overwrite_if_blank": False},
            "Assignee": {"source_column": "Assigned To", "target_column": "Assignee", "overwrite_if_blank": False},
            "Assignment group": {"source_column": "Assignment Group", "target_column": "Assignment group", "overwrite_if_blank": False},
            "Caused by Change": {"source_column": "Caused by Change", "target_column": "Caused by Change", "overwrite_if_blank": False},
        },
        "unmapped_fields": {},
        "business_rules": {
            "ic_rule": {
                "ic_source_column": "IC",
                "proposed_by_column": "Proposed By",
                "priority_column": "Priority",
                "priority_3_class": 3,
                "target_column": "IC",
            },
            "region_rule": {
                "ic_column": "IC",
                "lookup_ic_column": "IC",
                "lookup_region_column": "Region",
                "target_column": "Region",
                "case_sensitive": False,
            },
            "bank_rule": {
                "source_column": "Assignment group",
                "target_column": "Bank",
                "svb_keyword": "SVB",
                "svb_value": "SVB",
                "default_value": "",
                "case_sensitive": False,
            },
            "priority_normalization": {
                "extract_numeric": True,
                "separator": " - ",
                "priority_3_class": 3,
            },
        },
        "formula_columns": {"preserve_all_detected": True},
        "validation": {
            "required_servicenow_columns": [
                "Number", "Short Description", "Configuration Item",
                "Priority", "State", "IC", "Proposed By",
                "Assignment Group", "Assigned To", "Caused by Change",
                "Actual Incident Resolve",
            ],
            "required_master_columns": [
                "Number", "Short description", "Primary CI/Application",
                "Priority", "State", "Resolved", "Created", "IC",
                "Assignee", "Assignment group", "Caused by Change",
                "Region", "Bank",
            ],
            "required_ic_lookup_columns": ["IC", "Region"],
            "block_on_duplicate_numbers": True,
            "block_on_blank_numbers": True,
        },
    }


@pytest.fixture
def config(sample_config_data):
    """Provide a Config object for testing."""
    return Config(sample_config_data, base_dir=project_root)


@pytest.fixture
def sample_master_df():
    """Provide a sample Master Tracker DataFrame."""
    return pd.DataFrame([
        {"Number": "INC0010001", "Short Description": "App login failure",
         "Configuration Item": "App-Portal-001", "Priority": "1 - Critical",
         "State": "In Progress", "Resolved": None, "Created": "2026-03-15",
         "IC": "Alice Johnson", "Assignee": "Bob Smith",
         "Assignment group": "Tier 3 Application Support FCB",
         "Caused by Change": "CHG0001001", "Vendor Caused?": "",
         "Region": "US", "Bank": ""},
        {"Number": "INC0010002", "Short Description": "Database timeout",
         "Configuration Item": "DB-Report-002", "Priority": "2 - High",
         "State": "In Progress", "Resolved": None, "Created": "2026-04-10",
         "IC": "Carol Davis", "Assignee": "Dan Wilson",
         "Assignment group": "Tier 2 Database Support",
         "Caused by Change": "CHG0001002", "Vendor Caused?": "",
         "Region": "India", "Bank": ""},
        {"Number": "INC0010003", "Short Description": "Historical resolved incident",
         "Configuration Item": "Net-Core-003", "Priority": "3 - Moderate",
         "State": "Resolved", "Resolved": "2026-01-20", "Created": "2026-01-18",
         "IC": "Eve Martinez", "Assignee": "Frank Brown",
         "Assignment group": "Tier 1 Network Support",
         "Caused by Change": "", "Vendor Caused?": "",
         "Region": "US", "Bank": ""},
    ])


@pytest.fixture
def sample_servicenow_df():
    """Provide a sample ServiceNow DataFrame."""
    return pd.DataFrame([
        {"Number": "INC0010001", "Short Description": "App login failure - updated",
         "Configuration Item": "App-Portal-001", "Priority": "1 - Critical",
         "State": "Resolved", "IC": "Alice Johnson", "Proposed By": "Wendy Thomas",
         "Assignment Group": "Tier 3 Application Support FCB", "Assigned To": "Bob Smith",
         "Caused by Change": "CHG0001001", "Actual Incident Resolve": "2026-09-01"},
        {"Number": "INC0010004", "Short Description": "New incident",
         "Configuration Item": "New-App-004", "Priority": "2 - High",
         "State": "Open", "IC": "", "Proposed By": "Sam Parker",
         "Assignment Group": "Tier 2 Application Support FCB", "Assigned To": "Dan Wilson",
         "Caused by Change": "", "Actual Incident Resolve": None},
    ])


@pytest.fixture
def sample_ic_lookup_df():
    """Provide a sample IC Lookup DataFrame."""
    return pd.DataFrame([
        {"IC": "Alice Johnson", "Region": "US"},
        {"IC": "Bob Smith", "Region": "US"},
        {"IC": "Carol Davis", "Region": "India"},
        {"IC": "Eve Martinez", "Region": "US"},
        {"IC": "Sam Parker", "Region": "US"},
    ])


@pytest.fixture
def temp_dir():
    """Provide a temporary directory that is cleaned up after the test."""
    d = tempfile.mkdtemp(prefix="fcb_test_")
    yield d
    shutil.rmtree(d, ignore_errors=True)
