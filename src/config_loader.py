"""
Configuration Loader
Loads and validates the JSON configuration file for the FCB Incident Tracker.
"""

import json
import os
from pathlib import Path
from typing import Any, Dict, Optional


class ConfigError(Exception):
    """Raised when configuration is invalid or missing."""
    pass


class Config:
    """Immutable configuration container for the FCB Incident Tracker."""

    def __init__(self, config_data: Dict[str, Any], base_dir: Optional[str] = None):
        self._data = config_data
        self._base_dir = base_dir or os.getcwd()

    # --- Top-level settings ---

    @property
    def mode(self) -> str:
        return self._data.get("mode", "development")

    @property
    def is_production(self) -> bool:
        return self.mode == "production"

    @property
    def dry_run(self) -> bool:
        return self._data.get("dry_run", True)

    # --- Paths (resolved relative to base_dir) ---

    def _resolve_path(self, relative_path: str) -> str:
        if os.path.isabs(relative_path):
            return relative_path
        return os.path.normpath(os.path.join(self._base_dir, relative_path))

    @property
    def master_tracker_path(self) -> str:
        return self._resolve_path(self._data["paths"]["master_tracker"])

    @property
    def servicenow_input_path(self) -> str:
        return self._resolve_path(self._data["paths"]["servicenow_input"])

    @property
    def ic_lookup_path(self) -> str:
        return self._resolve_path(self._data["paths"]["ic_lookup"])

    @property
    def output_directory(self) -> str:
        return self._resolve_path(self._data["paths"].get("output_directory", "output"))

    @property
    def archive_directory(self) -> str:
        return self._resolve_path(self._data["paths"].get("archive_directory", "archive"))

    # --- Sheet names ---

    @property
    def master_sheet(self) -> str:
        return self._data["sheets"]["master_sheet"]

    @property
    def servicenow_sheet(self) -> str:
        return self._data["sheets"]["servicenow_sheet"]

    @property
    def ic_lookup_sheet(self) -> str:
        return self._data["sheets"]["ic_lookup_sheet"]

    # --- Tables (optional) ---

    @property
    def master_table(self) -> Optional[str]:
        return self._data.get("tables", {}).get("master_table")

    @property
    def servicenow_table(self) -> Optional[str]:
        return self._data.get("tables", {}).get("servicenow_table")

    @property
    def ic_lookup_table(self) -> Optional[str]:
        return self._data.get("tables", {}).get("ic_lookup_table")

    # --- Field mappings ---

    @property
    def field_mapping(self) -> Dict[str, Dict[str, Any]]:
        return self._data.get("field_mapping", {})

    @property
    def unmapped_fields(self) -> Dict[str, Dict[str, Any]]:
        return self._data.get("unmapped_fields", {})

    # --- Business rules ---

    @property
    def ic_rule(self) -> Dict[str, Any]:
        return self._data["business_rules"]["ic_rule"]

    @property
    def region_rule(self) -> Dict[str, Any]:
        return self._data["business_rules"]["region_rule"]

    @property
    def bank_rule(self) -> Dict[str, Any]:
        return self._data["business_rules"]["bank_rule"]

    @property
    def priority_normalization(self) -> Dict[str, Any]:
        return self._data["business_rules"]["priority_normalization"]

    # --- Formula columns ---

    @property
    def formula_columns(self) -> Dict[str, Any]:
        return self._data.get("formula_columns", {})

    @property
    def preserve_all_formulas(self) -> bool:
        return self.formula_columns.get("preserve_all_detected", True)

    # --- Validation settings ---

    @property
    def validation(self) -> Dict[str, Any]:
        return self._data.get("validation", {})

    @property
    def required_servicenow_columns(self) -> list:
        return self.validation.get("required_servicenow_columns", [])

    @property
    def required_master_columns(self) -> list:
        return self.validation.get("required_master_columns", [])

    @property
    def required_ic_lookup_columns(self) -> list:
        return self.validation.get("required_ic_lookup_columns", [])

    @property
    def block_on_duplicate_numbers(self) -> bool:
        return self.validation.get("block_on_duplicate_numbers", True)

    @property
    def block_on_blank_numbers(self) -> bool:
        return self.validation.get("block_on_blank_numbers", True)


def load_config(config_path: str, base_dir: Optional[str] = None) -> Config:
    """
    Load configuration from a JSON file.

    Args:
        config_path: Path to config.json
        base_dir: Base directory for resolving relative paths.
                  Defaults to the directory containing config_path.

    Returns:
        Config object

    Raises:
        ConfigError: If the config file is missing, unreadable, or invalid.
    """
    config_path = os.path.abspath(config_path)

    if not os.path.isfile(config_path):
        raise ConfigError(f"Configuration file not found: {config_path}")

    if base_dir is None:
        # Default base_dir to the project root (parent of config/)
        config_dir = os.path.dirname(config_path)
        base_dir = os.path.dirname(config_dir) if os.path.basename(config_dir) == "config" else config_dir

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        raise ConfigError(f"Invalid JSON in configuration file: {e}")
    except IOError as e:
        raise ConfigError(f"Cannot read configuration file: {e}")

    # Validate required top-level keys
    required_keys = ["paths", "sheets", "field_mapping", "business_rules"]
    missing = [k for k in required_keys if k not in data]
    if missing:
        raise ConfigError(f"Missing required configuration sections: {missing}")

    required_paths = ["master_tracker", "servicenow_input", "ic_lookup"]
    missing_paths = [p for p in required_paths if p not in data.get("paths", {})]
    if missing_paths:
        raise ConfigError(f"Missing required path settings: {missing_paths}")

    required_sheets = ["master_sheet", "servicenow_sheet", "ic_lookup_sheet"]
    missing_sheets = [s for s in required_sheets if s not in data.get("sheets", {})]
    if missing_sheets:
        raise ConfigError(f"Missing required sheet settings: {missing_sheets}")

    return Config(data, base_dir)
