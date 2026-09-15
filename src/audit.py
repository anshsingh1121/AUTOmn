"""
Audit Trail
Records execution metadata for each automation run.
Uses incident Numbers only — does not log sensitive descriptions.
"""

import json
import os
from datetime import datetime
from typing import Any, Dict, List, Optional


def create_audit_record(
    config,
    master_before_count: int,
    servicenow_count: int,
    recon_summary: Dict[str, int],
    validation_passed: bool,
    report_summary: Dict[str, Any],
    output_path: Optional[str] = None,
    archive_path: Optional[str] = None,
    errors: Optional[List[str]] = None,
    warnings: Optional[List[str]] = None
) -> Dict[str, Any]:
    """
    Create an audit record for this automation run.

    Args:
        config: Config object.
        master_before_count: Records in master before processing.
        servicenow_count: Records in ServiceNow input.
        recon_summary: Reconciliation summary dict.
        validation_passed: Whether all validations passed.
        report_summary: Full report summary dict.
        output_path: Path to the output file (if produced).
        archive_path: Path to the archive copy (if created).
        errors: List of error messages.
        warnings: List of warning messages.

    Returns:
        Audit record dict.
    """
    return {
        "timestamp": datetime.now().isoformat(),
        "mode": config.mode,
        "dry_run": config.dry_run,
        "source_files": {
            "master_tracker": os.path.basename(config.master_tracker_path),
            "servicenow_input": os.path.basename(config.servicenow_input_path),
            "ic_lookup": os.path.basename(config.ic_lookup_path),
        },
        "record_counts": {
            "master_before": master_before_count,
            "servicenow_input": servicenow_count,
            "updated": recon_summary.get("updated", 0),
            "new": recon_summary.get("new", 0),
            "historical_retained": recon_summary.get("historical_retained", 0),
            "total_output": recon_summary.get("total_output", 0),
        },
        "validation_passed": validation_passed,
        "output_file": os.path.basename(output_path) if output_path else None,
        "archive_file": os.path.basename(archive_path) if archive_path else None,
        "error_count": len(errors) if errors else 0,
        "warning_count": len(warnings) if warnings else 0,
        "status": "SUCCESS" if validation_passed else "FAILED",
    }


def save_audit_log(
    audit_record: Dict[str, Any],
    output_dir: str,
    filename: Optional[str] = None
) -> str:
    """
    Save the audit record to a JSON file inside a Year/Month folder structure.

    Args:
        audit_record: The audit record dict.
        output_dir: Base directory to save the audit log.
        filename: Optional filename. Defaults to audit_YYYYMMDD_HHMMSS.json.

    Returns:
        Path to the saved audit log.
    """
    now = datetime.now()
    nested_dir = os.path.join(output_dir, now.strftime("%Y"), now.strftime("%m_%B"))
    os.makedirs(nested_dir, exist_ok=True)

    if filename is None:
        timestamp = now.strftime("%Y%m%d_%H%M%S")
        filename = f"audit_{timestamp}.json"

    path = os.path.join(nested_dir, filename)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(audit_record, f, indent=2, default=str)

    return path
