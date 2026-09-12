"""
Report Generator
Produces reconciliation summaries, change reports, and validation summaries.
"""

from datetime import datetime
from typing import Any, Dict, List, Optional

from .reconciler import ReconciliationResult
from .validator import ValidationResult
from .normalizer import normalize_whitespace


def generate_reconciliation_summary(
    master_before_count: int,
    servicenow_count: int,
    recon: ReconciliationResult,
    input_validation: ValidationResult,
    output_validation: Optional[ValidationResult] = None
) -> Dict[str, Any]:
    """
    Generate a comprehensive reconciliation summary.

    Returns:
        Dict with all summary data for reporting.
    """
    summary = recon.summary()

    svb_count = sum(
        1 for r in (recon.updated_records + recon.new_records + recon.historical_records)
        if normalize_whitespace(r.get("Bank", "")) == "SVB"
    )

    blank_region_count = sum(
        1 for r in (recon.updated_records + recon.new_records)
        if normalize_whitespace(r.get("Region")) is None
    )

    report = {
        "summary": {
            "Run Timestamp": datetime.now().isoformat(),
            "Master Records Before": master_before_count,
            "ServiceNow Input Records": servicenow_count,
            "Matched (Updated)": summary["updated"],
            "New Records Added": summary["new"],
            "Historical Retained": summary["historical_retained"],
            "Total Output Records": summary["total_output"],
            "IC Lookup Misses": summary["ic_lookup_misses"],
            "Bank = SVB Count": svb_count,
            "Blank Region Count (new/updated)": blank_region_count,
            "Input Validation Errors": len(input_validation.errors),
            "Input Validation Warnings": len(input_validation.warnings),
            "Output Validation Errors": len(output_validation.errors) if output_validation else "N/A",
            "Output Validation Warnings": len(output_validation.warnings) if output_validation else "N/A",
            "Status": "PASSED" if (input_validation.is_valid and
                                   (output_validation is None or output_validation.is_valid))
                      else "FAILED",
        },
        "changes": dict(recon.change_details),
        "new_incidents": [
            normalize_whitespace(r.get("Number", ""))
            for r in recon.new_records
        ],
        "historical_incidents": [
            normalize_whitespace(r.get("Number", ""))
            for r in recon.historical_records
        ],
        "ic_misses": recon.ic_lookup_misses,
        "warnings": recon.warnings + input_validation.warnings + (
            output_validation.warnings if output_validation else []
        ),
        "errors": input_validation.errors + (
            output_validation.errors if output_validation else []
        ),
    }

    return report


def format_console_report(report: Dict[str, Any]) -> str:
    """
    Format the reconciliation report for console output.

    Returns:
        Formatted string.
    """
    lines = []
    lines.append("=" * 60)
    lines.append("FCB INCIDENT TRACKER — RECONCILIATION REPORT")
    lines.append("=" * 60)
    lines.append("")

    lines.append("RUN SUMMARY")
    lines.append("-" * 50)
    for key, value in report["summary"].items():
        lines.append(f"  {key:40s}: {value}")
    lines.append("")

    # New incidents
    new_incidents = report.get("new_incidents", [])
    if new_incidents:
        lines.append(f"NEW INCIDENTS ({len(new_incidents)})")
        lines.append("-" * 50)
        for num in new_incidents:
            lines.append(f"  + {num}")
        lines.append("")

    # Updated incidents with changes
    changes = report.get("changes", {})
    if changes:
        lines.append(f"UPDATED INCIDENTS ({len(changes)})")
        lines.append("-" * 50)
        for number, change_list in changes.items():
            lines.append(f"  {number}:")
            for change in change_list:
                lines.append(f"    → {change}")
        lines.append("")

    # Historical retained
    historical = report.get("historical_incidents", [])
    if historical:
        lines.append(f"HISTORICAL RETAINED ({len(historical)})")
        lines.append("-" * 50)
        for num in historical:
            lines.append(f"  ● {num}")
        lines.append("")

    # IC Lookup Misses
    ic_misses = report.get("ic_misses", [])
    if ic_misses:
        lines.append(f"IC LOOKUP MISSES ({len(ic_misses)})")
        lines.append("-" * 50)
        for miss in ic_misses:
            lines.append(f"  ⚠ {miss}")
        lines.append("")

    # Warnings
    warnings = report.get("warnings", [])
    if warnings:
        lines.append(f"WARNINGS ({len(warnings)})")
        lines.append("-" * 50)
        for w in warnings:
            lines.append(f"  ⚠ {w}")
        lines.append("")

    # Errors
    errors = report.get("errors", [])
    if errors:
        lines.append(f"ERRORS ({len(errors)})")
        lines.append("-" * 50)
        for e in errors:
            lines.append(f"  ✗ {e}")
        lines.append("")

    lines.append("=" * 60)
    status = report["summary"].get("Status", "UNKNOWN")
    lines.append(f"FINAL STATUS: {status}")
    lines.append("=" * 60)

    return "\n".join(lines)
