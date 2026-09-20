# FCB Monthly Incident Tracker Automation — Technical Developer Documentation

**Version:** 1.0  
**Last Updated:** September 2026  
**Python Version:** 3.9+  
**Platform:** Windows (requires Excel COM / win32com)

---

## Table of Contents

1. [Architecture Overview](#1-architecture-overview)
2. [Project Structure](#2-project-structure)
3. [Pipeline Stages](#3-pipeline-stages)
4. [Module Reference](#4-module-reference)
5. [Configuration Reference](#5-configuration-reference)
6. [Business Rules Engine](#6-business-rules-engine)
7. [Reconciliation Algorithm](#7-reconciliation-algorithm)
8. [Output Strategy (COM vs openpyxl)](#8-output-strategy-com-vs-openpyxl)
9. [Pivot Updater](#9-pivot-updater)
10. [Data Flow Diagram](#10-data-flow-diagram)
11. [Key Design Decisions](#11-key-design-decisions)
12. [Dependencies](#12-dependencies)
13. [Building the Standalone Executable](#13-building-the-standalone-executable)
14. [Testing](#14-testing)
15. [Known Limitations & Future Enhancements](#15-known-limitations--future-enhancements)

---

## 1. Architecture Overview

The system is an **8-stage deterministic data pipeline** that reconciles monthly ServiceNow incident exports against a Year-To-Date Master Tracker workbook.

```
INPUT → VALIDATION → NORMALIZATION → RECONCILIATION →
BUSINESS RULES → OUTPUT → VALIDATION → REPORTING → AUDIT
```

### Core Principles

| Principle | Implementation |
|---|---|
| **Never delete historical data** | Master records absent from ServiceNow are retained as "Historical" |
| **Non-blank overwrite protection** | ServiceNow blank fields never overwrite populated Master fields |
| **Non-destructive Region rule** | If IC Lookup fails, manually entered Regions are preserved |
| **Staging approach** | Output is always written to a new file; original is never modified |
| **Deterministic processing** | No AI/LLM; all business rules are explicit, auditable, and repeatable |
| **Pivot-safe writes** | Uses COM automation (not openpyxl) to preserve Pivot Tables, Charts, and Macros |

---

## 2. Project Structure

```
FCB_Incident_Tracker_Automation/
├── config/
│   └── config.json              # All configurable settings (paths, sheets, mappings, rules)
├── src/
│   ├── __init__.py              # Package marker
│   ├── main.py                  # Pipeline orchestrator and CLI entry point
│   ├── config_loader.py         # JSON config loader and Config class
│   ├── data_loader.py           # Excel/CSV readers (pandas + openpyxl)
│   ├── normalizer.py            # String normalization, priority parsing, blank detection
│   ├── validator.py             # Input/output validation with blocking errors and warnings
│   ├── reconciler.py            # Number-based incident matching engine
│   ├── business_rules.py        # IC, Region, Bank, Duration rules + field mapping
│   ├── output_writer.py         # COM-based Excel writer + openpyxl report writer
│   ├── pivot_updater.py         # Python-calculated pivot tables + COM chart rendering
│   ├── report_generator.py      # Reconciliation summary and console reporting
│   ├── audit.py                 # JSON audit trail generation
│   └── gui_picker.py            # Tkinter file selection dialog for non-tech users
├── tests/
│   └── test_integration.py      # Integration test suite
├── test_data/                   # Synthetic test data (never contains real corporate data)
├── output/                      # Generated output (gitignored)
├── archive/                     # Timestamped backups (gitignored)
├── build_exe.bat                # PyInstaller build script for standalone .exe
├── Run_Tracker_Update.bat       # Batch script launcher (for users with Python installed)
├── requirements.txt             # Python dependencies
├── .gitignore                   # Excludes output/, archive/, corporate data
├── USER_MANUAL.txt              # Non-technical user documentation
└── TECHNICAL_DOCS.md            # This file
```

---

## 3. Pipeline Stages

The main orchestrator (`src/main.py::run_pipeline()`) executes 8 sequential stages:

| Stage | Module | Description | Blocking? |
|-------|--------|-------------|-----------|
| 1. INPUT | `data_loader` | Load Master Tracker (.xlsm), ServiceNow export (.xlsx/.csv), IC Lookup | Yes — fatal on load failure |
| 2. VALIDATION | `validator` | Check required columns, duplicate/blank Numbers, IC Lookup integrity | Yes — blocks pipeline on errors |
| 3. NORMALIZATION | `business_rules` | Build normalized IC→Region lookup dictionary | No |
| 4. RECONCILIATION | `reconciler` | Match by Number: categorize as UPDATED / NEW / HISTORICAL | No |
| 5. OUTPUT BUILD | `reconciler` | Convert reconciliation result to output DataFrame | No |
| 6. OUTPUT VALIDATION | `validator` | Verify no historical data loss, no duplicate Numbers, count invariant | Yes — blocks write on failure |
| 7. WRITE | `output_writer` | COM-based write to Excel (preserves Pivots/Charts/Macros) | Skipped in dry-run or if no changes |
| 8. REPORTING | `report_generator`, `audit` | Console report, Excel reconciliation report, JSON audit log | No |

### Idempotency

The pipeline is **idempotent**. If run twice with the same inputs:
- Stage 4 detects zero changes (no new records, no field-level differences)
- Stage 7 skips output generation entirely
- Console shows: `"UP TO DATE (NO CHANGES)"`

---

## 4. Module Reference

### `src/main.py` — Pipeline Orchestrator

**Entry point:** `main()` → `run_pipeline(config)`

- Parses CLI arguments (`--config`, `--dry-run`, `--production`)
- Invokes the GUI file picker (`gui_picker.prompt_for_file()`) when paths are set to `"ASK"` or missing
- Executes 8 pipeline stages sequentially
- Auto-opens the output folder on Windows after completion
- Returns exit code: `0` (success), `1` (validation failure), `2` (fatal error)

### `src/config_loader.py` — Configuration

**Class:** `Config(config_data, base_dir)`

An immutable container wrapping the parsed JSON config. All paths are resolved relative to `base_dir` (defaults to the parent of the `config/` directory).

Key properties:
- `master_tracker_path`, `servicenow_input_path`, `ic_lookup_path` — resolved absolute paths
- `master_sheet`, `servicenow_sheet`, `ic_lookup_sheet` — Excel sheet names
- `field_mapping` — source→target column mapping dict
- `ic_rule`, `region_rule`, `bank_rule` — business rule configuration
- `dry_run`, `mode` — runtime mode control

**Function:** `load_config(config_path, base_dir)` → `Config`

Validates required top-level keys (`paths`, `sheets`, `field_mapping`, `business_rules`) and required sub-keys. Raises `ConfigError` on missing or invalid config.

### `src/data_loader.py` — Data Loading

| Function | Purpose |
|---|---|
| `load_excel_to_dataframe(path, sheet)` | Generic Excel→DataFrame reader via `pd.read_excel()` + openpyxl engine |
| `load_master_tracker(config)` | Returns `(DataFrame, openpyxl.Workbook)` — workbook kept for formula preservation |
| `load_servicenow(config)` | Reads .xlsx or .csv; auto-cleans non-INC rows (e.g., "Grand Total") |
| `load_ic_lookup(config)` | Reads the IC Lookup reference table |
| `detect_formula_columns(ws)` | Scans an openpyxl worksheet for cells starting with `=` |

### `src/normalizer.py` — Data Normalization

| Function | Purpose |
|---|---|
| `normalize_whitespace(value)` | Strip whitespace, convert "nan" to None |
| `extract_priority_class(value)` | Parse "3 - Moderate" → `3` (int) |
| `is_priority_3(value)` | Boolean check for Priority 3 |
| `normalize_for_matching(value)` | Lowercase + strip for dictionary key matching |
| `is_blank(value)` | `True` for None, "", "nan", whitespace-only |
| `contains_substring(value, sub)` | Case-insensitive substring check |

### `src/validator.py` — Validation Engine

**Class:** `ValidationResult`
- `errors: List[str]` — blocking errors (pipeline aborts)
- `warnings: List[str]` — non-blocking informational messages
- `is_valid: bool` — True when `len(errors) == 0`
- `merge(other)` — combine two results

| Function | Validates |
|---|---|
| `validate_required_columns(df, cols, name)` | All required columns present; warns on extra columns (schema drift) |
| `validate_number_column(df, col, name)` | No blank Numbers; no duplicate Numbers |
| `validate_ic_lookup(df, ic_col, region_col)` | IC/Region columns exist; no blank ICs; no duplicate ICs |
| `validate_output(output_df, master_df, col)` | **Critical invariants:** all historical Numbers retained; no duplicates; count >= master |

### `src/reconciler.py` — Reconciliation Engine

**Class:** `ReconciliationResult`
- `updated_records` — existing master records updated with ServiceNow data
- `new_records` — brand new incidents not in master
- `historical_records` — master records not matched by ServiceNow (retained)
- `change_details` — field-level diff per incident number
- `ic_lookup_misses` — ICs not found in lookup

**Function:** `reconcile(master_df, servicenow_df, ic_lookup, config)` → `ReconciliationResult`

Algorithm:
1. Index master records by normalized Number → O(n) hash map
2. For each ServiceNow record:
   - If Number in master index → apply field mapping + business rules → UPDATED
   - If Number not in index → create new record + business rules → NEW
3. For each master record not matched → HISTORICAL (retained unchanged)

**Function:** `reconciliation_to_dataframe(recon, columns)` → `pd.DataFrame`

Flattens the result into: updated + historical + new (new records appended at end).

### `src/business_rules.py` — Business Rules Engine

#### IC Determination (`determine_ic`)
```
IC populated? → YES → use IC
               → NO  → use Proposed By (fallback)
                      → if both blank → None
```

#### Region Determination (`determine_region` + `apply_all_business_rules`)
```
IC found in IC Lookup? → YES → use looked-up Region
                        → NO  → keep existing Region (NON-DESTRUCTIVE)
                                if no existing Region → blank
```

#### Bank Determination (`determine_bank`)
```
Assignment Group contains "SVB" (case-insensitive)? → YES → "SVB"
                                                     → NO  → blank
```

#### Duration Calculation
```
Actual Incident Start and Actual Incident Resolve both present?
  → YES → (Resolve - Start) in days as float (Excel serial format)
  → NO  → blank
```

#### Field Mapping Engine (`apply_field_mapping`)
- Iterates through configured source→target mappings
- Respects `overwrite_if_blank: false` — never overwrites populated Master fields with blank ServiceNow values
- Records field-level changes as strings: `"Field: 'old' -> 'new'"`

### `src/output_writer.py` — Output Writer

#### COM-Based Write (`write_output_workbook`)

**Why COM instead of openpyxl?**  
openpyxl silently destroys Pivot Tables, Charts, Macros, and Data Models when it saves a workbook. The COM approach:

1. `shutil.copy2()` the pristine master template → output path
2. Open the copy with an invisible `Excel.Application` via `win32com`
3. Read headers from row 1 of each target sheet
4. Detect formula columns from row 2 (R1C1 format) — excludes Duration
5. `ClearContents()` on old data rows (row 2 → end)
6. Write entire DataFrame as a 2D array via `Range.Value = [[...]]` (fast bulk write)
7. Re-apply preserved formulas via `FormulaR1C1`
8. Format date columns as `yyyy-mm-dd`
9. Resize `ListObjects` (Excel Tables) and re-apply table styles
10. Call `pivot_updater.update_pivots_sheet()` for chart generation
11. `Save()` + `Close()` + cleanup COM objects

**Type Safety:** `_to_com_safe(val)` converts numpy scalars, NaN, NaT, pd.Timestamp, and other Python types into COM-marshallable types. Dates are converted to Excel serial date floats to prevent COM text-alignment bugs.

#### Reconciliation Report (`write_reconciliation_report_excel`)

Uses openpyxl (safe for NEW workbooks) to create a simple 4-sheet report:
- Summary, New Incidents, Changes, IC Lookup Misses

### `src/pivot_updater.py` — Pivot Table & Chart Engine

Bypasses Excel's native Pivot Table engine (blocked by corporate Trust Center security policies) by calculating everything in Python and writing raw values via COM.

#### Table 1: Priority × Month
- `pd.crosstab(Month, Priority)` with ordered Categorical months (Jan→Dec)
- Drops zero-incident months
- Adds Grand Total column and row
- Writes to Pivots sheet starting at row 3

#### Table 2: Region × Month (Last 3 with data)
- Filters to non-blank Region rows only
- `pd.crosstab(Month, Region)`
- Drops zero-incident months, then `.tail(3)` for last 3 months with actual data
- Adds Grand Total column and row
- Writes below Table 1 with 8-row gap

#### Chart: IMT Region Wise Incident Volume
- `ws.Shapes.AddChart2(201, 52)` — Style 201, Stacked Column (xlColumnStacked=52)
- `SetSourceData` bound to Table 2's data range (excluding Grand Total row/column)
- Data labels enabled on each series
- Positioned to the right of Table 2

### `src/report_generator.py` — Reporting

| Function | Purpose |
|---|---|
| `generate_reconciliation_summary(...)` | Builds a comprehensive report dict with summary metrics, changes, new incidents, IC misses, warnings, errors |
| `format_console_report(report)` | Formats the report as a human-readable multi-section console string |

### `src/audit.py` — Audit Trail

| Function | Purpose |
|---|---|
| `create_audit_record(...)` | Builds a dict with timestamp, mode, source files, record counts, validation status |
| `save_audit_log(record, dir)` | Writes JSON to `output/YYYY/MM_Month/audit_YYYYMMDD_HHMMSS.json` |

Security: Only incident Numbers are logged. No descriptions, PII, or sensitive data appears in audit logs.

### `src/gui_picker.py` — File Selection Dialog

Uses Python's built-in `tkinter.filedialog.askopenfilename()` to display a native Windows file picker. The dialog is forced to appear on top of other windows (`-topmost`). No external GUI libraries required.

---

## 5. Configuration Reference

The configuration file `config/config.json` controls all runtime behavior.

### `paths` — File Locations

| Key | Type | Description |
|---|---|---|
| `master_tracker` | string | Path to master tracker, or `"ASK"` to trigger GUI file picker |
| `servicenow_input` | string | Path to ServiceNow export, or `"ASK"` to trigger GUI file picker |
| `ic_lookup` | string | Path to IC Lookup file (often same as master tracker) |
| `output_directory` | string | Base directory for output files (default: `"output"`) |
| `archive_directory` | string | Base directory for backup copies (default: `"archive"`) |

### `sheets` — Excel Sheet Names

| Key | Default | Description |
|---|---|---|
| `master_sheet` | `"Consolidated"` | Data sheet in the master tracker |
| `servicenow_sheet` | `"Page 1"` | Sheet name in the ServiceNow download |
| `ic_lookup_sheet` | `"IC Lookup"` | Sheet containing IC→Region mapping |
| `target_sheets` | `["Consolidated", "FCB"]` | Sheets to write output data into |

### `field_mapping` — Column Mappings

Each entry maps a ServiceNow source column to a Master Tracker target column:

```json
{
  "Short description": {
    "source_column": "Short Description",
    "target_column": "Short description",
    "overwrite_if_blank": false,
    "is_key": false
  }
}
```

- `overwrite_if_blank: false` — if ServiceNow value is blank, keep the existing Master value
- `is_key: true` — marks the matching key column (Number); never overwritten

### `business_rules` — Rule Configuration

See [Section 6](#6-business-rules-engine) for detailed rule documentation.

### `validation` — Data Quality Rules

| Key | Description |
|---|---|
| `required_servicenow_columns` | Columns that must exist in the ServiceNow export |
| `required_master_columns` | Columns that must exist in the Master Tracker |
| `required_ic_lookup_columns` | Columns that must exist in the IC Lookup sheet |
| `block_on_duplicate_numbers` | `true` = duplicate INC numbers are blocking errors |
| `block_on_blank_numbers` | `true` = blank INC numbers are blocking errors |

---

## 6. Business Rules Engine

All business rules are implemented in `src/business_rules.py` and are **fully deterministic** — no AI, no heuristics, no randomness.

### Rule 1: IC Determination

```
Input:  ServiceNow IC, Priority, Proposed By
Output: IC value (string or None)

if IC is not blank:
    return IC
elif Proposed By is not blank:
    return Proposed By
else:
    return None
```

### Rule 2: Region Lookup (Non-Destructive)

```
Input:  IC value, IC Lookup dictionary, existing Region
Output: Region value

if IC is found in IC Lookup:
    return IC Lookup[IC]     # Overwrite with authoritative value
else:
    return existing Region   # PRESERVE manual entry (do not erase)
```

This non-destructive behavior was a deliberate design decision for non-technical teams. It prevents chart breakage when IC Lookup is not up-to-date.

### Rule 3: Bank/SVB Determination

```
Input:  Assignment Group
Output: Bank value

if "SVB" in Assignment Group (case-insensitive):
    return "SVB"
else:
    return ""
```

### Rule 4: Duration Calculation

```
Input:  Actual Incident Start, Actual Incident Resolve
Output: Duration in Excel serial days (float)

if both Start and Resolve are valid dates:
    return (Resolve - Start).total_seconds() / 86400.0
else:
    return ""
```

---

## 7. Reconciliation Algorithm

```python
# Pseudocode
master_index = {normalize(row.Number): row for row in master_df}
matched = set()

for sn_row in servicenow_df:
    number = normalize(sn_row.Number)
    if number in master_index:
        # MATCH → update fields + apply business rules
        matched.add(number)
        updated_records.append(merge(master_index[number], sn_row))
    else:
        # NEW → create record + apply business rules
        new_records.append(create_new(sn_row))

for number, master_row in master_index:
    if number not in matched:
        # HISTORICAL → retain unchanged
        historical_records.append(master_row)

output = updated + historical + new  # New records appended at end
```

**Time Complexity:** O(M + S) where M = master records, S = ServiceNow records.

---

## 8. Output Strategy (COM vs openpyxl)

| Scenario | Technology | Reason |
|---|---|---|
| **Main output workbook** | `win32com` (COM) | Preserves Pivot Tables, Charts, Macros, Data Models, Conditional Formatting |
| **Reconciliation report** | `openpyxl` | Safe for brand-new workbooks (no existing pivots/charts to corrupt) |
| **Pivot calculations** | `pandas` | Corporate Trust Center blocks native Excel Pivot refresh; Python crosstab bypasses this |
| **Chart rendering** | `win32com` (COM) | `Shapes.AddChart2()` creates native Excel charts via the COM API |

### Why not openpyxl for everything?

openpyxl has a critical limitation: when it opens and saves an `.xlsm` file that contains Pivot Tables or Charts, it **silently destroys** them. This is a known and documented openpyxl limitation. The COM approach works around this by letting Excel itself handle the file I/O.

---

## 9. Pivot Updater

The Pivot Updater (`src/pivot_updater.py`) completely replaces Excel's native Pivot Table engine with Python-calculated summary tables.

### Why?

Corporate Trust Center policies block `RefreshAll()` on Pivot Tables in automated scripts. The Excel security banner requires manual user approval, which cannot be bypassed in an unattended automation.

### How it works

1. Gets or creates the "Pivots" sheet
2. Clears all cells, shapes (charts), and surviving pivot tables
3. Prepares data: parses Created dates, creates Categorical Month column
4. **Table 1:** `pd.crosstab(Month, Priority)` → writes to row 3+
5. **Table 2:** `pd.crosstab(Month, Region)` filtered to last 3 non-zero months → writes below Table 1
6. **Chart:** `AddChart2(201, 52)` stacked column chart bound to Table 2 data range

### Month handling

Months use `pd.Categorical` with explicit ordering `['Jan', ..., 'Dec']` to ensure correct chronological sorting regardless of data order.

---

## 10. Data Flow Diagram

```
┌──────────────────┐     ┌──────────────────┐     ┌──────────────────┐
│  Master Tracker  │     │ ServiceNow Export │     │   IC Lookup      │
│  (.xlsm/.xlsx)   │     │  (.xlsx/.csv)     │     │   (Excel sheet)  │
└────────┬─────────┘     └────────┬──────────┘     └────────┬─────────┘
         │                        │                          │
         ▼                        ▼                          ▼
    ┌─────────────────────────────────────────────────────────────┐
    │                    DATA LOADER                               │
    │  load_master_tracker()  load_servicenow()  load_ic_lookup() │
    └─────────────────────────────┬───────────────────────────────┘
                                  │
                                  ▼
    ┌─────────────────────────────────────────────────────────────┐
    │                    VALIDATOR                                 │
    │  validate_required_columns()                                 │
    │  validate_number_column()                                    │
    │  validate_ic_lookup()                                        │
    └─────────────────────────────┬───────────────────────────────┘
                                  │
                                  ▼
    ┌─────────────────────────────────────────────────────────────┐
    │                    RECONCILER                                │
    │  reconcile() → ReconciliationResult                         │
    │    ├─ UPDATED (matched by Number, fields merged)            │
    │    ├─ NEW (not in Master, created fresh)                    │
    │    └─ HISTORICAL (not in ServiceNow, retained)              │
    └─────────────────────────────┬───────────────────────────────┘
                                  │
                                  ▼
    ┌─────────────────────────────────────────────────────────────┐
    │                    BUSINESS RULES                            │
    │  apply_all_business_rules()                                  │
    │    ├─ IC Determination (IC → Proposed By → blank)           │
    │    ├─ Region Lookup (non-destructive)                       │
    │    ├─ Bank/SVB Determination                                │
    │    └─ Duration Calculation                                   │
    └─────────────────────────────┬───────────────────────────────┘
                                  │
                                  ▼
    ┌─────────────────────────────────────────────────────────────┐
    │                    OUTPUT WRITER (COM)                       │
    │  write_output_workbook()                                     │
    │    ├─ Copy template → output path                           │
    │    ├─ Open with invisible Excel (COM)                       │
    │    ├─ Write data arrays to Consolidated + FCB sheets        │
    │    ├─ Preserve formulas + table styles                      │
    │    ├─ update_pivots_sheet() → Tables + Chart                │
    │    └─ Save + Close + Cleanup                                │
    └─────────────────────────────┬───────────────────────────────┘
                                  │
                    ┌─────────────┼─────────────┐
                    ▼             ▼             ▼
          ┌──────────────┐ ┌───────────┐ ┌───────────┐
          │ Updated      │ │ Recon     │ │ Audit     │
          │ Master       │ │ Report    │ │ Log       │
          │ (.xlsm)      │ │ (.xlsx)   │ │ (.json)   │
          └──────────────┘ └───────────┘ └───────────┘
```

---

## 11. Key Design Decisions

### 1. COM over openpyxl for output
**Problem:** openpyxl silently destroys Pivot Tables, Charts, and Macros.  
**Decision:** Use `win32com.client.DispatchEx("Excel.Application")` to write data while preserving all Excel internals.  
**Trade-off:** Windows-only; requires Excel installed on the machine running the script.

### 2. Python-calculated pivot tables over Excel native pivots
**Problem:** Corporate Trust Center blocks automated `RefreshAll()` on pivot tables.  
**Decision:** Calculate summary tables in Python (`pd.crosstab`) and write raw values via COM.  
**Trade-off:** Pivot tables are static values, not interactive Excel PivotTable objects.

### 3. Non-destructive Region rule
**Problem:** New ServiceNow ICs not in IC Lookup caused all recent Regions to be blanked, breaking charts.  
**Decision:** If IC Lookup misses, preserve the existing manually-entered Region value.  
**Trade-off:** Region may be stale if the master was wrong; but this is far safer for non-technical users.

### 4. GUI file picker with `"ASK"` sentinel
**Problem:** Non-technical users cannot edit `config.json` file paths.  
**Decision:** Set paths to `"ASK"` in config; the pipeline detects this and opens a native Windows file picker.  
**Trade-off:** Requires a display (not suitable for headless/scheduled runs).

### 5. Idempotent output generation
**Problem:** Running the tool twice should not produce duplicate output files.  
**Decision:** Compare reconciliation results; if zero changes detected, skip all output generation.  
**Trade-off:** None — this is strictly better behavior.

### 6. Excel serial date conversion for COM
**Problem:** COM bridge sometimes marshals Python datetime objects as left-aligned text strings in Excel.  
**Decision:** Convert all dates to Excel serial date floats `(date - 1899-12-30).days` before writing.  
**Trade-off:** Requires explicit `NumberFormat = "yyyy-mm-dd"` on date columns.

---

## 12. Dependencies

### Runtime Dependencies

| Package | Version | Purpose |
|---|---|---|
| `pandas` | ≥1.5.0 | DataFrame operations, crosstab calculations |
| `openpyxl` | ≥3.0.0 | Excel reading (data_loader) and new workbook creation (reports) |
| `pywin32` | ≥300 | COM automation for pivot-safe Excel writes (Windows only) |

### Standard Library (no installation needed)

`argparse`, `json`, `os`, `sys`, `datetime`, `shutil`, `re`, `gc`, `dataclasses`, `tkinter`

### Build Dependencies

| Package | Purpose |
|---|---|
| `pyinstaller` | Compiles Python + dependencies into standalone .exe |

### Install all dependencies

```bash
pip install pandas openpyxl pywin32
```

---

## 13. Building the Standalone Executable

The `build_exe.bat` script automates the PyInstaller packaging process.

### Steps

```batch
@echo off
echo Installing Required Libraries (PyInstaller, Pandas, OpenPyXL)...
python -m pip install pyinstaller pandas openpyxl

echo Building the executable...
python -m PyInstaller --onefile --console --name "FCB_Tracker_Updater" ^
    --hidden-import openpyxl --hidden-import pandas src/main.py
```

### Output

The executable is generated at `dist/FCB_Tracker_Updater.exe` (~300-400 MB).

### Distribution Package

To distribute to end users, provide exactly:

```
FCB Automation Tool/
├── FCB_Tracker_Updater.exe    # The standalone executable
└── config/
    └── config.json            # Runtime configuration
```

### Notes

- The `.exe` bundles Python 3.12, pandas, openpyxl, and all dependencies
- `pywin32` is NOT bundled (it uses the system's COM runtime which is built into Windows)
- The `.exe` is 100% offline — no internet connection required at runtime
- Build must be performed on a Windows machine with Python installed

---

## 14. Testing

### Integration Tests

Located at `tests/test_integration.py`. Run with:

```bash
python -m pytest tests/ -v
```

The test suite uses synthetic test data in `test_data/` (never real corporate data).

### Manual Testing Checklist

| Test Case | Expected Result |
|---|---|
| First run with new ServiceNow data | New records added, output generated |
| Second identical run | "UP TO DATE (NO CHANGES)", no output generated |
| ServiceNow with unknown ICs | IC Lookup Misses reported; existing Regions preserved |
| Missing required column in ServiceNow | Validation error, pipeline aborts before output |
| Duplicate INC numbers in input | Validation error, pipeline aborts |
| Empty ServiceNow file | Zero new/updated records; historical records retained |
| Master with formulas (e.g., VLOOKUP) | Formulas preserved in output (except Duration) |
| Master with Pivot Tables and Charts | Pivots and charts preserved; Pivots sheet rebuilt |

---

## 15. Known Limitations & Future Enhancements

### Current Limitations

| Limitation | Impact | Workaround |
|---|---|---|
| Windows-only (COM requirement) | Cannot run on macOS/Linux | Use a Windows VM or build container |
| Requires Excel installed | COM needs Excel.Application | Install Excel or use Office 365 |
| Not suitable for headless/scheduled runs when paths = "ASK" | GUI picker needs a display | Set explicit paths in config.json |
| Pivot Tables in output are static values | Not interactive like native Excel PivotTables | Acceptable trade-off for automation reliability |
| Large .exe file size (~400 MB) | Slow to distribute initially | Build locally on corporate machine |

### Potential Future Enhancements

- **Power Automate Desktop wrapper** — trigger the .exe via Power Automate for scheduled runs
- **SharePoint integration** — auto-download ServiceNow exports from a SharePoint document library
- **Email notifications** — send reconciliation summary via Outlook after each run
- **Web dashboard** — Flask/Streamlit app for visual reporting
- **Cross-platform support** — replace COM with xlwings or a headless Excel engine
- **Incremental processing** — only process new ServiceNow records since last run timestamp
