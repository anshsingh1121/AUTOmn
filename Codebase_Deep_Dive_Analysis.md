# 🔍 FCB Incident Tracker Automation: Deep-Dive Codebase Analysis

This document provides a comprehensive, logical, and structural analysis of every major module in the `FCB_Incident_Tracker_Automation` codebase. It maps the Python implementation directly to the business requirements to ensure absolutely flawless execution.

---

## 1. System Architecture & Entry Point (`src/main.py`)

**Purpose:** The central orchestrator that ties all modules together.
**Logical Flow:**
1. **Argument Parsing:** Checks if the user is running `--dry-run` or `--production`.
2. **Configuration Loading:** Loads `config.json` via `config_loader.py`.
3. **Data Ingestion:** Uses `data_loader.py` to ingest the Master Tracker, the ServiceNow raw export, and the IC Lookup table.
4. **Validation (Pre-Flight):** Validates all input schemas using `validator.py`. If a critical column (like "Number" or "Priority") is missing, it aborts safely without corrupting data.
5. **Reconciliation Engine:** Calls `reconciler.py` to merge the data, applying all business rules from `business_rules.py`.
6. **Output & Archiving:** If in `--production` mode, writes the new updated workbook via `output_writer.py`, placing files securely in dynamic `YYYY/MM_Month` folders.
7. **Reporting & Auditing:** Generates a human-readable Reconciliation Excel Report and a JSON Audit Log.

---

## 2. Configuration Engine (`src/config_loader.py` & `config.json`)

**Purpose:** Extracts all hardcoded variables out of the code, making the script resilient to future column name changes.
**Deep Dive:**
- **`field_mapping`**: This is the heart of the schema translation. It dictates exactly how a column in ServiceNow maps to a column in your Master Tracker (e.g., `"Short Description" -> "Short description"`).
- **`overwrite_if_blank` safety constraint**: Inside the field mapping, every field is paired with `"overwrite_if_blank": false`. This is a strict safety net implemented in `reconciler.py` ensuring that if a ServiceNow export accidentally contains a blank value for a historical incident, the script will *refuse* to overwrite the existing valid data in your master tracker.

---

## 3. Data Processing & Normalization (`src/normalizer.py`)

**Purpose:** Cleans chaotic text data before rules are applied.
**Key Functions:**
- `normalize_whitespace(val)`: Strips trailing spaces, removes double spaces, and safely handles completely blank `NaN` fields.
- `extract_priority_class(priority_str)`: Uses Regex to safely pull the integer out of strings like `"3 - Moderate"`. This prevents crashes if ServiceNow ever changes the spelling of "Moderate".
- `contains_substring(value, target)`: A completely case-insensitive search tool used primarily to hunt for the word `"SVB"` inside the Assignment Group string, regardless of how it's typed.

---

## 4. The Core Business Logic (`src/business_rules.py`)

**Purpose:** Applies the highly specific FCB operational rules.

### Rule A: IC Determination (`determine_ic`)
- **Line-by-line Logic:**
  1. Does `servicenow_ic` contain text? If yes -> Return it.
  2. If blank, does `proposed_by` contain text? If yes -> Return it.
  3. If both are blank -> Return `None` (Blank).

### Rule B: Region Lookup
- **Line-by-line Logic:**
  1. Takes the resolved IC value.
  2. Normalizes the text (case-insensitive).
  3. Checks the `ic_lookup` dictionary (loaded from your lookup sheet).
  4. Returns the matching Region (e.g., "India" or "US"). If no match is found, leaves it blank (and throws a warning in the audit log).

### Rule C: Bank / SVB Detection
- **Line-by-line Logic:**
  1. Examines the `Assignment Group` string.
  2. Uses `contains_substring(..., "svb")`.
  3. If found -> Bank = `"SVB"`. Otherwise -> Bank = `"FCB"`.

### Rule D: Explicit Duration Calculation
- **Line-by-line Logic:**
  1. Safely grabs `Actual Incident Start` and `Actual Incident Resolve`.
  2. Converts both to strict Python Datetime objects using `pd.to_datetime`.
  3. Subtracts them to get a Python `TimeDelta`.
  4. Calculates `duration_td.total_seconds() / 86400.0`. 
  *Why 86400.0?* Because Excel natively stores time dynamically as days (1 Day = 1.0). By writing the pure float value into the cell (and keeping it off the `known_formulas` blocklist), Excel will instantly render it as `[h]:mm:ss` or any duration format applied to that column.

---

## 5. The Merge Engine (`src/reconciler.py`)

**Purpose:** Synthesizes the master baseline with the new ServiceNow export.
**Key Logic:**
- **Matching:** Operates strictly on the `Number` field (e.g., "INC0001234").
- **State Machine Iteration:**
  1. Iterates through the old master rows. If a row exists in ServiceNow, it triggers the update block. If a row *does not* exist in ServiceNow (historical), it triggers the retention block, explicitly retaining the data so historical records are never lost.
  2. Iterates through the ServiceNow rows. If a row was not in the master tracker, it triggers the `create_new_record` block.
- **Output:** Returns a strictly formatted DataFrame containing the complete, historically preserved, newly appended, fully updated dataset.

---

## 6. The Excel Writer (`src/output_writer.py`)

**Purpose:** The most complex file. Handles writing data to `.xlsm` while preserving macros, formatting, Pivot Tables, and Excel Tables.

### The Write Strategy:
1. **Dynamic Copying:** It uses `shutil.copy2` to duplicate your master `.xlsm` file securely into the `archive/` folder, protecting the original.
2. **Sheet Detection:** It iterates through the `target_sheets` array (`"Consolidated"` and `"FCB"`).
3. **Formula Protection (`_detect_formula_patterns`)**: 
   - It physically scans the existing data rows before deleting them.
   - If it detects an Excel formula (e.g., `=A2+B2`), it saves that column's formula.
   - *Exception*: `Duration` is explicitly blacklisted from formula detection here, ensuring Python's calculation overrides legacy formulas.
4. **Header Mapping (`_write_dataframe_to_sheet`)**:
   - Instead of blindly dumping data starting at A1, it reads Row 1. If the `"FCB"` sheet has 15 columns, and the `"Consolidated"` sheet has 20 columns, the script dynamically maps the data strictly to the columns that exist.
5. **The Pivot & Table Auto-Expansion:**
   - **ListObjects (Excel Tables):** Uses `get_column_letter` and `ws.max_row` to stretch the `tbl.ref` boundary of every formatted Excel Table in the sheet to encompass the new data.
   - **Pivot Caches:** Digs into `pivot.cacheDefinition.cacheSource.worksheetSource` and updates the `ref` boundary (e.g., from `A1:R180` -> `A1:R250`).
   - **Refresh Trigger:** Sets `refreshOnLoad = True` on every Pivot cache so the backend auto-syncs upon opening.

---

## Conclusion
The codebase strictly adheres to defensive programming principles. By entirely decoupling configuration from logic, implementing deep validation checks, and relying on `openpyxl`'s low-level XML modification capabilities, the script achieves exactly what Power Automate and Power Query could not: **Fully automated processing without destroying macros or pivot references.**
