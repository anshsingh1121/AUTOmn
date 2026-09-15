# 🔍 FCB Incident Tracker Automation: Deep-Dive Codebase Analysis (V2)

This document provides a highly detailed, line-by-line logical analysis and architectural breakdown of the *entire* `FCB_Incident_Tracker_Automation` Python codebase. It has been explicitly updated to reflect the new **Pure-COM Native Architecture**, which completely eliminates `openpyxl` corruption risks.

---

## 1. The Orchestrator (`src/main.py`)
**Purpose:** The central nervous system of the script.
**Line-by-Line Flow:**
1. **Argument Parsing (`argparse`)**: Listens for `--dry-run` or `--production`.
2. **`config_loader.load_config()`**: Bootstraps the application rules from `config.json`.
3. **`data_loader.load_excel_to_dataframe()`**: Reads your Master Tracker, ServiceNow extract, and IC Lookup into Pandas Memory. *Note: `openpyxl` is used here strictly in read-only mode, which is 100% safe.*
4. **`validator.validate_input_data()`**: Acts as a pre-flight checklist. If a required column is missing from the ServiceNow extract, the script intentionally crashes here, protecting your Master Tracker from being overwritten with bad data.
5. **`reconciler.reconcile_data()`**: Hands the datasets to the merge engine.
6. **Output Branching**:
   - If `--dry-run`, only prints stats.
   - If `--production`, it fires the `output_writer.py` engine to write the new files natively and securely.
7. **`audit.save_audit_log()`**: Saves a JSON trace of every single change made.

---

## 2. Configuration & Decoupling (`config.json` & `src/config_loader.py`)
**Purpose:** Removes hardcoded strings from Python so you can change column headers without rewriting code.
**Critical Logics:**
- **`field_mapping`**: Maps `Source (ServiceNow)` -> `Target (Master)`.
- **`overwrite_if_blank = false`**: An absolute safety mechanism. If a ServiceNow ticket is accidentally exported with a blank description, the script is strictly forbidden from overwriting your existing, valid Master Tracker description.

---

## 3. Data Cleansing (`src/normalizer.py`)
**Purpose:** Sanitizes chaotic human-entered data before the rules engine touches it.
**Critical Functions:**
- `normalize_whitespace(val)`: Strips trailing spaces, preventing "John " from failing to match "John".
- `extract_priority_class(priority_str)`: Uses Regex `r'^(\d+)'` to safely pull the integer `3` out of `"3 - Moderate"`. This future-proofs the code if ServiceNow changes "Moderate" to "Med".
- `contains_substring(value, target)`: Used to hunt for `"svb"` inside the Assignment Group. It converts everything to `.lower()`, ensuring "SVB", "svb", and "sVb" all trigger correctly.

---

## 4. The Brain: Business Rules (`src/business_rules.py`)
**Purpose:** Translates FCB's highly specific operational requirements into Python conditionals.

### Rule A: IC Determination (`determine_ic`)
- **Step 1:** `if not is_blank(servicenow_ic): return servicenow_ic` (If IC exists, prioritize it).
- **Step 2:** `if not is_blank(proposed_by): return proposed_by` (If IC is blank, fall back to Proposed By *unconditionally*).
- **Step 3:** `return None` (If both are blank, it remains blank).

### Rule B: Region Lookup
- Takes the resolved IC, normalizes it, and maps it strictly against the internal dictionary built from your `IC Lookup` sheet. Unmatched users are flagged in the Audit log.

### Rule C: Bank detection
- Looks at the `Assignment Group`. If `contains_substring(..., "svb")` is True, Bank = "SVB". Else, Bank = "FCB".

### Rule D: Native Duration Calculation
- Grabs `Actual Incident Start` and `Actual Incident Resolve`.
- Subtracts them: `duration_td = resolve_dt - start_dt`
- **Crucial Math:** `duration_td.total_seconds() / 86400.0`. Excel natively measures time in days. By writing the raw decimal (e.g., `1.5` days) directly into the cell, Excel immediately interprets and formats it perfectly without needing an `=F2-G2` formula.

---

## 5. The Merge Engine (`src/reconciler.py`)
**Purpose:** The surgical tool that merges historical data with new data.
**Logical Flow:**
1. Loads the Master Tracker into a dictionary indexed by `Number` (e.g., INC0001).
2. Iterates over the Master records. If the record exists in the new ServiceNow export, it updates the fields (respecting the `overwrite_if_blank` rule).
3. If the record does *not* exist in the new export, it is actively **retained**, ensuring historical incidents from January are never lost when you download a September report.
4. Iterates over the ServiceNow export to find completely new `Number`s. It creates new records, applies the Business Rules, and appends them to the master list.

---

## 6. The Native COM Output Engine (`src/output_writer.py`) *[MASSIVE UPDATE]*
**Purpose:** Writes the data into Excel using **Native Windows COM**, completely bypassing the `openpyxl` library which is infamous for destroying Pivot Charts and complex data models.

### The Complete Write Strategy (Line-by-Line):
1. **Pristine Backup:** It uses `shutil.copy2` to duplicate your pristine Master Tracker into the new `YYYY/MM_Month` folder. This guarantees the file isn't rebuilt from scratch; it's an exact clone of your original.
2. **Launch Native Excel:** Uses `win32com.client.DispatchEx("Excel.Application")` to launch an invisible robot instance of actual Microsoft Excel on your Windows machine.
3. **Dynamic Formula Preservation:**
   - It iterates through your target sheets (`Consolidated`, `FCB`).
   - It reads Row 2 of your sheet and checks `cell.HasFormula`.
   - If a column has a formula, it saves that exact formula string in Python memory (e.g., `existing_formulas[5] = '=RC[-1]+RC[-2]'`).
4. **Data Clearing:** It safely calls `ws.Range(...).ClearContents()` on rows 2 and below. This deletes the old data but mathematically preserves all of your cell colors, borders, and conditional formatting.
5. **High-Speed 2D Array Injection:**
   - Instead of writing cell-by-cell (which is very slow), Python builds a massive 2D matrix in memory.
   - For every cell, it injects either the pure Python value or the dynamic Formula string.
   - It dumps the entire matrix into Excel instantly using `ws.Range(...).FormulaR1C1 = write_data`.
6. **Dynamic Table & Pivot Expansion:**
   - **Tables:** Loops through `ws.ListObjects` and calls `.Resize(new_range)` natively.
   - **Pivots:** Loops through `wb.PivotCaches()`, intercepts the `SourceData` string, calculates the new maximum rows, and manually stretches the cache boundary natively.
7. **The Ultimate Refresh:** Calls `wb.RefreshAll()` and `excel.CalculateUntilAsyncQueriesDone()`. Because this is driven by the actual Microsoft Excel engine, your Pivot Tables and Pivot Charts are perfectly refreshed and completely protected from corruption.

---

## 7. Audit Logging (`src/audit.py`)
**Purpose:** Absolute transparency.
- Writes a highly detailed `.json` file outlining exactly which `Number`s were added, updated, retained, and which `IC`s failed the Region lookup, allowing for rapid debugging of data discrepancies.

---

## Conclusion
This architecture is now Enterprise-Grade. By isolating the read-phase to Python data structures and handing the write-phase entirely over to the native Microsoft Excel executable via `win32com`, the script guarantees absolute data integrity, historical preservation, and total protection of your complex Pivot Charts.
