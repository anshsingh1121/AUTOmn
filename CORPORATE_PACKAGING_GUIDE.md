# Corporate Deployment & Standalone Application Packaging Guide

This guide details how to build and execute the standalone **IMT Incident Tracker Dashboard** application in an enterprise corporate environment where:
- Direct `git pull` / clone is **prohibited**.
- Files are manually copied to corporate workstations.
- Application startup time must be instantaneous (**`--onedir` mode with `_internal/`**).
- Absolute enterprise security norms apply (zero cloud egress, air-gapped, localhost-only, non-admin execution).
- Complete portability and zero-corruption Excel processing are required.

---

## 1. Quick Step-by-Step for Corporate System

### Step 1: Copy Project Files to Corporate Machine
Copy the entire `FCB_Incident_Tracker_Automation` folder via corporate-approved file transfer (USB, network share, or zip archive) to your corporate drive, for example:
```
C:\Automation\FCB_Incident_Tracker_Automation\
```

### Step 2: Build the Standalone Application
Open a standard Command Prompt (or PowerShell) in that folder and run:
```cmd
build_dashboard_exe.bat
```
This batch script automatically:
1. Verifies local Python (3.9+) and build dependencies.
2. Packages the application using `IMT_Dashboard.spec` in **`--onedir` mode**.
3. Places all Python binaries, packages, and Streamlit assets in `dist\IMT_Dashboard\_internal\`.
4. Copies a customizable `config/config.json` next to `IMT_Dashboard.exe`.
5. Pre-creates the `dist\IMT_Dashboard\output\` folder.

### Step 3: Run the Application
In `dist\IMT_Dashboard\`:
- Simply **double-click `IMT_Dashboard.exe`** (or `Launch_Dashboard.bat`).
- The application will start immediately (~1-2 seconds) and automatically pop open your default web browser to:
  ```
  http://127.0.0.1:8501
  ```
- No terminal commands or browser navigation required!

---

## 2. Directory Architecture & Portability

The packaged application is completely portable. You can move the entire `dist\IMT_Dashboard` directory anywhere on the machine or distribute it to colleagues:

```
dist\IMT_Dashboard\
  ├── IMT_Dashboard.exe             <-- Double-click to launch
  ├── Launch_Dashboard.bat          <-- Optional batch launcher helper
  ├── README_HOW_TO_RUN.txt         <-- Quick reference guide
  ├── config\
  │     └── config.json             <-- External editable settings (override)
  ├── output\                       <-- Default target folder for reconciled files
  └── _internal\                    <-- All libraries, Streamlit, Pandas, Plotly
        ├── app.py
        ├── src\
        └── ...
```

### Key Portability Enhancements:
1. **Dynamic Path Resolution**: No paths are hardcoded to any user profile or machine. Paths dynamically resolve relative to the executable location.
2. **Custom Output Folder**: The sidebar includes a **"Target Output Folder"** field where users can visually verify or change the output directory (e.g. to a SharePoint sync drive or local desktop folder).
3. **One-Click Explorer**: Upon completion, a one-click button **"📂 Open Output Folder in Explorer"** opens Windows File Explorer directly to the output folder.

---

## 3. Solution to Previous Defects

### Defect 1: "Output file directory not getting created / files in wrong directory"
- **Cause**: Previously, file writing was guarded by an overly restrictive condition `if (recon_res.new_records or recon_res.change_details)`. When no new records were added, the output writing was completely skipped and `os.makedirs` was never called. Furthermore, paths were resolved to an obscure internal directory without being surfaced in the UI.
- **Resolution**:
  - The output directory is created unconditionally when reconciliation is executed with save enabled.
  - The reconciled Master Tracker, Reconciliation Audit Report (`.xlsx`), and Audit Log (`.json`) are always generated so the operator has concrete artifacts verifying the run.
  - The exact absolute destination folder is displayed prominently in the Streamlit UI with an "Open Folder in Explorer" button.

### Defect 2: "Corrupted format on output files"
- **Cause**: Previously, when Windows COM encountered an error or pywin32 was absent, the fallback invoked `pd.ExcelWriter(..., engine='openpyxl')`. For `.xlsm` workbooks, this stripped the VBA macros, destroyed all Pivot Tables and Charts, and caused Excel to report file corruption. Additionally, COM shutdown lacked proper reference releases, causing `0x800706be` RPC crashes.
- **Resolution**:
  - Cleaned up COM lifecycle: every worksheet, range, and workbook reference is explicitly closed, garbage-collected, and uninitialized before quitting Excel.
  - Implemented a bulletproof openpyxl fallback that uses `keep_vba=True` when updating `.xlsm` files. It preserves all extra sheets (such as `Pivots` and `IC Lookup`), formatting, tables, and macros.
  - Updated the UI download button so it downloads the actual, fully preserved output workbook rather than an empty pandas slice.

---

## 4. Corporate Security & Compliance Norms

1. **Air-Gap / Offline Execution**:
   - `gatherUsageStats = false` is enforced in Streamlit configuration and environment flags.
   - Server strictly binds to `127.0.0.1` (localhost only).
   - Cross-Origin Resource Sharing (CORS) is disabled and XSRF protection is enabled.
   - Zero network egress or external API calls.
2. **Least-Privilege Execution**:
   - No administrative privileges required. Runs in standard user space.
3. **Data Protection**:
   - Zero hardcoded credentials or API keys.
   - Audit logs contain incident numbers only (`INC*`) with zero PII.
