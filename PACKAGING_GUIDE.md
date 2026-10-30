# CMAA Standalone Packaging Guide (--onefile)
## Corporate Production-Ready Architecture

This guide provides end-to-end instructions for compiling the **Change Management Audit Automation (CMAA)** project into a single, portable, platform-independent executable using PyInstaller.

---

## 1. How the Standalone `--onefile` Architecture Works

When compiled with `CMAA.spec`, PyInstaller packs the entire Python runtime, all third-party libraries (`openpyxl`, `requests`, `certifi`, `urllib3`, `rapidfuzz`, etc.), audit questions rules, and the master template (`QAR Report.xlsx`) into a **single binary executable**:

- **Windows:** `dist/CMAA_Audit.exe`
- **Linux / macOS:** `dist/CMAA_Audit`

---

## 2. Production Credential Management in `constants.py`

In enterprise corporate environments, handling credentials securely across compiled executables is critical:

### A. Three Flexible Credential Modes Supported:
1. **Bundled Mode (Standalone Zero-Dependency):**
   - Place your corporate credentials directly into `constants.py` (`USERNAME = "..."`, `PASSWORD = "..."`, `INSTANCE_URL = "..."`) before running `build_cmaa_exe.bat`.
   - The compiled `CMAA_Audit.exe` will have the credentials baked into the binary. You can distribute the single `.exe` file anywhere—no extra files needed.
2. **Hot-Override Mode (No Recompilation on Password Rotation):**
   - If your service account password changes or you switch between Test and Production instances, you **do NOT need to recompile**.
   - Simply place a `constants.py` file with the new credentials in the same directory as `CMAA_Audit.exe` (or current working directory).
   - `CMAA_Audit.exe` dynamically detects external `constants.py` at startup, loads the overrides, and logs `[CONFIG] Loaded external configuration override from: constants.py`.
3. **Environment Variable Mode:**
   - Standard variables (`SN_USERNAME`, `SN_PASSWORD`, `SN_INSTANCE_URL`, `SN_VERIFY_SSL`) are supported and override defaults.

### B. Defense-in-Depth Leak Protection:
- **Auth Header Sanitization:** `api_logger.py` intercepts all HTTP calls and masks Basic Auth headers and `auth` credentials as `("username", "********")`.
- **URL Sanitization:** Embedded credentials in URLs (e.g. `https://user:pass@host`) are automatically redacted.
- **Deep Payload Scrubbing:** Any occurrence of `constants.PASSWORD` appearing in response payloads, error traces, or diagnostic dumps is scrubbed and replaced with `********` before writing to `api_logs/`.

---

## 3. Handling Codebase Modifications in Corporate Environments

If you modify any business logic or configuration files on your corporate machine:
- `constants.py` (URLs, credentials, keywords, exception templates)
- `questions_data.py` (Q1 through Q19 compliance evaluation rules)
- `sc_sheet_data.py` (Standard Change Template governance)
- `excel_handler.py` (Excel formatting, colors, column definitions)
- `servicenow_api.py` (ServiceNow query parameters, pagination)
- `QAR Report.xlsx` (Excel template structure)

### The 5-Step Corporate Build System:
Running `build_cmaa_exe.bat` automatically handles the full lifecycle:
1. **Dependency Verification:** Verifies Python 3.9+ and auto-installs requirements from `requirements-cmaa.txt`.
2. **Clean Cache & Bytecode:** Purges old `build/`, `dist/`, and all recursive `__pycache__` folders across the project so stale `.pyc` files are never packaged.
3. **Pre-Build Integrity Validation:** Runs `python validate_reconstruction.py` offline tests (regexes, CLI argument parser, date ranges, Excel loading, audit logic, security masking). If any syntax or logic error was introduced, the build halts immediately with descriptive error output to prevent building a broken `.exe`.
4. **Clean Standalone Compilation:** Invokes `PyInstaller CMAA.spec --clean --noconfirm` to compile the single-file binary.
5. **Post-Build Smoke Test:** Automatically runs `dist\CMAA_Audit.exe --help` to confirm the binary extracts and boots cleanly on the target OS.

---

## 4. Complete List of Files to Copy to Corporate System

To build and run this application on your corporate system, copy the following files:

| # | File Name | Purpose |
|:---:|---|---|
| **1** | `populate_cr_audit.py` | Main CLI and interactive application entry point with `sys._MEIPASS` template resolver |
| **2** | `constants.py` | ServiceNow endpoints, table constants, credentials, and runtime override engine |
| **3** | `questions_data.py` | Core compliance logic for audit questions Q1 through Q19 |
| **4** | `servicenow_api.py` | ServiceNow REST API client and data assembly |
| **5** | `sc_sheet_data.py` | Standard Change Template governance rules |
| **6** | `excel_handler.py` | Excel template population, cell formatting, and color fills |
| **7** | `api_logger.py` | Intercepts HTTP requests, masks passwords, and writes structured JSON logs |
| **8** | `QAR Report.xlsx` | Master structural Excel audit report template |
| **9** | `CMAA.spec` | PyInstaller specification file for `--onefile` standalone compilation |
| **10** | `requirements-cmaa.txt` | Python build & runtime dependencies |
| **11** | `build_cmaa_exe.bat` | 1-click Windows corporate build script with validation and verification |
| **12** | `build_cmaa.sh` | 1-click Linux/macOS corporate build script |
| **13** | `changes.txt` | Sample change-number list for testing |
| **14** | `validate_reconstruction.py`| Offline unit test to verify logic and security before compiling |
| **15** | `PACKAGING_GUIDE.md` | Complete documentation |

---

## 5. How to Build on Corporate System (Windows)

### Option A: 1-Click Batch Build
Open Command Prompt in your project directory and run:
```cmd
build_cmaa_exe.bat
```

### Option B: Manual Command Line Build
```cmd
python -m pip install -r requirements-cmaa.txt
python validate_reconstruction.py
python -m PyInstaller CMAA.spec --clean --noconfirm
```

Compiled Output:
```text
dist\CMAA_Audit.exe
```

---

## 6. How to Build on Linux / macOS
```bash
chmod +x build_cmaa.sh
./build_cmaa.sh
```
Binary output: `dist/CMAA_Audit`

---

## 7. How to Run the Compiled Executable

### Method 1: Double-Click (Interactive Console Mode)
Double-click `CMAA_Audit.exe` directly in Windows Explorer. You will see:
```text
=================================================================
      Change Management Audit Automation (CMAA) - Standalone
=================================================================

Select an Audit Execution Mode:
  [1] Audit specific Change Request(s) (e.g. CHG6892954 or changes.txt)
  [2] Audit by Date Range with Sampling (e.g. 2026-10-01 to 2026-10-31)
  [3] Exit

Enter choice [1/2/3]:
```

### Method 2: Command Prompt / Terminal (CLI Mode)

**Specific Change Numbers:**
```cmd
dist\CMAA_Audit.exe --change-nums CHG6892954 CHG6957477
```

**Using a text file of change numbers:**
```cmd
dist\CMAA_Audit.exe --change-nums changes.txt
```

**Date Range Audit (Random 20% Sampling):**
```cmd
dist\CMAA_Audit.exe --start-date 2026-10-01 --end-date 2026-10-31 --percentage 20 --tester "Auditor Name"
```

**Custom Output File:**
```cmd
dist\CMAA_Audit.exe --change-nums changes.txt --output "report\QAR_Report_Final.xlsx"
```
