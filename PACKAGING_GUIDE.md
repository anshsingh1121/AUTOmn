# CMAA Standalone Packaging Guide (--onefile)

This guide provides end-to-end instructions for compiling the **Change Management Audit Automation (CMAA)** project into a single, portable, platform-independent executable using PyInstaller.

---

## 1. How the Standalone `--onefile` Architecture Works

When compiled with `CMAA.spec`, PyInstaller packs the entire Python runtime, all third-party libraries (`openpyxl`, `requests`, `rapidfuzz`, etc.), your audit rules, and the master template (`QAR Report.xlsx`) into a **single binary executable**:

- **Windows:** `dist/CMAA_Audit.exe`
- **Linux / macOS:** `dist/CMAA_Audit`

### Key Corporate & Portability Features:
1. **Self-Contained Template:** `QAR Report.xlsx` is bundled inside the executable. When the program runs, it automatically resolves the template from `sys._MEIPASS`. If a user places a custom `QAR Report.xlsx` in their working directory, the application seamlessly uses the local one first.
2. **Dual Execution Modes:**
   - **CLI Mode:** When invoked with arguments (e.g. from scheduled tasks, batch scripts, or CI/CD pipelines), it runs non-interactively.
   - **Interactive Console Mode:** When double-clicked in Windows Explorer or run without arguments, it displays an interactive prompt, guides the user through audit options, and pauses before closing so output can be reviewed.
3. **Safe Credential Handling:** Passwords from `constants.py` are masked automatically as `********` by `api_logger.py` before any log files are written.
4. **Current Working Directory Outputs:** Reports (`report/`) and API payload audits (`api_logs/`) are written directly to the folder where the user executes the command.

---

## 2. Complete List of Files to Copy to Corporate System

To build and run this application on your corporate system, copy the following files:

| # | File Name | Purpose |
|:---:|---|---|
| **1** | `populate_cr_audit.py` | Main CLI and interactive application entry point with `sys._MEIPASS` template resolver |
| **2** | `constants.py` | ServiceNow endpoints, table constants, credentials, and safe logger import |
| **3** | `questions_data.py` | Core compliance logic for audit questions Q1 through Q19 |
| **4** | `servicenow_api.py` | ServiceNow REST API client and data assembly |
| **5** | `sc_sheet_data.py` | Standard Change Template governance rules |
| **6** | `excel_handler.py` | Excel template population, cell formatting, and color fills |
| **7** | `api_logger.py` | Intercepts HTTP requests, masks passwords, and writes structured JSON logs |
| **8** | `QAR Report.xlsx` | Master structural Excel audit report template |
| **9** | `CMAA.spec` | PyInstaller specification file for `--onefile` standalone compilation |
| **10** | `build_cmaa_exe.bat` | 1-click Windows build script |
| **11** | `build_cmaa.sh` | 1-click Linux/macOS build script |
| **12** | `changes.txt` | Sample change-number list for testing |
| **13** | `validate_reconstruction.py`| Offline unit test to verify logic before compiling |

---

## 3. How to Build on Corporate System (Windows)

### Option A: 1-Click Batch Build
Open Command Prompt in your project directory and run:
```cmd
build_cmaa_exe.bat
```

### Option B: Manual Command Line Build
If you already have your corporate virtual environment activated:
```cmd
python -m pip install pyinstaller openpyxl requests rapidfuzz
python -m PyInstaller CMAA.spec --clean --noconfirm
```

Upon completion, your standalone executable will be located at:
```text
dist\CMAA_Audit.exe
```

---

## 4. How to Build on Linux / macOS
```bash
chmod +x build_cmaa.sh
./build_cmaa.sh
```
Binary output: `dist/CMAA_Audit`

---

## 5. How to Run the Compiled Executable

### Method 1: Double-Click (Interactive Mode)
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
CMAA_Audit.exe --change-nums CHG6892954 CHG6957477
```

**Using a text file of change numbers:**
```cmd
CMAA_Audit.exe --change-nums changes.txt
```

**Date Range Audit (Random 20% Sampling):**
```cmd
CMAA_Audit.exe --start-date 2026-10-01 --end-date 2026-10-31 --percentage 20 --tester "Auditor Name"
```

**Custom Output File:**
```cmd
CMAA_Audit.exe --change-nums changes.txt --output "My_Custom_Audit.xlsx"
```
