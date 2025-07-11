# File Upload Script - README

This document provides detailed instructions for setting up, running, and troubleshooting the `FileUploadScript.ps1`.  
It covers system requirements, setup steps, configuration details, scheduling with Windows Task Scheduler, file processing workflow, example configurations, and troubleshooting guidance.

---

## 📖 Table of Contents

1. [System Requirements](#system-requirements)
2. [Setup Instructions](#setup-instructions)
3. [Configuration Details](#configuration-details)
4. [Running the Script](#running-the-script)
5. [Task Scheduler Setup](#task-scheduler-setup)
6. [File Processing Workflow](#file-processing-workflow)
7. [Script Functions Overview](#script-functions-overview)
8. [Sample Directory Structure](#sample-directory-structure)
9. [Logging and Troubleshooting](#logging-and-troubleshooting)
10. [Example User List and Configuration](#example-user-list-and-configuration)
11. [Behavioral Logic & Script Defaults](#behavioral-logic--script-defaults)
12. [Changelog](#changelog)

---

## System Requirements

### Hardware
- Windows Operating System (Windows 10 or higher recommended)
- At least 4GB RAM and 1GB free disk space

### Software
- PowerShell 7+
- Java Runtime Environment (JRE) version 11 or later
- [ImportExcel PowerShell Module](https://github.com/dfinke/ImportExcel) (automatically installed by the script if not already installed)
- Internet connection (required for API calls and ImportExcel installation)
- Administrator permissions required to run the script

### Required Files
- `settings.json` (global settings file)
- `FileUploadScript.ps1` (main execution script)
- `config.json` (per-app configuration file)

---

## Setup Instructions

1. **Download the Required Files**
   - Ensure `FileUploadScript.ps1`, `settings.json`, and per-app `config.json` are present in their respective directories.
   - If these files are missing, manually create `settings.json` and `config.json` using the examples below.

2. **Install Required Software**
   - Install PowerShell 7+ and Java JRE 11+.
   - Run this command in PowerShell to install the ImportExcel module (optional if no internet — script attempts it automatically):
```powershell
Install-Module ImportExcel -Scope CurrentUser
```

3. **Configure `settings.json`**
   - Edit `settings.json` to match your environment:
```json
{
  "ParentDirectory": "C:\\DataProcessing",
  "AppFolder": "C:\\DataProcessing\\Apps",
  "FileUploadUtility": "C:\\Tools\\sailpoint-file-upload.jar",
  "ClientID": "YourClientID",
  "ClientSecret": "YourClientSecret",
  "tenant": "YourTenantURL",
  "enableFileDeletion": true,
  "DaysToKeepFiles": 30,
  "AppFilter": "",
  "ExecutionLogDir": "C:\\DataProcessing\\ExecutionLog",
  "isDebug": false
}
```

---

## Running the Script

Run manually in PowerShell:
```powershell
pwsh.exe .\FileUploadScript.ps1
```

Logs and processed files will appear in the configured folders.

---

## Task Scheduler Setup

1. Open **Task Scheduler** → Create Basic Task.
2. Select a trigger (daily, weekly, etc.).
3. Action → Start a Program → `pwsh.exe`
4. Add arguments:
```
-File "C:\Path\To\FileUploadScript.ps1"
```
5. Save and test.

---

## File Processing Workflow

1. **Ensure ImportExcel Module is Available**
   - Checks if `ImportExcel` is installed.
   - If missing, attempts install; if failed, logs warning and continues if already pre-installed.

2. **Load Master Settings**
   - Reads `settings.json`.
   - Validates required parameters like `AppFolder`, `ClientID`, `ClientSecret`, and `tenant`.

3. **Identify App Folders**
   - Scans `AppFolder` for subdirectories.
   - If `AppFilter` is set, only matching folders are processed.

4. **For Each App Folder:**
   - **Load `config.json`**
     - Reads app-specific settings.
     - If missing/invalid, logs error and skips.

   - **Find Latest File to Process**
     - Looks for the most recent `.csv`, `.txt`, `.xls`, or `.xlsx`.
     - If `.xls`, converts to `.xlsx` (with `.xls` removed).

   - **Import & Clean Data**
     - Reads data starting at `headerRow`, from specified `sheetNumber` if Excel.
     - Trims `trimTopRows`, `trimBottomRows`, `trimLeftColumns`, `trimRightColumns`.
     - Merges `columnsToMerge` into `mergedColumnName`.
     - Drops columns listed in `dropColumns`.

   - **Process Roles & Entitlements**
     - Assigns roles (`Admin`, `User`) based on `adminColumnName` and `adminColumnValue`.
     - Marks users as disabled if `disableField` matches `disableValue`.
     - If `groupTypes` is set, assigns entitlements based on those columns.
     - If `booleanColumnList` is set, combines columns with `booleanColumnValue` into `Role`.

   - **Export Data**
     - Saves processed intermediate file.
     - Saves upload-ready file as `[sourceID]_upload_YYYYMMDD.csv`.

   - **Upload to SailPoint**
     - If `isUpload=true`, uploads via SailPoint File Upload Utility.
     - Logs upload status.

   - **Archive & Cleanup**
     - Moves original & processed files to `Archive` folder.
     - Deletes original file if upload successful or `isDebug=true`.
     - Deletes archived & log files older than `DaysToKeepFiles` if `enableFileDeletion=true`.

5. **Write Execution Logs**
   - Logs script execution summary including number of apps processed, skipped, and errors.

---

## Script Functions Overview

| Function Name | Purpose |
|---------------|---------|
| `Ensure-ImportExcelModule` | Ensures ImportExcel is installed |
| `Load-MasterSettings` | Loads `settings.json` |
| `Write-Log` | Logs events, warnings, errors |
| `Get-FileData` | Reads file data (CSV, TXT, XLS/XLSX) |
| `Trim-Data` | Trims/merges/drops columns |
| `Process-ImportedData` | Adds roles, disabled flags, entitlements |
| `Upload-ToSailPoint` | Uploads to SailPoint |
| `Archive-File` | Moves files to Archive |
| `Remove-OldFiles` | Deletes old archived files |
| `Remove-OldLogFiles` | Deletes old log files |
| `Remove-OriginalFile` | Deletes original file |
| `Process-FilesInAppFolder` | Main processing for app folder |

---

## Sample Directory Structure

```
C:\DataProcessing\
├── settings.json
├── ExecutionLog\
│   └── ExecutionLog_YYYYMMDD.csv
├── Apps\
│   ├── App1\
│   │   ├── config.json
│   │   ├── Log\
│   │   │   └── Log_App1_YYYYMMDD.csv
│   │   ├── input.csv
│   │   ├── Archive\
│   │   │   ├── Original_YYYYMMDD.csv
│   │   │   ├── Processed_YYYYMMDD.csv
│   │   │   └── [sourceID]_upload_YYYYMMDD.csv
├── FileUploadScript.ps1
```

---

## Logging and Troubleshooting

### Log File Locations
- **Execution Log**: `ExecutionLog\ExecutionLog_YYYYMMDD.csv`
- **App Logs**: `Apps/<AppName>/Log/Log_<AppName>_YYYYMMDD.csv`

### Common Issues & Solutions

| Issue | Solution |
|-------|----------|
| Missing ImportExcel | Run `Install-Module ImportExcel -Scope CurrentUser` |
| Missing Java | Install JDK 11+ and ensure `java` in PATH |
| Invalid JSON | Validate using online tool or PowerShell `ConvertFrom-Json` |
| Script exits unexpectedly | Check logs for errors |
| Upload fails | Verify credentials & service availability |
| No files found | Ensure files exist in app folder and `AppFilter` is correct |

---

## Example User List and Configuration

### Sample `users.csv`
```csv
FirstName,LastName,Email,Role,Status,Group
John,Doe,john.doe@example.com,User,Active,HR
Jane,Smith,jane.smith@example.com,Admin,Active,IT
Bob,Brown,bob.brown@example.com,User,Inactive,Finance
```

### Sample `config.json`
```json
{
  "sourceID": "550e8400-e29b-41d4-a716-446655440000",
  "disableField": "Status",
  "disableValue": ["Inactive"],
  "groupTypes": "",
  "groupDelimiter": ",",
  "isUpload": true,
  "headerRow": 2,
  "trimTopRows": 2,
  "trimBottomRows": 1,
  "trimLeftColumns": 1,
  "trimRightColumns": 0,
  "dropColumns": "Email",
  "columnsToMerge": "FirstName,LastName",
  "mergedColumnName": "FullName",
  "adminColumnName": "Role",
  "adminColumnValue": "Admin",
  "sheetNumber": 1,
  "booleanColumnList": "Entitlement1,Entitlement2",
  "booleanColumnValue": "Y"
}
```

---

## Configuration Parameters

| Parameter           | Description | Valid Inputs | Example Values |
|---------------------|-------------|-------------|---------------|
| **`sourceID`** | **UUID value pulled from the SailPoint connection settings URL.** This uniquely identifies the app's integration source. | UUID (String) | `"550e8400-e29b-41d4-a716-446655440000"` |
| **`disableField`** | Column used to determine if a user should be disabled. | Column Name from CSV/Excel | `"Status"`, `"EmployeeState"` |
| **`disableValue`** | Values in `disableField` that indicate an inactive user. | Array of Strings | `["Inactive", "Terminated"]` |
| **`groupTypes`** | Column(s) that contain entitlement or group data. **If left blank, defaults to `Role` column.** | Column Name(s) or `""` | `"Group"`, `"Department"`, `""` |
| **`groupDelimiter`** | Separator used in `groupTypes` if multiple values exist. | String | `","`, `"|"` |
| **`isUpload`** | Whether to upload processed data to SailPoint. | Boolean | `true`, `false` |
| **`headerRow`** | The row where column headers start. | Integer (`>= 1`) | `1`, `2`, `3` |
| **`trimTopRows`** | Number of rows **after header row** to remove. | Integer (`>= 0`) | `0`, `1`, `5` |
| **`trimBottomRows`** | Number of rows at the bottom to remove. | Integer (`>= 0`) | `0`, `2`, `10` |
| **`trimLeftColumns`** | Number of leftmost columns to remove. | Integer (`>= 0`) | `0`, `1`, `3` |
| **`trimRightColumns`** | Number of rightmost columns to remove. | Integer (`>= 0`) | `0`, `1`, `2` |
| **`dropColumns`** | List of columns to remove from processing. | Comma-separated column names | `"Email,PhoneNumber"` |
| **`columnsToMerge`** | Columns that should be merged into a new column. | Comma-separated column names | `"FirstName,LastName"` |
| **`mergedColumnName`** | Name of the new merged column. | String | `"FullName"` |
| **`adminColumnName`** | Column used to determine admin users. | Column Name from CSV/Excel | `"Role"` |
| **`adminColumnValue`** | Value in `adminColumnName` that indicates an admin user. | String | `"Admin"`, `"SuperUser"` |
| **`sheetNumber`** | Defines which worksheet to process in Excel files (1-based index). | Integer (`>= 1`) | `1`, `2`, `3` |

---

## Behavioral Logic & Script Defaults

✅ If `groupTypes` is blank, assigns `Role = User` by default.  
✅ Users with `adminColumnName=AdminColumnValue` are assigned `Role = Admin`.  
✅ `.xls` files are automatically converted to `.xlsx`.  
✅ Boolean entitlement columns (e.g., Y/N) are converted to `Role` field.  
✅ If `isDebug=true`, deletes original file even if upload fails.

---

## Changelog

| Date | Change |
|------|--------|
| 11/18/2024 | Added admin role assignment with `adminColumnName` & `adminColumnValue` |
| 11/21/2024 | Added boolean entitlement column processing |
| 12/17/2024 | Enhanced AppFilter & offline ImportExcel handling |
| 1/13/2025 | Added deletion of old archive files |
| 2/21/2025 | Added `.xls`→`.xlsx` conversion & sheet selection |
| 7/10/2025 | Added original and log file cleanup logic |

---

For additional help, review the log files and run manually before scheduling.