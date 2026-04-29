# SailPoint File Upload Utility - README

This document provides comprehensive instructions for setting up, running, and troubleshooting the SailPoint File Upload Utility.  
It covers system requirements, setup steps, configuration details, scheduling options, file processing workflow, and troubleshooting guidance.

---

##  Table of Contents

- [SailPoint File Upload Utility - README](#sailpoint-file-upload-utility---readme)
  - [Table of Contents](#table-of-contents)
  - [System Requirements](#system-requirements)
    - [Hardware](#hardware)
    - [Software](#software)
    - [Required Files](#required-files)
  - [Quick Start Guide](#quick-start-guide)
    - [For New Users (Recommended)](#for-new-users-recommended)
    - [For Experienced Users](#for-experienced-users)
  - [GUI Management Console](#gui-management-console)
    - [Features](#features)
    - [How to Use](#how-to-use)
      - [Tab 1  App Management](#tab-1--app-management)
      - [Tab 2  Settings](#tab-2--settings)
  - [Setup Instructions](#setup-instructions)
  - [Automated Scheduling](#automated-scheduling)
    - [Windows Task Scheduler Setup](#windows-task-scheduler-setup)
    - [Linux/Unix Cron Job Setup (for SailPoint IQ Service on Linux)](#linuxunix-cron-job-setup-for-sailpoint-iq-service-on-linux)
    - [Best Practices for Automated Scheduling](#best-practices-for-automated-scheduling)
  - [Running the Script](#running-the-script)
    - [Manual Execution](#manual-execution)
    - [Via GUI](#via-gui)
    - [Scheduled Execution](#scheduled-execution)
  - [File Processing Workflow](#file-processing-workflow)
  - [Script Functions Overview](#script-functions-overview)
  - [Sample Directory Structure](#sample-directory-structure)
  - [Logging and Troubleshooting](#logging-and-troubleshooting)
    - [Log File Locations](#log-file-locations)
    - [Common Issues \& Solutions](#common-issues--solutions)
  - [Example User List and Configuration](#example-user-list-and-configuration)
    - [Sample `users.csv`](#sample-userscsv)
    - [Sample `config.json`](#sample-configjson)
  - [Configuration Parameters](#configuration-parameters)
  - [Behavioral Logic \& Script Defaults](#behavioral-logic--script-defaults)
  - [Changelog](#changelog)

---

## System Requirements

### Hardware
- Windows Operating System (Windows 10 or higher recommended) or Windows Server
- At least 4GB RAM and 1GB free disk space

### Software
- **PowerShell 7+** (required)
- **Java Runtime Environment (JRE) version 11 or later** (required for file upload utility)
- **[ImportExcel PowerShell Module](https://github.com/dfinke/ImportExcel)** (automatically installed by the script if not already installed)
- Internet connection (required for initial API calls and ImportExcel installation)
- Administrator permissions may be required for initial setup

### Required Files
- `SailpointUtilityGUI.ps1` (GUI management console  **recommended for all setup and daily use**)
- `FileUploadScript.ps1` (main execution script  runs headless for scheduled tasks)
- `DirectoryCreateScriptv3.ps1` (standalone directory creation script  also integrated into the GUI)
- `CheckPrerequisites.ps1` (optional  validates environment before first run)
- `FileMatcher.ps1` (utility  matches files to app folders by name pattern)
- `settings.json` (global settings file  created automatically by the GUI if missing)
- `config.json` (per-app configuration file  created automatically during directory creation)
- SailPoint File Upload Utility JAR file (e.g., `sailpoint-file-upload-utility-4.1.0.jar`)

---

## Quick Start Guide

### For New Users (Recommended)

1. **Verify prerequisites** (optional but recommended):
   ```powershell
   .\CheckPrerequisites.ps1
   ```

2. **Launch the GUI**:
   ```powershell
   .\SailpointUtilityGUI.ps1
   ```

3. **Configure Settings** (Settings tab)
   - Fill in your SailPoint tenant name (or Custom Tenant URL for vanity domains)
   - Set directory paths (Parent Directory, App Folder, JAR path, Execution Log Directory)
   - Enter Client ID and Client Secret
   - Click **Save Settings**

4. **Create Directory Structure** (App Management tab)
   - Click **App Management**
   - A popup lists all Delimited File sources from your tenant
   - Check the apps you want to manage; toggle **Enable Upload** as desired
   - Click **Apply Changes**  folders and `config.json` files are created automatically

5. **Process and Upload Files**
   - Place source files in the appropriate `Import/[AppName]/` folders
   - Select an app in the dropdown on the App Management tab
   - Click **Upload Files** to process and upload for that app

### For Experienced Users

1. Edit `settings.json` directly with your configuration
2. Run `.\DirectoryCreateScriptv3.ps1` to create directory structure
3. Run `.\FileUploadScript.ps1` manually or via Task Scheduler

---

## GUI Management Console

The **SailpointUtilityGUI.ps1** provides a tabbed management interface for the entire file upload workflow.

### Features

- **Tabbed Interface**: Separate tabs for App Management and Settings keep the workspace organized
- **Persistent Settings**: All configuration is saved to `settings.json` and loaded on startup
- **App Management Popup**: View all Delimited File sources from SailPoint, create/remove directories, and toggle upload enabled per app in one dialog
- **Inline Config Editor**: Select any app from the dropdown to view and edit its `config.json` directly in the GUI  no text editor required
- **Per-App Log Viewer**: See the most recent log entries for the selected app in real time
- **Execution Log Viewer**: Browse and view historical execution log files from a dropdown
- **Upload User List**: Browse to copy a source file directly into an app folder
- **Operation Log**: Click ** View Operation Log** in the header to see a full session log

### How to Use

1. **Launch the GUI**
   ```powershell
   .\SailpointUtilityGUI.ps1
   ```

#### Tab 1  App Management

| Control | Description |
|---------|-------------|
| **App Dropdown** | Select an app to load its `config.json` and most recent log |
| **Refresh Apps** | Reload the list of app folders from disk |
| **App Management** | Open the App Management popup to create/delete directories and set upload flags |
| **Create New Source** | Wizard to create a new Delimited File source in SailPoint ISC and set up its local folder structure |
| **Save Config** | Save edits made in the inline config editor back to the app's `config.json` |
| **Reload Config** | Discard unsaved edits and reload the config from disk |
| **Upload Files** | Run `FileUploadScript.ps1` scoped to the selected app (processes & uploads) |
| **Process Only** | Process files for the selected app without uploading (ignores `isUpload` setting) |
| **Upload User List** | Browse for a source file and copy it into the selected app's folder |
| **Upload Schema** | 2-step wizard to upload an account schema to the selected source in SailPoint ISC |
| **Reset Source** | Cascade-reset the selected source in SailPoint ISC: clears entitlements → accounts → correlation config → account schema |
| **Open App Log Folder** | Open the app's `Log` directory in Windows Explorer |
| **App Logs panel** | Displays the most recent log file for the selected app |

**App Management Popup columns:**
- ** (checkbox)**: Include this app  checked apps have directories created; unchecked apps have directories removed
- **App Name**: Source name from SailPoint
- **Directory Status**: Shows ` Exists` if the folder already exists, or `Not Created`
- **Enable Upload**: Toggles `isUpload` in the app's `config.json`

#### Tab 2  Settings

| Section | Fields |
|---------|--------|
| **Directory Locations** | Parent Directory, App Folder, File Upload Utility JAR, Execution Log Directory |
| **SailPoint Client Credentials** | Tenant, Custom Tenant URL (vanity URLs), Client ID, Client Secret |
| **Options** | Days to Keep Files, Debug Mode, Enable File Deletion |
| **Save Settings** | Persists all changes to `settings.json` |
| **Execution Log Viewer** | Select an execution log from the dropdown to view it in the right panel |

> **Tenant vs Custom Tenant URL**: Use **Tenant** for standard IdentityNow tenants (e.g., `mycompany` from `mycompany.identitynow.com`). Use **Custom Tenant URL** for vanity/partner domains (e.g., `https://partner7354.identitynow-demo.com`). Only one is required.

---

## Setup Instructions

1. **Check Prerequisites** (recommended on first run)
   ```powershell
   .\CheckPrerequisites.ps1
   ```
   This verifies PowerShell 7+, Java 11+, ImportExcel, the JAR file path, and required scripts.

2. **Install Required Software**
   - Install PowerShell 7+ and Java JRE 11+.
   - Run this command in PowerShell to install the ImportExcel module (the main script attempts this automatically on first run):
   ```powershell
   Install-Module ImportExcel -Scope CurrentUser
   ```

3. **Configure `settings.json`**
   - Edit `settings.json` to match your environment, or use the GUI Settings tab:
```json
{
  "ParentDirectory": "C:\\DataProcessing",
  "AppFolder": "C:\\DataProcessing\\Import",
  "FileUploadUtility": "C:\\Tools\\sailpoint-file-upload-utility-4.1.0.jar",
  "ClientID": "YourClientID",
  "ClientSecret": "YourClientSecret",
  "tenant": "YourTenantName",
  "tenantUrl": "",
  "enableFileDeletion": true,
  "DaysToKeepFiles": 30,
  "AppFilter": "",
  "ExecutionLogDir": "C:\\DataProcessing\\ExecutionLog",
  "isDebug": false
}
```
   > **Note:** Set either `tenant` (e.g. `mycompany`) **or** `tenantUrl` (e.g. `https://partner7354.identitynow-demo.com`). If `tenant` is set, it takes priority.

---

## Automated Scheduling

The **FileUploadScript.ps1** can run headless (without GUI) for automated scheduling on SailPoint IQ Service servers or any Windows server.

### Windows Task Scheduler Setup

1. **Open Task Scheduler**
   - Press `Win + R`, type `taskschd.msc`, and press Enter

2. **Create a New Task**
   - Click "Create Task" (not "Create Basic Task" for more options)
   - **General Tab:**
     - Name: "SailPoint File Upload"
     - Description: "Automated SailPoint file processing and upload"
     - Select "Run whether user is logged on or not"
     - Check "Run with highest privileges" (if needed)

3. **Triggers Tab**
   - Click "New..."
   - Choose your schedule (Daily, Weekly, etc.)
   - Example: Daily at 2:00 AM
   - Click "OK"

4. **Actions Tab**
   - Click "New..."
   - Action: "Start a program"
   - Program/script: `pwsh.exe` (or full path: `C:\Program Files\PowerShell\7\pwsh.exe`)
   - Add arguments:
     ```
     -NoProfile -ExecutionPolicy Bypass -File "C:\Powershell\SailpointFileUploadUtility\FileUploadScript.ps1"
     ```
   > **Note:** `FileUploadScript.ps1` reads `settings.json` from the directory it is run from. Set the "Start in" field to the script's folder, or use an absolute path to `settings.json`.
   - Start in (optional): `C:\Powershell\SailpointFileUploadUtility`
   - Click "OK"

5. **Conditions Tab** (Optional)
   - Uncheck "Start the task only if the computer is on AC power" if needed

6. **Settings Tab**
   - Check "Run task as soon as possible after a scheduled start is missed"
   - Check "If the task fails, restart every: 10 minutes"
   - Click "OK"

7. **Enter Credentials**
   - Enter the username and password for the account that will run the task
   - Click "OK"

8. **Test the Task**
   - Right-click on your task and select "Run"
   - Check the execution logs in `ExecutionLogDir` to verify it ran successfully

### Linux/Unix Cron Job Setup (for SailPoint IQ Service on Linux)

If your SailPoint IQ Service runs on Linux with PowerShell Core installed:

1. **Edit crontab**
   ```bash
   crontab -e
   ```

2. **Add schedule** (example: daily at 2:00 AM)
   ```
   0 2 * * * /usr/bin/pwsh -NoProfile -ExecutionPolicy Bypass -File /opt/sailpoint/FileUploadScript.ps1 >> /var/log/sailpoint-upload.log 2>&1
   ```

3. **Save and exit**

4. **Verify cron job**
   ```bash
   crontab -l
   ```

### Best Practices for Automated Scheduling

- **Test manually first**: Always test the FileUploadScript.ps1 manually before scheduling
- **Monitor logs**: Regularly check the ExecutionLog directory for errors
- **Set appropriate file retention**: Use `DaysToKeepFiles` to prevent disk space issues
- **Enable file deletion**: Set `enableFileDeletion: true` in production environments
- **Use AppFilter**: Process specific apps if needed to reduce execution time
- **Debug mode off**: Set `isDebug: false` for production to avoid keeping debug files

---

## Running the Script

### Manual Execution

Run manually in PowerShell:
```powershell
cd C:\Powershell\SailpointFileUploadUtility
.\FileUploadScript.ps1
```

### Via GUI

Use the GUI Management Console for one-time or testing:
```powershell
.\SailpointUtilityGUI.ps1
```
On the **App Management** tab, select the app from the dropdown and click **Upload Files**.

### Scheduled Execution

The script runs completely headless (no GUI) when executed by Task Scheduler or cron.
All operations are logged to files in the `ExecutionLogDir`.

Logs and processed files will appear in the configured folders.

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
| `Ensure-ImportExcelModule` | Ensures ImportExcel is installed, auto-installs if missing |
| `Load-MasterSettings` | Loads and validates `settings.json` |
| `Get-BaseApiUrl` | Constructs the SailPoint API base URL from `tenant` or `tenantUrl` |
| `Write-Log` | Logs events, warnings, and errors to CSV log files |
| `Get-FileData` | Reads and imports CSV, TXT, XLS, or XLSX files |
| `Trim-Data` | Trims rows/columns, merges columns, drops unwanted columns |
| `Process-ImportedData` | Adds `IIQDisabled` flag, assigns roles, expands group entitlements |
| `Upload-ToSailPoint` | Invokes the SailPoint File Upload Utility JAR to upload processed data |
| `Archive-File` | Moves processed files to the app's `Archive` subfolder |
| `Remove-OldFiles` | Deletes archived files older than `DaysToKeepFiles` |
| `Remove-OldLogFiles` | Deletes log files older than `DaysToKeepFiles` |
| `Remove-OriginalFile` | Deletes the original source file from the app folder after processing |
| `Process-FilesInAppFolder` | Orchestrates the full processing pipeline for a single app folder |

---

## Sample Directory Structure

```
C:\DataProcessing\
 settings.json
 ExecutionLog\
    ExecutionLog_YYYYMMDD.csv
 Apps\
    App1\
       config.json
       Log\
          Log_App1_YYYYMMDD.csv
       input.csv
       Archive\
          Original_YYYYMMDD.csv
          Processed_YYYYMMDD.csv
          [sourceID]_upload_YYYYMMDD.csv
 FileUploadScript.ps1
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
| **`mergeDelimiter`** | Separator inserted between values when merging columns. If blank, values are concatenated with no separator. | String | `" "` (space), `"-"`, `" \| "` |
| **`adminColumnName`** | Column used to determine admin users. | Column Name from CSV/Excel | `"Role"` |
| **`adminColumnValue`** | Value in `adminColumnName` that indicates an admin user. | String | `"Admin"`, `"SuperUser"` |
| **`sheetNumber`** | Defines which worksheet to process in Excel files (1-based index). | Integer (`>= 1`) | `1`, `2`, `3` |
| **`schema`** | Comma-separated list of column names to include in output. If blank, all columns are included. | Comma-separated column names or `""` | `"employeeID,firstName,lastName,email"`, `""` |
| **`booleanColumnList`** | Comma-separated columns containing boolean entitlement flags. Combined with `booleanColumnValue` to produce a `Role` column. | Comma-separated column names or `""` | `"Entitlement1,Entitlement2"` |
| **`booleanColumnValue`** | The value in `booleanColumnList` columns that indicates the entitlement is active. | String | `"Y"`, `"Yes"`, `"1"`, `"TRUE"` |

---

## Behavioral Logic & Script Defaults

 If `groupTypes` is blank, assigns `Role = User` by default.  
 Users with `adminColumnName` matching `adminColumnValue` are assigned `Role = Admin`.  
 `.xls` files are automatically converted to `.xlsx` before processing.  
 Boolean entitlement columns (e.g., Y/N flags) are converted to a `Role` column via `booleanColumnList` + `booleanColumnValue`.  
 If `isDebug=true`, the original source file is deleted even if the upload fails.  
 If upload succeeds, the original source file is always deleted from the app folder.  
 If upload fails and `isDebug=false`, the original source file is **retained** for retry.  
 When multiple files are present in an app folder, only the **most recently modified** file is processed.  
 Archive, log, and execution log file cleanup is governed by `enableFileDeletion` and `DaysToKeepFiles`.  
 API URL is constructed from `tenant` (standard tenants) or derived from `tenantUrl` (vanity/partner domains).

---

## Changelog

| Date | Change |
|------|--------|
| 11/18/2024 | Added admin role assignment with `adminColumnName` & `adminColumnValue` |
| 11/21/2024 | Added boolean entitlement column (`booleanColumnList` / `booleanColumnValue`) processing |
| 12/17/2024 | Enhanced `AppFilter` logic & offline `ImportExcel` handling |
| 1/13/2025 | Added automatic deletion of old archive files via `enableFileDeletion` + `DaysToKeepFiles` |
| 2/21/2025 | Added `.xls`  `.xlsx` auto-conversion and `sheetNumber` selection |
| 7/10/2025 | Added original file and log file cleanup logic; added `tenantUrl` (vanity URL) support |
| 2/18/2026 | GUI redesigned with tabbed interface (App Management + Settings tabs); added inline `config.json` editor grouped by category; added App Management popup with directory status, per-app Enable Upload toggle, and Select/Deselect All; added per-app log viewer and execution log viewer with date dropdown; added Upload Files (single-app) and Upload User List buttons; added Operation Log popup in header |
| 4/27/2026 | Fixed Reset Source API step order to correct cascade (entitlements → accounts → correlation → schema), preventing HTTP 400 "referenced by other configuration" errors when the schema was reset before entitlements |
| 4/27/2026 | Fixed `groupDelimiter` save corruption — delimiter values containing commas (e.g., `","`) are now stored as plain strings instead of being converted to arrays, restoring row-splitting behavior |
| 4/27/2026 | Documented Create New Source wizard, Upload Schema wizard, Reset Source button, Process Only button, and `mergeDelimiter` config field; updated `.gitignore` to exclude CSV and Excel data files |

---

For additional help, review the log files and run manually before scheduling.