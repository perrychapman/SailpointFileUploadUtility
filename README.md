# SailPoint File Upload Utility - README

This document provides comprehensive instructions for setting up, running, and troubleshooting the SailPoint File Upload Utility.  
It covers system requirements, setup steps, configuration details, scheduling options, file processing workflow, and troubleshooting guidance.

---

## 📖 Table of Contents

1. [System Requirements](#system-requirements)
2. [Quick Start Guide](#quick-start-guide)
3. [GUI Management Console](#gui-management-console)
4. [Setup Instructions](#setup-instructions)
5. [Configuration Details](#configuration-details)
6. [Automated Scheduling](#automated-scheduling)
7. [File Processing Workflow](#file-processing-workflow)
8. [Script Functions Overview](#script-functions-overview)
9. [Sample Directory Structure](#sample-directory-structure)
10. [Logging and Troubleshooting](#logging-and-troubleshooting)
11. [Example User List and Configuration](#example-user-list-and-configuration)
12. [Behavioral Logic & Script Defaults](#behavioral-logic--script-defaults)
13. [Changelog](#changelog)

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
- `SailpointUtilityGUI.ps1` (GUI management console - **recommended for initial setup**)
- `FileUploadScript.ps1` (main execution script - can run headless for scheduled tasks)
- `DirectoryCreateScriptv3.ps1` (directory creation script - integrated into GUI)
- `settings.json` (global settings file - created automatically if missing)
- `config.json` (per-app configuration file - created automatically during directory creation)
- SailPoint File Upload Utility JAR file (e.g., `sailpoint-file-upload-utility-4.1.0.jar`)

---

## Quick Start Guide

### For New Users (Recommended)

1. **Double-click `LaunchUtility.bat`** to launch the GUI
   - (Or run `.\SailpointUtilityGUI.ps1` from PowerShell)

2. **Configure Settings**
   - Fill in your SailPoint tenant information
   - Set directory paths
   - Enter Client ID and Client Secret
   - Click "Save Settings"

3. **Create Directory Structure**
   - Click "1. Create/Update Directories"
   - This creates folders for all Delimited File sources in your tenant

4. **Process and Upload Files**
   - Place files in the appropriate app folders
   - Click "2. Run File Upload Process"

### For Experienced Users

1. Edit `settings.json` directly with your configuration
2. Run `.\DirectoryCreateScriptv3.ps1` to create directory structure
3. Run `.\FileUploadScript.ps1` manually or via Task Scheduler

---

## GUI Management Console

The **SailpointUtilityGUI.ps1** provides a user-friendly interface for managing the entire file upload workflow.

### Features

- **Persistent Settings**: All configuration changes are saved to `settings.json` and persist across sessions
- **Directory Creation**: One-click setup of all required folder structures
- **File Upload Processing**: Execute file processing and uploads from the GUI
- **Real-time Logging**: View operation logs in real-time within the interface
- **Quick Access**: Open settings file or Import folder directly from the GUI

### How to Use

1. **Launch the GUI**
   ```powershell
   .\SailpointUtilityGUI.ps1
   ```

2. **Configure Settings** (Left Panel)
   - **Directory Locations**: Set paths for parent directory, app folder, JAR file, and log directory
   - **SailPoint Credentials**: Enter your tenant name, Client ID, and Client Secret
   - **Options**: Configure app filter, file retention days, debug mode, and file deletion

3. **Save Settings**
   - Click "Save Settings" to persist your configuration

4. **Execute Actions** (Right Panel)
   - **Create/Update Directories**: Sets up folder structure for all Delimited File sources
   - **Run File Upload Process**: Processes and uploads files from configured folders
   - **Open settings.json**: Quick access to edit the configuration file
   - **Open Import Folder**: Navigate to the Import directory in Windows Explorer

5. **Monitor Operations** (Bottom Panel)
   - View real-time logs of all operations
   - Logs display timestamps, operation type, and status

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
Then click "2. Run File Upload Process"

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