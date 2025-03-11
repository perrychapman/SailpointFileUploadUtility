# File Upload Script - ReadMe

This document provides detailed instructions for setting up, running, and troubleshooting the `FileUploadScript.ps1`. It covers system requirements, setup steps, configuration details, scheduling with Windows Task Scheduler, and troubleshooting guidance.

---

## Table of Contents
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

---

## System Requirements

### Hardware
- Windows Operating System (Windows 10 or higher recommended)
- At least 4GB RAM and 1GB free disk space

### Software
- PowerShell 7+
- Java Runtime Environment (JRE) version 11 or later
- [ImportExcel PowerShell Module](https://github.com/dfinke/ImportExcel) (Automatically installed by the script if not already installed)
- Internet connection required for API calls to SailPoint
- Administrator permissions required to run the script

### Required Files
- `settings.json` (global settings file)
- `FileUploadScript.ps1` (main execution script)
- `config.json` (per-app configuration file)

---

## Setup Instructions

1. **Download the Required Files**
   - Ensure `FileUploadScript.ps1`, `settings.json`, and `config.json` are available in the designated processing directory.
   - If these files are missing, manually create `settings.json` and `config.json` based on the provided examples.

2. **Install Required Software**
   - Ensure PowerShell 7+ and Java JRE 11+ are installed.
   - Run the following command in PowerShell to install the ImportExcel module:
     ```powershell
     Install-Module ImportExcel -Scope CurrentUser
     ```

3. **Configure `settings.json`**
   - Modify `settings.json` to match your processing environment.
   - Example `settings.json`:
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
       "AppFilter": ""
     }
     ```

---

## File Processing Workflow

1. **Ensure ImportExcel Module is Available**
2. **Load Master Settings from `settings.json`**
3. **Fetch App Folders to Process**
4. **For Each App Folder:**
   - Load `config.json`.
   - Identify the latest file to process (CSV, TXT, XLS, XLSX).
   - Convert XLS to XLSX if applicable.
   - Import and clean up file data.
   - Process roles and entitlements.
   - Export processed data.
   - Upload to SailPoint if `isUpload` is `true`.
   - Archive original and processed files.
   - Clean up old archived files.
5. **Write Execution Logs**

---

## Script Functions Overview

| Function Name | Purpose |
|--------------|---------|
| `Ensure-ImportExcelModule` | Checks and installs the ImportExcel module if missing. |
| `Load-MasterSettings` | Reads global settings from `settings.json`. |
| `Write-Log` | Logs events, errors, and warnings to execution logs. |
| `Get-FileData` | Reads and processes input files based on config settings. |
| `Trim-Data` | Cleans, merges, and trims data as per `config.json`. |
| `Process-ImportedData` | Assigns roles, entitlements, and disables users based on settings. |
| `Upload-ToSailPoint` | Uses the SailPoint File Upload Utility to send processed files. |
| `Archive-File` | Moves original and processed files to the `Archive` directory. |
| `Remove-OldFiles` | Deletes archived files older than the specified retention period. |
| `Process-FilesInAppFolder` | Main function that processes files in each app folder. |

---

## Sample Directory Structure

```
C:\DataProcessing\
│-- settings.json
│-- ExecutionLog\
│   │-- ExecutionLog_YYYYMMDD.csv
│-- AppFolder1\
│   │-- config.json
│   │-- Log\
│   │   │-- Log_AppFolder1_YYYYMMDD.csv
│   │-- input.csv
│   │-- Archive\
│   │   │-- Processed_YYYYMMDD.csv
│   │   │-- Original_YYYYMMDD.csv
│   │   │-- [sourceID]_upload_YYYYMMDD.csv
│-- FileUploadScript.ps1
```

---

## Logging and Troubleshooting

### Log File Locations
- **Execution Log**: `ExecutionLog\ExecutionLog_YYYYMMDD.csv`
- **App Logs**: `AppFolder\Log\Log_AppName_YYYYMMDD.csv`

### Common Issues & Solutions
| Issue                         | Solution |
|--------------------------------|----------|
| Missing `ImportExcel` module  | Run `Install-Module ImportExcel -Scope CurrentUser` |
| Missing Java                  | Ensure Open JDK 11+ is installed and `java` is in `PATH`. |
| Config or settings JSON error | Validate JSON using an online tool or PowerShell (`ConvertFrom-Json`). |
| Script exits unexpectedly     | Check log files for error messages. |
| Files not uploading           | Ensure API credentials are correct and service is accessible. |

---

## Example User List and Configuration

### Sample `users.csv`
```
FirstName,LastName,Email,Role,Status,Group
John,Doe,john.doe@example.com,User,Active,HR
Jane,Smith,jane.smith@example.com,Admin,Active,IT
Bob,Brown,bob.brown@example.com,User,Inactive,Finance
```

### Corresponding `config.json`
```json
{
    "sourceID": "78910",
    "disableField": "Status",
    "disableValue": ["Inactive"],
    "groupTypes": "Group",
    "groupDelimiter": ",",
    "isUpload": true,
    "headerRow": 1,
    "dropColumns": "Email",
    "adminColumnName": "Role",
    "adminColumnValue": "Admin"
}
```

### Expected Output File
```
FirstName,LastName,Role,IIQDisabled,Group
John,Doe,User,false,HR
Jane,Smith,Admin,false,IT
Bob,Brown,User,true,Finance
```

---

This documentation provides an overview of the script's functionality and structure. For additional help, refer to the log files or test execution manually before automating with Task Scheduler.

