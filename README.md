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
   - The script checks if the `ImportExcel` module is installed.
   - If not installed, it attempts to install it. If installation fails, a warning is logged.

2. **Load Master Settings from `settings.json`**
   - Reads global settings from `settings.json`.
   - Validates that required parameters like `AppFolder`, `ClientID`, `ClientSecret`, and `tenant` are available.
   - If required parameters are missing, the script logs an error and exits.

3. **Fetch App Folders to Process**
   - Identifies all subdirectories within the designated `AppFolder`.
   - If an `AppFilter` is defined, only folders matching the filter will be processed.
   - If no valid app folders are found, a warning is logged, and processing stops.

4. **For Each App Folder:**
   - **Load `config.json`**
     - Reads application-specific settings for file processing and upload behavior.
     - Logs an error and skips the folder if `config.json` is missing or invalid.

   - **Identify the Latest File to Process**
     - Finds the most recently modified CSV, TXT, XLS, or XLSX file.
     - If multiple files are found, only the most recent one is processed.
     - If no valid file is found, logs an error and skips processing.

   - **Convert XLS to XLSX (if applicable)**
     - If an `.xls` file is found, it is converted to `.xlsx`.
     - The original `.xls` file is removed after conversion.

   - **Import and Clean Up File Data**
     - Reads file data based on format (CSV, TXT, or Excel).
     - Removes extra rows and columns as per `config.json` settings (`trimTopRows`, `trimBottomRows`, etc.).
     - Merges specified columns if `columnsToMerge` is defined.
     - Drops unnecessary columns as per `dropColumns`.
     - If Boolean processing is enabled, it converts entitlement fields into a `Role` column.

   - **Process Roles and Entitlements**
     - Assigns roles (`Admin`, `User`, etc.) based on `adminColumnName` and `adminColumnValue`.
     - Determines disabled users based on `disableField` and `disableValue`.
     - If `groupTypes` is defined, users are assigned to multiple entitlements.

   - **Export Processed Data**
     - Saves the processed file in CSV format.
     - Logs an error if the file cannot be created.

   - **Upload to SailPoint (if `isUpload` is `true`)**
     - Calls the `Upload-ToSailPoint` function to upload the processed file.
     - Logs success or failure messages related to the upload.

   - **Archive Original and Processed Files**
     - Moves both original and processed files to the `Archive` folder.
     - Files are named as `[sourceID]_upload_YYYYMMDD.csv` for tracking.

   - **Clean Up Old Archived Files**
     - If `enableFileDeletion` is `true`, removes archived files older than `DaysToKeepFiles`.
     - Logs deleted files and any errors that occur.

5. **Write Execution Logs**
   - Logs script execution details, errors, and warnings to the `ExecutionLog` directory.
   - Summarizes the number of apps processed, skipped, and any errors encountered.

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
## Sample Config.json File

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
    "sheetNumber": 1
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

✅ **If `groupTypes` is left blank (`""`), the script will create a `Role` column and assign `"User"` as the default value.**  
✅ **If `adminColumnName` is provided, any user with `adminColumnValue` in that column will have `"Admin"` as their Role.**  
✅ **Example:**
   - If **`groupTypes`** is `""` and `adminColumnName: "Role"` with `adminColumnValue: "Admin"`,  
     - Users with `"Admin"` in the `"Role"` column will be assigned `"Admin"`,  
     - All others will be assigned `"User"`.  

---

## Example Input & Expected Output

### **Input File (`users.csv`)**
```csv
FirstName,LastName,Email,Role,Status,Group
John,Doe,john.doe@example.com,User,Active,HR
Jane,Smith,jane.smith@example.com,Admin,Active,IT
Bob,Brown,bob.brown@example.com,User,Inactive,Finance
```

### **Expected Processed Output (`processed.csv`)**
```csv
FirstName,LastName,FullName,IIQDisabled,Role,Group
John,Doe,John Doe,false,User,HR
Jane,Smith,Jane Smith,false,Admin,IT
Bob,Brown,Bob Brown,true,User,Finance
```

---

This documentation provides an overview of the script's functionality and structure. For additional help, refer to the log files or test execution manually before automating with Task Scheduler.

