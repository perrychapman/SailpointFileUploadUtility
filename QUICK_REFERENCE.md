# SailPoint File Upload Utility - Quick Reference Guide

##  Quick Start

### First Time Setup
1. *(Optional)* Run `.\CheckPrerequisites.ps1` to verify your environment
2. Run `.\SailpointUtilityGUI.ps1` to launch the GUI
3. Go to the **Settings** tab  configure credentials and paths, then click **Save Settings**
4. Go to the **App Management** tab  click **App Management**, select apps, click **Apply Changes**
5. Place source files in `Import/[AppName]/` folders
6. Select an app in the dropdown and click **Upload Files**

### Daily Operations
1. Place updated source files in `Import/[AppName]/` folders
2. Launch GUI (`.\SailpointUtilityGUI.ps1`), select the app, click **Upload Files**
3.  OR  let the scheduled Task Scheduler job run automatically

##  File Locations

| File/Folder | Purpose |
|-------------|---------|
| `SailpointUtilityGUI.ps1` | Main GUI interface (tabbed: App Management + Settings) |
| `FileUploadScript.ps1` | Headless upload script (used by GUI and Task Scheduler) |
| `DirectoryCreateScriptv3.ps1` | Standalone directory creation script |
| `CheckPrerequisites.ps1` | Environment validation (PowerShell, Java, JAR, settings) |
| `FileMatcher.ps1` | Utility to match files to app folders by name pattern |
| `settings.json` | Global configuration (auto-created by GUI if missing) |
| `Import/[AppName]/` | Drop source files here for processing |
| `Import/[AppName]/config.json` | Per-app configuration (auto-created) |
| `Import/[AppName]/Archive/` | Processed files stored here |
| `Import/[AppName]/Log/` | Per-app log files (CSV) |
| `ExecutionLog/` | Master execution logs (CSV) |

##  Key Settings (settings.json)

```json
{
  "ParentDirectory": "C:\\Powershell\\FileUploadUtility",
  "AppFolder": "C:\\Powershell\\FileUploadUtility\\Import",
  "FileUploadUtility": "path\\to\\sailpoint-file-upload-utility.jar",
  "ExecutionLogDir": ".\\ExecutionLog",
  "tenant": "yourtenantname",
  "tenantUrl": "",
  "ClientID": "your-client-id",
  "ClientSecret": "your-client-secret",
  "enableFileDeletion": false,
  "DaysToKeepFiles": 30,
  "AppFilter": "",
  "isDebug": false
}
```

### Important Settings Explained

- **tenant**: Your SailPoint tenant name (e.g., `mycompany` from `mycompany.identitynow.com`)
- **tenantUrl**: Use instead of `tenant` for vanity/partner URLs (e.g., `https://partner7354.identitynow-demo.com`). Only one is needed; `tenant` takes priority if both are set.
- **enableFileDeletion**: Set to `true` in production to auto-delete old archived and log files
- **DaysToKeepFiles**: How many days to keep archived files and logs before deletion
- **AppFilter**: Process only a specific app (leave blank for all apps)
- **isDebug**: Set to `true` for troubleshooting; deletes original file even if upload fails

##  Per-App Configuration (config.json)

Each app folder has its own `config.json` (created automatically by the GUI):

```json
{
  "sourceID": "abc123xyz",
  "isUpload": false,
  "headerRow": 1,
  "sheetNumber": 1,
  "trimTopRows": 0,
  "trimBottomRows": 0,
  "trimLeftColumns": 0,
  "trimRightColumns": 0,
  "schema": "",
  "dropColumns": "",
  "columnsToMerge": "",
  "mergedColumnName": "",
  "groupTypes": "",
  "groupDelimiter": ",",
  "disableField": "",
  "disableValue": [""],
  "adminColumnName": "",
  "adminColumnValue": "",
  "booleanColumnList": "",
  "booleanColumnValue": ""
}
```

### Common Configuration Tasks

| Task | Setting | Example |
|------|---------|---------|
| Skip rows after header | `trimTopRows` | `2` |
| Skip rows at bottom | `trimBottomRows` | `1` |
| Remove leftmost columns | `trimLeftColumns` | `1` |
| Remove rightmost columns | `trimRightColumns` | `2` |
| Select Excel sheet | `sheetNumber` | `2` |
| Enable upload to SailPoint | `isUpload` | `true` |
| Specify output columns | `schema` | `"id,name,email"` |
| Remove specific columns | `dropColumns` | `"PhoneNumber,Fax"` |
| Merge columns | `columnsToMerge` + `mergedColumnName` | `"FirstName,LastName"`  `"FullName"` |
| Disable users by status | `disableField` + `disableValue` | `"Status"` + `["Inactive","Terminated"]` |
| Assign admin role | `adminColumnName` + `adminColumnValue` | `"Role"` + `"Admin"` |
| Boolean entitlement flags | `booleanColumnList` + `booleanColumnValue` | `"Ent1,Ent2"` + `"Y"` |

##  GUI Reference

### App Management Tab
| Button | Action |
|--------|--------|
| **App Management** | Opens popup to create/remove app directories and toggle upload |
| **Refresh Apps** | Reloads app folder list from disk |
| **Save Config** | Saves inline config editor changes to `config.json` |
| **Reload Config** | Reloads config from disk (discards unsaved changes) |
| **Upload Files** | Processes and uploads files for the selected app |
| **Upload User List** | Browse for a file and copy it to the selected app folder |
| **Open App Log Folder** | Opens the app's `Log/` folder in Explorer |

### Settings Tab
| Section | Key Fields |
|---------|-----------|
| Directory Locations | Parent Directory, App Folder, JAR path, Execution Log Dir |
| SailPoint Credentials | Tenant, Custom Tenant URL, Client ID, Client Secret |
| Options | Days to Keep Files, Debug Mode, Enable File Deletion |

##  Workflow

### One-Time Setup
1. Launch GUI  Settings tab  configure all fields  Save Settings
2. App Management tab  App Management popup  select apps  Apply Changes

### Per-App Config
1. Select app in dropdown
2. Edit fields in the inline config editor
3. Click **Save Config**

### Automated Schedule
- Task Scheduler runs `FileUploadScript.ps1` headless (no GUI needed)
- `AppFilter` in `settings.json` is temporarily set by the GUI's **Upload Files** button and reset after

##  Troubleshooting

### Check Logs
- **Master Log**: `ExecutionLog/ExecutionLog_YYYYMMDD.csv` (view in Settings tab)
- **Per-App Log**: `Import/[AppName]/Log/Log_[AppName]_YYYYMMDD.csv` (view in App Management tab)

### Common Issues

| Issue | Solution |
|-------|----------|
| "ImportExcel module not found" | Run: `Install-Module ImportExcel -Scope CurrentUser` |
| "Java not recognized" | Install JDK 11+ and add to PATH |
| "Config file not found" | Open App Management, select app, click Apply Changes |
| Upload fails | Check Client ID/Secret in Settings; verify `sourceID` in config |
| No files processed | Verify files are in `Import/[AppName]/`; check `AppFilter` |
| Cannot connect to SailPoint | Verify `tenant` or `tenantUrl` and credentials are correct |
| `.xls` file not processed | Script auto-converts to `.xlsx`; ensure Java is installed |

### Debug Mode
1. Set `isDebug: true` in Settings tab (or `settings.json`)
2. Run the script  a `raw_api_response.json` is also saved in the parent directory
3. Check the Operation Log (click  in the GUI header) and per-app logs

##  Scheduling

### Windows Task Scheduler
```
Program:   pwsh.exe
Arguments: -NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\FileUploadScript.ps1"
Start In:  C:\Path\To\SailpointFileUploadUtility
Trigger:   Daily at 2:00 AM
```

### Linux Cron
```
0 2 * * * /usr/bin/pwsh -NoProfile -ExecutionPolicy Bypass -File /opt/sailpoint/FileUploadScript.ps1
```

##  Best Practices

 **DO:**
- Run `CheckPrerequisites.ps1` before first use
- Test with `isDebug: true` and `isUpload: false` first
- Use the GUI inline config editor rather than editing `config.json` manually
- Enable `enableFileDeletion` in production to manage disk space
- Monitor logs regularly via the GUI's log viewer panels
- Keep `DaysToKeepFiles` between 3090 days
- Use `AppFilter` to test a single app before running all

 **DON'T:**
- Leave `isDebug: true` in production
- Manually edit or move files in `Archive/` folders
- Delete the `ExecutionLog/` folder
- Change `sourceID` after initial setup
- Run multiple instances of `FileUploadScript.ps1` simultaneously
- Commit `settings.json` to version control (it contains credentials)

##  Support

For issues or questions:
1. Run `.\CheckPrerequisites.ps1` to rule out environment issues
2. Check logs in `ExecutionLog/` and per-app `Log/` folders (viewable in GUI)
3. Review this guide and README.md
4. Test with a single app first using `AppFilter`

##  Security Notes

- `settings.json` contains sensitive credentials (Client ID and Client Secret)
- Protect this file with appropriate file system permissions
- Never commit `settings.json` to version control  it is listed in `.gitignore`
- Use a dedicated service account for scheduled tasks
- Rotate Client Secret regularly per your organization's security policy
