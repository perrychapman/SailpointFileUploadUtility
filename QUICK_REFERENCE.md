# SailPoint File Upload Utility - Quick Reference Guide

## 🚀 Quick Start

### First Time Setup
1. **Double-click `LaunchUtility.bat`** (launches the GUI automatically)
2. Configure your settings (tenant, credentials, paths)
3. Click "Save Settings"
4. Click "1. Create/Update Directories"
5. Place files in app folders (under Import directory)
6. Click "2. Run File Upload Process"

**That's it!** For daily use, just double-click `LaunchUtility.bat` whenever you need to run operations.

## 📁 File Locations

| File/Folder | Purpose |
|-------------|---------|
| `SailpointUtilityGUI.ps1` | Main GUI interface |
| `FileUploadScript.ps1` | Headless upload script (for scheduling) |
| `DirectoryCreateScriptv3.ps1` | Creates directory structure |
| `LaunchUtility.bat` | Easy launcher for Windows users |
| `settings.json` | Global configuration (auto-created) |
| `Import/[AppName]/` | Drop files here for processing |
| `Import/[AppName]/config.json` | Per-app configuration |
| `Import/[AppName]/Archive/` | Processed files stored here |
| `Import/[AppName]/Log/` | Per-app log files |
| `ExecutionLog/` | Master execution logs |

## ⚙️ Key Settings (settings.json)

```json
{
  "ParentDirectory": "C:\\Powershell\\FileUploadUtility",
  "AppFolder": "C:\\Powershell\\FileUploadUtility\\Import",
  "FileUploadUtility": "path\\to\\sailpoint-file-upload-utility.jar",
  "tenant": "yourtenantname",
  "ClientID": "your-client-id",
  "ClientSecret": "your-client-secret",
  "enableFileDeletion": false,
  "DaysToKeepFiles": 30,
  "AppFilter": "",
  "isDebug": false,
  "ExecutionLogDir": ".\\ExecutionLog"
}
```

### Important Settings Explained

- **enableFileDeletion**: Set to `true` in production to auto-delete old files
- **DaysToKeepFiles**: How many days to keep archived files before deletion
- **AppFilter**: Process only specific apps (leave blank for all apps)
- **isDebug**: Set to `true` for troubleshooting, `false` for production

## 📝 Per-App Configuration (config.json)

Each app folder has its own `config.json`:

```json
{
  "sourceID": "abc123xyz",
  "isUpload": true,
  "headerRow": 1,
  "trimTopRows": 0,
  "trimBottomRows": 0,
  "schema": "employeeID,firstName,lastName,email",
  "groupTypes": "Role",
  "groupDelimiter": ",",
  "disableField": "status",
  "disableValue": ["Terminated", "Inactive"],
  "sheetNumber": 1
}
```

### Common Configuration Tasks

**Skip rows:** Set `trimTopRows: 2` to skip 2 rows after header

**Select Excel sheet:** Set `sheetNumber: 2` for second sheet

**Enable upload:** Set `isUpload: true` when ready to upload to SailPoint

**Specify columns:** Set `schema: "id,name,email"` to select specific columns

## 🔄 Workflow

### One-Time Setup
1. Run GUI → Configure settings → Create directories

### Daily Operations
1. Place files in `Import/[AppName]/` folders
2. Files are processed automatically (if scheduled) OR click "Run File Upload" in GUI
3. Original files → moved to Archive (deleted after X days if enabled)
4. Check logs for errors

### Automated Schedule
- Set up Windows Task Scheduler to run `FileUploadScript.ps1` daily
- Script runs headless (no GUI needed)
- Logs written to `ExecutionLogDir`

## 🔍 Troubleshooting

### Check Logs
- **Master Log**: `ExecutionLog/ExecutionLog_YYYYMMDD.csv`
- **Per-App Log**: `Import/[AppName]/Log/Log_[AppName]_YYYYMMDD.csv`

### Common Issues

| Issue | Solution |
|-------|----------|
| "ImportExcel module not found" | Run: `Install-Module ImportExcel -Scope CurrentUser` |
| "Java not recognized" | Install JDK 11+ and add to PATH |
| "Config file not found" | Run directory creation first |
| Upload fails | Check Client ID/Secret, verify sourceID in config.json |
| No files processed | Check AppFilter setting, verify files are in correct folder |

### Debug Mode
1. Set `isDebug: true` in settings.json
2. Run the script
3. Check detailed logs and raw API responses

## 📅 Scheduling

### Windows Task Scheduler
```
Program: pwsh.exe
Arguments: -NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\FileUploadScript.ps1"
Trigger: Daily at 2:00 AM
```

### Linux Cron
```
0 2 * * * /usr/bin/pwsh -File /opt/sailpoint/FileUploadScript.ps1
```

## 🎯 Best Practices

✅ **DO:**
- Test with `isDebug: true` first
- Use GUI for initial setup
- Enable `enableFileDeletion` in production
- Monitor logs regularly
- Keep `DaysToKeepFiles` reasonable (30-90 days)
- Set `isUpload: false` initially, then `true` when ready

❌ **DON'T:**
- Leave debug mode on in production
- Manually edit files in Archive folders
- Delete ExecutionLog folder
- Change sourceID after initial setup
- Run multiple instances simultaneously

## 📞 Support

For issues or questions:
1. Check logs in ExecutionLog and per-app Log folders
2. Review this guide and README.md
3. Verify settings.json and config.json are correct
4. Test with a single app first (use AppFilter)

## 🔐 Security Notes

- `settings.json` contains sensitive credentials
- Protect this file with appropriate file system permissions
- Never commit settings.json to version control
- Use a dedicated service account for scheduled tasks
- Rotate Client Secret regularly per security policy
