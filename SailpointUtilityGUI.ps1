# =====================================================================
# PowerShell Script: SailpointUtilityGUI.ps1
# Description: Unified GUI for SailPoint File Upload Utility
#              Manages directory creation and file upload operations
# Requirements: PowerShell 7+
#               ImportExcel module installed
#               Open JDK 11+ installed on host machine
# =====================================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Global variables
$script:logTextBoxInitialized = $false
$script:logTextBox = $null
$script:settingsPath = ".\settings.json"

# ------------------------
# Settings Management
# ------------------------

function Get-BaseApiUrl {
    param (
        [pscustomobject]$Settings
    )

    if (![string]::IsNullOrWhiteSpace($Settings.tenant)) {
        return "https://$($Settings.tenant).api.identitynow.com"
    }

    if (![string]::IsNullOrWhiteSpace($Settings.tenantUrl)) {
        $vanityUrl = $Settings.tenantUrl.TrimEnd('/')
        $uri = [Uri]$vanityUrl
        $hostName = $uri.Host

        if ($hostName -match '^([^.]+)\.(.+)$') {
            $sub = $matches[1]
            $domain = $matches[2]
            return "https://$sub.api.$domain"
        } else {
            Write-Error "Invalid tenantUrl format: $($Settings.tenantUrl)"
            return $null
        }
    }

    Write-Error "Both tenant and tenantUrl are missing in settings.json"
    return $null
}

function Create-DefaultSettings {
    $defaultSettings = @{
        ParentDirectory    = "C:\Powershell\FileUploadUtility"
        AppFolder          = "C:\Powershell\FileUploadUtility\Import"
        SourceDirectory    = ""
        FileUploadUtility  = "C:\Powershell\sailpoint-file-upload-utility-4.1.0.jar"
        ClientSecret       = "Secret"
        ClientID           = "ClientID"
        tenant             = "tenant"
        tenantUrl          = ""
        isDebug            = $true
        AppFilter          = ""
        enableFileDeletion = $false
        DaysToKeepFiles    = 30
        ExecutionLogDir    = ".\ExecutionLog"
    }

    $defaultSettings | ConvertTo-Json -Depth 10 | Set-Content -Path $script:settingsPath -Force
    Write-Log "Default settings.json created."
    return $defaultSettings
}

function Load-Settings {
    if (-not (Test-Path -Path $script:settingsPath)) {
        Write-Log "Settings file not found. Creating default settings..."
        return Create-DefaultSettings
    }

    try {
        $settings = Get-Content -Path $script:settingsPath | ConvertFrom-Json
        
        # Add ExecutionLogDir if missing
        if (-not $settings.PSObject.Properties['ExecutionLogDir']) {
            $settings | Add-Member -MemberType NoteProperty -Name 'ExecutionLogDir' -Value '.\ExecutionLog' -Force
        }
        
        # Add tenantUrl if missing (for vanity URLs)
        if (-not $settings.PSObject.Properties['tenantUrl']) {
            $settings | Add-Member -MemberType NoteProperty -Name 'tenantUrl' -Value '' -Force
        }
        
        # Add SourceDirectory if missing
        if (-not $settings.PSObject.Properties['SourceDirectory']) {
            $settings | Add-Member -MemberType NoteProperty -Name 'SourceDirectory' -Value '' -Force
        }
        
        Write-Log "Settings loaded successfully."
        return $settings
    }
    catch {
        Write-Log "ERROR: Failed to parse settings JSON. Creating default settings..."
        return Create-DefaultSettings
    }
}

function Save-Settings {
    param (
        [PSObject]$settings
    )

    try {
        $settings | ConvertTo-Json -Depth 10 | Set-Content -Path $script:settingsPath -Force
        Write-Log "Settings saved successfully."
        return $true
    }
    catch {
        Write-Log "ERROR: Failed to save settings. $_"
        return $false
    }
}

# ------------------------
# Logging Function
# ------------------------

function Write-Log {
    param (
        [string]$message,
        [string]$type = 'INFO'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Color-code based on log type for better readability
    $typeFormatted = $type.PadRight(7)
    $logMessage = "[$timestamp] [$typeFormatted] $message"

    if ($script:logTextBoxInitialized) {
        # Prepend new log to show most recent first
        $script:operationLogText = "$logMessage`r`n" + $script:operationLogText
    }
    else {
        Write-Host $logMessage
    }
}

# ------------------------
# Directory Creation Logic
# ------------------------

function Run-DirectoryCreation {
    param (
        [PSObject]$settings,
        [array]$selectedSources = @()
    )

    Write-Log "=== Starting Directory Creation Process ===" "INFO"

    $parentDirectory = $settings.ParentDirectory
    $tenant = $settings.tenant
    
    # Get base API URL using the helper function
    $tenantUrl = Get-BaseApiUrl -Settings $settings
    
    if ($null -eq $tenantUrl) {
        Write-Log "ERROR: Cannot determine API URL. Please provide either Tenant or Custom Tenant URL." "ERROR"
        [System.Windows.MessageBox]::Show("Configuration Error: Both 'Tenant' and 'Custom Tenant URL' fields are empty.`n`nPlease provide either:`n- Tenant name (e.g., 'mycompany')`n- OR Custom Tenant URL (e.g., 'https://partner7354.identitynow-demo.com')", "Configuration Required", "OK", "Error")
        return $false
    }
    
    Write-Log "Using API URL: $tenantUrl"
    
    $clientID = $settings.ClientID
    $clientSecret = $settings.ClientSecret
    
    # Use tenant for CSV naming, or extract from vanity URL if tenant is empty
    if ([string]::IsNullOrWhiteSpace($tenant)) {
        # Try to extract tenant name from vanity URL for CSV naming
        $csvTenant = ($settings.tenantUrl -replace 'https?://', '') -split '\.' | Select-Object -First 1
    }
    else {
        $csvTenant = $tenant
    }
    
    $csvPath = Join-Path -Path $parentDirectory -ChildPath "AppList_$csvTenant.csv"

    # Ensure ExecutionLog folder exists
    $executionLogPath = $settings.ExecutionLogDir
    if (-not (Test-Path -Path $executionLogPath)) {
        New-Item -Path $executionLogPath -ItemType Directory | Out-Null
        Write-Log "Created ExecutionLog folder at $executionLogPath"
    }

    # Get OAuth token
    $authHeader = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("${clientID}:${clientSecret}"))

    try {
        Write-Log "Retrieving OAuth token from: $tenantUrl/oauth/token"
        $response = Invoke-RestMethod -Uri "$tenantUrl/oauth/token" -Method Post -Headers @{ Authorization = "Basic $authHeader" } -Body @{
            grant_type = "client_credentials"
        }
        $accessToken = $response.access_token
        Write-Log "OAuth token retrieved successfully."
    }
    catch {
        Write-Log "ERROR: Failed to retrieve OAuth token. $_" "ERROR"
        return $false
    }

    # Set headers for API request
    $headers = @{ 
        Authorization = "Bearer $accessToken"
        Accept        = "application/json"
    }

    # Fetch sources from the API
    try {
        Write-Log "Fetching sources from SailPoint API..."
        $rawSources = Invoke-RestMethod -Uri "$tenantUrl/beta/sources" -Method Get -Headers $headers -ContentType "application/json;charset=utf-8"
        
        if ($settings.isDebug) {
            $rawJsonPath = ".\raw_api_response.json"
            $jsonResponse = $rawSources | ConvertTo-Json -Depth 10
            $jsonResponse | Out-File -FilePath $rawJsonPath -Encoding UTF8
            Write-Log "Raw API response saved to $rawJsonPath"
        }

        $parsedSources = @($rawSources)
        $sourcesList = @()

        # Extract relevant fields
        foreach ($source in $parsedSources) {
            if ($null -ne $source -and $source.PSObject.Properties["id"]) {
                $sourceType = $source.connectorName
                $sourceID = $source.id
                $sourceName = $source.name -replace '/', '-'
                $sourcesList += [PSCustomObject]@{
                    SourceID   = $sourceID
                    SourceName = $sourceName
                    SourceType = $sourceType
                }
            }
        }

        if ($sourcesList.Count -eq 0) {
            Write-Log "No valid sources found." "WARNING"
            return $false
        }

        Write-Log "Successfully retrieved $($sourcesList.Count) sources."

        # Check if AppList.csv already exists
        $existingSourcesList = @()
        if (Test-Path -Path $csvPath) {
            $existingSourcesList = Import-Csv -Path $csvPath
            Write-Log "Loaded existing AppList.csv with $($existingSourcesList.Count) entries."
        }

        # Append only new sources
        $newSourcesList = @()
        foreach ($newSource in $sourcesList) {
            $isDuplicate = $existingSourcesList | Where-Object { $_.SourceID -eq $newSource.SourceID }
            if (-not $isDuplicate) {
                $newSourcesList += $newSource
            }
        }

        # Save updated list
        if ($newSourcesList.Count -gt 0) {
            $combinedList = $existingSourcesList + $newSourcesList
            $combinedList | Export-Csv -Path $csvPath -NoTypeInformation
            Write-Log "AppList_$csvTenant.csv updated with $($newSourcesList.Count) new sources."
        }
        else {
            Write-Log "No new sources to add. AppList_$csvTenant.csv is up to date."
        }
    }
    catch {
        Write-Log "ERROR: Failed to fetch or process sources. $_" "ERROR"
        return $false
    }

    # Import and filter sources for Delimited File type
    $folders = Import-Csv -Path $csvPath | Sort-Object SourceName
    $delimitedFileSources = $folders | Where-Object { $_.SourceType -eq "Delimited File" }
    
    # Further filter by selected sources if provided
    if ($selectedSources.Count -gt 0) {
        $delimitedFileSources = $delimitedFileSources | Where-Object { $selectedSources -contains $_.SourceName }
        Write-Log "Filtered to $($delimitedFileSources.Count) selected sources."
    }
    
    Write-Log "Found $($delimitedFileSources.Count) Delimited File sources."

    # Create Import directory if it doesn't exist
    $importDirectory = Join-Path -Path $parentDirectory -ChildPath "Import"
    if (-not (Test-Path -Path $importDirectory)) {
        New-Item -Path $importDirectory -ItemType Directory | Out-Null
        Write-Log "Created Import folder at $importDirectory"
    }

    # Create folder structure for Delimited File sources
    $createdCount = 0
    $skippedCount = 0
    
    foreach ($folder in $delimitedFileSources) {
        $sourceName = $folder.SourceName
        $sourceID = $folder.SourceID
        $appFolderPath = Join-Path -Path $importDirectory -ChildPath $sourceName

        # Check if the APP folder already exists
        if (-not (Test-Path -Path $appFolderPath)) {
            New-Item -Path $appFolderPath -ItemType Directory | Out-Null
            Write-Log "Created folder: $sourceName"

            # Create subfolders
            $logFolderPath = Join-Path -Path $appFolderPath -ChildPath "Log"
            $archiveFolderPath = Join-Path -Path $appFolderPath -ChildPath "Archive"
            
            New-Item -Path $logFolderPath -ItemType Directory | Out-Null
            New-Item -Path $archiveFolderPath -ItemType Directory | Out-Null
            
            $createdCount++
        }
        else {
            $skippedCount++
        }

        # Create the config.json file only if it doesn't exist
        $configFilePath = Join-Path -Path $appFolderPath -ChildPath "config.json"
        if (-not (Test-Path -Path $configFilePath)) {
            $configContent = @"
{
    "sourceID": "$sourceID",
    "disableField": "",
    "disableValue": [""],
    "groupTypes": "",
    "groupDelimiter": "",
    "isUpload": false,
    "headerRow": 1,
    "trimTopRows": 0,
    "trimBottomRows": 0,
    "trimLeftColumns": 0,
    "trimRightColumns": 0,
    "dropColumns": "",
    "columnsToMerge": "",
    "mergedColumnName": "",
    "mergeDelimiter": "",
    "adminColumnName": "",
    "adminColumnValue": "",
    "schema": "",
    "booleanColumnList": "",
    "booleanColumnValue": "",
    "sheetNumber": 1
}
"@
            Set-Content -Path $configFilePath -Value $configContent
        }
    }

    Write-Log "=== Directory Creation Complete ===" "INFO"
    Write-Log "Created: $createdCount folders | Skipped: $skippedCount existing folders" "INFO"
    
    return $true
}

# ------------------------
# File Match Logic (Single App)
# ------------------------

function Write-ExecutionLog {
    # Writes a CSV row to today's execution log file, matching the format used by FileUploadScript.ps1.
    param (
        [PSObject]$settings,
        [string]$message,
        [string]$logType = 'INFO'
    )
    try {
        $logDir = $settings.ExecutionLogDir
        if ([string]::IsNullOrWhiteSpace($logDir)) { return }
        if (-not (Test-Path -Path $logDir)) { New-Item -Path $logDir -ItemType Directory | Out-Null }
        $logFilePath = Join-Path -Path $logDir -ChildPath ("ExecutionLog_" + (Get-Date -Format 'yyyyMMdd') + ".csv")
        $entry = [PSCustomObject]@{
            'Date/Time'   = (Get-Date -Format 'HH:mm:ss')
            'Log Type'    = $logType
            'Log Details' = $message
        }
        if (-not (Test-Path -Path $logFilePath)) {
            $entry | Export-Csv -Path $logFilePath -NoTypeInformation
        } else {
            $entry | Export-Csv -Path $logFilePath -NoTypeInformation -Append
        }
    }
    catch { <# non-fatal #> }
}

function Get-SourceFilesForApp {
    # Returns files from SourceDirectory whose names match AppName.
    # First tries an exact substring match; if nothing found, falls back to
    # requiring ALL significant tokens (3+ chars) of the app name to appear
    # in the filename (case-insensitive).
    param (
        [string]$SourceDirectory,
        [string]$AppName
    )

    $allFiles = Get-ChildItem -Path $SourceDirectory -File

    # Exact match
    $matched = $allFiles | Where-Object { $_.Name -imatch [regex]::Escape($AppName) }
    if ($matched) { return $matched }

    # Loose match — all significant tokens must appear in the filename
    $tokens = $AppName -split '[\s\-_]+' | Where-Object { $_.Length -ge 3 }
    if ($tokens.Count -gt 0) {
        $matched = $allFiles | Where-Object {
            $name = $_.Name
            ($tokens | Where-Object { $name -imatch [regex]::Escape($_) }).Count -eq $tokens.Count
        }
    }

    return $matched
}

function Run-FileMatchOnly {
    param (
        [PSObject]$settings,
        [string]$appName
    )

    Write-Log "=== Starting File Match for App: $appName ===" "INFO"

    if ([string]::IsNullOrWhiteSpace($settings.SourceDirectory)) {
        Write-Log "ERROR: Source Directory is not configured in Settings." "ERROR"
        [System.Windows.MessageBox]::Show("Source Directory is not configured.`n`nPlease set the 'Source Directory' path in the Settings tab first.", "Configuration Required", "OK", "Warning")
        return $false
    }

    if (-not (Test-Path -Path $settings.SourceDirectory)) {
        Write-Log "ERROR: Source Directory does not exist: $($settings.SourceDirectory)" "ERROR"
        [System.Windows.MessageBox]::Show("Source Directory does not exist:`n$($settings.SourceDirectory)", "Directory Not Found", "OK", "Error")
        return $false
    }

    $appFolderPath = Join-Path -Path $settings.AppFolder -ChildPath $appName

    if (-not (Test-Path -Path $appFolderPath)) {
        Write-Log "ERROR: App folder not found: $appFolderPath" "ERROR"
        return $false
    }

    $matchingFiles = Get-SourceFilesForApp -SourceDirectory $settings.SourceDirectory -AppName $appName

    if (-not $matchingFiles -or $matchingFiles.Count -eq 0) {
        Write-Log "No matching source files found for '$appName' in $($settings.SourceDirectory)." "INFO"
        [System.Windows.MessageBox]::Show("No files matching '$appName' were found in:`n$($settings.SourceDirectory)", "No Files Found", "OK", "Information")
        return $false
    }

    $mostRecentFile = $matchingFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $destinationPath = Join-Path -Path $appFolderPath -ChildPath $mostRecentFile.Name

    try {
        Copy-Item -Path $mostRecentFile.FullName -Destination $destinationPath -Force
        $msg = "Source file match: Copied '$($mostRecentFile.Name)' from $($settings.SourceDirectory) to $appFolderPath."
        Write-Log $msg "INFO"
        Write-ExecutionLog -settings $settings -message $msg -logType 'INFO'
        [System.Windows.MessageBox]::Show("File matched and copied successfully:`n`n$($mostRecentFile.Name)`n`nDestination: $appFolderPath", "Match Complete", "OK", "Information")
        return $true
    }
    catch {
        $msg = "Source file match: Failed to copy '$($mostRecentFile.Name)' to $appFolderPath. $_"
        Write-Log $msg "ERROR"
        Write-ExecutionLog -settings $settings -message $msg -logType 'ERROR'
        return $false
    }
}

function Run-MatchAllFiles {
    param (
        [PSObject]$settings
    )

    Write-Log "=== Starting Match All Files ===" "INFO"
    Write-ExecutionLog -settings $settings -message "Match All Files started from $($settings.SourceDirectory)." -logType 'INFO'
    if ([string]::IsNullOrWhiteSpace($settings.SourceDirectory)) {
        Write-Log "ERROR: Source Directory is not configured in Settings." "ERROR"
        [System.Windows.MessageBox]::Show("Source Directory is not configured.`n`nPlease set the 'Source Directory' path in the Settings tab first.", "Configuration Required", "OK", "Warning")
        return
    }

    if (-not (Test-Path -Path $settings.SourceDirectory)) {
        Write-Log "ERROR: Source Directory does not exist: $($settings.SourceDirectory)" "ERROR"
        [System.Windows.MessageBox]::Show("Source Directory does not exist:`n$($settings.SourceDirectory)", "Directory Not Found", "OK", "Error")
        return
    }

    if ([string]::IsNullOrWhiteSpace($settings.AppFolder)) {
        Write-Log "ERROR: App Folder is not configured in Settings." "ERROR"
        [System.Windows.MessageBox]::Show("App Folder is not configured.`n`nPlease set the 'App Folder' path in the Settings tab first.", "Configuration Required", "OK", "Warning")
        return
    }

    $appFolders = Get-ChildItem -Path $settings.AppFolder -Directory -ErrorAction SilentlyContinue
    if (-not $appFolders -or $appFolders.Count -eq 0) {
        Write-Log "No app folders found in $($settings.AppFolder)." "INFO"
        [System.Windows.MessageBox]::Show("No app folders found in:`n$($settings.AppFolder)", "No Apps Found", "OK", "Information")
        return
    }

    $succeeded = [System.Collections.Generic.List[string]]::new()
    $skipped   = [System.Collections.Generic.List[string]]::new()
    $failed    = [System.Collections.Generic.List[string]]::new()

    foreach ($appFolder in $appFolders) {
        $appName = $appFolder.Name
        $matchingFiles = Get-SourceFilesForApp -SourceDirectory $settings.SourceDirectory -AppName $appName

        if (-not $matchingFiles -or $matchingFiles.Count -eq 0) {
            $msg = "Source file match: No matching files for '$appName'. Skipping."
            Write-Log $msg "INFO"
            Write-ExecutionLog -settings $settings -message $msg -logType 'INFO'
            $skipped.Add($appName)
            continue
        }

        $mostRecentFile = $matchingFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $destinationPath = Join-Path -Path $appFolder.FullName -ChildPath $mostRecentFile.Name

        try {
            Copy-Item -Path $mostRecentFile.FullName -Destination $destinationPath -Force
            $msg = "Source file match: Copied '$($mostRecentFile.Name)' from $($settings.SourceDirectory) to $($appFolder.FullName)."
            Write-Log $msg "INFO"
            Write-ExecutionLog -settings $settings -message $msg -logType 'INFO'
            $succeeded.Add("$appName  ←  $($mostRecentFile.Name)")
        }
        catch {
            $msg = "Source file match: Failed to copy '$($mostRecentFile.Name)' to $($appFolder.FullName). $_"
            Write-Log $msg "ERROR"
            Write-ExecutionLog -settings $settings -message $msg -logType 'ERROR'
            $failed.Add($appName)
        }
    }

    $summary = ""
    if ($succeeded.Count -gt 0) {
        $summary += "Copied ($($succeeded.Count)):`n" + ($succeeded -join "`n") + "`n`n"
    }
    if ($skipped.Count -gt 0) {
        $summary += "No match found ($($skipped.Count)):`n" + ($skipped -join ", ") + "`n`n"
    }
    if ($failed.Count -gt 0) {
        $summary += "Errors ($($failed.Count)):`n" + ($failed -join ", ")
    }

    $icon = if ($failed.Count -gt 0) { "Warning" } else { "Information" }
    [System.Windows.MessageBox]::Show($summary.TrimEnd(), "Match All Files Complete", "OK", $icon)

    $completionMsg = "Match All Files complete — Copied: $($succeeded.Count), Skipped: $($skipped.Count), Failed: $($failed.Count)."
    Write-Log "=== $completionMsg ===" "INFO"
    Write-ExecutionLog -settings $settings -message $completionMsg -logType 'INFO'
}

# ------------------------
# File Upload Logic (Single App)
# ------------------------

function Run-SingleAppUpload {
    param (
        [PSObject]$settings,
        [string]$appName,
        [switch]$ProcessOnly
    )

    if ($ProcessOnly) {
        Write-Log "=== Starting Process Only (No Upload) for App: $appName ===" "INFO"
    } else {
        Write-Log "=== Starting File Upload for App: $appName ===" "INFO"
    }

    try {
        # Set AppFilter to the selected app and save settings so FileUploadScript.ps1 picks it up
        $settings.AppFilter = $appName
        
        # Save the updated settings to config.json
        $saveResult = Save-Settings -settings $settings
        if (-not $saveResult) {
            Write-Log "ERROR: Failed to save settings with AppFilter=$appName" "ERROR"
            return $false
        }
        
        Write-Log "AppFilter set to '$appName' and saved to config.json" "INFO"
        
        # FileUploadScript.ps1 is in the same directory as this GUI script
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "FileUploadScript.ps1"
        
        if (-not (Test-Path -Path $scriptPath)) {
            Write-Log "ERROR: FileUploadScript.ps1 not found at $scriptPath" "ERROR"
            Write-Log "Script directory (PSScriptRoot): $PSScriptRoot" "INFO"
            return $false
        }

        Write-Log "Executing FileUploadScript.ps1 for app '$appName' from: $scriptPath" "INFO"
        
        # Execute the script with ExecutionPolicy Bypass to avoid signing issues
        # The script will read config.json which now has AppFilter set to the selected app
        # -ForceUpload bypasses the app's isUpload setting (GUI always uploads when user clicks Upload Files)
        # -ProcessOnly skips upload regardless of isUpload (for headless/review runs)
        # When run headlessly without either flag, isUpload from config.json is respected
        if ($ProcessOnly) {
            $output = pwsh.exe -ExecutionPolicy Bypass -WorkingDirectory $PSScriptRoot -File $scriptPath -ProcessOnly 2>&1
        } else {
            $output = pwsh.exe -ExecutionPolicy Bypass -WorkingDirectory $PSScriptRoot -File $scriptPath -ForceUpload 2>&1
        }

        # Treat a non-zero exit code as a failure
        $exitCode = $LASTEXITCODE
        
        # Display output in log
        foreach ($line in $output) {
            if ($line -match "ERROR") {
                Write-Log $line "ERROR"
            }
            elseif ($line -match "WARNING") {
                Write-Log $line "WARNING"
            }
            else {
                Write-Log $line "INFO"
            }
        }

        if ($ProcessOnly) {
            Write-Log "=== Process Only Complete for $appName (no upload performed) ===" "INFO"
        } else {
            Write-Log "=== File Upload Process Complete for $appName ===" "INFO"
        }
        
        # Reset AppFilter and save
        $settings.AppFilter = ""
        Save-Settings -settings $settings
        Write-Log "AppFilter reset to empty" "INFO"

        if ($exitCode -ne 0) {
            Write-Log "ERROR: FileUploadScript.ps1 exited with code $exitCode" "ERROR"
            return $false
        }
        
        return $true
    }
    catch {
        Write-Log "ERROR: Failed to execute FileUploadScript.ps1. $_" "ERROR"
        return $false
    }
}

# ------------------------
# GUI Creation
# ------------------------

function Show-MainWindow {
    $settings = Load-Settings

    # Create the Window
    $window = New-Object Windows.Window
    $window.Title = "SailPoint File Upload Utility"
    $window.Width = 1320
    $window.Height = 860
    $window.MinWidth = 1050
    $window.MinHeight = 650
    $window.WindowStartupLocation = 'CenterScreen'
    $window.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
    $window.FontFamily = New-Object Windows.Media.FontFamily("Segoe UI")

    # Create main grid
    $mainGrid = New-Object Windows.Controls.Grid
    $mainGrid.Margin = '0'
    $window.Content = $mainGrid

    # Define rows - header, content (tabs), and status bar
    $headerRow = New-Object Windows.Controls.RowDefinition
    $headerRow.Height = "Auto"
    $contentRow = New-Object Windows.Controls.RowDefinition
    $contentRow.Height = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
    $statusBarRow = New-Object Windows.Controls.RowDefinition
    $statusBarRow.Height = "Auto"
    $mainGrid.RowDefinitions.Add($headerRow)
    $mainGrid.RowDefinitions.Add($contentRow)
    $mainGrid.RowDefinitions.Add($statusBarRow)

    # ===== HEADER SECTION =====
    $headerPanel = New-Object Windows.Controls.Border
    $headerPanel.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
    $headerPanel.Padding = '24,14,24,14'
    [Windows.Controls.Grid]::SetRow($headerPanel, 0)
    $mainGrid.Children.Add($headerPanel)

    $headerDock = New-Object Windows.Controls.DockPanel
    $headerDock.LastChildFill = $true
    $headerPanel.Child = $headerDock

    # Right side: Operation Log button
    $headerButtonsRight = New-Object Windows.Controls.StackPanel
    $headerButtonsRight.Orientation = 'Horizontal'
    $headerButtonsRight.VerticalAlignment = 'Center'
    [Windows.Controls.DockPanel]::SetDock($headerButtonsRight, 'Right')
    $headerDock.Children.Add($headerButtonsRight)

    $viewOperationLogButton = New-Object Windows.Controls.Button
    $viewOperationLogButton.Content = "Operation Log"
    $viewOperationLogButton.Padding = '14,8'
    $viewOperationLogButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
    $viewOperationLogButton.Foreground = [System.Windows.Media.Brushes]::White
    $viewOperationLogButton.FontWeight = 'Normal'
    $viewOperationLogButton.FontSize = 13
    $viewOperationLogButton.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#54C0E8")
    $viewOperationLogButton.BorderThickness = '1'
    $viewOperationLogButton.Cursor = 'Hand'
    $headerButtonsRight.Children.Add($viewOperationLogButton)

    # Left side: title and subtitle
    $headerLeft = New-Object Windows.Controls.StackPanel
    $headerLeft.Orientation = 'Vertical'
    $headerLeft.VerticalAlignment = 'Center'
    $headerDock.Children.Add($headerLeft)

    $titleLabel = New-Object Windows.Controls.Label
    $titleLabel.Content = "SailPoint File Upload Utility"
    $titleLabel.FontSize = 22
    $titleLabel.FontWeight = 'SemiBold'
    $titleLabel.Foreground = [System.Windows.Media.Brushes]::White
    $titleLabel.Padding = '0,0,0,2'
    $headerLeft.Children.Add($titleLabel)

    $subtitleLabel = New-Object Windows.Controls.Label
    $subtitleLabel.Content = "Management Console  -  Configure credentials, manage app directories, and upload files"
    $subtitleLabel.FontSize = 12
    $subtitleLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#54C0E8")
    $subtitleLabel.Padding = '0'
    $headerLeft.Children.Add($subtitleLabel)

    # Initialize log storage (but no visible textbox)
    $script:operationLogText = ""
    $script:logTextBoxInitialized = $true

    # ===== STATUS BAR =====
    $statusBar = New-Object Windows.Controls.Border
    $statusBar.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
    $statusBar.Padding = '18,5'
    [Windows.Controls.Grid]::SetRow($statusBar, 2)
    $mainGrid.Children.Add($statusBar)

    $statusDock = New-Object Windows.Controls.DockPanel
    $statusDock.LastChildFill = $false
    $statusBar.Child = $statusDock

    $script:statusLabel = New-Object Windows.Controls.TextBlock
    $script:statusLabel.Text = "Ready"
    $script:statusLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#54C0E8")
    $script:statusLabel.FontSize = 12
    $script:statusLabel.VerticalAlignment = 'Center'
    [Windows.Controls.DockPanel]::SetDock($script:statusLabel, 'Left')
    $statusDock.Children.Add($script:statusLabel)

    $statusVersion = New-Object Windows.Controls.TextBlock
    $statusVersion.Text = "SailPoint ISC  -  File Upload Utility"
    $statusVersion.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
    $statusVersion.FontSize = 12
    $statusVersion.VerticalAlignment = 'Center'
    [Windows.Controls.DockPanel]::SetDock($statusVersion, 'Right')
    $statusDock.Children.Add($statusVersion)

    # ===== TAB CONTROL =====
    $tabControl = New-Object Windows.Controls.TabControl
    $tabControl.Background = [System.Windows.Media.Brushes]::White
    $tabControl.BorderThickness = '0'
    $tabControl.Margin = '0'
    [Windows.Controls.Grid]::SetRow($tabControl, 1)
    $mainGrid.Children.Add($tabControl)

    # =============================================================================
    # TAB 1: APP MANAGEMENT
    # =============================================================================
    $tab1 = New-Object Windows.Controls.TabItem
    $tab1.Header = "  App Management  "
    $tab1.FontSize = 12
    $tab1.FontWeight = 'SemiBold'
    $tab1.Padding = '4,6'
    $tabControl.Items.Add($tab1)

    $tab1MainPanel = New-Object Windows.Controls.DockPanel
    $tab1MainPanel.Background = [System.Windows.Media.Brushes]::White
    $tab1.Content = $tab1MainPanel

    # HEADER ROW - App selection dropdown and action buttons
    $headerPanel = New-Object Windows.Controls.DockPanel
    $headerPanel.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFFFFF")
    $headerPanel.Margin = '15,15,15,10'
    $headerPanel.LastChildFill = $false
    [Windows.Controls.DockPanel]::SetDock($headerPanel, 'Top')
    $tab1MainPanel.Children.Add($headerPanel)

    # Action buttons on the right
    $actionButtonsPanel = New-Object Windows.Controls.StackPanel
    $actionButtonsPanel.Orientation = 'Horizontal'
    $actionButtonsPanel.Margin = '10,5,0,5'
    [Windows.Controls.DockPanel]::SetDock($actionButtonsPanel, 'Right')
    $headerPanel.Children.Add($actionButtonsPanel)

    # Refresh button
    $refreshAppsButton = New-Object Windows.Controls.Button
    $refreshAppsButton.Content = "Refresh Apps"
    $refreshAppsButton.Padding = '12,6'
    $refreshAppsButton.Margin = '0,0,8,0'
    $refreshAppsButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
    $refreshAppsButton.Foreground = [System.Windows.Media.Brushes]::White
    $refreshAppsButton.FontWeight = 'SemiBold'
    $refreshAppsButton.FontSize = 12
    $refreshAppsButton.BorderThickness = '0'
    $refreshAppsButton.Cursor = 'Hand'
    $actionButtonsPanel.Children.Add($refreshAppsButton)

    # Directory Creation Button
    $createDirsButton = New-Object Windows.Controls.Button
    $createDirsButton.Content = "Manage Apps"
    $createDirsButton.Padding = '12,6'
    $createDirsButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
    $createDirsButton.Foreground = [System.Windows.Media.Brushes]::White
    $createDirsButton.FontWeight = 'SemiBold'
    $createDirsButton.FontSize = 12
    $createDirsButton.BorderThickness = '0'
    $createDirsButton.Cursor = 'Hand'
    $actionButtonsPanel.Children.Add($createDirsButton)

    # New Source Button
    $createSourceButton = New-Object Windows.Controls.Button
    $createSourceButton.Content = "New Source"
    $createSourceButton.Padding = '12,6'
    $createSourceButton.Margin = '8,0,0,0'
    $createSourceButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#2E7D32")
    $createSourceButton.Foreground = [System.Windows.Media.Brushes]::White
    $createSourceButton.FontWeight = 'SemiBold'
    $createSourceButton.FontSize = 12
    $createSourceButton.BorderThickness = '0'
    $createSourceButton.Cursor = 'Hand'
    $actionButtonsPanel.Children.Add($createSourceButton)

    # App selection dropdown on the left
    $appSelectorPanel = New-Object Windows.Controls.StackPanel
    $appSelectorPanel.Orientation = 'Horizontal'
    $appSelectorPanel.Margin = '0,5,0,5'
    [Windows.Controls.DockPanel]::SetDock($appSelectorPanel, 'Left')
    $headerPanel.Children.Add($appSelectorPanel)

    $appLabel = New-Object Windows.Controls.Label
    $appLabel.Content = "Select App:"
    $appLabel.FontWeight = 'SemiBold'
    $appLabel.FontSize = 12
    $appLabel.VerticalAlignment = 'Center'
    $appLabel.Margin = '0,0,8,0'
    $appSelectorPanel.Children.Add($appLabel)

    $appDropdown = New-Object Windows.Controls.ComboBox
    $appDropdown.MinWidth = 250
    $appDropdown.Padding = '8,4'
    $appDropdown.FontSize = 12
    $appDropdown.Background = [System.Windows.Media.Brushes]::White
    $appDropdown.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
    $appDropdown.BorderThickness = '1'
    $appSelectorPanel.Children.Add($appDropdown)
    $script:appDropdown = $appDropdown

    $openAppFolderButton = New-Object Windows.Controls.Button
    $openAppFolderButton.Content = "Open App Folder"
    $openAppFolderButton.Padding = '10,4'
    $openAppFolderButton.Margin = '10,0,0,0'
    $openAppFolderButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
    $openAppFolderButton.Foreground = [System.Windows.Media.Brushes]::White
    $openAppFolderButton.FontWeight = 'SemiBold'
    $openAppFolderButton.FontSize = 12
    $openAppFolderButton.BorderThickness = '0'
    $openAppFolderButton.Cursor = 'Hand'
    $openAppFolderButton.IsEnabled = $false
    $openAppFolderButton.ToolTip = "Open this app's directory in Windows Explorer to access source files, processed output, and logs"
    $appSelectorPanel.Children.Add($openAppFolderButton)

    # CONTENT GRID - 2 columns: Config Editor (left) and Logs (right)
    $contentGrid = New-Object Windows.Controls.Grid
    $contentGrid.Margin = '15,10,15,15'
    $tab1MainPanel.Children.Add($contentGrid)

    $col1 = New-Object Windows.Controls.ColumnDefinition
    $col1.Width = New-Object Windows.GridLength(1, 'Star')
    $contentGrid.ColumnDefinitions.Add($col1)

    $col2 = New-Object Windows.Controls.ColumnDefinition
    $col2.Width = New-Object Windows.GridLength(1.2, 'Star')
    $contentGrid.ColumnDefinitions.Add($col2)

    # LEFT PANEL: Col 0 wrapper (config editor + run panel below)
    $configColumnWrapper = New-Object Windows.Controls.DockPanel
    $configColumnWrapper.Margin = '0,0,8,0'
    $configColumnWrapper.LastChildFill = $true
    [Windows.Controls.Grid]::SetColumn($configColumnWrapper, 0)
    $contentGrid.Children.Add($configColumnWrapper)

    # RUN PANEL - docked to the bottom of the wrapper
    $runPanelBorder = New-Object Windows.Controls.Border
    $runPanelBorder.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F0F4FA")
    $runPanelBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
    $runPanelBorder.BorderThickness = '2'
    $runPanelBorder.Padding = '10,8'
    $runPanelBorder.Margin = '0,6,0,0'
    [Windows.Controls.DockPanel]::SetDock($runPanelBorder, 'Bottom')
    $configColumnWrapper.Children.Add($runPanelBorder)

    $runPanelInner = New-Object Windows.Controls.DockPanel
    $runPanelInner.LastChildFill = $false
    $runPanelBorder.Child = $runPanelInner

    $runPanelLabel = New-Object Windows.Controls.TextBlock
    $runPanelLabel.Text = "Run"
    $runPanelLabel.FontSize = 12
    $runPanelLabel.FontWeight = 'SemiBold'
    $runPanelLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
    $runPanelLabel.VerticalAlignment = 'Center'
    $runPanelLabel.Margin = '0,0,14,0'
    [Windows.Controls.DockPanel]::SetDock($runPanelLabel, 'Left')
    $runPanelInner.Children.Add($runPanelLabel)

    $uploadSchemaButton = New-Object Windows.Controls.Button
    $uploadSchemaButton.Content = "Upload Schema"
    $uploadSchemaButton.Padding = '12,6'
    $uploadSchemaButton.Margin = '0,0,6,0'
    $uploadSchemaButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
    $uploadSchemaButton.Foreground = [System.Windows.Media.Brushes]::White
    $uploadSchemaButton.FontWeight = 'SemiBold'
    $uploadSchemaButton.FontSize = 12
    $uploadSchemaButton.BorderThickness = '0'
    $uploadSchemaButton.Cursor = 'Hand'
    $uploadSchemaButton.IsEnabled = $false
    $uploadSchemaButton.ToolTip = "Build the account schema from the most recent processed file and push it to SailPoint ISC"
    [Windows.Controls.DockPanel]::SetDock($uploadSchemaButton, 'Left')
    $runPanelInner.Children.Add($uploadSchemaButton)

    $processOnlyButton = New-Object Windows.Controls.Button
    $processOnlyButton.Content = "Process Only"
    $processOnlyButton.Padding = '12,6'
    $processOnlyButton.Margin = '0,0,6,0'
    $processOnlyButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
    $processOnlyButton.Foreground = [System.Windows.Media.Brushes]::White
    $processOnlyButton.FontWeight = 'SemiBold'
    $processOnlyButton.FontSize = 12
    $processOnlyButton.BorderThickness = '0'
    $processOnlyButton.Cursor = 'Hand'
    $processOnlyButton.IsEnabled = $false
    $processOnlyButton.ToolTip = "Transform the source file and generate the upload CSV without uploading to SailPoint - useful for previewing output"
    [Windows.Controls.DockPanel]::SetDock($processOnlyButton, 'Left')
    $runPanelInner.Children.Add($processOnlyButton)

    $matchFilesButton = New-Object Windows.Controls.Button
    $matchFilesButton.Content = "Match Files"
    $matchFilesButton.Padding = '12,6'
    $matchFilesButton.Margin = '0,0,6,0'
    $matchFilesButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#5C6BC0")
    $matchFilesButton.Foreground = [System.Windows.Media.Brushes]::White
    $matchFilesButton.FontWeight = 'SemiBold'
    $matchFilesButton.FontSize = 12
    $matchFilesButton.BorderThickness = '0'
    $matchFilesButton.Cursor = 'Hand'
    $matchFilesButton.IsEnabled = $false
    $matchFilesButton.ToolTip = "Copy the most recent matching file from the Source Directory into this app's folder (requires Source Directory to be configured in Settings)"
    [Windows.Controls.DockPanel]::SetDock($matchFilesButton, 'Left')
    $runPanelInner.Children.Add($matchFilesButton)

    $uploadAppButton = New-Object Windows.Controls.Button
    $uploadAppButton.Content = "Upload to ISC"
    $uploadAppButton.Padding = '12,6'
    $uploadAppButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
    $uploadAppButton.Foreground = [System.Windows.Media.Brushes]::White
    $uploadAppButton.FontWeight = 'SemiBold'
    $uploadAppButton.FontSize = 12
    $uploadAppButton.BorderThickness = '0'
    $uploadAppButton.Cursor = 'Hand'
    $uploadAppButton.IsEnabled = $false
    $uploadAppButton.ToolTip = "Process the source file and upload the result to SailPoint ISC (always uploads regardless of isUpload setting)"
    [Windows.Controls.DockPanel]::SetDock($uploadAppButton, 'Left')
    $runPanelInner.Children.Add($uploadAppButton)

    $resetSourceButton = New-Object Windows.Controls.Button
    $resetSourceButton.Content = "Reset Source"
    $resetSourceButton.Padding = '12,6'
    $resetSourceButton.Margin = '6,0,0,0'
    $resetSourceButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#B22222")
    $resetSourceButton.Foreground = [System.Windows.Media.Brushes]::White
    $resetSourceButton.FontWeight = 'SemiBold'
    $resetSourceButton.FontSize = 12
    $resetSourceButton.BorderThickness = '0'
    $resetSourceButton.Cursor = 'Hand'
    $resetSourceButton.IsEnabled = $false
    $resetSourceButton.ToolTip = "DESTRUCTIVE: Remove all accounts, entitlements, and clear the account schema and correlation config for this source in SailPoint ISC"
    [Windows.Controls.DockPanel]::SetDock($resetSourceButton, 'Left')
    $runPanelInner.Children.Add($resetSourceButton)

    # $uploadUserListButton - kept for event-handler compatibility; not shown in UI
    $uploadUserListButton = New-Object Windows.Controls.Button
    $uploadUserListButton.IsEnabled = $false

    # CONFIG EDITOR GROUPBOX - fills remaining space in wrapper
    $configEditorBox = New-Object Windows.Controls.GroupBox
    $configEditorHeaderTB = New-Object Windows.Controls.TextBlock
    $configEditorHeaderTB.Text = "App Configuration"
    $configEditorHeaderTB.FontWeight = 'SemiBold'
    $configEditorHeaderTB.FontSize = 13
    $configEditorHeaderTB.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $configEditorBox.Header = $configEditorHeaderTB
    $configEditorBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
    $configEditorBox.BorderThickness = '2'
    $configEditorBox.Margin = '0'
    $configEditorBox.Padding = '12'
    $configColumnWrapper.Children.Add($configEditorBox)

    $configEditorPanel = New-Object Windows.Controls.DockPanel
    $configEditorBox.Content = $configEditorPanel

    # Save / Reload row inside the config panel
    $configButtonsPanel = New-Object Windows.Controls.StackPanel
    $configButtonsPanel.Orientation = 'Horizontal'
    $configButtonsPanel.Margin = '0,0,0,8'
    [Windows.Controls.DockPanel]::SetDock($configButtonsPanel, 'Top')
    $configEditorPanel.Children.Add($configButtonsPanel)

    $saveConfigButton = New-Object Windows.Controls.Button
    $saveConfigButton.Content = "Save Config"
    $saveConfigButton.Padding = '12,6'
    $saveConfigButton.Margin = '0,0,6,0'
    $saveConfigButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
    $saveConfigButton.Foreground = [System.Windows.Media.Brushes]::White
    $saveConfigButton.FontWeight = 'SemiBold'
    $saveConfigButton.FontSize = 12
    $saveConfigButton.BorderThickness = '0'
    $saveConfigButton.Cursor = 'Hand'
    $saveConfigButton.IsEnabled = $false
    $saveConfigButton.ToolTip = "Save the current configuration changes to this app's config.json"
    $configButtonsPanel.Children.Add($saveConfigButton)

    $reloadConfigButton = New-Object Windows.Controls.Button
    $reloadConfigButton.Content = "Reload Config"
    $reloadConfigButton.Padding = '12,6'
    $reloadConfigButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
    $reloadConfigButton.Foreground = [System.Windows.Media.Brushes]::White
    $reloadConfigButton.FontWeight = 'SemiBold'
    $reloadConfigButton.FontSize = 12
    $reloadConfigButton.BorderThickness = '0'
    $reloadConfigButton.Cursor = 'Hand'
    $reloadConfigButton.IsEnabled = $false
    $reloadConfigButton.ToolTip = "Discard unsaved changes and reload config.json from disk"
    $configButtonsPanel.Children.Add($reloadConfigButton)

    # Config editor - ScrollViewer with form fields
    $configScrollViewer = New-Object Windows.Controls.ScrollViewer
    $configScrollViewer.VerticalScrollBarVisibility = 'Auto'
    $configScrollViewer.HorizontalScrollBarVisibility = 'Disabled'
    $configScrollViewer.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFFFFF")
    $configScrollViewer.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
    $configScrollViewer.BorderThickness = '1'
    $configScrollViewer.Padding = '10'
    $configEditorPanel.Children.Add($configScrollViewer)
    
    $script:configFieldsPanel = New-Object Windows.Controls.StackPanel
    $script:configFieldsPanel.Margin = '0'

    # Empty-state hint shown when no app is selected
    $emptyStatePanel = New-Object Windows.Controls.StackPanel
    $emptyStatePanel.Margin = '10,40,10,0'
    $emptyStatePanel.HorizontalAlignment = 'Center'

    $emptyStateIcon = New-Object Windows.Controls.TextBlock
    $emptyStateIcon.Text = ""
    $emptyStateIcon.FontSize = 40
    $emptyStateIcon.HorizontalAlignment = 'Center'
    $emptyStateIcon.Margin = '0,0,0,12'
    $emptyStatePanel.Children.Add($emptyStateIcon)

    $emptyStateTitle = New-Object Windows.Controls.TextBlock
    $emptyStateTitle.Text = "No app selected"
    $emptyStateTitle.FontSize = 15
    $emptyStateTitle.FontWeight = 'SemiBold'
    $emptyStateTitle.HorizontalAlignment = 'Center'
    $emptyStateTitle.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
    $emptyStateTitle.Margin = '0,0,0,6'
    $emptyStatePanel.Children.Add($emptyStateTitle)

    $emptyStateHint = New-Object Windows.Controls.TextBlock
    $emptyStateHint.Text = "Select an app from the dropdown above to view and edit its configuration."
    $emptyStateHint.FontSize = 12
    $emptyStateHint.TextWrapping = 'Wrap'
    $emptyStateHint.HorizontalAlignment = 'Center'
    $emptyStateHint.TextAlignment = 'Center'
    $emptyStateHint.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
    $emptyStatePanel.Children.Add($emptyStateHint)

    $script:configFieldsPanel.Children.Add($emptyStatePanel)
    $script:emptyStatePanel = $emptyStatePanel
    $configScrollViewer.Content = $script:configFieldsPanel
    
    # Dictionary to store config field controls
    $script:configFields = @{}

    # RIGHT PANEL: Logs
    $logViewerBox = New-Object Windows.Controls.GroupBox
    $logViewerHeaderTB = New-Object Windows.Controls.TextBlock
    $logViewerHeaderTB.Text = "App Logs"
    $logViewerHeaderTB.FontWeight = 'SemiBold'
    $logViewerHeaderTB.FontSize = 13
    $logViewerHeaderTB.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $logViewerBox.Header = $logViewerHeaderTB
    $logViewerBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#D60EB5")
    $logViewerBox.BorderThickness = '2'
    $logViewerBox.Margin = '8,0,0,0'
    $logViewerBox.Padding = '12'
    [Windows.Controls.Grid]::SetColumn($logViewerBox, 1)
    $contentGrid.Children.Add($logViewerBox)

    $logViewerPanel = New-Object Windows.Controls.DockPanel
    $logViewerBox.Content = $logViewerPanel

    # Log viewer buttons at top
    $logButtonsPanel = New-Object Windows.Controls.StackPanel
    $logButtonsPanel.Orientation = 'Horizontal'
    $logButtonsPanel.Margin = '0,0,0,10'
    [Windows.Controls.DockPanel]::SetDock($logButtonsPanel, 'Top')
    $logViewerPanel.Children.Add($logButtonsPanel)

    $openAppLogButton = New-Object Windows.Controls.Button
    $openAppLogButton.Content = "Open Log Folder"
    $openAppLogButton.Padding = '10,6'
    $openAppLogButton.Margin = '0,0,8,0'
    $openAppLogButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
    $openAppLogButton.Foreground = [System.Windows.Media.Brushes]::White
    $openAppLogButton.FontWeight = 'SemiBold'
    $openAppLogButton.FontSize = 12
    $openAppLogButton.BorderThickness = '0'
    $openAppLogButton.Cursor = 'Hand'
    $openAppLogButton.IsEnabled = $false
    $logButtonsPanel.Children.Add($openAppLogButton)

    $appLogFileSelectorLabel = New-Object Windows.Controls.Label
    $appLogFileSelectorLabel.Content = "Log File:"
    $appLogFileSelectorLabel.FontWeight = 'SemiBold'
    $appLogFileSelectorLabel.FontSize = 12
    $appLogFileSelectorLabel.VerticalAlignment = 'Center'
    $appLogFileSelectorLabel.Margin = '12,0,6,0'
    $logButtonsPanel.Children.Add($appLogFileSelectorLabel)

    $script:appLogFileDropdown = New-Object Windows.Controls.ComboBox
    $script:appLogFileDropdown.MinWidth = 230
    $script:appLogFileDropdown.Padding = '6,3'
    $script:appLogFileDropdown.FontSize = 12
    $script:appLogFileDropdown.Background = [System.Windows.Media.Brushes]::White
    $script:appLogFileDropdown.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
    $script:appLogFileDropdown.BorderThickness = '1'
    $script:appLogFileDropdown.IsEnabled = $false
    $logButtonsPanel.Children.Add($script:appLogFileDropdown)

    # Log viewer textbox
    $script:appLogTextBox = New-Object Windows.Controls.TextBox
    $script:appLogTextBox.AcceptsReturn = $true
    $script:appLogTextBox.IsReadOnly = $true
    $script:appLogTextBox.TextWrapping = 'Wrap'
    $script:appLogTextBox.VerticalScrollBarVisibility = 'Auto'
    $script:appLogTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFFFFF")
    $script:appLogTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $script:appLogTextBox.Padding = '8,6'
    $script:appLogTextBox.FontSize = 12
    $script:appLogTextBox.FontFamily = New-Object Windows.Media.FontFamily("Consolas")
    $script:appLogTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
    $script:appLogTextBox.BorderThickness = '1'
    $script:appLogTextBox.Text = "Select an app to view log information..."
    $logViewerPanel.Children.Add($script:appLogTextBox)

    # =============================================================================
    # TAB 2: SETTINGS
    # =============================================================================
    $tab2 = New-Object Windows.Controls.TabItem
    $tab2.Header = "  Settings  "
    $tab2.FontSize = 13
    $tab2.FontWeight = 'SemiBold'
    $tab2.Padding = '4,6'
    $tabControl.Items.Add($tab2)

    $tab2MainPanel = New-Object Windows.Controls.DockPanel
    $tab2MainPanel.Background = [System.Windows.Media.Brushes]::White
    $tab2.Content = $tab2MainPanel

    # CONTENT GRID - 2 columns: Settings (left) and Log Viewer (right)
    $contentGrid2 = New-Object Windows.Controls.Grid
    $contentGrid2.Margin = '15,10,15,15'
    $tab2MainPanel.Children.Add($contentGrid2)

    $col1_2 = New-Object Windows.Controls.ColumnDefinition
    $col1_2.Width = New-Object Windows.GridLength(1, 'Star')
    $contentGrid2.ColumnDefinitions.Add($col1_2)

    $col2_2 = New-Object Windows.Controls.ColumnDefinition
    $col2_2.Width = New-Object Windows.GridLength(1.2, 'Star')
    $contentGrid2.ColumnDefinitions.Add($col2_2)

    # LEFT PANEL: Settings
    $settingsScrollViewer = New-Object Windows.Controls.ScrollViewer
    $settingsScrollViewer.VerticalScrollBarVisibility = 'Auto'
    $settingsScrollViewer.HorizontalScrollBarVisibility = 'Disabled'
    $settingsScrollViewer.Padding = '0,0,8,0'
    [Windows.Controls.Grid]::SetColumn($settingsScrollViewer, 0)
    $contentGrid2.Children.Add($settingsScrollViewer)

    $tab2StackPanel = New-Object Windows.Controls.StackPanel
    $tab2StackPanel.Margin = '0'
    $settingsScrollViewer.Content = $tab2StackPanel

    # Directory Locations GroupBox
    $directoryBox = New-Object Windows.Controls.GroupBox
    $dirBoxHeaderTB = New-Object Windows.Controls.TextBlock
    $dirBoxHeaderTB.Text = "Directory Locations"
    $dirBoxHeaderTB.FontWeight = 'SemiBold'
    $dirBoxHeaderTB.FontSize = 13
    $dirBoxHeaderTB.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $directoryBox.Header = $dirBoxHeaderTB
    $directoryBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
    $directoryBox.BorderThickness = '2'
    $directoryBox.Margin = '0,0,0,20'
    $directoryBox.Padding = '15,10,15,15'
    $directoryPanel = New-Object Windows.Controls.StackPanel
    $directoryPanel.Orientation = 'Vertical'
    $directoryBox.Content = $directoryPanel
    $tab2StackPanel.Children.Add($directoryBox)

    # Create text boxes and browse buttons for directory settings
    $directoryFields = @(
        @{ 
            Label = "Parent Directory"
            Name = "ParentDirectory"
            Value = $settings.ParentDirectory
            Type = "Folder"
            Tooltip = "Root directory where Import folder and AppList CSV will be created. Example: C:\SailPoint"
        },
        @{ 
            Label = "App Folder"
            Name = "AppFolder"
            Value = $settings.AppFolder
            Type = "Folder"
            Tooltip = "Directory containing app-specific folders with config.json files. Example: C:\SailPoint\Import"
        },
        @{ 
            Label = "Source Directory"
            Name = "SourceDirectory"
            Value = $settings.SourceDirectory
            Type = "Folder"
            Tooltip = "Optional directory to pull source files from. Files whose names contain the app folder name will be copied in automatically before processing. Leave blank to disable."
        },
        @{ 
            Label = "File Upload Utility JAR"
            Name = "FileUploadUtility"
            Value = $settings.FileUploadUtility
            Type = "File"
            Tooltip = "Path to SailPoint file-upload-utility.jar (version 4.1.0 or higher). Download from SailPoint GitHub."
        },
        @{ 
            Label = "Execution Log Directory"
            Name = "ExecutionLogDir"
            Value = $settings.ExecutionLogDir
            Type = "Folder"
            Tooltip = "Directory where execution logs will be stored. Example: C:\SailPoint\ExecutionLog"
        }
    )

    $script:textBoxes = @{}

    foreach ($field in $directoryFields) {
        $label = New-Object Windows.Controls.Label
        $label.Content = $field.Label
        $label.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
        $label.FontWeight = 'Normal'
        $label.Margin = '0,8,0,3'
        $label.FontSize = 12
        $label.ToolTip = $field.Tooltip
        $directoryPanel.Children.Add($label)

        # Create horizontal stack for textbox + browse button
        $pathStack = New-Object Windows.Controls.DockPanel
        $pathStack.LastChildFill = $true

        $browseButton = New-Object Windows.Controls.Button
        $browseButton.Content = "..."
        $browseButton.Width = 35
        $browseButton.Height = 28
        $browseButton.Margin = '5,0,0,0'
        $browseButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
        $browseButton.Foreground = [System.Windows.Media.Brushes]::White
        $browseButton.FontSize = 13
        $browseButton.BorderThickness = '0'
        $browseButton.Cursor = 'Hand'
        $browseButton.ToolTip = "Browse for " + $field.Type
        [Windows.Controls.DockPanel]::SetDock($browseButton, [Windows.Controls.Dock]::Right)
        $pathStack.Children.Add($browseButton)

        $textBox = New-Object Windows.Controls.TextBox
        $textBox.Text = $field.Value
        $textBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFFFFF")
        $textBox.Foreground = [System.Windows.Media.Brushes]::Black
        $textBox.Padding = '8,6'
        $textBox.FontSize = 12
        $textBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
        $textBox.ToolTip = $field.Tooltip
        $textBox.BorderThickness = '1'
        $pathStack.Children.Add($textBox)
        
        $script:textBoxes[$field.Name] = $textBox

        # Add browse button click handler - capture field info directly
        $currentField = $field
        $currentTextBox = $textBox
        $browseButton.Add_Click({
            if ($currentField.Type -eq "Folder") {
                $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
                $folderBrowser.Description = "Select " + $currentField.Name
                $folderBrowser.ShowNewFolderButton = $true
                
                # Convert relative path to absolute to avoid FolderBrowserDialog errors
                if ($currentTextBox.Text) {
                    $absolutePath = $currentTextBox.Text
                    if (-not [System.IO.Path]::IsPathRooted($absolutePath)) {
                        $absolutePath = Join-Path $PSScriptRoot $absolutePath
                        $absolutePath = [System.IO.Path]::GetFullPath($absolutePath)
                    }
                    if (Test-Path $absolutePath) {
                        $folderBrowser.SelectedPath = $absolutePath
                    }
                }
                
                $result = $folderBrowser.ShowDialog()
                if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                    $currentTextBox.Text = $folderBrowser.SelectedPath
                }
            }
            else {
                $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog
                $fileBrowser.Title = "Select " + $currentField.Name
                $fileBrowser.Filter = "JAR Files (*.jar)|*.jar|All Files (*.*)|*.*"
                
                if ($currentTextBox.Text -and (Test-Path (Split-Path $currentTextBox.Text -Parent))) {
                    $fileBrowser.InitialDirectory = Split-Path $currentTextBox.Text -Parent
                    $fileBrowser.FileName = Split-Path $currentTextBox.Text -Leaf
                }
                
                $result = $fileBrowser.ShowDialog()
                if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                    $currentTextBox.Text = $fileBrowser.FileName
                }
            }
        }.GetNewClosure())

        $directoryPanel.Children.Add($pathStack)
    }

    # Client Credentials GroupBox
    $credentialsBox = New-Object Windows.Controls.GroupBox
    $credBoxHeaderTB = New-Object Windows.Controls.TextBlock
    $credBoxHeaderTB.Text = "SailPoint Client Credentials"
    $credBoxHeaderTB.FontWeight = 'SemiBold'
    $credBoxHeaderTB.FontSize = 13
    $credBoxHeaderTB.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $credentialsBox.Header = $credBoxHeaderTB
    $credentialsBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
    $credentialsBox.BorderThickness = '2'
    $credentialsBox.Margin = '0,0,0,20'
    $credentialsBox.Padding = '15,10,15,15'
    $credentialsPanel = New-Object Windows.Controls.StackPanel
    $credentialsPanel.Orientation = 'Vertical'
    $credentialsBox.Content = $credentialsPanel
    $tab2StackPanel.Children.Add($credentialsBox)

    # Tenant
    $tenantLabel = New-Object Windows.Controls.Label
    $tenantLabel.Content = "Tenant"
    $tenantLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $tenantLabel.FontWeight = 'Normal'
    $tenantLabel.Margin = '0,8,0,3'
    $tenantLabel.FontSize = 12
    $tenantLabel.ToolTip = "Your SailPoint tenant name (e.g., 'mycompany' from mycompany.identitynow.com). Required if Custom Tenant URL is empty."
    $credentialsPanel.Children.Add($tenantLabel)

    $tenantTextBox = New-Object Windows.Controls.TextBox
    $tenantTextBox.Text = $settings.tenant
    $tenantTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFFFFF")
    $tenantTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $tenantTextBox.Padding = '8,6'
    $tenantTextBox.FontSize = 12
    $tenantTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
    $tenantTextBox.BorderThickness = '1'
    $tenantTextBox.ToolTip = "Your SailPoint tenant name (e.g., 'mycompany' from mycompany.identitynow.com). Required if Custom Tenant URL is empty."
    $credentialsPanel.Children.Add($tenantTextBox)
    $script:textBoxes['tenant'] = $tenantTextBox

    # Tenant URL (Vanity URL - Optional)
    $tenantUrlLabel = New-Object Windows.Controls.Label
    $tenantUrlLabel.Content = "Custom Tenant URL (optional)"
    $tenantUrlLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $tenantUrlLabel.FontWeight = 'Normal'
    $tenantUrlLabel.Margin = '0,8,0,3'
    $tenantUrlLabel.FontSize = 12
    $tenantUrlLabel.ToolTip = "For vanity URLs (e.g., https://partner7354.identitynow-demo.com). System will construct API URL as https://partner7354.api.identitynow-demo.com"
    $credentialsPanel.Children.Add($tenantUrlLabel)

    $tenantUrlTextBox = New-Object Windows.Controls.TextBox
    $tenantUrlTextBox.Text = $settings.tenantUrl
    $tenantUrlTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFFFFF")
    $tenantUrlTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $tenantUrlTextBox.Padding = '8,6'
    $tenantUrlTextBox.FontSize = 12
    $tenantUrlTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
    $tenantUrlTextBox.BorderThickness = '1'
    $tenantUrlTextBox.ToolTip = "For vanity URLs (e.g., https://partner7354.identitynow-demo.com). System will construct API URL as https://partner7354.api.identitynow-demo.com"
    $credentialsPanel.Children.Add($tenantUrlTextBox)
    $script:textBoxes['tenantUrl'] = $tenantUrlTextBox

    # Client ID
    $clientIDLabel = New-Object Windows.Controls.Label
    $clientIDLabel.Content = "Client ID"
    $clientIDLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $clientIDLabel.FontWeight = 'Normal'
    $clientIDLabel.Margin = '0,8,0,3'
    $clientIDLabel.FontSize = 12
    $clientIDLabel.ToolTip = "OAuth Client ID from your SailPoint API credentials (Admin > API Management > Create Token)"
    $credentialsPanel.Children.Add($clientIDLabel)

    $clientIDTextBox = New-Object Windows.Controls.TextBox
    $clientIDTextBox.Text = $settings.ClientID
    $clientIDTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFFFFF")
    $clientIDTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $clientIDTextBox.Padding = '8,6'
    $clientIDTextBox.FontSize = 12
    $clientIDTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
    $clientIDTextBox.BorderThickness = '1'
    $clientIDTextBox.ToolTip = "OAuth Client ID from your SailPoint API credentials (Admin > API Management > Create Token)"
    $credentialsPanel.Children.Add($clientIDTextBox)
    $script:textBoxes['ClientID'] = $clientIDTextBox

    # Client Secret
    $clientSecretLabel = New-Object Windows.Controls.Label
    $clientSecretLabel.Content = "Client Secret"
    $clientSecretLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $clientSecretLabel.FontWeight = 'Normal'
    $clientSecretLabel.Margin = '0,8,0,3'
    $clientSecretLabel.FontSize = 12
    $clientSecretLabel.ToolTip = "OAuth Client Secret from your SailPoint API credentials (keep this secure!)"
    $credentialsPanel.Children.Add($clientSecretLabel)

    $clientSecretPasswordBox = New-Object Windows.Controls.PasswordBox
    $clientSecretPasswordBox.Password = $settings.ClientSecret
    $clientSecretPasswordBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFFFFF")
    $clientSecretPasswordBox.Foreground = [System.Windows.Media.Brushes]::Black
    $clientSecretPasswordBox.Padding = '8,6'
    $clientSecretPasswordBox.FontSize = 12
    $clientSecretPasswordBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
    $clientSecretPasswordBox.BorderThickness = '1'
    $clientSecretPasswordBox.ToolTip = "OAuth Client Secret from your SailPoint API credentials (keep this secure!)"
    $credentialsPanel.Children.Add($clientSecretPasswordBox)
    $script:clientSecretBox = $clientSecretPasswordBox

    # Options GroupBox
    $optionsBox = New-Object Windows.Controls.GroupBox
    $optBoxHeaderTB = New-Object Windows.Controls.TextBlock
    $optBoxHeaderTB.Text = "Options"
    $optBoxHeaderTB.FontWeight = 'SemiBold'
    $optBoxHeaderTB.FontSize = 13
    $optBoxHeaderTB.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $optionsBox.Header = $optBoxHeaderTB
    $optionsBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
    $optionsBox.BorderThickness = '2'
    $optionsBox.Margin = '0,0,0,20'
    $optionsBox.Padding = '15,10,15,15'
    $optionsPanel = New-Object Windows.Controls.StackPanel
    $optionsPanel.Orientation = 'Vertical'
    $optionsBox.Content = $optionsPanel
    $tab2StackPanel.Children.Add($optionsBox)

    # Days to Keep Files
    $daysLabel = New-Object Windows.Controls.Label
    $daysLabel.Content = "Days to Keep Files"
    $daysLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $daysLabel.FontWeight = 'Normal'
    $daysLabel.Margin = '0,8,0,3'
    $daysLabel.FontSize = 12
    $daysLabel.ToolTip = "Number of days to retain processed files in the Archive folder before automatic deletion. Default: 30 days."
    $optionsPanel.Children.Add($daysLabel)

    $daysTextBox = New-Object Windows.Controls.TextBox
    $daysTextBox.Text = $settings.DaysToKeepFiles.ToString()
    $daysTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFFFFF")
    $daysTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $daysTextBox.Padding = '8,6'
    $daysTextBox.FontSize = 12
    $daysTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
    $daysTextBox.BorderThickness = '1'
    $daysTextBox.ToolTip = "Number of days to retain processed files in the Archive folder before automatic deletion. Default: 30 days."
    $optionsPanel.Children.Add($daysTextBox)
    $script:textBoxes['DaysToKeepFiles'] = $daysTextBox

    # Debug Mode Checkbox
    $debugCheckBox = New-Object Windows.Controls.CheckBox
    $debugCheckBox.Content = "Debug Mode"
    $debugCheckBox.IsChecked = $settings.isDebug
    $debugCheckBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $debugCheckBox.FontWeight = 'Normal'
    $debugCheckBox.Margin = '0,12,0,8'
    $debugCheckBox.FontSize = 12
    $debugCheckBox.ToolTip = "Enable verbose logging for troubleshooting. Logs additional details during file processing and upload."
    $optionsPanel.Children.Add($debugCheckBox)
    $script:debugCheckBox = $debugCheckBox

    # Enable File Deletion Checkbox
    $fileDeletionCheckBox = New-Object Windows.Controls.CheckBox
    $fileDeletionCheckBox.Content = "Enable File Deletion"
    $fileDeletionCheckBox.IsChecked = $settings.enableFileDeletion
    $fileDeletionCheckBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $fileDeletionCheckBox.FontWeight = 'Normal'
    $fileDeletionCheckBox.Margin = '0,0,0,8'
    $fileDeletionCheckBox.FontSize = 12
    $fileDeletionCheckBox.ToolTip = "Allow automatic deletion of archived files older than the specified retention period. Uncheck to keep all files indefinitely."
    $optionsPanel.Children.Add($fileDeletionCheckBox)
    $script:fileDeletionCheckBox = $fileDeletionCheckBox

    # Save Settings Button Row
    $settingsButtonPanel = New-Object Windows.Controls.StackPanel
    $settingsButtonPanel.Orientation = 'Horizontal'
    $settingsButtonPanel.Margin = '0,20,0,15'
    $tab2StackPanel.Children.Add($settingsButtonPanel)

    # Save Settings Button
    $saveButton = New-Object Windows.Controls.Button
    $saveButton.Content = "Save Settings"
    $saveButton.Padding = '15,10'
    $saveButton.Margin = '0,0,12,0'
    $saveButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
    $saveButton.Foreground = [System.Windows.Media.Brushes]::White
    $saveButton.FontWeight = 'Bold'
    $saveButton.FontSize = 14
    $saveButton.BorderThickness = '0'
    $saveButton.Cursor = 'Hand'
    $settingsButtonPanel.Children.Add($saveButton)

    # Match All Files Button
    $matchAllFilesButton = New-Object Windows.Controls.Button
    $matchAllFilesButton.Content = "Match All Files"
    $matchAllFilesButton.Padding = '15,10'
    $matchAllFilesButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#5C6BC0")
    $matchAllFilesButton.Foreground = [System.Windows.Media.Brushes]::White
    $matchAllFilesButton.FontWeight = 'Bold'
    $matchAllFilesButton.FontSize = 14
    $matchAllFilesButton.BorderThickness = '0'
    $matchAllFilesButton.Cursor = 'Hand'
    $matchAllFilesButton.ToolTip = "Copy the most recent matching file from the Source Directory into every app folder (matches by filename containing the app folder name)"
    $settingsButtonPanel.Children.Add($matchAllFilesButton)

    $saveButton.Add_Click({
        # Update settings object
        $settings.ParentDirectory = $script:textBoxes['ParentDirectory'].Text
        $settings.AppFolder = $script:textBoxes['AppFolder'].Text
        $settings.SourceDirectory = $script:textBoxes['SourceDirectory'].Text
        $settings.FileUploadUtility = $script:textBoxes['FileUploadUtility'].Text
        $settings.ExecutionLogDir = $script:textBoxes['ExecutionLogDir'].Text
        $settings.tenant = $script:textBoxes['tenant'].Text
        $settings.tenantUrl = $script:textBoxes['tenantUrl'].Text
        $settings.ClientID = $script:textBoxes['ClientID'].Text
        $settings.ClientSecret = $script:clientSecretBox.Password
        $settings.DaysToKeepFiles = [int]$script:textBoxes['DaysToKeepFiles'].Text
        $settings.isDebug = $script:debugCheckBox.IsChecked
        $settings.enableFileDeletion = $script:fileDeletionCheckBox.IsChecked

        if (Save-Settings -settings $settings) {
            [System.Windows.MessageBox]::Show("Settings saved successfully!", "Success", "OK", "Information")
        }
        else {
            [System.Windows.MessageBox]::Show("Failed to save settings. Check the log for details.", "Error", "OK", "Error")
        }
    })

    $matchAllFilesButton.Add_Click({
        $currentSettings = Load-Settings
        if ($null -eq $currentSettings) {
            [System.Windows.MessageBox]::Show("Failed to load settings!", "Error", "OK", "Error")
            return
        }
        Run-MatchAllFiles -settings $currentSettings
    })

    # RIGHT PANEL: Execution Log Viewer
    $execLogViewerBox = New-Object Windows.Controls.GroupBox
    $execLogHeaderTB = New-Object Windows.Controls.TextBlock
    $execLogHeaderTB.Text = "Execution Log Content"
    $execLogHeaderTB.FontWeight = 'SemiBold'
    $execLogHeaderTB.FontSize = 13
    $execLogHeaderTB.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#000000")
    $execLogViewerBox.Header = $execLogHeaderTB
    $execLogViewerBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#D60EB5")
    $execLogViewerBox.BorderThickness = '2'
    $execLogViewerBox.Margin = '8,0,0,0'
    $execLogViewerBox.Padding = '12'
    [Windows.Controls.Grid]::SetColumn($execLogViewerBox, 1)
    $contentGrid2.Children.Add($execLogViewerBox)

    $execLogViewerPanel = New-Object Windows.Controls.DockPanel
    $execLogViewerBox.Content = $execLogViewerPanel

    # Buttons / selector panel inside the GroupBox (matches App Logs panel layout)
    $execLogButtonsPanel = New-Object Windows.Controls.StackPanel
    $execLogButtonsPanel.Orientation = 'Horizontal'
    $execLogButtonsPanel.Margin = '0,0,0,8'
    [Windows.Controls.DockPanel]::SetDock($execLogButtonsPanel, 'Top')
    $execLogViewerPanel.Children.Add($execLogButtonsPanel)

    # Refresh logs button
    $refreshLogsButton = New-Object Windows.Controls.Button
    $refreshLogsButton.Content = "Refresh Logs"
    $refreshLogsButton.Padding = '10,5'
    $refreshLogsButton.Margin = '0,0,8,0'
    $refreshLogsButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
    $refreshLogsButton.Foreground = [System.Windows.Media.Brushes]::White
    $refreshLogsButton.FontWeight = 'SemiBold'
    $refreshLogsButton.FontSize = 12
    $refreshLogsButton.BorderThickness = '0'
    $refreshLogsButton.Cursor = 'Hand'
    $execLogButtonsPanel.Children.Add($refreshLogsButton)

    # Open log folder button
    $openLogFolderButton = New-Object Windows.Controls.Button
    $openLogFolderButton.Content = "Open Log Folder"
    $openLogFolderButton.Padding = '10,5'
    $openLogFolderButton.Margin = '0,0,16,0'
    $openLogFolderButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
    $openLogFolderButton.Foreground = [System.Windows.Media.Brushes]::White
    $openLogFolderButton.FontWeight = 'SemiBold'
    $openLogFolderButton.FontSize = 12
    $openLogFolderButton.BorderThickness = '0'
    $openLogFolderButton.Cursor = 'Hand'
    $execLogButtonsPanel.Children.Add($openLogFolderButton)

    # Log file selector label
    $logLabel = New-Object Windows.Controls.Label
    $logLabel.Content = "Log File:"
    $logLabel.FontWeight = 'SemiBold'
    $logLabel.FontSize = 12
    $logLabel.VerticalAlignment = 'Center'
    $logLabel.Margin = '0,0,6,0'
    $logLabel.Padding = '0'
    $execLogButtonsPanel.Children.Add($logLabel)

    # Log file selector dropdown
    $execLogDateDropdown = New-Object Windows.Controls.ComboBox
    $execLogDateDropdown.MinWidth = 230
    $execLogDateDropdown.Padding = '6,3'
    $execLogDateDropdown.FontSize = 12
    $execLogDateDropdown.Background = [System.Windows.Media.Brushes]::White
    $execLogDateDropdown.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
    $execLogDateDropdown.BorderThickness = '1'
    $execLogDateDropdown.IsEnabled = $false
    $execLogButtonsPanel.Children.Add($execLogDateDropdown)
    $script:execLogDateDropdown = $execLogDateDropdown

    # Execution log viewer textbox
    $execLogTextBox = New-Object Windows.Controls.TextBox
    $execLogTextBox.AcceptsReturn = $true
    $execLogTextBox.IsReadOnly = $true
    $execLogTextBox.TextWrapping = 'Wrap'
    $execLogTextBox.VerticalScrollBarVisibility = 'Auto'
    $execLogTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFFFFF")
    $execLogTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $execLogTextBox.Padding = '8,6'
    $execLogTextBox.FontSize = 12
    $execLogTextBox.FontFamily = New-Object Windows.Media.FontFamily("Consolas")
    $execLogTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
    $execLogTextBox.BorderThickness = '1'
    $execLogTextBox.Text = "Select an execution log from the dropdown above to view its contents..."
    $execLogViewerPanel.Children.Add($execLogTextBox)
    $script:execLogTextBox = $execLogTextBox

    # Function to reload app log file list and auto-select the most recent entry
    $reloadAppLogs = {
        if ($script:selectedAppName -and $script:currentAppLogFolder -and (Test-Path $script:currentAppLogFolder)) {
            $logFiles = Get-ChildItem -Path $script:currentAppLogFolder -Filter "*.csv" | Sort-Object Name -Descending
            $script:appLogFileDropdown.Items.Clear()
            foreach ($lf in $logFiles) { $script:appLogFileDropdown.Items.Add($lf.Name) | Out-Null }
            $script:appLogFileDropdown.IsEnabled = ($logFiles.Count -gt 0)
            if ($logFiles.Count -gt 0) {
                $script:appLogFileDropdown.SelectedIndex = 0  # triggers SelectionChanged to load content
            } else {
                $script:appLogTextBox.Text = "No log files found"
            }
        }
    }

    # Load the selected app log file when the dropdown selection changes
    $script:appLogFileDropdown.Add_SelectionChanged({
        if ($null -ne $script:appLogFileDropdown.SelectedItem -and $script:currentAppLogFolder) {
            $selectedFile = $script:appLogFileDropdown.SelectedItem.ToString()
            $logPath = Join-Path $script:currentAppLogFolder $selectedFile
            if (Test-Path $logPath) {
                try {
                    $logData = Import-Csv -Path $logPath -ErrorAction Stop
                    $logLines = @($logData) | ForEach-Object {
                        $logType = $_.'Log Type'.PadRight(7)
                        "$($_.'Date/Time') [$logType] $($_.'Log Details')"
                    }
                    [array]::Reverse($logLines)
                    $script:appLogTextBox.Text = "$selectedFile`r`n`r`n" + ($logLines -join "`r`n")
                }
                catch {
                    $script:appLogTextBox.Text = "Error loading log file: $_"
                }
            } else {
                $script:appLogTextBox.Text = "Log file not found: $logPath"
            }
        }
    })

    # Function to load execution log files
    $loadExecutionLogDates = {
        $script:execLogDateDropdown.Items.Clear()
        $script:execLogDateDropdown.IsEnabled = $false
        
        $currentSettings = Load-Settings
        if ([string]::IsNullOrWhiteSpace($currentSettings.ExecutionLogDir)) {
            $script:execLogTextBox.Text = "Execution Log Directory not configured.`n`nPlease set the 'Execution Log Directory' path in the Settings section on the left."
            return
        }
        
        if (-not (Test-Path $currentSettings.ExecutionLogDir)) {
            $script:execLogTextBox.Text = "Execution Log Directory does not exist:`n$($currentSettings.ExecutionLogDir)`n`nPlease verify the path in Settings or create the directory."
            return
        }

        # Get all CSV log files in the execution log directory
        $logFiles = Get-ChildItem -Path $currentSettings.ExecutionLogDir -Filter "*.csv" -File -ErrorAction SilentlyContinue | 
            Sort-Object Name -Descending
        
        if ($null -eq $logFiles -or $logFiles.Count -eq 0) {
            $script:execLogTextBox.Text = "No execution log files (*.csv) found in:`n$($currentSettings.ExecutionLogDir)`n`nLog files will appear here once the file upload script runs."
            return
        }
        
        foreach ($file in $logFiles) {
            $script:execLogDateDropdown.Items.Add($file.Name)
        }
        
        $script:execLogDateDropdown.IsEnabled = $true
        # Auto-select the most recent (first item)
        if ($script:execLogDateDropdown.Items.Count -gt 0) {
            $script:execLogDateDropdown.SelectedIndex = 0
        }
    }

    # Event handler for date dropdown selection changed
    $execLogDateDropdown.Add_SelectionChanged({
        if ($null -ne $script:execLogDateDropdown.SelectedItem) {
            $selectedFile = $script:execLogDateDropdown.SelectedItem.ToString()
            $currentSettings = Load-Settings
            $logPath = Join-Path $currentSettings.ExecutionLogDir $selectedFile
            
            if (Test-Path $logPath) {
                try {
                    # Import CSV and format with most recent first
                    $logData = Import-Csv -Path $logPath -ErrorAction Stop
                    $header = "Execution Log: $selectedFile`r`n`r`n"
                    
                    # Format each row with padding for better alignment
                    $logLines = @($logData) | ForEach-Object {
                        $logType = $_.'Log Type'.PadRight(7)
                        "$($_.'Date/Time') [$logType] $($_.'Log Details')"
                    }
                    
                    # Reverse to show most recent first
                    [array]::Reverse($logLines)
                    $script:execLogTextBox.Text = $header + ($logLines -join "`r`n")
                }
                catch {
                    # If CSV import fails, try reading as raw text
                    try {
                        $logContent = Get-Content -Path $logPath -Raw -ErrorAction Stop
                        $script:execLogTextBox.Text = "Execution Log: $selectedFile`r`n`r`n" + $logContent
                    }
                    catch {
                        $script:execLogTextBox.Text = "Error loading log file: $_"
                    }
                }
            }
            else {
                $script:execLogTextBox.Text = "Log file not found: $logPath"
            }
        }
    })

    # Refresh logs button handler
    $refreshLogsButton.Add_Click({
        & $loadExecutionLogDates
    })

    # Open log folder button handler
    $openLogFolderButton.Add_Click({
        $currentSettings = Load-Settings
        if (-not [string]::IsNullOrWhiteSpace($currentSettings.ExecutionLogDir) -and (Test-Path $currentSettings.ExecutionLogDir)) {
            Start-Process explorer.exe -ArgumentList $currentSettings.ExecutionLogDir
        }
        else {
            [System.Windows.MessageBox]::Show("Execution Log Directory not configured or does not exist.", "Error", "OK", "Error")
        }
    })

    # Load execution log dates on startup
    & $loadExecutionLogDates

    # Directory creation button handler
    $createDirsButton.Add_Click({
        Write-Log "Fetching sources from SailPoint..." "INFO"
        
        # Reload settings in case they changed
        $currentSettings = Load-Settings
        
        # Validate AppFolder is configured first
        if ([string]::IsNullOrWhiteSpace($currentSettings.AppFolder)) {
            Write-Log "ERROR: App Folder is not configured." "ERROR"
            [System.Windows.MessageBox]::Show("App Folder is not configured. Please set it in the Settings tab first.", "Configuration Required", "OK", "Warning")
            return
        }
        
        # Get base API URL
        $tenantUrl = Get-BaseApiUrl -Settings $currentSettings
        if ($null -eq $tenantUrl) {
            [System.Windows.MessageBox]::Show("Configuration Error: Please configure Tenant or Custom Tenant URL in Settings.", "Configuration Required", "OK", "Error")
            return
        }
        
        # Get OAuth token
        $authHeader = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($currentSettings.ClientID):$($currentSettings.ClientSecret)"))
        try {
            $response = Invoke-RestMethod -Uri "$tenantUrl/oauth/token" -Method Post -Headers @{ Authorization = "Basic $authHeader" } -Body @{
                grant_type = 'client_credentials'
            }
            $accessToken = $response.access_token
        }
        catch {
            Write-Log "ERROR: Failed to retrieve OAuth token. $_" "ERROR"
            [System.Windows.MessageBox]::Show("Failed to authenticate with SailPoint. Please check your credentials.", "Authentication Error", "OK", "Error")
            return
        }
        
        # Fetch sources
        $headers = @{ 
            Authorization = "Bearer $accessToken"
            Accept        = "application/json"
        }
        
        try {
            $rawSources = Invoke-RestMethod -Uri "$tenantUrl/beta/sources" -Method Get -Headers $headers -ContentType "application/json;charset=utf-8"
            $parsedSources = @($rawSources)
            $sourcesList = @()
            
            foreach ($source in $parsedSources) {
                if ($source.type -eq "DelimitedFile") {
                    $sourcesList += [PSCustomObject]@{
                        SourceName = $source.name
                        SourceID   = $source.id
                        SourceType = $source.type
                    }
                }
            }
            
            if ($sourcesList.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No Delimited File sources found in SailPoint.", "No Sources", "OK", "Information")
                return
            }
            
            Write-Log "Retrieved $($sourcesList.Count) Delimited File sources from SailPoint." "INFO"
            
            # Create selection popup window
            $selectionWindow = New-Object Windows.Window
            $selectionWindow.Title = "App Management"
            $selectionWindow.Width = 800
            $selectionWindow.Height = 600
            $selectionWindow.WindowStartupLocation = 'CenterScreen'
            $selectionWindow.Background = [System.Windows.Media.Brushes]::White
            
            $mainPanel = New-Object Windows.Controls.StackPanel
            $mainPanel.Margin = '15'
            $selectionWindow.Content = $mainPanel
            
            # Header
            $headerPanel = New-Object Windows.Controls.DockPanel
            $headerPanel.Margin = '0,0,0,10'
            $mainPanel.Children.Add($headerPanel)
            
            $headerLabel = New-Object Windows.Controls.Label
            $headerLabel.Content = "App Management - $($sourcesList.Count) sources"
            $headerLabel.FontSize = 17
            $headerLabel.FontWeight = 'Bold'
            [Windows.Controls.DockPanel]::SetDock($headerLabel, 'Left')
            $headerPanel.Children.Add($headerLabel)
            
            # Select/Deselect All buttons
            $buttonStack = New-Object Windows.Controls.StackPanel
            $buttonStack.Orientation = 'Horizontal'
            [Windows.Controls.DockPanel]::SetDock($buttonStack, 'Right')
            $headerPanel.Children.Add($buttonStack)
            
            $selectAllBtn = New-Object Windows.Controls.Button
            $selectAllBtn.Content = "Select All"
            $selectAllBtn.Padding = '10,5'
            $selectAllBtn.Margin = '0,0,5,0'
            $selectAllBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
            $selectAllBtn.Foreground = [System.Windows.Media.Brushes]::White
            $selectAllBtn.FontSize = 12
            $selectAllBtn.BorderThickness = '0'
            $selectAllBtn.Cursor = 'Hand'
            $buttonStack.Children.Add($selectAllBtn)
            
            $deselectAllBtn = New-Object Windows.Controls.Button
            $deselectAllBtn.Content = "Deselect All"
            $deselectAllBtn.Padding = '10,5'
            $deselectAllBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
            $deselectAllBtn.Foreground = [System.Windows.Media.Brushes]::White
            $deselectAllBtn.FontSize = 12
            $deselectAllBtn.BorderThickness = '0'
            $deselectAllBtn.Cursor = 'Hand'
            $buttonStack.Children.Add($deselectAllBtn)
            
            # Description
            $descLabel = New-Object Windows.Controls.TextBlock
            $descLabel.Text = "Select apps to create/maintain directories for. Toggle 'Enable Upload' to control whether processed files are uploaded to SailPoint."
            $descLabel.TextWrapping = 'Wrap'
            $descLabel.Margin = '0,0,0,10'
            $descLabel.FontSize = 12
            $descLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
            $mainPanel.Children.Add($descLabel)
            
            # Create outer border for the entire table
            $tableBorder = New-Object Windows.Controls.Border
            $tableBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
            $tableBorder.BorderThickness = '1'
            $tableBorder.Margin = '0,0,0,15'
            
            # Create Grid with SharedSizeScope for column alignment
            $tableGrid = New-Object Windows.Controls.Grid
            [Windows.Controls.Grid]::SetIsSharedSizeScope($tableGrid, $true)
            $tableBorder.Child = $tableGrid
            
            # Define rows: header and content
            $headerRow = New-Object Windows.Controls.RowDefinition
            $headerRow.Height = [Windows.GridLength]::Auto
            $contentRow = New-Object Windows.Controls.RowDefinition
            $contentRow.Height = New-Object Windows.GridLength(340)
            $tableGrid.RowDefinitions.Add($headerRow)
            $tableGrid.RowDefinitions.Add($contentRow)
            
            # Column headers in a grid
            $headerGrid = New-Object Windows.Controls.Grid
            $headerGrid.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
            $headerGrid.Height = 35
            [Windows.Controls.Grid]::SetRow($headerGrid, 0)
            
            $col1 = New-Object Windows.Controls.ColumnDefinition
            $col1.Width = New-Object Windows.GridLength(60)
            $col2 = New-Object Windows.Controls.ColumnDefinition
            $col2.Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
            $col3 = New-Object Windows.Controls.ColumnDefinition
            $col3.Width = New-Object Windows.GridLength(150)
            $col4 = New-Object Windows.Controls.ColumnDefinition
            $col4.Width = New-Object Windows.GridLength(150)
            $headerGrid.ColumnDefinitions.Add($col1)
            $headerGrid.ColumnDefinitions.Add($col2)
            $headerGrid.ColumnDefinitions.Add($col3)
            $headerGrid.ColumnDefinitions.Add($col4)
            
            # Create bordered labels for header
            $headerCheckBorder = New-Object Windows.Controls.Border
            $headerCheckBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
            $headerCheckBorder.BorderThickness = '0,0,1,0'
            [Windows.Controls.Grid]::SetColumn($headerCheckBorder, 0)
            $headerCheckLabel = New-Object Windows.Controls.Label
            $headerCheckLabel.Content = ""
            $headerCheckLabel.FontWeight = 'Bold'
            $headerCheckLabel.FontSize = 13
            $headerCheckLabel.HorizontalAlignment = 'Center'
            $headerCheckLabel.VerticalAlignment = 'Center'
            $headerCheckLabel.Padding = '5'
            $headerCheckBorder.Child = $headerCheckLabel
            $headerGrid.Children.Add($headerCheckBorder)
            
            $headerNameBorder = New-Object Windows.Controls.Border
            $headerNameBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
            $headerNameBorder.BorderThickness = '0,0,1,0'
            [Windows.Controls.Grid]::SetColumn($headerNameBorder, 1)
            $headerNameLabel = New-Object Windows.Controls.Label
            $headerNameLabel.Content = "App Name"
            $headerNameLabel.FontWeight = 'Bold'
            $headerNameLabel.FontSize = 12
            $headerNameLabel.VerticalAlignment = 'Center'
            $headerNameLabel.Padding = '10,5'
            $headerNameBorder.Child = $headerNameLabel
            $headerGrid.Children.Add($headerNameBorder)
            
            $headerStatusBorder = New-Object Windows.Controls.Border
            $headerStatusBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
            $headerStatusBorder.BorderThickness = '0,0,1,0'
            [Windows.Controls.Grid]::SetColumn($headerStatusBorder, 2)
            $headerStatusLabel = New-Object Windows.Controls.Label
            $headerStatusLabel.Content = "Directory Status"
            $headerStatusLabel.FontWeight = 'Bold'
            $headerStatusLabel.FontSize = 12
            $headerStatusLabel.HorizontalAlignment = 'Center'
            $headerStatusLabel.VerticalAlignment = 'Center'
            $headerStatusLabel.Padding = '5'
            $headerStatusBorder.Child = $headerStatusLabel
            $headerGrid.Children.Add($headerStatusBorder)
            
            $headerUploadLabel = New-Object Windows.Controls.Label
            $headerUploadLabel.Content = "Enable Upload"
            $headerUploadLabel.FontWeight = 'Bold'
            $headerUploadLabel.FontSize = 12
            $headerUploadLabel.HorizontalAlignment = 'Center'
            $headerUploadLabel.VerticalAlignment = 'Center'
            $headerUploadLabel.Padding = '5'
            [Windows.Controls.Grid]::SetColumn($headerUploadLabel, 3)
            $headerGrid.Children.Add($headerUploadLabel)
            
            # Add separator line below header
            $headerSeparator = New-Object Windows.Controls.Border
            $headerSeparator.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
            $headerSeparator.BorderThickness = '0,0,0,1'
            $headerSeparator.Child = $headerGrid
            $tableGrid.Children.Add($headerSeparator)
            
            # ScrollViewer with app list
            $scrollViewer = New-Object Windows.Controls.ScrollViewer
            $scrollViewer.VerticalScrollBarVisibility = 'Auto'
            [Windows.Controls.Grid]::SetRow($scrollViewer, 1)
            
            $appListPanel = New-Object Windows.Controls.StackPanel
            $appListPanel.Margin = '0'
            $appListPanel.Background = [System.Windows.Media.Brushes]::White
            $scrollViewer.Content = $appListPanel
            $tableGrid.Children.Add($scrollViewer)
            
            $mainPanel.Children.Add($tableBorder)
            
            # Create grid rows for each source
            $checkboxes = @{}
            $uploadToggles = @{}
            
            # Check which directories already exist and load their configs
            $existingDirs = @()
            if (-not [string]::IsNullOrWhiteSpace($currentSettings.AppFolder) -and (Test-Path $currentSettings.AppFolder)) {
                $existingDirs = Get-ChildItem -Path $currentSettings.AppFolder -Directory | Select-Object -ExpandProperty Name
            }
            
            foreach ($source in $sourcesList | Sort-Object @{Expression={$existingDirs -contains $_.SourceName}; Descending=$true}, SourceName) {
                $appName = $source.SourceName
                $dirExists = $existingDirs -contains $appName
                
                # Load config if directory exists
                $uploadEnabled = $false
                if ($dirExists) {
                    $configPath = Join-Path $currentSettings.AppFolder "$appName\config.json"
                    if (Test-Path $configPath) {
                        try {
                            $appConfig = Get-Content -Path $configPath -Raw | ConvertFrom-Json
                            $uploadEnabled = $appConfig.isUpload -eq $true
                        } catch {
                            Write-Log "Warning: Could not load config for $appName" "WARNING"
                        }
                    }
                }
                
                # Create row grid with border
                $rowGrid = New-Object Windows.Controls.Grid
                $rowGrid.Margin = '0'
                $rowGrid.Background = [System.Windows.Media.Brushes]::White
                $rowGrid.Height = 38
                
                # Add border to row
                $rowBorder = New-Object Windows.Controls.Border
                $rowBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
                $rowBorder.BorderThickness = '0,0,0,1'
                $rowBorder.Child = $rowGrid
                
                $rowCol1 = New-Object Windows.Controls.ColumnDefinition
                $rowCol1.Width = New-Object Windows.GridLength(60)
                $rowCol2 = New-Object Windows.Controls.ColumnDefinition
                $rowCol2.Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
                $rowCol3 = New-Object Windows.Controls.ColumnDefinition
                $rowCol3.Width = New-Object Windows.GridLength(150)
                $rowCol4 = New-Object Windows.Controls.ColumnDefinition
                $rowCol4.Width = New-Object Windows.GridLength(150)
                $rowGrid.ColumnDefinitions.Add($rowCol1)
                $rowGrid.ColumnDefinitions.Add($rowCol2)
                $rowGrid.ColumnDefinitions.Add($rowCol3)
                $rowGrid.ColumnDefinitions.Add($rowCol4)
                
                # Checkbox column with border
                $checkboxBorder = New-Object Windows.Controls.Border
                $checkboxBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
                $checkboxBorder.BorderThickness = '0,0,1,0'
                [Windows.Controls.Grid]::SetColumn($checkboxBorder, 0)
                $checkbox = New-Object Windows.Controls.CheckBox
                $checkbox.IsChecked = $dirExists
                $checkbox.HorizontalAlignment = 'Center'
                $checkbox.VerticalAlignment = 'Center'
                $checkbox.Tag = $appName
                $checkboxBorder.Child = $checkbox
                $rowGrid.Children.Add($checkboxBorder)
                $checkboxes[$appName] = $checkbox
                
                # App name column with border
                $nameBorder = New-Object Windows.Controls.Border
                $nameBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
                $nameBorder.BorderThickness = '0,0,1,0'
                [Windows.Controls.Grid]::SetColumn($nameBorder, 1)
                $nameLabel = New-Object Windows.Controls.Label
                $nameLabel.Content = $appName
                $nameLabel.FontSize = 12
                $nameLabel.VerticalAlignment = 'Center'
                $nameLabel.Padding = '10,5'
                $nameBorder.Child = $nameLabel
                $rowGrid.Children.Add($nameBorder)
                
                # Status column with border
                $statusBorder = New-Object Windows.Controls.Border
                $statusBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
                $statusBorder.BorderThickness = '0,0,1,0'
                [Windows.Controls.Grid]::SetColumn($statusBorder, 2)
                $statusLabel = New-Object Windows.Controls.Label
                if ($dirExists) {
                    $statusLabel.Content = "Exists"
                    $statusLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
                } else {
                    $statusLabel.Content = "Not Created"
                    $statusLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
                }
                $statusLabel.FontSize = 12
                $statusLabel.HorizontalAlignment = 'Center'
                $statusLabel.VerticalAlignment = 'Center'
                $statusLabel.Padding = '5,2'
                $statusBorder.Child = $statusLabel
                $rowGrid.Children.Add($statusBorder)
                
                # Upload toggle checkbox (no border on last column)
                # Upload toggle checkbox (no border on last column)
                $uploadToggle = New-Object Windows.Controls.CheckBox
                $uploadToggle.IsChecked = $uploadEnabled
                $uploadToggle.Content = if ($uploadEnabled) { "Enabled" } else { "Disabled" }
                $uploadToggle.HorizontalAlignment = 'Center'
                $uploadToggle.VerticalAlignment = 'Center'
                $uploadToggle.FontSize = 12
                $uploadToggle.Tag = $appName
                [Windows.Controls.Grid]::SetColumn($uploadToggle, 3)
                $rowGrid.Children.Add($uploadToggle)
                $uploadToggles[$appName] = $uploadToggle
                
                # Update label text when toggled
                $uploadToggle.Add_Checked({
                    $this.Content = "Enabled"
                }.GetNewClosure())
                $uploadToggle.Add_Unchecked({
                    $this.Content = "Disabled"
                }.GetNewClosure())
                
                $appListPanel.Children.Add($rowBorder)
            }
            
            # Select/Deselect All handlers
            $selectAllBtn.Add_Click({
                foreach ($cb in $checkboxes.Values) {
                    $cb.IsChecked = $true
                }
            }.GetNewClosure())
            
            $deselectAllBtn.Add_Click({
                foreach ($cb in $checkboxes.Values) {
                    $cb.IsChecked = $false
                }
            }.GetNewClosure())
            
            # Bottom buttons
            $bottomPanel = New-Object Windows.Controls.StackPanel
            $bottomPanel.Orientation = 'Horizontal'
            $bottomPanel.HorizontalAlignment = 'Center'
            $bottomPanel.Margin = '0,10,0,0'
            $mainPanel.Children.Add($bottomPanel)
            
            $cancelBtn = New-Object Windows.Controls.Button
            $cancelBtn.Content = "Cancel"
            $cancelBtn.Padding = '15,8'
            $cancelBtn.Margin = '0,0,10,0'
            $cancelBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
            $cancelBtn.Foreground = [System.Windows.Media.Brushes]::White
            $cancelBtn.FontSize = 12
            $cancelBtn.BorderThickness = '0'
            $cancelBtn.Cursor = 'Hand'
            $bottomPanel.Children.Add($cancelBtn)
            
            $createBtn = New-Object Windows.Controls.Button
            $createBtn.Content = "Apply Changes"
            $createBtn.Padding = '15,8'
            $createBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
            $createBtn.Foreground = [System.Windows.Media.Brushes]::White
            $createBtn.FontWeight = 'Bold'
            $createBtn.FontSize = 12
            $createBtn.BorderThickness = '0'
            $createBtn.Cursor = 'Hand'
            $bottomPanel.Children.Add($createBtn)
            
            $cancelBtn.Add_Click({
                $selectionWindow.Close()
            }.GetNewClosure())
            
            $createBtn.Add_Click({
                try {
                    # Get selected and unselected sources
                    $selectedSourceNames = @()
                    $unselectedSourceNames = @()
                    $uploadSettings = @{}
                    
                    foreach ($cb in $checkboxes.Values) {
                        $appName = $cb.Tag
                        if ($cb.IsChecked) {
                            $selectedSourceNames += $appName
                        }
                        else {
                            $unselectedSourceNames += $appName
                        }
                        # Store upload setting for each app
                        $uploadSettings[$appName] = $uploadToggles[$appName].IsChecked
                    }
                    
                    # Close window first
                    $selectionWindow.Close()
                    
                    # Validate AppFolder is configured
                    if ([string]::IsNullOrWhiteSpace($currentSettings.AppFolder)) {
                        Write-Log "ERROR: App Folder is not configured. Please configure it in Settings first." "ERROR"
                        [System.Windows.MessageBox]::Show("App Folder is not configured. Please set it in the Settings tab first.", "Configuration Required", "OK", "Warning")
                        return
                    }
                    
                    $changesMade = $false
                    
                    # Delete directories for unselected sources that exist
                    if ($unselectedSourceNames.Count -gt 0) {
                        foreach ($sourceName in $unselectedSourceNames) {
                            $dirPath = Join-Path $currentSettings.AppFolder $sourceName
                            if (Test-Path $dirPath) {
                                try {
                                    Remove-Item -Path $dirPath -Recurse -Force
                                    Write-Log "Deleted directory: $sourceName" "INFO"
                                    $changesMade = $true
                                }
                                catch {
                                    Write-Log "ERROR: Failed to delete directory $sourceName`: $_" "ERROR"
                                }
                            }
                        }
                    }
                    
                    # Create directories for selected sources
                    if ($selectedSourceNames.Count -gt 0) {
                        Write-Log "Creating directories for $($selectedSourceNames.Count) selected sources..." "INFO"
                        $result = Run-DirectoryCreation -settings $currentSettings -selectedSources $selectedSourceNames
                        
                        if ($result) {
                            $changesMade = $true
                            Write-Log "Directory creation completed successfully!" "INFO"
                        }
                        else {
                            Write-Log "Directory creation encountered errors. Check the log above." "ERROR"
                        }
                    }
                    
                    # Update isUpload setting in config.json for all selected apps
                    foreach ($appName in $selectedSourceNames) {
                        $configPath = Join-Path $currentSettings.AppFolder "$appName\config.json"
                        if (Test-Path $configPath) {
                            try {
                                $appConfigContent = Get-Content -Path $configPath -Raw
                                $appConfig = $appConfigContent | ConvertFrom-Json
                                $newUploadValue = $uploadSettings[$appName]
                                
                                # Only update if value changed
                                if ($appConfig.isUpload -ne $newUploadValue) {
                                    $appConfig.isUpload = $newUploadValue
                                    $appConfig | ConvertTo-Json -Depth 10 | Out-File -FilePath $configPath -Encoding UTF8
                                    Write-Log "Updated upload setting for ${appName}: isUpload = $newUploadValue" "INFO"
                                    $changesMade = $true
                                }
                            }
                            catch {
                                Write-Log "ERROR: Failed to update config for ${appName}: $_" "ERROR"
                            }
                        }
                    }
                    
                    # Refresh app list after changes
                    if ($changesMade) {
                        Write-Log "App management completed. Please refresh the page to see updated app list." "INFO"
                        [System.Windows.MessageBox]::Show("Changes applied successfully!`n`nPlease click the 'Refresh App List' button to see the updated apps.", "Success", "OK", "Information")
                    }
                    else {
                        Write-Log "No changes were needed." "INFO"
                    }
                }
                catch {
                    Write-Log "ERROR: Apply changes failed: $_" "ERROR"
                    [System.Windows.MessageBox]::Show("Error applying changes: $_", "Error", "OK", "Error")
                }
            })
            
            # Show the selection window
            $selectionWindow.ShowDialog() | Out-Null
        }
        catch {
            Write-Log "ERROR: Failed to fetch sources. $_" "ERROR"
            [System.Windows.MessageBox]::Show("Failed to fetch sources from SailPoint: $_", "Error", "OK", "Error")
        }
    })

    # Create Source Wizard
    $createSourceButton.Add_Click({
        $currentSettings = Load-Settings

        if ([string]::IsNullOrWhiteSpace($currentSettings.AppFolder)) {
            [System.Windows.MessageBox]::Show("App Folder is not configured. Please set it in the Settings tab first.", "Configuration Required", "OK", "Warning")
            return
        }

        $tenantUrl = Get-BaseApiUrl -Settings $currentSettings
        if ($null -eq $tenantUrl) {
            [System.Windows.MessageBox]::Show("Configuration Error: Please configure Tenant or Custom Tenant URL in Settings.", "Configuration Required", "OK", "Error")
            return
        }

        # Capture script-level functions for use inside .GetNewClosure() blocks
        $writeLog      = ${function:Write-Log}
        $loadAppFoldersFn = $loadAppFolders

        # Shared mutable state passed between inner closures via hashtable reference
        $wizardState = @{
            Step        = 1
            SourceName  = ""
            Description = ""
            SourceFile  = ""
        }

        # === WIZARD WINDOW ===
        $wizardWindow = New-Object Windows.Window
        $wizardWindow.Title = "Create New Source"
        $wizardWindow.Width = 540
        $wizardWindow.Height = 490
        $wizardWindow.WindowStartupLocation = 'CenterScreen'
        $wizardWindow.Background = [System.Windows.Media.Brushes]::White
        $wizardWindow.ResizeMode = 'NoResize'

        $wizardMain = New-Object Windows.Controls.DockPanel
        $wizardWindow.Content = $wizardMain

        # --- Header ---
        $wzHeader = New-Object Windows.Controls.Border
        $wzHeader.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
        $wzHeader.Padding = '20,14'
        [Windows.Controls.DockPanel]::SetDock($wzHeader, 'Top')
        $wizardMain.Children.Add($wzHeader)

        $wzHeaderDP = New-Object Windows.Controls.DockPanel
        $wzHeader.Child = $wzHeaderDP

        $wzStepLabel = New-Object Windows.Controls.TextBlock
        $wzStepLabel.Text = "Step 1 of 3"
        $wzStepLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#54C0E8")
        $wzStepLabel.FontSize = 11
        $wzStepLabel.VerticalAlignment = 'Center'
        [Windows.Controls.DockPanel]::SetDock($wzStepLabel, 'Right')
        $wzHeaderDP.Children.Add($wzStepLabel)

        $wzTitleText = New-Object Windows.Controls.TextBlock
        $wzTitleText.Text = "Create New Source"
        $wzTitleText.Foreground = [System.Windows.Media.Brushes]::White
        $wzTitleText.FontSize = 17
        $wzTitleText.FontWeight = 'SemiBold'
        $wzHeaderDP.Children.Add($wzTitleText)

        # --- Footer ---
        $wzFooter = New-Object Windows.Controls.Border
        $wzFooter.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F0F4FA")
        $wzFooter.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
        $wzFooter.BorderThickness = '0,1,0,0'
        $wzFooter.Padding = '20,12'
        [Windows.Controls.DockPanel]::SetDock($wzFooter, 'Bottom')
        $wizardMain.Children.Add($wzFooter)

        $wzFooterDP = New-Object Windows.Controls.DockPanel
        $wzFooterDP.LastChildFill = $false
        $wzFooter.Child = $wzFooterDP

        $wzCancelBtn = New-Object Windows.Controls.Button
        $wzCancelBtn.Content = "Cancel"
        $wzCancelBtn.Padding = '14,7'
        $wzCancelBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
        $wzCancelBtn.Foreground = [System.Windows.Media.Brushes]::White
        $wzCancelBtn.FontSize = 12
        $wzCancelBtn.BorderThickness = '0'
        $wzCancelBtn.Cursor = 'Hand'
        [Windows.Controls.DockPanel]::SetDock($wzCancelBtn, 'Left')
        $wzFooterDP.Children.Add($wzCancelBtn)

        $wzNextBtn = New-Object Windows.Controls.Button
        $wzNextBtn.Content = "Next"
        $wzNextBtn.Padding = '14,7'
        $wzNextBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
        $wzNextBtn.Foreground = [System.Windows.Media.Brushes]::White
        $wzNextBtn.FontWeight = 'SemiBold'
        $wzNextBtn.FontSize = 12
        $wzNextBtn.BorderThickness = '0'
        $wzNextBtn.Cursor = 'Hand'
        [Windows.Controls.DockPanel]::SetDock($wzNextBtn, 'Right')
        $wzFooterDP.Children.Add($wzNextBtn)

        $wzBackBtn = New-Object Windows.Controls.Button
        $wzBackBtn.Content = "Back"
        $wzBackBtn.Padding = '14,7'
        $wzBackBtn.Margin = '0,0,8,0'
        $wzBackBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
        $wzBackBtn.Foreground = [System.Windows.Media.Brushes]::White
        $wzBackBtn.FontSize = 12
        $wzBackBtn.BorderThickness = '0'
        $wzBackBtn.Cursor = 'Hand'
        $wzBackBtn.IsEnabled = $false
        [Windows.Controls.DockPanel]::SetDock($wzBackBtn, 'Right')
        $wzFooterDP.Children.Add($wzBackBtn)

        # --- Content area ---
        $wzContent = New-Object Windows.Controls.Border
        $wzContent.Padding = '25,20'
        $wizardMain.Children.Add($wzContent)

        $stepsContainer = New-Object Windows.Controls.StackPanel
        $wzContent.Child = $stepsContainer

        # ---- STEP 1: Source Details ----
        $step1 = New-Object Windows.Controls.StackPanel

        $s1Title = New-Object Windows.Controls.TextBlock
        $s1Title.Text = "Source Details"
        $s1Title.FontSize = 15
        $s1Title.FontWeight = 'SemiBold'
        $s1Title.Margin = '0,0,0,6'
        $step1.Children.Add($s1Title)

        $s1InfoText = New-Object Windows.Controls.TextBlock
        $s1InfoText.Text = "Provide a name and description for the new source. The source name will also be used for the local app folder."
        $s1InfoText.TextWrapping = 'Wrap'
        $s1InfoText.FontSize = 12
        $s1InfoText.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
        $s1InfoText.Margin = '0,0,0,18'
        $step1.Children.Add($s1InfoText)

        $s1NameLbl = New-Object Windows.Controls.TextBlock
        $s1NameLbl.Text = "Source Name *"
        $s1NameLbl.FontWeight = 'SemiBold'
        $s1NameLbl.FontSize = 12
        $s1NameLbl.Margin = '0,0,0,4'
        $step1.Children.Add($s1NameLbl)

        $s1NameBox = New-Object Windows.Controls.TextBox
        $s1NameBox.Padding = '8,6'
        $s1NameBox.FontSize = 12
        $s1NameBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
        $s1NameBox.BorderThickness = '1'
        $s1NameBox.Margin = '0,0,0,14'
        $step1.Children.Add($s1NameBox)

        $s1DescLbl = New-Object Windows.Controls.TextBlock
        $s1DescLbl.Text = "Description"
        $s1DescLbl.FontWeight = 'SemiBold'
        $s1DescLbl.FontSize = 12
        $s1DescLbl.Margin = '0,0,0,4'
        $step1.Children.Add($s1DescLbl)

        $s1DescBox = New-Object Windows.Controls.TextBox
        $s1DescBox.Padding = '8,6'
        $s1DescBox.FontSize = 12
        $s1DescBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
        $s1DescBox.BorderThickness = '1'
        $s1DescBox.Height = 70
        $s1DescBox.TextWrapping = 'Wrap'
        $s1DescBox.AcceptsReturn = $true
        $s1DescBox.VerticalScrollBarVisibility = 'Auto'
        $step1.Children.Add($s1DescBox)

        $s1ValidationLbl = New-Object Windows.Controls.TextBlock
        $s1ValidationLbl.Text = ""
        $s1ValidationLbl.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#C62828")
        $s1ValidationLbl.FontSize = 11
        $s1ValidationLbl.Margin = '0,8,0,0'
        $step1.Children.Add($s1ValidationLbl)

        $stepsContainer.Children.Add($step1)

        # ---- STEP 2: Source File ----
        $step2 = New-Object Windows.Controls.StackPanel
        $step2.Visibility = 'Collapsed'

        $s2Title = New-Object Windows.Controls.TextBlock
        $s2Title.Text = "Source File"
        $s2Title.FontSize = 15
        $s2Title.FontWeight = 'SemiBold'
        $s2Title.Margin = '0,0,0,6'
        $step2.Children.Add($s2Title)

        $s2InfoText = New-Object Windows.Controls.TextBlock
        $s2InfoText.TextWrapping = 'Wrap'
        $s2InfoText.FontSize = 12
        $s2InfoText.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
        $s2InfoText.Margin = '0,0,0,18'
        $step2.Children.Add($s2InfoText)

        $s2FileLbl = New-Object Windows.Controls.TextBlock
        $s2FileLbl.Text = "Source File (CSV or Excel)"
        $s2FileLbl.FontWeight = 'SemiBold'
        $s2FileLbl.FontSize = 12
        $s2FileLbl.Margin = '0,0,0,4'
        $step2.Children.Add($s2FileLbl)

        $s2PickerRow = New-Object Windows.Controls.DockPanel
        $s2PickerRow.Margin = '0,0,0,8'
        $step2.Children.Add($s2PickerRow)

        $s2BrowseBtn = New-Object Windows.Controls.Button
        $s2BrowseBtn.Content = "Browse..."
        $s2BrowseBtn.Padding = '12,6'
        $s2BrowseBtn.Margin = '8,0,0,0'
        $s2BrowseBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
        $s2BrowseBtn.Foreground = [System.Windows.Media.Brushes]::White
        $s2BrowseBtn.FontSize = 12
        $s2BrowseBtn.BorderThickness = '0'
        $s2BrowseBtn.Cursor = 'Hand'
        [Windows.Controls.DockPanel]::SetDock($s2BrowseBtn, 'Right')
        $s2PickerRow.Children.Add($s2BrowseBtn)

        $s2FileBox = New-Object Windows.Controls.TextBox
        $s2FileBox.Padding = '8,6'
        $s2FileBox.FontSize = 12
        $s2FileBox.IsReadOnly = $true
        $s2FileBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
        $s2FileBox.BorderThickness = '1'
        $s2FileBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F5F5F5")
        $s2FileBox.Text = "No file selected"
        $s2PickerRow.Children.Add($s2FileBox)

        $s2SkipNote = New-Object Windows.Controls.TextBlock
        $s2SkipNote.Text = "Optional - you can place the source file in the app folder manually after creation."
        $s2SkipNote.TextWrapping = 'Wrap'
        $s2SkipNote.FontSize = 11
        $s2SkipNote.FontStyle = 'Italic'
        $s2SkipNote.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
        $step2.Children.Add($s2SkipNote)

        $stepsContainer.Children.Add($step2)

        # ---- STEP 3: Review & Create ----
        $step3 = New-Object Windows.Controls.StackPanel
        $step3.Visibility = 'Collapsed'

        $s3Title = New-Object Windows.Controls.TextBlock
        $s3Title.Text = "Review & Create"
        $s3Title.FontSize = 15
        $s3Title.FontWeight = 'SemiBold'
        $s3Title.Margin = '0,0,0,6'
        $step3.Children.Add($s3Title)

        $s3InfoText = New-Object Windows.Controls.TextBlock
        $s3InfoText.Text = "Review the details below and click 'Create Source' to proceed."
        $s3InfoText.TextWrapping = 'Wrap'
        $s3InfoText.FontSize = 12
        $s3InfoText.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
        $s3InfoText.Margin = '0,0,0,14'
        $step3.Children.Add($s3InfoText)

        $s3ReviewBorder = New-Object Windows.Controls.Border
        $s3ReviewBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
        $s3ReviewBorder.BorderThickness = '1'
        $s3ReviewBorder.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F5F8FF")
        $s3ReviewBorder.Padding = '14,12'
        $s3ReviewBorder.Margin = '0,0,0,14'
        $step3.Children.Add($s3ReviewBorder)

        $s3ReviewPanel = New-Object Windows.Controls.StackPanel
        $s3ReviewBorder.Child = $s3ReviewPanel

        $s3ReviewName = New-Object Windows.Controls.TextBlock
        $s3ReviewName.FontSize = 12
        $s3ReviewName.Margin = '0,0,0,5'
        $s3ReviewPanel.Children.Add($s3ReviewName)

        $s3ReviewDesc = New-Object Windows.Controls.TextBlock
        $s3ReviewDesc.FontSize = 12
        $s3ReviewDesc.TextWrapping = 'Wrap'
        $s3ReviewDesc.Margin = '0,0,0,5'
        $s3ReviewPanel.Children.Add($s3ReviewDesc)

        $s3ReviewFolder = New-Object Windows.Controls.TextBlock
        $s3ReviewFolder.FontSize = 12
        $s3ReviewFolder.TextWrapping = 'Wrap'
        $s3ReviewFolder.Margin = '0,0,0,5'
        $s3ReviewPanel.Children.Add($s3ReviewFolder)

        $s3ReviewFile = New-Object Windows.Controls.TextBlock
        $s3ReviewFile.FontSize = 12
        $s3ReviewFile.TextWrapping = 'Wrap'
        $s3ReviewPanel.Children.Add($s3ReviewFile)

        $s3StatusLbl = New-Object Windows.Controls.TextBlock
        $s3StatusLbl.Text = ""
        $s3StatusLbl.TextWrapping = 'Wrap'
        $s3StatusLbl.FontSize = 12
        $s3StatusLbl.Margin = '0,0,0,0'
        $step3.Children.Add($s3StatusLbl)

        $stepsContainer.Children.Add($step3)

        # ---- Event Handlers ----

        $wzCancelBtn.Add_Click({
            $wizardWindow.Close()
        }.GetNewClosure())

        $s2BrowseBtn.Add_Click({
            $ofd = New-Object Microsoft.Win32.OpenFileDialog
            $ofd.Title = "Select Source File"
            $ofd.Filter = "Data Files (*.csv;*.xlsx;*.xls)|*.csv;*.xlsx;*.xls|All Files (*.*)|*.*"
            if ($ofd.ShowDialog() -eq $true) {
                $s2FileBox.Text = $ofd.FileName
                $wizardState.SourceFile = $ofd.FileName
            }
        }.GetNewClosure())

        $wzBackBtn.Add_Click({
            $wizardState.Step--
            $step1.Visibility = 'Collapsed'
            $step2.Visibility = 'Collapsed'
            $step3.Visibility = 'Collapsed'
            if ($wizardState.Step -eq 1) {
                $step1.Visibility = 'Visible'
                $wzStepLabel.Text = "Step 1 of 3"
                $wzBackBtn.IsEnabled = $false
                $wzNextBtn.Content = "Next"
                $wzNextBtn.IsEnabled = $true
            }
            elseif ($wizardState.Step -eq 2) {
                $step2.Visibility = 'Visible'
                $wzStepLabel.Text = "Step 2 of 3"
                $wzBackBtn.IsEnabled = $true
                $wzNextBtn.Content = "Next"
                $wzNextBtn.IsEnabled = $true
            }
        }.GetNewClosure())

        $wzNextBtn.Add_Click({
            if ($wizardState.Step -eq 1) {
                # --- Validate step 1 ---
                $sourceName = $s1NameBox.Text.Trim()
                if ([string]::IsNullOrWhiteSpace($sourceName)) {
                    $s1ValidationLbl.Text = "Source name is required."
                    return
                }
                $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
                foreach ($ch in $invalidChars) {
                    if ($sourceName.IndexOf($ch) -ge 0) {
                        $s1ValidationLbl.Text = "Source name contains invalid characters."
                        return
                    }
                }
                $targetFolder = Join-Path $currentSettings.AppFolder $sourceName
                if (Test-Path $targetFolder) {
                    $s1ValidationLbl.Text = "A folder named '$sourceName' already exists in the App Folder."
                    return
                }
                $s1ValidationLbl.Text = ""
                $wizardState.SourceName  = $sourceName
                $wizardState.Description = $s1DescBox.Text.Trim()

                $s2InfoText.Text = "Select the source CSV or Excel file for '$($wizardState.SourceName)'. It will be copied to: $(Join-Path $currentSettings.AppFolder $wizardState.SourceName)"
                $wizardState.Step = 2
                $step1.Visibility = 'Collapsed'
                $step2.Visibility = 'Visible'
                $wzStepLabel.Text = "Step 2 of 3"
                $wzBackBtn.IsEnabled = $true
                $wzNextBtn.Content = "Next"

            }
            elseif ($wizardState.Step -eq 2) {
                # --- Advance to review ---
                $selFile = $s2FileBox.Text
                $wizardState.SourceFile = if ($selFile -eq "No file selected") { "" } else { $selFile }

                $descDisplay   = if ([string]::IsNullOrWhiteSpace($wizardState.Description)) { "(none)" } else { $wizardState.Description }
                $fileDisplay   = if ([string]::IsNullOrWhiteSpace($wizardState.SourceFile))  { "(none - add manually later)" } else { $wizardState.SourceFile }
                $folderDisplay = Join-Path $currentSettings.AppFolder $wizardState.SourceName

                $s3ReviewName.Text   = "Source Name:  $($wizardState.SourceName)"
                $s3ReviewDesc.Text   = "Description:  $descDisplay"
                $s3ReviewFolder.Text = "App Folder:   $folderDisplay"
                $s3ReviewFile.Text   = "Source File:  $fileDisplay"
                $s3StatusLbl.Text    = ""

                $wizardState.Step = 3
                $step2.Visibility = 'Collapsed'
                $step3.Visibility = 'Visible'
                $wzStepLabel.Text = "Step 3 of 3"
                $wzBackBtn.IsEnabled = $true
                $wzNextBtn.Content = "Create Source"

            }
            elseif ($wizardState.Step -eq 3) {
                # === PERFORM CREATION ===
                $wzNextBtn.IsEnabled  = $false
                $wzBackBtn.IsEnabled  = $false
                $wzCancelBtn.IsEnabled = $false
                $s3StatusLbl.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
                $s3StatusLbl.Text = "Authenticating with SailPoint..."

                try {
                    # --- Authenticate ---
                    $authHeader = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($currentSettings.ClientID):$($currentSettings.ClientSecret)"))
                    $tokenResp  = Invoke-RestMethod -Uri "$tenantUrl/oauth/token" -Method Post `
                        -Headers @{ Authorization = "Basic $authHeader" } `
                        -Body @{ grant_type = 'client_credentials' }
                    $accessToken = $tokenResp.access_token

                    # --- Decode JWT to get owner identity ID ---
                    $s3StatusLbl.Text = "Resolving owner identity..."
                    $jwtPayload = $accessToken.Split('.')[1]
                    $pad = 4 - ($jwtPayload.Length % 4)
                    if ($pad -ne 4) { $jwtPayload += '=' * $pad }
                    $claims = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($jwtPayload)) | ConvertFrom-Json
                    $ownerIdentityId   = $claims.identity_id
                    $ownerIdentityName = $claims.sub

                    if ([string]::IsNullOrWhiteSpace($ownerIdentityId)) {
                        throw "Could not determine owner identity from access token. Ensure you are using a Personal Access Token."
                    }

                    # --- Create source in SailPoint ISC ---
                    $s3StatusLbl.Text = "Creating source in SailPoint ISC..."
                    $apiHeaders = @{
                        Authorization  = "Bearer $accessToken"
                        'Content-Type' = 'application/json'
                        Accept         = 'application/json'
                    }
                    $sourceBody = @{
                        name        = $wizardState.SourceName
                        description = $wizardState.Description
                        connector   = "delimited-file"
                        owner       = @{
                            type = "IDENTITY"
                            id   = $ownerIdentityId
                            name = $ownerIdentityName
                        }
                    } | ConvertTo-Json -Depth 5

                    $newSource   = Invoke-RestMethod -Uri "$tenantUrl/v3/sources?provisionAsCsv=true" -Method Post -Headers $apiHeaders -Body $sourceBody
                    $newSourceId = $newSource.id

                    if ([string]::IsNullOrWhiteSpace($newSourceId)) {
                        throw "Source was created but no ID was returned. Please verify in SailPoint ISC."
                    }
                    & $writeLog "Created SailPoint source '$($wizardState.SourceName)' with ID: $newSourceId" "INFO"

                    # --- Create local folder structure ---
                    $s3StatusLbl.Text = "Creating local folder structure..."
                    $newAppFolder = Join-Path $currentSettings.AppFolder $wizardState.SourceName
                    New-Item -Path $newAppFolder                         -ItemType Directory | Out-Null
                    New-Item -Path (Join-Path $newAppFolder "Log")     -ItemType Directory | Out-Null
                    New-Item -Path (Join-Path $newAppFolder "Archive") -ItemType Directory | Out-Null

                    # --- Copy source file (if provided) ---
                    if (-not [string]::IsNullOrWhiteSpace($wizardState.SourceFile) -and (Test-Path $wizardState.SourceFile)) {
                        $s3StatusLbl.Text = "Copying source file..."
                        $destFile = Join-Path $newAppFolder (Split-Path $wizardState.SourceFile -Leaf)
                        Copy-Item -Path $wizardState.SourceFile -Destination $destFile
                        & $writeLog "Copied source file to: $destFile" "INFO"
                    }

                    # --- Write config.json ---
                    $s3StatusLbl.Text = "Writing config.json..."
                    $cfgObj = [ordered]@{
                        sourceID           = $newSourceId
                        disableField       = ""
                        disableValue       = @("")
                        groupTypes         = ""
                        groupDelimiter     = ""
                        isUpload           = $false
                        headerRow          = 1
                        trimTopRows        = 0
                        trimBottomRows     = 0
                        trimLeftColumns    = 0
                        trimRightColumns   = 0
                        dropColumns        = ""
                        columnsToMerge     = ""
                        mergedColumnName   = ""
                        mergeDelimiter     = ""
                        adminColumnName    = ""
                        adminColumnValue   = ""
                        schema             = ""
                        booleanColumnList  = ""
                        booleanColumnValue = ""
                        sheetNumber        = 1
                    }
                    $cfgObj | ConvertTo-Json -Depth 3 | Out-File -FilePath (Join-Path $newAppFolder "config.json") -Encoding utf8

                    & $writeLog "Source creation wizard completed: '$($wizardState.SourceName)' (ID: $newSourceId)" "INFO"

                    $s3StatusLbl.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#2E7D32")
                    $s3StatusLbl.Text       = "Source '$($wizardState.SourceName)' created successfully. (ID: $newSourceId)"
                    $wzNextBtn.IsEnabled    = $false
                    $wzCancelBtn.IsEnabled  = $true
                    $wzCancelBtn.Content    = "Close"

                    # Refresh the app dropdown to include the new source
                    & $loadAppFoldersFn
                }
                catch {
                    $errMsg = $_.ToString()
                    # Try to extract a friendly message from the JSON error body
                    try {
                        $errJson = $_ | ConvertFrom-Json -ErrorAction Stop
                        $firstMsg = $errJson.messages | Where-Object { $_.locale -eq 'en-US' } | Select-Object -First 1
                        if ($firstMsg) { $errMsg = $firstMsg.text }
                        elseif ($errJson.detailCode) { $errMsg = $errJson.detailCode }
                    } catch {}
                    $s3StatusLbl.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#C62828")
                    $s3StatusLbl.Text = "Error: $errMsg"
                    & $writeLog "ERROR: Source creation failed - $errMsg" "ERROR"
                    $wzNextBtn.IsEnabled   = $true
                    $wzBackBtn.IsEnabled   = $true
                    $wzCancelBtn.IsEnabled = $true
                }
            }
        }.GetNewClosure())

        $wizardWindow.ShowDialog() | Out-Null
    })

    # Function to load app folders into dropdown
    $loadAppFolders = {
        $script:appDropdown.Items.Clear()
        $script:selectedAppName = $null
        $appFolderPath = $script:textBoxes['AppFolder'].Text
        
        # Clear config editor
        $script:configFieldsPanel.Children.Clear()
        $script:configFields.Clear()
        # Show empty-state hint
        if ($null -ne $script:emptyStatePanel) { $script:configFieldsPanel.Children.Add($script:emptyStatePanel) | Out-Null }
        $saveConfigButton.IsEnabled = $false
        $reloadConfigButton.IsEnabled = $false
        $processOnlyButton.IsEnabled = $false
        $matchFilesButton.IsEnabled = $false
        $uploadAppButton.IsEnabled = $false
        $uploadUserListButton.IsEnabled = $false
        $openAppLogButton.IsEnabled = $false
        $openAppFolderButton.IsEnabled = $false
        $uploadSchemaButton.IsEnabled = $false
        $resetSourceButton.IsEnabled = $false
        $script:appLogTextBox.Text = "Select an app to view log information..."
        $script:statusLabel.Text = "Ready"
        
        if ([string]::IsNullOrWhiteSpace($appFolderPath) -or -not (Test-Path $appFolderPath)) {
            $script:appDropdown.IsEnabled = $false
            return
        }

        try {
            $appDirs = Get-ChildItem -Path $appFolderPath -Directory | Sort-Object Name
            
            if ($appDirs.Count -eq 0) {
                $script:appDropdown.IsEnabled = $false
                $script:selectedAppLabel.Content = "No app folders found in: $appFolderPath"
            }
            else {
                # Populate dropdown with app folders that have config.json
                $appsWithConfig = @()
                foreach ($dir in $appDirs) {
                    $configPath = Join-Path $dir.FullName "config.json"
                    if (Test-Path $configPath) {
                        $script:appDropdown.Items.Add($dir.Name) | Out-Null
                        $appsWithConfig += $dir.Name
                    }
                }
                
                if ($script:appDropdown.Items.Count -gt 0) {
                    $script:appDropdown.IsEnabled = $true
                }
                else {
                    $script:appDropdown.IsEnabled = $false
                }
            }
        }
        catch {
            $script:appDropdown.IsEnabled = $false
        }
    }

    # Dropdown SelectionChanged handler - Load config when app is selected
    $script:appDropdown.Add_SelectionChanged({
        if ($script:appDropdown.SelectedItem) {
            $folderName = $script:appDropdown.SelectedItem
            $script:selectedAppName = $folderName
            
            try {
                $appFolderPath = $script:textBoxes['AppFolder'].Text
                
                if ([string]::IsNullOrWhiteSpace($appFolderPath)) {
                    throw "App Folder path is not configured. Please set it in Settings tab."
                }
                
                $configFilePath = Join-Path -Path $appFolderPath -ChildPath "$folderName\config.json"
                $logFolderPath = Join-Path -Path $appFolderPath -ChildPath "$folderName\Log"
                
                if (-not (Test-Path $configFilePath)) {
                    throw "Config file not found: $configFilePath"
                }
                
                $configContent = Get-Content -Path $configFilePath -Raw
                $configJson = $configContent | ConvertFrom-Json
                
                # Clear existing fields
                $script:configFieldsPanel.Children.Clear()
                $script:configFields.Clear()
                $script:statusLabel.Text = "App: $folderName"
                
                # Define field metadata with friendly names, tooltips, and grouping
                $fieldMetadata = @{
                    'sourceID' = @{ Group='Source Configuration'; Label='Source ID'; Tooltip='UUID value from SailPoint connection settings URL - uniquely identifies the integration source' }
                    'isUpload' = @{ Group='Source Configuration'; Label='Enable Upload'; Tooltip='Whether to upload processed data to SailPoint after processing' }
                    'schema' = @{ Group='Source Configuration'; Label='Schema'; Tooltip='Optional schema definition for the source' }
                    
                    'headerRow' = @{ Group='File Structure'; Label='Header Row'; Tooltip='The row number where column headers start (1-based index)' }
                    'sheetNumber' = @{ Group='File Structure'; Label='Sheet Number'; Tooltip='Which worksheet to process in Excel files (1-based index, default: 1)' }
                    'trimTopRows' = @{ Group='File Structure'; Label='Trim Top Rows'; Tooltip='Number of rows to remove AFTER the header row' }
                    'trimBottomRows' = @{ Group='File Structure'; Label='Trim Bottom Rows'; Tooltip='Number of rows to remove from the bottom of the file' }
                    'trimLeftColumns' = @{ Group='File Structure'; Label='Trim Left Columns'; Tooltip='Number of leftmost columns to remove' }
                    'trimRightColumns' = @{ Group='File Structure'; Label='Trim Right Columns'; Tooltip='Number of rightmost columns to remove' }
                    'dropColumns' = @{ Group='File Structure'; Label='Drop Columns'; Tooltip='Comma-separated list of columns to remove from processing (e.g., "Email,PhoneNumber")' }
                    
                    'columnsToMerge' = @{ Group='Column Operations'; Label='Columns to Merge'; Tooltip='Comma-separated columns to merge into a new column (e.g., "FirstName,LastName")' }
                    'mergedColumnName' = @{ Group='Column Operations'; Label='Merged Column Name'; Tooltip='Name of the new merged column (e.g., "FullName")' }
                    'mergeDelimiter' = @{ Group='Column Operations'; Label='Merge Delimiter'; Tooltip='Separator placed between merged column values (e.g., " ", "-", " - ", ", "). Leave blank to use a single space.' }
                    
                    'disableField' = @{ Group='User Status & Roles'; Label='Disable Field'; Tooltip='Column name used to determine if a user should be disabled (e.g., "Status")' }
                    'disableValue' = @{ Group='User Status & Roles'; Label='Disable Values'; Tooltip='Comma-separated values that indicate an inactive user (e.g., "Inactive,Terminated")' }
                    'adminColumnName' = @{ Group='User Status & Roles'; Label='Admin Column Name'; Tooltip='Column used to identify admin users (e.g., "Role")' }
                    'adminColumnValue' = @{ Group='User Status & Roles'; Label='Admin Column Value'; Tooltip='Value that indicates an admin user (e.g., "Admin", "SuperUser")' }
                    
                    'groupTypes' = @{ Group='Entitlements & Groups'; Label='Group Types Column'; Tooltip='Column(s) containing entitlement or group data. If blank, defaults to "Role" column' }
                    'groupDelimiter' = @{ Group='Entitlements & Groups'; Label='Group Delimiter'; Tooltip='Separator used in group columns if multiple values exist (e.g., "," or "|")' }
                    'booleanColumnList' = @{ Group='Entitlements & Groups'; Label='Boolean Columns'; Tooltip='Comma-separated list of boolean entitlement columns (e.g., "Entitlement1,Entitlement2")' }
                    'booleanColumnValue' = @{ Group='Entitlements & Groups'; Label='Boolean True Value'; Tooltip='Value that indicates true in boolean columns (e.g., "Y", "Yes", "1")' }
                }
                
                # Group fields by category
                $groups = @{}
                $configJson.PSObject.Properties | ForEach-Object {
                    $propName = $_.Name
                    $propValue = $_.Value
                    
                    # Determine group (use 'Other' if not in metadata)
                    $groupName = if ($fieldMetadata.ContainsKey($propName)) {
                        $fieldMetadata[$propName].Group
                    } else {
                        'Other'
                    }
                    
                    if (-not $groups.ContainsKey($groupName)) {
                        $groups[$groupName] = @()
                    }
                    
                    $groups[$groupName] += @{
                        Name = $propName
                        Value = $propValue
                        Metadata = if ($fieldMetadata.ContainsKey($propName)) { $fieldMetadata[$propName] } else { @{ Label=$propName; Tooltip=$propName } }
                    }
                }
                
                # Helper: create a label + control stack panel for a config field
                $makeFieldStack = {
                    param($pName, $pValue, $pMeta)
                    $fs = New-Object Windows.Controls.StackPanel
                    $fs.Margin = '0,0,10,12'
                    $lbl = New-Object Windows.Controls.Label
                    $lbl.Content = $pMeta.Label
                    $lbl.FontWeight = 'SemiBold'
                    $lbl.FontSize = 12
                    $lbl.Margin = '0,0,0,2'
                    $lbl.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
                    $lbl.ToolTip = $pMeta.Tooltip
                    $fs.Children.Add($lbl) | Out-Null
                    if ($pValue -is [bool]) {
                        $ctrl = New-Object Windows.Controls.CheckBox
                        $ctrl.IsChecked = $pValue
                        $ctrl.ToolTip = $pMeta.Tooltip
                        $fs.Children.Add($ctrl) | Out-Null
                        $script:configFields[$pName] = $ctrl
                    }
                    elseif ($pValue -is [array]) {
                        $ctrl = New-Object Windows.Controls.TextBox
                        $ctrl.Text = if ($pValue -and $pValue.Count -gt 0) { $pValue -join ', ' } else { "" }
                        $ctrl.Padding = '6,4'
                        $ctrl.FontSize = 12
                        $ctrl.Background = [System.Windows.Media.Brushes]::White
                        $ctrl.ToolTip = $pMeta.Tooltip + " (comma-separated values)"
                        $fs.Children.Add($ctrl) | Out-Null
                        $script:configFields[$pName] = $ctrl
                    }
                    else {
                        $ctrl = New-Object Windows.Controls.TextBox
                        $ctrl.Text = if ($null -ne $pValue) { $pValue.ToString() } else { "" }
                        $ctrl.Padding = '6,4'
                        $ctrl.FontSize = 12
                        $ctrl.Background = [System.Windows.Media.Brushes]::White
                        $ctrl.ToolTip = $pMeta.Tooltip
                        if ($pName -eq 'sourceID') {
                            $ctrl.IsReadOnly = $true
                            $ctrl.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#E8E8E8")
                            $ctrl.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#666666")
                            $ctrl.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
                            $ctrl.BorderThickness = '1'
                            $ctrl.FontStyle = 'Italic'
                            $ctrl.ToolTip = $pMeta.Tooltip + " (Read-only - managed by system)"
                        }
                        $fs.Children.Add($ctrl) | Out-Null
                        $script:configFields[$pName] = $ctrl
                    }
                    return $fs
                }

                # Explicit paired row layouts per group (fields on the same row)
                $groupRowLayouts = @{
                    'Source Configuration' = @(
                        @{ Fields = @('sourceID');                         FullWidth = $true }
                        @{ Fields = @('isUpload', 'schema') }
                    )
                    'File Structure' = @(
                        @{ Fields = @('headerRow',       'sheetNumber') }
                        @{ Fields = @('trimTopRows',     'trimBottomRows') }
                        @{ Fields = @('trimLeftColumns', 'trimRightColumns') }
                        @{ Fields = @('dropColumns');                      FullWidth = $true }
                    )
                    'Column Operations' = @(
                        @{ Fields = @('columnsToMerge', 'mergedColumnName') }
                        @{ Fields = @('mergeDelimiter') }
                    )
                    'User Status & Roles' = @(
                        @{ Fields = @('disableField',    'disableValue') }
                        @{ Fields = @('adminColumnName', 'adminColumnValue') }
                    )
                    'Entitlements & Groups' = @(
                        @{ Fields = @('groupTypes',       'groupDelimiter') }
                        @{ Fields = @('booleanColumnList','booleanColumnValue') }
                    )
                }

                # Define group order
                $groupOrder = @('Source Configuration', 'File Structure', 'Column Operations', 'User Status & Roles', 'Entitlements & Groups', 'Other')

                # Create UI for each group
                foreach ($groupName in $groupOrder) {
                    if (-not $groups.ContainsKey($groupName)) { continue }

                    $groupHeader = New-Object Windows.Controls.Label
                    $groupHeader.Content = $groupName
                    $groupHeader.FontWeight = 'Bold'
                    $groupHeader.FontSize = 12
                    $groupHeader.Margin = '0,15,0,8'
                    $groupHeader.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
                    $script:configFieldsPanel.Children.Add($groupHeader)

                    $separator = New-Object Windows.Controls.Separator
                    $separator.Margin = '0,0,0,10'
                    $script:configFieldsPanel.Children.Add($separator)

                    $groupGrid = New-Object Windows.Controls.Grid
                    $groupGrid.Margin = '0,0,0,0'
                    $gc1 = New-Object Windows.Controls.ColumnDefinition
                    $gc1.Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
                    $gc2 = New-Object Windows.Controls.ColumnDefinition
                    $gc2.Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
                    $groupGrid.ColumnDefinitions.Add($gc1)
                    $groupGrid.ColumnDefinitions.Add($gc2)

                    if ($groupRowLayouts.ContainsKey($groupName)) {
                        # Build name->field lookup
                        $byName = @{}
                        foreach ($f in $groups[$groupName]) { $byName[$f.Name] = $f }

                        $rowIdx = 0
                        foreach ($rowDef in $groupRowLayouts[$groupName]) {
                            $rowFields = @($rowDef.Fields | Where-Object { $byName.ContainsKey($_) })
                            if ($rowFields.Count -eq 0) { continue }

                            $rd = New-Object Windows.Controls.RowDefinition
                            $rd.Height = [Windows.GridLength]::Auto
                            $groupGrid.RowDefinitions.Add($rd)

                            $fullWidth = $rowDef.FullWidth -eq $true
                            $colIdx = 0
                            foreach ($fname in $rowFields) {
                                $f = $byName[$fname]
                                $fs = & $makeFieldStack $f.Name $f.Value $f.Metadata
                                [Windows.Controls.Grid]::SetRow($fs, $rowIdx)
                                [Windows.Controls.Grid]::SetColumn($fs, $colIdx)
                                if ($fullWidth) { [Windows.Controls.Grid]::SetColumnSpan($fs, 2) }
                                $groupGrid.Children.Add($fs)
                                if (-not $fullWidth) { $colIdx++ }
                            }
                            $rowIdx++
                        }

                        # Any fields not in the layout: append 2-per-row
                        $layoutNames = @($groupRowLayouts[$groupName] | ForEach-Object { $_.Fields } | ForEach-Object { $_ })
                        $remaining = @($groups[$groupName] | Where-Object { $layoutNames -notcontains $_.Name })
                        $colIdx = 0
                        foreach ($field in $remaining) {
                            if ($colIdx -eq 0) {
                                $rd = New-Object Windows.Controls.RowDefinition
                                $rd.Height = [Windows.GridLength]::Auto
                                $groupGrid.RowDefinitions.Add($rd)
                            }
                            $fs = & $makeFieldStack $field.Name $field.Value $field.Metadata
                            [Windows.Controls.Grid]::SetRow($fs, $rowIdx)
                            [Windows.Controls.Grid]::SetColumn($fs, $colIdx)
                            $groupGrid.Children.Add($fs)
                            $colIdx++
                            if ($colIdx -ge 2) { $colIdx = 0; $rowIdx++ }
                        }
                    }
                    else {
                        # Fallback: sequential 2-per-row (for 'Other' group)
                        $rowIdx = 0
                        $colIdx = 0
                        foreach ($field in $groups[$groupName]) {
                            if ($colIdx -eq 0) {
                                $rd = New-Object Windows.Controls.RowDefinition
                                $rd.Height = [Windows.GridLength]::Auto
                                $groupGrid.RowDefinitions.Add($rd)
                            }
                            $fs = & $makeFieldStack $field.Name $field.Value $field.Metadata
                            [Windows.Controls.Grid]::SetRow($fs, $rowIdx)
                            [Windows.Controls.Grid]::SetColumn($fs, $colIdx)
                            $groupGrid.Children.Add($fs)
                            $colIdx++
                            if ($colIdx -ge 2) { $colIdx = 0; $rowIdx++ }
                        }
                    }

                    $script:configFieldsPanel.Children.Add($groupGrid)
                }
                
                $saveConfigButton.IsEnabled = $true
                $reloadConfigButton.IsEnabled = $true
                $processOnlyButton.IsEnabled = $true
                $matchFilesButton.IsEnabled = $true
                $uploadAppButton.IsEnabled = $true
                $uploadUserListButton.IsEnabled = $true
                $script:currentConfigPath = $configFilePath
                
                # Enable app log button and app folder button
                $openAppLogButton.IsEnabled = $true
                $script:currentAppLogFolder = $logFolderPath
                $openAppFolderButton.IsEnabled = $true
                $uploadSchemaButton.IsEnabled = $true
                $resetSourceButton.IsEnabled = $true
                $script:currentAppFolderPath = Join-Path $appFolderPath $folderName

                # Populate log file selector and load most recent
                if (Test-Path $logFolderPath) {
                    & $reloadAppLogs
                }
                else {
                    $script:appLogFileDropdown.Items.Clear()
                    $script:appLogFileDropdown.IsEnabled = $false
                    $script:appLogTextBox.Text = "Log folder not found"
                }
            }
            catch {
                $script:configFieldsPanel.Children.Clear()
                if ($null -ne $script:emptyStatePanel) { $script:configFieldsPanel.Children.Add($script:emptyStatePanel) | Out-Null }
                $saveConfigButton.IsEnabled = $false
                $reloadConfigButton.IsEnabled = $false
                $processOnlyButton.IsEnabled = $false
                $uploadAppButton.IsEnabled = $false
                $uploadUserListButton.IsEnabled = $false
                $openAppLogButton.IsEnabled = $false
                $openAppFolderButton.IsEnabled = $false
                $uploadSchemaButton.IsEnabled = $false
                $resetSourceButton.IsEnabled = $false
                $script:appLogFileDropdown.Items.Clear()
                $script:appLogFileDropdown.IsEnabled = $false
                $script:appLogTextBox.Text = "Error loading app: $_"
                $script:statusLabel.Text = "Error loading app"
            }
        }
    })

    # Upload App Button Handler
    $uploadAppButton.Add_Click({
        if ($script:selectedAppName) {
            try {
                Write-Log "Starting upload for app: $script:selectedAppName" "INFO"
                $currentSettings = Load-Settings
                
                if ($null -eq $currentSettings) {
                    [System.Windows.MessageBox]::Show("Failed to load settings!", "Upload Error", "OK", "Error")
                    return
                }

                # --- Schema validation: block upload if ISC schema doesn't exist or doesn't match the processed file ---
                $archivePath = Join-Path $script:currentAppFolderPath "Archive"
                if (Test-Path $archivePath) {
                    $uploadFiles = Get-ChildItem -Path $archivePath -Filter "*_upload_file_*.csv" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
                    if ($uploadFiles.Count -gt 0) {
                        $latestFile = $uploadFiles[0]
                        $sourceID = $null
                        if ($script:configFields.ContainsKey('sourceID')) {
                            $sourceID = $script:configFields['sourceID'].Text.Trim()
                        }
                        if (-not [string]::IsNullOrWhiteSpace($sourceID)) {
                            try {
                                Write-Log "Validating ISC schema before upload for source $sourceID..." "INFO"
                                $fileRows = Import-Csv -Path $latestFile.FullName
                                $fileHeaders = if ($fileRows.Count -gt 0) { @($fileRows[0].PSObject.Properties.Name) } else { @() }

                                if ($fileHeaders.Count -gt 0) {
                                    $tenantUrl = Get-BaseApiUrl -Settings $currentSettings
                                    $authHdr = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($currentSettings.ClientID):$($currentSettings.ClientSecret)"))
                                    $tokenResp = Invoke-RestMethod -Uri "$tenantUrl/oauth/token" -Method Post `
                                        -Headers @{ Authorization = "Basic $authHdr" } `
                                        -Body @{ grant_type = 'client_credentials' }
                                    $iscHeaders = @{ Authorization = "Bearer $($tokenResp.access_token)"; Accept = 'application/json' }
                                    $schemas = Invoke-RestMethod -Uri "$tenantUrl/v2025/sources/$sourceID/schemas" -Method Get -Headers $iscHeaders
                                    $accountSchema = $schemas | Where-Object { $_.name -eq 'account' } | Select-Object -First 1

                                    if ($null -eq $accountSchema) {
                                        Write-Log "Upload blocked: no account schema in ISC for source $sourceID" "WARNING"
                                        [System.Windows.MessageBox]::Show(
                                            "Upload blocked: no account schema found in ISC for this source.`n`nPlease click 'Upload Schema' to push the schema to SailPoint before uploading data.",
                                            "Schema Required", "OK", "Warning")
                                        return
                                    }

                                    $schemaAttrNames = @($accountSchema.attributes | ForEach-Object { $_.name })
                                    $missingAttrs = @($fileHeaders | Where-Object { $schemaAttrNames -notcontains $_ })
                                    if ($missingAttrs.Count -gt 0) {
                                        Write-Log "Upload blocked: processed file columns not in ISC schema: $($missingAttrs -join ', ')" "WARNING"
                                        [System.Windows.MessageBox]::Show(
                                            "Upload blocked: the following columns in the processed file are not in the ISC account schema:`n`n$($missingAttrs -join ', ')`n`nPlease click 'Upload Schema' to update the schema before uploading data.",
                                            "Schema Mismatch", "OK", "Warning")
                                        return
                                    }
                                    Write-Log "Schema validation passed: all $($fileHeaders.Count) columns present in ISC schema." "INFO"
                                }
                            }
                            catch {
                                Write-Log "WARNING: Schema pre-check failed (will proceed with upload): $_" "WARNING"
                                # Non-blocking: if the API call itself fails (network, auth), let the upload continue
                            }
                        }
                    }
                }

                # Show confirmation dialog
                $result = [System.Windows.MessageBox]::Show("Upload files for app '$script:selectedAppName'?", "Confirm Upload", "YesNo", "Question")
                
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    Write-Log "User confirmed upload for $script:selectedAppName" "INFO"
                    $success = Run-SingleAppUpload -settings $currentSettings -appName $script:selectedAppName
                    
                    # Refresh logs after upload
                    & $reloadAppLogs
                    & $loadExecutionLogDates
                    
                    if ($success) {
                        [System.Windows.MessageBox]::Show("Upload completed for $script:selectedAppName. Check logs for details.", "Upload Complete", "OK", "Information")
                    }
                    else {
                        [System.Windows.MessageBox]::Show("Upload failed for $script:selectedAppName. Check logs for details.", "Upload Failed", "OK", "Warning")
                    }
                }
                else {
                    Write-Log "Upload cancelled by user for $script:selectedAppName" "INFO"
                }
            }
            catch {
                Write-Log "ERROR during upload for $script:selectedAppName : $_" "ERROR"
                [System.Windows.MessageBox]::Show("Upload error: $_", "Upload Error", "OK", "Error")
            }
        }
    })

    # Process Only Button Handler
    $processOnlyButton.Add_Click({
        if ($script:selectedAppName) {
            try {
                Write-Log "Starting process-only run for app: $script:selectedAppName" "INFO"
                $currentSettings = Load-Settings

                if ($null -eq $currentSettings) {
                    [System.Windows.MessageBox]::Show("Failed to load settings!", "Process Error", "OK", "Error")
                    return
                }

                $result = [System.Windows.MessageBox]::Show(
                    "Process files for app '$script:selectedAppName' without uploading to SailPoint?`n`nThe processed file will be created in the app folder so you can review it before uploading.",
                    "Confirm Process Only", "YesNo", "Question")

                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    Write-Log "User confirmed process-only run for $script:selectedAppName" "INFO"
                    $success = Run-SingleAppUpload -settings $currentSettings -appName $script:selectedAppName -ProcessOnly

                    # Refresh logs after run
                    & $reloadAppLogs
                    & $loadExecutionLogDates

                    if ($success) {
                        [System.Windows.MessageBox]::Show("Processing completed for $script:selectedAppName. No upload was performed.`n`nCheck the app log to review the processed file output.", "Process Complete", "OK", "Information")
                    }
                    else {
                        [System.Windows.MessageBox]::Show("Processing failed for $script:selectedAppName. Check logs for details.", "Process Failed", "OK", "Warning")
                    }
                }
                else {
                    Write-Log "Process-only run cancelled by user for $script:selectedAppName" "INFO"
                }
            }
            catch {
                Write-Log "ERROR during process-only run for $script:selectedAppName : $_" "ERROR"
                [System.Windows.MessageBox]::Show("Process error: $_", "Process Error", "OK", "Error")
            }
        }
    })

    # Match Files Button Handler
    $matchFilesButton.Add_Click({
        if ($script:selectedAppName) {
            try {
                $currentSettings = Load-Settings

                if ($null -eq $currentSettings) {
                    [System.Windows.MessageBox]::Show("Failed to load settings!", "Error", "OK", "Error")
                    return
                }

                Run-FileMatchOnly -settings $currentSettings -appName $script:selectedAppName

                # Refresh logs after run
                & $reloadAppLogs
            }
            catch {
                Write-Log "ERROR during file match for $script:selectedAppName : $_" "ERROR"
                [System.Windows.MessageBox]::Show("File match error: $_", "Error", "OK", "Error")
            }
        }
    })

    # Upload User List Button Handler
    $uploadUserListButton.Add_Click({
        if ($script:selectedAppName) {
            try {
                $appFolderPath = $script:textBoxes['AppFolder'].Text
                
                if ([string]::IsNullOrWhiteSpace($appFolderPath)) {
                    [System.Windows.MessageBox]::Show("App Folder path is not configured. Please set it in Settings tab.", "Upload Error", "OK", "Error")
                    return
                }
                
                $targetAppFolder = Join-Path -Path $appFolderPath -ChildPath $script:selectedAppName
                
                if (-not (Test-Path $targetAppFolder)) {
                    [System.Windows.MessageBox]::Show("App folder not found: $targetAppFolder", "Upload Error", "OK", "Error")
                    return
                }
                
                # Open file dialog to select user list file
                $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
                $openFileDialog.Title = "Select User List File"
                $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
                $openFileDialog.FilterIndex = 1
                
                if ($openFileDialog.ShowDialog() -eq $true) {
                    $sourceFile = $openFileDialog.FileName
                    $fileName = [System.IO.Path]::GetFileName($sourceFile)
                    $destinationFile = Join-Path -Path $targetAppFolder -ChildPath $fileName
                    
                    # Ask for confirmation if file already exists
                    if (Test-Path $destinationFile) {
                        $overwriteResult = [System.Windows.MessageBox]::Show("File '$fileName' already exists in the app folder. Do you want to overwrite it?", "File Exists", "YesNo", "Warning")
                        if ($overwriteResult -ne [System.Windows.MessageBoxResult]::Yes) {
                            Write-Log "User cancelled overwrite of existing file: $fileName" "INFO"
                            return
                        }
                    }
                    
                    # Copy the file
                    Copy-Item -Path $sourceFile -Destination $destinationFile -Force
                    Write-Log "User list file uploaded: $fileName to $script:selectedAppName" "INFO"
                    [System.Windows.MessageBox]::Show("User list file '$fileName' uploaded successfully to app folder!", "Upload Success", "OK", "Information")
                }
            }
            catch {
                Write-Log "ERROR uploading user list file: $_" "ERROR"
                [System.Windows.MessageBox]::Show("Failed to upload user list file!`n`nError: $_", "Upload Error", "OK", "Error")
            }
        }
    })

    # Refresh apps button handler
    $refreshAppsButton.Add_Click({
        & $loadAppFolders
    })

    # Save config button handler
    $saveConfigButton.Add_Click({
        if ($null -ne $script:currentConfigPath -and $script:configFields.Count -gt 0) {
            try {
                # Build JSON object from form fields
                $configObj = [ordered]@{}

                # Fields whose values must always be persisted as plain strings.
                # These can legitimately contain commas (e.g. groupDelimiter = ",") so
                # the generic comma-split-to-array logic must NOT apply to them.
                $scalarStringFields = @(
                    'sourceID', 'disableField',
                    'groupDelimiter', 'mergeDelimiter', 'mergedColumnName',
                    'adminColumnName', 'adminColumnValue', 'booleanColumnValue'
                )
                
                foreach ($key in $script:configFields.Keys) {
                    $control = $script:configFields[$key]
                    
                    if ($control -is [Windows.Controls.CheckBox]) {
                        $configObj[$key] = $control.IsChecked
                    }
                    elseif ($control -is [Windows.Controls.TextBox]) {
                        $value = $control.Text.Trim()
                        
                        # Scalar string fields: always saved as-is (never split on commas)
                        if ($scalarStringFields -contains $key) {
                            $configObj[$key] = $value
                        }
                        # Try to parse as number
                        elseif ([int]::TryParse($value, [ref]([int]$numValue))) {
                            $configObj[$key] = $numValue
                        }
                        # Check if it's an array (comma-separated)
                        elseif ($value.Contains(',')) {
                            $configObj[$key] = @($value -split ',' | ForEach-Object { $_.Trim() })
                        }
                        # Check if it's an empty array indicator
                        elseif ($value -eq '' -or $value -eq '[]') {
                            $configObj[$key] = @()
                        }
                        else {
                            $configObj[$key] = $value
                        }
                    }
                }
                
                # Convert to JSON and save
                $jsonContent = $configObj | ConvertTo-Json -Depth 10
                $jsonContent | Out-File -FilePath $script:currentConfigPath -Encoding UTF8
                [System.Windows.MessageBox]::Show("Config saved successfully!", "Success", "OK", "Information")
                Write-Log "Saved config: $script:currentConfigPath" "INFO"
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to save config!`n`nError: $_", "Save Error", "OK", "Error")
            }
        }
    })

    # Reload config button handler
    $reloadConfigButton.Add_Click({
        if ($null -ne $script:currentConfigPath -and (Test-Path $script:currentConfigPath)) {
            try {
                $configContent = Get-Content -Path $script:currentConfigPath -Raw
                $configJson = $configContent | ConvertFrom-Json
                
                # Reload form fields with saved values
                foreach ($key in $script:configFields.Keys) {
                    $control = $script:configFields[$key]
                    $value = $configJson.$key
                    
                    if ($control -is [Windows.Controls.CheckBox]) {
                        $control.IsChecked = $value
                    }
                    elseif ($control -is [Windows.Controls.TextBox]) {
                        if ($value -is [array]) {
                            $control.Text = ($value -join ', ')
                        }
                        else {
                            $control.Text = $value.ToString()
                        }
                    }
                }
                
                Write-Log "Reloaded config from disk" "INFO"
            }
            catch {
                [System.Windows.MessageBox]::Show("Error reloading config: $_", "Error", "OK", "Error")
            }
        }
    })

    # Open app log folder button handler
    $openAppLogButton.Add_Click({
        if ($script:currentAppLogFolder -and (Test-Path $script:currentAppLogFolder)) {
            Start-Process explorer.exe -ArgumentList $script:currentAppLogFolder
        }
        else {
            [System.Windows.MessageBox]::Show("App log folder not found.", "Error", "OK", "Error")
        }
    })

    # Open app folder button handler
    $openAppFolderButton.Add_Click({
        if ($script:currentAppFolderPath -and (Test-Path $script:currentAppFolderPath)) {
            Start-Process explorer.exe -ArgumentList $script:currentAppFolderPath
        }
        else {
            [System.Windows.MessageBox]::Show("App folder not found.", "Error", "OK", "Error")
        }
    })

    # Upload Schema Button Handler
    $uploadSchemaButton.Add_Click({
        if (-not $script:selectedAppName) { return }

        try {
            # Get sourceID from current config fields
            $sourceID = $null
            if ($script:configFields.ContainsKey('sourceID')) {
                $sourceID = $script:configFields['sourceID'].Text.Trim()
            }
            if ([string]::IsNullOrWhiteSpace($sourceID)) {
                [System.Windows.MessageBox]::Show("No Source ID found for this app. Please ensure a config is loaded.", "Upload Schema", "OK", "Warning")
                return
            }

            # Find Archive folder
            $archivePath = Join-Path $script:currentAppFolderPath "Archive"
            if (-not (Test-Path $archivePath)) {
                [System.Windows.MessageBox]::Show("Archive folder not found:`n$archivePath`n`nPlease run 'Process Only' or 'Upload Files' first to generate a processed file.", "Upload Schema", "OK", "Warning")
                return
            }

            # Find most recent upload file
            $uploadFiles = Get-ChildItem -Path $archivePath -Filter "*_upload_file_*.csv" -File | Sort-Object LastWriteTime -Descending
            if ($uploadFiles.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No processed upload files found in:`n$archivePath`n`nPlease run 'Process Only' or 'Upload Files' first to generate a processed file.", "Upload Schema", "OK", "Warning")
                return
            }

            $latestFile = $uploadFiles[0]
            Write-Log "Reading schema columns from: $($latestFile.Name)" "INFO"

            # Read all rows - needed both for headers and multi-valued column detection
            $allRows = @(Import-Csv -Path $latestFile.FullName)
            if ($allRows.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No data rows found in:`n$($latestFile.Name)", "Upload Schema", "OK", "Warning")
                return
            }
            $headers = @($allRows[0].PSObject.Properties.Name)
            if ($headers.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No column headers found in:`n$($latestFile.Name)", "Upload Schema", "OK", "Warning")
                return
            }

            # Authenticate with SailPoint and fetch identity attributes before opening wizard
            $currentSettings = Load-Settings
            $tenantUrl = Get-BaseApiUrl -Settings $currentSettings
            if ($null -eq $tenantUrl) {
                [System.Windows.MessageBox]::Show("Cannot determine API URL. Please check Settings.", "Upload Schema", "OK", "Error")
                return
            }

            $authHeader = [System.Convert]::ToBase64String(
                [System.Text.Encoding]::ASCII.GetBytes("$($currentSettings.ClientID):$($currentSettings.ClientSecret)")
            )

            Write-Log "Retrieving OAuth token for schema upload..." "INFO"
            $tokenResponse = Invoke-RestMethod -Uri "$tenantUrl/oauth/token" -Method Post `
                -Headers @{ Authorization = "Basic $authHeader" } `
                -Body @{ grant_type = "client_credentials" }
            $accessToken = $tokenResponse.access_token

            $apiHeaders = @{
                Authorization = "Bearer $accessToken"
                Accept        = "application/json"
            }

            # Fetch identity attributes from SailPoint API for correlation step
            Write-Log "Fetching identity attributes from SailPoint..." "INFO"
            $allIdentityAttrs = @()
            try {
                $identAttrResp = Invoke-RestMethod -Uri "$tenantUrl/beta/identity-attributes?includeSystem=true&includeSilent=true" -Method Get -Headers $apiHeaders
                if ($identAttrResp -is [System.Array]) {
                    $allIdentityAttrs = @($identAttrResp | Sort-Object { if ($_.displayName) { $_.displayName } else { $_.name } })
                } elseif ($null -ne $identAttrResp -and $identAttrResp.PSObject.Properties['items']) {
                    $allIdentityAttrs = @($identAttrResp.items | Sort-Object { if ($_.displayName) { $_.displayName } else { $_.name } })
                }
                Write-Log "Fetched $($allIdentityAttrs.Count) identity attributes." "INFO"
            }
            catch {
                $errDetail = $_.ToString()
                try { $errDetail = ($_ | ConvertFrom-Json -ErrorAction Stop).messages[0].text } catch {}
                Write-Log "WARNING: Could not fetch identity attributes from API: $errDetail" "WARNING"
            }

            # === SCHEMA UPLOAD WIZARD (2 steps) ===
            $schemaWizard = New-Object Windows.Window
            $schemaWizard.Title = "Upload Schema - $script:selectedAppName"
            $schemaWizard.Width = 580
            $schemaWizard.Height = 560
            $schemaWizard.WindowStartupLocation = 'CenterOwner'
            $schemaWizard.Owner = $window
            $schemaWizard.ResizeMode = 'NoResize'
            $schemaWizard.Background = [System.Windows.Media.Brushes]::White
            $schemaWizard.FontFamily = New-Object Windows.Media.FontFamily("Segoe UI")

            $wzMain = New-Object Windows.Controls.DockPanel
            $schemaWizard.Content = $wzMain

            # --- Header ---
            $wzHeader = New-Object Windows.Controls.Border
            $wzHeader.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
            $wzHeader.Padding = '20,14'
            [Windows.Controls.DockPanel]::SetDock($wzHeader, 'Top')
            $wzMain.Children.Add($wzHeader) | Out-Null

            $wzHeaderDP = New-Object Windows.Controls.DockPanel
            $wzHeader.Child = $wzHeaderDP

            $wzStepLabel = New-Object Windows.Controls.TextBlock
            $wzStepLabel.Text = "Step 1 of 2"
            $wzStepLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#54C0E8")
            $wzStepLabel.FontSize = 11
            $wzStepLabel.VerticalAlignment = 'Center'
            [Windows.Controls.DockPanel]::SetDock($wzStepLabel, 'Right')
            $wzHeaderDP.Children.Add($wzStepLabel) | Out-Null

            $wzTitleText = New-Object Windows.Controls.TextBlock
            $wzTitleText.Text = "Upload Schema to SailPoint"
            $wzTitleText.Foreground = [System.Windows.Media.Brushes]::White
            $wzTitleText.FontSize = 17
            $wzTitleText.FontWeight = 'SemiBold'
            $wzHeaderDP.Children.Add($wzTitleText) | Out-Null

            # --- Footer ---
            $wzFooter = New-Object Windows.Controls.Border
            $wzFooter.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F0F4FA")
            $wzFooter.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
            $wzFooter.BorderThickness = '0,1,0,0'
            $wzFooter.Padding = '20,12'
            [Windows.Controls.DockPanel]::SetDock($wzFooter, 'Bottom')
            $wzMain.Children.Add($wzFooter) | Out-Null

            $wzFooterDP = New-Object Windows.Controls.DockPanel
            $wzFooterDP.LastChildFill = $false
            $wzFooter.Child = $wzFooterDP

            $wzCancelBtn = New-Object Windows.Controls.Button
            $wzCancelBtn.Content = "Cancel"
            $wzCancelBtn.Padding = '14,7'
            $wzCancelBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
            $wzCancelBtn.Foreground = [System.Windows.Media.Brushes]::White
            $wzCancelBtn.FontSize = 12
            $wzCancelBtn.BorderThickness = '0'
            $wzCancelBtn.Cursor = 'Hand'
            [Windows.Controls.DockPanel]::SetDock($wzCancelBtn, 'Left')
            $wzFooterDP.Children.Add($wzCancelBtn) | Out-Null

            $wzNextBtn = New-Object Windows.Controls.Button
            $wzNextBtn.Content = "Next >"
            $wzNextBtn.Padding = '14,7'
            $wzNextBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0033A1")
            $wzNextBtn.Foreground = [System.Windows.Media.Brushes]::White
            $wzNextBtn.FontWeight = 'SemiBold'
            $wzNextBtn.FontSize = 12
            $wzNextBtn.BorderThickness = '0'
            $wzNextBtn.Cursor = 'Hand'
            [Windows.Controls.DockPanel]::SetDock($wzNextBtn, 'Right')
            $wzFooterDP.Children.Add($wzNextBtn) | Out-Null

            $wzBackBtn = New-Object Windows.Controls.Button
            $wzBackBtn.Content = "< Back"
            $wzBackBtn.Padding = '14,7'
            $wzBackBtn.Margin = '0,0,8,0'
            $wzBackBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
            $wzBackBtn.Foreground = [System.Windows.Media.Brushes]::White
            $wzBackBtn.FontSize = 12
            $wzBackBtn.BorderThickness = '0'
            $wzBackBtn.Cursor = 'Hand'
            $wzBackBtn.IsEnabled = $false
            [Windows.Controls.DockPanel]::SetDock($wzBackBtn, 'Right')
            $wzFooterDP.Children.Add($wzBackBtn) | Out-Null

            # --- Content Area ---
            $wzContent = New-Object Windows.Controls.Border
            $wzContent.Padding = '25,20'
            $wzMain.Children.Add($wzContent) | Out-Null

            $stepsContainer = New-Object Windows.Controls.StackPanel
            $wzContent.Child = $stepsContainer

            # ---- STEP 1: Configure Schema ----
            $step1Panel = New-Object Windows.Controls.StackPanel

            $s1Title = New-Object Windows.Controls.TextBlock
            $s1Title.Text = "Configure Schema"
            $s1Title.FontSize = 15
            $s1Title.FontWeight = 'SemiBold'
            $s1Title.Margin = '0,0,0,4'
            $step1Panel.Children.Add($s1Title) | Out-Null

            $s1SubInfo = New-Object Windows.Controls.TextBlock
            $s1SubInfo.Text = "Source file: $($latestFile.Name)   |   $($headers.Count) columns detected across $($allRows.Count) data rows"
            $s1SubInfo.FontSize = 11
            $s1SubInfo.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#555555")
            $s1SubInfo.TextWrapping = 'Wrap'
            $s1SubInfo.Margin = '0,0,0,16'
            $step1Panel.Children.Add($s1SubInfo) | Out-Null

            $s1IdLbl = New-Object Windows.Controls.TextBlock
            $s1IdLbl.Text = "Identity Attribute (unique identifier column) *"
            $s1IdLbl.FontWeight = 'SemiBold'
            $s1IdLbl.FontSize = 12
            $s1IdLbl.Margin = '0,0,0,4'
            $step1Panel.Children.Add($s1IdLbl) | Out-Null

            $idAttrDropdown = New-Object Windows.Controls.ComboBox
            $idAttrDropdown.FontSize = 12
            $idAttrDropdown.Margin = '0,0,0,14'
            $idAttrDropdown.Padding = '6,4'
            foreach ($h in $headers) { $idAttrDropdown.Items.Add($h) | Out-Null }
            $idAttrDropdown.SelectedIndex = 0
            $step1Panel.Children.Add($idAttrDropdown) | Out-Null

            $s1DispLbl = New-Object Windows.Controls.TextBlock
            $s1DispLbl.Text = "Display Attribute (human-readable name column) *"
            $s1DispLbl.FontWeight = 'SemiBold'
            $s1DispLbl.FontSize = 12
            $s1DispLbl.Margin = '0,0,0,4'
            $step1Panel.Children.Add($s1DispLbl) | Out-Null

            $dispAttrDropdown = New-Object Windows.Controls.ComboBox
            $dispAttrDropdown.FontSize = 12
            $dispAttrDropdown.Margin = '0,0,0,8'
            $dispAttrDropdown.Padding = '6,4'
            foreach ($h in $headers) { $dispAttrDropdown.Items.Add($h) | Out-Null }
            $dispAttrDropdown.SelectedIndex = 0
            $step1Panel.Children.Add($dispAttrDropdown) | Out-Null

            $s1ValidationLbl = New-Object Windows.Controls.TextBlock
            $s1ValidationLbl.Text = ""
            $s1ValidationLbl.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#C62828")
            $s1ValidationLbl.FontSize = 11
            $s1ValidationLbl.Margin = '0,4,0,0'
            $step1Panel.Children.Add($s1ValidationLbl) | Out-Null

            $stepsContainer.Children.Add($step1Panel) | Out-Null

            # ---- STEP 2: Configure Correlation ----
            $step2Panel = New-Object Windows.Controls.StackPanel
            $step2Panel.Visibility = 'Collapsed'

            $s2Title = New-Object Windows.Controls.TextBlock
            $s2Title.Text = "Configure Correlation"
            $s2Title.FontSize = 15
            $s2Title.FontWeight = 'SemiBold'
            $s2Title.Margin = '0,0,0,4'
            $step2Panel.Children.Add($s2Title) | Out-Null

            $s2InfoText = New-Object Windows.Controls.TextBlock
            $s2InfoText.Text = "Map a source account attribute to a SailPoint identity attribute for account correlation."
            $s2InfoText.TextWrapping = 'Wrap'
            $s2InfoText.FontSize = 12
            $s2InfoText.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#415364")
            $s2InfoText.Margin = '0,0,0,14'
            $step2Panel.Children.Add($s2InfoText) | Out-Null

            # Two-column layout: source attr (left) | identity attr search + list (right)
            $corrOuterGrid = New-Object Windows.Controls.Grid
            $corrGC1 = New-Object Windows.Controls.ColumnDefinition
            $corrGC1.Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
            $corrGC2 = New-Object Windows.Controls.ColumnDefinition
            $corrGC2.Width = New-Object Windows.GridLength(16, [Windows.GridUnitType]::Pixel)
            $corrGC3 = New-Object Windows.Controls.ColumnDefinition
            $corrGC3.Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
            $corrOuterGrid.ColumnDefinitions.Add($corrGC1) | Out-Null
            $corrOuterGrid.ColumnDefinitions.Add($corrGC2) | Out-Null
            $corrOuterGrid.ColumnDefinitions.Add($corrGC3) | Out-Null

            # Left: Source Attribute
            $corrLeftPanel = New-Object Windows.Controls.StackPanel
            $corrLeftPanel.VerticalAlignment = 'Top'
            [Windows.Controls.Grid]::SetColumn($corrLeftPanel, 0)
            $corrOuterGrid.Children.Add($corrLeftPanel) | Out-Null

            $corrSrcLbl = New-Object Windows.Controls.TextBlock
            $corrSrcLbl.Text = "Source Attribute"
            $corrSrcLbl.FontWeight = 'SemiBold'
            $corrSrcLbl.FontSize = 12
            $corrSrcLbl.Margin = '0,0,0,4'
            $corrSrcLbl.ToolTip = "The column in the source file to match against the identity attribute"
            $corrLeftPanel.Children.Add($corrSrcLbl) | Out-Null

            $corrSrcDropdown = New-Object Windows.Controls.ComboBox
            $corrSrcDropdown.FontSize = 12
            $corrSrcDropdown.Padding = '6,4'
            $corrSrcDropdown.ToolTip = "The column in the source file to match against the identity attribute"
            $corrSrcDropdown.Items.Add("(skip)") | Out-Null
            foreach ($h in $headers) { $corrSrcDropdown.Items.Add($h) | Out-Null }
            $corrSrcDropdown.SelectedIndex = 0
            $corrLeftPanel.Children.Add($corrSrcDropdown) | Out-Null

            # Right: Identity Attribute with search
            $corrRightPanel = New-Object Windows.Controls.StackPanel
            [Windows.Controls.Grid]::SetColumn($corrRightPanel, 2)
            $corrOuterGrid.Children.Add($corrRightPanel) | Out-Null

            $corrIdLbl = New-Object Windows.Controls.TextBlock
            $corrIdLbl.Text = "Identity Attribute"
            $corrIdLbl.FontWeight = 'SemiBold'
            $corrIdLbl.FontSize = 12
            $corrIdLbl.Margin = '0,0,0,4'
            $corrIdLbl.ToolTip = "The SailPoint identity attribute to correlate against"
            $corrRightPanel.Children.Add($corrIdLbl) | Out-Null

            $identAttrSearchBox = New-Object Windows.Controls.TextBox
            $identAttrSearchBox.FontSize = 12
            $identAttrSearchBox.Padding = '6,4'
            $identAttrSearchBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0071CE")
            $identAttrSearchBox.BorderThickness = '1'
            $identAttrSearchBox.Margin = '0,0,0,4'
            $identAttrSearchBox.ToolTip = "Type to filter identity attributes"
            $corrRightPanel.Children.Add($identAttrSearchBox) | Out-Null

            $identityAttrListBox = New-Object Windows.Controls.ListBox
            $identityAttrListBox.FontSize = 12
            $identityAttrListBox.Height = 140
            $identityAttrListBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DAE1E9")
            $identityAttrListBox.BorderThickness = '1'
            $corrRightPanel.Children.Add($identityAttrListBox) | Out-Null

            $step2Panel.Children.Add($corrOuterGrid) | Out-Null

            $s2ValidationLbl = New-Object Windows.Controls.TextBlock
            $s2ValidationLbl.Text = ""
            $s2ValidationLbl.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#C62828")
            $s2ValidationLbl.FontSize = 11
            $s2ValidationLbl.Margin = '0,6,0,0'
            $step2Panel.Children.Add($s2ValidationLbl) | Out-Null

            $stepsContainer.Children.Add($step2Panel) | Out-Null

            # Populate identity attr listbox initially (no filter)
            foreach ($iAttr in $allIdentityAttrs) {
                $dText = if (-not [string]::IsNullOrWhiteSpace($iAttr.displayName)) {
                    "$($iAttr.displayName) ($($iAttr.name))"
                } else {
                    $iAttr.name
                }
                $lbi = New-Object Windows.Controls.ListBoxItem
                $lbi.Content = $dText
                $lbi.Tag = $iAttr.name
                $identityAttrListBox.Items.Add($lbi) | Out-Null
            }
            if ($identityAttrListBox.Items.Count -eq 0) {
                $lbi = New-Object Windows.Controls.ListBoxItem
                $lbi.Content = "(No identity attributes loaded from API)"
                $lbi.IsEnabled = $false
                $identityAttrListBox.Items.Add($lbi) | Out-Null
            } elseif ($identityAttrListBox.Items.Count -gt 0) {
                $identityAttrListBox.SelectedIndex = 0
            }

            # Search/filter handler for identity attr listbox
            $identAttrSearchBox.Add_TextChanged({
                $filterText = $identAttrSearchBox.Text
                $identityAttrListBox.Items.Clear()
                foreach ($iAttr in $allIdentityAttrs) {
                    $dText = if (-not [string]::IsNullOrWhiteSpace($iAttr.displayName)) {
                        "$($iAttr.displayName) ($($iAttr.name))"
                    } else {
                        $iAttr.name
                    }
                    if ([string]::IsNullOrEmpty($filterText) -or
                        $dText.ToLower().Contains($filterText.ToLower()) -or
                        $iAttr.name.ToLower().Contains($filterText.ToLower())) {
                        $lbi = New-Object Windows.Controls.ListBoxItem
                        $lbi.Content = $dText
                        $lbi.Tag = $iAttr.name
                        $identityAttrListBox.Items.Add($lbi) | Out-Null
                    }
                }
                if ($identityAttrListBox.Items.Count -eq 0) {
                    $lbi = New-Object Windows.Controls.ListBoxItem
                    $lbi.Content = "(No matches)"
                    $lbi.IsEnabled = $false
                    $identityAttrListBox.Items.Add($lbi) | Out-Null
                } elseif ($identityAttrListBox.Items.Count -gt 0 -and $identityAttrListBox.Items[0].IsEnabled) {
                    $identityAttrListBox.SelectedIndex = 0
                }
            }.GetNewClosure())

            # ---- Wizard State and Events ----
            $schemaWizState = @{ Confirmed = $false; Result = $null }

            $wzCancelBtn.Add_Click({
                $schemaWizard.Close()
            }.GetNewClosure())

            $wzBackBtn.Add_Click({
                $step1Panel.Visibility  = 'Visible'
                $step2Panel.Visibility  = 'Collapsed'
                $wzStepLabel.Text       = "Step 1 of 2"
                $wzBackBtn.IsEnabled    = $false
                $wzNextBtn.Content      = "Next >"
                $s2ValidationLbl.Text   = ""
            }.GetNewClosure())

            $wzNextBtn.Add_Click({
                if ($step2Panel.Visibility -eq 'Collapsed') {
                    # Advance to step 2
                    $s1ValidationLbl.Text  = ""
                    $step1Panel.Visibility = 'Collapsed'
                    $step2Panel.Visibility = 'Visible'
                    $wzStepLabel.Text      = "Step 2 of 2"
                    $wzBackBtn.IsEnabled   = $true
                    $wzNextBtn.Content     = "Finish"
                } else {
                    # Finish: validate and collect results
                    $selectedSrc       = $corrSrcDropdown.SelectedItem
                    $selectedIdentItem = $identityAttrListBox.SelectedItem
                    if ($selectedSrc -ne "(skip)") {
                        if ($null -eq $selectedIdentItem -or
                            -not ($selectedIdentItem -is [Windows.Controls.ListBoxItem]) -or
                            -not $selectedIdentItem.IsEnabled) {
                            $s2ValidationLbl.Text = "Please select an identity attribute, or set Source Attribute to '(skip)'."
                            return
                        }
                    }
                    $s2ValidationLbl.Text = ""
                    $schemaWizState.Result = @{
                        IdentityAttribute = $idAttrDropdown.SelectedItem
                        DisplayAttribute  = $dispAttrDropdown.SelectedItem
                        CorrSrcAttr       = $selectedSrc
                        CorrIdentAttr     = if ($selectedSrc -eq "(skip)") { "" } else { $selectedIdentItem.Tag }
                    }
                    $schemaWizState.Confirmed = $true
                    $schemaWizard.Close()
                }
            }.GetNewClosure())

            $schemaWizard.ShowDialog() | Out-Null

            if (-not $schemaWizState.Confirmed) {
                Write-Log "Schema upload cancelled by user." "INFO"
                return
            }

            $identityAttribute  = $schemaWizState.Result.IdentityAttribute
            $displayAttribute   = $schemaWizState.Result.DisplayAttribute
            $corrSrcAttr        = $schemaWizState.Result.CorrSrcAttr
            $corrIdentAttr      = $schemaWizState.Result.CorrIdentAttr

            # Analyse all rows to detect multi-valued columns.
            # A column is multi-valued if any single identity value appears with more than one
            # distinct non-empty value in that column - meaning one account has multiple values.
            Write-Log "Analyzing $($allRows.Count) rows for multi-valued columns (identity attribute: '$identityAttribute')..." "INFO"
            $multiValuedColumns = @{}
            foreach ($col in $headers) { $multiValuedColumns[$col] = $false }
            foreach ($col in $headers) {
                if ($col -eq $identityAttribute) { continue }
                $valueMap = @{}
                foreach ($row in $allRows) {
                    $idVal = $row.$identityAttribute
                    if ([string]::IsNullOrEmpty($idVal)) { continue }
                    $colVal = $row.$col
                    if (-not $valueMap.ContainsKey($idVal)) {
                        $valueMap[$idVal] = [System.Collections.Generic.HashSet[string]]::new()
                    }
                    if (-not [string]::IsNullOrEmpty($colVal)) {
                        $valueMap[$idVal].Add($colVal) | Out-Null
                    }
                    if ($valueMap[$idVal].Count -gt 1) {
                        $multiValuedColumns[$col] = $true
                        break
                    }
                }
            }
            $mvNames = @($multiValuedColumns.Keys | Where-Object { $multiValuedColumns[$_] })
            if ($mvNames.Count -gt 0) {
                Write-Log "Multi-valued columns detected: $($mvNames -join ', ')" "INFO"
            } else {
                Write-Log "No multi-valued columns detected (each identity value has at most one value per column)." "INFO"
            }
            # Per-entitlement-column detection diagnostic (logged after entitlement detection below)

            # List schemas for the source and find the "account" schema
            Write-Log "Fetching existing schemas for source $sourceID..." "INFO"
            $schemas = Invoke-RestMethod -Uri "$tenantUrl/v2025/sources/$sourceID/schemas" -Method Get -Headers $apiHeaders
            $accountSchema = $schemas | Where-Object { $_.name -eq 'account' } | Select-Object -First 1

            if ($null -eq $accountSchema) {
                [System.Windows.MessageBox]::Show("No 'account' schema found on source $sourceID.`n`nA default schema must exist before it can be updated. Please verify the source is configured correctly in SailPoint.", "Upload Schema", "OK", "Warning")
                return
            }

            $schemaId = $accountSchema.id
            Write-Log "Found account schema: $schemaId" "INFO"

            # Look for a group/entitlement schema on this source (required for isGroup: true per SailPoint API)
            $groupSchemaRef = $schemas | Where-Object { $_.name -ne 'account' } | Select-Object -First 1
            if ($null -ne $groupSchemaRef) {
                Write-Log "Found group schema for entitlement reference: '$($groupSchemaRef.name)' ($($groupSchemaRef.id))" "INFO"
            } else {
                Write-Log "No group schema found on source - entitlement columns will have isGroup: false (isEntitlement and isMultiValued still applied)" "INFO"
            }

            # Determine entitlement columns from the app's config.json
            $entitlementColumns = @()
            try {
                $appConfigJson = Get-Content $script:currentConfigPath -Raw | ConvertFrom-Json

                # Parse groupTypes (comma-string or array)
                $cfgGroupTypes = @()
                if ($appConfigJson.groupTypes -is [System.Array]) {
                    $cfgGroupTypes = @($appConfigJson.groupTypes | Where-Object { $_ })
                } elseif (-not [string]::IsNullOrWhiteSpace($appConfigJson.groupTypes)) {
                    $cfgGroupTypes = $appConfigJson.groupTypes -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                }

                # Parse booleanColumnList
                $cfgBooleanColumns = @()
                if ($appConfigJson.booleanColumnList -is [System.Array]) {
                    $cfgBooleanColumns = @($appConfigJson.booleanColumnList | Where-Object { $_ })
                } elseif (-not [string]::IsNullOrWhiteSpace($appConfigJson.booleanColumnList)) {
                    $cfgBooleanColumns = $appConfigJson.booleanColumnList -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                }

                $cfgAdminColumnName  = if ($appConfigJson.adminColumnName  -is [System.Array]) { $appConfigJson.adminColumnName[0]  } else { $appConfigJson.adminColumnName }
                $cfgAdminColumnValue = if ($appConfigJson.adminColumnValue -is [System.Array]) { $appConfigJson.adminColumnValue[0] } else { $appConfigJson.adminColumnValue }

                if ($cfgBooleanColumns.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($appConfigJson.booleanColumnValue)) {
                    # booleanColumn processing removes boolean columns and creates a single "Role" column
                    $entitlementColumns = @("Role")
                    Write-Log "Detected boolean entitlement processing - 'Role' marked as entitlement." "INFO"
                } elseif ($cfgGroupTypes.Count -gt 0) {
                    # groupTypes columns are the entitlement columns in the output file
                    $entitlementColumns = $cfgGroupTypes
                    Write-Log "Detected groupTypes entitlements: $($cfgGroupTypes -join ', ')" "INFO"
                } elseif (-not [string]::IsNullOrWhiteSpace($cfgAdminColumnName)) {
                    # adminColumn processing adds a "Role" column (Admin/User)
                    $entitlementColumns = @("Role")
                    Write-Log "Detected admin-column processing - 'Role' marked as entitlement." "INFO"
                } else {
                    # Default: no groupTypes or boolean columns configured.
                    # Process-ImportedData always inserts a 'Role' column in this default path.
                    # Mark it as an entitlement if it is present in the upload file.
                    if ($headers -contains "Role") {
                        $entitlementColumns = @("Role")
                        Write-Log "Detected default 'Role' column in processed file - marked as entitlement." "INFO"
                    }
                }
            }
            catch {
                Write-Log "WARNING: Could not read app config for entitlement detection - all attributes will be plain STRING: $_" "WARNING"
            }

            # Diagnostic: report data-driven multi-valued result for each entitlement column
            foreach ($col in $entitlementColumns) {
                $colInFile = $headers -contains $col
                if ($colInFile) {
                    $detectedMv = $multiValuedColumns[$col]
                    Write-Log "Entitlement column '$col': isMultiValued=$detectedMv (data-driven from $($allRows.Count) rows)" "INFO"
                } else {
                    Write-Log "Entitlement column '$col' from config was NOT found in upload file headers - will be absent from schema." "WARNING"
                }
            }

            # Build attributes array from file headers, flagging entitlement columns
            # Per SailPoint API docs: isGroup: true requires a 'schema' CONNECTOR_SCHEMA reference.
            # Without it the API silently ignores isGroup/isMultiValued. Use the group schema ref if
            # one exists on the source; otherwise set isGroup: false (isEntitlement/isMultiValued still apply).
            $attributes = @()
            foreach ($col in $headers) {
                $isEnt = $entitlementColumns -contains $col
                # Purely data-driven: true only when the file shows >1 distinct value per identity.
                # Entitlement columns detected via adminColumnName produce single-valued roles (e.g., Admin/User)
                # and must NOT be forced to multi-valued.
                $isMv  = $multiValuedColumns[$col]
                $hasGroupRef = $isEnt -and ($null -ne $groupSchemaRef)
                $attr = [ordered]@{
                    name          = $col
                    type          = "STRING"
                    isMulti       = $isMv
                    isEntitlement = $isEnt
                    isGroup       = $hasGroupRef
                }
                if ($hasGroupRef) {
                    $attr['schema'] = [ordered]@{
                        type = "CONNECTOR_SCHEMA"
                        id   = $groupSchemaRef.id
                        name = $groupSchemaRef.name
                    }
                }
                $attributes += $attr
            }

            $entCount = ($attributes | Where-Object { $_.isEntitlement }).Count
            Write-Log "Schema attributes: $($attributes.Count) total, $entCount entitlement column(s)." "INFO"

            # Build full schema replacement body (PUT requires the schema id)
            $schemaBody = [ordered]@{
                id                 = $schemaId
                name               = "account"
                nativeObjectType   = "User"
                identityAttribute  = $identityAttribute
                displayAttribute   = $displayAttribute
                includePermissions = $false
                attributes         = $attributes
            }

            $schemaJson = $schemaBody | ConvertTo-Json -Depth 5
            Write-Log "Uploading schema with $($attributes.Count) attributes to source $sourceID..." "INFO"
            # Log every entitlement attribute so the full payload can be verified
            $entAttrs = @($attributes | Where-Object { $_.isEntitlement })
            if ($entAttrs.Count -gt 0) {
                foreach ($ea in $entAttrs) {
                    Write-Log "Entitlement attribute payload: $($ea | ConvertTo-Json -Depth 3 -Compress)" "INFO"
                }
            } else {
                Write-Log "No entitlement attributes in payload (check config entitlement detection)." "WARNING"
            }
            $result = Invoke-RestMethod -Uri "$tenantUrl/v2025/sources/$sourceID/schemas/$schemaId" `
                -Method Put -Headers $apiHeaders -Body $schemaJson -ContentType "application/json"

            Write-Log "Schema uploaded successfully. Attributes: $($result.attributes.Count), Identity: $identityAttribute, Display: $displayAttribute" "INFO"

            # Log what ISC actually stored for entitlement columns (verify the PUT response)
            $returnedEntAttrs = @($result.attributes | Where-Object { $_.isEntitlement })
            if ($returnedEntAttrs.Count -gt 0) {
                foreach ($rea in $returnedEntAttrs) {
                    Write-Log "ISC returned entitlement attribute: $($rea | ConvertTo-Json -Depth 3 -Compress)" "INFO"
                }
            } else {
                Write-Log "ISC returned 0 entitlement attributes - the API may have stripped them." "WARNING"
            }

            $entNames = ($attributes | Where-Object { $_.isEntitlement } | ForEach-Object { $_.name }) -join ", "
            $entSummary = if ($entCount -gt 0) { "`nEntitlement Columns ($entCount): $entNames" } else { "`nEntitlement Columns: none detected" }

            # --- Correlation Config ---
            $corrConfigured = $false
            $corrSummary = ""
            if ($corrSrcAttr -ne "(skip)" -and -not [string]::IsNullOrWhiteSpace($corrSrcAttr) -and -not [string]::IsNullOrWhiteSpace($corrIdentAttr)) {
                Write-Log "Configuring correlation: identity attr '$corrIdentAttr' (property) = source attr '$corrSrcAttr' (value)..." "INFO"
                try {
                    $corrConfigGet = Invoke-RestMethod -Uri "$tenantUrl/beta/sources/$sourceID/correlation-config" -Method Get -Headers $apiHeaders
                    $corrConfigId   = $corrConfigGet.id
                    $corrConfigName = if ($corrConfigGet.name) { $corrConfigGet.name } else { "correlationConfig" }
                    $corrBody = [ordered]@{
                        id   = $corrConfigId
                        name = $corrConfigName
                        attributeAssignments = @(
                            [ordered]@{
                                property     = $corrIdentAttr
                                value        = $corrSrcAttr
                                operation    = "EQ"
                                complex      = $false
                                ignoreCase   = $false
                                matchMode    = "ANYWHERE"
                                filterString = ""
                            }
                        )
                    }
                    $corrJson = $corrBody | ConvertTo-Json -Depth 5
                    $null = Invoke-RestMethod -Uri "$tenantUrl/beta/sources/$sourceID/correlation-config" `
                        -Method Put -Headers $apiHeaders -Body $corrJson -ContentType "application/json"
                    Write-Log "Correlation config updated successfully: identity='$corrIdentAttr' (property), source='$corrSrcAttr' (value)" "INFO"
                    $corrConfigured = $true
                    $corrSummary = "`nCorrelation: identity=$corrIdentAttr / source=$corrSrcAttr"
                }
                catch {
                    Write-Log "WARNING: Correlation config update failed: $_" "WARNING"
                    $corrSummary = "`nCorrelation: failed - check log"
                }
            }

            [System.Windows.MessageBox]::Show(
                "Schema uploaded successfully!`n`nSource: $script:selectedAppName`nAttributes: $($result.attributes.Count)`nIdentity Attribute: $identityAttribute`nDisplay Attribute: $displayAttribute$entSummary$corrSummary",
                "Upload Schema", "OK", "Information"
            )
        }
        catch {
            Write-Log "ERROR uploading schema: $_" "ERROR"
            [System.Windows.MessageBox]::Show("Schema upload failed!`n`nError: $_", "Upload Schema", "OK", "Error")
        }
    })

    # Reset Source Button Handler
    $resetSourceButton.Add_Click({
        if (-not $script:selectedAppName) { return }

        try {
            # Get sourceID from current config fields
            $sourceID = $null
            if ($script:configFields.ContainsKey('sourceID')) {
                $sourceID = $script:configFields['sourceID'].Text.Trim()
            }
            if ([string]::IsNullOrWhiteSpace($sourceID)) {
                [System.Windows.MessageBox]::Show("No Source ID found for this app. Please ensure a config is loaded.", "Reset Source", "OK", "Warning")
                return
            }

            # Strong warning confirmation dialog
            $warnResult = [System.Windows.MessageBox]::Show(
                "WARNING: This operation is DESTRUCTIVE and IRREVERSIBLE!`n`nThe following will be permanently removed from SailPoint ISC for source '$script:selectedAppName' (ID: $sourceID):`n`n  - All Accounts`n  - All Entitlements`n  - Account Schema attributes`n  - Correlation configuration`n`nThis cannot be undone. Are you absolutely sure you want to proceed?",
                "DESTRUCTIVE: Reset Source",
                "YesNo",
                "Warning"
            )
            if ($warnResult -ne [System.Windows.MessageBoxResult]::Yes) {
                Write-Log "Reset Source cancelled by user." "INFO"
                return
            }

            # Second confirmation
            $confirm2 = [System.Windows.MessageBox]::Show(
                "Final confirmation: Reset ALL data for '$script:selectedAppName'?`n`nClick Yes to proceed with the reset.",
                "Confirm Reset",
                "YesNo",
                "Warning"
            )
            if ($confirm2 -ne [System.Windows.MessageBoxResult]::Yes) {
                Write-Log "Reset Source cancelled by user at second confirmation." "INFO"
                return
            }

            Write-Log "Starting source reset for '$script:selectedAppName' (sourceID: $sourceID)..." "WARNING"

            # Get auth token
            $currentSettings = Load-Settings
            $tenantUrl = Get-BaseApiUrl -Settings $currentSettings
            if ($null -eq $tenantUrl) {
                [System.Windows.MessageBox]::Show("Cannot determine API URL. Please check Settings.", "Reset Source", "OK", "Error")
                return
            }

            $authHeader = [System.Convert]::ToBase64String(
                [System.Text.Encoding]::ASCII.GetBytes("$($currentSettings.ClientID):$($currentSettings.ClientSecret)")
            )

            Write-Log "Retrieving OAuth token for source reset..." "INFO"
            $tokenResponse = Invoke-RestMethod -Uri "$tenantUrl/oauth/token" -Method Post `
                -Headers @{ Authorization = "Basic $authHeader" } `
                -Body @{ grant_type = "client_credentials" }
            $accessToken = $tokenResponse.access_token

            $apiHeaders = @{
                Authorization  = "Bearer $accessToken"
                Accept         = "application/json"
                "Content-Type" = "application/json"
            }

            $resetErrors = @()

            # Step 1: Reset Entitlements
            # Must run first — entitlement objects reference schema attributes; clearing them
            # before the schema allows the schema PUT to succeed without "referenced by" errors.
            Write-Log "Step 1/4: Resetting entitlements..." "INFO"
            try {
                $entResult = Invoke-RestMethod -Uri "$tenantUrl/beta/entitlements/reset/sources/$sourceID" `
                    -Method Post -Headers $apiHeaders
                $entTaskId = if ($entResult.id) { $entResult.id } else { "(no task id)" }
                Write-Log "Entitlements reset initiated. Task ID: $entTaskId" "INFO"
            }
            catch {
                $msg = "Entitlements reset failed: $_"
                Write-Log $msg "ERROR"
                $resetErrors += $msg
            }

            # Step 2: Reset Accounts
            # Run after entitlements; account objects must be cleared before the schema can
            # be safely wiped (accounts hold references to schema attributes).
            Write-Log "Step 2/4: Resetting accounts..." "INFO"
            try {
                $acctResult = Invoke-RestMethod -Uri "$tenantUrl/v2025/sources/$sourceID/remove-accounts" `
                    -Method Post -Headers $apiHeaders
                $taskId = if ($acctResult.id) { $acctResult.id } else { "(no task id)" }
                Write-Log "Accounts reset initiated. Task ID: $taskId" "INFO"
            }
            catch {
                $msg = "Accounts reset failed: $_"
                Write-Log $msg "ERROR"
                $resetErrors += $msg
            }

            # Step 3: Reset Correlation Config (clear attribute assignments)
            # Clear correlation before schema so the schema attributes are no longer
            # referenced by any correlation rule when we attempt to wipe them.
            Write-Log "Step 3/4: Resetting correlation config..." "INFO"
            try {
                $corrConfig = Invoke-RestMethod -Uri "$tenantUrl/beta/sources/$sourceID/correlation-config" `
                    -Method Get -Headers $apiHeaders
                $corrId   = $corrConfig.id
                $corrName = if ($corrConfig.name) { $corrConfig.name } else { "correlationConfig" }
                $corrResetBody = [ordered]@{
                    id                   = $corrId
                    name                 = $corrName
                    attributeAssignments = @()
                }
                $corrJson = $corrResetBody | ConvertTo-Json -Depth 5
                $null = Invoke-RestMethod -Uri "$tenantUrl/beta/sources/$sourceID/correlation-config" `
                    -Method Put -Headers $apiHeaders -Body $corrJson -ContentType "application/json"
                Write-Log "Correlation config cleared. Config ID: $corrId" "INFO"
            }
            catch {
                $msg = "Correlation config reset failed: $_"
                Write-Log $msg "ERROR"
                $resetErrors += $msg
            }

            # Step 4: Reset Account Schema (clear all attributes)
            # Runs last — all entitlements, accounts, and correlation references have been
            # removed, so the schema attributes can now be deleted without dependency errors.
            Write-Log "Step 4/4: Resetting account schema..." "INFO"
            try {
                $schemas = Invoke-RestMethod -Uri "$tenantUrl/v2025/sources/$sourceID/schemas" -Method Get -Headers $apiHeaders
                $accountSchema = $schemas | Where-Object { $_.name -eq 'account' } | Select-Object -First 1
                if ($null -eq $accountSchema) {
                    Write-Log "No 'account' schema found on source $sourceID - skipping schema reset." "WARNING"
                    $resetErrors += "Account schema not found on source - schema reset skipped."
                }
                else {
                    $schemaId = $accountSchema.id
                    $schemaResetBody = [ordered]@{
                        id                 = $schemaId
                        name               = "account"
                        nativeObjectType   = "User"
                        identityAttribute  = if ($accountSchema.identityAttribute) { $accountSchema.identityAttribute } else { "id" }
                        displayAttribute   = if ($accountSchema.displayAttribute) { $accountSchema.displayAttribute } else { "id" }
                        includePermissions = $false
                        attributes         = @()
                    }
                    $schemaJson = $schemaResetBody | ConvertTo-Json -Depth 5
                    $schemaResult = Invoke-RestMethod -Uri "$tenantUrl/v2025/sources/$sourceID/schemas/$schemaId" `
                        -Method Put -Headers $apiHeaders -Body $schemaJson -ContentType "application/json"
                    Write-Log "Account schema cleared. Schema ID: $schemaId, attributes remaining: $($schemaResult.attributes.Count)" "INFO"
                }
            }
            catch {
                $msg = "Account schema reset failed: $_"
                Write-Log $msg "ERROR"
                $resetErrors += $msg
            }

            if ($resetErrors.Count -gt 0) {
                $errorList = $resetErrors -join "`n"
                [System.Windows.MessageBox]::Show(
                    "Source reset completed with errors for '$script:selectedAppName'.`n`nThe following steps failed:`n$errorList`n`nCheck the log for details.",
                    "Reset Source - Partial Failure", "OK", "Warning"
                )
            }
            else {
                Write-Log "Source reset completed successfully for '$script:selectedAppName'." "INFO"
                [System.Windows.MessageBox]::Show(
                    "Source reset completed successfully for '$script:selectedAppName'.`n`nAll 4 reset steps were initiated:`n  - Accounts removed`n  - Entitlements removed`n  - Account schema cleared`n  - Correlation config cleared`n`nNote: Account and entitlement removal tasks run asynchronously in SailPoint ISC.",
                    "Reset Source", "OK", "Information"
                )
            }
        }
        catch {
            Write-Log "ERROR during source reset: $_" "ERROR"
            [System.Windows.MessageBox]::Show("Source reset failed!`n`nError: $_", "Reset Source", "OK", "Error")
        }
    })

    # Load app folders on startup
    & $loadAppFolders

    # View Operation Log button handler
    $viewOperationLogButton.Add_Click({
        # Create popup window
        $logWindow = New-Object Windows.Window
        $logWindow.Title = "Operation Log"
        $logWindow.Width = 900
        $logWindow.Height = 600
        $logWindow.WindowStartupLocation = 'CenterScreen'
        $logWindow.Background = [System.Windows.Media.Brushes]::White
        
        $logWindowPanel = New-Object Windows.Controls.DockPanel
        $logWindowPanel.Margin = '15'
        $logWindow.Content = $logWindowPanel
        
        # Header with title and close button
        $logHeaderPanel = New-Object Windows.Controls.DockPanel
        $logHeaderPanel.Margin = '0,0,0,10'
        $logHeaderPanel.LastChildFill = $false
        [Windows.Controls.DockPanel]::SetDock($logHeaderPanel, 'Top')
        $logWindowPanel.Children.Add($logHeaderPanel)
        
        # Title
        $logTitleLabel = New-Object Windows.Controls.Label
        $logTitleLabel.Content = "Operation Log"
        $logTitleLabel.FontSize = 16
        $logTitleLabel.FontWeight = 'Bold'
        [Windows.Controls.DockPanel]::SetDock($logTitleLabel, 'Left')
        $logHeaderPanel.Children.Add($logTitleLabel)
        
        # Close button
        $closeBtn = New-Object Windows.Controls.Button
        $closeBtn.Content = "Close"
        $closeBtn.Padding = '10,6'
        $closeBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DC3545")
        $closeBtn.Foreground = [System.Windows.Media.Brushes]::White
        $closeBtn.FontWeight = 'SemiBold'
        $closeBtn.FontSize = 12
        $closeBtn.BorderThickness = '0'
        $closeBtn.Cursor = 'Hand'
        [Windows.Controls.DockPanel]::SetDock($closeBtn, 'Right')
        $logHeaderPanel.Children.Add($closeBtn)
        
        $closeBtn.Add_Click({
            $logWindow.Close()
        })
        
        # Log text box
        $logTextBox = New-Object Windows.Controls.TextBox
        $logTextBox.AcceptsReturn = $true
        $logTextBox.IsReadOnly = $true
        $logTextBox.TextWrapping = 'Wrap'
        $logTextBox.VerticalScrollBarVisibility = 'Auto'
        $logTextBox.HorizontalScrollBarVisibility = 'Auto'
        $logTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1E1E1E")
        $logTextBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#D4D4D4")
        $logTextBox.Padding = '15'
        $logTextBox.FontSize = 12
        $logTextBox.FontFamily = 'Consolas'
        $logTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
        $logTextBox.BorderThickness = '1'
        $logTextBox.Text = $script:operationLogText
        $logWindowPanel.Children.Add($logTextBox)
        
        # Show window
        $logWindow.ShowDialog() | Out-Null
    })

    # Show initial message in operation log
    Write-Log "***************************************************************" "INFO"
    Write-Log "  SailPoint File Upload Utility - Management Console" "INFO"
    Write-Log "***************************************************************" "INFO"
    Write-Log "" "INFO"
    Write-Log "Ready. Use the tabs above to navigate:" "INFO"
    Write-Log "  Settings - Configure credentials and paths" "INFO"
    Write-Log "  App Management - Manage directories, edit configs, and upload files" "INFO"
    Write-Log "" "INFO"
    Write-Log "Click an app name to view/edit its config. Click the upload button to upload." "INFO"
    Write-Log "" "INFO"

    # Show the window
    $window.ShowDialog() | Out-Null
}

# ------------------------
# Main Entry Point
# ------------------------

Show-MainWindow
