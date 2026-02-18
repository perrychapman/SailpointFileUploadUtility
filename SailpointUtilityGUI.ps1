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
# File Upload Logic (Single App)
# ------------------------

function Run-SingleAppUpload {
    param (
        [PSObject]$settings,
        [string]$appName
    )

    Write-Log "=== Starting File Upload for App: $appName ===" "INFO"

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
        $output = pwsh.exe -ExecutionPolicy Bypass -File $scriptPath 2>&1
        
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

        Write-Log "=== File Upload Process Complete for $appName ===" "INFO"
        
        # Reset AppFilter and save
        $settings.AppFilter = ""
        Save-Settings -settings $settings
        Write-Log "AppFilter reset to empty" "INFO"
        
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
    $window.Title = "SailPoint File Upload Utility - Management Console"
    $window.Width = 1200
    $window.Height = 800
    $window.MinWidth = 900
    $window.MinHeight = 600
    $window.WindowStartupLocation = 'CenterScreen'
    $window.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F5F5F5")

    # Create main grid
    $mainGrid = New-Object Windows.Controls.Grid
    $mainGrid.Margin = '0'
    $window.Content = $mainGrid

    # Define rows - header and content (tabs)
    $headerRow = New-Object Windows.Controls.RowDefinition
    $headerRow.Height = "Auto"
    $contentRow = New-Object Windows.Controls.RowDefinition
    $contentRow.Height = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
    $mainGrid.RowDefinitions.Add($headerRow)
    $mainGrid.RowDefinitions.Add($contentRow)

    # ===== HEADER SECTION =====
    $headerPanel = New-Object Windows.Controls.Border
    $headerPanel.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $headerPanel.Padding = '20,15'
    [Windows.Controls.Grid]::SetRow($headerPanel, 0)
    $mainGrid.Children.Add($headerPanel)

    $headerStack = New-Object Windows.Controls.StackPanel
    $headerPanel.Child = $headerStack

    $titleLabel = New-Object Windows.Controls.Label
    $titleLabel.Content = "SailPoint File Upload Utility"
    $titleLabel.FontSize = 24
    $titleLabel.FontWeight = 'Bold'
    $titleLabel.Foreground = [System.Windows.Media.Brushes]::White
    $headerStack.Children.Add($titleLabel)

    $subtitleLabel = New-Object Windows.Controls.Label
    $subtitleLabel.Content = "Management Console - Configure, Manage Directories, and Upload Files"
    $subtitleLabel.FontSize = 12
    $subtitleLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#AAAAAA")
    $subtitleLabel.Margin = '-5,0,0,0'
    $headerStack.Children.Add($subtitleLabel)

    # View Operation Log button in header
    $viewOperationLogButton = New-Object Windows.Controls.Button
    $viewOperationLogButton.Content = "📋 View Operation Log"
    $viewOperationLogButton.Padding = '12,6'
    $viewOperationLogButton.Margin = '0,8,0,0'
    $viewOperationLogButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#6C757D")
    $viewOperationLogButton.Foreground = [System.Windows.Media.Brushes]::White
    $viewOperationLogButton.FontWeight = 'SemiBold'
    $viewOperationLogButton.FontSize = 11
    $viewOperationLogButton.BorderThickness = '0'
    $viewOperationLogButton.Cursor = 'Hand'
    $headerStack.Children.Add($viewOperationLogButton)

    # Initialize log storage (but no visible textbox)
    $script:operationLogText = ""
    $script:logTextBoxInitialized = $true

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
    $tab1.Header = "App Management"
    $tab1.FontSize = 13
    $tab1.FontWeight = 'SemiBold'
    $tabControl.Items.Add($tab1)

    $tab1MainPanel = New-Object Windows.Controls.DockPanel
    $tab1MainPanel.Background = [System.Windows.Media.Brushes]::White
    $tab1.Content = $tab1MainPanel

    # HEADER ROW - App selection dropdown and action buttons
    $headerPanel = New-Object Windows.Controls.DockPanel
    $headerPanel.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F8F9FA")
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
    $refreshAppsButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#28A745")
    $refreshAppsButton.Foreground = [System.Windows.Media.Brushes]::White
    $refreshAppsButton.FontWeight = 'SemiBold'
    $refreshAppsButton.FontSize = 10
    $refreshAppsButton.BorderThickness = '0'
    $refreshAppsButton.Cursor = 'Hand'
    $actionButtonsPanel.Children.Add($refreshAppsButton)

    # Directory Creation Button
    $createDirsButton = New-Object Windows.Controls.Button
    $createDirsButton.Content = "App Management"
    $createDirsButton.Padding = '12,6'
    $createDirsButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $createDirsButton.Foreground = [System.Windows.Media.Brushes]::White
    $createDirsButton.FontWeight = 'SemiBold'
    $createDirsButton.FontSize = 10
    $createDirsButton.BorderThickness = '0'
    $createDirsButton.Cursor = 'Hand'
    $actionButtonsPanel.Children.Add($createDirsButton)

    # App selection dropdown on the left
    $appSelectorPanel = New-Object Windows.Controls.StackPanel
    $appSelectorPanel.Orientation = 'Horizontal'
    $appSelectorPanel.Margin = '0,5,0,5'
    [Windows.Controls.DockPanel]::SetDock($appSelectorPanel, 'Left')
    $headerPanel.Children.Add($appSelectorPanel)

    $appLabel = New-Object Windows.Controls.Label
    $appLabel.Content = "Select App:"
    $appLabel.FontWeight = 'SemiBold'
    $appLabel.FontSize = 11
    $appLabel.VerticalAlignment = 'Center'
    $appLabel.Margin = '0,0,8,0'
    $appSelectorPanel.Children.Add($appLabel)

    $appDropdown = New-Object Windows.Controls.ComboBox
    $appDropdown.MinWidth = 250
    $appDropdown.Padding = '8,4'
    $appDropdown.FontSize = 11
    $appDropdown.Background = [System.Windows.Media.Brushes]::White
    $appDropdown.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $appDropdown.BorderThickness = '1'
    $appSelectorPanel.Children.Add($appDropdown)
    $script:appDropdown = $appDropdown

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

    # LEFT PANEL: Config Editor
    $configEditorBox = New-Object Windows.Controls.GroupBox
    $configEditorBox.Header = "App Configuration"
    $configEditorBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $configEditorBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#28A745")
    $configEditorBox.BorderThickness = '2'
    $configEditorBox.Margin = '0,0,8,0'
    $configEditorBox.Padding = '12'
    $configEditorBox.FontWeight = 'Bold'
    [Windows.Controls.Grid]::SetColumn($configEditorBox, 0)
    $contentGrid.Children.Add($configEditorBox)

    $configEditorPanel = New-Object Windows.Controls.DockPanel
    $configEditorBox.Content = $configEditorPanel

    # Config action buttons at top
    $configButtonsPanel = New-Object Windows.Controls.StackPanel
    $configButtonsPanel.Orientation = 'Horizontal'
    $configButtonsPanel.Margin = '0,0,0,10'
    [Windows.Controls.DockPanel]::SetDock($configButtonsPanel, 'Top')
    $configEditorPanel.Children.Add($configButtonsPanel)

    $saveConfigButton = New-Object Windows.Controls.Button
    $saveConfigButton.Content = "Save Config"
    $saveConfigButton.Padding = '12,6'
    $saveConfigButton.Margin = '0,0,8,0'
    $saveConfigButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#28A745")
    $saveConfigButton.Foreground = [System.Windows.Media.Brushes]::White
    $saveConfigButton.FontWeight = 'SemiBold'
    $saveConfigButton.FontSize = 10
    $saveConfigButton.BorderThickness = '0'
    $saveConfigButton.Cursor = 'Hand'
    $saveConfigButton.IsEnabled = $false
    $configButtonsPanel.Children.Add($saveConfigButton)

    $reloadConfigButton = New-Object Windows.Controls.Button
    $reloadConfigButton.Content = "Reload Config"
    $reloadConfigButton.Padding = '12,6'
    $reloadConfigButton.Margin = '0,0,8,0'
    $reloadConfigButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#6C757D")
    $reloadConfigButton.Foreground = [System.Windows.Media.Brushes]::White
    $reloadConfigButton.FontWeight = 'SemiBold'
    $reloadConfigButton.FontSize = 10
    $reloadConfigButton.BorderThickness = '0'
    $reloadConfigButton.Cursor = 'Hand'
    $reloadConfigButton.IsEnabled = $false
    $configButtonsPanel.Children.Add($reloadConfigButton)

    $uploadAppButton = New-Object Windows.Controls.Button
    $uploadAppButton.Content = "Upload Files"
    $uploadAppButton.Padding = '14,7'
    $uploadAppButton.Margin = '0,0,8,0'
    $uploadAppButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $uploadAppButton.Foreground = [System.Windows.Media.Brushes]::White
    $uploadAppButton.FontWeight = 'SemiBold'
    $uploadAppButton.FontSize = 11
    $uploadAppButton.BorderThickness = '0'
    $uploadAppButton.Cursor = 'Hand'
    $uploadAppButton.IsEnabled = $false
    $configButtonsPanel.Children.Add($uploadAppButton)

    $uploadUserListButton = New-Object Windows.Controls.Button
    $uploadUserListButton.Content = "Upload User List"
    $uploadUserListButton.Padding = '14,7'
    $uploadUserListButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#17A2B8")
    $uploadUserListButton.Foreground = [System.Windows.Media.Brushes]::White
    $uploadUserListButton.FontWeight = 'SemiBold'
    $uploadUserListButton.FontSize = 11
    $uploadUserListButton.BorderThickness = '0'
    $uploadUserListButton.Cursor = 'Hand'
    $uploadUserListButton.IsEnabled = $false
    $configButtonsPanel.Children.Add($uploadUserListButton)

    # Config editor - ScrollViewer with form fields
    $configScrollViewer = New-Object Windows.Controls.ScrollViewer
    $configScrollViewer.VerticalScrollBarVisibility = 'Auto'
    $configScrollViewer.HorizontalScrollBarVisibility = 'Disabled'
    $configScrollViewer.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $configScrollViewer.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $configScrollViewer.BorderThickness = '1'
    $configScrollViewer.Padding = '10'
    $configEditorPanel.Children.Add($configScrollViewer)
    
    $script:configFieldsPanel = New-Object Windows.Controls.StackPanel
    $script:configFieldsPanel.Margin = '0'
    $configScrollViewer.Content = $script:configFieldsPanel
    
    # Dictionary to store config field controls
    $script:configFields = @{}

    # RIGHT PANEL: Logs
    $logViewerBox = New-Object Windows.Controls.GroupBox
    $logViewerBox.Header = "App Logs"
    $logViewerBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $logViewerBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFC107")
    $logViewerBox.BorderThickness = '2'
    $logViewerBox.Margin = '8,0,0,0'
    $logViewerBox.Padding = '12'
    $logViewerBox.FontWeight = 'Bold'
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
    $openAppLogButton.Content = "Open App Log Folder"
    $openAppLogButton.Padding = '10,6'
    $openAppLogButton.Margin = '0,0,8,0'
    $openAppLogButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#6C757D")
    $openAppLogButton.Foreground = [System.Windows.Media.Brushes]::White
    $openAppLogButton.FontWeight = 'SemiBold'
    $openAppLogButton.FontSize = 10
    $openAppLogButton.BorderThickness = '0'
    $openAppLogButton.Cursor = 'Hand'
    $openAppLogButton.IsEnabled = $false
    $logButtonsPanel.Children.Add($openAppLogButton)

    # Log viewer textbox
    $script:appLogTextBox = New-Object Windows.Controls.TextBox
    $script:appLogTextBox.AcceptsReturn = $true
    $script:appLogTextBox.IsReadOnly = $true
    $script:appLogTextBox.TextWrapping = 'Wrap'
    $script:appLogTextBox.VerticalScrollBarVisibility = 'Auto'
    $script:appLogTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $script:appLogTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $script:appLogTextBox.Padding = '8,6'
    $script:appLogTextBox.FontSize = 9
    $script:appLogTextBox.FontFamily = New-Object Windows.Media.FontFamily("Consolas")
    $script:appLogTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $script:appLogTextBox.BorderThickness = '1'
    $script:appLogTextBox.Text = "Select an app to view log information..."
    $logViewerPanel.Children.Add($script:appLogTextBox)

    # =============================================================================
    # TAB 2: SETTINGS
    # =============================================================================
    $tab2 = New-Object Windows.Controls.TabItem
    $tab2.Header = "Settings"
    $tab2.FontSize = 13
    $tab2.FontWeight = 'SemiBold'
    $tabControl.Items.Add($tab2)

    $tab2MainPanel = New-Object Windows.Controls.DockPanel
    $tab2MainPanel.Background = [System.Windows.Media.Brushes]::White
    $tab2.Content = $tab2MainPanel

    # HEADER ROW - Execution log selector and buttons
    $headerPanel2 = New-Object Windows.Controls.DockPanel
    $headerPanel2.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F8F9FA")
    $headerPanel2.Margin = '15,15,15,10'
    $headerPanel2.LastChildFill = $false
    [Windows.Controls.DockPanel]::SetDock($headerPanel2, 'Top')
    $tab2MainPanel.Children.Add($headerPanel2)

    # Action buttons on the right
    $headerButtonsPanel = New-Object Windows.Controls.StackPanel
    $headerButtonsPanel.Orientation = 'Horizontal'
    $headerButtonsPanel.Margin = '10,5,0,5'
    [Windows.Controls.DockPanel]::SetDock($headerButtonsPanel, 'Right')
    $headerPanel2.Children.Add($headerButtonsPanel)

    # Refresh logs button
    $refreshLogsButton = New-Object Windows.Controls.Button
    $refreshLogsButton.Content = "Refresh Logs"
    $refreshLogsButton.Padding = '12,6'
    $refreshLogsButton.Margin = '0,0,8,0'
    $refreshLogsButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#28A745")
    $refreshLogsButton.Foreground = [System.Windows.Media.Brushes]::White
    $refreshLogsButton.FontWeight = 'SemiBold'
    $refreshLogsButton.FontSize = 10
    $refreshLogsButton.BorderThickness = '0'
    $refreshLogsButton.Cursor = 'Hand'
    $headerButtonsPanel.Children.Add($refreshLogsButton)

    # Open log folder button
    $openLogFolderButton = New-Object Windows.Controls.Button
    $openLogFolderButton.Content = "Open Log Folder"
    $openLogFolderButton.Padding = '12,6'
    $openLogFolderButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#6C757D")
    $openLogFolderButton.Foreground = [System.Windows.Media.Brushes]::White
    $openLogFolderButton.FontWeight = 'SemiBold'
    $openLogFolderButton.FontSize = 10
    $openLogFolderButton.BorderThickness = '0'
    $openLogFolderButton.Cursor = 'Hand'
    $headerButtonsPanel.Children.Add($openLogFolderButton)

    # Log selector on the left
    $logSelectorPanel = New-Object Windows.Controls.StackPanel
    $logSelectorPanel.Orientation = 'Horizontal'
    $logSelectorPanel.Margin = '0,5,0,5'
    [Windows.Controls.DockPanel]::SetDock($logSelectorPanel, 'Left')
    $headerPanel2.Children.Add($logSelectorPanel)

    $logLabel = New-Object Windows.Controls.Label
    $logLabel.Content = "Execution Log:"
    $logLabel.FontWeight = 'SemiBold'
    $logLabel.FontSize = 11
    $logLabel.VerticalAlignment = 'Center'
    $logLabel.Margin = '0,0,8,0'
    $logSelectorPanel.Children.Add($logLabel)

    $execLogDateDropdown = New-Object Windows.Controls.ComboBox
    $execLogDateDropdown.MinWidth = 300
    $execLogDateDropdown.Padding = '8,4'
    $execLogDateDropdown.FontSize = 11
    $execLogDateDropdown.Background = [System.Windows.Media.Brushes]::White
    $execLogDateDropdown.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $execLogDateDropdown.BorderThickness = '1'
    $logSelectorPanel.Children.Add($execLogDateDropdown)
    $script:execLogDateDropdown = $execLogDateDropdown

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
    $directoryBox.Header = "Directory Locations"
    $directoryBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $directoryBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $directoryBox.BorderThickness = '2'
    $directoryBox.Margin = '0,0,0,20'
    $directoryBox.Padding = '15,10,15,15'
    $directoryBox.FontWeight = 'Bold'
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
        $label.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
        $label.FontWeight = 'Normal'
        $label.Margin = '0,8,0,3'
        $label.FontSize = 11
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
        $browseButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
        $browseButton.Foreground = [System.Windows.Media.Brushes]::White
        $browseButton.FontSize = 12
        $browseButton.BorderThickness = '0'
        $browseButton.Cursor = 'Hand'
        $browseButton.ToolTip = "Browse for " + $field.Type
        [Windows.Controls.DockPanel]::SetDock($browseButton, [Windows.Controls.Dock]::Right)
        $pathStack.Children.Add($browseButton)

        $textBox = New-Object Windows.Controls.TextBox
        $textBox.Text = $field.Value
        $textBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
        $textBox.Foreground = [System.Windows.Media.Brushes]::Black
        $textBox.Padding = '8,6'
        $textBox.FontSize = 11
        $textBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
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
    $credentialsBox.Header = "SailPoint Client Credentials"
    $credentialsBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $credentialsBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $credentialsBox.BorderThickness = '2'
    $credentialsBox.Margin = '0,0,0,20'
    $credentialsBox.Padding = '15,10,15,15'
    $credentialsBox.FontWeight = 'Bold'
    $credentialsPanel = New-Object Windows.Controls.StackPanel
    $credentialsPanel.Orientation = 'Vertical'
    $credentialsBox.Content = $credentialsPanel
    $tab2StackPanel.Children.Add($credentialsBox)

    # Tenant
    $tenantLabel = New-Object Windows.Controls.Label
    $tenantLabel.Content = "Tenant"
    $tenantLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
    $tenantLabel.FontWeight = 'Normal'
    $tenantLabel.Margin = '0,8,0,3'
    $tenantLabel.FontSize = 11
    $tenantLabel.ToolTip = "Your SailPoint tenant name (e.g., 'mycompany' from mycompany.identitynow.com). Required if Custom Tenant URL is empty."
    $credentialsPanel.Children.Add($tenantLabel)

    $tenantTextBox = New-Object Windows.Controls.TextBox
    $tenantTextBox.Text = $settings.tenant
    $tenantTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $tenantTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $tenantTextBox.Padding = '8,6'
    $tenantTextBox.FontSize = 11
    $tenantTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $tenantTextBox.BorderThickness = '1'
    $tenantTextBox.ToolTip = "Your SailPoint tenant name (e.g., 'mycompany' from mycompany.identitynow.com). Required if Custom Tenant URL is empty."
    $credentialsPanel.Children.Add($tenantTextBox)
    $script:textBoxes['tenant'] = $tenantTextBox

    # Tenant URL (Vanity URL - Optional)
    $tenantUrlLabel = New-Object Windows.Controls.Label
    $tenantUrlLabel.Content = "Custom Tenant URL (optional)"
    $tenantUrlLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
    $tenantUrlLabel.FontWeight = 'Normal'
    $tenantUrlLabel.Margin = '0,8,0,3'
    $tenantUrlLabel.FontSize = 11
    $tenantUrlLabel.ToolTip = "For vanity URLs (e.g., https://partner7354.identitynow-demo.com). System will construct API URL as https://partner7354.api.identitynow-demo.com"
    $credentialsPanel.Children.Add($tenantUrlLabel)

    $tenantUrlTextBox = New-Object Windows.Controls.TextBox
    $tenantUrlTextBox.Text = $settings.tenantUrl
    $tenantUrlTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $tenantUrlTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $tenantUrlTextBox.Padding = '8,6'
    $tenantUrlTextBox.FontSize = 11
    $tenantUrlTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $tenantUrlTextBox.BorderThickness = '1'
    $tenantUrlTextBox.ToolTip = "For vanity URLs (e.g., https://partner7354.identitynow-demo.com). System will construct API URL as https://partner7354.api.identitynow-demo.com"
    $credentialsPanel.Children.Add($tenantUrlTextBox)
    $script:textBoxes['tenantUrl'] = $tenantUrlTextBox

    # Client ID
    $clientIDLabel = New-Object Windows.Controls.Label
    $clientIDLabel.Content = "Client ID"
    $clientIDLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
    $clientIDLabel.FontWeight = 'Normal'
    $clientIDLabel.Margin = '0,8,0,3'
    $clientIDLabel.FontSize = 11
    $clientIDLabel.ToolTip = "OAuth Client ID from your SailPoint API credentials (Admin > API Management > Create Token)"
    $credentialsPanel.Children.Add($clientIDLabel)

    $clientIDTextBox = New-Object Windows.Controls.TextBox
    $clientIDTextBox.Text = $settings.ClientID
    $clientIDTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $clientIDTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $clientIDTextBox.Padding = '8,6'
    $clientIDTextBox.FontSize = 11
    $clientIDTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $clientIDTextBox.BorderThickness = '1'
    $clientIDTextBox.ToolTip = "OAuth Client ID from your SailPoint API credentials (Admin > API Management > Create Token)"
    $credentialsPanel.Children.Add($clientIDTextBox)
    $script:textBoxes['ClientID'] = $clientIDTextBox

    # Client Secret
    $clientSecretLabel = New-Object Windows.Controls.Label
    $clientSecretLabel.Content = "Client Secret"
    $clientSecretLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
    $clientSecretLabel.FontWeight = 'Normal'
    $clientSecretLabel.Margin = '0,8,0,3'
    $clientSecretLabel.FontSize = 11
    $clientSecretLabel.ToolTip = "OAuth Client Secret from your SailPoint API credentials (keep this secure!)"
    $credentialsPanel.Children.Add($clientSecretLabel)

    $clientSecretPasswordBox = New-Object Windows.Controls.PasswordBox
    $clientSecretPasswordBox.Password = $settings.ClientSecret
    $clientSecretPasswordBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $clientSecretPasswordBox.Foreground = [System.Windows.Media.Brushes]::Black
    $clientSecretPasswordBox.Padding = '8,6'
    $clientSecretPasswordBox.FontSize = 11
    $clientSecretPasswordBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $clientSecretPasswordBox.BorderThickness = '1'
    $clientSecretPasswordBox.ToolTip = "OAuth Client Secret from your SailPoint API credentials (keep this secure!)"
    $credentialsPanel.Children.Add($clientSecretPasswordBox)
    $script:clientSecretBox = $clientSecretPasswordBox

    # Options GroupBox
    $optionsBox = New-Object Windows.Controls.GroupBox
    $optionsBox.Header = "Options"
    $optionsBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $optionsBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $optionsBox.BorderThickness = '2'
    $optionsBox.Margin = '0,0,0,20'
    $optionsBox.Padding = '15,10,15,15'
    $optionsBox.FontWeight = 'Bold'
    $optionsPanel = New-Object Windows.Controls.StackPanel
    $optionsPanel.Orientation = 'Vertical'
    $optionsBox.Content = $optionsPanel
    $tab2StackPanel.Children.Add($optionsBox)

    # Days to Keep Files
    $daysLabel = New-Object Windows.Controls.Label
    $daysLabel.Content = "Days to Keep Files"
    $daysLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
    $daysLabel.FontWeight = 'Normal'
    $daysLabel.Margin = '0,8,0,3'
    $daysLabel.FontSize = 11
    $daysLabel.ToolTip = "Number of days to retain processed files in the Archive folder before automatic deletion. Default: 30 days."
    $optionsPanel.Children.Add($daysLabel)

    $daysTextBox = New-Object Windows.Controls.TextBox
    $daysTextBox.Text = $settings.DaysToKeepFiles.ToString()
    $daysTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $daysTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $daysTextBox.Padding = '8,6'
    $daysTextBox.FontSize = 11
    $daysTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $daysTextBox.BorderThickness = '1'
    $daysTextBox.ToolTip = "Number of days to retain processed files in the Archive folder before automatic deletion. Default: 30 days."
    $optionsPanel.Children.Add($daysTextBox)
    $script:textBoxes['DaysToKeepFiles'] = $daysTextBox

    # Debug Mode Checkbox
    $debugCheckBox = New-Object Windows.Controls.CheckBox
    $debugCheckBox.Content = "Debug Mode"
    $debugCheckBox.IsChecked = $settings.isDebug
    $debugCheckBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
    $debugCheckBox.FontWeight = 'Normal'
    $debugCheckBox.Margin = '0,12,0,8'
    $debugCheckBox.FontSize = 11
    $debugCheckBox.ToolTip = "Enable verbose logging for troubleshooting. Logs additional details during file processing and upload."
    $optionsPanel.Children.Add($debugCheckBox)
    $script:debugCheckBox = $debugCheckBox

    # Enable File Deletion Checkbox
    $fileDeletionCheckBox = New-Object Windows.Controls.CheckBox
    $fileDeletionCheckBox.Content = "Enable File Deletion"
    $fileDeletionCheckBox.IsChecked = $settings.enableFileDeletion
    $fileDeletionCheckBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
    $fileDeletionCheckBox.FontWeight = 'Normal'
    $fileDeletionCheckBox.Margin = '0,0,0,8'
    $fileDeletionCheckBox.FontSize = 11
    $fileDeletionCheckBox.ToolTip = "Allow automatic deletion of archived files older than the specified retention period. Uncheck to keep all files indefinitely."
    $optionsPanel.Children.Add($fileDeletionCheckBox)
    $script:fileDeletionCheckBox = $fileDeletionCheckBox

    # Save Settings Button
    $saveButton = New-Object Windows.Controls.Button
    $saveButton.Content = "Save Settings"
    $saveButton.Padding = '15,10'
    $saveButton.Margin = '0,20,0,15'
    $saveButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#28A745")
    $saveButton.Foreground = [System.Windows.Media.Brushes]::White
    $saveButton.FontWeight = 'Bold'
    $saveButton.FontSize = 13
    $saveButton.BorderThickness = '0'
    $saveButton.Cursor = 'Hand'
    $tab2StackPanel.Children.Add($saveButton)

    $saveButton.Add_Click({
        # Update settings object
        $settings.ParentDirectory = $script:textBoxes['ParentDirectory'].Text
        $settings.AppFolder = $script:textBoxes['AppFolder'].Text
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

    # RIGHT PANEL: Execution Log Viewer
    $execLogViewerBox = New-Object Windows.Controls.GroupBox
    $execLogViewerBox.Header = "Execution Log Content"
    $execLogViewerBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $execLogViewerBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FFC107")
    $execLogViewerBox.BorderThickness = '2'
    $execLogViewerBox.Margin = '8,0,0,0'
    $execLogViewerBox.Padding = '12'
    $execLogViewerBox.FontWeight = 'Bold'
    [Windows.Controls.Grid]::SetColumn($execLogViewerBox, 1)
    $contentGrid2.Children.Add($execLogViewerBox)

    # Execution log viewer textbox
    $execLogTextBox = New-Object Windows.Controls.TextBox
    $execLogTextBox.AcceptsReturn = $true
    $execLogTextBox.IsReadOnly = $true
    $execLogTextBox.TextWrapping = 'Wrap'
    $execLogTextBox.VerticalScrollBarVisibility = 'Auto'
    $execLogTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $execLogTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $execLogTextBox.Padding = '8,6'
    $execLogTextBox.FontSize = 9
    $execLogTextBox.FontFamily = New-Object Windows.Media.FontFamily("Consolas")
    $execLogTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $execLogTextBox.BorderThickness = '1'
    $execLogTextBox.Text = "Select an execution log from the dropdown above to view its contents..."
    $execLogViewerBox.Content = $execLogTextBox
    $script:execLogTextBox = $execLogTextBox

    # Function to reload app logs for currently selected app
    $reloadAppLogs = {
        if ($script:selectedAppName -and $script:currentAppLogFolder -and (Test-Path $script:currentAppLogFolder)) {
            $logFiles = Get-ChildItem -Path $script:currentAppLogFolder -Filter "*.csv" | Sort-Object LastWriteTime -Descending
            if ($logFiles.Count -gt 0) {
                $latestLog = Import-Csv -Path $logFiles[0].FullName
                $logText = "Latest log: $($logFiles[0].Name)`r`n`r`n"
                # Reverse to show most recent first, add padding for better readability
                $logLines = @($latestLog) | ForEach-Object { 
                    $logType = $_.'Log Type'.PadRight(7)
                    "$($_.'Date/Time') [$logType] $($_.'Log Details')" 
                }
                [array]::Reverse($logLines)
                $logText += $logLines -join "`r`n"
                $script:appLogTextBox.Text = $logText
            }
        }
    }

    # Function to load execution log files
    $loadExecutionLogDates = {
        $script:execLogDateDropdown.Items.Clear()
        
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
            $headerLabel.FontSize = 16
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
            $selectAllBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#28A745")
            $selectAllBtn.Foreground = [System.Windows.Media.Brushes]::White
            $selectAllBtn.FontSize = 11
            $selectAllBtn.BorderThickness = '0'
            $selectAllBtn.Cursor = 'Hand'
            $buttonStack.Children.Add($selectAllBtn)
            
            $deselectAllBtn = New-Object Windows.Controls.Button
            $deselectAllBtn.Content = "Deselect All"
            $deselectAllBtn.Padding = '10,5'
            $deselectAllBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#6C757D")
            $deselectAllBtn.Foreground = [System.Windows.Media.Brushes]::White
            $deselectAllBtn.FontSize = 11
            $deselectAllBtn.BorderThickness = '0'
            $deselectAllBtn.Cursor = 'Hand'
            $buttonStack.Children.Add($deselectAllBtn)
            
            # Description
            $descLabel = New-Object Windows.Controls.TextBlock
            $descLabel.Text = "Select apps to create/maintain directories for. Toggle 'Enable Upload' to control whether processed files are uploaded to SailPoint."
            $descLabel.TextWrapping = 'Wrap'
            $descLabel.Margin = '0,0,0,10'
            $descLabel.FontSize = 11
            $descLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#666666")
            $mainPanel.Children.Add($descLabel)
            
            # Create outer border for the entire table
            $tableBorder = New-Object Windows.Controls.Border
            $tableBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DEE2E6")
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
            $headerGrid.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#E9ECEF")
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
            $headerCheckBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DEE2E6")
            $headerCheckBorder.BorderThickness = '0,0,1,0'
            [Windows.Controls.Grid]::SetColumn($headerCheckBorder, 0)
            $headerCheckLabel = New-Object Windows.Controls.Label
            $headerCheckLabel.Content = "✓"
            $headerCheckLabel.FontWeight = 'Bold'
            $headerCheckLabel.FontSize = 12
            $headerCheckLabel.HorizontalAlignment = 'Center'
            $headerCheckLabel.VerticalAlignment = 'Center'
            $headerCheckLabel.Padding = '5'
            $headerCheckBorder.Child = $headerCheckLabel
            $headerGrid.Children.Add($headerCheckBorder)
            
            $headerNameBorder = New-Object Windows.Controls.Border
            $headerNameBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DEE2E6")
            $headerNameBorder.BorderThickness = '0,0,1,0'
            [Windows.Controls.Grid]::SetColumn($headerNameBorder, 1)
            $headerNameLabel = New-Object Windows.Controls.Label
            $headerNameLabel.Content = "App Name"
            $headerNameLabel.FontWeight = 'Bold'
            $headerNameLabel.FontSize = 11
            $headerNameLabel.VerticalAlignment = 'Center'
            $headerNameLabel.Padding = '10,5'
            $headerNameBorder.Child = $headerNameLabel
            $headerGrid.Children.Add($headerNameBorder)
            
            $headerStatusBorder = New-Object Windows.Controls.Border
            $headerStatusBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DEE2E6")
            $headerStatusBorder.BorderThickness = '0,0,1,0'
            [Windows.Controls.Grid]::SetColumn($headerStatusBorder, 2)
            $headerStatusLabel = New-Object Windows.Controls.Label
            $headerStatusLabel.Content = "Directory Status"
            $headerStatusLabel.FontWeight = 'Bold'
            $headerStatusLabel.FontSize = 11
            $headerStatusLabel.HorizontalAlignment = 'Center'
            $headerStatusLabel.VerticalAlignment = 'Center'
            $headerStatusLabel.Padding = '5'
            $headerStatusBorder.Child = $headerStatusLabel
            $headerGrid.Children.Add($headerStatusBorder)
            
            $headerUploadLabel = New-Object Windows.Controls.Label
            $headerUploadLabel.Content = "Enable Upload"
            $headerUploadLabel.FontWeight = 'Bold'
            $headerUploadLabel.FontSize = 11
            $headerUploadLabel.HorizontalAlignment = 'Center'
            $headerUploadLabel.VerticalAlignment = 'Center'
            $headerUploadLabel.Padding = '5'
            [Windows.Controls.Grid]::SetColumn($headerUploadLabel, 3)
            $headerGrid.Children.Add($headerUploadLabel)
            
            # Add separator line below header
            $headerSeparator = New-Object Windows.Controls.Border
            $headerSeparator.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DEE2E6")
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
            
            foreach ($source in $sourcesList | Sort-Object SourceName) {
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
                $rowBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DEE2E6")
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
                $checkboxBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DEE2E6")
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
                $nameBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DEE2E6")
                $nameBorder.BorderThickness = '0,0,1,0'
                [Windows.Controls.Grid]::SetColumn($nameBorder, 1)
                $nameLabel = New-Object Windows.Controls.Label
                $nameLabel.Content = $appName
                $nameLabel.FontSize = 11
                $nameLabel.VerticalAlignment = 'Center'
                $nameLabel.Padding = '10,5'
                $nameBorder.Child = $nameLabel
                $rowGrid.Children.Add($nameBorder)
                
                # Status column with border
                $statusBorder = New-Object Windows.Controls.Border
                $statusBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DEE2E6")
                $statusBorder.BorderThickness = '0,0,1,0'
                [Windows.Controls.Grid]::SetColumn($statusBorder, 2)
                $statusLabel = New-Object Windows.Controls.Label
                if ($dirExists) {
                    $statusLabel.Content = "✓ Exists"
                    $statusLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#28A745")
                } else {
                    $statusLabel.Content = "Not Created"
                    $statusLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#999999")
                }
                $statusLabel.FontSize = 10
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
                $uploadToggle.FontSize = 10
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
            $cancelBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#6C757D")
            $cancelBtn.Foreground = [System.Windows.Media.Brushes]::White
            $cancelBtn.FontSize = 12
            $cancelBtn.BorderThickness = '0'
            $cancelBtn.Cursor = 'Hand'
            $bottomPanel.Children.Add($cancelBtn)
            
            $createBtn = New-Object Windows.Controls.Button
            $createBtn.Content = "Apply Changes"
            $createBtn.Padding = '15,8'
            $createBtn.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
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
            }.GetNewClosure())
            
            # Show the selection window
            $selectionWindow.ShowDialog() | Out-Null
        }
        catch {
            Write-Log "ERROR: Failed to fetch sources. $_" "ERROR"
            [System.Windows.MessageBox]::Show("Failed to fetch sources from SailPoint: $_", "Error", "OK", "Error")
        }
    })

    # Function to load app folders into dropdown
    $loadAppFolders = {
        $script:appDropdown.Items.Clear()
        $script:selectedAppName = $null
        $appFolderPath = $script:textBoxes['AppFolder'].Text
        
        # Clear config editor
        $script:configFieldsPanel.Children.Clear()
        $script:configFields.Clear()
        $saveConfigButton.IsEnabled = $false
        $reloadConfigButton.IsEnabled = $false
        $uploadAppButton.IsEnabled = $false
        $uploadUserListButton.IsEnabled = $false
        $openAppLogButton.IsEnabled = $false
        $script:appLogTextBox.Text = "Select an app to view log information..."
        
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
                
                # Define group order
                $groupOrder = @('Source Configuration', 'File Structure', 'Column Operations', 'User Status & Roles', 'Entitlements & Groups', 'Other')
                
                # Create UI for each group
                foreach ($groupName in $groupOrder) {
                    if (-not $groups.ContainsKey($groupName)) { continue }
                    
                    # Group header
                    $groupHeader = New-Object Windows.Controls.Label
                    $groupHeader.Content = $groupName
                    $groupHeader.FontWeight = 'Bold'
                    $groupHeader.FontSize = 12
                    $groupHeader.Margin = '0,15,0,8'
                    $groupHeader.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#0366D6")
                    $script:configFieldsPanel.Children.Add($groupHeader)
                    
                    # Add separator
                    $separator = New-Object Windows.Controls.Separator
                    $separator.Margin = '0,0,0,10'
                    $script:configFieldsPanel.Children.Add($separator)
                    
                    # Create grid for fields in this group (2 columns)
                    $groupGrid = New-Object Windows.Controls.Grid
                    $groupGrid.Margin = '0,0,0,0'
                    
                    $col1 = New-Object Windows.Controls.ColumnDefinition
                    $col1.Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
                    $col2 = New-Object Windows.Controls.ColumnDefinition
                    $col2.Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
                    $groupGrid.ColumnDefinitions.Add($col1)
                    $groupGrid.ColumnDefinitions.Add($col2)
                    
                    $rowIndex = 0
                    $columnIndex = 0
                    
                    # Add fields to grid
                    foreach ($field in $groups[$groupName]) {
                        $propName = $field.Name
                        $propValue = $field.Value
                        $metadata = $field.Metadata
                        
                        # Create stack panel for label and control
                        $fieldStack = New-Object Windows.Controls.StackPanel
                        $fieldStack.Margin = '0,0,10,12'
                        
                        # Create label
                        $label = New-Object Windows.Controls.Label
                        $label.Content = $metadata.Label
                        $label.FontWeight = 'SemiBold'
                        $label.FontSize = 10
                        $label.Margin = '0,0,0,2'
                        $label.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
                        $label.ToolTip = $metadata.Tooltip
                        $fieldStack.Children.Add($label)
                        
                        # Create appropriate input control based on value type
                        if ($propValue -is [bool]) {
                            $checkBox = New-Object Windows.Controls.CheckBox
                            $checkBox.IsChecked = $propValue
                            $checkBox.ToolTip = $metadata.Tooltip
                            $fieldStack.Children.Add($checkBox)
                            $script:configFields[$propName] = $checkBox
                        }
                        elseif ($propValue -is [array]) {
                            $textBox = New-Object Windows.Controls.TextBox
                            if ($propValue -and $propValue.Count -gt 0) {
                                $textBox.Text = ($propValue -join ', ')
                            }
                            else {
                                $textBox.Text = ""
                            }
                            $textBox.Padding = '6,4'
                            $textBox.FontSize = 10
                            $textBox.Background = [System.Windows.Media.Brushes]::White
                            $textBox.ToolTip = $metadata.Tooltip + " (comma-separated values)"
                            $fieldStack.Children.Add($textBox)
                            $script:configFields[$propName] = $textBox
                        }
                        else {
                            $textBox = New-Object Windows.Controls.TextBox
                            if ($null -ne $propValue) {
                                $textBox.Text = $propValue.ToString()
                            }
                            else {
                                $textBox.Text = ""
                            }
                            $textBox.Padding = '6,4'
                            $textBox.FontSize = 10
                            $textBox.Background = [System.Windows.Media.Brushes]::White
                            $textBox.ToolTip = $metadata.Tooltip
                            
                            # Make sourceID field read-only
                            if ($propName -eq 'sourceID') {
                                $textBox.IsReadOnly = $true
                                $textBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#E8E8E8")
                                $textBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#666666")
                                $textBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
                                $textBox.BorderThickness = '1'
                                $textBox.FontStyle = 'Italic'
                                $textBox.ToolTip = $metadata.Tooltip + " (Read-only - managed by system)"
                            }
                            
                            $fieldStack.Children.Add($textBox)
                            $script:configFields[$propName] = $textBox
                        }
                        
                        # Add row definition if needed
                        if ($columnIndex -eq 0) {
                            $rowDef = New-Object Windows.Controls.RowDefinition
                            $rowDef.Height = [Windows.GridLength]::Auto
                            $groupGrid.RowDefinitions.Add($rowDef)
                        }
                        
                        # Set grid position
                        [Windows.Controls.Grid]::SetRow($fieldStack, $rowIndex)
                        [Windows.Controls.Grid]::SetColumn($fieldStack, $columnIndex)
                        $groupGrid.Children.Add($fieldStack)
                        
                        # Move to next column/row
                        $columnIndex++
                        if ($columnIndex -ge 2) {
                            $columnIndex = 0
                            $rowIndex++
                        }
                    }
                    
                    $script:configFieldsPanel.Children.Add($groupGrid)
                }
                
                $saveConfigButton.IsEnabled = $true
                $reloadConfigButton.IsEnabled = $true
                $uploadAppButton.IsEnabled = $true
                $uploadUserListButton.IsEnabled = $true
                $script:currentConfigPath = $configFilePath
                
                # Enable app log button
                $openAppLogButton.IsEnabled = $true
                $script:currentAppLogFolder = $logFolderPath
                
                # Load most recent log file
                if (Test-Path $logFolderPath) {
                    $logFiles = Get-ChildItem -Path $logFolderPath -Filter "*.csv" | Sort-Object LastWriteTime -Descending
                    if ($logFiles.Count -gt 0) {
                        $latestLog = Import-Csv -Path $logFiles[0].FullName
                        $logText = "Latest log: $($logFiles[0].Name)`r`n`r`n"
                        # Reverse to show most recent first, add padding for better readability
                        $logLines = @($latestLog) | ForEach-Object { 
                            $logType = $_.'Log Type'.PadRight(7)
                            "$($_.'Date/Time') [$logType] $($_.'Log Details')" 
                        }
                        [array]::Reverse($logLines)
                        $logText += $logLines -join "`r`n"
                        $script:appLogTextBox.Text = $logText
                    }
                    else {
                        $script:appLogTextBox.Text = "No log files found"
                    }
                }
                else {
                    $script:appLogTextBox.Text = "Log folder not found"
                }
            }
            catch {
                $script:selectedAppLabel.Content = "Error loading config: $_"
                $script:configFieldsPanel.Children.Clear()
                $saveConfigButton.IsEnabled = $false
                $reloadConfigButton.IsEnabled = $false
                $uploadAppButton.IsEnabled = $false
                $uploadUserListButton.IsEnabled = $false
                $openAppLogButton.IsEnabled = $false
                $script:appLogTextBox.Text = "Error loading app: $_"
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
                
                foreach ($key in $script:configFields.Keys) {
                    $control = $script:configFields[$key]
                    
                    if ($control -is [Windows.Controls.CheckBox]) {
                        $configObj[$key] = $control.IsChecked
                    }
                    elseif ($control -is [Windows.Controls.TextBox]) {
                        $value = $control.Text.Trim()
                        
                        # Try to parse as number
                        $numValue = 0
                        if ([int]::TryParse($value, [ref]$numValue)) {
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
        $closeBtn.FontSize = 10
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
        $logTextBox.FontSize = 11
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
