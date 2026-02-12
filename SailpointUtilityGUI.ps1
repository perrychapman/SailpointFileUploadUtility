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
    $logMessage = "[$timestamp] [$type] $message"

    if ($script:logTextBoxInitialized -and $script:logTextBox) {
        $script:logTextBox.Dispatcher.Invoke([action]{
            $script:logTextBox.AppendText("$logMessage`r`n")
            $script:logTextBox.ScrollToEnd()
        })
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
        [PSObject]$settings
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
            Write-Log "AppList_$tenant.csv updated with $($newSourcesList.Count) new sources."
        }
        else {
            Write-Log "No new sources to add. AppList_$tenant.csv is up to date."
        }
    }
    catch {
        Write-Log "ERROR: Failed to fetch or process sources. $_" "ERROR"
        return $false
    }

    # Import and filter sources for Delimited File type
    $folders = Import-Csv -Path $csvPath | Sort-Object SourceName
    $delimitedFileSources = $folders | Where-Object { $_.SourceType -eq "Delimited File" }
    
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
    "isMonarch": false,
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
# File Upload Logic
# ------------------------

function Run-FileUpload {
    param (
        [PSObject]$settings
    )

    Write-Log "=== Starting File Upload Process ===" "INFO"

    try {
        # Build arguments for FileUploadScript.ps1
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "FileUploadScript.ps1"
        
        if (-not (Test-Path -Path $scriptPath)) {
            Write-Log "ERROR: FileUploadScript.ps1 not found at $scriptPath" "ERROR"
            return $false
        }

        Write-Log "Executing FileUploadScript.ps1..."
        
        # Execute the script and capture output
        $output = & $scriptPath 2>&1
        
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

        Write-Log "=== File Upload Process Complete ===" "INFO"
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

    # Define rows - header, content, log
    $headerRow = New-Object Windows.Controls.RowDefinition
    $headerRow.Height = "Auto"
    $contentRow = New-Object Windows.Controls.RowDefinition
    $contentRow.Height = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
    $logRow = New-Object Windows.Controls.RowDefinition
    $logRow.Height = New-Object Windows.GridLength(200)
    $mainGrid.RowDefinitions.Add($headerRow)
    $mainGrid.RowDefinitions.Add($contentRow)
    $mainGrid.RowDefinitions.Add($logRow)

    # Define 3 columns for content section: Settings | App Browser | Actions
    $leftColumn = New-Object Windows.Controls.ColumnDefinition
    $leftColumn.Width = New-Object Windows.GridLength(1, [Windows.GridUnitType]::Star)
    $leftColumn.MinWidth = 320
    $mainGrid.ColumnDefinitions.Add($leftColumn)

    $middleColumn = New-Object Windows.Controls.ColumnDefinition
    $middleColumn.Width = New-Object Windows.GridLength(1.3, [Windows.GridUnitType]::Star)
    $middleColumn.MinWidth = 380
    $mainGrid.ColumnDefinitions.Add($middleColumn)

    $rightColumn = New-Object Windows.Controls.ColumnDefinition
    $rightColumn.Width = New-Object Windows.GridLength(0.7, [Windows.GridUnitType]::Star)
    $rightColumn.MinWidth = 200
    $mainGrid.ColumnDefinitions.Add($rightColumn)

    # ===== HEADER SECTION =====
    $headerPanel = New-Object Windows.Controls.Border
    $headerPanel.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $headerPanel.Padding = '20,15'
    [Windows.Controls.Grid]::SetRow($headerPanel, 0)
    [Windows.Controls.Grid]::SetColumn($headerPanel, 0)
    [Windows.Controls.Grid]::SetColumnSpan($headerPanel, 3)
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
    $subtitleLabel.Content = "Management Console - Configure, Create Directories, and Upload Files"
    $subtitleLabel.FontSize = 12
    $subtitleLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#AAAAAA")
    $subtitleLabel.Margin = '-5,0,0,0'
    $headerStack.Children.Add($subtitleLabel)

    # ===== LEFT PANEL - Settings =====
    $leftBorder = New-Object Windows.Controls.Border
    $leftBorder.Background = [System.Windows.Media.Brushes]::White
    $leftBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#E0E0E0")
    $leftBorder.BorderThickness = '0,0,1,0'
    $leftBorder.Margin = '0'
    [Windows.Controls.Grid]::SetColumn($leftBorder, 0)
    [Windows.Controls.Grid]::SetRow($leftBorder, 1)
    $mainGrid.Children.Add($leftBorder)

    $leftPanel = New-Object Windows.Controls.ScrollViewer
    $leftPanel.VerticalScrollBarVisibility = 'Auto'
    $leftPanel.HorizontalScrollBarVisibility = 'Disabled'
    $leftPanel.Padding = '15'
    $leftBorder.Child = $leftPanel

    $leftStackPanel = New-Object Windows.Controls.StackPanel
    $leftStackPanel.Margin = '0'
    $leftPanel.Content = $leftStackPanel

    # Settings Section Title
    $settingsTitle = New-Object Windows.Controls.Label
    $settingsTitle.Content = "⚙️ Configuration Settings"
    $settingsTitle.FontSize = 16
    $settingsTitle.FontWeight = 'Bold'
    $settingsTitle.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $settingsTitle.Margin = '0,0,0,15'
    $leftStackPanel.Children.Add($settingsTitle)

    # Directory Locations GroupBox
    $directoryBox = New-Object Windows.Controls.GroupBox
    $directoryBox.Header = "📁 Directory Locations"
    $directoryBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $directoryBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $directoryBox.BorderThickness = '2'
    $directoryBox.Margin = '0,0,0,15'
    $directoryBox.Padding = '15,10,15,15'
    $directoryBox.FontWeight = 'Bold'
    $directoryPanel = New-Object Windows.Controls.StackPanel
    $directoryPanel.Orientation = 'Vertical'
    $directoryBox.Content = $directoryPanel
    $leftStackPanel.Children.Add($directoryBox)

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
        $browseButton.Content = "📁"
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

        # Add browse button click handler
        $browseButton.Add_Click({
            param($sender, $e)
            $fieldName = $null
            $fieldType = $null
            
            # Find which field this button belongs to
            foreach ($f in $directoryFields) {
                if ($script:textBoxes[$f.Name] -eq $textBox) {
                    $fieldName = $f.Name
                    $fieldType = $f.Type
                    break
                }
            }
            
            if ($fieldType -eq "Folder") {
                $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
                $folderBrowser.Description = "Select " + $fieldName
                $folderBrowser.ShowNewFolderButton = $true
                
                if ($textBox.Text -and (Test-Path $textBox.Text)) {
                    $folderBrowser.SelectedPath = $textBox.Text
                }
                
                if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $textBox.Text = $folderBrowser.SelectedPath
                }
            }
            else {
                $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog
                $fileBrowser.Title = "Select " + $fieldName
                $fileBrowser.Filter = "JAR Files (*.jar)|*.jar|All Files (*.*)|*.*"
                
                if ($textBox.Text -and (Test-Path (Split-Path $textBox.Text -Parent))) {
                    $fileBrowser.InitialDirectory = Split-Path $textBox.Text -Parent
                    $fileBrowser.FileName = Split-Path $textBox.Text -Leaf
                }
                
                if ($fileBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $textBox.Text = $fileBrowser.FileName
                }
            }
        }.GetNewClosure())

        $directoryPanel.Children.Add($pathStack)
    }

    # Client Credentials GroupBox
    $credentialsBox = New-Object Windows.Controls.GroupBox
    $credentialsBox.Header = "🔐 SailPoint Client Credentials"
    $credentialsBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $credentialsBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $credentialsBox.BorderThickness = '2'
    $credentialsBox.Margin = '0,0,0,15'
    $credentialsBox.Padding = '15,10,15,15'
    $credentialsBox.FontWeight = 'Bold'
    $credentialsPanel = New-Object Windows.Controls.StackPanel
    $credentialsPanel.Orientation = 'Vertical'
    $credentialsBox.Content = $credentialsPanel
    $leftStackPanel.Children.Add($credentialsBox)

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
    $optionsBox.Header = "⚡ Options"
    $optionsBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $optionsBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $optionsBox.BorderThickness = '2'
    $optionsBox.Margin = '0,0,0,15'
    $optionsBox.Padding = '15,10,15,15'
    $optionsBox.FontWeight = 'Bold'
    $optionsPanel = New-Object Windows.Controls.StackPanel
    $optionsPanel.Orientation = 'Vertical'
    $optionsBox.Content = $optionsPanel
    $leftStackPanel.Children.Add($optionsBox)

    # App Filter
    $appFilterLabel = New-Object Windows.Controls.Label
    $appFilterLabel.Content = "App Filter (optional - for file upload)"
    $appFilterLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
    $appFilterLabel.FontWeight = 'Normal'
    $appFilterLabel.Margin = '0,8,0,3'
    $appFilterLabel.FontSize = 11
    $appFilterLabel.ToolTip = "Process only specific app folders during upload. Leave empty to process all apps. Separate multiple apps with commas."
    $optionsPanel.Children.Add($appFilterLabel)

    $appFilterTextBox = New-Object Windows.Controls.TextBox
    $appFilterTextBox.Text = $settings.AppFilter
    $appFilterTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $appFilterTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $appFilterTextBox.Padding = '8,6'
    $appFilterTextBox.FontSize = 11
    $appFilterTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $appFilterTextBox.BorderThickness = '1'
    $appFilterTextBox.ToolTip = "Process only specific app folders during upload. Leave empty to process all apps. Separate multiple apps with commas."
    $optionsPanel.Children.Add($appFilterTextBox)
    $script:textBoxes['AppFilter'] = $appFilterTextBox

    # Days to Keep Files
    $daysLabel = New-Object Windows.Controls.Label
    $daysLabel.Content = "Days to Keep Files"
    $daysLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#333333")
    $daysLabel.FontWeight = 'Normal'
    $daysLabel.Margin = '0,8,0,3'
    $daysLabel.FontSize = 11
    $daysLabel.ToolTip = "Number of days to retain processed files in the Archive folder before automatic deletion. Default: 90 days."
    $optionsPanel.Children.Add($daysLabel)

    $daysTextBox = New-Object Windows.Controls.TextBox
    $daysTextBox.Text = $settings.DaysToKeepFiles.ToString()
    $daysTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $daysTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $daysTextBox.Padding = '8,6'
    $daysTextBox.FontSize = 11
    $daysTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $daysTextBox.BorderThickness = '1'
    $daysTextBox.ToolTip = "Number of days to retain processed files in the Archive folder before automatic deletion. Default: 90 days."
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
    $saveButton.Content = "💾 Save Settings"
    $saveButton.Padding = '15,10'
    $saveButton.Margin = '0,20,0,15'
    $saveButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#28A745")
    $saveButton.Foreground = [System.Windows.Media.Brushes]::White
    $saveButton.FontWeight = 'Bold'
    $saveButton.FontSize = 13
    $saveButton.BorderThickness = '0'
    $saveButton.Cursor = 'Hand'
    $leftStackPanel.Children.Add($saveButton)

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
        $settings.AppFilter = $script:textBoxes['AppFilter'].Text
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

    # ===== MIDDLE PANEL - App Folder Browser & Config Editor =====
    $middleBorder = New-Object Windows.Controls.Border
    $middleBorder.Background = [System.Windows.Media.Brushes]::White
    $middleBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#E0E0E0")
    $middleBorder.BorderThickness = '0,0,1,0'
    $middleBorder.Margin = '0'
    [Windows.Controls.Grid]::SetColumn($middleBorder, 1)
    [Windows.Controls.Grid]::SetRow($middleBorder, 1)
    $mainGrid.Children.Add($middleBorder)

    $middleScrollViewer = New-Object Windows.Controls.ScrollViewer
    $middleScrollViewer.VerticalScrollBarVisibility = 'Auto'
    $middleScrollViewer.HorizontalScrollBarVisibility = 'Disabled'
    $middleScrollViewer.Padding = '15'
    $middleBorder.Child = $middleScrollViewer

    $middleStackPanel = New-Object Windows.Controls.StackPanel
    $middleStackPanel.Margin = '0'
    $middleScrollViewer.Content = $middleStackPanel

    # App Browser Section Title
    $appBrowserTitle = New-Object Windows.Controls.Label
    $appBrowserTitle.Content = "📂 App Folder Browser"
    $appBrowserTitle.FontSize = 16
    $appBrowserTitle.FontWeight = 'Bold'
    $appBrowserTitle.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $appBrowserTitle.Margin = '0,0,0,15'
    $middleStackPanel.Children.Add($appBrowserTitle)

    # Refresh button
    $refreshAppsButton = New-Object Windows.Controls.Button
    $refreshAppsButton.Content = "🔄 Refresh App List"
    $refreshAppsButton.Padding = '12,8'
    $refreshAppsButton.Margin = '0,0,0,15'
    $refreshAppsButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $refreshAppsButton.Foreground = [System.Windows.Media.Brushes]::White
    $refreshAppsButton.FontWeight = 'SemiBold'
    $refreshAppsButton.FontSize = 11
    $refreshAppsButton.BorderThickness = '0'
    $refreshAppsButton.Cursor = 'Hand'
    $middleStackPanel.Children.Add($refreshAppsButton)

    # App folders list
    $script:appFoldersList = New-Object Windows.Controls.ListBox
    $script:appFoldersList.Height = 220
    $script:appFoldersList.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $script:appFoldersList.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $script:appFoldersList.BorderThickness = '1'
    $script:appFoldersList.Margin = '0,0,0,15'
    $script:appFoldersList.FontSize = 11
    $middleStackPanel.Children.Add($script:appFoldersList)

    # Config Editor Section
    $configEditorBox = New-Object Windows.Controls.GroupBox
    $configEditorBox.Header = "⚙️ Config.json Editor"
    $configEditorBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $configEditorBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#28A745")
    $configEditorBox.BorderThickness = '2'
    $configEditorBox.Margin = '0,0,0,15'
    $configEditorBox.Padding = '12'
    $configEditorBox.FontWeight = 'Bold'
    $configEditorPanel = New-Object Windows.Controls.StackPanel
    $configEditorPanel.Orientation = 'Vertical'
    $configEditorBox.Content = $configEditorPanel
    $middleStackPanel.Children.Add($configEditorBox)

    # Selected app label
    $script:selectedAppLabel = New-Object Windows.Controls.Label
    $script:selectedAppLabel.Content = "No app selected"
    $script:selectedAppLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#666666")
    $script:selectedAppLabel.FontWeight = 'Normal'
    $script:selectedAppLabel.FontSize = 11
    $script:selectedAppLabel.Margin = '0,0,0,10'
    $configEditorPanel.Children.Add($script:selectedAppLabel)

    # Config editor textbox
    $script:configTextBox = New-Object Windows.Controls.TextBox
    $script:configTextBox.Height = 180
    $script:configTextBox.AcceptsReturn = $true
    $script:configTextBox.TextWrapping = 'NoWrap'
    $script:configTextBox.VerticalScrollBarVisibility = 'Auto'
    $script:configTextBox.HorizontalScrollBarVisibility = 'Auto'
    $script:configTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F9F9F9")
    $script:configTextBox.Foreground = [System.Windows.Media.Brushes]::Black
    $script:configTextBox.Padding = '8,6'
    $script:configTextBox.FontSize = 10
    $script:configTextBox.FontFamily = New-Object Windows.Media.FontFamily("Consolas")
    $script:configTextBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#CCCCCC")
    $script:configTextBox.BorderThickness = '1'
    $script:configTextBox.IsEnabled = $false
    $configEditorPanel.Children.Add($script:configTextBox)

    # Config editor buttons
    $configButtonsPanel = New-Object Windows.Controls.StackPanel
    $configButtonsPanel.Orientation = 'Horizontal'
    $configButtonsPanel.Margin = '0,10,0,0'
    $configEditorPanel.Children.Add($configButtonsPanel)

    $saveConfigButton = New-Object Windows.Controls.Button
    $saveConfigButton.Content = "💾 Save Config"
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
    $reloadConfigButton.Content = "🔄 Reload"
    $reloadConfigButton.Padding = '12,6'
    $reloadConfigButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#6C757D")
    $reloadConfigButton.Foreground = [System.Windows.Media.Brushes]::White
    $reloadConfigButton.FontWeight = 'SemiBold'
    $reloadConfigButton.FontSize = 10
    $reloadConfigButton.BorderThickness = '0'
    $reloadConfigButton.Cursor = 'Hand'
    $reloadConfigButton.IsEnabled = $false
    $configButtonsPanel.Children.Add($reloadConfigButton)

    # Function to load app folders
    $loadAppFolders = {
        $script:appFoldersList.Items.Clear()
        $appFolderPath = $script:textBoxes['AppFolder'].Text
        
        if ([string]::IsNullOrWhiteSpace($appFolderPath) -or -not (Test-Path $appFolderPath)) {
            $script:appFoldersList.Items.Add("⚠️ App Folder path not valid")
            return
        }

        try {
            $appDirs = Get-ChildItem -Path $appFolderPath -Directory | Sort-Object Name
            
            if ($appDirs.Count -eq 0) {
                $script:appFoldersList.Items.Add("📭 No app folders found")
            }
            else {
                foreach ($dir in $appDirs) {
                    $configPath = Join-Path $dir.FullName "config.json"
                    if (Test-Path $configPath) {
                        $script:appFoldersList.Items.Add("📁 " + $dir.Name)
                    }
                    else {
                        $script:appFoldersList.Items.Add("❌ " + $dir.Name + " (no config.json)")
                    }
                }
            }
        }
        catch {
            $script:appFoldersList.Items.Add("❌ Error loading folders: $_")
        }
    }

    # Refresh apps button handler
    $refreshAppsButton.Add_Click({
        & $loadAppFolders
    })

    # App folder selection handler
    $script:appFoldersList.Add_SelectionChanged({
        $selectedItem = $script:appFoldersList.SelectedItem
        
        if ($null -eq $selectedItem -or $selectedItem -like "⚠️*" -or $selectedItem -like "📭*" -or $selectedItem -like "❌*") {
            $script:selectedAppLabel.Content = "No valid app selected"
            $script:configTextBox.Text = ""
            $script:configTextBox.IsEnabled = $false
            $saveConfigButton.IsEnabled = $false
            $reloadConfigButton.IsEnabled = $false
            return
        }

        # Extract folder name (remove emoji prefix)
        $folderName = $selectedItem -replace '^📁\s+', ''
        $appFolderPath = $script:textBoxes['AppFolder'].Text
        $configPath = Join-Path (Join-Path $appFolderPath $folderName) "config.json"
        
        if (Test-Path $configPath) {
            try {
                $configContent = Get-Content -Path $configPath -Raw
                $script:configTextBox.Text = $configContent
                $script:configTextBox.IsEnabled = $true
                $script:selectedAppLabel.Content = "Editing: $folderName/config.json"
                $saveConfigButton.IsEnabled = $true
                $reloadConfigButton.IsEnabled = $true
                $script:currentConfigPath = $configPath
            }
            catch {
                $script:selectedAppLabel.Content = "Error loading config: $_"
                $script:configTextBox.Text = ""
                $script:configTextBox.IsEnabled = $false
                $saveConfigButton.IsEnabled = $false
                $reloadConfigButton.IsEnabled = $false
            }
        }
    })

    # Save config button handler
    $saveConfigButton.Add_Click({
        if ($null -ne $script:currentConfigPath) {
            try {
                # Validate JSON before saving
                $jsonTest = $script:configTextBox.Text | ConvertFrom-Json
                
                $script:configTextBox.Text | Out-File -FilePath $script:currentConfigPath -Encoding UTF8
                [System.Windows.MessageBox]::Show("Config saved successfully!", "Success", "OK", "Information")
                Write-Log "Saved config: $script:currentConfigPath" "INFO"
            }
            catch {
                [System.Windows.MessageBox]::Show("Invalid JSON format! Please fix syntax errors.`n`nError: $_", "Validation Error", "OK", "Error")
            }
        }
    })

    # Reload config button handler
    $reloadConfigButton.Add_Click({
        if ($null -ne $script:currentConfigPath -and (Test-Path $script:currentConfigPath)) {
            try {
                $configContent = Get-Content -Path $script:currentConfigPath -Raw
                $script:configTextBox.Text = $configContent
                Write-Log "Reloaded config from disk" "INFO"
            }
            catch {
                [System.Windows.MessageBox]::Show("Error reloading config: $_", "Error", "OK", "Error")
            }
        }
    })

    # Load app folders on startup
    & $loadAppFolders

    # ===== RIGHT PANEL - Action Buttons =====
    $rightBorder = New-Object Windows.Controls.Border
    $rightBorder.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#FAFAFA")
    $rightBorder.Margin = '0'
    [Windows.Controls.Grid]::SetColumn($rightBorder, 2)
    [Windows.Controls.Grid]::SetRow($rightBorder, 1)
    $mainGrid.Children.Add($rightBorder)

    $rightScrollViewer = New-Object Windows.Controls.ScrollViewer
    $rightScrollViewer.VerticalScrollBarVisibility = 'Auto'
    $rightScrollViewer.HorizontalScrollBarVisibility = 'Disabled'
    $rightScrollViewer.Padding = '15'
    $rightBorder.Child = $rightScrollViewer

    $rightPanel = New-Object Windows.Controls.StackPanel
    $rightScrollViewer.Content = $rightPanel

    # Action Buttons Title
    $actionsLabel = New-Object Windows.Controls.Label
    $actionsLabel.Content = "🚀 Run"
    $actionsLabel.FontSize = 14
    $actionsLabel.FontWeight = 'Bold'
    $actionsLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $actionsLabel.Margin = '0,0,0,12'
    $rightPanel.Children.Add($actionsLabel)

    # Info text
    $infoText = New-Object Windows.Controls.TextBlock
    $infoText.Text = "Execute these actions after configuring settings:"
    $infoText.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#666666")
    $infoText.TextWrapping = 'Wrap'
    $infoText.Margin = '0,0,0,15'
    $infoText.FontSize = 10
    $rightPanel.Children.Add($infoText)

    # Directory Creation Button
    $directoryButton = New-Object Windows.Controls.Button
    $directoryButton.Content = "📂 Create/Update`nDirectories"
    $directoryButton.Height = 55
    $directoryButton.Margin = '0,0,0,10'
    $directoryButton.Padding = '10,8'
    $directoryButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")
    $directoryButton.Foreground = [System.Windows.Media.Brushes]::White
    $directoryButton.FontSize = 11
    $directoryButton.FontWeight = 'SemiBold'
    $directoryButton.BorderThickness = '0'
    $directoryButton.Cursor = 'Hand'
    $directoryButton.HorizontalContentAlignment = 'Center'
    $rightPanel.Children.Add($directoryButton)

    $directoryButton.Add_Click({
        $script:logTextBox.Clear()
        Write-Log "Starting Directory Creation..." "INFO"
        
        # Reload settings in case they changed
        $currentSettings = Load-Settings
        $result = Run-DirectoryCreation -settings $currentSettings
        
        if ($result) {
            Write-Log "Directory creation completed successfully!" "INFO"
        }
        else {
            Write-Log "Directory creation encountered errors. Check the log above." "ERROR"
        }
    })

    # Description for Directory Creation
    $dirDescLabel = New-Object Windows.Controls.TextBlock
    $dirDescLabel.Text = "Connects to SailPoint and creates folder structures for Delimited File sources."
    $dirDescLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#999999")
    $dirDescLabel.TextWrapping = 'Wrap'
    $dirDescLabel.Margin = '0,0,0,20'
    $dirDescLabel.FontSize = 9
    $dirDescLabel.LineHeight = 14
    $rightPanel.Children.Add($dirDescLabel)

    # File Upload Button
    $uploadButton = New-Object Windows.Controls.Button
    $uploadButton.Content = "⬆️ Run File`nUpload Process"
    $uploadButton.Height = 55
    $uploadButton.Margin = '0,0,0,10'
    $uploadButton.Padding = '10,8'
    $uploadButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#28A745")
    $uploadButton.Foreground = [System.Windows.Media.Brushes]::White
    $uploadButton.FontSize = 11
    $uploadButton.FontWeight = 'SemiBold'
    $uploadButton.BorderThickness = '0'
    $uploadButton.Cursor = 'Hand'
    $uploadButton.HorizontalContentAlignment = 'Center'
    $uploadButton.HorizontalAlignment = 'Stretch'
    $rightPanel.Children.Add($uploadButton)

    $uploadButton.Add_Click({
        $script:logTextBox.Clear()
        Write-Log "Starting File Upload Process..." "INFO"
        
        # Reload settings in case they changed
        $currentSettings = Load-Settings
        $result = Run-FileUpload -settings $currentSettings
        
        if ($result) {
            Write-Log "File upload process completed!" "INFO"
        }
        else {
            Write-Log "File upload process encountered errors. Check the log above." "ERROR"
        }
    })

    # Description for File Upload
    $uploadDescLabel = New-Object Windows.Controls.TextBlock
    $uploadDescLabel.Text = "Processes and uploads CSV/Excel files from app folders to SailPoint."
    $uploadDescLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#999999")
    $uploadDescLabel.TextWrapping = 'Wrap'
    $uploadDescLabel.Margin = '0,0,0,20'
    $uploadDescLabel.FontSize = 9
    $uploadDescLabel.LineHeight = 14
    $rightPanel.Children.Add($uploadDescLabel)

    # Separator
    $separator = New-Object Windows.Controls.Separator
    $separator.Margin = '0,10,0,15'
    $separator.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DDDDDD")
    $rightPanel.Children.Add($separator)

    # Quick Actions Label
    $quickActionsLabel = New-Object Windows.Controls.Label
    $quickActionsLabel.Content = "🔧 Quick Links"
    $quickActionsLabel.FontSize = 11
    $quickActionsLabel.FontWeight = 'SemiBold'
    $quickActionsLabel.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $quickActionsLabel.Margin = '0,0,0,8'
    $rightPanel.Children.Add($quickActionsLabel)

    # Open Settings File Button
    $openSettingsButton = New-Object Windows.Controls.Button
    $openSettingsButton.Content = "📄 Settings File"
    $openSettingsButton.Height = 32
    $openSettingsButton.Margin = '0,0,0,6'
    $openSettingsButton.Padding = '10,6'
    $openSettingsButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#6C757D")
    $openSettingsButton.Foreground = [System.Windows.Media.Brushes]::White
    $openSettingsButton.FontSize = 10
    $openSettingsButton.BorderThickness = '0'
    $openSettingsButton.Cursor = 'Hand'
    $rightPanel.Children.Add($openSettingsButton)

    $openSettingsButton.Add_Click({
        if (Test-Path $script:settingsPath) {
            Start-Process notepad.exe -ArgumentList $script:settingsPath
        }
        else {
            [System.Windows.MessageBox]::Show("settings.json file not found.", "Error", "OK", "Error")
        }
    })

    # Open Import Folder Button
    $openImportButton = New-Object Windows.Controls.Button
    $openImportButton.Content = "📁 Import Folder"
    $openImportButton.Height = 32
    $openImportButton.Margin = '0,0,0,6'
    $openImportButton.Padding = '10,6'
    $openImportButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#6C757D")
    $openImportButton.Foreground = [System.Windows.Media.Brushes]::White
    $openImportButton.FontSize = 10
    $openImportButton.BorderThickness = '0'
    $openImportButton.Cursor = 'Hand'
    $rightPanel.Children.Add($openImportButton)

    $openImportButton.Add_Click({
        $importPath = $script:textBoxes['AppFolder'].Text
        if (Test-Path $importPath) {
            Start-Process explorer.exe -ArgumentList $importPath
        }
        else {
            [System.Windows.MessageBox]::Show("Import folder not found. Create directories first.", "Error", "OK", "Error")
        }
    })

    # ===== BOTTOM PANEL - Log Output =====
    $logBorder = New-Object Windows.Controls.Border
    $logBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DDDDDD")
    $logBorder.BorderThickness = '0,1,0,0'
    $logBorder.Background = [System.Windows.Media.Brushes]::White
    $logBorder.Margin = '0'
    [Windows.Controls.Grid]::SetRow($logBorder, 2)
    [Windows.Controls.Grid]::SetColumnSpan($logBorder, 3)
    $mainGrid.Children.Add($logBorder)

    $logPanel = New-Object Windows.Controls.DockPanel
    $logBorder.Child = $logPanel

    # Log Header
    $logHeaderBorder = New-Object Windows.Controls.Border
    $logHeaderBorder.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#F5F5F5")
    $logHeaderBorder.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#DDDDDD")
    $logHeaderBorder.BorderThickness = '0,0,0,1'
    $logHeaderBorder.Padding = '15,10'
    [Windows.Controls.DockPanel]::SetDock($logHeaderBorder, 'Top')
    $logPanel.Children.Add($logHeaderBorder)

    $logHeader = New-Object Windows.Controls.Label
    $logHeader.Content = "📋 Operation Log"
    $logHeader.FontSize = 13
    $logHeader.FontWeight = 'Bold'
    $logHeader.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")
    $logHeader.Padding = '0'
    $logHeaderBorder.Child = $logHeader

    # Log TextBox with ScrollViewer
    $logScrollViewer = New-Object Windows.Controls.ScrollViewer
    $logScrollViewer.VerticalScrollBarVisibility = 'Auto'
    $logScrollViewer.HorizontalScrollBarVisibility = 'Auto'
    $logPanel.Children.Add($logScrollViewer)

    $script:logTextBox = New-Object Windows.Controls.TextBox
    $script:logTextBox.AcceptsReturn = $true
    $script:logTextBox.IsReadOnly = $true
    $script:logTextBox.FontFamily = 'Consolas'
    $script:logTextBox.FontSize = 11
    $script:logTextBox.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1E1E1E")
    $script:logTextBox.Foreground = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#D4D4D4")
    $script:logTextBox.Padding = '15'
    $script:logTextBox.BorderThickness = '0'
    $script:logTextBox.TextWrapping = 'Wrap'
    $logScrollViewer.Content = $script:logTextBox

    $script:logTextBoxInitialized = $true

    # Show initial message
    Write-Log "═══════════════════════════════════════════════════════════════" "INFO"
    Write-Log "  SailPoint File Upload Utility - Management Console" "INFO"
    Write-Log "═══════════════════════════════════════════════════════════════" "INFO"
    Write-Log "" "INFO"
    Write-Log "Ready to begin. Select an action from the right panel." "INFO"
    Write-Log "" "INFO"

    # Show the window
    $window.ShowDialog() | Out-Null
}

# ------------------------
# Main Entry Point
# ------------------------

Show-MainWindow
