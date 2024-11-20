Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Function to run the main script logic
function Run-Script {
    # Load master settings
    $masterSettingsPath = ".\settings.json"

    if (-not (Test-Path -Path $masterSettingsPath)) {
        Write-Log "ERROR: Master settings file not found."
        return
    }

    # Parse the settings file
    try {
        $SettingsObject = Get-Content -Path $masterSettingsPath | ConvertFrom-Json
    } catch {
        Write-Log "ERROR: Failed to parse master settings JSON."
        return
    }

    # Set variables from settings
    $parentDirectory = $SettingsObject.ParentDirectory
    $tenant = $SettingsObject.tenant
    $tenantUrl = "https://$tenant.api.identitynow.com"
    $clientID = $SettingsObject.ClientID
    $clientSecret = $SettingsObject.ClientSecret
    $csvPath = Join-Path -Path $parentDirectory -ChildPath "AppList_$tenant.csv"

    # Ensure ExecutionLog folder exists
    $executionLogPath = Join-Path -Path $parentDirectory -ChildPath "ExecutionLog"
    if (-not (Test-Path -Path $executionLogPath)) {
        New-Item -Path $executionLogPath -ItemType Directory
        Write-Log "Created ExecutionLog folder."
    } else {
        Write-Log "ExecutionLog folder already exists. Skipping creation."
    }

    # Get OAuth token
    $authHeader = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("${clientID}:${clientSecret}"))

    try {
        $response = Invoke-RestMethod -Uri "$tenantUrl/oauth/token" -Method Post -Headers @{ Authorization = "Basic $authHeader" } -Body @{
            grant_type = "client_credentials"
        }
        $accessToken = $response.access_token
        Write-Log "OAuth token retrieved successfully."
    } catch {
        Write-Log "ERROR: Failed to retrieve OAuth token."
        return
    }

    # Set headers for API request
    $headers = @{ Authorization = "Bearer $accessToken" }

    #Fetch sources from the API
    try {
        $rawSources = Invoke-RestMethod -Uri "https://$tenant.api.identitynow.com/beta/sources" -Method Get -Headers $headers -ContentType "application/json;charset=utf-8"
        $rawJsonPath = "$parentDirectory\raw_api_response.json"
        $rawSources | Out-File -FilePath $rawJsonPath -Encoding UTF8
        Write-Log "Raw API response saved."

        $parsedSources = $rawSources | ConvertFrom-Json -AsHashtable

        $sourcesList = @()

        # Extract relevant fields
        foreach ($source in $parsedSources) {
            if ($source -is [hashtable]) {
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
            Write-Log "No valid sources found."
            return
        }

        # Check if AppList.csv already exists, load existing entries
        $existingSourcesList = @()
        if (Test-Path -Path $csvPath) {
            $existingSourcesList = Import-Csv -Path $csvPath
        }

        # Append only new sources
        $newSourcesList = @()
        foreach ($newSource in $sourcesList) {
            $isDuplicate = $existingSourcesList | Where-Object { $_.SourceID -eq $newSource.SourceID }
            if (-not $isDuplicate) {
                $newSourcesList += $newSource
            }
        }

        # Append new sources to the CSV if there are any
        if ($newSourcesList.Count -gt 0) {
            $combinedList = $existingSourcesList + $newSourcesList
            $combinedList | Export-Csv -Path $csvPath -NoTypeInformation
            Write-Log "AppList_$tenant.csv updated with new sources."
        } else {
            Write-Log "No new sources to add. No changes made to AppList_$tenant.csv."
        }

    } catch {
        Write-Log "ERROR: Failed to fetch or process sources. Details: $_"
        Write-Log "Error at line number: $($_.InvocationInfo.ScriptLineNumber)"
        Write-Log "Error message: $($_.Exception.Message)"
        return
    }

    # Import and filter sources for Delimited File type
    $folders = Import-Csv -Path $csvPath | Sort-Object SourceName
    $delimitedFileSources = $folders | Where-Object { $_.SourceType -eq "Delimited File" }

    # Create Import directory if it doesn't exist
    $importDirectory = Join-Path -Path $parentDirectory -ChildPath "Import"
    if (-not (Test-Path -Path $importDirectory)) {
        New-Item -Path $importDirectory -ItemType Directory
        Write-Log "Created Import folder."
    }

    # Create folder structure for Delimited File sources
    $counter = 1
    foreach ($folder in $delimitedFileSources) {
        $sourceName = $folder.SourceName
        $sourceID = $folder.SourceID
        $appFolderName = "$sourceName"
        $appFolderPath = Join-Path -Path $importDirectory -ChildPath $appFolderName

        # Check if the APP folder already exists, if so, skip folder creation
        if (-not (Test-Path -Path $appFolderPath)) {
            New-Item -Path $appFolderPath -ItemType Directory
            Write-Log "Created folder: $appFolderPath"

            # Create subfolders
            $logFolderPath = Join-Path -Path $appFolderPath -ChildPath "Log"
            $archiveFolderPath = Join-Path -Path $appFolderPath -ChildPath "Archive"
            if (-not (Test-Path -Path $logFolderPath)) {
                New-Item -Path $logFolderPath -ItemType Directory
                Write-Log "Created Log folder."
            }
            if (-not (Test-Path -Path $archiveFolderPath)) {
                New-Item -Path $archiveFolderPath -ItemType Directory
                Write-Log "Created Archive folder."
            }
        } else {
            Write-Log "Folder $appFolderPath already exists. Skipping folder creation."
        }

        <# Create the config.json file only if it doesn't already exist
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
}
"@
            Set-Content -Path $configFilePath -Value $configContent
            Write-Log "Config file created at: $configFilePath"
        } else {
            Write-Log "Config file already exists. Skipping creation."
        }#>


        # Check if isDebug is false, then delete raw_api_response.json
        if (-not $SettingsObject.isDebug) {
            if (Test-Path -Path $rawJsonPath) {
                try {
                    Remove-Item -Path $rawJsonPath -Force
                    Write-Log "Debugging turned off. Deleted raw_api_response.json"
                } catch {
                    Write-Log "ERROR: Failed to delete raw_api_response.json. Details: $_"
                }
            }
        } else {
            Write-Log "Debug mode is enabled (isDebug = true). raw_api_response.json not deleted."
        }

        $counter++
    }
}


# Function to create the default settings.json if it doesn't exist
function Create-DefaultSettings {
    $defaultSettings = @{
        ParentDirectory    = "C:\Powershell\FileUploadUtility"
        AppFolder          = "C:\Powershell\FileUploadUtility\Import"
        FileUploadUtility  = "C:\Powershell\sailpoint-file-upload-utility-4.1.0.jar"
        ClientSecret       = "Secret"
        ClientID           = "ClientID"
        tenant             = "tenant"
        isDebug            = $true  # Default to true for debugging
    }

    # Convert to JSON and save to settings.json
    $defaultSettings | ConvertTo-Json -Compress | Set-Content -Path $settingsPath -Force
    Write-Log "Default settings.json created."
}

# Function to append logs to the log panel with a timestamp
function Write-Log {
    param ($message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    if ($logTextBoxInitialized -and $logTextBox) {
        $logTextBox.Dispatcher.Invoke([action]{
            $logTextBox.AppendText("[$timestamp] $message`r`n")
            $logTextBox.ScrollToEnd()
        })
    } else {
        Write-Host "[$timestamp] $message"
    }
}

# Load existing settings or create a new settings.json if not found
$settingsPath = ".\settings.json"
if (-not (Test-Path -Path $settingsPath)) {
    Write-Host "Settings file not found. Creating a default settings file..."
    Create-DefaultSettings
}

try {
    $settings = Get-Content -Path $settingsPath | ConvertFrom-Json
} catch {
    Write-Host "ERROR: Failed to parse settings JSON."
    exit 1
}

# Create the Window
$window = New-Object Windows.Window
$window.Title = "SailPoint FileUploadUtility Directory Creation"
$window.Width = 900
$window.Height = 600
$window.WindowStartupLocation = 'CenterScreen'

# Create a Grid Layout
$grid = New-Object Windows.Controls.Grid
$grid.Margin = '10'
$window.Content = $grid

# Define Columns for left and right panels
$leftColumn = New-Object Windows.Controls.ColumnDefinition
$leftColumn.Width = "2*"  # Left panel takes up 2/5 of the width
$rightColumn = New-Object Windows.Controls.ColumnDefinition
$rightColumn.Width = "3*"  # Right panel takes up 3/5 of the width
$grid.ColumnDefinitions.Add($leftColumn)
$grid.ColumnDefinitions.Add($rightColumn)

# Create a StackPanel for the left panel with GuidePoint Security theme colors
$leftPanel = New-Object Windows.Controls.StackPanel
$leftPanel.Margin = '0'
$leftPanel.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#1C1D30")  # Navy Blue background
$leftPanel.VerticalAlignment = 'Stretch'
$leftPanel.HorizontalAlignment = 'Stretch'
[Windows.Controls.Grid]::SetColumn($leftPanel, 0)
$grid.Children.Add($leftPanel)

# Directory Locations GroupBox with white text and blue accents
$directoryBox = New-Object Windows.Controls.GroupBox
$directoryBox.Header = "Directory Locations"
$directoryBox.Foreground = [System.Windows.Media.Brushes]::White  # White text
$directoryBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")  # Bright Blue border
$directoryBox.Margin = '10'
$directoryPanel = New-Object Windows.Controls.StackPanel
$directoryPanel.Orientation = 'Vertical'
$directoryBox.Content = $directoryPanel
$leftPanel.Children.Add($directoryBox)

# Add fields for directories in Directory Locations group
$fields = @(
    @{ Label = "Parent Directory"; FieldName = "ParentDirectory"; FieldValue = $settings.ParentDirectory },
    @{ Label = "App Folder"; FieldName = "AppFolder"; FieldValue = $settings.AppFolder },
    @{ Label = "File Upload Utility"; FieldName = "FileUploadUtility"; FieldValue = $settings.FileUploadUtility }
)

foreach ($field in $fields) {
    $label = New-Object Windows.Controls.Label
    $label.Content = $field.Label
    $label.Foreground = [System.Windows.Media.Brushes]::White  # White text for labels
    $directoryPanel.Children.Add($label)

    $textBox = New-Object Windows.Controls.TextBox
    $textBox.Width = 300
    $textBox.Text = $field.FieldValue
    $textBox.Background = [System.Windows.Media.Brushes]::LightGray  # Light gray text box
    $textBox.Foreground = [System.Windows.Media.Brushes]::Black  # Black text inside text boxes
    $directoryPanel.Children.Add($textBox)

    Set-Variable -Name ($field.FieldName + "TextBox") -Value $textBox
}

# Client Credentials GroupBox with white text and blue accents
$credentialsBox = New-Object Windows.Controls.GroupBox
$credentialsBox.Header = "Client Credentials"
$credentialsBox.Foreground = [System.Windows.Media.Brushes]::White  # White text
$credentialsBox.BorderBrush = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")  # Bright Blue border
$credentialsBox.Margin = '10'
$credentialsPanel = New-Object Windows.Controls.StackPanel
$credentialsPanel.Orientation = 'Vertical'
$credentialsBox.Content = $credentialsPanel
$leftPanel.Children.Add($credentialsBox)

# Add fields for Client Credentials in Client Credentials group
$fields = @(
    @{ Label = "Tenant"; FieldName = "tenant"; FieldValue = $settings.tenant },
    @{ Label = "Client ID"; FieldName = "ClientID"; FieldValue = $settings.ClientID }
)

foreach ($field in $fields) {
    $label = New-Object Windows.Controls.Label
    $label.Content = $field.Label
    $label.Foreground = [System.Windows.Media.Brushes]::White  # White text for labels
    $credentialsPanel.Children.Add($label)

    $textBox = New-Object Windows.Controls.TextBox
    $textBox.Width = 300
    $textBox.Text = $field.FieldValue
    $textBox.Background = [System.Windows.Media.Brushes]::LightGray  # Light gray text box
    $textBox.Foreground = [System.Windows.Media.Brushes]::Black  # Black text inside text boxes
    $credentialsPanel.Children.Add($textBox)

    Set-Variable -Name ($field.FieldName + "TextBox") -Value $textBox
}

# Add Client Secret field with password handling
$clientSecretLabel = New-Object Windows.Controls.Label
$clientSecretLabel.Content = "Client Secret"
$clientSecretLabel.Foreground = [System.Windows.Media.Brushes]::White  # White text for label
$credentialsPanel.Children.Add($clientSecretLabel)

$clientSecretTextBox = New-Object Windows.Controls.PasswordBox
$clientSecretTextBox.Width = 300
$clientSecretTextBox.Password = $settings.ClientSecret
$clientSecretTextBox.Background = [System.Windows.Media.Brushes]::LightGray  # Light gray text box
$clientSecretTextBox.Foreground = [System.Windows.Media.Brushes]::Black  # Black text inside password box
$credentialsPanel.Children.Add($clientSecretTextBox)

# Centered and Stretched Run Script Button
$submitButton = New-Object Windows.Controls.Button
$submitButton.Content = "Run Script"
$submitButton.Width = 300
$submitButton.Padding = '10'
$submitButton.HorizontalAlignment = 'Center'
$submitButton.Background = New-Object Windows.Media.SolidColorBrush (New-Object Windows.Media.ColorConverter).ConvertFromString("#007BFF")  # Bright Blue background for the button
$submitButton.Foreground = [System.Windows.Media.Brushes]::White  # White text for the button
$submitButton.FontSize = 16
$submitButton.FontWeight = 'Bold'
$submitButton.Margin = '10'
$leftPanel.Children.Add($submitButton)

# Create TextBox for Logs in the right panel
$logTextBox = New-Object Windows.Controls.TextBox
$logTextBox.AcceptsReturn = $true
$logTextBox.VerticalScrollBarVisibility = 'Auto'
$logTextBox.HorizontalScrollBarVisibility = 'Auto'
$logTextBox.IsReadOnly = $true
$logTextBox.Margin = '0'
[Windows.Controls.Grid]::SetColumn($logTextBox, 1)
$grid.Children.Add($logTextBox)

# Mark the $logTextBox as initialized
$logTextBoxInitialized = $true

# Button Click Event
$submitButton.Add_Click({
    $tenant = $tenantTextBox.Text
    $clientID = $ClientIDTextBox.Text
    $clientSecret = $clientSecretTextBox.Password

    # Validate input fields
    if ([string]::IsNullOrEmpty($tenant) -or [string]::IsNullOrEmpty($clientID) -or [string]::IsNullOrEmpty($clientSecret)) {
        Write-Log "Error: All fields are required."
        return
    }

    # Update settings
    $settings.ParentDirectory = $ParentDirectoryTextBox.Text
    $settings.AppFolder = $AppFolderTextBox.Text
    $settings.FileUploadUtility = $FileUploadUtilityTextBox.Text
    $settings.tenant = $tenant
    $settings.ClientID = $clientID
    $settings.ClientSecret = $clientSecret

    # Save settings
    try {
        $settings | ConvertTo-Json -Compress | Set-Content -Path $settingsPath -Force
        Write-Log "Settings saved successfully."
    } catch {
        Write-Log "Error: Failed to save settings. $_"
    }

    # Run the script logic
    Run-Script
})

# Show Window
$window.ShowDialog()
