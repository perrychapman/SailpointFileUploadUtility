# =====================================================================
# PowerShell Script: FileUploadScript.ps1
# Description: Processes CSV and Excel files using config settings by app,
#              prepares data for upload, and then uploads processed file
#              to Sailpoint
# Requirements: Powershell 7+
#               ImportExcel module installed
#               Open JDK 11+ installed on host machine
# =====================================================================

# =====================================================================
# Changelog
#
# [11/18/2024]
# - Added adminColumnName and adminColumnValue for admin role assignment.
# - Updated Process-ImportedData to set Role as "Admin" or "User" based on column value.
# - Added support for txt file handling.
# 
# [11/21/2024]
# - Added booleanValue processing for files.
#
# [12/17/2024]
# - Added fixes for AppFilter logic
# - Added improvements to Ensure-ImportExcelModule for no internet connectivity scenarios.
# =====================================================================


# ------------------------
# 1. Ensure ImportExcel Module
# ------------------------

function Ensure-ImportExcelModule {
    param (
        [string]$ModuleName = 'ImportExcel'
    )

    $importedModule = Get-Module -Name $ModuleName
    if ($importedModule) {
        Write-Host "Module '$ModuleName' is already imported."
        return
    }

    $availableModule = Get-Module -ListAvailable -Name $ModuleName
    if ($availableModule) {
        Write-Host "Module '$ModuleName' is already installed locally."
        try {
            Import-Module $ModuleName -ErrorAction Stop
            Write-Host "Module '$ModuleName' imported successfully."
            return
        }
        catch {
            Write-Error "Failed to import module '$ModuleName'. Error: $_"
            exit 1
        }
    }

    # Check for PowerShell repository availability
    $repository = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
    if (-not $repository -or $repository.SourceLocation -eq $null) {
        Write-Warning "PowerShell Gallery repository is not available. Module '$ModuleName' cannot be installed. Ensure the module is pre-installed on this server."
        return
    }

    # Attempt to install the module
    Write-Host "Module '$ModuleName' is not installed. Attempting to install..."
    try {
        Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "Module '$ModuleName' installed successfully."
        Import-Module $ModuleName -ErrorAction Stop
        Write-Host "Module '$ModuleName' imported successfully."
    }
    catch {
        Write-Warning "Failed to install module '$ModuleName' due to error: $_"
        Write-Warning "Please ensure the module is installed manually or PowerShell Gallery access is available."
    }
}

# ------------------------
# 2. Load Master Settings
# ------------------------

function Load-MasterSettings {
    param (
        [string]$masterSettingsPath
    )

    if (-not (Test-Path -Path $masterSettingsPath)) {
        Write-Error "Master settings file not found at path: $masterSettingsPath"
        exit 1
    }

    try {
        $SettingsObject = Get-Content -Path $masterSettingsPath | ConvertFrom-Json
    }
    catch {
        Write-Error "Failed to parse master settings JSON. Error: $_"
        exit 1
    }

    return $SettingsObject
}

# ------------------------
# 3. Define Logging Function
# ------------------------

function Write-Log {
    param(
        [string]$logDetails,       # Log message
        [string]$logFilePath,      # Path to the log file
        [string]$logType = 'INFO'  # Log type (INFO/ERROR/WARNING)
    )

    $currentDateTime = Get-Date -Format "HH:mm:ss"
    $logEntry = [PSCustomObject]@{
        'Date/Time'   = $currentDateTime
        'Log Type'    = $logType
        'Log Details' = $logDetails
    }

    if (-not (Test-Path -Path $logFilePath)) {
        $logEntry | Export-Csv -Path $logFilePath -NoTypeInformation
    }
    else {
        $logEntry | Export-Csv -Path $logFilePath -NoTypeInformation -Append
    }
}

# ------------------------
# 4. Define Trimming, Dropping, and Merging Function
# ------------------------

function Trim-Data {
    param(
        [array]$data,                 # Array of PSCustomObject
        [int]$trimTopRows = 0,        # Number of top rows to trim after header
        [int]$trimBottomRows = 0,     # Number of bottom rows to trim
        [int]$trimLeftColumns = 0,    # Number of leftmost columns to trim
        [int]$trimRightColumns = 0,   # Number of rightmost columns to trim
        [array]$dropColumns,          # Columns to drop
        [array]$columnsToMerge,       # Columns to merge
        [string]$mergedColumnName,    # Name of the new merged column
        [string]$AppLogFilePath       # Log file path for logging
    )

    # Trim top rows
    if ($trimTopRows -gt 0) {
        $data = $data | Select-Object -Skip $trimTopRows
    }

    # Trim bottom rows
    if ($trimBottomRows -gt 0) {
        if ($trimBottomRows -lt $data.Count) {
            $data = $data | Select-Object -First ($data.Count - $trimBottomRows)
        }
        else {
            Write-Log -logDetails "Trim bottom rows ($trimBottomRows) exceeds data count ($($data.Count)). Skipping bottom trim." -logFilePath $AppLogFilePath -logType 'WARNING'
        }
    }

    # Trim right columns
    if ($trimRightColumns -gt 0 -and $data.Count -gt 0) {
        $allColumns = $data[0].PSObject.Properties.Name
        if ($allColumns.Count -gt $trimRightColumns) {
            $columnsToKeep = $allColumns[0..($allColumns.Count - $trimRightColumns - 1)]
            $data = $data | Select-Object -Property $columnsToKeep
        }
        else {
            Write-Log -logDetails "Trim right columns ($trimRightColumns) exceeds available columns ($($allColumns.Count)). Skipping right trim." -logFilePath $AppLogFilePath -logType 'WARNING'
        }
    }

    # Trim left columns
    if ($trimLeftColumns -gt 0 -and $data.Count -gt 0) {
        $allColumns = $data[0].PSObject.Properties.Name
        if ($allColumns.Count -gt $trimLeftColumns) {
            $columnsToKeep = $allColumns[$trimLeftColumns..($allColumns.Count - 1)]
            $data = $data | Select-Object -Property $columnsToKeep
        }
        else {
            Write-Log -logDetails "Trim left columns ($trimLeftColumns) exceeds available columns ($($allColumns.Count)). Skipping left trim." -logFilePath $AppLogFilePath -logType 'WARNING'
        }
    }

    # Merge specified columns into a new column
    if ($columnsToMerge.Count -gt 0 -and $mergedColumnName) {
        try {
            foreach ($row in $data) {
                $mergedValue = ($columnsToMerge | ForEach-Object { 
                    $value = $row."$_"
                    if (![string]::IsNullOrEmpty($value)) { $value }
                }) -join ' - '
                $row | Add-Member -MemberType NoteProperty -Name $mergedColumnName -Value $mergedValue -Force
            }
            Write-Log -logDetails "Merged columns ($($columnsToMerge -join ', ')) into $mergedColumnName." -logFilePath $AppLogFilePath -logType 'INFO'
        }
        catch {
            Write-Log -logDetails "Failed to merge columns: $($columnsToMerge -join ', '). Error: $_" -logFilePath $AppLogFilePath -logType 'ERROR'
        }
    }
    else {
        Write-Log -logDetails "No columns specified for merging or missing mergedColumnName. Skipping merge." -logFilePath $AppLogFilePath -logType 'WARNING'
    }
    
    if($columnsToMerge){
        $dropColumns += $columnsToMerge
    }
    # Drop specified columns (case-insensitive)
    try {
        if ($dropColumns.Count -gt 0 -and $data.Count -gt 0) {
            $columnsInData = $data[0].PSObject.Properties.Name
            $columnsToExclude = $columnsInData | Where-Object { $dropColumns -contains $_.ToLower() }

            if ($columnsToExclude.Count -gt 0) {
                $data = $data | Select-Object -Property * -ExcludeProperty $columnsToExclude
                Write-Log -logDetails "Dropped columns: $($columnsToExclude -join ', ')" -logFilePath $AppLogFilePath -logType 'INFO'
            }
            else {
                Write-Log -logDetails "No matching columns found to drop: $($dropColumns -join ', ')" -logFilePath $AppLogFilePath -logType 'WARNING'
            }
        }
    }
    catch {
        Write-Log -logDetails "Failed to drop specified columns. Error: $_" -logFilePath $AppLogFilePath -logType 'ERROR'
    }

    return $data
}

# ------------------------
# 5. Process Imported Data
# ------------------------

function Process-ImportedData {
    param (
        [array]$users,
        [array]$schema,
        [array]$groupTypes,
        [string]$disableField,
        [array]$disableValues,
        [string]$groupDelimiter,
        [string]$AppLogFilePath,
        [string]$adminColumnName,
        [string]$adminColumnValue
    )

    if ($groupTypes -ne $null -and $groupTypes.Count -gt 0) {
        Write-Log -logDetails "Group Types: $groupTypes" -logFilePath $AppLogFilePath
    }
    else {
        Write-Log -logDetails "No group types defined. Adding default Role column." -logFilePath $AppLogFilePath
    }

    # Ensure schema is initialized to an empty array if not provided
    if (-not $schema) { $schema = @() }

    $table = @()

    foreach ($user in $users) {
        $temp_user = [PSCustomObject]@{}

        if ($schema.Count -eq 0) {
            foreach ($object_property in $user.PSObject.Properties) {
                $schema += $object_property.Name
            }
        }

        foreach ($attribute in $schema) {
            $userAttribute = $user."$attribute"
            if ($groupTypes -contains $attribute) {
                $temp_user | Add-Member -Name $attribute -MemberType NoteProperty -Value $null
            }
            elseif ($schema -contains $attribute) {
                $temp_user | Add-Member -Name $attribute -MemberType NoteProperty -Value $userAttribute
            } 
        }

        $IIQDisabled = if ($disableField -ne "" -and $disableValues -contains $user."$disableField") { "true" } else { "false" }
        $temp_user | Add-Member -Name "IIQDisabled" -MemberType NoteProperty -Value $IIQDisabled -Force

        if ($groupTypes -ne $null -and $groupTypes.Count -gt 0) {
            foreach ($type in $groupTypes) {
                $usergrptype = $user."$type"
                if ($usergrptype) {
                    $splitGroups = if ([string]::IsNullOrEmpty($groupDelimiter)) {
                        @($usergrptype)
                    } else {
                        $usergrptype.Split($groupDelimiter)
                    }

                    foreach ($grp in $splitGroups) {
                        $temp_user.$type = $grp.Trim()
                        $table += $temp_user.PSObject.Copy()
                    }
                    $temp_user.$type = $null
                }
            }
        }
        else {
            # Check if adminColumnName and adminColumnValue match user data, and set Role accordingly
            if ($adminColumnName -and $adminColumnValue -and $user.PSObject.Properties[$adminColumnName] -and $user.PSObject.Properties[$adminColumnName].Value -eq $adminColumnValue) {
                $temp_user | Add-Member -Name "Role" -MemberType NoteProperty -Value "Admin" -Force
            }
            else {
                $temp_user | Add-Member -Name "Role" -MemberType NoteProperty -Value "User" -Force
            }
            $table += $temp_user.PSObject.Copy()
        }
         
    }

    return $table
}

# ------------------------
# 6. Upload to SailPoint
# ------------------------

function Upload-ToSailPoint {
    param (
        [string]$newFile,
        [string]$sourceID,
        [string]$clientURL,
        [string]$fileUploadUtility,
        [string]$ClientID,
        [string]$ClientSecret,
        [string]$AppLogFilePath
    )

    if(Test-Path -Path $newFile) {
        try {
            $output = & java -jar $fileUploadUtility --url $clientURL --clientId $ClientID --clientSecret $ClientSecret --file $newFile -v | out-string

            if ($output -and $output.Contains("error")) {
                Write-Log -logDetails "Error during upload for SourceID: $sourceID. Output: $output" -logFilePath $AppLogFilePath -logType 'ERROR'
            } else {
                Write-Log -logDetails "Upload completed successfully for SourceID: $sourceID." -logFilePath $AppLogFilePath -logType 'INFO'
                $script:uploadCount++
            }
        }
        catch {
            Write-Log -logDetails "Failed to upload file to SailPoint. Error: $_" -logFilePath $AppLogFilePath -logType 'ERROR'
        }
    } else {
        Write-Log -logDetails "File path for upload file not found. Upload failed." -logFilePath $AppLogFilePath -logType 'ERROR'
        $script:errorCount++
    }
}

# ------------------------
# 7. File Archival
# ------------------------

function Archive-File {
    param (
        [string]$file,
        [string]$archivePath,
        [string]$AppLogFilePath
    )

    try {
        Move-Item -Path $file -Destination $archivePath -Force
        Write-Log -logDetails "File moved to Archive: $file" -logFilePath $AppLogFilePath -logType 'INFO'
    }
    catch {
        Write-Log -logDetails "Failed to move file to Archive. Error: $_" -logFilePath $AppLogFilePath -logType 'ERROR'
    }
}

# ------------------------
# 8. Main Script Execution
# ------------------------

function Process-FilesInAppFolder {
    param (
        [string]$AppFolderPath,
        [PSObject]$AppConfig,
        [string]$AppLogFilePath,
        [PSObject]$SettingsObject
    )

    $tenant = $SettingsObject.tenant
    $clientURL = "https://$tenant.api.identitynow.com"
    $sourceID = $AppConfig.sourceID
    $isMonarch = $AppConfig.isMonarch
    $disableField = $AppConfig.disableField
    $disableValues = if ($AppConfig.disableValue -is [System.Collections.IEnumerable]) { $AppConfig.disableValue } else { @($AppConfig.disableValue) }
    $groupTypes = $AppConfig.groupTypes.Split(',') | ForEach-Object { $_.Trim() }
    $groupDelimiter = $AppConfig.groupDelimiter
    $isUpload = $AppConfig.isUpload
    $headerRow = $AppConfig.headerRow
    $trimTopRows = $AppConfig.trimTopRows
    $trimBottomRows = $AppConfig.trimBottomRows
    $trimLeftColumns = $AppConfig.trimLeftColumns
    $trimRightColumns = $AppConfig.trimRightColumns
    $schema = $AppConfig.schema.Split(',') | ForEach-Object { $_.Trim() }
    $dropColumns = $AppConfig.dropColumns.Split(',') | ForEach-Object { $_.Trim().ToLower() }
    $columnsToMerge = if ($AppConfig.columnsToMerge -ne $null) { $AppConfig.columnsToMerge.Split(',') | ForEach-Object { $_.Trim() } } else { @() }
    $mergedColumnName = $AppConfig.mergedColumnName
    $adminColumnName = $AppConfig.adminColumnName
    $adminColumnValue = $AppConfig.adminColumnValue
    $booleanColumns = if ($AppConfig.booleanColumnList -ne $null) {$AppConfig.booleanColumnList.Split(",").Trim() } else { @() }
    $booleanValue = $AppConfig.booleanColumnValue

    $checkPath = if ($isMonarch) { Join-Path -Path $AppFolderPath -ChildPath "MonarchProcessed" } else { $AppFolderPath }

    Write-Log -logDetails "Checking for CSV and Excel files in $checkPath..." -logFilePath $AppLogFilePath

    $files = Get-ChildItem -Path $checkPath -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '(?i)\.(csv|xlsx|xls|txt)$' }

    if (-not $files) {
        Write-Log -logDetails "No CSV or Excel files found in $checkPath. Import cancelled." -logFilePath $AppLogFilePath -logType 'ERROR'
        return
    }
    else {
        Write-Log -logDetails "Found $($files.Count) file(s) in $checkPath." -logFilePath $AppLogFilePath -logType 'INFO'
    }

    # Only process most recently modified file in App Folder
    if($files.Count -gt 1) {
        $recentFile = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        Write-Log -logDetails "Multiple files found. Processing most recently modified file: $($recentFile.Name)" -logFilePath $AppLogFilePath -logType 'INFO'
    } else {
        $recentFile = $files[0]
        Write-Log -logDetails "Processing file: $($recentFile.Name)" -logFilePath $AppLogFilePath -logType 'INFO'
    }

    $file = $recentFile
    $fileExtension = $file.Extension.ToLower()
    $current_date = Get-Date -Format yyyy_MM_dd_HH.mm
    $archivePath = Join-Path -Path $AppFolderPath -ChildPath "Archive"
    $newFile = Join-Path -Path $archivePath -ChildPath ("$sourceID" + "_upload_file_$current_date.csv")
    $originalFile = Join-Path -Path $archivePath -ChildPath "Original_$current_date$($file.Extension)"
    $processedFile = Join-Path -Path $archivePath -ChildPath "Processed_$current_date.csv"

    # Ensure Archive directory exists
    try {
        New-Item -ItemType Directory -Force -Path $archivePath | Out-Null
    }
    catch {
        Write-Log -logDetails "Failed to create Archive directory at $archivePath. Error: $_" -logFilePath $AppLogFilePath -logType 'ERROR'
        continue
    }

    # Archive the original file
    try {
        Copy-Item -Path $file.FullName -Destination $originalFile -Force
        Write-Log -logDetails "Original file archived to $originalFile." -logFilePath $AppLogFilePath -logType 'INFO'
    }
    catch {
        Write-Log -logDetails "Failed to archive original file. Error: $_" -logFilePath $AppLogFilePath -logType 'ERROR'
        continue
    }

    # Import and process the data
    $users = if (($fileExtension -eq ".csv") -or ($fileExtension -eq ".txt")) { Import-Csv -Path $file.FullName } else { Import-Excel -Path $file.FullName -StartRow $headerRow }

    if(!$schema){
        $users = Trim-Data -data $users -trimTopRows $trimTopRows -trimBottomRows $trimBottomRows -trimLeftColumns $trimLeftColumns -trimRightColumns $trimRightColumns -dropColumns $dropColumns -columnsToMerge $columnsToMerge -mergedColumnName $mergedColumnName -AppLogFilePath $AppLogFilePath
    } else {
        $users = Trim-Data -data $users -trimTopRows $trimTopRows -trimBottomRows $trimBottomRows -columnsToMerge $columnsToMerge -mergedColumnName $mergedColumnName -AppLogFilePath $AppLogFilePath
    }

    if ($users.Count -eq 0) {
        Write-Log -logDetails "No data found in $($file.Name) after trimming. Skipping processing." -logFilePath $AppLogFilePath -logType 'WARNING'
        continue
    }

    # Boolean value processing
    # Modify processed file to replace entitlement columns with Role column
    if ($booleanColumns.Count -gt 0 -and $booleanValue) {
        $groupTypes = "Role"
        $groupDelimiter = ","

        try {
            $users = $users | ForEach-Object {
                # Extract matched entitlements
                $entitlements = @()
                foreach ($col in $booleanColumns) {
                    if ($_.PSObject.Properties[$col] -and $_."$col" -eq $booleanValue) {
                        $entitlements += $col
                    }
                }

                # Add Role column with comma-separated entitlements
                $_ | Add-Member -Name "Role" -MemberType NoteProperty -Value ($entitlements -join ", ") -Force

                # Remove the original entitlement columns
                foreach ($col in $booleanColumns) {
                    $_.PSObject.Properties.Remove($col)
                }

                $_
            }

            # Log successful transformation
            Write-Log -logDetails "Replaced entitlement columns with Role column in processed file." -logFilePath $AppLogFilePath -logType 'INFO'
        }
        catch {
            # Log any errors
            Write-Log -logDetails "Error processing entitlement columns: $_" -logFilePath $AppLogFilePath -logType 'ERROR'
        }
    }


    # Export the processed data (before grouping)
    try {
        $users | Export-Csv -Path $processedFile -NoTypeInformation
        Write-Log -logDetails "Processed data exported to $processedFile." -logFilePath $AppLogFilePath -logType 'INFO'
    }
    catch {
        Write-Log -logDetails "Failed to export processed data to CSV. Error: $_" -logFilePath $AppLogFilePath -logType 'ERROR'
        continue
    }

    # Further process the data (group entitlements)
    $table = Process-ImportedData -users $users -groupTypes $groupTypes -disableField $disableField -disableValues $disableValues -groupDelimiter $groupDelimiter -AppLogFilePath $AppLogFilePath -schema $schema -adminColumnName $adminColumnName -adminColumnValue $adminColumnValue

    try {
        $table | Export-Csv -Path $newFile -NoTypeInformation

        # Check if the upload file was created successfully
        if (Test-Path -Path $newFile) {
            Write-Log -logDetails "Processed data for upload exported to $newFile." -logFilePath $AppLogFilePath -logType 'INFO'
        }
        else {
            Write-Log -logDetails "Failed to create upload file at $newFile." -logFilePath $AppLogFilePath -logType 'ERROR'
            continue
        }
    }
    catch {
        Write-Log -logDetails "Failed to export data for upload. Error: $_" -logFilePath $AppLogFilePath -logType 'ERROR'
        continue
    }

    # Upload to SailPoint
    if ($isUpload -and $null -ne $sourceID -and $sourceID -ne "") {
        Upload-ToSailPoint -newFile $newFile -sourceID $sourceID -clientURL $clientURL -fileUploadUtility $SettingsObject.FileUploadUtility -ClientID $SettingsObject.ClientID -ClientSecret $SettingsObject.ClientSecret -AppLogFilePath $AppLogFilePath
    }

    # Archive processed file
    Archive-File -file $processedFile -archivePath $archivePath -AppLogFilePath $AppLogFilePath
}

# Ensure ImportExcel module is available
Ensure-ImportExcelModule
# Path to the master settings JSON file
$masterSettingsPath = ".\settings.json"
$SettingsObject = Load-MasterSettings -masterSettingsPath $masterSettingsPath

# Validate essential settings
$requiredSettings = @('AppFolder', 'ParentDirectory', 'tenant', 'FileUploadUtility', 'ClientID', 'ClientSecret')
foreach ($setting in $requiredSettings) {
    if (-not $SettingsObject.PSObject.Properties.Match($setting)) {
        Write-Error "Missing required setting: $setting"
        exit 1
    }
}

# Define global variables from master settings
$parentDirectory = $SettingsObject.ParentDirectory
$targetDirectory = $SettingsObject.AppFolder
$logFileDate = Get-Date -Format "yyyyMMdd"
$executionLogFileName = "ExecutionLog_$logFileDate.csv"
$executionLogFilePath = "./ExecutionLog/$executionLogFileName"
$tenant = $SettingsObject.tenant
$fileUploadUtility = $SettingsObject.FileUploadUtility
$clientURL = "https://$tenant.api.identitynow.com"
$ClientID = $SettingsObject.ClientID
$ClientSecret = $SettingsObject.ClientSecret

# Main Execution
$startTime = Get-Date
$totalAppCount = 0
$processedCount = 0
$skippedCount = 0
$errorCount = 0
$uploadCount = 0
$AppFilter = $SettingsObject.AppFilter

if ([string]::IsNullOrWhiteSpace($AppFilter)) {
    # If AppFilter is null or empty, process all folders
    $AppFilter = ".*"
    Write-Log -logDetails "Script started at $($startTime.ToString("MM/dd/yyyy HH:mm:ss"))" -logFilePath $executionLogFilePath -logType 'INFO'
} else {
    # Escape special regex characters and convert to regex for substring match
    $escapedFilter = [regex]::Escape($AppFilter)
    $AppFilter = ".*$escapedFilter.*"
    Write-Log -logDetails "Script started at $($startTime.ToString("MM/dd/yyyy HH:mm:ss")) with filter on '$($SettingsObject.AppFilter)'." -logFilePath $executionLogFilePath -logType 'INFO'
}

# Fetch and filter app folders using wildcard filter
$AppFolders = Get-ChildItem -Path $targetDirectory -Directory | Where-Object { $_.Name -match $AppFilter }
$totalAppCount = $AppFolders.Count

if ($totalAppCount -eq 0) {
    Write-Log -logDetails "No folders matching filter '$($AppFilter)' found in $targetDirectory. Script exiting." -logFilePath $executionLogFilePath -logType 'WARNING'
    exit
}

foreach ($AppFolder in $AppFolders) {
    $startAppTime = Get-Date
    $AppFolderPath = $AppFolder.FullName
    $AppFolderName = $AppFolder.Name
    $configFilePath = Join-Path -Path $AppFolderPath -ChildPath "config.json"
    $appLogFileName = "Log_$AppFolderName" + "_$logFileDate.csv"
    $AppLogFilePath = Join-Path -Path "$AppFolderPath\Log" -ChildPath $appLogFileName

    Write-Log -logDetails "Processing started for $AppFolderName..." -logFilePath $executionLogFilePath -logType 'INFO'

    if (-not (Test-Path -Path $configFilePath)) {
        Write-Log -logDetails "Config file not found for $AppFolderName. Import skipped." -logFilePath $AppLogFilePath -logType 'ERROR'
        Write-Log -logDetails "Config file not found for $AppFolderName. Import skipped." -logFilePath $executionLogFilePath -logType 'WARNING'
        $skippedCount++
        continue
    }

    $AppConfig = Get-Content -Path $configFilePath | ConvertFrom-Json

    try {
        # Process files in the folder
        Process-FilesInAppFolder -AppFolderPath $AppFolderPath -AppConfig $AppConfig -AppLogFilePath $AppLogFilePath -SettingsObject $SettingsObject
        
        $endAppTime = Get-Date
        $appDuration = New-TimeSpan -Start $startAppTime -End $endAppTime
        Write-Log -logDetails "Processing completed for $AppFolderName. Duration: $($appDuration.ToString())" -logFilePath $executionLogFilePath -logType 'INFO'
        Write-Log -logDetails "Processing completed for $AppFolderName. Duration: $($appDuration.ToString())" -logFilePath $AppLogFilePath -logType 'INFO'
        $processedCount++
    }
    catch {
        Write-Log -logDetails "Error during processing for $AppFolderName. Error: $_" -logFilePath $executionLogFilePath -logType 'ERROR'
        $errorCount++
    }
}

$endTime = Get-Date
$scriptDuration = New-TimeSpan -Start $startTime -End $endTime
Write-Log -logDetails "Script completed at $($endTime.ToString("MM/dd/yyyy HH:mm:ss")). Duration: $($scriptDuration.ToString()). Total Apps: $totalAppCount. Processed: $processedCount. Skipped: $skippedCount. Errors: $errorCount. Successful Uploads: $uploadCount." -logFilePath $executionLogFilePath -logType 'INFO'
