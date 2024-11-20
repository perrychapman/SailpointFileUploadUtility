# Define the source directory for files and the target directory for folders
$sourceDirectory = "C:\Path\To\SourceFiles"
$folderDirectory = "C:\Path\To\TargetFolders"

# Get all the folders in the target directory with the pattern *APP-#
$folders = Get-ChildItem -Path $folderDirectory -Directory | Where-Object { $_.Name -match "APP-\d+" }

# Get all the files in the source directory with the pattern *APP-#
$files = Get-ChildItem -Path $sourceDirectory -File | Where-Object { $_.Name -match "APP-\d+" }

# Loop through each file
foreach ($file in $files) {
    # Extract the APP-# pattern (with any number of digits) from the file name
    $pattern = ($file.Name -match "APP-\d+") ? $matches[0] : $null

    if ($pattern) {
        # Find the corresponding folder with the same pattern
        $targetFolder = $folders | Where-Object { $_.Name -eq $pattern }

        if ($targetFolder) {
            # Define the target path
            $targetPath = Join-Path -Path $targetFolder.FullName -ChildPath $file.Name

            # Check if the file already exists in the target folder
            if (-not (Test-Path -Path $targetPath)) {
                # Copy the file to the target folder
                Copy-Item -Path $file.FullName -Destination $targetPath
                Write-Host "Copied $($file.Name) to $($targetFolder.FullName)"
            } else {
                Write-Host "File $($file.Name) already exists in $($targetFolder.FullName)"
            }
        } else {
            Write-Host "No matching folder found for $($file.Name)"
        }
    }
}
