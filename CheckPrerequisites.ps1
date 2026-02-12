# =====================================================================
# Environment Validation Script
# Checks prerequisites before running SailPoint File Upload Utility
# =====================================================================

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  SailPoint Utility - Prerequisites Check" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$allChecksPassed = $true

# Check PowerShell Version
Write-Host "Checking PowerShell version..." -NoNewline
$psVersion = $PSVersionTable.PSVersion
if ($psVersion.Major -ge 7) {
    Write-Host " OK (v$($psVersion.Major).$($psVersion.Minor))" -ForegroundColor Green
} else {
    Write-Host " FAILED (v$($psVersion.Major).$($psVersion.Minor))" -ForegroundColor Red
    Write-Host "  Required: PowerShell 7 or higher" -ForegroundColor Yellow
    Write-Host "  Download from: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Yellow
    $allChecksPassed = $false
}

# Check Java Installation
Write-Host "Checking Java installation..." -NoNewline
try {
    $javaVersion = & java -version 2>&1 | Select-Object -First 1
    if ($javaVersion -match 'version "(\d+)') {
        $javaVer = [int]$matches[1]
        if ($javaVer -ge 11) {
            Write-Host " OK (Java $javaVer)" -ForegroundColor Green
        } else {
            Write-Host " WARNING (Java $javaVer - 11+ recommended)" -ForegroundColor Yellow
        }
    } else {
        Write-Host " OK" -ForegroundColor Green
    }
} catch {
    Write-Host " FAILED" -ForegroundColor Red
    Write-Host "  Java is not installed or not in PATH" -ForegroundColor Yellow
    Write-Host "  Download JDK 11+ from: https://adoptium.net/" -ForegroundColor Yellow
    $allChecksPassed = $false
}

# Check ImportExcel Module
Write-Host "Checking ImportExcel module..." -NoNewline
$importExcel = Get-Module -ListAvailable -Name ImportExcel
if ($importExcel) {
    Write-Host " OK (v$($importExcel.Version))" -ForegroundColor Green
} else {
    Write-Host " NOT INSTALLED" -ForegroundColor Yellow
    Write-Host "  Will be installed automatically on first run" -ForegroundColor Cyan
    Write-Host "  Or install manually: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Cyan
}

# Check for settings.json
Write-Host "Checking settings.json..." -NoNewline
if (Test-Path ".\settings.json") {
    Write-Host " FOUND" -ForegroundColor Green
    try {
        $settings = Get-Content ".\settings.json" | ConvertFrom-Json
        
        # Validate required fields
        $requiredFields = @('tenant', 'ClientID', 'ClientSecret', 'FileUploadUtility', 'AppFolder')
        $missingFields = @()
        foreach ($field in $requiredFields) {
            if (-not $settings.PSObject.Properties[$field] -or [string]::IsNullOrWhiteSpace($settings.$field)) {
                $missingFields += $field
            }
        }
        
        if ($missingFields.Count -gt 0) {
            Write-Host "  WARNING: Missing or empty fields: $($missingFields -join ', ')" -ForegroundColor Yellow
        }
        
        # Check if JAR file exists
        if ($settings.FileUploadUtility) {
            Write-Host "Checking JAR file..." -NoNewline
            if (Test-Path $settings.FileUploadUtility) {
                Write-Host " OK" -ForegroundColor Green
            } else {
                Write-Host " NOT FOUND" -ForegroundColor Red
                Write-Host "  Path: $($settings.FileUploadUtility)" -ForegroundColor Yellow
                Write-Host "  Download from SailPoint documentation" -ForegroundColor Yellow
                $allChecksPassed = $false
            }
        }
        
        # Check if directories exist
        if ($settings.AppFolder) {
            Write-Host "Checking App folder..." -NoNewline
            if (Test-Path $settings.AppFolder) {
                Write-Host " OK" -ForegroundColor Green
                
                # Count app folders
                $appFolders = Get-ChildItem -Path $settings.AppFolder -Directory -ErrorAction SilentlyContinue
                if ($appFolders) {
                    Write-Host "  Found $($appFolders.Count) app folder(s)" -ForegroundColor Cyan
                } else {
                    Write-Host "  No app folders found - run directory creation first" -ForegroundColor Yellow
                }
            } else {
                Write-Host " NOT FOUND" -ForegroundColor Yellow
                Write-Host "  Will be created during directory setup" -ForegroundColor Cyan
            }
        }
        
    } catch {
        Write-Host "  ERROR: Invalid JSON format" -ForegroundColor Red
        Write-Host "  $_" -ForegroundColor Yellow
        $allChecksPassed = $false
    }
} else {
    Write-Host " NOT FOUND" -ForegroundColor Yellow
    Write-Host "  Will be created with defaults on first GUI launch" -ForegroundColor Cyan
}

# Check for required scripts
Write-Host "Checking required scripts..." -NoNewline
$requiredScripts = @(
    "SailpointUtilityGUI.ps1",
    "FileUploadScript.ps1",
    "DirectoryCreateScriptv3.ps1"
)

$missingScripts = @()
foreach ($script in $requiredScripts) {
    if (-not (Test-Path ".\$script")) {
        $missingScripts += $script
    }
}

if ($missingScripts.Count -eq 0) {
    Write-Host " OK" -ForegroundColor Green
} else {
    Write-Host " MISSING" -ForegroundColor Red
    Write-Host "  Missing files: $($missingScripts -join ', ')" -ForegroundColor Yellow
    $allChecksPassed = $false
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan

if ($allChecksPassed) {
    Write-Host "  All critical checks PASSED!" -ForegroundColor Green
    Write-Host "  You're ready to run the utility." -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "  1. Run: .\SailpointUtilityGUI.ps1" -ForegroundColor White
    Write-Host "  2. Configure your settings" -ForegroundColor White
    Write-Host "  3. Create directories" -ForegroundColor White
    Write-Host "  4. Upload files" -ForegroundColor White
} else {
    Write-Host "  Some checks FAILED!" -ForegroundColor Red
    Write-Host "  Please address the issues above." -ForegroundColor Yellow
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Offer to open GUI if all checks passed
if ($allChecksPassed) {
    $response = Read-Host "Launch GUI now? (Y/N)"
    if ($response -eq 'Y' -or $response -eq 'y') {
        .\SailpointUtilityGUI.ps1
    }
}
