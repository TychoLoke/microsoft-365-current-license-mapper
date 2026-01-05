<#
    .SYNOPSIS
    Generates CSV reference files containing SKU and Service Plan data for Microsoft 365 licensing.

    .DESCRIPTION
    This script retrieves SKU (Stock Keeping Unit) and Service Plan information from your
    Microsoft 365 tenant and creates CSV files with friendly display names. These files are
    required as input for the main MicrosoftLicenseMapper.ps1 script.

    The script performs the following operations:
    1. Connects to Microsoft Graph with MFA support
    2. Imports Microsoft's product names and service plan identifiers CSV
    3. Retrieves all SKUs currently in use in your tenant
    4. Maps SKU IDs to human-readable display names
    5. Exports two CSV files:
       - SkuDataComplete.csv: SKU information with display names
       - ServicePlanDataComplete.csv: Service plan information with friendly names

    .PARAMETER None
    This script uses hardcoded file paths that can be modified in the configuration section

    .EXAMPLE
    .\MicrosoftLicenseMapperCSV.ps1
    Runs the script with default settings and file paths

    .NOTES
    Author: Tycho Loke
    Website: https://currentcloud.net
    Blog: https://tycholoke.com
    Version: 1.3
    Updated: 27/03/2024

    Prerequisites:
    - Microsoft.Graph PowerShell module
    - Microsoft 365 admin account with Directory.Read.All permissions
    - "Product names and service plan identifiers for licensing" CSV from Microsoft
      Download from: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference

    .LINK
    https://github.com/TychoLoke/microsoft-365-current-license-mapper
#>

#region Script Initialization

# Error handling preference - stop on any error
$ErrorActionPreference = "Stop"

Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  Microsoft 365 SKU/Service Plan Data Export" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

#endregion

#region Verify Microsoft Graph Module

Write-Host "Checking for Microsoft Graph PowerShell module..." -ForegroundColor Yellow

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Microsoft Graph module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "Microsoft Graph module installed successfully!" -ForegroundColor Green
    } catch {
        Write-Host "Error: Failed to install Microsoft Graph module." -ForegroundColor Red
        Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""

        # Check if this is an execution policy issue
        if ($_.Exception.Message -match "not digitally signed" -or $_.Exception.Message -match "execution policy") {
            Write-Host "This appears to be a PowerShell execution policy issue." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "To resolve this, run PowerShell as Administrator and execute:" -ForegroundColor Cyan
            Write-Host "  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor White
            Write-Host ""
            Write-Host "Alternatively, install the module manually:" -ForegroundColor Cyan
            Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser -SkipPublisherCheck" -ForegroundColor White
        } else {
            Write-Host "Please install manually: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
        }

        Write-Host ""
        Read-Host "Press Enter to exit"
        Exit
    }
} else {
    Write-Host "Microsoft Graph module found!" -ForegroundColor Green
}

# Import the required Microsoft Graph modules
Write-Host "Loading Microsoft Graph modules..." -ForegroundColor Yellow
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Write-Host "Microsoft Graph modules loaded successfully!" -ForegroundColor Green
} catch {
    Write-Host "Error: Failed to load Microsoft Graph modules." -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "The module may not be properly installed. Please try:" -ForegroundColor Yellow
    Write-Host "  Uninstall-Module Microsoft.Graph -AllVersions" -ForegroundColor White
    Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser -SkipPublisherCheck" -ForegroundColor White
    Write-Host ""
    Read-Host "Press Enter to exit"
    Exit
}

# Verify that the required cmdlets are available
Write-Host "Verifying required cmdlets..." -ForegroundColor Yellow
if (-not (Get-Command Get-MgSubscribedSku -ErrorAction SilentlyContinue)) {
    Write-Host "Error: Get-MgSubscribedSku cmdlet is not available." -ForegroundColor Red
    Write-Host "The Microsoft Graph module may not be properly installed." -ForegroundColor Red
    Write-Host ""
    Write-Host "Please reinstall the module:" -ForegroundColor Yellow
    Write-Host "  Uninstall-Module Microsoft.Graph -AllVersions" -ForegroundColor White
    Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser -SkipPublisherCheck" -ForegroundColor White
    Write-Host ""
    Read-Host "Press Enter to exit"
    Exit
}
Write-Host "All required cmdlets are available!" -ForegroundColor Green

Write-Host ""

#endregion

#region Microsoft Graph Connection

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
Write-Host "You will be prompted to sign in with your Microsoft 365 admin account." -ForegroundColor Gray
Write-Host ""

try {
    Connect-MgGraph -Scopes "Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph!" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "Error: Failed to connect to Microsoft Graph." -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please ensure:" -ForegroundColor Yellow
    Write-Host "  - You have valid Microsoft 365 admin credentials" -ForegroundColor Yellow
    Write-Host "  - Your account has Directory.Read.All permissions" -ForegroundColor Yellow
    Write-Host "  - MFA is properly configured if required" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    Exit
}

#endregion

#region Import Microsoft Reference CSV

Write-Host "Importing Microsoft's product and service plan reference data..." -ForegroundColor Yellow

# File path configuration - modify this if using a different location
$csvPath = "C:\temp\Product names and service plan identifiers for licensing.csv"

if (-Not (Test-Path $csvPath)) {
    Write-Host "Error: CSV file not found at $csvPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please download the file from:" -ForegroundColor Yellow
    Write-Host "https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference" -ForegroundColor Cyan
    Write-Host "and save it to: $csvPath" -ForegroundColor Yellow
    Disconnect-MgGraph
    Read-Host "Press Enter to exit"
    Exit
}

try {
    [array]$Identifiers = Import-Csv -Path $csvPath
    Write-Host "Successfully imported $($Identifiers.Count) reference entries!" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "Error: Failed to import CSV file." -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph
    Read-Host "Press Enter to exit"
    Exit
}

#endregion

#region Process SKU and Service Plan Mappings

Write-Host "Processing SKU and Service Plan mappings..." -ForegroundColor Yellow

# Create lookup arrays with friendly display names
[array]$SKU_friendly = $Identifiers | Select-Object GUID, String_Id, Product_Display_Name -Unique
[array]$SP_friendly = $Identifiers | Select-Object Service_Plan_Id, Service_Plan_Name, Service_Plans_Included_Friendly_Names -Unique

Write-Host "Created lookup tables for $($SKU_friendly.Count) SKUs and $($SP_friendly.Count) service plans" -ForegroundColor Green
Write-Host ""

#endregion

#region Retrieve Tenant SKU Data

Write-Host "Retrieving subscribed SKUs from your Microsoft 365 tenant..." -ForegroundColor Yellow

try {
    [Array]$Skus = Get-MgSubscribedSku -ErrorAction Stop
    Write-Host "Successfully retrieved $($Skus.Count) SKUs from tenant!" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "Error: Unable to fetch SKU data from Microsoft Graph." -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph
    Read-Host "Press Enter to exit"
    Exit
}

#endregion

#region Export SKU Data to CSV

Write-Host "Exporting SKU data with friendly names..." -ForegroundColor Yellow

# Output file path - modify if using a different location
$skuCsvPath = "C:\temp\SkuDataComplete.csv"

try {
    $Skus | Select-Object SkuId, SkuPartNumber, `
        @{Name = "DisplayName"; Expression = {
            ($SKU_friendly | Where-Object -Property GUID -eq $_.SkuId).Product_Display_Name
        }} | Export-Csv -NoTypeInformation -Path $skuCsvPath

    Write-Host "SKU data exported successfully!" -ForegroundColor Green
    Write-Host "Location: $skuCsvPath" -ForegroundColor Cyan
    Write-Host ""
} catch {
    Write-Host "Error: Failed to export SKU data." -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph
    Read-Host "Press Enter to exit"
    Exit
}

#endregion

#region Build and Export Service Plan Data

Write-Host "Building service plan data with friendly names..." -ForegroundColor Yellow

# Build comprehensive service plan list from all SKUs
$SPData = [System.Collections.Generic.List[Object]]::new()

ForEach ($S in $Skus) {
    ForEach ($SP in $S.ServicePlans) {
        $SPLine = [PSCustomObject][Ordered]@{
            ServicePlanId          = $SP.ServicePlanId
            ServicePlanName        = $SP.ServicePlanName
            ServicePlanDisplayName = ($SP_friendly | Where-Object {
                $_.Service_Plan_Id -eq $SP.ServicePlanId
            }).Service_Plans_Included_Friendly_Names | Select-Object -First 1
        }
        $SPData.Add($SPLine)
    }
}

Write-Host "Processed $($SPData.Count) service plan entries" -ForegroundColor Green

# Output file path - modify if using a different location
$servicePlanCsvPath = "C:\Temp\ServicePlanDataComplete.csv"

try {
    $SPData | Sort-Object ServicePlanId -Unique | Export-Csv -NoTypeInformation -Path $servicePlanCsvPath
    Write-Host "Service Plan data exported successfully!" -ForegroundColor Green
    Write-Host "Location: $servicePlanCsvPath" -ForegroundColor Cyan
    Write-Host ""
} catch {
    Write-Host "Error: Failed to export Service Plan data." -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph
    Read-Host "Press Enter to exit"
    Exit
}

#endregion

#region Completion

Write-Host "===============================================" -ForegroundColor Green
Write-Host "    CSV Generation Completed Successfully!" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Generated Files:" -ForegroundColor Cyan
Write-Host "  1. SKU Data:          $skuCsvPath" -ForegroundColor White
Write-Host "  2. Service Plan Data: $servicePlanCsvPath" -ForegroundColor White
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Cyan
Write-Host "  1. (Optional) Edit SkuDataComplete.csv to add pricing information:" -ForegroundColor Yellow
Write-Host "     - Add a 'Price' column with monthly license costs" -ForegroundColor Gray
Write-Host "     - Add a 'Currency' column (e.g., USD, EUR, GBP)" -ForegroundColor Gray
Write-Host "  2. Run MicrosoftLicenseMapper.ps1 to generate license reports" -ForegroundColor Yellow
Write-Host ""

Disconnect-MgGraph
Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to exit"

#endregion

<#
    DISCLAIMER:
    This script is provided as-is without warranty of any kind. Always test in a non-production
    environment before deploying to production. The author and contributors are not responsible
    for any data loss, service disruption, or issues arising from the use of this script.

    Never run scripts downloaded from the Internet without first validating the code and
    understanding its functionality. Review and customize this script to meet your organization's
    specific needs and security requirements.
#>
