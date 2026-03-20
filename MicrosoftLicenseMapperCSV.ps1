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

    .PARAMETER ProductCsvPath
    Path to Microsoft's "Product names and service plan identifiers for licensing" CSV file.

    .PARAMETER OutputDirectory
    Directory where the generated SKU and service plan CSV files will be written.

    .EXAMPLE
    .\MicrosoftLicenseMapperCSV.ps1
    Runs the script with default settings and file paths

    .NOTES
    Author: Tycho Loke
    Website: https://currentcloud.net
    Blog: https://tycholoke.com
    Version: 2.0
    Updated: 05/01/2026

    Prerequisites:
    - PowerShell 7.0 or higher
    - Microsoft.Graph PowerShell module
    - Microsoft 365 admin account with Directory.Read.All permissions
    - "Product names and service plan identifiers for licensing" CSV from Microsoft
      Download from: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference

    .LINK
    https://github.com/TychoLoke/microsoft-365-current-license-mapper
#>

#requires -Version 7.0

[CmdletBinding()]
param(
    [string]$ProductCsvPath = "C:\temp\Product names and service plan identifiers for licensing.csv",
    [string]$OutputDirectory = "C:\temp"
)

function Write-InfoMessage {
    param([string]$Message)
    Write-Information $Message -InformationAction Continue
}

function Write-SuccessMessage {
    param([string]$Message)
    Write-Information $Message -InformationAction Continue
}

#region PowerShell Version Check

# Verify PowerShell 7.0 or higher
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "This script requires PowerShell 7.0 or higher. Current version: $($PSVersionTable.PSVersion)"
    Write-InfoMessage "Download PowerShell from https://github.com/PowerShell/PowerShell/releases"
    Write-InfoMessage "Or install via command: winget install Microsoft.PowerShell"
    exit 1
}

#endregion

#region Script Initialization

# Error handling preference - stop on any error
$ErrorActionPreference = "Stop"

Write-InfoMessage "==============================================="
Write-InfoMessage "  Microsoft 365 SKU/Service Plan Data Export"
Write-InfoMessage "==============================================="
Write-Output ""

#endregion

#region Verify Microsoft Graph Module

Write-InfoMessage "Checking for Microsoft Graph PowerShell module..."

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-InfoMessage "Microsoft Graph module not found. Installing..."
    try {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-SuccessMessage "Microsoft Graph module installed successfully!"
    } catch {
        Write-Error "Failed to install Microsoft Graph module. $($_.Exception.Message)"
        Write-Output ""

        # Check if this is an execution policy issue
        if ($_.Exception.Message -match "not digitally signed" -or $_.Exception.Message -match "execution policy") {
            Write-InfoMessage "This appears to be a PowerShell execution policy issue."
            Write-Output ""
            Write-InfoMessage "To resolve this, run: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser"
            Write-Output ""
            Write-InfoMessage "Alternatively, install manually: Install-Module Microsoft.Graph -Scope CurrentUser -SkipPublisherCheck"
        } else {
            Write-InfoMessage "Please install manually: Install-Module Microsoft.Graph -Scope CurrentUser"
        }

        Write-Output ""
        exit 1
    }
} else {
    Write-SuccessMessage "Microsoft Graph module found!"
}

# Import the required Microsoft Graph modules
Write-InfoMessage "Loading Microsoft Graph modules..."
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Write-SuccessMessage "Microsoft Graph modules loaded successfully!"
} catch {
    Write-Error "Failed to load Microsoft Graph modules. $($_.Exception.Message)"
    Write-InfoMessage "The module may not be properly installed. Try reinstalling Microsoft.Graph."
    Write-Output ""
    exit 1
}

# Verify that the required cmdlets are available
Write-InfoMessage "Verifying required cmdlets..."
if (-not (Get-Command Get-MgSubscribedSku -ErrorAction SilentlyContinue)) {
    Write-Error "Get-MgSubscribedSku cmdlet is not available. The Microsoft Graph module may not be properly installed."
    Write-InfoMessage "Please reinstall Microsoft.Graph."
    Write-Output ""
    exit 1
}
Write-SuccessMessage "All required cmdlets are available!"

Write-Output ""

#endregion

#region Microsoft Graph Connection

Write-InfoMessage "Connecting to Microsoft Graph..."
Write-InfoMessage "You will be prompted to sign in with your Microsoft 365 admin account."
Write-Output ""

try {
    Connect-MgGraph -Scopes "Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-SuccessMessage "Successfully connected to Microsoft Graph!"
    Write-Output ""
} catch {
    Write-Error "Failed to connect to Microsoft Graph. $($_.Exception.Message)"
    Write-InfoMessage "Please ensure you have valid Microsoft 365 admin credentials, Directory.Read.All permissions, and MFA configured if required."
    exit 1
}

#endregion

#region Import Microsoft Reference CSV

Write-Host "Importing Microsoft's product and service plan reference data..." -ForegroundColor Yellow

if (-Not (Test-Path $ProductCsvPath)) {
    Write-Host "Error: CSV file not found at $ProductCsvPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please download the file from:" -ForegroundColor Yellow
    Write-Host "https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference" -ForegroundColor Cyan
    Write-Host "and save it to: $ProductCsvPath" -ForegroundColor Yellow
    Disconnect-MgGraph
    Exit
}

try {
    [array]$Identifiers = Import-Csv -Path $ProductCsvPath
    Write-Host "Successfully imported $($Identifiers.Count) reference entries!" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "Error: Failed to import CSV file." -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph
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
    Exit
}

#endregion

#region Export SKU Data to CSV

Write-Host "Exporting SKU data with friendly names..." -ForegroundColor Yellow

if (-not (Test-Path -Path $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}

$skuCsvPath = Join-Path -Path $OutputDirectory -ChildPath "SkuDataComplete.csv"

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

$servicePlanCsvPath = Join-Path -Path $OutputDirectory -ChildPath "ServicePlanDataComplete.csv"

try {
    $SPData | Sort-Object ServicePlanId -Unique | Export-Csv -NoTypeInformation -Path $servicePlanCsvPath
    Write-Host "Service Plan data exported successfully!" -ForegroundColor Green
    Write-Host "Location: $servicePlanCsvPath" -ForegroundColor Cyan
    Write-Host ""
} catch {
    Write-Host "Error: Failed to export Service Plan data." -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph
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
