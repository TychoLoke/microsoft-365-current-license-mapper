<#
    .SYNOPSIS
    Script to create CSV files for SKUs and Service Plans in a Microsoft 365 tenant.
    For CreateCSVFilesForSKUsAndServicePlans.PS1.

    .DESCRIPTION
    This script connects to Microsoft Graph to retrieve SKU and service plan information,
    exporting them into CSV files for further editing and usage in licensing reports.

    .AUTHOR
    Tycho Loke
    Website: https://currentcloud.net
    Blog: https://tycholoke.com

    .NOTES
    Version: 1.3
    Updated: [Date]
#>

# Prevent PowerShell from closing on error
$ErrorActionPreference = "Stop"

# Ensure Microsoft Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph module..."
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}

# Attempt to Connect to Microsoft Graph with MFA Support
try {
    Write-Host "Connecting to Microsoft Graph with MFA support..."
    Connect-MgGraph -Scopes "Directory.Read.All" -NoWelcome
    Write-Host "Successfully connected to Microsoft Graph!" -ForegroundColor Green
} catch {
    Write-Host "Error: Failed to connect to Microsoft Graph. Check your credentials and MFA settings." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    Exit
}

# Import the Product names and service plan identifiers for licensing
$csvPath = "C:\temp\Product names and service plan identifiers for licensing.csv"
if (-Not (Test-Path $csvPath)) {
    Write-Host "Error: CSV file not found at $csvPath. Please ensure the file is available." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    Exit
}
[array]$Identifiers = Import-Csv -Path $csvPath

# Select SKUs with friendly display name
[array]$SKU_friendly = $Identifiers | Select-Object GUID, String_Id, Product_Display_Name -Unique

# Select service plans with friendly display name
[array]$SP_friendly = $Identifiers | Select-Object Service_Plan_Id, Service_Plan_Name, Service_Plans_Included_Friendly_Names -Unique

# Get products used in tenant
try {
    Write-Host "Fetching subscribed SKUs from Microsoft 365 tenant..."
    [Array]$Skus = Get-MgSubscribedSku
    Write-Host "Successfully retrieved SKU data!" -ForegroundColor Green
} catch {
    Write-Host "Error: Unable to fetch SKU data from Microsoft Graph." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    Exit
}

# Generate CSV of all product SKUs used in tenant
$skuCsvPath = "C:\temp\SkuDataComplete.csv"
try {
    $Skus | Select-Object SkuId, SkuPartNumber, @{Name = "DisplayName"; Expression = { ($SKU_friendly | Where-Object -Property GUID -eq $_.SkuId).Product_Display_Name } } | Export-Csv -NoTypeInformation -Path $skuCsvPath
    Write-Host "SKU data exported to: $skuCsvPath" -ForegroundColor Green
} catch {
    Write-Host "Error: Failed to export SKU data." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    Exit
}

# Generate list of all service plans used in SKUs in tenant
$SPData = [System.Collections.Generic.List[Object]]::new()
ForEach ($S in $Skus) {
    ForEach ($SP in $S.ServicePlans) {
        $SPLine = [PSCustomObject][Ordered]@{
            ServicePlanId          = $SP.ServicePlanId
            ServicePlanName        = $SP.ServicePlanName
            # Use 'Service_Plans_Included_Friendly_Names' from $SP_friendly for 'ServicePlanDisplayName'
            ServicePlanDisplayName = ($SP_friendly | Where-Object { $_.Service_Plan_Id -eq $SP.ServicePlanId }).Service_Plans_Included_Friendly_Names | Select-Object -First 1
        }
        $SPData.Add($SPLine)
    }
}

$servicePlanCsvPath = "C:\Temp\ServicePlanDataComplete.csv"
try {
    $SPData | Sort-Object ServicePlanId -Unique | Export-Csv -NoTypeInformation -Path $servicePlanCsvPath
    Write-Host "Service Plan data exported to: $servicePlanCsvPath" -ForegroundColor Green
} catch {
    Write-Host "Error: Failed to export Service Plan data." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    Exit
}

Write-Host "Script execution completed successfully!" -ForegroundColor Green
Read-Host "Press Enter to exit"
