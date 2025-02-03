<#
    .SYNOPSIS
    Enhanced script to create CSV and HTML reports for SKUs, Service Plans, and license assignments in a Microsoft 365 tenant.
    Designed for better public usability.

    .DESCRIPTION
    This script connects to Microsoft Graph to retrieve SKU and service plan information, generates structured reports,
    and provides enhanced error handling with user prompts. Outputs results in both CSV and HTML formats.

    .AUTHOR
    Tycho Loke
    Updated by ChatGPT for public usability improvements.

    .NOTES
    Version: 2.1
    Updated: [Date]
#>

# Ensure the script runs with the required permissions
$ErrorActionPreference = "Stop"

Function Connect-ToM365 {
    try {
        Connect-MgGraph -Scope "Directory.Read.All, AuditLog.Read.All, User.Read.All" -NoWelcome
        Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
    } catch {
        Write-Host "Failed to connect to Microsoft Graph. Ensure you have the correct permissions." -ForegroundColor Red
        Exit
    }
}

Function Import-CSVData {
    param (
        [string]$FilePath
    )
    if (!(Test-Path $FilePath)) {
        Write-Host "Error: CSV file not found at $FilePath. Ensure you have downloaded the necessary file." -ForegroundColor Red
        Exit
    }
    return Import-Csv -Path $FilePath
}

Function Generate-SKUReport {
    param (
        [array]$Identifiers, 
        [array]$Skus
    )
    
    $SKU_friendly = $Identifiers | Select-Object GUID, String_Id, Product_Display_Name -Unique
    $Skus | Select-Object SkuId, SkuPartNumber, @{Name = "DisplayName"; Expression = { ($SKU_friendly | Where-Object -Property GUID -eq $_.SkuId).Product_Display_Name } } |
    Export-Csv -NoTypeInformation "C:\temp\SkuDataComplete.csv"
    Write-Host "SKU Data Exported Successfully." -ForegroundColor Green
}

Function Generate-ServicePlanReport {
    param (
        [array]$Identifiers, 
        [array]$Skus
    )
    
    $SP_friendly = $Identifiers | Select-Object Service_Plan_Id, Service_Plan_Name, Service_Plans_Included_Friendly_Names -Unique
    $SPData = [System.Collections.Generic.List[Object]]::new()
    
    ForEach ($S in $Skus) {
        ForEach ($SP in $S.ServicePlans) {
            $SPData.Add([PSCustomObject]@{
                ServicePlanId = $SP.ServicePlanId
                ServicePlanName = $SP.ServicePlanName
                ServicePlanDisplayName = ($SP_friendly | Where-Object { $_.Service_Plan_Id -eq $SP.ServicePlanId }).Service_Plans_Included_Friendly_Names | Select-Object -First 1 
            })
        }
    }
    $SPData | Sort-Object ServicePlanId -Unique | Export-csv "C:\Temp\ServicePlanDataComplete.csv" -NoTypeInformation
    Write-Host "Service Plan Data Exported Successfully." -ForegroundColor Green
}

Function Generate-LicenseReport {
    $Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0" -All -Property DisplayName, UserPrincipalName, AssignedLicenses
    $Report = [System.Collections.Generic.List[Object]]::new()
    
    ForEach ($User in $Users) {
        $LicenseNames = ($User.AssignedLicenses | ForEach-Object { $_.SkuId }) -join ", "
        $Report.Add([PSCustomObject]@{
            DisplayName = $User.DisplayName
            UserPrincipalName = $User.UserPrincipalName
            Licenses = $LicenseNames
        })
    }
    
    $CsvReportPath = "C:\temp\Microsoft365LicensesReport.csv"
    $HtmlReportPath = "C:\temp\Microsoft365LicensesReport.html"
    
    # Export CSV
    $Report | Export-Csv -NoTypeInformation $CsvReportPath
    
    # Export HTML
    $HtmlContent = $Report | ConvertTo-Html -Title "Microsoft 365 License Report" | Out-String
    $HtmlContent | Out-File $HtmlReportPath -Encoding UTF8
    
    Write-Host "Microsoft 365 License Report Generated Successfully in CSV and HTML formats." -ForegroundColor Green
}

# Main Execution
Connect-ToM365

$csvFilePath = "C:\temp\Product names and service plan identifiers for licensing.csv"
$Identifiers = Import-CSVData -FilePath $csvFilePath
$Skus = Get-MgSubscribedSku

Generate-SKUReport -Identifiers $Identifiers -Skus $Skus
Generate-ServicePlanReport -Identifiers $Identifiers -Skus $Skus
Generate-LicenseReport

Write-Host "Script Execution Completed Successfully." -ForegroundColor Cyan
