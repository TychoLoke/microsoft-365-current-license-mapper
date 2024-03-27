<#
    .SYNOPSIS
    Script to create CSV files for SKUs and Service Plans in a Microsoft 365 tenant.
    For `CreateCSVFilesForSKUsAndServicePlans.PS1`.

    .DESCRIPTION
    This script connects to Microsoft Graph to retrieve SKU and service plan information,
    exporting them into CSV files for further editing and usage in licensing reports.

    .AUTHOR
    Tycho Loke
    Website: https://currentcloud.net
    Blog: https://tycholoke.com
    Tycho Loke is the creator of this script, dedicated to streamlining Microsoft 365 license
    management and reporting. For more scripts and insights, visit the author's website and blog.

    .NOTES
    Version: 1.0
    Updated: [Date]
#>

Connect-MgGraph -Scope Directory.Read.All -NoWelcome

#Import the Product names and service plan identifiers for licensing CSV file downloaded from https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
# Remember to move the CSV file downloaded from Microsoft to c:\temp\
[array]$Identifiers = Import-Csv -Path "C:\temp\Product names and service plan identifiers for licensing.csv"
#select all SKUs with friendly display name
[array]$SKU_friendly = $identifiers | Select-Object GUID, String_Id, Product_Display_Name -Unique
#select the service plans with friendly display name 
[array]$SP_friendly = $identifiers | Select-Object Service_Plan_Id, Service_Plan_Name, Service_Plans_Included_Friendly_Names -Unique

# Get prpducts used in tenant
[Array]$Skus = Get-MgSubscribedSku

# Generate CSV of all product SKUs used in tenant
$Skus | Select-Object SkuId, SkuPartNumber, @{Name = "DisplayName"; Expression = { ($SKU_friendly | Where-object -Property GUID -eq $_.SkuId).Product_Display_Name } } | Export-Csv -NoTypeInformation c:\temp\SkuDataComplete.csv
# Generate list of all service plans used in SKUs in tenant
$SPData = [System.Collections.Generic.List[Object]]::new()
ForEach ($S in $Skus) {
    ForEach ($SP in $S.ServicePlans) {
        $SPLine = [PSCustomObject][Ordered]@{  
            ServicePlanId          = $SP.ServicePlanId
            ServicePlanName        = $SP.ServicePlanName
            #use 'Service_Plans_Included_Friendly_Names' from $SKU_friendly for 'ServicePlanDisplayName'
            ServicePlanDisplayName = ($SP_friendly | Where-Object { $_.Service_Plan_Id -eq $SP.ServicePlanId }).Service_Plans_Included_Friendly_Names | Select-Object -First 1 
        }
        $SPData.Add($SPLine)
    }
}
$SPData | Sort-Object ServicePlanId -Unique | Export-csv c:\Temp\ServicePlanDataComplete.csv -NoTypeInformation
