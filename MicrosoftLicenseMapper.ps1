<#
    .SYNOPSIS
    Generates comprehensive Microsoft 365 license reports with cost analysis and usage insights.

    .DESCRIPTION
    This script connects to Microsoft Graph to analyze and report on Microsoft 365 license assignments.
    It generates detailed reports in both CSV and HTML formats, including:
    - User license assignments (direct and group-based)
    - Cost analysis by user, department, and country
    - Duplicate license detection
    - Inactive account identification
    - Service plan visibility

    .PARAMETER SkuDataPath
    Path to the SKU data CSV file (default: C:\temp\SkuDataComplete.csv)

    .PARAMETER ServicePlanPath
    Path to the service plan data CSV file (default: C:\temp\ServicePlanDataComplete.csv)

    .PARAMETER CSVOutputFile
    Path for the CSV output report (default: C:\temp\Microsoft365LicensesReport.CSV)

    .PARAMETER HtmlReportFile
    Path for the HTML output report (default: C:\temp\Microsoft365LicensesReport.html)

    .EXAMPLE
    .\MicrosoftLicenseMapper.ps1
    Runs the script with default file paths

    .NOTES
    Author: Tycho Löke
    Copyright: (c) 2026 Tycho Löke. All rights reserved.
    Website: https://tycholoke.com
    Portfolio: https://currentcloud.net
    Version: 2.1
    Updated: 05/01/2026

    Requires:
    - PowerShell 7.0 or higher
    - Microsoft.Graph PowerShell module
    - Appropriate Microsoft 365 admin permissions
    - Pre-generated SKU and Service Plan CSV files (run MicrosoftLicenseMapperCSV.ps1 first)

    .LINK
    https://github.com/TychoLoke/microsoft-365-current-license-mapper
    https://tycholoke.com

    .COPYRIGHT
    Copyright (c) 2026 Tycho Löke (tycholoke.com). All rights reserved.
    This script is provided as-is without warranty. Unauthorized redistribution
    or modification without attribution is prohibited.
#>

#requires -Version 7.0

#region PowerShell Version Check

# Verify PowerShell 7.0 or higher
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "===============================================" -ForegroundColor Red
    Write-Host "   PowerShell Version Error" -ForegroundColor Red
    Write-Host "===============================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "This script requires PowerShell 7.0 or higher." -ForegroundColor Yellow
    Write-Host "Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please download and install PowerShell 7 from:" -ForegroundColor Cyan
    Write-Host "https://github.com/PowerShell/PowerShell/releases" -ForegroundColor White
    Write-Host ""
    Write-Host "Or install via command:" -ForegroundColor Cyan
    Write-Host "  winget install Microsoft.PowerShell" -ForegroundColor White
    Write-Host ""
    Exit 1
}

#endregion

Function Get-LicenseCosts {
  <#
  .SYNOPSIS
  Calculates the annual cost of licenses assigned to a user account.

  .DESCRIPTION
  This function computes the total annual licensing cost for a given set of licenses
  by looking up pricing information in the global PricingHashTable.

  .PARAMETER Licenses
  Array of license SKU IDs to calculate costs for

  .OUTPUTS
  Returns the total annual cost as a decimal value
  #>
  [cmdletbinding()]
  Param( [array]$Licenses )

  [int]$Costs = 0

  ForEach ($License in $Licenses) {
    Try {
      [string]$LicenseCost = $PricingHashTable[$License]

      # Convert monthly cost to cents to avoid floating-point precision issues
      # (e.g., licenses costing $16.40/month)
      [float]$LicenseCostCents = [float]$LicenseCost * 100

      If ($LicenseCostCents -gt 0) {
        # Calculate annual cost (monthly cost * 12 months)
        [float]$AnnualCost = $LicenseCostCents * 12

        # Add to cumulative total
        $Costs = $Costs + ($AnnualCost)
      }
    }
    Catch {
      Write-Host ("Warning: Unable to find pricing for license SKU {0}" -f $License) -ForegroundColor Yellow
    }
  }

  # Convert back from cents to currency units
  Return ($Costs / 100)
} 

#region Script Configuration and Initialization

# Script metadata
[datetime]$RunDate = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
$Version = "2.1"

# Default currency (can be overridden by Currency column in SkuDataComplete.csv)
[string]$Currency = "EUR"

Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  Microsoft 365 License Mapper v$Version" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

# File paths - Modify these if using different locations
$SkuDataPath = "C:\temp\SkuDataComplete.csv"
$ServicePlanPath = "C:\temp\ServicePlanDataComplete.csv"
$CSVOutputFile = "C:\temp\Microsoft365LicensesReport.CSV"
$HtmlReportFile = "C:\temp\Microsoft365LicensesReport.html"

# Initialize counters
$UnlicensedAccounts = 0

#endregion

#region Microsoft Graph Connection

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
Write-Host "You will be prompted to sign in with your Microsoft 365 admin account." -ForegroundColor Gray
Write-Host ""

Try {
  Connect-MgGraph -Scope "Directory.AccessAsUser.All, Directory.Read.All, AuditLog.Read.All" -NoWelcome -ErrorAction Stop
  Write-Host "Successfully connected to Microsoft Graph!" -ForegroundColor Green
  Write-Host ""
}
Catch {
  Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
  Write-Host "Error: $_" -ForegroundColor Red
  Write-Host ""
  Write-Host "Please ensure you have:" -ForegroundColor Yellow
  Write-Host "  - Valid Microsoft 365 admin credentials" -ForegroundColor Yellow
  Write-Host "  - Appropriate permissions (Global Admin, Global Reader, or License Admin)" -ForegroundColor Yellow
  Write-Host "  - Microsoft Graph PowerShell SDK installed" -ForegroundColor Yellow
  Exit
}

#endregion

#region Validate Required Files

Write-Host "Validating required CSV files..." -ForegroundColor Yellow

If ((Test-Path $skuDataPath) -eq $False) {
  Write-Host ("Can't find the product data file ({0}). Exiting..." -f $skuDataPath) -ForegroundColor Red
  Write-Host "Please run MicrosoftLicenseMapperCSV.ps1 first to generate the required CSV files." -ForegroundColor Yellow
  Disconnect-MgGraph
  Exit
}

If ((Test-Path $servicePlanPath) -eq $False) {
  Write-Host ("Can't find the service plan data file ({0}). Exiting..." -f $servicePlanPath) -ForegroundColor Red
  Write-Host "Please run MicrosoftLicenseMapperCSV.ps1 first to generate the required CSV files." -ForegroundColor Yellow
  Disconnect-MgGraph
  Exit
}

Write-Host "Required files found!" -ForegroundColor Green
Write-Host ""

#endregion

#region Load and Process SKU Data

Write-Host "Loading SKU and pricing data..." -ForegroundColor Yellow

# Import SKU data from CSV
$ImportSkus = Import-CSV $skuDataPath

# Initialize hash tables for SKU and pricing lookups
$SkuHashTable = @{}
$PricingHashTable = @{}


# Build SKU lookup hash table (maps SKU IDs to friendly display names)
ForEach ($Line in $ImportSkus) {
  If (-not [string]::IsNullOrWhiteSpace($Line.SkuId)) {
    If (-not $SkuHashTable.ContainsKey([string]$Line.SkuId)) {
      $SkuHashTable.Add([string]$Line.SkuId, [string]$Line.DisplayName)
    } Else {
      Write-Host ("Warning: Duplicate SKU ID detected and skipped: " + $Line.SkuId) -ForegroundColor Yellow
    }
  } Else {
    Write-Host "Warning: Found an entry with null or empty SkuId, skipping..." -ForegroundColor Yellow
  }
}

# Check if pricing information is available and populate pricing hash table
$PricingInfoAvailable = $False

If ($ImportSkus[0].Price) {
  Write-Host "Pricing information detected - cost analysis will be included in reports" -ForegroundColor Green
  $PricingInfoAvailable = $True
  $Global:PricingHashTable = @{}

  ForEach ($Line in $ImportSkus) {
    If (-not [string]::IsNullOrWhiteSpace($Line.SkuId) -and -not [string]::IsNullOrWhiteSpace($Line.Price)) {
      $PricingHashTable.Add([string]$Line.SkuId, [string]$Line.Price)
    }
  }

  # Override default currency if specified in CSV
  If ($ImportSkus[0].Currency) {
    [string]$Currency = ($ImportSkus[0].Currency)
    Write-Host "Currency set to: $Currency" -ForegroundColor Cyan
  }
} Else {
  Write-Host "No pricing information found - cost analysis will be unavailable" -ForegroundColor Yellow
  Write-Host "To enable cost analysis, add 'Price' and 'Currency' columns to SkuDataComplete.csv" -ForegroundColor Gray
}

Write-Host ""

#endregion

#region Retrieve Licensed User Accounts

Write-Host "Retrieving licensed user accounts from Microsoft 365..." -ForegroundColor Yellow

$Users = Get-MgUser -All -ConsistencyLevel eventual -CountVariable Records `
  -Property id, displayName, userPrincipalName, country, department, assignedLicenses, `
  licenseAssignmentStates, createdDateTime, jobTitle, signInActivity, companyName | `
  Where-Object { $_.AssignedLicenses.Count -gt 0 } | Sort-Object DisplayName

If (!($Users)) {
  Write-Host "No licensed user accounts found in the tenant." -ForegroundColor Yellow
  Disconnect-MgGraph
  Exit
}
Else {
  Write-Host ("{0} licensed user accounts found!" -f $Users.Count) -ForegroundColor Green
  Write-Host ""
}

# Get organization information and unique department/country values
[array]$Departments = $Users.Department | Sort-Object -Unique
[array]$Countries = $Users.Country | Sort-Object -Unique
$OrgName = (Get-MgOrganization).DisplayName

# Initialize tracking variables
$DuplicateSKUsAccounts = 0
$DuplicateSKULicenses = 0
$LicenseErrorCount = 0
$Report = [System.Collections.Generic.List[Object]]::new()
$i = 0
[float]$TotalUserLicenseCosts = 0
[float]$TotalBoughtLicenseCosts = 0

#endregion

#region Process Each User Account

Write-Host "Processing license assignments for each user..." -ForegroundColor Cyan
Write-Host ""

ForEach ($User in $Users) {
  $UnusedAccountWarning = "OK"; $i++; $UserCosts = 0
  $ErrorMsg = ""; $LastLicenseChange = ""
  Write-Host ("Processing account {0} {1}/{2}" -f $User.UserPrincipalName, $i, $Users.Count)
  If ([string]::IsNullOrWhiteSpace($User.licenseAssignmentStates) -eq $False) {
    # Only process account if it has some licenses
    [array]$LicenseInfo = $Null; [array]$DisabledPlans = $Null; 
    #  Find out if any of the user's licenses are assigned via group-based licensing
    [array]$GroupAssignments = $User.licenseAssignmentStates | 
      Where-Object { $null -ne $_.AssignedByGroup -and $_.State -eq "Active" }
    #  Find out if any of the user's licenses are assigned via group-based licensing and having an error
    [array]$GroupErrorAssignments = $User.licenseAssignmentStates | 
      Where-Object { $Null -ne $_.AssignedByGroup -and $_.State -eq "Error" }
    [array]$GroupLicensing = $Null
    # Find out when the last license change was made
    If ([string]::IsNullOrWhiteSpace($User.licenseAssignmentStates.lastupdateddatetime) -eq $False) {
      $LastLicenseChange = Get-Date(($user.LicenseAssignmentStates.lastupdateddatetime | Measure-Object -Maximum).Maximum) -format g
    }
    # Figure out group-based licensing assignments if any exist
    ForEach ($G in $GroupAssignments) {
      $GroupName = (Get-MgGroup -GroupId $G.AssignedByGroup).DisplayName
      $GroupProductName = $SkuHashTable[$G.SkuId]
      $GroupLicensing += ("{0} assigned from {1}" -f $GroupProductName, $GroupName)
    }
    ForEach ($G in $GroupErrorAssignments) {
      $GroupName = (Get-MgGroup -GroupId $G.AssignedByGroup).DisplayName
      $GroupProductName = $SkuHashTable[$G.SkuId]
      $ErrorMsg = $G.Error
      $LicenseErrorCount++
      $GroupLicensing += ("{0} assigned from {1} BUT ERROR {2}!" -f $GroupProductName, $GroupName, $ErrorMsg)
    }
    $GroupLicensingAssignments = $GroupLicensing -Join ", "

    #  Find out if any of the user's licenses are assigned via direct licensing
    [array]$DirectAssignments = $User.licenseAssignmentStates | 
      Where-Object { $null -eq $_.AssignedByGroup -and $_.State -eq "Active" }

    # Figure out details of direct assigned licenses
    [array]$UserLicenses = $User.AssignedLicenses
    ForEach ($License in $DirectAssignments) {
      If ($SkuHashTable.ContainsKey($License.SkuId) -eq $True) {
        # We found a match in the SKU hash table
        $LicenseInfo += $SkuHashTable.Item($License.SkuId) 
      } Else {
        # Nothing found in the SKU hash table, so output the SkuID
        $LicenseInfo += $License.SkuId
      }
    }

# Iterate over each license in the user's assigned licenses
ForEach ($License in $UserLicenses) {
    # Check if the license has any disabled plans
    If (-not [string]::IsNullOrWhiteSpace($License.DisabledPlans)) {
        # Iterate over each disabled plan in the current license
        ForEach ($DisabledPlan in $License.DisabledPlans) {
            # Ensure $ServicePlanHashTable is not null before checking it
            If ($null -ne $ServicePlanHashTable -and $ServicePlanHashTable.ContainsKey($DisabledPlan)) {
                # We found a match in the Service Plans hash table
                $DisabledPlans += $ServicePlanHashTable.Item($DisabledPlan)
            }
            Else {
                # Handle the case where the plan is not found or ServicePlanHashTable is null
                Write-Host "Warning: ServicePlanHashTable is null or does not contain the plan: $DisabledPlan"
                # Optionally collect these for later review or logging
                $DisabledPlans += $DisabledPlan
            }
        }
    }
}


    # Detect if any duplicate licenses are assigned (direct and group-based)
    # Build a list of assigned SKUs
    $SkuUserReport = [System.Collections.Generic.List[Object]]::new()
    ForEach ($S in $DirectAssignments) {
      $ReportLine = [PSCustomObject][Ordered]@{ 
        User   = $User.Id
        Name   = $User.DisplayName 
        Sku    = $S.SkuId
        Method = "Direct" 
      }
      $SkuUserReport.Add($ReportLine)
    }
    ForEach ($S in $GroupAssignments) {
      $ReportLine = [PSCustomObject][Ordered]@{ 
        User   = $User.Id
        Name   = $User.DisplayName
        Sku    = $S.SkuId
        Method = "Group" 
      }
      $SkuUserReport.Add($ReportLine)
    }

    # Check if any duplicates exist
    [array]$DuplicateSkus = $SkuUserReport | Group-Object Sku | 
      Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name

    # If duplicates exist, resolve their SKU IDs into Product names and generate a warning for the report
    [string]$DuplicateWarningReport = "N/A"
    If ($DuplicateSkus) {
      [array]$DuplicateSkuNames = $Null
      $DuplicateSKUsAccounts++
      $DuplicateSKULicenses = $DuplicateSKULicenses + $DuplicateSKUs.Count
      ForEach ($DS in $DuplicateSkus) {
        $SkuName = $SkuHashTable[$DS]
        $DuplicateSkuNames += $SkuName
      }
      $DuplicateWarningReport = ("Warning: Duplicate licenses detected for: {0}" -f ($DuplicateSkuNames -join ", "))
    }
  } Else { 
      $UnlicensedAccounts++
  }

  $LastSignIn = $User.SignInActivity.LastSignInDateTime
  $LastNonInteractiveSignIn = $User.SignInActivity.LastNonInteractiveSignInDateTime

  If (-not $LastSignIn -and -not $LastNonInteractiveSignIn) {
      $DaysSinceLastSignIn = "Never"
      $UnusedAccountWarning = "Never logged in - Cleanup candidate"
      $LastAccess = "Never"
  } Else {
    # Get the newest date, if both dates contain values
    If ($LastSignIn -and $LastNonInteractiveSignIn) {
      If ($LastSignIn -gt $LastNonInteractiveSignIn) {
          $CompareDate = $LastSignIn
      } Else {
          $CompareDate = $LastNonInteractiveSignIn
      }
    } Elseif ($LastSignIn) {
        # Only $LastSignIn has a value
        $CompareDate = $LastSignIn
    } Else {
        # Only $LastNonInteractiveSignIn has a value
        $CompareDate = $LastNonInteractiveSignIn
    }

    $DaysSinceLastSignIn = ($RunDate - $CompareDate).Days
    $LastAccess = Get-Date($CompareDate) -format g

    # Enhanced status categorization for cleanup scenarios
    If ($DaysSinceLastSignIn -gt 180) {
      $UnusedAccountWarning = "Inactive 180+ days - High priority cleanup"
    } ElseIf ($DaysSinceLastSignIn -gt 90) {
      $UnusedAccountWarning = "Inactive 90+ days - Cleanup candidate"
    } ElseIf ($DaysSinceLastSignIn -gt 60) {
      $UnusedAccountWarning = "Inactive 60+ days - Review recommended"
    } ElseIf ($DaysSinceLastSignIn -gt 30) {
      $UnusedAccountWarning = "Inactive 30+ days - Monitor"
    }
  }

  $AccountCreatedDate = $null
  If ($User.CreatedDateTime) {
    $AccountCreatedDate = Get-Date($User.CreatedDateTime) -format 'dd-MMM-yyyy HH:mm' 
  }

  # Report information
  [string]$DisabledPlans = $DisabledPlans -join ", " 
  [string]$LicenseInfo = $LicenseInfo -join (", ")

  If ($PricingInfoAvailable) { 
    # Output report line with pricing info
    [float]$UserCosts = Get-LicenseCosts -Licenses $UserLicenses.SkuId
    $TotalUserLicenseCosts = $TotalUserLicenseCosts + $UserCosts
    $ReportLine = [PSCustomObject][Ordered]@{  
      User                       = $User.DisplayName
      UPN                        = $User.UserPrincipalName
      Country                    = $User.Country
      Department                 = $User.Department
      Title                      = $User.JobTitle
      Company                    = $User.companyName
      "Direct assigned licenses" = $LicenseInfo
      "Disabled Plans"           = $DisabledPlans 
      "Group based licenses"     = $GroupLicensingAssignments
      "Annual License Costs"     = ("{0} {1}" -f $Currency, ($UserCosts.toString('F2')))
      "Last license change"      = $LastLicenseChange
      "Account created"          = $AccountCreatedDate
      "Last Signin"              = $LastAccess
      "Days since last signin"   = $DaysSinceLastSignIn
      "Duplicates detected"      = $DuplicateWarningReport
      Status                     = $UnusedAccountWarning
      UserCosts                  = $UserCosts  
    }
  } Else { 
    # No pricing information
    $ReportLine = [PSCustomObject][Ordered]@{  
      User                       = $User.DisplayName
      UPN                        = $User.UserPrincipalName
      Country                    = $User.Country
      Department                 = $User.Department
      Title                      = $User.JobTitle
      Company                    = $User.companyName
      "Direct assigned licenses" = $LicenseInfo
      "Disabled Plans"           = $DisabledPlans 
      "Group based licenses"     = $GroupLicensingAssignments
      "Last license change"      = $LastLicenseChange
      "Account created"          = $AccountCreatedDate
      "Last Signin"              = $LastAccess
      "Days since last signin"   = $DaysSinceLastSignIn
      "Duplicates detected"      = $DuplicateWarningReport
      Status                     = $UnusedAccountWarning
    }
  }  
  $Report.Add($ReportLine)
} # End ForEach Users

$UnderusedAccounts = $Report | Where-Object { $_.Status -ne "OK" }
$PercentUnderusedAccounts = ($UnderUsedAccounts.Count / $Report.Count).toString("P")

# Enhanced cleanup statistics - Fixed to be cumulative
$NeverLoggedIn = $Report | Where-Object { $_.'Last Signin' -eq "Never" }
$Inactive180Plus = $Report | Where-Object { $_.Status -like "*180+ days*" }
# Inactive 90+ includes both 90-179 days AND 180+ days
$Inactive90Plus = $Report | Where-Object { $_.Status -like "*90+ days*" -or $_.Status -like "*180+ days*" }
# Inactive 60+ includes 60-89 days AND 90+ days AND 180+ days
$Inactive60Plus = $Report | Where-Object { $_.Status -like "*60+ days*" -or $_.Status -like "*90+ days*" -or $_.Status -like "*180+ days*" }
# Inactive 30+ includes all inactive categories
$Inactive30Plus = $Report | Where-Object { $_.Status -like "*30+ days*" -or $_.Status -like "*60+ days*" -or $_.Status -like "*90+ days*" -or $_.Status -like "*180+ days*" }
$HighPriorityCleanup = $Report | Where-Object { $_.Status -like "*Cleanup candidate*" -or $_.Status -like "*High priority*" -or $_.'Last Signin' -eq "Never" }

# This code grabs the SKU summary for the tenant and uses the data to create a SKU summary usage segment for the HTML report
$SkuReport = [System.Collections.Generic.List[Object]]::new()
[array]$SkuSummary = Get-MgSubscribedSku | Select-Object SkuId, ConsumedUnits, PrepaidUnits
$SkuSummary = $SkuSummary | Where-Object { $_.ConsumedUnits -ne 0 }
ForEach ($S in $SkuSummary) {
  $SKUCost = Get-LicenseCosts -Licenses $S.SkuId
  $SkuDisplayName = $SkuHashtable[$S.SkuId]
  If ($S.PrepaidUnits.Enabled -le $S.ConsumedUnits ) {
    $BoughtUnits = $S.ConsumedUnits 
  } Else {
    $BoughtUnits = $S.PrepaidUnits.Enabled
  }
  If ($PricingInfoAvailable) {
    $SKUTotalCost = ($SKUCost * $BoughtUnits)
    $SkuReportLine = [PSCustomObject][Ordered]@{  
      "SKU Id"                = $S.SkuId
      "SKU Name"              = $SkuDisplayName 
      "Units Used"            = $S.ConsumedUnits 
      "Units Purchased"       = $BoughtUnits
      "Annual license costs"  = $SKUTotalCost
      "Annual licensing cost" = ("{0} {1}" -f $Currency, ('{0:N2}' -f $SKUTotalCost))
    }
  } Else {
    $SkuReportLine = [PSCustomObject][Ordered]@{  
      "SKU Id"          = $S.SkuId
      "SKU Name"        = $SkuDisplayName 
      "Units Used"      = $S.ConsumedUnits 
      "Units Purchased" = $BoughtUnits
    }
  }
  $SkuReport.Add($SkuReportLine) 
  $TotalBoughtLicenseCosts = $TotalBoughtLicenseCosts + $SKUTotalCost
}

If ($PricingInfoAvailable) {
  $AverageCostPerUser = ($TotalUserLicenseCosts / $Users.Count)
  $AverageCostPerUserOutput = ("{0} {1}" -f $Currency, ('{0:N2}' -f $AverageCostPerUser))
  $TotalUserLicenseCostsOutput = ("{0} {1}" -f $Currency, ('{0:N2}' -f $TotalUserLicenseCosts))
  $TotalBoughtLicenseCostsOutput = ("{0} {1}" -f $Currency, ('{0:N2}' -f $TotalBoughtLicenseCosts))
  $PercentBoughtLicensesUsed = ($TotalUserLicenseCosts / $TotalBoughtLicenseCosts).toString('P')
  $SkuReport = $SkuReport | Sort-Object "Annual license costs" -Descending
} Else {
  $SkuReport = $SkuReport | Sort-Object "SKU Name" -Descending
}

If ($PricingInfoAvailable) { 
  # Generate the department analysis
  $DepartmentReport = [System.Collections.Generic.List[Object]]::new()
  ForEach ($Department in $Departments) {
    $DepartmentRecords = $Report | Where-Object Department -match $Department
    $DepartmentReportLine = [PSCustomObject][Ordered]@{
      Department  = $Department
      Accounts    = $DepartmentRecords.count
      Costs       = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($DepartmentRecords | Measure-Object UserCosts -Sum).Sum))
      AverageCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($DepartmentRecords | Measure-Object UserCosts -Average).Average))
    } 
    $DepartmentReport.Add($DepartmentReportLine)
  }
  $DepartmentHTML = $DepartmentReport | ConvertTo-HTML -Fragment
  # Anyone without a department?
  [array]$NoDepartments = $Report | Where-Object { $null -eq $_.Department }
  $NoDepartmentCosts = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($NoDepartments | Measure-Object UserCosts -Sum).Sum))

  # Generate the country analysis
  $CountryReport = [System.Collections.Generic.List[Object]]::new()
  ForEach ($Country in $Countries) {
    $CountryRecords = $Report | Where-Object Country -match $Country
    $CountryReportLine = [PSCustomObject][Ordered]@{
      Country     = $Country
      Accounts    = $CountryRecords.count
      Costs       = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CountryRecords | Measure-Object UserCosts -Sum).Sum))
      AverageCost = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($CountryRecords | Measure-Object UserCosts -Average).Average))
    } 
    $CountryReport.Add($CountryReportLine)
  }
  $CountryHTML = $CountryReport | ConvertTo-HTML -Fragment
  # Anyone without a country?
  [array]$NoCountry = $Report | Where-Object { $null -eq $_.Country }
  $NoCountryCosts = ("{0} {1}" -f $Currency, ('{0:N2}' -f ($NoCountry | Measure-Object UserCosts -Sum).Sum))
}

# Generate table rows dynamically
#region Generate Professional HTML Report with Dark Mode and Charts

# Create the HTML report with advanced features
$HtmlHead = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 License Report - $OrgName</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        :root {
            --primary-color: #0078d4;
            --primary-hover: #106ebe;
            --success-color: #10893e;
            --success-light: #dff6dd;
            --warning-color: #ff8c00;
            --warning-light: #fff4ce;
            --danger-color: #d13438;
            --danger-light: #fde7e9;
            --info-color: #0099bc;
            --info-light: #cff4fc;
            --dark-bg: #1e1e1e;
            --light-bg: #f8f9fa;
            --card-bg: #ffffff;
            --text-primary: #323130;
            --text-secondary: #605e5c;
            --border-color: #e1dfdd;
            --shadow: 0 3px 12px rgba(0,0,0,0.12);
            --shadow-hover: 0 6px 20px rgba(0,0,0,0.18);
            --header-gradient-start: #0078d4;
            --header-gradient-end: #005a9e;
            --accent-purple: #8b5cf6;
            --accent-orange: #ff6b35;
            --accent-teal: #14b8a6;
            --accent-pink: #ec4899;
        }

        [data-theme="dark"] {
            --primary-color: #60a5fa;
            --primary-hover: #3b82f6;
            --success-color: #34d399;
            --success-light: rgba(52, 211, 153, 0.15);
            --warning-color: #fbbf24;
            --warning-light: rgba(251, 191, 36, 0.15);
            --danger-color: #f87171;
            --danger-light: rgba(248, 113, 113, 0.15);
            --info-color: #22d3ee;
            --info-light: rgba(34, 211, 238, 0.15);
            --dark-bg: #0f0f0f;
            --light-bg: #1a1a1a;
            --card-bg: #262626;
            --text-primary: #f5f5f5;
            --text-secondary: #a3a3a3;
            --border-color: #404040;
            --shadow: 0 4px 16px rgba(0,0,0,0.6);
            --shadow-hover: 0 8px 24px rgba(0,0,0,0.8);
            --header-gradient-start: #1e40af;
            --header-gradient-end: #1e3a8a;
            --accent-purple: #a78bfa;
            --accent-orange: #fb923c;
            --accent-teal: #2dd4bf;
            --accent-pink: #f472b6;
        }

        html {
            scroll-behavior: smooth;
        }

        body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif;
            font-size: 14px;
            line-height: 1.6;
            color: var(--text-primary);
            background: var(--light-bg);
            padding: 20px;
            transition: background 0.3s ease, color 0.3s ease;
        }

        .container {
            max-width: 1600px;
            margin: 0 auto;
            background: var(--card-bg);
            border-radius: 12px;
            box-shadow: var(--shadow);
            overflow: hidden;
            transition: background 0.3s ease;
        }

        /* Header Styles with Banner Image */
        .header {
            background-image: linear-gradient(rgba(0, 120, 212, 0.85), rgba(0, 90, 158, 0.90)), url('https://p1-ofp.static.pub/ShareResource/na/faqs/img/microsoft-Office-365-sub-hero-banner.jpg');
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
            color: white;
            padding: 60px 40px;
            text-align: center;
            position: relative;
        }

        .header::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(135deg, rgba(0, 120, 212, 0.75) 0%, rgba(0, 90, 158, 0.80) 100%);
            z-index: 0;
        }

        .header > * {
            position: relative;
            z-index: 1;
        }

        .header h1 {
            font-size: 42px;
            font-weight: 600;
            margin-bottom: 12px;
            text-shadow: 0 3px 6px rgba(0,0,0,0.4);
            letter-spacing: -0.5px;
        }

        .header h2 {
            font-size: 24px;
            font-weight: 500;
            opacity: 1;
            margin-bottom: 10px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
        }

        .header h3 {
            font-size: 15px;
            font-weight: 400;
            opacity: 0.95;
            text-shadow: 0 1px 3px rgba(0,0,0,0.3);
        }

        .header .brought-by {
            margin-top: 20px;
            padding-top: 20px;
            border-top: 1px solid rgba(255, 255, 255, 0.3);
            font-size: 14px;
            opacity: 0.95;
        }

        .header .brought-by a {
            color: #fff;
            text-decoration: none;
            font-weight: 600;
            transition: all 0.3s ease;
            border-bottom: 1px solid rgba(255, 255, 255, 0.5);
        }

        .header .brought-by a:hover {
            opacity: 0.8;
            border-bottom-color: #fff;
        }

        .header .kofi-link {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: rgba(255, 94, 77, 0.9);
            padding: 6px 16px;
            border-radius: 20px;
            margin-left: 8px;
            font-weight: 600;
            font-size: 13px;
            transition: all 0.3s ease;
            border: 2px solid rgba(255, 255, 255, 0.3);
        }

        .header .kofi-link:hover {
            background: rgba(255, 94, 77, 1);
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
            border-color: rgba(255, 255, 255, 0.6);
        }

        /* Toolbar */
        .toolbar {
            background: var(--card-bg);
            border-bottom: 1px solid var(--border-color);
            padding: 15px 30px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 15px;
            transition: background 0.3s ease;
        }

        .search-box {
            display: flex;
            align-items: center;
            gap: 10px;
            flex: 1;
            max-width: 500px;
        }

        .search-box input {
            flex: 1;
            padding: 10px 15px;
            border: 2px solid var(--border-color);
            border-radius: 6px;
            font-size: 14px;
            background: var(--card-bg);
            color: var(--text-primary);
            transition: all 0.3s ease;
        }

        .search-box input:focus {
            outline: none;
            border-color: var(--primary-color);
            box-shadow: 0 0 0 3px rgba(0,120,212,0.1);
        }

        .toolbar-buttons {
            display: flex;
            gap: 10px;
            align-items: center;
        }

        .btn {
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.3s ease;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }

        .btn-primary {
            background: var(--primary-color);
            color: white;
        }

        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,120,212,0.3);
        }

        .btn-secondary {
            background: var(--card-bg);
            color: var(--text-primary);
            border: 2px solid var(--border-color);
        }

        .btn-secondary:hover {
            background: var(--light-bg);
            transform: translateY(-1px);
        }

        .filter-btn:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 6px 12px rgba(0,0,0,0.15) !important;
        }

        .filter-btn:active {
            transform: translateY(0) !important;
        }

        .theme-toggle {
            background: var(--card-bg);
            border: 2px solid var(--border-color);
            color: var(--text-primary);
            width: 44px;
            height: 44px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            transition: all 0.3s ease;
        }

        .theme-toggle:hover {
            transform: rotate(180deg);
            background: var(--primary-color);
            color: white;
            border-color: var(--primary-color);
        }

        /* Dashboard Cards */
        .dashboard {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 20px;
            padding: 30px;
            background: var(--light-bg);
        }

        .stat-card {
            background: var(--card-bg);
            border-radius: 16px;
            padding: 28px;
            box-shadow: var(--shadow);
            border-left: 5px solid var(--primary-color);
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
            overflow: hidden;
        }

        .stat-card::before {
            content: '';
            position: absolute;
            top: -50%;
            right: -20%;
            width: 200px;
            height: 200px;
            background: radial-gradient(circle, var(--primary-color) 0%, transparent 70%);
            opacity: 0.08;
            transition: all 0.4s ease;
        }

        .stat-card:hover {
            transform: translateY(-6px) scale(1.02);
            box-shadow: var(--shadow-hover);
        }

        .stat-card:hover::before {
            opacity: 0.12;
            transform: scale(1.2);
        }

        .stat-card.success {
            border-left-color: var(--success-color);
            background: linear-gradient(135deg, var(--card-bg) 0%, var(--success-light) 100%);
        }
        .stat-card.success::before { background: radial-gradient(circle, var(--success-color) 0%, transparent 70%); }

        .stat-card.warning {
            border-left-color: var(--warning-color);
            background: linear-gradient(135deg, var(--card-bg) 0%, var(--warning-light) 100%);
        }
        .stat-card.warning::before { background: radial-gradient(circle, var(--warning-color) 0%, transparent 70%); }

        .stat-card.danger {
            border-left-color: var(--danger-color);
            background: linear-gradient(135deg, var(--card-bg) 0%, var(--danger-light) 100%);
        }
        .stat-card.danger::before { background: radial-gradient(circle, var(--danger-color) 0%, transparent 70%); }

        .stat-card.info {
            border-left-color: var(--info-color);
            background: linear-gradient(135deg, var(--card-bg) 0%, var(--info-light) 100%);
        }
        .stat-card.info::before { background: radial-gradient(circle, var(--info-color) 0%, transparent 70%); }

        .stat-card .label {
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--text-secondary);
            margin-bottom: 8px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .stat-card .value {
            font-size: 36px;
            font-weight: 300;
            color: var(--text-primary);
            line-height: 1.2;
            margin: 10px 0;
        }

        .stat-card .subtitle {
            font-size: 13px;
            color: var(--text-secondary);
            margin-top: 8px;
        }

        /* Content Sections */
        .content {
            padding: 30px;
            background: var(--light-bg);
        }

        .section {
            margin-bottom: 40px;
            background: var(--card-bg);
            border-radius: 12px;
            padding: 30px;
            box-shadow: var(--shadow);
            transition: background 0.3s ease;
        }

        .section-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            flex-wrap: wrap;
            gap: 15px;
        }

        .section-title {
            font-size: 24px;
            font-weight: 400;
            color: var(--text-primary);
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .section-title i {
            color: var(--primary-color);
        }

        /* Chart Containers */
        .charts-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 30px;
            margin-bottom: 30px;
        }

        .chart-container {
            background: var(--card-bg);
            border-radius: 12px;
            padding: 20px;
            box-shadow: var(--shadow);
            transition: all 0.3s ease;
        }

        .chart-container:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 12px rgba(0,0,0,0.15);
        }

        .chart-title {
            font-size: 16px;
            font-weight: 600;
            color: var(--text-primary);
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        /* Modern Table Styles */
        .table-wrapper {
            background: var(--card-bg);
            border-radius: 12px;
            overflow: hidden;
            box-shadow: var(--shadow);
        }

        .table-controls {
            padding: 15px 20px;
            background: var(--light-bg);
            border-bottom: 1px solid var(--border-color);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
        }

        .table-info {
            font-size: 13px;
            color: var(--text-secondary);
        }

        .table-container {
            overflow-x: auto;
            max-height: 600px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            background: var(--card-bg);
            font-size: 13px;
        }

        thead {
            background: var(--light-bg);
            position: sticky;
            top: 0;
            z-index: 10;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        th {
            padding: 16px 12px;
            text-align: left;
            font-weight: 600;
            color: var(--text-primary);
            border-bottom: 2px solid var(--border-color);
            cursor: pointer;
            user-select: none;
            transition: background 0.2s;
            white-space: nowrap;
        }

        th:hover {
            background: var(--card-bg);
        }

        th.sortable::after {
            content: ' ↕';
            opacity: 0.3;
            font-size: 10px;
        }

        th.sorted-asc::after {
            content: ' ↑';
            opacity: 1;
            color: var(--primary-color);
        }

        th.sorted-desc::after {
            content: ' ↓';
            opacity: 1;
            color: var(--primary-color);
        }

        tbody tr {
            border-bottom: 1px solid var(--border-color);
            transition: all 0.2s;
        }

        tbody tr:nth-child(even) {
            background: rgba(0, 0, 0, 0.02);
        }

        [data-theme="dark"] tbody tr:nth-child(even) {
            background: rgba(255, 255, 255, 0.02);
        }

        tbody tr:hover {
            background: var(--light-bg) !important;
            transform: scale(1.005);
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }

        tbody tr.hidden {
            display: none;
        }

        td {
            padding: 14px 12px;
            color: var(--text-primary);
        }

        /* Status Badges */
        .badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .badge-success {
            background: rgba(16, 124, 16, 0.1);
            color: var(--success-color);
        }

        .badge-warning {
            background: rgba(255, 185, 0, 0.1);
            color: var(--warning-color);
        }

        .badge-danger {
            background: rgba(209, 52, 56, 0.1);
            color: var(--danger-color);
        }

        /* Footer */
        .footer {
            background: var(--dark-bg);
            color: white;
            padding: 30px;
            text-align: center;
            font-size: 13px;
        }

        .footer a {
            color: var(--info-color);
            text-decoration: none;
            transition: opacity 0.3s;
        }

        .footer a:hover {
            opacity: 0.8;
        }

        /* Loading Spinner */
        .loading {
            display: inline-block;
            width: 20px;
            height: 20px;
            border: 3px solid rgba(0,120,212,0.3);
            border-radius: 50%;
            border-top-color: var(--primary-color);
            animation: spin 1s linear infinite;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        /* Print Styles */
        @media print {
            body {
                background: white;
                padding: 0;
            }

            .toolbar, .theme-toggle, .btn {
                display: none !important;
            }

            .container {
                box-shadow: none;
            }

            .section {
                page-break-inside: avoid;
            }

            .charts-grid {
                page-break-inside: avoid;
            }
        }

        /* Responsive Design */
        @media (max-width: 768px) {
            .header {
                padding: 30px 20px;
            }

            .header h1 {
                font-size: 24px;
            }

            .dashboard {
                grid-template-columns: 1fr;
                padding: 20px;
            }

            .content {
                padding: 20px;
            }

            .section {
                padding: 20px;
            }

            .charts-grid {
                grid-template-columns: 1fr;
            }

            .toolbar {
                padding: 15px;
            }

            .search-box {
                max-width: 100%;
            }

            th, td {
                padding: 10px 8px;
                font-size: 12px;
            }
        }

        /* Utility Classes */
        .text-center { text-align: center; }
        .text-right { text-align: right; }
        .mt-2 { margin-top: 20px; }
        .mb-2 { margin-bottom: 20px; }
        .hidden { display: none; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-chart-line"></i> Microsoft 365 License Report</h1>
            <h2>$OrgName</h2>
            <h3><i class="far fa-clock"></i> Generated: $RunDate</h3>
            <div class="brought-by">
                <i class="fas fa-code"></i> Brought to you by <a href="https://tycholoke.com" target="_blank">Tycho Löke</a> from <a href="https://tycholoke.com" target="_blank">tycholoke.com</a>
                <br>
                <span style="font-size: 13px; margin-top: 8px; display: inline-block;">
                    Want to support Tycho?
                    <a href="https://ko-fi.com/tycholoke" target="_blank" class="kofi-link">
                        <i class="fas fa-heart"></i> Support on Ko-fi
                    </a>
                </span>
            </div>
        </div>

        <div class="toolbar">
            <div class="search-box">
                <i class="fas fa-search" style="color: var(--text-secondary);"></i>
                <input type="text" id="globalSearch" placeholder="Search across all tables...">
            </div>
            <div class="toolbar-buttons">
                <button class="btn btn-primary" onclick="exportToCSV()">
                    <i class="fas fa-download"></i> Export CSV
                </button>
                <button class="btn btn-secondary" onclick="window.print()">
                    <i class="fas fa-print"></i> Print
                </button>
                <button class="theme-toggle" onclick="toggleTheme()" title="Toggle Dark Mode">
                    <i class="fas fa-moon" id="themeIcon"></i>
                </button>
            </div>
        </div>

        <!-- Enhanced Cleanup Filters Section -->
        <div class="toolbar" style="background: linear-gradient(135deg, var(--light-bg) 0%, var(--card-bg) 100%); border-top: none; padding: 15px 30px; box-shadow: inset 0 2px 4px rgba(0,0,0,0.05);">
            <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap;">
                <span style="font-weight: 700; color: var(--text-primary); font-size: 14px; display: flex; align-items: center; gap: 8px;">
                    <i class="fas fa-filter"></i> Smart Cleanup Filters
                    <span style="font-size: 11px; font-weight: 400; color: var(--text-secondary); margin-left: 4px;">(Click to filter accounts)</span>
                </span>
            </div>
            <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-top: 10px;">
                <button class="btn btn-secondary filter-btn" style="font-size: 12px; padding: 7px 14px; transition: all 0.3s ease;" onclick="applyFilter('all')">
                    <i class="fas fa-list"></i> Show All
                </button>
                <button class="btn btn-secondary filter-btn" style="font-size: 12px; padding: 7px 14px; background: rgba(209, 52, 56, 0.15); color: var(--danger-color); border-color: var(--danger-color); transition: all 0.3s ease;" onclick="applyFilter('never')">
                    <i class="fas fa-user-slash"></i> Never Logged In <span style="background: var(--danger-color); color: white; padding: 2px 6px; border-radius: 10px; font-size: 10px; margin-left: 4px;">$($NeverLoggedIn.Count)</span>
                </button>
                <button class="btn btn-secondary filter-btn" style="font-size: 12px; padding: 7px 14px; background: rgba(209, 52, 56, 0.15); color: var(--danger-color); border-color: var(--danger-color); transition: all 0.3s ease;" onclick="applyFilter('cleanup')">
                    <i class="fas fa-exclamation-circle"></i> High Priority <span style="background: var(--danger-color); color: white; padding: 2px 6px; border-radius: 10px; font-size: 10px; margin-left: 4px;">$($HighPriorityCleanup.Count)</span>
                </button>
                <button class="btn btn-secondary filter-btn" style="font-size: 12px; padding: 7px 14px; background: rgba(251, 191, 36, 0.15); color: var(--warning-color); border-color: var(--warning-color); transition: all 0.3s ease;" onclick="applyFilter('30days')">
                    <i class="fas fa-calendar"></i> 30+ Days <span style="background: var(--warning-color); color: white; padding: 2px 6px; border-radius: 10px; font-size: 10px; margin-left: 4px;">$($Inactive30Plus.Count)</span>
                </button>
                <button class="btn btn-secondary filter-btn" style="font-size: 12px; padding: 7px 14px; background: rgba(251, 146, 60, 0.15); color: var(--accent-orange); border-color: var(--accent-orange); transition: all 0.3s ease;" onclick="applyFilter('60days')">
                    <i class="fas fa-clock"></i> 60+ Days <span style="background: var(--accent-orange); color: white; padding: 2px 6px; border-radius: 10px; font-size: 10px; margin-left: 4px;">$($Inactive60Plus.Count)</span>
                </button>
                <button class="btn btn-secondary filter-btn" style="font-size: 12px; padding: 7px 14px; background: rgba(239, 68, 68, 0.15); color: #ef4444; border-color: #ef4444; transition: all 0.3s ease;" onclick="applyFilter('90days')">
                    <i class="fas fa-history"></i> 90+ Days <span style="background: #ef4444; color: white; padding: 2px 6px; border-radius: 10px; font-size: 10px; margin-left: 4px;">$($Inactive90Plus.Count)</span>
                </button>
                <button class="btn btn-secondary filter-btn" style="font-size: 12px; padding: 7px 14px; background: rgba(220, 38, 38, 0.15); color: #dc2626; border-color: #dc2626; transition: all 0.3s ease;" onclick="applyFilter('180days')">
                    <i class="fas fa-ban"></i> 180+ Days <span style="background: #dc2626; color: white; padding: 2px 6px; border-radius: 10px; font-size: 10px; margin-left: 4px;">$($Inactive180Plus.Count)</span>
                </button>
                <button class="btn btn-secondary filter-btn" style="font-size: 12px; padding: 7px 14px; background: rgba(139, 92, 246, 0.15); color: var(--accent-purple); border-color: var(--accent-purple); transition: all 0.3s ease;" onclick="applyFilter('duplicates')">
                    <i class="fas fa-copy"></i> Duplicates <span style="background: var(--accent-purple); color: white; padding: 2px 6px; border-radius: 10px; font-size: 10px; margin-left: 4px;">$DuplicateSKUsAccounts</span>
                </button>
            </div>
            <div id="filterStatus" style="font-size: 13px; color: var(--text-secondary); margin-top: 12px; font-weight: 500; display: flex; align-items: center; gap: 6px;">
                <i class="fas fa-info-circle"></i> <span id="filterStatusText">Showing all <span id="filteredCount" style="font-weight: 700; color: var(--primary-color);">$($Report.Count)</span> accounts</span>
            </div>
        </div>
"@

# Build Dashboard Summary Cards with icons
$DashboardHTML = @"
        <div class="dashboard">
            <div class="stat-card success">
                <div class="label"><i class="fas fa-users"></i> Total Licensed Accounts</div>
                <div class="value">$($Report.Count)</div>
                <div class="subtitle">Active user licenses</div>
            </div>
            <div class="stat-card warning">
                <div class="label"><i class="fas fa-user-clock"></i> Underused Accounts</div>
                <div class="value">$($UnderUsedAccounts.Count)</div>
                <div class="subtitle">$PercentUnderusedAccounts of total</div>
            </div>
            <div class="stat-card danger">
                <div class="label"><i class="fas fa-user-slash"></i> Never Logged In</div>
                <div class="value">$($NeverLoggedIn.Count)</div>
                <div class="subtitle">Immediate cleanup candidates</div>
            </div>
            <div class="stat-card danger">
                <div class="label"><i class="fas fa-exclamation-circle"></i> High Priority Cleanup</div>
                <div class="value">$($HighPriorityCleanup.Count)</div>
                <div class="subtitle">Never used or 90+ days inactive</div>
            </div>
            <div class="stat-card warning">
                <div class="label"><i class="fas fa-calendar"></i> Inactive 30+ Days</div>
                <div class="value">$($Inactive30Plus.Count)</div>
                <div class="subtitle">Requires monitoring</div>
            </div>
            <div class="stat-card warning">
                <div class="label"><i class="fas fa-clock"></i> Inactive 60+ Days</div>
                <div class="value">$($Inactive60Plus.Count)</div>
                <div class="subtitle">Review recommended</div>
            </div>
            <div class="stat-card danger">
                <div class="label"><i class="fas fa-history"></i> Inactive 90+ Days</div>
                <div class="value">$($Inactive90Plus.Count)</div>
                <div class="subtitle">Cleanup candidates</div>
            </div>
            <div class="stat-card danger">
                <div class="label"><i class="fas fa-ban"></i> Inactive 180+ Days</div>
                <div class="value">$($Inactive180Plus.Count)</div>
                <div class="subtitle">Critical - immediate action</div>
            </div>
            <div class="stat-card info">
                <div class="label"><i class="fas fa-exclamation-triangle"></i> Duplicate Licenses</div>
                <div class="value">$DuplicateSKULicenses</div>
                <div class="subtitle">$DuplicateSKUsAccounts accounts affected</div>
            </div>
            <div class="stat-card info">
                <div class="label"><i class="fas fa-bug"></i> License Errors</div>
                <div class="value">$LicenseErrorCount</div>
                <div class="subtitle">Assignment errors</div>
            </div>
"@

# Add pricing cards if available
If ($PricingInfoAvailable) {
  $DashboardHTML += @"
            <div class="stat-card success">
                <div class="label"><i class="fas fa-dollar-sign"></i> Total License Cost</div>
                <div class="value">$TotalBoughtLicenseCostsOutput</div>
                <div class="subtitle">Annual tenant cost</div>
            </div>
            <div class="stat-card info">
                <div class="label"><i class="fas fa-money-bill-wave"></i> Assigned License Cost</div>
                <div class="value">$TotalUserLicenseCostsOutput</div>
                <div class="subtitle">$PercentBoughtLicensesUsed utilized</div>
            </div>
            <div class="stat-card">
                <div class="label"><i class="fas fa-user-tag"></i> Average Cost Per User</div>
                <div class="value">$AverageCostPerUserOutput</div>
                <div class="subtitle">Per licensed account</div>
            </div>
"@
}

$DashboardHTML += @"
        </div>

        <div class="content">
            <!-- Charts Section -->
            <div class="section">
                <div class="section-header">
                    <h2 class="section-title"><i class="fas fa-chart-pie"></i> Visual Analytics Dashboard</h2>
                </div>
                <div class="charts-grid">
                    <div class="chart-container">
                        <div class="chart-title"><i class="fas fa-chart-pie"></i> License Distribution</div>
                        <canvas id="licenseDistributionChart"></canvas>
                    </div>
                    <div class="chart-container">
                        <div class="chart-title"><i class="fas fa-chart-bar"></i> Top 10 Licenses by Usage</div>
                        <canvas id="topLicensesChart"></canvas>
                    </div>
                    <div class="chart-container">
                        <div class="chart-title"><i class="fas fa-exclamation-triangle"></i> Account Status Distribution</div>
                        <canvas id="accountStatusChart"></canvas>
                    </div>
                    <div class="chart-container">
                        <div class="chart-title"><i class="fas fa-clock"></i> Inactive Account Analysis</div>
                        <canvas id="inactiveAccountsChart"></canvas>
                    </div>
"@

# Add cost analysis charts if pricing is available
If ($PricingInfoAvailable) {
  $DashboardHTML += @"
                    <div class="chart-container">
                        <div class="chart-title"><i class="fas fa-building"></i> License Costs by Department</div>
                        <canvas id="departmentCostsChart"></canvas>
                    </div>
                    <div class="chart-container">
                        <div class="chart-title"><i class="fas fa-globe"></i> License Costs by Country</div>
                        <canvas id="countryCostsChart"></canvas>
                    </div>
                    <div class="chart-container">
                        <div class="chart-title"><i class="fas fa-dollar-sign"></i> Cost Utilization Overview</div>
                        <canvas id="costUtilizationChart"></canvas>
                    </div>
                    <div class="chart-container">
                        <div class="chart-title"><i class="fas fa-money-bill-wave"></i> Top 10 Most Expensive Licenses</div>
                        <canvas id="topCostlyLicensesChart"></canvas>
                    </div>
"@
}

$DashboardHTML += @"
                </div>
            </div>
"@

# User Licenses Table with enhanced wrapper
$UserTableHTML = $Report | ConvertTo-Html -Fragment
$UserTableHTML = $UserTableHTML -replace '<table>', '<table id="userTable" class="sortable">'
$UserTableHTML = $UserTableHTML -replace '<th>', '<th class="sortable">'

$HtmlBody1 = @"
            <!-- User License Details Section -->
            <div class="section">
                <div class="section-header">
                    <h2 class="section-title"><i class="fas fa-users"></i> User License Details</h2>
                </div>
                <div class="table-wrapper">
                    <div class="table-controls">
                        <div class="table-info">
                            <span id="userTableCount">$($Report.Count)</span> users found
                        </div>
                        <input type="text" id="userTableSearch" placeholder="Filter users..." style="padding: 8px 12px; border: 2px solid var(--border-color); border-radius: 6px; background: var(--card-bg); color: var(--text-primary);">
                    </div>
                    <div class="table-container">
                        $UserTableHTML
                    </div>
                </div>
            </div>
"@

# SKU Distribution Table with enhanced wrapper
$SkuTableHTML = $SkuReport | Select-Object "SKU Id", "SKU Name", "Units used", "Units purchased", "Annual licensing cost" | ConvertTo-Html -Fragment
$SkuTableHTML = $SkuTableHTML -replace '<table>', '<table id="skuTable" class="sortable">'
$SkuTableHTML = $SkuTableHTML -replace '<th>', '<th class="sortable">'

$HtmlBody2 = @"
            <!-- Product License Distribution Section -->
            <div class="section">
                <div class="section-header">
                    <h2 class="section-title"><i class="fas fa-box"></i> Product License Distribution</h2>
                </div>
                <div class="table-wrapper">
                    <div class="table-controls">
                        <div class="table-info">
                            <span id="skuTableCount">$($SkuReport.Count)</span> products found
                        </div>
                        <input type="text" id="skuTableSearch" placeholder="Filter products..." style="padding: 8px 12px; border: 2px solid var(--border-color); border-radius: 6px; background: var(--card-bg); color: var(--text-primary);">
                    </div>
                    <div class="table-container">
                        $SkuTableHTML
                    </div>
                </div>
            </div>
"@

# Cost Analysis Section with enhanced styling
$HtmlTail = ""

If ($PricingInfoAvailable) {
  # Department Analysis Table
  $DeptTableHTML = $DepartmentHTML -replace '<table>', '<table id="departmentTable" class="sortable">'
  $DeptTableHTML = $DeptTableHTML -replace '<th>', '<th class="sortable">'

  # Country Analysis Table
  $CountryTableHTML = $CountryHTML -replace '<table>', '<table id="countryTable" class="sortable">'
  $CountryTableHTML = $CountryTableHTML -replace '<th>', '<th class="sortable">'

  $HTMLTail = @"
            <!-- Department Cost Analysis Section -->
            <div class="section">
                <div class="section-header">
                    <h2 class="section-title"><i class="fas fa-building"></i> License Costs by Department</h2>
                </div>
                <div class="table-wrapper">
                    <div class="table-controls">
                        <div class="table-info">
                            <span id="departmentTableCount">$($DepartmentReport.Count)</span> departments found
                        </div>
                        <input type="text" id="departmentTableSearch" placeholder="Filter departments..." style="padding: 8px 12px; border: 2px solid var(--border-color); border-radius: 6px; background: var(--card-bg); color: var(--text-primary);">
                    </div>
                    <div class="table-container">
                        $DeptTableHTML
                    </div>
                </div>
                <p class="mt-2" style="color: var(--text-secondary);"><i class="fas fa-info-circle"></i> <strong>Accounts without department:</strong> $NoDepartmentCosts</p>
            </div>

            <!-- Country Cost Analysis Section -->
            <div class="section">
                <div class="section-header">
                    <h2 class="section-title"><i class="fas fa-globe"></i> License Costs by Country</h2>
                </div>
                <div class="table-wrapper">
                    <div class="table-controls">
                        <div class="table-info">
                            <span id="countryTableCount">$($CountryReport.Count)</span> countries found
                        </div>
                        <input type="text" id="countryTableSearch" placeholder="Filter countries..." style="padding: 8px 12px; border: 2px solid var(--border-color); border-radius: 6px; background: var(--card-bg); color: var(--text-primary);">
                    </div>
                    <div class="table-container">
                        $CountryTableHTML
                    </div>
                </div>
                <p class="mt-2" style="color: var(--text-secondary);"><i class="fas fa-info-circle"></i> <strong>Accounts without country:</strong> $NoCountryCosts</p>
            </div>
"@
}

# Prepare chart data for JavaScript
$SkuChartData = $SkuReport | Select-Object -First 10 "SKU Name", "Units used"
$SkuChartLabels = ($SkuChartData | ForEach-Object { """$($_.'SKU Name')""" }) -join ","
$SkuChartValues = ($SkuChartData | ForEach-Object { $_.'Units used' }) -join ","

$AllSkuLabels = ($SkuReport | ForEach-Object { """$($_.'SKU Name')""" }) -join ","
$AllSkuValues = ($SkuReport | ForEach-Object { $_.'Units used' }) -join ","

# Prepare department and country data for charts if pricing available
$DeptChartLabels = ""
$DeptChartValues = ""
$CountryChartLabels = ""
$CountryChartValues = ""
$TopCostlyLicensesLabels = ""
$TopCostlyLicensesValues = ""

If ($PricingInfoAvailable) {
  $DeptChartLabels = ($DepartmentReport | ForEach-Object { """$($_.Department)""" }) -join ","
  $DeptChartValues = ($DepartmentReport | ForEach-Object {
    [float]($_.Costs -replace '[^\d.]', '')
  }) -join ","

  $CountryChartLabels = ($CountryReport | ForEach-Object { """$($_.Country)""" }) -join ","
  $CountryChartValues = ($CountryReport | ForEach-Object {
    [float]($_.Costs -replace '[^\d.]', '')
  }) -join ","

  # Prepare top 10 most expensive licenses data
  $TopCostlyLicenses = $SkuReport | Select-Object -First 10 "SKU Name", "Annual license costs"
  $TopCostlyLicensesLabels = ($TopCostlyLicenses | ForEach-Object { """$($_.'SKU Name')""" }) -join ","
  $TopCostlyLicensesValues = ($TopCostlyLicenses | ForEach-Object { $_.'Annual license costs' }) -join ","
}

# Add comprehensive JavaScript with all features
$ScriptBlock = @"
        </div>
        <div class="footer">
            <p><i class="fas fa-code"></i> Microsoft 365 License Mapper v$Version | Generated: $RunDate</p>
            <p><i class="fas fa-building"></i> Report for: $OrgName</p>
            <p style="margin-top: 15px; font-size: 12px;">
                <i class="fas fa-copyright"></i> Copyright $(Get-Date -Format yyyy) Tycho Löke | All Rights Reserved
            </p>
            <p style="margin-top: 8px; font-size: 11px; opacity: 0.9;">
                Created by <strong>Tycho Löke</strong> |
                <a href="https://tycholoke.com" target="_blank" style="font-weight: 600;">TychoLoke.com</a> |
                <a href="https://currentcloud.net" target="_blank">CurrentCloud.net</a>
            </p>
            <p style="margin-top: 8px; font-size: 10px; opacity: 0.7;">
                <i class="fas fa-info-circle"></i> This tool is provided as-is for administrative purposes.
                Visit <a href="https://tycholoke.com" target="_blank">tycholoke.com</a> for updates and documentation.
            </p>
        </div>
    </div>

    <script>
        // ========================================
        // DARK MODE FUNCTIONALITY
        // ========================================
        function toggleTheme() {
            const html = document.documentElement;
            const themeIcon = document.getElementById('themeIcon');
            const currentTheme = html.getAttribute('data-theme');
            const newTheme = currentTheme === 'dark' ? 'light' : 'dark';

            html.setAttribute('data-theme', newTheme);
            localStorage.setItem('theme', newTheme);

            // Update icon
            if (newTheme === 'dark') {
                themeIcon.className = 'fas fa-sun';
            } else {
                themeIcon.className = 'fas fa-moon';
            }

            // Update charts colors for dark mode
            if (window.chartInstances) {
                Object.values(window.chartInstances).forEach(chart => {
                    if (chart) {
                        updateChartTheme(chart, newTheme);
                    }
                });
            }
        }

        function updateChartTheme(chart, theme) {
            const textColor = theme === 'dark' ? '#e8e8e8' : '#323130';
            const gridColor = theme === 'dark' ? '#404040' : '#edebe9';

            if (chart.options.scales) {
                if (chart.options.scales.x) {
                    chart.options.scales.x.ticks.color = textColor;
                    chart.options.scales.x.grid.color = gridColor;
                }
                if (chart.options.scales.y) {
                    chart.options.scales.y.ticks.color = textColor;
                    chart.options.scales.y.grid.color = gridColor;
                }
            }

            if (chart.options.plugins && chart.options.plugins.legend) {
                chart.options.plugins.legend.labels.color = textColor;
            }

            chart.update();
        }

        // Load saved theme on page load
        document.addEventListener('DOMContentLoaded', function() {
            const savedTheme = localStorage.getItem('theme') || 'light';
            const html = document.documentElement;
            const themeIcon = document.getElementById('themeIcon');

            if (savedTheme === 'dark') {
                html.setAttribute('data-theme', 'dark');
                themeIcon.className = 'fas fa-sun';
            }
        });

        // ========================================
        // CHART.JS INITIALIZATION
        // ========================================
        window.chartInstances = {};

        document.addEventListener('DOMContentLoaded', function() {
            const theme = document.documentElement.getAttribute('data-theme') || 'light';
            const textColor = theme === 'dark' ? '#e8e8e8' : '#323130';
            const gridColor = theme === 'dark' ? '#404040' : '#edebe9';

            // Enhanced vibrant color palette
            const colors = [
                'rgba(96, 165, 250, 0.9)',    // Bright Blue
                'rgba(52, 211, 153, 0.9)',    // Emerald Green
                'rgba(251, 191, 36, 0.9)',    // Amber
                'rgba(248, 113, 113, 0.9)',   // Red
                'rgba(34, 211, 238, 0.9)',    // Cyan
                'rgba(167, 139, 250, 0.9)',   // Purple
                'rgba(251, 146, 60, 0.9)',    // Orange
                'rgba(45, 212, 191, 0.9)',    // Teal
                'rgba(244, 114, 182, 0.9)',   // Pink
                'rgba(132, 204, 22, 0.9)',    // Lime
                'rgba(59, 130, 246, 0.9)',    // Blue
                'rgba(236, 72, 153, 0.9)',    // Hot Pink
                'rgba(139, 92, 246, 0.9)',    // Violet
                'rgba(34, 197, 94, 0.9)',     // Green
                'rgba(249, 115, 22, 0.9)'     // Deep Orange
            ];

            // License Distribution Pie Chart
            const licenseDistCtx = document.getElementById('licenseDistributionChart');
            if (licenseDistCtx) {
                window.chartInstances.licenseDist = new Chart(licenseDistCtx, {
                    type: 'pie',
                    data: {
                        labels: [$AllSkuLabels],
                        datasets: [{
                            data: [$AllSkuValues],
                            backgroundColor: colors,
                            borderWidth: 2,
                            borderColor: theme === 'dark' ? '#2d2d2d' : '#ffffff'
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        plugins: {
                            legend: {
                                position: 'right',
                                labels: {
                                    color: textColor,
                                    boxWidth: 15,
                                    padding: 10,
                                    font: { size: 11 }
                                }
                            },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        const label = context.label || '';
                                        const value = context.parsed || 0;
                                        const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                        const percentage = ((value / total) * 100).toFixed(1);
                                        return label + ': ' + value + ' (' + percentage + '%)';
                                    }
                                }
                            }
                        }
                    }
                });
            }

            // Top 10 Licenses Bar Chart
            const topLicensesCtx = document.getElementById('topLicensesChart');
            if (topLicensesCtx) {
                // Generate gradient colors for bars
                const barColors = [$SkuChartValues].map((_, i) => colors[i % colors.length]);

                window.chartInstances.topLicenses = new Chart(topLicensesCtx, {
                    type: 'bar',
                    data: {
                        labels: [$SkuChartLabels],
                        datasets: [{
                            label: 'Units Used',
                            data: [$SkuChartValues],
                            backgroundColor: barColors,
                            borderColor: barColors.map(c => c.replace('0.9)', '1)')),
                            borderWidth: 2,
                            borderRadius: 8,
                            borderSkipped: false
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        indexAxis: 'y',
                        plugins: {
                            legend: { display: false },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return 'Units: ' + context.parsed.x;
                                    }
                                }
                            }
                        },
                        scales: {
                            x: {
                                beginAtZero: true,
                                ticks: { color: textColor },
                                grid: { color: gridColor }
                            },
                            y: {
                                ticks: {
                                    color: textColor,
                                    font: { size: 10 }
                                },
                                grid: { display: false }
                            }
                        }
                    }
                });
            }

            // Account Status Distribution Chart
            const accountStatusCtx = document.getElementById('accountStatusChart');
            if (accountStatusCtx) {
                window.chartInstances.accountStatus = new Chart(accountStatusCtx, {
                    type: 'doughnut',
                    data: {
                        labels: ['Active (OK)', 'Underused', 'Never Logged In', 'High Priority Cleanup'],
                        datasets: [{
                            data: [
                                $($Report.Count - $UnderUsedAccounts.Count),
                                $($UnderUsedAccounts.Count - $HighPriorityCleanup.Count),
                                $($NeverLoggedIn.Count),
                                $($HighPriorityCleanup.Count - $NeverLoggedIn.Count)
                            ],
                            backgroundColor: [
                                'rgba(52, 211, 153, 0.9)',    // Green for OK
                                'rgba(251, 191, 36, 0.9)',     // Amber for Underused
                                'rgba(248, 113, 113, 0.9)',    // Red for Never Logged In
                                'rgba(251, 146, 60, 0.9)'      // Orange for High Priority
                            ],
                            borderWidth: 3,
                            borderColor: theme === 'dark' ? '#2d2d2d' : '#ffffff'
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        plugins: {
                            legend: {
                                position: 'bottom',
                                labels: {
                                    color: textColor,
                                    boxWidth: 15,
                                    padding: 12,
                                    font: { size: 12 }
                                }
                            },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        const label = context.label || '';
                                        const value = context.parsed || 0;
                                        const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                        const percentage = ((value / total) * 100).toFixed(1);
                                        return label + ': ' + value + ' accounts (' + percentage + '%)';
                                    }
                                }
                            }
                        }
                    }
                });
            }

            // Inactive Account Analysis Chart
            const inactiveAccountsCtx = document.getElementById('inactiveAccountsChart');
            if (inactiveAccountsCtx) {
                window.chartInstances.inactiveAccounts = new Chart(inactiveAccountsCtx, {
                    type: 'bar',
                    data: {
                        labels: ['Never Logged', '30+ Days', '60+ Days', '90+ Days', '180+ Days'],
                        datasets: [{
                            label: 'Number of Accounts',
                            data: [
                                $($NeverLoggedIn.Count),
                                $($Inactive30Plus.Count),
                                $($Inactive60Plus.Count),
                                $($Inactive90Plus.Count),
                                $($Inactive180Plus.Count)
                            ],
                            backgroundColor: [
                                'rgba(248, 113, 113, 0.9)',    // Red
                                'rgba(251, 191, 36, 0.9)',     // Amber
                                'rgba(251, 146, 60, 0.9)',     // Orange
                                'rgba(239, 68, 68, 0.9)',      // Dark Red
                                'rgba(220, 38, 38, 0.9)'       // Darker Red
                            ],
                            borderColor: [
                                'rgba(248, 113, 113, 1)',
                                'rgba(251, 191, 36, 1)',
                                'rgba(251, 146, 60, 1)',
                                'rgba(239, 68, 68, 1)',
                                'rgba(220, 38, 38, 1)'
                            ],
                            borderWidth: 2,
                            borderRadius: 8,
                            borderSkipped: false
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        plugins: {
                            legend: { display: false },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return 'Accounts: ' + context.parsed.y;
                                    }
                                }
                            }
                        },
                        scales: {
                            x: {
                                ticks: { color: textColor },
                                grid: { display: false }
                            },
                            y: {
                                beginAtZero: true,
                                ticks: {
                                    color: textColor,
                                    stepSize: 1
                                },
                                grid: { color: gridColor }
                            }
                        }
                    }
                });
            }

"@</invoke>

# Add department and country charts if pricing is available
If ($PricingInfoAvailable) {
  $ScriptBlock += @"
            // Department Costs Chart
            const deptCostsCtx = document.getElementById('departmentCostsChart');
            if (deptCostsCtx) {
                const deptColors = [$DeptChartValues].map((_, i) => colors[i % colors.length]);

                window.chartInstances.deptCosts = new Chart(deptCostsCtx, {
                    type: 'bar',
                    data: {
                        labels: [$DeptChartLabels],
                        datasets: [{
                            label: 'Annual Costs ($Currency)',
                            data: [$DeptChartValues],
                            backgroundColor: deptColors,
                            borderColor: deptColors.map(c => c.replace('0.9)', '1)')),
                            borderWidth: 2,
                            borderRadius: 8,
                            borderSkipped: false
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        plugins: {
                            legend: { display: true, labels: { color: textColor } },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return '$Currency ' + context.parsed.y.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
                                    }
                                }
                            }
                        },
                        scales: {
                            x: {
                                ticks: { color: textColor },
                                grid: { display: false }
                            },
                            y: {
                                beginAtZero: true,
                                ticks: {
                                    color: textColor,
                                    callback: function(value) {
                                        return '$Currency ' + value.toLocaleString();
                                    }
                                },
                                grid: { color: gridColor }
                            }
                        }
                    }
                });
            }

            // Country Costs Chart
            const countryCostsCtx = document.getElementById('countryCostsChart');
            if (countryCostsCtx) {
                const countryColors = [$CountryChartValues].map((_, i) => colors[i % colors.length]);

                window.chartInstances.countryCosts = new Chart(countryCostsCtx, {
                    type: 'bar',
                    data: {
                        labels: [$CountryChartLabels],
                        datasets: [{
                            label: 'Annual Costs ($Currency)',
                            data: [$CountryChartValues],
                            backgroundColor: countryColors,
                            borderColor: countryColors.map(c => c.replace('0.9)', '1)')),
                            borderWidth: 2,
                            borderRadius: 8,
                            borderSkipped: false
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        plugins: {
                            legend: { display: true, labels: { color: textColor } },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return '$Currency ' + context.parsed.y.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
                                    }
                                }
                            }
                        },
                        scales: {
                            x: {
                                ticks: { color: textColor },
                                grid: { display: false }
                            },
                            y: {
                                beginAtZero: true,
                                ticks: {
                                    color: textColor,
                                    callback: function(value) {
                                        return '$Currency ' + value.toLocaleString();
                                    }
                                },
                                grid: { color: gridColor }
                            }
                        }
                    }
                });
            }

            // Cost Utilization Overview Chart (Doughnut)
            const costUtilizationCtx = document.getElementById('costUtilizationChart');
            if (costUtilizationCtx) {
                const totalCost = $TotalBoughtLicenseCosts;
                const assignedCost = $TotalUserLicenseCosts;
                const unusedCost = totalCost - assignedCost;

                window.chartInstances.costUtilization = new Chart(costUtilizationCtx, {
                    type: 'doughnut',
                    data: {
                        labels: ['Assigned Licenses', 'Unused Capacity'],
                        datasets: [{
                            data: [assignedCost, unusedCost],
                            backgroundColor: [
                                'rgba(52, 211, 153, 0.9)',    // Green for Assigned
                                'rgba(248, 113, 113, 0.9)'     // Red for Unused
                            ],
                            borderWidth: 3,
                            borderColor: theme === 'dark' ? '#2d2d2d' : '#ffffff'
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        plugins: {
                            legend: {
                                position: 'bottom',
                                labels: {
                                    color: textColor,
                                    boxWidth: 15,
                                    padding: 12,
                                    font: { size: 12 }
                                }
                            },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        const label = context.label || '';
                                        const value = context.parsed || 0;
                                        const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                        const percentage = ((value / total) * 100).toFixed(1);
                                        return label + ': $Currency ' + value.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + ' (' + percentage + '%)';
                                    }
                                }
                            }
                        }
                    }
                });
            }

            // Top 10 Most Expensive Licenses Chart
            const topCostlyLicensesCtx = document.getElementById('topCostlyLicensesChart');
            if (topCostlyLicensesCtx) {
                const costlyLicenseColors = [$TopCostlyLicensesValues].map((_, i) => colors[i % colors.length]);

                window.chartInstances.topCostlyLicenses = new Chart(topCostlyLicensesCtx, {
                    type: 'bar',
                    data: {
                        labels: [$TopCostlyLicensesLabels],
                        datasets: [{
                            label: 'Annual Cost ($Currency)',
                            data: [$TopCostlyLicensesValues],
                            backgroundColor: costlyLicenseColors,
                            borderColor: costlyLicenseColors.map(c => c.replace('0.9)', '1)')),
                            borderWidth: 2,
                            borderRadius: 8,
                            borderSkipped: false
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: true,
                        indexAxis: 'y',
                        plugins: {
                            legend: { display: false },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return '$Currency ' + context.parsed.x.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
                                    }
                                }
                            }
                        },
                        scales: {
                            x: {
                                beginAtZero: true,
                                ticks: {
                                    color: textColor,
                                    callback: function(value) {
                                        return '$Currency ' + value.toLocaleString();
                                    }
                                },
                                grid: { color: gridColor }
                            },
                            y: {
                                ticks: {
                                    color: textColor,
                                    font: { size: 10 }
                                },
                                grid: { display: false }
                            }
                        }
                    }
                });
            }

"@
}

$ScriptBlock += @"
        });

        // ========================================
        // TABLE SORTING FUNCTIONALITY
        // ========================================
        document.addEventListener('DOMContentLoaded', function() {
            const tables = document.querySelectorAll('table.sortable');

            tables.forEach(table => {
                const headers = table.querySelectorAll('th.sortable');

                headers.forEach((header, index) => {
                    header.addEventListener('click', () => {
                        sortTable(table, index, header);
                    });
                });
            });
        });

        function sortTable(table, columnIndex, header) {
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr:not(.hidden)'));
            const allRows = Array.from(tbody.querySelectorAll('tr'));

            // Get current sort state
            const currentOrder = header.classList.contains('sorted-asc') ? 'asc' :
                                header.classList.contains('sorted-desc') ? 'desc' : 'none';

            // Remove all sort indicators
            table.querySelectorAll('th').forEach(th => {
                th.classList.remove('sorted-asc', 'sorted-desc');
            });

            // Determine new sort order
            let newOrder = 'asc';
            if (currentOrder === 'asc') {
                newOrder = 'desc';
            }

            // Add sort indicator
            header.classList.add('sorted-' + newOrder);

            // Sort rows
            rows.sort((a, b) => {
                const aVal = a.cells[columnIndex].textContent.trim();
                const bVal = b.cells[columnIndex].textContent.trim();

                // Try to parse as number
                const aNum = parseFloat(aVal.replace(/[^0-9.-]/g, ''));
                const bNum = parseFloat(bVal.replace(/[^0-9.-]/g, ''));

                let comparison = 0;
                if (!isNaN(aNum) && !isNaN(bNum)) {
                    comparison = aNum - bNum;
                } else {
                    comparison = aVal.localeCompare(bVal, undefined, {numeric: true, sensitivity: 'base'});
                }

                return newOrder === 'asc' ? comparison : -comparison;
            });

            // Reorder rows in DOM
            rows.forEach(row => tbody.appendChild(row));
        }

        // ========================================
        // SEARCH/FILTER FUNCTIONALITY
        // ========================================
        document.addEventListener('DOMContentLoaded', function() {
            // Global search
            const globalSearch = document.getElementById('globalSearch');
            if (globalSearch) {
                globalSearch.addEventListener('input', function(e) {
                    const searchTerm = e.target.value.toLowerCase();
                    const tables = document.querySelectorAll('table');

                    tables.forEach(table => {
                        filterTable(table, searchTerm);
                    });
                });
            }

            // Individual table searches
            setupTableSearch('userTableSearch', 'userTable', 'userTableCount');
            setupTableSearch('skuTableSearch', 'skuTable', 'skuTableCount');
            setupTableSearch('departmentTableSearch', 'departmentTable', 'departmentTableCount');
            setupTableSearch('countryTableSearch', 'countryTable', 'countryTableCount');
        });

        function setupTableSearch(searchId, tableId, countId) {
            const searchInput = document.getElementById(searchId);
            const table = document.getElementById(tableId);
            const countSpan = document.getElementById(countId);

            if (searchInput && table) {
                searchInput.addEventListener('input', function(e) {
                    const searchTerm = e.target.value.toLowerCase();
                    const visibleCount = filterTable(table, searchTerm);

                    if (countSpan) {
                        countSpan.textContent = visibleCount;
                    }
                });
            }
        }

        function filterTable(table, searchTerm) {
            if (!table) return 0;

            const tbody = table.querySelector('tbody');
            if (!tbody) return 0;

            const rows = tbody.querySelectorAll('tr');
            let visibleCount = 0;

            rows.forEach(row => {
                const text = row.textContent.toLowerCase();
                if (text.includes(searchTerm)) {
                    row.classList.remove('hidden');
                    visibleCount++;
                } else {
                    row.classList.add('hidden');
                }
            });

            return visibleCount;
        }

        // ========================================
        // QUICK FILTER FUNCTIONALITY
        // ========================================
        function applyFilter(filterType) {
            const userTable = document.getElementById('userTable');
            if (!userTable) return;

            const tbody = userTable.querySelector('tbody');
            const thead = userTable.querySelector('thead');
            const rows = tbody.querySelectorAll('tr');
            let visibleCount = 0;
            let filterDescription = '';

            // Find column indices by header text
            const headers = Array.from(thead.querySelectorAll('th')).map(th => th.textContent.trim().replace(/[↕↑↓]/g, '').trim());
            const statusColIndex = headers.indexOf('Status');
            const duplicatesColIndex = headers.indexOf('Duplicates detected');

            // Verify we found the columns
            if (statusColIndex === -1) {
                console.error('Status column not found in table headers');
                return;
            }

            // Clear global search when applying a filter
            const globalSearch = document.getElementById('globalSearch');
            if (globalSearch) globalSearch.value = '';

            rows.forEach(row => {
                const cells = row.querySelectorAll('td');
                let showRow = false;

                // Get the Status and Duplicates columns using the found indices
                const statusCell = cells[statusColIndex];
                const duplicatesCell = duplicatesColIndex !== -1 ? cells[duplicatesColIndex] : null;
                const statusText = statusCell ? statusCell.textContent.trim() : '';
                const duplicatesText = duplicatesCell ? duplicatesCell.textContent.trim() : '';

                switch(filterType) {
                    case 'all':
                        showRow = true;
                        filterDescription = 'all';
                        break;
                    case 'never':
                        showRow = statusText.includes('Never logged in');
                        filterDescription = 'never logged in accounts';
                        break;
                    case 'cleanup':
                        showRow = statusText.includes('Cleanup candidate') ||
                                  statusText.includes('High priority cleanup') ||
                                  statusText.includes('Never logged in');
                        filterDescription = 'high priority cleanup candidates';
                        break;
                    case '30days':
                        showRow = statusText.includes('30+ days') ||
                                  statusText.includes('60+ days') ||
                                  statusText.includes('90+ days') ||
                                  statusText.includes('180+ days');
                        filterDescription = 'accounts inactive 30+ days';
                        break;
                    case '60days':
                        showRow = statusText.includes('60+ days') ||
                                  statusText.includes('90+ days') ||
                                  statusText.includes('180+ days');
                        filterDescription = 'accounts inactive 60+ days';
                        break;
                    case '90days':
                        showRow = statusText.includes('90+ days') ||
                                  statusText.includes('180+ days');
                        filterDescription = 'accounts inactive 90+ days';
                        break;
                    case '180days':
                        showRow = statusText.includes('180+ days');
                        filterDescription = 'accounts inactive 180+ days';
                        break;
                    case 'duplicates':
                        showRow = duplicatesText.includes('Warning: Duplicate');
                        filterDescription = 'accounts with duplicate licenses';
                        break;
                }

                if (showRow) {
                    row.classList.remove('hidden');
                    visibleCount++;
                } else {
                    row.classList.add('hidden');
                }
            });

            // Update filter status with enhanced styling
            const filterStatus = document.getElementById('filterStatus');
            const filterStatusText = document.getElementById('filterStatusText');
            const filteredCount = document.getElementById('filteredCount');

            if (filteredCount) {
                filteredCount.textContent = visibleCount;
            }

            if (filterStatusText) {
                if (filterType === 'all') {
                    filterStatusText.innerHTML = 'Showing all <span id="filteredCount" style="font-weight: 700; color: var(--primary-color);">' + visibleCount + '</span> accounts';
                    if (filterStatus) {
                        filterStatus.style.color = 'var(--text-secondary)';
                    }
                } else {
                    filterStatusText.innerHTML = '<strong style="color: var(--primary-color);">Filter Active:</strong> Showing <span id="filteredCount" style="font-weight: 700; color: var(--primary-color);">' + visibleCount + '</span> ' + filterDescription;
                    if (filterStatus) {
                        filterStatus.style.color = 'var(--primary-color)';
                    }
                }
            }

            // Update user table count
            const userTableCount = document.getElementById('userTableCount');
            if (userTableCount) {
                userTableCount.textContent = visibleCount;
            }

            // Highlight active filter button
            document.querySelectorAll('.filter-btn').forEach(btn => {
                btn.style.transform = 'scale(1)';
                btn.style.boxShadow = 'none';
            });
            if (filterType !== 'all') {
                const activeBtn = event?.target?.closest('button');
                if (activeBtn) {
                    activeBtn.style.transform = 'scale(1.05)';
                    activeBtn.style.boxShadow = '0 4px 8px rgba(0,0,0,0.2)';
                }
            }
        }

        // ========================================
        // EXPORT TO CSV FUNCTIONALITY
        // ========================================
        function exportToCSV() {
            const tables = document.querySelectorAll('table');
            let csvContent = 'Microsoft 365 License Report - $OrgName\n';
            csvContent += 'Generated: $RunDate\n\n';

            tables.forEach((table, index) => {
                // Get table title from section
                const section = table.closest('.section');
                const title = section ? section.querySelector('.section-title')?.textContent.trim() || 'Table ' + (index + 1) : 'Table ' + (index + 1);

                csvContent += title + '\n';

                // Get headers
                const headers = Array.from(table.querySelectorAll('thead th')).map(th => {
                    return '"' + th.textContent.trim().replace(/[↕↑↓]/g, '').trim() + '"';
                });
                csvContent += headers.join(',') + '\n';

                // Get visible rows only
                const rows = table.querySelectorAll('tbody tr:not(.hidden)');
                rows.forEach(row => {
                    const cells = Array.from(row.querySelectorAll('td')).map(td => {
                        return '"' + td.textContent.trim().replace(/"/g, '""') + '"';
                    });
                    csvContent += cells.join(',') + '\n';
                });

                csvContent += '\n';
            });

            // Create download link
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            const url = URL.createObjectURL(blob);

            link.setAttribute('href', url);
            link.setAttribute('download', 'M365_License_Report_' + new Date().toISOString().split('T')[0] + '.csv');
            link.style.visibility = 'hidden';

            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    </script>
</body>
</html>
"@

# Assemble the complete HTML report
$HtmlReport = $HtmlHead + $DashboardHTML + $HtmlBody1 + $HtmlBody2 + $HtmlTail + $ScriptBlock
$HtmlReport | Out-File $HtmlReportFile -Encoding UTF8

Write-Host "Professional HTML report with advanced features generated successfully!" -ForegroundColor Green
Write-Host "  Features: Dark Mode, Interactive Charts, Advanced Search, Export to CSV" -ForegroundColor Cyan

#endregion



$Report | Export-CSV -NoTypeInformation $CSVOutputFile

# Display completion summary
Write-Host ""
Write-Host "===============================================" -ForegroundColor Green
Write-Host "   License Report Generation Complete!" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Output Files:" -ForegroundColor Cyan
Write-Host "  CSV Report:  $CSVOutputFile" -ForegroundColor White
Write-Host "  HTML Report: $HtmlReportFile" -ForegroundColor White
Write-Host ""
Write-Host "Report Summary:" -ForegroundColor Cyan
Write-Host "  Licensed Accounts:     $($Report.Count)" -ForegroundColor White
Write-Host "  Underused Accounts:    $($UnderUsedAccounts.Count)" -ForegroundColor White
Write-Host "  Duplicate Licenses:    $DuplicateSKULicenses" -ForegroundColor White
Write-Host "  License Errors:        $LicenseErrorCount" -ForegroundColor White
Write-Host ""
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
Disconnect-MgGraph
Write-Host "Session ended successfully." -ForegroundColor Green
Write-Host ""

<#
    DISCLAIMER:
    This script is provided as-is without warranty of any kind. Always test in a non-production
    environment before deploying to production. The author and contributors are not responsible
    for any data loss, service disruption, or issues arising from the use of this script.

    Never run scripts downloaded from the Internet without first validating the code and
    understanding its functionality. Review and customize this script to meet your organization's
    specific needs and security requirements.
#>
