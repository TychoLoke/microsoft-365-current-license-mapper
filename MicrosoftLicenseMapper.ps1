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
    Author: Tycho Loke
    Website: https://currentcloud.net
    Blog: https://tycholoke.com
    Version: 2.0
    Updated: 05/01/2026

    Requires:
    - PowerShell 7.0 or higher
    - Microsoft.Graph PowerShell module
    - Appropriate Microsoft 365 admin permissions
    - Pre-generated SKU and Service Plan CSV files (run MicrosoftLicenseMapperCSV.ps1 first)

    .LINK
    https://github.com/TychoLoke/microsoft-365-current-license-mapper
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
$Version = "2.0"

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
      $DaysSinceLastSignIn = "Unknown"
      $UnusedAccountWarning = ("Unknown last sign-in for account")
      $LastAccess = "Unknown"
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
    If ($DaysSinceLastSignIn -gt 60) { 
      $UnusedAccountWarning = ("Account unused for {0} days - check!" -f $DaysSinceLastSignIn) 
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
            --success-color: #107c10;
            --warning-color: #ffb900;
            --danger-color: #d13438;
            --info-color: #00b7c3;
            --dark-bg: #1e1e1e;
            --light-bg: #f5f5f5;
            --card-bg: #ffffff;
            --text-primary: #323130;
            --text-secondary: #605e5c;
            --border-color: #edebe9;
            --shadow: 0 2px 8px rgba(0,0,0,0.1);
            --header-gradient-start: #0078d4;
            --header-gradient-end: #005a9e;
        }

        [data-theme="dark"] {
            --primary-color: #4da6ff;
            --success-color: #6cc24a;
            --warning-color: #ffdd87;
            --danger-color: #ff6b6b;
            --info-color: #4dd4e1;
            --dark-bg: #121212;
            --light-bg: #1e1e1e;
            --card-bg: #2d2d2d;
            --text-primary: #e8e8e8;
            --text-secondary: #b3b3b3;
            --border-color: #404040;
            --shadow: 0 2px 12px rgba(0,0,0,0.5);
            --header-gradient-start: #1a5490;
            --header-gradient-end: #0d3a6e;
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

        /* Header Styles */
        .header {
            background: linear-gradient(135deg, var(--header-gradient-start) 0%, var(--header-gradient-end) 100%);
            color: white;
            padding: 40px;
            text-align: center;
            position: relative;
        }

        .header h1 {
            font-size: 36px;
            font-weight: 300;
            margin-bottom: 10px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }

        .header h2 {
            font-size: 22px;
            font-weight: 400;
            opacity: 0.95;
            margin-bottom: 8px;
        }

        .header h3 {
            font-size: 14px;
            font-weight: 300;
            opacity: 0.85;
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
            border-radius: 12px;
            padding: 24px;
            box-shadow: var(--shadow);
            border-left: 4px solid var(--primary-color);
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }

        .stat-card::before {
            content: '';
            position: absolute;
            top: 0;
            right: 0;
            width: 100px;
            height: 100px;
            opacity: 0.05;
            font-size: 80px;
            font-family: 'Font Awesome 6 Free';
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .stat-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 8px 16px rgba(0,0,0,0.15);
        }

        .stat-card.success { border-left-color: var(--success-color); }
        .stat-card.warning { border-left-color: var(--warning-color); }
        .stat-card.danger { border-left-color: var(--danger-color); }
        .stat-card.info { border-left-color: var(--info-color); }

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

        tbody tr:hover {
            background: var(--light-bg);
            transform: scale(1.01);
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
                    <h2 class="section-title"><i class="fas fa-chart-pie"></i> Visual Analytics</h2>
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

If ($PricingInfoAvailable) {
  $DeptChartLabels = ($DepartmentReport | ForEach-Object { """$($_.Department)""" }) -join ","
  $DeptChartValues = ($DepartmentReport | ForEach-Object {
    [float]($_.Costs -replace '[^\d.]', '')
  }) -join ","

  $CountryChartLabels = ($CountryReport | ForEach-Object { """$($_.Country)""" }) -join ","
  $CountryChartValues = ($CountryReport | ForEach-Object {
    [float]($_.Costs -replace '[^\d.]', '')
  }) -join ","
}

# Add comprehensive JavaScript with all features
$ScriptBlock = @"
        </div>
        <div class="footer">
            <p><i class="fas fa-code"></i> Microsoft 365 License Mapper v$Version | Generated: $RunDate</p>
            <p><i class="fas fa-building"></i> Report for: $OrgName</p>
            <p style="margin-top: 10px; font-size: 11px; opacity: 0.8;">
                <a href="https://currentcloud.net" target="_blank">CurrentCloud.net</a> |
                <a href="https://tycholoke.com" target="_blank">TychoLoke.com</a>
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

            // Chart color palette
            const colors = [
                'rgba(0, 120, 212, 0.8)',
                'rgba(16, 124, 16, 0.8)',
                'rgba(255, 185, 0, 0.8)',
                'rgba(209, 52, 56, 0.8)',
                'rgba(0, 183, 195, 0.8)',
                'rgba(138, 43, 226, 0.8)',
                'rgba(255, 140, 0, 0.8)',
                'rgba(0, 128, 128, 0.8)',
                'rgba(255, 20, 147, 0.8)',
                'rgba(50, 205, 50, 0.8)'
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
                window.chartInstances.topLicenses = new Chart(topLicensesCtx, {
                    type: 'bar',
                    data: {
                        labels: [$SkuChartLabels],
                        datasets: [{
                            label: 'Units Used',
                            data: [$SkuChartValues],
                            backgroundColor: 'rgba(0, 120, 212, 0.8)',
                            borderColor: 'rgba(0, 120, 212, 1)',
                            borderWidth: 2
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

"@

# Add department and country charts if pricing is available
If ($PricingInfoAvailable) {
  $ScriptBlock += @"
            // Department Costs Chart
            const deptCostsCtx = document.getElementById('departmentCostsChart');
            if (deptCostsCtx) {
                window.chartInstances.deptCosts = new Chart(deptCostsCtx, {
                    type: 'bar',
                    data: {
                        labels: [$DeptChartLabels],
                        datasets: [{
                            label: 'Annual Costs ($Currency)',
                            data: [$DeptChartValues],
                            backgroundColor: 'rgba(16, 124, 16, 0.8)',
                            borderColor: 'rgba(16, 124, 16, 1)',
                            borderWidth: 2
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
                window.chartInstances.countryCosts = new Chart(countryCostsCtx, {
                    type: 'bar',
                    data: {
                        labels: [$CountryChartLabels],
                        datasets: [{
                            label: 'Annual Costs ($Currency)',
                            data: [$CountryChartValues],
                            backgroundColor: 'rgba(0, 183, 195, 0.8)',
                            borderColor: 'rgba(0, 183, 195, 1)',
                            borderWidth: 2
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
