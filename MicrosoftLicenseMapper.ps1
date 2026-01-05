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
#region Generate Modern HTML Report

# Create the HTML report with modern styling
$HtmlHead = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 License Report - $OrgName</title>
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
        }

        body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif;
            font-size: 14px;
            line-height: 1.6;
            color: var(--text-primary);
            background: var(--light-bg);
            padding: 20px;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 8px;
            box-shadow: var(--shadow);
            overflow: hidden;
        }

        /* Header Styles */
        .header {
            background: linear-gradient(135deg, #0078d4 0%, #005a9e 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }

        .header h1 {
            font-size: 32px;
            font-weight: 300;
            margin-bottom: 10px;
        }

        .header h2 {
            font-size: 20px;
            font-weight: 400;
            opacity: 0.95;
            margin-bottom: 8px;
        }

        .header h3 {
            font-size: 14px;
            font-weight: 300;
            opacity: 0.85;
        }

        /* Dashboard Cards */
        .dashboard {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            padding: 30px;
            background: var(--light-bg);
        }

        .stat-card {
            background: var(--card-bg);
            border-radius: 8px;
            padding: 24px;
            box-shadow: var(--shadow);
            border-left: 4px solid var(--primary-color);
            transition: transform 0.2s, box-shadow 0.2s;
        }

        .stat-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
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
        }

        .stat-card .value {
            font-size: 32px;
            font-weight: 300;
            color: var(--text-primary);
            line-height: 1.2;
        }

        .stat-card .subtitle {
            font-size: 13px;
            color: var(--text-secondary);
            margin-top: 8px;
        }

        /* Content Sections */
        .content {
            padding: 30px;
        }

        .section {
            margin-bottom: 40px;
        }

        .section-title {
            font-size: 24px;
            font-weight: 400;
            color: var(--text-primary);
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid var(--primary-color);
        }

        /* Modern Table Styles */
        .table-container {
            overflow-x: auto;
            border-radius: 8px;
            box-shadow: var(--shadow);
            margin-bottom: 20px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            font-size: 13px;
        }

        thead {
            background: linear-gradient(to bottom, #f8f8f8, #f0f0f0);
            position: sticky;
            top: 0;
            z-index: 10;
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
        }

        th:hover {
            background: #e8e8e8;
        }

        th::after {
            content: ' ↕';
            opacity: 0.3;
            font-size: 10px;
        }

        tbody tr {
            border-bottom: 1px solid var(--border-color);
            transition: background 0.2s;
        }

        tbody tr:hover {
            background: #f9f9f9;
        }

        tbody tr:nth-child(even) {
            background: #fafafa;
        }

        tbody tr:nth-child(even):hover {
            background: #f5f5f5;
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
            background: #e7f4e7;
            color: var(--success-color);
        }

        .badge-warning {
            background: #fff4ce;
            color: #8a6d00;
        }

        .badge-danger {
            background: #fde7e9;
            color: var(--danger-color);
        }

        .badge-info {
            background: #cef0f5;
            color: #006f7a;
        }

        /* Footer */
        .footer {
            background: var(--dark-bg);
            color: white;
            padding: 20px;
            text-align: center;
            font-size: 12px;
        }

        .footer a {
            color: var(--info-color);
            text-decoration: none;
        }

        /* Print Styles */
        @media print {
            body {
                background: white;
                padding: 0;
            }

            .container {
                box-shadow: none;
            }

            .stat-card:hover,
            tbody tr:hover,
            th:hover {
                transform: none;
                background: initial;
            }

            .table-container {
                box-shadow: none;
            }

            thead {
                position: static;
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

            .table-container {
                font-size: 12px;
            }

            th, td {
                padding: 10px 8px;
            }
        }

        /* Utility Classes */
        .text-center { text-align: center; }
        .text-right { text-align: right; }
        .mt-2 { margin-top: 20px; }
        .mb-2 { margin-bottom: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Microsoft 365 License Report</h1>
            <h2>$OrgName</h2>
            <h3>Generated: $RunDate</h3>
        </div>
"@

# Build Dashboard Summary Cards
$DashboardHTML = @"
        <div class="dashboard">
            <div class="stat-card success">
                <div class="label">Total Licensed Accounts</div>
                <div class="value">$($Report.Count)</div>
                <div class="subtitle">Active user licenses</div>
            </div>
            <div class="stat-card warning">
                <div class="label">Underused Accounts</div>
                <div class="value">$($UnderUsedAccounts.Count)</div>
                <div class="subtitle">$PercentUnderusedAccounts of total</div>
            </div>
            <div class="stat-card danger">
                <div class="label">Duplicate Licenses</div>
                <div class="value">$DuplicateSKULicenses</div>
                <div class="subtitle">$DuplicateSKUsAccounts accounts affected</div>
            </div>
            <div class="stat-card info">
                <div class="label">License Errors</div>
                <div class="value">$LicenseErrorCount</div>
                <div class="subtitle">Assignment errors</div>
            </div>
"@

# Add pricing cards if available
If ($PricingInfoAvailable) {
  $DashboardHTML += @"
            <div class="stat-card success">
                <div class="label">Total License Cost</div>
                <div class="value">$TotalBoughtLicenseCostsOutput</div>
                <div class="subtitle">Annual tenant cost</div>
            </div>
            <div class="stat-card info">
                <div class="label">Assigned License Cost</div>
                <div class="value">$TotalUserLicenseCostsOutput</div>
                <div class="subtitle">$PercentBoughtLicensesUsed utilized</div>
            </div>
            <div class="stat-card">
                <div class="label">Average Cost Per User</div>
                <div class="value">$AverageCostPerUserOutput</div>
                <div class="subtitle">Per licensed account</div>
            </div>
"@
}

$DashboardHTML += @"
        </div>
"@

# User Licenses Table
$UserTableHTML = $Report | ConvertTo-Html -Fragment
$UserTableHTML = $UserTableHTML -replace '<table>', '<div class="table-container"><table>'
$UserTableHTML = $UserTableHTML -replace '</table>', '</table></div>'

$HtmlBody1 = @"
        <div class="content">
            <div class="section">
                <h2 class="section-title">User License Details</h2>
                $UserTableHTML
            </div>
"@

# SKU Distribution Table
$SkuTableHTML = $SkuReport | Select-Object "SKU Id", "SKU Name", "Units used", "Units purchased", "Annual licensing cost" | ConvertTo-Html -Fragment
$SkuTableHTML = $SkuTableHTML -replace '<table>', '<div class="table-container"><table>'
$SkuTableHTML = $SkuTableHTML -replace '</table>', '</table></div>'

$HtmlBody2 = @"
            <div class="section">
                <h2 class="section-title">Product License Distribution</h2>
                $SkuTableHTML
            </div>
"@

# Cost Analysis Section
$HtmlTail = ""

If ($PricingInfoAvailable) {
  # Department Analysis Table
  $DeptTableHTML = $DepartmentHTML -replace '<table>', '<div class="table-container"><table>'
  $DeptTableHTML = $DeptTableHTML -replace '</table>', '</table></div>'

  # Country Analysis Table
  $CountryTableHTML = $CountryHTML -replace '<table>', '<div class="table-container"><table>'
  $CountryTableHTML = $CountryTableHTML -replace '</table>', '</table></div>'

  $HTMLTail = @"
            <div class="section">
                <h2 class="section-title">License Costs by Department</h2>
                $DeptTableHTML
                <p class="mt-2"><strong>Accounts without department:</strong> $NoDepartmentCosts</p>
            </div>

            <div class="section">
                <h2 class="section-title">License Costs by Country</h2>
                $CountryTableHTML
                <p class="mt-2"><strong>Accounts without country:</strong> $NoCountryCosts</p>
            </div>
"@
}

# Add JavaScript for table sorting
$ScriptBlock = @"
        </div>
        <div class="footer">
            <p>Microsoft 365 License Mapper v$Version | Generated: $RunDate</p>
            <p>Report for: $OrgName</p>
        </div>
    </div>

    <script>
        // Simple table sorting functionality
        document.addEventListener('DOMContentLoaded', function() {
            const tables = document.querySelectorAll('table');

            tables.forEach(table => {
                const headers = table.querySelectorAll('th');

                headers.forEach((header, index) => {
                    header.addEventListener('click', () => {
                        sortTable(table, index);
                    });
                });
            });
        });

        function sortTable(table, column) {
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));

            const currentSort = table.dataset.sortColumn;
            const currentOrder = table.dataset.sortOrder || 'asc';

            let newOrder = 'asc';
            if (currentSort === column.toString() && currentOrder === 'asc') {
                newOrder = 'desc';
            }

            rows.sort((a, b) => {
                const aVal = a.cells[column].textContent.trim();
                const bVal = b.cells[column].textContent.trim();

                const aNum = parseFloat(aVal.replace(/[^0-9.-]/g, ''));
                const bNum = parseFloat(bVal.replace(/[^0-9.-]/g, ''));

                let comparison = 0;
                if (!isNaN(aNum) && !isNaN(bNum)) {
                    comparison = aNum - bNum;
                } else {
                    comparison = aVal.localeCompare(bVal);
                }

                return newOrder === 'asc' ? comparison : -comparison;
            });

            rows.forEach(row => tbody.appendChild(row));

            table.dataset.sortColumn = column;
            table.dataset.sortOrder = newOrder;
        }
    </script>
</body>
</html>
"@

# Assemble the complete HTML report
$HtmlReport = $HtmlHead + $DashboardHTML + $HtmlBody1 + $HtmlBody2 + $HtmlTail + $ScriptBlock
$HtmlReport | Out-File $HtmlReportFile -Encoding UTF8

Write-Host "Modern HTML report generated successfully!" -ForegroundColor Green

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
