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
    Version: 1.8
    Updated: 27/03/2024

    Requires:
    - Microsoft.Graph PowerShell module
    - Appropriate Microsoft 365 admin permissions
    - Pre-generated SKU and Service Plan CSV files (run MicrosoftLicenseMapperCSV.ps1 first)

    .LINK
    https://github.com/TychoLoke/microsoft-365-current-license-mapper
#>


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
$Version = "1.8"

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
# Create the HTML report
$HtmlHead = "<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Microsoft 365 License Report</h1></p>
           <p><h2><b>For the " + $Orgname + " tenant</b></h2></p>
           <p><h3>Generated: " + $RunDate + "</h3></p></div>"

$HtmlBody1 = $Report | ConvertTo-Html -Fragment
$HtmlBody1 = $HTMLBody1 + "<p>Report created for: " + $OrgName + "</p>" +
"<p>Created: " + $RunDate + "<p>" +
"<p>-----------------------------------------------------------------------------------------------------------------------------</p>" +  
"<p>Number of licensed user accounts found:    " + $Report.Count + "</p>" +
"<p>Number of underused user accounts found:   " + $UnderUsedAccounts.Count + "</p>" +
"<p>Percent underused user accounts:           " + $PercentUnderusedAccounts + "</p>" +
"<p>Accounts detected with duplicate licenses: " + $DuplicateSKUsAccounts + "</p>" +
"<p>Count of duplicate licenses:               " + $DuplicateSKULicenses + "</p>" +
"<p>Count of errors:                           " + $LicenseErrorCount + "</p>" +
"<p>-----------------------------------------------------------------------------------------------------------------------------</p>"


$HtmlBody2 = $SkuReport | Select-Object "SKU Id", "SKU Name", "Units used", "Units purchased", "Annual licensing cost" | ConvertTo-Html -Fragment
$HtmlSkuSeparator = "<p><h2>Product License Distribution</h2></p>"

$HtmlTail = "<p></p>"
# Add Cost analysis if pricing information is available

If ($PricingInfoAvailable) {
  $HTMLTail = $HTMLTail + "<h2>Licensing Cost Analysis</h2>" +
  "<p>Total licensing cost for tenant:             " + $TotalBoughtLicenseCostsOutput + "</p>" +
  "<p>Total cost for assigned licenses:            " + $TotalUserLicenseCostsOutput + "</p>" +
  "<p>Percent bought licenses assigned to users:   " + $PercentBoughtLicensesUsed + "</p>" +
  "<p>Average licensing cost per user:             " + $AverageCostPerUserOutput + "</p>" +
  "<p><h2>License Costs by Country</h2></p>" + $CountryHTML +
  "<p>License costs for users without a country:   " + $NoCountryCosts +
  "<p><h2>License Costs by Department</h2></p>" + $DepartmentHTML +
  "<p>License costs for users without a department: " + $NoDepartmentCosts
}
          
$HTMLTail = $HTMLTail + "<p>Microsoft 365 Licensing Report<b> " + $Version + "</b></p>"	

$HtmlReport = $Htmlhead + $Htmlbody1 + $HtmlSkuSeparator + $HtmlBody2 + $Htmltail
$HtmlReport | Out-File $HtmlReportFile -Encoding UTF8



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
