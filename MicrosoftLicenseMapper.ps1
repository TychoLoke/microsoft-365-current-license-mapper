<#
    .SYNOPSIS
    Script to report on licenses assigned to Azure AD user accounts using Microsoft Graph PowerShell SDK cmdlets.
    For MicrosoftLicenseMapper.ps1.

    .DESCRIPTION
    Generates detailed and summarized reports of Microsoft 365 licenses assigned to users,
    including cost analysis and usage insights, outputting to CSV and HTML formats.

    .AUTHOR
    Tycho Loke
    Website: https://currentcloud.net
    Blog: https://tycholoke.com
    As the maintainer and creator of this script, Tycho Loke offers tools for effective license
    management within Microsoft 365 environments. Explore more resources and get in touch
    through the author's website and blog.

    .NOTES
    Version: 1.0
    Updated: 27/03/2024
#>


Function Get-LicenseCosts {
  # Function to calculate the annual costs of the licenses assigned to a user account  
  [cmdletbinding()]
  Param( [array]$Licenses )
  [int]$Costs = 0
  ForEach ($License in $Licenses) {
    Try {
      [string]$LicenseCost = $PricingHashTable[$License]
      # Monthly cost in cents (because some licenses cost sums like 16.40)
      [float]$LicenseCostCents = [float]$LicenseCost * 100
      If ($LicenseCostCents -gt 0) {
        # Compute annual cost for the license
        [float]$AnnualCost = $LicenseCostCents * 12
        # Add to the cumulative license costs
        $Costs = $Costs + ($AnnualCost)
        # Write-Host ("License {0} Cost {1} running total {2}" -f $License, $LicenseCost, $Costs)
      }
    }
    Catch {
      Write-Host ("Error finding license {0} in pricing table - please check" -f $License)
    }
  }
  # Return 
  Return ($Costs / 100)
} 

[datetime]$RunDate = Get-Date -format "dd-MMM-yyyy HH:mm:ss"
$Version = "1.8"

# Default currency - can be overwritten by a value read into the $ImportSkus array
[string]$Currency = "EUR"

# Connect to the Graph, specifing the tenant and profile to use - Add your tenant identifier here
Connect-MgGraph -Scope "Directory.AccessAsUser.All, Directory.Read.All, AuditLog.Read.All" -NoWelcome


$SkuDataPath = "C:\temp\SkuDataComplete.csv"
$ServicePlanPath = "C:\temp\ServicePlanDataComplete.csv"
$CSVOutputFile = "c:\temp\Microsoft365LicensesReport.CSV"
$HtmlReportFile = "c:\temp\Microsoft365LicensesReport.html"
$UnlicensedAccounts = 0

If ((Test-Path $skuDataPath) -eq $False) {
  Write-Host ("Can't find the product data file ({0}). Exiting..." -f $skuDataPath) ; break 
}
If ((Test-Path $servicePlanPath) -eq $False) {
  Write-Host ("Can't find the serivice plan data file ({0}). Exiting..." -f $servicePlanPath) ; break 
}
   
$ImportSkus = Import-CSV $skuDataPath
$SkuHashTable = @{}

# Before your loop, initialize the hashtable
$PricingHashTable = @{}


ForEach ($Line in $ImportSkus) {
  If (-not [string]::IsNullOrWhiteSpace($Line.SkuId)) {
    If (-not $SkuHashTable.ContainsKey([string]$Line.SkuId)) {
      $SkuHashTable.Add([string]$Line.SkuId, [string]$Line.DisplayName)
    } Else {
      Write-Host ("Duplicate SKU ID detected and skipped: " + $Line.SkuId)
    }
  } Else {
    Write-Host "Found an entry with null or empty SkuId, skipping..."
  }
}


# If pricing information is in the $ImportSkus array, we can add the information to the report. We prepare to do this
# by setting the $PricingInfoAvailable to $true and populating the $PricingHashTable
$PricingInfoAvailable = $true

If ($ImportSkus[0].Price) {
  $PricingInfoAvailable = $true
  $Global:PricingHashTable = @{}
  ForEach ($Line in $ImportSkus) { 
    $PricingHashTable.Add([string]$Line.SkuId, [string]$Line.Price) 
  }
  If ($ImportSkus[0].Currency) {
    [string]$Currency = ($ImportSkus[0].Currency)
  }
}

# Find tenant accounts - but filtered so that we only fetch those with licenses
Write-Host "Finding licensed user accounts..."

$Users = Get-MgUser -All -ConsistencyLevel eventual -CountVariable Records `
  -Property id, displayName, userPrincipalName, country, department, assignedLicenses, `
  licenseAssignmentStates, createdDateTime, jobTitle, signInActivity, companyName | `
  Where-Object { $_.AssignedLicenses.Count -gt 0 } | Sort-Object DisplayName


If (!($Users)) { 
  Write-Host "No licensed user accounts found - exiting"; break 
}
Else { 
  Write-Host ("{0} Licensed user accounts found - now processing their license data..." -f $Users.Count) 
}

[array]$Departments = $Users.Department | Sort-Object -Unique
[array]$Countries = $Users.Country | Sort-Object -Unique
$OrgName = (Get-MgOrganization).DisplayName
$DuplicateSKUsAccounts = 0; $DuplicateSKULicenses = 0; $LicenseErrorCount = 0
$Report = [System.Collections.Generic.List[Object]]::new()
$i = 0
[float]$TotalUserLicenseCosts = 0
[float]$TotalBoughtLicenseCosts = 0

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
Write-Host ""
Write-Host "All done. Output files are" $CSVOutputFile "and" $ReportFile

Disconnect-MgGraph

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.
