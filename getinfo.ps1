# Initiate user-friendly interface
Write-Host "Welcome to the Azure AD License and User Info Exporter!" -ForegroundColor Cyan
Write-Host "Let's prepare your data for export." -ForegroundColor Green

# Prompt for Azure AD connection
Write-Host "Establishing connection to Azure AD. Please sign in when prompted..." -ForegroundColor Yellow
Connect-AzureAD

# Prompt for the directory path where the CSV will be saved
$DirectoryPath = Read-Host -Prompt "Please enter the directory path where you'd like to save the export CSV (e.g., C:\Exports)"

# Validate the provided directory path
if (-not (Test-Path -Path $DirectoryPath)) {
    Write-Host "The specified directory path does not exist. Please ensure the directory exists and try again." -ForegroundColor Red
    exit
}

# Generate the filename using the current date and time
$CurrentDateTime = Get-Date -Format "yyyy-MM-dd-HHmmss"
$FileName = "export-$CurrentDateTime.csv"
$ExportPath = Join-Path -Path $DirectoryPath -ChildPath $FileName

# Rest of the script remains the same...

# Result array initialization
$Result = @()

# Fetch all subscribed license Skus to Microsoft services
$AllSkus = Get-AzureADSubscribedSku

# Fetch all Azure AD Users with the required properties
$AllUsers = Get-AzureADUser -All $true | Select DisplayName, UserPrincipalName, AssignedLicenses, AssignedPlans, ObjectId

# User iteration for license detail resolution
foreach ($User in $AllUsers) {
    $AssignedLicenses = @()
    $LicensedServices = @()

    if ($User.AssignedLicenses.Count -ne 0) {
        # Resolve license SKU details
        foreach ($License in $User.AssignedLicenses) {
            $SkuInfo = $AllSkus | Where { $_.SkuId -eq $License.SkuId }
            $AssignedLicenses += $SkuInfo.SkuPartNumber
        }

        # Resolve assigned service plans
        foreach ($ServicePlan in $User.AssignedPlans) {
            $LicensedServices += $ServicePlan.Service
        }
    }

    # Compiling user details
    $Result += New-Object PSObject -property $([ordered]@{
        UserName           = $User.DisplayName
        UserPrincipalName  = $User.UserPrincipalName
        UserId             = $User.ObjectId
        IsLicensed         = if ($User.AssignedLicenses.Count -ne 0) { $true } else { $false }
        Licenses           = $AssignedLicenses -join ","
        LicensedServices   = ($LicensedServices | Sort-Object | Get-Unique) -join ","
    })
}

# Exporting the result to CSV
$Result | Export-CSV $ExportPath -NoTypeInformation -Encoding UTF8
Write-Host "Export completed successfully! The data has been saved to $ExportPath" -ForegroundColor Green
