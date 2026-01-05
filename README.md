# Microsoft 365 Current License Mapper

[![PowerShell](https://img.shields.io/badge/PowerShell-7.0%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
  - [Preparing Your CSV File](#preparing-your-csv-file)
  - [Mapping Licenses with MicrosoftLicenseMapper.ps1](#mapping-licenses-with-microsoftlicensemapperps1)
- [Configuration](#configuration)
- [Output Files](#output-files)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)
- [Author](#author)

## Overview

The Microsoft 365 Current License Mapper is a comprehensive PowerShell solution designed for system administrators and IT professionals to efficiently audit, map, and manage Microsoft 365 licenses within their organizations. This tool provides detailed insights into license allocation, costs, and usage patterns, enabling better resource optimization and cost management.

## Features

- **Modern Interactive HTML Reports**: Beautiful, responsive web-based reports with sortable tables and dashboard cards
- **Advanced Cleanup Filtering**: Quick filter presets for tenant cleanup scenarios:
  - Never logged in accounts
  - High priority cleanup candidates (never used or 90+ days inactive)
  - Inactive accounts (30+, 60+, 90+, 180+ days)
  - Duplicate license assignments
  - Enhanced status categorization for better cleanup decision-making
- **Comprehensive License Reporting**: Generate detailed reports of all Microsoft 365 licenses assigned to users
- **Cost Analysis**: Calculate and analyze licensing costs by user, department, and country (when pricing data is available)
- **Dual-Assignment Detection**: Identify users with duplicate license assignments (both direct and group-based)
- **Usage Analytics**: Track user sign-in activity and identify underutilized accounts with granular inactivity levels
- **Multiple Output Formats**: Export reports in both CSV and modern HTML formats with interactive features
- **Group-Based License Support**: Full visibility into both direct and group-based license assignments
- **Service Plan Visibility**: View enabled and disabled service plans for each license
- **Automated Data Collection**: Automatically retrieve SKU and service plan information from your tenant
- **Responsive Design**: HTML reports work seamlessly on desktop, tablet, and mobile devices

## Prerequisites

Before using this tool, ensure you have the following:

- **PowerShell 7.0 or Higher**: This tool requires PowerShell 7.0+
  - **Windows**: Download from [PowerShell GitHub Releases](https://github.com/PowerShell/PowerShell/releases) or install via:
    ```powershell
    winget install Microsoft.PowerShell
    ```
  - **macOS**: Install via Homebrew:
    ```bash
    brew install powershell/tap/powershell
    ```
  - **Linux**: Follow the [installation guide for your distribution](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux)
  - **Verify Installation**: Check your version with `$PSVersionTable`
- **Microsoft Graph PowerShell SDK**: Version 2.0 or higher
  ```powershell
  Install-Module -Name Microsoft.Graph -Scope CurrentUser
  ```
- **Microsoft 365 Permissions**: An account with one of the following roles:
  - Global Administrator
  - Global Reader
  - License Administrator
  - User Administrator
- **Microsoft Graph API Permissions**:
  - `Directory.Read.All`
  - `Directory.AccessAsUser.All` (for the main reporting script)
  - `AuditLog.Read.All` (for sign-in activity data)
- **Storage**: Writable access to `C:\temp\` directory (or modify script paths as needed)
- **CSV Reference File**: Download the "Product names and service plan identifiers for licensing" CSV from [Microsoft's official documentation](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference)

## Installation

1. **Clone the Repository**:
   ```powershell
   git clone https://github.com/TychoLoke/microsoft-365-current-license-mapper.git
   cd microsoft-365-current-license-mapper
   ```

2. **Install Microsoft Graph PowerShell SDK** (if not already installed):
   ```powershell
   Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
   ```

3. **Create Working Directory**:
   ```powershell
   New-Item -Path "C:\temp" -ItemType Directory -Force
   ```

## Usage

### Step 1: Preparing Your CSV File

Before generating license reports, you need to create reference CSV files containing SKU and service plan information for your tenant.

1. **Download Microsoft's Reference CSV**:
   - Navigate to [Microsoft's Licensing Service Plan Reference](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference)
   - Download the "Product names and service plan identifiers for licensing" CSV file
   - Save it to `C:\temp\`

2. **Run the CSV Preparation Script**:
   ```powershell
   .\MicrosoftLicenseMapperCSV.ps1
   ```

3. **What This Script Does**:
   - Connects to Microsoft Graph with MFA support
   - Imports the Microsoft reference CSV file
   - Retrieves all SKUs currently used in your tenant
   - Maps SKU IDs to friendly display names
   - Generates two output files:
     - `C:\temp\SkuDataComplete.csv` - SKU information with display names
     - `C:\temp\ServicePlanDataComplete.csv` - Service plan information with friendly names

4. **(Optional) Enhance the CSV Files**:
   - Open the generated CSV files
   - Add pricing information in a `Price` column (monthly cost per license)
   - Add a `Currency` column (e.g., "USD", "EUR", "GBP")
   - This enables cost analysis features in the main report

### Step 2: Generate License Reports

After preparing your CSV files, you can now generate comprehensive license reports.

1. **Verify Prerequisites**:
   - Ensure `SkuDataComplete.csv` and `ServicePlanDataComplete.csv` exist in `C:\temp\`
   - Confirm you have the required Microsoft Graph permissions

2. **Run the License Mapping Script**:
   ```powershell
   .\MicrosoftLicenseMapper.ps1
   ```

3. **Authentication**:
   - The script will prompt you to sign in to Microsoft Graph
   - Use an account with appropriate administrative permissions
   - Approve the requested permissions when prompted

4. **Processing**:
   - The script will process all licensed user accounts in your tenant
   - Progress will be displayed for each user account
   - Processing time depends on the number of licensed users

## Configuration

### Customizing File Paths

By default, the scripts use `C:\temp\` for all input and output files. To use different paths, modify the following variables in each script:

**In `MicrosoftLicenseMapperCSV.ps1`**:
```powershell
$csvPath = "C:\temp\Product names and service plan identifiers for licensing.csv"
$skuCsvPath = "C:\temp\SkuDataComplete.csv"
$servicePlanCsvPath = "C:\Temp\ServicePlanDataComplete.csv"
```

**In `MicrosoftLicenseMapper.ps1`**:
```powershell
$SkuDataPath = "C:\temp\SkuDataComplete.csv"
$ServicePlanPath = "C:\temp\ServicePlanDataComplete.csv"
$CSVOutputFile = "c:\temp\Microsoft365LicensesReport.CSV"
$HtmlReportFile = "c:\temp\Microsoft365LicensesReport.html"
```

### Adding Pricing Information

To enable cost analysis features:

1. Open `C:\temp\SkuDataComplete.csv` in Excel or a text editor
2. Add a column named `Price` with the monthly cost per license (in decimal format, e.g., "12.50")
3. Add a column named `Currency` with your currency code (e.g., "USD", "EUR", "GBP")
4. Save the file

## Tenant Cleanup Features

The tool now includes powerful cleanup filtering capabilities to help you identify and manage underutilized licenses:

### Cleanup Categories

1. **Never Logged In Accounts** - Users who have never accessed their account
   - Immediate cleanup candidates for license recovery
   - Clear indicator in the Status column

2. **High Priority Cleanup** - Accounts with critical inactivity
   - Never logged in users
   - Inactive for 90+ days
   - Inactive for 180+ days

3. **Inactive Account Tiers**
   - **30+ days**: Monitor for potential issues
   - **60+ days**: Review recommended
   - **90+ days**: Cleanup candidate
   - **180+ days**: High priority cleanup

### Using Quick Filters

The HTML report includes one-click filter buttons in the toolbar:
- Click any filter button to instantly show only matching accounts
- Filter status shows the count of filtered results
- Click "Show All" to reset filters
- Combine with the search box for advanced filtering

### Recommended Cleanup Workflow

1. **Start with "Never Logged In"** - These are safe cleanup candidates if confirmed
2. **Review "High Priority Cleanup"** - Focus on accounts inactive 90+ days
3. **Check "Inactive 60+ Days"** - Evaluate if licenses can be reassigned
4. **Review "Duplicate Licenses"** - Eliminate wasteful duplicate assignments
5. **Export filtered results** to CSV for documentation and approval processes

## Output Files

### CSV Report (`Microsoft365LicensesReport.CSV`)

Contains detailed information for each licensed user:
- User display name and UPN
- Country and department
- Job title and company
- Direct assigned licenses
- Disabled service plans
- Group-based licenses
- Annual license costs (if pricing enabled)
- Last license change date
- Account creation date
- Last sign-in information
- Inactive account warnings
- Duplicate license warnings

### HTML Report (`Microsoft365LicensesReport.html`)

A modern, interactive web-based report featuring:

**Dashboard Overview:**
- Interactive summary cards displaying key metrics at a glance
- Color-coded statistics (success, warning, danger, info)
- Responsive grid layout adapting to screen sizes
- Cost analysis cards (when pricing data is available)

**Data Tables:**
- **User License Details**: Complete list of all licensed users with sortable columns
- **Product License Distribution**: SKU usage, costs, and utilization
- **License Costs by Department**: Breakdown by organizational department
- **License Costs by Country**: Breakdown by geographic location

**Key Features:**
- **Sortable Tables**: Click any column header to sort data ascending/descending
- **Responsive Design**: Optimized for desktop, tablet, and mobile viewing
- **Modern UI**: Microsoft Fluent-inspired design with gradient headers
- **Hover Effects**: Interactive elements with smooth transitions
- **Print-Friendly**: Optimized CSS for professional printing
- **Sticky Headers**: Table headers remain visible while scrolling

**Metrics Displayed:**
- Total licensed accounts
- Underused accounts (count and percentage)
- Never logged in accounts
- High priority cleanup candidates
- Inactive accounts (60+ days, 90+ days, 180+ days)
- Duplicate licenses detected
- License assignment errors
- Total licensing costs (when pricing enabled)
- Average cost per user
- License utilization percentage

**Cleanup Filtering:**
- **Quick Filter Buttons**: One-click filtering for common cleanup scenarios
- **Never Logged In**: Identify accounts that have never been accessed
- **High Priority Cleanup**: Accounts never used or inactive for 90+ days
- **Inactive Filters**: Filter by 60+, 90+ days of inactivity
- **Duplicate Licenses**: Find accounts with duplicate license assignments
- **Smart Status Categories**: Enhanced status messages for better decision-making:
  - "Never logged in - Cleanup candidate"
  - "Inactive 180+ days - High priority cleanup"
  - "Inactive 90+ days - Cleanup candidate"
  - "Inactive 60+ days - Review recommended"
  - "Inactive 30+ days - Monitor"

## Troubleshooting

### Common Issues

**Issue**: "This script requires PowerShell 7.0 or higher"
- **Solution**: You're running an older version of PowerShell (likely Windows PowerShell 5.1)
- **Fix**: Install PowerShell 7+ using one of these methods:
  - **Windows**: `winget install Microsoft.PowerShell` or download from [PowerShell releases](https://github.com/PowerShell/PowerShell/releases)
  - **macOS**: `brew install powershell/tap/powershell`
  - **Linux**: Follow the [Linux installation guide](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux)
- After installation, run the scripts using `pwsh` instead of `powershell`:
  ```powershell
  pwsh .\MicrosoftLicenseMapperCSV.ps1
  ```

**Issue**: "Can't find the product data file"
- **Solution**: Ensure you've run `MicrosoftLicenseMapperCSV.ps1` first to generate the required CSV files

**Issue**: "Failed to connect to Microsoft Graph"
- **Solution**: Verify your credentials and ensure MFA is properly configured
- Check that you have the required administrative roles

**Issue**: "Insufficient privileges to complete the operation"
- **Solution**: Ensure your account has one of the required roles (Global Administrator, Global Reader, License Administrator, or User Administrator)

**Issue**: Pricing information not showing in reports
- **Solution**: Add `Price` and `Currency` columns to `SkuDataComplete.csv` as described in the Configuration section

**Issue**: Script execution is disabled
- **Solution**: Run PowerShell as Administrator and execute:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

### Best Practices

- **Test First**: Always run scripts in a test environment before production use
- **Regular Backups**: Keep backups of your CSV files after customization
- **Schedule Reports**: Consider scheduling monthly reports to track license usage trends
- **Review Inactive Accounts**: Regularly review accounts flagged as underused to optimize costs
- **Audit Duplicates**: Investigate and resolve duplicate license assignments
- **Update Reference Data**: Periodically re-download Microsoft's reference CSV to ensure accuracy


## Contributing

Contributions are welcome and greatly appreciated! This project benefits from community feedback and enhancements.

### How to Contribute

1. **Fork the Repository**
   ```bash
   git clone https://github.com/TychoLoke/microsoft-365-current-license-mapper.git
   ```

2. **Create a Feature Branch**
   ```bash
   git checkout -b feature/YourFeatureName
   ```

3. **Make Your Changes**
   - Follow PowerShell best practices
   - Add comments for complex logic
   - Test thoroughly before committing

4. **Commit Your Changes**
   ```bash
   git commit -m "Add: Brief description of your changes"
   ```

5. **Push to Your Fork**
   ```bash
   git push origin feature/YourFeatureName
   ```

6. **Open a Pull Request**
   - Provide a clear description of the changes
   - Reference any related issues
   - Include testing details

### Reporting Issues

If you encounter bugs or have feature requests:
- Open an issue on GitHub
- Use descriptive titles
- Include reproduction steps for bugs
- Provide environment details (PowerShell version, OS, etc.)

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for full details.

## Author

**Tycho Löke**
- Website: [https://tycholoke.com](https://tycholoke.com)
- Portfolio: [https://currentcloud.net](https://currentcloud.net)
- GitHub: [@TychoLoke](https://github.com/TychoLoke)

**Copyright © 2026 Tycho Löke. All rights reserved.**

This tool is copyrighted by Tycho Löke and available at [tycholoke.com](https://tycholoke.com). While licensed under MIT for use and modification, attribution to the original author must be maintained in all distributions and derivative works.

---

## Disclaimer

This tool is provided as-is for informational and administrative purposes. Always test scripts in a non-production environment before deploying to production. The author and contributors are not responsible for any data loss or issues arising from the use of this tool.

## Acknowledgments

- Microsoft Graph PowerShell SDK team
- Microsoft 365 community
- Contributors and users providing feedback

---

**Note**: This is an independent project and is not officially affiliated with or endorsed by Microsoft Corporation.
