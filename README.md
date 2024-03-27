# Microsoft 365 Current License Mapper

## Overview
The Microsoft 365 Current License Mapper is a PowerShell tool designed for system administrators and IT professionals to efficiently map and manage Microsoft 365 licenses within their organizations. This tool simplifies the identification and allocation of licenses, ensuring that users have appropriate access to Microsoft 365 services.

## Features

- **License Mapping**: Automatically maps current Microsoft 365 licenses to users based on input data.
- **CSV Support**: Leverages CSV files for input, facilitating the management of large sets of user data.
- **User-Friendly**: Accessible for administrators with varying levels of PowerShell experience.

## Prerequisites

Before starting, ensure you have:

- PowerShell 5.0 or higher installed on your machine.
- Administrative access to your Microsoft 365 tenant.
- A CSV file with user information for license mapping. Follow the "Script for CSV Execution" section to generate this file if needed.

## Installation

To use the Microsoft 365 Current License Mapper, clone this repository to your local machine:

```shell
git clone https://github.com/TychoLoke/microsoft-365-current-license-mapper.git
cd microsoft-365-current-license-mapper
```

## Usage

### Preparing Your CSV File

This script requires a CSV file that contains product names and service plan identifiers for licensing, which can be downloaded from the Microsoft documentation. Here's how to prepare and use this file with the script:

1. Download the "Product names and service plan identifiers for licensing" CSV file from [Microsoft's official documentation](https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference).

2. Move the downloaded CSV file to a designated folder, for example, `C:\temp\`.

3. Ensure you have the Microsoft Graph PowerShell SDK installed and configured. If not, you can install it using PowerShellGet with the following command:
    ```powershell
    Install-Module -Name Microsoft.Graph -Scope CurrentUser
    ```

4. Execute the `MicrosoftLicenseMapperCSV.ps1` script in PowerShell, which performs the following operations:
    - Connects to the Microsoft Graph, specifying the tenant and profile to use. Add your tenant identifier to the `Connect-MgGraph` cmdlet in the script.
    - Imports the product names and service plan identifiers from the CSV file you placed in `C:\temp\`.
    - Selects all SKUs with friendly display names and all service plans with friendly display names.
    - Retrieves all products used in your tenant and generates a CSV file (`SkuDataComplete.csv`) containing all product SKUs used in your tenant.
    - Generates a list of all service plans used in SKUs in your tenant and exports this list to another CSV file (`ServicePlanDataComplete.csv`).

To run the script, navigate to the directory containing `MicrosoftLicenseMapperCSV.ps1` and execute it:

```powershell
.\MicrosoftLicenseMapperCSV.ps1
```

Follow any prompts or instructions that appear during the execution. The script will generate two CSV files in `C:\temp\`:
- `SkuDataComplete.csv`: Contains SKU information for all products used in your tenant.
- `ServicePlanDataComplete.csv`: Contains service plan information for all SKUs used in your tenant.

These files can then be edited to add additional display name information and used to generate a comprehensive licensing report for your Microsoft 365 tenant.

### Mapping Licenses with `MicrosoftLicenseMapper.ps1`

After preparing your CSV files as described in the previous sections, you're ready to create a comprehensive report of licenses assigned to Azure AD user accounts using the Microsoft Graph PowerShell SDK cmdlets. Here’s how you can use the `MicrosoftLicenseMapper.ps1` script:

1. **Prepare the Environment**: Ensure you have the Microsoft Graph PowerShell SDK installed and your CSV files (`SkuDataComplete.csv` and `ServicePlanDataComplete.csv`) ready in `C:\temp\`.

2. **Execute the Script**: Run the license mapping script with the prepared CSV files. Open PowerShell and navigate to the directory containing `MicrosoftLicenseMapper.ps1`, then execute the script:

    ```powershell
    .\MicrosoftLicenseMapper.ps1
    ```

3. **Follow On-screen Instructions**: The script will connect to the Microsoft Graph and begin processing the license data for users in your tenant. Follow any additional prompts or instructions that appear.

4. **Review the Reports**: Upon completion, the script will generate two reports:
   - A CSV file (`Microsoft365LicensesReport.CSV`) containing detailed license assignment information.
   - An HTML file (`Microsoft365LicensesReport.html`) that provides a visual summary of license usage, costs (if pricing information is available), and other analytics.

5. **Post-Execution Steps**:
   - Review the generated reports located in `C:\temp\` to analyze the license distribution and costs in your Microsoft 365 tenant.
   - For detailed analysis, the HTML report provides breakdowns by department and country, including underused accounts and potential cost savings.

### Notes

- **Customization**: Before running the script, you may need to customize the CSV file paths or add your tenant identifier to the `Connect-MgGraph` cmdlet as instructed within the script comments.
- **Security**: Ensure you run this script in a secure environment, and review the permissions required for the Microsoft Graph SDK cmdlets used in the script.
- **Testing**: Always test scripts in a non-production environment before running them on your live tenant data to prevent unintended consequences.

Remember, this script offers valuable insights into your Microsoft 365 license usage and can help in optimizing license allocations and reducing costs. Follow the setup and execution steps carefully to ensure accurate reporting.


## Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement". Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

Distributed under the MIT License. See `LICENSE` for more information.
