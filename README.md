# Microsoft CoPilot License Calculator

A PowerShell-based GUI calculator for estimating Microsoft CoPilot licensing costs. This tool provides an interactive interface to calculate both monthly and annual costs for different types of CoPilot subscriptions.

## Features

- Interactive Windows Forms-based GUI
- Calculates costs for multiple CoPilot license types:
  - E5 + CoPilot ($84.75/user/month)
  - E3 + CoPilot ($63.75/user/month)
  - Business Premium + CoPilot ($52/user/month)
  - Security CoPilot ($4/SCU/hour)
- Provides both monthly and annual cost breakdowns
- Export functionality to save calculations in CSV format

## Requirements

- Windows PowerShell 5.1 or later
- Windows Forms assembly
- System.Drawing assembly

## Installation

1. Clone this repository or download the `CoPilotLicenseCalculator.ps1` file
2. Ensure you have the required PowerShell version and assemblies installed
3. Run the script directly from PowerShell

## Usage

```powershell
.\CoPilotLicenseCalculator.ps1
```

The script will launch an interactive GUI where you can:
1. Input the number of users for each license type
2. Specify Security CoPilot SCUs and running hours
3. Calculate total costs using the "Calculate" button
4. Export results to CSV format using the "Export to CSV" button

## Export Format

The CSV export includes the following columns:
- License Type
- Number of Users
- Monthly Cost
- Annual Cost

## Version Information

- Version: 1.2
- Author: John O'Neill Sr.
- Creation Date: 2024-02-12
- Last Modified: 2024-11-12

## Additional Resources

- [Microsoft 365 Copilot Documentation](https://learn.microsoft.com/microsoft-365/copilot)
- [Microsoft 365 Enterprise Plans and Pricing](https://www.microsoft.com/en-us/microsoft-365/enterprise/microsoft365-plans-and-pricing)
- [Microsoft 365 Business Products Comparison](https://www.microsoft.com/en-us/microsoft-365/business/compare-all-microsoft-365-business-products)

## Components

- Microsoft.PowerShell.Management
- System.Windows.Forms
- System.Drawing

## License Costs

The calculator uses the following pricing:
- E5 + CoPilot: $54.75 + $30 = $84.75/user/month
- E3 + CoPilot: $33.75 + $30 = $63.75/user/month
- Business Premium + CoPilot: $22 + $30 = $52/user/month
- Security CoPilot: $4/SCU/hour

