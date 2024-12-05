<#
.SYNOPSIS
    Creates an interactive GUI calculator for Microsoft CoPilot licensing costs.

.DESCRIPTION
    This script provides a Windows Forms-based calculator for estimating Microsoft CoPilot licensing costs.
    It allows users to input the number of licenses needed for different types of CoPilot subscriptions
    and calculates both monthly and annual costs. The calculator includes export functionality for saving
    results to CSV format.

.PARAMETER None
    This script does not accept any parameters as it runs interactively through a GUI.

.INPUTS
    None. This script does not accept pipeline input.

.OUTPUTS
    Optional CSV file containing license calculations with the following columns:
    - License Type
    - Number of Users
    - Monthly Cost
    - Annual Cost

.EXAMPLE
    .\CoPilotLicenseCalculator.ps1
    Launches the interactive GUI calculator for Microsoft CoPilot licensing costs.

.NOTES
    Version:        1.2
    Author:         John O'Neill Sr.
    Creation Date:  2024-02-12
    Last Modified:  2024-11-12
    
    License costs used in calculations:
    - E5 + CoPilot: $54.75 + $30 = $84.75/user/month
    - E3 + CoPilot: $33.75 + $30 = $63.75/user/month
    - Business Premium + CoPilot: $22 + $30 = $52/user/month
    - Security CoPilot: $4/SCU/hour

    Requirements:
    - Windows PowerShell 5.1 or later
    - Windows Forms assembly
    - System.Drawing assembly

.LINK
    https://learn.microsoft.com/microsoft-365/copilot
    https://www.microsoft.com/en-us/microsoft-365/enterprise/microsoft365-plans-and-pricing
    https://www.microsoft.com/en-us/microsoft-365/business/compare-all-microsoft-365-business-products


.COMPONENT
    Microsoft.PowerShell.Management
    System.Windows.Forms
    System.Drawing

.FUNCTIONALITY
    - Calculates Microsoft CoPilot licensing costs
    - Provides monthly and annual cost breakdowns
    - Exports calculations to CSV
    - Interactive GUI interface
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Microsoft CoPilot License Calculator'
$form.Size = New-Object System.Drawing.Size(600,700)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Title Label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(20,20)
$titleLabel.Size = New-Object System.Drawing.Size(550,30)
$titleLabel.Text = 'Microsoft CoPilot License Calculator'
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI',14,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($titleLabel)

# E5 Users
$e5Label = New-Object System.Windows.Forms.Label
$e5Label.Location = New-Object System.Drawing.Point(20,70)
$e5Label.Size = New-Object System.Drawing.Size(280,20)
$e5Label.Text = 'Number of E5 + CoPilot Users ($84.75/user):'
$form.Controls.Add($e5Label)

$e5TextBox = New-Object System.Windows.Forms.TextBox
$e5TextBox.Location = New-Object System.Drawing.Point(300,70)
$e5TextBox.Size = New-Object System.Drawing.Size(100,20)
$e5TextBox.Text = "0"
$form.Controls.Add($e5TextBox)

# E3 Users
$e3Label = New-Object System.Windows.Forms.Label
$e3Label.Location = New-Object System.Drawing.Point(20,110)
$e3Label.Size = New-Object System.Drawing.Size(280,20)
$e3Label.Text = 'Number of E3 + CoPilot Users ($63.75/user):'
$form.Controls.Add($e3Label)

$e3TextBox = New-Object System.Windows.Forms.TextBox
$e3TextBox.Location = New-Object System.Drawing.Point(300,110)
$e3TextBox.Size = New-Object System.Drawing.Size(100,20)
$e3TextBox.Text = "0"
$form.Controls.Add($e3TextBox)

# Business Premium Users
$bpLabel = New-Object System.Windows.Forms.Label
$bpLabel.Location = New-Object System.Drawing.Point(20,150)
$bpLabel.Size = New-Object System.Drawing.Size(280,20)
$bpLabel.Text = 'Number of Business Premium Users ($52/user):'
$form.Controls.Add($bpLabel)

$bpTextBox = New-Object System.Windows.Forms.TextBox
$bpTextBox.Location = New-Object System.Drawing.Point(300,150)
$bpTextBox.Size = New-Object System.Drawing.Size(100,20)
$bpTextBox.Text = "0"
$form.Controls.Add($bpTextBox)

# Security SCUs
$secLabel = New-Object System.Windows.Forms.Label
$secLabel.Location = New-Object System.Drawing.Point(20,190)
$secLabel.Size = New-Object System.Drawing.Size(280,20)
$secLabel.Text = 'Number of Security CoPilot SCUs ($4/SCU/hour):'
$form.Controls.Add($secLabel)

$secTextBox = New-Object System.Windows.Forms.TextBox
$secTextBox.Location = New-Object System.Drawing.Point(300,190)
$secTextBox.Size = New-Object System.Drawing.Size(100,20)
$secTextBox.Text = "0"
$form.Controls.Add($secTextBox)

# SCU Hours per Day
$hoursLabel = New-Object System.Windows.Forms.Label
$hoursLabel.Location = New-Object System.Drawing.Point(20,230)
$hoursLabel.Size = New-Object System.Drawing.Size(280,20)
$hoursLabel.Text = 'Number of Hours SCUs Run per Day:'
$form.Controls.Add($hoursLabel)

$hoursTextBox = New-Object System.Windows.Forms.TextBox
$hoursTextBox.Location = New-Object System.Drawing.Point(300,230)
$hoursTextBox.Size = New-Object System.Drawing.Size(100,20)
$hoursTextBox.Text = "0"
$form.Controls.Add($hoursTextBox)

# Results Group Box
$resultsGroup = New-Object System.Windows.Forms.GroupBox
$resultsGroup.Location = New-Object System.Drawing.Point(20,310)
$resultsGroup.Size = New-Object System.Drawing.Size(540,300)
$resultsGroup.Text = "Cost Breakdown"
$form.Controls.Add($resultsGroup)

# Results Labels
$monthlyE5Label = New-Object System.Windows.Forms.Label
$monthlyE5Label.Location = New-Object System.Drawing.Point(10,30)
$monthlyE5Label.Size = New-Object System.Drawing.Size(500,20)
$monthlyE5Label.Text = 'E5 + CoPilot Monthly Cost: $0'
$resultsGroup.Controls.Add($monthlyE5Label)

$monthlyE3Label = New-Object System.Windows.Forms.Label
$monthlyE3Label.Location = New-Object System.Drawing.Point(10,60)
$monthlyE3Label.Size = New-Object System.Drawing.Size(500,20)
$monthlyE3Label.Text = 'E3 + CoPilot Monthly Cost: $0'
$resultsGroup.Controls.Add($monthlyE3Label)

$monthlyBPLabel = New-Object System.Windows.Forms.Label
$monthlyBPLabel.Location = New-Object System.Drawing.Point(10,90)
$monthlyBPLabel.Size = New-Object System.Drawing.Size(500,20)
$monthlyBPLabel.Text = 'Business Premium Monthly Cost: $0'
$resultsGroup.Controls.Add($monthlyBPLabel)

$monthlySecLabel = New-Object System.Windows.Forms.Label
$monthlySecLabel.Location = New-Object System.Drawing.Point(10,120)
$monthlySecLabel.Size = New-Object System.Drawing.Size(500,20)
$monthlySecLabel.Text = 'Security CoPilot Monthly (31 Days) Cost: $0'
$resultsGroup.Controls.Add($monthlySecLabel)

$totalMonthlyLabel = New-Object System.Windows.Forms.Label
$totalMonthlyLabel.Location = New-Object System.Drawing.Point(10,160)
$totalMonthlyLabel.Size = New-Object System.Drawing.Size(500,20)
$totalMonthlyLabel.Text = 'Total Monthly Cost: $0'
$totalMonthlyLabel.Font = New-Object System.Drawing.Font($totalMonthlyLabel.Font, [System.Drawing.FontStyle]::Bold)
$resultsGroup.Controls.Add($totalMonthlyLabel)

$totalAnnualLabel = New-Object System.Windows.Forms.Label
$totalAnnualLabel.Location = New-Object System.Drawing.Point(10,190)
$totalAnnualLabel.Size = New-Object System.Drawing.Size(500,20)
$totalAnnualLabel.Text = 'Total Annual Cost: $0'
$totalAnnualLabel.Font = New-Object System.Drawing.Font($totalAnnualLabel.Font, [System.Drawing.FontStyle]::Bold)
$resultsGroup.Controls.Add($totalAnnualLabel)

# Export Button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(10,230)
$exportButton.Size = New-Object System.Drawing.Size(150,30)
$exportButton.Text = 'Export to CSV'
$resultsGroup.Controls.Add($exportButton)

# Calculate Button
$calculateButton = New-Object System.Windows.Forms.Button
$calculateButton.Location = New-Object System.Drawing.Point(20,270)
$calculateButton.Size = New-Object System.Drawing.Size(150,30)
$calculateButton.Text = 'Calculate'
$calculateButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,212)
$calculateButton.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($calculateButton)

# Calculate Button Click Event
$calculateButton.Add_Click({
    try {
        $e5Users = [int]$e5TextBox.Text
        $e3Users = [int]$e3TextBox.Text
        $bpUsers = [int]$bpTextBox.Text
        $secSCUs = [int]$secTextBox.Text
        $hoursPerDay = [int]$hoursTextBox.Text

        $e5Cost = $e5Users * 84.75
        $e3Cost = $e3Users * 63.75
        $bpCost = $bpUsers * 52
        $secCost = $secSCUs * 4 * $hoursPerDay * 31 # Assuming 31 days in a month

        $totalMonthly = $e5Cost + $e3Cost + $bpCost + $secCost
        $totalAnnual = $totalMonthly * 12

        $monthlyE5Label.Text = "E5 + CoPilot Monthly Cost: $" + $e5Cost.ToString('N2')
        $monthlyE3Label.Text = "E3 + CoPilot Monthly Cost: $" + $e3Cost.ToString('N2')
        $monthlyBPLabel.Text = "Business Premium Monthly Cost: $" + $bpCost.ToString('N2')
        $monthlySecLabel.Text = "Security CoPilot Monthly (31 Days) Cost: $" + $secCost.ToString('N2')
        $totalMonthlyLabel.Text = "Total Monthly Cost: $" + $totalMonthly.ToString('N2')
        $totalAnnualLabel.Text = "Total Annual Cost: $" + $totalAnnual.ToString('N2')

        $script:lastCalculation = @{
            E5Users = $e5Users
            E3Users = $e3Users
            BPUsers = $bpUsers
            SecuritySCUs = $secSCUs
            HoursPerDay = $hoursPerDay
            E5Cost = $e5Cost
            E3Cost = $e3Cost
            BPCost = $bpCost
            SecurityCost = $secCost
            TotalMonthly = $totalMonthly
            TotalAnnual = $totalAnnual
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Please enter valid numbers.", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Export Button Click Event
$exportButton.Add_Click({
    if ($script:lastCalculation) {
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        $SaveFileDialog.DefaultExt = "csv"
        $SaveFileDialog.AddExtension = $true

        if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # Create an array of custom objects for each license type
            $exportData = @(
                [PSCustomObject]@{
                    'License Type' = 'E5 + CoPilot'
                    'Users' = $script:lastCalculation.E5Users
                    'Monthly Cost' = "$" + $script:lastCalculation.E5Cost.ToString('N2')
                    'Annual Cost' = "$" + ($script:lastCalculation.E5Cost * 12).ToString('N2')
                },
                [PSCustomObject]@{
                    'License Type' = 'E3 + CoPilot'
                    'Users' = $script:lastCalculation.E3Users
                    'Monthly Cost' = "$" + $script:lastCalculation.E3Cost.ToString('N2')
                    'Annual Cost' = "$" + ($script:lastCalculation.E3Cost * 12).ToString('N2')
                },
                [PSCustomObject]@{
                    'License Type' = 'Business Premium + CoPilot'
                    'Users' = $script:lastCalculation.BPUsers
                    'Monthly Cost' = "$" + $script:lastCalculation.BPCost.ToString('N2')
                    'Annual Cost' = "$" + ($script:lastCalculation.BPCost * 12).ToString('N2')
                },
                [PSCustomObject]@{
                    'License Type' = 'Security CoPilot'
                    'SCUs' = $script:lastCalculation.SecuritySCUs
                    'Hours Per Day' = $script:lastCalculation.HoursPerDay
                    'Monthly Cost' = "$" + $script:lastCalculation.SecurityCost.ToString('N2')
                    'Annual Cost' = "$" + ($script:lastCalculation.SecurityCost * 12).ToString('N2')
                },
                [PSCustomObject]@{
                    'License Type' = 'Total'
                    'Users' = ($script:lastCalculation.E5Users + 
                             $script:lastCalculation.E3Users + 
                             $script:lastCalculation.BPUsers)
                    'SCUs' = $script:lastCalculation.SecuritySCUs
                    'Monthly Cost' = "$" + $script:lastCalculation.TotalMonthly.ToString('N2')
                    'Annual Cost' = "$" + $script:lastCalculation.TotalAnnual.ToString('N2')
                }
            )

            $exportData | Export-Csv -Path $SaveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Export completed successfully!", "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please calculate costs before exporting.", "Warning",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# Show the form
$form.ShowDialog()
