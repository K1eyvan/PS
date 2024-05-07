# Import the PnP.PowerShell module
Import-Module PnP.PowerShell

# Optional: Import the module for handling Excel files
Import-Module ImportExcel

# Path to the Excel file
$excelPath = "C:\path\to\your\file.xlsx"

# Load the Excel data
$siteUrls = Import-Excel -Path $excelPath

# Your sensitivity label to apply
$labelName = "Internal Only"

# Loop through each site URL
foreach ($site in $siteUrls) {
    # Connect to the SharePoint site
    Connect-PnPOnline -Url $site.URL -Interactive

    # Set the sensitivity label
    Set-PnPSite -SensitivityLabel $labelName

    # Disconnect the session
    Disconnect-PnPOnline
}

# Output completion
Write-Host "All sites have been labeled as '$labelName'"
