# Required Modules
Import-Module ImportExcel  # Ensure you have the ImportExcel module installed
Import-Module SharePointPnPPowerShellOnline  # Ensure PnP PowerShell is installed

# Function to connect to SharePoint
function ConnectToSharePoint {
    $adminUrl = "https://<your-tenant>-admin.sharepoint.com"
    $userCredential = Get-Credential
    Connect-PnPOnline -Url $adminUrl -Credentials $userCredential
}

# Read Group IDs from Excel
$excelPath = "C:\Path\To\Your\ExcelFile.xlsx"
$groups = Import-Excel -Path $excelPath

# Connect to SharePoint
ConnectToSharePoint

# Prepare results list
$results = @()

# Retrieve SharePoint URLs using Group IDs and store results
foreach ($group in $groups) {
    $groupId = $group.ID  # Ensure your Excel has a column named 'ID' for Group IDs
    try {
        $siteUrl = Get-PnPSite -GroupId $groupId -ErrorAction Stop
        $results += [PSCustomObject]@{
            GroupID = $groupId
            SharePointURL = $siteUrl.Url
        }
    }
    catch {
        $results += [PSCustomObject]@{
            GroupID = $groupId
            SharePointURL = "Not Found or Access Denied"
        }
    }
}

# Export results to CSV
$results | Export-Csv -Path "C:\Path\To\Output\SharePointURLs.csv" -NoTypeInformation

Write-Output "Results have been saved to 'C:\Path\To\
