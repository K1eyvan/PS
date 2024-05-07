# Required Modules
Import-Module ImportExcel  # Ensure you have the ImportExcel module installed
Import-Module SharePointPnPPowerShellOnline  # Ensure PnP PowerShell is installed
Import-Module MicrosoftTeams  # Ensure the Teams PowerShell module is installed

# Function to connect to SharePoint and Teams
function ConnectToServices {
    $adminUrl = "https://<your-tenant>-admin.sharepoint.com"
    $userCredential = Get-Credential
    Connect-PnPOnline -Url $adminUrl -Credentials $userCredential
    Connect-MicrosoftTeams -Credential $userCredential
}

# Read Channel IDs from Excel
$excelPath = "C:\Path\To\Your\ExcelFile.xlsx"
$channels = Import-Excel -Path $excelPath

# Connect to Services
ConnectToServices

# Prepare results list
$results = @()

# Retrieve SharePoint URLs using Channel IDs and store results
foreach ($channel in $channels) {
    $channelId = $channel.ID  # Ensure your Excel has a column named 'ID' for Channel IDs
    try {
        $teamId = (Get-TeamChannel -GroupId <YourTeamId> | Where-Object { $_.Id -eq $channelId }).GroupId
        $site = Get-PnPSite -GroupId $teamId -ErrorAction Stop
        $channelFolderUrl = $site.Url + "/Shared Documents/" + (Get-TeamChannel -GroupId $teamId | Where-Object { $_.Id -eq $channelId }).DisplayName
        $results += [PSCustomObject]@{
            ChannelID = $channelId
            SharePointFolderURL = $channelFolderUrl
        }
    }
    catch {
        $results += [PSCustomObject]@{
            ChannelID = $channelId
            SharePointFolderURL = "Not Found or Access Denied"
        }
    }
}

# Export results to CSV
$results | Export-Csv -Path "C:\Path\To\Output\ChannelURLs.csv" -NoTypeInformation

Write-Output "Results have been saved to 'C:\Path\To\Output\ChannelURLs.csv'"
