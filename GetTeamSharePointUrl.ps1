
Import-Module ImportExcel  # Ensure you have the ImportExcel module installed
Import-Module SharePointPnPPowerShellOnline  # Ensure PnP PowerShell is installed


function ConnectToSharePoint {
    $adminUrl = "https://<your-tenant>-admin.sharepoint.com"
    $userCredential = Get-Credential
    Connect-PnPOnline -Url $adminUrl -Credentials $userCredential
}

# Read Teams from Excel
$excelPath = "C:\Path\To\Your\ExcelFile.xlsx"
$teams = Import-Excel -Path $excelPath


ConnectToSharePoint


$results = @()

foreach ($team in $teams) {
    $site = Get-PnPTenantSite -Detailed | Where-Object { $_.Url -like "*$($team.ID)*" }
    if ($site) {
        $results += [PSCustomObject]@{
            TeamName = $team.Name
            SharePointURL = $site.Url
        }
    } else {
        $results += [PSCustomObject]@{
            TeamName = $team.Name
            SharePointURL = "Not Found"
        }
    }
}

$results | Export-Csv -Path "C:\Path\To\Output\SharePointURLs.csv" -NoTypeInformation

Write-Output "Results have been saved to 'C:\Path\To\Output\SharePointURLs.csv'"
