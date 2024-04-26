# Get Team Sites
$tenantAdminUrl = "https://yourTenantName-admin.sharepoint.com"
Connect-PnPOnline -Url $tenantAdminUrl -Interactive

$outputFilePath = "C:\Temp\teams_channels.csv"
$teams = Get-PnPTeamsTeam
$data = @()
foreach ($team in $teams) {
    $channels = Get-PnPTeamsChannel -Team $team.GroupId
    foreach ($channel in $channels) {
        $data += [PSCustomObject]@{
            TeamName = $team.DisplayName
            ChannelName = $channel.DisplayName
        }
    }
}

$data | Export-Csv -Path $outputFilePath -NoTypeInformation
Write-Host "Data exported to $outputFilePath"
