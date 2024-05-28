param
(
    [Parameter(Mandatory = $true)][string] $Tenant,
    [Parameter(Mandatory = $true)][string] $ClientId,
    [Parameter(Mandatory = $true)][string] $CertificatePath,
    [Parameter(Mandatory = $true)][securestring] $CertPassword
)
 
 
if ($CertPassword -eq $null) {
    $CertPassword = (Read-Host -AsSecureString)
}
 
$excelPath = "TeamsAndChannels2024-05-07-10-01-03.xlsx"
$teams = Import-Excel -Path $excelPath
$Timestamp = Get-Date -format "yyyy-MM-dd-hh-mm-ss"
#ConnectToSharePoint
 
$tenantAdminConn = Connect-PnPOnline -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword $CertPassword -Url "https://$tenant-admin.sharepoint.com" -Tenant "$Tenant.onmicrosoft.com" -ReturnConnection
Write-Host "Connected!"
 
write-host $teams.Count
 
$results = @()
 
foreach ($channel in $teams) {
    $channelId = $channel.ChannelID  
   
    try {
        $teamId = (Get-TeamChannel -GroupId $channel.TeamID | Where-Object { $_.Id -eq $channelId }).GroupId
       
        $site = Get-PnPSite -GroupId $teamId -Connection $tenantAdminConn -ErrorAction Stop
        $channelFolderUrl = $site.Url + "/Shared Documents/" + (Get-TeamChannel -GroupId $teamId | Where-Object { $_.Id -eq $channelId }).DisplayName
        $results += [PSCustomObject]@{
            ChannelID           = $channelId
            SharePointFolderURL = $channelFolderUrl
        }
    }
    catch {
        $results += [PSCustomObject]@{
            ChannelID           = $channelId
            SharePointFolderURL = "Not Found or Access Denied"
        }
    }
}
 
$results | Export-Csv -Path "TeamsChannel-$Timestamp.csv" -NoTypeInformation

