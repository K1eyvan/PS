$ApplicationId = ""
$SecuredPassword = ""
$tenantID = ""
$ReportFolder = "C:\Users\colhug01\OneDrive - Robert Half\Documents\Scripts\Reports\"
function Get-SharePointURL {
    param (
        $TeamID,
        $ChannelID
    )
    $FolderURL = Get-MgTeamChannelFileFolder -TeamId $TeamId -ChannelId $ChannelId | Select-Object -ExpandProperty WebUrl
    $SplitURL = $FolderURL -split '/'
    Return ($SplitURL[0..4] -join '/')   
}

$SecuredPasswordPassword = ConvertTo-SecureString -String $SecuredPassword -AsPlainText -Force

$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationId, $SecuredPasswordPassword
Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential -NoWelcome

$DateString = Get-Date -Format "yyyy-MM-dd-hh-mm-ss"
$ReportPath = Join-Path -Path $ReportFolder -ChildPath "$($DateString)-TeamsChannelSites.JSON"
$ReportHashTable = @{}

$Teams = Get-MgTeam -All
Foreach ($Team in $Teams){
    $ReportHashTable[$Team.ID] = @{
        "TeamID" = $Team.Id;
        "TeamName" = $Team.DisplayName;
        "Channels" = @();
    }
    $Channels = Get-MgTeamChannel -TeamId $team.id -All
    $StandardChannelURL = $null
    Foreach ($Channel in $Channels)
    {
        If ($Channel.MembershipType -eq "standard"){
            if ($StandardChannelURL){
                $SharePointSiteURL = $StandardChannelURL
            }
            else {
                $SharePointSiteURL = Get-SharePointURL -TeamID $Team.Id -ChannelID $Channel.Id
                $StandardChannelURL = $SharePointSiteURL
            }
        }
        else {
            $SharePointSiteURL = Get-SharePointURL -TeamID $Team.Id -ChannelID $Channel.Id
        }

        $ReportHashTable[$Team.Id].Channels += @{
            "ChannelId" = $channel.Id;
            "ChannelName" = $channel.DisplayName;
            "ChannelType" = $channel.MembershipType;
            "SharePointSiteURL" = $SharePointSiteURL
        }
    }
}

$JSONContent = $ReportHashTable | ConvertTo-Json -Depth 3
$JSONContent | Set-Content -Path $ReportPath
Disconnect-Graph