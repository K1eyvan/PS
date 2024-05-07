# Required Modules
Install-Module -Name ImportExcel -Scope CurrentUser -Force
Install-Module -Name MSAL.PS -Scope CurrentUser -Force

# Constants
$tenantId = "your-tenant-id"
$clientId = "your-application-id"
$excelPath = "C:\path\to\your\file.xlsx"
$sheetName = "Sheet1"
$certThumbprint = "YOUR_CERT_THUMBPRINT"

# Authenticate and Acquire Token using Certificate
$cert = Get-ChildItem -Path Cert:\CurrentUser\My\ | Where-Object {$_.Thumbprint -eq $certThumbprint}
$tokenRequest = @{
    ClientId     = $clientId
    TenantId     = $tenantId
    Certificate  = $cert
    Scope        = "https://graph.microsoft.com/.default"
}
$authToken = Get-MsalToken @tokenRequest

# Read Excel File
$teamsData = Import-Excel -Path $excelPath -WorksheetName $sheetName

# Add SharePointURL column if not exists
if (-not $teamsData.psobject.properties.name -contains "SharePointURL") {
    $teamsData | Add-Member -MemberType NoteProperty -Name "SharePointURL" -Value $null
}

# Get SharePoint URLs and update Excel
foreach ($team in $teamsData) {
    $teamId = $team.TeamID
    if ($teamId) {
        $uri = "https://graph.microsoft.com/v1.0/teams/$teamId"
        $headers = @{
            Authorization = "Bearer $($authToken.AccessToken)"
            Accept        = "application/json"
        }
        try {
            $teamDetails = Invoke-RestMethod -Uri $uri -Headers $headers
            $team.SharePointURL = $teamDetails.webUrl
        } catch {
            Write-Warning "Failed to retrieve SharePoint URL for Team ID: $teamId"
            $team.SharePointURL = "Error retrieving URL"
        }
    }
}

# Save updated data back to Excel
$teamsData | Export-Excel -Path $excelPath -WorksheetName $sheetName -KillExcel -AutoSize

Write-Output "Excel file has been updated with SharePoint URLs."
