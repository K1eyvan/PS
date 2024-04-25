$tenantId = "544-44baff944166"
$clientId = "1333d3504e58e7a"
$clientSecret = "CE~8Q~gXSWbHNhUcSV~Ahy.2HOqCEsdhl"
$scope = "https://graph.microsoft.com/.default"
$body = @{
    grant_type    = "client_credentials"
    scope         = $scope
    client_id     = $clientId
    client_secret = $clientSecret
}
$oauth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body
$token = $oauth.access_token

$siteUrl = "https://k1eyvan.sharepoint.com/sites/APProval"
$labelId = "614a3022-85c5-41b0-8544-44baff944166"  # Sensitivity label ID

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

$body = @{
    sensitivityLabel = @{
        labelId = $labelId
    }
} | ConvertTo-Json

$graphEndpoint = "https://graph.microsoft.com/v1.0/sites/root"
$response = Invoke-RestMethod -Uri $graphEndpoint -Headers $headers -Method Patch -Body $body
