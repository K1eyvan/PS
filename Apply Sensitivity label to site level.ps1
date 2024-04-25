$tenantId = "ui"
$clientId = "your-client-id"
$clientSecret = "your-client-secret"
$scope = "https://graph.microsoft.com/.default"
$body = @{
    grant_type    = "client_credentials"
    scope         = $scope
    client_id     = $clientId
    client_secret = $clientSecret
}
$oauth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body
$token = $oauth.access_token

$siteUrl = "https://yourtenant.sharepoint.com/sites/yoursite"
$labelId = "your-label-id"  # Sensitivity label ID

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
