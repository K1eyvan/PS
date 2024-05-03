# Load the used helpers in the PowerShell session
Get-ChildItem -Path ".\Functions\*.ps1" -Recurse | ForEach-Object {
    . $_.FullName
}
# Set here the common name of the certificate (Leave as is if no specific need)
$commonName = "valo-auth-cert"
# Set here the number of years the certificate should be valid
$validityYears = 100
# The certificate is valid from today
$startDate = [System.DateTime]::Today.ToString("yyyy-MM-dd")
# Until the specified number of years
$endDate = [System.DateTime]::Today.AddYears($validityYears).ToString("yyyy-MM-dd")
# Set here the password of the certificate in plain text
# Change it according to your preferences
$certPasswordPlain = "valo-password"
$certPassword = ConvertTo-SecureString -String $certPasswordPlain -Force -AsPlainText
# Create the certificate
$certGenerationResult = Create-SelfSignedCertificate -CommonName $commonName -StartDate $startDate -EndDate $endDate -Force -Password $certPassword
$cer = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cer.Import($certGenerationResult.cerPath)
Write-Host "PFX File    : $($certGenerationResult.pfxPath)"
Write-Host "CER File    : $($certGenerationResult.cerPath)"
Write-Host "Thumbprint  : $($cer.Thumbprint)"

# Open the Windows Explorer in the folder containing the certificates
explorer .\Temp\Certificate