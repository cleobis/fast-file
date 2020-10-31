# Script to generate the self-signed certificate for use on GitHub actions.
# After running, you can view the certificate by running `mmc`, File > Add/Remove Snap-in... 
# and choosing certificates for the current user. The certificate will appear under Personal 
# > Certificates. Copy the base-64 encoded output from the console into the GitHub Actions
# secret.
#
# Set the inputPwd variable before running. 
$expiry = Get-Date
$expiry = $expiry.AddYears(3)

$path = "cert:\CurrentUser\My"
$cert = New-SelfSignedCertificate -CertStoreLocation $path `
    -Subject "github.com/cleobis/fast-file" -Type CodeSigningCert `
    -notafter $expiry

if (!(Test-Path variable:inputPwd)) {
    throw "Password not set."
}
$pwd = ConvertTo-SecureString -String $inputPwd -Force -AsPlainText
$path = $path + "\" + $cert.thumbprint
Export-PfxCertificate -cert $path -FilePath "fast-file.pfx" -Password $pwd

$content = get-content "fast-file.pfx" -Encoding Byte
$base64 = [System.Convert]::ToBase64String($content)

$base64