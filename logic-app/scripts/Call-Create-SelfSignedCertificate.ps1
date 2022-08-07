<#
    .SYNOPSIS
        Script to create a self-signed certificate.

    .DESCRIPTION
        This script allows you to create a self-signed certificate and its private key information file.
        The certificate can be used to authenticate identities, like Azure app registrations.
        The script also automatically removes the created certificate from the user's keystore.

    .PARAMETER certificateName <string> [required]
        The name of the certificate.

    .PARAMETER certificatePwd <string> [required]
        The password used for certificate encryption.

    .PARAMETER monthsValid <int> [required]
        The number of months before the certificate becomes invalid.

    .PARAMETER folderPath <string> [required]
        The (relative) folder path the certificate should be created in.
#>

Remove-Module -Name Create-SelfSignedCertificate -ErrorAction SilentlyContinue -Force
Import-Module .\modules\Create-SelfSignedCertificate.psm1 -ErrorAction Stop -WarningAction SilentlyContinue

$certificateName = "prod"
$certificatePwd = ""
$monthsValid = 36
$folderPath = ""

Create-SelfSignedCertificate -certificateName $certificateName -certificatePwd $certificatePwd -monthsValid $monthsValid -folderPath $folderPath
