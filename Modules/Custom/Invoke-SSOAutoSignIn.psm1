# CRED1 Auto SSO Sign On
function Invoke-SSOAutoSignIn {
    param(
        $Target,
        $Driver
    )
    $Username_Field = Get-SeElement -By ID "username"
    $Password_Field = Get-SeElement -By ID "password"
    $Button = Get-SeElement -By CSSSelector ".btn"
    try {
        $Credentials = Get-StoredCredential -Target "cred1"
        Start-Sleep -Seconds 10
        Invoke-SeKeys -Element $Username_Field -Keys $Credentials.username
        Invoke-SeKeys -Element $Password_Field -Keys $(ConvertFrom-SecureString $Credentials.Password -AsPlainText)
    }
    catch {
        Write-Log "An error occurred while retrieving the stored credentials from credential manager"
        return $false
    }
    $Button.click()
    $TrustButton = Get-SeElement -By ID "trust-browser-button" -Timeout 300
    if ($TrustButton) {
        $TrustButton.click()
    }
    else {
        Write-Log "Error occurred when authenticating"
        return $false
    }
    while($Driver.SeUrl -notlike "*$Target*") {
        Write-Log "Waiting for 2FA to be approved..."
        Write-Log "URL: $($Driver.SeURL)"
        Write-Log "Target: $Target"
        Start-Sleep -Seconds 2
    }
    Start-Sleep -Seconds 10
    return $true
}
