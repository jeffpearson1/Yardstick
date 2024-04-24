# CRED1 Auto SSO Sign On
function Invoke-SSOAutoSignIn {
    param(
        $Target,
        $Driver
    )
    $Username_Field = Get-SeElement -By ID "username"
    $Password_Field = Get-SeElement -By ID "password"
    $Button = Get-SeElement -By CSSSelector ".btn"
    $Credentials = Get-StoredCredential -Target "cred1"
    Invoke-SeKeys -Element $Username_Field -Keys $Credentials.username
    Invoke-SeKeys -Element $Password_Field -Keys $(ConvertFrom-SecureString $Credentials.Password -AsPlainText)
    $Button.click()
    while($Driver.SeUrl -notlike "*$Target*") {
        Write-Log "Waiting for 2FA to be approved..."
        Write-Log "URL: $($Driver.SeURL)"
        Write-Log "Target: $Target"
        Start-Sleep -Seconds 2
    }
    Start-Sleep -Seconds 10
}
