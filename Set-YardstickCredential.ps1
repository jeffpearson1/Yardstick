<#
.SYNOPSIS
    Stores, inspects, or removes the Intune app-registration credentials Yardstick
    uses to authenticate to Microsoft Graph.

.DESCRIPTION
    Yardstick reads its TenantID, ClientID and ClientSecret from Windows Credential
    Manager rather than from preferences.yaml. This script is the supported way to
    populate or rotate them. The client secret is collected as a SecureString and is
    never echoed, logged, or written to disk in plaintext.

    Run with no parameters to enter (or re-enter) the credentials interactively.

.PARAMETER Show
    Display the stored credential metadata (tenant, client, expiration). The secret
    itself is masked.

.PARAMETER Remove
    Delete the stored credential from Windows Credential Manager.

.PARAMETER RefreshExpiration
    Connect to Graph and refresh the recorded client secret expiration date. Requires
    the app registration to hold Application.Read.All.

.PARAMETER Smtp
    Work with the SMTP account credential used when emailDeliveryMethod is 'smtp',
    instead of the Intune credential. Combine with -Show or -Remove. Only needed when
    the SMTP server requires authentication; without it Yardstick relays anonymously.

.PARAMETER Target
    Override the Credential Manager target name. Defaults to the `credentialTarget`
    preference, or "Yardstick:IntuneGraph" (with -Smtp: the `smtpCredentialTarget`
    preference, or "Yardstick:Smtp").

.EXAMPLE
    .\Set-YardstickCredential.ps1
    Prompts for tenant, client and secret, and stores them securely.

.EXAMPLE
    .\Set-YardstickCredential.ps1 -Show

.EXAMPLE
    .\Set-YardstickCredential.ps1 -RefreshExpiration

.EXAMPLE
    .\Set-YardstickCredential.ps1 -Smtp
    Prompts for the SMTP user name and password, and stores them securely.
#>
[CmdletBinding(DefaultParameterSetName = 'Set')]
param(
    [Parameter(ParameterSetName = 'Show')]
    [switch]$Show,

    [Parameter(ParameterSetName = 'Remove')]
    [switch]$Remove,

    [Parameter(ParameterSetName = 'Refresh')]
    [switch]$RefreshExpiration,

    [Parameter(ParameterSetName = 'Set')]
    [Parameter(ParameterSetName = 'Show')]
    [Parameter(ParameterSetName = 'Remove')]
    [switch]$Smtp,

    [string]$Target
)

$ErrorActionPreference = 'Stop'

$Global:LogLocation = $PSScriptRoot
$Global:LogFile = 'YLog.log'

Import-Module powershell-yaml -ErrorAction Stop
Import-Module "$PSScriptRoot\Modules\YardstickGraph.psm1" -Scope Global -Force
Import-Module "$PSScriptRoot\Modules\YardstickIntune.psm1" -Scope Global -Force
Import-Module "$PSScriptRoot\Modules\YardstickSupport.psm1" -Scope Global -Force
Import-Module "$PSScriptRoot\Modules\YardstickCredential.psm1" -Scope Global -Force

$Prefs = $null
. (Join-Path $PSScriptRoot 'ProjectLayout.ps1')
$prefsPath = Get-YardstickPreferencesPath -ProjectRoot $PSScriptRoot
if (Test-Path $prefsPath) {
    $Prefs = Get-Content -LiteralPath $prefsPath | ConvertFrom-Yaml
}

if ($Smtp) {
    if (-not $Target) {
        $Target = Get-YardstickSmtpCredentialTarget -Preferences $Prefs
    }

    switch ($PSCmdlet.ParameterSetName) {
        'Show' {
            $credential = Get-YardstickSmtpCredential -Target $Target
            if (-not $credential) {
                Write-Host "No SMTP credentials are stored under target '$Target'." -ForegroundColor Yellow
                exit 1
            }
            [PSCustomObject]@{
                Target   = $Target
                UserName = $credential.UserName
                Password = '********'
            } | Format-List
        }

        'Remove' {
            Remove-YardstickSmtpCredential -Target $Target | Out-Null
        }

        default {
            Register-YardstickSmtpCredential -Target $Target | Out-Null
            Write-Host "SMTP credentials stored. Run .\Test-EmailNotification.ps1 -Send to verify delivery." -ForegroundColor Green
        }
    }
    return
}

if (-not $Target) {
    $Target = Get-YardstickCredentialTarget -Preferences $Prefs
}

switch ($PSCmdlet.ParameterSetName) {
    'Show' {
        $credential = Get-YardstickIntuneCredential -Target $Target
        if (-not $credential) {
            Write-Host "No Yardstick credentials are stored under target '$Target'." -ForegroundColor Yellow
            exit 1
        }
        [PSCustomObject]@{
            Target          = $credential.Target
            TenantID        = $credential.TenantID
            ClientID        = $credential.ClientID
            ClientSecret    = '********'
            SecretExpiresOn = if ($credential.SecretExpiresOn) { $credential.SecretExpiresOn.ToString('yyyy-MM-dd') } else { '(unknown)' }
            DaysRemaining   = if ($credential.SecretExpiresOn) { ($credential.SecretExpiresOn.Date - (Get-Date).Date).Days } else { $null }
        } | Format-List
    }

    'Remove' {
        Remove-YardstickIntuneCredential -Target $Target | Out-Null
    }

    'Refresh' {
        $credential = Initialize-YardstickIntuneCredential -Preferences $Prefs
        Connect-YardstickGraph -Force
        $credential = Update-YardstickSecretExpiration -Credential $credential
        Test-YardstickSecretExpiration -Credential $credential -Preferences $Prefs -NoEmail | Format-List
    }

    default {
        $credential = Register-YardstickIntuneCredential -Target $Target
        Write-Host "Credentials stored. Validating against Microsoft Graph..." -ForegroundColor Cyan
        try {
                $Global:TenantID = $credential.TenantID
            $Global:ClientID = $credential.ClientID
            $Global:ClientSecret = $credential.ClientSecret
            Connect-YardstickGraph -Force
            Write-Host "Authentication succeeded." -ForegroundColor Green
            $credential = Update-YardstickSecretExpiration -Credential $credential
            Test-YardstickSecretExpiration -Credential $credential -Preferences $Prefs -NoEmail | Out-Null
        } catch {
            Write-Warning "Credentials were stored, but authentication failed: $($_.Exception.Message)"
            exit 1
        }
    }
}
