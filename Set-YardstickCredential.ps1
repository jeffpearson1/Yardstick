<#
.SYNOPSIS
    Stores, inspects, or removes the Intune app-registration credentials Yardstick
    uses to authenticate to Microsoft Graph, the optional SMTP credential, the
    optional GitHub API token, and the secret vaults recipes read with Get-Secrets.

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

.PARAMETER GitHub
    Work with the GitHub API token instead of the Intune credential. Combine with
    -Show (which also asks GitHub how many calls are left) or -Remove. Optional:
    without a token, GitHub allows 60 API calls an hour, about one full run. Use a
    fine-grained personal access token with read-only access to public repositories
    and no permissions. Stored under "Yardstick:GitHub".

.PARAMETER Secret
    Work with the named recipe secret vault (the one a recipe reads with
    Get-Secrets -VaultName) instead of the Intune credential. Prompts for secret
    names and values; combine with -Show to list the stored names or -Remove to
    delete the vault. Vaults are stored under "Yardstick:Secret:<name>".

.PARAMETER Target
    Override the Credential Manager target name. Defaults to the `credentialTarget`
    preference, or "Yardstick:IntuneGraph" (with -Smtp: the `smtpCredentialTarget`
    preference, or "Yardstick:Smtp"). Not used with -Secret or -GitHub.

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

.EXAMPLE
    .\Set-YardstickCredential.ps1 -GitHub
    Prompts for a GitHub API token, stores it securely, and checks it with GitHub.

.EXAMPLE
    .\Set-YardstickCredential.ps1 -Secret crowdstrikefalcon
    Prompts for the secrets the crowdstrikefalcon recipe reads, and stores them securely.
#>
#Requires -Version 7.0

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

    [Parameter(ParameterSetName = 'Set')]
    [Parameter(ParameterSetName = 'Show')]
    [Parameter(ParameterSetName = 'Remove')]
    [switch]$GitHub,

    [Parameter(ParameterSetName = 'Set')]
    [Parameter(ParameterSetName = 'Show')]
    [Parameter(ParameterSetName = 'Remove')]
    [ValidateNotNullOrEmpty()]
    [string]$Secret,

    [string]$Target
)

$ErrorActionPreference = 'Stop'

$Global:LogLocation = $PSScriptRoot
$Global:LogFile = 'YLog.log'

Import-Module powershell-yaml -ErrorAction Stop
Import-Module "$PSScriptRoot\Yardstick.psd1" -Force

if (@($Smtp, $GitHub, [bool]$Secret | Where-Object { $_ }).Count -gt 1) {
    throw "-Smtp, -GitHub and -Secret cannot be used together."
}

if ($Secret) {
    switch ($PSCmdlet.ParameterSetName) {
        'Show' {
            $vault = Get-YardstickSecretVault -VaultName $Secret
            if ($null -eq $vault) {
                Write-Host "No secret vault '$Secret' is stored under target '$(Get-YardstickSecretVaultTarget -VaultName $Secret)'." -ForegroundColor Yellow
                exit 1
            }
            [PSCustomObject]@{
                Target  = Get-YardstickSecretVaultTarget -VaultName $Secret
                Secrets = (@($vault.Keys | Sort-Object) | ForEach-Object { "$_ = ********" }) -join "`n"
            } | Format-List
        }

        'Remove' {
            Remove-YardstickSecretVault -VaultName $Secret | Out-Null
        }

        default {
            Register-YardstickSecretVault -VaultName $Secret | Out-Null
            Write-Host "Secret vault '$Secret' stored." -ForegroundColor Green
        }
    }
    return
}

if ($GitHub) {
    switch ($PSCmdlet.ParameterSetName) {
        'Show' {
            $token = Get-YardstickGitHubToken
            if (-not $token) {
                Write-Host "No GitHub API token is stored, so Yardstick calls GitHub anonymously (60 calls an hour)." -ForegroundColor Yellow
                exit 1
            }
            $rateLimit = $null
            try {
                $rateLimit = Get-YardstickGitHubRateLimit
            } catch {
                Write-Warning "Could not check the token with GitHub: $($_.Exception.Message)"
            }
            [PSCustomObject]@{
                Token         = "********$($token.Substring([math]::Max(0, $token.Length - 4)))"
                Authenticated = if ($rateLimit) { $rateLimit.Authenticated } else { '(unknown)' }
                Remaining     = if ($rateLimit) { "$($rateLimit.Remaining) of $($rateLimit.Limit)" } else { '(unknown)' }
                Resets        = if ($rateLimit -and $rateLimit.Reset) { $rateLimit.Reset.ToString('HH:mm') } else { '(unknown)' }
            } | Format-List
        }

        'Remove' {
            Remove-YardstickGitHubToken | Out-Null
        }

        default {
            Register-YardstickGitHubToken | Out-Null
            Write-Host "GitHub API token stored. Checking it with GitHub..." -ForegroundColor Cyan
            try {
                $rateLimit = Get-YardstickGitHubRateLimit
            } catch {
                Write-Warning "The token was stored, but GitHub could not be reached to check it: $($_.Exception.Message)"
                exit 1
            }
            if (-not $rateLimit.Authenticated) {
                Write-Warning "The token was stored, but GitHub rejected it. Check it was copied in full and has not expired."
                exit 1
            }
            Write-Host "GitHub accepted the token: $($rateLimit.Remaining) of $($rateLimit.Limit) calls left this hour." -ForegroundColor Green
        }
    }
    return
}

# Credentials can be stored before preferences.yaml exists, so a missing file
# just means the built-in defaults (credential target names) apply.
$Prefs = Import-YardstickPreferences -ProjectRoot $PSScriptRoot -AllowMissing

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
