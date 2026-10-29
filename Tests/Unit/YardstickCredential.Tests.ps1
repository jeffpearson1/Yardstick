BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickCredential.psm1" -Force

    # Unique per-run target so a failed test cannot clobber a real credential.
    $script:Target = "Yardstick:PesterTest-$([guid]::NewGuid().ToString('N').Substring(0,8))"
}

AfterAll {
    Remove-YardstickIntuneCredential -Target $script:Target -ErrorAction SilentlyContinue | Out-Null
}

Describe "Get-YardstickCredentialTarget" {
    It "Falls back to the default target when preferences are empty" {
        Get-YardstickCredentialTarget -Preferences @{} | Should -Be 'Yardstick:IntuneGraph'
    }

    It "Falls back to the default target when preferences are null" {
        Get-YardstickCredentialTarget -Preferences $null | Should -Be 'Yardstick:IntuneGraph'
    }

    It "Honours credentialTarget from preferences" {
        Get-YardstickCredentialTarget -Preferences @{ credentialTarget = 'Custom:Target' } | Should -Be 'Custom:Target'
    }

    It "Ignores a blank credentialTarget" {
        Get-YardstickCredentialTarget -Preferences @{ credentialTarget = '   ' } | Should -Be 'Yardstick:IntuneGraph'
    }
}

Describe "Credential storage round-trip" {
    AfterEach {
        Remove-YardstickIntuneCredential -Target $script:Target -ErrorAction SilentlyContinue | Out-Null
    }

    It "Returns null when no credential is stored" {
        Get-YardstickIntuneCredential -Target $script:Target | Should -BeNullOrEmpty
    }

    It "Round-trips tenant, client and secret" {
        Set-YardstickIntuneCredential -Target $script:Target -TenantID 'tenant-1' -ClientID 'client-1' -ClientSecret 'secret-1'
        $cred = Get-YardstickIntuneCredential -Target $script:Target
        $cred.TenantID | Should -Be 'tenant-1'
        $cred.ClientID | Should -Be 'client-1'
        $cred.ClientSecret | Should -Be 'secret-1'
        $cred.Target | Should -Be $script:Target
    }

    It "Accepts a SecureString secret" {
        $secure = ConvertTo-SecureString 'secure-secret' -AsPlainText -Force
        Set-YardstickIntuneCredential -Target $script:Target -TenantID 't' -ClientID 'c' -ClientSecret $secure
        (Get-YardstickIntuneCredential -Target $script:Target).ClientSecret | Should -Be 'secure-secret'
    }

    It "Round-trips the expiration date without timezone drift" {
        $expires = [datetime]::new(2030, 5, 17, 9, 30, 0, [System.DateTimeKind]::Local)
        Set-YardstickIntuneCredential -Target $script:Target -TenantID 't' -ClientID 'c' -ClientSecret 's' -SecretExpiresOn $expires
        (Get-YardstickIntuneCredential -Target $script:Target).SecretExpiresOn | Should -Be $expires
    }

    It "Leaves the expiration null when it was never supplied" {
        Set-YardstickIntuneCredential -Target $script:Target -TenantID 't' -ClientID 'c' -ClientSecret 's'
        (Get-YardstickIntuneCredential -Target $script:Target).SecretExpiresOn | Should -BeNullOrEmpty
    }

    It "Rejects an empty secret" {
        { Set-YardstickIntuneCredential -Target $script:Target -TenantID 't' -ClientID 'c' -ClientSecret '' } | Should -Throw
    }

    It "Reports whether a credential was actually deleted" {
        Set-YardstickIntuneCredential -Target $script:Target -TenantID 't' -ClientID 'c' -ClientSecret 's'
        Remove-YardstickIntuneCredential -Target $script:Target | Should -BeTrue
        Remove-YardstickIntuneCredential -Target $script:Target | Should -BeFalse
    }
}

Describe "Initialize-YardstickIntuneCredential" {
    AfterEach {
        Remove-YardstickIntuneCredential -Target $script:Target -ErrorAction SilentlyContinue | Out-Null
    }

    It "Migrates legacy preferences.yaml values into Credential Manager" {
        $prefs = @{ credentialTarget = $script:Target; TenantID = 'legacy-t'; ClientID = 'legacy-c'; ClientSecret = 'legacy-s' }
        $cred = Initialize-YardstickIntuneCredential -Preferences $prefs -NoPrompt

        $cred.TenantID | Should -Be 'legacy-t'
        (Get-YardstickIntuneCredential -Target $script:Target).ClientSecret | Should -Be 'legacy-s'
        $Global:TenantID | Should -Be 'legacy-t'
        $Global:ClientID | Should -Be 'legacy-c'
        $Global:ClientSecret | Should -Be 'legacy-s'
    }

    It "Prefers the stored credential over legacy preference values" {
        Set-YardstickIntuneCredential -Target $script:Target -TenantID 'stored-t' -ClientID 'stored-c' -ClientSecret 'stored-s'
        $prefs = @{ credentialTarget = $script:Target; TenantID = 'legacy-t'; ClientID = 'legacy-c'; ClientSecret = 'legacy-s' }

        (Initialize-YardstickIntuneCredential -Preferences $prefs -NoPrompt -WarningAction SilentlyContinue).ClientSecret | Should -Be 'stored-s'
    }

    It "Throws instead of prompting when nothing is stored and -NoPrompt is used" {
        $prefs = @{ credentialTarget = $script:Target }
        { Initialize-YardstickIntuneCredential -Preferences $prefs -NoPrompt } |
            Should -Throw -ExpectedMessage "*Set-YardstickCredential.ps1*"
    }
}

Describe "Test-YardstickSecretExpiration" {
    It "Reports no warning for a secret with plenty of life left" {
        $cred = [PSCustomObject]@{ Target = $script:Target; TenantID = 't'; ClientID = 'c'; ClientSecret = 's'; SecretExpiresOn = (Get-Date).AddDays(90); LastNotifiedOn = $null }
        $result = Test-YardstickSecretExpiration -Credential $cred -Preferences @{ credentialExpirationWarningDays = 30 } -NoEmail
        $result.IsExpiring | Should -BeFalse
        $result.IsExpired | Should -BeFalse
        $result.DaysRemaining | Should -Be 90
    }

    It "Flags a secret inside the warning window" {
        $cred = [PSCustomObject]@{ Target = $script:Target; TenantID = 't'; ClientID = 'c'; ClientSecret = 's'; SecretExpiresOn = (Get-Date).AddDays(5); LastNotifiedOn = $null }
        $result = Test-YardstickSecretExpiration -Credential $cred -Preferences @{ credentialExpirationWarningDays = 30 } -NoEmail -WarningAction SilentlyContinue
        $result.IsExpiring | Should -BeTrue
        $result.IsExpired | Should -BeFalse
    }

    It "Flags an already-expired secret" {
        $cred = [PSCustomObject]@{ Target = $script:Target; TenantID = 't'; ClientID = 'c'; ClientSecret = 's'; SecretExpiresOn = (Get-Date).AddDays(-2); LastNotifiedOn = $null }
        $result = Test-YardstickSecretExpiration -Credential $cred -Preferences @{} -NoEmail -WarningAction SilentlyContinue
        $result.IsExpired | Should -BeTrue
        $result.IsExpiring | Should -BeTrue
    }

    It "Defaults the warning window to 30 days" {
        $cred = [PSCustomObject]@{ Target = $script:Target; TenantID = 't'; ClientID = 'c'; ClientSecret = 's'; SecretExpiresOn = (Get-Date).AddDays(29); LastNotifiedOn = $null }
        (Test-YardstickSecretExpiration -Credential $cred -Preferences @{} -NoEmail -WarningAction SilentlyContinue).IsExpiring | Should -BeTrue

        $cred.SecretExpiresOn = (Get-Date).AddDays(31)
        (Test-YardstickSecretExpiration -Credential $cred -Preferences @{} -NoEmail).IsExpiring | Should -BeFalse
    }

    It "Warns but does not fail when the expiration is unknown" {
        $cred = [PSCustomObject]@{ Target = $script:Target; TenantID = 't'; ClientID = 'c'; ClientSecret = 's'; SecretExpiresOn = $null; LastNotifiedOn = $null }
        $result = Test-YardstickSecretExpiration -Credential $cred -Preferences @{} -NoEmail -WarningAction SilentlyContinue
        $result.DaysRemaining | Should -BeNullOrEmpty
        $result.IsExpiring | Should -BeFalse
    }
}

Describe "Get-YardstickAdminEmailRecipient" {
    It "Prefers adminEmailRecipient" {
        Get-YardstickAdminEmailRecipient -Preferences @{ adminEmailRecipient = 'admin@x.com'; emailRecipient = 'reports@x.com' } |
            Should -Be @('admin@x.com')
    }

    It "Falls back to emailRecipient" {
        Get-YardstickAdminEmailRecipient -Preferences @{ emailRecipient = 'reports@x.com' } | Should -Be @('reports@x.com')
    }

    It "Returns nothing when neither is configured" {
        Get-YardstickAdminEmailRecipient -Preferences @{} | Should -BeNullOrEmpty
    }
}

Describe "Send-YardstickCredentialExpiryEmail" {
    BeforeAll {
        $script:cred = [PSCustomObject]@{ Target = 't'; TenantID = 't'; ClientID = 'c'; ClientSecret = 's'; SecretExpiresOn = (Get-Date).AddDays(5); LastNotifiedOn = $null }
    }

    It "Does nothing when email notifications are disabled" {
        Send-YardstickCredentialExpiryEmail -Preferences @{ emailNotificationEnabled = $false; adminEmailRecipient = 'admin@x.com' } -Credential $script:cred -DaysRemaining 5 |
            Should -BeFalse
    }

    It "Does nothing when no recipient is configured" {
        Send-YardstickCredentialExpiryEmail -Preferences @{ emailNotificationEnabled = $true } -Credential $script:cred -DaysRemaining 5 -WarningAction SilentlyContinue |
            Should -BeFalse
    }

    It "Honours credentialExpirationEmailEnabled over emailNotificationEnabled" {
        Send-YardstickCredentialExpiryEmail -Preferences @{ emailNotificationEnabled = $true; credentialExpirationEmailEnabled = $false; adminEmailRecipient = 'admin@x.com' } -Credential $script:cred -DaysRemaining 5 |
            Should -BeFalse
    }
}

Describe "Get-YardstickSmtpCredentialTarget" {
    It "Falls back to the default target when preferences are empty" {
        Get-YardstickSmtpCredentialTarget -Preferences @{} | Should -Be 'Yardstick:Smtp'
    }

    It "Falls back to the default target when preferences are null" {
        Get-YardstickSmtpCredentialTarget -Preferences $null | Should -Be 'Yardstick:Smtp'
    }

    It "Honours smtpCredentialTarget from preferences" {
        Get-YardstickSmtpCredentialTarget -Preferences @{ smtpCredentialTarget = 'Custom:Smtp' } | Should -Be 'Custom:Smtp'
    }

    It "Ignores a blank smtpCredentialTarget" {
        Get-YardstickSmtpCredentialTarget -Preferences @{ smtpCredentialTarget = '   ' } | Should -Be 'Yardstick:Smtp'
    }
}

Describe "SMTP credential storage round-trip" {
    BeforeAll {
        $script:SmtpTarget = "Yardstick:PesterSmtp-$([guid]::NewGuid().ToString('N').Substring(0,8))"
    }

    AfterEach {
        Remove-YardstickSmtpCredential -Target $script:SmtpTarget -ErrorAction SilentlyContinue | Out-Null
    }

    It "Returns null when no credential is stored" {
        Get-YardstickSmtpCredential -Target $script:SmtpTarget | Should -BeNullOrEmpty
    }

    It "Round-trips the user name and password as a PSCredential" {
        Set-YardstickSmtpCredential -Target $script:SmtpTarget -UserName 'smtp-user' -Password 'smtp-pass'
        $cred = Get-YardstickSmtpCredential -Target $script:SmtpTarget
        $cred | Should -BeOfType [System.Management.Automation.PSCredential]
        $cred.UserName | Should -Be 'smtp-user'
        $cred.GetNetworkCredential().Password | Should -Be 'smtp-pass'
    }

    It "Accepts a SecureString password" {
        $secure = ConvertTo-SecureString 'secure-pass' -AsPlainText -Force
        Set-YardstickSmtpCredential -Target $script:SmtpTarget -UserName 'u' -Password $secure
        (Get-YardstickSmtpCredential -Target $script:SmtpTarget).GetNetworkCredential().Password | Should -Be 'secure-pass'
    }

    It "Rejects an empty password" {
        { Set-YardstickSmtpCredential -Target $script:SmtpTarget -UserName 'u' -Password '' } | Should -Throw
    }

    It "Reports whether a credential was removed" {
        Set-YardstickSmtpCredential -Target $script:SmtpTarget -UserName 'u' -Password 'p'
        Remove-YardstickSmtpCredential -Target $script:SmtpTarget | Should -BeTrue
        Remove-YardstickSmtpCredential -Target $script:SmtpTarget | Should -BeFalse
    }
}

Describe "GitHub token storage round-trip" {
    BeforeAll {
        $script:GitHubTarget = "Yardstick:PesterGitHub-$([guid]::NewGuid().ToString('N').Substring(0,8))"
    }

    AfterEach {
        Remove-YardstickGitHubToken -Target $script:GitHubTarget -ErrorAction SilentlyContinue | Out-Null
    }

    It "Returns null when no token is stored" {
        Get-YardstickGitHubToken -Target $script:GitHubTarget | Should -BeNullOrEmpty
    }

    It "Round-trips the token, trimmed" {
        Set-YardstickGitHubToken -Target $script:GitHubTarget -Token "  github_pat_11ABC  "
        Get-YardstickGitHubToken -Target $script:GitHubTarget | Should -BeExactly 'github_pat_11ABC'
    }

    It "Accepts a SecureString token" {
        $secure = ConvertTo-SecureString 'github_pat_secure' -AsPlainText -Force
        Set-YardstickGitHubToken -Target $script:GitHubTarget -Token $secure
        Get-YardstickGitHubToken -Target $script:GitHubTarget | Should -BeExactly 'github_pat_secure'
    }

    It "Rejects an empty token" {
        { Set-YardstickGitHubToken -Target $script:GitHubTarget -Token '  ' } | Should -Throw
    }

    It "Reports whether a token was removed" {
        Set-YardstickGitHubToken -Target $script:GitHubTarget -Token 'github_pat_x'
        Remove-YardstickGitHubToken -Target $script:GitHubTarget | Should -BeTrue
        Remove-YardstickGitHubToken -Target $script:GitHubTarget | Should -BeFalse
    }
}

Describe "Send-YardstickCredentialExpiryEmail over SMTP" {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\..\Modules\YardstickSupport.psm1" -Global -Force
        $script:cred = [PSCustomObject]@{ Target = 't'; TenantID = 't'; ClientID = 'c'; ClientSecret = 's'; SecretExpiresOn = (Get-Date).AddDays(5); LastNotifiedOn = $null }
        $script:smtpPrefs = @{ emailNotificationEnabled = $true; emailDeliveryMethod = 'smtp'; adminEmailRecipient = 'admin@x.com' }
    }

    It "Sends through Send-YardstickSmtpMessage instead of Outlook" {
        Mock Send-YardstickSmtpMessage {} -ModuleName YardstickCredential
        Mock Test-OutlookAvailability { $true } -ModuleName YardstickCredential

        Send-YardstickCredentialExpiryEmail -Preferences $script:smtpPrefs -Credential $script:cred -DaysRemaining 5 |
            Should -BeTrue

        Should -Invoke Send-YardstickSmtpMessage -ModuleName YardstickCredential -Times 1 -Exactly -ParameterFilter {
            $To -contains 'admin@x.com' -and $Subject -like 'Yardstick: *'
        }
        Should -Invoke Test-OutlookAvailability -ModuleName YardstickCredential -Times 0 -Exactly
    }

    It "Returns false when the SMTP send fails" {
        Mock Send-YardstickSmtpMessage { throw 'relay denied' } -ModuleName YardstickCredential

        Send-YardstickCredentialExpiryEmail -Preferences $script:smtpPrefs -Credential $script:cred -DaysRemaining 5 -WarningAction SilentlyContinue |
            Should -BeFalse
    }
}

Describe "Secret vault storage round-trip" {
    BeforeAll {
        $script:Vault = "PesterVault-$([guid]::NewGuid().ToString('N').Substring(0,8))"
    }

    AfterEach {
        Remove-YardstickSecretVault -VaultName $script:Vault -ErrorAction SilentlyContinue | Out-Null
    }

    It "Names the target after the vault" {
        Get-YardstickSecretVaultTarget -VaultName 'crowdstrikefalcon' | Should -Be 'Yardstick:Secret:crowdstrikefalcon'
    }

    It "Returns null when the vault is not stored" {
        Get-YardstickSecretVault -VaultName $script:Vault | Should -BeNullOrEmpty
    }

    It "Round-trips every secret as a hashtable" {
        $secure = ConvertTo-SecureString 'secure-value' -AsPlainText -Force
        Set-YardstickSecretVault -VaultName $script:Vault -Secrets @{ client_id = 'id-1'; client_secret = $secure }
        $vault = Get-YardstickSecretVault -VaultName $script:Vault
        $vault | Should -BeOfType [hashtable]
        $vault.client_id | Should -Be 'id-1'
        $vault.client_secret | Should -Be 'secure-value'
    }

    It "Rejects an empty vault" {
        { Set-YardstickSecretVault -VaultName $script:Vault -Secrets @{} } | Should -Throw
    }

    It "Reports whether a vault was removed" {
        Set-YardstickSecretVault -VaultName $script:Vault -Secrets @{ a = 'b' }
        Remove-YardstickSecretVault -VaultName $script:Vault | Should -BeTrue
        Remove-YardstickSecretVault -VaultName $script:Vault | Should -BeFalse
    }
}

Describe "Get-Secrets" {
    BeforeAll {
        Import-Module powershell-yaml -ErrorAction Stop
        $script:Vault = "PesterVault-$([guid]::NewGuid().ToString('N').Substring(0,8))"
    }

    AfterEach {
        Remove-YardstickSecretVault -VaultName $script:Vault -ErrorAction SilentlyContinue | Out-Null
        Remove-Item -Path (Join-Path $TestDrive '*.yaml') -ErrorAction SilentlyContinue
    }

    It "Reads the vault from Credential Manager" {
        Set-YardstickSecretVault -VaultName $script:Vault -Secrets @{ customer_id = 'cid-1' }
        (Get-Secrets -VaultName $script:Vault -SecretsDir $TestDrive).customer_id | Should -Be 'cid-1'
    }

    It "Prefers Credential Manager over a leftover YAML file" {
        Set-YardstickSecretVault -VaultName $script:Vault -Secrets @{ customer_id = 'stored' }
        Set-Content -Path (Join-Path $TestDrive "$script:Vault.yaml") -Value 'customer_id: legacy'
        (Get-Secrets -VaultName $script:Vault -SecretsDir $TestDrive).customer_id | Should -Be 'stored'
    }

    It "Moves a legacy YAML vault into Credential Manager" {
        Set-Content -Path (Join-Path $TestDrive "$script:Vault.yaml") -Value "client_id: legacy-id`nclient_secret: legacy-secret"
        $secrets = Get-Secrets -VaultName $script:Vault -SecretsDir $TestDrive -WarningAction SilentlyContinue
        $secrets.client_secret | Should -Be 'legacy-secret'
        (Get-YardstickSecretVault -VaultName $script:Vault).client_id | Should -Be 'legacy-id'
    }

    It "Throws with the setup command when the vault is stored nowhere" {
        { Get-Secrets -VaultName $script:Vault -SecretsDir $TestDrive } |
            Should -Throw -ExpectedMessage "*Set-YardstickCredential.ps1 -Secret $script:Vault*"
    }
}
