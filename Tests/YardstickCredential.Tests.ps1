BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    Import-Module "$PSScriptRoot\..\Modules\YardstickCredential.psm1" -Force

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
