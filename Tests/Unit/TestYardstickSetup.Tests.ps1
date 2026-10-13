BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $script:repoRoot = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
    Import-Module powershell-yaml -ErrorAction Stop

    # A minimal Yardstick install in TestDrive, so the real preferences file,
    # scheduled task and Intune credential on this machine are never consulted.
    # The repo's own script and modules check it, so no second copy of a module
    # is ever loaded.
    $script:root = Join-Path $TestDrive 'Yardstick'
    New-Item -ItemType Directory -Path (Join-Path $script:root 'Recipes') -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $script:repoRoot 'Yardstick.Project.psd1') -Destination $script:root
    Set-Content -LiteralPath (Join-Path $script:root 'Recipes\demo.yaml') -Value "id: demo`ndisplayName: Demo App`niconFile: demo.png"
    $script:prefsPath = Join-Path $script:root 'Local\preferences.yaml'

    # Mocks need the commands to exist already. The script re-imports these, but a
    # Pester mock is an alias, which wins over the re-imported functions.
    foreach ($module in 'YardstickSupport', 'YardstickCredential', 'YardstickConfig', 'YardstickSetup', 'YardstickIntune') {
        Import-Module (Join-Path $script:repoRoot "Modules\$module.psm1") -Force
    }

    function Set-TestPreferences {
        param([string]$Text = "defaultOwner: 'IT'")
        New-Item -ItemType Directory -Path (Split-Path $script:prefsPath) -Force | Out-Null
        Set-Content -LiteralPath $script:prefsPath -Value $Text
    }

    function Invoke-HealthCheck {
        param([switch]$Offline)
        $results = & (Join-Path $script:repoRoot 'Test-YardstickSetup.ps1') -ProjectRoot $script:root -PassThru -Offline:$Offline 6>$null
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Results = @($results) }
    }
}

AfterAll {
    Remove-Variable -Name YardstickTestSecretExpiresOn -Scope Global -ErrorAction SilentlyContinue
    # The script imports the modules into the global session state while this file
    # imported them into Pester's, so two copies of each are loaded now. Unload both
    # or Mock -ModuleName in the test files that run next refuses to choose.
    Get-Module -Name 'YardstickSupport', 'YardstickCredential', 'YardstickConfig', 'YardstickSetup', 'YardstickIntune', 'YardstickGraph' -All |
        Remove-Module -Force -ErrorAction SilentlyContinue
}

Describe "Test-YardstickSetup.ps1" {
    BeforeEach {
        Remove-Item -LiteralPath $script:prefsPath -ErrorAction SilentlyContinue
        # Mock bodies run inside the health check, where $script: means its scope.
        $Global:YardstickTestSecretExpiresOn = (Get-Date).AddDays(200)
        Mock Get-YardstickIntuneCredential {
            [pscustomobject]@{ Target = 'Yardstick:IntuneGraph'; TenantID = 'tenant'; ClientID = 'client'; ClientSecret = 'secret'; SecretExpiresOn = $Global:YardstickTestSecretExpiresOn; LastNotifiedOn = $null }
        }
        Mock Get-YardstickAppTokenRoles { @('DeviceManagementApps.ReadWrite.All', 'DeviceManagementConfiguration.Read.All', 'DeviceManagementRBAC.Read.All') }
        Mock Invoke-RestMethod { [pscustomobject]@{ value = @() } }
        Mock Get-ScheduledTask { $null }
        Mock Test-OutlookAvailability { $true }
    }

    It "Fails with a pointer to the installer when there is no preferences file" {
        $run = Invoke-HealthCheck -Offline
        $run.ExitCode | Should -Be 1
        $failure = $run.Results | Where-Object Level -eq 'Fail'
        $failure.Check | Should -Be 'Preferences file'
        $failure.Message | Should -Match 'Install-Yardstick\.ps1'
    }

    It "Passes a minimal setup and creates nothing but its scratch files" {
        Set-TestPreferences
        New-Item -ItemType Directory -Path (Join-Path $script:root 'Icons') -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:root 'Icons\demo.png') -Value 'png'

        $run = Invoke-HealthCheck
        ($run.Results | Where-Object Level -eq 'Fail').Message | Should -BeNullOrEmpty
        $run.ExitCode | Should -Be 0
        ($run.Results | Where-Object Check -eq 'Missing icons').Level | Should -Be 'OK'
        ($run.Results | Where-Object Check -eq 'Intune').Level | Should -Be 'OK'
        Get-ChildItem -LiteralPath (Join-Path $script:root 'Icons') -Filter '*.tmp' | Should -BeNullOrEmpty
        Test-Path (Join-Path $script:root 'Artifacts\Temp') | Should -BeFalse
    }

    It "Names a required permission that has no admin consent" {
        Set-TestPreferences
        Mock Get-YardstickAppTokenRoles { @('DeviceManagementApps.ReadWrite.All', 'DeviceManagementConfiguration.Read.All') }
        $run = Invoke-HealthCheck
        $run.ExitCode | Should -Be 1
        ($run.Results | Where-Object Level -eq 'Fail').Check | Should -Be 'DeviceManagementRBAC.Read.All'
        ($run.Results | Where-Object Check -eq 'Group.Read.All').Level | Should -Be 'Info'
    }

    It "Fails when the credential cannot sign in" {
        Set-TestPreferences
        Mock Get-YardstickAppTokenRoles { throw 'AADSTS7000215: Invalid client secret provided.' }
        $run = Invoke-HealthCheck
        $run.ExitCode | Should -Be 1
        ($run.Results | Where-Object Check -eq 'Graph sign-in').Message | Should -Match 'AADSTS7000215'
    }

    It "Does not sign in with -Offline" {
        Set-TestPreferences
        $run = Invoke-HealthCheck -Offline
        Should -Invoke Get-YardstickAppTokenRoles -Times 0 -Exactly
        ($run.Results | Where-Object Check -eq 'Graph').Level | Should -Be 'Skip'
    }

    It "Fails when no credential is stored for this user" {
        Set-TestPreferences
        Mock Get-YardstickIntuneCredential { $null }
        $run = Invoke-HealthCheck
        $run.ExitCode | Should -Be 1
        ($run.Results | Where-Object Check -eq 'Credential').Message | Should -Match 'AppRegistration'
    }

    It "Fails an expired secret and warns about one close to expiry" {
        Set-TestPreferences
        $Global:YardstickTestSecretExpiresOn = (Get-Date).AddDays(-1)
        ((Invoke-HealthCheck -Offline).Results | Where-Object Check -eq 'Secret expiry').Level | Should -Be 'Fail'
        $Global:YardstickTestSecretExpiresOn = (Get-Date).AddDays(5)
        ((Invoke-HealthCheck -Offline).Results | Where-Object Check -eq 'Secret expiry').Level | Should -Be 'Warn'
    }

    It "Reports invalid preferences as failures" {
        Set-TestPreferences "defaultOwner: 'IT'`ndefaultDeploymentGroups: ['<REPLACE_WITH_GROUP_ID>']"
        $run = Invoke-HealthCheck -Offline
        $run.ExitCode | Should -Be 1
        (($run.Results | Where-Object Level -eq 'Fail').Message -join "`n") | Should -Match 'defaultDeploymentGroups'
    }

    It "Fails Outlook email when Outlook is missing" {
        Set-TestPreferences "emailNotificationEnabled: true`nemailDeliveryMethod: outlook`nemailRecipient: 'a@contoso.com'"
        Mock Test-OutlookAvailability { $false }
        ((Invoke-HealthCheck -Offline).Results | Where-Object Check -eq 'Outlook').Level | Should -Be 'Fail'
    }

    It "Warns when the scheduled task runs as another account or another folder" {
        Set-TestPreferences
        Mock Get-ScheduledTask {
            [pscustomobject]@{
                State     = 'Ready'
                Principal = [pscustomobject]@{ UserId = 'CONTOSO\svc-yardstick' }
                Actions   = @([pscustomobject]@{ Arguments = '-File "D:\Other\Yardstick.ps1" -All -NoInteractive' })
            }
        }
        Mock Get-ScheduledTaskInfo { [pscustomobject]@{ LastTaskResult = 1; LastRunTime = (Get-Date).AddHours(-3); NextRunTime = (Get-Date).AddHours(21) } }

        $results = (Invoke-HealthCheck -Offline).Results
        ($results | Where-Object Check -eq 'Task account').Message | Should -Match 'svc-yardstick'
        ($results | Where-Object Check -eq 'Task target').Level | Should -Be 'Warn'
        ($results | Where-Object Check -eq 'Last run').Message | Should -Match '0x1\b'
    }
}
