BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $script:repoRoot = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
    Import-Module powershell-yaml -ErrorAction Stop
    Import-Module "$script:repoRoot\Modules\YardstickConfig.psm1" -Force
    Import-Module "$script:repoRoot\Modules\YardstickSetup.psm1" -Force

    # A throwaway Yardstick root with the real layout manifest and preferences template.
    function New-TestProjectRoot {
        $root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path (Join-Path $root 'Config') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $root 'Recipes') | Out-Null
        Copy-Item "$script:repoRoot\Yardstick.Project.psd1" $root
        Copy-Item "$script:repoRoot\Config\Preferences.example.yaml" (Join-Path $root 'Config')
        $root
    }

    function New-TestJwt {
        param([hashtable]$Claims)
        $encode = {
            param($Object)
            [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($Object | ConvertTo-Json -Compress))).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        }
        "$(& $encode @{ alg = 'none' }).$(& $encode $Claims).signature"
    }

    # An error shaped like the one Invoke-RestMethod raises for an HTTP failure.
    function New-TestHttpError {
        param([int]$Status, [string]$Body)
        $record = [System.Management.Automation.ErrorRecord]::new(
            [Exception]::new("Response status code does not indicate success: $Status (Error)."), 'Http', 'NotSpecified', $null)
        if ($Body) { $record.ErrorDetails = [System.Management.Automation.ErrorDetails]::new($Body) }
        $record
    }

    function Set-TestSession {
        param($Session)
        InModuleScope YardstickSetup -Parameters @{ Session = $Session } {
            param($Session)
            $Script:SetupSession = $Session
        }
    }

    $script:restSession = [pscustomobject]@{
        Mode = 'Rest'; TenantId = 'tenant-1'; Account = 'admin@contoso.com'
        AccessToken = 'delegated-token'; ExpiresOn = [datetimeoffset]::UtcNow.AddHours(1)
    }

    # Graph's service principal with the app roles Yardstick asks for.
    $script:graphSp = [pscustomobject]@{
        id       = 'graph-sp'
        appId    = '00000003-0000-0000-c000-000000000000'
        appRoles = @(
            [pscustomobject]@{ value = 'DeviceManagementApps.ReadWrite.All'; id = 'role-apps' }
            [pscustomobject]@{ value = 'DeviceManagementConfiguration.Read.All'; id = 'role-config' }
            [pscustomobject]@{ value = 'DeviceManagementRBAC.Read.All'; id = 'role-rbac' }
            [pscustomobject]@{ value = 'Group.Read.All'; id = 'role-group' }
            [pscustomobject]@{ value = 'Application.Read.All'; id = 'role-app-read' }
            [pscustomobject]@{ value = 'User.Read.All'; id = 'role-user' }
        )
    }
}

AfterAll {
    Set-YardstickSetupMode
    Set-TestSession $null
}


Describe "ConvertFrom-YardstickJwt" {
    It "Decodes the claims of a base64url token without padding" {
        $token = New-TestJwt @{ tid = 'abc'; roles = @('A', 'B'); upn = 'x+y/z@contoso.com' }
        $claims = ConvertFrom-YardstickJwt -Token $token
        $claims.tid | Should -Be 'abc'
        $claims.roles | Should -Be @('A', 'B')
        $claims.upn | Should -Be 'x+y/z@contoso.com'
    }

    It "Rejects a value that is not a JWT" {
        { ConvertFrom-YardstickJwt -Token 'not-a-token' } | Should -Throw '*not a JWT*'
    }
}


Describe "Get-YardstickRequiredGraphPermissions" {
    It "Lists the two permissions a normal run needs" {
        (Get-YardstickRequiredGraphPermissions -RequiredOnly).Name | Should -Be @(
            'DeviceManagementApps.ReadWrite.All'
            'DeviceManagementRBAC.Read.All'
        )
    }

    It "Marks Group.Read.All and Application.Read.All optional" {
        $optional = Get-YardstickRequiredGraphPermissions | Where-Object { -not $_.Required }
        $optional.Name | Should -Be @('Group.Read.All', 'Application.Read.All')
    }

    It "Gives a reason for every permission" {
        Get-YardstickRequiredGraphPermissions | ForEach-Object { $_.Reason | Should -Not -BeNullOrEmpty }
    }
}


Describe "Test-YardstickGraphRoleGranted" {
    It "Finds a role granted by name" {
        Test-YardstickGraphRoleGranted -Roles @('DeviceManagementRBAC.Read.All') -Name 'DeviceManagementRBAC.Read.All' | Should -BeTrue
    }

    It "Counts a ReadWrite.All role as covering the matching Read.All one" {
        Test-YardstickGraphRoleGranted -Roles @('DeviceManagementRBAC.ReadWrite.All') -Name 'DeviceManagementRBAC.Read.All' | Should -BeTrue
    }

    It "Does not count a Read.All role as covering ReadWrite.All" {
        Test-YardstickGraphRoleGranted -Roles @('DeviceManagementApps.Read.All') -Name 'DeviceManagementApps.ReadWrite.All' | Should -BeFalse
    }

    It "Does not count a ReadWrite.All role on a different resource" {
        Test-YardstickGraphRoleGranted -Roles @('DeviceManagementConfiguration.ReadWrite.All') -Name 'DeviceManagementRBAC.Read.All' | Should -BeFalse
    }

    It "Reports a role as missing when nothing is granted" {
        Test-YardstickGraphRoleGranted -Roles @() -Name 'DeviceManagementRBAC.Read.All' | Should -BeFalse
    }
}


Describe "Prompt helpers" {
    AfterEach { Set-YardstickSetupMode }

    Context "Non-interactive" {
        BeforeEach { Set-YardstickSetupMode -NonInteractive }

        It "Returns defaults without prompting" {
            Mock Read-Host { throw 'prompted' } -ModuleName YardstickSetup
            Read-YardstickSetupValue -Prompt 'Owner' -Default 'IT' | Should -Be 'IT'
            Read-YardstickSetupYesNo -Prompt 'Go?' -Default $false | Should -BeFalse
            Read-YardstickSetupChoice -Prompt 'Pick' -Choices @('a', 'b') -Default @('b') | Should -Be 'b'
        }

        It "Fails a required value that has no default" {
            { Read-YardstickSetupValue -Prompt 'Owner' -Required } | Should -Throw '*required*'
        }
    }

    Context "Interactive" {
        It "Accepts the default on Enter" {
            Mock Read-Host { '' } -ModuleName YardstickSetup
            Read-YardstickSetupValue -Prompt 'Owner' -Default 'IT' | Should -Be 'IT'
            Read-YardstickSetupYesNo -Prompt 'Go?' -Default $true | Should -BeTrue
        }

        It "Picks numbered choices" {
            Mock Read-Host { '2' } -ModuleName YardstickSetup
            Read-YardstickSetupChoice -Prompt 'Pick' -Choices @('a', 'b', 'c') | Should -Be 'b'
        }

        It "Picks several choices when allowed" {
            Mock Read-Host { '3, 1' } -ModuleName YardstickSetup
            Read-YardstickSetupChoice -Prompt 'Pick' -Choices @('a', 'b', 'c') -Multiple | Should -Be @('c', 'a')
        }

        It "Asks again after an out-of-range answer" {
            $script:answers = [System.Collections.Generic.Queue[string]]::new([string[]]@('9', '1'))
            Mock Read-Host { $script:answers.Dequeue() } -ModuleName YardstickSetup
            Read-YardstickSetupChoice -Prompt 'Pick' -Choices @('a', 'b') | Should -Be 'a'
            Should -Invoke Read-Host -ModuleName YardstickSetup -Times 2 -Exactly
        }
    }
}


Describe "Install-YardstickModule" {
    BeforeAll {
        # Load these before Get-Module is mocked, which would stop them autoloading.
        Import-Module PowerShellGet, Microsoft.PowerShell.PSResourceGet -ErrorAction SilentlyContinue
    }
    BeforeEach {
        # Get-Module reports the module as installed once an installer has "run".
        $Global:YardstickTestModuleInstalled = $false
        Mock Get-Module { if ($Global:YardstickTestModuleInstalled) { [pscustomobject]@{ Name = 'demo' } } } -ModuleName YardstickSetup -ParameterFilter { $Name -eq 'demo' }
        Mock Write-YardstickSetupStatus { } -ModuleName YardstickSetup
    }
    AfterAll { Remove-Variable -Name YardstickTestModuleInstalled -Scope Global -ErrorAction SilentlyContinue }

    It "Does nothing when the module is already installed" {
        $Global:YardstickTestModuleInstalled = $true
        Mock Install-PSResource { throw 'not expected' } -ModuleName YardstickSetup
        Mock Install-Module { throw 'not expected' } -ModuleName YardstickSetup
        Install-YardstickModule -Name 'demo' | Should -BeTrue
        Should -Invoke Install-PSResource -ModuleName YardstickSetup -Times 0
        Should -Invoke Install-Module -ModuleName YardstickSetup -Times 0
    }

    It "Installs with PSResourceGet when it is available" {
        Mock Install-PSResource { $Global:YardstickTestModuleInstalled = $true } -ModuleName YardstickSetup
        Mock Install-Module { throw 'not expected' } -ModuleName YardstickSetup
        Install-YardstickModule -Name 'demo' | Should -BeTrue
        Should -Invoke Install-PSResource -ModuleName YardstickSetup -Times 1 -ParameterFilter { $Scope -eq 'CurrentUser' -and $TrustRepository }
        Should -Invoke Install-Module -ModuleName YardstickSetup -Times 0
    }

    It "Falls back to Install-Module when PSResourceGet fails" {
        Mock Install-PSResource { throw 'repository not found' } -ModuleName YardstickSetup
        Mock Install-Module { $Global:YardstickTestModuleInstalled = $true } -ModuleName YardstickSetup
        Install-YardstickModule -Name 'demo' | Should -BeTrue
        Should -Invoke Install-Module -ModuleName YardstickSetup -Times 1 -ParameterFilter { $Scope -eq 'CurrentUser' }
    }

    It "Reports both errors and returns false when neither works" {
        Mock Install-PSResource { throw 'psresource failed' } -ModuleName YardstickSetup
        Mock Install-Module { throw 'no repositories' } -ModuleName YardstickSetup
        Install-YardstickModule -Name 'demo' | Should -BeFalse
        Should -Invoke Write-YardstickSetupStatus -ModuleName YardstickSetup -Times 1 -ParameterFilter { $Level -eq 'Fail' -and $Message -match 'psresource failed' -and $Message -match 'no repositories' }
    }

    It "Installs over an older version with -Update, allowing prereleases with -Prerelease" {
        $Global:YardstickTestModuleInstalled = $true
        Mock Install-PSResource { } -ModuleName YardstickSetup
        Mock Install-Module { throw 'not expected' } -ModuleName YardstickSetup
        Install-YardstickModule -Name 'demo' -Prerelease -Update | Should -BeTrue
        Should -Invoke Install-PSResource -ModuleName YardstickSetup -Times 1 -ParameterFilter { $Prerelease }
    }

    It "Passes -AllowPrerelease to the Install-Module fallback" {
        Mock Install-PSResource { throw 'repository not found' } -ModuleName YardstickSetup
        Mock Install-Module { $Global:YardstickTestModuleInstalled = $true } -ModuleName YardstickSetup
        Install-YardstickModule -Name 'demo' -Prerelease | Should -BeTrue
        Should -Invoke Install-Module -ModuleName YardstickSetup -Times 1 -ParameterFilter { $AllowPrerelease }
    }

    It "Passes -SkipPublisherCheck to the Install-Module fallback only when asked" {
        Mock Install-PSResource { throw 'repository not found' } -ModuleName YardstickSetup
        Mock Install-Module { $Global:YardstickTestModuleInstalled = $true } -ModuleName YardstickSetup
        Install-YardstickModule -Name 'demo' | Should -BeTrue
        Should -Invoke Install-Module -ModuleName YardstickSetup -Times 1 -Exactly -ParameterFilter { -not $SkipPublisherCheck }

        $Global:YardstickTestModuleInstalled = $false
        Install-YardstickModule -Name 'demo' -SkipPublisherCheck | Should -BeTrue
        Should -Invoke Install-Module -ModuleName YardstickSetup -Times 1 -Exactly -ParameterFilter { $SkipPublisherCheck }
    }
}


Describe "ConvertTo-YardstickYamlValue" {
    It "Formats <Label>" -TestCases @(
        @{ Label = 'true'; Value = $true; Expected = 'true' }
        @{ Label = 'false'; Value = $false; Expected = 'false' }
        @{ Label = 'an integer'; Value = 587; Expected = '587' }
        @{ Label = 'a string'; Value = 'IT Services'; Expected = "'IT Services'" }
        @{ Label = 'an apostrophe'; Value = "O'Brien"; Expected = "'O''Brien'" }
        @{ Label = 'a list'; Value = @('a', 'b'); Expected = "['a', 'b']" }
        @{ Label = 'an empty string'; Value = ''; Expected = '' }
    ) {
        ConvertTo-YardstickYamlValue $Value | Should -Be $Expected
    }

    It "Formats null as blank and an empty list as []" {
        ConvertTo-YardstickYamlValue $null | Should -Be ''
        ConvertTo-YardstickYamlValue @() | Should -Be '[]'
    }

    It "Round-trips a UNC path through ConvertFrom-Yaml" {
        $path = '\\server\share\Yardstick Backup'
        ("key: " + (ConvertTo-YardstickYamlValue $path) | ConvertFrom-Yaml).key | Should -Be $path
    }
}


Describe "Set-YardstickPreferenceValue" {
    BeforeEach {
        $script:file = Join-Path $TestDrive "$([guid]::NewGuid().ToString('N')).yaml"
    }

    It "Replaces the key's line and leaves comments alone" {
        Set-Content -LiteralPath $script:file -Value "# Owner comment`ndefaultOwner:`n# after`nother: 1"
        Set-YardstickPreferenceValue -Path $script:file -Key 'defaultOwner' -Value 'IT'
        Get-Content -LiteralPath $script:file | Should -Be @('# Owner comment', "defaultOwner: 'IT'", '# after', 'other: 1')
    }

    It "Uncomments a commented-out key" {
        Set-Content -LiteralPath $script:file -Value "# Backup: Artifacts\Backup`nother: 1"
        Set-YardstickPreferenceValue -Path $script:file -Key 'Backup' -Value 'D:\Backup'
        (Get-Content -LiteralPath $script:file)[0] | Should -Be "Backup: 'D:\Backup'"
    }

    It "Prefers the active line over a commented one" {
        Set-Content -LiteralPath $script:file -Value "# smtpPort: 587`nsmtpPort: 25"
        Set-YardstickPreferenceValue -Path $script:file -Key 'smtpPort' -Value 2525
        Get-Content -LiteralPath $script:file | Should -Be @('# smtpPort: 587', 'smtpPort: 2525')
    }

    It "Does not match a key that merely ends with the same name" {
        Set-Content -LiteralPath $script:file -Value "# adminEmailSubject: x`nemailSubject: old"
        Set-YardstickPreferenceValue -Path $script:file -Key 'emailSubject' -Value 'new'
        Get-Content -LiteralPath $script:file | Should -Be @('# adminEmailSubject: x', "emailSubject: 'new'")
    }

    It "Replaces a block list under the key" {
        Set-Content -LiteralPath $script:file -Value "defaultScopeTags:`n  - One`n  - Two`n# next`nnext: 1"
        Set-YardstickPreferenceValue -Path $script:file -Key 'defaultScopeTags' -Value @('Three')
        Get-Content -LiteralPath $script:file | Should -Be @("defaultScopeTags: ['Three']", '# next', 'next: 1')
    }

    It "Appends a key the file does not have" {
        Set-Content -LiteralPath $script:file -Value "a: 1"
        Set-YardstickPreferenceValue -Path $script:file -Key 'b' -Value $true
        Get-Content -LiteralPath $script:file | Should -Be @('a: 1', 'b: true')
    }

    It "Writes a blank value for null" {
        Set-Content -LiteralPath $script:file -Value "adminEmailRecipient: 'a@b.c'"
        Set-YardstickPreferenceValue -Path $script:file -Key 'adminEmailRecipient' -Value $null
        Get-Content -LiteralPath $script:file | Should -Be @('adminEmailRecipient:')
    }

    It "Keeps CRLF line endings" {
        [IO.File]::WriteAllText($script:file, "a: 1`r`nb: 2`r`n")
        Set-YardstickPreferenceValue -Path $script:file -Key 'a' -Value 3
        [IO.File]::ReadAllText($script:file) | Should -Be "a: 3`r`nb: 2`r`n"
    }

    It "Keeps LF line endings" {
        [IO.File]::WriteAllText($script:file, "a: 1`nb: 2`n")
        Set-YardstickPreferenceValue -Path $script:file -Key 'b' -Value 'x'
        [IO.File]::ReadAllText($script:file) | Should -Be "a: 1`nb: 'x'`n"
    }

    It "Leaves the shipped template valid after every value the installer sets" {
        $root = New-TestProjectRoot
        $path = (New-YardstickPreferencesFile -ProjectRoot $root).Path
        $values = [ordered]@{
            defaultOwner            = "IT Services (O'Neil)"
            defaultMinOSVersion     = 'W11_22H2'
            Backup                  = '\\server\share\Backup'
            SoftwareDropbox         = 'D:\Dropbox'
            SoftwareArchive         = 'D:\Archive'
            defaultScopeTags        = @('Default', 'Workstations')
            defaultDeploymentGroups = @('11111111-2222-3333-4444-555555555555')
            emailRecipient          = @('a@contoso.com', 'b@contoso.com')
            adminEmailRecipient     = 'admin@contoso.com'
            emailSendFromAddress    = 'yardstick@contoso.com'
            smtpServer              = 'smtp.contoso.com'
            smtpPort                = 587
            smtpUseSsl              = $true
            emailDeliveryMethod     = 'smtp'
            emailNotificationEnabled = $true
        }
        foreach ($key in $values.Keys) {
            Set-YardstickPreferenceValue -Path $path -Key $key -Value $values[$key]
        }

        $prefs = Import-YardstickPreferences -ProjectRoot $root
        foreach ($key in $values.Keys) {
            $prefs[$key] | Should -Be $values[$key] -Because $key
        }
        $check = Test-YardstickPreferences -Preferences $prefs
        $check.Errors | Should -BeNullOrEmpty
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'FOLDER LOCATIONS'
    }
}


Describe "New-YardstickPreferencesFile" {
    BeforeEach { $script:root = New-TestProjectRoot }

    It "Copies the template to Local\preferences.yaml" {
        $result = New-YardstickPreferencesFile -ProjectRoot $script:root
        $result.Created | Should -BeTrue
        $result.Path | Should -Be (Join-Path $script:root 'Local\preferences.yaml')
        Get-Content -LiteralPath $result.Path -Raw | Should -Be (Get-Content -LiteralPath (Join-Path $script:root 'Config\Preferences.example.yaml') -Raw)
    }

    It "Leaves an existing file alone" {
        New-YardstickPreferencesFile -ProjectRoot $script:root | Out-Null
        $path = Join-Path $script:root 'Local\preferences.yaml'
        Set-Content -LiteralPath $path -Value 'defaultOwner: Mine'
        $result = New-YardstickPreferencesFile -ProjectRoot $script:root
        $result.Created | Should -BeFalse
        Get-Content -LiteralPath $path | Should -Be 'defaultOwner: Mine'
    }

    It "Uses a legacy root preferences.yaml instead of creating a new file" {
        $legacy = Join-Path $script:root 'preferences.yaml'
        Set-Content -LiteralPath $legacy -Value 'defaultOwner: Legacy'
        $result = New-YardstickPreferencesFile -ProjectRoot $script:root
        $result.Created | Should -BeFalse
        $result.Path | Should -Be $legacy
        Test-Path (Join-Path $script:root 'Local\preferences.yaml') | Should -BeFalse
    }

    It "Explains a missing template" {
        Remove-Item (Join-Path $script:root 'Config\Preferences.example.yaml')
        { New-YardstickPreferencesFile -ProjectRoot $script:root } | Should -Throw '*template*'
    }
}


Describe "Get-YardstickSetupGraphError" {
    It "Reads a Graph error body" {
        $err = Get-YardstickSetupGraphError (New-TestHttpError 403 '{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges."}}')
        $err.StatusCode | Should -Be 403
        $err.Code | Should -Be 'Authorization_RequestDenied'
        $err.Message | Should -Be 'Insufficient privileges.'
    }

    It "Reads an OAuth error body" {
        $err = Get-YardstickSetupGraphError (New-TestHttpError 400 '{"error":"authorization_pending","error_description":"Waiting"}')
        $err.Code | Should -Be 'authorization_pending'
        $err.Message | Should -Be 'Waiting'
    }
}


Describe "Graph requests with the setup session" {
    AfterEach { Set-TestSession $null }

    It "Refuses to run before sign-in" {
        Set-TestSession $null
        { Invoke-YardstickSetupGraphRequest -Uri 'v1.0/me' } | Should -Throw '*Not signed in*'
    }

    It "Sends the delegated token to a Graph-relative URI" {
        Set-TestSession $script:restSession
        Mock Invoke-RestMethod { [pscustomobject]@{ ok = $true } } -ModuleName YardstickSetup
        Invoke-YardstickSetupGraphRequest -Method POST -Uri 'v1.0/applications' -Body @{ displayName = 'Y' } | Out-Null
        Should -Invoke Invoke-RestMethod -ModuleName YardstickSetup -Times 1 -Exactly -ParameterFilter {
            $Uri -eq 'https://graph.microsoft.com/v1.0/applications' -and $Method -eq 'POST' -and
            $Headers.Authorization -eq 'Bearer delegated-token' -and ($Body | ConvertFrom-Json).displayName -eq 'Y'
        }
    }

    It "Uses the app's own token in App mode" {
        Use-YardstickSetupAppCredential -Credential ([pscustomobject]@{ TenantID = 't'; ClientID = 'c'; ClientSecret = 's' }) | Out-Null
        Mock Connect-YardstickGraph { [pscustomobject]@{ AccessToken = 'app-token' } } -ModuleName YardstickSetup
        Mock Invoke-RestMethod { [pscustomobject]@{ value = @() } } -ModuleName YardstickSetup
        Invoke-YardstickSetupGraphRequest -Uri 'beta/deviceManagement/roleScopeTags' | Out-Null
        Should -Invoke Connect-YardstickGraph -ModuleName YardstickSetup -ParameterFilter { $TenantID -eq 't' -and $ClientID -eq 'c' -and $ClientSecret -eq 's' }
        Should -Invoke Invoke-RestMethod -ModuleName YardstickSetup -ParameterFilter { $Headers.Authorization -eq 'Bearer app-token' }
    }

    It "Follows @odata.nextLink" {
        Set-TestSession $script:restSession
        Mock Invoke-RestMethod {
            if ($Uri -like '*page2*') { [pscustomobject]@{ value = @([pscustomobject]@{ displayName = 'Two' }) } }
            else { [pscustomobject]@{ value = @([pscustomobject]@{ displayName = 'One' }); '@odata.nextLink' = 'https://graph.microsoft.com/beta/page2' } }
        } -ModuleName YardstickSetup
        Get-YardstickScopeTagNames | Should -Be @('One', 'Two')
    }

    It "Finds a group by name with an escaped filter" {
        Set-TestSession $script:restSession
        Mock Invoke-RestMethod { [pscustomobject]@{ value = @([pscustomobject]@{ id = 'g1'; displayName = "IT's Test" }) } } -ModuleName YardstickSetup
        (Find-YardstickGroup -Name "IT's Test").id | Should -Be 'g1'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickSetup -ParameterFilter {
            [uri]::UnescapeDataString($Uri) -like "*groups?`$filter=displayName eq 'IT''s Test'*"
        }
    }

    It "Reads a group directly by object ID" {
        Set-TestSession $script:restSession
        Mock Invoke-RestMethod { [pscustomobject]@{ id = '11111111-2222-3333-4444-555555555555'; displayName = 'Test' } } -ModuleName YardstickSetup
        $group = @(Find-YardstickGroup -Name '11111111-2222-3333-4444-555555555555')
        $group.Count | Should -Be 1
        Should -Invoke Invoke-RestMethod -ModuleName YardstickSetup -ParameterFilter { $Uri -like '*/v1.0/groups/11111111-2222-3333-4444-555555555555?*' }
    }
}


Describe "Connect-YardstickSetupGraph (device code)" {
    BeforeEach {
        Set-TestSession $null
        Mock Get-Module { $null } -ModuleName YardstickSetup -ParameterFilter { $Name -eq 'Microsoft.Graph.Authentication' }
        Mock Start-Sleep { } -ModuleName YardstickSetup
        Mock Write-Host { } -ModuleName YardstickSetup
        $script:tokenCalls = 0
        $script:jwt = New-TestJwt @{ tid = 'tenant-from-token'; upn = 'admin@contoso.com' }
        Mock Invoke-RestMethod {
            if ($Uri -like '*/devicecode') {
                return [pscustomobject]@{ device_code = 'dc'; user_code = 'ABCD'; verification_uri = 'https://microsoft.com/devicelogin'; interval = 1; expires_in = 900 }
            }
            $script:tokenCalls++
            if ($script:tokenCalls -lt 3) { throw (New-TestHttpError 400 '{"error":"authorization_pending"}') }
            [pscustomobject]@{ access_token = $script:jwt; expires_in = 3600 }
        } -ModuleName YardstickSetup
    }
    AfterEach { Set-TestSession $null }

    It "Polls until sign-in completes and takes the tenant from the token" {
        $session = Connect-YardstickSetupGraph
        $session.Mode | Should -Be 'Rest'
        $session.TenantId | Should -Be 'tenant-from-token'
        $session.Account | Should -Be 'admin@contoso.com'
        $script:tokenCalls | Should -Be 3
        Should -Invoke Invoke-RestMethod -ModuleName YardstickSetup -ParameterFilter {
            $Uri -like '*/organizations/oauth2/v2.0/devicecode' -and $Body.client_id -eq '14d82eec-204b-4c2f-b7e8-296a70dab67e' -and
            $Body.scope -match 'Application\.ReadWrite\.All' -and $Body.scope -match 'AppRoleAssignment\.ReadWrite\.All'
        }
    }

    It "Reuses the session on a second call" {
        Connect-YardstickSetupGraph | Out-Null
        Connect-YardstickSetupGraph | Out-Null
        Should -Invoke Invoke-RestMethod -ModuleName YardstickSetup -Times 1 -Exactly -ParameterFilter { $Uri -like '*/devicecode' }
    }

    It "Does not treat an app session as a sign-in" {
        Use-YardstickSetupAppCredential -Credential ([pscustomobject]@{ TenantID = 't'; ClientID = 'c'; ClientSecret = 's' }) | Out-Null
        (Connect-YardstickSetupGraph).Mode | Should -Be 'Rest'
    }

    It "Stops when sign-in is declined" {
        Mock Invoke-RestMethod {
            if ($Uri -like '*/devicecode') {
                return [pscustomobject]@{ device_code = 'dc'; user_code = 'ABCD'; verification_uri = 'u'; interval = 1; expires_in = 900 }
            }
            throw (New-TestHttpError 400 '{"error":"authorization_declined"}')
        } -ModuleName YardstickSetup
        { Connect-YardstickSetupGraph } | Should -Throw '*declined*'
    }
}


Describe "App registration building blocks" {
    It "Maps permission names to the tenant's app role IDs" {
        $map = Resolve-YardstickGraphAppRoleIds -ServicePrincipal $script:graphSp -PermissionName @('DeviceManagementRBAC.Read.All', 'DeviceManagementApps.ReadWrite.All')
        @($map.Keys) | Should -Be @('DeviceManagementRBAC.Read.All', 'DeviceManagementApps.ReadWrite.All')
        $map['DeviceManagementApps.ReadWrite.All'] | Should -Be 'role-apps'
    }

    It "Rejects an unknown permission name" {
        { Resolve-YardstickGraphAppRoleIds -ServicePrincipal $script:graphSp -PermissionName @('Nope.All') } | Should -Throw "*'Nope.All'*"
    }

    It "Builds requiredResourceAccess from scratch" {
        $rra = New-YardstickRequiredResourceAccess -Existing $null -RoleIds ([ordered]@{ A = 'id-a'; B = 'id-b' })
        @($rra).Count | Should -Be 1
        $rra[0].resourceAppId | Should -Be '00000003-0000-0000-c000-000000000000'
        $rra[0].resourceAccess.id | Should -Be @('id-a', 'id-b')
        $rra[0].resourceAccess.type | Should -Be @('Role', 'Role')
    }

    It "Merges into what the app already requests without duplicates" {
        $existing = @(
            [pscustomobject]@{ resourceAppId = 'other-api'; resourceAccess = @([pscustomobject]@{ id = 'x'; type = 'Scope' }) }
            [pscustomobject]@{ resourceAppId = '00000003-0000-0000-c000-000000000000'; resourceAccess = @([pscustomobject]@{ id = 'id-a'; type = 'Role' }) }
        )
        $rra = New-YardstickRequiredResourceAccess -Existing $existing -RoleIds ([ordered]@{ A = 'id-a'; B = 'id-b' })
        @($rra).Count | Should -Be 2
        ($rra | Where-Object resourceAppId -eq 'other-api').resourceAccess.id | Should -Be 'x'
        ($rra | Where-Object resourceAppId -eq '00000003-0000-0000-c000-000000000000').resourceAccess.id | Should -Be @('id-a', 'id-b')
    }

    It "Builds the admin consent URL" {
        Get-YardstickAdminConsentUrl -TenantId 't1' -ClientId 'c1' | Should -Be 'https://login.microsoftonline.com/t1/adminconsent?client_id=c1'
    }

    Context "Grant-YardstickAppRoles" {
        BeforeEach {
            Set-TestSession $script:restSession
            $script:posted = [System.Collections.Generic.List[object]]::new()
            Mock Start-Sleep { } -ModuleName YardstickSetup
        }
        AfterEach { Set-TestSession $null }

        It "Assigns only the roles not already granted" {
            Mock Invoke-RestMethod {
                if ($Method -eq 'GET') { return [pscustomobject]@{ value = @([pscustomobject]@{ resourceId = 'graph-sp'; appRoleId = 'id-a' }) } }
                $script:posted.Add(($Body | ConvertFrom-Json)); [pscustomobject]@{ id = 'assignment' }
            } -ModuleName YardstickSetup

            $result = Grant-YardstickAppRoles -ServicePrincipalId 'app-sp' -ResourceServicePrincipalId 'graph-sp' -RoleIds ([ordered]@{ A = 'id-a'; B = 'id-b' })
            $result.AlreadyGranted | Should -Be @('A')
            $result.Granted | Should -Be @('B')
            $result.Failed | Should -BeNullOrEmpty
            $script:posted.Count | Should -Be 1
            $script:posted[0].principalId | Should -Be 'app-sp'
            $script:posted[0].resourceId | Should -Be 'graph-sp'
            $script:posted[0].appRoleId | Should -Be 'id-b'
            Should -Invoke Invoke-RestMethod -ModuleName YardstickSetup -ParameterFilter { $Method -eq 'POST' -and $Uri -like '*/servicePrincipals/graph-sp/appRoleAssignedTo' }
        }

        It "Reports roles it is not allowed to grant without retrying" {
            Mock Invoke-RestMethod {
                if ($Method -eq 'GET') { return [pscustomobject]@{ value = @() } }
                throw (New-TestHttpError 403 '{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges."}}')
            } -ModuleName YardstickSetup

            $result = Grant-YardstickAppRoles -ServicePrincipalId 'app-sp' -ResourceServicePrincipalId 'graph-sp' -RoleIds ([ordered]@{ A = 'id-a' })
            $result.Failed | Should -Be @('A')
            Should -Invoke Start-Sleep -ModuleName YardstickSetup -Times 0 -Exactly
        }

        It "Retries while the new service principal replicates" {
            $script:attempts = 0
            Mock Invoke-RestMethod {
                if ($Method -eq 'GET') { return [pscustomobject]@{ value = @() } }
                $script:attempts++
                if ($script:attempts -lt 3) { throw (New-TestHttpError 400 '{"error":{"code":"Request_BadRequest","message":"Principal not found."}}') }
                [pscustomobject]@{ id = 'assignment' }
            } -ModuleName YardstickSetup

            $result = Grant-YardstickAppRoles -ServicePrincipalId 'app-sp' -ResourceServicePrincipalId 'graph-sp' -RoleIds ([ordered]@{ A = 'id-a' })
            $result.Granted | Should -Be @('A')
            Should -Invoke Start-Sleep -ModuleName YardstickSetup -Times 2 -Exactly
        }
    }

    Context "New-YardstickClientSecret" {
        BeforeEach { Set-TestSession $script:restSession }
        AfterEach { Set-TestSession $null }

        It "Adds a password with the requested lifetime" {
            Mock Invoke-RestMethod { [pscustomobject]@{ secretText = 's3cret'; keyId = 'k1'; endDateTime = '2027-10-12T00:00:00Z' } } -ModuleName YardstickSetup
            $secret = New-YardstickClientSecret -ApplicationObjectId 'app-obj' -Months 6
            $secret.SecretText | Should -Be 's3cret'
            $secret.EndDateTime | Should -Be ([datetime]'2027-10-12T00:00:00Z').ToUniversalTime()
            Should -Invoke Invoke-RestMethod -ModuleName YardstickSetup -ParameterFilter {
                $requested = ([datetime](($Body | ConvertFrom-Json).passwordCredential.endDateTime)).ToUniversalTime()
                $Uri -like '*/v1.0/applications/app-obj/addPassword' -and
                [math]::Abs(($requested - [datetime]::UtcNow.AddMonths(6)).TotalHours) -lt 1
            }
        }
    }
}


Describe "Wait-YardstickAppCredential" {
    BeforeEach { Mock Start-Sleep { } -ModuleName YardstickSetup }

    It "Retries until the token carries every required role" {
        $script:roleCalls = 0
        Mock Get-YardstickAppTokenRoles {
            $script:roleCalls++
            switch ($script:roleCalls) {
                1 { throw 'AADSTS7000215: Invalid client secret provided.' }
                2 { @('DeviceManagementApps.ReadWrite.All') }
                default { @('DeviceManagementApps.ReadWrite.All', 'DeviceManagementRBAC.Read.All') }
            }
        } -ModuleName YardstickSetup

        $result = Wait-YardstickAppCredential -TenantId 't' -ClientId 'c' -ClientSecret 's' -TimeoutSeconds 600 -PollSeconds 1
        $result.Success | Should -BeTrue
        $script:roleCalls | Should -Be 3
    }

    It "Accepts a ReadWrite.All role in place of a required Read.All one" {
        Mock Get-YardstickAppTokenRoles { @('DeviceManagementApps.ReadWrite.All', 'DeviceManagementRBAC.ReadWrite.All') } -ModuleName YardstickSetup
        $result = Wait-YardstickAppCredential -TenantId 't' -ClientId 'c' -ClientSecret 's' -TimeoutSeconds 0
        $result.Success | Should -BeTrue
    }

    It "Gives up after the timeout and names what is missing" {
        Mock Get-YardstickAppTokenRoles { @('DeviceManagementApps.ReadWrite.All') } -ModuleName YardstickSetup
        $result = Wait-YardstickAppCredential -TenantId 't' -ClientId 'c' -ClientSecret 's' -TimeoutSeconds 0
        $result.Success | Should -BeFalse
        $result.Missing | Should -Be @('DeviceManagementRBAC.Read.All')
    }
}


Describe "New-YardstickAppRegistration" {
    BeforeEach {
        Set-TestSession $script:restSession
        Mock Start-Sleep { } -ModuleName YardstickSetup
        Mock Write-Host { } -ModuleName YardstickSetup
        Mock Set-YardstickIntuneCredential { } -ModuleName YardstickSetup
        Mock Wait-YardstickAppCredential { [pscustomobject]@{ Success = $true; Roles = @(); Missing = @(); Error = $null } } -ModuleName YardstickSetup

        $script:requests = [System.Collections.Generic.List[object]]::new()
        $script:existingApps = @()
        $script:denyConsent = $false
        Mock Invoke-RestMethod {
            $parsed = if ($Body) { $Body | ConvertFrom-Json } else { $null }
            $script:requests.Add([pscustomobject]@{ Method = $Method; Uri = [uri]::UnescapeDataString($Uri); Body = $parsed })
            switch -Regex ("$Method $Uri") {
                "^GET .*servicePrincipals\(appId='00000003" { return $script:graphSp }
                '^GET .*/applications\?' { return [pscustomobject]@{ value = $script:existingApps } }
                '^POST .*/v1\.0/applications$' { return [pscustomobject]@{ id = 'app-obj'; appId = 'app-id' } }
                '^PATCH .*/applications/' { return $null }
                "^GET .*servicePrincipals\(appId='app-id'\)" { throw (New-TestHttpError 404 '{"error":{"code":"Request_ResourceNotFound","message":"Not found"}}') }
                '^POST .*/v1\.0/servicePrincipals$' { return [pscustomobject]@{ id = 'app-sp'; appId = 'app-id' } }
                '^GET .*/appRoleAssignments' { return [pscustomobject]@{ value = @() } }
                '^POST .*/appRoleAssignedTo' {
                    if ($script:denyConsent) { throw (New-TestHttpError 403 '{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges."}}') }
                    return [pscustomobject]@{ id = 'assignment' }
                }
                '/addPassword$' { return [pscustomobject]@{ secretText = 's3cret'; keyId = 'k1'; endDateTime = '2027-10-12T00:00:00Z' } }
                default { throw "Unexpected request: $Method $Uri" }
            }
        } -ModuleName YardstickSetup
    }
    AfterEach { Set-TestSession $null }

    It "Creates a single-tenant app requesting exactly the required Graph roles" {
        $result = New-YardstickAppRegistration -DisplayName 'Yardstick'
        $result.Created | Should -BeTrue
        $result.ClientId | Should -Be 'app-id'
        $result.TenantId | Should -Be 'tenant-1'

        $create = $script:requests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match '/v1\.0/applications$' }
        $create.Body.displayName | Should -Be 'Yardstick'
        $create.Body.signInAudience | Should -Be 'AzureADMyOrg'
        $graph = @($create.Body.requiredResourceAccess)
        $graph.Count | Should -Be 1
        $graph[0].resourceAppId | Should -Be '00000003-0000-0000-c000-000000000000'
        @($graph[0].resourceAccess.id) | Should -Be @('role-apps', 'role-rbac')
        @($graph[0].resourceAccess.type | Select-Object -Unique) | Should -Be @('Role')
    }

    It "Adds the optional roles only when asked" {
        New-YardstickAppRegistration -IncludeOptionalPermissions | Out-Null
        $create = $script:requests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match '/v1\.0/applications$' }
        @($create.Body.requiredResourceAccess[0].resourceAccess.id) | Should -Be @('role-apps', 'role-rbac', 'role-group', 'role-app-read')
    }

    It "Creates the service principal, grants each role and adds a secret" {
        $result = New-YardstickAppRegistration
        ($script:requests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match '/v1\.0/servicePrincipals$' }).Body.appId | Should -Be 'app-id'
        $grants = @($script:requests | Where-Object { $_.Uri -match 'appRoleAssignedTo$' })
        $grants.Count | Should -Be 2
        $grants.Body.principalId | Select-Object -Unique | Should -Be 'app-sp'
        $grants.Body.resourceId | Select-Object -Unique | Should -Be 'graph-sp'
        $result.Consent.Granted | Should -Be @('DeviceManagementApps.ReadWrite.All', 'DeviceManagementRBAC.Read.All')
        ($script:requests | Where-Object { $_.Uri -match '/applications/app-obj/addPassword$' }) | Should -Not -BeNullOrEmpty
    }

    It "Stores the secret with its expiry under the credential target" {
        New-YardstickAppRegistration -CredentialTarget 'Yardstick:Test' | Out-Null
        Should -Invoke Set-YardstickIntuneCredential -ModuleName YardstickSetup -Times 1 -Exactly -ParameterFilter {
            $Target -eq 'Yardstick:Test' -and $TenantID -eq 'tenant-1' -and $ClientID -eq 'app-id' -and
            $ClientSecret -eq 's3cret' -and $SecretExpiresOn -eq ([datetime]'2027-10-12T00:00:00Z').ToUniversalTime()
        }
        Should -Invoke Wait-YardstickAppCredential -ModuleName YardstickSetup -Times 1 -Exactly -ParameterFilter { $ClientSecret -eq 's3cret' }
    }

    It "Reuses an existing app with the same name and updates its permissions" {
        $script:existingApps = @([pscustomobject]@{ id = 'app-obj'; appId = 'app-id'; displayName = 'Yardstick'; requiredResourceAccess = @() })
        $result = New-YardstickAppRegistration
        $result.Created | Should -BeFalse
        ($script:requests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match '/v1\.0/applications$' }) | Should -BeNullOrEmpty
        $patch = $script:requests | Where-Object { $_.Method -eq 'PATCH' }
        $patch.Uri | Should -Match '/v1\.0/applications/app-obj$'
        @($patch.Body.requiredResourceAccess[0].resourceAccess.id) | Should -Be @('role-apps', 'role-rbac')
    }

    It "Refuses to pick between apps with the same name" {
        $script:existingApps = @(
            [pscustomobject]@{ id = 'a1'; appId = 'x1'; displayName = 'Yardstick'; requiredResourceAccess = @() }
            [pscustomobject]@{ id = 'a2'; appId = 'x2'; displayName = 'Yardstick'; requiredResourceAccess = @() }
        )
        { New-YardstickAppRegistration } | Should -Throw '*More than one*'
    }

    It "Returns the consent URL and skips validation when consent cannot be granted" {
        $script:denyConsent = $true
        $result = New-YardstickAppRegistration
        $result.Consent.Failed.Count | Should -Be 2
        $result.ConsentUrl | Should -Be 'https://login.microsoftonline.com/tenant-1/adminconsent?client_id=app-id'
        $result.Validation | Should -BeNullOrEmpty
        Should -Invoke Wait-YardstickAppCredential -ModuleName YardstickSetup -Times 0 -Exactly
        Should -Invoke Set-YardstickIntuneCredential -ModuleName YardstickSetup -Times 1 -Exactly
    }

    It "Requires a signed-in (not app) session" {
        Use-YardstickSetupAppCredential -Credential ([pscustomobject]@{ TenantID = 't'; ClientID = 'c'; ClientSecret = 's' }) | Out-Null
        { New-YardstickAppRegistration } | Should -Throw '*Not signed in*'
    }
}


Describe "Get-YardstickScheduledTaskDefinition" {
    It "Runs every recipe unattended by default" {
        $definition = Get-YardstickScheduledTaskDefinition -ProjectRoot 'C:\Yardstick' -PwshPath 'C:\pwsh\pwsh.exe'
        $definition.Execute | Should -Be 'C:\pwsh\pwsh.exe'
        $definition.Argument | Should -Be '-NoProfile -ExecutionPolicy Bypass -File "C:\Yardstick\Yardstick.ps1" -All -NoInteractive'
        $definition.WorkingDirectory | Should -Be 'C:\Yardstick'
    }

    It "Runs one recipe group when given" {
        $definition = Get-YardstickScheduledTaskDefinition -ProjectRoot '\\server\share\Yardstick' -Group 'Lab Apps'
        $definition.Argument | Should -Be '-NoProfile -ExecutionPolicy Bypass -File "\\server\share\Yardstick\Yardstick.ps1" -Group "Lab Apps" -NoInteractive'
    }

    It "Uses arguments that bind to Yardstick.ps1's parameter sets" {
        $command = Get-Command "$script:repoRoot\Yardstick.ps1"
        $sets = $command.Parameters['NoInteractive'].ParameterSets.Keys
        $sets | Should -Contain 'AllApps'
        $sets | Should -Contain 'GroupApps'
    }
}


Describe "Register-YardstickScheduledTask" {
    BeforeEach {
        Mock Register-ScheduledTask { } -ModuleName YardstickSetup
    }

    It "Runs as the current user only while they are signed in, without a credential" {
        Register-YardstickScheduledTask -ProjectRoot 'C:\Yardstick' -At ([datetime]'03:00')
        Should -Invoke Register-ScheduledTask -ModuleName YardstickSetup -Times 1 -Exactly -ParameterFilter {
            $Principal.LogonType -eq 'Interactive' -and $Principal.RunLevel -eq 'Limited' -and
            $Principal.UserId -eq "$env:USERDOMAIN\$env:USERNAME" -and -not $Password
        }
    }

    It "Runs whether or not the account is signed in, with a credential" {
        $credential = [pscredential]::new('CONTOSO\runner', (ConvertTo-SecureString 'secret' -AsPlainText -Force))
        Register-YardstickScheduledTask -ProjectRoot 'C:\Yardstick' -At ([datetime]'03:00') -Credential $credential
        Should -Invoke Register-ScheduledTask -ModuleName YardstickSetup -Times 1 -Exactly -ParameterFilter {
            $User -eq 'CONTOSO\runner' -and $Password -eq 'secret' -and -not $Principal
        }
    }
}


Describe "Test-YardstickCredentialGuard" {
    It "Is <Expected> when the running security services are <Services>" -TestCases @(
        @{ Services = @(1, 2); Expected = $true }
        @{ Services = @(2); Expected = $false }
        @{ Services = @(); Expected = $false }
    ) {
        param($Services, $Expected)
        Mock Get-CimInstance { [pscustomobject]@{ SecurityServicesRunning = $Services } } -ModuleName YardstickSetup
        Test-YardstickCredentialGuard | Should -Be $Expected
    }

    It "Is false when Device Guard cannot be queried" {
        Mock Get-CimInstance { throw 'Invalid namespace' } -ModuleName YardstickSetup
        Test-YardstickCredentialGuard | Should -BeFalse
    }
}


Describe "Test-YardstickAclWritable" {
    BeforeAll {
        $script:user = 'S-1-5-21-1000-2000-3000-1001'
        $script:group = 'S-1-5-32-545'
        function New-TestRule {
            param([string]$Sid, [string]$Rights, [string]$Type = 'Allow', [string]$Inherit = 'None', [string]$Propagate = 'None')
            [Security.AccessControl.FileSystemAccessRule]::new([Security.Principal.SecurityIdentifier]::new($Sid), $Rights, $Inherit, $Propagate, $Type)
        }
        function New-TestAcl {
            param([object[]]$Rule)
            $acl = [Security.AccessControl.DirectorySecurity]::new()
            foreach ($r in $Rule) { $acl.AddAccessRule($r) }
            $acl
        }
    }

    It "Is writable when the user or one of its groups may modify it" {
        Test-YardstickAclWritable -Acl (New-TestAcl (New-TestRule $script:user Modify)) -Sid $script:user | Should -BeTrue
        Test-YardstickAclWritable -Acl (New-TestAcl (New-TestRule $script:group FullControl)) -Sid $script:user, $script:group | Should -BeTrue
    }

    It "Is not writable with read access only, or access for someone else" {
        Test-YardstickAclWritable -Acl (New-TestAcl (New-TestRule $script:user ReadAndExecute)) -Sid $script:user | Should -BeFalse
        Test-YardstickAclWritable -Acl (New-TestAcl (New-TestRule 'S-1-5-32-544' FullControl)) -Sid $script:user | Should -BeFalse
    }

    It "Lets a deny entry win over an allow entry" {
        $acl = New-TestAcl (New-TestRule $script:group FullControl), (New-TestRule $script:user WriteData -Type Deny)
        Test-YardstickAclWritable -Acl $acl -Sid $script:user, $script:group | Should -BeFalse
    }

    It "Ignores entries that only apply to new children" {
        $acl = New-TestAcl (New-TestRule $script:user FullControl -Inherit 'ContainerInherit, ObjectInherit' -Propagate InheritOnly)
        Test-YardstickAclWritable -Acl $acl -Sid $script:user | Should -BeFalse
    }
}


Describe "Get-YardstickUnwritablePath and Grant-YardstickFolderAccess" {
    BeforeEach {
        # A SID with no access to anything, standing in for the user setup runs as.
        $script:stranger = 'S-1-5-21-1000-2000-3000-4242'
        $script:tree = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path "$script:tree\Local", "$script:tree\.git\objects" -Force | Out-Null
        Set-Content -LiteralPath "$script:tree\YInstall.log" -Value 'log'
        Set-Content -LiteralPath "$script:tree\Local\preferences.yaml" -Value 'x: 1'
        Set-Content -LiteralPath "$script:tree\.git\objects\pack" -Value 'git'
    }

    It "Returns nothing when this session can change everything" {
        Get-YardstickUnwritablePath -ProjectRoot $script:tree | Should -BeNullOrEmpty
    }

    It "Lists the folder and everything in it, except .git, for a SID without access" {
        $paths = @(Get-YardstickUnwritablePath -ProjectRoot $script:tree -Sid $script:stranger)
        $paths | Should -Be @($script:tree, "$script:tree\Local", "$script:tree\YInstall.log", "$script:tree\Local\preferences.yaml")
    }

    It "Grants full control of the folder and its contents" {
        Grant-YardstickFolderAccess -Path $script:tree -Sid $script:stranger
        Get-YardstickUnwritablePath -ProjectRoot $script:tree -Sid $script:stranger | Should -BeNullOrEmpty
        $rule = (Get-Acl -LiteralPath $script:tree).Access | Where-Object { $_.IdentityReference.Value -eq $script:stranger }
        $rule.FileSystemRights | Should -Be 'FullControl'
        $rule.InheritanceFlags | Should -Be 'ContainerInherit, ObjectInherit'
    }

    It "Grants each listed item that does not inherit from the folder" {
        $file = "$script:tree\YInstall.log"
        $acl = Get-Acl -LiteralPath $file
        $acl.SetAccessRuleProtection($true, $true)
        Set-Acl -LiteralPath $file -AclObject $acl

        Grant-YardstickFolderAccess -Path $script:tree -Sid $script:stranger
        Get-YardstickUnwritablePath -ProjectRoot $script:tree -Sid $script:stranger | Should -Be $file

        Grant-YardstickFolderAccess -Path $script:tree -Sid $script:stranger -Item $file, "$script:tree\missing.txt"
        Get-YardstickUnwritablePath -ProjectRoot $script:tree -Sid $script:stranger | Should -BeNullOrEmpty
    }

    It "Changes nothing with -WhatIf" {
        Grant-YardstickFolderAccess -Path $script:tree -Sid $script:stranger -WhatIf
        @(Get-YardstickUnwritablePath -ProjectRoot $script:tree -Sid $script:stranger).Count | Should -Be 4
    }
}


Describe "Test-YardstickUacElevated" {
    It "Reads the elevation type of this session's token" {
        Get-YardstickTokenElevationType | Should -BeIn 'Default', 'Full', 'Limited'
    }

    It "Is false without administrator rights, without reading the token" {
        Mock Test-YardstickAdministrator { $false } -ModuleName YardstickSetup
        Mock Get-YardstickTokenElevationType { 'Full' } -ModuleName YardstickSetup
        Test-YardstickUacElevated | Should -BeFalse
        Should -Invoke Get-YardstickTokenElevationType -ModuleName YardstickSetup -Times 0 -Exactly
    }

    It "Is true only for a UAC-elevated token" {
        Mock Test-YardstickAdministrator { $true } -ModuleName YardstickSetup
        Mock Get-YardstickTokenElevationType { 'Full' } -ModuleName YardstickSetup
        Test-YardstickUacElevated | Should -BeTrue
        Mock Get-YardstickTokenElevationType { 'Default' } -ModuleName YardstickSetup
        Test-YardstickUacElevated | Should -BeFalse
    }
}


Describe "Invoke-YardstickElevatedSetup" {
    BeforeEach {
        # The mock stands in for the elevated process: it reads the request the
        # function wrote and answers with $Global:YardstickTestReply.
        $Global:YardstickTestReply = $null
        $Global:YardstickTestSeen = $null
        Mock Start-Process {
            $text = $ArgumentList -join ' '
            if ($text -notmatch '-File "(?<script>[^"]+)" -RequestPath "(?<request>[^"]+)" -ResultPath "(?<result>[^"]+)" -Key (?<key>\S+)$') { throw "Unexpected arguments: $text" }
            $Global:YardstickTestSeen = [pscustomobject]@{
                Script  = $Matches.script
                Request = $Matches.request
                Key     = $Matches.key
                Actions = @(Get-Content -LiteralPath $Matches.request -Raw | ConvertFrom-Json)
            }
            if ($Global:YardstickTestReply) {
                ConvertTo-Json -InputObject $Global:YardstickTestReply -Depth 4 | Set-Content -LiteralPath $Matches.result
            }
        } -ModuleName YardstickSetup
    }

    AfterAll {
        Remove-Variable -Name YardstickTestReply, YardstickTestSeen -Scope Global -ErrorAction SilentlyContinue
    }

    It "Runs every action in one elevated process and returns each result in order" {
        $Global:YardstickTestReply = @{ Results = @(@{ Success = $true; Message = $null }, @{ Success = $false; Message = 'no Outlook' }); Error = $null }
        $results = @(Invoke-YardstickElevatedSetup -ProjectRoot 'C:\Yardstick' -Action @{ Type = 'GrantFolderAccess'; Path = 'C:\Yardstick' }, @{ Type = 'DisableOutlookPrompts' })

        Should -Invoke Start-Process -ModuleName YardstickSetup -Times 1 -Exactly -ParameterFilter { $Verb -eq 'RunAs' -and $Wait -and $FilePath -eq (Join-Path $PSHOME 'pwsh.exe') }
        $Global:YardstickTestSeen.Script | Should -Be 'C:\Yardstick\Private\Invoke-YardstickElevatedSetup.ps1'
        $Global:YardstickTestSeen.Actions.Type | Should -Be @('GrantFolderAccess', 'DisableOutlookPrompts')
        $results.Success | Should -Be @($true, $false)
        $results[1].Message | Should -Be 'no Outlook'
        Test-Path -LiteralPath (Split-Path $Global:YardstickTestSeen.Request) | Should -BeFalse
    }

    It "Passes a credential as a user name and a password only the key decrypts" {
        $Global:YardstickTestReply = @{ Results = @(@{ Success = $true }); Error = $null }
        $credential = [pscredential]::new('CONTOSO\runner', (ConvertTo-SecureString 'P@ssw0rd!' -AsPlainText -Force))
        Invoke-YardstickElevatedSetup -ProjectRoot 'C:\Yardstick' -Action @{ Type = 'RegisterScheduledTask'; Credential = $credential } | Out-Null

        $sent = $Global:YardstickTestSeen.Actions[0]
        $sent.PSObject.Properties.Name | Should -Not -Contain 'Credential'
        $sent.UserName | Should -Be 'CONTOSO\runner'
        $sent.Password | Should -Not -Match 'P@ssw0rd'
        $key = [Convert]::FromBase64String($Global:YardstickTestSeen.Key)
        (ConvertTo-SecureString -String $sent.Password -Key $key | ConvertFrom-SecureString -AsPlainText) | Should -Be 'P@ssw0rd!'
    }

    It "Fails every action when the administrator prompt is declined" {
        Mock Start-Process { throw 'The operation was canceled by the user.' } -ModuleName YardstickSetup
        $results = @(Invoke-YardstickElevatedSetup -ProjectRoot 'C:\Yardstick' -Action @{ Type = 'A' }, @{ Type = 'B' })
        $results.Success | Should -Be @($false, $false)
        $results[0].Message | Should -Match 'did not grant administrator rights.*canceled by the user'
    }

    It "Fails the actions the elevated process did not report" {
        $results = @(Invoke-YardstickElevatedSetup -ProjectRoot 'C:\Yardstick' -Action @{ Type = 'A' })
        $results[0].Success | Should -BeFalse
        $results[0].Message | Should -Match 'without reporting a result'

        $Global:YardstickTestReply = @{ Results = @(@{ Success = $true }); Error = 'Could not load Yardstick.psd1' }
        $results = @(Invoke-YardstickElevatedSetup -ProjectRoot 'C:\Yardstick' -Action @{ Type = 'A' }, @{ Type = 'B' })
        $results.Success | Should -Be @($true, $false)
        $results[1].Message | Should -Be 'Could not load Yardstick.psd1'
    }
}


Describe "Invoke-YardstickElevatedAction" {
    BeforeAll {
        # The elevated script gets each action back from JSON.
        function ConvertTo-TestAction { param([hashtable]$Action) $Action | ConvertTo-Json | ConvertFrom-Json }
    }

    It "Creates the scheduled task with the decrypted credential" {
        Mock Register-YardstickScheduledTask { } -ModuleName YardstickSetup
        $key = [byte[]](1..32)
        $password = ConvertFrom-SecureString -SecureString (ConvertTo-SecureString 'secret' -AsPlainText -Force) -Key $key
        $action = ConvertTo-TestAction @{ Type = 'RegisterScheduledTask'; ProjectRoot = 'C:\Yardstick'; At = '05:30'; Group = 'Lab'; UserName = 'CONTOSO\runner'; Password = $password }

        Invoke-YardstickElevatedAction -Action $action -Key $key
        Should -Invoke Register-YardstickScheduledTask -ModuleName YardstickSetup -Times 1 -Exactly -ParameterFilter {
            $ProjectRoot -eq 'C:\Yardstick' -and $Group -eq 'Lab' -and $At.Hour -eq 5 -and $At.Minute -eq 30 -and
            $Credential.UserName -eq 'CONTOSO\runner' -and $Credential.GetNetworkCredential().Password -eq 'secret'
        }
    }

    It "Grants folder access" {
        Mock Grant-YardstickFolderAccess { } -ModuleName YardstickSetup
        Invoke-YardstickElevatedAction -Action (ConvertTo-TestAction @{ Type = 'GrantFolderAccess'; Path = 'C:\Yardstick'; Sid = 'S-1-5-21-1'; Item = @('C:\Yardstick\YInstall.log') })
        Should -Invoke Grant-YardstickFolderAccess -ModuleName YardstickSetup -Times 1 -Exactly -ParameterFilter {
            $Path -eq 'C:\Yardstick' -and $Sid -eq 'S-1-5-21-1' -and "$Item" -eq 'C:\Yardstick\YInstall.log'
        }
    }

    It "Rejects an unknown action" {
        { Invoke-YardstickElevatedAction -Action (ConvertTo-TestAction @{ Type = 'FormatDisk' }) } | Should -Throw "*Unknown setup action 'FormatDisk'*"
    }
}


Describe "Private\Invoke-YardstickElevatedSetup.ps1" {
    It "Runs each action and reports every result, including failures" {
        $work = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path "$work\Target" -Force | Out-Null
        $sid = 'S-1-5-21-1000-2000-3000-4242'
        ConvertTo-Json -InputObject @(
            @{ Type = 'GrantFolderAccess'; Path = "$work\Target"; Sid = $sid; Item = @() }
            @{ Type = 'FormatDisk' }
        ) | Set-Content -LiteralPath "$work\request.json"

        # A separate process, as in setup: the script imports the whole module.
        $arguments = "-NoProfile -File `"$script:repoRoot\Private\Invoke-YardstickElevatedSetup.ps1`" -RequestPath `"$work\request.json`" -ResultPath `"$work\result.json`""
        Start-Process -FilePath (Join-Path $PSHOME 'pwsh.exe') -ArgumentList $arguments -Wait -NoNewWindow

        $reply = Get-Content -LiteralPath "$work\result.json" -Raw | ConvertFrom-Json
        $reply.Error | Should -BeNullOrEmpty
        $reply.Results.Success | Should -Be @($true, $false)
        $reply.Results[1].Message | Should -Match 'Unknown setup action'
        Get-YardstickUnwritablePath -ProjectRoot "$work\Target" -Sid $sid | Should -BeNullOrEmpty
    }
}
