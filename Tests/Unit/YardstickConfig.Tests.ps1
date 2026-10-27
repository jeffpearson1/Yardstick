BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $script:repoRoot = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
    Import-Module powershell-yaml -ErrorAction Stop
    Import-Module "$script:repoRoot\Modules\YardstickConfig.psm1" -Force

    # A throwaway Yardstick root: the real layout manifest plus a Recipes folder, so
    # the loader resolves defaults exactly as it would in a fresh clone.
    function New-TestProjectRoot {
        $root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $root | Out-Null
        Copy-Item "$script:repoRoot\Yardstick.Project.psd1" $root
        New-Item -ItemType Directory -Path (Join-Path $root 'Recipes') | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $root 'Local') | Out-Null
        $root
    }

    function Set-TestPreferences {
        param([string]$Root, [string]$Yaml)
        Set-Content -LiteralPath (Join-Path $Root 'Local\preferences.yaml') -Value $Yaml
    }

    # ValidateSet values declared on a parameter in YardstickIntune.psm1.
    function Get-IntuneValidateSet {
        param([string]$Function, [string]$Parameter)
        $ast = [Management.Automation.Language.Parser]::ParseFile("$script:repoRoot\Modules\YardstickIntune.psm1", [ref]$null, [ref]$null)
        $fn = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Function }, $true)
        $param = $fn.Body.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq $Parameter }
        $attr = $param.Attributes | Where-Object { $_.TypeName.Name -eq 'ValidateSet' }
        @($attr.PositionalArguments | ForEach-Object { $_.Value })
    }
}

Describe "Get-YardstickDefaultPreferences" {
    BeforeEach {
        $script:root = New-TestProjectRoot
    }

    It "Defaults every folder to the standard layout under the project root" {
        $defaults = Get-YardstickDefaultPreferences -ProjectRoot $script:root
        $defaults.Recipes | Should -Be (Join-Path $script:root 'Recipes')
        $defaults.Temp | Should -Be (Join-Path $script:root 'Artifacts\Temp')
        $defaults.Buildspace | Should -Be (Join-Path $script:root 'Artifacts\BuildSpace')
        $defaults.Published | Should -Be (Join-Path $script:root 'Artifacts\Packages')
        $defaults.Scripts | Should -Be (Join-Path $script:root 'Scripts')
        $defaults.Tools | Should -Be (Join-Path $script:root 'Tools')
        $defaults.Secrets | Should -Be (Join-Path $script:root 'Local\Secrets')
    }

    It "Falls back to the legacy root Secrets folder when only that one exists" {
        New-Item -ItemType Directory -Path (Join-Path $script:root 'Secrets') | Out-Null
        $defaults = Get-YardstickDefaultPreferences -ProjectRoot $script:root
        $defaults.Secrets | Should -Be (Join-Path $script:root 'Secrets')
    }

    It "Prefers Local\Secrets when both Secrets folders exist" {
        New-Item -ItemType Directory -Path (Join-Path $script:root 'Secrets') | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $script:root 'Local\Secrets') | Out-Null
        $defaults = Get-YardstickDefaultPreferences -ProjectRoot $script:root
        $defaults.Secrets | Should -Be (Join-Path $script:root 'Local\Secrets')
    }

    It "Does not default keys whose absence has its own meaning" {
        $defaults = Get-YardstickDefaultPreferences -ProjectRoot $script:root
        foreach ($key in 'Backup', 'SoftwareDropbox', 'SoftwareArchive', 'adminEmailRecipient', 'credentialExpirationEmailEnabled') {
            $defaults.Contains($key) | Should -BeFalse -Because "$key must stay unset unless the user sets it"
        }
    }

    It "Leaves shared sync off by default" {
        $defaults = Get-YardstickDefaultPreferences -ProjectRoot $script:root
        $defaults.syncSharedContent | Should -BeFalse
        $defaults.Contains('SharedIcons') | Should -BeFalse
        $defaults.Contains('SharedRecipeOverrides') | Should -BeFalse
    }

    It "Supplies defaults that pass validation" {
        $defaults = Get-YardstickDefaultPreferences -ProjectRoot $script:root
        $prefs = @{}
        foreach ($key in $defaults.Keys) { $prefs[$key] = $defaults[$key] }
        (Test-YardstickPreferences -Preferences $prefs).Errors | Should -BeNullOrEmpty
    }
}

Describe "ConvertFrom-YardstickQuotedPath" {
    It "Removes one pair of matching quotes and the spaces around the path" {
        ConvertFrom-YardstickQuotedPath ' "C:\Program Files\Icons" ' | Should -Be 'C:\Program Files\Icons'
        ConvertFrom-YardstickQuotedPath "'\\server\share\Icons'" | Should -Be '\\server\share\Icons'
    }

    It "Leaves unquoted, unmatched and empty values alone" {
        ConvertFrom-YardstickQuotedPath 'C:\Icons' | Should -Be 'C:\Icons'
        ConvertFrom-YardstickQuotedPath '"C:\Icons' | Should -Be '"C:\Icons'
        ConvertFrom-YardstickQuotedPath "`"C:\Icons'" | Should -Be "`"C:\Icons'"
        ConvertFrom-YardstickQuotedPath '' | Should -Be ''
        ConvertFrom-YardstickQuotedPath $null | Should -Be ''
    }
}

Describe "Import-YardstickPreferences" {
    BeforeEach {
        $script:root = New-TestProjectRoot
    }

    It "Is not a development-only instance unless the file says so" {
        (Import-YardstickPreferences -ProjectRoot $script:root -AllowMissing).developmentOnly | Should -BeFalse
        Set-TestPreferences -Root $script:root -Yaml 'developmentOnly: true'
        (Import-YardstickPreferences -ProjectRoot $script:root).developmentOnly | Should -BeTrue
    }

    It "Drops quotes around folder and file paths" {
        $icons = Join-Path $TestDrive 'Quoted Icons'
        Set-TestPreferences -Root $script:root -Yaml @"
Icons: '"$icons"'
SharedRecipeOverrides: '"\\server\share\recipeOverrides.yaml"'
"@
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.Icons | Should -Be $icons
        $prefs.SharedRecipeOverrides | Should -Be '\\server\share\recipeOverrides.yaml'
    }

    It "Throws a message that points at the installer when no file exists" {
        { Import-YardstickPreferences -ProjectRoot $script:root } | Should -Throw "*Install-Yardstick.ps1*"
    }

    It "Returns the defaults with -AllowMissing when no file exists" {
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root -AllowMissing
        $prefs | Should -BeOfType [hashtable]
        $prefs.defaultInstallExperience | Should -Be 'system'
        $prefs.Temp | Should -Be (Join-Path $script:root 'Artifacts\Temp')
    }

    It "Treats an empty file as all defaults" {
        Set-TestPreferences -Root $script:root -Yaml ''
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.defaultNumVersionsToKeep | Should -Be 2
    }

    It "Lets explicit values win over defaults" {
        Set-TestPreferences -Root $script:root -Yaml @"
defaultNumVersionsToKeep: 5
defaultRestartBehavior: force
emailNotificationEnabled: true
defaultAutoUpdate: false
"@
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.defaultNumVersionsToKeep | Should -Be 5
        $prefs.defaultRestartBehavior | Should -Be 'force'
        $prefs.emailNotificationEnabled | Should -BeTrue
        $prefs.defaultAutoUpdate | Should -BeFalse
    }

    It "Reads a renamed key under its old name" {
        Set-TestPreferences -Root $script:root -Yaml @"
dependentApplicationBlacklist:
  - Old Bundle
"@
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        @($prefs.dependentApplicationBlocklist) | Should -Be @('Old Bundle')
    }

    It "Prefers the new name when a file has both" {
        Set-TestPreferences -Root $script:root -Yaml @"
dependentApplicationBlacklist:
  - Old Bundle
dependentApplicationBlocklist:
  - New Bundle
"@
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        @($prefs.dependentApplicationBlocklist) | Should -Be @('New Bundle')
    }

    It "Fills blank keys from the defaults" {
        Set-TestPreferences -Root $script:root -Yaml @"
defaultInstallExperience:
emailSubject: "   "
Temp:
"@
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.defaultInstallExperience | Should -Be 'system'
        $prefs.emailSubject | Should -Be 'Yardstick Application Update Report'
        $prefs.Temp | Should -Be (Join-Path $script:root 'Artifacts\Temp')
    }

    It "Leaves optional folders blank when they are blank" {
        Set-TestPreferences -Root $script:root -Yaml @"
Backup:
SoftwareDropbox:
"@
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.Backup | Should -BeNullOrEmpty
        $prefs.SoftwareDropbox | Should -BeNullOrEmpty
        $prefs.ContainsKey('SoftwareArchive') | Should -BeFalse
    }

    It "Resolves the shared icon folder and shared overrides file" {
        Set-TestPreferences -Root $script:root -Yaml @"
syncSharedContent: true
SharedIcons: Shared\Icons
SharedRecipeOverrides: \\server\share\Yardstick\recipeOverrides.yaml
"@
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.syncSharedContent | Should -BeTrue
        $prefs.SharedIcons | Should -Be (Join-Path $script:root 'Shared\Icons')
        $prefs.SharedRecipeOverrides | Should -Be '\\server\share\Yardstick\recipeOverrides.yaml'
    }

    It "Resolves relative folder paths against the project root" {
        Set-TestPreferences -Root $script:root -Yaml @"
Temp: Work\Temp
Backup: .\Backups
"@
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.Temp | Should -Be (Join-Path $script:root 'Work\Temp')
        $prefs.Backup | Should -Be (Join-Path $script:root 'Backups')
    }

    It "Passes absolute and UNC paths through unchanged" {
        Set-TestPreferences -Root $script:root -Yaml @"
Temp: C:\YardstickWork\Temp
Backup: \\fileserver\share\Yardstick Backup
"@
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.Temp | Should -Be 'C:\YardstickWork\Temp'
        $prefs.Backup | Should -Be '\\fileserver\share\Yardstick Backup'
    }

    It "Expands environment variables in folder paths" {
        $env:YARDSTICK_TEST_DIR = 'D:\YsTest'
        try {
            # Single-quoted in YAML: double quotes would read "\T" as an escape sequence.
            Set-TestPreferences -Root $script:root -Yaml "Temp: '%YARDSTICK_TEST_DIR%\Temp'"
            $prefs = Import-YardstickPreferences -ProjectRoot $script:root
            $prefs.Temp | Should -Be 'D:\YsTest\Temp'
        } finally {
            Remove-Item Env:\YARDSTICK_TEST_DIR
        }
    }

    It "Keeps a legacy file with absolute paths exactly as written" {
        Set-TestPreferences -Root $script:root -Yaml @"
Recipes: \\server\Yardstick\Recipes
Temp: \\server\Yardstick\Temp
Backup: \\server\Backups
defaultMinOSVersion: W11_22H2
defaultScopeTags:
  - Workstations
defaultDeploymentGroups:
  - 0b0e2f1c-7b1a-4e33-9d43-6f0f6f1a2b3c
"@
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.Recipes | Should -Be '\\server\Yardstick\Recipes'
        $prefs.Temp | Should -Be '\\server\Yardstick\Temp'
        $prefs.Backup | Should -Be '\\server\Backups'
        $prefs.defaultMinOSVersion | Should -Be 'W11_22H2'
        @($prefs.defaultScopeTags) | Should -Be @('Workstations')
        @($prefs.defaultDeploymentGroups) | Should -Be @('0b0e2f1c-7b1a-4e33-9d43-6f0f6f1a2b3c')
    }

    It "Reads the legacy root preferences.yaml when Local\preferences.yaml is absent" {
        Set-Content -LiteralPath (Join-Path $script:root 'preferences.yaml') -Value 'defaultOwner: Legacy Owner'
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.defaultOwner | Should -Be 'Legacy Owner'
    }

    It "Reads an explicit -Path" {
        $file = Join-Path $script:root 'other.yaml'
        Set-Content -LiteralPath $file -Value 'defaultOwner: Explicit'
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root -Path $file
        $prefs.defaultOwner | Should -Be 'Explicit'
    }

    It "Returns a case-insensitive hashtable" {
        Set-TestPreferences -Root $script:root -Yaml 'defaultOwner: Someone'
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.DEFAULTOWNER | Should -Be 'Someone'
        $prefs.temp | Should -Be $prefs.Temp
    }

    It "Leaves template placeholders unresolved so validation can report them" {
        Set-TestPreferences -Root $script:root -Yaml 'Temp: <REPLACE_WITH_YARDSTICK_ROOT>\Artifacts\Temp'
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        $prefs.Temp | Should -Be '<REPLACE_WITH_YARDSTICK_ROOT>\Artifacts\Temp'
    }

    It "Creates the working folders with -CreateFolders, but never Recipes" {
        Remove-Item (Join-Path $script:root 'Recipes')
        Set-TestPreferences -Root $script:root -Yaml ''
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root -CreateFolders
        foreach ($key in 'Temp', 'Buildspace', 'Published', 'Scripts', 'Icons', 'Tools') {
            Test-Path -LiteralPath $prefs[$key] -PathType Container | Should -BeTrue -Because "$key should be created"
        }
        Test-Path -LiteralPath $prefs.Recipes | Should -BeFalse
    }

    It "Does not create folders without -CreateFolders" {
        Set-TestPreferences -Root $script:root -Yaml ''
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root
        Test-Path -LiteralPath $prefs.Temp | Should -BeFalse
    }

    It "Loads the shipped example file without errors" {
        $prefs = Import-YardstickPreferences -ProjectRoot $script:root -Path "$script:repoRoot\Config\Preferences.example.yaml"
        $result = Test-YardstickPreferences -Preferences $prefs
        $result.Errors | Should -BeNullOrEmpty
        $result.Warnings | Should -BeNullOrEmpty
    }
}

Describe "Initialize-YardstickFolders" {
    It "Returns only the folders it created" {
        $existing = Join-Path $TestDrive 'init-existing'
        New-Item -ItemType Directory -Path $existing | Out-Null
        $new = Join-Path $TestDrive 'init-new'
        $created = @(Initialize-YardstickFolders -Preferences @{ Temp = $existing; Tools = $new })
        $created | Should -Be @($new)
    }

    It "Skips blank values and placeholders" {
        $created = @(Initialize-YardstickFolders -Preferences @{ Temp = ''; Tools = '<REPLACE_WITH_YARDSTICK_ROOT>\Tools' })
        $created | Should -BeNullOrEmpty
    }

    It "Honours -WhatIf" {
        $target = Join-Path $TestDrive 'init-whatif'
        Initialize-YardstickFolders -Preferences @{ Temp = $target } -WhatIf | Out-Null
        Test-Path -LiteralPath $target | Should -BeFalse
    }
}

Describe "Test-YardstickPreferences" {
    BeforeAll {
        $script:validRoot = New-TestProjectRoot
    }

    BeforeEach {
        $script:prefs = Import-YardstickPreferences -ProjectRoot $script:validRoot -AllowMissing
    }

    It "Accepts the defaults" {
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.IsValid | Should -BeTrue
        $result.Errors | Should -BeNullOrEmpty
    }

    It "Reports an unreplaced placeholder by key" {
        $script:prefs.Tools = '<REPLACE_WITH_YARDSTICK_ROOT>\Tools'
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.IsValid | Should -BeFalse
        $result.Errors | Should -Match "'Tools'.*<REPLACE_WITH_YARDSTICK_ROOT>"
    }

    It "Warns about a renamed key without failing" {
        $script:prefs.dependentApplicationBlacklist = @('Old Bundle')
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.IsValid | Should -BeTrue
        $result.Warnings | Should -Match "'dependentApplicationBlacklist' has been renamed to 'dependentApplicationBlocklist'"
    }

    It "Reports a placeholder inside a list" {
        $script:prefs.defaultScopeTags = @('Default', '<REPLACE_WITH_SCOPE_TAG>')
        (Test-YardstickPreferences -Preferences $script:prefs).Errors | Should -Match 'defaultScopeTags'
    }

    It "Reports a missing Recipes folder" {
        $script:prefs.Recipes = Join-Path $TestDrive 'no-such-recipes'
        (Test-YardstickPreferences -Preferences $script:prefs).Errors | Should -Match 'Recipes folder'
    }

    It "Reports a value Intune does not accept for <Key>" -TestCases @(
        @{ Key = 'defaultInstallExperience'; Value = 'machine' }
        @{ Key = 'defaultRestartBehavior'; Value = 'reboot' }
        @{ Key = 'defaultArchitecture'; Value = 'amd64' }
        @{ Key = 'defaultMinOSVersion'; Value = 'W11_25H2_typo' }
    ) {
        param($Key, $Value)
        $script:prefs[$Key] = $Value
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.IsValid | Should -BeFalse
        $result.Errors | Should -Match "'$Key' is '$Value'"
    }

    It "Reports a non-GUID entry in <Key>" -TestCases @(
        @{ Key = 'defaultDeploymentGroups' }
        @{ Key = 'defaultGroupsSkipAutoUpdates' }
    ) {
        param($Key)
        $script:prefs[$Key] = @('0b0e2f1c-7b1a-4e33-9d43-6f0f6f1a2b3c', 'Group IDs go here')
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.Errors | Should -HaveCount 1
        $result.Errors[0] | Should -Match "'$Key' contains 'Group IDs go here'"
    }

    It "Accepts empty group lists" {
        $script:prefs.defaultDeploymentGroups = @()
        $script:prefs.defaultGroupsSkipAutoUpdates = $null
        (Test-YardstickPreferences -Preferences $script:prefs).IsValid | Should -BeTrue
    }

    It "Reports defaultNumVersionsToKeep below 1" {
        $script:prefs.defaultNumVersionsToKeep = 0
        (Test-YardstickPreferences -Preferences $script:prefs).Errors | Should -Match 'defaultNumVersionsToKeep'
    }

    It "Reports a non-numeric date offset" {
        $script:prefs.defaultDeadlineDateOffset = 'next week'
        (Test-YardstickPreferences -Preferences $script:prefs).Errors | Should -Match 'defaultDeadlineDateOffset'
    }

    It "Warns with Windows' reason when an optional folder cannot be opened" {
        $script:prefs.Backup = Join-Path $TestDrive 'no-such-backup'
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.IsValid | Should -BeTrue
        $result.Warnings | Should -HaveCount 1
        $result.Warnings[0] | Should -Match "^The Backup folder '.*no-such-backup' is not reachable: .+"
    }

    It "Accepts an optional folder that exists" {
        $script:prefs.Backup = (New-Item -ItemType Directory -Path (Join-Path $TestDrive 'backup-ok') -Force).FullName
        (Test-YardstickPreferences -Preferences $script:prefs).Warnings | Should -BeNullOrEmpty
    }

    It "Says nothing about email settings while email is disabled" {
        $script:prefs.emailDeliveryMethod = 'smtp'
        $script:prefs.smtpServer = $null
        (Test-YardstickPreferences -Preferences $script:prefs).Warnings | Should -BeNullOrEmpty
    }

    It "Warns about missing SMTP settings when email is enabled" {
        $script:prefs.emailNotificationEnabled = $true
        $script:prefs.emailDeliveryMethod = 'smtp'
        $script:prefs.emailRecipient = 'it@example.org'
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.IsValid | Should -BeTrue
        $result.Warnings | Should -HaveCount 2
        ($result.Warnings -join "`n") | Should -Match 'smtpServer'
        ($result.Warnings -join "`n") | Should -Match 'emailSendFromAddress'
    }

    It "Warns when email is enabled with no recipient" {
        $script:prefs.emailNotificationEnabled = $true
        (Test-YardstickPreferences -Preferences $script:prefs).Warnings | Should -Match 'emailRecipient'
    }

    It "Warns about an unreachable optional folder" {
        $script:prefs.Backup = Join-Path $TestDrive 'no-such-backup'
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.IsValid | Should -BeTrue
        $result.Warnings | Should -Match 'Backup folder'
    }

    It "Ignores unreachable shared paths while shared sync is off" {
        $script:prefs.SharedIcons = Join-Path $TestDrive 'no-such-icons'
        $script:prefs.SharedRecipeOverrides = Join-Path $TestDrive 'no-such-overrides.yaml'
        (Test-YardstickPreferences -Preferences $script:prefs).Warnings | Should -BeNullOrEmpty
    }

    It "Warns when shared sync is on with nothing to sync" {
        $script:prefs.syncSharedContent = $true
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.IsValid | Should -BeTrue
        $result.Warnings | Should -Match "neither 'SharedIcons' nor 'SharedRecipeOverrides'"
    }

    It "Warns about unreachable shared paths while shared sync is on" {
        $script:prefs.syncSharedContent = $true
        $script:prefs.SharedIcons = Join-Path $TestDrive 'no-such-icons'
        $script:prefs.SharedRecipeOverrides = Join-Path $TestDrive 'no-such-share\recipeOverrides.yaml'
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.IsValid | Should -BeTrue
        $result.Warnings | Should -HaveCount 2
        ($result.Warnings -join "`n") | Should -Match 'SharedIcons folder'
        ($result.Warnings -join "`n") | Should -Match 'SharedRecipeOverrides file'
    }

    It "Accepts a shared overrides file that does not exist yet in a reachable folder" {
        $script:prefs.syncSharedContent = $true
        $script:prefs.SharedRecipeOverrides = Join-Path $TestDrive 'not-created-yet.yaml'
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.IsValid | Should -BeTrue
        $result.Warnings | Should -BeNullOrEmpty
    }

    It "Rejects a shared overrides path that is a folder" {
        $script:prefs.syncSharedContent = $true
        $script:prefs.SharedRecipeOverrides = $TestDrive
        $result = Test-YardstickPreferences -Preferences $script:prefs
        $result.IsValid | Should -BeFalse
        $result.Errors | Should -Match 'is a folder'
    }
}

Describe "Allowed values stay in step with YardstickIntune" {
    It "<Key> matches the ValidateSet on <Function> -<Parameter>" -TestCases @(
        @{ Key = 'defaultInstallExperience'; Function = 'Add-YardstickWin32App'; Parameter = 'InstallExperience' }
        @{ Key = 'defaultRestartBehavior'; Function = 'Add-YardstickWin32App'; Parameter = 'RestartBehavior' }
        @{ Key = 'defaultArchitecture'; Function = 'New-YardstickWin32AppRequirementRule'; Parameter = 'Architecture' }
        @{ Key = 'defaultMinOSVersion'; Function = 'New-YardstickWin32AppRequirementRule'; Parameter = 'MinimumSupportedWindowsRelease' }
    ) {
        param($Key, $Function, $Parameter)
        $expected = Get-IntuneValidateSet -Function $Function -Parameter $Parameter
        $expected | Should -Not -BeNullOrEmpty
        $actual = InModuleScope YardstickConfig -Parameters @{ Key = $Key } { param($Key) $Script:AllowedValues[$Key] }
        $actual | Should -Be $expected
    }
}
