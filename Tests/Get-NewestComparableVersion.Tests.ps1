BeforeAll {
    Import-Module "$PSScriptRoot\..\Modules\VersionPro.psm1" -Force
    Import-Module "$PSScriptRoot\..\Modules\YardstickSupport.psm1" -Force -DisableNameChecking

    function New-App {
        param([string]$DisplayName, $DisplayVersion)
        [PSCustomObject]@{ displayName = $DisplayName; displayVersion = $DisplayVersion }
    }

    # Get-SameAppAllVersions returns ", @(...)". The unary comma only survives as a
    # wrapper until the pipeline unrolls it on return, so callers receive the inner
    # array. Returning from a function here reproduces that exactly -- a direct
    # "$x = , @(...)" assignment would leave the array nested and test nothing real.
    function New-ExistingVersions {
        param([object[]]$Apps)
        return , @($Apps)
    }
}

Describe "Get-NewestComparableVersion" {
    It "Returns the newest version when several apps exist" {
        $apps = New-ExistingVersions @(
            (New-App 'Steam' '3.0.0'),
            (New-App 'Steam (N-1)' '2.10.91.91'),
            (New-App 'Steam (N-2)' '2.9.0.0')
        )
        Get-NewestComparableVersion -ExistingVersions $apps | Should -Be '3.0.0'
    }

    It "Returns the whole version, not its first character, when exactly one app exists" {
        # $ExistingVersions.displayVersion[0] unrolls to a scalar string here and
        # returns [Char] '2'. That silently compared against the wrong value and
        # could republish an older build as if it were newer.
        $apps = New-ExistingVersions @((New-App 'Steam' '2.10.91.91'))

        # Sanity check: the old expression really does return a single character.
        $apps.displayVersion[0] | Should -Be '2'

        $result = Get-NewestComparableVersion -ExistingVersions $apps
        $result | Should -Be '2.10.91.91'
        $result | Should -BeOfType [string]
    }

    It "Skips a leading app with an empty displayVersion" {
        # This is the shape that made the steam recipe fail with
        # "Version2 is null or empty - cannot compare versions".
        $apps = New-ExistingVersions @(
            (New-App 'Ω DETECT - Steam' ''),
            (New-App 'Steam' '2.10.91.91')
        )
        Get-NewestComparableVersion -ExistingVersions $apps | Should -Be '2.10.91.91'
    }

    It "Skips a leading app with a null displayVersion" {
        $apps = New-ExistingVersions @(
            (New-App 'Ω DETECT - Steam' $null),
            (New-App 'Steam' '2.10.91.91')
        )
        Get-NewestComparableVersion -ExistingVersions $apps | Should -Be '2.10.91.91'
    }

    It "Returns null when the only app has no usable version" {
        $apps = New-ExistingVersions @((New-App 'Steam' ''))
        Get-NewestComparableVersion -ExistingVersions $apps | Should -BeNullOrEmpty
    }

    It "Returns null for an empty result set" {
        Get-NewestComparableVersion -ExistingVersions (New-ExistingVersions @()) | Should -BeNullOrEmpty
    }

    It "Returns null for null input" {
        Get-NewestComparableVersion -ExistingVersions $null | Should -BeNullOrEmpty
    }

    It "Produces a value Compare-AppVersions accepts rather than throwing" {
        $apps = New-ExistingVersions @(
            (New-App 'Ω DETECT - Steam' ''),
            (New-App 'Steam' '2.10.91.91')
        )
        $newest = Get-NewestComparableVersion -ExistingVersions $apps
        { Compare-AppVersions '2.10.92.0' $newest } | Should -Not -Throw
        Compare-AppVersions '2.10.92.0' $newest | Should -Be 1
    }
}

Describe "Yardstick.ps1 version comparison" {
    BeforeAll {
        $Global:YardstickSource = Get-Content "$PSScriptRoot\..\Yardstick.ps1" -Raw
    }

    It "Never indexes into a displayVersion projection" {
        # .displayVersion[0] is the unrolling trap this fix removed. Comment lines
        # are stripped first -- the fix documents the old expression in a comment.
        $code = (Get-Content "$PSScriptRoot\..\Yardstick.ps1") |
            Where-Object { $_ -notmatch '^\s*#' }
        ($code | Select-String -Pattern '\.displayVersion\[0\]' -SimpleMatch:$false) |
            Should -BeNullOrEmpty
    }

    It "Guards Compare-AppVersions against a blank existing version" {
        $Global:YardstickSource | Should -Match 'Get-NewestComparableVersion'
        $Global:YardstickSource | Should -Match 'elseif \(-not \$NewestExistingVersion\)'
    }
}
