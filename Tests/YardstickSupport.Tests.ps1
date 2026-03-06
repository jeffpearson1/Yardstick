BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $modulePath = "$PSScriptRoot\..\Modules\YardstickSupport.psm1"
    Import-Module $modulePath -Force
}

Describe "ArrayToString" {
    It "Converts single-element array" {
        ArrayToString @("hello") | Should -Be "@(hello)"
    }

    It "Converts multi-element array" {
        ArrayToString @("a","b","c") | Should -Be "@(a,b,c)"
    }

    It "Handles numeric values" {
        ArrayToString @(1,2,3) | Should -Be "@(1,2,3)"
    }
}

Describe "Format-FileDetectionVersion" {
    It "1-segment pads to 4" {
        Format-FileDetectionVersion -Version "1" | Should -Be "1.0.0.0"
    }

    It "2-segment pads to 4" {
        Format-FileDetectionVersion -Version "1.2" | Should -Be "1.2.0.0"
    }

    It "3-segment pads to 4" {
        Format-FileDetectionVersion -Version "1.2.3" | Should -Be "1.2.3.0"
    }

    It "4-segment returns as-is" {
        Format-FileDetectionVersion -Version "1.2.3.4" | Should -Be "1.2.3.4"
    }

    It "5+ segment truncates to first 4" {
        Format-FileDetectionVersion -Version "1.2.3.4.5" | Should -Be "1.2.3.4"
    }
}

Describe "Test-VersionExcluded" {
    It "No lock returns false (version allowed)" {
        Test-VersionExcluded -Version "1.2.3" -VersionLock $null | Should -BeFalse
    }

    It "Empty string lock returns false (version allowed)" {
        Test-VersionExcluded -Version "1.2.3" -VersionLock "" | Should -BeFalse
    }

    It "Version matching lock pattern returns false (version allowed)" {
        Test-VersionExcluded -Version "1.2.3" -VersionLock "1.2.x" | Should -BeFalse
    }

    It "Version outside lock pattern returns true (version excluded)" {
        Test-VersionExcluded -Version "2.0.0" -VersionLock "1.x.x" | Should -BeTrue
    }

    It "Exact match returns false (version allowed)" {
        Test-VersionExcluded -Version "1.2.3" -VersionLock "1.2.3" | Should -BeFalse
    }

    It "Capital X wildcard works" {
        Test-VersionExcluded -Version "1.2.3" -VersionLock "1.X.X" | Should -BeFalse
    }

    It "Wildcard in minor only" {
        Test-VersionExcluded -Version "1.3.0" -VersionLock "1.x.0" | Should -BeFalse
    }

    It "Major mismatch is excluded" {
        Test-VersionExcluded -Version "2.2.3" -VersionLock "1.2.x" | Should -BeTrue
    }

    It "Null or empty version returns true (excluded as safety)" {
        Test-VersionExcluded -Version "" -VersionLock "1.x" | Should -BeTrue
    }
}

Describe "Compare-AppVersions" {
    It "Equal versions return 0" {
        Compare-AppVersions -Version1 "1.2.3" -Version2 "1.2.3" | Should -Be 0
    }

    It "Greater returns 1" {
        Compare-AppVersions -Version1 "2.0.0" -Version2 "1.9.9" | Should -Be 1
    }

    It "Lesser returns -1" {
        Compare-AppVersions -Version1 "1.0.0" -Version2 "1.0.1" | Should -Be -1
    }

    It "Different segment counts with implicit zeros" {
        Compare-AppVersions -Version1 "1.2" -Version2 "1.2.0" | Should -Be 0
    }

    It "Strips non-numeric prefixes" {
        Compare-AppVersions -Version1 "v1.2.3" -Version2 "1.2.3" | Should -Be 0
    }

    It "Strips alpha suffixes" {
        Compare-AppVersions -Version1 "1.2.3-beta" -Version2 "1.2.3" | Should -Be 0
    }

    It "Handles large version numbers" {
        Compare-AppVersions -Version1 "2024.1.15" -Version2 "2024.1.14" | Should -Be 1
    }

    It "Throws on null or empty Version1" {
        { Compare-AppVersions -Version1 "" -Version2 "1.0.0" } | Should -Throw "*null or empty*"
    }

    It "Throws on null or empty Version2" {
        { Compare-AppVersions -Version1 "1.0.0" -Version2 "" } | Should -Throw "*null or empty*"
    }
}
