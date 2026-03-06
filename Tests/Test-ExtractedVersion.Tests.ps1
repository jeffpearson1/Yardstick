BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $modulePath = "$PSScriptRoot\..\Modules\YardstickSupport.psm1"
    Import-Module $modulePath -Force
}

Describe "Test-ExtractedVersion" {
    Context "Valid versions" {
        It "Accepts a standard 3-segment version" {
            $result = Test-ExtractedVersion -Version "1.2.3" -ApplicationId "testapp"
            $result.IsValid | Should -BeTrue
            $result.Errors.Count | Should -Be 0
        }

        It "Accepts a 4-segment version" {
            $result = Test-ExtractedVersion -Version "1.2.3.4" -ApplicationId "testapp"
            $result.IsValid | Should -BeTrue
        }

        It "Accepts a 2-segment version" {
            $result = Test-ExtractedVersion -Version "25.01" -ApplicationId "testapp"
            $result.IsValid | Should -BeTrue
        }

        It "Accepts a version with pre-release suffix" {
            $result = Test-ExtractedVersion -Version "1.2.3-beta" -ApplicationId "testapp"
            $result.IsValid | Should -BeTrue
        }

        It "Accepts a version with v prefix" {
            $result = Test-ExtractedVersion -Version "v2.0.0" -ApplicationId "testapp"
            $result.IsValid | Should -BeTrue
        }

        It "Accepts a year-based version" {
            $result = Test-ExtractedVersion -Version "2025.1.3" -ApplicationId "testapp"
            $result.IsValid | Should -BeTrue
        }
    }

    Context "Null and empty versions" {
        It "Rejects null version" {
            $result = Test-ExtractedVersion -Version $null -ApplicationId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Should -Contain "Version is null or empty after preDownloadScript execution"
        }

        It "Rejects empty string version" {
            $result = Test-ExtractedVersion -Version "" -ApplicationId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Should -Contain "Version is null or empty after preDownloadScript execution"
        }

        It "Rejects whitespace-only version" {
            $result = Test-ExtractedVersion -Version "   " -ApplicationId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Should -Contain "Version is null or empty after preDownloadScript execution"
        }
    }

    Context "HTML contamination" {
        It "Rejects version with HTML tags" {
            $result = Test-ExtractedVersion -Version "<b>1.2.3</b>" -ApplicationId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Where-Object { $_ -match "HTML/XML" } | Should -Not -BeNullOrEmpty
        }

        It "Rejects version with angle brackets" {
            $result = Test-ExtractedVersion -Version ">1.2.3<" -ApplicationId "testapp"
            $result.IsValid | Should -BeFalse
        }
    }

    Context "Invalid characters" {
        It "Rejects version with curly braces" {
            $result = Test-ExtractedVersion -Version "{1.2.3}" -ApplicationId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Where-Object { $_ -match "invalid characters" } | Should -Not -BeNullOrEmpty
        }

        It "Rejects version with no digits" {
            $result = Test-ExtractedVersion -Version "abc.def" -ApplicationId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Should -Contain "Version contains no digits: 'abc.def'"
        }
    }

    Context "Excessive length" {
        It "Rejects excessively long version string" {
            $longVersion = "1." + ("0." * 50) + "1"
            $result = Test-ExtractedVersion -Version $longVersion -ApplicationId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Where-Object { $_ -match "suspiciously long" } | Should -Not -BeNullOrEmpty
        }
    }

    Context "Whitespace warnings" {
        It "Warns on leading whitespace" {
            $result = Test-ExtractedVersion -Version " 1.2.3" -ApplicationId "testapp"
            $result.IsValid | Should -BeTrue
            $result.Warnings | Where-Object { $_ -match "whitespace" } | Should -Not -BeNullOrEmpty
        }

        It "Warns on trailing whitespace" {
            $result = Test-ExtractedVersion -Version "1.2.3 " -ApplicationId "testapp"
            $result.IsValid | Should -BeTrue
            $result.Warnings | Where-Object { $_ -match "whitespace" } | Should -Not -BeNullOrEmpty
        }
    }

    Context "Major version jump detection" {
        It "Warns when major version drops significantly" {
            $result = Test-ExtractedVersion -Version "1.0.0" -ApplicationId "testapp" -ExistingVersion "25.0.1"
            $result.IsValid | Should -BeTrue
            $result.Warnings | Where-Object { $_ -match "Major version dropped" } | Should -Not -BeNullOrEmpty
        }

        It "Does not warn on normal version increment" {
            $result = Test-ExtractedVersion -Version "26.0.0" -ApplicationId "testapp" -ExistingVersion "25.0.1"
            $result.IsValid | Should -BeTrue
            $result.Warnings.Count | Should -Be 0
        }

        It "Does not warn when no existing version provided" {
            $result = Test-ExtractedVersion -Version "1.0.0" -ApplicationId "testapp"
            $result.IsValid | Should -BeTrue
            $result.Warnings.Count | Should -Be 0
        }
    }
}
