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

Describe "Test-YardstickTransientNetworkError" {
    BeforeAll {
        function New-NetworkErrorRecord {
            param([Parameter(Mandatory)][System.Exception]$Exception)
            [System.Management.Automation.ErrorRecord]::new(
                $Exception, 'TestError', [System.Management.Automation.ErrorCategory]::ConnectionError, $null)
        }
    }

    It "Treats a connection timeout as transient" {
        $record = New-NetworkErrorRecord -Exception ([System.Net.Sockets.SocketException]::new(
            [int][System.Net.Sockets.SocketError]::TimedOut))
        Test-YardstickTransientNetworkError -ErrorRecord $record | Should -BeTrue
    }

    It "Treats a refused connection as transient" {
        $record = New-NetworkErrorRecord -Exception ([System.Net.Sockets.SocketException]::new(
            [int][System.Net.Sockets.SocketError]::ConnectionRefused))
        Test-YardstickTransientNetworkError -ErrorRecord $record | Should -BeTrue
    }

    It "Finds a socket fault wrapped in an HttpRequestException" {
        # This is the shape Get-RedirectedUrl actually sees: HttpClient wraps the
        # socket error, and PowerShell wraps that again in the error record.
        $socket = [System.Net.Sockets.SocketException]::new([int][System.Net.Sockets.SocketError]::TimedOut)
        $http = [System.Net.Http.HttpRequestException]::new('The operation has timed out.', $socket)
        $record = New-NetworkErrorRecord -Exception ([System.Management.Automation.MethodInvocationException]::new(
            'Exception calling "GetResult"', $http))
        Test-YardstickTransientNetworkError -ErrorRecord $record | Should -BeTrue
    }

    It "Treats a DNS failure as transient" {
        $record = New-NetworkErrorRecord -Exception ([System.Net.WebException]::new(
            'The remote name could not be resolved.', [System.Net.WebExceptionStatus]::NameResolutionFailure))
        Test-YardstickTransientNetworkError -ErrorRecord $record | Should -BeTrue
    }

    It "Does not treat a protocol error as transient" {
        $record = New-NetworkErrorRecord -Exception ([System.Net.WebException]::new(
            'The remote server returned an error: (404) Not Found.', [System.Net.WebExceptionStatus]::ProtocolError))
        Test-YardstickTransientNetworkError -ErrorRecord $record | Should -BeFalse
    }

    It "Does not treat a recipe's own failure as transient" {
        $record = New-NetworkErrorRecord -Exception ([System.InvalidOperationException]::new(
            "The vendor feed returned an invalid ProductVersion: ''."))
        Test-YardstickTransientNetworkError -ErrorRecord $record | Should -BeFalse
    }

    It "Accepts a bare exception as well as an error record" {
        Test-YardstickTransientNetworkError -ErrorRecord ([System.TimeoutException]::new()) | Should -BeTrue
    }

    It "Returns false for null" {
        Test-YardstickTransientNetworkError -ErrorRecord $null | Should -BeFalse
    }

    It "Survives a self-referential inner exception" {
        # Some .NET exceptions return themselves from InnerException; walking that
        # chain naively hangs forever.
        $looping = [pscustomobject]@{ Message = 'loop' }
        $looping | Add-Member -MemberType ScriptProperty -Name InnerException -Value { $looping }
        Test-YardstickTransientNetworkError -ErrorRecord ([pscustomobject]@{ Exception = $looping }) | Should -BeFalse
    }
}

Describe "Remove-YardstickApp delete retry" {
    BeforeAll {
        $Script:App = [pscustomobject]@{ id = 'app-1'; DisplayName = 'Google Chrome (N-1)' }
    }

    It "Retries while Intune is still catching up, then succeeds" {
        InModuleScope YardstickSupport {
            $Script:attempts = 0
            Mock Clear-YardstickAppLink { @() }
            Mock Write-Log {}
            Mock Start-Sleep {}
            Mock Get-YardstickWin32App { $null }
            Mock Remove-YardstickWin32App {
                $Script:attempts++
                if ($Script:attempts -lt 3) { throw 'Cannot delete this app at this time. Try again shortly.' }
            }

            Remove-YardstickApp -App ([pscustomobject]@{ id = 'app-1'; DisplayName = 'Google Chrome (N-1)' })

            $Script:attempts | Should -Be 3
        }
    }

    It "Retries the relationship-validation wording Intune also uses" {
        InModuleScope YardstickSupport {
            $Script:attempts = 0
            Mock Clear-YardstickAppLink { @() }
            Mock Write-Log {}
            Mock Start-Sleep {}
            Mock Get-YardstickWin32App { $null }
            Mock Remove-YardstickWin32App {
                $Script:attempts++
                if ($Script:attempts -lt 2) { throw 'MetadataBehaviorValidateRelationshipsForAppToDeleteAndThrowAsync failed' }
            }

            Remove-YardstickApp -App ([pscustomobject]@{ id = 'app-1'; DisplayName = 'App' })

            $Script:attempts | Should -Be 2
        }
    }

    It "Gives up after the last attempt" {
        InModuleScope YardstickSupport {
            $Script:attempts = 0
            Mock Clear-YardstickAppLink { @() }
            Mock Write-Log {}
            Mock Start-Sleep {}
            Mock Get-YardstickWin32App { $null }
            Mock Remove-YardstickWin32App {
                $Script:attempts++
                throw 'Cannot delete this app at this time. Try again shortly.'
            }

            { Remove-YardstickApp -App ([pscustomobject]@{ id = 'app-1'; DisplayName = 'App' }) } |
                Should -Throw '*Cannot delete this app at this time*'
            $Script:attempts | Should -Be 4
        }
    }

    It "Does not retry an unrelated failure" {
        InModuleScope YardstickSupport {
            $Script:attempts = 0
            Mock Clear-YardstickAppLink { @() }
            Mock Write-Log {}
            Mock Start-Sleep {}
            Mock Get-YardstickWin32App { $null }
            Mock Remove-YardstickWin32App {
                $Script:attempts++
                throw 'Forbidden'
            }

            { Remove-YardstickApp -App ([pscustomobject]@{ id = 'app-1'; DisplayName = 'App' }) } |
                Should -Throw '*Forbidden*'
            $Script:attempts | Should -Be 1
        }
    }
}
