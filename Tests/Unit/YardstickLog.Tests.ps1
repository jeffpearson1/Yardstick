BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $script:modulePath = [IO.Path]::GetFullPath("$PSScriptRoot\..\..\Modules\YardstickLog.psm1")
    Import-Module $script:modulePath -Force

    function Get-LogLine {
        @(Get-Content -LiteralPath (Join-Path $Global:LogLocation $Global:LogFile))
    }

    # Write-Host lands in the information stream; return just the echoed text.
    function Get-ConsoleLine {
        param([scriptblock]$Script)
        @(& $Script 6>&1 | ForEach-Object { "$_" })
    }

    $script:stamp = '\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}'
}

AfterAll {
    Clear-YardstickLogContext
}

Describe "Write-Log" {
    BeforeEach {
        Clear-YardstickLogContext
        Write-Log -Init
    }

    It "starts a run with the header the report parser splits on" {
        $lines = Get-LogLine
        $lines.Count | Should -Be 3
        $lines[0] | Should -Match '^#{5,}$'
        $lines[1] | Should -Match "^LOGGING STARTED AT $script:stamp$"
        $lines[2] | Should -Match '^#{5,}$'
    }

    It "writes the original line format when no level or app is given" {
        Write-Log "Starting update for 7zip..." 6>$null
        (Get-LogLine)[-1] | Should -Match "^$script:stamp - Starting update for 7zip\.\.\.$"
    }

    It "adds the severity prefix for -Level <Level>" -TestCases @(
        @{ Level = 'Error'; Prefix = 'ERROR: ' }
        @{ Level = 'Warning'; Prefix = 'WARNING: ' }
        @{ Level = 'Info'; Prefix = '' }
        @{ Level = 'Detail'; Prefix = '' }
    ) {
        Write-Log "something happened" -Level $Level 6>$null
        (Get-LogLine)[-1] | Should -Match "^$script:stamp - $([regex]::Escape($Prefix))something happened$"
    }

    It "tags lines with the app while a context is set, ahead of the severity" {
        Set-YardstickLogContext '7zip.arm64'
        Write-Log "Download failed" -Level Error 6>$null
        (Get-LogLine)[-1] | Should -Match "^$script:stamp - \[7zip\.arm64\] ERROR: Download failed$"
    }

    It "leaves the tag off with -NoContext and after the context is cleared" {
        Set-YardstickLogContext 'firefox'
        Write-Log "Backed up other-app" -NoContext 6>$null
        (Get-LogLine)[-1] | Should -Match "^$script:stamp - Backed up other-app$"

        Clear-YardstickLogContext
        Write-Log "Run finished." 6>$null
        (Get-LogLine)[-1] | Should -Match "^$script:stamp - Run finished\.$"
    }

    It "echoes Info, Warning and Error lines to the console" {
        $console = Get-ConsoleLine { Write-Log "info line"; Write-Log "warn line" -Level Warning; Write-Log "error line" -Level Error }
        $console.Count | Should -Be 3
        $console[1] | Should -Match 'WARNING: warn line$'
    }

    It "writes Detail lines to the file but keeps them off the console by default" {
        $console = Get-ConsoleLine { Write-Log "routine step" -Level Detail }
        $console | Should -BeNullOrEmpty
        (Get-LogLine)[-1] | Should -Match 'routine step$'
    }

    It "shows Detail lines on the console after -Init -ShowDetail, until the next -Init" {
        Write-Log -Init -ShowDetail
        Get-ConsoleLine { Write-Log "routine step" -Level Detail } | Should -Match 'routine step$'

        Write-Log -Init
        Get-ConsoleLine { Write-Log "routine step" -Level Detail } | Should -BeNullOrEmpty
    }

    It "writes nothing for empty content" {
        Write-Log "" 6>$null
        (Get-LogLine).Count | Should -Be 3
    }

    It "still echoes to the console when the log location is not set" {
        $location = $Global:LogLocation
        try {
            $Global:LogLocation = $null
            $console = Get-ConsoleLine { Write-Log "no file for this" }
            $console[-1] | Should -Match "^$script:stamp - no file for this$"
        } finally {
            $Global:LogLocation = $location
        }
    }

    It "loses no lines when two threads append at the same time" {
        $jobs = 1..2 | ForEach-Object {
            Start-ThreadJob -ArgumentList $script:modulePath, $Global:LogLocation, $Global:LogFile, $_ -ScriptBlock {
                param($ModulePath, $Location, $File, $Thread)
                Import-Module $ModulePath
                $Global:LogLocation = $Location
                $Global:LogFile = $File
                1..150 | ForEach-Object { Write-Log "thread $Thread line $_" 6>$null }
            }
        }
        $jobs | Wait-Job -Timeout 120 | Receive-Job -ErrorAction Stop 6>$null | Out-Null
        $jobs | Remove-Job -Force

        @(Get-LogLine | Where-Object { $_ -match ' - thread \d line \d+$' }).Count | Should -Be 300
    }
}

Describe "Set-YardstickLogContext" {
    It "treats an empty app as clearing the context" {
        Set-YardstickLogContext 'firefox'
        Set-YardstickLogContext ''
        Get-YardstickLogContext | Should -BeNullOrEmpty
    }
}
