BeforeAll {
    $script:modulePath = [IO.Path]::GetFullPath("$PSScriptRoot\..\..\Modules\YardstickUserAgent.psm1")
    Import-Module $script:modulePath -Force

    function Get-FeedResponse {
        param([string]$Version)
        [pscustomobject]@{ versions = @([pscustomobject]@{ version = $Version }) }
    }
}

Describe "Get-YardstickChromeUserAgent" {
    BeforeEach {
        # Each test starts with an empty session cache.
        Import-Module $script:modulePath -Force
    }

    It "Uses the major version from Google's stable feed in Chrome's reduced format" {
        Mock Invoke-RestMethod { Get-FeedResponse '163.0.8250.12' } -ModuleName YardstickUserAgent
        Get-YardstickChromeUserAgent |
            Should -Be 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/163.0.0.0 Safari/537.36'
    }

    It "Reads the feed once per session" {
        Mock Invoke-RestMethod { Get-FeedResponse '163.0.8250.12' } -ModuleName YardstickUserAgent
        $null = Get-YardstickChromeUserAgent
        $null = Get-YardstickChromeUserAgent
        Should -Invoke Invoke-RestMethod -ModuleName YardstickUserAgent -Times 1 -Exactly
    }

    It "Reads the feed again when asked to refresh" {
        Mock Invoke-RestMethod { Get-FeedResponse '163.0.8250.12' } -ModuleName YardstickUserAgent
        $null = Get-YardstickChromeUserAgent
        $null = Get-YardstickChromeUserAgent -Refresh
        Should -Invoke Invoke-RestMethod -ModuleName YardstickUserAgent -Times 2 -Exactly
    }

    It "Falls back to the estimate when the feed is unreachable, and caches that too" {
        Mock Invoke-RestMethod { throw 'No such host is known.' } -ModuleName YardstickUserAgent
        $expected = Get-YardstickEstimatedChromeMajorVersion
        Get-YardstickChromeMajorVersion | Should -Be $expected
        $null = Get-YardstickChromeMajorVersion
        Should -Invoke Invoke-RestMethod -ModuleName YardstickUserAgent -Times 1 -Exactly
    }

    It "Falls back to the estimate when the feed returns something unusable" {
        Mock Invoke-RestMethod { [pscustomobject]@{ versions = @() } } -ModuleName YardstickUserAgent
        Get-YardstickChromeMajorVersion | Should -Be (Get-YardstickEstimatedChromeMajorVersion)
    }

    It "Never goes below the last verified version" {
        Mock Invoke-RestMethod { Get-FeedResponse '7.0.500.0' } -ModuleName YardstickUserAgent
        Get-YardstickChromeMajorVersion | Should -BeGreaterOrEqual 157
    }
}

Describe "Get-YardstickEstimatedChromeMajorVersion" {
    It "Returns the last verified version on the day it was verified" {
        Get-YardstickEstimatedChromeMajorVersion -Date '2026-10-23' | Should -Be 157
    }

    It "Adds one release per four weeks after that" {
        Get-YardstickEstimatedChromeMajorVersion -Date '2026-11-19' | Should -Be 157
        Get-YardstickEstimatedChromeMajorVersion -Date '2026-11-20' | Should -Be 158
        Get-YardstickEstimatedChromeMajorVersion -Date '2027-10-23' | Should -Be 170
    }

    It "Does not go backwards for an earlier date" {
        Get-YardstickEstimatedChromeMajorVersion -Date '2020-01-01' | Should -Be 157
    }
}
