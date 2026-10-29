BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickGitHub.psm1" -Force

    # Mock bodies may run outside this file's scope, so the helper is global.
    function global:New-GitHubHttpError {
        param([int]$Status, [hashtable]$Headers = @{}, [string]$Body, [string]$Uri)
        $response = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]$Status)
        foreach ($name in $Headers.Keys) {
            $response.Headers.TryAddWithoutValidation($name, [string]$Headers[$name]) | Out-Null
        }
        if ($Uri) { $response.RequestMessage = [System.Net.Http.HttpRequestMessage]::new('GET', $Uri) }
        $exception = [Microsoft.PowerShell.Commands.HttpResponseException]::new("Response status code does not indicate success: $Status.", $response)
        # Invoke-RestMethod puts the response body in ErrorDetails.
        $record = [System.Management.Automation.ErrorRecord]::new($exception, 'WebCmdletWebResponseException', 'InvalidOperation', $null)
        if ($Body) { $record.ErrorDetails = [System.Management.Automation.ErrorDetails]::new($Body) }
        $record
    }
}

AfterAll {
    Remove-Item Function:\New-GitHubHttpError -ErrorAction SilentlyContinue
    Remove-Variable GitHubTestCalls -Scope Global -ErrorAction SilentlyContinue
}

Describe "Invoke-YardstickGitHubApi" {
    BeforeEach {
        Reset-YardstickGitHubSession
        $Global:GitHubTestCalls = 0
        Mock Get-YardstickGitHubToken { $null } -ModuleName YardstickGitHub
        Mock Invoke-RestMethod { [pscustomobject]@{ tag_name = 'v1.0.0' } } -ModuleName YardstickGitHub
        Mock Start-Sleep {} -ModuleName YardstickGitHub
        Mock Write-Log {} -ModuleName YardstickGitHub
    }

    It "Sends the stored token, Yardstick's user agent and the API version" {
        Mock Get-YardstickGitHubToken { 'github_pat_test' } -ModuleName YardstickGitHub
        (Invoke-YardstickGitHubApi 'repos/owner/app/releases/latest').tag_name | Should -Be 'v1.0.0'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter {
            $Uri -eq 'https://api.github.com/repos/owner/app/releases/latest' -and
            $Headers.Authorization -eq 'Bearer github_pat_test' -and
            $Headers['User-Agent'] -match '^Yardstick ' -and
            $Headers['X-GitHub-Api-Version'] -eq '2022-11-28'
        }
    }

    It "Calls anonymously when no token is stored" {
        Invoke-YardstickGitHubApi 'repos/owner/app/releases' | Out-Null
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter {
            -not $Headers.ContainsKey('Authorization')
        }
    }

    It "Reads the token from Credential Manager once per session" {
        Mock Get-YardstickGitHubToken { 'github_pat_test' } -ModuleName YardstickGitHub
        Invoke-YardstickGitHubApi 'repos/owner/app/releases' | Out-Null
        Invoke-YardstickGitHubApi 'repos/owner/other/releases' | Out-Null
        Should -Invoke Get-YardstickGitHubToken -ModuleName YardstickGitHub -Times 1 -Exactly
    }

    It "Accepts a full API URL or a path with a leading slash" {
        Invoke-YardstickGitHubApi 'https://api.github.com/repos/owner/app/releases?per_page=30' | Out-Null
        Invoke-YardstickGitHubApi '/rate_limit' | Out-Null
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter { $Uri -eq 'https://api.github.com/repos/owner/app/releases?per_page=30' }
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter { $Uri -eq 'https://api.github.com/rate_limit' }
    }

    It "Refuses <Uri>, so the token cannot leave api.github.com" -ForEach @(
        @{ Uri = 'https://github.com/owner/app/releases' }
        @{ Uri = 'http://api.github.com/repos/owner/app' }
        @{ Uri = 'https://raw.githubusercontent.com/owner/app/main/file.yaml' }
        @{ Uri = 'https://api.github.com.example.com/repos/owner/app' }
    ) {
        Mock Get-YardstickGitHubToken { 'github_pat_test' } -ModuleName YardstickGitHub
        { Invoke-YardstickGitHubApi $Uri } | Should -Throw '*only calls https://api.github.com*'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 0 -Exactly
    }

    It "Follows the global web request timeout" {
        $saved = $Global:PSDefaultParameterValues['Invoke-RestMethod:TimeoutSec']
        try {
            $Global:PSDefaultParameterValues['Invoke-RestMethod:TimeoutSec'] = 33
            Invoke-YardstickGitHubApi 'rate_limit' | Out-Null
            Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter { $TimeoutSec -eq 33 }
        } finally {
            if ($null -eq $saved) {
                $Global:PSDefaultParameterValues.Remove('Invoke-RestMethod:TimeoutSec')
            } else {
                $Global:PSDefaultParameterValues['Invoke-RestMethod:TimeoutSec'] = $saved
            }
        }
    }

    It "Drops a rejected token for the session and retries anonymously" {
        Mock Get-YardstickGitHubToken { 'github_pat_expired' } -ModuleName YardstickGitHub
        Mock Invoke-RestMethod {
            if ($Headers.ContainsKey('Authorization')) { throw (New-GitHubHttpError -Status 401) }
            [pscustomobject]@{ tag_name = 'v1.0.0' }
        } -ModuleName YardstickGitHub

        (Invoke-YardstickGitHubApi 'repos/owner/app/releases/latest').tag_name | Should -Be 'v1.0.0'
        Invoke-YardstickGitHubApi 'repos/owner/other/releases/latest' | Out-Null

        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter { $Headers.ContainsKey('Authorization') }
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 2 -Exactly -ParameterFilter { -not $Headers.ContainsKey('Authorization') }
        Should -Invoke Write-Log -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter { $Level -eq 'Warning' -and $Content -match 'rejected' }
    }

    It "Passes other errors through unchanged and keeps calling GitHub" {
        Mock Invoke-RestMethod { throw (New-GitHubHttpError -Status 404) } -ModuleName YardstickGitHub
        { Invoke-YardstickGitHubApi 'repos/owner/missing/releases' } | Should -Throw '*404*'
        { Invoke-YardstickGitHubApi 'repos/owner/missing/releases' } | Should -Throw '*404*'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 2 -Exactly
    }

    It "Holds off until the rate limit resets once it is used up" {
        Reset-YardstickGitHubSession -DeferThrottledCalls
        $reset = [DateTimeOffset]::Now.AddMinutes(37)
        $Global:GitHubTestReset = $reset.ToUnixTimeSeconds()
        Mock Invoke-RestMethod {
            throw (New-GitHubHttpError -Status 403 -Headers @{ 'x-ratelimit-remaining' = '0'; 'x-ratelimit-reset' = $Global:GitHubTestReset })
        } -ModuleName YardstickGitHub

        { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } |
            Should -Throw "*GitHub throttled*wait until $($reset.LocalDateTime.ToString('HH:mm:ss'))*"
        { Invoke-YardstickGitHubApi 'repos/owner/other/releases' } | Should -Throw '*Skipped*wait until*'

        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly
        Should -Invoke Write-Log -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter {
            $Level -eq 'Warning' -and $Content -match 'rate limit used up' -and $Content -match 'Set-YardstickCredential\.ps1 -GitHub'
        }
        (Get-YardstickGitHubHold).ToString('s') | Should -Be $reset.LocalDateTime.ToString('s')
        Remove-Variable GitHubTestReset -Scope Global
    }

    It "Honours Retry-After across calls and never sleeps in a run" {
        Reset-YardstickGitHubSession -DeferThrottledCalls
        Mock Invoke-RestMethod { throw (New-GitHubHttpError -Status 429 -Headers @{ 'retry-after' = '120' }) } -ModuleName YardstickGitHub

        $before = Get-Date
        { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } | Should -Throw '*GitHub throttled*'
        { Invoke-YardstickGitHubApi 'repos/owner/other/releases' } | Should -Throw '*Skipped*'

        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly
        Should -Invoke Start-Sleep -ModuleName YardstickGitHub -Times 0 -Exactly
        $wait = ((Get-YardstickGitHubHold) - $before).TotalSeconds
        $wait | Should -BeGreaterOrEqual 119
        $wait | Should -BeLessThan 125
    }

    It "Throws an error the runner can recognise, and reports the throttle" {
        Reset-YardstickGitHubSession -DeferThrottledCalls
        Mock Invoke-RestMethod { throw (New-GitHubHttpError -Status 429 -Headers @{ 'retry-after' = '30' }) } -ModuleName YardstickGitHub

        $caught = $null
        try { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } catch { $caught = $_ }
        $caught.FullyQualifiedErrorId | Should -BeLike 'YardstickGitHubThrottled*'

        $throttle = Get-YardstickGitHubThrottle
        $throttle.Host | Should -Be 'api.github.com'
        $throttle.Status | Should -Be 429
        $throttle.Uri | Should -Be 'https://api.github.com/repos/owner/app/releases'

        Clear-YardstickGitHubThrottle
        Get-YardstickGitHubThrottle | Should -BeNullOrEmpty
        Get-YardstickGitHubHold | Should -Not -BeNullOrEmpty -Because 'clearing the throttle keeps the hold'
    }

    It "Reports a held call made by a later recipe as a throttle too" {
        Reset-YardstickGitHubSession -DeferThrottledCalls
        Mock Invoke-RestMethod { throw (New-GitHubHttpError -Status 429 -Headers @{ 'retry-after' = '30' }) } -ModuleName YardstickGitHub
        { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } | Should -Throw
        Clear-YardstickGitHubThrottle

        { Invoke-YardstickGitHubApi 'repos/owner/other/releases' } | Should -Throw '*Skipped*'
        $throttle = Get-YardstickGitHubThrottle
        $throttle.Uri | Should -Be 'https://api.github.com/repos/owner/other/releases'
        $throttle.RetryAt | Should -Be (Get-YardstickGitHubHold)
    }

    It "Logs GitHub's explanation and the request ID" {
        Reset-YardstickGitHubSession -DeferThrottledCalls
        Mock Invoke-RestMethod {
            throw (New-GitHubHttpError -Status 403 -Headers @{ 'x-github-request-id' = 'ABCD:1234' } -Body '{"message":"You have exceeded a secondary rate limit. Please wait a few minutes.","documentation_url":"https://docs.github.com"}')
        } -ModuleName YardstickGitHub

        { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } | Should -Throw '*GitHub throttled*'
        Should -Invoke Write-Log -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter {
            $Level -eq 'Warning' -and
            $Content -match 'HTTP 403' -and
            $Content -match 'GitHub said: You have exceeded a secondary rate limit\. Please wait a few minutes\.' -and
            $Content -match 'request ID ABCD:1234'
        }
    }

    It "Waits a minute for a secondary limit without Retry-After, doubling each time" {
        Reset-YardstickGitHubSession -DeferThrottledCalls
        Mock Invoke-RestMethod { throw (New-GitHubHttpError -Status 429) } -ModuleName YardstickGitHub

        $waits = foreach ($attempt in 1..3) {
            $before = Get-Date
            { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } | Should -Throw
            [math]::Round(((Get-YardstickGitHubHold) - $before).TotalSeconds)
            # Let the next call through, as if the wait had passed.
            InModuleScope YardstickGitHub { $Script:HoldUntil = $null }
        }
        $waits | Should -Be @(60, 120, 240)
    }

    It "Keeps doubling when a call gets through between throttles" {
        Reset-YardstickGitHubSession -DeferThrottledCalls
        Mock Invoke-RestMethod {
            $Global:GitHubTestCalls++
            if ($Global:GitHubTestCalls -eq 2) { return [pscustomobject]@{ tag_name = 'v1.0.0' } }
            throw (New-GitHubHttpError -Status 429)
        } -ModuleName YardstickGitHub

        { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } | Should -Throw
        InModuleScope YardstickGitHub { $Script:HoldUntil = $null }
        Invoke-YardstickGitHubApi 'repos/owner/app/releases' | Out-Null
        $before = Get-Date
        { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } | Should -Throw
        [math]::Round(((Get-YardstickGitHubHold) - $before).TotalSeconds) | Should -Be 120
    }

    It "Starts the doubling again once GitHub has been quiet for 15 minutes" {
        Reset-YardstickGitHubSession -DeferThrottledCalls
        Mock Invoke-RestMethod { throw (New-GitHubHttpError -Status 429) } -ModuleName YardstickGitHub

        { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } | Should -Throw
        # As if that hold had ended 16 minutes ago.
        InModuleScope YardstickGitHub {
            $Script:HoldUntil = $null
            $Script:LastHoldUntil = (Get-Date).AddMinutes(-16)
        }
        $before = Get-Date
        { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } | Should -Throw
        [math]::Round(((Get-YardstickGitHubHold) - $before).TotalSeconds) | Should -Be 60
    }

    It "Never shortens a hold GitHub already asked for" {
        InModuleScope YardstickGitHub {
            Register-GitHubThrottle -HostName 'api.github.com' -Uri 'u' -Status 429 -Headers @{ 'retry-after' = '600' } | Out-Null
            Register-GitHubThrottle -HostName 'api.github.com' -Uri 'u' -Status 429 -Headers @{ 'retry-after' = '5' } | Out-Null
        }
        ((Get-YardstickGitHubHold) - (Get-Date)).TotalSeconds | Should -BeGreaterThan 590
    }

    It "Outside a run, waits out a short hold in place and tries again" {
        Mock Invoke-RestMethod {
            $Global:GitHubTestCalls++
            if ($Global:GitHubTestCalls -eq 1) { throw (New-GitHubHttpError -Status 403 -Headers @{ 'retry-after' = '5' }) }
            [pscustomobject]@{ tag_name = 'v1.0.0' }
        } -ModuleName YardstickGitHub

        (Invoke-YardstickGitHubApi 'repos/owner/app/releases/latest').tag_name | Should -Be 'v1.0.0'
        Should -Invoke Start-Sleep -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter { $Seconds -ge 4 -and $Seconds -le 5 }
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 2 -Exactly
    }

    It "Outside a run, gives up after three waits" {
        Mock Invoke-RestMethod { throw (New-GitHubHttpError -Status 429 -Headers @{ 'retry-after' = '5' }) } -ModuleName YardstickGitHub

        { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } | Should -Throw '*GitHub throttled*'
        Should -Invoke Start-Sleep -ModuleName YardstickGitHub -Times 3 -Exactly
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 4 -Exactly
    }

    It "Outside a run, does not sleep through a long wait" {
        Mock Invoke-RestMethod { throw (New-GitHubHttpError -Status 403 -Headers @{ 'retry-after' = '900' }) } -ModuleName YardstickGitHub

        { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } | Should -Throw '*GitHub throttled*'
        Should -Invoke Start-Sleep -ModuleName YardstickGitHub -Times 0 -Exactly
    }

    It "Does not retry a 403 that is not a rate limit" {
        Mock Invoke-RestMethod { throw (New-GitHubHttpError -Status 403) } -ModuleName YardstickGitHub
        { Invoke-YardstickGitHubApi 'repos/owner/app/releases' } | Should -Throw '*403*'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly
        Should -Invoke Start-Sleep -ModuleName YardstickGitHub -Times 0 -Exactly
    }
}

Describe "Invoke-YardstickGitHubRaw" {
    BeforeEach {
        Reset-YardstickGitHubSession -DeferThrottledCalls
        Mock Get-YardstickGitHubToken { 'github_pat_test' } -ModuleName YardstickGitHub
        Mock Invoke-RestMethod { 'PackageIdentifier: Amazon.Kindle' } -ModuleName YardstickGitHub
        Mock Start-Sleep {} -ModuleName YardstickGitHub
        Mock Write-Log {} -ModuleName YardstickGitHub
    }

    It "Reads the file with Yardstick's user agent and never sends the token" {
        $url = 'https://raw.githubusercontent.com/microsoft/winget-pkgs/master/manifests/a/Amazon/Kindle/1.0/Amazon.Kindle.installer.yaml'
        Invoke-YardstickGitHubRaw $url | Should -Be 'PackageIdentifier: Amazon.Kindle'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter {
            $Uri -eq 'https://raw.githubusercontent.com/microsoft/winget-pkgs/master/manifests/a/Amazon/Kindle/1.0/Amazon.Kindle.installer.yaml' -and
            $Headers['User-Agent'] -match '^Yardstick ' -and
            -not $Headers.ContainsKey('Authorization')
        }
        Should -Invoke Get-YardstickGitHubToken -ModuleName YardstickGitHub -Times 0 -Exactly
    }

    It "Refuses <Uri>" -ForEach @(
        @{ Uri = 'https://api.github.com/repos/owner/app' }
        @{ Uri = 'http://raw.githubusercontent.com/owner/app/main/file.yaml' }
        @{ Uri = 'manifests/a/Amazon/Kindle/1.0/Amazon.Kindle.installer.yaml' }
    ) {
        { Invoke-YardstickGitHubRaw $Uri } | Should -Throw '*only reads https://raw.githubusercontent.com*'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 0 -Exactly
    }

    It "Shares one hold with the API, so a throttled read stops API calls too" {
        Mock Invoke-RestMethod {
            if ($Uri -like 'https://raw.githubusercontent.com/*') {
                throw (New-GitHubHttpError -Status 429 -Headers @{ 'retry-after' = '60' } -Body '429: Too Many Requests')
            }
            [pscustomobject]@{ tag_name = 'v1.0.0' }
        } -ModuleName YardstickGitHub

        { Invoke-YardstickGitHubRaw 'https://raw.githubusercontent.com/owner/app/main/a.yaml' } | Should -Throw '*GitHub throttled*'
        { Invoke-YardstickGitHubRaw 'https://raw.githubusercontent.com/owner/app/main/b.yaml' } | Should -Throw '*Skipped*'
        { Invoke-YardstickGitHubApi 'repos/owner/app/releases/latest' } | Should -Throw '*Skipped*before calling GitHub again*'

        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly
        Get-YardstickGitHubHold | Should -Not -BeNullOrEmpty
        Should -Invoke Write-Log -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter {
            $Content -match 'GitHub said: 429: Too Many Requests' -and $Content -match 'No more calls to GitHub until'
        }
    }

    It "Lets a throttled API call stop reads too" {
        Mock Invoke-RestMethod {
            throw (New-GitHubHttpError -Status 429 -Body '{"message":"gitmon refuses to schedule us: too-many-processes"}')
        } -ModuleName YardstickGitHub

        { Invoke-YardstickGitHubApi 'repos/microsoft/winget-pkgs/contents/manifests/a/Amazon/Kindle' } | Should -Throw '*GitHub throttled*'
        { Invoke-YardstickGitHubRaw 'https://raw.githubusercontent.com/microsoft/winget-pkgs/master/x.yaml' } | Should -Throw '*Skipped*'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGitHub -Times 1 -Exactly
    }
}

Describe "Get-YardstickGitHubThrottle with a recipe's own error" {
    BeforeEach {
        Reset-YardstickGitHubSession -DeferThrottledCalls
        Mock Write-Log {} -ModuleName YardstickGitHub
    }

    It "Recognises a 429 from <HostName> and holds every GitHub call" -ForEach @(
        @{ Uri = 'https://raw.githubusercontent.com/microsoft/winget-pkgs/master/x.yaml'; HostName = 'raw.githubusercontent.com' }
        @{ Uri = 'https://github.com/owner/app/releases/download/v1/app.msi'; HostName = 'github.com' }
        @{ Uri = 'https://objects.githubusercontent.com/github-production-release-asset/1'; HostName = 'objects.githubusercontent.com' }
    ) {
        $record = New-GitHubHttpError -Status 429 -Headers @{ 'retry-after' = '90' } -Uri $Uri
        $throttle = Get-YardstickGitHubThrottle -ErrorRecord $record
        $throttle.Host | Should -Be $HostName
        Get-YardstickGitHubHold | Should -Not -BeNullOrEmpty
    }

    It "Recognises the error after a recipe wraps it" {
        $inner = (New-GitHubHttpError -Status 429 -Headers @{ 'retry-after' = '90' } -Uri 'https://raw.githubusercontent.com/o/r/main/x.yaml').Exception
        $record = [System.Management.Automation.ErrorRecord]::new([Exception]::new('Could not read the manifest', $inner), 'x', 'NotSpecified', $null)
        (Get-YardstickGitHubThrottle -ErrorRecord $record).Host | Should -Be 'raw.githubusercontent.com'
    }

    It "Ignores <Name>" -ForEach @(
        @{ Name = 'a 429 from another host'; Record = { New-GitHubHttpError -Status 429 -Headers @{ 'retry-after' = '90' } -Uri 'https://example.com/file' } }
        @{ Name = 'a 404 from GitHub'; Record = { New-GitHubHttpError -Status 404 -Uri 'https://raw.githubusercontent.com/owner/app/main/x.yaml' } }
        @{ Name = 'an error without a response'; Record = { [System.Management.Automation.ErrorRecord]::new([Exception]::new('boom'), 'x', 'NotSpecified', $null) } }
        @{ Name = 'no error at all'; Record = { $null } }
    ) {
        Get-YardstickGitHubThrottle -ErrorRecord (& $Record) | Should -BeNullOrEmpty
        Get-YardstickGitHubHold | Should -BeNullOrEmpty
    }
}

Describe "Get-YardstickGitHubHold" {
    BeforeEach {
        Reset-YardstickGitHubSession
    }

    It "Returns the hold while it is in force and forgets it once it has passed" {
        InModuleScope YardstickGitHub { $Script:HoldUntil = (Get-Date).AddMinutes(5) }
        ((Get-YardstickGitHubHold) - (Get-Date)).TotalMinutes | Should -BeGreaterThan 4.9

        InModuleScope YardstickGitHub { $Script:HoldUntil = (Get-Date).AddMinutes(-1) }
        Get-YardstickGitHubHold | Should -BeNullOrEmpty
        InModuleScope YardstickGitHub { $Script:HoldUntil } | Should -BeNullOrEmpty
    }
}

Describe "Token expiry warning" {
    BeforeEach {
        Reset-YardstickGitHubSession
        Mock Write-Log {} -ModuleName YardstickGitHub
    }

    It "Warns once when the token expires within 14 days" {
        InModuleScope YardstickGitHub {
            $soon = @{ 'github-authentication-token-expiration' = @((Get-Date).ToUniversalTime().AddDays(5).ToString('yyyy-MM-dd HH:mm:ss') + ' UTC') }
            Test-GitHubTokenExpiry -Headers $soon
            Test-GitHubTokenExpiry -Headers $soon
        }
        Should -Invoke Write-Log -ModuleName YardstickGitHub -Times 1 -Exactly -ParameterFilter { $Level -eq 'Warning' -and $Content -match 'expires on' }
    }

    It "Stays quiet for a token with time left or no expiry" {
        InModuleScope YardstickGitHub {
            Test-GitHubTokenExpiry -Headers @{ 'github-authentication-token-expiration' = @((Get-Date).ToUniversalTime().AddDays(90).ToString('yyyy-MM-dd HH:mm:ss') + ' UTC') }
            Test-GitHubTokenExpiry -Headers @{}
        }
        Should -Invoke Write-Log -ModuleName YardstickGitHub -Times 0 -Exactly
    }
}

Describe "Get-YardstickGitHubRateLimit" {
    BeforeEach {
        Reset-YardstickGitHubSession
    }

    It "Reports the core limit and whether a token was used" {
        Mock Get-YardstickGitHubToken { 'github_pat_test' } -ModuleName YardstickGitHub
        Mock Invoke-RestMethod {
            [pscustomobject]@{ resources = [pscustomobject]@{ core = [pscustomobject]@{ limit = 5000; remaining = 4990; reset = 1893456000 } } }
        } -ModuleName YardstickGitHub

        $limit = Get-YardstickGitHubRateLimit
        $limit.Authenticated | Should -BeTrue
        $limit.Limit | Should -Be 5000
        $limit.Remaining | Should -Be 4990
        $limit.Reset | Should -Be ([DateTimeOffset]::FromUnixTimeSeconds(1893456000).LocalDateTime)
    }

    It "Reports an anonymous session" {
        Mock Get-YardstickGitHubToken { $null } -ModuleName YardstickGitHub
        Mock Invoke-RestMethod {
            [pscustomobject]@{ resources = [pscustomobject]@{ core = [pscustomobject]@{ limit = 60; remaining = 59; reset = 1893456000 } } }
        } -ModuleName YardstickGitHub

        $limit = Get-YardstickGitHubRateLimit
        $limit.Authenticated | Should -BeFalse
        $limit.Limit | Should -Be 60
    }
}
