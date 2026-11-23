BeforeAll {
    # Earlier test files can leave extra copies of these modules loaded (YardstickSupport
    # imports them globally, YardstickSetup nests YardstickGraph), and Mock -ModuleName
    # refuses to pick between copies. Start from a single fresh copy.
    Get-Module -Name YardstickSupport, YardstickSetup, YardstickIntune, YardstickGraph -All | Remove-Module -Force -ErrorAction SilentlyContinue
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickGraph.psm1" -Force
}

AfterAll {
    Remove-Variable TenantID, ClientID, ClientSecret, Token, AuthenticationHeader -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Connect-YardstickGraph' {
    BeforeEach {
        Remove-Variable Token, AuthenticationHeader -Scope Global -ErrorAction SilentlyContinue
        InModuleScope YardstickGraph { $script:AccessToken = $null }
        Mock Invoke-RestMethod -ModuleName YardstickGraph {
            [pscustomobject]@{ access_token = 'test-token'; token_type = 'Bearer'; expires_in = 3600 }
        }
    }

    It 'uses client credentials and publishes the compatibility header' {
        Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret' | Out-Null
        $Global:AuthenticationHeader.Authorization | Should -Be 'Bearer test-token'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter {
            $Uri -eq 'https://login.microsoftonline.com/tenant/oauth2/v2.0/token' -and
            $Body.grant_type -eq 'client_credentials' -and
            $Body.scope -eq 'https://graph.microsoft.com/.default'
        }
    }

    It 'reuses a token that is safely inside its lifetime' {
        Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret' | Out-Null
        Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret' | Out-Null
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly
    }

    It 'reports missing credentials clearly in a fresh process' {
        Remove-Variable TenantID, ClientID, ClientSecret -Scope Global -ErrorAction SilentlyContinue
        { Connect-YardstickGraph } | Should -Throw '*TenantID, ClientID, and ClientSecret*'
    }

    It 'does not emit the token unless PassThru is requested' {
        $output = Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret'
        $output | Should -BeNullOrEmpty
        $cached = Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret'
        $cached | Should -BeNullOrEmpty
        (Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret' -PassThru).AccessToken | Should -Be 'test-token'
    }
}

Describe 'Invoke-YardstickGraphRequest' {
    BeforeEach {
        $Global:Token = [pscustomobject]@{ ExpiresOn = [datetimeoffset]::UtcNow.AddHours(1) }
        $Global:AuthenticationHeader = @{ Authorization = 'Bearer cached'; ExpiresOn = [datetimeoffset]::UtcNow.AddHours(1) }
    }

    It 'unwraps and follows OData collection pages' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph {
            if ($Uri -eq 'https://graph.microsoft.com/beta/next') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'two' }) }
            }
            [pscustomobject]@{ value = @([pscustomobject]@{ id = 'one' }); '@odata.nextLink' = 'https://graph.microsoft.com/beta/next' }
        }
        $result = @(Invoke-YardstickGraphRequest -Resource 'deviceAppManagement/mobileApps')
        $result.id | Should -Be @('one', 'two')
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 2 -Exactly
    }

    It 'returns an empty collection without manufacturing an item' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph { [pscustomobject]@{ value = @() } }
        @(Invoke-YardstickGraphRequest -Resource 'deviceAppManagement/mobileApps').Count | Should -Be 0
    }

    It 'serializes request bodies as JSON' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph { [pscustomobject]@{ id = 'created' } }
        $result = Invoke-YardstickGraphRequest -Method Post -ApiVersion v1.0 -Resource 'deviceAppManagement/mobileApps' -Body @{ name = 'App' }
        $result.id | Should -Be 'created'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Post' -and $ContentType -eq 'application/json' -and $Body -match '"name":"App"'
        }
    }
}

Describe 'Invoke-YardstickGraphRequest retries' {
    BeforeAll {
        function New-GraphHttpError {
            param([int]$Status, [int]$RetryAfter = 0)
            $message = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]$Status)
            $message.Headers.Add('request-id', 'req-123')
            if ($RetryAfter) { $message.Headers.Add('Retry-After', "$RetryAfter") }
            [Microsoft.PowerShell.Commands.HttpResponseException]::new("Response status code does not indicate success: $Status.", $message)
        }
    }

    BeforeEach {
        $Global:Token = [pscustomobject]@{ ExpiresOn = [datetimeoffset]::UtcNow.AddHours(1) }
        $Global:AuthenticationHeader = @{ Authorization = 'Bearer cached'; ExpiresOn = [datetimeoffset]::UtcNow.AddHours(1) }
        Mock Start-Sleep -ModuleName YardstickGraph {}
        Mock Write-YardstickGraphLog -ModuleName YardstickGraph {}
        Mock Connect-YardstickGraph -ModuleName YardstickGraph {}
        Reset-YardstickGraphStatistics
        $script:calls = 0
    }

    It 'retries a 504 and logs the status, attempt and both request ids' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph {
            $script:calls++
            if ($script:calls -lt 3) { throw (New-GraphHttpError -Status 504) }
            [pscustomobject]@{ value = @([pscustomobject]@{ id = 'ok' }) }
        }
        @(Invoke-YardstickGraphRequest -Resource 'deviceAppManagement/mobileApps').id | Should -Be @('ok')
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 3 -Exactly
        Should -Invoke Write-YardstickGraphLog -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter {
            $Level -eq 'Detail' -and $Message -match 'returned 504' -and $Message -match 'attempt 1 of 6' -and
            $Message -match 'request-id req-123' -and $Message -match 'client-request-id [0-9a-f-]{36}'
        }
        Should -Invoke Write-YardstickGraphLog -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter {
            $Level -eq 'Warning' -and $Message -match 'attempt 2 of 6'
        }
        (Get-YardstickGraphStatistics).Retried | Should -Be 2
        (Get-YardstickGraphStatistics).Requests | Should -Be 3
    }

    It 'logs a warning and counts it when it gives up' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph { throw (New-GraphHttpError -Status 503) }
        { Invoke-YardstickGraphRequest -Resource 'x' -MaxAttempts 2 } | Should -Throw
        Should -Invoke Write-YardstickGraphLog -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter { $Level -eq 'Warning' -and $Message -match 'Giving up' }
        (Get-YardstickGraphStatistics).GaveUp | Should -Be 1
    }

    It 'does not retry a 4xx' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph { throw (New-GraphHttpError -Status 400) }
        { Invoke-YardstickGraphRequest -Resource 'x' } | Should -Throw
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly
    }

    It 'does not resend a request after a 5xx with -NoRetryOnServerError' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph { throw (New-GraphHttpError -Status 504) }
        { Invoke-YardstickGraphRequest -Method Post -Resource 'deviceAppManagement/mobileApps' -Body @{} -NoRetryOnServerError } | Should -Throw
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly
    }

    It 'still retries a 429 with -NoRetryOnServerError' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph {
            $script:calls++
            if ($script:calls -eq 1) { throw (New-GraphHttpError -Status 429 -RetryAfter 3) }
            [pscustomobject]@{ id = 'created' }
        }
        (Invoke-YardstickGraphRequest -Method Post -Resource 'x' -Body @{} -NoRetryOnServerError).id | Should -Be 'created'
        Should -Invoke Start-Sleep -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter { $Seconds -eq 3 }
    }

    It 'caps a long Retry-After at 120 seconds' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph {
            $script:calls++
            if ($script:calls -eq 1) { throw (New-GraphHttpError -Status 429 -RetryAfter 600) }
            [pscustomobject]@{ id = 'ok' }
        }
        $null = Invoke-YardstickGraphRequest -Resource 'x'
        Should -Invoke Start-Sleep -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter { $Seconds -eq 120 }
    }

    It 'treats a 404 on a retried DELETE as done' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph {
            $script:calls++
            if ($script:calls -eq 1) { throw (New-GraphHttpError -Status 504) }
            throw (New-GraphHttpError -Status 404)
        }
        { Invoke-YardstickGraphRequest -Method Delete -Resource 'deviceAppManagement/mobileApps/1' } | Should -Not -Throw
    }

    It 'still throws a 404 on a first DELETE attempt' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph { throw (New-GraphHttpError -Status 404) }
        { Invoke-YardstickGraphRequest -Method Delete -Resource 'deviceAppManagement/mobileApps/1' } | Should -Throw
    }

    It 'refreshes the token once on a 401' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph {
            $script:calls++
            if ($script:calls -eq 1) { throw (New-GraphHttpError -Status 401) }
            [pscustomobject]@{ id = 'ok' }
        }
        (Invoke-YardstickGraphRequest -Resource 'x').id | Should -Be 'ok'
        Should -Invoke Connect-YardstickGraph -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter { $Force }
    }

    It 'refreshes a token about to expire before a retry' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph {
            $script:calls++
            if ($script:calls -eq 1) {
                $Global:Token = [pscustomobject]@{ ExpiresOn = [datetimeoffset]::UtcNow.AddMinutes(1) }
                throw (New-GraphHttpError -Status 504)
            }
            [pscustomobject]@{ id = 'ok' }
        }
        InModuleScope YardstickGraph { $script:AccessToken = $null }
        $null = Invoke-YardstickGraphRequest -Resource 'x'
        Should -Invoke Connect-YardstickGraph -ModuleName YardstickGraph -Times 1 -Exactly
    }

    It 'repeats only the page that failed' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph {
            if ($Uri -eq 'https://graph.microsoft.com/beta/next') {
                $script:calls++
                if ($script:calls -eq 1) { throw (New-GraphHttpError -Status 504) }
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'two' }) }
            }
            [pscustomobject]@{ value = @([pscustomobject]@{ id = 'one' }); '@odata.nextLink' = 'https://graph.microsoft.com/beta/next' }
        }
        @(Invoke-YardstickGraphRequest -Resource 'x').id | Should -Be @('one', 'two')
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter { $Uri -ne 'https://graph.microsoft.com/beta/next' }
    }

    It 'sends a fresh client-request-id and a 60-second timeout on every attempt' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph { [pscustomobject]@{ id = 'ok' } }
        $null = Invoke-YardstickGraphRequest -Resource 'x'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter {
            $TimeoutSec -eq 60 -and $Headers['client-request-id'] -match '^[0-9a-f-]{36}$'
        }
    }
}

Describe 'Disconnect-YardstickGraph' {
    It 'clears the token and the credentials from the session' {
        $Global:TenantID = 'tenant'
        $Global:ClientID = 'client'
        $Global:ClientSecret = 'secret'
        $Global:Token = [pscustomobject]@{ ExpiresOn = [datetimeoffset]::UtcNow.AddHours(1) }
        $Global:AuthenticationHeader = @{ Authorization = 'Bearer cached' }

        Disconnect-YardstickGraph

        Get-Variable -Name TenantID, ClientID, ClientSecret, Token, AuthenticationHeader, YardstickCredential -Scope Global -ErrorAction SilentlyContinue |
            Should -BeNullOrEmpty
    }

    It 'is safe to call when nothing is connected' {
        { Disconnect-YardstickGraph } | Should -Not -Throw
    }
}
