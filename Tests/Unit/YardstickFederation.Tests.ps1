BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickFederation.psm1" -Force -Global
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickGraph.psm1" -Force
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickCredential.psm1" -Force

    $script:RepoRoot = (Resolve-Path "$PSScriptRoot\..\..").Path
    $script:FederationVariables = @(
        'YARDSTICK_SERVICE_CONNECTION_ID',
        'YARDSTICK_TENANT_ID',
        'YARDSTICK_CLIENT_ID',
        'YARDSTICK_FEDERATED_TOKEN_FILE',
        'SYSTEM_OIDCREQUESTURI',
        'SYSTEM_ACCESSTOKEN',
        'ACTIONS_ID_TOKEN_REQUEST_URL',
        'ACTIONS_ID_TOKEN_REQUEST_TOKEN',
        'GITHUB_ACTIONS'
    )
    $script:SavedEnvironment = @{}
    foreach ($name in $script:FederationVariables) {
        $script:SavedEnvironment[$name] = [Environment]::GetEnvironmentVariable($name)
    }

    $script:ServiceConnectionId = '11111111-2222-3333-4444-555555555555'
    $script:ClientId = '66666666-7777-8888-9999-000000000000'
    $script:JobAccessToken = 'job-access-token-do-not-log'

    function script:Clear-FederationEnvironment {
        foreach ($name in $script:FederationVariables) {
            [Environment]::SetEnvironmentVariable($name, $null)
        }
    }

    function script:Set-FederationEnvironment {
        param([hashtable]$Override = @{})
        $values = @{
            YARDSTICK_SERVICE_CONNECTION_ID = $script:ServiceConnectionId
            YARDSTICK_TENANT_ID             = 'tenant-fed'
            YARDSTICK_CLIENT_ID             = $script:ClientId
            SYSTEM_OIDCREQUESTURI           = 'https://dev.azure.com/org/project/_apis/distributedtask/hubs/build/plans/plan/jobs/job/oidctoken'
            SYSTEM_ACCESSTOKEN              = $script:JobAccessToken
        }
        foreach ($key in $Override.Keys) { $values[$key] = $Override[$key] }
        foreach ($key in $values.Keys) { [Environment]::SetEnvironmentVariable($key, $values[$key]) }
    }

    function script:Set-GitHubActionsEnvironment {
        param([hashtable]$Override = @{})
        $values = @{
            YARDSTICK_TENANT_ID            = 'tenant-fed'
            YARDSTICK_CLIENT_ID            = $script:ClientId
            GITHUB_ACTIONS                 = 'true'
            ACTIONS_ID_TOKEN_REQUEST_URL   = 'https://pipelines.actions.githubusercontent.com/abc/idtoken/def?api-version=2.0'
            ACTIONS_ID_TOKEN_REQUEST_TOKEN = $script:JobAccessToken
        }
        foreach ($key in $Override.Keys) { $values[$key] = $Override[$key] }
        foreach ($key in $values.Keys) { [Environment]::SetEnvironmentVariable($key, $values[$key]) }
    }

    # A JWT-shaped assertion whose payload carries the claims Entra matches on.
    function script:New-TestJwt {
        param([string]$Subject = 'sc://org/project/connection')
        $encode = {
            param($Text)
            [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        }
        $header = & $encode '{"alg":"RS256","typ":"JWT"}'
        $payload = & $encode ('{{"iss":"https://vstoken.dev.azure.com/org-guid","sub":"{0}","aud":"api://AzureADTokenExchange"}}' -f $Subject)
        return "$header.$payload.signature"
    }

    function script:New-HttpErrorRecord {
        param([int]$StatusCode, [string]$Body)
        $response = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]$StatusCode)
        $exception = [Microsoft.PowerShell.Commands.HttpResponseException]::new(
            "Response status code does not indicate success: $StatusCode.", $response)
        $record = [System.Management.Automation.ErrorRecord]::new(
            $exception, 'WebCmdletWebResponseException', [System.Management.Automation.ErrorCategory]::InvalidOperation, $null)
        if ($Body) { $record.ErrorDetails = [System.Management.Automation.ErrorDetails]::new($Body) }
        return $record
    }
}

AfterAll {
    foreach ($name in $script:FederationVariables) {
        [Environment]::SetEnvironmentVariable($name, $script:SavedEnvironment[$name])
    }
    Remove-Variable TenantID, ClientID, ClientSecret, Token, AuthenticationHeader, YardstickCredential -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Get-YardstickFederatedIdentity' {
    BeforeEach { Clear-FederationEnvironment }

    It 'returns nothing when no federated variable is set' {
        Get-YardstickFederatedIdentity | Should -BeNullOrEmpty
    }

    It 'rejects a variable that is set but blank instead of falling back to a secret' {
        [Environment]::SetEnvironmentVariable('YARDSTICK_TENANT_ID', '   ')
        { Get-YardstickFederatedIdentity } | Should -Throw '*YARDSTICK_TENANT_ID is set but blank*does not fall back to a client secret*'
    }

    It 'returns the identity when every input is present' {
        Set-FederationEnvironment
        $identity = Get-YardstickFederatedIdentity
        $identity.TenantID | Should -Be 'tenant-fed'
        $identity.ClientID | Should -Be $script:ClientId
        $identity.Provider | Should -Be 'AzureDevOps'
        $identity.ServiceConnectionID | Should -Be $script:ServiceConnectionId
        $identity.RequestUri | Should -Match '/oidctoken$'
    }

    It 'never carries the job access token' {
        Set-FederationEnvironment
        $identity = Get-YardstickFederatedIdentity
        ($identity | ConvertTo-Json -Compress) | Should -Not -Match ([regex]::Escape($script:JobAccessToken))
    }

    It 'is switched on by <Name> alone and then names the missing inputs' -TestCases @(
        @{ Name = 'YARDSTICK_SERVICE_CONNECTION_ID'; Value = '11111111-2222-3333-4444-555555555555'; Expected = 'SYSTEM_ACCESSTOKEN is not available' }
        @{ Name = 'YARDSTICK_FEDERATED_TOKEN_FILE'; Value = 'Z:\no\such\token'; Expected = 'does not exist' }
        @{ Name = 'YARDSTICK_TENANT_ID'; Value = 'tenant'; Expected = 'no ID token source is set' }
        @{ Name = 'YARDSTICK_CLIENT_ID'; Value = '66666666-7777-8888-9999-000000000000'; Expected = 'no ID token source is set' }
    ) {
        [Environment]::SetEnvironmentVariable($Name, $Value)
        { Get-YardstickFederatedIdentity } | Should -Throw '*does not fall back to a client secret*'
        try { Get-YardstickFederatedIdentity } catch { $message = $_.Exception.Message }
        $message | Should -Match $Expected
        if ($Name -ne 'YARDSTICK_TENANT_ID') { $message | Should -Match 'YARDSTICK_TENANT_ID is not set' }
    }

    It 'is not switched on by a GitHub Actions job alone' {
        Set-GitHubActionsEnvironment -Override @{ YARDSTICK_TENANT_ID = $null; YARDSTICK_CLIENT_ID = $null }
        Get-YardstickFederatedIdentity | Should -BeNullOrEmpty
    }

    It 'uses GitHub Actions when it offers an ID token and no other source is set' {
        Set-GitHubActionsEnvironment
        $identity = Get-YardstickFederatedIdentity
        $identity.Provider | Should -Be 'GitHubActions'
        $identity.RequestUri | Should -BeLike 'https://pipelines.actions.githubusercontent.com/*'
        ($identity | ConvertTo-Json -Compress) | Should -Not -Match ([regex]::Escape($script:JobAccessToken))
    }

    It 'asks for id-token: write in a GitHub Actions job that offers no ID token' {
        Set-GitHubActionsEnvironment -Override @{ ACTIONS_ID_TOKEN_REQUEST_URL = $null; ACTIONS_ID_TOKEN_REQUEST_TOKEN = $null }
        { Get-YardstickFederatedIdentity } | Should -Throw "*permissions: id-token: write*"
    }

    It 'uses a token file' {
        $file = Join-Path $TestDrive 'token'
        Set-Content -LiteralPath $file -Value 'id-token'
        [Environment]::SetEnvironmentVariable('YARDSTICK_TENANT_ID', 'tenant-fed')
        [Environment]::SetEnvironmentVariable('YARDSTICK_CLIENT_ID', $script:ClientId)
        [Environment]::SetEnvironmentVariable('YARDSTICK_FEDERATED_TOKEN_FILE', $file)
        $identity = Get-YardstickFederatedIdentity
        $identity.Provider | Should -Be 'File'
        $identity.TokenFile | Should -Be $file
    }

    It 'prefers an explicit source to GitHub Actions' {
        $file = Join-Path $TestDrive 'token'
        Set-Content -LiteralPath $file -Value 'id-token'
        Set-GitHubActionsEnvironment -Override @{ YARDSTICK_FEDERATED_TOKEN_FILE = $file }
        (Get-YardstickFederatedIdentity).Provider | Should -Be 'File'
    }

    It 'refuses both a token file and a service connection' {
        $file = Join-Path $TestDrive 'token'
        Set-Content -LiteralPath $file -Value 'id-token'
        Set-FederationEnvironment -Override @{ YARDSTICK_FEDERATED_TOKEN_FILE = $file }
        { Get-YardstickFederatedIdentity } | Should -Throw '*are both set*'
    }

    It 'rejects a SYSTEM_ACCESSTOKEN left as an unresolved macro' {
        Set-FederationEnvironment -Override @{ SYSTEM_ACCESSTOKEN = '$(System.AccessToken)' }
        { Get-YardstickFederatedIdentity } | Should -Throw '*SYSTEM_ACCESSTOKEN: $(System.AccessToken)*'
    }

    It 'rejects a YARDSTICK variable left as an unresolved macro' {
        Set-FederationEnvironment -Override @{ YARDSTICK_SERVICE_CONNECTION_ID = '$(serviceConnectionId)' }
        { Get-YardstickFederatedIdentity } | Should -Throw "*YARDSTICK_SERVICE_CONNECTION_ID is the unresolved pipeline variable*"
    }

    It 'rejects a client ID that is not a GUID' {
        Set-FederationEnvironment -Override @{ YARDSTICK_CLIENT_ID = 'not-a-guid' }
        { Get-YardstickFederatedIdentity } | Should -Throw "*YARDSTICK_CLIENT_ID 'not-a-guid' is not a GUID*"
    }

    It 'rejects a relative OIDC request URI' {
        Set-FederationEnvironment -Override @{ SYSTEM_OIDCREQUESTURI = '/oidctoken' }
        { Get-YardstickFederatedIdentity } | Should -Throw '*SYSTEM_OIDCREQUESTURI is not an absolute https URI*'
    }

    It 'rejects a plain-HTTP <Name> on another host, so the job token is never sent unencrypted' -TestCases @(
        @{ Name = 'SYSTEM_OIDCREQUESTURI'; Set = { Set-FederationEnvironment -Override @{ SYSTEM_OIDCREQUESTURI = 'http://dev.azure.com/org/oidctoken' } } }
        @{ Name = 'ACTIONS_ID_TOKEN_REQUEST_URL'; Set = { Set-GitHubActionsEnvironment -Override @{ ACTIONS_ID_TOKEN_REQUEST_URL = 'http://pipelines.actions.githubusercontent.com/idtoken' } } }
    ) {
        & $Set
        { Get-YardstickFederatedIdentity } | Should -Throw "*$Name is not an absolute https URI*"
    }

    It 'allows plain HTTP to this machine' {
        Set-GitHubActionsEnvironment -Override @{ ACTIONS_ID_TOKEN_REQUEST_URL = 'http://localhost:8080/idtoken' }
        (Get-YardstickFederatedIdentity).Provider | Should -Be 'GitHubActions'
    }

    It 'rejects a token file given as a relative path' {
        Set-Content -LiteralPath (Join-Path $TestDrive 'token') -Value 'id-token'
        Push-Location $TestDrive
        try {
            [Environment]::SetEnvironmentVariable('YARDSTICK_TENANT_ID', 'tenant-fed')
            [Environment]::SetEnvironmentVariable('YARDSTICK_CLIENT_ID', $script:ClientId)
            [Environment]::SetEnvironmentVariable('YARDSTICK_FEDERATED_TOKEN_FILE', 'token')
            { Get-YardstickFederatedIdentity } | Should -Throw "*YARDSTICK_FEDERATED_TOKEN_FILE 'token' is not a full path*"
        } finally {
            Pop-Location
        }
    }
}

Describe 'Retry delay' {
    It 'waits as long as Retry-After asks, up to a minute' -TestCases @(
        @{ Seconds = 30; Expected = 30 }
        @{ Seconds = 600; Expected = 60 }
    ) {
        $record = New-HttpErrorRecord -StatusCode 429
        $record.Exception.Response.Headers.RetryAfter = [System.Net.Http.Headers.RetryConditionHeaderValue]::new([timespan]::FromSeconds($Seconds))
        InModuleScope YardstickFederation -Parameters @{ Record = $record } { Get-FederationRetryDelay -ErrorRecord $Record -Attempt 1 } | Should -Be $Expected
    }

    It 'reads a Retry-After date' {
        $record = New-HttpErrorRecord -StatusCode 503
        $record.Exception.Response.Headers.RetryAfter = [System.Net.Http.Headers.RetryConditionHeaderValue]::new([datetimeoffset]::UtcNow.AddSeconds(20))
        $delay = InModuleScope YardstickFederation -Parameters @{ Record = $record } { Get-FederationRetryDelay -ErrorRecord $Record -Attempt 1 }
        $delay | Should -BeGreaterOrEqual 18
        $delay | Should -BeLessOrEqual 21
    }

    It 'backs off exponentially without Retry-After' {
        $record = New-HttpErrorRecord -StatusCode 503
        InModuleScope YardstickFederation -Parameters @{ Record = $record } { Get-FederationRetryDelay -ErrorRecord $Record -Attempt 2 } | Should -Be 4
    }

    It 'waits the Retry-After time between attempts' {
        Clear-FederationEnvironment
        Set-FederationEnvironment
        $identity = Get-YardstickFederatedIdentity
        $script:Calls = 0
        Mock Start-Sleep -ModuleName YardstickFederation {}
        Mock Invoke-RestMethod -ModuleName YardstickFederation {
            $script:Calls++
            if ($script:Calls -eq 1) {
                $record = New-HttpErrorRecord -StatusCode 429
                $record.Exception.Response.Headers.RetryAfter = [System.Net.Http.Headers.RetryConditionHeaderValue]::new([timespan]::FromSeconds(30))
                throw $record
            }
            [pscustomobject]@{ oidcToken = 'id-token' }
        }
        Get-YardstickFederatedAssertion -Identity $identity | Should -Be 'id-token'
        Should -Invoke Start-Sleep -ModuleName YardstickFederation -Times 1 -Exactly -ParameterFilter { $Seconds -eq 30 }
    }
}

Describe 'Get-YardstickFederatedAssertion' {
    BeforeEach {
        Clear-FederationEnvironment
        Set-FederationEnvironment
        $script:Identity = Get-YardstickFederatedIdentity
        Mock Start-Sleep -ModuleName YardstickFederation {}
    }

    It 'requests an ID token for the service connection with the job access token' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation { [pscustomobject]@{ oidcToken = 'id-token' } }
        Get-YardstickFederatedAssertion -Identity $script:Identity | Should -Be 'id-token'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Post' -and
            $Uri -eq "https://dev.azure.com/org/project/_apis/distributedtask/hubs/build/plans/plan/jobs/job/oidctoken?api-version=7.1&serviceConnectionId=11111111-2222-3333-4444-555555555555" -and
            $Headers.Authorization -eq 'Bearer job-access-token-do-not-log' -and
            $ContentType -eq 'application/json' -and
            $Body -eq '{}'
        }
    }

    It 'appends to a request URI that already has a query string' {
        $script:Identity.RequestUri = 'https://dev.azure.com/org/oidctoken?existing=1'
        Mock Invoke-RestMethod -ModuleName YardstickFederation { [pscustomobject]@{ oidcToken = 'id-token' } }
        Get-YardstickFederatedAssertion -Identity $script:Identity | Out-Null
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 1 -Exactly -ParameterFilter {
            $Uri -eq 'https://dev.azure.com/org/oidctoken?existing=1&api-version=7.1&serviceConnectionId=11111111-2222-3333-4444-555555555555'
        }
    }

    It 'reads the job access token at request time' {
        [Environment]::SetEnvironmentVariable('SYSTEM_ACCESSTOKEN', 'rotated-job-token')
        Mock Invoke-RestMethod -ModuleName YardstickFederation { [pscustomobject]@{ oidcToken = 'id-token' } }
        Get-YardstickFederatedAssertion -Identity $script:Identity | Out-Null
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 1 -Exactly -ParameterFilter {
            $Headers.Authorization -eq 'Bearer rotated-job-token'
        }
    }

    It 'retries a transient Azure DevOps failure' {
        $script:Calls = 0
        Mock Invoke-RestMethod -ModuleName YardstickFederation {
            $script:Calls++
            if ($script:Calls -eq 1) { throw (New-HttpErrorRecord -StatusCode 503) }
            [pscustomobject]@{ oidcToken = 'id-token' }
        }
        Get-YardstickFederatedAssertion -Identity $script:Identity | Should -Be 'id-token'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 2 -Exactly
    }

    It 'gives up after three transient failures' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation { throw (New-HttpErrorRecord -StatusCode 500) }
        { Get-YardstickFederatedAssertion -Identity $script:Identity } | Should -Throw '*requesting an ID token from Azure DevOps failed (HTTP 500)*'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 3 -Exactly
    }

    It 'does not retry a rejected job token, and says what to check without echoing it' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation { throw (New-HttpErrorRecord -StatusCode 401 -Body 'Unauthorized') }
        try { Get-YardstickFederatedAssertion -Identity $script:Identity } catch { $message = $_.Exception.Message }
        $message | Should -Match 'HTTP 401'
        $message | Should -Match 'SYSTEM_ACCESSTOKEN is mapped'
        $message | Should -Not -Match ([regex]::Escape($script:JobAccessToken))
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 1 -Exactly
    }

    It 'points at the service connection when Azure DevOps refuses it' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation {
            throw (New-HttpErrorRecord -StatusCode 400 -Body '{"message":"Pipeline does not have permissions to use the referenced service connection."}')
        }
        try { Get-YardstickFederatedAssertion -Identity $script:Identity } catch { $message = $_.Exception.Message }
        $message | Should -Match 'service connection 11111111-2222-3333-4444-555555555555 exists'
        $message | Should -Match 'Pipeline does not have permissions'
    }

    It 'reads the message from an Azure DevOps error body that starts with a blank line' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation {
            throw (New-HttpErrorRecord -StatusCode 403 -Body "`r`n{`r`n  `"`$id`": `"1`",`r`n  `"innerException`": null,`r`n  `"message`": `"Only build identities are allowed to request OpenIdConnect tokens.`",`r`n  `"typeKey`": `"AccessDeniedException`"`r`n}")
        }
        try { Get-YardstickFederatedAssertion -Identity $script:Identity } catch { $message = $_.Exception.Message }
        $message | Should -Match 'Only build identities are allowed to request OpenIdConnect tokens\.$'
    }

    It 'keeps the whole body, on one line, when it has no message field' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation {
            throw (New-HttpErrorRecord -StatusCode 400 -Body "{`r`n  `"title`": `"One or more validation errors occurred.`",`r`n  `"status`": 400,`r`n  `"errors`": { `"serviceConnectionId`": [ `"The value is invalid.`" ] }`r`n}")
        }
        try { Get-YardstickFederatedAssertion -Identity $script:Identity } catch { $message = $_.Exception.Message }
        $message | Should -Match 'One or more validation errors occurred'
        $message | Should -Match 'The value is invalid'
        $message | Should -Not -Match "`n"
    }

    It 'fails when Azure DevOps returns no token' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation { [pscustomobject]@{ oidcToken = '' } }
        { Get-YardstickFederatedAssertion -Identity $script:Identity } | Should -Throw '*returned no ID token*'
    }

    It 'fails when the job access token has disappeared' {
        [Environment]::SetEnvironmentVariable('SYSTEM_ACCESSTOKEN', $null)
        Mock Invoke-RestMethod -ModuleName YardstickFederation {}
        { Get-YardstickFederatedAssertion -Identity $script:Identity } | Should -Throw '*SYSTEM_ACCESSTOKEN is no longer available*'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 0 -Exactly
    }
}

Describe 'Get-YardstickFederatedAssertion from GitHub Actions' {
    BeforeEach {
        Clear-FederationEnvironment
        Set-GitHubActionsEnvironment
        $script:Identity = Get-YardstickFederatedIdentity
        Mock Start-Sleep -ModuleName YardstickFederation {}
    }

    It 'requests an ID token for the Entra audience with the request token' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation { [pscustomobject]@{ value = 'id-token' } }
        Get-YardstickFederatedAssertion -Identity $script:Identity | Should -Be 'id-token'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Get' -and
            $Uri -eq 'https://pipelines.actions.githubusercontent.com/abc/idtoken/def?api-version=2.0&audience=api%3A%2F%2FAzureADTokenExchange' -and
            $Headers.Authorization -eq 'Bearer job-access-token-do-not-log'
        }
    }

    It 'points at id-token: write when GitHub refuses, without echoing the request token' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation { throw (New-HttpErrorRecord -StatusCode 403 -Body '{"message":"Forbidden"}') }
        try { Get-YardstickFederatedAssertion -Identity $script:Identity } catch { $message = $_.Exception.Message }
        $message | Should -Match 'requesting an ID token from GitHub Actions failed \(HTTP 403\)'
        $message | Should -Match 'id-token: write'
        $message | Should -Not -Match ([regex]::Escape($script:JobAccessToken))
    }

    It 'fails when GitHub returns no token' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation { [pscustomobject]@{ value = $null } }
        { Get-YardstickFederatedAssertion -Identity $script:Identity } | Should -Throw '*GitHub Actions returned no ID token*'
    }
}

Describe 'Get-YardstickFederatedAssertion from a token file' {
    BeforeEach {
        Clear-FederationEnvironment
        $script:TokenFile = Join-Path $TestDrive 'token'
        Set-Content -LiteralPath $script:TokenFile -Value "first-token`n"
        [Environment]::SetEnvironmentVariable('YARDSTICK_TENANT_ID', 'tenant-fed')
        [Environment]::SetEnvironmentVariable('YARDSTICK_CLIENT_ID', $script:ClientId)
        [Environment]::SetEnvironmentVariable('YARDSTICK_FEDERATED_TOKEN_FILE', $script:TokenFile)
        $script:Identity = Get-YardstickFederatedIdentity
    }

    It 'reads the file, trimmed, every time' {
        Get-YardstickFederatedAssertion -Identity $script:Identity | Should -Be 'first-token'
        Set-Content -LiteralPath $script:TokenFile -Value 'rotated-token'
        Get-YardstickFederatedAssertion -Identity $script:Identity | Should -Be 'rotated-token'
    }

    It 'fails on an empty file' {
        Set-Content -LiteralPath $script:TokenFile -Value ''
        { Get-YardstickFederatedAssertion -Identity $script:Identity } | Should -Throw '*token file returned no ID token*'
    }

    It 'fails when the file has gone' {
        Remove-Item -LiteralPath $script:TokenFile
        { Get-YardstickFederatedAssertion -Identity $script:Identity } | Should -Throw "*reading the ID token from*"
    }
}

Describe 'Get-YardstickJwtClaimSummary' {
    It 'describes issuer, subject and audience' {
        $summary = Get-YardstickJwtClaimSummary -Jwt (New-TestJwt -Subject 'sc://contoso/Packaging/intune')
        $summary | Should -Be "iss 'https://vstoken.dev.azure.com/org-guid', sub 'sc://contoso/Packaging/intune', aud 'api://AzureADTokenExchange'"
    }

    It 'returns nothing for a malformed token' {
        Get-YardstickJwtClaimSummary -Jwt 'not-a-jwt' | Should -BeNullOrEmpty
        Get-YardstickJwtClaimSummary -Jwt 'a.!!!.c' | Should -BeNullOrEmpty
    }
}

Describe 'Invoke-YardstickFederatedTokenRequest' {
    BeforeEach {
        Clear-FederationEnvironment
        Set-FederationEnvironment
        $script:Identity = Get-YardstickFederatedIdentity
        $script:Jwt = New-TestJwt
        Mock Start-Sleep -ModuleName YardstickFederation {}
    }

    It 'exchanges the ID token as a client assertion and never sends a secret' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation -ParameterFilter { $Uri -like '*oidctoken*' } { [pscustomobject]@{ oidcToken = $script:Jwt } }
        Mock Invoke-RestMethod -ModuleName YardstickFederation -ParameterFilter { $Uri -like 'https://login.microsoftonline.com/*' } {
            [pscustomobject]@{ access_token = 'graph-token'; token_type = 'Bearer'; expires_in = 3599 }
        }
        $response = Invoke-YardstickFederatedTokenRequest -TokenUri 'https://login.microsoftonline.com/tenant-fed/oauth2/v2.0/token' -ClientID $script:ClientId -Identity $script:Identity
        $response.access_token | Should -Be 'graph-token'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 1 -Exactly -ParameterFilter {
            $Uri -like 'https://login.microsoftonline.com/*' -and
            $Method -eq 'Post' -and
            $ContentType -eq 'application/x-www-form-urlencoded' -and
            $Body.client_id -eq '66666666-7777-8888-9999-000000000000' -and
            $Body.client_assertion -eq $script:Jwt -and
            $Body.client_assertion_type -eq 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer' -and
            $Body.grant_type -eq 'client_credentials' -and
            $Body.scope -eq 'https://graph.microsoft.com/.default' -and
            -not $Body.ContainsKey('client_secret')
        }
    }

    It 'reports the AADSTS code and the presented claims, but not the assertion' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation -ParameterFilter { $Uri -like '*oidctoken*' } { [pscustomobject]@{ oidcToken = $script:Jwt } }
        Mock Invoke-RestMethod -ModuleName YardstickFederation -ParameterFilter { $Uri -like 'https://login.microsoftonline.com/*' } {
            $body = @{
                error             = 'invalid_client'
                error_description = "AADSTS700213: No matching federated identity record found for presented assertion subject 'sc://org/project/connection'.`r`nTrace ID: abc`r`nCorrelation ID: def"
            } | ConvertTo-Json
            throw (New-HttpErrorRecord -StatusCode 400 -Body $body)
        }
        try {
            Invoke-YardstickFederatedTokenRequest -TokenUri 'https://login.microsoftonline.com/tenant-fed/oauth2/v2.0/token' -ClientID $script:ClientId -Identity $script:Identity
        } catch { $message = $_.Exception.Message }
        $message | Should -Match 'exchanging the Azure DevOps ID token with Entra ID failed \(HTTP 400, AADSTS700213\)'
        $message | Should -Match "sub 'sc://org/project/connection'"
        $message | Should -Not -Match 'Trace ID'
        $message | Should -Not -Match ([regex]::Escape($script:Jwt))
    }

    It 'requests a new ID token for every exchange' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation -ParameterFilter { $Uri -like '*oidctoken*' } { [pscustomobject]@{ oidcToken = $script:Jwt } }
        Mock Invoke-RestMethod -ModuleName YardstickFederation -ParameterFilter { $Uri -like 'https://login.microsoftonline.com/*' } {
            [pscustomobject]@{ access_token = 'graph-token'; token_type = 'Bearer'; expires_in = 3599 }
        }
        1..2 | ForEach-Object {
            Invoke-YardstickFederatedTokenRequest -TokenUri 'https://login.microsoftonline.com/t/oauth2/v2.0/token' -ClientID $script:ClientId -Identity $script:Identity | Out-Null
        }
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 2 -Exactly -ParameterFilter { $Uri -like '*oidctoken*' }
    }
}

Describe 'Connect-YardstickGraph with workload identity federation' {
    BeforeEach {
        Clear-FederationEnvironment
        Set-FederationEnvironment
        Remove-Variable TenantID, ClientID, ClientSecret, Token, AuthenticationHeader -Scope Global -ErrorAction SilentlyContinue
        InModuleScope YardstickGraph { $script:AccessToken = $null }
        $script:Jwt = New-TestJwt
        $script:GraphTokenCount = 0
        Mock Start-Sleep -ModuleName YardstickFederation {}
        Mock Invoke-RestMethod -ModuleName YardstickFederation -ParameterFilter { $Uri -like '*oidctoken*' } { [pscustomobject]@{ oidcToken = $script:Jwt } }
        Mock Invoke-RestMethod -ModuleName YardstickFederation -ParameterFilter { $Uri -like 'https://login.microsoftonline.com/*' } {
            $script:GraphTokenCount++
            [pscustomobject]@{ access_token = "graph-token-$($script:GraphTokenCount)"; token_type = 'Bearer'; expires_in = 3599 }
        }
        # The secret flow's own request lives in YardstickGraph; it must never run.
        Mock Invoke-RestMethod -ModuleName YardstickGraph { throw 'The client secret request was made in federated mode.' }
    }

    It 'signs in without any client secret' {
        Connect-YardstickGraph
        $Global:AuthenticationHeader.Authorization | Should -Be 'Bearer graph-token-1'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 1 -Exactly -ParameterFilter {
            $Uri -eq 'https://login.microsoftonline.com/tenant-fed/oauth2/v2.0/token'
        }
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 0 -Exactly
    }

    It 'ignores a stale client secret left in the session' {
        $Global:TenantID = 'tenant-old'
        $Global:ClientID = 'client-old'
        $Global:ClientSecret = 'stale-secret'
        Connect-YardstickGraph
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 1 -Exactly -ParameterFilter {
            $Uri -eq 'https://login.microsoftonline.com/tenant-fed/oauth2/v2.0/token' -and
            $Body.client_id -eq '66666666-7777-8888-9999-000000000000' -and
            -not $Body.ContainsKey('client_secret')
        }
    }

    It 'reuses a Graph token that is safely inside its lifetime' {
        Connect-YardstickGraph
        Connect-YardstickGraph
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 1 -Exactly -ParameterFilter { $Uri -like '*oidctoken*' }
    }

    It 'does not reuse a token from a client-secret sign-in once federation is on' {
        Clear-FederationEnvironment
        Mock Invoke-RestMethod -ModuleName YardstickGraph { [pscustomobject]@{ access_token = 'secret-token'; token_type = 'Bearer'; expires_in = 3600 } }
        Connect-YardstickGraph -TenantID 'tenant-old' -ClientID 'client-old' -ClientSecret 'secret'
        Set-FederationEnvironment
        Connect-YardstickGraph
        $Global:AuthenticationHeader.Authorization | Should -Be 'Bearer graph-token-1'
    }

    It 'does not reuse a token for a different app' {
        Connect-YardstickGraph
        Connect-YardstickGraph -ClientID '99999999-8888-7777-6666-555555555555'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 2 -Exactly -ParameterFilter { $Uri -like '*oidctoken*' }
    }

    It 'gets a fresh assertion on -Force' {
        Connect-YardstickGraph
        Connect-YardstickGraph -Force
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 2 -Exactly -ParameterFilter { $Uri -like '*oidctoken*' }
        $Global:AuthenticationHeader.Authorization | Should -Be 'Bearer graph-token-2'
    }

    It 'renews through Invoke-YardstickGraphRequest when the token nears expiry' {
        Connect-YardstickGraph
        InModuleScope YardstickGraph { $script:AccessToken.ExpiresOn = [datetimeoffset]::UtcNow.AddMinutes(2) }
        $Global:AuthenticationHeader.ExpiresOn = [datetimeoffset]::UtcNow.AddMinutes(2)
        Mock Invoke-RestMethod -ModuleName YardstickGraph -ParameterFilter { $Uri -like 'https://graph.microsoft.com/*' } { [pscustomobject]@{ id = 'app' } }

        (Invoke-YardstickGraphRequest -Resource 'deviceAppManagement/mobileApps/app').id | Should -Be 'app'

        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 2 -Exactly -ParameterFilter { $Uri -like '*oidctoken*' }
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter {
            $Uri -like 'https://graph.microsoft.com/*' -and $Headers.Authorization -eq 'Bearer graph-token-2'
        }
    }

    It 'uses an explicitly passed client secret instead' {
        Mock Invoke-RestMethod -ModuleName YardstickGraph { [pscustomobject]@{ access_token = 'secret-token'; token_type = 'Bearer'; expires_in = 3600 } }
        Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter { $Body.client_secret -eq 'secret' }
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 0 -Exactly
    }

    It 'fails without trying a secret when federation is only partly configured' {
        [Environment]::SetEnvironmentVariable('SYSTEM_ACCESSTOKEN', $null)
        $Global:TenantID = 'tenant'
        $Global:ClientID = 'client'
        $Global:ClientSecret = 'secret'
        { Connect-YardstickGraph } | Should -Throw '*does not fall back to a client secret*'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 0 -Exactly
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 0 -Exactly
    }

    It 'propagates an Entra failure instead of publishing a token' {
        Mock Invoke-RestMethod -ModuleName YardstickFederation -ParameterFilter { $Uri -like 'https://login.microsoftonline.com/*' } {
            throw (New-HttpErrorRecord -StatusCode 401 -Body '{"error":"invalid_client","error_description":"AADSTS700024: Client assertion is not within its valid time range."}')
        }
        { Connect-YardstickGraph } | Should -Throw '*AADSTS700024*'
        Get-Variable AuthenticationHeader -Scope Global -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
    }
}

Describe 'Connect-YardstickGraph without federation' {
    BeforeEach {
        Clear-FederationEnvironment
        Remove-Variable Token, AuthenticationHeader -Scope Global -ErrorAction SilentlyContinue
        InModuleScope YardstickGraph { $script:AccessToken = $null }
        Mock Invoke-RestMethod -ModuleName YardstickGraph { [pscustomobject]@{ access_token = 'secret-token'; token_type = 'Bearer'; expires_in = 3600 } }
        Mock Invoke-RestMethod -ModuleName YardstickFederation { throw 'federation must not be used' }
    }

    It 'keeps the client secret flow' {
        Connect-YardstickGraph -TenantID 'tenant' -ClientID 'client' -ClientSecret 'secret'
        Should -Invoke Invoke-RestMethod -ModuleName YardstickGraph -Times 1 -Exactly -ParameterFilter {
            $Body.client_secret -eq 'secret' -and -not $Body.ContainsKey('client_assertion')
        }
        Should -Invoke Invoke-RestMethod -ModuleName YardstickFederation -Times 0 -Exactly
    }
}

Describe 'Credential lifecycle in federated mode' {
    BeforeEach {
        Clear-FederationEnvironment
        Set-FederationEnvironment
        Remove-Variable TenantID, ClientID, ClientSecret, YardstickCredential -Scope Global -ErrorAction SilentlyContinue
        Mock Get-YardstickIntuneCredential -ModuleName YardstickCredential { throw 'Credential Manager must not be read in federated mode.' }
        Mock Register-YardstickIntuneCredential -ModuleName YardstickCredential { throw 'No prompt in federated mode.' }
        Mock Set-YardstickIntuneCredential -ModuleName YardstickCredential { throw 'Credential Manager must not be written in federated mode.' }
    }

    It 'Initialize-YardstickIntuneCredential skips Credential Manager under -NoPrompt' {
        $credential = Initialize-YardstickIntuneCredential -Preferences @{} -NoPrompt
        $credential.Federated | Should -BeTrue
        $credential.ClientSecret | Should -BeNullOrEmpty
        $Global:TenantID | Should -Be 'tenant-fed'
        $Global:ClientID | Should -Be $script:ClientId
        $Global:ClientSecret | Should -BeNullOrEmpty
        Should -Invoke Get-YardstickIntuneCredential -ModuleName YardstickCredential -Times 0 -Exactly
    }

    It 'does not migrate legacy secrets from preferences.yaml' {
        $prefs = @{ TenantID = 'legacy-tenant'; ClientID = 'legacy-client'; ClientSecret = 'legacy-secret' }
        $credential = Initialize-YardstickIntuneCredential -Preferences $prefs
        $credential.TenantID | Should -Be 'tenant-fed'
        $Global:ClientSecret | Should -BeNullOrEmpty
        Should -Invoke Set-YardstickIntuneCredential -ModuleName YardstickCredential -Times 0 -Exactly
    }

    It 'clears a client secret left over in the session' {
        $Global:ClientSecret = 'stale-secret'
        Initialize-YardstickIntuneCredential -Preferences @{} | Out-Null
        $Global:ClientSecret | Should -BeNullOrEmpty
    }

    It 'Update-YardstickSecretExpiration makes no Graph call' {
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickCredential { throw 'Application.Read.All is not granted.' }
        $credential = Initialize-YardstickIntuneCredential -Preferences @{}
        (Update-YardstickSecretExpiration -Credential $credential).Federated | Should -BeTrue
        Should -Invoke Invoke-YardstickGraphRequest -ModuleName YardstickCredential -Times 0 -Exactly
    }

    It 'Test-YardstickSecretExpiration neither warns nor emails' {
        Mock Send-YardstickCredentialExpiryEmail -ModuleName YardstickCredential { $true }
        Mock Write-Warning -ModuleName YardstickCredential {}
        $credential = Initialize-YardstickIntuneCredential -Preferences @{}
        $result = Test-YardstickSecretExpiration -Credential $credential -Preferences @{ adminEmailRecipient = 'admin@example.com' }
        $result.IsExpired | Should -BeFalse
        $result.IsExpiring | Should -BeFalse
        $result.Notified | Should -BeFalse
        Should -Invoke Send-YardstickCredentialExpiryEmail -ModuleName YardstickCredential -Times 0 -Exactly
        Should -Invoke Write-Warning -ModuleName YardstickCredential -Times 0 -Exactly
    }

    It 'fails fast when federation is partly configured' {
        [Environment]::SetEnvironmentVariable('YARDSTICK_CLIENT_ID', $null)
        { Initialize-YardstickIntuneCredential -Preferences @{} -NoPrompt } | Should -Throw '*YARDSTICK_CLIENT_ID is not set*'
        Should -Invoke Get-YardstickIntuneCredential -ModuleName YardstickCredential -Times 0 -Exactly
    }
}

Describe 'Credential lifecycle without federation' {
    BeforeEach { Clear-FederationEnvironment }

    It 'Initialize-YardstickIntuneCredential still reads Credential Manager' {
        Mock Get-YardstickIntuneCredential -ModuleName YardstickCredential {
            [pscustomobject]@{ Target = 't'; TenantID = 'tenant'; ClientID = 'client'; ClientSecret = 'secret'; SecretExpiresOn = $null; LastNotifiedOn = $null }
        }
        $credential = Initialize-YardstickIntuneCredential -Preferences @{} -NoPrompt
        $credential.ClientSecret | Should -Be 'secret'
        Test-YardstickFederatedCredential $credential | Should -BeFalse
        Should -Invoke Get-YardstickIntuneCredential -ModuleName YardstickCredential -Times 1 -Exactly
    }
}

# A second sign-in path, or an entry point that loads credentials some other way,
# would quietly need a client secret again in federated mode.
Describe 'Every sign-in goes through Connect-YardstickGraph' {
    BeforeAll {
        $script:SourceFiles = @(
            Get-ChildItem -Path $script:RepoRoot -Filter *.ps1 -File
            Get-ChildItem -Path (Join-Path $script:RepoRoot 'Modules') -Filter *.psm1 -File -Recurse
            Get-ChildItem -Path (Join-Path $script:RepoRoot 'Private') -Filter *.ps1 -File
        )
    }

    # YardstickSetup signs an administrator in for Install-Yardstick only; runs never load it for Graph.
    It 'Entra is only called from YardstickGraph, YardstickFederation and the installer' {
        $offenders = $script:SourceFiles |
            Where-Object { $_.Name -notin 'YardstickGraph.psm1', 'YardstickFederation.psm1', 'YardstickSetup.psm1' } |
            Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -match 'login\.microsoftonline\.com|client_secret|client_assertion' } |
            ForEach-Object Name
        $offenders | Should -BeNullOrEmpty -Because 'a new sign-in path would bypass workload identity federation'
    }

    It 'the runner loads credentials only through Initialize-YardstickIntuneCredential, before signing in' {
        $wrapper = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Yardstick.ps1') -Raw
        $wrapper | Should -Not -Match '\b(Get|Set|Register|Initialize)-YardstickIntuneCredential\b|Connect-YardstickGraph'
        $text = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Private\Invoke-YardstickRun.ps1') -Raw
        $text | Should -Not -Match '\b(Get|Set|Register)-YardstickIntuneCredential\b'
        $initialize = $text.IndexOf('Initialize-YardstickIntuneCredential')
        $connect = $text.IndexOf('Connect-YardstickGraph')
        $initialize | Should -BeGreaterThan -1
        $connect | Should -BeGreaterThan $initialize
    }
}
