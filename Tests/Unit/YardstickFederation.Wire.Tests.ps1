# Exercises the federated sign-in over real HTTP against local stand-ins for the
# Azure DevOps and GitHub Actions ID token endpoints and the Entra token endpoint,
# so what goes on the wire
# (method, headers, query string, form encoding) and how real error responses are
# parsed are tested without mocking Invoke-RestMethod.

BeforeAll {
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickFederation.psm1" -Force -Global

    $script:FederationVariables = 'YARDSTICK_SERVICE_CONNECTION_ID', 'YARDSTICK_TENANT_ID', 'YARDSTICK_CLIENT_ID',
        'YARDSTICK_FEDERATED_TOKEN_FILE', 'SYSTEM_OIDCREQUESTURI', 'SYSTEM_ACCESSTOKEN',
        'ACTIONS_ID_TOKEN_REQUEST_URL', 'ACTIONS_ID_TOKEN_REQUEST_TOKEN', 'GITHUB_ACTIONS'
    $script:SavedEnvironment = @{}
    foreach ($name in $script:FederationVariables) {
        $script:SavedEnvironment[$name] = [Environment]::GetEnvironmentVariable($name)
    }

    $probe = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    $probe.Start()
    $script:Port = ([System.Net.IPEndPoint]$probe.LocalEndpoint).Port
    $probe.Stop()
    $script:BaseUri = "http://localhost:$script:Port"

    $script:Requests = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:Listener = [System.Net.HttpListener]::new()
    $script:Listener.Prefixes.Add("$script:BaseUri/")
    $script:Listener.Start()

    # Routes: /oidc/ok, /oidc/unauthorized, /oidc/empty, /github/ok, /entra/ok, /entra/aadsts.
    $listener = $script:Listener
    $requests = $script:Requests
    $script:Server = Start-ThreadJob -ScriptBlock {
        $Listener = $using:listener
        $Requests = $using:requests
        while ($Listener.IsListening) {
            try { $context = $Listener.GetContext() } catch { break }
            $request = $context.Request
            $reader = [System.IO.StreamReader]::new($request.InputStream, $request.ContentEncoding)
            $body = $reader.ReadToEnd()
            $Requests.Enqueue([pscustomobject]@{
                Method        = $request.HttpMethod
                Path          = $request.Url.AbsolutePath
                Query         = $request.Url.Query
                Authorization = $request.Headers['Authorization']
                ContentType   = $request.ContentType
                Body          = $body
            })

            $status = 200
            $payload = switch -Wildcard ($request.Url.AbsolutePath) {
                '/oidc/ok' { '{"oidcToken":"eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJzYzovL29yZy9wcm9qL2Nvbm4ifQ.sig"}' }
                '/oidc/empty' { '{"oidcToken":null}' }
                '/oidc/unauthorized' { $status = 401; '{"message":"TF400813: The user is not authorized to access this resource."}' }
                '/github/ok' { '{"value":"eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJyZXBvOm9yZy9yZXBvIn0.sig"}' }
                '/entra/ok' { '{"token_type":"Bearer","expires_in":3599,"access_token":"graph-token"}' }
                '/entra/aadsts' {
                    $status = 400
                    '{"error":"invalid_client","error_description":"AADSTS70021: No matching federated identity record found for presented assertion.\r\nTrace ID: 1\r\nCorrelation ID: 2"}'
                }
                default { $status = 404; '{"message":"not found"}' }
            }
            $bytes = [Text.Encoding]::UTF8.GetBytes($payload)
            $context.Response.StatusCode = $status
            $context.Response.ContentType = 'application/json'
            $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
            $context.Response.Close()
        }
    }

    function script:Set-WireEnvironment {
        param([string]$OidcPath = '/oidc/ok')
        foreach ($name in $script:FederationVariables) { [Environment]::SetEnvironmentVariable($name, $null) }
        [Environment]::SetEnvironmentVariable('YARDSTICK_SERVICE_CONNECTION_ID', '11111111-2222-3333-4444-555555555555')
        [Environment]::SetEnvironmentVariable('YARDSTICK_TENANT_ID', 'tenant')
        [Environment]::SetEnvironmentVariable('YARDSTICK_CLIENT_ID', '66666666-7777-8888-9999-000000000000')
        [Environment]::SetEnvironmentVariable('SYSTEM_OIDCREQUESTURI', "$script:BaseUri$OidcPath")
        [Environment]::SetEnvironmentVariable('SYSTEM_ACCESSTOKEN', 'wire-job-token')
    }

    function script:Set-GitHubWireEnvironment {
        foreach ($name in $script:FederationVariables) { [Environment]::SetEnvironmentVariable($name, $null) }
        [Environment]::SetEnvironmentVariable('YARDSTICK_TENANT_ID', 'tenant')
        [Environment]::SetEnvironmentVariable('YARDSTICK_CLIENT_ID', '66666666-7777-8888-9999-000000000000')
        [Environment]::SetEnvironmentVariable('ACTIONS_ID_TOKEN_REQUEST_URL', "$script:BaseUri/github/ok?api-version=2.0")
        [Environment]::SetEnvironmentVariable('ACTIONS_ID_TOKEN_REQUEST_TOKEN', 'wire-request-token')
    }

    function script:Get-LastRequest {
        $items = $script:Requests.ToArray()
        return $items[-1]
    }
}

AfterAll {
    $script:Listener.Stop()
    $script:Listener.Close()
    $script:Server | Wait-Job -Timeout 5 | Out-Null
    $script:Server | Remove-Job -Force
    foreach ($name in $script:FederationVariables) {
        [Environment]::SetEnvironmentVariable($name, $script:SavedEnvironment[$name])
    }
}

Describe 'Federated sign-in over HTTP' {
    BeforeEach { Mock Start-Sleep -ModuleName YardstickFederation {} }

    It 'requests the ID token the way Azure DevOps expects' {
        Set-WireEnvironment
        $token = Get-YardstickFederatedAssertion -Identity (Get-YardstickFederatedIdentity)
        $token | Should -BeLike 'eyJ*.sig'

        $request = Get-LastRequest
        $request.Method | Should -Be 'POST'
        $request.Path | Should -Be '/oidc/ok'
        $request.Query | Should -Be '?api-version=7.1&serviceConnectionId=11111111-2222-3333-4444-555555555555'
        $request.Authorization | Should -Be 'Bearer wire-job-token'
        $request.ContentType | Should -BeLike 'application/json*'
        $request.Body | Should -Be '{}'
    }

    It 'reports a real 401 from Azure DevOps with its message' {
        Set-WireEnvironment -OidcPath '/oidc/unauthorized'
        try { Get-YardstickFederatedAssertion -Identity (Get-YardstickFederatedIdentity) } catch { $message = $_.Exception.Message }
        $message | Should -Match 'failed \(HTTP 401\)'
        $message | Should -Match 'TF400813'
        $message | Should -Not -Match 'wire-job-token'
    }

    It 'requests the ID token the way GitHub Actions expects' {
        Set-GitHubWireEnvironment
        $token = Get-YardstickFederatedAssertion -Identity (Get-YardstickFederatedIdentity)
        $token | Should -BeLike 'eyJ*.sig'

        $request = Get-LastRequest
        $request.Method | Should -Be 'GET'
        $request.Path | Should -Be '/github/ok'
        $request.Query | Should -Be '?api-version=2.0&audience=api%3A%2F%2FAzureADTokenExchange'
        $request.Authorization | Should -Be 'Bearer wire-request-token'
    }

    It 'treats a null oidcToken as a failure' {
        Set-WireEnvironment -OidcPath '/oidc/empty'
        { Get-YardstickFederatedAssertion -Identity (Get-YardstickFederatedIdentity) } | Should -Throw '*returned no ID token*'
    }

    It 'form-encodes the client assertion for Entra' {
        Set-WireEnvironment
        $response = Invoke-YardstickFederatedTokenRequest -TokenUri "$script:BaseUri/entra/ok" -ClientID '66666666-7777-8888-9999-000000000000' -Identity (Get-YardstickFederatedIdentity)
        $response.access_token | Should -Be 'graph-token'

        $request = Get-LastRequest
        $request.Method | Should -Be 'POST'
        $request.ContentType | Should -BeLike 'application/x-www-form-urlencoded*'
        $form = @{}
        foreach ($pair in $request.Body.Split('&')) {
            $key, $value = $pair.Split('=', 2)
            $form[[uri]::UnescapeDataString($key)] = [uri]::UnescapeDataString($value.Replace('+', ' '))
        }
        $form.grant_type | Should -Be 'client_credentials'
        $form.client_id | Should -Be '66666666-7777-8888-9999-000000000000'
        $form.scope | Should -Be 'https://graph.microsoft.com/.default'
        $form.client_assertion_type | Should -Be 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        $form.client_assertion | Should -BeLike 'eyJ*.sig'
        $form.ContainsKey('client_secret') | Should -BeFalse
    }

    It 'surfaces a real Entra AADSTS error on one line with the presented subject' {
        Set-WireEnvironment
        try {
            Invoke-YardstickFederatedTokenRequest -TokenUri "$script:BaseUri/entra/aadsts" -ClientID '66666666-7777-8888-9999-000000000000' -Identity (Get-YardstickFederatedIdentity)
        } catch { $message = $_.Exception.Message }
        $message | Should -Match '\(HTTP 400, AADSTS70021\)'
        $message | Should -Match "sub 'sc://org/proj/conn'"
        $message | Should -Not -Match 'Trace ID'
        $message | Should -Not -Match '\.sig'
    }
}
