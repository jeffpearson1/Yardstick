<#
.SYNOPSIS
Workload identity federation: signing in to Microsoft Graph without a client secret.

.DESCRIPTION
When the federated environment variables are set, Yardstick proves its identity to
Entra ID with a short-lived ID token instead of a client secret. The ID token comes
from one of three sources:

  Source           Selected by                         ID token
  ------           -----------                         --------
  File             YARDSTICK_FEDERATED_TOKEN_FILE       Read from that file
  Azure DevOps     YARDSTICK_SERVICE_CONNECTION_ID      Requested from SYSTEM_OIDCREQUESTURI
  GitHub Actions   ACTIONS_ID_TOKEN_REQUEST_URL         Requested from that URL

The ID token is sent to Entra as client_assertion in the same client-credentials
request the secret flow makes. A fresh one is obtained for every Graph token, so a
run is limited only by the job timeout. Nothing is cached here, and the job tokens
used to request an ID token are read from the environment at request time, never
stored.

Setting YARDSTICK_TENANT_ID, YARDSTICK_CLIENT_ID, YARDSTICK_SERVICE_CONNECTION_ID or
YARDSTICK_FEDERATED_TOKEN_FILE turns federated mode on, and then the tenant, the
client and exactly one source are required. Federated mode never falls back to a
client secret. See Docs/FederatedIdentity.md.
#>

Set-StrictMode -Version 3.0

Import-Module (Join-Path $PSScriptRoot 'YardstickLog.psm1')

$script:FederationTriggerVariables = @(
    'YARDSTICK_TENANT_ID',
    'YARDSTICK_CLIENT_ID',
    'YARDSTICK_SERVICE_CONNECTION_ID',
    'YARDSTICK_FEDERATED_TOKEN_FILE'
)
$script:AzureDevOpsOidcApiVersion = '7.1'
$script:EntraTokenExchangeAudience = 'api://AzureADTokenExchange'
$script:FederationMaxAttempts = 3
$script:FederationMaxRetryDelaySeconds = 60
$script:ClientAssertionType = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'


function Get-FederationEnvironmentValue {
    param([Parameter(Mandatory)][string]$Name)

    $value = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    return $value.Trim()
}


function Test-FederationRequestUri {
    # Plain HTTP only to this machine, so a job token is never sent unencrypted.
    param([string]$Uri)

    $parsed = $null
    if (-not [uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$parsed)) { return $false }
    return $parsed.Scheme -eq 'https' -or ($parsed.Scheme -eq 'http' -and $parsed.IsLoopback)
}


function Test-UnresolvedPipelineMacro {
    # Azure DevOps leaves "$(Name)" in place when the variable does not exist,
    # which is the usual symptom of a typo or of System.AccessToken not being mapped.
    param([string]$Value)
    return [bool]($Value -match '^\$\([^)]+\)$')
}


function Get-FederationSourceName {
    param([Parameter(Mandatory)][string]$Provider)

    switch ($Provider) {
        'AzureDevOps' { 'Azure DevOps' }
        'GitHubActions' { 'GitHub Actions' }
        default { 'token file' }
    }
}


function Get-YardstickFederatedIdentity {
    <#
    .SYNOPSIS
    Returns the federated sign-in settings, or $null when federated mode is off.

    .DESCRIPTION
    Federated mode is on when YARDSTICK_TENANT_ID, YARDSTICK_CLIENT_ID,
    YARDSTICK_SERVICE_CONNECTION_ID or YARDSTICK_FEDERATED_TOKEN_FILE is set. Once
    it is on, a missing, malformed or ambiguous input throws instead of silently
    using a client secret.

    The returned object never holds a token.
    #>
    [CmdletBinding()]
    param()

    $triggered = @($script:FederationTriggerVariables | Where-Object { -not [string]::IsNullOrEmpty([Environment]::GetEnvironmentVariable($_)) })
    if ($triggered.Count -eq 0) { return $null }

    $values = @{}
    foreach ($name in @($script:FederationTriggerVariables) + 'SYSTEM_OIDCREQUESTURI', 'SYSTEM_ACCESSTOKEN',
        'ACTIONS_ID_TOKEN_REQUEST_URL', 'ACTIONS_ID_TOKEN_REQUEST_TOKEN', 'GITHUB_ACTIONS') {
        $values[$name] = Get-FederationEnvironmentValue -Name $name
    }

    $problems = [System.Collections.Generic.List[string]]::new()
    foreach ($name in $script:FederationTriggerVariables) {
        if ($name -in $triggered -and -not $values[$name]) {
            $problems.Add("$name is set but blank")
        } elseif ($values[$name] -and (Test-UnresolvedPipelineMacro $values[$name])) {
            $problems.Add("$name is the unresolved pipeline variable '$($values[$name])'")
        }
    }
    foreach ($name in 'YARDSTICK_TENANT_ID', 'YARDSTICK_CLIENT_ID') {
        if ($name -notin $triggered) { $problems.Add("$name is not set") }
    }

    $guid = [guid]::Empty
    foreach ($name in 'YARDSTICK_CLIENT_ID', 'YARDSTICK_SERVICE_CONNECTION_ID') {
        $value = $values[$name]
        if ($value -and -not (Test-UnresolvedPipelineMacro $value) -and -not [guid]::TryParse($value, [ref]$guid)) {
            $problems.Add("$name '$value' is not a GUID")
        }
    }

    $provider = $null
    if ($values.YARDSTICK_FEDERATED_TOKEN_FILE -and $values.YARDSTICK_SERVICE_CONNECTION_ID) {
        $problems.Add('YARDSTICK_FEDERATED_TOKEN_FILE and YARDSTICK_SERVICE_CONNECTION_ID are both set; set only the one for where Yardstick runs')
    } elseif ($values.YARDSTICK_FEDERATED_TOKEN_FILE) {
        $provider = 'File'
    } elseif ($values.YARDSTICK_SERVICE_CONNECTION_ID) {
        $provider = 'AzureDevOps'
    } elseif ($values.ACTIONS_ID_TOKEN_REQUEST_URL) {
        $provider = 'GitHubActions'
    } elseif ($values.GITHUB_ACTIONS -eq 'true') {
        $problems.Add("GitHub Actions is not offering an ID token; give the job 'permissions: id-token: write'")
    } else {
        $problems.Add('no ID token source is set: set YARDSTICK_FEDERATED_TOKEN_FILE, set YARDSTICK_SERVICE_CONNECTION_ID in an Azure DevOps pipeline, or run in GitHub Actions with id-token: write')
    }

    $requestUri = $null
    $tokenFile = $null
    switch ($provider) {
        'File' {
            $tokenFile = $values.YARDSTICK_FEDERATED_TOKEN_FILE
            if (-not (Test-UnresolvedPipelineMacro $tokenFile)) {
                if (-not [IO.Path]::IsPathFullyQualified($tokenFile)) {
                    # Yardstick changes directory during a run, so a relative path would move.
                    $problems.Add("YARDSTICK_FEDERATED_TOKEN_FILE '$tokenFile' is not a full path")
                } elseif (-not (Test-Path -LiteralPath $tokenFile -PathType Leaf)) {
                    $problems.Add("YARDSTICK_FEDERATED_TOKEN_FILE '$tokenFile' does not exist")
                }
            }
        }
        'AzureDevOps' {
            $requestUri = $values.SYSTEM_OIDCREQUESTURI
            if (-not $requestUri) {
                $problems.Add('SYSTEM_OIDCREQUESTURI is not set (it is predefined in Azure DevOps jobs; Yardstick must run inside a pipeline job)')
            } elseif (-not (Test-FederationRequestUri $requestUri)) {
                $problems.Add('SYSTEM_OIDCREQUESTURI is not an absolute https URI')
            }
            if (-not $values.SYSTEM_ACCESSTOKEN -or (Test-UnresolvedPipelineMacro $values.SYSTEM_ACCESSTOKEN)) {
                $problems.Add('SYSTEM_ACCESSTOKEN is not available; map it on the step with "SYSTEM_ACCESSTOKEN: $(System.AccessToken)"')
            }
        }
        'GitHubActions' {
            $requestUri = $values.ACTIONS_ID_TOKEN_REQUEST_URL
            if (-not (Test-FederationRequestUri $requestUri)) {
                $problems.Add('ACTIONS_ID_TOKEN_REQUEST_URL is not an absolute https URI')
            }
            if (-not $values.ACTIONS_ID_TOKEN_REQUEST_TOKEN) {
                $problems.Add("ACTIONS_ID_TOKEN_REQUEST_TOKEN is not set; give the job 'permissions: id-token: write'")
            }
        }
    }

    if ($problems.Count -gt 0) {
        throw "Workload identity federation is enabled ($($triggered -join ', ') set) but is not configured correctly: $($problems -join '; '). Yardstick does not fall back to a client secret."
    }

    return [pscustomobject]@{
        TenantID            = $values.YARDSTICK_TENANT_ID
        ClientID            = $values.YARDSTICK_CLIENT_ID
        Provider            = $provider
        ServiceConnectionID = $values.YARDSTICK_SERVICE_CONNECTION_ID
        RequestUri          = $requestUri
        TokenFile           = $tokenFile
    }
}


function Get-YardstickFederatedCredential {
    <#
    .SYNOPSIS
    In federated mode, publishes the tenant and client IDs as the globals that
    Connect-YardstickGraph reads and returns a secret-less credential object.
    Returns $null when federated mode is off.

    .DESCRIPTION
    Called at the top of Initialize-YardstickIntuneCredential so Credential Manager
    is never touched. The object mirrors the shape of a stored credential, with
    Federated = $true, so the secret expiration functions can recognise it and do
    nothing.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'Connect-YardstickGraph and the entry scripts read these globals.')]
    [CmdletBinding()]
    param()

    $identity = Get-YardstickFederatedIdentity
    if (-not $identity) { return $null }

    $credential = [pscustomobject]@{
        Target          = $null
        TenantID        = $identity.TenantID
        ClientID        = $identity.ClientID
        ClientSecret    = $null
        SecretExpiresOn = $null
        LastNotifiedOn  = $null
        Federated       = $true
        Provider        = $identity.Provider
    }

    $Global:TenantID = $credential.TenantID
    $Global:ClientID = $credential.ClientID
    $Global:ClientSecret = $null
    $Global:YardstickCredential = $credential

    Write-Log "Using workload identity federation ($(Get-YardstickFederatedSourceDescription -Identity $identity)) for app $($credential.ClientID). No client secret is loaded." -NoContext
    return $credential
}


function Get-YardstickFederatedSourceDescription {
    <#
    .SYNOPSIS
    Describes where the ID token comes from, for logs and the health check.
    #>
    param([Parameter(Mandatory)]$Identity)

    switch ($Identity.Provider) {
        'AzureDevOps' { "Azure DevOps service connection $($Identity.ServiceConnectionID)" }
        'GitHubActions' { 'GitHub Actions' }
        default { "token file $($Identity.TokenFile)" }
    }
}


function Test-YardstickFederatedCredential {
    <#
    .SYNOPSIS
    True when a credential object came from Get-YardstickFederatedCredential.
    #>
    param($InputObject)

    return [bool]($InputObject -and $InputObject.PSObject.Properties['Federated'] -and $InputObject.Federated)
}


function Get-FederationErrorStatusCode {
    param([Parameter(Mandatory)]$ErrorRecord)

    $responseProperty = $ErrorRecord.Exception.PSObject.Properties['Response']
    $response = if ($responseProperty) { $responseProperty.Value } else { $null }
    if ($null -eq $response) { return $null }
    try {
        if ($response.StatusCode.value__) { return [int]$response.StatusCode.value__ }
        return [int]$response.StatusCode
    } catch {
        return $null
    }
}


function Get-FederationErrorDetail {
    param([Parameter(Mandatory)]$ErrorRecord)

    $raw = $null
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $raw = [string]$ErrorRecord.ErrorDetails.Message
    }
    if ([string]::IsNullOrWhiteSpace($raw)) { $raw = [string]$ErrorRecord.Exception.Message }

    $code = $null
    $detail = $raw
    try {
        $body = $raw | ConvertFrom-Json -ErrorAction Stop
        if ($body.PSObject.Properties['error_description'] -and $body.error_description) {
            $detail = [string]$body.error_description
            if ($body.PSObject.Properties['error'] -and $body.error) { $code = [string]$body.error }
        } elseif ($body.PSObject.Properties['message'] -and $body.message) {
            $detail = [string]$body.message
        } else {
            $detail = $body | ConvertTo-Json -Depth 10 -Compress
        }
    } catch {
        $null = $_
    }

    # Entra descriptions carry a trace ID, correlation ID and timestamp on later lines.
    $detail = (($detail -split "\r?\n") | Where-Object { $_.Trim() } | Select-Object -First 1)
    if ($detail) { $detail = $detail.Trim() }
    if ($detail -and $detail.Length -gt 1000) { $detail = $detail.Substring(0, 1000) + '...' }
    $aadsts = if ($detail -match '\b(AADSTS\d+)\b') { $Matches[1] } else { $null }

    return [pscustomobject]@{
        Code    = $code
        AadSts  = $aadsts
        Message = $detail
    }
}


function Get-FederationRetryDelay {
    param([Parameter(Mandatory)]$ErrorRecord, [Parameter(Mandatory)][int]$Attempt)

    $delay = [math]::Pow(2, $Attempt)
    $responseProperty = $ErrorRecord.Exception.PSObject.Properties['Response']
    $response = if ($responseProperty) { $responseProperty.Value } else { $null }
    $retryAfter = if ($response -and $response.PSObject.Properties['Headers'] -and $response.Headers.PSObject.Properties['RetryAfter']) { $response.Headers.RetryAfter }
    if ($retryAfter) {
        if ($null -ne $retryAfter.Delta) {
            $delay = $retryAfter.Delta.TotalSeconds
        } elseif ($null -ne $retryAfter.Date) {
            $delay = ($retryAfter.Date - [datetimeoffset]::UtcNow).TotalSeconds
        }
    }
    return [int][math]::Ceiling([math]::Min([math]::Max($delay, 1), $script:FederationMaxRetryDelaySeconds))
}


function Invoke-FederationRestMethod {
    # Invoke-RestMethod with retries for transient failures. Non-transient failures
    # and exhausted retries rethrow the original error record.
    param(
        [Parameter(Mandatory)][hashtable]$Parameters
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            # Explicit: module code doesn't see the runner's global -TimeoutSec default,
            # and an unresponsive endpoint would otherwise hang the run.
            return Invoke-RestMethod @Parameters -TimeoutSec 60 -ErrorAction Stop
        } catch {
            $status = Get-FederationErrorStatusCode -ErrorRecord $_
            $transient = ($null -eq $status) -or ($status -in 408, 429) -or ($status -ge 500 -and $status -le 599)
            if (-not $transient -or $attempt -ge $script:FederationMaxAttempts) { throw }
            Start-Sleep -Seconds (Get-FederationRetryDelay -ErrorRecord $_ -Attempt $attempt)
        }
    }
}


function Add-FederationQueryString {
    param([string]$Uri, [string]$Query)

    $separator = if ($Uri.Contains('?')) { '&' } else { '?' }
    return "$Uri$separator$Query"
}


function Get-AzureDevOpsAssertion {
    param([Parameter(Mandatory)]$Identity)

    $accessToken = Get-FederationEnvironmentValue -Name 'SYSTEM_ACCESSTOKEN'
    if (-not $accessToken -or (Test-UnresolvedPipelineMacro $accessToken)) {
        throw 'Workload identity federation: SYSTEM_ACCESSTOKEN is no longer available, so a new ID token cannot be requested from Azure DevOps.'
    }

    $uri = Add-FederationQueryString -Uri $Identity.RequestUri -Query ('api-version={0}&serviceConnectionId={1}' -f
        $script:AzureDevOpsOidcApiVersion, [uri]::EscapeDataString($Identity.ServiceConnectionID))

    try {
        $response = Invoke-FederationRestMethod -Parameters @{
            Uri         = $uri
            Method      = 'Post'
            ContentType = 'application/json'
            Headers     = @{ Authorization = "Bearer $accessToken" }
            # Without a body Azure DevOps answers 400 "The request is invalid".
            Body        = '{}'
        }
    } catch {
        $status = Get-FederationErrorStatusCode -ErrorRecord $_
        $detail = Get-FederationErrorDetail -ErrorRecord $_
        $hint = switch ($status) {
            401 { ' The job access token was rejected; check that SYSTEM_ACCESSTOKEN is mapped from $(System.AccessToken).' }
            { $_ -in 400, 403, 404 } { " Check that service connection $($Identity.ServiceConnectionID) exists, uses workload identity federation, is referenced by a task in this stage, and that this pipeline is permitted to use it." }
            default { '' }
        }
        $statusText = if ($status) { "HTTP $status" } else { 'no response' }
        throw "Workload identity federation: requesting an ID token from Azure DevOps failed ($statusText).$hint $($detail.Message)".TrimEnd()
    }

    if ($response -and $response.PSObject.Properties['oidcToken']) { return [string]$response.oidcToken }
    return $null
}


function Get-GitHubActionsAssertion {
    param([Parameter(Mandatory)]$Identity)

    $requestToken = Get-FederationEnvironmentValue -Name 'ACTIONS_ID_TOKEN_REQUEST_TOKEN'
    if (-not $requestToken) {
        throw 'Workload identity federation: ACTIONS_ID_TOKEN_REQUEST_TOKEN is no longer available, so a new ID token cannot be requested from GitHub Actions.'
    }

    $uri = Add-FederationQueryString -Uri $Identity.RequestUri -Query "audience=$([uri]::EscapeDataString($script:EntraTokenExchangeAudience))"

    try {
        $response = Invoke-FederationRestMethod -Parameters @{
            Uri     = $uri
            Method  = 'Get'
            Headers = @{ Authorization = "Bearer $requestToken" }
        }
    } catch {
        $status = Get-FederationErrorStatusCode -ErrorRecord $_
        $detail = Get-FederationErrorDetail -ErrorRecord $_
        $hint = if ($status -in 401, 403) { " Check that the job has 'permissions: id-token: write'." } else { '' }
        $statusText = if ($status) { "HTTP $status" } else { 'no response' }
        throw "Workload identity federation: requesting an ID token from GitHub Actions failed ($statusText).$hint $($detail.Message)".TrimEnd()
    }

    if ($response -and $response.PSObject.Properties['value']) { return [string]$response.value }
    return $null
}


function Get-FileAssertion {
    param([Parameter(Mandatory)]$Identity)

    try {
        return [string](Get-Content -LiteralPath $Identity.TokenFile -Raw -ErrorAction Stop)
    } catch {
        throw "Workload identity federation: reading the ID token from '$($Identity.TokenFile)' failed: $($_.Exception.Message)"
    }
}


function Get-YardstickFederatedAssertion {
    <#
    .SYNOPSIS
    Gets a fresh ID token from the identity's source.

    .OUTPUTS
    The ID token (a JWT) as a string. Callers must not log it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Identity
    )

    $idToken = switch ($Identity.Provider) {
        'AzureDevOps' { Get-AzureDevOpsAssertion -Identity $Identity }
        'GitHubActions' { Get-GitHubActionsAssertion -Identity $Identity }
        'File' { Get-FileAssertion -Identity $Identity }
        default { throw "Workload identity federation: unknown ID token source '$($Identity.Provider)'." }
    }

    if ([string]::IsNullOrWhiteSpace($idToken)) {
        throw "Workload identity federation: $(Get-FederationSourceName -Provider $Identity.Provider) returned no ID token."
    }
    return $idToken.Trim()
}


function Get-YardstickJwtClaimSummary {
    <#
    .SYNOPSIS
    Describes the non-secret identifying claims of a JWT (issuer, subject,
    audience), to compare with the app registration's federated credential.
    The signature is not checked and the token itself is never returned.
    #>
    param([string]$Jwt)

    try {
        $parts = $Jwt.Split('.')
        if ($parts.Count -lt 2) { return $null }
        $payload = $parts[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) {
            2 { $payload += '==' }
            3 { $payload += '=' }
        }
        $claims = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json -ErrorAction Stop
        $describe = foreach ($name in 'iss', 'sub', 'aud') {
            if ($claims.PSObject.Properties[$name]) { "$name '$(@($claims.$name) -join ',')'" }
        }
        if (-not $describe) { return $null }
        return ($describe -join ', ')
    } catch {
        return $null
    }
}


function Invoke-YardstickFederatedTokenRequest {
    <#
    .SYNOPSIS
    Gets a fresh ID token and exchanges it with Entra ID for a Graph token.

    .DESCRIPTION
    Used by Connect-YardstickGraph in place of the client-secret request. Returns
    Entra's token response unchanged, so the caller caches it exactly as it does
    in the secret flow.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TokenUri,
        [Parameter(Mandatory)][string]$ClientID,
        [Parameter(Mandatory)]$Identity,
        [string]$Scope = 'https://graph.microsoft.com/.default'
    )

    $assertion = Get-YardstickFederatedAssertion -Identity $Identity

    try {
        return Invoke-FederationRestMethod -Parameters @{
            Uri         = $TokenUri
            Method      = 'Post'
            ContentType = 'application/x-www-form-urlencoded'
            Body        = @{
                client_id             = $ClientID
                client_assertion_type = $script:ClientAssertionType
                client_assertion      = $assertion
                scope                 = $Scope
                grant_type            = 'client_credentials'
            }
        }
    } catch {
        $status = Get-FederationErrorStatusCode -ErrorRecord $_
        $detail = Get-FederationErrorDetail -ErrorRecord $_
        $statusText = if ($status) { "HTTP $status" } else { 'no response' }
        $codes = @($statusText, $detail.AadSts | Where-Object { $_ }) -join ', '
        $claims = Get-YardstickJwtClaimSummary -Jwt $assertion
        $claimText = if ($claims) { " The ID token presented $claims; the app registration's federated credential must match the issuer and subject." } else { '' }
        throw "Workload identity federation: exchanging the $(Get-FederationSourceName -Provider $Identity.Provider) ID token with Entra ID failed ($codes). $($detail.Message)$claimText"
    } finally {
        $assertion = $null
    }
}


Export-ModuleMember -Function @(
    'Get-YardstickFederatedIdentity',
    'Get-YardstickFederatedCredential',
    'Get-YardstickFederatedSourceDescription',
    'Test-YardstickFederatedCredential',
    'Get-YardstickFederatedAssertion',
    'Get-YardstickJwtClaimSummary',
    'Invoke-YardstickFederatedTokenRequest'
)
