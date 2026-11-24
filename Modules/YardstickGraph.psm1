Set-StrictMode -Version 3.0

Import-Module (Join-Path $PSScriptRoot 'YardstickFederation.psm1')
Import-Module (Join-Path $PSScriptRoot 'YardstickLog.psm1')

$script:AccessToken = $null
$script:GraphStatistics = $null

function Write-YardstickGraphLog {
    # One seam for this module's logging, so tests can mock it in one place.
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info', 'Detail', 'Warning', 'Error')][string]$Level = 'Info'
    )
    Write-Log $Message -Level $Level
}

function Reset-YardstickGraphStatistics {
    <#
    .SYNOPSIS
    Starts counting Graph requests and retries afresh (once per run).
    #>
    $script:GraphStatistics = [pscustomobject]@{
        Requests       = 0
        Retried        = 0
        RetrySleepSecs = 0
        GaveUp         = 0
    }
}

function Get-YardstickGraphStatistics {
    <#
    .SYNOPSIS
    Returns the Graph request, retry and give-up counts since the last reset.
    #>
    if (-not $script:GraphStatistics) { Reset-YardstickGraphStatistics }
    return $script:GraphStatistics
}

function Get-YardstickGraphResponseHeader {
    param($ErrorRecord, [Parameter(Mandatory)][string]$Name)
    try {
        $values = $null
        if ($ErrorRecord.Exception.Response.Headers.TryGetValues($Name, [ref]$values)) { return (@($values) -join ',') }
    } catch {
        $null = $_
    }
    return $null
}

function Format-YardstickGraphResourceForLog {
    # Decoded and shortened, so a log line stays readable. Never carries a token:
    # tokens travel in the Authorization header, not the URI.
    param([string]$Uri)
    $text = $Uri -replace '^https://graph\.microsoft\.com/', ''
    try { $text = [uri]::UnescapeDataString($text) } catch { $null = $_ }
    if ($text.Length -gt 160) { $text = $text.Substring(0, 157) + '...' }
    return $text
}

function Get-YardstickGlobalValue {
    param([Parameter(Mandatory)][string]$Name)
    $variable = Get-Variable -Name $Name -Scope Global -ErrorAction SilentlyContinue
    if ($variable) { return $variable.Value }
    return $null
}

function Get-YardstickGraphTokenExpiry {
    if ($script:AccessToken -and $script:AccessToken.ExpiresOn) {
        return [datetimeoffset]$script:AccessToken.ExpiresOn
    }
    $token = Get-YardstickGlobalValue -Name Token
    if ($token -and $token.ExpiresOn) {
        return [datetimeoffset]$token.ExpiresOn
    }
    $authenticationHeader = Get-YardstickGlobalValue -Name AuthenticationHeader
    if ($authenticationHeader -and $authenticationHeader.ExpiresOn) {
        return [datetimeoffset]$authenticationHeader.ExpiresOn
    }
    return $null
}

function Connect-YardstickGraph {
    <#
    .SYNOPSIS
    Acquires and caches a Microsoft Graph application token for Yardstick.

    .DESCRIPTION
    Uses the OAuth 2.0 client-credentials flow directly. Credentials default to
    the values populated by Initialize-YardstickIntuneCredential.

    When the workload identity federation variables are set and no -ClientSecret
    is passed, a fresh ID token is sent as the client assertion instead of a
    client secret. See Docs/FederatedIdentity.md.

    .PARAMETER PassThru
    Returns the cached token object. Omit it so the bearer token never lands on
    the console or in a transcript.
    #>
    [CmdletBinding()]
    param(
        [string]$TenantID,
        [string]$ClientID,
        [string]$ClientSecret,
        [switch]$Force,
        [switch]$PassThru
    )

    if (-not $PSBoundParameters.ContainsKey('TenantID')) { $TenantID = Get-YardstickGlobalValue -Name TenantID }
    if (-not $PSBoundParameters.ContainsKey('ClientID')) { $ClientID = Get-YardstickGlobalValue -Name ClientID }
    if (-not $PSBoundParameters.ContainsKey('ClientSecret')) { $ClientSecret = Get-YardstickGlobalValue -Name ClientSecret }

    $federatedIdentity = if (-not $PSBoundParameters.ContainsKey('ClientSecret')) { Get-YardstickFederatedIdentity }
    if ($federatedIdentity) {
        if (-not $PSBoundParameters.ContainsKey('TenantID')) { $TenantID = $federatedIdentity.TenantID }
        if (-not $PSBoundParameters.ContainsKey('ClientID')) { $ClientID = $federatedIdentity.ClientID }
    } elseif ([string]::IsNullOrWhiteSpace($TenantID) -or
        [string]::IsNullOrWhiteSpace($ClientID) -or
        [string]::IsNullOrWhiteSpace($ClientSecret)) {
        throw 'Required Graph credentials are not set: TenantID, ClientID, and ClientSecret are required.'
    }

    $context = '{0}|{1}|{2}' -f $TenantID, $ClientID, $(if ($federatedIdentity) { 'federated' } else { 'secret' })
    $expiresOn = Get-YardstickGraphTokenExpiry
    $authenticationHeader = Get-YardstickGlobalValue -Name AuthenticationHeader
    if (-not $Force -and $authenticationHeader -and $authenticationHeader.Authorization -and
        $expiresOn -and $expiresOn -gt [datetimeoffset]::UtcNow.AddMinutes(30) -and
        (-not $script:AccessToken -or $script:AccessToken.Context -eq $context)) {
        if ($PassThru) { return $script:AccessToken }
        return
    }

    $tokenUri = "https://login.microsoftonline.com/$([uri]::EscapeDataString($TenantID))/oauth2/v2.0/token"
    # Explicit rather than inherited: each module has its own session state, so the
    # $Global:PSDefaultParameterValues the runner sets for recipes is not visible
    # here. Without this, an unresponsive identity endpoint hangs the run forever.
    $tokenResponse = if ($federatedIdentity) {
        Invoke-YardstickFederatedTokenRequest -TokenUri $tokenUri -ClientID $ClientID -Identity $federatedIdentity
    } else { Invoke-RestMethod -Uri $tokenUri -Method Post -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 60 -Body @{
        client_id     = $ClientID
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    } -ErrorAction Stop }

    if (-not $tokenResponse.access_token) {
        throw 'Microsoft identity platform returned no access token.'
    }

    $lifetimeSeconds = if ($tokenResponse.expires_in) { [int]$tokenResponse.expires_in } else { 3600 }
    $expiresOn = [datetimeoffset]::UtcNow.AddSeconds($lifetimeSeconds)
    $script:AccessToken = [pscustomobject]@{
        AccessToken = [string]$tokenResponse.access_token
        TokenType   = if ($tokenResponse.token_type) { [string]$tokenResponse.token_type } else { 'Bearer' }
        ExpiresOn   = $expiresOn
        Context     = $context
    }

    # Keep these globals during the migration because credential and reporting
    # code already consumes the authentication header. The token itself remains
    # opaque to callers.
    $Global:Token = $script:AccessToken
    $Global:AuthenticationHeader = @{
        Authorization = "Bearer $($script:AccessToken.AccessToken)"
        ExpiresOn     = $expiresOn
    }

    if ($PassThru) { return $script:AccessToken }
}

function Disconnect-YardstickGraph {
    <#
    .SYNOPSIS
    Forgets the cached Graph token and the app registration credentials.

    .DESCRIPTION
    Invoke-Yardstick calls this when a run ends, so the client secret and a live
    token do not stay readable in the session that ran Yardstick.
    #>
    [CmdletBinding()]
    param()

    $script:AccessToken = $null
    foreach ($name in 'Token', 'AuthenticationHeader', 'TenantID', 'ClientID', 'ClientSecret', 'YardstickCredential') {
        Remove-Variable -Name $name -Scope Global -Force -ErrorAction SilentlyContinue
    }
}

function Get-YardstickGraphErrorStatusCode {
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

function Get-YardstickGraphRetryDelay {
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [Parameter(Mandatory)][int]$Attempt
    )

    try {
        $retryAfter = $ErrorRecord.Exception.Response.Headers.RetryAfter
        $cap = $script:MaxRetryAfterSeconds
        if ($retryAfter.Delta) {
            return [math]::Min($cap, [math]::Max(1, [int][math]::Ceiling($retryAfter.Delta.TotalSeconds)))
        }
        if ($retryAfter.Date) {
            return [math]::Min($cap, [math]::Max(1, [int][math]::Ceiling(($retryAfter.Date - [datetimeoffset]::UtcNow).TotalSeconds)))
        }
        $raw = $ErrorRecord.Exception.Response.Headers['Retry-After']
        if ($raw -as [int]) { return [math]::Min($cap, [math]::Max(1, [int]$raw)) }
    } catch {
        $null = $_
        # Fall through to exponential backoff.
    }
    return [math]::Min(30, [math]::Pow(2, $Attempt - 1))
}

# A Retry-After longer than this is not waited out inline: a recipe that would
# sleep for many minutes is better failed and retried by the next run.
$script:MaxRetryAfterSeconds = 120

function Invoke-YardstickGraphRequest {
    <#
    .SYNOPSIS
    Invokes Microsoft Graph with token refresh, paging, and transient retries.

    .DESCRIPTION
    Transient failures (no status, 408, 429, 5xx) are retried with backoff, up to
    -MaxAttempts per request (per page when paging). Every retry is logged with the
    status, the time the attempt took, and the request-id and client-request-id
    Microsoft support asks for: the first retry at Detail, later ones and giving
    up as warnings. Before 2026-11-23 retries were silent, and a degraded Intune
    hid 58 minutes of retried 504s inside "slow" lookups.

    .PARAMETER NoRetryOnServerError
    Retry only 429 (Graph refused the request without acting on it). Use for
    requests that create something, where a 5xx or a timeout may arrive after
    Graph did the work: resending would create a duplicate.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Resource,

        [ValidateSet('Get', 'Post', 'Patch', 'Put', 'Delete')]
        [string]$Method = 'Get',

        $Body,

        [ValidateSet('beta', 'v1.0')]
        [string]$ApiVersion = 'beta',

        [ValidateRange(1, 10)]
        [int]$MaxAttempts = 6,

        [switch]$NoRetryOnServerError
    )

    $stats = Get-YardstickGraphStatistics
    $uri = if ($Resource -match '^https://graph\.microsoft\.com/') {
        $Resource
    } else {
        "https://graph.microsoft.com/$ApiVersion/$($Resource.TrimStart('/'))"
    }

    $results = [System.Collections.Generic.List[object]]::new()
    do {
        $attempt = 0
        $refreshedOn401 = $false
        $requestRetried = $false
        do {
            $attempt++
            # Checked on every attempt, not once per call: retries and paging can
            # outlast a token that was nearly expired when the call began.
            $expiresOn = Get-YardstickGraphTokenExpiry
            $authenticationHeader = Get-YardstickGlobalValue -Name AuthenticationHeader
            if (-not $authenticationHeader -or -not $authenticationHeader.Authorization -or
                ($expiresOn -and $expiresOn -le [datetimeoffset]::UtcNow.AddMinutes(5))) {
                Connect-YardstickGraph -Force
                $authenticationHeader = Get-YardstickGlobalValue -Name AuthenticationHeader
            }
            $clientRequestId = [guid]::NewGuid().ToString()
            $headers = @{
                Authorization       = $authenticationHeader.Authorization
                Accept              = 'application/json'
                'client-request-id' = $clientRequestId
            }
            $stopwatch = [Diagnostics.Stopwatch]::StartNew()
            $stats.Requests++
            try {
                $params = @{
                    Uri         = $uri
                    Headers     = $headers
                    Method      = $Method
                    ErrorAction = 'Stop'
                    # Module session state does not see the runner's global
                    # PSDefaultParameterValues, so bound this here. Graph's gateway
                    # gives up on a slow backend at about 30 s, so a call that has
                    # not answered in 60 s is treated as transient and retried.
                    TimeoutSec  = 60
                }
                if ($null -ne $Body) {
                    $params.ContentType = 'application/json'
                    $params.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 20 -Compress }
                }
                $response = Invoke-RestMethod @params
                break
            } catch {
                $elapsed = $stopwatch.Elapsed.TotalSeconds
                $status = Get-YardstickGraphErrorStatusCode -ErrorRecord $_
                # A DELETE whose earlier attempt timed out may have gone through.
                if ($Method -eq 'Delete' -and $status -eq 404 -and $requestRetried) {
                    Write-YardstickGraphLog "Graph DELETE $(Format-YardstickGraphResourceForLog $uri) returned 404 on attempt $attempt; an earlier attempt already removed it." -Level Detail
                    return $null
                }
                # An expired or revoked token is fixed by one forced refresh.
                if ($status -eq 401 -and -not $refreshedOn401) {
                    $refreshedOn401 = $true
                    Write-YardstickGraphLog "Graph $($Method.ToUpper()) $(Format-YardstickGraphResourceForLog $uri) returned 401. Refreshing the token and trying again." -Level Detail
                    Connect-YardstickGraph -Force
                    $attempt--
                    continue
                }
                $transient = if ($NoRetryOnServerError) {
                    $status -eq 429
                } else {
                    ($null -eq $status) -or ($status -in 408, 429) -or ($status -ge 500 -and $status -le 599)
                }
                $statusText = if ($null -ne $status) { "$status" } else { $_.Exception.GetType().Name }
                $requestId = Get-YardstickGraphResponseHeader -ErrorRecord $_ -Name 'request-id'
                $ids = "request-id $(if ($requestId) { $requestId } else { 'none' }), client-request-id $clientRequestId"
                $what = "Graph $($Method.ToUpper()) $(Format-YardstickGraphResourceForLog $uri) returned $statusText after $([math]::Round($elapsed, 1)) s"
                if (-not $transient) { throw }
                if ($attempt -ge $MaxAttempts) {
                    $stats.GaveUp++
                    Write-YardstickGraphLog "$what (attempt $attempt of $MaxAttempts, $ids). Giving up." -Level Warning
                    throw
                }
                $delay = Get-YardstickGraphRetryDelay -ErrorRecord $_ -Attempt $attempt
                $stats.Retried++
                $stats.RetrySleepSecs += $delay
                $requestRetried = $true
                $level = if ($attempt -eq 1) { 'Detail' } else { 'Warning' }
                Write-YardstickGraphLog "$what (attempt $attempt of $MaxAttempts, $ids). Retrying in $delay s." -Level $level
                Start-Sleep -Seconds $delay
            }
        } while ($true)

        if ($null -ne $response -and $response.PSObject.Properties['value']) {
            foreach ($item in @($response.value)) { $results.Add($item) }
            $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
            $uri = if ($nextLinkProperty) { $nextLinkProperty.Value } else { $null }
            $Method = 'Get'
            $Body = $null
        } else {
            return $response
        }
    } while ($uri)

    return $results.ToArray()
}

Export-ModuleMember -Function Connect-YardstickGraph, Disconnect-YardstickGraph, Invoke-YardstickGraphRequest, Get-YardstickGraphStatistics, Reset-YardstickGraphStatistics
