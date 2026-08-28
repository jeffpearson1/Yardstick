Set-StrictMode -Version 3.0

$script:AccessToken = $null

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

    if ([string]::IsNullOrWhiteSpace($TenantID) -or
        [string]::IsNullOrWhiteSpace($ClientID) -or
        [string]::IsNullOrWhiteSpace($ClientSecret)) {
        throw 'Required Graph credentials are not set: TenantID, ClientID, and ClientSecret are required.'
    }

    $expiresOn = Get-YardstickGraphTokenExpiry
    $authenticationHeader = Get-YardstickGlobalValue -Name AuthenticationHeader
    if (-not $Force -and $authenticationHeader -and $authenticationHeader.Authorization -and
        $expiresOn -and $expiresOn -gt [datetimeoffset]::UtcNow.AddMinutes(30)) {
        if ($PassThru) { return $script:AccessToken }
        return
    }

    $tokenUri = "https://login.microsoftonline.com/$([uri]::EscapeDataString($TenantID))/oauth2/v2.0/token"
    $tokenResponse = Invoke-RestMethod -Uri $tokenUri -Method Post -ContentType 'application/x-www-form-urlencoded' -Body @{
        client_id     = $ClientID
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    } -ErrorAction Stop

    if (-not $tokenResponse.access_token) {
        throw 'Microsoft identity platform returned no access token.'
    }

    $lifetimeSeconds = if ($tokenResponse.expires_in) { [int]$tokenResponse.expires_in } else { 3600 }
    $expiresOn = [datetimeoffset]::UtcNow.AddSeconds($lifetimeSeconds)
    $script:AccessToken = [pscustomobject]@{
        AccessToken = [string]$tokenResponse.access_token
        TokenType   = if ($tokenResponse.token_type) { [string]$tokenResponse.token_type } else { 'Bearer' }
        ExpiresOn   = $expiresOn
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
        if ($retryAfter.Delta) {
            return [math]::Max(1, [int][math]::Ceiling($retryAfter.Delta.TotalSeconds))
        }
        if ($retryAfter.Date) {
            return [math]::Max(1, [int][math]::Ceiling(($retryAfter.Date - [datetimeoffset]::UtcNow).TotalSeconds))
        }
        $raw = $ErrorRecord.Exception.Response.Headers['Retry-After']
        if ($raw -as [int]) { return [math]::Max(1, [int]$raw) }
    } catch {
        $null = $_
        # Fall through to exponential backoff.
    }
    return [math]::Min(30, [math]::Pow(2, $Attempt - 1))
}

function Invoke-YardstickGraphRequest {
    <#
    .SYNOPSIS
    Invokes Microsoft Graph with token refresh, paging, and transient retries.
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
        [int]$MaxAttempts = 6
    )

    $expiresOn = Get-YardstickGraphTokenExpiry
    $authenticationHeader = Get-YardstickGlobalValue -Name AuthenticationHeader
    if (-not $authenticationHeader -or -not $authenticationHeader.Authorization -or
        ($expiresOn -and $expiresOn -le [datetimeoffset]::UtcNow.AddMinutes(5))) {
        Connect-YardstickGraph -Force
        $authenticationHeader = Get-YardstickGlobalValue -Name AuthenticationHeader
    }

    $headers = @{
        Authorization = $authenticationHeader.Authorization
        Accept        = 'application/json'
    }
    $uri = if ($Resource -match '^https://graph\.microsoft\.com/') {
        $Resource
    } else {
        "https://graph.microsoft.com/$ApiVersion/$($Resource.TrimStart('/'))"
    }

    $results = [System.Collections.Generic.List[object]]::new()
    do {
        $attempt = 0
        do {
            $attempt++
            try {
                $params = @{
                    Uri         = $uri
                    Headers     = $headers
                    Method      = $Method
                    ErrorAction = 'Stop'
                }
                if ($null -ne $Body) {
                    $params.ContentType = 'application/json'
                    $params.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 20 -Compress }
                }
                $response = Invoke-RestMethod @params
                break
            } catch {
                $status = Get-YardstickGraphErrorStatusCode -ErrorRecord $_
                $transient = ($null -eq $status) -or ($status -in 408, 429) -or ($status -ge 500 -and $status -le 599)
                if (-not $transient -or $attempt -ge $MaxAttempts) { throw }
                Start-Sleep -Seconds (Get-YardstickGraphRetryDelay -ErrorRecord $_ -Attempt $attempt)
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

Export-ModuleMember -Function Connect-YardstickGraph, Invoke-YardstickGraphRequest
