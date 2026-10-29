<#
.SYNOPSIS
GitHub REST API access for Yardstick and its recipes.

.DESCRIPTION
Every call Yardstick makes to api.github.com goes through Invoke-YardstickGitHubApi,
so the optional token stored with Set-YardstickCredential.ps1 -GitHub is attached
in one place. Anonymous calls are limited to 60 an hour per IP address, which a
single full run nearly uses up; a token raises that to 5,000. Files read from
raw.githubusercontent.com go through Invoke-YardstickGitHubRaw, which shares the
same throttling rules but never sends the token.

The token goes to api.github.com and nowhere else. It is deliberately not a
$PSDefaultParameterValues entry, which would hand it to every vendor a recipe calls.

GitHub's API terms and best practices shape the error handling:
- When GitHub throttles a call, Yardstick waits as long as GitHub asks before it
  calls GitHub again: the Retry-After header if there is one, the rate-limit
  reset time if the limit is used up, and otherwise a minute, doubling each time
  GitHub throttles again within 15 minutes of the last wait ending. GitHub warns
  that carrying on while limited can get an integration banned.
- The wait is a single hold on every GitHub host, kept for the whole session, so
  it applies to every recipe and not just the one that was throttled. One hold
  covers them all because api.github.com and raw.githubusercontent.com front the
  same busy backend: a throttled read of a winget-pkgs manifest is followed by
  throttled API calls for the same repository. Calls made during a hold fail here
  without a request.
- In a Yardstick run (Reset-YardstickGitHubSession -DeferThrottledCalls) the
  helper never sleeps: it throws, and the runner sets the recipe aside, carries on
  with the others and replays it once the hold has passed. Elsewhere, such as the
  recipe harness, a hold of up to five minutes is waited out in place.
- A rejected token is dropped for the session and the call retried anonymously,
  so an expired token degrades to the old behaviour instead of failing every
  GitHub recipe.
#>

$Script:GitHubApiHost = 'api.github.com'
$Script:GitHubRawHost = 'raw.githubusercontent.com'
$Script:GitHubUserAgent = 'Yardstick (+https://github.com/jeffpearson1/Yardstick)'
$Script:MinThrottleWaitSeconds = 60
$Script:MaxThrottleWaitSeconds = 1800
$Script:MaxInlineWaitSeconds = 300
$Script:MaxInlineWaits = 3
# A throttle this long after the last wait ended starts the doubling wait over.
$Script:StrikeResetMinutes = 15
$Script:TokenExpiryWarningDays = 14

# Hosts whose throttling responses are honoured when a recipe calls them directly
# and the runner hands the error to Get-YardstickGitHubThrottle.
$Script:GitHubHostPattern = '^(api\.github\.com|github\.com|codeload\.github\.com|([a-z0-9-]+\.)*githubusercontent\.com)$'

# Session state. Reset-YardstickGitHubSession puts these back.
$Script:TokenLoaded = $false
$Script:Token = $null
$Script:ExpiryWarned = $false
$Script:DeferThrottledCalls = $false
# The time GitHub asked Yardstick to wait until, for every GitHub host.
$Script:HoldUntil = $null
# When the latest hold ended or ends, kept after it passes, so a throttle soon
# afterwards keeps doubling the wait.
$Script:LastHoldUntil = $null
# Throttles since the doubling wait last started over.
$Script:ThrottleStrikes = 0
# The last throttle or hold that stopped a call; Clear-YardstickGitHubThrottle
# forgets it.
$Script:LastThrottle = $null


# No -Force, so already-loaded copies are reused.
Import-Module (Join-Path $PSScriptRoot 'YardstickLog.psm1')
Import-Module (Join-Path $PSScriptRoot 'YardstickCredential.psm1')


function Reset-YardstickGitHubSession {
    <#
    .SYNOPSIS
    Forgets the cached token and any hold, so the next call starts fresh.

    .DESCRIPTION
    Use after storing or removing the token in the same PowerShell session. The
    runner calls it at the start of every run.

    .PARAMETER DeferThrottledCalls
    Never wait for GitHub inside a call: throw at once, so the caller can do other
    work and retry once Get-YardstickGitHubHold has passed. Without it, a hold of
    up to five minutes is waited out in place.
    #>
    [CmdletBinding()]
    param(
        [switch]$DeferThrottledCalls
    )

    $Script:TokenLoaded = $false
    $Script:Token = $null
    $Script:ExpiryWarned = $false
    $Script:DeferThrottledCalls = [bool]$DeferThrottledCalls
    $Script:HoldUntil = $null
    $Script:LastHoldUntil = $null
    $Script:ThrottleStrikes = 0
    $Script:LastThrottle = $null
}


function Get-GitHubSessionToken {
    # Read once per session: Credential Manager is cheap, but there is no reason
    # to ask it on every one of a run's sixty-odd calls.
    if (-not $Script:TokenLoaded) {
        $Script:Token = Get-YardstickGitHubToken
        $Script:TokenLoaded = $true
    }
    return $Script:Token
}


function Get-GitHubHeaderValue {
    param(
        $Headers,
        [string]$Name
    )

    if ($null -eq $Headers) { return $null }

    # Invoke-RestMethod -ResponseHeadersVariable gives a dictionary of string arrays;
    # an HttpResponseException carries HttpResponseHeaders. Handle both.
    if ($Headers -is [System.Net.Http.Headers.HttpHeaders]) {
        $values = $null
        if ($Headers.TryGetValues($Name, [ref]$values)) { return @($values)[0] }
        return $null
    }
    foreach ($key in $Headers.Keys) {
        if ($key -eq $Name) { return @($Headers[$key])[0] }
    }
    return $null
}


function ConvertFrom-GitHubResetTime {
    param([string]$EpochSeconds)

    $seconds = 0L
    if ([long]::TryParse($EpochSeconds, [ref]$seconds) -and $seconds -gt 0) {
        return [DateTimeOffset]::FromUnixTimeSeconds($seconds).LocalDateTime
    }
    return $null
}


function Get-GitHubErrorMessage {
    param([string]$Body)

    if ([string]::IsNullOrWhiteSpace($Body)) { return $null }

    # The API answers with JSON ({"message": ..., "documentation_url": ...});
    # raw.githubusercontent.com with plain text.
    $message = $Body
    try {
        $parsed = $Body | ConvertFrom-Json -ErrorAction Stop
        if ($parsed.PSObject.Properties['message'] -and $parsed.message) { $message = [string]$parsed.message }
    } catch {
        $message = $Body
    }
    $message = ($message -replace '\s+', ' ').Trim()
    if ($message.Length -gt 300) { $message = $message.Substring(0, 300) + '...' }
    return $message
}


function Test-GitHubTokenExpiry {
    param($Headers)

    if ($Script:ExpiryWarned) { return }

    # Fine-grained tokens report their expiry on every response, for example
    # "2026-11-30 12:00:00 UTC".
    $expiration = Get-GitHubHeaderValue -Headers $Headers -Name 'github-authentication-token-expiration'
    if ([string]::IsNullOrWhiteSpace($expiration)) { return }

    $parsed = [DateTimeOffset]::MinValue
    $normalized = ($expiration -replace '\s+UTC$', ' +00:00').Trim()
    if (-not [DateTimeOffset]::TryParse($normalized, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) {
        return
    }

    $daysLeft = ($parsed.LocalDateTime.Date - (Get-Date).Date).Days
    if ($daysLeft -le $Script:TokenExpiryWarningDays) {
        $Script:ExpiryWarned = $true
        Write-Log "The GitHub API token expires on $($parsed.LocalDateTime.ToString('yyyy-MM-dd')) ($daysLeft days). Regenerate it on GitHub and store it with .\Set-YardstickCredential.ps1 -GitHub." -Level Warning
    }
}


function Register-GitHubThrottle {
    <#
    Records a throttling response as a hold on every GitHub host, and returns the
    throttle, or $null when the response is not GitHub asking Yardstick to slow
    down.
    #>
    param(
        [string]$HostName,
        [string]$Uri,
        [int]$Status,
        $Headers,
        [string]$Body,
        [bool]$Authenticated
    )

    if ($Status -ne 403 -and $Status -ne 429) { return $null }

    $remaining = Get-GitHubHeaderValue -Headers $Headers -Name 'x-ratelimit-remaining'
    # Not "-as [int]" on the raw value: $null -as [int] is 0, which would read a
    # missing header as "retry now".
    $retryAfterHeader = Get-GitHubHeaderValue -Headers $Headers -Name 'retry-after'
    $retryAfter = if ([string]::IsNullOrWhiteSpace($retryAfterHeader)) { $null } else { $retryAfterHeader -as [int] }
    $gitHubMessage = Get-GitHubErrorMessage -Body $Body
    $isSecondaryLimit = $Status -eq 429 -or "$gitHubMessage" -match 'rate limit'

    # A 403 with none of these is a real refusal (a private repository, say), not
    # a request to wait.
    if ($remaining -ne '0' -and $null -eq $retryAfter -and -not $isSecondaryLimit) { return $null }

    $now = Get-Date
    # A call that gets through after a wait does not mean GitHub has recovered, so
    # only a quiet spell starts the doubling over.
    if ($Script:LastHoldUntil -and $now -gt $Script:LastHoldUntil.AddMinutes($Script:StrikeResetMinutes)) {
        $Script:ThrottleStrikes = 0
    }
    $advice = ''
    if ($null -ne $retryAfter) {
        $retryAt = $now.AddSeconds([math]::Max($retryAfter, 1))
        $Script:ThrottleStrikes++
    } elseif ($remaining -eq '0') {
        $retryAt = ConvertFrom-GitHubResetTime (Get-GitHubHeaderValue -Headers $Headers -Name 'x-ratelimit-reset')
        if (-not $retryAt -or $retryAt -le $now) { $retryAt = $now.AddMinutes(1) }
        if ($HostName -eq $Script:GitHubApiHost -and -not $Authenticated) {
            $advice = ' Anonymous calls are limited to 60 an hour; store a token with .\Set-YardstickCredential.ps1 -GitHub to raise that to 5,000.'
        }
    } else {
        # GitHub's guidance for a secondary limit without Retry-After: wait at least
        # a minute, and exponentially longer if it keeps happening.
        $Script:ThrottleStrikes++
        $waitSeconds = [math]::Min($Script:MinThrottleWaitSeconds * [math]::Pow(2, $Script:ThrottleStrikes - 1), $Script:MaxThrottleWaitSeconds)
        $retryAt = $now.AddSeconds($waitSeconds)
    }

    # Never shorten a hold another response asked for.
    $existing = Get-YardstickGitHubHold
    if ($existing -and $existing -gt $retryAt) { $retryAt = $existing }
    $Script:HoldUntil = $retryAt
    $Script:LastHoldUntil = $retryAt

    $requestId = Get-GitHubHeaderValue -Headers $Headers -Name 'x-github-request-id'
    $details = @()
    if ($gitHubMessage) { $details += "GitHub said: $gitHubMessage" }
    if ($null -ne $retryAfter) { $details += "Retry-After: $retryAfter s" }
    if ($remaining -eq '0') { $details += 'rate limit used up' }
    if ($requestId) { $details += "request ID $requestId" }
    $detailText = if ($details) { " ($($details -join '; '))" } else { '' }
    Write-Log "GitHub throttled $Uri with HTTP $Status$detailText. No more calls to GitHub until $($retryAt.ToString('HH:mm:ss')).$advice" -Level Warning

    $Script:LastThrottle = [PSCustomObject]@{
        Host      = $HostName
        Uri       = $Uri
        Status    = $Status
        Message   = $gitHubMessage
        RequestId = $requestId
        RetryAt   = $retryAt
    }
    return $Script:LastThrottle
}


function New-GitHubThrottleError {
    param(
        [string]$Message,
        $Throttle
    )

    $exception = [System.Exception]::new($Message)
    return [System.Management.Automation.ErrorRecord]::new(
        $exception, 'YardstickGitHubThrottled', [System.Management.Automation.ErrorCategory]::LimitsExceeded, $Throttle)
}


function Invoke-GitHubRequest {
    param(
        [string]$RequestUri,
        [string]$HostName,
        [hashtable]$Headers,
        [switch]$UseToken
    )

    # Module code does not see the global -TimeoutSec default that
    # Set-YardstickWebRequestTimeout applies to recipes, so follow it explicitly.
    $timeoutSec = $Global:PSDefaultParameterValues['Invoke-RestMethod:TimeoutSec'] -as [int]
    if ($timeoutSec -le 0) { $timeoutSec = 120 }

    $inlineWaits = 0
    $throttledHere = $false
    while ($true) {
        $holdUntil = Get-YardstickGitHubHold
        if ($holdUntil) {
            $waitSeconds = [math]::Ceiling(($holdUntil - (Get-Date)).TotalSeconds)
            if ($Script:DeferThrottledCalls -or $inlineWaits -ge $Script:MaxInlineWaits -or $waitSeconds -gt $Script:MaxInlineWaitSeconds) {
                if (-not $Script:LastThrottle) {
                    $Script:LastThrottle = [PSCustomObject]@{
                        Host = $HostName; Uri = $RequestUri; Status = $null; Message = $null; RequestId = $null; RetryAt = $holdUntil
                    }
                }
                $Script:LastThrottle.RetryAt = $holdUntil
                $what = if ($throttledHere) { "GitHub throttled $RequestUri" } else { "Skipped $RequestUri" }
                throw (New-GitHubThrottleError -Throttle $Script:LastThrottle -Message "$what : GitHub asked Yardstick to wait until $($holdUntil.ToString('HH:mm:ss')) before calling GitHub again.")
            }
            Write-Log "Waiting $waitSeconds seconds, until $($holdUntil.ToString('HH:mm:ss')), as GitHub asked, before calling $RequestUri." -Level Warning
            Start-Sleep -Seconds $waitSeconds
            # Waited out, so lift it; a fresh throttle sets a new one.
            $Script:HoldUntil = $null
            $inlineWaits++
            continue
        }

        $requestHeaders = @{ 'User-Agent' = $Script:GitHubUserAgent }
        foreach ($name in $Headers.Keys) { $requestHeaders[$name] = $Headers[$name] }
        $token = if ($UseToken) { Get-GitHubSessionToken } else { $null }
        if ($token) { $requestHeaders['Authorization'] = "Bearer $token" }

        $responseHeaders = $null
        try {
            $restParams = @{
                Uri                     = $RequestUri
                Headers                 = $requestHeaders
                TimeoutSec              = $timeoutSec
                ResponseHeadersVariable = 'responseHeaders'
                ErrorAction             = 'Stop'
            }
            $result = Invoke-RestMethod @restParams
            if ($token) { Test-GitHubTokenExpiry -Headers $responseHeaders }
            return $result
        } catch {
            $response = $_.Exception.Response
            if ($null -eq $response) { throw }
            $status = [int]$response.StatusCode

            if ($status -eq 401 -and $token) {
                Write-Log "GitHub rejected the stored API token (401), so this session continues anonymously. Store a new token with .\Set-YardstickCredential.ps1 -GitHub." -Level Warning
                $Script:Token = $null
                continue
            }

            $throttleParams = @{
                HostName      = $HostName
                Uri           = $RequestUri
                Status        = $status
                Headers       = $response.Headers
                Body          = "$($_.ErrorDetails.Message)"
                Authenticated = [bool]$token
            }
            if (-not (Register-GitHubThrottle @throttleParams)) { throw }
            $throttledHere = $true
            # Back to the top, where the new hold is either waited out or thrown.
        }
    }
}


function Invoke-YardstickGitHubApi {
    <#
    .SYNOPSIS
    Sends a GET request to the GitHub REST API and returns the parsed response.

    .DESCRIPTION
    Use this for every api.github.com call, in Yardstick code and in recipes. It
    identifies Yardstick to GitHub, attaches the stored token when there is one,
    and waits as long as GitHub asks when it throttles a call. The result is what
    Invoke-RestMethod would have returned.

    Read files from raw.githubusercontent.com with Invoke-YardstickGitHubRaw.
    Release asset downloads are not API calls and do not need either.

    .PARAMETER Uri
    A full https://api.github.com/... URL, or a path relative to it such as
    'repos/owner/name/releases/latest'. Any other host is refused, so the token
    cannot be sent anywhere else.

    .EXAMPLE
    $release = Invoke-YardstickGitHubApi 'repos/brave/brave-browser/releases/latest'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri
    )

    if ($Uri -match '^[a-z][a-z0-9+.-]*://') {
        $parsedUri = [Uri]$Uri
        if ($parsedUri.Scheme -ne 'https' -or $parsedUri.Host -ne $Script:GitHubApiHost) {
            throw "Invoke-YardstickGitHubApi only calls https://$($Script:GitHubApiHost); '$Uri' is not a GitHub API URL."
        }
        $requestUri = $parsedUri.AbsoluteUri
    } else {
        $requestUri = "https://$($Script:GitHubApiHost)/$($Uri.TrimStart('/'))"
    }

    $headers = @{
        'Accept'               = 'application/vnd.github+json'
        'X-GitHub-Api-Version' = '2022-11-28'
    }
    Invoke-GitHubRequest -RequestUri $requestUri -HostName $Script:GitHubApiHost -Headers $headers -UseToken
}


function Invoke-YardstickGitHubRaw {
    <#
    .SYNOPSIS
    Reads a file from raw.githubusercontent.com, waiting as long as GitHub asks
    when it throttles.

    .DESCRIPTION
    raw.githubusercontent.com is not part of the API rate limit, but it does
    throttle, and the winget-pkgs manifests several recipes read are popular enough
    that it often does. Going through this keeps every recipe off the host while
    GitHub has asked Yardstick to wait. The token is never sent here. The result is
    what Invoke-RestMethod would have returned: usually the file's text.

    .PARAMETER Uri
    A full https://raw.githubusercontent.com/... URL. Any other host is refused.

    .EXAMPLE
    $manifest = Invoke-YardstickGitHubRaw "https://raw.githubusercontent.com/microsoft/winget-pkgs/master/manifests/a/Amazon/Kindle/$version/Amazon.Kindle.installer.yaml" | ConvertFrom-Yaml
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri
    )

    $parsedUri = $null
    if (-not [Uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$parsedUri) -or
        $parsedUri.Scheme -ne 'https' -or $parsedUri.Host -ne $Script:GitHubRawHost) {
        throw "Invoke-YardstickGitHubRaw only reads https://$($Script:GitHubRawHost); '$Uri' is not one of its URLs."
    }

    Invoke-GitHubRequest -RequestUri $parsedUri.AbsoluteUri -HostName $Script:GitHubRawHost -Headers @{}
}


function Get-YardstickGitHubHold {
    <#
    .SYNOPSIS
    Returns the time GitHub asked Yardstick to wait until, or nothing if it has not.

    .DESCRIPTION
    One hold covers every GitHub host: api.github.com, raw.githubusercontent.com,
    and the release download hosts when a recipe calls them directly.

    .OUTPUTS
    DateTime (local time), or nothing when there is no hold in force.
    #>
    [CmdletBinding()]
    [OutputType([datetime])]
    param()

    if ($Script:HoldUntil -and $Script:HoldUntil -le (Get-Date)) { $Script:HoldUntil = $null }
    return $Script:HoldUntil
}


function Get-YardstickGitHubThrottle {
    <#
    .SYNOPSIS
    Reports whether GitHub throttling stopped a call since the last
    Clear-YardstickGitHubThrottle, and until when to wait.

    .DESCRIPTION
    The runner calls this when a recipe fails, to tell "GitHub asked us to wait"
    apart from a real failure. It reads what the helpers recorded rather than the
    error, because recipes often catch an error and throw their own.

    Given -ErrorRecord, it also recognises a throttling response from a GitHub host
    that a recipe called directly, and records the hold as the helpers would have.

    .PARAMETER ErrorRecord
    The error the recipe failed with.

    .OUTPUTS
    PSCustomObject with Host, Uri, Status, Message, RequestId and RetryAt, or
    nothing.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        $ErrorRecord
    )

    if ($Script:LastThrottle) {
        # The hold may have grown since; report the time it is actually safe to retry.
        $hold = Get-YardstickGitHubHold
        if ($hold -and $hold -gt $Script:LastThrottle.RetryAt) { $Script:LastThrottle.RetryAt = $hold }
        return $Script:LastThrottle
    }

    if ($null -eq $ErrorRecord -or -not $ErrorRecord.PSObject.Properties['Exception']) { return $null }

    $exception = $ErrorRecord.Exception
    $depth = 0
    while ($exception -and $depth -lt 16) {
        $responseProperty = $exception.PSObject.Properties['Response']
        if ($responseProperty -and $responseProperty.Value) { break }
        $exception = $exception.InnerException
        $depth++
    }
    if (-not $exception) { return $null }

    $response = $exception.Response
    $requestUri = if ($response.PSObject.Properties['RequestMessage'] -and $response.RequestMessage) {
        $response.RequestMessage.RequestUri
    } elseif ($response.PSObject.Properties['ResponseUri']) {
        $response.ResponseUri
    }
    if (-not $requestUri -or $requestUri.Host -notmatch $Script:GitHubHostPattern) { return $null }

    $throttleParams = @{
        HostName      = $requestUri.Host
        Uri           = $requestUri.AbsoluteUri
        Status        = [int]$response.StatusCode
        Headers       = $response.Headers
        Body          = "$($ErrorRecord.ErrorDetails.Message)"
        Authenticated = $false
    }
    return Register-GitHubThrottle @throttleParams
}


function Clear-YardstickGitHubThrottle {
    <#
    .SYNOPSIS
    Forgets which throttle last stopped a call, keeping the holds themselves.

    .DESCRIPTION
    The runner calls this as each recipe starts, so Get-YardstickGitHubThrottle
    answers for that recipe alone.
    #>
    [CmdletBinding()]
    param()

    $Script:LastThrottle = $null
}


function Get-YardstickGitHubRateLimit {
    <#
    .SYNOPSIS
    Reports this session's GitHub API rate limit and whether it is authenticated.

    .DESCRIPTION
    GET /rate_limit does not count against the limit.

    .OUTPUTS
    PSCustomObject with Authenticated, Limit, Remaining and Reset (local time).
    #>
    [CmdletBinding()]
    param()

    $response = Invoke-YardstickGitHubApi 'rate_limit'
    $core = if ($response.resources.core) { $response.resources.core } else { $response.rate }

    [PSCustomObject]@{
        Authenticated = [bool](Get-GitHubSessionToken)
        Limit         = [int]$core.limit
        Remaining     = [int]$core.remaining
        Reset         = ConvertFrom-GitHubResetTime ([string]$core.reset)
    }
}


Export-ModuleMember -Function @(
    'Invoke-YardstickGitHubApi',
    'Invoke-YardstickGitHubRaw',
    'Get-YardstickGitHubRateLimit',
    'Get-YardstickGitHubHold',
    'Get-YardstickGitHubThrottle',
    'Clear-YardstickGitHubThrottle',
    'Reset-YardstickGitHubSession'
)
