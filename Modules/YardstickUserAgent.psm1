<#
The browser user agent Yardstick presents to vendor sites that refuse obvious
scripts. Some sites reject an outdated browser as firmly as they reject a
script: [Microsoft.PowerShell.Commands.PSUserAgent]::Chrome still claims
"Chrome/7.0", which Dropbox answers with 409/400. So the Chrome version here is
read from Google's own release feed rather than written down once and left to
age.

Callers outside the runner (Get-YardstickIcon.ps1) load this directly, so it
imports nothing and logs nothing - a failed lookup is not worth a log line, and
the fallback is good enough to send.
#>

# The format Chrome itself sends since its user-agent reduction: only the major
# version is real, and the Windows platform token is frozen at "10.0; Win64".
$script:Template = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{0}.0.0.0 Safari/537.36'
$script:VersionFeed = 'https://versionhistory.googleapis.com/v1/chrome/platforms/win64/channels/stable/versions?pageSize=1'

# Used only when the feed is unreachable. 157 was stable on 2026-10-23. Chrome
# has shipped faster than every four weeks since (150 to 157 took about 18 days
# a release), so assuming four weeks can run a little behind but never forward
# into a version that does not exist - and a little behind is still a current
# browser to any site that checks.
$script:KnownMajor = 157
$script:KnownDate = [datetime]'2026-10-23'
$script:ReleaseDays = 28

$script:Cached = $null


function Get-YardstickChromeMajorVersion {
    <#
    .SYNOPSIS
    Returns the current stable Chrome major version.

    .DESCRIPTION
    Reads Google's Chrome VersionHistory feed once per session and caches the
    answer, including a failed lookup - so a run with no route to Google pays
    the timeout once, not once per recipe. When the feed cannot be read, the
    version is estimated from the last one verified.

    .PARAMETER Refresh
    Ignore the cached answer and read the feed again.

    .PARAMETER TimeoutSec
    How long to wait for the feed before falling back to the estimate.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [switch] $Refresh,
        [ValidateRange(1, 120)]
        [int] $TimeoutSec = 10
    )

    if ($script:Cached -and -not $Refresh) { return $script:Cached }

    $major = $null
    try {
        $response = Invoke-RestMethod -Uri $script:VersionFeed -TimeoutSec $TimeoutSec -ErrorAction Stop
        $version = [string]@($response.versions)[0].version
        if ($version -match '^(\d+)\.') { $major = [int]$Matches[1] }
    } catch {
        Write-Verbose "Chrome version feed unavailable ($($_.Exception.Message)); estimating instead."
    }
    if (-not $major -or $major -lt $script:KnownMajor) {
        $major = Get-YardstickEstimatedChromeMajorVersion
    }

    $script:Cached = $major
    $major
}


function Get-YardstickEstimatedChromeMajorVersion {
    <#
    .SYNOPSIS
    Estimates the stable Chrome major version from the last verified release.

    .PARAMETER Date
    The date to estimate for. Defaults to today.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [datetime] $Date = (Get-Date)
    )

    $elapsed = [math]::Max(0, ($Date.Date - $script:KnownDate).TotalDays)
    $script:KnownMajor + [int][math]::Floor($elapsed / $script:ReleaseDays)
}


function Get-YardstickChromeUserAgent {
    <#
    .SYNOPSIS
    Returns a user agent string for the current stable Chrome on Windows.

    .DESCRIPTION
    Use this wherever a vendor site needs to see a browser. Sites that only
    want to identify the caller are better served by a plain 'Yardstick'
    agent; for the GitHub API, use Invoke-YardstickGitHubApi, which sets one.

    .PARAMETER Refresh
    Re-read the Chrome version instead of using this session's cached answer.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [switch] $Refresh
    )

    $script:Template -f (Get-YardstickChromeMajorVersion -Refresh:$Refresh)
}


Export-ModuleMember -Function Get-YardstickChromeMajorVersion, Get-YardstickEstimatedChromeMajorVersion, Get-YardstickChromeUserAgent
