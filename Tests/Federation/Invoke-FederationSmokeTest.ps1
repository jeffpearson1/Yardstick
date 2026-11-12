<#
.SYNOPSIS
Runs the real Yardstick.ps1 startup in workload identity federation mode against
local stand-ins for the ID token sources, Entra ID and Graph.

.DESCRIPTION
Proves, end to end and without any tenant, that an unattended run
(-All -NoInteractive -NoEmail) of the current working tree, for each ID token
source (Azure DevOps, GitHub Actions and a token file):

  - gets an ID token from that source; Azure DevOps and GitHub Actions are
    real HTTP endpoints on localhost
  - exchanges it with Entra as client_assertion, never sending client_secret
  - never reads Windows Credential Manager (a unique, empty credential target is
    configured, so any read would make the -NoInteractive run fail)
  - never calls the app registration lookup behind the secret-expiry check
  - exits 0

The working tree (tracked and untracked, not ignored) is copied to a temporary
folder with its own Local\preferences.yaml, so the checkout is not touched. The
run uses an empty recipe folder: no app is downloaded, packaged or uploaded.

Entra and Graph are faked by a global Invoke-RestMethod function in the child
process, which PowerShell resolves ahead of the cmdlet from inside modules too.

.PARAMETER Provider
The ID token sources to test. Defaults to all three.

.EXAMPLE
    pwsh -NoProfile -File Tests\Federation\Invoke-FederationSmokeTest.ps1
#>
[CmdletBinding()]
param(
    [ValidateSet('AzureDevOps', 'GitHubActions', 'File')]
    [string[]]$Provider = @('AzureDevOps', 'GitHubActions', 'File'),
    [switch]$KeepWorkspace
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$workspace = Join-Path ([IO.Path]::GetTempPath()) "yardstick-wif-smoke-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
$engine = Join-Path $workspace 'engine'
$failures = [System.Collections.Generic.List[string]]::new()

function Assert-Smoke {
    param([bool]$Condition, [string]$Message)
    if ($Condition) {
        Write-Host "  PASS  $Message" -ForegroundColor Green
    } else {
        Write-Host "  FAIL  $Message" -ForegroundColor Red
        $failures.Add($Message)
    }
}

# --- Copy the working tree -------------------------------------------------------
Write-Host "Copying working tree to $engine"
$files = @(& git -C $repoRoot ls-files --cached --others --exclude-standard)
if ($LASTEXITCODE -ne 0) { throw 'git ls-files failed.' }
foreach ($relative in $files) {
    if ($relative -like '.claude/*') { continue }
    $source = Join-Path $repoRoot $relative
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { continue }
    $destination = Join-Path $engine $relative
    New-Item -ItemType Directory -Path (Split-Path $destination) -Force | Out-Null
    Copy-Item -LiteralPath $source -Destination $destination
}

# --- Preferences: unattended, with an empty recipe folder ------------------------
$recipes = Join-Path $workspace 'Recipes'
foreach ($folder in $recipes, (Join-Path $workspace 'Icons'), (Join-Path $engine 'Local'),
    (Join-Path $engine 'Artifacts\Temp'), (Join-Path $engine 'Artifacts\BuildSpace'), (Join-Path $engine 'Artifacts\Packages')) {
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
}
$credentialTarget = "Yardstick:WifSmoke-$([guid]::NewGuid().ToString('N'))"
$preferences = (Get-Content -LiteralPath (Join-Path $engine 'Config\Preferences.example.yaml') -Raw).Replace('<REPLACE_WITH_YARDSTICK_ROOT>', $engine)
$preferences = $preferences -replace '(?m)^#?\s*Recipes:.*$', "Recipes: $recipes"
$preferences = $preferences -replace '(?m)^#?\s*Icons:.*$', "Icons: $(Join-Path $workspace 'Icons')"
# Without this the run falls back to the engine's own Recipes and downloads every real installer.
if ($preferences -notmatch "(?m)^Recipes: $([regex]::Escape($recipes))\s*$") {
    throw 'Could not point Recipes at the empty smoke-test folder; Config\Preferences.example.yaml has changed.'
}
$preferences = $preferences -replace '(?m)^credentialTarget:.*$', "credentialTarget: `"$credentialTarget`""
$preferences = $preferences -replace '(?m)^emailNotificationEnabled:.*$', 'emailNotificationEnabled: false'
Set-Content -LiteralPath (Join-Path $engine 'Local\preferences.yaml') -Value $preferences

# --- Fake ID token endpoints (real HTTP) -----------------------------------------
$probe = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
$probe.Start()
$port = ([System.Net.IPEndPoint]$probe.LocalEndpoint).Port
$probe.Stop()

$encode = { param($Text) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text)).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
$idToken = '{0}.{1}.smoke-signature' -f (& $encode '{"alg":"RS256"}'), (& $encode '{"iss":"https://issuer.smoke.invalid","sub":"smoke-subject","aud":"api://AzureADTokenExchange"}')
$jobToken = "smoke-job-token-$([guid]::NewGuid().ToString('N'))"
$serviceConnectionId = [guid]::NewGuid().ToString()
$clientId = [guid]::NewGuid().ToString()
$tokenFile = Join-Path $workspace 'id-token'
Set-Content -LiteralPath $tokenFile -Value $idToken -NoNewline

$tokenRequests = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()
$server = Start-ThreadJob -ScriptBlock {
    $Listener = $using:listener
    $Requests = $using:tokenRequests
    $IdToken = $using:idToken
    while ($Listener.IsListening) {
        try { $context = $Listener.GetContext() } catch { break }
        $Requests.Enqueue([pscustomobject]@{
            Method        = $context.Request.HttpMethod
            Path          = $context.Request.Url.AbsolutePath
            Query         = $context.Request.Url.Query
            Authorization = $context.Request.Headers['Authorization']
        })
        $payload = if ($context.Request.Url.AbsolutePath -like '/github/*') { @{ value = $IdToken } } else { @{ oidcToken = $IdToken } }
        $bytes = [Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Compress))
        $context.Response.ContentType = 'application/json'
        $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $context.Response.Close()
    }
}

# --- Child process bootstrap: fake Entra and Graph -------------------------------
$bootstrap = Join-Path $workspace 'bootstrap.ps1'
@'
param($Engine, $CallLog)
function global:Invoke-RestMethod {
    [CmdletBinding()]
    param($Uri, $Method = 'Get', $Headers, $Body, $ContentType, [switch]$UseBasicParsing, $TimeoutSec)
    $uriText = [string]$Uri
    if ($uriText -like 'https://login.microsoftonline.com/*') {
        $record = [ordered]@{ kind = 'entra'; uri = $uriText; keys = @($Body.Keys | Sort-Object)
            assertion = [string]$Body.client_assertion; assertionType = [string]$Body.client_assertion_type }
        $record | ConvertTo-Json -Compress | Add-Content -LiteralPath $CallLog
        return [pscustomobject]@{ token_type = 'Bearer'; expires_in = 3599; access_token = 'smoke-graph-token' }
    }
    if ($uriText -like 'https://graph.microsoft.com/*') {
        [ordered]@{ kind = 'graph'; uri = $uriText; method = [string]$Method; authorization = [string]$Headers.Authorization } |
            ConvertTo-Json -Compress | Add-Content -LiteralPath $CallLog
        return [pscustomobject]@{ value = @() }
    }
    $pass = @{} + $PSBoundParameters
    Microsoft.PowerShell.Utility\Invoke-RestMethod @pass
}
Set-Location $Engine
& (Join-Path $Engine 'Yardstick.ps1') -All -NoInteractive -NoEmail
exit $LASTEXITCODE
'@ | Set-Content -LiteralPath $bootstrap

$allVariables = 'YARDSTICK_TENANT_ID', 'YARDSTICK_CLIENT_ID', 'YARDSTICK_SERVICE_CONNECTION_ID', 'YARDSTICK_FEDERATED_TOKEN_FILE',
    'SYSTEM_OIDCREQUESTURI', 'SYSTEM_ACCESSTOKEN', 'ACTIONS_ID_TOKEN_REQUEST_URL', 'ACTIONS_ID_TOKEN_REQUEST_TOKEN', 'GITHUB_ACTIONS'
$saved = @{}
foreach ($name in $allVariables) { $saved[$name] = [Environment]::GetEnvironmentVariable($name) }

try {
    foreach ($source in $Provider) {
        $childEnvironment = @{
            YARDSTICK_TENANT_ID = 'smoke-tenant'
            YARDSTICK_CLIENT_ID = $clientId
        }
        $expectedRequest = $null
        switch ($source) {
            'AzureDevOps' {
                $childEnvironment.YARDSTICK_SERVICE_CONNECTION_ID = $serviceConnectionId
                $childEnvironment.SYSTEM_OIDCREQUESTURI = "http://localhost:$port/ado/oidctoken"
                $childEnvironment.SYSTEM_ACCESSTOKEN = $jobToken
                $expectedRequest = @{ Method = 'POST'; Path = '/ado/oidctoken'; Query = "?api-version=7.1&serviceConnectionId=$serviceConnectionId" }
            }
            'GitHubActions' {
                $childEnvironment.GITHUB_ACTIONS = 'true'
                $childEnvironment.ACTIONS_ID_TOKEN_REQUEST_URL = "http://localhost:$port/github/idtoken?api-version=2.0"
                $childEnvironment.ACTIONS_ID_TOKEN_REQUEST_TOKEN = $jobToken
                $expectedRequest = @{ Method = 'GET'; Path = '/github/idtoken'; Query = '?api-version=2.0&audience=api%3A%2F%2FAzureADTokenExchange' }
            }
            'File' {
                $childEnvironment.YARDSTICK_FEDERATED_TOKEN_FILE = $tokenFile
            }
        }
        foreach ($name in $allVariables) { [Environment]::SetEnvironmentVariable($name, $childEnvironment[$name]) }

        $callLog = Join-Path $workspace "calls-$source.jsonl"
        $logFile = Join-Path $engine 'YLog.log'
        Remove-Item -LiteralPath $logFile -ErrorAction SilentlyContinue
        $tokenRequests.Clear()

        Write-Host ''
        Write-Host "Running Yardstick.ps1 -All -NoInteractive -NoEmail with ID tokens from $source"
        $output = & pwsh -NoProfile -NonInteractive -File $bootstrap -Engine $engine -CallLog $callLog 2>&1
        $exitCode = $LASTEXITCODE

        $outputText = ($output | ForEach-Object { "$_" }) -join "`n"
        $calls = @(if (Test-Path -LiteralPath $callLog) { Get-Content -LiteralPath $callLog | ForEach-Object { $_ | ConvertFrom-Json } })
        $entraCalls = @($calls | Where-Object kind -eq 'entra')
        $graphCalls = @($calls | Where-Object kind -eq 'graph')
        $requests = @($tokenRequests.ToArray())

        Assert-Smoke ($exitCode -eq 0) "[$source] Yardstick exited 0 (got $exitCode)"
        if ($expectedRequest) {
            Assert-Smoke ($requests.Count -ge 1) "[$source] requested an ID token ($($requests.Count) request(s))"
            Assert-Smoke (@($requests | Where-Object { $_.Method -ne $expectedRequest.Method -or $_.Path -ne $expectedRequest.Path -or $_.Query -ne $expectedRequest.Query }).Count -eq 0) "[$source] every ID token request was $($expectedRequest.Method) $($expectedRequest.Path)$($expectedRequest.Query)"
            Assert-Smoke (@($requests | Where-Object { $_.Authorization -ne "Bearer $jobToken" }).Count -eq 0) "[$source] every ID token request carried the job token"
        } else {
            Assert-Smoke ($requests.Count -eq 0) "[$source] made no ID token request"
        }
        Assert-Smoke ($entraCalls.Count -ge 1) "[$source] exchanged the ID token with Entra ($($entraCalls.Count) request(s))"
        Assert-Smoke (@($entraCalls | Where-Object { $_.uri -ne 'https://login.microsoftonline.com/smoke-tenant/oauth2/v2.0/token' }).Count -eq 0) "[$source] Entra requests used the federated tenant"
        Assert-Smoke (@($entraCalls | Where-Object { 'client_secret' -in $_.keys }).Count -eq 0) "[$source] no Entra request sent client_secret"
        Assert-Smoke (@($entraCalls | Where-Object { $_.assertion -ne $idToken }).Count -eq 0) "[$source] every Entra request presented the ID token"
        Assert-Smoke (@($graphCalls | Where-Object { $_.uri -match '/applications' }).Count -eq 0) "[$source] did not look up the app registration for secret expiry"
        Assert-Smoke (@($graphCalls | Where-Object { $_.authorization -ne 'Bearer smoke-graph-token' }).Count -eq 0) "[$source] Graph calls used the federated token ($($graphCalls.Count) call(s))"
        Assert-Smoke ($outputText -notmatch 'Credential Manager under target') "[$source] did not look for a stored credential"
        Assert-Smoke ($outputText -notmatch [regex]::Escape($jobToken) -and $outputText -notmatch [regex]::Escape($idToken)) "[$source] no token appeared in the output"

        $logText = if (Test-Path -LiteralPath $logFile) { Get-Content -LiteralPath $logFile -Raw } else { '' }
        Assert-Smoke ($logText -match 'Using workload identity federation') "[$source] the log records federated sign-in"
        Assert-Smoke ($logText -notmatch [regex]::Escape($jobToken) -and $logText -notmatch [regex]::Escape($idToken)) "[$source] no token was written to the log"

        if ($exitCode -ne 0) {
            Write-Host 'Yardstick output:' -ForegroundColor Yellow
            Write-Host $outputText
        }
    }
} finally {
    foreach ($name in $saved.Keys) { [Environment]::SetEnvironmentVariable($name, $saved[$name]) }
    $listener.Stop()
    $listener.Close()
    $server | Wait-Job -Timeout 5 | Out-Null
    $server | Remove-Job -Force
}

if ($KeepWorkspace) {
    Write-Host "Workspace kept at $workspace"
} else {
    Remove-Item -LiteralPath $workspace -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ''
if ($failures.Count -gt 0) {
    Write-Host "$($failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host 'Federation smoke test passed.' -ForegroundColor Green
exit 0
