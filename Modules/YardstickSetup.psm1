# Yardstick setup: the building blocks behind Install-Yardstick.ps1.
#
# Each wizard step is a function here so it can be unit-tested and re-run on its own.
# Nothing in this module is used by a normal Yardstick run.
#
# Creating the app registration needs a *delegated* (signed-in admin) Graph token,
# which the runner never has - it only ever uses the app's own client-credentials
# token. That sign-in goes through Connect-MgGraph when Microsoft.Graph.Authentication
# is installed (browser sign-in), and otherwise through the OAuth device-code flow
# over plain REST, so the installer does not add a module dependency.

# No -Force: re-importing here would tear down the copies the installer imported
# globally, along with any state they hold.
Import-Module (Join-Path $PSScriptRoot 'YardstickConfig.psm1')
Import-Module (Join-Path $PSScriptRoot 'YardstickGraph.psm1')
Import-Module (Join-Path $PSScriptRoot 'YardstickCredential.psm1')

# Microsoft Graph's well-known application ID; the same in every tenant.
$Script:GraphResourceAppId = '00000003-0000-0000-c000-000000000000'
# Public client ID of "Microsoft Graph Command Line Tools", the app Connect-MgGraph
# signs in through. It supports the device-code flow and needs no app of our own.
$Script:SetupClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$Script:SetupScopes = @(
    'Application.ReadWrite.All'        # create the app registration and its secret
    'AppRoleAssignment.ReadWrite.All'  # grant admin consent to its Graph permissions
    'DeviceManagementRBAC.Read.All'    # list Intune scope tags
    'Group.Read.All'                   # look up the test deployment group by name
)
$Script:GraphBaseUri = 'https://graph.microsoft.com'

$Script:NonInteractive = $false
$Script:SetupSession = $null


#region Console helpers

function Set-YardstickSetupMode {
    <#
    .SYNOPSIS
    Switches the prompt helpers between asking and silently taking their defaults.
    #>
    [CmdletBinding()]
    param([switch]$NonInteractive)
    $Script:NonInteractive = [bool]$NonInteractive
}


function Write-YardstickSetupStep {
    param([Parameter(Mandatory)][string]$Title)
    Write-Host ''
    Write-Host ('=' * 70) -ForegroundColor DarkCyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host ('=' * 70) -ForegroundColor DarkCyan
}


function Write-YardstickSetupStatus {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('OK', 'Warn', 'Fail', 'Info', 'Skip')]
        [string]$Level = 'Info'
    )
    $tag, $color = switch ($Level) {
        'OK'   { '[ OK ]', 'Green' }
        'Warn' { '[WARN]', 'Yellow' }
        'Fail' { '[FAIL]', 'Red' }
        'Skip' { '[SKIP]', 'DarkGray' }
        default { '[ .. ]', 'Gray' }
    }
    Write-Host "$tag " -ForegroundColor $color -NoNewline
    Write-Host $Message
}


function Read-YardstickSetupValue {
    <#
    .SYNOPSIS
    Prompts for a value, offering a default that Enter accepts.
    #>
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default,
        [switch]$Required
    )

    if ($Script:NonInteractive) {
        if ($Required -and [string]::IsNullOrWhiteSpace($Default)) {
            throw "A value for '$Prompt' is required. Pass it as a parameter when using -NonInteractive."
        }
        return $Default
    }

    $label = if ([string]::IsNullOrWhiteSpace($Default)) { $Prompt } else { "$Prompt [$Default]" }
    while ($true) {
        $answer = Read-Host $label
        if ([string]::IsNullOrWhiteSpace($answer)) { $answer = $Default }
        if (-not $Required -or -not [string]::IsNullOrWhiteSpace($answer)) {
            if ($null -eq $answer) { return '' }
            return $answer.Trim()
        }
        Write-Host '  A value is required.' -ForegroundColor Yellow
    }
}


function Read-YardstickSetupYesNo {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [bool]$Default = $true
    )

    if ($Script:NonInteractive) { return $Default }

    $suffix = if ($Default) { '[Y/n]' } else { '[y/N]' }
    while ($true) {
        $answer = Read-Host "$Prompt $suffix"
        if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
        switch -Regex ($answer.Trim()) {
            '^(y|yes)$' { return $true }
            '^(n|no)$'  { return $false }
        }
        Write-Host '  Please answer y or n.' -ForegroundColor Yellow
    }
}


function Read-YardstickSetupChoice {
    <#
    .SYNOPSIS
    Shows a numbered list and returns the chosen item(s).

    .PARAMETER Multiple
    Accept a comma-separated list of numbers. An empty answer returns the defaults.
    #>
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [Parameter(Mandatory)][string[]]$Choices,
        [string[]]$Default = @(),
        [switch]$Multiple
    )

    if ($Script:NonInteractive) { return $Default }

    Write-Host $Prompt
    for ($i = 0; $i -lt $Choices.Count; $i++) {
        $marker = if ($Default -contains $Choices[$i]) { '*' } else { ' ' }
        Write-Host ("  {0}{1,3}. {2}" -f $marker, ($i + 1), $Choices[$i])
    }

    $hint = if ($Multiple) { 'Numbers, comma-separated' } else { 'Number' }
    if ($Default.Count -gt 0) { $hint += ' (Enter keeps the * items)' }
    while ($true) {
        $answer = Read-Host $hint
        if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
        $picked = @()
        $valid = $true
        foreach ($part in ($answer -split '[,\s]+' | Where-Object { $_ })) {
            $n = 0
            if ([int]::TryParse($part, [ref]$n) -and $n -ge 1 -and $n -le $Choices.Count) {
                $picked += $Choices[$n - 1]
            } else {
                $valid = $false
            }
        }
        if ($valid -and $picked.Count -gt 0 -and ($Multiple -or $picked.Count -eq 1)) {
            return @($picked | Select-Object -Unique)
        }
        Write-Host "  Enter a number between 1 and $($Choices.Count)." -ForegroundColor Yellow
    }
}

#endregion


#region Permissions and tokens

function Get-YardstickRequiredGraphPermissions {
    <#
    .SYNOPSIS
    The Microsoft Graph application permissions Yardstick's app registration uses.

    .DESCRIPTION
    Derived from every Graph endpoint the runner and its modules call. Required ones
    are needed by a normal run; optional ones only enable a side feature.

    Assignment filters need no permission of their own: the runner copies existing
    filters onto a new version by ID, through the app's assignments, and never looks
    one up by name.
    #>
    [CmdletBinding()]
    param([switch]$RequiredOnly)

    $all = @(
        [pscustomobject]@{ Name = 'DeviceManagementApps.ReadWrite.All'; Required = $true;  Reason = 'Create, update, assign and delete Win32 apps in Intune.' }
        [pscustomobject]@{ Name = 'DeviceManagementRBAC.Read.All';      Required = $true;  Reason = 'Resolve Intune scope tag names to IDs.' }
        [pscustomobject]@{ Name = 'Group.Read.All';                     Required = $false; Reason = 'Look up the test deployment group by name during setup.' }
        [pscustomobject]@{ Name = 'Application.Read.All';               Required = $false; Reason = 'Discover the client secret expiry date when none was recorded at setup.' }
    )
    if ($RequiredOnly) { return @($all | Where-Object Required) }
    $all
}


function Test-YardstickGraphRoleGranted {
    <#
    .SYNOPSIS
    Whether a Graph application permission is among the granted roles. A
    .ReadWrite.All role covers the matching .Read.All one, so an app that already
    holds the broader permission is not reported as missing the narrower one.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowEmptyCollection()][string[]]$Roles = @(),
        [Parameter(Mandatory)][string]$Name
    )

    if ($Roles -contains $Name) { return $true }
    ($Name -match '\.Read\.All$') -and ($Roles -contains ($Name -replace '\.Read\.All$', '.ReadWrite.All'))
}


function ConvertFrom-YardstickJwt {
    <#
    .SYNOPSIS
    Decodes the claims of a JWT access token. The signature is not checked.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Token)

    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { throw 'The value is not a JWT.' }
    $payload = $parts[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
    }
    $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload))
    $json | ConvertFrom-Json
}

#endregion


#region Prerequisites

function Install-YardstickModule {
    <#
    .SYNOPSIS
    Installs a PowerShell Gallery module for the current user when it is missing.
    Returns $true when the module is available afterwards.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$Name)

    if (Get-Module -ListAvailable -Name $Name) { return $true }
    if (-not $PSCmdlet.ShouldProcess($Name, 'Install for the current user from the PowerShell Gallery')) { return $false }
    # PowerShell 7.4+ ships PSResourceGet, which keeps its own repository list; on
    # some machines PowerShellGet's Install-Module has no PSGallery registered, so
    # try PSResourceGet first and fall back to Install-Module.
    $errors = @()
    if (Get-Command -Name Install-PSResource -ErrorAction SilentlyContinue) {
        try {
            Install-PSResource -Name $Name -Scope CurrentUser -TrustRepository -Quiet -ErrorAction Stop
            if (Get-Module -ListAvailable -Name $Name) { return $true }
        } catch {
            $errors += $_.Exception.Message
        }
    }
    try {
        # -Force skips the "untrusted repository" prompt for PSGallery.
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        if (Get-Module -ListAvailable -Name $Name) { return $true }
    } catch {
        $errors += $_.Exception.Message
    }
    Write-YardstickSetupStatus "Could not install '$Name': $($errors -join ' / ')" -Level Fail
    return $false
}


function Unblock-YardstickFiles {
    <#
    .SYNOPSIS
    Removes the "downloaded from the internet" mark from Yardstick's scripts and tools.

    .DESCRIPTION
    A zip downloaded from GitHub marks every file it contains, and PowerShell then
    refuses to load the modules under a RemoteSigned policy. Returns the count of
    files that were unblocked.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$ProjectRoot)

    $extensions = @('.ps1', '.psm1', '.psd1', '.exe', '.dll', '.yaml')
    $skip = '[\\/](\.git|Artifacts|Local)[\\/]'
    $count = 0
    Get-ChildItem -LiteralPath $ProjectRoot -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() -and $_.FullName -notmatch $skip } |
        ForEach-Object {
            $zone = Get-Item -LiteralPath $_.FullName -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue
            if ($zone -and $PSCmdlet.ShouldProcess($_.FullName, 'Unblock-File')) {
                Unblock-File -LiteralPath $_.FullName
                $count++
            }
        }
    $count
}


function Test-YardstickExecutionPolicy {
    <#
    .SYNOPSIS
    Returns $true when the effective execution policy lets Yardstick's unsigned
    scripts run.
    #>
    $policy = Get-ExecutionPolicy
    $policy -in @('Unrestricted', 'RemoteSigned', 'Bypass')
}

#endregion


#region Preferences file

function New-YardstickPreferencesFile {
    <#
    .SYNOPSIS
    Creates Local\preferences.yaml from the example when no preferences file exists.

    .OUTPUTS
    The path of the preferences file and whether it was just created.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$ProjectRoot)

    $path = Get-YardstickPreferencesPath -ProjectRoot $ProjectRoot
    if (Test-Path -LiteralPath $path -PathType Leaf) {
        return [pscustomobject]@{ Path = $path; Created = $false }
    }

    $example = Join-Path (Get-YardstickProjectLayout -ProjectRoot $ProjectRoot).ProjectRoot 'Config\Preferences.example.yaml'
    if (-not (Test-Path -LiteralPath $example -PathType Leaf)) {
        throw "The preferences template '$example' is missing. Re-download Yardstick or restore the Config folder."
    }
    if ($PSCmdlet.ShouldProcess($path, 'Create preferences file from Config\Preferences.example.yaml')) {
        $folder = Split-Path -Parent $path
        if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        Copy-Item -LiteralPath $example -Destination $path
    }
    [pscustomobject]@{ Path = $path; Created = $true }
}


function ConvertTo-YardstickYamlValue {
    <#
    .SYNOPSIS
    Formats a value for a single "key: value" line of YAML.

    .DESCRIPTION
    Strings are single-quoted, where a backslash is literal, so Windows paths survive
    unchanged. Lists become a flow sequence on the same line.
    #>
    param([AllowNull()][AllowEmptyCollection()]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [bool]) { return $Value.ToString().ToLowerInvariant() }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return ([IFormattable]$Value).ToString($null, [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = @($Value | ForEach-Object { ConvertTo-YardstickYamlValue ([string]$_) })
        return '[' + ($items -join ', ') + ']'
    }
    $text = [string]$Value
    if ($text.Length -eq 0) { return '' }
    "'" + $text.Replace("'", "''") + "'"
}


function Set-YardstickPreferenceValue {
    <#
    .SYNOPSIS
    Sets one top-level key in a preferences file without disturbing its comments.

    .DESCRIPTION
    ConvertTo-Yaml would rewrite the whole file and drop every comment, which are the
    file's documentation. Instead the existing "key:" line is replaced in place (a
    commented-out "# key:" line is uncommented), along with any block-list items
    under it. A key that is not in the file at all is appended.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Key,
        [AllowNull()][AllowEmptyCollection()][AllowEmptyString()]
        $Value
    )

    $text = [IO.File]::ReadAllText($Path)
    $newline = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }
    $endsWithNewline = $text.EndsWith("`n")
    $lines = [System.Collections.Generic.List[string]]::new()
    if ($text.Length -gt 0) {
        $lines.AddRange([string[]]($text.TrimEnd("`r", "`n") -split "`r?`n"))
    }

    $escaped = [regex]::Escape($Key)
    $index = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "^$escaped\s*:") { $index = $i; break }
    }
    if ($index -lt 0) {
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match "^#\s*$escaped\s*:") { $index = $i; break }
        }
    }

    $formatted = ConvertTo-YardstickYamlValue $Value
    # Not $newLine: variable names are case-insensitive, so it would clobber $newline.
    $entry = if ($formatted) { "${Key}: $formatted" } else { "${Key}:" }

    if ($index -ge 0) {
        $lines[$index] = $entry
        # Drop a block list ("- item" lines) or indented continuation under the key.
        while ($index + 1 -lt $lines.Count -and $lines[$index + 1] -match '^(\s*-(\s|$)|\s+[^#\s])') {
            $lines.RemoveAt($index + 1)
        }
    } else {
        $lines.Add($entry)
        $endsWithNewline = $true
    }

    $output = ($lines -join $newline)
    if ($endsWithNewline) { $output += $newline }
    if ($PSCmdlet.ShouldProcess($Path, "Set $Key")) {
        [IO.File]::WriteAllText($Path, $output, [Text.UTF8Encoding]::new($false))
    }
}

#endregion


#region Delegated Graph sign-in

function Get-YardstickSetupGraphError {
    <#
    .SYNOPSIS
    Pulls the HTTP status and Graph error code/message out of a failed request.
    #>
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $status = $null
    $response = $ErrorRecord.Exception.PSObject.Properties['Response']
    if ($response -and $response.Value -and $response.Value.PSObject.Properties['StatusCode']) {
        try { $status = [int]$response.Value.StatusCode } catch { $status = $null }
    }

    $code = $null
    $message = $ErrorRecord.Exception.Message
    $details = if ($ErrorRecord.ErrorDetails) { $ErrorRecord.ErrorDetails.Message } else { $null }
    if ($details) {
        try {
            $parsed = $details | ConvertFrom-Json -ErrorAction Stop
            if ($parsed.PSObject.Properties['error']) {
                if ($parsed.error -is [string]) {
                    # OAuth endpoints: { "error": "authorization_pending", "error_description": "..." }
                    $code = $parsed.error
                    if ($parsed.PSObject.Properties['error_description']) { $message = $parsed.error_description }
                } else {
                    $code = $parsed.error.code
                    $message = $parsed.error.message
                }
            }
        } catch {
            $message = $details
        }
    }
    if (-not $status -and $ErrorRecord.Exception.Message -match '\b(400|401|403|404|409|429)\b') { $status = [int]$Matches[1] }

    [pscustomobject]@{ StatusCode = $status; Code = $code; Message = $message }
}


function Connect-YardstickSetupGraph {
    <#
    .SYNOPSIS
    Signs an administrator in to Microsoft Graph for the setup steps.

    .DESCRIPTION
    Uses Connect-MgGraph (browser sign-in) when Microsoft.Graph.Authentication is
    installed, otherwise the device-code flow over REST: the user opens a URL, enters
    a short code and signs in there. The session is cached for the rest of setup.

    .PARAMETER TenantId
    Tenant to sign in to. By default the tenant of the account that signs in.

    .PARAMETER UseDeviceCode
    Use the device-code flow even when Microsoft.Graph.Authentication is installed.
    #>
    [CmdletBinding()]
    param(
        [string]$TenantId,
        [switch]$UseDeviceCode,
        [switch]$Force
    )

    if ($Script:SetupSession -and -not $Force -and (
            $Script:SetupSession.Mode -eq 'MgGraph' -or
            ($Script:SetupSession.Mode -eq 'Rest' -and $Script:SetupSession.ExpiresOn -gt [datetimeoffset]::UtcNow.AddMinutes(5)))) {
        return $Script:SetupSession
    }

    if (-not $UseDeviceCode -and (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $connectArgs = @{ Scopes = $Script:SetupScopes; NoWelcome = $true; ErrorAction = 'Stop' }
        if ($TenantId) { $connectArgs.TenantId = $TenantId }
        Connect-MgGraph @connectArgs
        $context = Get-MgContext
        $Script:SetupSession = [pscustomobject]@{
            Mode        = 'MgGraph'
            TenantId    = $context.TenantId
            Account     = $context.Account
            AccessToken = $null
            ExpiresOn   = $null
        }
        return $Script:SetupSession
    }

    $authority = if ($TenantId) { $TenantId } else { 'organizations' }
    $base = "https://login.microsoftonline.com/$([uri]::EscapeDataString($authority))/oauth2/v2.0"
    $scope = (($Script:SetupScopes | ForEach-Object { "$Script:GraphBaseUri/$_" }) + 'offline_access') -join ' '

    $device = Invoke-RestMethod -Method Post -Uri "$base/devicecode" -TimeoutSec 60 -ContentType 'application/x-www-form-urlencoded' -Body @{
        client_id = $Script:SetupClientId
        scope     = $scope
    } -ErrorAction Stop

    Write-Host ''
    Write-Host "Sign in as an administrator who can create app registrations and grant admin consent" -ForegroundColor Cyan
    Write-Host "(Global Administrator, or Privileged Role Administrator plus Application Administrator)." -ForegroundColor Cyan
    Write-Host ''
    Write-Host "  1. Open   $($device.verification_uri)" -ForegroundColor White
    Write-Host "  2. Enter  $($device.user_code)" -ForegroundColor White
    Write-Host ''
    Write-Host 'Waiting for sign-in...' -ForegroundColor DarkGray

    $interval = if ($device.interval) { [int]$device.interval } else { 5 }
    $deadline = [datetime]::UtcNow.AddSeconds([int]$device.expires_in)
    $token = $null
    while (-not $token) {
        if ([datetime]::UtcNow -gt $deadline) { throw 'The sign-in code expired before sign-in completed. Run the step again.' }
        Start-Sleep -Seconds $interval
        try {
            $token = Invoke-RestMethod -Method Post -Uri "$base/token" -TimeoutSec 60 -ContentType 'application/x-www-form-urlencoded' -Body @{
                grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id   = $Script:SetupClientId
                device_code = $device.device_code
            } -ErrorAction Stop
        } catch {
            $err = Get-YardstickSetupGraphError $_
            switch ($err.Code) {
                'authorization_pending' { continue }
                'slow_down' { $interval += 5; continue }
                'authorization_declined' { throw 'Sign-in was declined.' }
                'expired_token' { throw 'The sign-in code expired before sign-in completed. Run the step again.' }
                default {
                    $hint = ''
                    if ($err.Message -match 'AADSTS53003|AADSTS50105|device code') {
                        $hint = ' Your tenant may block device-code sign-in; install the Microsoft.Graph.Authentication module (Install-Module Microsoft.Graph.Authentication -Scope CurrentUser) to sign in through the browser instead.'
                    }
                    throw "Sign-in failed: $($err.Message)$hint"
                }
            }
        }
    }

    $claims = ConvertFrom-YardstickJwt -Token $token.access_token
    $Script:SetupSession = [pscustomobject]@{
        Mode        = 'Rest'
        TenantId    = $claims.tid
        Account     = if ($claims.PSObject.Properties['upn']) { $claims.upn } elseif ($claims.PSObject.Properties['preferred_username']) { $claims.preferred_username } else { $null }
        AccessToken = [string]$token.access_token
        ExpiresOn   = [datetimeoffset]::UtcNow.AddSeconds([int]$token.expires_in)
    }
    $Script:SetupSession
}


function Use-YardstickSetupAppCredential {
    <#
    .SYNOPSIS
    Makes setup lookups (scope tags, groups) use the stored app credential, for when
    nobody signed in during this run - for example with -UseExistingApp.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Credential)

    $Script:SetupSession = [pscustomobject]@{
        Mode         = 'App'
        TenantId     = $Credential.TenantID
        Account      = $Credential.ClientID
        AccessToken  = $null
        ExpiresOn    = $null
        ClientId     = $Credential.ClientID
        ClientSecret = $Credential.ClientSecret
    }
    $Script:SetupSession
}


function Get-YardstickSetupSession {
    <#
    .SYNOPSIS
    Returns the current setup session, or $null when there is none.
    #>
    $Script:SetupSession
}


function Invoke-YardstickSetupGraphRequest {
    <#
    .SYNOPSIS
    Sends a request with the setup (delegated) session.

    .PARAMETER Uri
    A path relative to https://graph.microsoft.com (for example "v1.0/applications")
    or a full URL such as an @odata.nextLink.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')]
        [string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Uri,
        $Body
    )

    if (-not $Script:SetupSession) { throw 'Not signed in to Microsoft Graph. Call Connect-YardstickSetupGraph first.' }
    $fullUri = if ($Uri -match '^https://') { $Uri } else { "$Script:GraphBaseUri/$($Uri.TrimStart('/'))" }
    $json = if ($null -ne $Body) { $Body | ConvertTo-Json -Depth 10 -Compress } else { $null }

    if ($Script:SetupSession.Mode -eq 'MgGraph') {
        $mgArgs = @{ Method = $Method; Uri = $fullUri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
        if ($json) { $mgArgs.Body = $json; $mgArgs.ContentType = 'application/json' }
        return Invoke-MgGraphRequest @mgArgs
    }

    $bearer = $Script:SetupSession.AccessToken
    if ($Script:SetupSession.Mode -eq 'App') {
        # Connect-YardstickGraph caches the token and renews it near expiry.
        $bearer = (Connect-YardstickGraph -TenantID $Script:SetupSession.TenantId -ClientID $Script:SetupSession.ClientId -ClientSecret $Script:SetupSession.ClientSecret -PassThru).AccessToken
    }
    $restArgs = @{
        Method      = $Method
        Uri         = $fullUri
        Headers     = @{ Authorization = "Bearer $bearer" }
        TimeoutSec  = 120
        ErrorAction = 'Stop'
    }
    if ($json) { $restArgs.Body = $json; $restArgs.ContentType = 'application/json' }
    Invoke-RestMethod @restArgs
}


function Get-YardstickSetupGraphCollection {
    <#
    .SYNOPSIS
    GETs a Graph collection, following @odata.nextLink, and returns every item.
    #>
    param([Parameter(Mandatory)][string]$Uri)

    $next = $Uri
    while ($next) {
        $page = Invoke-YardstickSetupGraphRequest -Method GET -Uri $next
        if ($page.PSObject.Properties['value']) { $page.value }
        $next = if ($page.PSObject.Properties['@odata.nextLink']) { $page.'@odata.nextLink' } else { $null }
    }
}


function Invoke-YardstickSetupRetry {
    <#
    .SYNOPSIS
    Retries an action that can fail while Entra replicates a newly created object.
    #>
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [int]$Attempts = 6,
        [int]$DelaySeconds = 10,
        [int[]]$RetryStatusCodes = @(400, 404)
    )

    for ($attempt = 1; ; $attempt++) {
        try {
            return & $Action
        } catch {
            $err = Get-YardstickSetupGraphError $_
            if ($attempt -ge $Attempts -or $err.StatusCode -notin $RetryStatusCodes) { throw }
            Write-Verbose "Attempt $attempt failed ($($err.Message)); retrying in $DelaySeconds s."
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

#endregion


#region App registration

function Get-YardstickGraphServicePrincipal {
    <#
    .SYNOPSIS
    Returns Microsoft Graph's service principal in the tenant, with its app roles.
    #>
    Invoke-YardstickSetupGraphRequest -Uri "v1.0/servicePrincipals(appId='$Script:GraphResourceAppId')?`$select=id,appId,appRoles"
}


function Resolve-YardstickGraphAppRoleIds {
    <#
    .SYNOPSIS
    Maps Graph application permission names to their app role IDs.

    .DESCRIPTION
    The IDs are read from the tenant rather than hard-coded so a typo cannot
    silently grant the wrong permission.
    #>
    param(
        [Parameter(Mandatory)]$ServicePrincipal,
        [Parameter(Mandatory)][string[]]$PermissionName
    )

    $map = [ordered]@{}
    foreach ($name in $PermissionName) {
        $role = @($ServicePrincipal.appRoles | Where-Object { $_.value -eq $name }) | Select-Object -First 1
        if (-not $role) { throw "Microsoft Graph has no application permission named '$name'." }
        $map[$name] = [string]$role.id
    }
    $map
}


function Find-YardstickApplication {
    param([Parameter(Mandatory)][string]$DisplayName)
    $filter = [uri]::EscapeDataString("displayName eq '$($DisplayName.Replace("'", "''"))'")
    @(Get-YardstickSetupGraphCollection -Uri "v1.0/applications?`$filter=$filter&`$select=id,appId,displayName,requiredResourceAccess")
}


function New-YardstickRequiredResourceAccess {
    <#
    .SYNOPSIS
    Builds an app's requiredResourceAccess with the Graph roles added to whatever
    it already requests.
    #>
    param(
        [AllowNull()]$Existing,
        [Parameter(Mandatory)][System.Collections.IDictionary]$RoleIds
    )

    $result = @()
    $graphEntry = $null
    foreach ($entry in @($Existing | Where-Object { $_ })) {
        $access = @($entry.resourceAccess | ForEach-Object { @{ id = [string]$_.id; type = [string]$_.type } })
        $item = @{ resourceAppId = [string]$entry.resourceAppId; resourceAccess = $access }
        if ($item.resourceAppId -eq $Script:GraphResourceAppId) { $graphEntry = $item }
        $result += $item
    }
    if (-not $graphEntry) {
        $graphEntry = @{ resourceAppId = $Script:GraphResourceAppId; resourceAccess = @() }
        $result += $graphEntry
    }
    foreach ($id in $RoleIds.Values) {
        if (-not ($graphEntry.resourceAccess | Where-Object { $_.id -eq $id -and $_.type -eq 'Role' })) {
            $graphEntry.resourceAccess += @{ id = $id; type = 'Role' }
        }
    }
    , $result
}


function Get-YardstickServicePrincipal {
    param([Parameter(Mandatory)][string]$AppId)
    try {
        Invoke-YardstickSetupGraphRequest -Uri "v1.0/servicePrincipals(appId='$AppId')?`$select=id,appId,displayName"
    } catch {
        if ((Get-YardstickSetupGraphError $_).StatusCode -eq 404) { return $null }
        throw
    }
}


function Grant-YardstickAppRoles {
    <#
    .SYNOPSIS
    Grants admin consent for application permissions by assigning the app roles.

    .OUTPUTS
    Granted, AlreadyGranted and Failed permission names. Failed is non-empty when the
    signed-in user lacks the rights to grant consent.
    #>
    param(
        [Parameter(Mandatory)][string]$ServicePrincipalId,
        [Parameter(Mandatory)][string]$ResourceServicePrincipalId,
        [Parameter(Mandatory)][System.Collections.IDictionary]$RoleIds
    )

    $existing = @(Get-YardstickSetupGraphCollection -Uri "v1.0/servicePrincipals/$ServicePrincipalId/appRoleAssignments" |
        Where-Object { $_.resourceId -eq $ResourceServicePrincipalId } |
        ForEach-Object { [string]$_.appRoleId })

    $granted = @(); $already = @(); $failed = @()
    foreach ($name in $RoleIds.Keys) {
        $roleId = $RoleIds[$name]
        if ($existing -contains $roleId) { $already += $name; continue }
        try {
            Invoke-YardstickSetupRetry -Action {
                Invoke-YardstickSetupGraphRequest -Method POST -Uri "v1.0/servicePrincipals/$ResourceServicePrincipalId/appRoleAssignedTo" -Body @{
                    principalId = $ServicePrincipalId
                    resourceId  = $ResourceServicePrincipalId
                    appRoleId   = $roleId
                }
            } | Out-Null
            $granted += $name
        } catch {
            $err = Get-YardstickSetupGraphError $_
            Write-Verbose "Granting $name failed: $($err.Message)"
            $failed += $name
        }
    }
    [pscustomobject]@{ Granted = $granted; AlreadyGranted = $already; Failed = $failed }
}


function Get-YardstickAdminConsentUrl {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId
    )
    "https://login.microsoftonline.com/$TenantId/adminconsent?client_id=$ClientId"
}


function New-YardstickClientSecret {
    <#
    .SYNOPSIS
    Adds a client secret to an app registration. Returns the secret text (only
    available now) and its expiry.
    #>
    param(
        [Parameter(Mandatory)][string]$ApplicationObjectId,
        [ValidateRange(1, 24)][int]$Months = 12,
        [string]$DisplayName = "Yardstick ($env:COMPUTERNAME)"
    )

    $endDate = [datetime]::UtcNow.AddMonths($Months)
    $result = Invoke-YardstickSetupGraphRequest -Method POST -Uri "v1.0/applications/$ApplicationObjectId/addPassword" -Body @{
        passwordCredential = @{
            displayName = $DisplayName
            endDateTime = $endDate.ToString('o')
        }
    }
    [pscustomobject]@{
        SecretText  = [string]$result.secretText
        KeyId       = [string]$result.keyId
        EndDateTime = if ($result.endDateTime) { ([datetime]$result.endDateTime).ToUniversalTime() } else { $endDate }
    }
}


function Get-YardstickAppTokenRoles {
    <#
    .SYNOPSIS
    Gets an app-only token with the given credentials and returns the Graph
    application permissions in it (its "roles" claim).
    #>
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )

    $token = Connect-YardstickGraph -TenantID $TenantId -ClientID $ClientId -ClientSecret $ClientSecret -Force -PassThru
    $claims = ConvertFrom-YardstickJwt -Token $token.AccessToken
    if ($claims.PSObject.Properties['roles']) { return @($claims.roles) }
    @()
}


function Wait-YardstickAppCredential {
    <#
    .SYNOPSIS
    Waits until a new secret works and its token carries the required permissions.

    .DESCRIPTION
    A new secret and new consent both take a little while to reach every sign-in
    server, so the first attempts can fail with "invalid client secret" or return a
    token without the roles yet.
    #>
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [string[]]$RequiredRoles = @((Get-YardstickRequiredGraphPermissions -RequiredOnly).Name),
        [int]$TimeoutSeconds = 180,
        [int]$PollSeconds = 15
    )

    $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
    $roles = @(); $lastError = $null
    while ($true) {
        try {
            $roles = @(Get-YardstickAppTokenRoles -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret)
            $lastError = $null
            $missing = @($RequiredRoles | Where-Object { -not (Test-YardstickGraphRoleGranted -Roles $roles -Name $_) })
            if ($missing.Count -eq 0) {
                return [pscustomobject]@{ Success = $true; Roles = $roles; Missing = @(); Error = $null }
            }
        } catch {
            $lastError = $_.Exception.Message
        }
        if ([datetime]::UtcNow.AddSeconds($PollSeconds) -gt $deadline) { break }
        Start-Sleep -Seconds $PollSeconds
    }
    [pscustomobject]@{
        Success = $false
        Roles   = $roles
        Missing = @($RequiredRoles | Where-Object { -not (Test-YardstickGraphRoleGranted -Roles $roles -Name $_) })
        Error   = $lastError
    }
}


function New-YardstickAppRegistration {
    <#
    .SYNOPSIS
    Creates (or reuses) Yardstick's Entra app registration, grants its Graph
    permissions, adds a client secret and stores it in Credential Manager.

    .DESCRIPTION
    Safe to run again: an existing app with the same display name is reused, missing
    permissions are added, and permissions already granted are left alone. Each run
    adds a new client secret, because an existing secret's value cannot be read back.
    Requires Connect-YardstickSetupGraph first.

    .OUTPUTS
    TenantId, ClientId, ObjectId, Created, SecretExpiresOn, Consent (the result of
    Grant-YardstickAppRoles), ConsentUrl and Validation (Wait-YardstickAppCredential).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$DisplayName = 'Yardstick',
        [ValidateRange(1, 24)][int]$SecretMonths = 12,
        [switch]$IncludeOptionalPermissions,
        [string]$CredentialTarget = 'Yardstick:IntuneGraph',
        [switch]$SkipValidation
    )

    $session = $Script:SetupSession
    if (-not $session -or $session.Mode -eq 'App') { throw 'Not signed in to Microsoft Graph. Call Connect-YardstickSetupGraph first.' }

    $permissions = if ($IncludeOptionalPermissions) { Get-YardstickRequiredGraphPermissions } else { Get-YardstickRequiredGraphPermissions -RequiredOnly }
    $graphSp = Get-YardstickGraphServicePrincipal
    $roleIds = Resolve-YardstickGraphAppRoleIds -ServicePrincipal $graphSp -PermissionName @($permissions.Name)

    $created = $false
    $apps = @(Find-YardstickApplication -DisplayName $DisplayName)
    if ($apps.Count -gt 1) {
        throw "More than one app registration is named '$DisplayName'. Pass -AppDisplayName with a unique name, or delete the duplicates."
    }
    if ($apps.Count -eq 1) {
        $app = $apps[0]
        Write-YardstickSetupStatus "Found existing app registration '$DisplayName' ($($app.appId))." -Level OK
        $rra = New-YardstickRequiredResourceAccess -Existing $app.requiredResourceAccess -RoleIds $roleIds
        if ($PSCmdlet.ShouldProcess($DisplayName, 'Update requested API permissions')) {
            Invoke-YardstickSetupGraphRequest -Method PATCH -Uri "v1.0/applications/$($app.id)" -Body @{ requiredResourceAccess = $rra } | Out-Null
        }
    } else {
        if (-not $PSCmdlet.ShouldProcess($DisplayName, 'Create app registration')) { return }
        $app = Invoke-YardstickSetupGraphRequest -Method POST -Uri 'v1.0/applications' -Body @{
            displayName            = $DisplayName
            signInAudience         = 'AzureADMyOrg'
            notes                  = 'Created by Yardstick Install-Yardstick.ps1. Publishes Win32 apps to Intune.'
            requiredResourceAccess = (New-YardstickRequiredResourceAccess -Existing $null -RoleIds $roleIds)
        }
        $created = $true
        Write-YardstickSetupStatus "Created app registration '$DisplayName' ($($app.appId))." -Level OK
    }

    $sp = Get-YardstickServicePrincipal -AppId $app.appId
    if (-not $sp) {
        $sp = Invoke-YardstickSetupRetry -Action {
            Invoke-YardstickSetupGraphRequest -Method POST -Uri 'v1.0/servicePrincipals' -Body @{ appId = $app.appId }
        }
    }

    $consent = Grant-YardstickAppRoles -ServicePrincipalId $sp.id -ResourceServicePrincipalId $graphSp.id -RoleIds $roleIds
    $consentUrl = Get-YardstickAdminConsentUrl -TenantId $session.TenantId -ClientId $app.appId

    $secret = New-YardstickClientSecret -ApplicationObjectId $app.id -Months $SecretMonths
    Set-YardstickIntuneCredential -Target $CredentialTarget -TenantID $session.TenantId -ClientID $app.appId -ClientSecret $secret.SecretText -SecretExpiresOn $secret.EndDateTime

    $validation = $null
    if (-not $SkipValidation -and $consent.Failed.Count -eq 0) {
        Write-YardstickSetupStatus 'Waiting for the new secret and permissions to take effect (up to 3 minutes)...'
        $validation = Wait-YardstickAppCredential -TenantId $session.TenantId -ClientId $app.appId -ClientSecret $secret.SecretText
    }

    [pscustomobject]@{
        TenantId        = $session.TenantId
        ClientId        = [string]$app.appId
        ObjectId        = [string]$app.id
        Created         = $created
        SecretExpiresOn = $secret.EndDateTime
        Consent         = $consent
        ConsentUrl      = $consentUrl
        Validation      = $validation
    }
}

#endregion


#region Intune defaults

function Get-YardstickScopeTagNames {
    <#
    .SYNOPSIS
    Lists Intune scope tag names. Works with a signed-in or an app session.
    #>
    $tags = Get-YardstickSetupGraphCollection -Uri 'beta/deviceManagement/roleScopeTags?$select=id,displayName'
    @($tags | ForEach-Object { [string]$_.displayName } | Sort-Object)
}


function Find-YardstickGroup {
    <#
    .SYNOPSIS
    Finds an Entra group by display name or object ID. Returns id and displayName.

    .DESCRIPTION
    With an app session this needs Group.Read.All, which Yardstick does not require;
    callers should fall back to asking for the group's object ID on a 403.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $guid = [guid]::Empty
    if ([guid]::TryParse($Name, [ref]$guid)) {
        $uri = "v1.0/groups/$guid`?`$select=id,displayName"
        $single = $true
    } else {
        $filter = [uri]::EscapeDataString("displayName eq '$($Name.Replace("'", "''"))'")
        $uri = "v1.0/groups?`$filter=$filter&`$select=id,displayName"
        $single = $false
    }

    $result = Invoke-YardstickSetupGraphRequest -Uri $uri
    if ($single) { return @($result) }
    @($result.value)
}

#endregion


#region Scheduled task

function Get-YardstickScheduledTaskDefinition {
    <#
    .SYNOPSIS
    Builds the command line the scheduled task runs.

    .PARAMETER Group
    Run one recipe group instead of every recipe.
    #>
    param(
        [Parameter(Mandatory)][string]$ProjectRoot,
        [string]$Group,
        [string]$PwshPath = (Join-Path $PSHOME 'pwsh.exe')
    )

    $script = Join-Path $ProjectRoot 'Yardstick.ps1'
    $selection = if ($Group) { "-Group `"$Group`"" } else { '-All' }
    [pscustomobject]@{
        Execute          = $PwshPath
        Argument         = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" $selection -NoInteractive"
        WorkingDirectory = $ProjectRoot
    }
}


function Register-YardstickScheduledTask {
    <#
    .SYNOPSIS
    Registers a daily scheduled task that runs Yardstick unattended.

    .DESCRIPTION
    The task runs as the given account whether or not it is signed in, so that
    account must be the one that stored the Intune credential: Credential Manager
    entries belong to a single Windows user.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$ProjectRoot,
        [Parameter(Mandatory)][datetime]$At,
        [Parameter(Mandatory)][pscredential]$Credential,
        [string]$Group,
        [string]$TaskName = 'Yardstick',
        [string]$TaskPath = '\Yardstick\'
    )

    $definition = Get-YardstickScheduledTaskDefinition -ProjectRoot $ProjectRoot -Group $Group
    $action = New-ScheduledTaskAction -Execute $definition.Execute -Argument $definition.Argument -WorkingDirectory $definition.WorkingDirectory
    $trigger = New-ScheduledTaskTrigger -Daily -At $At
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Hours 12) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    if ($PSCmdlet.ShouldProcess("$TaskPath$TaskName", 'Register scheduled task')) {
        Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Action $action -Trigger $trigger -Settings $settings `
            -User $Credential.UserName -Password $Credential.GetNetworkCredential().Password `
            -Description 'Runs Yardstick to publish application updates to Intune.' -Force -ErrorAction Stop
    }
}

#endregion


Export-ModuleMember -Function @(
    'Set-YardstickSetupMode'
    'Write-YardstickSetupStep'
    'Write-YardstickSetupStatus'
    'Read-YardstickSetupValue'
    'Read-YardstickSetupYesNo'
    'Read-YardstickSetupChoice'
    'Get-YardstickRequiredGraphPermissions'
    'Test-YardstickGraphRoleGranted'
    'ConvertFrom-YardstickJwt'
    'Install-YardstickModule'
    'Unblock-YardstickFiles'
    'Test-YardstickExecutionPolicy'
    'New-YardstickPreferencesFile'
    'ConvertTo-YardstickYamlValue'
    'Set-YardstickPreferenceValue'
    'Get-YardstickSetupGraphError'
    'Connect-YardstickSetupGraph'
    'Use-YardstickSetupAppCredential'
    'Get-YardstickSetupSession'
    'Invoke-YardstickSetupGraphRequest'
    'Get-YardstickGraphServicePrincipal'
    'Resolve-YardstickGraphAppRoleIds'
    'Find-YardstickApplication'
    'New-YardstickRequiredResourceAccess'
    'Grant-YardstickAppRoles'
    'Get-YardstickAdminConsentUrl'
    'New-YardstickClientSecret'
    'Get-YardstickAppTokenRoles'
    'Wait-YardstickAppCredential'
    'New-YardstickAppRegistration'
    'Get-YardstickScopeTagNames'
    'Find-YardstickGroup'
    'Get-YardstickScheduledTaskDefinition'
    'Register-YardstickScheduledTask'
)
