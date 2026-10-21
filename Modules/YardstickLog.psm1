#Requires -Version 7.0

<#
Yardstick's log: Write-Log and the parser that reads the log back for
Send-YardstickLogReport.ps1. Writer and reader live together because the
report depends on the exact line format.

Line format, one entry per line:

    MM/dd/yyyy HH:mm:ss - [app] ERROR: message

The "[app] " tag is present while Set-YardstickLogContext names an app, and the
"ERROR: " / "WARNING: " prefix comes from -Level. Both are optional, so a line
with neither is the original format. Runs begin with the header written by
Write-Log -Init.

This module imports nothing, so any component (and the watchdog and backup
threads, through YardstickSupport) can load it.
#>

$script:TimestampFormat = 'MM/dd/yyyy HH:mm:ss'
$script:HeaderRule = '#######################################################'
$script:Context = $null
$script:ShowDetail = $false
$script:WarnedUnset = $false


function Write-YardstickHost {
    # A console that a child process left in a bad state makes Write-Host throw a
    # HostException. Losing the echo is acceptable; losing the run is not.
    param([string]$Message, [string]$Color)
    try {
        if ($Color) { Write-Host $Message -ForegroundColor $Color } else { Write-Host $Message }
    } catch { }
}


function Add-YardstickLogText {
    # The watchdog thread appends to the same file as the run, so a write can
    # find the file briefly locked. Retry rather than drop the line.
    param([string]$Path, [string]$Text)
    for ($attempt = 1; ; $attempt++) {
        try {
            [System.IO.File]::AppendAllText($Path, $Text)
            return
        } catch [System.IO.IOException] {
            if ($attempt -ge 20) { throw }
            Start-Sleep -Milliseconds 25
        }
    }
}


function Write-Log {
    <#
    .SYNOPSIS
    Writes a timestamped entry to the Yardstick log and the console.

    .DESCRIPTION
    The log file is $Global:LogFile in $Global:LogLocation. When either is unset,
    or the file cannot be written, the entry still goes to the console.

    .PARAMETER Content
    The message. Don't start it with a severity; use -Level.

    .PARAMETER Level
    Info (default), Detail, Warning or Error. Warning and Error add the
    "WARNING: " / "ERROR: " prefix and are coloured on the console. Detail is a
    routine step: it is always written to the file but shown on the console only
    after Write-Log -Init -ShowDetail (or Set-YardstickLogContext -ShowDetail).

    .PARAMETER Init
    Start a new log: delete the file and write the run header.

    .PARAMETER ShowDetail
    With -Init, show Detail entries on the console for this run.

    .PARAMETER NoContext
    Leave off the "[app] " tag, for lines that don't belong to the current app.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [AllowEmptyString()]
        [string]$Content,

        [ValidateSet('Info', 'Detail', 'Warning', 'Error')]
        [string]$Level = 'Info',

        [switch]$Init,

        [switch]$ShowDetail,

        [switch]$NoContext
    )

    $now = Get-Date -Format $script:TimestampFormat
    $line = $null
    if ($Content) {
        $tag = if ($script:Context -and -not $NoContext) { "[$script:Context] " } else { '' }
        $prefix = switch ($Level) { 'Error' { 'ERROR: ' } 'Warning' { 'WARNING: ' } default { '' } }
        $line = "$now - $tag$prefix$Content"
    }
    $color = switch ($Level) { 'Error' { 'Red' } 'Warning' { 'Yellow' } default { $null } }
    $echo = $line -and ($Level -ne 'Detail' -or $script:ShowDetail)

    if ($Init) {
        $script:ShowDetail = [bool]$ShowDetail
        $echo = $line -and ($Level -ne 'Detail' -or $script:ShowDetail)
    }

    if (-not $Global:LogLocation -or -not $Global:LogFile) {
        if (-not $script:WarnedUnset) {
            Write-YardstickHost 'WARNING: $Global:LogLocation and $Global:LogFile are not set, so log entries go to the console only.' -Color Yellow
            $script:WarnedUnset = $true
        }
        if ($echo) { Write-YardstickHost $line -Color $color }
        return
    }

    $path = Join-Path $Global:LogLocation $Global:LogFile
    $newLine = [Environment]::NewLine
    try {
        if ($Init) {
            $header = "$script:HeaderRule${newLine}LOGGING STARTED AT $now$newLine$script:HeaderRule$newLine"
            try {
                # Overwriting keeps the file's owner and permissions. Deleting it would
                # let an elevated run replace it with one the user cannot write to.
                [System.IO.File]::WriteAllText($path, $header)
            } catch [System.UnauthorizedAccessException] {
                # Left read-only by an elevated run. The user can usually still delete
                # it from the Yardstick folder and start a file of their own.
                Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                Add-YardstickLogText -Path $path -Text $header
            }
        }
        if ($line) {
            Add-YardstickLogText -Path $path -Text "$line$newLine"
        }
    } catch {
        Write-YardstickHost "WARNING: Could not write to the log file ${path}: $_" -Color Yellow
    }

    if ($echo) { Write-YardstickHost $line -Color $color }
}


function Set-YardstickLogContext {
    <#
    .SYNOPSIS
    Tags every following log entry with an app ID until cleared.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)][AllowEmptyString()][AllowNull()][string]$App,
        [switch]$ShowDetail
    )
    $script:Context = if ($App) { $App } else { $null }
    if ($PSBoundParameters.ContainsKey('ShowDetail')) { $script:ShowDetail = [bool]$ShowDetail }
}


function Clear-YardstickLogContext {
    <#
    .SYNOPSIS
    Stops tagging log entries with an app ID.
    #>
    [CmdletBinding()]
    param()
    $script:Context = $null
}


function Get-YardstickLogContext {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    $script:Context
}


#region Parser (used by Send-YardstickLogReport.ps1)

# The stages the runner passes to Add-FailedApplication, which logs
# "<stage> failed: <message>" at Error.
$script:FailureStages = @(
    'Configuration', 'Recipe Validation', 'Manual Drop', 'Pre-Download Script',
    'Version Validation', 'Download', 'Download Script', 'Post-Download Script',
    'MSI Product Code Retrieval', 'Placeholder Replacement', 'Intune Upload',
    'Post-Run Script', 'General Processing', 'Processing'
)
$script:FailureLinePattern = '^ERROR: (' + (($script:FailureStages | ForEach-Object { [regex]::Escape($_) }) -join '|') + ') failed: ([\s\S]+)$'

function Get-YardstickLogRun {
    <#
    .SYNOPSIS
    Splits the log into runs and returns the requested one.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string[]]$Lines,
        [int]$Index
    )

    $starts = @()
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match '^LOGGING STARTED AT (.+)$') {
            $starts += [PSCustomObject]@{ Line = $i; StartTime = $Matches[1].Trim() }
        }
    }

    # A log fragment with no header is still a run.
    if ($starts.Count -eq 0) {
        return [PSCustomObject]@{ StartTime = 'unknown'; Lines = $Lines }
    }

    $resolved = if ($Index -lt 0) { $starts.Count + $Index } else { $Index }
    if ($resolved -lt 0 -or $resolved -ge $starts.Count) {
        throw "RunIndex $Index is out of range - the log contains $($starts.Count) run(s)."
    }

    $from = $starts[$resolved].Line
    $to = if ($resolved -lt $starts.Count - 1) { $starts[$resolved + 1].Line - 1 } else { $Lines.Count - 1 }

    return [PSCustomObject]@{
        StartTime = $starts[$resolved].StartTime
        Lines     = $Lines[$from..$to]
    }
}


function ConvertTo-YardstickLogEntry {
    <#
    .SYNOPSIS
    Turns raw log lines into timestamped entries, folding wrapped error text into
    the entry it belongs to.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Lines)

    $entries = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($line in $Lines) {
        if ($line -match '^#{5,}\s*$' -or $line -match '^LOGGING STARTED AT ') { continue }

        if ($line -match '^(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}) - ([\s\S]*)$') {
            $stamp = [datetime]::MinValue
            [void][datetime]::TryParseExact($Matches[1], 'MM/dd/yyyy HH:mm:ss',
                [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$stamp)
            $entries.Add([PSCustomObject]@{ Time = $stamp; Message = $Matches[2] })
        }
        elseif ($entries.Count -gt 0) {
            # Multi-line errors (stack traces, Graph response bodies) land here.
            $entries[$entries.Count - 1].Message += "`n$line"
        }
    }

    return $entries
}


function Split-NameAndVersion {
    <#
    .SYNOPSIS
    Separates "<display name> <version>" using the display name already seen in
    the log, falling back to the last whitespace-delimited token.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        [string]$KnownName
    )

    $text = $Text.Trim()
    if ($KnownName -and $text.StartsWith("$KnownName ")) {
        return @{ Name = $KnownName; Version = $text.Substring($KnownName.Length + 1).Trim() }
    }

    $split = $text.LastIndexOf(' ')
    if ($split -lt 1) {
        return @{ Name = $text; Version = 'Unknown' }
    }
    return @{ Name = $text.Substring(0, $split); Version = $text.Substring($split + 1) }
}


function Get-YardstickFailureStage {
    <#
    .SYNOPSIS
    Maps a logged error message back to the stage Add-FailedApplication would
    have recorded. Only needed for logs written before the failure line named
    the stage ("<stage> failed: <message>").
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$ErrorMessage)

    switch -Regex ($ErrorMessage) {
        'Unable to open parameters file|Failed to resolve base recipe' { return 'Configuration' }
        'Recipe validation failed|schema'                              { return 'Recipe Validation' }
        'pre-download PowerShell script'                               { return 'Pre-Download Script' }
        'Version validation failed|Version contains|version'           { return 'Version Validation' }
        'staged payload|dropbox'                                       { return 'Manual Drop' }
        'post download PowerShell script'                              { return 'Post-Download Script' }
        'download PowerShell script'                                   { return 'Download Script' }
        'URL is empty|Error downloading file'                          { return 'Download' }
        'MSI Product Code'                                             { return 'MSI Product Code Retrieval' }
        'placeholder'                                                  { return 'Placeholder Replacement' }
        # A missing icon no longer fails the app, but older logs still have the failure.
        'Icon file'                                                    { return 'Icon Retrieval' }
        'upload application to Intune'                                 { return 'Intune Upload' }
        'post run PowerShell script'                                   { return 'Post-Run Script' }
        'Unexpected error'                                             { return 'General Processing' }
        default                                                        { return 'Processing' }
    }
}


function New-YardstickAppState {
    param([Parameter(Mandatory)][string]$Id)

    return [PSCustomObject]@{
        Id           = $Id
        DisplayName  = $null
        Version      = $null
        Action       = 'Updated'
        Dependents   = @{}
        AutoParts    = [System.Collections.Generic.List[string]]::new()
        BackupOk     = $false
        BackupPruned = 0
        BackupStatus = $null
        Succeeded    = $false
        Failures     = [System.Collections.Generic.List[PSObject]]::new()
        ErrorLines   = [System.Collections.Generic.List[string]]::new()
    }
}


function ConvertFrom-YardstickLog {
    <#
    .SYNOPSIS
    Replays a run's log entries into per-application results.
    #>
    param([Parameter(Mandatory)][System.Collections.Generic.List[PSObject]]$Entries)

    $apps = [System.Collections.Generic.List[PSObject]]::new()
    $current = $null

    foreach ($entry in $Entries) {
        $m = $entry.Message

        # Lines logged while an app is processed carry its tag. Strip it only when it
        # names the current app, so an untagged message that happens to start with
        # brackets (older "[Label] ..." retry lines) is left alone.
        if ($current -and $m -match '^\[([^\]\s]+)\] ([\s\S]*)$' -and $Matches[1] -eq $current.Id) {
            $m = $Matches[2]
        }

        if ($m -match '^Starting update for (.+?)\.\.\.$') {
            if ($current) { $apps.Add($current) }
            $current = New-YardstickAppState -Id $Matches[1]
            continue
        }

        if (-not $current) { continue }

        switch -Regex ($m) {
            '^Retrieving all versions of application with display name: (.+)$' {
                $current.DisplayName = $Matches[1].Trim()
            }
            '^Checking if ([\s\S]+) is a new version\.\.\.$' {
                $parsed = Split-NameAndVersion -Text $Matches[1] -KnownName $current.DisplayName
                if (-not $current.DisplayName) { $current.DisplayName = $parsed.Name }
                $current.Version = $parsed.Version
            }
            'Running maintenance sweep\.$' {
                $current.Action = 'Maintained'
            }
            '^Tracked successful application: ([\s\S]+)$' {
                $parsed = Split-NameAndVersion -Text $Matches[1] -KnownName $current.DisplayName
                if (-not $current.DisplayName) { $current.DisplayName = $parsed.Name }
                $current.Version = $parsed.Version
                $current.Succeeded = $true
            }
            '^Done: ([\s\S]+) \(([^)]+)\)\.$' {
                $action = $Matches[2]
                $parsed = Split-NameAndVersion -Text $Matches[1] -KnownName $current.DisplayName
                if (-not $current.DisplayName) { $current.DisplayName = $parsed.Name }
                $current.Version = $parsed.Version
                $current.Action = $action
                $current.Succeeded = $true
            }
            '^Tracked failed application: (\S+) - ([\s\S]+)$' {
                $current.Failures.Add([PSCustomObject]@{ Id = $Matches[1]; ErrorMessage = $Matches[2] })
            }
            ($script:FailureLinePattern) {
                $current.Failures.Add([PSCustomObject]@{ Id = $current.Id; ErrorMessage = $Matches[2]; Stage = $Matches[1] })
            }
            '^ERROR: ([\s\S]+)$' {
                $current.ErrorLines.Add($Matches[1])
            }
            '^Attaching supersedence \(.+\) from .+ to (\d+) target\(s\)$' {
                $current.AutoParts.Add("supersedes $($Matches[1])")
            }
            '^Auto-update \(.+\) is set on (\d+) assignment\(s\)' {
                $current.AutoParts.Add("auto-update on $($Matches[1]) assignment(s)")
            }
            '^Pinning .+ anchor for ' {
                $current.AutoParts.Add('anchor: pinned')
            }
            '^Updated dependent app ([\s\S]+?) to reference [\s\S]+\.$' {
                $current.Dependents[$Matches[1]] = 'Updated'
            }
            '^(?:WARNING: )?Failed to update dependent app ([\s\S]+?) on attempt \d+: ([\s\S]+)$' {
                # A later attempt may still succeed and overwrite this.
                $current.Dependents[$Matches[1]] = "Failed ($($Matches[2]))"
            }
            '^Skipping dependent app ([\s\S]+?) because it is included in the blacklist\.$' {
                $current.Dependents[$Matches[1]] = 'Skipped (blacklisted)'
            }
            '^Backed up \S+ to ' {
                $current.BackupOk = $true
            }
            '^Pruned old backup ' {
                $current.BackupPruned++
            }
            '^WARNING: Backup of \S+ (?:failed|timed out)[\s\S]*$' {
                $current.BackupStatus = $m -replace '^WARNING: ', ''
            }
            '^WARNING: Could not queue backup for .+?: ([\s\S]+)$' {
                $current.BackupStatus = "not queued: $($Matches[1])"
            }
        }
    }

    if ($current) { $apps.Add($current) }

    foreach ($app in $apps) {
        if (-not $app.BackupStatus -and $app.BackupOk) {
            $app.BackupStatus = if ($app.BackupPruned -gt 0) { "ok ($($app.BackupPruned) pruned)" } else { 'ok' }
        }
    }

    return $apps
}

#endregion


Export-ModuleMember -Function @(
    'Write-Log'
    'Set-YardstickLogContext'
    'Clear-YardstickLogContext'
    'Get-YardstickLogContext'
    'Get-YardstickLogRun'
    'ConvertTo-YardstickLogEntry'
    'Split-NameAndVersion'
    'Get-YardstickFailureStage'
    'New-YardstickAppState'
    'ConvertFrom-YardstickLog'
)
