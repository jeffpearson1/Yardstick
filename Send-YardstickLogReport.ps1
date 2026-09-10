<#
.SYNOPSIS
Rebuilds the Yardstick email report from YLog.log and sends it after the fact.

.DESCRIPTION
Yardstick normally emails its summary at the end of a run, straight out of the
in-memory trackers. When that email is skipped (-NoEmail, a declined prompt, an
Outlook failure, or an aborted run) the results are gone, but the log still has
everything the report needs. This script replays a run out of YLog.log,
repopulates the same trackers Yardstick.ps1 uses, and hands them to
Send-YardstickEmailReport.

Only what Write-Log actually recorded can be reconstructed, so a few fields are
inferred: the failure stage comes from the error text, and the action is
"Maintained" when the run logged a -Repair maintenance sweep, otherwise
"Updated".

.PARAMETER LogPath
Log file to read. Defaults to YLog.log next to this script.

.PARAMETER RunIndex
Which run block in the log to report on. Log files can hold more than one
"LOGGING STARTED AT" header when Write-Log -Init was not called. Negative values
count back from the end; -1 (the default) is the most recent run.

.PARAMETER RunParameters
Text shown in the report header as the parameters of the run. Defaults to a note
naming the log file and run start time.

.PARAMETER Preview
Open the message in Outlook for inspection instead of sending it.

.PARAMETER HtmlOutputPath
Also write the rendered HTML body here for browser preview.

.PARAMETER ParseOnly
Print what was parsed out of the log and exit without touching Outlook.

.EXAMPLE
.\Send-YardstickLogReport.ps1 -ParseOnly

.EXAMPLE
.\Send-YardstickLogReport.ps1 -Preview -HtmlOutputPath .\LogReportPreview.html

.EXAMPLE
.\Send-YardstickLogReport.ps1 -LogPath D:\Archive\YLog-20260908.log
#>

[CmdletBinding()]
param(
    [string]$LogPath = "$PSScriptRoot\YLog.log",

    [int]$RunIndex = -1,

    [string]$RunParameters,

    [switch]$Preview,

    [string]$HtmlOutputPath,

    [switch]$ParseOnly
)

$ErrorActionPreference = 'Stop'

# Keep this script's own logging out of the file it is reading.
$Global:LogLocation = $PSScriptRoot
$Global:LogFile = 'YEmailReport.log'

Import-Module "$PSScriptRoot\Modules\YardstickSupport.psm1" -Scope Global -Force
Import-Module powershell-yaml -Force


function Get-YardstickLogRun {
    <#
    .SYNOPSIS
    Splits the log into runs and returns the requested one.
    #>
    param(
        [Parameter(Mandatory)][string[]]$Lines,
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
    param([Parameter(Mandatory)][string[]]$Lines)

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
    have recorded. The stage itself is never written to the log.
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
            '^Tracked failed application: (\S+) - ([\s\S]+)$' {
                $current.Failures.Add([PSCustomObject]@{ Id = $Matches[1]; ErrorMessage = $Matches[2] })
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
            '^Failed to update dependent app ([\s\S]+?) on attempt \d+: ([\s\S]+)$' {
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


function Publish-YardstickTrackerFromLog {
    <#
    .SYNOPSIS
    Feeds parsed results into the tracking lists Send-YardstickEmailReport reads.
    #>
    param([Parameter(Mandatory)][System.Collections.Generic.List[PSObject]]$Apps)

    Initialize-ApplicationTracker

    foreach ($app in $Apps) {
        $displayName = if ($app.DisplayName) { $app.DisplayName } else { $app.Id }
        $version = if ($app.Version) { $app.Version } else { 'Unknown' }

        if ($app.Succeeded) {
            $autoStatus = if ($app.AutoParts.Count -gt 0) { ($app.AutoParts | Select-Object -Unique) -join ', ' } else { $null }
            Add-SuccessfulApplication -ApplicationId $app.Id -DisplayName $displayName -Version $version `
                -Action $app.Action -Dependents $app.Dependents -AutoUpdateStatus $autoStatus

            if ($app.BackupStatus) {
                # BackupStatus is normally stamped on by Wait-YardstickBackup, so it
                # has to be set on the tracked object directly.
                $tracked = & (Get-Module YardstickSupport) { $Script:SuccessfulApplications }
                $entry = $tracked | Select-Object -Last 1
                if ($entry) { $entry.BackupStatus = $app.BackupStatus }
            }
        }

        $failures = @($app.Failures)
        if ($failures.Count -eq 0 -and -not $app.Succeeded -and $app.ErrorLines.Count -gt 0) {
            # Errors that killed the recipe before any tracking line was written.
            $failures = @([PSCustomObject]@{ Id = $app.Id; ErrorMessage = $app.ErrorLines[-1] })
        }

        foreach ($failure in $failures) {
            Add-FailedApplication -ApplicationId $failure.Id -DisplayName $displayName -Version $version `
                -ErrorMessage $failure.ErrorMessage -FailureStage (Get-YardstickFailureStage -ErrorMessage $failure.ErrorMessage)
        }
    }
}


###################################################
# MAIN
###################################################

if (-not (Test-Path -LiteralPath $LogPath)) {
    throw "Log file not found: $LogPath"
}

$prefs = Get-Content "$PSScriptRoot\Preferences.yaml" | ConvertFrom-Yaml

$run = Get-YardstickLogRun -Lines (Get-Content -LiteralPath $LogPath) -Index $RunIndex
$entries = ConvertTo-YardstickLogEntry -Lines $run.Lines
$apps = ConvertFrom-YardstickLog -Entries $entries

$succeeded = @($apps | Where-Object { $_.Succeeded })
$failed = @($apps | Where-Object { $_.Failures.Count -gt 0 -or (-not $_.Succeeded -and $_.ErrorLines.Count -gt 0) })

Write-Host "Parsed run started $($run.StartTime) from $LogPath"
Write-Host "  Recipes processed: $($apps.Count)"
Write-Host "  Successful: $($succeeded.Count)"
Write-Host "  Failed: $($failed.Count)"

if ($ParseOnly) {
    $succeeded | Select-Object Id, DisplayName, Version, Action, BackupStatus,
        @{ n = 'AutoUpdate'; e = { ($_.AutoParts | Select-Object -Unique) -join ', ' } },
        @{ n = 'Dependents'; e = { $_.Dependents.Count } } | Format-Table -AutoSize | Out-String | Write-Host
    $failed | Select-Object Id, DisplayName, Version,
        @{ n = 'Error'; e = { (($_.Failures | Select-Object -First 1).ErrorMessage -split "`n")[0] } } |
        Format-Table -AutoSize | Out-String | Write-Host
    return
}

if ($succeeded.Count -eq 0 -and $failed.Count -eq 0) {
    Write-Host "Nothing to report - no application results found in this run."
    return
}

Publish-YardstickTrackerFromLog -Apps $apps

if (-not $RunParameters) {
    $RunParameters = "Reconstructed from $(Split-Path $LogPath -Leaf) (run started $($run.StartTime))"
}

$reportArgs = @{
    Preferences   = $prefs
    RunParameters = $RunParameters
}
if ($HtmlOutputPath) { $reportArgs.HtmlOutputPath = $HtmlOutputPath }
if ($Preview) { $reportArgs.Preview = $true }

if ($Preview) {
    Write-Host "Opening the reconstructed report in Outlook for preview..."
} else {
    Write-Host "Sending the reconstructed report..."
}

Send-YardstickEmailReport @reportArgs

if ($HtmlOutputPath) {
    Write-Host "Browser preview written to $HtmlOutputPath"
}
