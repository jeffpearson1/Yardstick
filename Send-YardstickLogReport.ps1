<#
.SYNOPSIS
Rebuilds the Yardstick email report from YLog.log and sends it after the fact.

.DESCRIPTION
Yardstick normally emails its summary at the end of a run, straight out of the
in-memory trackers. When that email is skipped (-NoEmail, a declined prompt, a
delivery failure, or an aborted run) the results are gone, but the log still has
everything the report needs. This script replays a run out of YLog.log,
repopulates the same trackers Yardstick.ps1 uses, and hands them to
Send-YardstickEmailReport, which delivers through the configured
emailDeliveryMethod.

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
Open the message for inspection instead of sending it - in Outlook, or in the
default browser for SMTP delivery.

.PARAMETER HtmlOutputPath
Also write the rendered HTML body here for browser preview.

.PARAMETER ParseOnly
Print what was parsed out of the log and exit without sending anything.

.EXAMPLE
.\Send-YardstickLogReport.ps1 -ParseOnly

.EXAMPLE
.\Send-YardstickLogReport.ps1 -Preview -HtmlOutputPath .\LogReportPreview.html

.EXAMPLE
.\Send-YardstickLogReport.ps1 -LogPath D:\Archive\YLog-20260908.log
#>

#Requires -Version 7.0

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

Import-Module "$PSScriptRoot\Yardstick.psd1" -Force
Import-Module powershell-yaml -Force

# The log parser (Get-YardstickLogRun, ConvertFrom-YardstickLog, ...) lives in
# Modules/YardstickLog.psm1, next to the Write-Log that produces the format.


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
            # Newer logs name the stage; older ones only have the message to go on.
            $stage = if ($failure.Stage) { $failure.Stage } else { Get-YardstickFailureStage -ErrorMessage $failure.ErrorMessage }
            Add-FailedApplication -ApplicationId $failure.Id -DisplayName $displayName -Version $version `
                -ErrorMessage $failure.ErrorMessage -FailureStage $stage
        }
    }
}


###################################################
# MAIN
###################################################

if (-not (Test-Path -LiteralPath $LogPath)) {
    throw "Log file not found: $LogPath"
}

$prefs = Import-YardstickPreferences -ProjectRoot $PSScriptRoot

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
    Write-Host "Opening the reconstructed report for preview..."
} else {
    Write-Host "Sending the reconstructed report..."
}

Send-YardstickEmailReport @reportArgs

if ($HtmlOutputPath) {
    Write-Host "Browser preview written to $HtmlOutputPath"
}
