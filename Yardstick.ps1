<#
.DESCRIPTION
    Yardstick is a PowerShell script designed to automate the process of updating and managing Win32 applications in Microsoft Intune. It allows users to specify applications to update, either individually or in groups, and handles the downloading, packaging, and deployment of these applications.

.SYNOPSIS
    Adds and Updates Win32 applications in Microsoft Intune.

.PARAMETER ApplicationId
    The ID (or list of IDs) of the application(s) to update. Accepts a single string or a comma-separated array of strings.
    May only be used with -Force, -NoDelete, and -Repair parameters.

.PARAMETER Group
    The name of the group of applications to update. This parameter is used when updating a group of applications.
    May only be used with -NoInteractive, -Force, -NoDelete and -Repair parameters.

.PARAMETER All
    A switch parameter that indicates whether to update all applications in the repository. If specified, all applications will be processed.
    May only be used with -NoInteractive, -Force, -NoDelete and -Repair parameters.

.PARAMETER NoInteractive
    A switch parameter that, when specified, excludes interactive applications from the update process. This is denoted by recipes that are in the "Interactive" folder.
    May be used with all other parameters except ApplicationId.

.PARAMETER Force
    A switch parameter that forces the replacement of an application, even if it is already up-to-date.
    May be used with all other parameters.

.PARAMETER NoDelete
    A switch parameter that prevents the deletion of old versions of applications after an update. This is useful for debugging or testing purposes.
    May be used with all other parameters.

.PARAMETER NoEmail
    A switch parameter that suppresses the sending of email notifications after the update process is complete.
    May be used with all other parameters.

.PARAMETER Repair
    A switch parameter that repairs the application by renaming any incorrectly named applications to their correct format.
    It also runs the full maintenance sweep - retention pruning, assignment migration, supersedence and auto-update - on
    applications that have no update available, which a normal run skips entirely. Combine with -NoDelete to reconcile
    everything except deletions.
    May be used with all other parameters.

.EXAMPLE
    .\Yardstick.ps1 -All -Repair
    Updates everything that has a new version, and reconciles naming, retention, assignments and supersedence on
    everything that does not.

.EXAMPLE
    .\Yardstick.ps1 -All
    The recommended way to update all applications in the repository. This will update all applications.

.EXAMPLE
    .\Yardstick.ps1 -All -NoInteractive
    The recommended way to automate Yardstick. This will update all applications in the repository, excluding interactive applications.

.EXAMPLE
    .\Yardstick.ps1 -ApplicationId "ExampleApp" -Force -NoDelete
    This command updates the application with the ID "ExampleApp", forcing the update and preventing the deletion of old versions.

.EXAMPLE
    .\Yardstick.ps1 -ApplicationId "App1","App2","App3"
    This command updates three applications by their IDs.

.EXAMPLE
    .\Yardstick.ps1 -Group "ExampleGroup" -NoInteractive -Force
    This command updates all applications in the group "ExampleGroup", and forces the update.

.EXAMPLE
    .\Yardstick.ps1 -All -NoInteractive -Force -NoDelete
    This command updates all applications in the repository, excluding interactive applications, forcing the update and preventing the deletion of old versions.
#>

#Requires -Version 7.0

# This is a thin wrapper so existing command lines and the scheduled task keep
# working. The run itself is Private\Invoke-YardstickRun.ps1, called through
# Invoke-Yardstick in Yardstick.psm1.

# No set named is a SingleApp run with no -ApplicationId, which the runner answers
# with "Specify -ApplicationId, -Group or -All." rather than a binding error.
[CmdletBinding(DefaultParameterSetName = "SingleApp")]
param (
    [Alias("AppId", "AppIds")]
    [parameter(ParameterSetName="SingleApp")]
    [String[]] $ApplicationId,

    [parameter(ParameterSetName="GroupApps")]
    [String] $Group,

    [parameter(ParameterSetName="AllApps")]
    [Switch] $All,

    [parameter(ParameterSetName="GroupApps")]
    [parameter(ParameterSetName="AllApps")]
    [Switch] $NoInteractive,

    [parameter(ParameterSetName="SingleApp")]
    [parameter(ParameterSetName="GroupApps")]
    [parameter(ParameterSetName="AllApps")]
    [Switch] $Force,

    [parameter(ParameterSetName="SingleApp")]
    [parameter(ParameterSetName="GroupApps")]
    [parameter(ParameterSetName="AllApps")]
    [Switch] $NoDelete,

    [parameter(ParameterSetName="SingleApp")]
    [parameter(ParameterSetName="GroupApps")]
    [parameter(ParameterSetName="AllApps")]
    [Switch] $Repair,

    [Switch] $NoEmail
)

# -Force so a console that ran Yardstick before picks up any edits to the modules.
Import-Module (Join-Path $PSScriptRoot 'Yardstick.psd1') -Force

# The scheduled task reads the exit code: 0 on success, the runner's code otherwise.
try {
    Invoke-Yardstick @PSBoundParameters -ProjectRoot $PSScriptRoot
    exit 0
} catch {
    $exitCode = if ($LASTEXITCODE) { $LASTEXITCODE } else { 1 }
    Write-Error -ErrorRecord $_ -ErrorAction Continue
    exit $exitCode
}
