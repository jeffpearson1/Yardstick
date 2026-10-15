#Requires -Version 7.0

<#
Root module for Yardstick. Import-Module .\Yardstick.psd1 loads every component
module once and adds three commands: Invoke-Yardstick, Install-Yardstick and
Test-YardstickSetup.

The components are imported into the global session state, as the scripts have
always done. Recipes, the Custom helper modules and the backup and watchdog
threads all call component functions (Write-Log above all) through the global
scope, so they cannot be private to this module yet. What changes is that each
loads exactly once: components import each other without -Force, which reuses
the loaded copy instead of replacing it - replacing YardstickGraph also threw
away its cached access token.
#>

$script:ComponentModules = @(
    'YardstickGraph'
    'YardstickIntune'
    'YardstickSupport'
    'YardstickCredential'
    'YardstickConfig'
    'YardstickSetup'
)

foreach ($name in $script:ComponentModules) {
    Import-Module (Join-Path $PSScriptRoot "Modules\$name.psm1") -Global -ErrorAction Stop
}

# Removing (or force-reloading) Yardstick takes the components with it, so the
# next import loads fresh copies rather than reusing stale ones.
$ExecutionContext.SessionState.Module.OnRemove = {
    $names = @($script:ComponentModules)
    [array]::Reverse($names)
    foreach ($name in $names) {
        Get-Module -Name $name -All | Remove-Module -Force -ErrorAction SilentlyContinue
    }
}


function Get-YardstickScriptParameter {
    <#
    .SYNOPSIS
    Builds a dynamic parameter dictionary matching a script's own parameters.

    .DESCRIPTION
    Install-Yardstick and Test-YardstickSetup take exactly the parameters of the
    scripts they run. Copying them from the script at call time keeps the two in
    step without maintaining a second copy of each param block. Defaults are not
    copied; the script applies its own when a parameter is not passed.
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.RuntimeDefinedParameterDictionary])]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $common = [System.Management.Automation.PSCmdlet]::CommonParameters +
        [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
    $dictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
    foreach ($parameter in (Get-Command -Name $Path -CommandType ExternalScript).Parameters.Values) {
        if ($parameter.Name -in $common) { continue }
        $attributes = [System.Collections.ObjectModel.Collection[Attribute]]::new()
        foreach ($attribute in $parameter.Attributes) { $attributes.Add($attribute) }
        $dictionary.Add($parameter.Name,
            [System.Management.Automation.RuntimeDefinedParameter]::new($parameter.Name, $parameter.ParameterType, $attributes))
    }
    return $dictionary
}


function Invoke-Yardstick {
    <#
    .SYNOPSIS
    Adds and updates Win32 applications in Microsoft Intune.

    .DESCRIPTION
    The module form of .\Yardstick.ps1, with the same parameters. The run itself
    is Private\Invoke-YardstickRun.ps1, which stays a script because recipes
    assign its variables directly (see that script's help).

    Unlike the script, this leaves the current session as it found it: the
    current folder, the web request timeout defaults and any watchdog or backup
    threads are restored or stopped whichever way the run ends. A stuck recipe
    stage is logged but does not close the session; press Ctrl+C to stop it.

    A run that fails before processing any application (bad preferences, a
    missing recipe, Graph sign-in) ends with a terminating error. Failures of
    individual applications are reported in the log and the email report, as
    they always have been.

    .PARAMETER ApplicationId
    The ID (or list of IDs) of the application(s) to update.

    .PARAMETER Group
    The name of a group in RecipeGroups.yaml to update.

    .PARAMETER All
    Update every application in the recipes folder.

    .PARAMETER NoInteractive
    Skip recipes in the Interactive folder, and never prompt.

    .PARAMETER Force
    Replace applications even when they are already up to date.

    .PARAMETER NoDelete
    Keep old versions instead of pruning them after an update.

    .PARAMETER Repair
    Fix incorrectly named applications and run the full maintenance sweep on
    applications that have no update.

    .PARAMETER NoEmail
    Do not send the email report.

    .PARAMETER ProjectRoot
    The Yardstick folder holding preferences, recipes and the working folders.
    Defaults to the folder this module was loaded from.

    .EXAMPLE
    Invoke-Yardstick -ApplicationId 7zip -NoEmail

    .EXAMPLE
    Invoke-Yardstick -All -NoInteractive
    #>
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

        [Switch] $NoEmail,

        [String] $ProjectRoot = $PSScriptRoot
    )

    $arguments = @{} + $PSBoundParameters
    $arguments['ProjectRoot'] = $ProjectRoot

    # The runner sets these for every recipe's web calls. They belong to the run,
    # not to whoever called it.
    $timeoutKeys = 'Invoke-WebRequest:TimeoutSec', 'Invoke-RestMethod:TimeoutSec'
    $savedDefaults = @{}
    if ($null -ne $Global:PSDefaultParameterValues) {
        foreach ($key in $timeoutKeys) {
            if ($Global:PSDefaultParameterValues.Contains($key)) { $savedDefaults[$key] = $Global:PSDefaultParameterValues[$key] }
        }
    }

    $global:LASTEXITCODE = 0
    $exitCode = 1
    Push-Location
    try {
        & (Join-Path $PSScriptRoot 'Private\Invoke-YardstickRun.ps1') @arguments
        $exitCode = $global:LASTEXITCODE
    } finally {
        # An exit partway through the run skips the runner's own teardown. In a
        # process of its own that did not matter; here the session carries on, and a
        # watchdog thread left behind could end it hours later.
        Stop-YardstickWatchdog
        Wait-YardstickBackup -TimeoutSeconds 600 | Out-Null
        Pop-Location
        if ($null -ne $Global:PSDefaultParameterValues) {
            foreach ($key in $timeoutKeys) {
                if ($savedDefaults.ContainsKey($key)) {
                    $Global:PSDefaultParameterValues[$key] = $savedDefaults[$key]
                } else {
                    $Global:PSDefaultParameterValues.Remove($key)
                }
            }
        }
    }

    if ($exitCode -ne 0) {
        # Yardstick.ps1 hands this on as the process exit code.
        $global:LASTEXITCODE = $exitCode
        $log = Join-Path $Global:LogLocation $Global:LogFile
        throw "The Yardstick run stopped with exit code $exitCode. See $log for details."
    }
}


function Install-Yardstick {
    <#
    .SYNOPSIS
    Sets up Yardstick on this machine.

    .DESCRIPTION
    Runs the setup wizard, .\Install-Yardstick.ps1, with the same parameters.
    See Get-Help .\Install-Yardstick.ps1 -Full for the steps and parameters.

    .EXAMPLE
    Install-Yardstick

    .EXAMPLE
    Install-Yardstick -Step Email
    #>
    [CmdletBinding()]
    param()

    dynamicparam {
        Get-YardstickScriptParameter -Path (Join-Path $PSScriptRoot 'Install-Yardstick.ps1')
    }

    end {
        $global:LASTEXITCODE = 0
        & (Join-Path $PSScriptRoot 'Install-Yardstick.ps1') @PSBoundParameters
        if ($global:LASTEXITCODE -ne 0) {
            throw "Yardstick setup did not finish (exit code $global:LASTEXITCODE). See the messages above, or $(Join-Path $PSScriptRoot 'YInstall.log')."
        }
    }
}


function Test-YardstickSetup {
    <#
    .SYNOPSIS
    Checks that Yardstick is set up correctly on this machine.

    .DESCRIPTION
    Runs the health check, .\Test-YardstickSetup.ps1, with the same parameters.
    Failing checks are reported on screen and as a non-terminating error; use
    -PassThru for the individual results.

    .EXAMPLE
    Test-YardstickSetup

    .EXAMPLE
    Test-YardstickSetup -Offline -PassThru | Where-Object Level -eq 'Fail'
    #>
    [CmdletBinding()]
    param()

    dynamicparam {
        Get-YardstickScriptParameter -Path (Join-Path $PSScriptRoot 'Test-YardstickSetup.ps1')
    }

    end {
        $global:LASTEXITCODE = 0
        & (Join-Path $PSScriptRoot 'Test-YardstickSetup.ps1') @PSBoundParameters
        if ($global:LASTEXITCODE -ne 0) {
            Write-Error 'One or more Yardstick setup checks failed. See the results above.'
        }
    }
}


Export-ModuleMember -Function Invoke-Yardstick, Install-Yardstick, Test-YardstickSetup
