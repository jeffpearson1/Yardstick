<#
.SYNOPSIS
    The body of a Yardstick run. Start runs with .\Yardstick.ps1 or Invoke-Yardstick, not this script.

.DESCRIPTION
    Recipes are YAML scriptblocks that this script runs with Invoke-Command -NoNewScope. They assign
    runner variables such as $version, $url and $fileName without a scope qualifier, and read $Temp,
    $BuildSpace and $Tools the same way, so they must run at the top level of a script. Inside a module
    function those assignments would land in the function's scope and every recipe would silently stop
    working. That is why this is a script rather than a module function, and why Invoke-Yardstick calls
    it with & instead of dot-sourcing or inlining it.

    exit here ends only this script; Invoke-Yardstick turns a nonzero exit code into an error, and
    Yardstick.ps1 turns that back into the process exit code.

    Parameters are those of Yardstick.ps1 (see Get-Help .\Yardstick.ps1), plus -ProjectRoot.

.PARAMETER ProjectRoot
    The Yardstick folder holding preferences, recipes and the working folders. Defaults to the folder
    containing this module.
#>

#Requires -Version 7.0
using module ..\Modules\VersionPro.psm1

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

    [Switch] $NoEmail,

    [String] $ProjectRoot = (Split-Path -Parent $PSScriptRoot)
)

# So we can get redirected URLs
Add-Type -AssemblyName Microsoft.PowerShell.Commands.Utility

# Ensure non-terminating errors are caught by try/catch blocks
$ErrorActionPreference = 'Stop'

# A full filesystem path, since it is compared with (Get-Location).ProviderPath and
# handed to .NET APIs that cannot open PowerShell's provider-qualified UNC form.
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).ProviderPath

# Where the code lives, as opposed to the data. The same folder unless -ProjectRoot
# says otherwise.
$YardstickModuleRoot = (Resolve-Path -LiteralPath (Split-Path -Parent $PSScriptRoot)).ProviderPath

# Constants
$Global:LogLocation = "$ProjectRoot"
$Global:LogFile = "YLog.log"

# Invoke-Yardstick has already loaded the Yardstick module. Load it here as well
# so the script still works when someone runs it directly.
if (-not (Get-Command Invoke-Yardstick -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $YardstickModuleRoot 'Yardstick.psd1')
}

# Recipe helpers. Recipes reach these through the global scope, and they are
# re-imported each run so edits to them are picked up without a new session.
$CustomModuleImports = Get-ChildItem -Path (Join-Path $YardstickModuleRoot 'Modules\Custom\*.psm1')
foreach ($Module in $CustomModuleImports) {
    Import-Module "$($Module.FullName)" -Scope Global -Force
}

# Initialize script variables
$Script:Applications = [System.Collections.Generic.List[PSObject]]::new()

# Recipes import helpers with "using module .\Modules\...", which resolves against
# the current folder, so the run works from the module folder. Invoke-Yardstick
# puts the caller back where they were afterwards.
Set-Location $YardstickModuleRoot

# Initialize the log file. -Verbose shows the routine (Detail) log lines on the
# console; it is not meant to turn on every cmdlet's own verbose output.
$showDetail = $VerbosePreference -eq 'Continue'
$VerbosePreference = 'SilentlyContinue'
Write-Log -Init -ShowDetail:$showDetail

$yardstickVersion = try { (Import-PowerShellDataFile -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'Yardstick.psd1')).ModuleVersion } catch { 'unknown' }
Write-Log "Yardstick $yardstickVersion, PowerShell $($PSVersionTable.PSVersion)." -Level Detail
# Session 0 is a task running while nobody is signed in, where Outlook and some
# network shares are out of reach - worth knowing when a scheduled run misbehaves.
Write-Log ("Running as {0} in Windows session {1}." -f [Security.Principal.WindowsIdentity]::GetCurrent().Name, (Get-Process -Id $PID).SessionId) -Level Detail
if (Test-YardstickUacElevated) {
    Write-Log "This run has administrator rights, so the files it creates (this log, detection scripts) belong to the Administrators group and cannot be changed later without them. Run Yardstick from a normal PowerShell window, and its scheduled task only when you are signed in." -Level Warning
}

# Validate parameters
if ((-not $ApplicationId) -and (($All -eq $false) -and (!$Group))) {
    Write-Log "Specify -ApplicationId, -Group or -All." -Level Error
    exit 1
}

# powershell-yaml is needed to read the preferences file itself, so it is checked
# here rather than waiting for Test-Prerequisites below.
if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
    Write-Log "The PowerShell module 'powershell-yaml' is not installed. Install it with: Install-Module powershell-yaml -Scope CurrentUser (or run .\Install-Yardstick.ps1)." -Level Error
    exit 1
}

# Import preferences file. Absent or blank keys take built-in defaults, relative
# folder paths are anchored to this folder, and the working folders are created.
try {
    $Prefs = Import-YardstickPreferences -ProjectRoot $ProjectRoot -CreateFolders
} catch {
    Write-Log "Unable to load Yardstick preferences: $($_.Exception.Message)" -Level Error
    exit 1
}
$prefsResult = Test-YardstickPreferences -Preferences $Prefs
foreach ($w in $prefsResult.Warnings) { Write-Log $w -Level Warning }
if (-not $prefsResult.IsValid) {
    foreach ($e in $prefsResult.Errors) { Write-Log $e -Level Error }
    Write-Log "Preference checks failed. Fix preferences.yaml (or run .\Test-YardstickSetup.ps1) and try again." -Level Error
    exit 1
}

# Run-safety timeouts. These are read before anything else touches the network on
# purpose: every value here exists to stop the runner waiting forever, so they have
# to be in force for the whole run, not just the recipe loop.
$Script:WebRequestTimeoutSeconds = Get-YardstickTimeoutPreference -Value $Prefs.webRequestTimeoutSeconds -Default 120 -Name 'webRequestTimeoutSeconds' -AllowZero
$Script:RecipeStageTimeoutMinutes = Get-YardstickTimeoutPreference -Value $Prefs.recipeStageTimeoutMinutes -Default 60 -Name 'recipeStageTimeoutMinutes' -AllowZero
$Script:PackagingTimeoutMinutes = Get-YardstickTimeoutPreference -Value $Prefs.packagingTimeoutMinutes -Default 30 -Name 'packagingTimeoutMinutes'
$Script:BitsStallTimeoutMinutes = Get-YardstickTimeoutPreference -Value $Prefs.bitsStallTimeoutMinutes -Default 15 -Name 'bitsStallTimeoutMinutes'

Set-YardstickWebRequestTimeout -TimeoutSeconds $Script:WebRequestTimeoutSeconds

# Shared content: with syncSharedContent on, pull the team's icons and recipe
# overrides from the share before any recipe is read. An unreachable share only
# warns - the run carries on with the icons it has and the last synced overrides.
$SharedOverridesCachePath = $null
if ($Prefs.syncSharedContent -eq $true) {
    if (-not [string]::IsNullOrWhiteSpace($Prefs.SharedIcons)) {
        try {
            $iconSync = Sync-YardstickSharedIcons -Source $Prefs.SharedIcons -Destination $Prefs.Icons
            Write-Log "Synced $($iconSync.Copied.Count) icon(s) from $($Prefs.SharedIcons); $($iconSync.Unchanged) unchanged."
            foreach ($name in $iconSync.Copied) { Write-Log "Copied icon $name from the shared icon folder." -Level Detail }
            foreach ($failure in $iconSync.Failed) {
                Write-Log "Could not copy icon $($failure.Name) from $($Prefs.SharedIcons): $($failure.Error)" -Level Warning
            }
        } catch {
            Write-Log "Could not sync icons from $($Prefs.SharedIcons): $($_.Exception.Message) Using the icons already in $($Prefs.Icons)." -Level Warning
        }
    }

    $SharedOverridesCachePath = Get-YardstickSharedRecipeOverridesCachePath -ProjectRoot $ProjectRoot
    if (-not [string]::IsNullOrWhiteSpace($Prefs.SharedRecipeOverrides)) {
        try {
            $entryCount = Sync-YardstickSharedRecipeOverrides -Source $Prefs.SharedRecipeOverrides -CachePath $SharedOverridesCachePath
            Write-Log "Synced $entryCount shared recipe override(s) from $($Prefs.SharedRecipeOverrides)."
        } catch {
            $cache = Get-Item -LiteralPath $SharedOverridesCachePath -ErrorAction SilentlyContinue
            $fallback = if ($cache) {
                "Using the copy synced on $($cache.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))."
            } else {
                'There is no synced copy yet, so only the local overrides apply.'
            }
            Write-Log "Could not sync recipe overrides: $($_.Exception.Message) $fallback" -Level Warning
        }
    }
}

# Site-specific recipe fields, such as scope tags, live in Local\recipeOverrides.yaml
# so the shared recipes work in any tenant, layered over the synced shared copy when
# shared sync is on. A bad file stops the run: dropping an override would upload
# apps with the wrong scope tags.
$RecipeOverridesPath = Get-YardstickRecipeOverridesPath -ProjectRoot $ProjectRoot
try {
    $overrideSet = Import-YardstickEffectiveRecipeOverrides -Path $RecipeOverridesPath -SharedPath $SharedOverridesCachePath
    $RecipeOverrides = $overrideSet.Overrides
    if ($overrideSet.Sources.Count -gt 1) {
        Write-Log "Applying $($RecipeOverrides.Count) recipe override(s) from $($overrideSet.Sources -join ' and ')." -Level Detail
    }
} catch {
    Write-Log "Unable to load recipe overrides from ${RecipeOverridesPath}: $($_.Exception.Message)" -Level Error
    exit 1
}

# When each recipe last contacted its vendor, read by recipes that set minRunInterval.
$RecipeRunHistoryPath = Get-YardstickRecipeRunHistoryPath -ProjectRoot $ProjectRoot

# Validate prerequisites before importing external modules
$prereqResult = Test-Prerequisites -ToolsPath $Prefs.Tools
foreach ($w in $prereqResult.Warnings) { Write-Log $w -Level Warning }
if (-not $prereqResult.IsValid) {
    foreach ($e in $prereqResult.Errors) { Write-Log $e -Level Error }
    Write-Log "Prerequisite checks failed." -Level Error
    exit 1
}

# Selenium 4's ScriptsToProcess defines PowerShell classes/enums in the caller's scope,
# but the module's own session state cannot see them during parameter binding. Defining
# the types via Add-Type (C#) BEFORE importing the module places them in the .NET
# AppDomain where they are visible to every session state, including the module's.
if (-not ('ValidateURIAttribute' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Management.Automation;

public class ValidateURIAttribute : ValidateArgumentsAttribute {
    protected override void Validate(object arguments, EngineIntrinsics engineIntrinsics) {
        Uri outUri;
        if (Uri.TryCreate(arguments != null ? arguments.ToString() : "", UriKind.Absolute, out outUri)) { return; }
        throw new ValidationMetadataException("Incorrect StartURL please make sure the URL starts with http:// or https://");
    }
}
"@
}

if (-not ('SeBrowsers' -as [type])) {
    Add-Type -TypeDefinition @"
public enum SeBrowsers {
    Chrome,
    Edge,
    Firefox,
    InternetExplorer,
    MSEdge
}
"@
}

if (-not ('SeWindowState' -as [type])) {
    Add-Type -TypeDefinition @"
public enum SeWindowState {
    Headless,
    Default,
    Minimized,
    Maximized,
    Fullscreen
}
"@
}

if (-not ('SeBySelector' -as [type])) {
    Add-Type -TypeDefinition @"
public enum SeBySelector {
    ClassName,
    CssSelector,
    Id,
    LinkText,
    PartialLinkText,
    Name,
    TagName,
    XPath
}
"@
}

if (-not ('SeBySelect' -as [type])) {
    Add-Type -TypeDefinition @"
public enum SeBySelect {
    Index,
    Text,
    Value
}
"@
}

# Import external modules (after Selenium compat types are in the AppDomain)
Import-Module powershell-yaml -Scope Local
Import-Module Selenium -Scope Global -ErrorAction SilentlyContinue
Import-Module TUN.CredentialManager -Scope Local -ErrorAction SilentlyContinue



function Clear-RecipeVariableConstraint {
    <#
    .SYNOPSIS
    Removes type constraints left behind in script scope by a previous recipe.

    .DESCRIPTION
    Recipe scripts run via Invoke-Command -NoNewScope, so they share one variable
    scope for the entire run. A recipe that writes a left-side cast such as
    [xml]$manifest = ... attaches a permanent type constraint to that variable.
    The constraint outlives the recipe, and every later recipe that assigns a
    non-XML value to $manifest throws for the rest of the process -- which looks
    like a batch of unrelated recipes failing in a row, then succeeding when run
    individually in a fresh process.

    Removing the variable clears the constraint along with the value. Only
    type-constrained variables that recipes themselves introduced are removed,
    guarded two ways:

    1. The first call snapshots every variable name that already exists and
       removes nothing. Later calls only consider names absent from that snapshot.
    2. PowerShell's own preference variables ($ErrorActionPreference,
       $ProgressPreference, $PSDefaultParameterValues, ...) are type-constrained
       by design and are never removed. The snapshot alone is not enough here:
       they are materialised into script scope lazily, so they can be missing
       from the snapshot and present later, which would delete them and change
       the runner's behaviour mid-run.

    Accumulators such as $Script:FailedApplications live in the YardstickSupport
    module's own scope and are not visible here, so they are unaffected.
    #>
    $protected = @(
        # Runner-owned
        'Applications', 'Prefs', 'RecipeOverrides', 'Temp', 'BuildSpace', 'Published',
        'RecipeVariableBaseline',
        # PowerShell automatic variables that carry a type constraint
        'OutputEncoding', 'PSDefaultParameterValues', 'ErrorView',
        'FormatEnumerationLimit', 'MaximumHistoryCount', 'PSEmailServer',
        'PSSessionOption', 'PSSessionApplicationName', 'PSSessionConfigurationName',
        'PSModuleAutoLoadingPreference'
    )

    # First call establishes the baseline: nothing here came from a recipe.
    if (-not $Script:RecipeVariableBaseline) {
        $Script:RecipeVariableBaseline = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]@(Get-Variable -Scope Script | Select-Object -ExpandProperty Name),
            [System.StringComparer]::OrdinalIgnoreCase
        )
        return
    }

    foreach ($variable in Get-Variable -Scope Script) {
        if ($variable.Name -in $protected) { continue }
        # Covers $ErrorActionPreference, $ProgressPreference, $WarningPreference, etc.
        if ($variable.Name -like '*Preference') { continue }
        if ($Script:RecipeVariableBaseline.Contains($variable.Name)) { continue }
        if (-not $variable.Attributes.Count) { continue }

        # ArgumentTypeConverterAttribute is internal to PowerShell, so it cannot be
        # referenced as a type literal here -- match on the type name instead.
        $isConstrained = $variable.Attributes | Where-Object {
            $_.GetType().Name -eq 'ArgumentTypeConverterAttribute'
        }
        if (-not $isConstrained) { continue }

        Write-Log "Clearing leaked type constraint on `$$($variable.Name) from a previous recipe." -Level Detail
        Remove-Variable -Name $variable.Name -Scope Script -Force -ErrorAction SilentlyContinue
    }
}


function Restore-RunnerPathVariable {
    <#
    .SYNOPSIS
    Re-asserts the runner-owned folder variables that recipes can overwrite.

    .DESCRIPTION
    The folder locations are read from preferences once at startup, but recipe
    scripts run via Invoke-Command -NoNewScope and therefore share this scope.
    PowerShell variable names are case-insensitive, so a recipe that writes an
    innocuous-looking $temp overwrites the runner's own $Temp for the rest of the
    run. That is exactly what hxd did: it pointed $Temp at a folder under the
    user's TEMP and then deleted it in a finally block, which broke every later
    recipe that reads $Temp and silently stopped the per-recipe temp clear.

    Clear-RecipeVariableConstraint protects these names from removal but cannot
    restore a value, so snapshot them on the first call and put them back at the
    top of every recipe. A recipe clobber then lasts only for that recipe.
    #>
    $names = @(
        'Temp', 'BuildSpace', 'Scripts', 'Published', 'Recipes',
        'Icons', 'Tools', 'Secrets', 'SoftwareDropbox', 'SoftwareArchive'
    )

    if (-not $Script:RunnerPathBaseline) {
        $Script:RunnerPathBaseline = [ordered]@{}
        foreach ($name in $names) {
            $Script:RunnerPathBaseline[$name] = Get-Variable -Name $name -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        }
        return
    }

    foreach ($name in $names) {
        $expected = $Script:RunnerPathBaseline[$name]
        $current = Get-Variable -Name $name -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        # Compare as strings on purpose: these are all paths, and a recipe that put
        # an array here would make -eq return a filtered collection, which an if()
        # reads as true and would silently skip the restore.
        if ([string]$current -ceq [string]$expected) { continue }
        Write-Log "Restoring `$${name}: the recipe overwrote it with '$current'." -Level Warning
        Set-Variable -Name $name -Value $expected -Scope Script
    }
}


function Invoke-YardstickAppMaintenance {
    <#
    .SYNOPSIS
    Reconciles every version of the current recipe's app in Intune: anchor,
    (N-x) naming, assignment migration, retention pruning, supersedence and
    auto-update.

    .DESCRIPTION
    Runs after a successful upload, and - under -Repair - on runs where the
    vendor shipped nothing new. The desired state it enforces does not depend on
    a new version existing, but it used to be reachable only through the upload
    path, so a version that failed to prune (blocked by a dependent, refused by
    Intune, or stranded by a bug) got exactly one retry: the next release. Apps
    that update rarely accumulated (N-2)/(N-3) versions indefinitely.

    Every step is idempotent - renames are no-ops when the name is already right,
    assignment migration re-reads both ends, and Set-YardstickSupersedence
    replaces the parent's whole target set - so a sweep over an already-clean app
    changes nothing.

    Declared at script scope rather than in the module so the recipe's $Script:*
    variables and the caller's per-recipe state ($ProtectedDependencyAppIds,
    $DependentLinkOptions, $DependentUpdateStatus, $NoDelete) stay readable
    without threading twenty parameters through.

    .PARAMETER NewAppId
    The id of the app this run just uploaded. Omitted on a maintenance sweep,
    where the newest version already in Intune takes its place.

    .OUTPUTS
    A summary object, or $null when a sweep could not read the app list.
    #>
    param(
        [string]$NewAppId
    )

    # Refresh the version list from Intune (retry in case a new upload has not
    # yet propagated to Get-YardstickWin32App list results).
    Write-Log "Refreshing the list of versions in Intune..." -Level Detail
    if ($NewAppId) { Start-Sleep -Seconds 4 }
    try {
        $AllMatchingApps = Get-SameAppAllVersions $Script:DisplayName
        if ($NewAppId -and !($AllMatchingApps | Where-Object id -eq $NewAppId)) {
            Write-Log "The new app is not in the version list yet. Retrying in 5 seconds..." -Level Detail
            Start-Sleep -Seconds 5
            $AllMatchingApps = Get-SameAppAllVersions $Script:DisplayName
        }
        if ($NewAppId -and !($AllMatchingApps | Where-Object id -eq $NewAppId)) {
            # The list endpoint lags behind creation. Fetch the new app directly
            # by id and splice it in - skipping here would leave a freshly
            # uploaded app with no assignments and no supersedence, and the next
            # run would consider the version already published and skip it.
            Write-Log "The new app is still not in the version list. Fetching it directly by id." -Level Detail
            $DirectApp = Get-YardstickWin32App -Id $NewAppId -ErrorAction Stop
            if (-not $DirectApp) {
                throw "Intune returned no application for id $NewAppId"
            }
            $AllMatchingApps = @($DirectApp) + @($AllMatchingApps)
        }
    } catch {
        if (-not $NewAppId) {
            # Nothing was published this run, so there is no half-finished app to
            # protect. Give up on this recipe rather than the whole run.
            Write-Log "Could not fetch the existing versions of $($Script:DisplayName) for the maintenance sweep: $_" -Level Error
            return $null
        }
        Write-Log "Could not fetch the existing versions of $($Script:DisplayName) after uploading $NewAppId. Exiting: $_" -Level Error
        # exit does not wait for background threads, so drain the backup queued
        # by the caller before tearing the process down.
        Wait-YardstickBackup -TimeoutSeconds 120 | Out-Null
        Exit 4
    }

    if ($NewAppId) {
        $CurrentApp = $AllMatchingApps | Where-Object id -eq $NewAppId | Select-Object -First 1
    }
    else {
        # The current version is the one holding the bare display name; every
        # other kept version is numbered (N-x) and the anchor holds the reserved
        # Ω DETECT - name. Fall back to the highest version when naming is broken,
        # which is exactly the state a sweep is meant to repair.
        $anchorNameForPick = Get-DetectAnchorName -DisplayName $Script:DisplayName
        $CurrentApp = $AllMatchingApps | Where-Object DisplayName -eq $Script:DisplayName | Select-Object -First 1
        if (-not $CurrentApp) {
            $CurrentApp = $AllMatchingApps |
                Where-Object DisplayName -ne $anchorNameForPick |
                Sort-Object @{Expression = {[VersionPro]$_.displayVersion}; Descending = $true} |
                Select-Object -First 1
        }
        if (-not $CurrentApp) {
            Write-Log "No published version of $($Script:DisplayName) to maintain."
            return $null
        }
        Write-Log "Sweeping the versions of $($Script:DisplayName) against $($CurrentApp.DisplayName) $($CurrentApp.displayVersion)..."
    }

    $OtherApps = @($AllMatchingApps | Where-Object id -ne $CurrentApp.id)

    # 1. Anchor identification. Version-detection recipes (msi, file-by-
    #    version, registry-by-version) get one sticky Ω DETECT - anchor so a
    #    fast-moving app cannot prune away every version that could still
    #    detect a stale endpoint. The anchor is a permanent supersedence
    #    target: its detection rule matches a very wide version range, so
    #    superseding it is what actually pulls ancient installs forward.
    #
    #    Recognising an EXISTING anchor is unconditional, and deliberately so.
    #    Get-SameAppAllVersions returns the anchor by name whether or not this
    #    recipe is eligible to pin one, so gating the lookup on eligibility let
    #    the anchor fall through into $OtherApps as an ordinary old version.
    #    It is always the lowest version, so it sorted to the end of $ToKeep and
    #    got renamed "<DisplayName> (N-2)" under the default retention of 3 -
    #    and worse, it could then be pruned, or superseded with 'Replace'
    #    against its intentionally wide detection rule. That hit every recipe
    #    whose detection is not version-based (fileDetectionMethod: exists,
    #    detectionType: script, ...) and every recipe run after
    #    useDetectAnchor was turned off. Only the act of PINNING a new anchor
    #    is gated below; an anchor that already exists is always protected.
    $anchorName = Get-DetectAnchorName -DisplayName $Script:DisplayName
    $Anchor = $OtherApps | Where-Object DisplayName -eq $anchorName | Select-Object -First 1
    $anchorStatus = $null
    $anchorPinned = $false
    if ($Anchor) {
        $anchorStatus = "existing ($($Anchor.displayVersion))"
    }
    elseif ($Script:UseDetectAnchor -and (Test-IsVersionDetection -DetectionType $Script:DetectionType -FileDetectionMethod $Script:FileDetectionMethod -RegistryDetectionMethod $Script:RegistryDetectionMethod)) {
        if ($OtherApps.Count -gt 0) {
            # First run under the new model - pin the oldest surviving version.
            $AnchorCandidate = $OtherApps | Sort-Object @{Expression = {[VersionPro]$_.displayVersion}} | Select-Object -First 1
            Set-DetectAnchor -App $AnchorCandidate -DisplayName $Script:DisplayName
            # Re-fetch so the object reflects the new name. Fall back to the
            # candidate if Intune has not caught up yet - otherwise the app we
            # just pinned would be renamed straight back to (N-x) below.
            $Anchor = Get-DetectAnchor -DisplayName $Script:DisplayName
            if (-not $Anchor) { $Anchor = $AnchorCandidate }
            $anchorStatus = "pinned ($($Anchor.displayVersion))"
            $anchorPinned = $true
        } else {
            $anchorStatus = "n/a (no prior versions)"
        }
    }

    # 2. The anchor is off-limits to renaming (step 5) and pruning (step 8): it
    #    holds a reserved name forever, so it must never be numbered (N-x), and it
    #    must not count against NumVersionsToKeep. It is NOT off-limits to
    #    assignment migration - step 6 re-adds it to that sweep explicitly. The
    #    anchor is pinned from the oldest surviving version, which on the first run
    #    under this model is usually the previously-CURRENT app, so it carries that
    #    version's live assignments. Dropping it here and never re-adding it
    #    stranded them on an app that is never renamed, never pruned and never
    #    revisited.
    if ($Anchor) {
        $OtherApps = @($OtherApps | Where-Object id -ne $Anchor.id)
    }

    # 3. Same-version collision - a prior run (or -Force) can leave an app at
    #    the same version. Those apps are held aside: their assignments and
    #    dependencies still migrate in step 6, and they are deleted with the
    #    rest of the prunable versions in step 8. They are never renamed to
    #    (N-x) or superseded, because their version is not actually older.
    $SameVersionApps = @($OtherApps | Where-Object displayVersion -eq $CurrentApp.displayVersion)
    $SameVersionIds = @($SameVersionApps | Select-Object -ExpandProperty id)
    if ($SameVersionApps) {
        Write-Log "Found $($SameVersionApps.Count) existing app(s) already at version $($CurrentApp.displayVersion); they will be retired after their assignments are migrated."
        $OtherApps = @($OtherApps | Where-Object displayVersion -ne $CurrentApp.displayVersion)
    }

    # 4. Compute kept vs prunable. `NumVersionsToKeep` includes the current app,
    #    so we keep (NumVersionsToKeep - 1) of the older ones.
    $Sorted = @($OtherApps | Sort-Object @{Expression = {[VersionPro]$_.displayVersion}; Descending = $true})
    $ToKeepCount = [Math]::Max(0, $Script:NumVersionsToKeep - 1)
    $ToKeep = @()
    $ToPrune = @()
    if ($Sorted.Count -gt 0) {
        $ToKeep  = @($Sorted | Select-Object -First $ToKeepCount)
        $ToPrune = @($Sorted | Select-Object -Skip $ToKeepCount)
    }
    # Duplicates of the current version always get retired, regardless of retention.
    $ToPrune = @($SameVersionApps) + @($ToPrune)
    $ToPruneIds = @($ToPrune | Select-Object -ExpandProperty id)

    # 5. Rename kept apps to (N-1), (N-2), ...
    $renamedCount = 0
    for ($i = 0; $i -lt $ToKeep.Count; $i++) {
        $targetName = "$($Script:DisplayName) (N-$($i + 1))"
        if ($ToKeep[$i].DisplayName -ne $targetName) {
            try {
                Set-YardstickWin32App -Id $ToKeep[$i].Id -DisplayName $targetName | Out-Null
                Write-Log "Renamed $($ToKeep[$i].DisplayName) to $targetName."
                $renamedCount++
            } catch {
                Write-Log "Could not rename $($ToKeep[$i].DisplayName) to ${targetName}: $_" -Level Error
            }
        }
    }

    # 6. Intent-based assignment handling.
    #    - Required: MOVE from older versions to newest, as Yardstick has always
    #      done - required deployments install unconditionally, so consolidating
    #      them on a single app is correct.
    #    - Available: COPY to the newest and leave the source assignment in
    #      place, but only when the source will actually end up superseded.
    #      Intune builds the auto-update component on the device when a user
    #      installs from Company Portal, and documents that "any application
    #      assignment changes delete the component responsible for auto-updating
    #      the app" - so removing the available assignment from a superseded
    #      version would break auto-update for exactly the devices we want to
    #      update. Where supersedence will NOT be attached, copying would just
    #      leave a duplicate Company Portal listing with no update path, so we
    #      fall back to moving.
    #    Dependencies migrate exactly once per source app (during the required
    #    pass); the available pass runs with -SkipDependencies to avoid
    #    double-processing them. Dependencies that point at a version being
    #    pruned this run are dropped rather than carried onto the new app - the
    #    fresh link would block that version's own deletion in step 8.
    #
    #    The Ω DETECT - anchor is swept here too, on every run, even though steps 5
    #    and 8 leave it alone. Step 2 removed it from $OtherApps to keep it out of
    #    the (N-x) numbering and the retention count, not because its assignments
    #    are special - and excluding it from this loop is what stranded the
    #    assignments of any app whose only prior version got pinned.
    #    It is appended LAST so a (N-x) version wins any duplicate-target race:
    #    Intune rejects the second add for a target with "already exists", so the
    #    first assignment to land is the one whose notification and install-time
    #    settings survive, and a kept version's schedule is more current than the
    #    anchor's. The generic -CopyOnly:$willBeSuperseded rule below resolves
    #    correctly for it without a special case - the anchor is never in $ToPrune,
    #    so $isDoomed is false and the flag collapses to $Script:Supersedence, which
    #    is exactly "step 9 will supersede this app". With supersedence off there is
    #    no auto-update component to protect and no update path, so falling back to
    #    a move is right there too.
    #
    #    Dependencies are skipped for the anchor. Unlike an (N-x) version there is
    #    no deletion to unblock - the anchor is never pruned, so a dependent app
    #    pointing at it is not holding anything hostage. And its child dependency
    #    set is frozen at whatever the recipe declared when it was pinned;
    #    Add-YardstickWin32AppDependency replaces the target's whole set from a merged
    #    list, so migrating it would re-merge that stale set onto the current app on
    #    every run, resurrecting dependencies the recipe has since dropped.
    $allOlder = @($ToKeep) + @($ToPrune)
    # $Anchor can be the raw pre-rename candidate (see step 1) when Intune's lookup
    # is still stale. Only .id and .DisplayName are read, and both objects come from
    # Get-YardstickWin32App, so the shapes are interchangeable. The id guard is cheap
    # insurance: From -eq To would add the assignment, read it back as present, then
    # delete it off the app it was just confirmed on.
    if ($Anchor -and ($Anchor.id -ne $CurrentApp.id)) { $allOlder += $Anchor }
    foreach ($old in $allOlder) {
        $isDoomed = $ToPruneIds -contains $old.id
        $isAnchor = ($null -ne $Anchor) -and ($old.id -eq $Anchor.id)
        Write-Log "Moving assignments and dependencies from $($old.id) to $($CurrentApp.id)..."
        try {
            Move-AssignmentsAndDependencies -From $old -To $CurrentApp `
                -AvailableDateOffset $Script:AvailableDateOffset `
                -DeadlineDateOffset $Script:DeadlineDateOffset `
                -IntentFilter 'required' `
                -AllowDependentLinkUpdates $Script:AllowDependentLinkUpdates `
                -DependentLinkOptions $DependentLinkOptions `
                -DependentUpdateStatus $DependentUpdateStatus `
                -ProtectedSourceIds $ProtectedDependencyAppIds `
                -ExcludeDependencyTargetIds $ToPruneIds `
                -SkipDependencies:$isAnchor
        } catch {
            Write-Log "Could not move the required assignments from $($old.DisplayName): $_" -Level Error
        }

        $availOnOld = @(Get-YardstickAppAssignment -Id $old.id | Where-Object Intent -eq 'available')
        if ($availOnOld.Count -gt 0) {
            # A pruned app is never a supersedence target - step 9 only supersedes
            # the versions that survive - so there is no auto-update path to
            # protect and copying would just strand the assignment on an app
            # about to be deleted. Move it instead.
            $willBeSuperseded = $Script:Supersedence -and (-not $isDoomed)
            try {
                Move-AssignmentsAndDependencies -From $old -To $CurrentApp `
                    -AvailableDateOffset $Script:AvailableDateOffset `
                    -DeadlineDateOffset $Script:DeadlineDateOffset `
                    -ProtectedSourceIds $ProtectedDependencyAppIds `
                    -IntentFilter 'available' -CopyOnly:$willBeSuperseded -SkipDependencies
            } catch {
                Write-Log "Could not migrate the available assignments from $($old.DisplayName): $_" -Level Error
            }
        }
    }

    # 7. Default deployments (unchanged behavior). Adds default groups as
    #    available-intent assignments if not already present.
    if ($Script:DefaultDeploymentGroups) {
        $ID = $CurrentApp.Id
        $CurrentlyDeployedIDs = (Get-YardstickAppAssignment -Id $ID).GroupID
        foreach ($DeploymentGroupID in $Script:DefaultDeploymentGroups) {
            if (!($CurrentlyDeployedIDs -Contains $DeploymentGroupID)) {
                Write-Log "Deploying $($CurrentApp.DisplayName) to default group $DeploymentGroupID."
                Add-YardstickWin32AppAssignmentGroup -Include -ID $ID -GroupID $DeploymentGroupID -Intent "available" -Notification "hideAll" | Out-Null
            }
        }
    }

    # 8. Prune before wiring supersedence. Intune refuses to delete an app that
    #    still participates in a supersedence relationship, so removing the
    #    expired versions first keeps the graph pointing only at apps that
    #    survive this run.
    $Pruned = @()
    if (!$NoDelete) {
        foreach ($old in $ToPrune) {
            if ($ProtectedDependencyAppIds.Contains($old.id)) {
                Write-Log "Keeping $($old.displayName): dependent apps still target this version."
                continue
            }
            try {
                Remove-YardstickApp -App $old
                $Pruned += $old.id
            } catch {
                Write-Log "Could not remove $($old.DisplayName) ($($old.id)): $_" -Level Error
            }
        }
    }

    # 9. Supersedence. The newest version becomes the single superseding parent
    #    for every surviving older version plus the Ω DETECT - anchor. The anchor
    #    is forced to "Update" because a "Replace" against its deliberately wide
    #    detection rule would uninstall the app from every device that has any
    #    version of it installed.
    $supersededCount = 0
    if ($Script:Supersedence) {
        $SupersedenceTargets = @($ToKeep)
        # Anything we intended to prune but could not (protected by a dependent,
        # or the delete was refused) still needs to be superseded so devices on
        # it move forward. Same-version leftovers are excluded: superseding an
        # app that carries the identical version is meaningless.
        $SupersedenceTargets += @($ToPrune | Where-Object { ($Pruned -notcontains $_.id) -and ($SameVersionIds -notcontains $_.id) })
        if ($Anchor) { $SupersedenceTargets += $Anchor }
        $Type = if ($Script:UninstallPreviousVersion) { 'Replace' } else { 'Update' }
        try {
            $supersededCount = Set-YardstickSupersedence -NewApp $CurrentApp -SupersededApps $SupersedenceTargets -Type $Type `
                -UpdateOnlyIds @(if ($Anchor) { $Anchor.id })
        } catch {
            Write-Log "Could not configure supersedence for $($CurrentApp.DisplayName): $_" -Level Error
        }
    } else {
        Write-Log "Supersedence is disabled for this recipe." -Level Detail
    }

    # 10. Native auto-update. Intune only honours this on available-intent
    #     assignments, so required-only recipes are a no-op here by design.
    $autoUpdateApplied = 0
    if ($Script:AutoUpdateOnAssignment) {
        try {
            $autoUpdateApplied = Set-AssignmentAutoUpdate -AppId $CurrentApp.Id -Enabled $true -IntentFilter 'available' `
                -SkipGroupIds $Script:GroupsSkipAutoUpdates
        } catch {
            Write-Log "Could not enable auto-update on the assignments of $($CurrentApp.DisplayName): $_" -Level Error
        }
    }

    return [PSCustomObject]@{
        CurrentApp        = $CurrentApp
        SupersededCount   = $supersededCount
        AutoUpdateApplied = $autoUpdateApplied
        AnchorStatus      = $anchorStatus
        PrunedCount       = @($Pruned).Count
        RenamedCount      = $renamedCount
        # Steps 9 and 10 reconcile on every pass, so they are not evidence that
        # anything was wrong. Only the state changes an operator would want to
        # read about in a report count here.
        ChangedSomething  = (@($Pruned).Count -gt 0) -or ($renamedCount -gt 0) -or $anchorPinned
    }
}


function Invoke-YardstickSkip {
    <#
    .SYNOPSIS
    Handles a recipe that will publish nothing this run.

    .DESCRIPTION
    Logs why, and under -Repair reconciles the versions already in Intune anyway.
    Nothing to publish is not the same as nothing to do: retention, naming,
    assignments and supersedence are independent of whether the vendor shipped a
    new build, but that state used to be reachable only through the upload path,
    so a version that failed to prune waited for the next release to be retried -
    which for a slow-moving app can be never.

    Declared at script scope for the same reason as Invoke-YardstickAppMaintenance:
    it reads the caller's per-recipe state ($Repair, $AppId_Processing,
    $CurrentDisplayName, $DependentUpdateStatus) directly.

    .PARAMETER Reason
    Sentence explaining why nothing is being published, ending in a period.
    #>
    param(
        [Parameter(Mandatory)][string]$Reason
    )

    if (-not $Repair) {
        Write-Log "$Reason Skipping update."
        return
    }

    Write-Log "$Reason Running maintenance sweep."
    try {
        $Sweep = Invoke-YardstickAppMaintenance
        if ($Sweep -and $Sweep.ChangedSomething) {
            $sweepParts = @()
            if ($Sweep.PrunedCount -gt 0)     { $sweepParts += "pruned $($Sweep.PrunedCount)" }
            if ($Sweep.RenamedCount -gt 0)    { $sweepParts += "renamed $($Sweep.RenamedCount)" }
            if ($Sweep.SupersededCount -gt 0) { $sweepParts += "supersedes $($Sweep.SupersededCount)" }
            if ($Sweep.AnchorStatus)          { $sweepParts += "anchor: $($Sweep.AnchorStatus)" }
            Add-SuccessfulApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName `
                -Version $Sweep.CurrentApp.displayVersion -Action "Maintained" `
                -Dependents $DependentUpdateStatus -AutoUpdateStatus ($sweepParts -join ', ')
        }
    } catch {
        Write-Log "Maintenance sweep failed: $_" -Level Error
    }
}


function Set-ScriptVariables {
    <#
    .SYNOPSIS
    Sets script-scoped variables from application parameters and per-app preferences.

    .DESCRIPTION
    Sets per-application script variables from recipe parameters, falling back to
    global preferences with null coalescing. Folder locations and connection settings
    are set once at script startup and are not reassigned here.

    .PARAMETER Parameters
    Hashtable containing application-specific parameters from the YAML recipe.

    .PARAMETER Preferences
    Hashtable containing global preferences from the preferences.yaml file.
    #>
    param(
        [hashtable]$Parameters,
        [hashtable]$Preferences
    )

    # Set all variables from default preferences and the application recipe
    $Script:Url = if ($Parameters.urlRedirects -eq $true) {Get-RedirectedUrl $Parameters.url} else {$Parameters.url}
    $Script:Id = $Parameters.id
    $Script:Version = $Parameters.version
    $Script:FileDetectionVersion = $Parameters.fileDetectionVersion
    $Script:DisplayName = $Parameters.displayName
    $Script:DisplayVersion = $Parameters.displayVersion
    $Script:FileName = $Parameters.fileName
    $Script:FileDetectionPath = $Parameters.fileDetectionPath
    $Script:PreDownloadScript = if ($Parameters.preDownloadScript) { [ScriptBlock]::Create($Parameters.preDownloadScript) }
    $Script:DownloadScript = if ($Parameters.downloadScript) { [ScriptBlock]::Create($Parameters.downloadScript) }
    $Script:PostDownloadScript = if ($Parameters.postDownloadScript) { [ScriptBlock]::Create($Parameters.postDownloadScript) }
    $Script:PostRunScript = if ($Parameters.postRunScript) { [ScriptBlock]::Create($Parameters.postRunScript) }
    $Script:InstallScript = $Parameters.installScript
    $Script:UninstallScript = $Parameters.uninstallScript
    $Script:PowerShellInstallScript = $Parameters.powerShellInstallScript
    $Script:PowerShellUninstallScript = $Parameters.powerShellUninstallScript

    # Manual-drop recipes take their payload from the software dropbox instead of
    # a URL or download script. See Get-YardstickDropboxPayload.
    $Script:ManualDownload = if ($null -ne $Parameters.manualDownload) { [bool]$Parameters.manualDownload } else { $false }
    $Script:ManualDownloadFolder = if ($Parameters.manualDownloadFolder) { $Parameters.manualDownloadFolder } else { $Parameters.id }

    # Days to wait between runs, for vendors that throttle repeated requests.
    $Script:MinRunInterval = if ($null -ne $Parameters.minRunInterval) { [int]$Parameters.minRunInterval } else { 0 }

    # Use parameters or fall back to preferences
    $Script:ScopeTags = if ($null -ne $Parameters.scopeTags) { $Parameters.scopeTags } else { $Preferences.defaultScopeTags }
    $Script:Owner = if ($null -ne $Parameters.owner) { $Parameters.owner } else { $Preferences.defaultOwner }
    $Script:MaximumInstallationTimeInMinutes = if ($null -ne $Parameters.maximumInstallationTimeInMinutes) { $Parameters.maximumInstallationTimeInMinutes } else { $Preferences.defaultMaximumInstallationTimeInMinutes }
    $Script:MinOSVersion = if ($null -ne $Parameters.minOSVersion) { $Parameters.minOSVersion } else { $Preferences.defaultMinOSVersion }
    $Script:InstallExperience = if ($null -ne $Parameters.installExperience) { $Parameters.installExperience } else { $Preferences.defaultInstallExperience }
    $Script:RestartBehavior = if ($null -ne $Parameters.restartBehavior) { $Parameters.restartBehavior } else { $Preferences.defaultRestartBehavior }
    $Script:DefaultDeploymentGroups = if ($null -ne $Parameters.defaultDeploymentGroups) { $Parameters.defaultDeploymentGroups } else { $Preferences.defaultDeploymentGroups }
    $Script:AllowUserUninstall = if ($null -ne $Parameters.allowUserUninstall) { $Parameters.allowUserUninstall } else { $Preferences.defaultAllowUserUninstall }
    $Script:Is32BitApp = if ($null -ne $Parameters.is32BitApp) { $Parameters.is32BitApp } else { $Preferences.defaultIs32BitApp }
    $Script:Architecture = if ($null -ne $Parameters.architecture) { $Parameters.architecture } else { $Preferences.defaultArchitecture }
    $Script:DeadlineDateOffset = if ($null -ne $Parameters.deadlineDateOffset) { $Parameters.deadlineDateOffset } else { $Preferences.defaultDeadlineDateOffset }
    $Script:AvailableDateOffset = if ($null -ne $Parameters.availableDateOffset) { $Parameters.availableDateOffset } else { $Preferences.defaultAvailableDateOffset }
    $Script:AllowDependentLinkUpdates = if ($null -ne $Parameters.allowDependentLinkUpdates) { $Parameters.allowDependentLinkUpdates } else { $Preferences.defaultAllowDependentLinkUpdates }
    
    # Detection-related variables
    $Script:DetectionType = $Parameters.detectionType
    $Script:FileDetectionVersion = $Parameters.fileDetectionVersion
    $Script:FileDetectionMethod = $Parameters.fileDetectionMethod
    $Script:FileDetectionName = $Parameters.fileDetectionName
    $Script:FileDetectionOperator = $Parameters.fileDetectionOperator
    $Script:FileDetectionDateTime = $Parameters.fileDetectionDateTime
    $Script:FileDetectionValue = $Parameters.fileDetectionValue
    $Script:RegistryDetectionMethod = $Parameters.registryDetectionMethod
    $Script:RegistryDetectionKey = $Parameters.registryDetectionKey
    $Script:RegistryDetectionValueName = $Parameters.registryDetectionValueName
    $Script:RegistryDetectionValue = $Parameters.registryDetectionValue
    $Script:RegistryDetectionOperator = $Parameters.registryDetectionOperator
    # Extra file/registry rules that must ALL match alongside the primary rule above.
    # Entries are the recipe's own hashtables, so a pre/post-download script can set
    # e.g. $additionalDetectionRules[0].registryDetectionValue at runtime.
    $Script:AdditionalDetectionRules = @($Parameters.additionalDetectionRules | Where-Object { $_ })
    $Script:DetectionScript = $Parameters.detectionScript
    $Script:DetectionScriptFileExtension = if ($null -ne $Parameters.detectionScriptFileExtension) { $Parameters.detectionScriptFileExtension } else { $Preferences.defaultDetectionScriptFileExtension }
    $Script:DetectionScriptRunAs32Bit = if ($null -ne $Parameters.detectionScriptRunAs32Bit) { $Parameters.detectionScriptRunAs32Bit } else { $Preferences.defaultdetectionScriptRunAs32Bit }
    $Script:DetectionScriptEnforceSignatureCheck = if ($null -ne $Parameters.detectionScriptEnforceSignatureCheck) { $Parameters.detectionScriptEnforceSignatureCheck } else { $Preferences.defaultdetectionScriptEnforceSignatureCheck }
    $Script:DependentLinkUpdateEnabled = if ($null -ne $Parameters.dependentLinkUpdateEnabled) { $Parameters.dependentLinkUpdateEnabled } elseif ($null -ne $Preferences.dependentLinkUpdateEnabled) { $Preferences.dependentLinkUpdateEnabled } else { $true }
    $Script:DependentLinkUpdateRetryCount = if ($null -ne $Parameters.dependentLinkUpdateRetryCount) { $Parameters.dependentLinkUpdateRetryCount } elseif ($null -ne $Preferences.dependentLinkUpdateRetryCount) { $Preferences.dependentLinkUpdateRetryCount } else { 3 }
    $Script:DependentLinkUpdateRetryDelaySeconds = if ($null -ne $Parameters.dependentLinkUpdateRetryDelaySeconds) { $Parameters.dependentLinkUpdateRetryDelaySeconds } elseif ($null -ne $Preferences.dependentLinkUpdateRetryDelaySeconds) { $Preferences.dependentLinkUpdateRetryDelaySeconds } else { 5 }
    $Script:DependentLinkUpdateTimeoutSeconds = if ($null -ne $Parameters.dependentLinkUpdateTimeoutSeconds) { $Parameters.dependentLinkUpdateTimeoutSeconds } elseif ($null -ne $Preferences.dependentLinkUpdateTimeoutSeconds) { $Preferences.dependentLinkUpdateTimeoutSeconds } else { 60 }
    $Script:DependentApplicationBlocklist = if ($Parameters.dependentApplicationBlocklist) {
        $Parameters.dependentApplicationBlocklist
    } elseif ($Parameters.dependentApplicationBlacklist) {
        # The key's old name; Test-RecipeSchema asks for it to be renamed.
        $Parameters.dependentApplicationBlacklist
    } elseif ($Preferences.dependentApplicationBlocklist) {
        $Preferences.dependentApplicationBlocklist
    } else {
        @()
    }
    if (-not ($Script:DependentApplicationBlocklist -is [System.Collections.IEnumerable])) {
        $Script:DependentApplicationBlocklist = @($Script:DependentApplicationBlocklist)
    }
    
    # Additional variables
    $Script:IconFile = $Parameters.iconFile
    $Script:Description = $Parameters.description
    $Script:Publisher = $Parameters.publisher
    $Script:VersionLock = $Parameters.versionLock
    $Script:ProductCode = $null
    
    # Handle version lock logic
    if ($Script:VersionLock) {
        $Script:NumVersionsToKeep = 1
        if ($Parameters.numVersionsToKeep -gt 1) {
            Write-Log "versionLock is set, so numVersionsToKeep ($($Parameters.numVersionsToKeep)) is ignored." -Level Warning
        }
    } else {
        $Script:NumVersionsToKeep = if ($null -ne $Parameters.numVersionsToKeep) { $Parameters.numVersionsToKeep } else { $Preferences.defaultNumVersionsToKeep }
    }

    # Supersedence + auto-update wiring. `autoUpdate` is accepted as an alias
    # for `autoUpdateOnAssignment` so pre-existing recipes continue to work.
    $Script:Supersedence = if ($null -ne $Parameters.supersedence) {
        [bool]$Parameters.supersedence
    } elseif ($null -ne $Preferences.defaultSupersedence) {
        [bool]$Preferences.defaultSupersedence
    } else { $true }

    $Script:UninstallPreviousVersion = if ($null -ne $Parameters.uninstallPreviousVersion) {
        [bool]$Parameters.uninstallPreviousVersion
    } elseif ($null -ne $Preferences.defaultUninstallPreviousVersion) {
        [bool]$Preferences.defaultUninstallPreviousVersion
    } else { $false }

    $Script:AutoUpdateOnAssignment = if ($null -ne $Parameters.autoUpdateOnAssignment) {
        [bool]$Parameters.autoUpdateOnAssignment
    } elseif ($null -ne $Parameters.autoUpdate) {
        [bool]$Parameters.autoUpdate
    } elseif ($null -ne $Preferences.defaultAutoUpdate) {
        [bool]$Preferences.defaultAutoUpdate
    } else { $true }

    # Groups that opt out of native auto-update. Unlike every other recipe
    # setting this is a union, not an override: a group excluded tenant-wide
    # stays excluded even when a recipe supplies its own list.
    $Script:GroupsSkipAutoUpdates = @(
        @($Preferences.defaultGroupsSkipAutoUpdates) + @($Parameters.groupSkipAutoUpdates) |
            Where-Object { $_ } | Select-Object -Unique
    )

    $Script:UseDetectAnchor = if ($null -ne $Preferences.useDetectAnchor) {
        [bool]$Preferences.useDetectAnchor
    } else { $true }

    # Handle PowerShell Script batch handoff.
    # "powershell.exe" is left bare on purpose. It resolves to SysWOW64 under the
    # 32-bit Management Extension, which is wrong, but the correction belongs in
    # the script (see Add-NativeBitnessGuard) rather than here: a
    # "%SystemRoot%\sysnative\..." command line only resolves inside a 32-bit
    # process and fails outright under a 64-bit caller.
    if (($Script:PowerShellInstallScript) -and (!$Script:InstallScript)) {
        $Script:InstallScript = "powershell.exe -noprofile -executionpolicy bypass -file .\install.ps1"
    }
    if (($Script:PowerShellUninstallScript) -and (!$Script:UninstallScript)) {
        $Script:UninstallScript = "powershell.exe -noprofile -executionpolicy bypass -file .\uninstall.ps1"
    }
}



# Import Folder Locations from preferences
$Script:Temp = $Prefs.Temp
$Script:BuildSpace = $Prefs.Buildspace
$Script:Scripts = $Prefs.Scripts
$Script:Published = $Prefs.Published
$Script:Recipes = $Prefs.Recipes
$Script:Icons = $Prefs.Icons
$Script:Tools = $Prefs.Tools
$Script:Secrets = $Prefs.Secrets
$Script:SoftwareDropbox = $Prefs.SoftwareDropbox
$Script:SoftwareArchive = $Prefs.SoftwareArchive

# Backup of uploaded .intunewin files is optional - a blank or absent Backup key
# turns it off. A present-but-empty YAML key comes through as "", so this checks
# for whitespace rather than just $null.
$Script:Backup = $Prefs.Backup
$Script:BackupEnabled = -not [string]::IsNullOrWhiteSpace($Script:Backup)
$Script:BackupVersionsToKeep = if (($Prefs.backupVersionsToKeep -as [int]) -gt 0) { [int]$Prefs.backupVersionsToKeep } else { 3 }

# Import Intune Connection Settings from Windows Credential Manager. Legacy
# TenantID/ClientId/ClientSecret keys in preferences.yaml are migrated on first
# run; -NoInteractive runs fail fast rather than blocking on a prompt.
$Script:IntuneCredential = Initialize-YardstickIntuneCredential -Preferences $Prefs -NoPrompt:$NoInteractive

# Validate ApplicationId parameter
if ($ApplicationId) {
    foreach ($AppIdToValidate in $ApplicationId) {
        if (-not (Get-YardstickRecipeFile -RecipesPath $Recipes -Id $AppIdToValidate)) {
            $disabledHint = if (Get-YardstickRecipeFile -RecipesPath $Recipes -Id $AppIdToValidate -IncludeDisabled) { ' Its recipe is in a Disabled folder.' } else { '' }
            Write-Log "No recipe named $AppIdToValidate.$disabledHint" -Level Error
            exit 1
        }
    }
}

function Get-ApplicationsToProcess {
    <#
    .SYNOPSIS
    Determines which applications to process based on command-line parameters.
    
    .DESCRIPTION
    This function encapsulates the logic for selecting applications to process,
    whether processing a single app, a group, or all applications.
    
    .PARAMETER ApplicationId
    The ID of a single application to process.
    
    .PARAMETER Group
    The name of an application group to process.
    
    .PARAMETER All
    Switch to process all applications.
    
    .PARAMETER NoInteractive
    Switch to exclude interactive applications when processing all.
    #>
    param(
        [string[]]$ApplicationId,
        [string]$Group,
        [switch]$All,
        [switch]$NoInteractive
    )
    
    $applications = [System.Collections.Generic.List[PSObject]]::new()
    
    if ($All) {
        foreach ($recipeFile in Get-YardstickRecipeFile -RecipesPath $Recipes -ExcludeInteractive:$NoInteractive) {
            $applications.Add([String]$recipeFile.BaseName) | Out-Null
        }
    } elseif ($Group) {
        $groupsPath = Join-Path $ProjectRoot 'RecipeGroups.yaml'
        try {
            $groupFile = Get-Content -LiteralPath $groupsPath -Raw | ConvertFrom-Yaml
        } catch {
            Write-Log "Could not read ${groupsPath}: $_" -Level Error
            exit 3
        }
        $groupNames = if ($groupFile -is [System.Collections.IDictionary]) { @($groupFile.Keys) } else { @() }
        $groupName = $groupNames | Where-Object { $_ -eq $Group } | Select-Object -First 1
        if ($null -eq $groupName) {
            $available = if ($groupNames) { "Groups in RecipeGroups.yaml: $(($groupNames | Sort-Object) -join ', ')." } else { 'RecipeGroups.yaml has no groups.' }
            Write-Log "There is no group named '$Group'. $available" -Level Error
            exit 3
        }
        foreach ($Id in @($groupFile[$groupName])) {
            if (-not [string]::IsNullOrWhiteSpace([string]$Id)) {
                $applications.Add([String]$Id) | Out-Null
            }
        }
        if ($applications.Count -eq 0) {
            Write-Log "Group '$groupName' in RecipeGroups.yaml lists no recipes." -Level Warning
        }
    } else {
        foreach ($Id in $ApplicationId) {
            $applications.Add([String]$Id) | Out-Null
        }
    }
    return $applications
}



function Update-ScriptPlaceholders {
    <#
    .SYNOPSIS
    Replaces placeholders in script strings with actual values.
    
    .DESCRIPTION
    This function handles the replacement of common placeholders like <filename>,
    <productcode>, and <version> in install, uninstall, and detection scripts.
    
    .PARAMETER FileName
    The actual filename to replace <filename> placeholders with.
    
    .PARAMETER ProductCode
    The MSI product code to replace <productcode> placeholders with.
    
    .PARAMETER Version
    The version string to replace <version> placeholders with.
    #>
    param(
        [string]$FileName,
        [string]$ProductCode = "",
        [string]$Version
    )
    
    # Replace the <filename> placeholder with the actual filename
    if ($Script:InstallScript) {
        $Script:InstallScript = $Script:InstallScript.replace("<filename>", $FileName)
    }
    if ($Script:PowerShellInstallScript) {
        $Script:PowerShellInstallScript = $Script:PowerShellInstallScript.replace("<filename>", $FileName)
    }
    if ($Script:UninstallScript) {
        $Script:UninstallScript = $Script:UninstallScript.replace("<filename>", $FileName)
    }
    if ($Script:PowerShellUninstallScript) {
        $Script:PowerShellUninstallScript = $Script:PowerShellUninstallScript.replace("<filename>", $FileName)
    }
    if ($Script:DetectionScript) {
        $Script:DetectionScript = $Script:DetectionScript.replace("<filename>", $FileName)
    }
    if ($Script:RegistryDetectionKey) {
        $Script:RegistryDetectionKey = $Script:RegistryDetectionKey.replace("<filename>", $FileName)
    }

    # Replace the <productcode> placeholder with the actual product code
    if ($ProductCode) {
        if ($Script:InstallScript) {
            $Script:InstallScript = $Script:InstallScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:PowerShellInstallScript) {
            $Script:PowerShellInstallScript = $Script:PowerShellInstallScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:UninstallScript) {
            $Script:UninstallScript = $Script:UninstallScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:PowerShellUninstallScript) {
            $Script:PowerShellUninstallScript = $Script:PowerShellUninstallScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:DetectionScript) {
            $Script:DetectionScript = $Script:DetectionScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:RegistryDetectionKey) {
            $Script:RegistryDetectionKey = $Script:RegistryDetectionKey.replace("<productcode>", $ProductCode)
        }
    }
    
    # Replace <version> placeholder with the actual version
    if ($Script:InstallScript) {
        $Script:InstallScript = $Script:InstallScript.replace("<version>", $Version)
    }
    if ($Script:PowerShellInstallScript) {
        $Script:PowerShellInstallScript = $Script:PowerShellInstallScript.replace("<version>", $Version)
    }
    if ($Script:UninstallScript) {
        $Script:UninstallScript = $Script:UninstallScript.replace("<version>", $Version) 
    }
    if ($Script:PowerShellUninstallScript) {
        $Script:PowerShellUninstallScript = $Script:PowerShellUninstallScript.replace("<version>", $Version) 
    }
    if ($Script:DetectionScript) {
        $Script:DetectionScript = $Script:DetectionScript.replace("<version>", $Version) 
    } 
    if ($Script:RegistryDetectionKey) {
        $Script:RegistryDetectionKey = $Script:RegistryDetectionKey.replace("<version>", $Version)
    }

    # Additional detection rules get the same tokens in their path/key/value fields
    foreach ($rule in $Script:AdditionalDetectionRules) {
        foreach ($field in 'fileDetectionPath', 'fileDetectionName', 'fileDetectionVersion', 'registryDetectionKey', 'registryDetectionValue') {
            if ($rule[$field] -is [string]) {
                $value = $rule[$field].replace("<filename>", $FileName).replace("<version>", $Version)
                if ($ProductCode) {
                    $value = $value.replace("<productcode>", $ProductCode)
                }
                $rule[$field] = $value
            }
        }
    }
}



function New-DetectionRule {
    <#
    .SYNOPSIS
    Creates Intune Win32 app detection rules based on detection type.
    
    .DESCRIPTION
    This function creates the appropriate detection rule object based on the
    specified detection type (file, msi, registry, or script). For file, msi,
    and registry recipes, each entry in additionalDetectionRules adds another
    rule; Intune requires every rule to match. Test-RecipeSchema rejects
    additionalDetectionRules on script recipes, since Intune cannot combine them.
    
    .PARAMETER DetectionType
    The type of detection rule to create (file, msi, registry, script).
    
    .PARAMETER ProductCode
    The MSI product code (required for MSI detection type).
    #>
    param(
        [string]$DetectionType,
        [string]$ProductCode = ""
    )
    
    if ($DetectionType -eq "script") {
        if (!(Test-Path $Script:Scripts\$Script:Id)) {
            New-Item -Name $Script:Id -ItemType Directory -Path $Script:Scripts | Out-Null
        }
        $ScriptLocation = "$($Script:Scripts)\$($Script:Id)\$($Script:Version).$($Script:DetectionScriptFileExtension)"
        Set-Content -Path $ScriptLocation -Value $Script:DetectionScript -Force
        $DetRule = New-YardstickWin32AppDetectionRuleScript -ScriptFile $ScriptLocation -EnforceSignatureCheck $Script:DetectionScriptEnforceSignatureCheck -RunAs32Bit $Script:DetectionScriptRunAs32Bit
        return $DetRule
    }

    $PrimarySpec = @{
        detectionType              = $DetectionType
        fileDetectionMethod        = $Script:FileDetectionMethod
        fileDetectionPath          = $Script:FileDetectionPath
        fileDetectionName          = $Script:FileDetectionName
        fileDetectionOperator      = $Script:FileDetectionOperator
        fileDetectionVersion       = $Script:FileDetectionVersion
        fileDetectionDateTime      = $Script:FileDetectionDateTime
        fileDetectionValue         = $Script:FileDetectionValue
        registryDetectionMethod    = $Script:RegistryDetectionMethod
        registryDetectionKey       = $Script:RegistryDetectionKey
        registryDetectionValueName = $Script:RegistryDetectionValueName
        registryDetectionValue     = $Script:RegistryDetectionValue
        registryDetectionOperator  = $Script:RegistryDetectionOperator
    }
    $RuleArgs = @{
        ProductCode    = $ProductCode
        DefaultVersion = $Script:Version
        Is32BitApp     = [bool]$Script:Is32BitApp
    }
    $Rules = @(New-YardstickDetectionRuleFromSpec -Spec $PrimarySpec @RuleArgs)
    foreach ($AdditionalRule in $Script:AdditionalDetectionRules) {
        $Rules += New-YardstickDetectionRuleFromSpec -Spec $AdditionalRule @RuleArgs
    }
    if ($Rules.Count -gt 1) {
        Write-Log "Using $($Rules.Count) detection rules (all must match)." -Level Detail
    }
    return $Rules
}


function Invoke-Cleanup {
    <#
    .SYNOPSIS
    Performs cleanup operations on build and published directories.
    
    .DESCRIPTION
    This function cleans up temporary files and directories created during
    the application processing, with error handling for non-critical failures.
    #>
    # Backups read from the Published folder on a background thread, so make sure
    # none are still in flight before anything here deletes their source.
    Wait-YardstickBackup -TimeoutSeconds 600 | Out-Null

    Write-Log "Cleaning up the buildspace..." -Level Detail
    try {
        Get-ChildItem $BuildSpace -Exclude ".gitkeep" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Log "Could not clean up the buildspace completely: $_" -Level Warning
    }
    
    Write-Log "Removing .intunewin files..." -Level Detail
    try {
        Get-ChildItem $Published -Exclude ".gitkeep" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Log "Could not remove the .intunewin files completely: $_" -Level Warning
    }
}



# Get the list of applications to process
$Applications = Get-ApplicationsToProcess -ApplicationId $ApplicationId -Group $Group -All:$All -NoInteractive:$NoInteractive

# Initialize application tracking for email notifications
Initialize-ApplicationTracker

# Build run parameters string for email report
$RunParametersArray = @()
if ($ApplicationId) { $RunParametersArray += "-ApplicationId $($ApplicationId -join ',')" }
if ($Group) { $RunParametersArray += "-Group $Group" }
if ($All) { $RunParametersArray += "-All" }
if ($NoInteractive) { $RunParametersArray += "-NoInteractive" }
if ($Force) { $RunParametersArray += "-Force" }
if ($NoDelete) { $RunParametersArray += "-NoDelete" }
if ($Repair) { $RunParametersArray += "-Repair" }
$RunParameters = $RunParametersArray -join " "

# Authenticate up front so a bad or expired secret fails before any packaging work,
# then check how much life the secret has left.
try {
    Connect-YardstickGraph
    $Script:IntuneCredential = Update-YardstickSecretExpiration -Credential $Script:IntuneCredential
} catch {
    Write-Log "Could not authenticate to Microsoft Graph: $_" -Level Error
    Write-Log "If the client secret was rotated, store the new one with .\Set-YardstickCredential.ps1."
    exit 1
}
Test-YardstickSecretExpiration -Credential $Script:IntuneCredential -Preferences $Prefs -NoEmail:$NoEmail | Out-Null

# Last line of defence against a hung run. Everything below reports its stage to
# this; if one of them stops reporting for recipeStageTimeoutMinutes, the watchdog
# writes the diagnosis to the log and ends the process. Only when the process is
# the run's own, though: Invoke-Yardstick at a prompt would take the user's console
# down with it, so there the watchdog only logs.
Start-YardstickWatchdog -TimeoutMinutes $Script:RecipeStageTimeoutMinutes -NoExit:(-not (Test-YardstickDedicatedProcess))

# Main processing loop
Write-Log "Processing $(@($Applications).Count) recipe(s)."
foreach ($AppId_Processing in $Applications) {
    # This line stays untagged: the log report starts a new app at it. Everything
    # logged after it, until the finally below, carries the app's tag.
    Write-Log "Starting update for $AppId_Processing..."
    Set-YardstickLogContext $AppId_Processing
    Set-YardstickWatchdogStage -AppId $AppId_Processing -Stage 'Starting up'
    Set-Location $YardstickModuleRoot

    # Initialize variables for tracking
    $CurrentDisplayName = $AppId_Processing

    # Set-ScriptVariables nulls this for a recipe without a postRunScript, but the
    # configuration and validation failures below skip out before it ever runs. The
    # finally block still fires on those paths, so without this the previous
    # recipe's cleanup hook would run against this one.
    $Script:PostRunScript = $null

    try {
        # Refresh token if necessary
        Connect-YardstickGraph

        # Undo any folder variable the previous recipe overwrote before anything
        # in this iteration reads one.
        Restore-RunnerPathVariable

        # Likewise re-assert the web timeout defaults. A recipe is free to reassign
        # $PSDefaultParameterValues, and one that does would otherwise leave every
        # recipe after it able to hang indefinitely again.
        Set-YardstickWebRequestTimeout -TimeoutSeconds $Script:WebRequestTimeoutSeconds

        # Clear the temp file. -LiteralPath and a top-level enumeration on purpose:
        # Get-ChildItem $Temp -Exclude ".gitkeep" -Recurse quietly returns nothing
        # when $Temp does not exist, because -Exclude turns -Path into a wildcard
        # container query. That hid a missing temp directory for a whole run.
        Write-Log "Clearing the temp directory..." -Level Detail
        if (-not (Test-Path -LiteralPath $Temp -PathType Container)) {
            throw "The Yardstick temp directory '$Temp' does not exist."
        }
        Get-ChildItem -LiteralPath $Temp -Force |
            Where-Object Name -ne ".gitkeep" |
            Remove-Item -Recurse -Force

        # Open the YAML file and collect all necessary attributes
        $recipeFile = @(Get-YardstickRecipeFile -RecipesPath $Recipes -Id $AppId_Processing)[0]
        if (-not $recipeFile) {
            # Only a group can list a disabled recipe: -ApplicationId is checked
            # before the run and -All never lists one. Disabling is deliberate, so it
            # is not reported as a failure.
            if (Get-YardstickRecipeFile -RecipesPath $Recipes -Id $AppId_Processing -IncludeDisabled) {
                Write-Log "Skipping $AppId_Processing because its recipe is in a Disabled folder." -Level Warning
            } else {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $AppId_Processing -Version "Unknown" -ErrorMessage "There is no recipe named $AppId_Processing." -FailureStage "Configuration"
            }
            continue
        }
        try {
            $Parameters = Get-Content -LiteralPath $recipeFile.FullName -Raw | ConvertFrom-Yaml
        } catch {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $AppId_Processing -Version "Unknown" -ErrorMessage "Could not open the recipe: $_" -FailureStage "Configuration"
            continue
        }

        # Resolve base recipe inheritance if present
        if ($Parameters.ContainsKey('base')) {
            try {
                $Parameters = Merge-RecipeWithBase -Recipe $Parameters -RecipesPath $Recipes
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $AppId_Processing -Version "Unknown" -ErrorMessage "Could not resolve the base recipe: $_" -FailureStage "Configuration"
                continue
            }
        }

        # Apply this site's overrides from Local\recipeOverrides.yaml
        if ($RecipeOverrides.ContainsKey($AppId_Processing)) {
            $Parameters = Merge-RecipeWithOverride -Recipe $Parameters -Override $RecipeOverrides[$AppId_Processing]
            Write-Log "Applied recipe overrides: $(@($RecipeOverrides[$AppId_Processing].Keys) -join ', ')" -Level Detail
        }

        # A shared recipe that needs a value only this site knows ships with a
        # <YOUR_...> placeholder. Left in, it would upload an install command that
        # cannot work.
        $placeholders = @(Get-RecipePlaceholder -Recipe $Parameters)
        if ($placeholders) {
            $fields = ($placeholders | ForEach-Object { "$($_.Field) ($($_.Placeholder))" }) -join ', '
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $AppId_Processing -Version "Unknown" -ErrorMessage "The recipe needs this site's values for: $fields. Set those fields for $AppId_Processing in $RecipeOverridesPath." -FailureStage "Configuration"
            continue
        }

        # Validate recipe schema before processing
        $validation = Test-RecipeSchema -Recipe $Parameters -RecipeId $AppId_Processing
        foreach ($w in $validation.Warnings) { Write-Log "Recipe check: $w" -Level Warning }
        if (-not $validation.IsValid) {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $AppId_Processing -Version "Unknown" -ErrorMessage ($validation.Errors -join '; ') -FailureStage "Recipe Validation"
            continue
        }

        # Set all script variables from parameters and preferences
        Set-ScriptVariables -Parameters $parameters -Preferences $Prefs
        
        # Update tracking variables with actual values
        $CurrentDisplayName = $Script:DisplayName
        $DependentUpdateStatus = [ordered]@{}
        $ProtectedDependencyAppIds = [System.Collections.Generic.HashSet[string]]::new()
        $DependentLinkOptions = @{
            Enabled = [bool]$Script:DependentLinkUpdateEnabled
            RetryCount = [int]$Script:DependentLinkUpdateRetryCount
            RetryDelaySeconds = [int]$Script:DependentLinkUpdateRetryDelaySeconds
            TimeoutSeconds = [int]$Script:DependentLinkUpdateTimeoutSeconds
            Blocklist = $Script:DependentApplicationBlocklist
        }


        if ($Repair) {
            # Correct any naming discrepancies before we continue.
            # Skip the Ω DETECT - anchor - it deliberately holds a different name.
            $anchorName = Get-DetectAnchorName -DisplayName $Script:DisplayName
            $CurrentApps = @(Get-SameAppAllVersions $Script:DisplayName | Where-Object DisplayName -ne $anchorName)
            for ($i = 1; $i -lt $CurrentApps.Count; $i++) {
                if ($CurrentApps[$i].DisplayName -ne "$($Script:DisplayName) (N-$i)") {
                    Write-Log "Renaming $($CurrentApps[$i].DisplayName) to $($Script:DisplayName) (N-$i)..."
                    Set-YardstickWin32App -Id $CurrentApps[$i].Id -DisplayName "$($Script:DisplayName) (N-$i)" | Out-Null
                }
            }
        }


        # Clear stale $matches from prior recipe iterations to prevent cross-contamination
        $null = "reset" -match "reset"

        # Recipe scripts share one scope for the whole run, so a left-side cast in an
        # earlier recipe leaves a type constraint that breaks every later recipe using
        # that variable name. Drop any such constraint before this recipe runs.
        Clear-RecipeVariableConstraint

        # Manual-drop recipes package whatever an operator staged in the dropbox.
        # This has to resolve before the pre-download script, because that script is
        # where the recipe reads $dropboxPath to work out $version and $fileName.
        $dropboxPath = $null
        $dropboxFiles = @()
        if ($Script:ManualDownload) {
            try {
                $DropboxPayload = Get-YardstickDropboxPayload -DropboxRoot $Script:SoftwareDropbox -Folder $Script:ManualDownloadFolder
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version "Unknown" -ErrorMessage "$_" -FailureStage "Manual Drop"
                continue
            }
            if (-not $DropboxPayload) {
                Invoke-YardstickSkip -Reason "Nothing staged in the software dropbox for $($Script:Id)."
                continue
            }
            $dropboxPath = $DropboxPayload.Path
            $dropboxFiles = $DropboxPayload.Files
            Write-Log "Found $($dropboxFiles.Count) staged file(s) in $dropboxPath."
        }

        # Vendors that throttle repeated requests set minRunInterval. The run is
        # recorded before the vendor is contacted, not when the recipe succeeds: a
        # failed attempt still counts against the vendor's limit, and retrying it
        # every run is what gets a site to start refusing us.
        if ($Script:MinRunInterval -gt 0) {
            $runDue = Test-YardstickRecipeRunDue -Path $RecipeRunHistoryPath -Id $AppId_Processing -Days $Script:MinRunInterval
            if (-not $runDue.IsDue -and -not $Force) {
                Invoke-YardstickSkip -Reason "Last run $($runDue.LastRun.ToString('yyyy-MM-dd HH:mm')); minRunInterval is $($Script:MinRunInterval) day(s), so the next run is on or after $($runDue.NextRun.ToString('yyyy-MM-dd'))."
                continue
            }
            if (-not $runDue.IsDue) {
                Write-Log "-Force is set: running despite minRunInterval ($($Script:MinRunInterval) day(s)); the last run was $($runDue.LastRun.ToString('yyyy-MM-dd HH:mm'))."
            }
            try {
                Set-YardstickRecipeLastRun -Path $RecipeRunHistoryPath -Id $AppId_Processing
            } catch {
                Write-Log "Could not record this run in ${RecipeRunHistoryPath}: $($_.Exception.Message)" -Level Warning
            }
        }

        # Run the pre-download script
        if ($Script:PreDownloadScript) {
            Write-Log "Running the pre-download script..." -Level Detail
            Set-YardstickWatchdogStage -Stage 'Pre-Download Script'
            try {
                # Pre-download scripts are almost all version probes against a vendor
                # endpoint, so a dropped or timed-out connection fails a recipe that
                # would have succeeded a moment later. Only genuinely transient
                # network faults are replayed; anything else throws on the first try.
                $preDownloadAttempts = 3
                for ($preDownloadAttempt = 1; ; $preDownloadAttempt++) {
                    try {
                        Invoke-Command -ScriptBlock $Script:PreDownloadScript -NoNewScope
                        break
                    } catch {
                        if ($preDownloadAttempt -ge $preDownloadAttempts -or -not (Test-YardstickTransientNetworkError -ErrorRecord $_)) { throw }
                        Write-Log "Pre-download script failed (attempt $preDownloadAttempt/$preDownloadAttempts) on a transient network error: $($_.Exception.Message)" -Level Warning
                        Start-Sleep -Seconds ([math]::Pow(2, $preDownloadAttempt))
                    }
                }
                Write-Log "Ran the pre-download script." -Level Detail
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "$_" -FailureStage "Pre-Download Script"
                continue
            }
        } else {
            Write-Log "No pre-download script." -Level Detail
        }

        # Validate the extracted version before proceeding
        $ExistingVersions = Get-SameAppAllVersions $Script:DisplayName

        # Get-SameAppAllVersions returns newest first, but do not reach for
        # $ExistingVersions.displayVersion[0]: PowerShell unrolls a single-element
        # property projection to a scalar, so with exactly one existing app that
        # indexes into the version *string* and yields its first character. A blank
        # displayVersion also has to be skipped -- a renamed or manually created app
        # can have none, and passing that to Compare-AppVersions throws and fails the
        # whole recipe.
        $NewestExistingVersion = Get-NewestComparableVersion -ExistingVersions $ExistingVersions

        $versionValidation = Test-ExtractedVersion -Version $Script:Version -ApplicationId $AppId_Processing -ExistingVersion $NewestExistingVersion
        foreach ($w in $versionValidation.Warnings) {
            Write-Log "Version check: $w" -Level Warning
        }
        if (-not $versionValidation.IsValid) {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $(if ($null -ne $Script:Version) { $Script:Version } else { "null" }) -ErrorMessage ($versionValidation.Errors -join '; ') -FailureStage "Version Validation"
            continue
        }

        # Trim whitespace from version if present (warned but not blocked above)
        if ($Script:Version -ne $Script:Version.Trim()) {
            $Script:Version = $Script:Version.Trim()
            Write-Log "Trimmed whitespace from the version: '$($Script:Version)'." -Level Detail
        }

        # Check if there is an up-to-date version in the repo already
        Write-Log "Checking if $($Script:DisplayName) $($Script:Version) is a new version..."
        
        if (-not $ExistingVersions) {
            Write-Log "No existing versions of $($Script:DisplayName) in Intune. Continuing with the update." -Level Detail
            $VersionCompareResult = 0
        }
        elseif (-not $NewestExistingVersion) {
            Write-Log "No existing version of $($Script:DisplayName) has a version number to compare against. Continuing with the update."
            $VersionCompareResult = 0
        }
        else {
            $VersionCompareResult = Compare-AppVersions $Script:Version $NewestExistingVersion
        }
        
        # Check various conditions to determine if we should proceed
        $SkipReason = $null
        if ($Force) {
            Write-Log "-Force is set: updating $($Script:DisplayName) $($Script:Version) regardless of the versions in Intune."
        } elseif (Test-VersionExcluded -Version $Script:Version -VersionLock $Script:VersionLock) {
            $SkipReason = "Version is locked to $($Script:VersionLock)."
        } elseif (@($ExistingVersions.displayVersion) -contains $Script:Version) {
            $SkipReason = "$($Script:DisplayName) $($Script:Version) is already in Intune."
        } elseif ($VersionCompareResult -eq 1) {
            Write-Log "$($Script:DisplayName) $($Script:Version) is newer than $NewestExistingVersion. Updating."
        } elseif ($VersionCompareResult -eq -1) {
            $SkipReason = "$($Script:DisplayName) $($Script:Version) is older than the currently newest available version $NewestExistingVersion."
        }

        if ($SkipReason) {
            Invoke-YardstickSkip -Reason $SkipReason
            continue
        }


        # See if this has been run before. If there are previous files, move them to a folder called "Old"
        if (Test-Path $BuildSpace\$($Script:Id)) {
            if (-not (Test-Path $BuildSpace\Old)) {
                New-Item -Path $BuildSpace -ItemType Directory -Name "Old" | Out-Null
            }
            Write-Log "Moving the previous buildspace to Old..." -Level Detail
            Move-Item -Path $BuildSpace\$($Script:Id) $BuildSpace\Old\$($Script:Id)-$(Get-Date -Format "MMddyyhhmmss")
        }
        if (Test-Path $Scripts\$($Script:Id)) {
            if (-not (Test-Path $Scripts\Old)) {
                New-Item -Path $Scripts -ItemType Directory -Name "Old" | Out-Null
            }
            Write-Log "Moving the previous script folder to Old..." -Level Detail
            Move-Item -Path $Scripts\$($Script:Id) $Scripts\Old\$($Script:Id)-$(Get-Date -Format "MMddyyhhmmss")
        }

        # Make the new BUILDSPACE directory
        New-Item -Path $BuildSpace\$($Script:Id) -ItemType Directory -Name $Script:Version | Out-Null
        Set-Location $BuildSpace\$($Script:Id)\$Script:Version

        Update-ScriptPlaceholders -FileName $Script:FileName -ProductCode $productCode -Version $Script:Version 

        # Download the new installer
        Write-Log "Downloading..."
        Set-YardstickWatchdogStage -Stage 'Download'
        if ($Script:Url) {
            Write-Log "URL: $($Script:Url)" -Level Detail
        }
        if ((-not ($Script:Url)) -and (-not ($Script:DownloadScript)) -and (-not ($Script:ManualDownload))) {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "The recipe sets no URL, download script or manual drop." -FailureStage "Download"
            continue
        }
        
        if ($Script:ManualDownload) {
            Write-Log "Copying the staged payload from $dropboxPath..."
            try {
                Copy-Item -Path (Join-Path $dropboxPath '*') -Destination . -Recurse -Force
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Could not copy the staged payload from ${dropboxPath}: $_" -FailureStage "Download"
                continue
            }
        } elseif ($Script:DownloadScript) {
            # Pop in a finally: the continue below would otherwise leave this entry on
            # the location stack, and Invoke-Yardstick's own Pop-Location would then
            # return the caller to the build folder instead of where they started.
            Push-Location $BuildSpace\$Script:Id\$Script:Version
            try {
                Invoke-Command -ScriptBlock $Script:DownloadScript -NoNewScope
                Write-Log "Ran the download script." -Level Detail
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "$_" -FailureStage "Download Script"
                Write-Log "Download script contents: $Script:DownloadScript" -Level Detail
                continue
            } finally {
                Pop-Location
            }
        } else {
            try {
                # Bounded by a stall timeout rather than a plain Start-BitsTransfer,
                # which blocks the runner with no way to give up.
                Invoke-YardstickBitsDownload -Source "$($Script:Url)" `
                    -Destination "$BuildSpace\$Script:Id\$Script:Version\$Script:FileName" `
                    -StallTimeoutMinutes $Script:BitsStallTimeoutMinutes
                Write-Log "Downloaded via BITS." -Level Detail
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "$_" -FailureStage "Download"
                continue
            }
        }
        # Update Placeholders again in case the download script changed anything
        Update-ScriptPlaceholders -FileName $Script:FileName -ProductCode $ProductCode -Version $Script:Version 
        # Run the post-download script
        if ($Script:PostDownloadScript) {
            Write-Log "Running the post-download script..." -Level Detail
            Set-YardstickWatchdogStage -Stage 'Post-Download Script'
            Push-Location $BuildSpace\$Script:Id\$Script:Version
            try {
                Invoke-Command -ScriptBlock $Script:PostDownloadScript -NoNewScope
                Write-Log "Ran the post-download script." -Level Detail
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "$_" -FailureStage "Post-Download Script"
                continue
            } finally {
                Pop-Location
            }
        }
        # Handle script placeholder replacement
        $ProductCode = ""
        if ($Script:FileName -match "\.msi$") {
            $ProductCode = Get-MSIProductCode $BuildSpace\$Script:Id\$Script:Version\$Script:FileName | Where-Object { $_ -match "^\{[0-9A-Fa-f\-]{36}\}$" }
            Write-Log "Product code: $ProductCode" -Level Detail
            if (-not $ProductCode) {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Could not read the MSI product code from $Script:FileName." -FailureStage "MSI Product Code Retrieval"
                continue
            }
        }
        # Update placeholders for the final time
        Update-ScriptPlaceholders -FileName $Script:FileName -ProductCode $ProductCode -Version $Script:Version

        # DEBUG: Make sure that there aren't any placeholder strings left
        if ($Script:InstallScript -match "<filename>|<productcode>|<version>" -or
            $Script:PowerShellInstallScript -match "<filename>|<productcode>|<version>" -or
            $Script:UninstallScript -match "<filename>|<productcode>|<version>" -or
            $Script:PowerShellUninstallScript -match "<filename>|<productcode>|<version>" -or
            $Script:DetectionScript -match "<filename>|<productcode>|<version>") {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "One or more placeholder strings were not replaced in the scripts." -FailureStage "Placeholder Replacement"
            continue
        }

        # Write the contents of the install and uninstall scripts to files if they are PowerShell scripts.
        # Add-NativeBitnessGuard is applied here rather than at assignment so the
        # recipe-authored text stays intact for logging and the placeholder audit
        # above; only the file that actually ships in the package carries it.
        if ($Script:PowerShellInstallScript) {
            Set-Content -Path $BuildSpace\$Script:Id\$Script:Version\install.ps1 -Value (Add-NativeBitnessGuard -ScriptContent $Script:PowerShellInstallScript) -Force
        }

        if ($Script:PowerShellUninstallScript) {
            Set-Content -Path $BuildSpace\$Script:Id\$Script:Version\uninstall.ps1 -Value (Add-NativeBitnessGuard -ScriptContent $Script:PowerShellUninstallScript) -Force
        }

        # Generate the .intunewin file
        Set-Location $YardstickModuleRoot
        Write-Log "Packaging the .intunewin file..." -Level Detail
        Set-YardstickWatchdogStage -Stage 'Packaging'
        $App = New-YardstickWin32AppPackage -SourceFolder $BuildSpace\$Script:Id\$Script:Version -SetupFile $Script:FileName -OutputFolder $Published -Force -TimeoutMinutes $Script:PackagingTimeoutMinutes

        # Upload .intunewin file to Intune
        # Detection Types
        # Intune does not require an icon, so a missing file must not discard a
        # package that has already been downloaded and built.
        $Icon = $null
        $missingIcon = $null
        if (Test-Path -LiteralPath "$($Icons)\$($Script:IconFile)" -PathType Leaf) {
            $Icon = New-YardstickWin32AppIcon -FilePath "$($Icons)\$($Script:IconFile)"
        } else {
            $missingIcon = if ($Script:IconFile) { [string]$Script:IconFile } else { '(no iconFile set)' }
            Write-Log "Icon file $($Script:IconFile) is not in the Icons folder. Uploading without an icon." -Level Warning
        }
        if (-not $Script:FileDetectionVersion) {
            $Script:FileDetectionVersion = $Script:Version
        }

        $DetectionRule = New-DetectionRule -DetectionType $Script:DetectionType -ProductCode $ProductCode

        # Generate the min OS requirement rule
        $RequirementRule = New-YardstickWin32AppRequirementRule -Architecture $Script:Architecture -MinimumSupportedWindowsRelease $Script:MinOSVersion

        # Create the Intune App
        Write-Log "Uploading $Script:DisplayName $Script:Version to Intune..."
        Set-YardstickWatchdogStage -Stage 'Upload'
        Connect-YardstickGraph
        try {
            if ($Script:AllowUserUninstall) {
                $Win32App = Add-YardstickWin32App -FilePath $App.path -DisplayName $Script:DisplayName -Description $Script:Description -Publisher $Script:Publisher -InstallExperience $Script:InstallExperience -RestartBehavior $Script:RestartBehavior -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $Script:InstallScript -UninstallCommandLine $Script:UninstallScript -Icon $Icon -AppVersion "$Script:Version" -ScopeTagName $Script:ScopeTags -Owner $Script:Owner -MaximumInstallationTimeInMinutes $Script:MaximumInstallationTimeInMinutes -AllowAvailableUninstall
            } else {
                $Win32App = Add-YardstickWin32App -FilePath $App.path -DisplayName $Script:DisplayName -Description $Script:Description -Publisher $Script:Publisher -InstallExperience $Script:InstallExperience -RestartBehavior $Script:RestartBehavior -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $Script:InstallScript -UninstallCommandLine $Script:UninstallScript -Icon $Icon -AppVersion "$Script:Version" -ScopeTagName $Script:ScopeTags -Owner $Script:Owner -MaximumInstallationTimeInMinutes $Script:MaximumInstallationTimeInMinutes
            }
            Write-Log "Uploaded. Waiting for Intune to publish it..."
            $PublishTimeout = (Get-Date).AddMinutes(30)
            do {
                Start-Sleep -Seconds 15
                $LiveWin32App = Get-YardstickWin32App -Id $Win32App.id
                if ((Get-Date) -gt $PublishTimeout) {
                    throw "Timed out waiting for Intune to publish $Script:DisplayName after 30 minutes (state: $($LiveWin32App.publishingState))"
                }
            } while ($LiveWin32App.publishingState -ne "published")
            $UploadedAt = Get-Date
        } catch {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "$_" -FailureStage "Intune Upload"
            continue
        }

        # Hand the .intunewin off to a background thread that copies it to the
        # backup share and then deletes the local copy. Staging it out of
        # $Published first is what stops the next app's
        # New-YardstickWin32AppPackage -Force (which names its output from the setup
        # file, so two recipes can collide) from clobbering a file the thread is
        # still reading. Nothing in here may throw - a bad backup share must not
        # turn a successful upload into a failed application.
        if ($Script:BackupEnabled) {
            try {
                $BackupFileName = Get-YardstickBackupFileName -AppId $Script:Id -Version $Script:Version -Timestamp $UploadedAt
                $StageDir = Join-Path (Join-Path $Published '_BackupQueue') $Script:Id
                if (-not (Test-Path -LiteralPath $StageDir)) {
                    New-Item -ItemType Directory -Path $StageDir -Force | Out-Null
                }
                $StagedPath = Join-Path $StageDir $BackupFileName
                Invoke-WithRetry -Label "Stage .intunewin for backup" -MaxRetries 3 -DelaySeconds 1 `
                    -ScriptBlock { Move-Item -LiteralPath $App.path -Destination $StagedPath -Force } `
                    -VerifyBlock { Test-Path -LiteralPath $StagedPath } | Out-Null
                if (-not (Test-Path -LiteralPath $StagedPath)) {
                    throw "could not stage $($App.path)"
                }
                Start-YardstickBackup -SourcePath $StagedPath -BackupRoot $Script:Backup -AppId $Script:Id `
                    -FileName $BackupFileName -VersionsToKeep $Script:BackupVersionsToKeep
                Write-Log "Queued a backup of $BackupFileName to $Script:Backup." -Level Detail
            } catch {
                # Send-YardstickLogReport.ps1 matches "Could not queue backup for <id>: ".
                Write-Log "Could not queue backup for $($Script:Id): $_" -Level Warning
            }
        }

        # The payload is published, so retire it out of the dropbox. Leaving the
        # dropbox empty is what makes the next run skip this recipe rather than
        # republishing the same build. A failure here must not fail the app - the
        # operator just has to clear the folder by hand.
        if ($Script:ManualDownload) {
            try {
                Complete-YardstickDropbox -DropboxPath $dropboxPath -ArchiveRoot $Script:SoftwareArchive `
                    -Folder $Script:ManualDownloadFolder -Version $Script:Version
            } catch {
                Write-Log "Could not archive the manual drop: $_" -Level Warning
            }
        }


        ###################################################
        # SUPERSEDENCE, RETENTION, AUTO-UPDATE
        ###################################################

        $Maintenance = Invoke-YardstickAppMaintenance -NewAppId $Win32App.id
        $CurrentApp        = $Maintenance.CurrentApp
        $supersededCount   = $Maintenance.SupersededCount
        $autoUpdateApplied = $Maintenance.AutoUpdateApplied
        $anchorStatus      = $Maintenance.AnchorStatus

        $ActionPerformed = if ($Force) { "Force Updated" } elseif ($Repair) { "Repaired" } else { "Updated" }
        $autoStatusParts = @()
        if ($supersededCount -gt 0)   { $autoStatusParts += "supersedes $supersededCount" }
        if ($autoUpdateApplied -gt 0) { $autoStatusParts += "auto-update on $autoUpdateApplied assignment(s)" }
        if ($anchorStatus)            { $autoStatusParts += "anchor: $anchorStatus" }
        $autoStatus = if ($autoStatusParts.Count -gt 0) { $autoStatusParts -join ', ' } else { $null }
        Add-SuccessfulApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -Action $ActionPerformed -Dependents $DependentUpdateStatus -AutoUpdateStatus $autoStatus -MissingIcon $missingIcon

    } catch {
        # Catch any unexpected errors during processing
        Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "$_" -FailureStage "General Processing"
        if (-not((Get-Location).ProviderPath -eq $YardstickModuleRoot)) {
            Set-Location $YardstickModuleRoot
        }
    } finally {
        # In a finally because the stage handlers above skip to the next recipe with
        # `continue`, which leaves the try without ever entering the catch. The
        # post-run script is a recipe's cleanup hook - crowdstrikefalcon revokes its
        # OAuth token there - so it has to run on the failure path too, and the
        # backup drain has to happen before the next iteration starts deleting
        # .intunewin files.

        # Run the post-run script
        if ($Script:PostRunScript) {
            Write-Log "Running the post-run script..." -Level Detail
            Set-YardstickWatchdogStage -Stage 'Post-Run Script'
            try {
                Invoke-Command -ScriptBlock $Script:PostRunScript -NoNewScope
                Write-Log "Ran the post-run script." -Level Detail
            } catch {
                # Not Write-Error: $ErrorActionPreference is 'Stop', so a Write-Error
                # here would throw out of the finally and kill the run.
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "$_" -FailureStage "Post-Run Script"
            }
        }

        # Drain this app's backup before moving on, so at most one copy is ever in
        # flight and an abrupt exit in the next iteration cannot orphan it. The
        # thread still ran concurrently with all the supersedence and assignment work
        # above, which is where the time actually goes.
        Wait-YardstickBackup -TimeoutSeconds 600 | Out-Null
        Clear-YardstickLogContext
    }
}
Write-Log "Finished processing $(@($Applications).Count) recipe(s)."

# The watchdog stops at the end of the loop, not at the end of the script. What
# follows is the email report, which can sit on a Read-Host waiting for an
# operator to approve sending - killing a run for being slow to answer a prompt is
# exactly the behaviour nobody wants. Everything past here is either bounded
# already or deliberately interactive.
Stop-YardstickWatchdog

# Final drain. This has to happen before the email report is built, because
# Wait-YardstickBackup is what records each app's backup status on it.
Wait-YardstickBackup -TimeoutSeconds 600 | Out-Null

# Send email report if enabled. Checked here as well as in Send-YardstickEmailReport,
# so a run with email turned off neither prompts nor logs a send that never happens.
$sendReport = -not $NoEmail -and $Prefs.emailNotificationEnabled

# If prompt before sending is enabled, ask the user for confirmation.
# An unattended run has nobody to answer, so it sends without asking.
if ($sendReport -and $Prefs.emailPromptBeforeSending -and -not $NoInteractive) {
    # play notification sound if specified
    if ($Prefs.emailNotificationSoundFile -and (Test-Path $Prefs.emailNotificationSoundFile)) {
        try {
            $sound = New-Object System.Media.SoundPlayer($Prefs.emailNotificationSoundFile)
            $sound.Play()
        } catch {
            Write-Log "Could not play the notification sound: $_" -Level Warning
        }
    }
    try {
        $confirmation = Read-Host "Do you want to send the email report now? (Y/N)"
    } catch {
        # Read-Host throws when PowerShell was started with -NonInteractive.
        Write-Log "Could not ask whether to send the email report: $_" -Level Warning
        $confirmation = 'N'
    }
    # A flag rather than an early return, so cleanup below still runs.
    if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
        Write-Log "Not sending the email report: declined at the prompt."
        $sendReport = $false
    }
}

if ($sendReport) {
    try {
        # Select email delivery method based on preferences
        $emailMethod = if ($null -ne $Prefs.emailDeliveryMethod) { $Prefs.emailDeliveryMethod } else { "outlook" }

        switch ($emailMethod.ToLower()) {
            "outlook" {
                Write-Log "Sending the email report through Outlook..."
                Send-YardstickEmailReport -Preferences $Prefs -RunParameters $RunParameters
            }
            "smtp" {
                Write-Log "Sending the email report via SMTP ($($Prefs.smtpServer))..."
                Send-YardstickEmailReport -Preferences $Prefs -RunParameters $RunParameters
            }
            default {
                Write-Log "Unknown email delivery method '$emailMethod'. Using Outlook." -Level Warning
                Send-YardstickEmailReport -Preferences $Prefs -RunParameters $RunParameters
            }
        }
    } catch {
        Write-Log "Could not send the email report: $_" -Level Warning
    }
}

# Clean up
Invoke-Cleanup

# Explicit, so Invoke-Yardstick reads a success code rather than whatever the last
# native command in the run happened to leave in $LASTEXITCODE.
exit 0
