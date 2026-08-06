<#
.SYNOPSIS
    One-shot migration from Yardstick's assignment-migration model to the
    supersedence + auto-update model.

.DESCRIPTION
    For every recipe (or a single recipe when -ApplicationId is given), this
    script:

      1. Fetches all Intune versions of the app via Get-SameAppAllVersions.
      2. If the recipe is version-detection and no {DETECT} anchor exists yet,
         pins the oldest surviving version as `{DETECT} <DisplayName>`.
      3. COPIES available-intent assignments from every non-anchor version below
         the newest onto the newest app, leaving them on the source (removing an
         available assignment destroys the on-device auto-update component). This
         sweeps versions outside the keep window too - they never become
         supersedence targets, so their Company Portal installs would otherwise
         have no assignment on the newest app for auto-update to act on.
      4. Strips existing supersedence off every kept version so the newest app
         is the only superseding parent.
      5. Rebuilds supersedence on the newest app, targeting all kept versions
         plus the {DETECT} anchor. Uses `Replace` when
         `uninstallPreviousVersion` is true, else `Update`; the anchor is always
         `Update` so a Replace can never mass-uninstall the app.
      6. MOVES required-intent assignments (and dependencies) from each kept
         older version to the newest.
      7. Enables Intune's native auto-update
         (`autoUpdateSettings.autoUpdateSupersededAppsState = enabled`) on the
         newest app's available-intent assignments when `autoUpdateOnAssignment`
         (or legacy `autoUpdate`) is true. Runs last so it also covers the
         assignments copied in step 3. Intune only honours auto-update for
         available assignments.

    Idempotent - safe to re-run. Use -WhatIf for a dry run.

.PARAMETER ApplicationId
    Restrict to a single recipe id. If omitted, all recipes are processed.

.PARAMETER SkipMove
    Do not run the intent-based assignment migration (copy available / move
    required). Useful when assignments have already been consolidated by a
    previous migration pass.

.EXAMPLE
    .\Migrate-ToSupersedenceModel.ps1 -ApplicationId 7zip -WhatIf

.EXAMPLE
    .\Migrate-ToSupersedenceModel.ps1
#>
using module .\Modules\VersionPro.psm1
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ApplicationId,
    [switch]$SkipMove
)

$ErrorActionPreference = 'Stop'

Import-Module powershell-yaml -Force
Import-Module IntuneWin32App -Force
Import-Module "$PSScriptRoot\Modules\YardstickSupport.psm1" -Scope Global -Force

# Load preferences (same location Yardstick.ps1 uses)
$Prefs = Get-Content "$PSScriptRoot\preferences.yaml" | ConvertFrom-Yaml
$Global:TenantID     = $Prefs.TenantID
$Global:ClientID     = $Prefs.ClientId
$Global:ClientSecret = $Prefs.ClientSecret
$Global:LogLocation = "G:\Intune\YardstickDev"
$Global:LogFile = "SupersedenceMigration.log"
$Recipes = $Prefs.Recipes

# Auth
Connect-AutoMSIntuneGraph -Force

# Enumerate recipes
$recipeFiles = Get-ChildItem $Recipes -Recurse -File -Include *.yaml,*.yml |
    Where-Object { $_.FullName -notmatch '\\Disabled\\' }

if ($ApplicationId) {
    $recipeFiles = $recipeFiles | Where-Object BaseName -eq $ApplicationId
    if (-not $recipeFiles) {
        Write-Error "No recipe found with id '$ApplicationId'"
        exit 1
    }
}

foreach ($file in $recipeFiles) {
    $recipe = Get-Content $file.FullName | ConvertFrom-Yaml
    $recipe = Merge-RecipeWithBase -Recipe $recipe -RecipesPath $Recipes

    $displayName = $recipe.displayName
    if (-not $displayName) {
        Write-Warning "Skipping $($file.Name) - missing displayName"
        continue
    }

    Write-Host "`n=== $displayName ($($file.BaseName)) ===" -ForegroundColor Cyan

    $useAnchor              = if ($null -ne $Prefs.useDetectAnchor) { [bool]$Prefs.useDetectAnchor } else { $true }
    $supersedence           = if ($null -ne $recipe.supersedence) { [bool]$recipe.supersedence } elseif ($null -ne $Prefs.defaultSupersedence) { [bool]$Prefs.defaultSupersedence } else { $true }
    $uninstallPrev          = if ($null -ne $recipe.uninstallPreviousVersion) { [bool]$recipe.uninstallPreviousVersion } elseif ($null -ne $Prefs.defaultUninstallPreviousVersion) { [bool]$Prefs.defaultUninstallPreviousVersion } else { $false }
    $autoUpdate             = if ($null -ne $recipe.autoUpdateOnAssignment) { [bool]$recipe.autoUpdateOnAssignment } elseif ($null -ne $recipe.autoUpdate) { [bool]$recipe.autoUpdate } elseif ($null -ne $Prefs.defaultAutoUpdate) { [bool]$Prefs.defaultAutoUpdate } else { $true }
    $numKeep                = if ($null -ne $recipe.numVersionsToKeep) { [int]$recipe.numVersionsToKeep } else { [int]$Prefs.defaultNumVersionsToKeep }

    $all = Get-SameAppAllVersions $displayName
    if (-not $all -or $all.Count -eq 0) {
        Write-Host "  no versions in Intune - skipping" -ForegroundColor DarkGray
        continue
    }

    $newest = $all | Sort-Object @{Expression = {[VersionPro]$_.displayVersion}; Descending = $true} |
              Where-Object { $_.DisplayName -ne (Get-DetectAnchorName -DisplayName $displayName) } |
              Select-Object -First 1
    if (-not $newest) {
        Write-Host "  only the {DETECT} anchor exists - nothing to migrate" -ForegroundColor DarkGray
        continue
    }
    Write-Host "  newest: $($newest.displayVersion) ($($newest.id))"

    # 1. Anchor
    $anchor = $null
    if ($useAnchor -and (Test-IsVersionDetection -DetectionType $recipe.detectionType -FileDetectionMethod $recipe.fileDetectionMethod -RegistryDetectionMethod $recipe.registryDetectionMethod)) {
        $anchor = Get-DetectAnchor -DisplayName $displayName
        if (-not $anchor) {
            $others = @($all | Where-Object id -ne $newest.id)
            if ($others.Count -gt 0) {
                $oldest = $others | Sort-Object @{Expression = {[VersionPro]$_.displayVersion}} | Select-Object -First 1
                if ($PSCmdlet.ShouldProcess($oldest.DisplayName, "pin as {DETECT} anchor")) {
                    Set-DetectAnchor -App $oldest -DisplayName $displayName
                    # Hold on to the candidate in case the lookup is still stale -
                    # otherwise the app we just pinned would be treated as a normal
                    # target and could be superseded with 'Replace'.
                    $anchor = $oldest
                    $refreshed = Get-DetectAnchor -DisplayName $displayName
                    if ($refreshed) { $anchor = $refreshed }
                    Write-Host "  pinned {DETECT} anchor: $($anchor.displayVersion)" -ForegroundColor Green
                }
            }
        } else {
            Write-Host "  {DETECT} anchor already exists: $($anchor.displayVersion)"
        }
    }

    # Reload after possible rename
    $all = Get-SameAppAllVersions $displayName
    $nonAnchor = @($all | Where-Object { -not $anchor -or $_.id -ne $anchor.id })
    $allOlder = @($nonAnchor | Where-Object id -ne $newest.id |
                  Sort-Object @{Expression = {[VersionPro]$_.displayVersion}; Descending = $true})
    $olderKept = @($allOlder | Select-Object -First ([Math]::Max(0, $numKeep - 1)))

    # 2. Available-intent assignments, swept from EVERY non-anchor version below
    #    the newest rather than just the ones inside the keep window. Versions
    #    outside that window are never supersedence targets (Intune caps the graph
    #    at 10 nodes), so a Company Portal install from one of them would otherwise
    #    be left with no assignment on the newest app and nothing for auto-update
    #    to act on; the {DETECT} anchor still supplies the detection side. This runs
    #    before supersedence and auto-update so step 4 configures everything landed
    #    here.
    #
    #    Copy rather than move while supersedence is on: Intune documents that "any
    #    application assignment changes delete the component responsible for
    #    auto-updating the app", so stripping the available assignment off the
    #    version a user actually installed breaks auto-update for exactly the
    #    devices being targeted. With supersedence off there is no update path to
    #    protect, so fall back to moving instead of leaving a duplicate Company
    #    Portal listing.
    if (-not $SkipMove) {
        $verb = if ($supersedence) { 'copy' } else { 'move' }
        foreach ($old in $allOlder) {
            # KNOWN ISSUE: this count is 0 for any app with exactly one
            # assignment - see the note in Move-AssignmentsAndDependencies.
            $availCount = @(Get-IntuneWin32AppAssignment -Id $old.id | Where-Object Intent -eq 'available').Count
            if ($availCount -eq 0) { continue }

            if ($PSCmdlet.ShouldProcess($old.DisplayName, "$verb $availCount available assignment(s) to newest")) {
                try {
                    Move-AssignmentsAndDependencies -From $old -To $newest -AvailableDateOffset 0 -DeadlineDateOffset 0 `
                        -IntentFilter 'available' -CopyOnly:$supersedence -SkipDependencies
                } catch {
                    Write-Warning "  failed to $verb available assignments from $($old.DisplayName): $_"
                }
            }
        }
    }

    # 3. Supersedence over kept versions + the anchor. The anchor is forced to
    #    'Update' so a 'Replace' recipe cannot uninstall the app fleet-wide.
    if ($supersedence) {
        $type = if ($uninstallPrev) { 'Replace' } else { 'Update' }
        $targets = @($olderKept)
        if ($anchor) { $targets += $anchor }
        if ($PSCmdlet.ShouldProcess($newest.DisplayName, "attach supersedence ($type) over $($targets.Count) target(s)")) {
            Set-YardstickSupersedence -NewApp $newest -SupersededApps $targets -Type $type `
                -UpdateOnlyIds @(if ($anchor) { $anchor.id }) | Out-Null
        }
    }

    # 4. Migrate lingering required-intent assignments from the kept older
    #    versions to the newest. Required deployments install unconditionally, so
    #    consolidating them onto a single app is correct and they are moved, not
    #    copied. This pass also carries the dependency migration, so it runs for
    #    every kept old app even when there are no required assignments to move.
    #    Available intent was already handled in step 2.
    if (-not $SkipMove) {
        foreach ($old in $olderKept) {
            $reqCount = @(Get-IntuneWin32AppAssignment -Id $old.id | Where-Object Intent -eq 'required').Count

            if ($PSCmdlet.ShouldProcess($old.DisplayName, "move $reqCount required assignment(s) and any dependencies to newest")) {
                try {
                    Move-AssignmentsAndDependencies -From $old -To $newest -AvailableDateOffset 0 -DeadlineDateOffset 0 -IntentFilter 'required'
                } catch {
                    Write-Warning "  failed moving required assignments from $($old.DisplayName): $_"
                }
            }
        }
    }

    # 5. Native auto-update, applied last so it covers the available assignments
    #    copied in step 2. Intune only honours this for available intent, and
    #    Set-AssignmentAutoUpdate is a no-op when there are no such assignments.
    if ($autoUpdate) {
        if ($PSCmdlet.ShouldProcess($newest.DisplayName, "enable auto-update on available-intent assignments")) {
            Set-AssignmentAutoUpdate -AppId $newest.Id -Enabled $true -IntentFilter 'available' | Out-Null
        }
    }
}

Write-Host "`nMigration pass complete." -ForegroundColor Cyan
