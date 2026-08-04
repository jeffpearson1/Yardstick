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
      3. Strips existing supersedence off every kept version.
      4. Rebuilds supersedence on the newest app, targeting all other kept
         versions (excluding the anchor). Uses `Replace` when
         `uninstallPreviousVersion` is true, else `Update`. Only runs when the
         recipe involves available-intent assignments.
      5. Sets the assignment-level `autoUpdate` flag on the newest app's
         available-intent assignments when `autoUpdateOnAssignment` (or legacy
         `autoUpdate`) is true.
      6. For each older version: MOVES required-intent assignments (and
         dependencies) to the newest, then COPIES available-intent assignments
         to the newest (leaving them on the source).

    Idempotent - safe to re-run. Use -WhatIf for a dry run.

.PARAMETER ApplicationId
    Restrict to a single recipe id. If omitted, all recipes are processed.

.PARAMETER SkipMove
    Do not run the intent-based assignment migration (move required / copy
    available) as the final step. Useful when assignments have already been
    consolidated by a previous migration pass.

.EXAMPLE
    .\Migrate-ToSupersedenceModel.ps1 -ApplicationId 7zip -WhatIf

.EXAMPLE
    .\Migrate-ToSupersedenceModel.ps1
#>
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
              Where-Object { $_.DisplayName -ne "{DETECT} $displayName" } |
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
                    $anchor = Get-DetectAnchor -DisplayName $displayName
                    Write-Host "  pinned {DETECT} anchor: $($oldest.displayVersion)" -ForegroundColor Green
                }
            }
        } else {
            Write-Host "  {DETECT} anchor already exists: $($anchor.displayVersion)"
        }
    }

    # Reload after possible rename
    $all = Get-SameAppAllVersions $displayName
    $nonAnchor = @($all | Where-Object { -not $anchor -or $_.id -ne $anchor.id })
    $olderKept = @($nonAnchor | Where-Object id -ne $newest.id |
                   Sort-Object @{Expression = {[VersionPro]$_.displayVersion}; Descending = $true} |
                   Select-Object -First ([Math]::Max(0, $numKeep - 1)))

    # Detect whether any older version has available-intent assignments (or
    # whether the newest already does). Supersedence + autoUpdate only apply
    # when the recipe involves available deployments.
    $availableInPlay = $false
    foreach ($app in $all) {
        $avail = @(Get-IntuneWin32AppAssignment -Id $app.id | Where-Object Intent -eq 'available')
        if ($avail.Count -gt 0) { $availableInPlay = $true; break }
    }

    # 2. Supersedence (only when available assignments are involved)
    if ($supersedence -and $availableInPlay) {
        $type = if ($uninstallPrev) { 'Replace' } else { 'Update' }
        if ($PSCmdlet.ShouldProcess($newest.DisplayName, "attach supersedence ($type) over $($olderKept.Count) target(s)")) {
            Set-YardstickSupersedence -NewApp $newest -SupersededApps $olderKept -Type $type | Out-Null
        }
    } elseif ($supersedence) {
        Write-Host "  skipping supersedence - no available-intent assignments in play" -ForegroundColor DarkGray
    }

    # 3. autoUpdate on assignments (only for available intent)
    if ($autoUpdate -and $availableInPlay) {
        if ($PSCmdlet.ShouldProcess($newest.DisplayName, "set autoUpdate on available-intent assignments")) {
            Set-AssignmentAutoUpdate -AppId $newest.Id -Enabled $true -IntentFilter 'available' | Out-Null
        }
    }

    # 4. Migrate lingering assignments from older versions to newest, split
    #    by intent: MOVE required, COPY available.
    if (-not $SkipMove) {
        foreach ($old in $olderKept) {
            $assigns = Get-IntuneWin32AppAssignment -Id $old.id
            if (-not $assigns -or $assigns.Count -eq 0) { continue }

            $reqCount   = @($assigns | Where-Object Intent -eq 'required').Count
            $availCount = @($assigns | Where-Object Intent -eq 'available').Count

            if ($reqCount -gt 0 -and $PSCmdlet.ShouldProcess($old.DisplayName, "move $reqCount required assignment(s) to newest")) {
                try {
                    Move-AssignmentsAndDependencies -From $old -To $newest -AvailableDateOffset 0 -DeadlineDateOffset 0 -IntentFilter 'required'
                } catch {
                    Write-Warning "  failed moving required assignments from $($old.DisplayName): $_"
                }
            }
            if ($availCount -gt 0 -and $PSCmdlet.ShouldProcess($old.DisplayName, "copy $availCount available assignment(s) to newest")) {
                try {
                    Move-AssignmentsAndDependencies -From $old -To $newest -AvailableDateOffset 0 -DeadlineDateOffset 0 -IntentFilter 'available' -CopyOnly -SkipDependencies
                } catch {
                    Write-Warning "  failed copying available assignments from $($old.DisplayName): $_"
                }
            }
        }
    }
}

Write-Host "`nMigration pass complete." -ForegroundColor Cyan
