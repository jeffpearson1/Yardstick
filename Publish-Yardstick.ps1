<#
.SYNOPSIS
Publishes explicitly selected Yardstick development content to production.

.DESCRIPTION
Production is a deployment target, not the development workspace. This command
copies only allowlisted, Git-tracked runtime files and/or explicitly named
recipes. Existing destination files are backed up beneath Artifacts before they
are replaced. The command never deletes production files.

.EXAMPLE
.\Publish-Yardstick.ps1 -Runtime -WhatIf

.EXAMPLE
.\Publish-Yardstick.ps1 -RecipeId dragonframe, marcedit -WhatIf
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [switch]$Runtime,

    [ValidatePattern('^[A-Za-z0-9._-]+$')]
    [string[]]$RecipeId,

    [string]$ProductionPath
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ProjectLayout.ps1')
$layout = Get-YardstickProjectLayout -ProjectRoot $PSScriptRoot

if (-not $Runtime -and -not $RecipeId) {
    throw 'Specify -Runtime, one or more -RecipeId values, or both.'
}

if (-not $ProductionPath) {
    $ProductionPath = $layout.ProductionRoot
}
$productionRoot = (Resolve-Path -LiteralPath $ProductionPath -ErrorAction Stop).Path
$projectRoot = $layout.ProjectRoot
if ($productionRoot.Equals($projectRoot, [StringComparison]::OrdinalIgnoreCase) -or
    $productionRoot.StartsWith($projectRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "Production path must be outside the development project: $productionRoot"
}
if (-not (Test-Path -LiteralPath (Join-Path $productionRoot 'Yardstick.psd1') -PathType Leaf)) {
    throw "The production target does not look like Yardstick: $productionRoot"
}

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backupRoot = Join-Path $layout.ArtifactsPath "PublishBackups\$timestamp"
$reportRoot = Join-Path $layout.ArtifactsPath 'PublishReports'
$operations = [Collections.Generic.List[object]]::new()

function Publish-YardstickFile {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$RelativePath
    )

    $sourcePath = (Resolve-Path -LiteralPath $Source -ErrorAction Stop).Path
    $relativeWindowsPath = $RelativePath.Replace('/', '\')
    $destinationPath = [IO.Path]::GetFullPath((Join-Path $productionRoot $relativeWindowsPath))
    if (-not $destinationPath.StartsWith($productionRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Publish destination escaped the production root: $destinationPath"
    }

    if ($PSCmdlet.ShouldProcess($destinationPath, "Publish '$relativeWindowsPath'")) {
        $backupPath = $null
        if (Test-Path -LiteralPath $destinationPath -PathType Leaf) {
            $backupPath = Join-Path $backupRoot $relativeWindowsPath
            New-Item -ItemType Directory -Path (Split-Path -Parent $backupPath) -Force | Out-Null
            Copy-Item -LiteralPath $destinationPath -Destination $backupPath -Force
        }

        New-Item -ItemType Directory -Path (Split-Path -Parent $destinationPath) -Force | Out-Null
        Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
        $operations.Add([pscustomobject]@{
            RelativePath = $relativeWindowsPath
            Source       = $sourcePath
            Destination  = $destinationPath
            Backup       = $backupPath
        })
    }
}

if ($Runtime) {
    $runtimeSpecs = @(
        'Branding/**',
        'Modules/**',
        'Scripts/**',
        'Templates/**',
        'Tools/**',
        'LICENSE.md',
        'Deploy-YardstickApps.ps1',
        'Migrate-ToSupersedenceModel.ps1',
        'ProjectLayout.ps1',
        'RecipeGroups.yaml',
        'Send-YardstickLogReport.ps1',
        'Set-YardstickCredential.ps1',
        'Sync-RecipeGroup.ps1',
        'Test-EmailNotification.ps1',
        'Test-MgearDeployment.ps1',
        'Yardstick.Project.psd1',
        'Yardstick.ps1',
        'Yardstick.psd1'
    )
    $gitArgs = @('-C', $projectRoot, 'ls-files', '--') + $runtimeSpecs
    $runtimeFiles = @(& git @gitArgs)
    if ($LASTEXITCODE -ne 0) {
        throw 'Unable to enumerate the Git-tracked runtime allowlist.'
    }

    # These files must travel with the layout-aware runtime even before the
    # first commit that introduces them.
    foreach ($required in 'ProjectLayout.ps1', 'Yardstick.Project.psd1') {
        if ((Test-Path -LiteralPath (Join-Path $projectRoot $required)) -and $required -notin $runtimeFiles) {
            $runtimeFiles += $required
        }
    }

    foreach ($relativePath in ($runtimeFiles | Sort-Object -Unique)) {
        Publish-YardstickFile -Source (Join-Path $projectRoot $relativePath) -RelativePath $relativePath
    }
}

foreach ($id in $RecipeId) {
    $developmentRecipe = Join-Path $layout.DevelopmentRecipes "$id.yaml"
    $standardRecipe = Join-Path $layout.RecipesPath "$id.yaml"
    $recipePath = if (Test-Path -LiteralPath $developmentRecipe -PathType Leaf) {
        $developmentRecipe
    } elseif (Test-Path -LiteralPath $standardRecipe -PathType Leaf) {
        $standardRecipe
    } else {
        throw "Recipe '$id' was not found in Development or the standard recipe directory."
    }

    Publish-YardstickFile -Source $recipePath -RelativePath "Recipes\$id.yaml"

    if (-not (Get-Command ConvertFrom-Yaml -ErrorAction SilentlyContinue)) {
        Import-Module powershell-yaml -ErrorAction Stop
    }
    $recipe = Get-Content -LiteralPath $recipePath -Raw | ConvertFrom-Yaml
    if ($recipe.iconFile) {
        $iconPath = Join-Path $layout.IconsPath ([string]$recipe.iconFile)
        if (-not (Test-Path -LiteralPath $iconPath -PathType Leaf)) {
            throw "Recipe '$id' references missing icon '$($recipe.iconFile)'."
        }
        Publish-YardstickFile -Source $iconPath -RelativePath "Icons\$($recipe.iconFile)"
    }
}

if ($operations.Count) {
    New-Item -ItemType Directory -Path $reportRoot -Force | Out-Null
    $reportPath = Join-Path $reportRoot "publish-$timestamp.json"
    [pscustomobject]@{
        Timestamp      = (Get-Date).ToString('o')
        ProjectRoot    = $projectRoot
        ProductionRoot = $productionRoot
        Runtime        = [bool]$Runtime
        RecipeIds      = @($RecipeId)
        Operations     = $operations
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $reportPath -Encoding utf8
    Write-Host "Published $($operations.Count) file(s). Report: $reportPath"
} elseif ($WhatIfPreference) {
    Write-Host 'WhatIf completed; production was not changed.'
} else {
    Write-Host 'No files required publication.'
}
