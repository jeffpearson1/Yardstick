[CmdletBinding(DefaultParameterSetName = 'ById')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ById')]
    [ValidatePattern('^[A-Za-z0-9._-]+$')]
    [string] $Id,

    [Parameter(Mandatory, ParameterSetName = 'ByPath')]
    [string] $RecipePath,

    [string] $YardstickDevPath,

    [string] $RecipesPath,

    [string] $IconsPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..') -ErrorAction Stop).Path
. (Join-Path $projectRoot 'ProjectLayout.ps1')
$projectLayout = Get-YardstickProjectLayout -ProjectRoot $projectRoot
if (-not $YardstickDevPath) { $YardstickDevPath = $projectLayout.RuntimeRoot }
if (-not $RecipesPath) { $RecipesPath = $projectLayout.DevelopmentRecipes }
if (-not $IconsPath) { $IconsPath = $projectLayout.IconsPath }

$workspaceRecipes = (Resolve-Path -LiteralPath $RecipesPath -ErrorAction Stop).Path
$workspaceIcons = (Resolve-Path -LiteralPath $IconsPath -ErrorAction Stop).Path
$yardstickModule = Join-Path $YardstickDevPath 'Modules\YardstickSupport.psm1'
$yardstickRecipes = Join-Path $YardstickDevPath 'Recipes'

if ((Split-Path -Leaf $YardstickDevPath) -ne 'YardstickDev') {
    throw "Safety check failed: YardstickDevPath must point to a folder named 'YardstickDev'."
}
if (-not (Test-Path -LiteralPath $yardstickModule -PathType Leaf)) {
    throw "Yardstick support module not found at '$yardstickModule'."
}

Import-Module powershell-yaml -MinimumVersion 0.4.12 -ErrorAction Stop
Import-Module $yardstickModule -Force -ErrorAction Stop

if ($PSCmdlet.ParameterSetName -eq 'ById') {
    $RecipePath = Join-Path $workspaceRecipes "$Id.yaml"
}
$RecipePath = (Resolve-Path -LiteralPath $RecipePath -ErrorAction Stop).Path
$recipeId = [IO.Path]::GetFileNameWithoutExtension($RecipePath)
$recipe = Get-Content -LiteralPath $RecipePath -Raw | ConvertFrom-Yaml

if ($recipe.ContainsKey('base')) {
    $baseId = [string]$recipe.base
    $baseWorkspacePath = Join-Path $workspaceRecipes "$baseId.yaml"
    $baseLookupPath = if (Test-Path -LiteralPath $baseWorkspacePath) {
        $workspaceRecipes
    } else {
        $yardstickRecipes
    }
    $recipe = Merge-RecipeWithBase -Recipe $recipe -RecipesPath $baseLookupPath
}

$result = Test-RecipeSchema -Recipe $recipe -RecipeId $recipeId
$errors = [Collections.Generic.List[string]]::new()
foreach ($errorMessage in $result.Errors) {
    $errors.Add($errorMessage)
}

if (-not $recipe.ContainsKey('id') -or $recipe.id -cne $recipeId) {
    $errors.Add("Recipe id '$($recipe.id)' must exactly match filename '$recipeId'.")
}

if ($recipe.ContainsKey('iconFile') -and -not [string]::IsNullOrWhiteSpace($recipe.iconFile)) {
    $workspaceIcon = Join-Path $workspaceIcons $recipe.iconFile
    $yardstickIcon = Join-Path (Join-Path $YardstickDevPath 'Icons') $recipe.iconFile
    if (-not (Test-Path -LiteralPath $workspaceIcon) -and -not (Test-Path -LiteralPath $yardstickIcon)) {
        $errors.Add("Icon '$($recipe.iconFile)' was not found in '$workspaceIcons' or YardstickDev/Icons.")
    }
}

foreach ($warning in $result.Warnings) {
    Write-Warning $warning
}

if ($errors.Count -gt 0) {
    foreach ($errorMessage in $errors) {
        Write-Error $errorMessage -ErrorAction Continue
    }
    throw "Recipe '$recipeId' failed schema validation with $($errors.Count) error(s)."
}

[PSCustomObject]@{
    Id = $recipeId
    RecipePath = $RecipePath
    DetectionType = $recipe.detectionType
    Valid = $true
    WarningCount = @($result.Warnings).Count
}
