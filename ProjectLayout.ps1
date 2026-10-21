function Get-YardstickProjectLayout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProjectRoot
    )

    # ProviderPath, not Path: on a UNC share Path can come back provider-qualified
    # (Microsoft.PowerShell.Core\FileSystem::\\server\...), which breaks Join-Path.
    $root = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
    $manifestPath = Join-Path $root 'Yardstick.Project.psd1'
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "Yardstick project manifest was not found at '$manifestPath'."
    }

    $manifest = Import-PowerShellDataFile -LiteralPath $manifestPath
    $resolved = [ordered]@{ ProjectRoot = $root; ManifestPath = $manifestPath }
    foreach ($key in $manifest.Keys) {
        $value = [string]$manifest[$key]
        $resolved[$key] = if ([IO.Path]::IsPathRooted($value)) {
            [IO.Path]::GetFullPath($value)
        } else {
            [IO.Path]::GetFullPath((Join-Path $root $value))
        }
    }
    [pscustomobject]$resolved
}

function Get-YardstickPreferencesPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProjectRoot
    )

    $layout = Get-YardstickProjectLayout -ProjectRoot $ProjectRoot
    if (Test-Path -LiteralPath $layout.Preferences -PathType Leaf) {
        return $layout.Preferences
    }

    # Production installations created before the monorepo layout keep their
    # machine-local preferences beside Yardstick.ps1.
    $legacyPath = Join-Path $layout.ProjectRoot 'preferences.yaml'
    if (Test-Path -LiteralPath $legacyPath -PathType Leaf) {
        return $legacyPath
    }

    return $layout.Preferences
}

function Get-YardstickRecipeOverridesPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProjectRoot
    )

    $layout = Get-YardstickProjectLayout -ProjectRoot $ProjectRoot
    if ($layout.PSObject.Properties['RecipeOverrides']) {
        return $layout.RecipeOverrides
    }

    # A project manifest from before recipe overrides existed has no entry for them.
    return Join-Path $layout.ProjectRoot 'Local\recipeOverrides.yaml'
}
