function Get-YardstickProjectLayout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProjectRoot
    )

    $root = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).Path
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
