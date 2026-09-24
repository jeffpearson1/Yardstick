[CmdletBinding(DefaultParameterSetName = 'ById')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ById')]
    [ValidatePattern('^[A-Za-z0-9._-]+$')]
    [string] $Id,

    [Parameter(Mandatory, ParameterSetName = 'ByPath')]
    [string] $RecipePath,

    [ValidateSet('Schema', 'Download', 'Package')]
    [string] $Stage = 'Schema',

    [string] $YardstickDevPath,

    [string] $RecipesPath,

    [string] $IconsPath,

    [string] $ArtifactsPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..') -ErrorAction Stop).Path
. (Join-Path $projectRoot 'ProjectLayout.ps1')
$projectLayout = Get-YardstickProjectLayout -ProjectRoot $projectRoot
if (-not $YardstickDevPath) { $YardstickDevPath = $projectLayout.RuntimeRoot }
if (-not $RecipesPath) { $RecipesPath = $projectLayout.DevelopmentRecipes }
if (-not $IconsPath) { $IconsPath = $projectLayout.IconsPath }
if (-not $ArtifactsPath) { $ArtifactsPath = $projectLayout.HarnessArtifacts }

$schemaScript = Join-Path $PSScriptRoot 'Test-RecipeSchema.ps1'
$workspaceRecipes = (Resolve-Path -LiteralPath $RecipesPath -ErrorAction Stop).Path
$workspaceIcons = (Resolve-Path -LiteralPath $IconsPath -ErrorAction Stop).Path
$artifactsRoot = $ArtifactsPath
$runStamp = Get-Date -Format 'yyyyMMdd-HHmmss'

if ($PSCmdlet.ParameterSetName -eq 'ById') {
    $RecipePath = Join-Path $workspaceRecipes "$Id.yaml"
}
$RecipePath = (Resolve-Path -LiteralPath $RecipePath -ErrorAction Stop).Path
$recipeId = [IO.Path]::GetFileNameWithoutExtension($RecipePath)

& $schemaScript -RecipePath $RecipePath -YardstickDevPath $YardstickDevPath -RecipesPath $workspaceRecipes -IconsPath $workspaceIcons | Out-Host
if ($Stage -eq 'Schema') {
    return
}

Import-Module powershell-yaml -MinimumVersion 0.4.12 -ErrorAction Stop
Import-Module (Join-Path $YardstickDevPath 'Modules\YardstickSupport.psm1') -Force -ErrorAction Stop
Import-Module BitsTransfer -ErrorAction Stop

$recipe = Get-Content -LiteralPath $RecipePath -Raw | ConvertFrom-Yaml
if ($recipe.ContainsKey('base')) {
    $baseId = [string]$recipe.base
    $baseWorkspacePath = Join-Path $workspaceRecipes "$baseId.yaml"
    $baseLookupPath = if (Test-Path -LiteralPath $baseWorkspacePath) {
        $workspaceRecipes
    } else {
        Join-Path $YardstickDevPath 'Recipes'
    }
    $recipe = Merge-RecipeWithBase -Recipe $recipe -RecipesPath $baseLookupPath
}

$runRoot = Join-Path (Join-Path $artifactsRoot $recipeId) $runStamp
$BuildSpace = Join-Path $runRoot 'build'
$Temp = Join-Path $runRoot 'temp'
$Scripts = Join-Path $YardstickDevPath 'Scripts'
$Published = Join-Path $runRoot 'packages'
$Backup = $null
$Recipes = $workspaceRecipes
$Icons = $workspaceIcons
$Tools = Join-Path $YardstickDevPath 'Tools'
$Secrets = Join-Path $PSScriptRoot 'empty-secrets'

foreach ($path in @($BuildSpace, $Temp, $Published)) {
    New-Item -ItemType Directory -Path $path -Force | Out-Null
}

# Initialize the functionally required recipe variables so StrictMode reports a
# useful validation error instead of an undefined-variable error.
$url = $null
$version = $null
$fileName = $null
foreach ($key in $recipe.Keys) {
    Set-Variable -Name $key -Value $recipe[$key] -Scope Script
}
$id = $recipeId

if ($recipe.ContainsKey('urlRedirects') -and $recipe['urlRedirects'] -eq $true -and $url) {
    $response = Invoke-WebRequest -Uri $url -Method Head -MaximumRedirection 10
    $url = $response.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
}

if ($recipe.ContainsKey('preDownloadScript') -and $recipe['preDownloadScript']) {
    Push-Location $YardstickDevPath
    try {
        . ([ScriptBlock]::Create([string]$recipe['preDownloadScript']))
    } finally {
        Pop-Location
    }
}
if ([string]::IsNullOrWhiteSpace([string]$version)) {
    throw 'Recipe did not set $version in YAML or preDownloadScript.'
}

$harnessBuildDirectory = Join-Path (Join-Path $BuildSpace $id) ([string]$version)
New-Item -ItemType Directory -Path $harnessBuildDirectory -Force | Out-Null
Push-Location $harnessBuildDirectory
try {
    if ($recipe.ContainsKey('downloadScript') -and $recipe['downloadScript']) {
        . ([ScriptBlock]::Create([string]$recipe['downloadScript']))
    } else {
        if ([string]::IsNullOrWhiteSpace([string]$url)) {
            throw 'Recipe did not set $url and has no downloadScript.'
        }
        if ([string]::IsNullOrWhiteSpace([string]$fileName)) {
            $fileName = [IO.Path]::GetFileName(([Uri]$url).AbsolutePath)
        }
        if ([string]::IsNullOrWhiteSpace([string]$fileName)) {
            throw 'Recipe did not set $fileName and it could not be derived from the URL.'
        }
        Start-BitsTransfer -Source $url -Destination (Join-Path $harnessBuildDirectory $fileName)
    }

    if ($recipe.ContainsKey('postDownloadScript') -and $recipe['postDownloadScript']) {
        . ([ScriptBlock]::Create([string]$recipe['postDownloadScript']))
    }
} finally {
    Pop-Location
}

$installerPath = Join-Path $harnessBuildDirectory $fileName
if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
    throw "Expected installer was not created at '$installerPath'."
}

$installer = Get-Item -LiteralPath $installerPath
$signature = Get-AuthenticodeSignature -LiteralPath $installerPath
$summary = [ordered]@{
    Id = $id
    Version = [string]$version
    Installer = $installer.FullName
    SizeBytes = $installer.Length
    SHA256 = (Get-FileHash -LiteralPath $installerPath -Algorithm SHA256).Hash
    SignatureStatus = $signature.Status.ToString()
    Signer = if ($signature.SignerCertificate) { $signature.SignerCertificate.Subject } else { $null }
    FileVersion = $installer.VersionInfo.FileVersion
    ProductVersion = $installer.VersionInfo.ProductVersion
}

if ($installer.Extension -ieq '.msi') {
    $productCode = @(Get-MsiProductCode -FilePath $installerPath) |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
        Select-Object -Last 1
    $msiProductVersion = @(Get-MsiProperty -Path $installerPath -PropertyName ProductVersion) |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
        Select-Object -Last 1
    $summary['ProductCode'] = [string]$productCode
    $summary['MsiProductVersion'] = [string]$msiProductVersion
}

if ($Stage -eq 'Package') {
    Import-Module IntuneWin32App -MinimumVersion 1.5.0 -MaximumVersion 1.5.2 -ErrorAction Stop
    if ($recipe.ContainsKey('powerShellInstallScript') -and $recipe['powerShellInstallScript']) {
        Set-Content -LiteralPath (Join-Path $harnessBuildDirectory 'install.ps1') -Value $recipe['powerShellInstallScript']
    }
    if ($recipe.ContainsKey('powerShellUninstallScript') -and $recipe['powerShellUninstallScript']) {
        Set-Content -LiteralPath (Join-Path $harnessBuildDirectory 'uninstall.ps1') -Value $recipe['powerShellUninstallScript']
    }
    $package = New-IntuneWin32AppPackage -SourceFolder $harnessBuildDirectory -SetupFile $fileName -OutputFolder $Published -Force -Verbose
    $summary['Package'] = $package.Path
}

[PSCustomObject]$summary
