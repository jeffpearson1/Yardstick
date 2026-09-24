<#
.SYNOPSIS
Installs, detects, and uninstalls Yardstick recipes on this machine to prove
that a recipe actually works end to end.

.DESCRIPTION
For each recipe the harness runs these phases:

  1. Schema        - Test-RecipeSchema.ps1 (same rules the lab already enforces)
  2. Download      - urlRedirects, preDownloadScript, downloadScript,
                     postDownloadScript, installer presence, Authenticode
                     signature, SHA256, MSI product code
  3. PreDetect     - detection rule must report Not Detected before installing
  4. Install       - installScript / powerShellInstallScript through cmd.exe
  5. PostDetect    - detection rule must report Detected
  6. Reinstall     - optional idempotency pass (-IncludeReinstall)
  7. Uninstall     - uninstallScript / powerShellUninstallScript
  8. PostUninstall - detection rule must report Not Detected again
  9. Residue       - registry and directory leftovers after uninstall

Everything is written to artifacts\_install-tests\<timestamp>\ as JSON, a
markdown summary, and raw installer logs. Nothing is uploaded anywhere.

.EXAMPLE
    .\Invoke-RecipeInstallTest.ps1 -Id xmind

.EXAMPLE
    .\Invoke-RecipeInstallTest.ps1 -Id anydesk, discord -IncludeReinstall

.EXAMPLE
    .\Invoke-RecipeInstallTest.ps1 -All -WhatIf
#>
[CmdletBinding(DefaultParameterSetName = 'ById', SupportsShouldProcess)]
param(
    [Parameter(Mandatory, ParameterSetName = 'ById', Position = 0)]
    [ValidatePattern('^[A-Za-z0-9._-]+$')]
    [string[]] $Id,

    [Parameter(Mandatory, ParameterSetName = 'ByPath')]
    [string[]] $RecipePath,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch] $All,

    # Internal: newline-delimited recipe paths. Used by the elevated relaunch,
    # because -File cannot pass a string array reliably.
    [Parameter(Mandatory, ParameterSetName = 'FromFile')]
    [string] $RecipeListFile,

    # Install the app a second time before uninstalling, to prove the recipe is
    # idempotent and that detection still passes after an in-place reinstall.
    [switch] $IncludeReinstall,

    # Leave the application installed. Use only when debugging a single recipe.
    [switch] $SkipUninstall,

    # Delete each recipe's downloaded installer after the recipe finishes. Logs,
    # hashes, and results are kept. Worth using with -All, where retaining every
    # installer costs many gigabytes.
    [switch] $DiscardInstaller,

    # Continue even when the app is already detected before installation.
    # Without this the harness refuses to touch software it did not install.
    [switch] $Force,

    [ValidateRange(1, 240)]
    [int] $InstallTimeoutMinutes = 30,

    # How long to keep re-checking detection after install and uninstall before
    # calling it a failure. Silent installers frequently return early.
    [ValidateRange(0, 900)]
    [int] $DetectionSettleSeconds = 120,

    # How long to keep re-checking for uninstall leftovers before warning.
    [ValidateRange(0, 900)]
    [int] $ResidueSettleSeconds = 30,

    # Never relaunch elevated. System-context recipes are reported as skipped.
    [switch] $NoElevate,

    [string] $YardstickDevPath,

    [string] $RecipesPath,

    [string] $IconsPath,

    [string] $ArtifactsPath,

    # Internal: set on the elevated relaunch so the child does not loop.
    [switch] $ElevatedChild,

    # Internal: shared output directory between the parent and elevated child.
    [string] $RunRoot
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

# -WhatIf must only suppress the install and uninstall commands. Left as a
# preference variable it would also suppress the harness's own directory,
# variable, and report operations, which makes the run meaningless.
$requestedWhatIf = [bool]$WhatIfPreference
$WhatIfPreference = $false

$workspaceRecipes = (Resolve-Path -LiteralPath $RecipesPath -ErrorAction Stop).Path
$workspaceIcons = (Resolve-Path -LiteralPath $IconsPath -ErrorAction Stop).Path
$artifactsRoot = $ArtifactsPath
$schemaScript = Join-Path $PSScriptRoot 'Test-RecipeSchema.ps1'

Import-Module (Join-Path $PSScriptRoot 'RecipeTestSupport.psm1') -Force -ErrorAction Stop
Import-Module powershell-yaml -MinimumVersion 0.4.12 -ErrorAction Stop
Import-Module (Join-Path $YardstickDevPath 'Modules\YardstickSupport.psm1') -Force -ErrorAction Stop
Import-Module BitsTransfer -ErrorAction Stop

# Manual-download recipes use the same SoftwareDropbox setting as production
# Yardstick. The harness only reads and copies staged media; it never invokes
# Complete-YardstickDropbox or changes the source folder.
$yardstickPreferencesPath = Join-Path $YardstickDevPath 'Local\preferences.yaml'
if (-not (Test-Path -LiteralPath $yardstickPreferencesPath -PathType Leaf)) {
    $yardstickPreferencesPath = Join-Path $YardstickDevPath 'preferences.yaml'
}
$softwareDropbox = $null
if (Test-Path -LiteralPath $yardstickPreferencesPath -PathType Leaf) {
    $yardstickPreferences = Get-Content -LiteralPath $yardstickPreferencesPath -Raw | ConvertFrom-Yaml
    if ($yardstickPreferences.ContainsKey('SoftwareDropbox')) {
        $softwareDropbox = [string]$yardstickPreferences['SoftwareDropbox']
    }
}

#region recipe selection ------------------------------------------------------

$recipePaths = switch ($PSCmdlet.ParameterSetName) {
    'ById' { $Id | ForEach-Object { Join-Path $workspaceRecipes "$_.yaml" } }
    'ByPath' { $RecipePath }
    'All' { (Get-ChildItem -LiteralPath $workspaceRecipes -Filter '*.yaml' | Sort-Object Name).FullName }
    'FromFile' { Get-Content -LiteralPath $RecipeListFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } }
}
$recipePaths = @($recipePaths | ForEach-Object { (Resolve-Path -LiteralPath $_ -ErrorAction Stop).Path })
if ($recipePaths.Count -eq 0) { throw 'No recipes were selected.' }

function Read-Recipe {
    param([string] $Path)
    $recipe = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Yaml
    if ($recipe.ContainsKey('base')) {
        $baseId = [string]$recipe.base
        $baseWorkspacePath = Join-Path $workspaceRecipes "$baseId.yaml"
        $lookup = if (Test-Path -LiteralPath $baseWorkspacePath) { $workspaceRecipes } else { Join-Path $YardstickDevPath 'Recipes' }
        $recipe = Merge-RecipeWithBase -Recipe $recipe -RecipesPath $lookup
    }
    return $recipe
}

#endregion

#region elevation -------------------------------------------------------------

$isElevated = Test-IsElevated

# Recipes are partitioned by install context. A user-scope installer must not be
# run elevated: Spotify, for example, fails with exit 23 and logs "Invalid user
# account for install". So the unelevated session keeps the user-scope recipes
# and hands only the system-scope ones to an elevated child process.
$systemPaths = [Collections.Generic.List[string]]::new()
$userPaths = [Collections.Generic.List[string]]::new()
foreach ($path in $recipePaths) {
    $probe = Read-Recipe -Path $path
    $experience = if ($probe.ContainsKey('installExperience')) { [string]$probe['installExperience'] } else { 'system' }
    if ($experience -eq 'user') { $userPaths.Add($path) } else { $systemPaths.Add($path) }
}

if (-not $RunRoot) {
    $RunRoot = Join-Path $artifactsRoot ("_install-tests\{0}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
New-Item -ItemType Directory -Path $RunRoot -Force | Out-Null
$summaryJsonPath = Join-Path $RunRoot 'summary.json'
$summaryMarkdownPath = Join-Path $RunRoot 'summary.md'

# True when this process should hand the system-scope recipes to an elevated child.
# WhatIf never runs install/uninstall commands, so it can safely validate and
# download system-context recipes in the unelevated parent process.
$delegateSystemRecipes = ($systemPaths.Count -gt 0) -and (-not $isElevated) -and (-not $NoElevate) -and (-not $ElevatedChild) -and (-not $requestedWhatIf)

if ($systemPaths.Count -gt 0 -and -not $isElevated -and -not $requestedWhatIf -and ($NoElevate -or $ElevatedChild)) {
    Write-Warning 'Not elevated: system-context recipes will be reported as Skipped.'
}
if ($userPaths.Count -gt 0 -and $isElevated) {
    Write-Warning 'This session is elevated, so user-scope recipes will install elevated. Some user-scope installers reject that. Launch the harness unelevated to let it partition the run automatically.'
}

function Invoke-ElevatedBatch {
    <#
    .SYNOPSIS
    Runs the given recipes in an elevated child process and returns its results.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]] $Paths)

    $childRoot = Join-Path $RunRoot 'elevated'
    New-Item -ItemType Directory -Path $childRoot -Force | Out-Null

    $childArgs = [Collections.Generic.List[string]]::new()
    $childArgs.AddRange([string[]]@('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath))

    # The already-resolved recipe paths go through a file so that quoting and
    # array binding cannot be mangled by -File argument parsing.
    $listFile = Join-Path $childRoot 'recipes.txt'
    Set-Content -LiteralPath $listFile -Value $Paths -Encoding UTF8
    $childArgs.AddRange([string[]]@('-RecipeListFile', $listFile))
    if ($IncludeReinstall) { $childArgs.Add('-IncludeReinstall') }
    if ($SkipUninstall) { $childArgs.Add('-SkipUninstall') }
    if ($DiscardInstaller) { $childArgs.Add('-DiscardInstaller') }
    if ($Force) { $childArgs.Add('-Force') }
    if ($requestedWhatIf) { $childArgs.Add('-WhatIf') }
    $childArgs.AddRange([string[]]@('-InstallTimeoutMinutes', "$InstallTimeoutMinutes"))
    $childArgs.AddRange([string[]]@('-DetectionSettleSeconds', "$DetectionSettleSeconds"))
    $childArgs.AddRange([string[]]@('-ResidueSettleSeconds', "$ResidueSettleSeconds"))
    $childArgs.AddRange([string[]]@('-YardstickDevPath', $YardstickDevPath))
    $childArgs.AddRange([string[]]@('-RecipesPath', $workspaceRecipes))
    $childArgs.AddRange([string[]]@('-IconsPath', $workspaceIcons))
    $childArgs.AddRange([string[]]@('-ArtifactsPath', $artifactsRoot))
    $childArgs.AddRange([string[]]@('-RunRoot', $childRoot, '-ElevatedChild'))

    $psHost = [Environment]::ProcessPath
    if (-not $psHost -or -not (Test-Path -LiteralPath $psHost)) { $psHost = (Get-Process -Id $PID).Path }

    Write-Host ""
    Write-Host "Running $($Paths.Count) system-context recipe(s) elevated. Approve the UAC prompt once." -ForegroundColor Yellow
    Start-Process -FilePath $psHost -ArgumentList $childArgs -Verb RunAs -PassThru -Wait | Out-Null

    $childSummary = Join-Path $childRoot 'summary.json'
    if (-not (Test-Path -LiteralPath $childSummary)) {
        Write-Warning "The elevated run produced no summary at '$childSummary'."
        return @($Paths | ForEach-Object {
                [PSCustomObject]@{
                    Id      = [IO.Path]::GetFileNameWithoutExtension($_)
                    Version = ''
                    Overall = 'Fail'
                    Residue = $null
                    Phases  = @([PSCustomObject]@{
                            Name = 'Elevation'; Phase = 'Elevation'; Status = 'Fail'
                            Detail = 'The elevated child process did not produce results. UAC may have been declined.'
                            DurationSec = 0
                        })
                }
            })
    }
    $parsed = Get-Content -LiteralPath $childSummary -Raw | ConvertFrom-Json

    # The child writes its per-recipe evidence under <run>\elevated\<id>\. Move it
    # up to <run>\<id>\ so every recipe in the run has the same artifact shape
    # regardless of which context executed it; anything reading run output
    # directly then does not have to know that a merge happened. The child's own
    # summary and transcript stay under elevated\ as the raw record.
    foreach ($childRecipeDir in @(Get-ChildItem -LiteralPath $childRoot -Directory -ErrorAction SilentlyContinue)) {
        $target = Join-Path $RunRoot $childRecipeDir.Name
        if (Test-Path -LiteralPath $target) { continue }
        try {
            Move-Item -LiteralPath $childRecipeDir.FullName -Destination $target -ErrorAction Stop
        } catch {
            Write-Warning "Could not promote elevated evidence for '$($childRecipeDir.Name)': $_"
        }
    }

    Write-Host "Elevated results merged from $childSummary" -ForegroundColor DarkGray
    return @($parsed.Results)
}

$transcriptPath = Join-Path $RunRoot 'harness.transcript.log'
Start-Transcript -LiteralPath $transcriptPath -Force | Out-Null

#endregion

#region per-recipe test -------------------------------------------------------

$script:RecipeVariableNames = @()

function Clear-RecipeVariable {
    foreach ($name in $script:RecipeVariableNames) {
        Remove-Variable -Name $name -Scope Script -ErrorAction SilentlyContinue
    }
    $script:RecipeVariableNames = @()
}

function Wait-ForDetectionState {
    <#
    .SYNOPSIS
    Polls a recipe's detection rule until it reaches the expected state.

    .DESCRIPTION
    Many silent installers and NSIS-style uninstallers return control before the
    work on disk is finished, so a single detection check immediately after the
    command line exits produces false results. This polls until the expected
    state is reached or the settle window expires.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable] $DetectionArgs,
        [Parameter(Mandatory)][bool] $ExpectDetected,
        [Parameter(Mandatory)][string] $LogDirectory,
        [Parameter(Mandatory)][string] $LogName,
        [int] $TimeoutSeconds = 90,
        [int] $PollSeconds = 5
    )

    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    do {
        $check = Test-RecipeDetection @DetectionArgs -LogDirectory $LogDirectory -LogName $LogName
        if ($check.Detected -eq $ExpectDetected) { break }
        if ($stopwatch.Elapsed.TotalSeconds -ge $TimeoutSeconds) { break }
        Start-Sleep -Seconds $PollSeconds
    } while ($true)
    $stopwatch.Stop()

    $check | Add-Member -NotePropertyName SettleSeconds -NotePropertyValue ([math]::Round($stopwatch.Elapsed.TotalSeconds, 1)) -Force
    return $check
}

function New-PhaseResult {    param(
        [string] $Name,
        [ValidateSet('Pass', 'Fail', 'Skip', 'Warn')][string] $Status,
        [string] $Detail = '',
        [double] $DurationSec = 0
    )
    # Both Name and Phase carry the phase name. Phase is the original key and is
    # kept for compatibility; Name is what callers reaching for a conventional
    # identifier property expect, and its absence silently produced blank
    # columns in downstream reporting.
    [PSCustomObject]@{ Name = $Name; Phase = $Name; Status = $Status; Detail = $Detail; DurationSec = $DurationSec }
}

function Test-Recipe {
    [CmdletBinding()]
    param([string] $Path)

    $recipeId = [IO.Path]::GetFileNameWithoutExtension($Path)
    $recipeRoot = Join-Path $RunRoot $recipeId
    $logRoot = Join-Path $recipeRoot 'logs'
    $buildRoot = Join-Path $recipeRoot 'build'
    New-Item -ItemType Directory -Path $logRoot, $buildRoot -Force | Out-Null

    $phases = [Collections.Generic.List[psobject]]::new()
    $result = [ordered]@{
        Id                = $recipeId
        RecipePath        = $Path
        StartedAt         = (Get-Date).ToString('o')
        Version           = $null
        FileName          = $null
        Installer         = $null
        SHA256            = $null
        SignatureStatus   = $null
        Signer            = $null
        ProductCode       = $null
        InstallExperience = $null
        InstallCommand    = $null
        UninstallCommand  = $null
        Overall           = 'Fail'
        Phases            = $phases
        OutputDirectory   = $recipeRoot
        Residue           = $null
    }

    try {
        # -- Phase: Schema ---------------------------------------------------
        $sw = [Diagnostics.Stopwatch]::StartNew()
        try {
            & $schemaScript -RecipePath $Path -YardstickDevPath $YardstickDevPath -RecipesPath $workspaceRecipes -IconsPath $workspaceIcons | Out-Null
            $phases.Add((New-PhaseResult -Name 'Schema' -Status 'Pass' -DurationSec $sw.Elapsed.TotalSeconds))
        } catch {
            $phases.Add((New-PhaseResult -Name 'Schema' -Status 'Fail' -Detail "$_" -DurationSec $sw.Elapsed.TotalSeconds))
            return [PSCustomObject]$result
        }

        $recipe = Read-Recipe -Path $Path
        $installExperience = if ($recipe.ContainsKey('installExperience')) { [string]$recipe['installExperience'] } else { 'system' }
        $result['InstallExperience'] = $installExperience

        if ($installExperience -eq 'system' -and -not (Test-IsElevated) -and -not $requestedWhatIf) {
            $phases.Add((New-PhaseResult -Name 'Elevation' -Status 'Skip' -Detail 'System-context recipe requires an elevated harness run'))
            $result['Overall'] = 'Skipped'
            return [PSCustomObject]$result
        }

        # Recipe keys become variables so pre/post download scripts behave the
        # same way they do inside Yardstick and Invoke-RecipeLab.
        Clear-RecipeVariable
        foreach ($key in $recipe.Keys) {
            Set-Variable -Name $key -Value $recipe[$key] -Scope Script
            $script:RecipeVariableNames += $key
        }

        # Local copies of the three variables recipe scripts assign to. They must
        # be locals, because a dot-sourced pre/download script assigns into this
        # function's scope; reading the script-scope copies would miss those writes.
        $url = if ($recipe.ContainsKey('url')) { [string]$recipe['url'] } else { $null }
        $version = if ($recipe.ContainsKey('version')) { [string]$recipe['version'] } else { $null }
        $fileName = if ($recipe.ContainsKey('fileName')) { [string]$recipe['fileName'] } else { $null }

        # Resolve manual media before preDownloadScript, matching Yardstick's
        # production ordering so the recipe can inspect $dropboxPath/$dropboxFiles.
        $dropboxPath = $null
        $dropboxFiles = @()
        $isManualDownload = $recipe.ContainsKey('manualDownload') -and [bool]$recipe['manualDownload']
        if ($isManualDownload) {
            $manualFolder = if ($recipe.ContainsKey('manualDownloadFolder') -and $recipe['manualDownloadFolder']) {
                [string]$recipe['manualDownloadFolder']
            } else {
                $recipeId
            }
            $dropboxPayload = Get-YardstickDropboxPayload -DropboxRoot $softwareDropbox -Folder $manualFolder
            if (-not $dropboxPayload) {
                throw "Nothing is staged in the software dropbox for '$recipeId' (folder '$manualFolder')."
            }
            $dropboxPath = $dropboxPayload.Path
            $dropboxFiles = $dropboxPayload.Files
        }

        # Paths recipe scripts may reference.
        $script:BuildSpace = $buildRoot
        $script:Temp = Join-Path $recipeRoot 'temp'
        $script:Scripts = Join-Path $YardstickDevPath 'Scripts'
        $script:Tools = Join-Path $YardstickDevPath 'Tools'
        $script:Secrets = Join-Path $PSScriptRoot 'empty-secrets'
        $script:Recipes = $workspaceRecipes
        $script:Icons = $workspaceIcons
        New-Item -ItemType Directory -Path $script:Temp -Force | Out-Null

        # -- Phase: Download -------------------------------------------------
        # Recipe scripts are dot-sourced so their variable assignments land in
        # this scope, exactly as Yardstick's -NoNewScope invocation does. Their
        # streams are teed to a log rather than left in the pipeline: anything a
        # recipe writes would otherwise become part of Test-Recipe's own output
        # and corrupt the result object.
        $recipeScriptLog = Join-Path $logRoot 'recipe-scripts.log'
        $sw = [Diagnostics.Stopwatch]::StartNew()
        if ($recipe.ContainsKey('urlRedirects') -and $recipe['urlRedirects'] -eq $true -and $url) {
            $response = Invoke-WebRequest -Uri $url -Method Head -MaximumRedirection 10
            $url = $response.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
        }
        if ($recipe.ContainsKey('preDownloadScript') -and $recipe['preDownloadScript']) {
            Push-Location $YardstickDevPath
            try {
                . ([ScriptBlock]::Create([string]$recipe['preDownloadScript'])) *>&1 |
                    Tee-Object -FilePath $recipeScriptLog -Append | Out-Null
            } finally { Pop-Location }
        }
        if ([string]::IsNullOrWhiteSpace([string]$version)) {
            throw 'Recipe did not set $version in YAML or preDownloadScript.'
        }
        $result['Version'] = [string]$version

        $harnessBuildDirectory = Join-Path (Join-Path $buildRoot $recipeId) ([string]$version)
        New-Item -ItemType Directory -Path $harnessBuildDirectory -Force | Out-Null
        Push-Location $harnessBuildDirectory
        try {
            if ($isManualDownload) {
                Copy-Item -Path (Join-Path $dropboxPath '*') -Destination $harnessBuildDirectory -Recurse -Force
            } elseif ($recipe.ContainsKey('downloadScript') -and $recipe['downloadScript']) {
                . ([ScriptBlock]::Create([string]$recipe['downloadScript'])) *>&1 |
                    Tee-Object -FilePath $recipeScriptLog -Append | Out-Null
            } else {
                if ([string]::IsNullOrWhiteSpace([string]$url)) { throw 'Recipe did not set $url and has no downloadScript.' }
                if ([string]::IsNullOrWhiteSpace([string]$fileName)) {
                    $fileName = [IO.Path]::GetFileName(([Uri]$url).AbsolutePath)
                }
                if ([string]::IsNullOrWhiteSpace([string]$fileName)) {
                    throw 'Recipe did not set $fileName and it could not be derived from the URL.'
                }
                Start-BitsTransfer -Source $url -Destination (Join-Path $harnessBuildDirectory $fileName)
            }
            if ($recipe.ContainsKey('postDownloadScript') -and $recipe['postDownloadScript']) {
                . ([ScriptBlock]::Create([string]$recipe['postDownloadScript'])) *>&1 |
                    Tee-Object -FilePath $recipeScriptLog -Append | Out-Null
            }
        } finally {
            Pop-Location
        }

        $installerPath = Join-Path $harnessBuildDirectory ([string]$fileName)
        if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
            throw "Expected installer was not created at '$installerPath'."
        }
        $signature = Get-AuthenticodeSignature -LiteralPath $installerPath
        $result['FileName'] = [string]$fileName
        $result['Installer'] = $installerPath
        $result['SHA256'] = (Get-FileHash -LiteralPath $installerPath -Algorithm SHA256).Hash
        $result['SignatureStatus'] = $signature.Status.ToString()
        $result['Signer'] = if ($signature.SignerCertificate) { $signature.SignerCertificate.Subject } else { $null }

        if ([IO.Path]::GetExtension($installerPath) -ieq '.msi') {
            $productCode = @(Get-MsiProductCode -FilePath $installerPath) |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Select-Object -Last 1
            $result['ProductCode'] = [string]$productCode
        }

        $signatureNote = "signature=$($result['SignatureStatus'])"
        $downloadStatus = if ($signature.Status -eq 'Valid') { 'Pass' } else { 'Warn' }
        $phases.Add((New-PhaseResult -Name 'Download' -Status $downloadStatus `
                    -Detail "$fileName v$version $signatureNote" -DurationSec $sw.Elapsed.TotalSeconds))

        # -- Resolve install and uninstall command lines ----------------------
        $tokenArgs = @{ FileName = [string]$fileName; Version = [string]$version; ProductCode = [string]$result['ProductCode'] }

        $installCommand = $null
        if ($recipe.ContainsKey('installScript') -and $recipe['installScript']) {
            $installCommand = Expand-RecipeToken -Value ([string]$recipe['installScript']) @tokenArgs
        } elseif ($recipe.ContainsKey('powerShellInstallScript') -and $recipe['powerShellInstallScript']) {
            $content = Expand-RecipeToken -Value ([string]$recipe['powerShellInstallScript']) @tokenArgs
            Set-Content -LiteralPath (Join-Path $harnessBuildDirectory 'install.ps1') -Value $content -Encoding UTF8
            $installCommand = 'powershell.exe -noprofile -executionpolicy bypass -file .\install.ps1'
        }
        if (-not $installCommand) { throw 'Recipe has neither installScript nor powerShellInstallScript.' }
        $result['InstallCommand'] = $installCommand

        $uninstallCommand = $null
        if ($recipe.ContainsKey('uninstallScript') -and $recipe['uninstallScript']) {
            $uninstallCommand = Expand-RecipeToken -Value ([string]$recipe['uninstallScript']) @tokenArgs
        } elseif ($recipe.ContainsKey('powerShellUninstallScript') -and $recipe['powerShellUninstallScript']) {
            $content = Expand-RecipeToken -Value ([string]$recipe['powerShellUninstallScript']) @tokenArgs
            Set-Content -LiteralPath (Join-Path $harnessBuildDirectory 'uninstall.ps1') -Value $content -Encoding UTF8
            $uninstallCommand = 'powershell.exe -noprofile -executionpolicy bypass -file .\uninstall.ps1'
        }
        $result['UninstallCommand'] = $uninstallCommand

        # Yardstick runs recipe scripts with -NoNewScope, so a recipe may assign detection
        # fields ($fileDetectionPath, $fileDetectionVersion, $registryDetectionValue, ...) from
        # preDownloadScript or postDownloadScript. Those scripts are dot-sourced above, so the
        # assignments land here as locals. Fold them back into the recipe before detection runs,
        # otherwise the harness tests the literal YAML value and reports a false failure.
        foreach ($detectionKey in @(
                'fileDetectionPath', 'fileDetectionName', 'fileDetectionVersion', 'fileDetectionValue',
                'registryDetectionKey', 'registryDetectionValueName', 'registryDetectionValue')) {
            $assigned = Get-Variable -Name $detectionKey -Scope 0 -ValueOnly -ErrorAction SilentlyContinue
            if (-not [string]::IsNullOrWhiteSpace([string]$assigned)) {
                $recipe[$detectionKey] = [string]$assigned
            }
        }

        $detectionArgs = @{
            Recipe      = $recipe
            Version     = [string]$version
            ProductCode = [string]$result['ProductCode']
            FileName    = [string]$fileName
        }

        # -- Phase: PreDetect -------------------------------------------------
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $pre = Test-RecipeDetection @detectionArgs -LogDirectory $logRoot -LogName 'detect-pre'
        if ($pre.Detected) {
            $phases.Add((New-PhaseResult -Name 'PreDetect' -Status $(if ($Force) { 'Warn' } else { 'Fail' }) `
                        -Detail "Already detected before install ($($pre.Rule): $($pre.Detail))" -DurationSec $sw.Elapsed.TotalSeconds))
            if (-not $Force) {
                $result['Overall'] = 'Blocked'
                return [PSCustomObject]$result
            }
        } else {
            $phases.Add((New-PhaseResult -Name 'PreDetect' -Status 'Pass' `
                        -Detail "Not detected ($($pre.Rule): $($pre.Detail))" -DurationSec $sw.Elapsed.TotalSeconds))
        }

        if ($requestedWhatIf) {
            $phases.Add((New-PhaseResult -Name 'Install' -Status 'Skip' -Detail "WhatIf: would run '$installCommand'"))
            $phases.Add((New-PhaseResult -Name 'Uninstall' -Status 'Skip' -Detail "WhatIf: would run '$uninstallCommand'"))
            $result['Overall'] = 'WhatIf'
            return [PSCustomObject]$result
        }

        $timeoutMinutes = $InstallTimeoutMinutes
        if ($recipe.ContainsKey('maximumInstallationTimeInMinutes') -and $recipe['maximumInstallationTimeInMinutes']) {
            $timeoutMinutes = [int]$recipe['maximumInstallationTimeInMinutes']
        }
        $timeoutSeconds = $timeoutMinutes * 60

        $snapshotBefore = Get-SystemSnapshot

        # -- Phase: Install ---------------------------------------------------
        $install = Invoke-CommandLine -CommandLine $installCommand -WorkingDirectory $harnessBuildDirectory `
            -TimeoutSeconds $timeoutSeconds -LogDirectory $logRoot -LogName 'install'
        $installOk = (-not $install.TimedOut) -and ($install.ExitCode -in 0, 1641, 3010)
        $phases.Add((New-PhaseResult -Name 'Install' -Status $(if ($installOk) { 'Pass' } else { 'Fail' }) `
                    -Detail "exit=$($install.ExitCode) timedOut=$($install.TimedOut) cmd='$installCommand'" -DurationSec $install.DurationSec))

        # -- Phase: PostDetect ------------------------------------------------
        $post = Wait-ForDetectionState -DetectionArgs $detectionArgs -ExpectDetected $true `
            -LogDirectory $logRoot -LogName 'detect-post-install' -TimeoutSeconds $DetectionSettleSeconds
        $phases.Add((New-PhaseResult -Name 'PostInstallDetect' -Status $(if ($post.Detected) { 'Pass' } else { 'Fail' }) `
                    -Detail "$($post.Rule): $($post.Detail)" -DurationSec $post.SettleSeconds))

        # Taken after detection settles so late registry writes are captured.
        $snapshotAfterInstall = Get-SystemSnapshot

        # -- Phase: Reinstall (optional) --------------------------------------
        if ($IncludeReinstall -and $installOk) {
            $reinstall = Invoke-CommandLine -CommandLine $installCommand -WorkingDirectory $harnessBuildDirectory `
                -TimeoutSeconds $timeoutSeconds -LogDirectory $logRoot -LogName 'reinstall'
            $reinstallOk = (-not $reinstall.TimedOut) -and ($reinstall.ExitCode -in 0, 1641, 3010)
            $reDetect = Wait-ForDetectionState -DetectionArgs $detectionArgs -ExpectDetected $true `
                -LogDirectory $logRoot -LogName 'detect-post-reinstall' -TimeoutSeconds $DetectionSettleSeconds
            $phases.Add((New-PhaseResult -Name 'Reinstall' -Status $(if ($reinstallOk -and $reDetect.Detected) { 'Pass' } else { 'Fail' }) `
                        -Detail "exit=$($reinstall.ExitCode) detected=$($reDetect.Detected)" -DurationSec $reinstall.DurationSec))
        }

        # -- Phase: Uninstall -------------------------------------------------
        if ($SkipUninstall) {
            $phases.Add((New-PhaseResult -Name 'Uninstall' -Status 'Skip' -Detail '-SkipUninstall was specified; the app is still installed'))
        } elseif (-not $uninstallCommand) {
            $phases.Add((New-PhaseResult -Name 'Uninstall' -Status 'Fail' -Detail 'Recipe has neither uninstallScript nor powerShellUninstallScript'))
        } else {
            $uninstall = Invoke-CommandLine -CommandLine $uninstallCommand -WorkingDirectory $harnessBuildDirectory `
                -TimeoutSeconds $timeoutSeconds -LogDirectory $logRoot -LogName 'uninstall'
            $uninstallOk = (-not $uninstall.TimedOut) -and ($uninstall.ExitCode -in 0, 1605, 1641, 3010)
            $phases.Add((New-PhaseResult -Name 'Uninstall' -Status $(if ($uninstallOk) { 'Pass' } else { 'Fail' }) `
                        -Detail "exit=$($uninstall.ExitCode) timedOut=$($uninstall.TimedOut) cmd='$uninstallCommand'" -DurationSec $uninstall.DurationSec))

            # -- Phase: PostUninstallDetect ------------------------------------
            $postUninstall = Wait-ForDetectionState -DetectionArgs $detectionArgs -ExpectDetected $false `
                -LogDirectory $logRoot -LogName 'detect-post-uninstall' -TimeoutSeconds $DetectionSettleSeconds
            $phases.Add((New-PhaseResult -Name 'PostUninstallDetect' -Status $(if ($postUninstall.Detected) { 'Fail' } else { 'Pass' }) `
                        -Detail "$($postUninstall.Rule): $($postUninstall.Detail)" -DurationSec $postUninstall.SettleSeconds))

            # -- Phase: Residue -------------------------------------------------
            # Re-checked over a short window: some uninstallers leave a
            # self-deleting stub behind that disappears seconds later.
            $residueStopwatch = [Diagnostics.Stopwatch]::StartNew()
            do {
                $snapshotAfterUninstall = Get-SystemSnapshot
                $finalDelta = Compare-SystemSnapshot -Before $snapshotBefore -After $snapshotAfterUninstall
                $leftoverCount = @($finalDelta.AddedPackages).Count + @($finalDelta.AddedDirectories).Count
                if ($leftoverCount -eq 0) { break }
                if ($residueStopwatch.Elapsed.TotalSeconds -ge $ResidueSettleSeconds) { break }
                Start-Sleep -Seconds 10
            } while ($true)
            $residueStopwatch.Stop()

            $installedDelta = Compare-SystemSnapshot -Before $snapshotBefore -After $snapshotAfterInstall
            $result['Residue'] = [PSCustomObject]@{
                PackagesAddedByInstall = @($installedDelta.AddedPackages | Select-Object DisplayName, DisplayVersion, Scope, InstallLocation)
                PackagesLeftBehind     = @($finalDelta.AddedPackages | Select-Object DisplayName, DisplayVersion, Scope, InstallLocation)
                DirectoriesLeftBehind  = @($finalDelta.AddedDirectories)
                PackagesRemovedOverall = @($finalDelta.RemovedPackages | Select-Object DisplayName, DisplayVersion, Scope)
            }
            $phases.Add((New-PhaseResult -Name 'Residue' -Status $(if ($leftoverCount -eq 0) { 'Pass' } else { 'Warn' }) `
                        -Detail "$(@($finalDelta.AddedPackages).Count) ARP entries and $(@($finalDelta.AddedDirectories).Count) directories remain" `
                        -DurationSec ([math]::Round($residueStopwatch.Elapsed.TotalSeconds, 1))))
        }
    } catch {
        $phases.Add((New-PhaseResult -Name 'Error' -Status 'Fail' -Detail "$_"))
    } finally {
        Clear-RecipeVariable
        if ($DiscardInstaller -and (Test-Path -LiteralPath $buildRoot)) {
            Remove-Item -LiteralPath $buildRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    $statuses = @($phases | ForEach-Object { $_.Status })
    $result['Overall'] = if ($statuses -contains 'Fail') { 'Fail' }
    elseif ($result['Overall'] -in 'Skipped', 'Blocked', 'WhatIf') { $result['Overall'] }
    elseif ($statuses -contains 'Warn') { 'Pass with warnings' }
    else { 'Pass' }
    $result['CompletedAt'] = (Get-Date).ToString('o')

    $recipeResult = [PSCustomObject]$result
    $recipeResult | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $recipeRoot 'result.json') -Encoding UTF8
    return $recipeResult
}

#endregion

#region run -------------------------------------------------------------------

# When system-context recipes are delegated, this process only runs the
# user-scope ones; otherwise it runs everything it was given.
# The outer @() matters: a statement that yields a one-element array is unrolled
# on assignment, which would leave $localPaths as a bare string.
$localPaths = @(if ($delegateSystemRecipes) { $userPaths } else { $recipePaths })

$results = [Collections.Generic.List[psobject]]::new()
$index = 0
foreach ($path in $localPaths) {
    $index++
    $recipeId = [IO.Path]::GetFileNameWithoutExtension($path)
    Write-Host ""
    Write-Host "[$index/$($localPaths.Count)] $recipeId" -ForegroundColor Cyan
    $rawOutput = @(Test-Recipe -Path $path)
    $recipeResult = $rawOutput |
        Where-Object { $null -ne $_ -and $_.PSObject.Properties.Match('Phases').Count -gt 0 } |
        Select-Object -Last 1
    $strayOutput = @($rawOutput | Where-Object { $null -ne $_ -and $_.PSObject.Properties.Match('Phases').Count -eq 0 })

    if (-not $recipeResult) {
        # A recipe script that writes to the success stream pollutes Test-Recipe's output.
        # Never let that abort the whole run.
        $recipeResult = [PSCustomObject]@{
            Id      = $recipeId
            Version = ''
            Overall = 'Fail'
            Residue = $null
            Phases  = @([PSCustomObject]@{
                    Name        = 'Harness'
                    Phase       = 'Harness'
                    Status      = 'Fail'
                    Detail      = 'Test-Recipe returned no result object; the recipe likely emitted output to the success stream.'
                    DurationSec = 0
                })
        }
    } elseif ($strayOutput.Count -gt 0) {
        # Record it instead of discarding silently, so the leak stays diagnosable.
        $strayLog = Join-Path (Join-Path $RunRoot $recipeId) 'stray-output.log'
        $strayOutput | Out-String -Width 200 | Set-Content -LiteralPath $strayLog -Encoding UTF8 -ErrorAction SilentlyContinue
        $types = @($strayOutput | ForEach-Object { $_.GetType().Name } | Select-Object -Unique) -join ', '
        $recipeResult.Phases.Add([PSCustomObject]@{
                Name        = 'HarnessOutput'
                Phase       = 'HarnessOutput'
                Status      = 'Warn'
                Detail      = "$($strayOutput.Count) stray object(s) of type $types leaked into the harness pipeline; see stray-output.log"
                DurationSec = 0
            })
    }
    $recipeResult | Add-Member -NotePropertyName Context `
        -NotePropertyValue $(if ($isElevated) { 'elevated' } else { 'user' }) -Force
    $results.Add($recipeResult)
    foreach ($phase in $recipeResult.Phases) {
        $color = switch ($phase.Status) {
            'Pass' { 'Green' }
            'Warn' { 'Yellow' }
            'Skip' { 'DarkGray' }
            default { 'Red' }
        }
        Write-Host ("    {0,-20} {1,-5} {2}" -f $phase.Phase, $phase.Status, $phase.Detail) -ForegroundColor $color
    }
    Write-Host ("    => {0}" -f $recipeResult.Overall) -ForegroundColor $(if ($recipeResult.Overall -like 'Pass*') { 'Green' } else { 'Red' })
}

# The elevated batch runs after the user-scope ones so that a declined UAC prompt
# does not throw away results already gathered.
if ($delegateSystemRecipes) {
    foreach ($elevatedResult in (Invoke-ElevatedBatch -Paths $systemPaths)) {
        $elevatedResult | Add-Member -NotePropertyName Context -NotePropertyValue 'elevated' -Force
        $results.Add($elevatedResult)
    }
}

# Restore the caller's requested order regardless of which context ran each recipe.
$order = @{}
for ($i = 0; $i -lt $recipePaths.Count; $i++) { $order[[IO.Path]::GetFileNameWithoutExtension($recipePaths[$i])] = $i }
$results = [Collections.Generic.List[psobject]]@($results | Sort-Object { if ($order.ContainsKey($_.Id)) { $order[$_.Id] } else { [int]::MaxValue } })

$summary = [PSCustomObject]@{
    RunRoot     = $RunRoot
    StartedAt   = (Get-Date).ToString('o')
    Elevated    = (Test-IsElevated)
    Machine     = $env:COMPUTERNAME
    Results     = $results
}
$summary | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $summaryJsonPath -Encoding UTF8

$markdown = [Collections.Generic.List[string]]::new()
$markdown.Add("# Recipe install test run")
$markdown.Add("")
$markdown.Add("- Machine: $env:COMPUTERNAME")
$markdown.Add("- Elevated: $(Test-IsElevated)")
$markdown.Add("- Results: $RunRoot")
$markdown.Add("")
$markdown.Add("| Recipe | Version | Context | Result | Failed or warned phases |")
$markdown.Add("| --- | --- | --- | --- | --- |")
foreach ($item in $results) {
    $problems = @($item.Phases | Where-Object { $_.Status -in 'Fail', 'Warn' } | ForEach-Object { "$($_.Phase)=$($_.Status)" }) -join ', '
    $context = if ($item.PSObject.Properties.Match('Context').Count -gt 0) { $item.Context } else { '' }
    $markdown.Add("| $($item.Id) | $($item.Version) | $context | $($item.Overall) | $problems |")
}

$withResidue = @($results | Where-Object {
        $_.Residue -and (@($_.Residue.PackagesLeftBehind).Count + @($_.Residue.DirectoriesLeftBehind).Count) -gt 0
    })
if ($withResidue.Count -gt 0) {
    $markdown.Add("")
    $markdown.Add("## Uninstall leftovers")
    foreach ($item in $withResidue) {
        $markdown.Add("")
        $markdown.Add("### $($item.Id)")
        foreach ($package in $item.Residue.PackagesLeftBehind) {
            $markdown.Add("- Add/Remove Programs entry remains: $($package.DisplayName) $($package.DisplayVersion) ($($package.Scope))")
        }
        foreach ($directory in $item.Residue.DirectoriesLeftBehind) {
            $markdown.Add("- Path remains: $directory")
        }
    }
}
Set-Content -LiteralPath $summaryMarkdownPath -Value $markdown -Encoding UTF8

Stop-Transcript | Out-Null

Write-Host ""
Write-Host "Summary written to $summaryMarkdownPath" -ForegroundColor Cyan
$results | Select-Object Id, Version, Context, Overall | Format-Table -AutoSize | Out-Host

if (@($results | Where-Object { $_.Overall -eq 'Fail' }).Count -gt 0) { exit 1 }
exit 0

#endregion
