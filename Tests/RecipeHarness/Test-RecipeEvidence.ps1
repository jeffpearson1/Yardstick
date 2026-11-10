<#
.SYNOPSIS
Checks that every recipe a branch adds or changes carries passing Studio evidence.

.DESCRIPTION
This is the "Recipe evidence" pull request check (.github/workflows/recipe-evidence.yml),
and it runs the same way locally. For every recipe under Recipes/ that the branch
adds, modifies or renames into place compared with the merge base, it requires
Tests/RecipeEvidence/<id>.json to prove that the recipe passed a live install
test in Yardstick Development Studio with exactly the contents being merged, and
it runs Test-RecipeSchema.ps1 on the recipe.

Recipes and evidence are read from the head commit, not from the working tree, so
line-ending conversion and clean filters cannot change the result. Commit your
changes before running it locally.

Evidence failures become warnings when -Waived is set (the evidence-waived pull
request label). Schema failures always fail. See Docs/RecipeEvidence.md.

.PARAMETER BaseRef
The branch being merged into. Recipes are compared with its merge base with
HeadRef, so the history of both must be present (actions/checkout fetch-depth: 0).

.PARAMETER HeadRef
The commit being checked. Defaults to HEAD.

.PARAMETER Waived
Report evidence failures as warnings. Set by the evidence-waived label.

.PARAMETER RepositoryPath
The git repository to check. Defaults to the Yardstick project this script is in.

.PARAMETER SummaryPath
A file to append the Markdown results table to. Defaults to $env:GITHUB_STEP_SUMMARY.

.PARAMETER PassThru
Return the results as an object as well.

.EXAMPLE
.\Tests\RecipeHarness\Test-RecipeEvidence.ps1 -BaseRef origin/main
#>
[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string] $BaseRef = 'origin/main',

    [ValidateNotNullOrEmpty()]
    [string] $HeadRef = 'HEAD',

    [switch] $Waived,

    [string] $RepositoryPath,

    [string] $SummaryPath = $env:GITHUB_STEP_SUMMARY,

    [switch] $PassThru
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..') -ErrorAction Stop).ProviderPath
if (-not $RepositoryPath) { $RepositoryPath = $projectRoot }
$RepositoryPath = (Resolve-Path -LiteralPath $RepositoryPath -ErrorAction Stop).ProviderPath
Import-Module (Join-Path $projectRoot 'Modules\YardstickRecipeEvidence.psm1') -Force -ErrorAction Stop
$inActions = $env:GITHUB_ACTIONS -eq 'true'


function Invoke-Git {
    # Runs git in the repository and returns stdout as raw bytes. PowerShell's
    # native-command pipeline would decode it as text and change line endings,
    # which the recipe hash must see untouched.
    param(
        [Parameter(Mandatory)] [string[]] $Arguments,
        [switch] $AllowFailure
    )

    $startInfo = [System.Diagnostics.ProcessStartInfo]::new('git')
    foreach ($argument in @('-C', $RepositoryPath, '-c', 'core.quotepath=off') + $Arguments) {
        $startInfo.ArgumentList.Add($argument)
    }
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.UseShellExecute = $false

    $process = [System.Diagnostics.Process]::Start($startInfo)
    $stderr = $process.StandardError.ReadToEndAsync()
    $stdout = [System.IO.MemoryStream]::new()
    $process.StandardOutput.BaseStream.CopyTo($stdout)
    $process.WaitForExit()

    $result = [PSCustomObject]@{
        ExitCode = $process.ExitCode
        Bytes    = $stdout.ToArray()
        Error    = $stderr.Result.Trim()
    }
    if ($result.ExitCode -ne 0 -and -not $AllowFailure) {
        throw "git $($Arguments -join ' ') failed ($($result.ExitCode)): $($result.Error)"
    }
    $result
}


function Get-GitPathList {
    # Splits NUL-separated git output (-z) into paths.
    param([Parameter(Mandatory)] [string[]] $Arguments)

    $text = [System.Text.Encoding]::UTF8.GetString((Invoke-Git -Arguments $Arguments).Bytes)
    @($text.Split([char]0) | Where-Object { $_ })
}


function ConvertTo-AnnotationText {
    param([string] $Text)
    $Text.Replace('%', '%25').Replace("`r", '%0D').Replace("`n", '%0A')
}


function Write-CheckLine {
    param(
        [Parameter(Mandatory)] [ValidateSet('Info', 'Warning', 'Error')] [string] $Level,
        [Parameter(Mandatory)] [string] $Message,
        [string] $File
    )

    Write-Host $Message
    if ($inActions -and $Level -ne 'Info') {
        $command = if ($Level -eq 'Error') { 'error' } else { 'warning' }
        $location = if ($File) { " file=$File" } else { '' }
        Write-Host "::$command${location}::$(ConvertTo-AnnotationText $Message)"
    }
}


function Invoke-RecipeSchemaCheck {
    # Runs Test-RecipeSchema.ps1 on the recipe as committed at HeadRef. It is
    # written to a temporary <id>.yaml, since the schema script requires the id to
    # match the file name; base recipes are found in the repository's Recipes.
    param(
        [Parameter(Mandatory)] [string] $RecipePath,
        [Parameter(Mandatory)] [byte[]] $Bytes
    )

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) "RecipeEvidence-$([guid]::NewGuid().ToString('N'))"
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    try {
        $tempRecipe = Join-Path $tempRoot ([System.IO.Path]::GetFileName($RecipePath))
        [System.IO.File]::WriteAllBytes($tempRecipe, $Bytes)

        $schemaParams = @{
            RecipePath       = $tempRecipe
            RecipesPath      = Join-Path $RepositoryPath 'Recipes'
            YardstickDevPath = $projectRoot
            IconsPath        = Join-Path $RepositoryPath 'Icons'
            NoSync           = $true
            SkipIconCheck    = $true
        }
        $errors = [System.Collections.Generic.List[string]]::new()
        try {
            # Streamed, so the error lines written before the script throws are kept.
            & (Join-Path $PSScriptRoot 'Test-RecipeSchema.ps1') @schemaParams 2>&1 | ForEach-Object {
                if ($_ -is [System.Management.Automation.ErrorRecord]) { $errors.Add($_.ToString()) }
            }
        } catch {
            if ($errors.Count -eq 0) { $errors.Add($_.Exception.Message) }
        }
        $errors.ToArray()
    } finally {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}


# -- Which recipes changed ----------------------------------------------------

$mergeBase = Invoke-Git -Arguments @('merge-base', $BaseRef, $HeadRef) -AllowFailure
if ($mergeBase.ExitCode -ne 0) {
    throw "No merge base between '$BaseRef' and '$HeadRef'. Fetch the base branch with full history (actions/checkout fetch-depth: 0, or git fetch origin). $($mergeBase.Error)"
}

$range = "$BaseRef...$HeadRef"
$changed = @(Get-GitPathList -Arguments @('diff', '--name-only', '-z', '-M', '--diff-filter=AMR', $range, '--', 'Recipes/') |
    Where-Object { Test-RecipeEvidenceNeeded -Path $_ } | Sort-Object -Unique)
$deleted = @(Get-GitPathList -Arguments @('diff', '--name-only', '-z', '-M', '--diff-filter=D', $range, '--', 'Recipes/') |
    Where-Object { Test-RecipeEvidenceNeeded -Path $_ })

$results = [System.Collections.Generic.List[object]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()

if ($changed.Count -eq 0) {
    Write-CheckLine -Level Info -Message "No recipe changes compared with $BaseRef; nothing to check."
} else {
    Write-CheckLine -Level Info -Message "Checking $($changed.Count) changed recipe(s) against $BaseRef$(if ($Waived) { ' (evidence-waived: evidence failures are warnings)' })."
}

# -- Evidence and schema, per recipe -------------------------------------------

foreach ($recipePath in $changed) {
    $recipeId = Get-RecipeEvidenceId -RecipePath $recipePath
    $evidencePath = Get-RecipeEvidencePath -RecipePath $recipePath
    $recipeBytes = (Invoke-Git -Arguments @('cat-file', 'blob', "${HeadRef}:$recipePath")).Bytes
    $recipeHash = Get-RecipeContentHash -Bytes $recipeBytes

    $evidenceFailures = @()
    $verdict = ''
    $testedAt = ''
    $evidenceBlob = Invoke-Git -Arguments @('cat-file', 'blob', "${HeadRef}:$evidencePath") -AllowFailure
    if ($evidenceBlob.ExitCode -ne 0) {
        $evidenceFailures = @("No evidence file: $evidencePath is missing. Test the recipe in Yardstick Development Studio and publish it from there, which adds the file.")
    } else {
        $evidenceJson = [System.Text.Encoding]::UTF8.GetString($evidenceBlob.Bytes).TrimStart([char]0xFEFF)
        $check = Test-RecipeEvidenceContent -Json $evidenceJson -RecipePath $recipePath -RecipeSha256 $recipeHash
        $evidenceFailures = @($check.Failures)
        $verdict = $check.Verdict
        $testedAt = $check.TestedAt
    }

    $schemaFailures = @(Invoke-RecipeSchemaCheck -RecipePath $recipePath -Bytes $recipeBytes)

    $outcome = if ($schemaFailures) {
        'Fail'
    } elseif ($evidenceFailures -and $Waived) {
        'Waived'
    } elseif ($evidenceFailures) {
        'Fail'
    } else {
        'Pass'
    }
    $reasons = @($schemaFailures | ForEach-Object { "Schema: $_" }) + @($evidenceFailures)

    $results.Add([PSCustomObject]@{
        Id               = $recipeId
        RecipePath       = $recipePath
        EvidencePath     = $evidencePath
        Result           = $outcome
        Verdict          = $verdict
        TestedAt         = $testedAt
        RecipeSha256     = $recipeHash
        SchemaFailures   = $schemaFailures
        EvidenceFailures = $evidenceFailures
        Reasons          = $reasons
    })

    $about = if ($verdict) { ": $verdict, tested $testedAt" } else { '' }
    $level = switch ($outcome) { 'Pass' { 'Info' } 'Waived' { 'Warning' } default { 'Error' } }
    Write-CheckLine -Level Info -Message "[$($outcome.ToUpperInvariant())] $recipeId ($recipePath)$about"
    foreach ($reason in $reasons) {
        Write-CheckLine -Level $level -Message "  $recipeId - $reason" -File $recipePath
    }
}

# -- Stray evidence ------------------------------------------------------------

$evidenceFiles = @(Get-GitPathList -Arguments @('ls-tree', '-r', '--name-only', '-z', $HeadRef, '--', 'Tests/RecipeEvidence/') |
    Where-Object { $_ -like '*.json' })
if ($evidenceFiles) {
    $headIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($path in Get-GitPathList -Arguments @('ls-tree', '-r', '--name-only', '-z', $HeadRef, '--', 'Recipes/')) {
        if (Test-RecipeEvidenceNeeded -Path $path) { [void]$headIds.Add((Get-RecipeEvidenceId -RecipePath $path)) }
    }
    $deletedIds = @($deleted | ForEach-Object { Get-RecipeEvidenceId -RecipePath $_ })

    foreach ($evidencePath in $evidenceFiles) {
        $evidenceId = Get-RecipeEvidenceId -RecipePath $evidencePath
        if ($headIds.Contains($evidenceId)) { continue }
        $message = if ($deletedIds -ccontains $evidenceId) {
            "$evidencePath belongs to a recipe this branch deletes; remove it."
        } else {
            "$evidencePath matches no recipe under Recipes/; remove it or rename it to match its recipe."
        }
        $warnings.Add($message)
        Write-CheckLine -Level Warning -Message $message -File $evidencePath
    }
}

# -- Report --------------------------------------------------------------------

$failed = @($results | Where-Object Result -eq 'Fail')
$waivedCount = @($results | Where-Object Result -eq 'Waived').Count
$passed = $failed.Count -eq 0

if ($SummaryPath) {
    $summary = [System.Text.StringBuilder]::new()
    [void]$summary.AppendLine('## Recipe evidence')
    [void]$summary.AppendLine()
    if ($results.Count -eq 0) {
        [void]$summary.AppendLine("No recipe changes compared with ``$BaseRef``; nothing to check.")
    } else {
        if ($Waived) {
            [void]$summary.AppendLine('The `evidence-waived` label is set: evidence failures are reported as warnings. Schema failures still fail.')
            [void]$summary.AppendLine()
        }
        [void]$summary.AppendLine('| Recipe | Result | Verdict | Tested at | Reason |')
        [void]$summary.AppendLine('|---|---|---|---|---|')
        foreach ($result in $results) {
            $reason = (@($result.Reasons) -join '<br>').Replace('|', '\|')
            [void]$summary.AppendLine("| ``$($result.Id)`` | $($result.Result) | $($result.Verdict) | $($result.TestedAt) | $reason |")
        }
    }
    if ($warnings.Count -gt 0) {
        [void]$summary.AppendLine()
        [void]$summary.AppendLine('**Warnings**')
        [void]$summary.AppendLine()
        foreach ($warning in $warnings) { [void]$summary.AppendLine("- $warning") }
    }
    Add-Content -LiteralPath $SummaryPath -Value $summary.ToString() -Encoding utf8
}

if ($results.Count -gt 0) {
    $verdictLine = "$($results.Count) recipe(s) checked: $($results.Count - $failed.Count - $waivedCount) passed, $waivedCount waived, $($failed.Count) failed."
    Write-CheckLine -Level Info -Message $verdictLine
}

if ($PassThru) {
    [PSCustomObject]@{
        Passed   = $passed
        Recipes  = $results.ToArray()
        Warnings = $warnings.ToArray()
    }
}

if (-not $passed) { exit 1 }
exit 0
