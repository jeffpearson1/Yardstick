<#
.SYNOPSIS
Reads and checks recipe evidence files: the proof that a recipe passed a live
install test with exactly its current contents.

.DESCRIPTION
Yardstick Development Studio writes Tests/RecipeEvidence/<id>.json next to each
recipe it publishes. The "Recipe evidence" pull request check
(Tests/RecipeHarness/Test-RecipeEvidence.ps1) uses these functions to decide
whether a changed recipe may be merged, and the recipe sync can use them to
decide whether to take a recipe. See Docs/RecipeEvidence.md.

The evidence is self-reported by the Studio. Checking it proves a recipe carries
evidence for exactly these contents; it is a process gate, not a security
boundary.
#>

# The eight harness phases (Invoke-RecipeInstallTest.ps1) a validated recipe must
# have passed. Older Studio builds wrote the two detection phases under shorter
# names, so those are accepted as aliases.
$Script:RequiredEvidencePhases = @(
    'Schema', 'Download', 'PreDetect', 'Install',
    'PostInstallDetect', 'Uninstall', 'PostUninstallDetect', 'Residue'
)
$Script:EvidencePhaseAliases = @{
    PostDetect    = 'PostInstallDetect'
    PostUninstall = 'PostUninstallDetect'
}
$Script:ValidatedVerdicts = @('Validated', 'ValidatedWithWarnings')
$Script:EvidenceRetestHint = 're-run the live test in Yardstick Development Studio and publish again'


function Get-RecipeContentHash {
    <#
    .SYNOPSIS
    Returns the recipe hash recorded in an evidence file's recipeSha256.

    .DESCRIPTION
    SHA-256, as lowercase hex, of the recipe's bytes after dropping a leading
    UTF-8 byte order mark and replacing every CRLF with LF. A lone CR is kept.
    This makes the hash ignore git's line-ending conversion, so a checkout with
    core.autocrlf gets the same value. The Studio computes it the same way.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [byte[]]$Bytes
    )

    $start = 0
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        $start = 3
    }

    $normalized = [System.IO.MemoryStream]::new($Bytes.Length)
    for ($i = $start; $i -lt $Bytes.Length; $i++) {
        if ($Bytes[$i] -eq 0x0D -and $i + 1 -lt $Bytes.Length -and $Bytes[$i + 1] -eq 0x0A) {
            continue
        }
        $normalized.WriteByte($Bytes[$i])
    }

    $hash = [System.Security.Cryptography.SHA256]::HashData($normalized.ToArray())
    [Convert]::ToHexString($hash).ToLowerInvariant()
}


function Get-RecipeEvidenceId {
    <#
    .SYNOPSIS
    Returns the recipe id for a recipe path: its file name without the extension.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string]$RecipePath
    )

    [System.IO.Path]::GetFileNameWithoutExtension(($RecipePath -replace '\\', '/').Split('/')[-1])
}


function Get-RecipeEvidencePath {
    <#
    .SYNOPSIS
    Returns the repository-relative evidence file path for a recipe path.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string]$RecipePath
    )

    "Tests/RecipeEvidence/$(Get-RecipeEvidenceId -RecipePath $RecipePath).json"
}


function Test-RecipeEvidenceNeeded {
    <#
    .SYNOPSIS
    Returns true when a repository-relative path is a recipe that needs evidence.

    .DESCRIPTION
    Every .yaml or .yml file under Recipes/, except development recipes
    (Recipes/Development/), disabled recipes (any Disabled folder, which the
    runtime skips too) and the retired *.common.yaml copies.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $segments = @(($Path -replace '\\', '/').Split('/'))
    if ($segments.Count -lt 2 -or $segments[0] -cne 'Recipes') { return $false }
    $fileName = $segments[-1]
    if ($fileName -notmatch '\.ya?ml$') { return $false }
    if ($fileName -like '*.common.yaml') { return $false }
    if ($segments.Count -gt 2 -and $segments[1] -eq 'Development') { return $false }
    if ($segments.Count -gt 2 -and $segments[1..($segments.Count - 2)] -contains 'Disabled') { return $false }
    $true
}


function Test-RecipeEvidenceContent {
    <#
    .SYNOPSIS
    Checks an evidence file against the recipe it should prove.

    .DESCRIPTION
    Returns Failures (one message per broken rule, each saying how to fix it), and
    the Verdict and TestedAt the file reports, for display. An empty Failures list
    means the evidence proves this exact recipe passed a live test.

    .PARAMETER Json
    The evidence file's text.

    .PARAMETER RecipePath
    The recipe's repository-relative path, with forward slashes.

    .PARAMETER RecipeSha256
    Get-RecipeContentHash of the recipe as it will be merged or used.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Json,

        [Parameter(Mandatory)]
        [string]$RecipePath,

        [Parameter(Mandatory)]
        [string]$RecipeSha256
    )

    $failures = [System.Collections.Generic.List[string]]::new()
    $result = [PSCustomObject]@{
        Failures = $failures
        Verdict  = ''
        TestedAt = ''
        Evidence = $null
    }

    try {
        $evidence = $Json | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    } catch {
        $failures.Add("Evidence file is not valid JSON ($($_.Exception.Message)); $Script:EvidenceRetestHint.")
        return $result
    }
    if ($evidence -isnot [System.Collections.IDictionary]) {
        $failures.Add("Evidence file is not a JSON object; $Script:EvidenceRetestHint.")
        return $result
    }
    $result.Evidence = $evidence
    $result.Verdict = [string]$evidence['verdict']
    $testedAt = $evidence['testedAt']
    # ConvertFrom-Json turns ISO 8601 text into a DateTime; show it as written.
    $result.TestedAt = if ($testedAt -is [datetime]) {
        $testedAt.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
    } else {
        [string]$testedAt
    }

    $schemaVersion = $evidence['schemaVersion']
    if (($schemaVersion -isnot [long] -and $schemaVersion -isnot [int]) -or $schemaVersion -ne 1) {
        $failures.Add("schemaVersion is '$schemaVersion', but this check reads schemaVersion 1; publish again with a Studio that writes schemaVersion 1.")
        return $result
    }

    $recipeId = Get-RecipeEvidenceId -RecipePath $RecipePath
    if ([string]$evidence['recipeId'] -cne $recipeId) {
        $failures.Add("recipeId is '$($evidence['recipeId'])', but the recipe file is '$recipeId'; publish the recipe again from the Studio so the evidence names it.")
    }
    if ([string]$evidence['recipePath'] -cne $RecipePath) {
        $failures.Add("recipePath is '$($evidence['recipePath'])', but the recipe is at '$RecipePath'; publish the recipe again from the Studio at its new path.")
    }

    $recordedHash = [string]$evidence['recipeSha256']
    if ($recordedHash -notmatch '^[0-9a-fA-F]{64}$') {
        $failures.Add("recipeSha256 '$recordedHash' is not a SHA-256 hash; $Script:EvidenceRetestHint.")
    } elseif ($recordedHash -ne $RecipeSha256) {
        $failures.Add("Hash mismatch: the recipe changed since it was tested (evidence $($recordedHash.ToLowerInvariant()), recipe $RecipeSha256); $Script:EvidenceRetestHint.")
    }

    $validated = $evidence['validated']
    if ($validated -isnot [bool] -or -not $validated) {
        $reason = [string]$evidence['reason']
        $because = if ($reason) { ": $reason" } else { '' }
        $failures.Add("validated is not true; the recipe was published without a passing live test$because. Get the live test to pass, then $Script:EvidenceRetestHint.")
    }
    if ($Script:ValidatedVerdicts -cnotcontains $result.Verdict) {
        $failures.Add("verdict is '$($result.Verdict)', not Validated or ValidatedWithWarnings; get the live test to pass, then $Script:EvidenceRetestHint.")
    }

    $statuses = @{}
    foreach ($phase in @($evidence['phases'])) {
        if ($phase -isnot [System.Collections.IDictionary]) { continue }
        # The harness's summary.json writes Name and Status; the Studio may not
        # change their case when it copies them.
        $phaseValues = @{}
        foreach ($key in $phase.Keys) { $phaseValues[[string]$key] = $phase[$key] }
        $name = [string]$phaseValues['name']
        foreach ($alias in $Script:EvidencePhaseAliases.Keys) {
            if ($name -eq $alias) { $name = $Script:EvidencePhaseAliases[$alias] }
        }
        $required = $Script:RequiredEvidencePhases | Where-Object { $_ -eq $name }
        if (-not $required) { continue }
        if (-not $statuses.ContainsKey($required)) { $statuses[$required] = [System.Collections.Generic.List[string]]::new() }
        $statuses[$required].Add([string]$phaseValues['status'])
    }
    $missing = @($Script:RequiredEvidencePhases | Where-Object { -not $statuses.ContainsKey($_) })
    if ($missing) {
        $failures.Add("Phase(s) missing from the evidence: $($missing -join ', '); $Script:EvidenceRetestHint.")
    }
    foreach ($name in $Script:RequiredEvidencePhases) {
        if (-not $statuses.ContainsKey($name)) { continue }
        # Count, not truth: a missing or empty status is a single '', which is false.
        $bad = @($statuses[$name] | Where-Object { $_ -notin @('Pass', 'Warn') })
        if ($bad.Count -gt 0) {
            $shown = if ($bad[0]) { "'$($bad[0])'" } else { 'missing' }
            $failures.Add("Phase $name status is $shown, not Pass or Warn; fix the recipe so the live test passes, then $Script:EvidenceRetestHint.")
        }
    }

    $result
}


Export-ModuleMember -Function @(
    'Get-RecipeContentHash',
    'Get-RecipeEvidenceId',
    'Get-RecipeEvidencePath',
    'Test-RecipeEvidenceNeeded',
    'Test-RecipeEvidenceContent'
)
