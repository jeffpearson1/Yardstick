<#
.SYNOPSIS
Downloads recipes from GitHub repositories at the start of a run.

.DESCRIPTION
With syncRecipes on, every run first brings its recipes up to date from GitHub:
- recipeRepository, Yardstick's own recipe source, is synced into the Recipes
  folder itself;
- each of additionalRecipeRepositories is mirrored into its own folder under
  Local\RecipeRepos, which the run searches after Recipes, in the order listed.
  A recipe in Recipes wins over one with the same id in a linked repository.

Only recipes are synced: the .yaml and .yml files under the repository's Recipes
folder, apart from Development and *.common.yaml. Recipes in a Disabled folder are
synced too, where they stay: they never run, but an enabled recipe can use one as
its base. Every run lists the repository with one API call and compares the git
blob hash of each local file with it, so a recipe deleted or changed here is put
right; only files that differ are downloaded.

In Recipes, a file the sync wrote and that was edited here since is kept, with a
warning, and so is any recipe that is not in the repository. A file the sync
replaces or removes is copied to Artifacts\RecipeSyncBackups first.

With requireRecipeEvidence on, a new or changed recipe is only taken when its
Tests/RecipeEvidence/<id>.json in the same repository proves it passed a live
install test with exactly those contents (see Docs\RecipeEvidence.md). Otherwise
the copy already here stays.

Like the shared content sync, this never throws: a repository that cannot be
reached is reported and the run carries on with what was synced before.
#>

. (Join-Path (Split-Path -Parent $PSScriptRoot) 'ProjectLayout.ps1')
Import-Module (Join-Path $PSScriptRoot 'YardstickSupport.psm1')
Import-Module (Join-Path $PSScriptRoot 'YardstickGitHub.psm1')
Import-Module (Join-Path $PSScriptRoot 'YardstickRecipeEvidence.psm1')

# owner/name or owner/name@ref, as GitHub allows them.
$Script:RepositorySpecPattern = '^(?<owner>[A-Za-z0-9][A-Za-z0-9-]*)/(?<name>[A-Za-z0-9._-]+)(?:@(?<ref>[^\s@]+))?$'


function ConvertFrom-YardstickRecipeRepositorySpec {
    <#
    .SYNOPSIS
    Parses 'owner/name' or 'owner/name@ref' into its parts. Throws when it is not one.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Spec
    )

    $text = $Spec.Trim()
    if ($text -notmatch $Script:RepositorySpecPattern) {
        throw "'$Spec' is not a GitHub repository. Use owner/name, or owner/name@branch."
    }
    $owner = $Matches.owner
    $name = $Matches.name
    $ref = [string]$Matches['ref']
    # GitHub owner and repository names ignore case; refs do not. The folder name
    # is readable, and a hash of the exact identity keeps a-b/c apart from a/b-c,
    # and release/1.2 apart from release_1.2.
    $identity = "$($owner.ToLowerInvariant())/$($name.ToLowerInvariant())@$ref"
    $digest = [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData([System.Text.Encoding]::UTF8.GetBytes($identity))).Substring(0, 12).ToLowerInvariant()

    [PSCustomObject]@{
        Spec     = $text
        Owner    = $owner
        Name     = $name
        Ref      = $ref
        Identity = $identity
        Key      = "$($owner.ToLowerInvariant())-$($name.ToLowerInvariant())-$digest"
    }
}


function Get-YardstickRecipeRepository {
    <#
    .SYNOPSIS
    Lists the recipe repositories in preferences, the main one first.

    .DESCRIPTION
    Each has Role (Main or Linked), Destination (the folder its recipes go to) and
    StatePath (what the last sync recorded). A spec that cannot be parsed comes back
    with Error set instead.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Preferences,
        [Parameter(Mandatory)][string]$ProjectRoot
    )

    $cacheRoot = Get-YardstickRecipeRepositoryCachePath -ProjectRoot $ProjectRoot
    $entries = @()
    if (-not [string]::IsNullOrWhiteSpace([string]$Preferences.recipeRepository)) {
        $entries += , @('Main', [string]$Preferences.recipeRepository)
    }
    foreach ($spec in @($Preferences.additionalRecipeRepositories)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$spec)) { $entries += , @('Linked', [string]$spec) }
    }

    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($entry in $entries) {
        $role, $spec = $entry
        try {
            $parsed = ConvertFrom-YardstickRecipeRepositorySpec -Spec $spec
        } catch {
            [PSCustomObject]@{ Role = $role; Spec = $spec; Error = $_.Exception.Message }
            continue
        }
        if ($role -eq 'Linked' -and -not $seen.Add($parsed.Identity)) { continue }
        $destination = if ($role -eq 'Main') { [string]$Preferences.Recipes } else { Join-Path (Join-Path $cacheRoot $parsed.Key) 'Recipes' }
        [PSCustomObject]@{
            Role        = $role
            Spec        = $parsed.Spec
            Owner       = $parsed.Owner
            Name        = $parsed.Name
            Ref         = $parsed.Ref
            Key         = $parsed.Key
            Destination = $destination
            StatePath   = Join-Path $cacheRoot "$($role.ToLowerInvariant())-$($parsed.Key).state.json"
            Error       = $null
        }
    }
}


function Get-YardstickRecipeSearchPath {
    <#
    .SYNOPSIS
    Returns the folders a run looks for recipes in, in order of precedence.

    .DESCRIPTION
    The Recipes folder, then, when syncRecipes is on, the folder of each linked
    recipe repository that has been synced, in the order preferences list them.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)][hashtable]$Preferences,
        [Parameter(Mandatory)][string]$ProjectRoot
    )

    $paths = [System.Collections.Generic.List[string]]::new()
    $paths.Add([string]$Preferences.Recipes)
    if ($Preferences.syncRecipes -eq $true) {
        foreach ($repository in Get-YardstickRecipeRepository -Preferences $Preferences -ProjectRoot $ProjectRoot) {
            if ($repository.Role -ne 'Linked' -or $repository.Error) { continue }
            if (Test-Path -LiteralPath $repository.Destination -PathType Container) { $paths.Add($repository.Destination) }
        }
    }
    $paths.ToArray()
}


function Get-YardstickGitBlobSha {
    <#
    .SYNOPSIS
    Returns git's object id for a file's content: SHA-1 of "blob <length>\0" and
    the bytes, the same as git hash-object.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [byte[]]$Bytes
    )

    $hash = [System.Security.Cryptography.IncrementalHash]::CreateHash([System.Security.Cryptography.HashAlgorithmName]::SHA1)
    try {
        $hash.AppendData([System.Text.Encoding]::ASCII.GetBytes("blob $($Bytes.Length)`0"))
        $hash.AppendData($Bytes)
        [Convert]::ToHexString($hash.GetHashAndReset()).ToLowerInvariant()
    } finally {
        $hash.Dispose()
    }
}


function Read-RecipeSyncState {
    param([string]$Path, [string]$Spec)

    $empty = @{ spec = $Spec; ref = ''; treeSha = ''; syncedAt = ''; files = @{}; rejected = @{} }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $empty }
    try {
        $state = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable -ErrorAction Stop
    } catch {
        return $empty
    }
    # A different repository in the same preference starts again.
    if ($state -isnot [System.Collections.IDictionary] -or [string]$state['spec'] -ne $Spec) { return $empty }
    $files = @{}
    if ($state['files'] -is [System.Collections.IDictionary]) {
        foreach ($key in $state['files'].Keys) { $files[[string]$key] = [string]$state['files'][$key] }
    }
    # Recipes refused for want of evidence, by path: the blob and evidence they were
    # refused with, and why, so the same pair isn't downloaded again every run.
    $rejected = @{}
    if ($state['rejected'] -is [System.Collections.IDictionary]) {
        foreach ($key in $state['rejected'].Keys) {
            $entry = $state['rejected'][$key]
            if ($entry -is [System.Collections.IDictionary]) {
                $rejected[[string]$key] = @{ blobs = [string]$entry['blobs']; reason = [string]$entry['reason'] }
            }
        }
    }
    @{
        spec     = $Spec
        ref      = [string]$state['ref']
        treeSha  = [string]$state['treeSha']
        syncedAt = $state['syncedAt']
        files    = $files
        rejected = $rejected
    }
}


function Save-RecipeSyncState {
    param([string]$Path, [hashtable]$State)

    $json = $State | ConvertTo-Json -Depth 5
    Write-YardstickFileSafely -Path $Path -Content ([System.Text.Encoding]::UTF8.GetBytes($json))
}


function Get-RecipeSyncBlob {
    # Downloads one file through the API, so private repositories work with the
    # stored token, and checks it against the hash the listing gave.
    param([string]$ApiBase, [string]$Sha)

    $blob = Invoke-YardstickGitHubApi "$ApiBase/git/blobs/$Sha"
    if ([string]$blob.encoding -ne 'base64') { throw "GitHub returned blob $Sha as '$($blob.encoding)', not base64." }
    $bytes = [Convert]::FromBase64String(([string]$blob.content -replace '\s', ''))
    if ((Get-YardstickGitBlobSha -Bytes $bytes) -ne $Sha) { throw "Blob $Sha did not match its hash." }
    , $bytes
}


function Resolve-RecipeSyncPath {
    # Maps a repository path under Recipes/ to the local folder, refusing any path
    # that would land outside it. Git allows names Windows reads differently: a
    # backslash is a separator here, a colon starts a drive or stream, and a
    # trailing dot or space is dropped.
    param([string]$Root, [string]$RepositoryPath)

    $segments = @($RepositoryPath.Substring('Recipes/'.Length).Split('/'))
    # Counted, not truth-tested: an empty segment is '', which reads as false.
    $bad = @($segments | Where-Object {
        $_ -in '', '.', '..' -or $_ -match '[\:*?"<>|\x00-\x1f]' -or $_ -match '[. ]$'
    })
    if ($bad.Count -gt 0) { throw "Refusing the repository path '$RepositoryPath'." }

    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    $full = [System.IO.Path]::GetFullPath((Join-Path $Root ($segments -join [System.IO.Path]::DirectorySeparatorChar)))
    if (-not $full.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing the repository path '$RepositoryPath', which is outside $Root."
    }
    $full
}


function Test-RecipeSyncPath {
    # The repository files the sync takes: every recipe that needs evidence, and
    # disabled ones too. Those never run, but an enabled recipe can use one as its
    # base, so they are kept in their Disabled folder.
    param([string]$Path)

    if (Test-RecipeEvidenceNeeded -Path $Path) { return $true }
    $segments = @($Path.Split('/'))
    $segments.Count -gt 2 -and $segments[0] -ceq 'Recipes' -and $segments[1] -ne 'Development' -and
        $segments[-1] -match '\.ya?ml$' -and $segments[-1] -notlike '*.common.yaml' -and
        $segments[1..($segments.Count - 2)] -contains 'Disabled'
}


function Get-RecipeSyncEvidenceProblem {
    # Why a recipe's evidence does not prove these exact contents passed a live
    # test, or nothing when it does.
    param([string]$ApiBase, [string]$Path, [byte[]]$Bytes, [string]$EvidenceSha)

    $json = [System.Text.Encoding]::UTF8.GetString((Get-RecipeSyncBlob -ApiBase $ApiBase -Sha $EvidenceSha)).TrimStart([char]0xFEFF)
    $check = Test-RecipeEvidenceContent -Json $json -RecipePath $Path -RecipeSha256 (Get-RecipeContentHash -Bytes $Bytes)
    if ($check.Failures.Count) { return ($check.Failures -join ' ') }
}


function Backup-RecipeSyncFile {
    param([string]$Path, [string]$Root, [string]$BackupRoot, [switch]$Move)

    $relative = [System.IO.Path]::GetRelativePath($Root, $Path)
    $target = Join-Path $BackupRoot $relative
    New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
    if ($Move) {
        Move-Item -LiteralPath $Path -Destination $target -Force
    } else {
        Copy-Item -LiteralPath $Path -Destination $target -Force
    }
}


function Sync-YardstickRecipeRepository {
    <#
    .SYNOPSIS
    Brings one repository's recipes into its destination folder. Throws when
    GitHub cannot be reached; whatever was synced before that is kept and recorded.

    .PARAMETER Repository
    An entry from Get-YardstickRecipeRepository.

    .PARAMETER BackupRoot
    Where files the sync replaces or removes in Recipes are copied first.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Repository,
        [switch]$RequireEvidence,
        [Parameter(Mandatory)][string]$BackupRoot
    )

    $apiBase = "repos/$($Repository.Owner)/$($Repository.Name)"
    $isMain = $Repository.Role -eq 'Main'
    $result = [PSCustomObject]@{
        UpToDate  = $false
        Added     = [System.Collections.Generic.List[string]]::new()
        Updated   = [System.Collections.Generic.List[string]]::new()
        Removed   = [System.Collections.Generic.List[string]]::new()
        Unchanged = 0
        Kept      = [System.Collections.Generic.List[string]]::new()
        Rejected  = [System.Collections.Generic.List[string]]::new()
        Warnings  = [System.Collections.Generic.List[string]]::new()
    }

    $ref = $Repository.Ref
    if (-not $ref) { $ref = [string](Invoke-YardstickGitHubApi $apiBase).default_branch }
    $escapedRef = [uri]::EscapeDataString($ref).Replace('%2F', '/')
    $tree = Invoke-YardstickGitHubApi "$apiBase/git/trees/${escapedRef}?recursive=1"
    $state = Read-RecipeSyncState -Path $Repository.StatePath -Spec $Repository.Spec

    $remote = [ordered]@{}
    $evidence = @{}
    foreach ($item in @($tree.tree)) {
        if ($item.type -ne 'blob') { continue }
        $path = [string]$item.path
        if (Test-RecipeSyncPath -Path $path) {
            $remote[$path] = [string]$item.sha
        } elseif ($path -match '^Tests/RecipeEvidence/[^/]+\.json$') {
            $evidence[(Get-RecipeEvidenceId -RecipePath $path)] = [string]$item.sha
        }
    }
    # A listing GitHub cut short can't say what was removed.
    $complete = -not $tree.truncated
    if (-not $complete) {
        $result.Warnings.Add("GitHub returned only part of $($Repository.Spec), so recipes removed from it are not removed here.")
    }

    # Every run compares every file, even when the repository has not changed, so a
    # recipe deleted or edited here is put right. Only files that differ are
    # downloaded.
    $files = $state.files
    $rejected = $state.rejected
    $finished = $false
    try {
        foreach ($path in $remote.Keys) {
            $sha = $remote[$path]
            try {
                $local = Resolve-RecipeSyncPath -Root $Repository.Destination -RepositoryPath $path
            } catch {
                $result.Warnings.Add("Skipped $path from $($Repository.Spec): $($_.Exception.Message)")
                continue
            }
            $relative = $path.Substring('Recipes/'.Length)
            $localBytes = if (Test-Path -LiteralPath $local -PathType Leaf) { [System.IO.File]::ReadAllBytes($local) } else { $null }
            $localSha = if ($null -ne $localBytes) { Get-YardstickGitBlobSha -Bytes $localBytes } else { $null }

            if ($localSha -eq $sha) {
                $files[$path] = $sha
                $rejected.Remove($path)
                $result.Unchanged++
                continue
            }
            # In Recipes, a file the sync wrote and that has been edited since stays.
            if ($isMain -and $null -ne $localBytes -and $files.ContainsKey($path) -and $localSha -ne $files[$path]) {
                $result.Kept.Add($relative)
                continue
            }

            if ($RequireEvidence) {
                # Without evidence the recipe can't be taken, so it isn't downloaded;
                # nor is a recipe and evidence pair already refused.
                $evidenceSha = $evidence[(Get-RecipeEvidenceId -RecipePath $path)]
                if (-not $evidenceSha) {
                    $rejected.Remove($path)
                    $result.Rejected.Add("${relative}: it has no evidence file ($(Get-RecipeEvidencePath -RecipePath $path))")
                    continue
                }
                $blobs = "$sha $evidenceSha"
                if ($rejected.ContainsKey($path) -and $rejected[$path].blobs -eq $blobs) {
                    $result.Rejected.Add("${relative}: $($rejected[$path].reason)")
                    continue
                }
            }
            $bytes = Get-RecipeSyncBlob -ApiBase $apiBase -Sha $sha
            if ($RequireEvidence) {
                $problem = Get-RecipeSyncEvidenceProblem -ApiBase $apiBase -Path $path -Bytes $bytes -EvidenceSha $evidenceSha
                if ($problem) {
                    $rejected[$path] = @{ blobs = $blobs; reason = $problem }
                    $result.Rejected.Add("${relative}: $problem")
                    continue
                }
            }
            $rejected.Remove($path)

            if ($isMain -and $null -ne $localBytes) {
                Backup-RecipeSyncFile -Path $local -Root $Repository.Destination -BackupRoot $BackupRoot
            }
            # Only replaces what was read above, so an edit saved while the files
            # were downloading is kept.
            $readHash = if ($null -ne $localBytes) { [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($localBytes)) } else { '' }
            try {
                Write-YardstickFileSafely -Path $local -Content $bytes -ExpectedHash $readHash
            } catch {
                $now = if (Test-Path -LiteralPath $local -PathType Leaf) { [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData([System.IO.File]::ReadAllBytes($local))) } else { '' }
                if ($now -eq $readHash) { throw }
                $result.Kept.Add("$relative (changed here during the sync)")
                continue
            }
            if ($null -ne $localBytes) { $result.Updated.Add($relative) } else { $result.Added.Add($relative) }
            $files[$path] = $sha
        }

        if ($complete) {
            foreach ($path in @($rejected.Keys | Where-Object { -not $remote.Contains($_) })) { $rejected.Remove($path) }
            foreach ($path in @($files.Keys | Where-Object { -not $remote.Contains($_) })) {
                $local = Resolve-RecipeSyncPath -Root $Repository.Destination -RepositoryPath $path
                $relative = $path.Substring('Recipes/'.Length)
                if (Test-Path -LiteralPath $local -PathType Leaf) {
                    $unchanged = (Get-YardstickGitBlobSha -Bytes ([System.IO.File]::ReadAllBytes($local))) -eq $files[$path]
                    if ($isMain -and -not $unchanged) {
                        $result.Kept.Add("$relative (removed from the repository)")
                    } elseif ($isMain) {
                        Backup-RecipeSyncFile -Path $local -Root $Repository.Destination -BackupRoot $BackupRoot -Move
                        $result.Removed.Add($relative)
                    } else {
                        Remove-Item -LiteralPath $local -Force
                        $result.Removed.Add($relative)
                    }
                }
                $files.Remove($path)
            }

            # A linked repository's folder is a mirror: nothing else belongs there.
            if (-not $isMain -and (Test-Path -LiteralPath $Repository.Destination -PathType Container)) {
                foreach ($file in Get-ChildItem -LiteralPath $Repository.Destination -Recurse -File -Force) {
                    $repositoryPath = 'Recipes/' + [System.IO.Path]::GetRelativePath($Repository.Destination, $file.FullName).Replace('\', '/')
                    if (-not $remote.Contains($repositoryPath)) {
                        Remove-Item -LiteralPath $file.FullName -Force
                        $result.Removed.Add($repositoryPath.Substring('Recipes/'.Length))
                    }
                }
            }
        }
        $finished = $true
    } finally {
        # Recorded even when GitHub stopped the sync part way, so the next run
        # carries on from here.
        $state.ref = $ref
        $state.treeSha = if ($finished -and $complete) { [string]$tree.sha } else { '' }
        $state.syncedAt = (Get-Date).ToUniversalTime().ToString('o')
        $state.files = $files
        $state.rejected = $rejected
        Save-RecipeSyncState -Path $Repository.StatePath -State $state
    }

    $result.UpToDate = $result.Added.Count -eq 0 -and $result.Updated.Count -eq 0 -and $result.Removed.Count -eq 0
    $result
}


function Invoke-YardstickRecipeSync {
    <#
    .SYNOPSIS
    Syncs recipes from every configured repository, when syncRecipes is on. Never
    throws.

    .PARAMETER Preferences
    Preferences from Import-YardstickPreferences.

    .PARAMETER ProjectRoot
    The Yardstick folder.

    .OUTPUTS
    PSCustomObjects with Level (Info, Detail or Warning, as Write-Log takes) and
    Message, in order.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Preferences,
        [Parameter(Mandatory)][string]$ProjectRoot
    )

    $messages = [System.Collections.Generic.List[object]]::new()
    $say = { param($Level, $Message) $messages.Add([PSCustomObject]@{ Level = $Level; Message = $Message }) }
    if ($Preferences.syncRecipes -ne $true) { return }

    try {
        $repositories = @(Get-YardstickRecipeRepository -Preferences $Preferences -ProjectRoot $ProjectRoot)
    } catch {
        & $say 'Warning' "Could not read the recipe repositories from preferences: $($_.Exception.Message)"
        return $messages.ToArray()
    }
    if ($repositories.Count -eq 0) {
        & $say 'Warning' "syncRecipes is on but neither 'recipeRepository' nor 'additionalRecipeRepositories' is set, so there is nothing to sync."
        return $messages.ToArray()
    }

    $artifacts = Join-Path $ProjectRoot 'Artifacts'
    try {
        $layout = Get-YardstickProjectLayout -ProjectRoot $ProjectRoot
        if ($layout.PSObject.Properties['ArtifactsPath']) { $artifacts = $layout.ArtifactsPath }
    } catch {
        Write-Verbose "Using $artifacts for sync backups: $($_.Exception.Message)"
    }
    $backupStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $requireEvidence = $Preferences.requireRecipeEvidence -eq $true

    foreach ($repository in $repositories) {
        if ($repository.Error) {
            & $say 'Warning' "Recipe repository skipped: $($repository.Error)"
            continue
        }
        $backupRoot = Join-Path $artifacts "RecipeSyncBackups\$backupStamp\$($repository.Key)"
        try {
            $result = Sync-YardstickRecipeRepository -Repository $repository -RequireEvidence:$requireEvidence -BackupRoot $backupRoot
        } catch {
            $state = Read-RecipeSyncState -Path $repository.StatePath -Spec $repository.Spec
            $fallback = if ($state.syncedAt) {
                "Using the recipes synced on $(([datetime]$state.syncedAt).ToLocalTime().ToString('yyyy-MM-dd HH:mm'))."
            } else {
                'Nothing has been synced from it yet.'
            }
            & $say 'Warning' "Could not sync recipes from $($repository.Spec): $($_.Exception.Message) $fallback"
            continue
        }

        # Kept and refused recipes are reported every run, changed or not.
        if ($result.UpToDate) {
            & $say 'Info' "Recipes from $($repository.Spec) are up to date in $($repository.Destination)."
        } else {
            $summary = "Synced recipes from $($repository.Spec) into $($repository.Destination): $($result.Added.Count) added, $($result.Updated.Count) updated, $($result.Removed.Count) removed, $($result.Unchanged) unchanged"
            if ($result.Kept.Count) { $summary += ", $($result.Kept.Count) kept as edited here" }
            if ($result.Rejected.Count) { $summary += ", $($result.Rejected.Count) not taken without evidence" }
            & $say 'Info' "$summary."
        }
        foreach ($name in $result.Added) { & $say 'Detail' "Added recipe $name from $($repository.Spec)." }
        foreach ($name in $result.Updated) { & $say 'Detail' "Updated recipe $name from $($repository.Spec)." }
        foreach ($name in $result.Removed) { & $say 'Detail' "Removed recipe $name, which $($repository.Spec) no longer has." }
        if (($result.Updated.Count -or $result.Removed.Count) -and $repository.Role -eq 'Main') {
            & $say 'Detail' "The previous copies are in $backupRoot."
        }
        foreach ($name in $result.Kept) {
            & $say 'Warning' "Kept recipe $name, which was edited here since it was last synced from $($repository.Spec). Delete it, or undo the edit, to take the repository's version."
        }
        # One warning each for a few; for many, such as a repository without
        # evidence files, one warning and the reasons in the detail log.
        if ($result.Rejected.Count -gt 5) {
            $names = @($result.Rejected | Select-Object -First 5 | ForEach-Object { ($_ -split ':', 2)[0] })
            & $say 'Warning' "Did not take $($result.Rejected.Count) recipes from $($repository.Spec) (requireRecipeEvidence is on), including $($names -join ', '). Their evidence does not prove they passed a live test with these contents; the reasons are in the detail log."
            foreach ($rejection in $result.Rejected) { & $say 'Detail' "Did not take recipe $rejection" }
        } else {
            foreach ($rejection in $result.Rejected) {
                & $say 'Warning' "Did not take recipe $rejection (requireRecipeEvidence is on)."
            }
        }
        foreach ($warning in $result.Warnings) { & $say 'Warning' $warning }
    }

    $messages.ToArray()
}


Export-ModuleMember -Function @(
    'ConvertFrom-YardstickRecipeRepositorySpec',
    'Get-YardstickRecipeRepository',
    'Get-YardstickRecipeSearchPath',
    'Get-YardstickGitBlobSha',
    'Sync-YardstickRecipeRepository',
    'Invoke-YardstickRecipeSync'
)
