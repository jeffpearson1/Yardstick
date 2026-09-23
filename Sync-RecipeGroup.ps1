<#
.SYNOPSIS
    Rewrites a recipe group in RecipeGroups.yaml so it matches the recipes that
    currently live in the matching recipe folder.

.DESCRIPTION
    The Development group is meant to mirror the contents of Recipes\Development,
    but recipes get added to and promoted out of that folder without the group
    being updated. This script regenerates the group's entry list from the recipe
    files on disk and reports what changed.

    Only the targeted group's block is rewritten; every other group in the file is
    left untouched, comments and all. Recipes inside a Disabled folder are ignored,
    matching how Deploy-YardstickApps.ps1 resolves recipes.

.PARAMETER GroupName
    The group in RecipeGroups.yaml to sync. Defaults to "Development".

.PARAMETER RecipeFolder
    The folder to read recipe ids from. Defaults to the folder named after the
    group underneath the Recipes path from Preferences.yaml.

.PARAMETER GroupFile
    Path to the recipe group file. Defaults to RecipeGroups.yaml beside this script.

.PARAMETER AllowEmpty
    Permit the sync to run when the recipe folder contains no recipes. Without this
    the script stops, so a mistyped -RecipeFolder cannot silently empty a group.

.EXAMPLE
    .\Sync-RecipeGroup.ps1 -WhatIf
    Shows what would be added to and removed from the Development group.

.EXAMPLE
    .\Sync-RecipeGroup.ps1

.EXAMPLE
    .\Sync-RecipeGroup.ps1 -GroupName TestApps -RecipeFolder .\Recipes\Testing
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [String]$GroupName = 'Development',

    [String]$RecipeFolder,

    [String]$GroupFile = "$PSScriptRoot\RecipeGroups.yaml",

    [Switch]$AllowEmpty
)

$ErrorActionPreference = 'Stop'

Import-Module powershell-yaml -ErrorAction Stop

if (-not $RecipeFolder) {
    $prefsPath = Join-Path $PSScriptRoot 'Preferences.yaml'
    if (-not (Test-Path $prefsPath)) {
        Write-Error "Preferences.yaml not found. Pass -RecipeFolder explicitly."
        exit 1
    }
    $prefs = Get-Content $prefsPath | ConvertFrom-Yaml
    if ([string]::IsNullOrWhiteSpace($prefs.Recipes)) {
        Write-Error "Preferences.yaml does not define a Recipes path. Pass -RecipeFolder explicitly."
        exit 1
    }
    $RecipeFolder = Join-Path $prefs.Recipes $GroupName
}

if (-not (Test-Path $RecipeFolder -PathType Container)) {
    Write-Error "Recipe folder '$RecipeFolder' does not exist."
    exit 1
}
if (-not (Test-Path $GroupFile -PathType Leaf)) {
    Write-Error "Recipe group file '$GroupFile' does not exist."
    exit 1
}


function Get-RecipeIdsFromFolder {
    <#
    .SYNOPSIS
    Returns the sorted, de-duplicated recipe ids found under a folder. A recipe id
    is the file name with the .yaml/.yml extension stripped, so multi-part names
    like 7zip.arm64 survive intact.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [String]$Path
    )

    $ids = [System.Collections.Generic.List[String]]::new()
    $seen = [System.Collections.Generic.HashSet[String]]::new([StringComparer]::OrdinalIgnoreCase)

    $files = @(Get-ChildItem -Path $Path -Recurse -File -Include '*.yaml', '*.yml' |
        Where-Object { $_.FullName -notlike "*\Disabled\*" } |
        Sort-Object BaseName)

    foreach ($file in $files) {
        if ($seen.Add($file.BaseName)) {
            $ids.Add($file.BaseName) | Out-Null
        } else {
            Write-Warning "Duplicate recipe id '$($file.BaseName)' at $($file.FullName); keeping the first one found."
        }
    }

    return , $ids.ToArray()
}


function Set-GroupEntries {
    <#
    .SYNOPSIS
    Replaces one group's list of entries in the recipe group file, leaving the rest
    of the file byte-for-byte alone. Returns the new file content as a string.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [String[]]$Line,

        [Parameter(Mandatory = $true)]
        [String]$Name,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [String[]]$Entry,

        [Parameter(Mandatory = $true)]
        [String]$Newline
    )

    $keyPattern = "^$([regex]::Escape($Name))\s*:"
    $start = -1
    for ($i = 0; $i -lt $Line.Count; $i++) {
        if ($Line[$i] -match $keyPattern) { $start = $i; break }
    }

    # The group's block runs until the next line that starts in column 0 (the next
    # key, or a comment attached to it) or the end of the file.
    $indent = '  '
    if ($start -ge 0) {
        $end = $Line.Count
        for ($i = $start + 1; $i -lt $Line.Count; $i++) {
            if ($Line[$i] -match '^\S') { $end = $i; break }
            if ($Line[$i] -match '^(\s+)-\s') { $indent = $Matches[1] }
        }
    }

    $body = @($Entry | ForEach-Object { "$indent- $_" })

    if ($start -lt 0) {
        Write-Verbose "Group '$Name' is not in the file yet; appending it."
        $head = @($Line)
        while ($head.Count -gt 0 -and [string]::IsNullOrWhiteSpace($head[-1])) {
            $head = @($head[0..($head.Count - 2)])
        }
        return (@($head) + @("${Name}:") + $body) -join $Newline
    }

    $before = @($Line[0..$start])
    # Keep any blank lines that separated this block from the next group.
    while ($end - 1 -gt $start -and [string]::IsNullOrWhiteSpace($Line[$end - 1])) { $end-- }
    $after = if ($end -lt $Line.Count) { @($Line[$end..($Line.Count - 1)]) } else { @() }

    return (@($before) + $body + $after) -join $Newline
}


$FolderIds = Get-RecipeIdsFromFolder -Path $RecipeFolder
Write-Host "Found $($FolderIds.Count) recipe(s) in $RecipeFolder"

if ($FolderIds.Count -eq 0 -and -not $AllowEmpty) {
    Write-Error "No recipes found in '$RecipeFolder'. Re-run with -AllowEmpty if the group really should be emptied."
    exit 1
}

try {
    $Groups = Get-Content $GroupFile -Raw | ConvertFrom-Yaml
} catch {
    Write-Error "Unable to parse '$GroupFile': $_"
    exit 1
}

$CurrentIds = @()
if ($Groups -and $Groups.ContainsKey($GroupName) -and $Groups[$GroupName]) {
    $CurrentIds = @($Groups[$GroupName] | ForEach-Object { [String]$_ })
} else {
    Write-Warning "Group '$GroupName' is empty or not defined in $GroupFile; it will be created from the folder contents."
}

$Added = @($FolderIds | Where-Object { $_ -notin $CurrentIds })
$Removed = @($CurrentIds | Where-Object { $_ -notin $FolderIds })

foreach ($id in $Added) { Write-Host "  + $id" -ForegroundColor Green }
foreach ($id in $Removed) { Write-Host "  - $id" -ForegroundColor Yellow }

if ($Added.Count -eq 0 -and $Removed.Count -eq 0) {
    Write-Host "Group '$GroupName' already matches $RecipeFolder; nothing to do." -ForegroundColor Cyan
    return
}

if (-not $PSCmdlet.ShouldProcess($GroupFile, "Sync group '$GroupName' ($($Added.Count) added, $($Removed.Count) removed)")) {
    return
}

$Raw = Get-Content $GroupFile -Raw
$Newline = if ($Raw -match "`r`n") { "`r`n" } else { "`n" }
$TrailingNewline = if ($Raw -match "(`r`n|`n)$") { $Newline } else { '' }

$Updated = Set-GroupEntries -Line ($Raw -split "`r?`n") -Name $GroupName -Entry $FolderIds -Newline $Newline

# Parse the result before committing it, so a formatting bug cannot leave the file
# in a state Yardstick can no longer read.
try {
    $check = $Updated | ConvertFrom-Yaml
} catch {
    Write-Error "The rewritten group file failed to parse, so it was not saved: $_"
    exit 2
}
$checkIds = @(if ($check[$GroupName]) { $check[$GroupName] | ForEach-Object { [String]$_ } })
if (@(Compare-Object -ReferenceObject @($FolderIds) -DifferenceObject $checkIds -SyncWindow 0).Count -gt 0) {
    Write-Error "The rewritten group '$GroupName' does not match the folder contents, so the file was not saved."
    exit 2
}

[System.IO.File]::WriteAllText($GroupFile, $Updated + $TrailingNewline, [System.Text.UTF8Encoding]::new($false))
Write-Host "Updated '$GroupName' in $GroupFile - $($Added.Count) added, $($Removed.Count) removed, $($FolderIds.Count) total." -ForegroundColor Cyan
