<#
.SYNOPSIS
    Assigns the newest published version of one or more Yardstick recipes to an
    Entra ID group in Intune.

.DESCRIPTION
    Resolves recipes either by recipe group name (as defined in RecipeGroups.yaml)
    or by individual recipe id, looks up each recipe's display name, finds the
    newest matching Win32 app in Intune, and adds a group assignment for it.

    Both -Group and -ApplicationId may be supplied together; the resulting recipe
    list is the de-duplicated union of the two.

    Only the newest version is targeted. Detection anchors and (N-x) versions are
    ignored. The script is idempotent - a recipe whose newest version already
    carries the requested assignment is skipped.

.PARAMETER Group
    One or more recipe group names from RecipeGroups.yaml.

.PARAMETER ApplicationId
    One or more individual recipe ids (the recipe file name without extension).

.PARAMETER TargetGroup
    The Entra ID group to deploy to. Accepts either a group object ID (GUID) or a
    group display name, which is resolved through Graph.

.PARAMETER Intent
    Assignment intent. Defaults to "available".

.PARAMETER Notification
    End-user notification behaviour. Defaults to "hideAll".

.PARAMETER Exclude
    Create an exclusion assignment instead of an inclusion.

.EXAMPLE
    .\Deploy-YardstickApps.ps1 -Group Development -TargetGroup "00000000-1111-2222-3333-444444444444"

.EXAMPLE
    .\Deploy-YardstickApps.ps1 -Group TestApps -ApplicationId googlechrome,7zip -TargetGroup "Lab Workstations" -Intent required

.EXAMPLE
    .\Deploy-YardstickApps.ps1 -ApplicationId wingetautoupdate -TargetGroup "Lab Workstations" -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [String[]]$Group,

    [Alias("AppId", "AppIds")]
    [String[]]$ApplicationId,

    [Parameter(Mandatory = $true)]
    [String]$TargetGroup,

    [ValidateSet("available", "required", "uninstall")]
    [String]$Intent = "available",

    [ValidateSet("showAll", "showReboot", "hideAll")]
    [String]$Notification = "hideAll",

    [Switch]$Exclude
)

$ErrorActionPreference = 'Stop'

$Global:LogLocation = $PSScriptRoot
$Global:LogFile = 'YDeploy.log'

Import-Module powershell-yaml -ErrorAction Stop
Import-Module "$PSScriptRoot\Modules\YardstickSupport.psm1" -Scope Global -Force
Import-Module "$PSScriptRoot\Modules\YardstickCredential.psm1" -Scope Global -Force
Import-Module IntuneWin32App -ErrorAction Stop

Write-Log -Init

if (-not $Group -and -not $ApplicationId) {
    Write-Log "ERROR: Supply at least one of -Group or -ApplicationId."
    exit 1
}

try {
    $Prefs = Get-Content "$PSScriptRoot\Preferences.yaml" | ConvertFrom-Yaml
} catch {
    Write-Log "ERROR: Unable to open Preferences.yaml!"
    exit 1
}

$RecipesPath = $Prefs.Recipes
$GroupTargetType = '#microsoft.graph.groupAssignmentTarget'
$ExclusionTargetType = '#microsoft.graph.exclusionGroupAssignmentTarget'


function Resolve-RecipeIdList {
    <#
    .SYNOPSIS
    Builds the de-duplicated list of recipe ids from the requested recipe groups
    and any explicitly named ids, preserving the order they were requested in.
    #>
    param(
        [String[]]$GroupName,
        [String[]]$RecipeId
    )

    $ids = [System.Collections.Generic.List[String]]::new()
    $seen = [System.Collections.Generic.HashSet[String]]::new([StringComparer]::OrdinalIgnoreCase)

    if ($GroupName) {
        try {
            $groupFile = Get-Content "$PSScriptRoot\RecipeGroups.yaml" | ConvertFrom-Yaml
        } catch {
            Write-Log "ERROR: Unable to open RecipeGroups.yaml: $_"
            exit 3
        }
        foreach ($name in $GroupName) {
            if (-not $groupFile.ContainsKey($name)) {
                Write-Log "ERROR: Recipe group '$name' is not defined in RecipeGroups.yaml. Available groups: $($groupFile.Keys -join ', ')"
                exit 3
            }
            foreach ($entry in $groupFile[$name]) {
                if ($seen.Add([String]$entry)) { $ids.Add([String]$entry) | Out-Null }
            }
        }
    }

    foreach ($entry in $RecipeId) {
        if ($seen.Add([String]$entry)) { $ids.Add([String]$entry) | Out-Null }
    }

    return $ids
}


function Get-RecipeDisplayName {
    <#
    .SYNOPSIS
    Returns the Intune display name declared by a recipe, or $null when the recipe
    file cannot be found or parsed.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [String]$RecipeId
    )

    $pattern = "^$([regex]::Escape($RecipeId))\.ya{0,1}ml$"
    $file = @(Get-ChildItem -Path $RecipesPath -Recurse -File |
        Where-Object { $_.FullName -notlike "*\Disabled\*" -and $_.Name -match $pattern }) |
        Select-Object -First 1

    if (-not $file) {
        Write-Log "ERROR: No recipe file found for '$RecipeId' (excluding the Disabled folder)."
        return $null
    }

    try {
        $recipe = Get-Content $file.FullName | ConvertFrom-Yaml
    } catch {
        Write-Log "ERROR: Unable to parse recipe $($file.FullName): $_"
        return $null
    }

    if ($recipe.ContainsKey('base')) {
        try {
            $recipe = Merge-RecipeWithBase -Recipe $recipe -RecipesPath $RecipesPath
        } catch {
            Write-Log "ERROR: Failed to resolve base recipe for '$RecipeId': $_"
            return $null
        }
    }

    if ([string]::IsNullOrWhiteSpace($recipe.displayName)) {
        Write-Log "ERROR: Recipe '$RecipeId' does not declare a displayName."
        return $null
    }

    return [String]$recipe.displayName
}


function Resolve-TargetGroup {
    <#
    .SYNOPSIS
    Turns the -TargetGroup argument into an object ID, accepting either a GUID or
    an Entra group display name.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [String]$Identity
    )

    $guid = [guid]::Empty
    if ([guid]::TryParse($Identity, [ref]$guid)) {
        $group = Invoke-YardstickGraphRequest -Resource "groups/$Identity" -ApiVersion 'v1.0'
        return [PSCustomObject]@{ Id = $group.id; DisplayName = $group.displayName }
    }

    # OData escapes a single quote by doubling it.
    $filter = [uri]::EscapeDataString("displayName eq '$($Identity -replace "'", "''")'")
    $found = @(Invoke-YardstickGraphRequest -Resource "groups?`$filter=$filter&`$select=id,displayName" -ApiVersion 'v1.0')

    if ($found.Count -eq 0) {
        throw "No Entra group found with display name '$Identity'."
    }
    if ($found.Count -gt 1) {
        throw "Display name '$Identity' matches $($found.Count) Entra groups. Pass the group object ID instead."
    }
    return [PSCustomObject]@{ Id = $found[0].id; DisplayName = $found[0].displayName }
}


$RecipeIds = Resolve-RecipeIdList -GroupName $Group -RecipeId $ApplicationId
if ($RecipeIds.Count -eq 0) {
    Write-Log "ERROR: The requested recipe group(s) resolved to no recipes."
    exit 3
}
Write-Log "Resolved $($RecipeIds.Count) recipe(s) to deploy."

$Script:IntuneCredential = Initialize-YardstickIntuneCredential -Preferences $Prefs
try {
    Connect-AutoMSIntuneGraph
} catch {
    Write-Log "ERROR: Unable to authenticate to Microsoft Graph: $_"
    Write-Log "If the client secret was rotated, store the new one with .\Set-YardstickCredential.ps1"
    exit 1
}

try {
    $Target = Resolve-TargetGroup -Identity $TargetGroup
} catch {
    Write-Log "ERROR: $_"
    exit 1
}
Write-Log "Target group: $($Target.DisplayName) ($($Target.Id)) - intent '$Intent'$(if ($Exclude) { ' (exclusion)' })"

$TargetType = if ($Exclude) { $ExclusionTargetType } else { $GroupTargetType }
$Results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($RecipeId in $RecipeIds) {
    $record = [PSCustomObject]@{
        RecipeId    = $RecipeId
        DisplayName = $null
        Version     = $null
        Status      = 'Failed'
        Detail      = $null
    }
    $Results.Add($record) | Out-Null

    try {
        Connect-AutoMSIntuneGraph

        $displayName = Get-RecipeDisplayName -RecipeId $RecipeId
        if (-not $displayName) {
            $record.Detail = 'Recipe could not be resolved'
            continue
        }
        $record.DisplayName = $displayName

        $anchorName = Get-DetectAnchorName -DisplayName $displayName
        # Get-SameAppAllVersions returns its list wrapped in an outer array, so it
        # must be assigned before filtering - piping it straight into Where-Object
        # hands the whole list over as a single item.
        $allVersions = Get-SameAppAllVersions -DisplayName $displayName
        $versions = @($allVersions | Where-Object { $_.DisplayName -ne $anchorName })
        if ($versions.Count -eq 0) {
            Write-Log "ERROR: No Intune application found for '$displayName' ($RecipeId)."
            $record.Status = 'NotFound'
            $record.Detail = "No Intune app named '$displayName'"
            continue
        }

        $newest = $versions[0]
        $record.Version = $newest.displayVersion
        Write-Log "Newest version of '$displayName' is $($newest.displayVersion) ($($newest.id))"

        if (Test-YardstickAssignmentPresent -AppId $newest.id -CountKey $Target.Id -Intent $Intent -TargetType $TargetType) {
            Write-Log "Assignment already present on $($newest.DisplayName); nothing to do."
            $record.Status = 'AlreadyAssigned'
            continue
        }

        if (-not $PSCmdlet.ShouldProcess("$($newest.DisplayName) ($($newest.id))", "Assign to $($Target.DisplayName) with intent '$Intent'")) {
            $record.Status = 'Skipped'
            $record.Detail = 'WhatIf'
            continue
        }

        $assignmentParams = @{
            ID      = $newest.id
            GroupID = $Target.Id
            Intent  = $Intent
        }
        if ($Exclude) {
            $assignmentParams['Exclude'] = $true
        } else {
            $assignmentParams['Include'] = $true
            $assignmentParams['Notification'] = $Notification
        }

        $addWarnings = @()
        # The IntuneWin32App cmdlets abandon their Begin block with a bare `break`
        # on some failures, which is not catchable and would unwind this foreach.
        foreach ($breakGuard in 1) {
            Add-IntuneWin32AppAssignmentGroup @assignmentParams -WarningAction SilentlyContinue -WarningVariable addWarnings | Out-Null
        }
        foreach ($warning in $addWarnings) {
            Write-Log "WARNING from Intune while assigning $($newest.id): $warning"
        }

        if (Test-YardstickAssignmentPresent -AppId $newest.id -CountKey $Target.Id -Intent $Intent -TargetType $TargetType) {
            Write-Log "Assigned '$($newest.DisplayName)' to $($Target.DisplayName)."
            $record.Status = 'Assigned'
        } else {
            Write-Log "ERROR: Assignment of '$($newest.DisplayName)' could not be verified after the add."
            $record.Detail = 'Assignment not present after add'
        }
    } catch {
        Write-Log "ERROR: Failed to deploy '$RecipeId': $_"
        $record.Detail = "$_"
    }
}

Write-Log "----- Summary -----"
$Results | Format-Table RecipeId, DisplayName, Version, Status, Detail -AutoSize | Out-String -Width 4096 | Write-Host

$failed = @($Results | Where-Object Status -in @('Failed', 'NotFound'))
Write-Log "$(@($Results | Where-Object Status -eq 'Assigned').Count) assigned, $(@($Results | Where-Object Status -eq 'AlreadyAssigned').Count) already assigned, $($failed.Count) failed."
if ($failed.Count -gt 0) { exit 2 }
