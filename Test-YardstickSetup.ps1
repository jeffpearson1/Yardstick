#Requires -Version 7.0
<#
.SYNOPSIS
Checks that Yardstick is set up correctly on this machine.

.DESCRIPTION
Runs every check a working install depends on and prints one line per check:

    [ OK ]  the check passed
    [WARN]  an optional feature will not work, or something needs attention soon
    [FAIL]  normal runs will fail until this is fixed

It covers PowerShell and its modules, preferences.yaml, the working folders and
files the current account cannot change, the stored Intune credential and the
Graph permissions it actually has, the client secret's expiry,
IntuneWinAppUtil.exe, email delivery, the scheduled task and recipes without
icons.

Nothing is changed, except that a scratch file is written to (and removed from)
each working folder to prove it is writable. Run it after Install-Yardstick.ps1,
after changing preferences, or before a scheduled run as a pre-flight check.

On a development-only instance (developmentOnly: true), the Intune, email and
scheduled task checks are skipped.

The Intune credential is stored per Windows user, so run this as the account
that runs Yardstick (for example the scheduled task's account).

.PARAMETER Offline
Skip the checks that sign in to Microsoft Graph.

.PARAMETER PassThru
Also return the check results as objects (Area, Check, Level, Message).

.PARAMETER ProjectRoot
The Yardstick folder to check. Defaults to the folder this script is in. The
check always runs this script's own modules, whichever folder it checks.

.EXAMPLE
.\Test-YardstickSetup.ps1

.EXAMPLE
.\Test-YardstickSetup.ps1 -Offline

.OUTPUTS
Exit code 0 when nothing failed, 1 when at least one check failed.
#>
[CmdletBinding()]
param(
    [switch]$Offline,
    [switch]$PassThru,
    [string]$ProjectRoot = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = (Get-Item -LiteralPath $ProjectRoot).FullName

# Write-Log (used by the credential and support modules) needs somewhere to write.
# The installer calls this script with its own log already set up, so keep that.
if (-not $Global:LogLocation) { $Global:LogLocation = $ProjectRoot }
if (-not $Global:LogFile) { $Global:LogFile = 'YSetupCheck.log' }

$results = [System.Collections.Generic.List[object]]::new()
$developmentOnly = $false

function Add-Check {
    param(
        [Parameter(Mandatory)][string]$Check,
        [Parameter(Mandatory)][ValidateSet('OK', 'Warn', 'Fail', 'Info', 'Skip')][string]$Level,
        [Parameter(Mandatory)][string]$Message
    )
    $results.Add([pscustomobject]@{ Area = $script:Area; Check = $Check; Level = $Level; Message = $Message })
    if (Get-Command Write-YardstickSetupStatus -ErrorAction SilentlyContinue) {
        Write-YardstickSetupStatus -Message $Message -Level $Level
    } else {
        $tag = @{ OK = '[ OK ]'; Warn = '[WARN]'; Fail = '[FAIL]'; Skip = '[SKIP]'; Info = '[ .. ]' }[$Level]
        Write-Host "$tag $Message"
    }
}

function Start-Area {
    param([string]$Name)
    $script:Area = $Name
    Write-Host ''
    Write-Host $Name -ForegroundColor Cyan
}

function Test-FolderWritable {
    param([string]$Folder)
    $probe = Join-Path $Folder ".yardstick-write-test-$PID.tmp"
    try {
        [IO.File]::WriteAllText($probe, 'ok')
        Remove-Item -LiteralPath $probe -Force
        $true
    } catch {
        $false
    }
}

function Get-MissingIconCount {
    # Dot-sourced here so the icon helpers stay local to this function.
    param([string[]]$RecipeFolder, [string]$IconsFolder)
    . (Join-Path $PSScriptRoot 'Get-YardstickIcon.ps1') -LoadOnly
    $recipes = @(Get-RecipeIconInfo -RecipeFolder $RecipeFolder)
    [pscustomobject]@{
        Recipes = $recipes.Count
        Missing = @(Get-MissingIconRecipe -Recipe $recipes -IconsFolder $IconsFolder)
    }
}

Write-Host ''
Write-Host 'Yardstick setup check' -ForegroundColor Cyan
Write-Host "Folder: $ProjectRoot" -ForegroundColor DarkGray
Write-Host "User:   $env:USERDOMAIN\$env:USERNAME" -ForegroundColor DarkGray


#region PowerShell and modules

Start-Area 'PowerShell'
Add-Check 'PowerShell' OK "PowerShell $($PSVersionTable.PSVersion)"

if (-not (Get-Module -ListAvailable -Name 'powershell-yaml')) {
    Add-Check 'powershell-yaml' Fail "The powershell-yaml module is not installed. Run .\Install-Yardstick.ps1 -Step Prerequisites, or: Install-Module powershell-yaml -Scope CurrentUser"
    Write-Host ''
    Write-Host 'Fix the failure above, then run this check again.' -ForegroundColor Red
    if ($PassThru) { $results }
    exit 1
}

Import-Module powershell-yaml
# Reload the module so a session that imported it before an update runs the
# current code. Not when Test-YardstickSetup in the module runs this script: that
# would unload the module out from under the command.
$calledFromModule = [bool](Get-PSCallStack | Where-Object { $_.InvocationInfo.MyCommand.ModuleName -eq 'Yardstick' })
Import-Module (Join-Path $PSScriptRoot 'Yardstick.psd1') -Force:(-not $calledFromModule)

$prefs = Import-YardstickPreferences -ProjectRoot $ProjectRoot -AllowMissing
$prerequisites = Test-Prerequisites -ToolsPath $prefs.Tools
foreach ($message in $prerequisites.Errors) { Add-Check 'Prerequisites' Fail $message }
foreach ($message in $prerequisites.Warnings) { Add-Check 'Prerequisites' Warn $message }
if (-not $prerequisites.Errors) { Add-Check 'Prerequisites' OK 'Required modules and tools are available.' }

#endregion


#region Preferences and folders

Start-Area 'Preferences'
$prefsPath = Get-YardstickPreferencesPath -ProjectRoot $ProjectRoot
$havePreferences = $prefsPath -and (Test-Path -LiteralPath $prefsPath -PathType Leaf)
if ($havePreferences) {
    try {
        $prefs = Import-YardstickPreferences -ProjectRoot $ProjectRoot
        Add-Check 'Preferences file' OK "Preferences: $prefsPath"
        $developmentOnly = $prefs.developmentOnly -eq $true
        if ($developmentOnly) {
            Add-Check 'Instance' Info 'Development-only instance (developmentOnly: true): for the recipe harness, never connects to Intune.'
        }
    } catch {
        $havePreferences = $false
        Add-Check 'Preferences file' Fail "Could not read '$prefsPath': $($_.Exception.Message)"
    }
} else {
    Add-Check 'Preferences file' Fail 'No preferences.yaml found. Run .\Install-Yardstick.ps1 to create one.'
}

if ($havePreferences) {
    $check = Test-YardstickPreferences -Preferences $prefs
    foreach ($message in $check.Errors) { Add-Check 'Preferences' Fail $message }
    foreach ($message in $check.Warnings) { Add-Check 'Preferences' Warn $message }
    if (-not $check.Errors -and -not $check.Warnings) { Add-Check 'Preferences' OK 'All preferences are valid.' }
    if (-not $developmentOnly -and [string]::IsNullOrWhiteSpace([string]$prefs.defaultOwner)) {
        Add-Check 'Owner' Warn "defaultOwner is not set, so apps are uploaded without an owner. Run .\Install-Yardstick.ps1 -Step Preferences."
    }
}

Start-Area 'Folders'
foreach ($key in 'Temp', 'Buildspace', 'Published', 'Scripts', 'Icons', 'Tools') {
    $folder = [string]$prefs[$key]
    if ([string]::IsNullOrWhiteSpace($folder)) {
        Add-Check $key Fail "The $key folder is not set."
    } elseif (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        # Yardstick creates these on its next run, so a missing one is not a failure.
        Add-Check $key Warn "$key folder '$folder' does not exist yet. Yardstick creates it on the next run."
    } elseif (-not (Test-FolderWritable $folder)) {
        Add-Check $key Fail "$key folder '$folder' is not writable by $env:USERNAME."
    } else {
        Add-Check $key OK "$key`: $folder"
    }
}

# Files left by running setup or Yardstick as administrator, which this account can read but not change.
$unwritable = @(Get-YardstickUnwritablePath -ProjectRoot $ProjectRoot)
if ($unwritable.Count -gt 0) {
    Add-Check 'Permissions' Warn "$env:USERNAME cannot change $($unwritable.Count) item(s) in $ProjectRoot (for example $($unwritable[0])). Run .\Install-Yardstick.ps1 -Step Prerequisites to fix them."
} else {
    Add-Check 'Permissions' OK "$env:USERNAME can change every file in $ProjectRoot."
}

$recipeFolder = [string]$prefs.Recipes
if ($recipeFolder -and (Test-Path -LiteralPath $recipeFolder -PathType Container)) {
    $recipeCount = @(Get-YardstickRecipeFile -RecipesPath $recipeFolder).Count
    $level = if ($recipeCount) { 'OK' } else { 'Warn' }
    Add-Check 'Recipes' $level "$recipeCount recipe(s) in $recipeFolder"
}

# An override whose id matches no recipe is never applied, so a typo there would
# quietly upload the recipe with its own (or the default) scope tags.
$overridesPath = Get-YardstickRecipeOverridesPath -ProjectRoot $ProjectRoot
# With shared sync on, a run layers the local file over the last synced copy of the
# shared file, so these checks read both. The health check never syncs.
$sharedCachePath = $null
if ($prefs.syncSharedContent -eq $true) {
    $sharedCachePath = Get-YardstickSharedRecipeOverridesCachePath -ProjectRoot $ProjectRoot
    $sharedCache = Get-Item -LiteralPath $sharedCachePath -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace([string]$prefs.SharedRecipeOverrides)) {
        Add-Check 'Shared sync' OK 'Shared sync is on for icons only (SharedRecipeOverrides is not set).'
    } elseif ($sharedCache) {
        Add-Check 'Shared sync' OK "Shared sync is on. Overrides from $($prefs.SharedRecipeOverrides) were last synced $($sharedCache.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))."
    } else {
        Add-Check 'Shared sync' Warn "Shared sync is on, but overrides from $($prefs.SharedRecipeOverrides) have not been synced yet. The next run, or recipe harness script, syncs them."
    }
} else {
    Add-Check 'Shared sync' OK 'Shared sync is off (optional: syncSharedContent).'
}
try {
    $overrideSet = Import-YardstickEffectiveRecipeOverrides -Path $overridesPath -SharedPath $sharedCachePath
    $overrides = $overrideSet.Overrides
    $overrideSources = if ($overrideSet.Sources.Count) { $overrideSet.Sources -join ' and ' } else { $overridesPath }
    if ($overrides.Count -eq 0) {
        Add-Check 'Recipe overrides' OK "No recipe overrides (optional: $overridesPath)."
    } else {
        $recipeIds = @()
        if ($recipeFolder -and (Test-Path -LiteralPath $recipeFolder -PathType Container)) {
            $recipeIds = @(Get-ChildItem -LiteralPath $recipeFolder -Recurse -File -Include '*.yaml', '*.yml' | ForEach-Object BaseName)
        }
        $unmatched = @($overrides.Keys | Where-Object { $_ -notin $recipeIds } | Sort-Object)
        if ($unmatched.Count -gt 0) {
            Add-Check 'Recipe overrides' Warn "$($unmatched.Count) override(s) in $overrideSources name no recipe, so they are never applied: $($unmatched -join ', ')"
        } else {
            Add-Check 'Recipe overrides' OK "$($overrides.Count) recipe override(s) in $overrideSources"
        }
    }

    # Recipes that ship with a <YOUR_...> placeholder fail until this site's
    # overrides fill it in. Only those files are parsed, so this stays quick.
    if ($recipeFolder -and (Test-Path -LiteralPath $recipeFolder -PathType Container)) {
        $unfilled = foreach ($file in Get-YardstickRecipeFile -RecipesPath $recipeFolder) {
            if ((Get-Content -LiteralPath $file.FullName -Raw) -cnotmatch '<YOUR_[A-Z0-9_]+>') { continue }
            $recipe = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Yaml
            if ($overrides.Contains($file.BaseName)) {
                $recipe = Merge-RecipeWithOverride -Recipe $recipe -Override $overrides[$file.BaseName]
            }
            if (Get-RecipePlaceholder -Recipe $recipe) { $file.BaseName }
        }
        if ($unfilled) {
            Add-Check 'Recipe overrides' Warn "$(@($unfilled).Count) recipe(s) need values only your site knows and will fail until you set them in ${overridesPath}: $(@($unfilled) -join ', '). See Site-Specific Overrides in Docs\WritingRecipes.md."
        }
    }
} catch {
    Add-Check 'Recipe overrides' Fail "Could not read the recipe overrides: $($_.Exception.Message). Yardstick will not run until it is fixed."
}

#endregion


#region Intune credential and Graph permissions

Start-Area 'Intune connection'
if ($developmentOnly) {
    Add-Check 'Intune' Skip 'Development-only instance: Yardstick does not connect to Intune on this computer.'
} else {
    $target = Get-YardstickCredentialTarget -Preferences $prefs
    $credential = $null
    try {
        $credential = Get-YardstickIntuneCredential -Target $target
    } catch {
        Add-Check 'Credential' Fail "Could not read the Intune credential '$target': $($_.Exception.Message)"
    }

    if (-not $credential) {
        if (-not ($results | Where-Object Check -eq 'Credential')) {
            Add-Check 'Credential' Fail "No Intune credential '$target' is stored for $env:USERNAME. Run .\Install-Yardstick.ps1 -Step AppRegistration as this user."
        }
    } else {
        Add-Check 'Credential' OK "Credential '$target' (app $($credential.ClientID), tenant $($credential.TenantID))"

        $expiry = Test-YardstickSecretExpiration -Credential $credential -Preferences $prefs -NoEmail
        if (-not $expiry.ExpiresOn) {
            Add-Check 'Secret expiry' Warn "The client secret's expiry date is not known, so Yardstick cannot warn before it expires. Run .\Set-YardstickCredential.ps1 to record it."
        } elseif ($expiry.IsExpired) {
            Add-Check 'Secret expiry' Fail "The client secret expired on $($expiry.ExpiresOn.ToString('yyyy-MM-dd')). Run .\Install-Yardstick.ps1 -Step AppRegistration to create a new one."
        } elseif ($expiry.IsExpiring) {
            Add-Check 'Secret expiry' Warn "The client secret expires in $($expiry.DaysRemaining) day(s), on $($expiry.ExpiresOn.ToString('yyyy-MM-dd'))."
        } else {
            Add-Check 'Secret expiry' OK "The client secret expires on $($expiry.ExpiresOn.ToString('yyyy-MM-dd')) ($($expiry.DaysRemaining) days)."
        }

        if ($Offline) {
            Add-Check 'Graph' Skip 'Graph sign-in skipped (-Offline).'
        } else {
            $roles = $null
            try {
                $roles = @(Get-YardstickAppTokenRoles -TenantId $credential.TenantID -ClientId $credential.ClientID -ClientSecret $credential.ClientSecret)
                Add-Check 'Graph sign-in' OK 'Signed in to Microsoft Graph with the stored credential.'
            } catch {
                Add-Check 'Graph sign-in' Fail "Could not sign in to Microsoft Graph: $($_.Exception.Message)"
            }

            if ($null -ne $roles) {
                # The token's roles claim lists the permissions that have admin consent,
                # so this needs no extra Graph permission to check.
                foreach ($permission in Get-YardstickRequiredGraphPermissions) {
                    if (Test-YardstickGraphRoleGranted -Roles $roles -Name $permission.Name) {
                        Add-Check $permission.Name OK "$($permission.Name) is granted."
                    } elseif ($permission.Required) {
                        Add-Check $permission.Name Fail "$($permission.Name) is not granted ($($permission.Reason)). Grant it with admin consent on the app registration."
                    } else {
                        Add-Check $permission.Name Info "$($permission.Name) is not granted (optional: $($permission.Reason))."
                    }
                }

                if ($roles -contains 'DeviceManagementApps.ReadWrite.All') {
                    try {
                        Invoke-RestMethod -Method Get -Uri 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$top=1&$select=id' `
                            -Headers @{ Authorization = $Global:AuthenticationHeader.Authorization } -TimeoutSec 60 | Out-Null
                        Add-Check 'Intune' OK 'Intune apps can be read.'
                    } catch {
                        Add-Check 'Intune' Fail "Reading Intune apps failed: $((Get-YardstickSetupGraphError $_).Message)"
                    }
                }
            }
        }
    }
}

#endregion


#region GitHub API

# Recipes call the GitHub API on development-only instances too, so this is
# checked everywhere. A token is optional: without one, GitHub allows 60 calls an
# hour, about one full run.
Start-Area 'GitHub API'
$gitHubToken = $null
try {
    $gitHubToken = Get-YardstickGitHubToken
} catch {
    Add-Check 'GitHub token' Warn "Could not read the GitHub API token: $($_.Exception.Message)"
}
if (-not $gitHubToken) {
    if (-not ($results | Where-Object Check -eq 'GitHub token')) {
        Add-Check 'GitHub token' Info 'No GitHub API token is stored, so GitHub calls are anonymous and limited to 60 an hour. Run .\Set-YardstickCredential.ps1 -GitHub to store one.'
    }
} elseif ($Offline) {
    Add-Check 'GitHub token' OK 'A GitHub API token is stored. Checking it was skipped (-Offline).'
} else {
    try {
        Reset-YardstickGitHubSession
        $rateLimit = Get-YardstickGitHubRateLimit
        if ($rateLimit.Authenticated) {
            Add-Check 'GitHub token' OK "GitHub accepted the stored token: $($rateLimit.Remaining) of $($rateLimit.Limit) calls left this hour."
        } else {
            Add-Check 'GitHub token' Warn 'GitHub rejected the stored token, so runs fall back to anonymous calls. Store a new one with .\Set-YardstickCredential.ps1 -GitHub.'
        }
    } catch {
        Add-Check 'GitHub token' Warn "Could not check the GitHub API token: $($_.Exception.Message)"
    }
}

#endregion


#region Packaging tool

Start-Area 'Packaging'
$toolPath = Join-Path $prefs.Tools 'IntuneWinAppUtil.exe'
if (Test-Path -LiteralPath $toolPath -PathType Leaf) {
    try {
        Resolve-YardstickIntuneWinAppUtil -ToolsPath $prefs.Tools | Out-Null
        Add-Check 'IntuneWinAppUtil' OK "IntuneWinAppUtil.exe matches the pinned Microsoft release."
    } catch {
        Add-Check 'IntuneWinAppUtil' Fail "$($_.Exception.Message) Delete it and Yardstick will download a fresh copy."
    }
} else {
    Add-Check 'IntuneWinAppUtil' Info 'IntuneWinAppUtil.exe will be downloaded on the first run.'
}

#endregion


#region Email

Start-Area 'Email'
if ($developmentOnly) {
    Add-Check 'Email' Skip 'Development-only instance: no email reports are sent from this computer.'
} else {
    if ($prefs.emailNotificationEnabled -ne $true) {
        Add-Check 'Email' Info 'Email reports are turned off. Run .\Install-Yardstick.ps1 -Step Email to set them up.'
    } else {
        $method = Get-YardstickEmailDeliveryMethod -Preferences $prefs
        $recipients = @(@($prefs.emailRecipient) + @($prefs.adminEmailRecipient) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
        $to = if ($recipients.Count -gt 3) { "$(($recipients | Select-Object -First 3) -join '; ') and $($recipients.Count - 3) more" } else { $recipients -join '; ' }
        Add-Check 'Email' OK "Email reports are sent by $method to $to."
        if ($method -eq 'smtp') {
            $smtpTarget = Get-YardstickSmtpCredentialTarget -Preferences $prefs
            if (Get-YardstickSmtpCredential -Target $smtpTarget) {
                Add-Check 'SMTP credential' OK "SMTP sign-in '$smtpTarget' is stored."
            } else {
                Add-Check 'SMTP credential' Info "No SMTP sign-in is stored, so mail is relayed anonymously through $($prefs.smtpServer)."
            }
        } else {
            if (Test-OutlookAvailability) {
                Add-Check 'Outlook' OK 'Outlook is available.'
                if (Test-YardstickOutlookPromptExpected) {
                    Add-Check 'Outlook prompts' Warn 'Outlook will ask you to approve every email Yardstick sends. Run .\Install-Yardstick.ps1 -Step Email to turn the prompts off.'
                } else {
                    Add-Check 'Outlook prompts' OK 'Outlook sends without asking for approval.'
                }
            } else {
                Add-Check 'Outlook' Fail 'Email is set to use Outlook, but Outlook is not available. Install Outlook or switch to SMTP with .\Install-Yardstick.ps1 -Step Email.'
            }
        }
    }
}

#endregion


#region Scheduled task

Start-Area 'Scheduled task'
if ($developmentOnly) {
    Add-Check 'Scheduled task' Skip 'Development-only instance: Yardstick does not run on a schedule here.'
} else {
    $task = Get-ScheduledTask -TaskPath '\Yardstick\' -TaskName 'Yardstick' -ErrorAction SilentlyContinue
    if (-not $task) {
        Add-Check 'Scheduled task' Info 'No scheduled task. Run .\Install-Yardstick.ps1 -Step ScheduledTask to run Yardstick daily.'
    } else {
        $action = @($task.Actions)[0]
        $runAs = [string]$task.Principal.UserId
        $whenSignedIn = $task.Principal.LogonType -eq 'Interactive'
        $when = if ($whenSignedIn) { 'only while signed in' } else { 'whether or not signed in' }
        Add-Check 'Scheduled task' OK "Task \Yardstick\Yardstick runs as $runAs, $when ($($task.State))."
        $execute = if ($action) { ([string]$action.Execute).Trim('"') } else { '' }
        if ($execute -and -not (Test-Path -LiteralPath $execute)) {
            Add-Check 'Task program' Fail "The task starts $execute, which no longer exists (PowerShell was probably updated). Run .\Install-Yardstick.ps1 -Step ScheduledTask to replace the task."
        }
        if ($action -and "$($action.Arguments)" -notlike "*$(Join-Path $ProjectRoot 'Yardstick.ps1')*") {
            Add-Check 'Task target' Warn "The task runs '$($action.Arguments)', which is not this folder's Yardstick.ps1."
        }
        $shortRunAs = ($runAs -split '\\')[-1]
        if ($shortRunAs -and $shortRunAs -ne $env:USERNAME) {
            Add-Check 'Task account' Warn "The task runs as $runAs, not $env:USERNAME. Credentials are stored per user, so run this check (and the AppRegistration step) as $runAs."
        }
        if (-not $whenSignedIn) {
            if ($prefs.emailNotificationEnabled -eq $true -and (Get-YardstickEmailDeliveryMethod -Preferences $prefs) -eq 'outlook') {
                Add-Check 'Task email' Warn 'Outlook cannot send mail from a task that runs whether or not you are signed in. Use SMTP, or run .\Install-Yardstick.ps1 -Step ScheduledTask and choose to run only while signed in.'
            }
            if (Test-YardstickCredentialGuard) {
                Add-Check 'Task network' Warn 'Credential Guard is on, which can stop a task with a stored password from opening network shares ("The user name or password is incorrect"). If backups or the software dropbox fail in scheduled runs, run .\Install-Yardstick.ps1 -Step ScheduledTask and choose to run only while signed in.'
            }
        }

        $info = Get-ScheduledTaskInfo -TaskPath '\Yardstick\' -TaskName 'Yardstick'
        $next = if ($info.NextRunTime) { " Next run: $($info.NextRunTime.ToString('yyyy-MM-dd HH:mm'))." } else { '' }
        switch ($info.LastTaskResult) {
            0 { Add-Check 'Last run' OK "Last run on $($info.LastRunTime.ToString('yyyy-MM-dd HH:mm')) succeeded.$next" }
            267011 { Add-Check 'Last run' Info "The task has not run yet.$next" }
            267009 { Add-Check 'Last run' Info "The task is running now.$next" }
            default { Add-Check 'Last run' Warn ("Last run on $($info.LastRunTime.ToString('yyyy-MM-dd HH:mm')) returned 0x{0:X}. See the log in $($prefs.Temp).$next" -f $info.LastTaskResult) }
        }
    }
}

#endregion


#region Icons

Start-Area 'Icons'
try {
    $icons = Get-MissingIconCount -RecipeFolder @($prefs.Recipes) -IconsFolder $prefs.Icons
    if ($icons.Missing.Count -eq 0) {
        Add-Check 'Missing icons' OK "All $($icons.Recipes) recipe(s) have an icon."
    } else {
        $sample = ($icons.Missing | Select-Object -First 5 | ForEach-Object DisplayName) -join ', '
        $more = if ($icons.Missing.Count -gt 5) { ', ...' } else { '' }
        Add-Check 'Missing icons' Warn "$($icons.Missing.Count) recipe icon(s) missing ($sample$more). Apps upload without an icon; add them with .\Get-YardstickIcon.ps1."
    }
} catch {
    Add-Check 'Missing icons' Skip "Could not check icons: $($_.Exception.Message)"
}

#endregion


$failed = @($results | Where-Object Level -eq 'Fail').Count
$warned = @($results | Where-Object Level -eq 'Warn').Count
Write-Host ''
if ($failed) {
    Write-Host "$failed check(s) failed, $warned warning(s). Fix the failures before running Yardstick." -ForegroundColor Red
} elseif ($warned) {
    Write-Host "No failures, $warned warning(s)." -ForegroundColor Yellow
} else {
    Write-Host 'All checks passed.' -ForegroundColor Green
}

if ($PassThru) { $results }
exit ([int][bool]$failed)
