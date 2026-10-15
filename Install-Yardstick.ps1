<#
.SYNOPSIS
Sets up Yardstick on this machine.

.DESCRIPTION
Walks through everything a new installation needs, asking only for what has no
sensible default:

  Prerequisites    PowerShell modules, unblocking downloaded files, execution
                   policy and the IntuneWinAppUtil packaging tool.
  Preferences      Creates Local\preferences.yaml and the working folders, and asks
                   for the app owner and minimum Windows version.
  AppRegistration  Creates the Entra app registration Yardstick signs in with,
                   grants its Microsoft Graph permissions and stores the secret in
                   Windows Credential Manager. Or stores the details of an app you
                   created yourself (-UseExistingApp).
  IntuneDefaults   Scope tags and a test group for new app versions.
  Email            Optional email report after each run.
  ScheduledTask    Optional daily scheduled task.
  Verify           Runs Test-YardstickSetup.ps1.

Every step checks what is already in place, so the script is safe to run again -
for example after a git pull, or to change one setting with -Step.

.PARAMETER Step
Run only these steps. By default every step runs in order.

.PARAMETER NonInteractive
Never prompt. Values come from parameters, the existing preferences file, or the
built-in defaults; steps that need missing information are skipped.

.PARAMETER UseExistingApp
Store the details of an existing app registration instead of creating one.
Combine with -TenantId, -ClientId and -ClientSecret to avoid the prompts.

.PARAMETER AppDisplayName
Display name of the app registration to create or reuse. Default: Yardstick.

.PARAMETER SecretLifetimeMonths
Lifetime of the client secret that is created. Default: 12.

.PARAMETER IncludeOptionalPermissions
Also request Group.Read.All and Application.Read.All. Neither is needed for a
normal run; see Docs\Installation.md.

.PARAMETER UseDeviceCode
Sign in with a device code even when the Microsoft.Graph.Authentication module is
installed.

.PARAMETER TenantId
Tenant ID of an existing app registration (with -UseExistingApp).

.PARAMETER ClientId
Application (client) ID of an existing app registration (with -UseExistingApp).

.PARAMETER ClientSecret
Client secret of an existing app registration (with -UseExistingApp).

.PARAMETER ClientSecretExpiresOn
Expiry date of -ClientSecret, used for the expiry warning.

.PARAMETER Owner
Shown as the Owner of every app Yardstick creates in Intune.

.PARAMETER MinOSVersion
Default minimum Windows version for new apps, for example W10_22H2.

.PARAMETER ScopeTags
Intune scope tag names applied to every new app.

.PARAMETER TestGroup
Display name or object ID of an Entra group that gets every new app version as
Available.

.PARAMETER EmailMethod
none, outlook or smtp.

.PARAMETER EmailRecipient
Address(es) that receive the list of updated software.

.PARAMETER AdminEmailRecipient
Address(es) that receive the full report and alerts.

.PARAMETER EmailFrom
Sender address. Required for smtp.

.PARAMETER SmtpServer
SMTP server host name.

.PARAMETER SmtpPort
SMTP server port.

.PARAMETER SmtpUseSsl
Use TLS for SMTP.

.PARAMETER TaskCredential
Account the scheduled task runs as. Must be the account running this script,
because the stored credentials belong to it.

.PARAMETER TaskTime
Daily start time for the scheduled task, as HH:mm. Default: 03:00.

.PARAMETER TaskGroup
Recipe group (from RecipeGroups.yaml) for the scheduled task to run instead of
every recipe.

.EXAMPLE
.\Install-Yardstick.ps1

Runs the whole setup wizard.

.EXAMPLE
.\Install-Yardstick.ps1 -Step AppRegistration

Creates the app registration again, for example to replace an expiring secret.

.EXAMPLE
.\Install-Yardstick.ps1 -Step AppRegistration -UseExistingApp

Stores the details of an app registration that was created in the Entra portal.

.EXAMPLE
$secret = Read-Host -AsSecureString
.\Install-Yardstick.ps1 -NonInteractive -UseExistingApp -TenantId $tenant -ClientId $app -ClientSecret $secret -Owner 'IT Services'

Scripted install for an existing app registration.
#>

# Deliberately no "#Requires -Version 7.0": Windows PowerShell 5.1 has to be able
# to parse this file so it can explain how to get PowerShell 7. Keep the syntax
# 5.1-compatible (no ternaries, ?? or &&) and the file ASCII.
[CmdletBinding()]
param(
    [ValidateSet('Prerequisites', 'Preferences', 'AppRegistration', 'IntuneDefaults', 'Email', 'ScheduledTask', 'Verify')]
    [string[]]$Step,
    [switch]$NonInteractive,

    [switch]$UseExistingApp,
    [string]$AppDisplayName = 'Yardstick',
    [ValidateRange(1, 24)]
    [int]$SecretLifetimeMonths = 12,
    [switch]$IncludeOptionalPermissions,
    [switch]$UseDeviceCode,
    [string]$TenantId,
    [string]$ClientId,
    [SecureString]$ClientSecret,
    [datetime]$ClientSecretExpiresOn,

    [string]$Owner,
    [string]$MinOSVersion,
    [string[]]$ScopeTags,
    [string]$TestGroup,

    [ValidateSet('none', 'outlook', 'smtp')]
    [string]$EmailMethod,
    [string[]]$EmailRecipient,
    [string[]]$AdminEmailRecipient,
    [string]$EmailFrom,
    [string]$SmtpServer,
    [int]$SmtpPort,
    [switch]$SmtpUseSsl,

    [pscredential]$TaskCredential,
    [string]$TaskTime = '03:00',
    [string]$TaskGroup
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host ''
    Write-Host "Yardstick needs PowerShell 7 or later. This is Windows PowerShell $($PSVersionTable.PSVersion)." -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'Install PowerShell 7 with:'
    Write-Host '    winget install --id Microsoft.PowerShell --source winget'
    Write-Host '(or download it from https://aka.ms/powershell), then open "PowerShell 7"'
    Write-Host 'and run .\Install-Yardstick.ps1 again.'
    Write-Host ''
    exit 1
}
if (-not $IsWindows) {
    Write-Host 'Yardstick runs on Windows only: it stores credentials in Windows Credential Manager and packages apps with IntuneWinAppUtil.exe.' -ForegroundColor Yellow
    exit 1
}

$ErrorActionPreference = 'Stop'
$Root = (Get-Item -LiteralPath $PSScriptRoot).FullName
# The step functions need to know which parameters were passed to the script;
# inside a function $PSBoundParameters is that function's own.
$ScriptBound = $PSBoundParameters

$Global:LogLocation = $Root
$Global:LogFile = 'YInstall.log'

# Reload the module so a session that imported it before an update runs the
# current code. Not when Install-Yardstick in the module runs this script: that
# would unload the module out from under the command.
$calledFromModule = [bool](Get-PSCallStack | Where-Object { $_.InvocationInfo.MyCommand.ModuleName -eq 'Yardstick' })
Import-Module (Join-Path $Root 'Yardstick.psd1') -Force:(-not $calledFromModule)

Write-Log -Init
Set-YardstickSetupMode -NonInteractive:$NonInteractive

$AllSteps = @('Prerequisites', 'Preferences', 'AppRegistration', 'IntuneDefaults', 'Email', 'ScheduledTask', 'Verify')
$Selected = if ($Step) { @($AllSteps | Where-Object { $Step -contains $_ }) } else { $AllSteps }
$Summary = [System.Collections.Generic.List[object]]::new()


#region Helpers

function Get-SetupPreferences {
    Import-YardstickPreferences -ProjectRoot $Root -AllowMissing
}

function Get-SetupPreferencesPath {
    # Every step that writes a setting needs the file to exist first.
    (New-YardstickPreferencesFile -ProjectRoot $Root).Path
}

function Split-AddressList {
    param([string[]]$Value)
    @($Value | ForEach-Object { $_ -split '[;,]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function ConvertTo-PreferenceAddress {
    # One address stays a plain string, several become a list; Yardstick accepts both.
    param([string[]]$Address)
    $list = @(Split-AddressList $Address)
    if ($list.Count -eq 0) { return $null }
    if ($list.Count -eq 1) { return $list[0] }
    , $list
}

function Get-StoredIntuneCredential {
    $target = Get-YardstickCredentialTarget -Preferences (Get-SetupPreferences)
    Get-YardstickIntuneCredential -Target $target
}

function Test-StoredIntuneCredential {
    <#
    Returns $null when the credential works and carries every required permission,
    otherwise a short description of the problem.
    #>
    param([Parameter(Mandatory)]$Credential)
    try {
        $roles = @(Get-YardstickAppTokenRoles -TenantId $Credential.TenantID -ClientId $Credential.ClientID -ClientSecret $Credential.ClientSecret)
    } catch {
        return "it could not sign in: $($_.Exception.Message)"
    }
    $missing = @((Get-YardstickRequiredGraphPermissions -RequiredOnly).Name | Where-Object { -not (Test-YardstickGraphRoleGranted -Roles $roles -Name $_) })
    if ($missing.Count -gt 0) { return "it is missing these Graph permissions: $($missing -join ', ')" }
    $null
}

function Write-PermissionTable {
    param([switch]$IncludeOptional)
    $permissions = if ($IncludeOptional) { Get-YardstickRequiredGraphPermissions } else { Get-YardstickRequiredGraphPermissions -RequiredOnly }
    foreach ($permission in $permissions) {
        Write-Host ("    {0,-40} {1}" -f $permission.Name, $permission.Reason)
    }
}

function Write-CredentialOwnerReminder {
    Write-Host ''
    Write-Host "The credential is stored for $env:USERDOMAIN\$env:USERNAME only. Run Yardstick (and its" -ForegroundColor DarkGray
    Write-Host 'scheduled task) as this account, or repeat this step as the account that will.' -ForegroundColor DarkGray
}

function Add-Summary {
    param([string]$Name, [string]$Result, [string]$Detail)
    $Summary.Add([pscustomobject]@{ Step = $Name; Result = $Result; Detail = $Detail })
}

#endregion


#region Steps

function Invoke-PrerequisitesStep {
    Write-YardstickSetupStep 'Prerequisites'

    Write-YardstickSetupStatus "PowerShell $($PSVersionTable.PSVersion)" -Level OK

    if (Install-YardstickModule -Name 'powershell-yaml') {
        Write-YardstickSetupStatus 'powershell-yaml module is installed.' -Level OK
    } else {
        throw "The powershell-yaml module is required. Install it with: Install-Module powershell-yaml -Scope CurrentUser"
    }

    $optional = @('Selenium', 'TUN.CredentialManager') | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
    if ($optional) {
        Write-YardstickSetupStatus "Optional modules not installed: $($optional -join ', '). They are only needed for Adobe recipes that sign in through a browser." -Level Info
        if (Read-YardstickSetupYesNo -Prompt 'Install them now?' -Default $false) {
            foreach ($name in $optional) {
                if (Install-YardstickModule -Name $name) { Write-YardstickSetupStatus "$name installed." -Level OK }
            }
        }
    }

    $unblocked = Unblock-YardstickFiles -ProjectRoot $Root
    if ($unblocked -gt 0) {
        Write-YardstickSetupStatus "Unblocked $unblocked downloaded file(s)." -Level OK
    }

    if (Test-YardstickExecutionPolicy) {
        Write-YardstickSetupStatus "Execution policy ($(Get-ExecutionPolicy)) allows Yardstick's scripts." -Level OK
    } else {
        Write-YardstickSetupStatus "Execution policy '$(Get-ExecutionPolicy)' blocks Yardstick's unsigned scripts." -Level Warn
        if (-not $NonInteractive -and (Read-YardstickSetupYesNo -Prompt 'Set it to RemoteSigned for your account?' -Default $true)) {
            try {
                Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
                Write-YardstickSetupStatus 'Execution policy set to RemoteSigned for the current user.' -Level OK
            } catch {
                Write-YardstickSetupStatus "Could not change it (it may be set by Group Policy): $($_.Exception.Message)" -Level Warn
            }
        }
    }

    $prefs = Get-SetupPreferences
    try {
        $tool = Resolve-YardstickIntuneWinAppUtil -ToolsPath $prefs.Tools
        Write-YardstickSetupStatus "IntuneWinAppUtil.exe is ready ($tool)." -Level OK
    } catch {
        Write-YardstickSetupStatus "Could not get IntuneWinAppUtil.exe: $($_.Exception.Message) Yardstick will try again on its first run." -Level Warn
    }

    $check = Test-Prerequisites -ToolsPath $prefs.Tools
    foreach ($e in $check.Errors) { Write-YardstickSetupStatus $e -Level Fail }
    foreach ($w in $check.Warnings | Where-Object { $_ -notmatch 'Optional PowerShell module' }) { Write-YardstickSetupStatus $w -Level Warn }
    if (-not $check.IsValid) { throw 'Some required prerequisites are missing (see above).' }
    'Done'
}


function Invoke-PreferencesStep {
    Write-YardstickSetupStep 'Preferences'

    $file = New-YardstickPreferencesFile -ProjectRoot $Root
    if ($file.Created) {
        Write-YardstickSetupStatus "Created $($file.Path) from Config\Preferences.example.yaml." -Level OK
    } else {
        Write-YardstickSetupStatus "Using existing $($file.Path)." -Level OK
    }
    $path = $file.Path
    $prefs = Get-SetupPreferences

    foreach ($folder in @(Initialize-YardstickFolders -Preferences $prefs)) {
        Write-YardstickSetupStatus "Created folder $folder" -Level OK
    }

    Write-Host ''
    Write-Host 'Owner is shown on every app Yardstick creates in Intune - usually your team name.' -ForegroundColor DarkGray
    $newOwner = if ($Owner) { $Owner } else { Read-YardstickSetupValue -Prompt 'App owner' -Default ([string]$prefs.defaultOwner) }
    if ($newOwner -ne [string]$prefs.defaultOwner) { Set-YardstickPreferenceValue -Path $path -Key 'defaultOwner' -Value $newOwner }

    $osChoices = @(Get-YardstickPreferenceAllowedValues -Key 'defaultMinOSVersion')
    $newOs = $MinOSVersion
    if (-not $newOs) {
        $newOs = @(Read-YardstickSetupChoice -Prompt 'Minimum Windows version for new apps:' -Choices $osChoices -Default @([string]$prefs.defaultMinOSVersion))[0]
    }
    if ($osChoices -notcontains $newOs) { throw "'$newOs' is not a valid minimum OS version. Use one of: $($osChoices -join ', ')" }
    if ($newOs -ne [string]$prefs.defaultMinOSVersion) { Set-YardstickPreferenceValue -Path $path -Key 'defaultMinOSVersion' -Value $newOs }

    if (-not $NonInteractive) {
        Write-Host ''
        Write-Host 'Optional folders: a backup location for every uploaded package, and a dropbox for' -ForegroundColor DarkGray
        Write-Host 'recipes whose installers have to be downloaded by hand (manualDownload: true).' -ForegroundColor DarkGray
        if (Read-YardstickSetupYesNo -Prompt 'Set up optional folders now?' -Default $false) {
            foreach ($key in @('Backup', 'SoftwareDropbox', 'SoftwareArchive')) {
                $value = Read-YardstickSetupValue -Prompt "$key folder (blank to leave off)" -Default ([string]$prefs[$key])
                if ($value -ne [string]$prefs[$key]) { Set-YardstickPreferenceValue -Path $path -Key $key -Value $value }
            }
        }
    }

    $check = Test-YardstickPreferences -Preferences (Get-SetupPreferences)
    foreach ($w in $check.Warnings) { Write-YardstickSetupStatus $w -Level Warn }
    foreach ($e in $check.Errors) { Write-YardstickSetupStatus $e -Level Fail }
    if (-not $check.IsValid) { throw "Fix the errors above in $path, then run .\Install-Yardstick.ps1 -Step Preferences again." }
    Write-YardstickSetupStatus 'Preferences are valid.' -Level OK
    'Done'
}


function Invoke-AppRegistrationStep {
    Write-YardstickSetupStep 'Intune app registration'

    $prefs = Get-SetupPreferences
    $target = Get-YardstickCredentialTarget -Preferences $prefs

    $haveExplicit = $TenantId -and $ClientId -and $ClientSecret
    $existing = Get-YardstickIntuneCredential -Target $target
    if ($existing -and -not $haveExplicit) {
        $problem = Test-StoredIntuneCredential -Credential $existing
        if (-not $problem) {
            Write-YardstickSetupStatus "Credentials for app $($existing.ClientID) are stored and working." -Level OK
            if (-not (Read-YardstickSetupYesNo -Prompt 'Replace them (for example to renew the secret)?' -Default $false)) {
                return 'Already configured'
            }
        } else {
            Write-YardstickSetupStatus "Stored credentials for app $($existing.ClientID) need attention: $problem" -Level Warn
        }
    }

    $createApp = $false
    if ($haveExplicit -or $UseExistingApp) {
        $createApp = $false
    } elseif ($NonInteractive) {
        Write-YardstickSetupStatus 'Skipped: pass -UseExistingApp -TenantId -ClientId -ClientSecret, or run this step interactively.' -Level Skip
        return 'Skipped'
    } else {
        Write-Host ''
        Write-Host 'Yardstick signs in to Intune with its own Entra app registration, which needs these' -ForegroundColor DarkGray
        Write-Host 'Microsoft Graph application permissions:' -ForegroundColor DarkGray
        Write-PermissionTable -IncludeOptional:$IncludeOptionalPermissions
        Write-Host ''
        $choice = Read-YardstickSetupChoice -Prompt 'How do you want to set it up?' -Choices @(
            'Create the app registration for me (sign in as a Global Administrator)'
            'I already have an app registration (enter its tenant ID, client ID and secret)'
        ) -Default @('Create the app registration for me (sign in as a Global Administrator)')
        $createApp = ($choice -like 'Create*')
    }

    if (-not $createApp) {
        if ($haveExplicit) {
            $expires = $null
            if ($ScriptBound.ContainsKey('ClientSecretExpiresOn')) { $expires = $ClientSecretExpiresOn }
            Set-YardstickIntuneCredential -Target $target -TenantID $TenantId -ClientID $ClientId -ClientSecret $ClientSecret -SecretExpiresOn $expires
        } else {
            Register-YardstickIntuneCredential -Target $target -TenantID $TenantId -ClientID $ClientId | Out-Null
        }
        $stored = Get-YardstickIntuneCredential -Target $target
        $problem = Test-StoredIntuneCredential -Credential $stored
        Write-CredentialOwnerReminder
        if ($problem) {
            Write-YardstickSetupStatus "Credentials stored, but $problem" -Level Warn
            Write-Host '  Grant the permissions listed above (with admin consent) to the app, then run .\Test-YardstickSetup.ps1.' -ForegroundColor DarkGray
            return 'Stored with warnings'
        }
        Write-YardstickSetupStatus 'Credentials stored and working.' -Level OK
        return 'Done'
    }

    $displayName = Read-YardstickSetupValue -Prompt 'App registration name' -Default $AppDisplayName -Required
    $months = Read-YardstickSetupValue -Prompt 'Client secret lifetime in months (1-24)' -Default ([string]$SecretLifetimeMonths) -Required
    $monthsValue = 0
    if (-not [int]::TryParse($months, [ref]$monthsValue) -or $monthsValue -lt 1 -or $monthsValue -gt 24) {
        throw "'$months' is not a number of months between 1 and 24."
    }

    $session = Connect-YardstickSetupGraph -TenantId $TenantId -UseDeviceCode:$UseDeviceCode
    Write-YardstickSetupStatus "Signed in as $($session.Account) (tenant $($session.TenantId))." -Level OK

    try {
        $result = New-YardstickAppRegistration -DisplayName $displayName -SecretMonths $monthsValue -IncludeOptionalPermissions:$IncludeOptionalPermissions -CredentialTarget $target
    } catch {
        $err = Get-YardstickSetupGraphError $_
        if ($err.StatusCode -eq 403) {
            Write-YardstickSetupStatus "Your account is not allowed to do this: $($err.Message)" -Level Fail
            Write-Host '  Ask an administrator to create the app registration (see Docs\Installation.md), then run:' -ForegroundColor DarkGray
            Write-Host '    .\Install-Yardstick.ps1 -Step AppRegistration -UseExistingApp' -ForegroundColor DarkGray
        }
        throw
    }

    if ($result.Consent.Granted) { Write-YardstickSetupStatus "Granted: $($result.Consent.Granted -join ', ')" -Level OK }
    if ($result.Consent.AlreadyGranted) { Write-YardstickSetupStatus "Already granted: $($result.Consent.AlreadyGranted -join ', ')" -Level OK }
    Write-YardstickSetupStatus "Client secret stored in Credential Manager; it expires $($result.SecretExpiresOn.ToLocalTime().ToString('yyyy-MM-dd'))." -Level OK
    Write-CredentialOwnerReminder

    if ($result.Consent.Failed.Count -gt 0) {
        Write-YardstickSetupStatus "Could not grant admin consent for: $($result.Consent.Failed -join ', '). Your account needs Global Administrator or Privileged Role Administrator." -Level Warn
        Write-Host '  Ask an administrator to open this link and accept:' -ForegroundColor DarkGray
        Write-Host "    $($result.ConsentUrl)" -ForegroundColor White
        Write-Host '  then run .\Test-YardstickSetup.ps1 to confirm.' -ForegroundColor DarkGray
        return 'Waiting for admin consent'
    }
    if ($result.Validation -and -not $result.Validation.Success) {
        $detail = if ($result.Validation.Missing) { "missing $($result.Validation.Missing -join ', ')" } else { $result.Validation.Error }
        Write-YardstickSetupStatus "The new credentials are not working yet ($detail). Changes in Entra can take a few minutes; run .\Test-YardstickSetup.ps1 shortly." -Level Warn
        return 'Created; not yet verified'
    }
    Write-YardstickSetupStatus 'The app signs in and has every required permission.' -Level OK
    'Done'
}


function Invoke-IntuneDefaultsStep {
    Write-YardstickSetupStep 'Intune defaults'

    if (-not (Get-YardstickSetupSession)) {
        $credential = Get-StoredIntuneCredential
        if (-not $credential) {
            Write-YardstickSetupStatus 'Skipped: no Intune credentials are stored yet. Run the AppRegistration step first.' -Level Skip
            return 'Skipped'
        }
        Use-YardstickSetupAppCredential -Credential $credential | Out-Null
    }

    $path = Get-SetupPreferencesPath
    $prefs = Get-SetupPreferences

    # Scope tags
    $currentTags = @($prefs.defaultScopeTags | Where-Object { $_ })
    if ($ScopeTags) {
        Set-YardstickPreferenceValue -Path $path -Key 'defaultScopeTags' -Value @($ScopeTags)
        Write-YardstickSetupStatus "Scope tags set to: $($ScopeTags -join ', ')" -Level OK
    } elseif (-not $NonInteractive) {
        try {
            $tags = @(Get-YardstickScopeTagNames)
            if ($tags.Count -le 1) {
                Write-YardstickSetupStatus 'This tenant only has the Default scope tag; new apps will use it.' -Level OK
            } else {
                $default = if ($currentTags.Count) { $currentTags } else { @('Default') }
                $picked = @(Read-YardstickSetupChoice -Prompt 'Scope tags for new apps:' -Choices $tags -Default $default -Multiple)
                # Only "Default" is the same as no setting, which keeps the file tidy.
                $value = if ($picked.Count -eq 1 -and $picked[0] -eq 'Default') { @() } else { $picked }
                Set-YardstickPreferenceValue -Path $path -Key 'defaultScopeTags' -Value $value
                Write-YardstickSetupStatus "Scope tags: $($picked -join ', ')" -Level OK
            }
        } catch {
            Write-YardstickSetupStatus "Could not list scope tags: $($_.Exception.Message)" -Level Warn
        }
    }

    # Test group
    $currentGroups = @($prefs.defaultDeploymentGroups | Where-Object { $_ })
    $groupName = $TestGroup
    if (-not $groupName -and -not $NonInteractive) {
        Write-Host ''
        Write-Host 'A test group receives every new app version as Available, so you can try updates' -ForegroundColor DarkGray
        Write-Host 'before they reach anyone else. Usually an IT staff device or user group.' -ForegroundColor DarkGray
        if ($currentGroups.Count -gt 0) {
            Write-YardstickSetupStatus "Currently: $($currentGroups -join ', ')" -Level Info
            if (Read-YardstickSetupYesNo -Prompt 'Keep it?' -Default $true) { return 'Done' }
        }
        $groupName = Read-YardstickSetupValue -Prompt 'Test group name or object ID (blank for none)'
    }
    if (-not $groupName) { return 'Done' }

    $groupId = $null
    while (-not $groupId) {
        try {
            $found = @(Find-YardstickGroup -Name $groupName)
            if ($found.Count -eq 1) {
                $groupId = [string]$found[0].id
                Write-YardstickSetupStatus "Test group: $($found[0].displayName) ($groupId)" -Level OK
            } elseif ($found.Count -gt 1) {
                $labels = @($found | ForEach-Object { "$($_.displayName) ($($_.id))" })
                $pick = @(Read-YardstickSetupChoice -Prompt "Several groups are named '$groupName':" -Choices $labels)[0]
                if (-not $pick) { throw "Several groups are named '$groupName'; pass its object ID instead." }
                $groupId = [regex]::Match($pick, '\(([0-9a-fA-F-]{36})\)$').Groups[1].Value
            } else {
                Write-YardstickSetupStatus "No group named '$groupName' was found." -Level Warn
                if ($NonInteractive) { return 'Group not found' }
                $groupName = Read-YardstickSetupValue -Prompt 'Test group name or object ID (blank to skip)'
                if (-not $groupName) { return 'Done' }
            }
        } catch {
            $err = Get-YardstickSetupGraphError $_
            if ($err.StatusCode -ne 403) { throw }
            # App-only sessions lack Group.Read.All by design; take the ID on trust.
            $guid = [guid]::Empty
            if ([guid]::TryParse($groupName, [ref]$guid)) {
                $groupId = $guid.ToString()
                Write-YardstickSetupStatus "Could not look the group up (no Group.Read.All); using object ID $groupId as entered." -Level Warn
            } else {
                Write-YardstickSetupStatus 'Groups cannot be searched by name with the app credentials (no Group.Read.All).' -Level Warn
                if ($NonInteractive) { return 'Pass -TestGroup as an object ID' }
                $groupName = Read-YardstickSetupValue -Prompt "Enter the group's object ID from the Entra portal (blank to skip)"
                if (-not $groupName) { return 'Done' }
            }
        }
    }
    Set-YardstickPreferenceValue -Path $path -Key 'defaultDeploymentGroups' -Value @($groupId)
    'Done'
}


function Invoke-EmailStep {
    Write-YardstickSetupStep 'Email notifications'

    $path = Get-SetupPreferencesPath
    $prefs = Get-SetupPreferences

    $method = $EmailMethod
    if (-not $method) {
        if ($NonInteractive) {
            Write-YardstickSetupStatus 'Skipped: pass -EmailMethod to configure email.' -Level Skip
            return 'Skipped'
        }
        $current = if ($prefs.emailNotificationEnabled) { [string]$prefs.emailDeliveryMethod } else { 'none' }
        Write-Host 'Yardstick can email a report after each run listing what was updated.' -ForegroundColor DarkGray
        Write-Host 'Use SMTP for scheduled runs: Outlook needs a signed-in desktop session.' -ForegroundColor DarkGray
        $labels = [ordered]@{
            'smtp'    = 'SMTP server (recommended for scheduled runs)'
            'outlook' = 'Outlook on this computer'
            'none'    = 'No email'
        }
        $picked = @(Read-YardstickSetupChoice -Prompt 'How should reports be sent?' -Choices @($labels.Values) -Default @($labels[$current]))[0]
        $method = @($labels.Keys | Where-Object { $labels[$_] -eq $picked })[0]
    }

    if ($method -eq 'none') {
        Set-YardstickPreferenceValue -Path $path -Key 'emailNotificationEnabled' -Value $false
        Write-YardstickSetupStatus 'Email reports are off.' -Level OK
        return 'Off'
    }

    $to = $EmailRecipient
    if (-not $to) {
        $to = Read-YardstickSetupValue -Prompt 'Send the list of updated software to (separate several with ;)' -Default ((Split-AddressList $prefs.emailRecipient) -join '; ') -Required
    }
    $admin = $AdminEmailRecipient
    if (-not $admin -and -not $NonInteractive) {
        Write-Host 'The admin report adds failures and alerts such as secret expiry. Leave blank to send' -ForegroundColor DarkGray
        Write-Host 'one full report to the address(es) above instead.' -ForegroundColor DarkGray
        $admin = Read-YardstickSetupValue -Prompt 'Send the full admin report to (optional)' -Default ((Split-AddressList $prefs.adminEmailRecipient) -join '; ')
    }

    Set-YardstickPreferenceValue -Path $path -Key 'emailRecipient' -Value (ConvertTo-PreferenceAddress $to)
    Set-YardstickPreferenceValue -Path $path -Key 'adminEmailRecipient' -Value (ConvertTo-PreferenceAddress $admin)

    if ($method -eq 'smtp') {
        $server = $SmtpServer
        if (-not $server) { $server = Read-YardstickSetupValue -Prompt 'SMTP server' -Default ([string]$prefs.smtpServer) -Required }
        $port = $SmtpPort
        if (-not $port) { $port = [int](Read-YardstickSetupValue -Prompt 'SMTP port' -Default ([string]$prefs.smtpPort) -Required) }
        $ssl = [bool]$SmtpUseSsl
        if (-not $ScriptBound.ContainsKey('SmtpUseSsl')) {
            $ssl = Read-YardstickSetupYesNo -Prompt 'Use TLS (STARTTLS)?' -Default ([bool]$prefs.smtpUseSsl -or $port -eq 587)
        }
        $from = $EmailFrom
        if (-not $from) { $from = Read-YardstickSetupValue -Prompt 'Send from address' -Default ([string]$prefs.emailSendFromAddress) -Required }

        Set-YardstickPreferenceValue -Path $path -Key 'smtpServer' -Value $server
        Set-YardstickPreferenceValue -Path $path -Key 'smtpPort' -Value ([int]$port)
        Set-YardstickPreferenceValue -Path $path -Key 'smtpUseSsl' -Value $ssl
        Set-YardstickPreferenceValue -Path $path -Key 'emailSendFromAddress' -Value $from

        $smtpTarget = Get-YardstickSmtpCredentialTarget -Preferences (Get-SetupPreferences)
        $hasSmtpCredential = [bool](Get-YardstickSmtpCredential -Target $smtpTarget)
        if (-not $NonInteractive) {
            $prompt = if ($hasSmtpCredential) { 'An SMTP sign-in is stored. Replace it?' } else { 'Does the server need a user name and password?' }
            if (Read-YardstickSetupYesNo -Prompt $prompt -Default $false) {
                Register-YardstickSmtpCredential -Target $smtpTarget | Out-Null
            }
        }
    } else {
        $from = $EmailFrom
        if (-not $from -and -not $NonInteractive) {
            $from = Read-YardstickSetupValue -Prompt 'Send on behalf of (optional; blank sends as your Outlook account)' -Default ([string]$prefs.emailSendFromAddress)
        }
        Set-YardstickPreferenceValue -Path $path -Key 'emailSendFromAddress' -Value $from
        if (-not (Test-OutlookAvailability)) {
            Write-YardstickSetupStatus 'Outlook is not available on this computer, so reports cannot be sent until it is.' -Level Warn
        }
    }

    Set-YardstickPreferenceValue -Path $path -Key 'emailDeliveryMethod' -Value $method
    Set-YardstickPreferenceValue -Path $path -Key 'emailNotificationEnabled' -Value $true
    Write-YardstickSetupStatus "Email reports will be sent by $method." -Level OK

    if (-not $NonInteractive -and (Read-YardstickSetupYesNo -Prompt 'Send a sample report now to check it works?' -Default $true)) {
        try {
            & (Join-Path $Root 'Test-EmailNotification.ps1') -Send | Out-Null
        } catch {
            Write-YardstickSetupStatus "The sample report could not be sent: $($_.Exception.Message)" -Level Warn
            return 'Saved; test email failed'
        } finally {
            # The email test points the shared log at its own file.
            $Global:LogLocation = $Root
            $Global:LogFile = 'YInstall.log'
        }
    }
    'Done'
}


function Invoke-ScheduledTaskStep {
    Write-YardstickSetupStep 'Scheduled task'

    $existingTask = Get-ScheduledTask -TaskPath '\Yardstick\' -TaskName 'Yardstick' -ErrorAction SilentlyContinue
    if ($existingTask) {
        $info = Get-ScheduledTaskInfo -InputObject $existingTask -ErrorAction SilentlyContinue
        Write-YardstickSetupStatus "A Yardstick task already exists (runs as $($existingTask.Principal.UserId); last result $($info.LastTaskResult))." -Level OK
        if (-not (Read-YardstickSetupYesNo -Prompt 'Replace it?' -Default $false)) { return 'Already configured' }
    } elseif ($NonInteractive -and -not $TaskCredential) {
        Write-YardstickSetupStatus 'Skipped: pass -TaskCredential to create the scheduled task.' -Level Skip
        return 'Skipped'
    } elseif (-not $NonInteractive) {
        if (-not (Read-YardstickSetupYesNo -Prompt 'Run Yardstick automatically every day?' -Default $true)) { return 'Skipped' }
    }

    $time = Read-YardstickSetupValue -Prompt 'Daily start time (HH:mm)' -Default $TaskTime -Required
    $at = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($time, 'H:mm', [Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$at)) {
        throw "'$time' is not a time in HH:mm format."
    }

    $group = $TaskGroup
    if (-not $group -and -not $NonInteractive) {
        $groupsFile = Join-Path $Root 'RecipeGroups.yaml'
        $groups = @()
        if (Test-Path -LiteralPath $groupsFile) {
            $parsed = Get-Content -LiteralPath $groupsFile -Raw | ConvertFrom-Yaml
            if ($parsed -is [hashtable]) { $groups = @($parsed.Keys | Sort-Object) }
        }
        if ($groups.Count -gt 0) {
            $all = 'All recipes'
            $picked = @(Read-YardstickSetupChoice -Prompt 'Which recipes should the task run?' -Choices (@($all) + $groups) -Default @($all))[0]
            if ($picked -ne $all) { $group = $picked }
        }
    }

    $prefs = Get-SetupPreferences
    if ($prefs.emailNotificationEnabled -and $prefs.emailDeliveryMethod -eq 'outlook') {
        Write-YardstickSetupStatus 'Email is set to Outlook, which usually fails when nobody is signed in. Consider SMTP (-Step Email).' -Level Warn
    }
    if ($Root.StartsWith('\\')) {
        Write-YardstickSetupStatus "Yardstick is on a network share ($Root). The task account needs access to it; mapped drive letters are not available to scheduled tasks." -Level Warn
    }

    $credential = $TaskCredential
    if (-not $credential) {
        Write-Host ''
        Write-Host 'The task runs whether or not you are signed in, which needs your Windows password.' -ForegroundColor DarkGray
        Write-Host 'It is stored by Task Scheduler, not by Yardstick.' -ForegroundColor DarkGray
        $credential = Get-Credential -UserName "$env:USERDOMAIN\$env:USERNAME" -Message 'Account the Yardstick task runs as'
        if (-not $credential) { return 'Skipped' }
    }
    $taskUser = $credential.UserName
    if ($taskUser -notmatch [regex]::Escape($env:USERNAME) + '$') {
        Write-YardstickSetupStatus "The task will run as $taskUser, but the Intune credentials are stored for $env:USERNAME. Sign in as $taskUser and run .\Install-Yardstick.ps1 -Step AppRegistration as well." -Level Warn
    }

    try {
        Register-YardstickScheduledTask -ProjectRoot $Root -At $at -Credential $credential -Group $group | Out-Null
    } catch {
        if ($_.Exception.Message -match 'Access is denied|0x80070005') {
            throw 'Windows refused to create the task. Run PowerShell as administrator and run .\Install-Yardstick.ps1 -Step ScheduledTask again.'
        }
        throw
    }
    $what = if ($group) { "recipe group '$group'" } else { 'all recipes' }
    Write-YardstickSetupStatus "Scheduled task \Yardstick\Yardstick runs $what daily at $($at.ToString('HH:mm')) as $taskUser." -Level OK
    'Done'
}


function Invoke-VerifyStep {
    Write-YardstickSetupStep 'Verify'
    $healthCheck = Join-Path $Root 'Test-YardstickSetup.ps1'
    if (-not (Test-Path -LiteralPath $healthCheck)) {
        Write-YardstickSetupStatus 'Test-YardstickSetup.ps1 is missing.' -Level Warn
        return 'Skipped'
    }
    & $healthCheck
    if ($LASTEXITCODE -ne 0) { return 'Problems found' }
    'Done'
}

#endregion


Write-Host ''
Write-Host 'Yardstick setup' -ForegroundColor Cyan
Write-Host "Folder: $Root" -ForegroundColor DarkGray

$stopped = $false
foreach ($name in $Selected) {
    if ($stopped) { Add-Summary $name 'Not run' ''; continue }
    try {
        $result = & "Invoke-${name}Step"
        Add-Summary $name ([string]@($result)[-1]) ''
    } catch {
        Write-YardstickSetupStatus $_.Exception.Message -Level Fail
        Write-Log "Install step $name failed: $($_.Exception.Message)"
        Add-Summary $name 'Failed' $_.Exception.Message
        # Nothing later works without the YAML module or a valid preferences file.
        if ($name -in @('Prerequisites', 'Preferences')) { $stopped = $true }
    }
}

Write-YardstickSetupStep 'Summary'
foreach ($row in $Summary) {
    $level = switch -Wildcard ($row.Result) {
        'Done' { 'OK' }
        'Already configured' { 'OK' }
        'Off' { 'OK' }
        'Skipped' { 'Skip' }
        'Not run' { 'Skip' }
        'Failed' { 'Fail' }
        default { 'Warn' }
    }
    Write-YardstickSetupStatus ("{0,-16} {1}" -f $row.Step, $row.Result) -Level $level
}

$failed = @($Summary | Where-Object { $_.Result -eq 'Failed' })
Write-Host ''
if ($failed.Count -gt 0) {
    Write-Host "Fix the problems above and run .\Install-Yardstick.ps1 -Step $($failed[0].Step) to retry." -ForegroundColor Yellow
    exit 1
}
if (-not $Step) {
    Write-Host 'Next, try a first run with a single recipe:' -ForegroundColor Cyan
    Write-Host '    .\Yardstick.ps1 -ApplicationId 7zip -NoEmail'
    Write-Host ''
    Write-Host 'Recipes without an icon upload without one; .\Get-YardstickIcon.ps1 helps find icons.' -ForegroundColor DarkGray
    Write-Host 'Run .\Test-YardstickSetup.ps1 at any time to re-check this installation.' -ForegroundColor DarkGray
}
exit 0
