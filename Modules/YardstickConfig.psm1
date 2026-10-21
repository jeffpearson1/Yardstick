# Yardstick preferences: one loader shared by every entry script.
#
# A preferences file used to have to spell out every folder as an absolute path and
# every default* key, because the runner had no fallback for most of them - a missing
# defaultInstallExperience only surfaced as a parameter-binding error at upload time.
# This module supplies a built-in default for each of those keys, resolves relative
# folder paths against the Yardstick root, and validates the result up front, so a
# preferences file only has to carry what is genuinely site-specific.

. (Join-Path (Split-Path -Parent $PSScriptRoot) 'ProjectLayout.ps1')

# Folder keys that always have a value. Each maps to a property of the layout
# manifest (Yardstick.Project.psd1) or, for folders the manifest does not name, to a
# path relative to the Yardstick root.
$Script:DefaultedFolderKeys = [ordered]@{
    Recipes    = 'RecipesPath'
    Icons      = 'IconsPath'
    Buildspace = 'BuildSpace'
    Published  = 'Published'
    Temp       = 'Temp'
    Scripts    = 'Scripts'
    Tools      = 'Tools'
    Secrets    = 'Secrets'
}

# Folder keys that are off unless set. A blank value is meaningful here (it disables
# the feature), so these are resolved when present but never defaulted.
$Script:OptionalFolderKeys = @('SoftwareDropbox', 'SoftwareArchive', 'Backup')

# Folders the runner writes into and therefore creates when they are missing.
# Recipes is deliberately absent: an empty recipe folder is a misconfiguration, not
# something to paper over.
$Script:CreatedFolderKeys = @('Temp', 'Buildspace', 'Published', 'Scripts', 'Icons', 'Tools')

# Mirrors the ValidateSet attributes on Add-YardstickWin32App and
# New-YardstickWin32AppRequirementRule (Modules\YardstickIntune.psm1). A unit test
# keeps the two in step.
$Script:AllowedValues = @{
    defaultInstallExperience = @('system', 'user')
    defaultRestartBehavior   = @('basedOnReturnCode', 'allow', 'suppress', 'force')
    defaultArchitecture      = @('none', 'x64', 'x86', 'arm64', 'x64x86', 'AllWithARM64')
    defaultMinOSVersion      = @('W10_1607', 'W10_1703', 'W10_1709', 'W10_1803', 'W10_1809', 'W10_1903', 'W10_1909', 'W10_2004', 'W10_20H2', 'W10_21H1', 'W10_21H2', 'W10_22H2', 'W11_21H2', 'W11_22H2', 'W11_23H2', 'W11_24H2')
}

# Preference lists that Yardstick sends to Graph as Entra group object IDs.
$Script:GroupIdListKeys = @('defaultDeploymentGroups', 'defaultGroupsSkipAutoUpdates')


function Get-YardstickDefaultPreferences {
    <#
    .SYNOPSIS
    Returns the built-in value for every preference that has a safe default.

    .DESCRIPTION
    Only keys whose absence used to break a run, or that already had a hard-coded
    fallback in the runner, are listed. Keys where "unset" has its own meaning -
    Backup, SoftwareDropbox, adminEmailRecipient, credentialExpirationEmailEnabled -
    are deliberately left out so an omitted key behaves exactly as it always has.

    .PARAMETER ProjectRoot
    The Yardstick root folder. Default folder paths are resolved against it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProjectRoot
    )

    $layout = Get-YardstickProjectLayout -ProjectRoot $ProjectRoot

    $defaults = [ordered]@{}
    foreach ($key in $Script:DefaultedFolderKeys.Keys) {
        $layoutKey = $Script:DefaultedFolderKeys[$key]
        $defaults[$key] = if ($layout.PSObject.Properties[$layoutKey]) {
            $layout.$layoutKey
        } else {
            Join-Path $layout.ProjectRoot $layoutKey
        }
    }

    # Installations that predate the Local\ layout keep their secret vaults in the
    # root Secrets folder. Get-Secrets makes the same choice.
    $legacySecrets = Join-Path $layout.ProjectRoot 'Secrets'
    if (-not (Test-Path -LiteralPath $defaults.Secrets -PathType Container) -and
        (Test-Path -LiteralPath $legacySecrets -PathType Container)) {
        $defaults.Secrets = $legacySecrets
    }

    $builtIn = [ordered]@{
        backupVersionsToKeep                        = 3
        credentialTarget                            = 'Yardstick:IntuneGraph'
        credentialExpirationWarningDays             = 30
        credentialExpirationEmailIntervalHours      = 24

        defaultNumVersionsToKeep                    = 2
        defaultSupersedence                         = $true
        defaultUninstallPreviousVersion             = $false
        defaultAutoUpdate                           = $true
        useDetectAnchor                             = $true
        defaultMaximumInstallationTimeInMinutes     = 60
        defaultMinOSVersion                         = 'W10_20H2'
        defaultInstallExperience                    = 'system'
        defaultRestartBehavior                      = 'suppress'
        defaultAllowUserUninstall                   = $true
        defaultIs32BitApp                           = $false
        defaultArchitecture                         = 'none'
        defaultRequiresApproval                     = $false
        defaultDetectionScriptFileExtension         = 'ps1'
        defaultDetectionScriptRunAs32Bit            = $false
        defaultDetectionScriptEnforceSignatureCheck = $false
        defaultAllowDependentLinkUpdates            = $true
        defaultDeadlineDateOffset                   = 0
        defaultAvailableDateOffset                  = 0

        webRequestTimeoutSeconds                    = 120
        recipeStageTimeoutMinutes                   = 60
        packagingTimeoutMinutes                     = 30
        bitsStallTimeoutMinutes                     = 15

        emailNotificationEnabled                    = $false
        emailDeliveryMethod                         = 'outlook'
        emailSubject                                = 'Yardstick Application Update Report'
        emailSenderName                             = 'Yardstick Automation'
        emailPromptBeforeSending                    = $false
        smtpPort                                    = 25
        smtpUseSsl                                  = $false
        smtpCredentialTarget                        = 'Yardstick:Smtp'

        dependentLinkUpdateEnabled                  = $true
        dependentLinkUpdateRetryCount               = 3
        dependentLinkUpdateRetryDelaySeconds        = 5
        dependentLinkUpdateTimeoutSeconds           = 60
    }
    foreach ($key in $builtIn.Keys) {
        $defaults[$key] = $builtIn[$key]
    }

    $defaults
}


function Test-YardstickBlankValue {
    param($Value)
    $null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))
}


function Resolve-YardstickPreferencePath {
    <#
    .SYNOPSIS
    Expands environment variables in a folder preference and anchors a relative
    path to the Yardstick root. Absolute and UNC paths pass through unchanged.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ProjectRoot
    )

    $expanded = [Environment]::ExpandEnvironmentVariables($Path.Trim())
    if ($expanded -match '<REPLACE_') {
        # Leave template placeholders untouched so validation can name them.
        return $expanded
    }
    if ([IO.Path]::IsPathRooted($expanded)) {
        return $expanded
    }
    [IO.Path]::GetFullPath((Join-Path $ProjectRoot $expanded))
}


function Import-YardstickPreferences {
    <#
    .SYNOPSIS
    Loads preferences.yaml, fills in built-in defaults and resolves folder paths.

    .DESCRIPTION
    Returns the same hashtable the runner has always used as $Prefs. A key set in
    the file always wins; a key that is absent or blank takes its built-in default.
    Relative folder paths are anchored to the Yardstick root, so a preferences file
    can say "Temp: Artifacts\Temp" instead of an absolute path.

    .PARAMETER ProjectRoot
    The Yardstick root folder.

    .PARAMETER Path
    Read this preferences file instead of the one the project layout points at.

    .PARAMETER AllowMissing
    Return the built-in defaults when no preferences file exists, instead of
    throwing. For helper scripts that can run before setup is complete.

    .PARAMETER CreateFolders
    Create the working folders (Temp, Buildspace, Published, Scripts, Icons, Tools)
    when they do not exist.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProjectRoot,

        [string]$Path,

        [switch]$AllowMissing,

        [switch]$CreateFolders
    )

    $root = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
    if (-not $Path) {
        $Path = Get-YardstickPreferencesPath -ProjectRoot $root
    }

    $prefs = $null
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        if (-not (Get-Command ConvertFrom-Yaml -ErrorAction SilentlyContinue)) {
            throw "The PowerShell module 'powershell-yaml' is required to read '$Path'. Install it with: Install-Module powershell-yaml -Scope CurrentUser (or run .\Install-Yardstick.ps1)."
        }
        $prefs = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Yaml
    } elseif (-not $AllowMissing) {
        throw "No Yardstick preferences file was found at '$Path'. Run .\Install-Yardstick.ps1 to create one, or copy Config\Preferences.example.yaml to that location."
    }

    # An empty file parses to $null; a hashtable from ConvertFrom-Yaml is already
    # case-insensitive, which the runner has always relied on.
    if ($prefs -isnot [hashtable]) {
        $prefs = @{}
    }

    $defaults = Get-YardstickDefaultPreferences -ProjectRoot $root
    foreach ($key in $defaults.Keys) {
        if (-not $prefs.ContainsKey($key) -or (Test-YardstickBlankValue $prefs[$key])) {
            $prefs[$key] = $defaults[$key]
        }
    }

    foreach ($key in @($Script:DefaultedFolderKeys.Keys) + $Script:OptionalFolderKeys) {
        if (-not (Test-YardstickBlankValue $prefs[$key])) {
            $prefs[$key] = Resolve-YardstickPreferencePath -Path ([string]$prefs[$key]) -ProjectRoot $root
        }
    }

    if ($CreateFolders) {
        Initialize-YardstickFolders -Preferences $prefs | Out-Null
    }

    $prefs
}


function Initialize-YardstickFolders {
    <#
    .SYNOPSIS
    Creates the working folders a run writes into. Returns the folders it created.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Preferences
    )

    foreach ($key in $Script:CreatedFolderKeys) {
        $folder = [string]$Preferences[$key]
        if ([string]::IsNullOrWhiteSpace($folder) -or $folder -match '<REPLACE_') {
            continue
        }
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
            if ($PSCmdlet.ShouldProcess($folder, "Create $key folder")) {
                New-Item -ItemType Directory -Path $folder -Force | Out-Null
                $folder
            }
        }
    }
}


function Get-YardstickPreferenceAllowedValues {
    <#
    .SYNOPSIS
    Returns the values a preference key accepts, or nothing when it is free-form.
    #>
    param([Parameter(Mandatory)][string]$Key)
    foreach ($name in $Script:AllowedValues.Keys) {
        if ($name -eq $Key) { return $Script:AllowedValues[$name] }
    }
}


function Get-YardstickFolderProblem {
    <#
    .SYNOPSIS
    Returns why a folder cannot be opened, or nothing when it can.

    .DESCRIPTION
    Test-Path says only $false, which reads the same for a mistyped path and a
    share that refused the account's credentials. Listing the folder surfaces
    Windows' own reason.
    #>
    param([Parameter(Mandatory)][string]$Path)
    try {
        $entries = [System.IO.Directory]::EnumerateFileSystemEntries($Path).GetEnumerator()
        try { [void]$entries.MoveNext() } finally { $entries.Dispose() }
    } catch {
        $_.Exception.GetBaseException().Message.Trim()
    }
}


function Test-YardstickPreferences {
    <#
    .SYNOPSIS
    Validates a loaded preferences hashtable.

    .DESCRIPTION
    Errors are settings that would make every recipe fail - a value Intune will
    reject, a template placeholder that was never filled in, a missing recipe
    folder. Warnings are settings that only affect an optional feature such as
    email or backups.

    .OUTPUTS
    PSCustomObject with IsValid, Errors and Warnings, the same shape as
    Test-Prerequisites.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Preferences
    )

    $errors = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    foreach ($key in @($Preferences.Keys)) {
        foreach ($item in @($Preferences[$key])) {
            if ($item -is [string] -and $item -match '<REPLACE_[A-Z_]*>') {
                $errors.Add("Preference '$key' still contains the template placeholder '$($Matches[0])'. Replace it with a real value, or delete the key to use the default.")
                break
            }
        }
    }

    $recipes = [string]$Preferences.Recipes
    if ([string]::IsNullOrWhiteSpace($recipes) -or -not (Test-Path -LiteralPath $recipes -PathType Container)) {
        $errors.Add("The Recipes folder '$recipes' does not exist.")
    }

    foreach ($key in $Script:AllowedValues.Keys) {
        $value = $Preferences[$key]
        if (-not (Test-YardstickBlankValue $value) -and ([string]$value -notin $Script:AllowedValues[$key])) {
            $errors.Add("Preference '$key' is '$value', which Intune does not accept. Use one of: $($Script:AllowedValues[$key] -join ', ').")
        }
    }

    foreach ($key in $Script:GroupIdListKeys) {
        foreach ($item in @($Preferences[$key])) {
            if (Test-YardstickBlankValue $item) { continue }
            $guid = [guid]::Empty
            if (-not [guid]::TryParse([string]$item, [ref]$guid)) {
                $errors.Add("Preference '$key' contains '$item', which is not an Entra group object ID (a GUID).")
            }
        }
    }

    $numVersions = $Preferences.defaultNumVersionsToKeep -as [int]
    if ($null -eq $numVersions -or $numVersions -lt 1) {
        $errors.Add("Preference 'defaultNumVersionsToKeep' is '$($Preferences.defaultNumVersionsToKeep)'. It must be a whole number of at least 1.")
    }

    foreach ($key in 'defaultDeadlineDateOffset', 'defaultAvailableDateOffset', 'defaultMaximumInstallationTimeInMinutes') {
        $value = $Preferences[$key]
        if (-not (Test-YardstickBlankValue $value) -and $null -eq ($value -as [int])) {
            $errors.Add("Preference '$key' is '$value'. It must be a whole number.")
        }
    }

    if ($Preferences.emailNotificationEnabled -eq $true) {
        $method = ([string]$Preferences.emailDeliveryMethod).ToLowerInvariant()
        if ($method -notin 'outlook', 'smtp') {
            $warnings.Add("emailDeliveryMethod '$($Preferences.emailDeliveryMethod)' is not 'outlook' or 'smtp'. Outlook will be used.")
        }
        if ($method -eq 'smtp') {
            foreach ($key in 'smtpServer', 'emailSendFromAddress') {
                if (Test-YardstickBlankValue $Preferences[$key]) {
                    $warnings.Add("emailDeliveryMethod is smtp but '$key' is not set, so the email report cannot be sent.")
                }
            }
        }
        if ((Test-YardstickBlankValue $Preferences.emailRecipient) -and (Test-YardstickBlankValue $Preferences.adminEmailRecipient)) {
            $warnings.Add("Email notifications are enabled but neither 'emailRecipient' nor 'adminEmailRecipient' is set.")
        }
    }

    foreach ($key in $Script:OptionalFolderKeys) {
        $folder = [string]$Preferences[$key]
        if (-not [string]::IsNullOrWhiteSpace($folder) -and $folder -notmatch '<REPLACE_' -and
            ($problem = Get-YardstickFolderProblem -Path $folder)) {
            $warnings.Add("The $key folder '$folder' is not reachable: $problem")
        }
    }

    [PSCustomObject]@{
        IsValid  = ($errors.Count -eq 0)
        Errors   = [string[]]$errors
        Warnings = [string[]]$warnings
    }
}


Export-ModuleMember -Function @(
    'Get-YardstickDefaultPreferences',
    'Import-YardstickPreferences',
    'Initialize-YardstickFolders',
    'Test-YardstickPreferences',
    'Get-YardstickPreferenceAllowedValues',
    'Get-YardstickProjectLayout',
    'Get-YardstickPreferencesPath',
    'Get-YardstickRecipeOverridesPath'
)
