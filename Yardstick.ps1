<#
.DESCRIPTION
    Yardstick is a PowerShell script designed to automate the process of updating and managing Win32 applications in Microsoft Intune. It allows users to specify applications to update, either individually or in groups, and handles the downloading, packaging, and deployment of these applications.

.SYNOPSIS
    Adds and Updates Win32 applications in Microsoft Intune.

.PARAMETER ApplicationId
    The ID (or list of IDs) of the application(s) to update. Accepts a single string or a comma-separated array of strings.
    May only be used with -Force, -NoDelete, and -Repair parameters.

.PARAMETER Group
    The name of the group of applications to update. This parameter is used when updating a group of applications.
    May only be used with -Force, -NoDelete and -Repair parameters.

.PARAMETER All
    A switch parameter that indicates whether to update all applications in the repository. If specified, all applications will be processed.
    May only be used with -NoInteractive, -Force, -NoDelete and -Repair parameters.

.PARAMETER NoInteractive
    A switch parameter that, when specified, excludes interactive applications from the update process. This is denoted by recipes that are in the "Interactive" folder.
    May be used with all other parameters except ApplicationId.

.PARAMETER Force
    A switch parameter that forces the replacement of an application, even if it is already up-to-date.
    May be used with all other parameters.

.PARAMETER NoDelete
    A switch parameter that prevents the deletion of old versions of applications after an update. This is useful for debugging or testing purposes.
    May be used with all other parameters.

.PARAMETER NoEmail
    A switch parameter that suppresses the sending of email notifications after the update process is complete.
    May be used with all other parameters.

.PARAMETER Repair
    A switch parameter that repairs the application by renaming any incorrectly named applications to their correct format.
    May be used with all other parameters.

.EXAMPLE
    .\Yardstick.ps1 -All
    The recommended way to update all applications in the repository. This will update all applications.

.EXAMPLE
    .\Yardstick.ps1 -All -NoInteractive
    The recommended way to automate Yardstick. This will update all applications in the repository, excluding interactive applications.

.EXAMPLE
    .\Yardstick.ps1 -ApplicationId "ExampleApp" -Force -NoDelete
    This command updates the application with the ID "ExampleApp", forcing the update and preventing the deletion of old versions.

.EXAMPLE
    .\Yardstick.ps1 -ApplicationId "App1","App2","App3"
    This command updates three applications by their IDs.

.EXAMPLE
    .\Yardstick.ps1 -Group "ExampleGroup" -NoInteractive -Force
    This command updates all applications in the group "ExampleGroup", and forces the update.

.EXAMPLE
    .\Yardstick.ps1 -All -NoInteractive -Force -NoDelete
    This command updates all applications in the repository, excluding interactive applications, forcing the update and preventing the deletion of old versions.
#>

using module .\Modules\VersionPro.psm1

param (
    [Alias("AppId", "AppIds")]
    [parameter(ParameterSetName="SingleApp")]
    [String[]] $ApplicationId,

    [parameter(ParameterSetName="GroupApps")]
    [String] $Group,

    [parameter(ParameterSetName="AllApps")]
    [Switch] $All,

    [parameter(ParameterSetName="AllApps")]
    [Switch] $NoInteractive,

    [parameter(ParameterSetName="SingleApp")]
    [parameter(ParameterSetName="GroupApps")]
    [parameter(ParameterSetName="AllApps")]
    [Switch] $Force,

    [parameter(ParameterSetName="SingleApp")]
    [parameter(ParameterSetName="GroupApps")]
    [parameter(ParameterSetName="AllApps")]
    [Switch] $NoDelete,

    [parameter(ParameterSetName="SingleApp")]
    [parameter(ParameterSetName="GroupApps")]
    [parameter(ParameterSetName="AllApps")]
    [Switch] $Repair,

    [Switch] $NoEmail
)

# So we can get redirected URLs
Add-Type -AssemblyName Microsoft.PowerShell.Commands.Utility

# Ensure non-terminating errors are caught by try/catch blocks
$ErrorActionPreference = 'Stop'

# Constants
$Global:LogLocation = "$PSScriptRoot"
$Global:LogFile = "YLog.log"

# Import core modules first (needed for Write-Log and Test-Prerequisites)
Import-Module ${PSScriptRoot}\Modules\YardstickSupport.psm1 -Scope Global -Force

$CustomModuleImports = Get-ChildItem -Path $PSScriptRoot\Modules\Custom\*.psm1
foreach ($Module in $CustomModuleImports) {
    # Force allows us to reload them for testing
    Import-Module "$($Module.FullName)" -Scope Global -Force
}

# Initialize script variables
$Script:Applications = [System.Collections.Generic.List[PSObject]]::new()

# So we can pop at the end
Set-Location $PSScriptRoot

# Initialize the log file
Write-Log -Init

# Validate parameters
if ((-not $ApplicationId) -and (($All -eq $false) -and (!$Group))) {
    Write-Log "Please provide parameter -ApplicationId"
    exit 1
}

# Import preferences file:
try {
    $Prefs = Get-Content $PSScriptRoot\Preferences.yaml | ConvertFrom-Yaml
} catch {
    Write-Error "Unable to open preferences.yaml!"
    exit 1
}

# Validate prerequisites before importing external modules
$prereqResult = Test-Prerequisites -ToolsPath $Prefs.Tools
foreach ($w in $prereqResult.Warnings) { Write-Log "WARNING: $w" }
if (-not $prereqResult.IsValid) {
    foreach ($e in $prereqResult.Errors) { Write-Log "ERROR: $e" }
    Write-Log "Prerequisite checks failed. Exiting."
    exit 1
}

# Selenium 4's ScriptsToProcess defines PowerShell classes/enums in the caller's scope,
# but the module's own session state cannot see them during parameter binding. Defining
# the types via Add-Type (C#) BEFORE importing the module places them in the .NET
# AppDomain where they are visible to every session state, including the module's.
if (-not ('ValidateURIAttribute' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Management.Automation;

public class ValidateURIAttribute : ValidateArgumentsAttribute {
    protected override void Validate(object arguments, EngineIntrinsics engineIntrinsics) {
        Uri outUri;
        if (Uri.TryCreate(arguments != null ? arguments.ToString() : "", UriKind.Absolute, out outUri)) { return; }
        throw new ValidationMetadataException("Incorrect StartURL please make sure the URL starts with http:// or https://");
    }
}
"@
}

if (-not ('SeBrowsers' -as [type])) {
    Add-Type -TypeDefinition @"
public enum SeBrowsers {
    Chrome,
    Edge,
    Firefox,
    InternetExplorer,
    MSEdge
}
"@
}

if (-not ('SeWindowState' -as [type])) {
    Add-Type -TypeDefinition @"
public enum SeWindowState {
    Headless,
    Default,
    Minimized,
    Maximized,
    Fullscreen
}
"@
}

if (-not ('SeBySelector' -as [type])) {
    Add-Type -TypeDefinition @"
public enum SeBySelector {
    ClassName,
    CssSelector,
    Id,
    LinkText,
    PartialLinkText,
    Name,
    TagName,
    XPath
}
"@
}

if (-not ('SeBySelect' -as [type])) {
    Add-Type -TypeDefinition @"
public enum SeBySelect {
    Index,
    Text,
    Value
}
"@
}

# Import external modules (after Selenium compat types are in the AppDomain)
Import-Module powershell-yaml -Scope Local
Import-Module IntuneWin32App -Scope Local
Import-Module Selenium -Scope Global -ErrorAction SilentlyContinue
Import-Module TUN.CredentialManager -Scope Local -ErrorAction SilentlyContinue



function Set-ScriptVariables {
    <#
    .SYNOPSIS
    Sets script-scoped variables from application parameters and per-app preferences.

    .DESCRIPTION
    Sets per-application script variables from recipe parameters, falling back to
    global preferences with null coalescing. Folder locations and connection settings
    are set once at script startup and are not reassigned here.

    .PARAMETER Parameters
    Hashtable containing application-specific parameters from the YAML recipe.

    .PARAMETER Preferences
    Hashtable containing global preferences from the preferences.yaml file.
    #>
    param(
        [hashtable]$Parameters,
        [hashtable]$Preferences
    )

    # Set all variables from default preferences and the application recipe
    $Script:Url = if ($Parameters.urlRedirects -eq $true) {Get-RedirectedUrl $Parameters.url} else {$Parameters.url}
    $Script:Id = $Parameters.id
    $Script:Version = $Parameters.version
    $Script:FileDetectionVersion = $Parameters.fileDetectionVersion
    $Script:DisplayName = $Parameters.displayName
    $Script:DisplayVersion = $Parameters.displayVersion
    $Script:FileName = $Parameters.fileName
    $Script:FileDetectionPath = $Parameters.fileDetectionPath
    $Script:PreDownloadScript = if ($Parameters.preDownloadScript) { [ScriptBlock]::Create($Parameters.preDownloadScript) }
    $Script:DownloadScript = if ($Parameters.downloadScript) { [ScriptBlock]::Create($Parameters.downloadScript) }
    $Script:PostDownloadScript = if ($Parameters.postDownloadScript) { [ScriptBlock]::Create($Parameters.postDownloadScript) }
    $Script:PostRunScript = if ($Parameters.postRunScript) { [ScriptBlock]::Create($Parameters.postRunScript) }
    $Script:InstallScript = $Parameters.installScript
    $Script:UninstallScript = $Parameters.uninstallScript
    $Script:PowerShellInstallScript = $Parameters.powerShellInstallScript
    $Script:PowerShellUninstallScript = $Parameters.powerShellUninstallScript
    
    # Use parameters or fall back to preferences
    $Script:ScopeTags = if ($null -ne $Parameters.scopeTags) { $Parameters.scopeTags } else { $Preferences.defaultScopeTags }
    $Script:Owner = if ($null -ne $Parameters.owner) { $Parameters.owner } else { $Preferences.defaultOwner }
    $Script:MaximumInstallationTimeInMinutes = if ($null -ne $Parameters.maximumInstallationTimeInMinutes) { $Parameters.maximumInstallationTimeInMinutes } else { $Preferences.defaultMaximumInstallationTimeInMinutes }
    $Script:MinOSVersion = if ($null -ne $Parameters.minOSVersion) { $Parameters.minOSVersion } else { $Preferences.defaultMinOSVersion }
    $Script:InstallExperience = if ($null -ne $Parameters.installExperience) { $Parameters.installExperience } else { $Preferences.defaultInstallExperience }
    $Script:RestartBehavior = if ($null -ne $Parameters.restartBehavior) { $Parameters.restartBehavior } else { $Preferences.defaultRestartBehavior }
    $Script:AvailableGroups = if ($null -ne $Parameters.availableGroups) { $Parameters.availableGroups } else { $Preferences.defaultAvailableGroups }
    $Script:RequiredGroups = if ($null -ne $Parameters.requiredGroups) { $Parameters.requiredGroups } else { $Preferences.defaultRequiredGroups }
    $Script:DefaultDeploymentGroups = if ($null -ne $Parameters.defaultDeploymentGroups) { $Parameters.defaultDeploymentGroups } else { $Preferences.defaultDeploymentGroups }
    $Script:AllowUserUninstall = if ($null -ne $Parameters.allowUserUninstall) { $Parameters.allowUserUninstall } else { $Preferences.defaultAllowUserUninstall }
    $Script:Is32BitApp = if ($null -ne $Parameters.is32BitApp) { $Parameters.is32BitApp } else { $Preferences.defaultIs32BitApp }
    $Script:Architecture = if ($null -ne $Parameters.architecture) { $Parameters.architecture } else { $Preferences.defaultArchitecture }
    $Script:DeadlineDateOffset = if ($null -ne $Parameters.deadlineDateOffset) { $Parameters.deadlineDateOffset } else { $Preferences.defaultDeadlineDateOffset }
    $Script:AvailableDateOffset = if ($null -ne $Parameters.availableDateOffset) { $Parameters.availableDateOffset } else { $Preferences.defaultAvailableDateOffset }
    $Script:AllowDependentLinkUpdates = if ($null -ne $Parameters.allowDependentLinkUpdates) { $Parameters.allowDependentLinkUpdates } else { $Preferences.defaultAllowDependentLinkUpdates }
    
    # Detection-related variables
    $Script:DetectionType = $Parameters.detectionType
    $Script:FileDetectionVersion = $Parameters.fileDetectionVersion
    $Script:FileDetectionMethod = $Parameters.fileDetectionMethod
    $Script:FileDetectionName = $Parameters.fileDetectionName
    $Script:FileDetectionOperator = $Parameters.fileDetectionOperator
    $Script:FileDetectionDateTime = $Parameters.fileDetectionDateTime
    $Script:FileDetectionValue = $Parameters.fileDetectionValue
    $Script:RegistryDetectionMethod = $Parameters.registryDetectionMethod
    $Script:RegistryDetectionKey = $Parameters.registryDetectionKey
    $Script:RegistryDetectionValueName = $Parameters.registryDetectionValueName
    $Script:RegistryDetectionValue = $Parameters.registryDetectionValue
    $Script:RegistryDetectionOperator = $Parameters.registryDetectionOperator
    $Script:DetectionScript = $Parameters.detectionScript
    $Script:DetectionScriptFileExtension = if ($null -ne $Parameters.detectionScriptFileExtension) { $Parameters.detectionScriptFileExtension } else { $Preferences.defaultDetectionScriptFileExtension }
    $Script:DetectionScriptRunAs32Bit = if ($null -ne $Parameters.detectionScriptRunAs32Bit) { $Parameters.detectionScriptRunAs32Bit } else { $Preferences.defaultdetectionScriptRunAs32Bit }
    $Script:DetectionScriptEnforceSignatureCheck = if ($null -ne $Parameters.detectionScriptEnforceSignatureCheck) { $Parameters.detectionScriptEnforceSignatureCheck } else { $Preferences.defaultdetectionScriptEnforceSignatureCheck }
    $Script:DependentLinkUpdateEnabled = if ($null -ne $Parameters.dependentLinkUpdateEnabled) { $Parameters.dependentLinkUpdateEnabled } elseif ($null -ne $Preferences.dependentLinkUpdateEnabled) { $Preferences.dependentLinkUpdateEnabled } else { $true }
    $Script:DependentLinkUpdateRetryCount = if ($null -ne $Parameters.dependentLinkUpdateRetryCount) { $Parameters.dependentLinkUpdateRetryCount } elseif ($null -ne $Preferences.dependentLinkUpdateRetryCount) { $Preferences.dependentLinkUpdateRetryCount } else { 3 }
    $Script:DependentLinkUpdateRetryDelaySeconds = if ($null -ne $Parameters.dependentLinkUpdateRetryDelaySeconds) { $Parameters.dependentLinkUpdateRetryDelaySeconds } elseif ($null -ne $Preferences.dependentLinkUpdateRetryDelaySeconds) { $Preferences.dependentLinkUpdateRetryDelaySeconds } else { 5 }
    $Script:DependentLinkUpdateTimeoutSeconds = if ($null -ne $Parameters.dependentLinkUpdateTimeoutSeconds) { $Parameters.dependentLinkUpdateTimeoutSeconds } elseif ($null -ne $Preferences.dependentLinkUpdateTimeoutSeconds) { $Preferences.dependentLinkUpdateTimeoutSeconds } else { 60 }
    $Script:DependentApplicationBlacklist = if ($Parameters.dependentApplicationBlacklist) {
        $Parameters.dependentApplicationBlacklist
    } elseif ($Preferences.dependentApplicationBlacklist) {
        $Preferences.dependentApplicationBlacklist
    } else {
        @()
    }
    if (-not ($Script:DependentApplicationBlacklist -is [System.Collections.IEnumerable])) {
        $Script:DependentApplicationBlacklist = @($Script:DependentApplicationBlacklist)
    }
    
    # Additional variables
    $Script:IconFile = $Parameters.iconFile
    $Script:Description = $Parameters.description
    $Script:Publisher = $Parameters.publisher
    $Script:Arm64FilterName = $Preferences.arm64FilterName
    $Script:Amd64FilterName = $Preferences.amd64FilterName
    $Script:VersionLock = $Parameters.versionLock
    $Script:ProductCode = $null
    
    # Handle version lock logic
    if ($Script:VersionLock) {
        $Script:NumVersionsToKeep = 1
        if ($Parameters.numVersionsToKeep -gt 1) {
            Write-Log "Warning: Version lock is set, but numVersionsToKeep is set to $($Parameters.numVersionsToKeep). This will be ignored."
        }
    } else {
        $Script:NumVersionsToKeep = if ($null -ne $Parameters.numVersionsToKeep) { $Parameters.numVersionsToKeep } else { $Preferences.defaultNumVersionsToKeep }
    }

    # Supersedence + auto-update wiring. `autoUpdate` is accepted as an alias
    # for `autoUpdateOnAssignment` so pre-existing recipes continue to work.
    $Script:Supersedence = if ($null -ne $Parameters.supersedence) {
        [bool]$Parameters.supersedence
    } elseif ($null -ne $Preferences.defaultSupersedence) {
        [bool]$Preferences.defaultSupersedence
    } else { $true }

    $Script:UninstallPreviousVersion = if ($null -ne $Parameters.uninstallPreviousVersion) {
        [bool]$Parameters.uninstallPreviousVersion
    } elseif ($null -ne $Preferences.defaultUninstallPreviousVersion) {
        [bool]$Preferences.defaultUninstallPreviousVersion
    } else { $false }

    $Script:AutoUpdateOnAssignment = if ($null -ne $Parameters.autoUpdateOnAssignment) {
        [bool]$Parameters.autoUpdateOnAssignment
    } elseif ($null -ne $Parameters.autoUpdate) {
        [bool]$Parameters.autoUpdate
    } elseif ($null -ne $Preferences.defaultAutoUpdate) {
        [bool]$Preferences.defaultAutoUpdate
    } else { $true }

    # Groups that opt out of native auto-update. Unlike every other recipe
    # setting this is a union, not an override: a group excluded tenant-wide
    # stays excluded even when a recipe supplies its own list.
    $Script:GroupsSkipAutoUpdates = @(
        @($Preferences.defaultGroupsSkipAutoUpdates) + @($Parameters.groupSkipAutoUpdates) |
            Where-Object { $_ } | Select-Object -Unique
    )

    $Script:UseDetectAnchor = if ($null -ne $Preferences.useDetectAnchor) {
        [bool]$Preferences.useDetectAnchor
    } else { $true }

    # Handle PowerShell Script batch handoff
    if (($Script:PowerShellInstallScript) -and (!$Script:InstallScript)) {
        $Script:InstallScript = "powershell.exe -noprofile -executionpolicy bypass -file .\install.ps1"
    }
    if (($Script:PowerShellUninstallScript) -and (!$Script:UninstallScript)) {
        $Script:UninstallScript = "powershell.exe -noprofile -executionpolicy bypass -file .\uninstall.ps1"
    }
}



# Import Folder Locations from preferences
$Script:Temp = $Prefs.Temp
$Script:BuildSpace = $Prefs.Buildspace
$Script:Scripts = $Prefs.Scripts
$Script:Published = $Prefs.Published
$Script:Recipes = $Prefs.Recipes
$Script:Icons = $Prefs.Icons
$Script:Tools = $Prefs.Tools
$Script:Secrets = $Prefs.Secrets

# Backup of uploaded .intunewin files is optional - a blank or absent Backup key
# turns it off. A present-but-empty YAML key comes through as "", so this checks
# for whitespace rather than just $null.
$Script:Backup = $Prefs.Backup
$Script:BackupEnabled = -not [string]::IsNullOrWhiteSpace($Script:Backup)
$Script:BackupVersionsToKeep = if (($Prefs.backupVersionsToKeep -as [int]) -gt 0) { [int]$Prefs.backupVersionsToKeep } else { 3 }

# Import Intune Connection Settings
$Global:TenantID = $Prefs.TenantID
$Global:ClientID = $Prefs.ClientID
$Global:ClientSecret = $Prefs.ClientSecret
$Global:ScriptRoot = $PSScriptRoot

# Validate ApplicationId parameter
if ($ApplicationId) {
    foreach ($AppIdToValidate in $ApplicationId) {
        $MatchingApps = Get-ChildItem $Recipes | Where-Object Name -ne 'Disabled' | Get-ChildItem -File | Where-Object Name -match "$AppIdToValidate\.ya{0,1}ml"
        if ($null -eq $MatchingApps) {
            Write-Log "ERROR: Application $AppIdToValidate not found. (Excluding Disabled Folder)"
            exit 1
        }
    }
}

function Get-ApplicationsToProcess {
    <#
    .SYNOPSIS
    Determines which applications to process based on command-line parameters.
    
    .DESCRIPTION
    This function encapsulates the logic for selecting applications to process,
    whether processing a single app, a group, or all applications.
    
    .PARAMETER ApplicationId
    The ID of a single application to process.
    
    .PARAMETER Group
    The name of an application group to process.
    
    .PARAMETER All
    Switch to process all applications.
    
    .PARAMETER NoInteractive
    Switch to exclude interactive applications when processing all.
    #>
    param(
        [string[]]$ApplicationId,
        [string]$Group,
        [switch]$All,
        [switch]$NoInteractive
    )
    
    $applications = [System.Collections.Generic.List[PSObject]]::new()
    
    if ($All) {
        $folderFilter = if ($NoInteractive) { 
            { $_.Directory -notlike '*Disabled*' -and $_.Directory -notlike '*Interactive*' }
        } else { 
            { $_.Directory -notlike '*Disabled*' }
        }
        
        $ApplicationFullNames = (Get-ChildItem $Recipes -Recurse -File | Where-Object $folderFilter | Where-Object Name -match ".*\.ya{0,1}ml").Name
        foreach ($Application in $ApplicationFullNames) {
            $applications.Add([String]($Application -split "\.ya{0,1}ml")[0]) | Out-Null
        }
    } elseif ($Group) {
        try {
            $groupFile = Get-Content $PSScriptRoot\RecipeGroups.yaml | ConvertFrom-Yaml
            $groupFile[$Group] | ForEach-Object {
                $applications.Add([String]$_) | Out-Null
            }
        } catch {
            Write-Log "ERROR: There was an issue importing the application group! Exiting."
            exit 3
        }
    } else {
        foreach ($Id in $ApplicationId) {
            $applications.Add([String]$Id) | Out-Null
        }
    }
    return $applications
}



function Update-ScriptPlaceholders {
    <#
    .SYNOPSIS
    Replaces placeholders in script strings with actual values.
    
    .DESCRIPTION
    This function handles the replacement of common placeholders like <filename>,
    <productcode>, and <version> in install, uninstall, and detection scripts.
    
    .PARAMETER FileName
    The actual filename to replace <filename> placeholders with.
    
    .PARAMETER ProductCode
    The MSI product code to replace <productcode> placeholders with.
    
    .PARAMETER Version
    The version string to replace <version> placeholders with.
    #>
    param(
        [string]$FileName,
        [string]$ProductCode = "",
        [string]$Version
    )
    
    # Replace the <filename> placeholder with the actual filename
    if ($Script:InstallScript) {
        $Script:InstallScript = $Script:InstallScript.replace("<filename>", $FileName)
    }
    if ($Script:PowerShellInstallScript) {
        $Script:PowerShellInstallScript = $Script:PowerShellInstallScript.replace("<filename>", $FileName)
    }
    if ($Script:UninstallScript) {
        $Script:UninstallScript = $Script:UninstallScript.replace("<filename>", $FileName)
    }
    if ($Script:PowerShellUninstallScript) {
        $Script:PowerShellUninstallScript = $Script:PowerShellUninstallScript.replace("<filename>", $FileName)
    }
    if ($Script:DetectionScript) {
        $Script:DetectionScript = $Script:DetectionScript.replace("<filename>", $FileName)
    }
    if ($Script:RegistryDetectionKey) {
        $Script:RegistryDetectionKey = $Script:RegistryDetectionKey.replace("<filename>", $FileName)
    }

    # Replace the <productcode> placeholder with the actual product code
    if ($ProductCode) {
        if ($Script:InstallScript) {
            $Script:InstallScript = $Script:InstallScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:PowerShellInstallScript) {
            $Script:PowerShellInstallScript = $Script:PowerShellInstallScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:UninstallScript) {
            $Script:UninstallScript = $Script:UninstallScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:PowerShellUninstallScript) {
            $Script:PowerShellUninstallScript = $Script:PowerShellUninstallScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:DetectionScript) {
            $Script:DetectionScript = $Script:DetectionScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:RegistryDetectionKey) {
            $Script:RegistryDetectionKey = $Script:RegistryDetectionKey.replace("<productcode>", $ProductCode)
        }
    }
    
    # Replace <version> placeholder with the actual version
    if ($Script:InstallScript) {
        $Script:InstallScript = $Script:InstallScript.replace("<version>", $Version)
    }
    if ($Script:PowerShellInstallScript) {
        $Script:PowerShellInstallScript = $Script:PowerShellInstallScript.replace("<version>", $Version)
    }
    if ($Script:UninstallScript) {
        $Script:UninstallScript = $Script:UninstallScript.replace("<version>", $Version) 
    }
    if ($Script:PowerShellUninstallScript) {
        $Script:PowerShellUninstallScript = $Script:PowerShellUninstallScript.replace("<version>", $Version) 
    }
    if ($Script:DetectionScript) {
        $Script:DetectionScript = $Script:DetectionScript.replace("<version>", $Version) 
    } 
    if ($Script:RegistryDetectionKey) {
        $Script:RegistryDetectionKey = $Script:RegistryDetectionKey.replace("<version>", $Version)
    }
}



function New-DetectionRule {
    <#
    .SYNOPSIS
    Creates Intune Win32 app detection rules based on detection type.
    
    .DESCRIPTION
    This function creates the appropriate detection rule object based on the
    specified detection type (file, msi, registry, or script).
    
    .PARAMETER DetectionType
    The type of detection rule to create (file, msi, registry, script).
    
    .PARAMETER ProductCode
    The MSI product code (required for MSI detection type).
    #>
    param(
        [string]$DetectionType,
        [string]$ProductCode = ""
    )
    
    switch ($DetectionType) {
        "file" {
            switch ($Script:FileDetectionMethod) {
                "exists" {
                    return New-IntuneWin32AppDetectionRuleFile -Existence -DetectionType "exists" -Path $Script:FileDetectionPath -FileOrFolder $Script:FileDetectionName
                }
                "modified" {
                    return New-IntuneWin32AppDetectionRuleFile -DateModified -Path $Script:FileDetectionPath -FileOrFolder $Script:FileDetectionName -Operator $Script:FileDetectionOperator -DateTimeValue $Script:FileDetectionDateTime
                }
                "created" {
                    return New-IntuneWin32AppDetectionRuleFile -DateCreated -Path $Script:FileDetectionPath -FileOrFolder $Script:FileDetectionName -Operator $Script:FileDetectionOperator -DateTimeValue (Get-Date $Script:FileDetectionDateTime)
                }
                "version" {
                    return New-IntuneWin32AppDetectionRuleFile -Version -Path $Script:FileDetectionPath -FileOrFolder $Script:FileDetectionName -Operator $Script:FileDetectionOperator -VersionValue $Script:FileDetectionVersion
                }
                "size" {
                    return New-IntuneWin32AppDetectionRuleFile -Size -Path $Script:FileDetectionPath -FileOrFolder $Script:FileDetectionName -Operator $Script:FileDetectionOperator -SizeinMBValue $Script:FileDetectionValue
                }
            }
        }
        "msi" {
            return New-IntuneWin32AppDetectionRuleMsi -ProductCode $ProductCode -ProductVersion $Script:FileDetectionVersion
        }
        "registry" {
            switch ($Script:RegistryDetectionMethod) {
                "exists" {
                    if ($Script:RegistryDetectionValue) {
                        return New-IntuneWin32AppDetectionRuleRegistry -Existence -KeyPath $Script:RegistryDetectionKey -ValueName $Script:RegistryDetectionValueName -DetectionType "exists"
                    } else {
                        return New-IntuneWin32AppDetectionRuleRegistry -Existence -KeyPath $Script:RegistryDetectionKey -DetectionType "exists"
                    }
                }
                "version" {
                    return New-IntuneWin32AppDetectionRuleRegistry -VersionComparison -KeyPath $Script:RegistryDetectionKey -ValueName $Script:RegistryDetectionValueName -Check32BitOn64System $Script:Is32BitApp -VersionComparisonOperator $Script:RegistryDetectionOperator -VersionComparisonValue $Script:RegistryDetectionValue
                }
                "integer" {
                    return New-IntuneWin32AppDetectionRuleRegistry -IntegerComparison -KeyPath $Script:RegistryDetectionKey -ValueName $Script:RegistryDetectionValueName -Check32BitOn64System $Script:Is32BitApp -IntegerComparisonOperator $Script:RegistryDetectionOperator -IntegerComparisonValue $Script:RegistryDetectionValue
                }
                "string" {
                    return New-IntuneWin32AppDetectionRuleRegistry -StringComparison -KeyPath $Script:RegistryDetectionKey -ValueName $Script:RegistryDetectionValueName -Check32BitOn64System $Script:Is32BitApp -StringComparisonOperator $Script:RegistryDetectionOperator -StringComparisonValue $Script:RegistryDetectionValue
                }
            }
        }
        "script" {
            if (!(Test-Path $Script:Scripts\$Script:Id)) {
                $item = New-Item -Name $Script:Id -ItemType Directory -Path $Script:Scripts | Out-Null
            }
            $ScriptLocation = "$($Script:Scripts)\$($Script:Id)\$($Script:Version).$($Script:DetectionScriptFileExtension)"
            Set-Content -Path $ScriptLocation -Value $Script:DetectionScript -Force
            $DetRule = New-IntuneWin32AppDetectionRuleScript -ScriptFile $ScriptLocation -EnforceSignatureCheck $Script:DetectionScriptEnforceSignatureCheck -RunAs32Bit $Script:DetectionScriptRunAs32Bit
            return $DetRule
        }
    }
}


function Invoke-Cleanup {
    <#
    .SYNOPSIS
    Performs cleanup operations on build and published directories.
    
    .DESCRIPTION
    This function cleans up temporary files and directories created during
    the application processing, with error handling for non-critical failures.
    #>
    # Backups read from the Published folder on a background thread, so make sure
    # none are still in flight before anything here deletes their source.
    Wait-YardstickBackup -TimeoutSeconds 600 | Out-Null

    Write-Log "Cleaning up the Buildspace..."
    try {
        Get-ChildItem $BuildSpace -Exclude ".gitkeep" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Log "Warning: Could not clean buildspace completely: $_"
    }
    
    Write-Log "Removing .intunewin files..."
    try {
        Get-ChildItem $Published -Exclude ".gitkeep" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Log "Warning: Could not clean published files completely: $_"
    }
}



# Get the list of applications to process
$Applications = Get-ApplicationsToProcess -ApplicationId $ApplicationId -Group $Group -All:$All -NoInteractive:$NoInteractive

# Initialize application tracking for email notifications
Initialize-ApplicationTracker

# Build run parameters string for email report
$RunParametersArray = @()
if ($ApplicationId) { $RunParametersArray += "-ApplicationId $($ApplicationId -join ',')" }
if ($Group) { $RunParametersArray += "-Group $Group" }
if ($All) { $RunParametersArray += "-All" }
if ($NoInteractive) { $RunParametersArray += "-NoInteractive" }
if ($Force) { $RunParametersArray += "-Force" }
if ($NoDelete) { $RunParametersArray += "-NoDelete" }
if ($Repair) { $RunParametersArray += "-Repair" }
$RunParameters = $RunParametersArray -join " "

# Main processing loop
foreach ($AppId_Processing in $Applications) {
    Write-Log "Starting update for $AppId_Processing..."
    Set-Location $PSScriptRoot
    
    # Initialize variables for tracking
    $CurrentDisplayName = $AppId_Processing
    
    try {
        # Refresh token if necessary
        Connect-AutoMSIntuneGraph
        
        # Clear the temp file
        Write-Log "Clearing the temp directory..."
        Get-ChildItem $Temp -Exclude ".gitkeep" -Recurse | Remove-Item -Recurse -Force

        # Open the YAML file and collect all necessary attributes
        try {
            $AppName = (Get-ChildItem $Recipes -Force -Recurse | Where-Object Name -ne 'Disabled' | Get-ChildItem -File -Recurse | Where-Object Name -match "^$AppId_Processing\.ya{0,1}ml")[0].FullName
            $Parameters = Get-Content "$AppName" | ConvertFrom-Yaml
        } catch {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $AppId_Processing -Version "Unknown" -ErrorMessage "Unable to open parameters file for $AppId_Processing" -FailureStage "Configuration"
            Write-Error "Unable to open parameters file for $AppId_Processing"
            continue
        }

        # Resolve base recipe inheritance if present
        if ($Parameters.ContainsKey('base')) {
            try {
                $Parameters = Merge-RecipeWithBase -Recipe $Parameters -RecipesPath $Recipes
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $AppId_Processing -Version "Unknown" -ErrorMessage "Failed to resolve base recipe: $_" -FailureStage "Configuration"
                Write-Error "Failed to resolve base recipe for ${ApplicationId}: $_"
                continue
            }
        }

        # Validate recipe schema before processing
        $validation = Test-RecipeSchema -Recipe $Parameters -RecipeId $AppId_Processing
        foreach ($w in $validation.Warnings) { Write-Log "WARNING: [Recipe $AppId_Processing] $w" }
        if (-not $validation.IsValid) {
            $errorMsg = "Recipe validation failed: $($validation.Errors -join '; ')"
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $AppId_Processing -Version "Unknown" -ErrorMessage $errorMsg -FailureStage "Recipe Validation"
            Write-Error "[Recipe $AppId_Processing] $errorMsg"
            continue
        }

        # Set all script variables from parameters and preferences
        Set-ScriptVariables -Parameters $parameters -Preferences $Prefs
        
        # Update tracking variables with actual values
        $CurrentDisplayName = $Script:DisplayName
        $DependentUpdateStatus = [ordered]@{}
        $ProtectedDependencyAppIds = [System.Collections.Generic.HashSet[string]]::new()
        $DependentLinkOptions = @{
            Enabled = [bool]$Script:DependentLinkUpdateEnabled
            RetryCount = [int]$Script:DependentLinkUpdateRetryCount
            RetryDelaySeconds = [int]$Script:DependentLinkUpdateRetryDelaySeconds
            TimeoutSeconds = [int]$Script:DependentLinkUpdateTimeoutSeconds
            Blacklist = $Script:DependentApplicationBlacklist
        }


        if ($Repair) {
            # Correct any naming discrepancies before we continue.
            # Skip the Ω DETECT - anchor - it deliberately holds a different name.
            $anchorName = Get-DetectAnchorName -DisplayName $Script:DisplayName
            $CurrentApps = @(Get-SameAppAllVersions $Script:DisplayName | Where-Object DisplayName -ne $anchorName)
            for ($i = 1; $i -lt $CurrentApps.Count; $i++) {
                if ($CurrentApps[$i].DisplayName -ne "$($Script:DisplayName) (N-$i)") {
                    Write-Log "Setting name for $($Script:DisplayName) (N-$i)"
                    Set-IntuneWin32App -Id $CurrentApps[$i].Id -DisplayName "$($Script:DisplayName) (N-$i)"
                }
            }
        }


        # Clear stale $matches from prior recipe iterations to prevent cross-contamination
        $null = "reset" -match "reset"

        # Run the pre-download script
        if ($Script:PreDownloadScript) {
            Write-Log "Running pre-download script..."
            try {
                Invoke-Command -ScriptBlock $Script:PreDownloadScript -NoNewScope
                Write-Log "Pre-download script ran successfully."
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Error while running pre-download PowerShell script: $_" -FailureStage "Pre-Download Script"
                Write-Error "Error while running pre-download PowerShell script"
                continue
            }
        } else {
            Write-Log "Skipping Pre-download script"
        }

        # Validate the extracted version before proceeding
        $ExistingVersions = Get-SameAppAllVersions $Script:DisplayName
        $existingVersionForCheck = if ($ExistingVersions -and $ExistingVersions.Count -gt 0) { $ExistingVersions.displayVersion[0] } else { $null }

        $versionValidation = Test-ExtractedVersion -Version $Script:Version -ApplicationId $AppId_Processing -ExistingVersion $existingVersionForCheck
        foreach ($w in $versionValidation.Warnings) {
            Write-Log "WARNING: [Version Check $AppId_Processing] $w"
        }
        if (-not $versionValidation.IsValid) {
            $versionErrorMsg = "Version validation failed: $($versionValidation.Errors -join '; ')"
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $(if ($null -ne $Script:Version) { $Script:Version } else { "null" }) -ErrorMessage $versionErrorMsg -FailureStage "Version Validation"
            Write-Error "[Version Check $AppId_Processing] $versionErrorMsg"
            continue
        }

        # Trim whitespace from version if present (warned but not blocked above)
        if ($Script:Version -ne $Script:Version.Trim()) {
            $Script:Version = $Script:Version.Trim()
            Write-Log "Trimmed whitespace from version: '$($Script:Version)'"
        }

        # Check if there is an up-to-date version in the repo already
        Write-Log "Checking if $($Script:DisplayName) $($Script:Version) is a new version..."
        
        if (-not $ExistingVersions) {
            Write-Log "No existing versions found for $($Script:DisplayName). Continuing with update."
            $VersionCompareResult = 0
        }
        else {
            $VersionCompareResult = Compare-AppVersions $Script:Version $($ExistingVersions.displayVersion[0])
        }
        
        # Check various conditions to determine if we should proceed
        if ($Force) {
            Write-Log "Force flag is set. Forcing update of $($Script:DisplayName) $($Script:Version)"
        } elseif (Test-VersionExcluded -Version $Script:Version -VersionLock $Script:VersionLock) {
            Write-Log "Version is locked to $($Script:VersionLock). Skipping update."
            continue
        } elseif ($ExistingVersions.displayVersion -contains $Script:Version) {
            Write-Log "$($Script:Id) $($Script:DisplayName) $($Script:Version) is already in the repo. Skipping update."
            continue
        } elseif ($VersionCompareResult -eq 1) {
            Write-Log "$($Script:DisplayName) $($Script:Version) is a newer version. Continuing with update."
        } elseif ($VersionCompareResult -eq -1) {
            Write-Log "$($Script:DisplayName) $($Script:Version) is older than the currently newest available version $($ExistingVersions.displayVersion[0]). Skipping update."
            continue
        }


        # See if this has been run before. If there are previous files, move them to a folder called "Old"
        if (Test-Path $BuildSpace\$($Script:Id)) {
            if (-not (Test-Path $BuildSpace\Old)) {
                New-Item -Path $BuildSpace -ItemType Directory -Name "Old"
            }
            Write-Log "Removing old Buildspace..."
            Move-Item -Path $BuildSpace\$($Script:Id) $BuildSpace\Old\$($Script:Id)-$(Get-Date -Format "MMddyyhhmmss")
        }
        if (Test-Path $Scripts\$($Script:Id)) {
            if (-not (Test-Path $Scripts\Old)) {
                New-Item -Path $Scripts -ItemType Directory -Name "Old"
            }
            Write-Log "Removing old script space..."
            Move-Item -Path $Scripts\$($Script:Id) $Scripts\Old\$($Script:Id)-$(Get-Date -Format "MMddyyhhmmss")
        }

        # Make the new BUILDSPACE directory
        New-Item -Path $BuildSpace\$($Script:Id) -ItemType Directory -Name $Script:Version
        Set-Location $BuildSpace\$($Script:Id)\$Script:Version

        Update-ScriptPlaceholders -FileName $Script:FileName -ProductCode $productCode -Version $Script:Version 

        # Download the new installer
        Write-Log "Starting download..."
        if ($Script:Url) {
            Write-Log "URL: $($Script:Url)"
        }
        if ((-not ($Script:Url)) -and (-not ($Script:DownloadScript))) {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "URL is empty - cannot continue" -FailureStage "Download"
            Write-Error "URL is empty - cannot continue."
            continue
        }
        
        if ($Script:DownloadScript) {
            Push-Location $BuildSpace\$Script:Id\$Script:Version
            try {
                Invoke-Command -ScriptBlock $Script:DownloadScript -NoNewScope
                Write-Log "Download script ran successfully."
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Error while running download PowerShell script: $_" -FailureStage "Download Script"
                Write-Error "Error while running download PowerShell script: $_"
                Write-Error "Script Contents: $Script:DownloadScript"
                continue
            }
            Pop-Location
        } else {
            try {
                Start-BitsTransfer -Source "$($Script:Url)" -Destination "$BuildSpace\$Script:Id\$Script:Version\$Script:FileName"
                Write-Log "File downloaded successfully using BITS transfer."
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Error downloading file: $_" -FailureStage "Download"
                Write-Error "Error downloading file: $_"
                continue
            }
        }
        # Update Placeholders again in case the download script changed anything
        Update-ScriptPlaceholders -FileName $Script:FileName -ProductCode $ProductCode -Version $Script:Version 
        # Run the post-download script
        if ($Script:PostDownloadScript) {
            Write-Log "Running post download script..."
            Push-Location $BuildSpace\$Script:Id\$Script:Version
            try {
                Invoke-Command -ScriptBlock $Script:PostDownloadScript -NoNewScope
                Write-Log "Post download script ran successfully."
            } catch {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Error while running post download PowerShell script: $_" -FailureStage "Post-Download Script"
                Write-Error "Error while running post download PowerShell script: $_"
                continue
            }
            Pop-Location
        }
        # Handle script placeholder replacement
        $ProductCode = ""
        if ($Script:FileName -match "\.msi$") {
            $ProductCode = Get-MSIProductCode $BuildSpace\$Script:Id\$Script:Version\$Script:FileName | Where-Object { $_ -match "^\{[0-9A-Fa-f\-]{36}\}$" }
            Write-Log "Product Code: $ProductCode"
            if (-not $ProductCode) {
                Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Could not determine MSI Product Code from $Script:FileName" -FailureStage "MSI Product Code Retrieval"
                Write-Error "Could not determine MSI Product Code from $Script:FileName"
                continue
            }
        }
        # Update placeholders for the final time
        Update-ScriptPlaceholders -FileName $Script:FileName -ProductCode $ProductCode -Version $Script:Version

        # DEBUG: Make sure that there aren't any placeholder strings left
        if ($Script:InstallScript -match "<filename>|<productcode>|<version>" -or
            $Script:PowerShellInstallScript -match "<filename>|<productcode>|<version>" -or
            $Script:UninstallScript -match "<filename>|<productcode>|<version>" -or
            $Script:PowerShellUninstallScript -match "<filename>|<productcode>|<version>" -or
            $Script:DetectionScript -match "<filename>|<productcode>|<version>") {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "One or more placeholder strings were not replaced in scripts." -FailureStage "Placeholder Replacement"
            Write-Error "One or more placeholder strings were not replaced in scripts."
            continue
        }

        # Write the contents of the install and uninstall scripts to files if they are PowerShell scripts
        if ($Script:PowerShellInstallScript) {
            Set-Content -Path $BuildSpace\$Script:Id\$Script:Version\install.ps1 -Value $Script:PowerShellInstallScript -Force
        }

        if ($Script:PowerShellUninstallScript) {
            Set-Content -Path $BuildSpace\$Script:Id\$Script:Version\uninstall.ps1 -Value $Script:PowerShellUninstallScript -Force
        }

        # Generate the .intunewin file
        Set-Location $PSScriptRoot
        Write-Log "Generating .intunewin file..."
        $App = New-IntuneWin32AppPackage -SourceFolder $BuildSpace\$Script:Id\$Script:Version -SetupFile $Script:FileName -OutputFolder $Published -Force

        # Upload .intunewin file to Intune
        # Detection Types
        if (!(Test-Path "$($Icons)\$($Script:IconFile)")) {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Icon file $($Script:IconFile) not found in Icons folder." -FailureStage "Icon Retrieval"
            Write-Error "Icon file $($Script:IconFile) not found in Icons folder."
            continue
        }
        $Icon = New-IntuneWin32AppIcon -FilePath "$($Icons)\$($Script:IconFile)"
        if (-not $Script:FileDetectionVersion) {
            $Script:FileDetectionVersion = $Script:Version
        }

        $DetectionRule = New-DetectionRule -DetectionType $Script:DetectionType -ProductCode $ProductCode

        # Generate the min OS requirement rule
        $RequirementRule = New-IntuneWin32AppRequirementRule -Architecture $Script:Architecture -MinimumSupportedWindowsRelease $Script:MinOSVersion

        # Create the Intune App
        Write-Log "Uploading $Script:DisplayName to Intune..."
        Connect-AutoMSIntuneGraph
        try {
            if ($Script:AllowUserUninstall) {
                $Win32App = Add-IntuneWin32App -FilePath $App.path -DisplayName $Script:DisplayName -Description $Script:Description -Publisher $Script:Publisher -InstallExperience $Script:InstallExperience -RestartBehavior $Script:RestartBehavior -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $Script:InstallScript -UninstallCommandLine $Script:UninstallScript -Icon $Icon -AppVersion "$Script:Version" -ScopeTagName $Script:ScopeTags -Owner $Script:Owner -MaximumInstallationTimeInMinutes $Script:MaximumInstallationTimeInMinutes -AllowAvailableUninstall
            } else {
                $Win32App = Add-IntuneWin32App -FilePath $App.path -DisplayName $Script:DisplayName -Description $Script:Description -Publisher $Script:Publisher -InstallExperience $Script:InstallExperience -RestartBehavior $Script:RestartBehavior -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $Script:InstallScript -UninstallCommandLine $Script:UninstallScript -Icon $Icon -AppVersion "$Script:Version" -ScopeTagName $Script:ScopeTags -Owner $Script:Owner -MaximumInstallationTimeInMinutes $Script:MaximumInstallationTimeInMinutes
            }
            Write-Log "Successfully uploaded $Script:DisplayName to Intune"
            Write-Log "Waiting for Intune to process the uploaded application..."
            $PublishTimeout = (Get-Date).AddMinutes(30)
            do {
                Start-Sleep -Seconds 15
                $LiveWin32App = Get-IntuneWin32App -Id $Win32App.id
                if ((Get-Date) -gt $PublishTimeout) {
                    throw "Timed out waiting for Intune to publish $Script:DisplayName after 30 minutes (state: $($LiveWin32App.publishingState))"
                }
            } while ($LiveWin32App.publishingState -ne "published")
            $UploadedAt = Get-Date
        } catch {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Failed to upload application to Intune: $_" -FailureStage "Intune Upload"
            Write-Error "Failed to upload application to Intune: $_"
            continue
        }

        # Hand the .intunewin off to a background thread that copies it to the
        # backup share and then deletes the local copy. Staging it out of
        # $Published first is what stops the next app's
        # New-IntuneWin32AppPackage -Force (which names its output from the setup
        # file, so two recipes can collide) from clobbering a file the thread is
        # still reading. Nothing in here may throw - a bad backup share must not
        # turn a successful upload into a failed application.
        if ($Script:BackupEnabled) {
            try {
                $BackupFileName = Get-YardstickBackupFileName -AppId $Script:Id -Version $Script:Version -Timestamp $UploadedAt
                $StageDir = Join-Path (Join-Path $Published '_BackupQueue') $Script:Id
                if (-not (Test-Path -LiteralPath $StageDir)) {
                    New-Item -ItemType Directory -Path $StageDir -Force | Out-Null
                }
                $StagedPath = Join-Path $StageDir $BackupFileName
                Invoke-WithRetry -Label "Stage .intunewin for backup" -MaxRetries 3 -DelaySeconds 1 `
                    -ScriptBlock { Move-Item -LiteralPath $App.path -Destination $StagedPath -Force } `
                    -VerifyBlock { Test-Path -LiteralPath $StagedPath } | Out-Null
                if (-not (Test-Path -LiteralPath $StagedPath)) {
                    throw "could not stage $($App.path)"
                }
                Start-YardstickBackup -SourcePath $StagedPath -BackupRoot $Script:Backup -AppId $Script:Id `
                    -FileName $BackupFileName -VersionsToKeep $Script:BackupVersionsToKeep
                Write-Log "Queued backup of $BackupFileName to $Script:Backup"
            } catch {
                Write-Log "WARNING: Could not queue backup for $($Script:Id): $_"
            }
        }


        ###################################################
        # SUPERSEDENCE, RETENTION, AUTO-UPDATE
        ###################################################

        # Refresh the version list from Intune (retry in case the new upload
        # has not yet propagated to Get-IntuneWin32App list results).
        Write-Log "Updating local application manifest..."
        Start-Sleep -Seconds 4
        try {
            $AllMatchingApps = Get-SameAppAllVersions $Script:DisplayName
            if (!($AllMatchingApps | Where-Object id -eq $Win32App.id)) {
                Write-Log "Newly created app not found in list of all matching apps. Waiting 5 seconds and trying again..."
                Start-Sleep -Seconds 5
                $AllMatchingApps = Get-SameAppAllVersions $Script:DisplayName
            }
            if (!($AllMatchingApps | Where-Object id -eq $Win32App.id)) {
                # The list endpoint lags behind creation. Fetch the new app directly
                # by id and splice it in - skipping here would leave a freshly
                # uploaded app with no assignments and no supersedence, and the next
                # run would consider the version already published and skip it.
                Write-Log "Newly created app still missing from the list. Fetching it directly by id."
                $DirectApp = Get-IntuneWin32App -Id $Win32App.id -ErrorAction Stop
                if (-not $DirectApp) {
                    throw "Intune returned no application for id $($Win32App.id)"
                }
                $AllMatchingApps = @($DirectApp) + @($AllMatchingApps)
            }
        } catch {
            Write-Log "There was an error fetching information about existing applications. Exiting"
            # exit does not wait for background threads, so drain the backup we
            # queued a few lines above before tearing the process down.
            Wait-YardstickBackup -TimeoutSeconds 120 | Out-Null
            Exit 4
        }

        $CurrentApp = $AllMatchingApps | Where-Object id -eq $Win32App.id | Select-Object -First 1
        $OtherApps  = @($AllMatchingApps | Where-Object id -ne $CurrentApp.id)

        # 1. Anchor identification. Version-detection recipes (msi, file-by-
        #    version, registry-by-version) get one sticky Ω DETECT - anchor so a
        #    fast-moving app cannot prune away every version that could still
        #    detect a stale endpoint. The anchor is a permanent supersedence
        #    target: its detection rule matches a very wide version range, so
        #    superseding it is what actually pulls ancient installs forward.
        $Anchor = $null
        $anchorStatus = $null
        if ($Script:UseDetectAnchor -and (Test-IsVersionDetection -DetectionType $Script:DetectionType -FileDetectionMethod $Script:FileDetectionMethod -RegistryDetectionMethod $Script:RegistryDetectionMethod)) {
            $Anchor = Get-DetectAnchor -DisplayName $Script:DisplayName
            if (-not $Anchor -and $OtherApps.Count -gt 0) {
                # First run under the new model - pin the oldest surviving version.
                $AnchorCandidate = $OtherApps | Sort-Object @{Expression = {[VersionPro]$_.displayVersion}} | Select-Object -First 1
                Set-DetectAnchor -App $AnchorCandidate -DisplayName $Script:DisplayName
                # Re-fetch so the object reflects the new name. Fall back to the
                # candidate if Intune has not caught up yet - otherwise the app we
                # just pinned would be renamed straight back to (N-x) below.
                $Anchor = Get-DetectAnchor -DisplayName $Script:DisplayName
                if (-not $Anchor) { $Anchor = $AnchorCandidate }
                $anchorStatus = "pinned ($($Anchor.displayVersion))"
            } elseif ($Anchor) {
                $anchorStatus = "existing ($($Anchor.displayVersion))"
            } else {
                $anchorStatus = "n/a (no prior versions)"
            }
        }

        # 2. The anchor is off-limits to renaming (step 5) and pruning (step 8): it
        #    holds a reserved name forever, so it must never be numbered (N-x), and it
        #    must not count against NumVersionsToKeep. It is NOT off-limits to
        #    assignment migration - step 6 re-adds it to that sweep explicitly. The
        #    anchor is pinned from the oldest surviving version, which on the first run
        #    under this model is usually the previously-CURRENT app, so it carries that
        #    version's live assignments. Dropping it here and never re-adding it
        #    stranded them on an app that is never renamed, never pruned and never
        #    revisited.
        if ($Anchor) {
            $OtherApps = @($OtherApps | Where-Object id -ne $Anchor.id)
        }

        # 3. Same-version collision - a prior run (or -Force) can leave an app at
        #    the same version. Those apps are held aside: their assignments and
        #    dependencies still migrate in step 6, and they are deleted with the
        #    rest of the prunable versions in step 8. They are never renamed to
        #    (N-x) or superseded, because their version is not actually older.
        $SameVersionApps = @($OtherApps | Where-Object displayVersion -eq $CurrentApp.displayVersion)
        $SameVersionIds = @($SameVersionApps | Select-Object -ExpandProperty id)
        if ($SameVersionApps) {
            Write-Log "Found $($SameVersionApps.Count) existing app(s) already at version $($CurrentApp.displayVersion); they will be retired after their assignments are migrated."
            $OtherApps = @($OtherApps | Where-Object displayVersion -ne $CurrentApp.displayVersion)
        }

        # 4. Compute kept vs prunable. `NumVersionsToKeep` includes the newly
        #    uploaded app, so we keep (NumVersionsToKeep - 1) of the older ones.
        $Sorted = @($OtherApps | Sort-Object @{Expression = {[VersionPro]$_.displayVersion}; Descending = $true})
        $ToKeepCount = [Math]::Max(0, $Script:NumVersionsToKeep - 1)
        $ToKeep = @()
        $ToPrune = @()
        if ($Sorted.Count -gt 0) {
            $ToKeep  = @($Sorted | Select-Object -First $ToKeepCount)
            $ToPrune = @($Sorted | Select-Object -Skip $ToKeepCount)
        }
        # Duplicates of the current version always get retired, regardless of retention.
        $ToPrune = @($SameVersionApps) + @($ToPrune)
        $ToPruneIds = @($ToPrune | Select-Object -ExpandProperty id)

        # 5. Rename kept apps to (N-1), (N-2), ...
        for ($i = 0; $i -lt $ToKeep.Count; $i++) {
            $targetName = "$($Script:DisplayName) (N-$($i + 1))"
            if ($ToKeep[$i].DisplayName -ne $targetName) {
                try {
                    Set-IntuneWin32App -Id $ToKeep[$i].Id -DisplayName $targetName | Out-Null
                    Write-Log "Renamed $($ToKeep[$i].DisplayName) -> $targetName"
                } catch {
                    Write-Log "ERROR: Failed to rename $($ToKeep[$i].Id) to $targetName : $_"
                }
            }
        }

        # 6. Intent-based assignment handling.
        #    - Required: MOVE from older versions to newest, as Yardstick has always
        #      done - required deployments install unconditionally, so consolidating
        #      them on a single app is correct.
        #    - Available: COPY to the newest and leave the source assignment in
        #      place, but only when the source will actually end up superseded.
        #      Intune builds the auto-update component on the device when a user
        #      installs from Company Portal, and documents that "any application
        #      assignment changes delete the component responsible for auto-updating
        #      the app" - so removing the available assignment from a superseded
        #      version would break auto-update for exactly the devices we want to
        #      update. Where supersedence will NOT be attached, copying would just
        #      leave a duplicate Company Portal listing with no update path, so we
        #      fall back to moving.
        #    Dependencies migrate exactly once per source app (during the required
        #    pass); the available pass runs with -SkipDependencies to avoid
        #    double-processing them. Dependencies that point at a version being
        #    pruned this run are dropped rather than carried onto the new app - the
        #    fresh link would block that version's own deletion in step 8.
        #
        #    The Ω DETECT - anchor is swept here too, on every run, even though steps 5
        #    and 8 leave it alone. Step 2 removed it from $OtherApps to keep it out of
        #    the (N-x) numbering and the retention count, not because its assignments
        #    are special - and excluding it from this loop is what stranded the
        #    assignments of any app whose only prior version got pinned.
        #    It is appended LAST so a (N-x) version wins any duplicate-target race:
        #    Intune rejects the second add for a target with "already exists", so the
        #    first assignment to land is the one whose notification and install-time
        #    settings survive, and a kept version's schedule is more current than the
        #    anchor's. The generic -CopyOnly:$willBeSuperseded rule below resolves
        #    correctly for it without a special case - the anchor is never in $ToPrune,
        #    so $isDoomed is false and the flag collapses to $Script:Supersedence, which
        #    is exactly "step 9 will supersede this app" (line 1222). With supersedence
        #    off there is no auto-update component to protect and no update path, so
        #    falling back to a move is right there too.
        #
        #    Dependencies are skipped for the anchor. Unlike an (N-x) version there is
        #    no deletion to unblock - the anchor is never pruned, so a dependent app
        #    pointing at it is not holding anything hostage. And its child dependency
        #    set is frozen at whatever the recipe declared when it was pinned;
        #    Add-IntuneWin32AppDependency replaces the target's whole set from a merged
        #    list, so migrating it would re-merge that stale set onto the current app on
        #    every run, resurrecting dependencies the recipe has since dropped.
        $allOlder = @($ToKeep) + @($ToPrune)
        # $Anchor can be the raw pre-rename candidate (see step 1) when Intune's lookup
        # is still stale. Only .id and .DisplayName are read, and both objects come from
        # Get-IntuneWin32App, so the shapes are interchangeable. The id guard is cheap
        # insurance: From -eq To would add the assignment, read it back as present, then
        # delete it off the app it was just confirmed on.
        if ($Anchor -and ($Anchor.id -ne $CurrentApp.id)) { $allOlder += $Anchor }
        foreach ($old in $allOlder) {
            $isDoomed = $ToPruneIds -contains $old.id
            $isAnchor = ($null -ne $Anchor) -and ($old.id -eq $Anchor.id)
            try {
                Move-AssignmentsAndDependencies -From $old -To $CurrentApp `
                    -AvailableDateOffset $Script:AvailableDateOffset `
                    -DeadlineDateOffset $Script:DeadlineDateOffset `
                    -IntentFilter 'required' `
                    -AllowDependentLinkUpdates $Script:AllowDependentLinkUpdates `
                    -DependentLinkOptions $DependentLinkOptions `
                    -DependentUpdateStatus $DependentUpdateStatus `
                    -ProtectedSourceIds $ProtectedDependencyAppIds `
                    -ExcludeDependencyTargetIds $ToPruneIds `
                    -SkipDependencies:$isAnchor
            } catch {
                Write-Log "ERROR: Failed required-intent move from $($old.DisplayName): $_"
            }

            $availOnOld = @(Get-IntuneWin32AppAssignment -Id $old.id | Where-Object Intent -eq 'available')
            if ($availOnOld.Count -gt 0) {
                # A pruned app is never a supersedence target - step 9 only supersedes
                # the versions that survive - so there is no auto-update path to
                # protect and copying would just strand the assignment on an app
                # about to be deleted. Move it instead.
                $willBeSuperseded = $Script:Supersedence -and (-not $isDoomed)
                try {
                    Move-AssignmentsAndDependencies -From $old -To $CurrentApp `
                        -AvailableDateOffset $Script:AvailableDateOffset `
                        -DeadlineDateOffset $Script:DeadlineDateOffset `
                        -ProtectedSourceIds $ProtectedDependencyAppIds `
                        -IntentFilter 'available' -CopyOnly:$willBeSuperseded -SkipDependencies
                } catch {
                    Write-Log "ERROR: Failed available-intent migration from $($old.DisplayName): $_"
                }
            }
        }

        # 7. Default deployments (unchanged behavior). Adds default groups as
        #    available-intent assignments if not already present.
        if ($Script:DefaultDeploymentGroups) {
            $ID = $CurrentApp.Id
            $CurrentlyDeployedIDs = (Get-IntuneWin32AppAssignment -Id $ID).GroupID
            foreach ($DeploymentGroupID in $Script:DefaultDeploymentGroups) {
                if (!($CurrentlyDeployedIDs -Contains $DeploymentGroupID)) {
                    Write-Log "Deploying $ID to $DeploymentGroupID because it is in the default list"
                    Add-IntuneWin32AppAssignmentGroup -Include -ID $ID -GroupID $DeploymentGroupID -Intent "available" -Notification "hideAll" | Out-Null
                }
            }
        }

        # 8. Prune before wiring supersedence. Intune refuses to delete an app that
        #    still participates in a supersedence relationship, so removing the
        #    expired versions first keeps the graph pointing only at apps that
        #    survive this run.
        $Pruned = @()
        if (!$NoDelete) {
            foreach ($old in $ToPrune) {
                if ($ProtectedDependencyAppIds.Contains($old.id)) {
                    Write-Log "Skipping removal of $($old.displayName) because dependent applications are still targeting this version."
                    continue
                }
                try {
                    Remove-YardstickApp -App $old
                    $Pruned += $old.id
                } catch {
                    Write-Log "ERROR: Failed to remove $($old.DisplayName) ($($old.id)): $_"
                }
            }
        }

        # 9. Supersedence. The newest version becomes the single superseding parent
        #    for every surviving older version plus the Ω DETECT - anchor. The anchor
        #    is forced to "Update" because a "Replace" against its deliberately wide
        #    detection rule would uninstall the app from every device that has any
        #    version of it installed.
        $supersededCount = 0
        if ($Script:Supersedence) {
            $SupersedenceTargets = @($ToKeep)
            # Anything we intended to prune but could not (protected by a dependent,
            # or the delete was refused) still needs to be superseded so devices on
            # it move forward. Same-version leftovers are excluded: superseding an
            # app that carries the identical version is meaningless.
            $SupersedenceTargets += @($ToPrune | Where-Object { ($Pruned -notcontains $_.id) -and ($SameVersionIds -notcontains $_.id) })
            if ($Anchor) { $SupersedenceTargets += $Anchor }
            $Type = if ($Script:UninstallPreviousVersion) { 'Replace' } else { 'Update' }
            try {
                $supersededCount = Set-YardstickSupersedence -NewApp $CurrentApp -SupersededApps $SupersedenceTargets -Type $Type `
                    -UpdateOnlyIds @(if ($Anchor) { $Anchor.id })
            } catch {
                Write-Log "ERROR: Failed to configure supersedence for $($CurrentApp.DisplayName): $_"
            }
        } else {
            Write-Log "Supersedence disabled for this recipe - skipping"
        }

        # 10. Native auto-update. Intune only honours this on available-intent
        #     assignments, so required-only recipes are a no-op here by design.
        $autoUpdateApplied = 0
        if ($Script:AutoUpdateOnAssignment) {
            try {
                $autoUpdateApplied = Set-AssignmentAutoUpdate -AppId $CurrentApp.Id -Enabled $true -IntentFilter 'available' `
                    -SkipGroupIds $Script:GroupsSkipAutoUpdates
            } catch {
                Write-Log "ERROR: Failed to enable auto-update on assignments for $($CurrentApp.DisplayName): $_"
            }
        }

        $ActionPerformed = if ($Force) { "Force Updated" } elseif ($Repair) { "Repaired" } else { "Updated" }
        $autoStatusParts = @()
        if ($supersededCount -gt 0)   { $autoStatusParts += "supersedes $supersededCount" }
        if ($autoUpdateApplied -gt 0) { $autoStatusParts += "auto-update on $autoUpdateApplied assignment(s)" }
        if ($anchorStatus)            { $autoStatusParts += "anchor: $anchorStatus" }
        $autoStatus = if ($autoStatusParts.Count -gt 0) { $autoStatusParts -join ', ' } else { $null }
        Add-SuccessfulApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -Action $ActionPerformed -Dependents $DependentUpdateStatus -AutoUpdateStatus $autoStatus
        Write-Log "Updates complete for $Script:DisplayName"

    } catch {
        # Catch any unexpected errors during processing
        Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Unexpected error during processing: $_" -FailureStage "General Processing"
        Write-Log "ERROR: Unexpected error processing $AppId_Processing : $_"
        if (-not((Get-Location).Path -eq $PSScriptRoot)) {
            Set-Location $PSScriptRoot
        }
    }
    # Run the post-run script
    if ($Script:PostRunScript) {
        Write-Log "Running post run script..."
        try {
            Invoke-Command -ScriptBlock $Script:PostRunScript -NoNewScope
            Write-Log "Post run script ran successfully."
        } catch {
            Add-FailedApplication -ApplicationId $AppId_Processing -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Error while running post run PowerShell script: $_" -FailureStage "Post-Run Script"
            Write-Error "Error while running post run PowerShell script: $_"
        }
    }

    # Drain this app's backup before moving on, so at most one copy is ever in
    # flight and an abrupt exit in the next iteration cannot orphan it. The
    # thread still ran concurrently with all the supersedence and assignment work
    # above, which is where the time actually goes.
    Wait-YardstickBackup -TimeoutSeconds 600 | Out-Null
}

# Final drain. This has to happen before the email report is built, because
# Wait-YardstickBackup is what records each app's backup status on it - and
# because the email block can return early, skipping Invoke-Cleanup entirely.
Wait-YardstickBackup -TimeoutSeconds 600 | Out-Null

# Send email report if enabled
if (-not $NoEmail) {
    try {
        # Select email delivery method based on preferences
        $emailMethod = if ($null -ne $Prefs.emailDeliveryMethod) { $Prefs.emailDeliveryMethod } else { "outlook" }
        
        switch ($emailMethod.ToLower()) {
            "mailkit" {
                Write-Log "ERROR: MailKit email delivery method is not yet implemented. Set emailDeliveryMethod to 'outlook' in preferences.yaml or remove it to use the default."
                throw "MailKit email delivery method is not yet implemented."
            }
            "outlook" {
                # If prompt before sending is enabled, ask the user for confirmation
                if ($Prefs.emailPromptBeforeSending) {
                    # play notification sound if specified
                    if ($Prefs.emailNotificationSoundFile -and (Test-Path $Prefs.emailNotificationSoundFile)) {
                        try {
                            $sound = New-Object System.Media.SoundPlayer($Prefs.emailNotificationSoundFile)
                            $sound.Play()
                        } catch {
                            Write-Log "WARNING: Failed to play notification sound: $_"
                        }
                    }
                    $confirmation = Read-Host "Do you want to send the email report now? (Y/N)"
                    if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
                        Write-Log "Email report sending skipped by user."
                        return
                    }
                }
                Write-Log "Sending email report using Outlook"
                Send-YardstickEmailReport -Preferences $Prefs -RunParameters $RunParameters
            }
            default {
                Write-Log "WARNING: Unknown email delivery method '$emailMethod'. Defaulting to Outlook."
                Send-YardstickEmailReport -Preferences $Prefs -RunParameters $RunParameters
            }
        }
    } catch {
        Write-Log "WARNING: Failed to send email report: $_"
    }
}

# Clean up
Invoke-Cleanup

# Return to the original directory
Set-Location $PSScriptRoot
