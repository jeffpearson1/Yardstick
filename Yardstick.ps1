<#
.DESCRIPTION
    Yardstick is a PowerShell script designed to automate the process of updating and managing Win32 applications in Microsoft Intune. It allows users to specify applications to update, either individually or in groups, and handles the downloading, packaging, and deployment of these applications.

.SYNOPSIS
    Adds and Updates Win32 applications in Microsoft Intune.

.PARAMETER ApplicationId
    The ID of the application to update. This parameter is used when updating a single application.
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
    .\Yardstick.ps1 -Group "ExampleGroup" -NoInteractive -Force
    This command updates all applications in the group "ExampleGroup", and forces the update.

.EXAMPLE
    .\Yardstick.ps1 -All -NoInteractive -Force -NoDelete
    This command updates all applications in the repository, excluding interactive applications, forcing the update and preventing the deletion of old versions.
#>

using module .\Modules\Custom\AdobeDownloader.psm1

param (
    [Alias("AppId")]
    [parameter(ParameterSetName="SingleApp")]
    [String] $ApplicationId = "None",

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

# Constants
$Global:LogLocation = "$PSScriptRoot"
$Global:LogFile = "YLog.log"

# Modules required:
# powershell-yaml
# IntuneWin32App
# Selenium
Import-Module powershell-yaml -Scope Local
Import-Module IntuneWin32App -Scope Local
Import-Module Selenium -Scope Local
Import-Module TUN.CredentialManager -Scope Local
Import-Module ${PSScriptRoot}\Modules\YardstickSupport.psm1 -Scope Global -Force

$CustomModules = Get-ChildItem -Path $PSScriptRoot\Modules\Custom\*.psm1
foreach ($Module in $CustomModules) {
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
if ("None" -eq $ApplicationId -and (($All -eq $false) -and (!$Group))) {
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



function Set-ScriptVariables {
    <#
    .SYNOPSIS
    Sets script-scoped variables from application parameters and global preferences.
    
    .DESCRIPTION
    This function centralizes the logic for setting script variables, using application-specific
    parameters when available and falling back to global preferences with null coalescing.
    
    .PARAMETER Parameters
    Hashtable containing application-specific parameters from the YAML recipe.
    
    .PARAMETER Preferences
    Hashtable containing global preferences from the preferences.yaml file.
    #>
    param(
        [hashtable]$Parameters,
        [hashtable]$Preferences
    )
    
    # Import Folder Locations
    $Script:Temp = $Preferences.Temp
    $Script:BuildSpace = $Preferences.Buildspace
    $Script:Scripts = $Preferences.Scripts
    $Script:Published = $Preferences.Published
    $Script:Recipes = $Preferences.Recipes
    $Script:Icons = $Preferences.Icons
    $Script:Tools = $Preferences.Tools
    $Script:Secrets = $Preferences.Secrets
    $Script:CustomModules = $Preferences.CustomModules

    # Import Intune Connection Settings
    $Global:TenantID = $Preferences.TenantID
    $Global:ClientID = $Preferences.ClientID
    $Global:ClientSecret = $Preferences.ClientSecret

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
    
    # Use parameters or fall back to preferences with null coalescing
    $Script:ScopeTags = $Parameters.scopeTags ?? $Preferences.defaultScopeTags
    $Script:Owner = $Parameters.owner ?? $Preferences.defaultOwner
    $Script:MaximumInstallationTimeInMinutes = $Parameters.maximumInstallationTimeInMinutes ?? $Preferences.defaultMaximumInstallationTimeInMinutes
    $Script:MinOSVersion = $Parameters.minOSVersion ?? $Preferences.defaultMinOSVersion
    $Script:InstallExperience = $Parameters.installExperience ?? $Preferences.defaultInstallExperience
    $Script:RestartBehavior = $Parameters.restartBehavior ?? $Preferences.defaultRestartBehavior
    $Script:AvailableGroups = $Parameters.availableGroups ?? $Preferences.defaultAvailableGroups
    $Script:RequiredGroups = $Parameters.requiredGroups ?? $Preferences.defaultRequiredGroups
    $Script:DefaultDeploymentGroups = $Parameters.defaultDeploymentGroups ?? $Preferences.defaultDeploymentGroups
    $Script:AllowUserUninstall = $Parameters.allowUserUninstall ?? $Preferences.defaultAllowUserUninstall
    $Script:Is32BitApp = $Parameters.is32BitApp ?? $Preferences.defaultIs32BitApp
    $Script:Architecture = $Parameters.architecture ?? $Preferences.defaultArchitecture
    $Script:DeadlineDateOffset = $Parameters.deadlineDateOffset ?? $Preferences.defaultDeadlineDateOffset
    $Script:AvailableDateOffset = $Parameters.availableDateOffset ?? $Preferences.defaultAvailableDateOffset
    $Script:AllowDependentLinkUpdates = $Parameters.allowDependentLinkUpdates ?? $Preferences.defaultAllowDependentLinkUpdates
    
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
    $Script:DetectionScriptFileExtension = $Parameters.detectionScriptFileExtension ?? $Preferences.defaultDetectionScriptFileExtension
    $Script:DetectionScriptRunAs32Bit = $Parameters.detectionScriptRunAs32Bit ?? $Preferences.defaultdetectionScriptRunAs32Bit
    $Script:DetectionScriptEnforceSignatureCheck = $Parameters.detectionScriptEnforceSignatureCheck ?? $Preferences.defaultdetectionScriptEnforceSignatureCheck
    $Script:DependentLinkUpdateEnabled = $Parameters.dependentLinkUpdateEnabled ?? ($Preferences.dependentLinkUpdateEnabled ?? $true)
    $Script:DependentLinkUpdateRetryCount = $Parameters.dependentLinkUpdateRetryCount ?? ($Preferences.dependentLinkUpdateRetryCount ?? 3)
    $Script:DependentLinkUpdateRetryDelaySeconds = $Parameters.dependentLinkUpdateRetryDelaySeconds ?? ($Preferences.dependentLinkUpdateRetryDelaySeconds ?? 5)
    $Script:DependentLinkUpdateTimeoutSeconds = $Parameters.dependentLinkUpdateTimeoutSeconds ?? ($Preferences.dependentLinkUpdateTimeoutSeconds ?? 60)
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
        $Script:NumVersionsToKeep = $Parameters.numVersionsToKeep ?? $Preferences.defaultNumVersionsToKeep
    }

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

# Import Intune Connection Settings
$Global:TenantID = $Prefs.TenantID
$Global:ClientID = $Prefs.ClientID
$Global:ClientSecret = $Prefs.ClientSecret
$Global:ScriptRoot = $PSScriptRoot

# Validate ApplicationId parameter
if ($ApplicationId -ne "None") {
    $MatchingApps = Get-ChildItem $Recipes | Where-Object Name -ne 'Disabled' | Get-ChildItem -File | Where-Object Name -match "$ApplicationId\.ya{0,1}ml"
    if ($null -eq $MatchingApps) {
        Write-Log "ERROR: Application $ApplicationId not found. (Excluding Disabled Folder)"
        exit 1
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
        [string]$ApplicationId,
        [string]$Group,
        [switch]$All,
        [switch]$NoInteractive
    )
    
    $applications = [System.Collections.Generic.List[PSObject]]::new()
    
    if ($All) {
        $folderFilter = if ($NoInteractive) { 
            { $_.Name -ne 'Disabled' -and $_.Name -ne 'Interactive' }
        } else { 
            { $_.Name -ne 'Disabled' }
        }
        
        $ApplicationFullNames = (Get-ChildItem $Recipes | Where-Object $folderFilter | Get-ChildItem -File).Name
        foreach ($Application in $ApplicationFullNames) {
            $applications.Add($($Application -split "\.ya{0,1}ml")[0]) | Out-Null
        }
    } elseif ($Group) {
        try {
            $groupFile = Get-Content $PSScriptRoot\RecipeGroups.yaml | ConvertFrom-Yaml
            $groupFile[$Group] | ForEach-Object {
                $applications.Add($_) | Out-Null
            }
        } catch {
            Write-Log "ERROR: There was an issue importing the application group! Exiting."
            exit 3
        }
    } else {
        $applications.Add($ApplicationId) | Out-Null
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
if ($ApplicationId -ne "None") { $RunParametersArray += "-ApplicationId $ApplicationId" }
if ($Group) { $RunParametersArray += "-Group $Group" }
if ($All) { $RunParametersArray += "-All" }
if ($NoInteractive) { $RunParametersArray += "-NoInteractive" }
if ($Force) { $RunParametersArray += "-Force" }
if ($NoDelete) { $RunParametersArray += "-NoDelete" }
if ($Repair) { $RunParametersArray += "-Repair" }
$RunParameters = $RunParametersArray -join " "

# Main processing loop
foreach ($ApplicationId in $Applications) {
    Write-Log "Starting update for $ApplicationId..."
    
    # Initialize variables for tracking
    $CurrentDisplayName = $ApplicationId
    
    try {
        # Refresh token if necessary
        Connect-AutoMSIntuneGraph
        
        # Clear the temp file
        Write-Log "Clearing the temp directory..."
        Get-ChildItem $Temp -Exclude ".gitkeep" -Recurse | Remove-Item -Recurse -Force

        # Open the YAML file and collect all necessary attributes
        try {
            $AppName = (Get-ChildItem $Recipes -Force -Recurse | Where-Object Name -ne 'Disabled' | Get-ChildItem -File -Recurse | Where-Object Name -match "^$ApplicationId\.ya{0,1}ml")[0].FullName
            $Parameters = Get-Content "$AppName" | ConvertFrom-Yaml
        } catch {
            Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $ApplicationId -Version "Unknown" -ErrorMessage "Unable to open parameters file for $ApplicationId" -FailureStage "Configuration"
            Write-Error "Unable to open parameters file for $ApplicationId"
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
            # Correct any naming discrepancies before we continue
            # Rename any apps if they are named incorrectly
            $CurrentApps = Get-SameAppAllVersions $Script:DisplayName
            for ($i = 1; $i -lt $CurrentApps.Count; $i++) {
                if ($CurrentApps[$i].DisplayName -ne "$($Script:DisplayName) (N-$i)") {
                    Write-Log "Setting name for $($Script:DisplayName) (N-$i)"
                    Set-IntuneWin32App -Id $CurrentApps[$i].Id -DisplayName "$($Script:DisplayName) (N-$i)"
                }
            }
        }


        # Run the pre-download script
        if ($Script:PreDownloadScript) {
            Write-Log "Running pre-download script..."
            try {
                Invoke-Command -ScriptBlock $Script:PreDownloadScript -NoNewScope
                Write-Log "Pre-download script ran successfully."
            } catch {
                Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Error while running pre-download PowerShell script: $_" -FailureStage "Pre-Download Script"
                Write-Error "Error while running pre-download PowerShell script"
                continue
            }
        } else {
            Write-Log "Skipping Pre-download script"
        }
        # Check if there is an up-to-date version in the repo already
        Write-Log "Checking if $($Script:DisplayName) $($Script:Version) is a new version..."
        $ExistingVersions = Get-SameAppAllVersions $Script:DisplayName
        
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
        } elseif (Get-VersionLocked -Version $Script:Version -VersionLock $Script:VersionLock) {
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
            Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "URL is empty - cannot continue" -FailureStage "Download"
            Write-Error "URL is empty - cannot continue."
            continue
        }
        
        if ($Script:DownloadScript) {
            Push-Location $BuildSpace\$Script:Id\$Script:Version
            try {
                Invoke-Command -ScriptBlock $Script:DownloadScript -NoNewScope
                Write-Log "Download script ran successfully."
            } catch {
                Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Error while running download PowerShell script: $_" -FailureStage "Download Script"
                Write-Error "Error while running download PowerShell script: $_"
                Write-Error "Script Contents: $Script:DownloadScript"
                continue
            }
            Pop-Location
        } else {
            try {
                Start-BitsTransfer -Source $Script:Url -Destination $BuildSpace\$Script:Id\$Script:Version\$Script:FileName
                Write-Log "File downloaded successfully using BITS transfer."
            } catch {
                Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Error downloading file: $_" -FailureStage "Download"
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
                Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Error while running post download PowerShell script: $_" -FailureStage "Post-Download Script"
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
                Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Could not determine MSI Product Code from $Script:FileName" -FailureStage "MSI Product Code Retrieval"
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
            Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "One or more placeholder strings were not replaced in scripts." -FailureStage "Placeholder Replacement"
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
            Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Icon file $($Script:IconFile) not found in Icons folder." -FailureStage "Icon Retrieval"
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
        } catch {
            Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Failed to upload application to Intune: $_" -FailureStage "Intune Upload"
            Write-Error "Failed to upload application to Intune: $_"
            continue
        }


        ###################################################
        # MIGRATE OLD DEPLOYMENTS
        ###################################################

        # Check if any existing applications have the same version so we can delete them
        $ToRemove = $ExistingVersions | where-object displayVersion -eq $Script:Version
        if ($ToRemove) {
            # Remove conflicting versions
            Write-Log "Removing conflicting versions"
            $CurrentApp = Get-IntuneWin32App -Id $Win32App.id
            foreach ($RemoveApp in $ToRemove) {
                Write-Log "Moving assignments before removal..."
                Move-AssignmentsAndDependencies -From $RemoveApp -To $CurrentApp -AvailableDateOffset $Script:AvailableDateOffset -DeadlineDateOffset $Script:DeadlineDateOffset -AllowDependentLinkUpdates:$Script:AllowDependentLinkUpdates -DependentLinkOptions $DependentLinkOptions -DependentUpdateStatus $DependentUpdateStatus -ProtectedSourceIds $ProtectedDependencyAppIds
                if ($ProtectedDependencyAppIds.Contains($RemoveApp.id)) {
                    Write-Log "Skipping removal of $($RemoveApp.DisplayName) because dependent applications are still targeting this version."
                }
                else {
                    Write-Log "Removing App with ID $($RemoveApp.id)"
                    Remove-IntuneWin32App -Id $RemoveApp.id
                }
            }
        }

        # Define the current version, the version that is one older but shares the same name, and all the ones older than that
        Write-Log "Updating local application manifest..."
        Start-Sleep -Seconds 4
        try {
            $AllMatchingApps = Get-SameAppAllVersions $Script:DisplayName
        } catch {
            Write-Log "There was an error fetching information about existing applications. Exiting"
            Exit 4
        }
        $CurrentApp = $AllMatchingApps[0]
        $AllOldApps = $AllMatchingApps[1..$($AllMatchingApps.count - 1)]
        # For any apps that share the same version as the current one, move all their deployments to the current one and remove the rest from the list
        $SameVersionApps = $AllOldApps | Where-Object displayVersion -eq $CurrentApp.displayVersion
        if ($SameVersionApps) {
            foreach ($App in $SameVersionApps) {
                Write-Log "Moving assignments from $($App.id) to $($CurrentApp.id)"
                Move-AssignmentsAndDependencies -From $App -To $CurrentApp -AvailableDateOffset $Script:AvailableDateOffset -DeadlineDateOffset $Script:DeadlineDateOffset -AllowDependentLinkUpdates:$Script:AllowDependentLinkUpdates -DependentLinkOptions $DependentLinkOptions -DependentUpdateStatus $DependentUpdateStatus -ProtectedSourceIds $ProtectedDependencyAppIds
            }
            $AllOldApps = $AllOldApps | Where-Object displayVersion -ne $CurrentApp.displayVersion
        }

        
        # All apps that will be assigned N-1 - these currently do not have the N-<version> suffix
        $NMinusOneApps = $AllOldApps | Where-Object displayName -eq $Script:DisplayName
        # All apps that will be assigned N-2 and older - these currently have the N-<version> suffix
        $NMinusTwoAndOlderApps = $AllOldApps | Where-Object displayName -ne $Script:DisplayName

        # Start with the N-1 app first and move all its deployments to the newest one
        if ($NMinusOneApps) {
            foreach ($NMinusOneApp in $NMinusOneApps) {
                if ($CurrentApp) {
                    Write-Log "Moving assignments from $($NMinusOneApp.id) to $($CurrentApp.id)"
                    Move-AssignmentsAndDependencies -From $NMinusOneApp -To $CurrentApp -AvailableDateOffset $Script:AvailableDateOffset -DeadlineDateOffset $Script:DeadlineDateOffset -AllowDependentLinkUpdates:$Script:AllowDependentLinkUpdates -DependentLinkOptions $DependentLinkOptions -DependentUpdateStatus $DependentUpdateStatus -ProtectedSourceIds $ProtectedDependencyAppIds
                }
                else {
                    Write-Log "There was an error fetching information about the current application. Exiting"
                    Exit 5
                }
            }
        }
        if ($($NMinusOneApps.count) -gt 1) {
            $NewestNMinusOneApp = ($NMinusOneApps | Sort-Object createdDateTime -Descending)[0]
        }
        for ($i = 0; $i -lt $NMinusTwoAndOlderApps.count; $i++) {
            # Move all the deployments up one number
            if ($i -eq 0) {
                # Move the app assignments in the 0 position to the nminusoneapp
                if ($NMinusTwoAndOlderApps[$i] -and $NewestNMinusOneApp) {
                    Move-AssignmentsAndDependencies -From $NMinusTwoAndOlderApps[$i] -To $NewestNMinusOneApp -AvailableDateOffset $Script:AvailableDateOffset -DeadlineDateOffset $Script:DeadlineDateOffset
                }
            }
            else {
                if ($NMinusTwoAndOlderApps[$i] -and $NMinusTwoAndOlderApps[$i - 1]) {
                    Move-AssignmentsAndDependencies -From $NMinusTwoAndOlderApps[$i] -To $NMinusTwoAndOlderApps[$i - 1] -AvailableDateOffset $Script:AvailableDateOffset -DeadlineDateOffset $Script:DeadlineDateOffset
                }
            }
        }

        # Add the default deployments if they're not already there from the migration
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

        # Rename all the old applications to have the appropriate N-<versions behind> in them
        for ($i = 1; $i -lt $AllMatchingApps.count; $i++) {
            Set-IntuneWin32App -Id $AllMatchingApps[$i].Id -DisplayName "$($Script:DisplayName) (N-$i)"
            if ($?) {
                Write-Log "Successfully set Application Name to $($Script:DisplayName) (N-$i)"
            } else {
                Write-Log "ERROR: Failed to update display name to $($Script:DisplayName) (N-$i)"
            }
        }

        # Remove all the old versions 
        if (!$NoDelete) {
            for ($i = $Script:NumVersionsToKeep - 1; $i -lt $AllOldApps.count; $i++) {
                if ($ProtectedDependencyAppIds.Contains($AllOldApps[$i].id)) {
                    Write-Log "Skipping removal of $($AllOldApps[$i].displayName) because dependent applications are still targeting this version."
                    continue
                }
                Write-Log "Removing old app with id $($AllOldApps[$i].id)"
                Remove-IntuneWin32App -Id $AllOldApps[$i].id
            }
            
            # If we get here, the application was processed successfully
            $ActionPerformed = if ($Force) { "Force Updated" } elseif ($Repair) { "Repaired" } else { "Updated" }
            Add-SuccessfulApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -Action $ActionPerformed -Dependents $DependentUpdateStatus
            Write-Log "Updates complete for $Script:DisplayName"
        }

    } catch {
        # Catch any unexpected errors during processing
        Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Unexpected error during processing: $_" -FailureStage "General Processing"
        Write-Log "ERROR: Unexpected error processing $ApplicationId : $_"
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
            Add-FailedApplication -ApplicationId $ApplicationId -DisplayName $CurrentDisplayName -Version $Script:Version -ErrorMessage "Error while running post run PowerShell script: $_" -FailureStage "Post-Run Script"
            Write-Error "Error while running post run PowerShell script: $_"
        }
    }
}

# Send email report if enabled
if (-not $NoEmail) {
    try {
        # Select email delivery method based on preferences
        $emailMethod = $Prefs.emailDeliveryMethod ?? "outlook"
        
        switch ($emailMethod.ToLower()) {
            "mailkit" {
                Write-Log "Sending email report using PoshMailKit"
                Send-YardstickEmailReportMailKit -Preferences $Prefs -RunParameters $RunParameters
            }
            "outlook" {
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
