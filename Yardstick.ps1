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
    [Switch] $Repair
)

# Constants
$Global:LOG_LOCATION = "$PSScriptRoot"
$Global:LOG_FILE = "YLog.log"

# Modules required:
# powershell-yaml
# IntuneWin32App
# Selenium
Import-Module powershell-yaml -Scope Local
Import-Module IntuneWin32App -Scope Local
Import-Module Selenium -Scope Local
Import-Module TUN.CredentialManager -Scope Local
Import-Module $PSScriptRoot\Modules\YardstickSupport.psm1 -Scope Global -Force

$CustomModules = Get-ChildItem -Path $PSScriptRoot\Modules\Custom\*.psm1
foreach ($Module in $CustomModules) {
    # Force allows us to reload them for testing
    Import-Module "$($Module.FullName)" -Scope Global -Force
}

# Initialize script variables
$Script:Applications = [System.Collections.Generic.List[PSObject]]::new()

# So we can pop at the end
Push-Location $PSScriptRoot

# Initialize the log file
Write-Log -Init

# Validate parameters
if ("None" -eq $ApplicationId -and (($All -eq $false) -and (!$Group))) {
    Write-Log "Please provide parameter -ApplicationId"
    exit 1
}

# Import preferences file:
try {
    $prefs = Get-Content $PSScriptRoot\Preferences.yaml | ConvertFrom-Yaml
}
catch {
    Write-Error "Unable to open preferences.yaml!"
    exit 1
}

# Function to set script variables from parameters and preferences
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
    $Script:TEMP = $Preferences.Temp
    $Script:BUILDSPACE = $Preferences.Buildspace
    $Script:SCRIPTS = $Preferences.Scripts
    $Script:PUBLISHED = $Preferences.Published
    $Script:RECIPES = $Preferences.Recipes
    $Script:ICONS = $Preferences.Icons
    $Script:TOOLS = $Preferences.Tools
    $Script:SECRETS = $Preferences.Secrets

    # Import Intune Connection Settings
    $Global:TENANT_ID = $Preferences.TenantID
    $Global:CLIENT_ID = $Preferences.ClientID
    $Global:CLIENT_SECRET = $Preferences.ClientSecret

    # Set all variables from default preferences and the application recipe
    $Script:url = if ($Parameters.urlRedirects -eq $true) {Get-RedirectedUrl $Parameters.url} else {$Parameters.url}
    $Script:id = $Parameters.id
    $Script:version = $Parameters.version
    $Script:fileDetectionVersion = $Parameters.fileDetectionVersion
    $Script:displayName = $Parameters.displayName
    $Script:displayVersion = $Parameters.displayVersion
    $Script:fileName = $Parameters.fileName
    $Script:fileDetectionPath = $Parameters.fileDetectionPath
    $Script:preDownloadScript = if ($Parameters.preDownloadScript) { [ScriptBlock]::Create($Parameters.preDownloadScript)}
    $Script:postDownloadScript = if ($Parameters.postDownloadScript) { [ScriptBlock]::Create($Parameters.postDownloadScript) }
    $Script:downloadScript = if ($Parameters.downloadScript) { [ScriptBlock]::Create($Parameters.downloadScript) }
    $Script:installScript = $Parameters.installScript
    $Script:uninstallScript = $Parameters.uninstallScript
    
    # Use parameters or fall back to preferences with null coalescing
    $Script:scopeTags = $Parameters.scopeTags ?? $Preferences.defaultScopeTags
    $Script:owner = $Parameters.owner ?? $Preferences.defaultOwner
    $Script:maximumInstallationTimeInMinutes = $Parameters.maximumInstallationTimeInMinutes ?? $Preferences.defaultMaximumInstallationTimeInMinutes
    $Script:minOSVersion = $Parameters.minOSVersion ?? $Preferences.defaultMinOSVersion
    $Script:installExperience = $Parameters.installExperience ?? $Preferences.defaultInstallExperience
    $Script:restartBehavior = $Parameters.restartBehavior ?? $Preferences.defaultRestartBehavior
    $Script:availableGroups = $Parameters.availableGroups ?? $Preferences.defaultAvailableGroups
    $Script:requiredGroups = $Parameters.requiredGroups ?? $Preferences.defaultRequiredGroups
    $Script:defaultDeploymentGroups = $Parameters.defaultDeploymentGroups ?? $Preferences.defaultDeploymentGroups
    $Script:allowUserUninstall = $Parameters.allowUserUninstall ?? $Preferences.defaultAllowUserUninstall
    $Script:is32BitApp = $Parameters.is32BitApp ?? $Preferences.defaultIs32BitApp
    $Script:architecture = $Parameters.architecture ?? $Preferences.defaultArchitecture
    $Script:deadlineDateOffset = $Parameters.deadlineDateOffset ?? $Preferences.defaultDeadlineDateOffset
    $Script:availableDateOffset = $Parameters.availableDateOffset ?? $Preferences.defaultAvailableDateOffset
    
    # Detection-related variables
    $Script:detectionType = $Parameters.detectionType
    $Script:fileDetectionVersion = $Parameters.fileDetectionVersion
    $Script:fileDetectionMethod = $Parameters.fileDetectionMethod
    $Script:fileDetectionName = $Parameters.fileDetectionName
    $Script:fileDetectionOperator = $Parameters.fileDetectionOperator
    $Script:fileDetectionDateTime = $Parameters.fileDetectionDateTime
    $Script:fileDetectionValue = $Parameters.fileDetectionValue
    $Script:registryDetectionMethod = $Parameters.registryDetectionMethod
    $Script:registryDetectionKey = $Parameters.registryDetectionKey
    $Script:registryDetectionValueName = $Parameters.registryDetectionValueName
    $Script:registryDetectionValue = $Parameters.registryDetectionValue
    $Script:registryDetectionOperator = $Parameters.registryDetectionOperator
    $Script:detectionScript = $Parameters.detectionScript
    $Script:detectionScriptFileExtension = $Parameters.detectionScriptFileExtension ?? $Preferences.defaultDetectionScriptFileExtension
    $Script:detectionScriptRunAs32Bit = $Parameters.detectionScriptRunAs32Bit ?? $Preferences.defaultdetectionScriptRunAs32Bit
    $Script:detectionScriptEnforceSignatureCheck = $Parameters.detectionScriptEnforceSignatureCheck ?? $Preferences.defaultdetectionScriptEnforceSignatureCheck
    
    # Additional variables
    $Script:iconFile = $Parameters.iconFile
    $Script:description = $Parameters.description
    $Script:publisher = $Parameters.publisher
    $Script:arm64FilterName = $Preferences.arm64FilterName
    $Script:amd64FilterName = $Preferences.amd64FilterName
    $Script:versionLock = $Parameters.versionLock
    
    # Handle version lock logic
    if ($Script:versionLock) {
        $Script:numVersionsToKeep = 1
        if($Parameters.numVersionsToKeep -gt 1) {
            Write-Log "Warning: Version lock is set, but numVersionsToKeep is set to $($Parameters.numVersionsToKeep). This will be ignored."
        }
    }
    else {
        $Script:numVersionsToKeep = $Parameters.numVersionsToKeep ?? $Preferences.defaultNumVersionsToKeep
    }
}

# Import Folder Locations from preferences
$Script:TEMP = $prefs.Temp
$Script:BUILDSPACE = $prefs.Buildspace
$Script:SCRIPTS = $prefs.Scripts
$Script:PUBLISHED = $prefs.Published
$Script:RECIPES = $prefs.Recipes
$Script:ICONS = $prefs.Icons
$Script:TOOLS = $prefs.Tools
$Script:SECRETS = $prefs.Secrets

# Import Intune Connection Settings
$Global:TENANT_ID = $prefs.TenantID
$Global:CLIENT_ID = $prefs.ClientID
$Global:CLIENT_SECRET = $prefs.ClientSecret

# Validate ApplicationId parameter
if ($ApplicationId -ne "None") {
    $MatchingApps = Get-ChildItem $RECIPES | Where-Object Name -ne 'Disabled' | Get-ChildItem -File | Where-Object Name -match "$ApplicationId\.ya{0,1}ml"
    if ($null -eq $MatchingApps) {
        Write-Log "ERROR: Application $ApplicationId not found. (Excluding Disabled Folder)"
        exit 1
    }
}

# Function to get applications based on parameters
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
        
        $ApplicationFullNames = (Get-ChildItem $RECIPES | Where-Object $folderFilter | Get-ChildItem -File).Name
        foreach ($Application in $ApplicationFullNames) {
            $applications.Add($($Application -split "\.ya{0,1}ml")[0]) | Out-Null
        }
    }
    elseif ($Group) {
        try {
            $groupFile = Get-Content $PSScriptRoot\RecipeGroups.yaml | ConvertFrom-Yaml
            $groupFile[$Group] | ForEach-Object {
                $applications.Add($_) | Out-Null
            }
        }
        catch {
            Write-Log "ERROR: There was an issue importing the application group! Exiting."
            exit 3
        }
    }
    else {
        $applications.Add($ApplicationId) | Out-Null
    }
    
    return $applications
}

# Function to replace placeholders in scripts
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
    if ($Script:installScript) {
        $Script:installScript = $Script:installScript.replace("<filename>", $FileName)
    }
    if ($Script:uninstallScript) {
        $Script:uninstallScript = $Script:uninstallScript.replace("<filename>", $FileName)
    }
    if ($Script:detectionScript) {
        $Script:detectionScript = $Script:detectionScript.replace("<filename>", $FileName)
    }
    if ($Script:registryDetectionKey) {
        $Script:registryDetectionKey = $Script:registryDetectionKey.replace("<filename>", $FileName)
    }

    # Replace the <productcode> placeholder with the actual product code
    if ($ProductCode) {
        if ($Script:installScript) {
            $Script:installScript = $Script:installScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:uninstallScript) {
            $Script:uninstallScript = $Script:uninstallScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:detectionScript) {
            $Script:detectionScript = $Script:detectionScript.replace("<productcode>", $ProductCode)
        }
        if ($Script:registryDetectionKey) {
            $Script:registryDetectionKey = $Script:registryDetectionKey.replace("<productcode>", $ProductCode)
        }
    }
    
    # Replace <version> placeholder with the actual version
    if ($Script:installScript) {
        $Script:installScript = $Script:installScript.replace("<version>", $Version)
    }
    if ($Script:uninstallScript) {
        $Script:uninstallScript = $Script:uninstallScript.replace("<version>", $Version) 
    }
    if ($Script:detectionScript) {
        $Script:detectionScript = $Script:detectionScript.replace("<version>", $Version) 
    } 
    if ($Script:registryDetectionKey) {
        $Script:registryDetectionKey = $Script:registryDetectionKey.replace("<version>", $Version)
    }
}

# Function to create detection rules
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
            switch ($Script:fileDetectionMethod) {
                "exists" {
                    return New-IntuneWin32AppDetectionRuleFile -Existence -DetectionType "exists" -Path $Script:fileDetectionPath -FileOrFolder $Script:fileDetectionName
                }
                "modified" {
                    return New-IntuneWin32AppDetectionRuleFile -DateModified -Path $Script:fileDetectionPath -FileOrFolder $Script:fileDetectionName -Operator $Script:fileDetectionOperator -DateTimeValue $Script:fileDetectionDateTime
                }
                "created" {
                    return New-IntuneWin32AppDetectionRuleFile -DateCreated -Path $Script:fileDetectionPath -FileOrFolder $Script:fileDetectionName -Operator $Script:fileDetectionOperator -DateTimeValue (Get-Date $Script:fileDetectionDateTime)
                }
                "version" {
                    return New-IntuneWin32AppDetectionRuleFile -Version -Path $Script:fileDetectionPath -FileOrFolder $Script:fileDetectionName -Operator $Script:fileDetectionOperator -VersionValue $Script:fileDetectionVersion
                }
                "size" {
                    return New-IntuneWin32AppDetectionRuleFile -Size -Path $Script:fileDetectionPath -FileOrFolder $Script:fileDetectionName -Operator $Script:fileDetectionOperator -SizeinMBValue $Script:fileDetectionValue
                }
            }
        }
        "msi" {
            return New-IntuneWin32AppDetectionRuleMsi -ProductCode $ProductCode -ProductVersion $Script:fileDetectionVersion
        }
        "registry" {
            switch ($Script:registryDetectionMethod) {
                "exists" {
                    if ($Script:registryDetectionValue) {
                        return New-IntuneWin32AppDetectionRuleRegistry -Existence -KeyPath $Script:registryDetectionKey -ValueName $Script:registryDetectionValueName -DetectionType "exists"
                    }
                    else {
                        return New-IntuneWin32AppDetectionRuleRegistry -Existence -KeyPath $Script:registryDetectionKey -DetectionType "exists"
                    }
                }
                "version" {
                    return New-IntuneWin32AppDetectionRuleRegistry -VersionComparison -KeyPath $Script:registryDetectionKey -ValueName $Script:registryDetectionValueName -Check32BitOn64System $Script:is32BitApp -VersionComparisonOperator $Script:registryDetectionOperator -VersionComparisonValue $Script:registryDetectionValue
                }
                "integer" {
                    return New-IntuneWin32AppDetectionRuleRegistry -IntegerComparison -KeyPath $Script:registryDetectionKey -ValueName $Script:registryDetectionValueName -Check32BitOn64System $Script:is32BitApp -IntegerComparisonOperator $Script:registryDetectionOperator -IntegerComparisonValue $Script:registryDetectionValue
                }
                "string" {
                    return New-IntuneWin32AppDetectionRuleRegistry -StringComparison -KeyPath $Script:registryDetectionKey -ValueName $Script:registryDetectionValueName -Check32BitOn64System $Script:is32BitApp -StringComparisonOperator $Script:registryDetectionOperator -StringComparisonValue $Script:registryDetectionValue
                }
            }
        }
        "script" {
            if (!(Test-Path $Script:SCRIPTS\$Script:id)) {
                New-Item -Name $Script:id -ItemType Directory -Path $Script:SCRIPTS
            }
            $ScriptLocation = "$($Script:SCRIPTS)\$($Script:id)\$($Script:version).$($Script:detectionScriptFileExtension)"
            Write-Output $Script:detectionScript | Out-File $ScriptLocation
            return New-IntuneWin32AppDetectionRuleScript -ScriptFile $ScriptLocation -EnforceSignatureCheck $Script:detectionScriptEnforceSignatureCheck -RunAs32Bit $Script:detectionScriptRunAs32Bit
        }
    }
}

# Function to add deployment with architecture filtering
function Add-DeploymentWithArchitecture {
    <#
    .SYNOPSIS
    Adds a deployment assignment with optional architecture filtering.
    
    .DESCRIPTION
    This function handles the assignment of applications to groups with
    optional architecture-specific filtering for ARM64 or AMD64.
    
    .PARAMETER AppId
    The ID of the application to deploy.
    
    .PARAMETER GroupId
    The ID of the group to deploy to.
    
    .PARAMETER Architecture
    The target architecture (arm64, amd64, or default for no filtering).
    #>
    param(
        [string]$AppId,
        [string]$GroupId,
        [string]$Architecture
    )
    
    switch ($Architecture) {
        "arm64" {
            Write-Log "Architecture filter: ARM64"
            Add-IntuneWin32AppAssignmentGroup -Include -ID $AppId -GroupID $GroupId -Intent "available" -Notification "hideAll" -FilterMode Include -FilterName "$($Script:arm64FilterName)" | Out-Null
        }
        "amd64" {
            Write-Log "Architecture filter: AMD64"
            Add-IntuneWin32AppAssignmentGroup -Include -ID $AppId -GroupID $GroupId -Intent "available" -Notification "hideAll" -FilterMode Include -FilterName "$($Script:amd64FilterName)" | Out-Null
        }
        default {
            Write-Log "No architecture filter selected"
            Add-IntuneWin32AppAssignmentGroup -Include -ID $AppId -GroupID $GroupId -Intent "available" -Notification "hideAll" | Out-Null
        }
    }
}

# Function to perform cleanup operations
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
        Get-ChildItem $BUILDSPACE -Exclude ".gitkeep" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Log "Warning: Could not clean buildspace completely: $_"
    }
    
    Write-Log "Removing .intunewin files..."
    try {
        Get-ChildItem $PUBLISHED -Exclude ".gitkeep" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Log "Warning: Could not clean published files completely: $_"
    }
}

# Get the list of applications to process
$Applications = Get-ApplicationsToProcess -ApplicationId $ApplicationId -Group $Group -All:$All -NoInteractive:$NoInteractive

# Main processing loop
foreach ($ApplicationId in $Applications) {
    Write-Log "Starting update for $ApplicationId..."
    # Refresh token if necessary
    Connect-AutoMSIntuneGraph
    
    # Clear the temp file
    Write-Log "Clearing the temp directory..."
    Get-ChildItem $TEMP -Exclude ".gitkeep" -Recurse | Remove-Item -Recurse -Force

    # Open the YAML file and collect all necessary attributes
    try {
        $appName = (Get-ChildItem $RECIPES -Force -Recurse | Where-Object Name -ne 'Disabled' | Get-ChildItem -File -Recurse | Where-Object Name -match "^$ApplicationId\.ya{0,1}ml")[0].FullName
        $parameters = Get-Content "$appName" | ConvertFrom-Yaml
    }
    catch {
        Write-Error "Unable to open parameters file for $ApplicationId"
        continue
    }

    # Set all script variables from parameters and preferences
    Set-ScriptVariables -Parameters $parameters -Preferences $prefs


    if ($Repair) {
        # Correct any naming discrepancies before we continue
        # Rename any apps if they are named incorrectly
        $CurrentApps = Get-SameAppAllVersions $displayName
        for ($i = 1; $i -lt $CurrentApps.Count; $i++) {
            if ($CurrentApps[$i].DisplayName -ne "$displayName (N-$i)") {
                Write-Log "Setting name for $displayName (N-$i)"
                Set-IntuneWin32App -Id $CurrentApps[$i].Id -DisplayName "$displayName (N-$i)"
            }
        }
    }


    # Run the pre-download script
    if ($preDownloadScript) {
        Write-Log "Running pre-download script..."
        try {
            Invoke-Command -ScriptBlock $preDownloadScript -NoNewScope
        }
        catch {
            Write-Error "Error while running pre-download PowerShell script"
            continue
        }
        Write-Log "Pre-download script ran successfully."
    }
    else {
        Write-Log "Skipping Pre-download script"
    }


    # Check if there is an up-to-date version in the repo already
    Write-Log "Checking if $displayName $version is a new version..."
    $ExistingVersions = Get-SameAppAllVersions $displayName
    
    if (-not $ExistingVersions) {
        Write-Log "No existing versions found for $displayName. Continuing with update."
        $VersionCompareResult = 0
    }
    else {
        $VersionCompareResult = Compare-AppVersions $version $($ExistingVersions.displayVersion[0])
    }
    
    # Check various conditions to determine if we should proceed
    if ($Force) {
        Write-Log "Force flag is set. Forcing update of $displayName $version"
    }
    elseif (Get-VersionLocked -Version $version -VersionLock $versionLock) {
        Write-Log "Version is locked to $versionLock. Skipping update."
        continue
    }
    elseif ($ExistingVersions.displayVersion -contains $version) {
        Write-Log "$id $displayName $version is already in the repo. Skipping update."
        continue
    }
    elseif ($VersionCompareResult -eq 1) {
        Write-Log "$displayName $version is a newer version. Continuing with update."
    }
    elseif ($VersionCompareResult -eq -1) {
        Write-Log "$displayName $version is older than the currently newest available version $($ExistingVersions.displayVersion[0]). Skipping update."
        continue
    }


    # See if this has been run before. If there are previous files, move them to a folder called "Old"
    if (Test-Path $BUILDSPACE\$id) {
        if (-not (Test-Path $BUILDSPACE\Old)) {
            New-Item -Path $BUILDSPACE -ItemType Directory -Name "Old"
        }
        Write-Log "Removing old Buildspace..."
        Move-Item -Path $BUILDSPACE\$id $BUILDSPACE\Old\$id-$(Get-Date -Format "MMddyyhhmmss")
    }
    if (Test-Path $SCRIPTS\$id) {
        if (-not (Test-Path $SCRIPTS\Old)) {
            New-Item -Path $SCRIPTS -ItemType Directory -Name "Old"
        }
        Write-Log "Removing old script space..."
        Move-Item -Path $SCRIPTS\$id $SCRIPTS\Old\$id-$(Get-Date -Format "MMddyyhhmmss")
    }

    # Make the new BUILDSPACE directory
    New-Item -Path $BUILDSPACE\$id -ItemType Directory -Name $version
    Push-Location $BUILDSPACE\$id\$version

    # Download the new installer
    Write-Log "Starting download..."
    Write-Log "URL: $url"
    if (!$url) {
        Write-Error "URL is empty - cannot continue."
        continue
    }
    
    if ($downloadScript) {
        try {
            Invoke-Command -ScriptBlock $downloadScript -NoNewScope
            Write-Log "Download script ran successfully."
        }
        catch {
            Write-Error "Error while running download PowerShell script: $_"
            continue
        }
    }
    else {
        try {
            Start-BitsTransfer -Source $url -Destination $BUILDSPACE\$id\$version\$fileName
            Write-Log "File downloaded successfully using BITS transfer."
        }
        catch {
            Write-Error "Error downloading file: $_"
            continue
        }
    }


    # Run the post-download script
    if ($postDownloadScript) {
        Write-Log "Running post download script..."
        try {
            Invoke-Command -ScriptBlock $postDownloadScript -NoNewScope
            Write-Log "Post download script ran successfully."
        }
        catch {
            Write-Error "Error while running post download PowerShell script: $_"
            continue
        }
    }
    

    # Handle script placeholder replacement
    $productCode = ""
    if ($fileName -match "\.msi$") {
        $productCode = Get-MSIProductCode $BUILDSPACE\$id\$version\$fileName
        Write-Log "Product Code: $productCode"
    }
    
    Update-ScriptPlaceholders -FileName $fileName -ProductCode $productCode -Version $version  


    # Generate the .intunewin file
    Pop-Location
    Write-Log "Generating .intunewin file..."
    $app = New-IntuneWin32AppPackage -SourceFolder $BUILDSPACE\$id\$version -SetupFile $fileName -OutputFolder $PUBLISHED -Force

    # Upload .intunewin file to Intune
    # Detection Types
    $Icon = New-IntuneWin32AppIcon -FilePath "$($ICONS)\$($iconFile)"
    if (-not $fileDetectionVersion) {
        $fileDetectionVersion = $version
    }

    $DetectionRule = New-DetectionRule -DetectionType $detectionType -ProductCode $productCode

    # Generate the min OS requirement rule
    $RequirementRule = New-IntuneWin32AppRequirementRule -Architecture "All" -MinimumSupportedWindowsRelease $minOSVersion



    # Create the Intune App
    Write-Log "Uploading $displayName to Intune..."
    Connect-AutoMSIntuneGraph
    if ($allowUserUninstall) {
        $Win32App = Add-IntuneWin32App -FilePath $app.path -DisplayName $displayName -Description $description -Publisher $publisher -InstallExperience $installExperience -RestartBehavior $restartBehavior -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $installScript -UninstallCommandLine $uninstallScript -Icon $Icon -AppVersion "$version" -ScopeTagName $scopeTags -Owner $owner -MaximumInstallationTimeInMinutes $maximumInstallationTimeInMinutes -AllowAvailableUninstall
    }
    else {
        $Win32App = Add-IntuneWin32App -FilePath $app.path -DisplayName $displayName -Description $description -Publisher $publisher -InstallExperience $installExperience -RestartBehavior $restartBehavior -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $installScript -UninstallCommandLine $uninstallScript -Icon $Icon -AppVersion "$version" -ScopeTagName $scopeTags -Owner $owner -MaximumInstallationTimeInMinutes $maximumInstallationTimeInMinutes
    }



    ###################################################
    # MIGRATE OLD DEPLOYMENTS
    ###################################################


    # Check if any existing applications have the same version so we can delete them
    $ToRemove = $ExistingVersions | where-object displayVersion -eq $version
    if ($ToRemove) {
        # Remove conflicting versions
        Write-Log "Removing conflicting versions"
        $CurrentApp = Get-IntuneWin32App -Id $Win32App.id
        foreach ($RemoveApp in $ToRemove) {
            Write-Log "Moving assignments before removal..."
            Move-AssignmentsAndDependencies -From $RemoveApp -To $CurrentApp -AvailableDateOffset $Script:availableDateOffset -DeadlineDateOffset $Script:deadlineDateOffset
            Write-Log "Removing App with ID $($RemoveApp.id)"
            Remove-IntuneWin32App -Id $RemoveApp.id
        }
    }

    # Define the current version, the version that is one older but shares the same name, and all the ones older than that
    Write-Log "Updating local application manifest..."
    Start-Sleep -Seconds 4
    $AllMatchingApps = Get-SameAppAllVersions $displayName 
    $CurrentApp = Get-IntuneWin32App -ID $Win32App.id
    $AllOldApps = $AllMatchingApps | Where-Object id -ne $Win32App.Id | Sort-Object displayName
    $NMinusOneApps = $AllOldApps | Where-Object displayName -eq $displayName
    $NMinusTwoAndOlderApps = $AllOldApps | Where-Object displayName -ne $displayName

    

    # Start with the N-1 app first and move all its deployments to the newest one
    if ($NMinusOneApps) {
        foreach ($NMinusOneApp in $NMinusOneApps) {
            if ($CurrentApp) {
                Write-Log "Moving assignments from $($NMinusOneApp.id) to $($CurrentApp.id)"
                Move-AssignmentsAndDependencies -From $NMinusOneApp -To $CurrentApp -AvailableDateOffset $Script:availableDateOffset -DeadlineDateOffset $Script:deadlineDateOffset

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
                Move-AssignmentsAndDependencies -From $NMinusTwoAndOlderApps[$i] -To $NewestNMinusOneApp -AvailableDateOffset $Script:availableDateOffset -DeadlineDateOffset $Script:deadlineDateOffset
            }
        }
        else {
            if ($NMinusTwoAndOlderApps[$i] -and $NMinusTwoAndOlderApps[$i - 1]) {
                Move-AssignmentsAndDependencies -From $NMinusTwoAndOlderApps[$i] -To $NMinusTwoAndOlderApps[$i - 1] -AvailableDateOffset $Script:availableDateOffset -DeadlineDateOffset $Script:deadlineDateOffset
            }
        }
    }

    # Add the default deployments if they're not already there from the migration
    if ($Script:defaultDeploymentGroups) {
        $ID = $CurrentApp.Id
        $CurrentlyDeployedIDs = (Get-IntuneWin32AppAssignment -Id $ID).GroupID
        foreach ($DeploymentGroupID in $Script:defaultDeploymentGroups) {
            if (!($CurrentlyDeployedIDs -Contains $DeploymentGroupID)) {
                Write-Log "Deploying $ID to $DeploymentGroupID because it is in the default list"
                Add-DeploymentWithArchitecture -AppId $ID -GroupId $DeploymentGroupID -Architecture $Script:architecture
            }
        }
    }

    # Rename all the old applications to have the appropriate N-<versions behind> in them
    for ($i = 1; $i -lt $AllMatchingApps.count; $i++) {
        Set-IntuneWin32App -Id $AllMatchingApps[$i].Id -DisplayName "$($displayName) (N-$i)"
        if ($?) {
            Write-Log "Successfully set Application Name to $($displayName) (N-$i)"
        }
        else {
            Write-Log "ERROR: Failed to update display name to $($displayName) (N-$i)"
        }
    }
    
    # Remove all the old versions 
    if (!$NoDelete) {
        for ($i = $numVersionsToKeep - 1; $i -lt $AllOldApps.count; $i++) {
            Write-Log "Removing old app with id $($AllOldApps[$i].id)"
            Remove-IntuneWin32App -Id $AllOldApps[$i].id
        }
    }
    Write-Log "Updates complete for $displayName"
}

# Clean up
Invoke-Cleanup

# Return to the original directory
Pop-Location