using module .\Modules\Custom\AdobeDownloader.psm1

param (
    [Alias("AppId")]
    [String] $ApplicationId = "None",
    [String] $Group,
    [Switch] $Force,
	[Switch] $All,
    [Switch] $NoDelete,
    [Switch] $Repair
)



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

# Constants
$Global:LOG_LOCATION = "$PSScriptRoot"
$Global:LOG_FILE = "YLog.log"

# So we can pop at the end
Push-Location $PSScriptRoot

# Initialize the log file
Write-Log -Init

if ("None" -eq $ApplicationId -and (($All -eq $false) -and (!$Group))) {
    Write-Log "Please provide parameter -ApplicationId"
    exit 1
}
$applications = [System.Collections.ArrayList]::new()

# Import preferences file:
try {
    $prefs = Get-Content $PSScriptRoot\Preferences.yaml | ConvertFrom-Yaml
}
catch {
    Write-Error "Unable to open preferences.yaml!"
    exit 1
}

# Import Folder Locations
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

if ($ApplicationId -ne "None") {
    if (-not (Test-Path "$RECIPES\$ApplicationId.yaml")) {
        Write-Log "ERROR: Application $ApplicationId not found in $RECIPES"
        exit 1
    }
}


# Start updating each application if all are selected
if ($All) {
	$ApplicationFullNames = $(Get-ChildItem $RECIPES -File).Name
	foreach ($Application in $ApplicationFullNames) {
		$Applications.Add($($Application -split ".yaml")[0]) | Out-Null
	}
}
elseif ($Group) {
    # Open the RecipeGroups.yaml file and read in the appropriate list of application ids
    try {
        $groupFile = Get-Content $PSScriptRoot\RecipeGroups.yaml | ConvertFrom-Yaml
        $groupFile[$Group] | ForEach-Object {
            $Applications.Add($_) | Out-Null
        }
    
    }
    catch {
        Write-Log "ERROR: There was an issue importing the application group! Exiting."
        exit 3
    }
}
else {
    # Just do the one specified if -All is not specified
    $Applications.Add($ApplicationId) | Out-Null
}


foreach ($ApplicationId in $Applications) {
    Write-Log "Starting update for $ApplicationId..."
    # Refresh token if necessary
    Connect-AutoMSIntuneGraph
    
    # Clear the temp file
    Write-Log "Clearing the temp directory..."
    Get-ChildItem $TEMP -Exclude ".gitkeep" -Recurse | Remove-Item -Recurse -Force

    # Open the YAML file and collect all necessary attributes
    try {
        $parameters = Get-Content "$RECIPES\$ApplicationId.yaml" | ConvertFrom-Yaml
    }
    catch {
        Write-Error "Unable to open parameters file for $ApplicationId"
        continue
    }
    $Script:url = if ($parameters.urlRedirects -eq $true) {Get-RedirectedUrl $parameters.url} else {$parameters.url}
    $Script:id = $parameters.id
    $Script:version = $parameters.version
    $Script:fileDetectionVersion = $parameters.fileDetectionVersion
    $Script:displayName = $parameters.displayName
    $Script:displayVersion = $parameters.displayVersion
    $Script:fileName = $parameters.fileName
    $Script:fileDetectionPath = $parameters.fileDetectionPath
    $Script:preDownloadScript = if ($parameters.preDownloadScript) { [ScriptBlock]::Create($parameters.preDownloadScript)}
    $Script:postDownloadScript = if ($parameters.postDownloadScript) { [ScriptBlock]::Create($parameters.postDownloadScript) }
    $Script:downloadScript = if ($parameters.downloadScript) { [ScriptBlock]::Create($parameters.downloadScript) }
    $Script:installScript = $parameters.installScript
    $Script:uninstallScript = $parameters.uninstallScript
    $Script:scopeTags = if($parameters.scopeTags) {$parameters.scopeTags} else {$prefs.defaultScopeTags}
    $Script:owner = if($parameters.owner) {$parameters.owner} else {$prefs.defaultOwner}
    $Script:maximumInstallationTimeInMinutes = if($parameters.maximumInstallationTimeInMinutes) {$parameters.maximumInstallationTimeInMinutes} else {$prefs.defaultMaximumInstallationTimeInMinutes}
    $Script:minOSVersion = if($parameters.minOSVersion) {$parameters.minOSVersion} else {$prefs.defaultMinOSVersion}
    $Script:installExperience = if($parameters.installExperience) {$parameters.installExperience} else {$prefs.defaultInstallExperience}
    $Script:restartBehavior = if($parameters.restartBehavior) {$parameters.restartBehavior} else {$prefs.defaultRestartBehavior}
    $Script:availableGroups = if($parameters.availableGroups) {$parameters.availableGroups} else {$prefs.defaultAvailableGroups}
    $Script:requiredGroups = if($parameters.requiredGroups) {$parameters.requiredGroups} else {$prefs.defaultRequiredGroups}
    $Script:defaultDeploymentGroups = if($parameters.defaultDeploymentGroups) {$parameters.defaultDeploymentGroups} else {$prefs.defaultDeploymentGroups}
    $Script:detectionType = $parameters.detectionType
    $Script:fileDetectionVersion = $parameters.fileDetectionVersion
    $Script:fileDetectionMethod = $parameters.fileDetectionMethod
    $Script:fileDetectionName = $parameters.fileDetectionName
    $Script:fileDetectionOperator = $parameters.fileDetectionOperator
    $Script:fileDetectionDateTime = $parameters.fileDetectionDateTime
    $Script:fileDetectionValue = $parameters.fileDetectionValue
    $Script:registryDetectionMethod = $parameters.registryDetectionMethod
    $Script:registryDetectionKey = $parameters.registryDetectionKey
    $Script:registryDetectionValueName = $parameters.registryDetectionValueName
    $Script:registryDetectionValue = $parameters.registryDetectionValue
    $Script:registryDetectionOperator = $parameters.registryDetectionOperator
    $Script:detectionScript = $parameters.detectionScript
    $Script:detectionScriptFileExtension = if($parameters.detectionScriptFileExtension) {$parameters.detectionScriptFileExtension} else {$prefs.defaultDetectionScriptFileExtension}
    $Script:detectionScriptRunAs32Bit = if($parameters.detectionScriptRunAs32Bit) {$parameters.detectionScriptRunAs32Bit} else {$prefs.defaultdetectionScriptRunAs32Bit}
    $Script:detectionScriptEnforceSignatureCheck = if($parameters.detectionScriptEnforceSignatureCheck) {$parameters.detectionScriptEnforceSignatureCheck} else {$prefs.defaultdetectionScriptEnforceSignatureCheck}
    $Script:allowUserUninstall = if($parameters.allowUserUninstall) {$parameters.allowUserUninstall} else {$prefs.defaultAllowUserUninstall}
    $Script:iconFile = $parameters.iconFile
    $Script:description = $parameters.description
    $Script:publisher = $parameters.publisher
    $Script:is32BitApp = if($parameters.is32BitApp) {$parameters.is32BitApp} else {$prefs.defaultIs32BitApp}
    $Script:numVersionsToKeep = if($parameters.numVersionsToKeep) {$parameters.numVersionsToKeep} else {$prefs.defaultNumVersionsToKeep}


    if ($Repair) {
        # Correct any naming discrepancies before we continue
        # Rename any apps if they are named incorrectly
        $CurrentApps = Get-SameAppAllVersions $DisplayName
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
        Invoke-Command -ScriptBlock $preDownloadScript -NoNewScope

        if (!$?) {
                Write-Error "Error while running pre-download PowerShell script"
                continue
        }
        else {
            Write-Log "Pre-download script ran successfully."
        }
    }
    else {
        Write-Log "Skipping Pre-download script"
    }


    # Check if there is an up-to-date version in the repo already
    Write-Log "Checking if $displayName $version is a new version..."
    $ExistingVersions = Get-SameAppAllVersions $DisplayName
    if ($ExistingVersions.displayVersion -contains $version) {
        if ($force) {
            Write-Log "Package up-to-date. -Force applied. Recreating package."
        }
        else {
            Write-Log "$id already up-to-date!"
            continue
        }
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
    Set-Location $BUILDSPACE\$id\$version

    # Download the new installer
    Write-Log "Starting download..."
    Write-Log "URL: $url"
    if (!$url) {
        Write-Error "URL is empty - cannot continue."
        continue
    }
    if ($downloadScript) {
        Invoke-Command -ScriptBlock $downloadScript -NoNewScope
        if (!$?) {
            Write-Error "Error while running download PowerShell script"
            continue
        }
        else {
            Write-Log "Download script ran successfully."
        }
    }
    else {
        Start-BitsTransfer -Source $url -Destination $BUILDSPACE\$id\$version\$fileName
    }


    # Run the post-download script
    if ($postDownloadScript) {
        Write-Log "Running post download script..."
        $result = Invoke-Command -ScriptBlock $postDownloadScript -NoNewScope
        if (!$? -or ($result -eq $false)) {
                Write-Error "Error while running post download PowerShell script"
                continue
        }
        else {
            Write-Log "Post download script ran successfully."
        }
    }
    

    # Script Files:
    # Replace the <filename> placeholder with the actual filename
    if ($installScript) {
        $installScript = $installScript.replace("<filename>", $fileName)
    }
    if ($uninstallScript) {
        $uninstallScript = $uninstallScript.replace("<filename>", $fileName)
    }
    if ($detectionScript) {
        $detectionScript = $detectionScript.replace("<filename>", $fileName)
    }
    if ($registryDetectionKey) {
        $registryDetectionKey = $registryDetectionKey.replace("<filename>", $fileName)
    }


    # Replace the <productcode> placeholder with the actual product code
    if ($filename -match "\.msi$") {
        $ProductCode = Get-MSIProductCode $BUILDSPACE\$id\$version\$fileName
        Write-Log "Product Code: $ProductCode"
        if ($installScript) {
            $installScript = $installScript.replace("<productcode>", $productCode)
        }
        if ($uninstallScript) {
            $uninstallScript = $uninstallScript.replace("<productcode>", $productCode)
        }
        if ($detectionScript) {
            $detectionScript = $detectionScript.replace("<productcode>", $productCode)
        }
        if ($registryDetectionKey) {
            $registryDetectionKey = $registryDetectionKey.replace("<productcode>", $productCode)
        }
    }

    
    # Replace <version> placeholder with the actual version
    if ($installScript) {
        $installScript = $installScript.replace("<version>", $version)
    }
    if ($uninstallScript) {
        $uninstallScript = $uninstallScript.replace("<version>", $version) 
    }
    if ($detectionScript) {
        $detectionScript = $detectionScript.replace("<version>", $version) 
    } 
    if ($registryDetectionKey) {
        $registryDetectionKey = $registryDetectionKey.replace("<version>", $version)
    }  


    # Generate the .intunewin file
    Set-Location $PSScriptRoot
    Write-Log "Generating .intunewin file..."
    $app = New-IntuneWin32AppPackage -SourceFolder $BUILDSPACE\$id\$version -SetupFile $filename -OutputFolder $PUBLISHED -Force

    # Upload .intunewin file to Intune
    # Detection Types
    $Icon = New-IntuneWin32AppIcon -FilePath "$($ICONS)\$($iconFile)"
    if ( -not ($fileDetectionVersion)) {
        $fileDetectionVersion = $version
    }

    if ($detectionType -eq "file") {
        if ($fileDetectionMethod -eq "exists") {
            $DetectionRule = New-IntuneWin32AppDetectionRuleFile -Existence -DetectionType "exists" -Path $fileDetectionPath -FileOrFolder $fileDetectionName
        }
        elseif ($fileDetectionMethod -eq "modified") {
            $DetectionRule = New-IntuneWin32AppDetectionRuleFile -DateModified -Path $fileDetectionPath -FileOrFolder $fileDetectionName -Operator $fileDetectionOperator -DateTimeValue $fileDetectionDateTime
        }
        elseif ($fileDetectionMethod -eq "created") {
            $DetectionRule = New-IntuneWin32AppDetectionRuleFile -DateCreated -Path $fileDetectionPath -FileOrFolder $fileDetectionName -Operator $fileDetectionOperator -DateTimeValue (Get-Date $fileDetectionDateTime)
        }
        elseif ($fileDetectionMethod -eq "version") {
            $DetectionRule = New-IntuneWin32AppDetectionRuleFile -Version -Path $fileDetectionPath -FileOrFolder $fileDetectionName -Operator $fileDetectionOperator -VersionValue $fileDetectionVersion
        }
        elseif ($fileDetectionMethod -eq "size") {
            $DetectionRule = New-IntuneWin32AppDetectionRuleFile -Size -Path $fileDetectionPath -FileOrFolder $fileDetectionName -Operator $fileDetectionOperator -SizeinMBValue $fileDetectionValue
        }
    }
    elseif ($detectionType -eq "msi") {
        $DetectionRule = New-IntuneWin32AppDetectionRuleMsi -ProductCode "$ProductCode" -ProductVersion $fileDetectionVersion 
    }
    elseif ($detectionType -eq "registry") {
        if($registryDetectionMethod -eq "exists") {
            if ($registryDetectionValue) {
                $DetectionRule = New-IntuneWin32AppDetectionRuleRegistry -Existence -KeyPath $registryDetectionKey -ValueName $registryDetectionValueName -DetectionType "exists"
            }
            else {
                $DetectionRule = New-IntuneWin32AppDetectionRuleRegistry -Existence -KeyPath $registryDetectionKey -DetectionType "exists"
            }
        }
        elseif($registryDetectionMethod -eq "version") {
            $DetectionRule = New-IntuneWin32AppDetectionRuleRegistry -VersionComparison -KeyPath $registryDetectionKey -ValueName $registryDetectionValueName -Check32BitOn64System $is32BitApp -VersionComparisonOperator $registryDetectionOperator -VersionComparisonValue $registryDetectionValue
        }
        elseif($registryDetectionMethod -eq "integer") {
            $DetectionRule = New-IntuneWin32AppDetectionRuleRegistry -IntegerComparison -KeyPath $registryDetectionKey -ValueName $registryDetectionValue -Check32BitOn64System $is32BitApp -IntegerComparisonOperator $registryDetectionOperator -IntegerComparisonValue $registryDetectionValue
        }
        elseif($registryDetectionMethod -eq "string") {
            $DetectionRule = New-IntuneWin32AppDetectionRuleRegistry -StringComparison -KeyPath $registryDetectionKey -ValueName $registryDetectionValue -Check32BitOn64System $is32BitApp -StringComparisonOperator $registryDetectionOperator -StringComparisonValue $registryDetectionValue
        }
    }
    elseif ($detectionType -eq "script") {
        if (!(Test-Path $SCRIPTS\$id)) {
            New-Item -Name $id -ItemType Directory -Path $SCRIPTS
        }
        $ScriptLocation = "$SCRIPTS\$id\$version.$detectionScriptFileExtension"
        Write-Output $detectionScript | Out-File $ScriptLocation
        $DetectionRule = New-IntuneWin32AppDetectionRuleScript -ScriptFile $ScriptLocation -EnforceSignatureCheck $detectionScriptEnforceSignatureCheck -RunAs32Bit $detectionScriptRunAs32Bit
    }

    # Generate the min OS requirement rule
    $RequirementRule = New-IntuneWin32AppRequirementRule -Architecture "All" -MinimumSupportedWindowsRelease $minOSVersion



    # Create the Intune App
    Write-Log "Uploading $displayName to Intune..."
    Connect-AutoMSIntuneGraph
    if ($allowUserUninstall) {
        $Win32App = Add-IntuneWin32App -FilePath $app.path -DisplayName $DisplayName -Description $description -Publisher $publisher -InstallExperience $installExperience -RestartBehavior $restartBehavior -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $InstallScript -UninstallCommandLine $UninstallScript -Icon $Icon -AppVersion "$Version" -ScopeTagName $ScopeTags -Owner $owner -MaximumInstallationTimeInMinutes $maximumInstallationTimeInMinutes -AllowAvailableUninstall
    }
    else {
        $Win32App = Add-IntuneWin32App -FilePath $app.path -DisplayName $DisplayName -Description $description -Publisher $publisher -InstallExperience $installExperience -RestartBehavior $restartBehavior -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $InstallScript -UninstallCommandLine $UninstallScript -Icon $Icon -AppVersion "$Version" -ScopeTagName $ScopeTags -Owner $owner -MaximumInstallationTimeInMinutes $maximumInstallationTimeInMinutes
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
        foreach ($removeapp in $ToRemove) {
            Write-Log "Moving assignments before removal..."
            Move-Assignments -From $removeapp -To $CurrentApp
            Write-Log "Removing App with ID $($removeapp.id)"
            Remove-IntuneWin32App -Id $removeapp.id
        }
    }

    # Define the current version, the version that is one older but shares the same name, and all the ones older than that
    Write-Log "Updating local application manifest..."
    Start-Sleep -Seconds 4
    $AllMatchingApps = Get-SameAppAllVersions $DisplayName 
    $CurrentApp = Get-IntuneWin32App -ID $Win32App.id
    $AllOldApps = $AllMatchingApps | Where-Object id -ne $Win32App.Id | Sort-Object displayName
    $NMinusOneApps = $AllOldApps | Where-Object displayName -eq $displayName
    $NMinusTwoAndOlderApps = $AllOldApps | Where-Object displayName -ne $displayName

    

    # Start with the N-1 app first and move all its deployments to the newest one
    if ($NMinusOneApps) {
        foreach ($NMinusOneApp in $NMinusOneApps) {
            if ($CurrentApp) {
                Write-Log "Moving assignments from $($NMinusOneApp.id) to $($CurrentApp.id)"
                Move-Assignments -From $NMinusOneApp -To $CurrentApp
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
                Move-Assignments -From $NMinusTwoAndOlderApps[$i] -To $NewestNMinusOneApp
            }
        }
        else {
            if ($NMinusTwoAndOlderApps[$i] -and $NMinusTwoAndOlderApps[$i - 1]) {
                Move-Assignments -From $NMinusTwoAndOlderApps[$i] -To $NMinusTwoAndOlderApps[$i - 1]
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
                Add-IntuneWin32AppAssignmentGroup -Include -ID $id -GroupID $DeploymentGroupID -Intent "available" -Notification "hideAll" | Out-Null
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
Write-Log "Cleaning up the Buildspace..."
Get-ChildItem $BUILDSPACE -Exclude ".gitkeep" -Recurse | Remove-Item -Recurse -Force
Write-Log "Removing .intunewin files..."
Get-ChildItem $PUBLISHED -Exclude ".gitkeep" -Recurse | Remove-Item -Recurse -Force

# Return to the original directory
Pop-Location
