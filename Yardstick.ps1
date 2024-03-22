param (
    [Alias("AppId")]
    [String] $ApplicationId = "None",
    [Switch] $Force,
	[Switch] $All,
    [Switch] $NoDelete,
    [Switch] $Repair
)

# Modules required:
# powershell-yaml
# IntuneWin32App
# Selenium
Import-Module powershell-yaml
Import-Module IntuneWin32App
Import-Module Selenium

# Constants
$LOG_LOCATION = "$PSScriptRoot"
$LOG_FILE = "YLog.log"

# Useful Functions

# Write-Log
# Prints both to the console and also to the log
# Param: [String] $ContentToWriteToLog, [Switch] $Init (Erase log and print big timestamp for new script run)
function Write-Log {
    param(
        [String]$Content,
        [Switch]$Init
    )
    Push-Location $LOG_LOCATION
    if($Init) {
        if (Test-Path $LOG_LOCATION\$LOG_FILE) {
            Remove-Item $LOG_LOCATION\$LOG_FILE -Force
        }
        Write-Output "#######################################################" | Out-File $LOG_FILE
        Write-Output "LOGGING STARTED AT $(Get-Date -Format "MM/dd/yyyy HH:mm:ss")" | Out-File $LOG_FILE -Append
        Write-Output "#######################################################" | Out-File $LOG_FILE -Append
    }
    if ($Content) {
        $Content = "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss") - $Content"
        Write-Output $Content | Out-File $LOG_FILE -Append
        Write-Output $Content
    }
    Pop-Location
}


# ArrayToString
# Converts an array to a string
function ArrayToString() {
    param (
        [Array] $array
    )
    $arrayString = "@("
    foreach ($value in $array) {
        $arrayString = "$($arrayString)$([String]$value),"
    }
    $arrayString = $arrayString.TrimEnd(",")
    $arrayString = "$arrayString)"
    return [String]$arrayString
}

# Get-RedirectedUrl
# Follows redirects and returns the actual URL of a resource
# Param: [String] URL
# Return: [String] The last URL that is linked to that does not redirect
function Get-RedirectedUrl() {
    param (
        [Parameter(Mandatory=$true)]
        [String]$URL
    )

    $request = [System.Net.WebRequest]::Create($url)
    $request.AllowAutoRedirect=$false
    $response=$request.GetResponse()

    If ($response.StatusCode -eq "Found")
    {
        $response.GetResponseHeader("Location")
    }
}


# Get-MsiProductCode
# Returns the product code of an MSI file in the current directory only
# Param: [String] $filePath
# Return: [String] The MSI Product Code
function Get-MsiProductCode() {
    param (
        [Parameter(Mandatory=$true)]
        [String]$filePath
    )
    # Read property from MSI database
    $windowsInstallerObject = New-Object -ComObject WindowsInstaller.Installer
    $msiDatabase = $windowsInstallerObject.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstallerObject, @($filePath, 0))
    $query = "SELECT Value FROM Property WHERE Property = 'ProductCode'"
    $view = $msiDatabase.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $msiDatabase, ($query))
    $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
    $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
    $value = $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1)
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstallerObject) 
    return [String]$value
}

# Connect-AutoMSIntuneGraph
# Automatically connects to the Microsoft Graph API using the Intune module
# Param: None
# Returns: None
function Connect-AutoMSIntuneGraph() {
    # Check if the current token is invalid
    if (-not $Global:Token) {
        Write-Log "Getting an Intune Graph Client APItoken..."
        $Global:Token = Connect-MSIntuneGraph -TenantID $TenantId -ClientID $ClientId -ClientSecret $clientSecret
    }
    elseif ($Global:Token.ExpiresOn.ToLocalTime() -lt (Get-Date)) {
        # If not, get a new token
        Write-Log "Token is expired. Refreshing token..."
        $Global:Token = Connect-MSIntuneGraph -TenantID $TenantId -ClientID $ClientId -ClientSecret $clientSecret
        Write-Log "Token refreshed. New Token Expires at: $($Global:Token.ExpiresOn.ToLocalTime())"
    }
    elseif ($Global:Token.ExpiresOn.AddMinutes(-20).ToLocalTime() -lt (Get-Date)) {
        # For whatever reason, this API stops working 10 minutes before a token refresh
        Write-Log "Token expires soon - Refreshing token..."
        # Required to force a refresh
        Clear-MsalTokenCache
        $Global:Token = Connect-MSIntuneGraph -TenantID $TenantId -ClientID $ClientId -ClientSecret $clientSecret
        Write-Log "Token refreshed. New Token Expires at: $($Global:Token.ExpiresOn.ToLocalTime())"
    }
    else {
        Write-Log "Token is still valid. Skipping token refresh."
    } 
}

# MOVE-ASSIGNMENTS
# Gets all assignments for an application and adds them to a target (To)
# Removes successfully moved assignments from the (From) application
# Param: [Application]$From, [Application]$To
# Return: None
function Move-Assignments {
    param(
        [Parameter(Mandatory, Position=0)]
        [System.Object] $From,
        [Parameter(Mandatory, Position=1)]
        [System.Object] $To
    )
    $FromAssignments = Get-IntuneWin32AppAssignment -Id $From.id
    $FromAvailable = ($FromAssignments | Where-Object intent -eq "available").groupId
    $FromRequired = ($FromAssignments | Where-Object intent -eq "required").groupId
    if ($FromAvailable) {
        foreach ($groupId in $FromAvailable) {
            $maxRetries = 3
            $try = 0
            $successfullyAdded = $false
            while (!$successfullyAdded -and ($try -lt $maxRetries)) {
                Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $groupId -Intent "available" -Notification "hideAll" | Out-Null
                # Check that it worked
                $AssignedGroups = Get-IntuneWin32AppAssignment -ID $To.id
                if($AssignedGroups | Where-Object GroupID -eq $groupID) {
                    $successfullyAdded = $true
                    Write-Log "$GroupID (Available) added successfully to $($To.id)"
                    Write-Log "Removing $groupID available assignment from $($From.id)"
                    Remove-IntuneWin32AppAssignmentGroup -ID $From.id -GroupID $groupId | Out-Null
                }
                else {
                    # Wait 1 second and retry
                    Start-Sleep -Seconds 1
                }
            }
            
            if (!$successfullyAdded) {
                Write-Log "ERROR: Cannot add available app assignment to $($To.id) for group $groupID. Will not remove original assignment."
            }
        }
        
    }
    if ($FromRequired) {
        foreach ($groupId in $FromRequired) {
            $maxRetries = 3
            $try = 0
            $successfullyAdded = $false
            while (!$successfullyAdded -and ($try -lt $maxRetries)) {
                Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $groupId -Intent "available" -Notification "hideAll" | Out-Null
                # Check that it worked
                $AssignedGroups = Get-IntuneWin32AppAssignment -ID $To.id
                if($AssignedGroups | Where-Object GroupID -eq $groupID) {
                    $successfullyAdded = $true
                    Write-Log "$GroupID (Required) added successfully to $($To.id)"
                    Write-Log "Removing $groupID required assignment from $($From.id)"
                    Remove-IntuneWin32AppAssignmentGroup -ID $From.id -GroupID $groupId | Out-Null
                }
                else {
                    # Wait 1 second and retry
                    Start-Sleep -Seconds 1
                }
            }
            
            if (!$successfullyAdded) {
                Write-Log "ERROR: Cannot add required app assignment to $($To.id) for group $groupID. Will not remove original assignment."
            }
        }
    }
}


# Get-SameAppAllVersions
# Returns all versions of an app
# Accounts for edge cases where an application name might be similar to others, i.e. Mozilla Firefox vs. Mozilla Firefox ESR
# Param: [String] DisplayName
# Return: @(PSCustomObject) 
function Get-SameAppAllVersions($DisplayName) {
    $AllSimilarApps = Get-IntuneWin32App -DisplayName "$DisplayName*"
    return ($AllSimilarApps | Where-Object {($_.DisplayName -eq $DisplayName) -or ($_.DisplayName -like "$DisplayName (*")})
}


# So we can pop at the end
Push-Location

# Initialize the log file
Write-Log -Init

if ("None" -eq $ApplicationId -and $All -eq $false) {
    Write-Log "Please provide parameter -ApplicationId"
    exit 1
}
$applications = [System.Collections.ArrayList]::new()
$testApps = "$PSScriptRoot\TestApps"
$buildSpace = "$PSScriptRoot\BuildSpace"
$scriptSpace = "$PSScriptRoot\Scripts"
$publishedApps = "$PSScriptRoot\Apps"
$autoPackagerRecipes = "$PSScriptRoot\Recipes"
$iconPath = "$PSScriptRoot\Icons"
$toolsDir = "$PSScriptRoot\Tools"
$tempDir = "$PSScriptRoot\Temp"

# Import preferences file:
$prefs = Get-Content $PSScriptRoot\Preferences.yaml | ConvertFrom-Yaml
$tenantId = $prefs.tenantId
$clientId = $prefs.clientId
$clientSecret = $prefs.clientSecret

# $scopeTags = $prefs.scopeTags

if ($ApplicationId -ne "None") {
    if (-not (Test-Path "$autoPackagerRecipes\$ApplicationId.yaml")) {
        Write-Log "Application $ApplicationId not found in $autoPackagerRecipes"
        exit 1
    }
}


# Start updating each application if all are selected
if ($All) {
	$ApplicationFullNames = $(Get-ChildItem $autoPackagerRecipes -File).Name
	foreach ($Application in $ApplicationFullNames) {
		$Applications.Add($($Application -split ".yaml")[0]) | Out-Null
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
    Get-ChildItem $TempDir -Exclude ".gitkeep" -Recurse | Remove-Item -Recurse -Force

    # Open the YAML file and collect all necessary attributes
    $parameters = Get-Content "$autoPackagerRecipes\$ApplicationId.yaml" | ConvertFrom-Yaml
    $url = if ($parameters.urlRedirects) {Get-RedirectedUrl $parameters.url} else {$parameters.url}
    $id = $parameters.id
    $version = $parameters.version
    $fileDetectionVersion = $parameters.fileDetectionVersion
    $displayName = $parameters.displayName
    $displayVersion = $parameters.displayVersion
    $fileName = $parameters.fileName
    $fileDetectionPath = $parameters.fileDetectionPath
    $preDownloadScript = $parameters.preDownloadScript
    $postDownloadScript = $parameters.postDownloadScript
    $downloadScript = $parameters.downloadScript
    $installScript = $parameters.installScript
    $uninstallScript = $parameters.uninstallScript
    $scopeTags = if($parameters.scopeTags) {$parameters.scopeTags} else {$prefs.defaultScopeTags}
    $owner = if($parameters.owner) {$parameters.owner} else {$prefs.defaultOwner}
    $maximumInstallationTimeInMinutes = if($parameters.maximumInstallationTimeInMinutes) {$parameters.maximumInstallationTimeInMinutes} else {$prefs.defaultMaximumInstallationTimeInMinutes}
    $minOSVersion = if($parameters.minOSVersion) {$parameters.minOSVersion} else {$prefs.defaultMinOSVersion}
    $installExperience = if($parameters.installExperience) {$parameters.installExperience} else {$prefs.defaultInstallExperience}
    $restartBehavior = if($parameters.restartBehavior) {$parameters.restartBehavior} else {$prefs.defaultRestartBehavior}
    $availableGroups = if($parameters.availableGroups) {$parameters.availableGroups} else {$prefs.defaultAvailableGroups}
    $requiredGroups = if($parameters.requiredGroups) {$parameters.requiredGroups} else {$prefs.defaultRequiredGroups}
    $detectionType = $parameters.detectionType
    $fileDetectionVersion = $parameters.fileDetectionVersion
    $fileDetectionMethod = $parameters.fileDetectionMethod
    $fileDetectionName = $parameters.fileDetectionName
    $fileDetectionOperator = $parameters.fileDetectionOperator
    $fileDetectionDateTime = $parameters.fileDetectionDateTime
    $fileDetectionValue = $parameters.fileDetectionValue
    $registryDetectionMethod = $parameters.registryDetectionMethod
    $registryDetectionKey = $parameters.registryDetectionKey
    $registryDetectionValueName = $parameters.registryDetectionValueName
    $registryDetectionValue = $parameters.registryDetectionValue
    $registryDetectionOperator = $parameters.registryDetectionOperator
    $detectionScript = $parameters.detectionScript
    $DetectionScriptFileExtension = if($parameters.detectionScriptFileExtension) {$parameters.detectionScriptFileExtension} else {$prefs.defaultDetectionScriptFileExtension}
    $detectionScriptRunAs32Bit = if($parameters.detectionScriptRunAs32Bit) {$parameters.detectionScriptRunAs32Bit} else {$prefs.defaultdetectionScriptRunAs32Bit}
    $detectionScriptEnforceSignatureCheck = if($parameters.detectionScriptEnforceSignatureCheck) {$parameters.detectionScriptEnforceSignatureCheck} else {$prefs.defaultdetectionScriptEnforceSignatureCheck}
    $allowUserUninstall = if($parameters.allowUserUninstall) {$parameters.allowUserUninstall} else {$prefs.defaultAllowUserUninstall}
    $iconFile = $parameters.iconFile
    $description = $parameters.description
    $publisher = $parameters.publisher
    $is32BitApp = if($parameters.is32BitApp) {$parameters.is32BitApp} else {$prefs.defaultIs32BitApp}
    $numVersionsToKeep = if($parameters.numVersionsToKeep) {$parameters.numVersionsToKeep} else {$prefs.defaultNumVersionsToKeep}


    if ($Repair) {
        # Correct any naming discrepancies before we continue
        # Rename any apps if they are named incorrectly
        $CurrentApps = Get-SameAppAllVersions $DisplayName | Sort-Object displayVersion -descending
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
        Invoke-Expression $preDownloadScript | Out-Null

        if (!$?) {
                Write-Error "Error while running pre-download PowerShell script"
                continue
        }
        else {
            Write-Log "Pre-download script ran successfully."
        }
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
    if (Test-Path $buildSpace\$id) {
        if (-not (Test-Path $buildSpace\Old)) {
            New-Item -Path $buildSpace -ItemType Directory -Name "Old"
        }
        Write-Log "Removing old buildspace..."
        Move-Item -Path $buildSpace\$id $buildSpace\Old\$id-$(Get-Date -Format "MMddyyhhmmss")
    }
    if (Test-Path $scriptSpace\$id) {
        if (-not (Test-Path $scriptSpace\Old)) {
            New-Item -Path $scriptSpace -ItemType Directory -Name "Old"
        }
        Write-Log "Removing old script space..."
        Move-Item -Path $scriptSpace\$id $scriptSpace\Old\$id-$(Get-Date -Format "MMddyyhhmmss")
    }

    # Make the new buildspace directory
    New-Item -Path $buildSpace\$id -ItemType Directory -Name $version
    Set-Location $buildSpace\$id\$version

    # Download the new installer
    Write-Log "Starting download..."
    Write-Log "URL: $url"
    if (!$url) {
        Write-Error "URL is empty - cannot continue."
        break
    }
    if ($downloadScript) {
        Invoke-Expression $downloadScript | Out-Null
        if (!$?) {
            Write-Error "Error while running download PowerShell script"
            break
        }
        else {
            Write-Log "Download script ran successfully."
        }
    }
    else {
        Start-BitsTransfer -Source $url -Destination $buildSpace\$id\$version\$fileName
    }


    # Run the post-download script
    if ($postDownloadScript) {
        Write-Log "Running post download script..."
        Invoke-Expression $postDownloadScript | Out-Null
        if (!$?) {
                Write-Error "Error while running post download PowerShell script"
                break
        }
        else {
            Write-Log "Post download script ran successfully."
        }
    }
    
    # Script Files:
    # Replace the <filename> placeholder with the actual filename
    $installScript = $installScript.replace("<filename>", $fileName)
    $uninstallScript = $uninstallScript.replace("<filename>", $fileName)

    # Replace the <productcode> placeholder with the actual product code
    if ($filename -match "\.msi$") {
        $ProductCode = Get-MSIProductCode $buildSpace\$id\$version\$fileName
        Write-Log "Product Code: $ProductCode"
        $installScript = $installScript.replace("<productcode>", $ProductCode)
        $uninstallScript = $uninstallScript.replace("<productcode>", $ProductCode)
    }

    # Replace <version> placeholder with the actual version
    $installScript = $installScript.replace("<version>", $version)
    $uninstallScript = $uninstallScript.replace("<version>", $version)  
    
    if ($registryDetectionKey) {
        $registryDetectionKey = $registryDetectionKey.replace("<filename>", $fileName)
        $registryDetectionKey = $registryDetectionKey.replace("<version>", $version)
        if ($ProductCode) {
            $registryDetectionKey = $registryDetectionKey.replace("<productcode>", $productCode)
        }
    }

    # Generate the .intunewin file
    Set-Location $PSScriptRoot
    Write-Log "Generating .intunewin file..."
    $app = New-IntuneWin32AppPackage -SourceFolder $buildSpace\$id\$version -SetupFile $filename -OutputFolder $publishedApps -Force

    # Upload .intunewin file to Intune
    # Detection Types
    $Icon = New-IntuneWin32AppIcon -FilePath "$($iconPath)\$($iconFile)"
    if ($fileDetectionVersion) {
        $fileDetectionVersion = $fileDetectionVersion
    }
    else {
        if ( -not ($fileDetectionVersion)) {
            $fileDetectionVersion = $version
        }
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
        if (!(Test-Path $scriptSpace\$id)) {
            New-Item -Name $id -ItemType Directory -Path $scriptSpace
        }
        $ScriptLocation = "$scriptSpace\$id\$version.$detectionScriptFileExtension"
        Write-Output $detectionScript | Out-File $ScriptLocation
        $DetectionRule = New-IntuneWin32AppDetectionRuleScript -ScriptFile $ScriptLocation -EnforceSignatureCheck $detectionScriptEnforceSignatureCheck -RunAs32Bit $detectionScriptRunAs32Bit
    }

    # Generate the min OS requirement rule
    $RequirementRule = New-IntuneWin32AppRequirementRule -Architecture "All" -MinimumSupportedWindowsRelease $minOSVersion

    

    # Create the Intune App
    Write-Log "Uploading $displayName to Intune..."
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
    $AllMatchingApps = Get-SameAppAllVersions $DisplayName | Sort-Object DisplayVersion -Descending
    $CurrentApp = $AllMatchingApps | Where-Object id -eq $Win32App.Id
    $AllOldApps = $AllMatchingApps | Where-Object id -ne $CurrentApp.Id
    $NMinusOneApp = $AllOldApps | Where-Object displayName -eq $displayName
    $NMinusTwoAndOlderApps = $AllOldApps | Where-Object displayName -ne $displayName

    

    # Start with the N-1 app first and move all its deployments to the newest one
    if ($NMinusOneApp) {
        Write-Log "Moving assignments from $($NMinusOneApp.id) to $($CurrentApp.id)"
        Move-Assignments -From $NMinusOneApp -To $CurrentApp
        for ($i = 0; $i -lt $NMinusTwoAndOlderApps.count; $i++) {
            # Move all the deployments up one number
            if ($i -eq 0) {
                # Move the app assignments in the 0 position to the nminusoneapp
                Move-Assignments -From $NMinusTwoAndOlderApps[$i] -To $NMinusOneApp
            }
            else {
                Move-Assignments -From $NMinusTwoAndOlderApps[$i] -To $NMinusTwoAndOlderApps[$i - 1]
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
        for ($i = $numVersionsToKeep; $i -lt $AllMatchingApps.count; $i++) {
            Write-Log "Removing old app with id $($AllMatchingApps[$i].id)"
            Remove-IntuneWin32App -Id $AllMatchingApps[$i].id
        }
    }
    Write-Log "Updates complete for $displayName"
}   

# Clean up
Write-Log "Cleaning up the buildspace..."
Get-ChildItem $buildSpace -Exclude ".gitkeep" -Recurse | Remove-Item -Recurse -Force
Write-Log "Removing .intunewin files..."
Get-ChildItem $publishedApps -Exclude ".gitkeep" -Recurse | Remove-Item -Recurse -Force
Write-Log "Removing temporary script files..."
Get-ChildItem $scriptSpace -Exclude ".gitkeep" -Recurse | Remove-Item -Recurse -Force

# Return to the original directory
Pop-Location
