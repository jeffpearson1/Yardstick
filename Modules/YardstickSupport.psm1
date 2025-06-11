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
        Write-Output "#######################################################" | Out-File $LOG_FILE -Append
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
    $userAgent = [Microsoft.PowerShell.Commands.PSUserAgent]::Chrome
    $httpClient = [System.Net.Http.HttpClient]::new()
    $httpClient.DefaultRequestHeaders.UserAgent.ParseAdd($userAgent)

    # Get the redirected url object
    $Response = $httpClient.GetAsync($url).GetAwaiter().GetResult()
    if ($Response.StatusCode -eq "OK") {
        $RedirectedURL = $response.RequestMessage.RequestUri.AbsoluteUri
    }
    else {
        throw "There was an issue getting the redirected URL."
    }
    return $RedirectedURL
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
        $Global:Token = Connect-MSIntuneGraph -TenantID $TENANT_ID -ClientID $CLIENT_ID -ClientSecret $CLIENT_SECRET
    }
    elseif ($Global:Token.ExpiresOn.ToLocalTime() -lt (Get-Date)) {
        # If not, get a new token
        Write-Log "Token is expired. Refreshing token..."
        $Global:Token = Connect-MSIntuneGraph -TenantID $TENANT_ID -ClientID $CLIENT_ID -ClientSecret $CLIENT_SECRET
        Write-Log "Token refreshed. New Token Expires at: $($Global:Token.ExpiresOn.ToLocalTime())"
    }
    elseif ($Global:Token.ExpiresOn.AddMinutes(-30).ToLocalTime() -lt (Get-Date)) {
        # For whatever reason, this API stops working 10 minutes before a token refresh
        # Set at 30 minutes in case we are uploading large files. 
        Write-Log "Token expires soon - Refreshing token..."
        # Required to force a refresh
        Clear-MsalTokenCache
        $Global:Token = Connect-MSIntuneGraph -TenantID $TENANT_ID -ClientID $CLIENT_ID -ClientSecret $CLIENT_SECRET
        Write-Log "Token refreshed. New Token Expires at: $($Global:Token.ExpiresOn.ToLocalTime())"
    }
    else {
        Write-Log "Token is still valid. Skipping token refresh."
    } 
}

# MOVE-ASSIGNMENTS
# Gets all assignments for an application and adds them to a target (To)
# Removes successfully moved assignments from the (From) application
# Also reassigns any app dependencies that are assigned to the From application
# Param: [Application]$From, [Application]$To
# Return: None
function Move-Assignments {
    param(
        [Parameter(Mandatory, Position=0)]
        [System.Object] $From,
        [Parameter(Mandatory, Position=1)]
        [System.Object] $To,
        [Parameter(Position=2)]
        [String] $ArchitectureFilterName
    )
    $FromAssignments = Get-IntuneWin32AppAssignment -Id $From.id
    $FromDependencies = Get-IntuneWin32AppDependency -Id $From.id
    $FromAvailable = ($FromAssignments | Where-Object intent -eq "available").groupId
    $FromRequired = ($FromAssignments | Where-Object intent -eq "required").groupId
    if ($FromAvailable) {
        foreach ($groupId in $FromAvailable) {
            $maxRetries = 3
            $try = 0
            $successfullyAdded = $false
            while (!$successfullyAdded -and ($try -lt $maxRetries)) {
                if ($ArchitectureFilterName) {
                    Write-Log "Using architecture filter for $ArchitectureFilterName"
                    Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $groupId -Intent "available" -Notification "hideAll" -FilterMode Include -FilterName "$ArchitectureFilterName" | Out-Null
                }
                else {
                    Write-Log "Not using an architecture filter: $Architecture"
                    Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $groupId -Intent "available" -Notification "hideAll" | Out-Null
                }
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
            while (!$successfullyAdded -and ($try++ -lt $maxRetries)) {
                Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $groupId -Intent "required" -Notification "hideAll" | Out-Null
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
    if($FromDependencies) {
        foreach ($dependency in $FromDependencies) {
            $maxRetries = 3
            $try = 0
            $successfullyAdded = $false
            $dependentApps = Get-IntuneWin32AppDependency -ID $dependency.sourceId
            if ($dependentApps -and $dependentApps.Count -gt 1) {
                while (!$successfullyAdded -and ($try++ -lt $maxRetries)) {
                    Write-Log "Multiple dependencies found for $($dependency.sourceId)"
                    foreach ($dependentApp in $dependentApps) {
                        $appDependenciesToMigrate = Get-IntuneWin32AppDependency -ID $dependentApp.id
                        # Clear all existing dependencies
                        Remove-IntuneWin32AppDependency -ID $dependentApp.targetId
                        # Add all dependencies back that are not the one we are moving
                        foreach ($appDependency in $appDependenciesToMigrate) {
                            if ($appDependency.targetId -ne $dependency.targetId) {
                                $toAdd = New-IntuneWin32AppDependency -ID $appDependency.id -DependencyType $appDependency.dependencyType
                                Add-IntuneWin32AppDependency -ID $dependentApp.targetId -Dependency $toAdd | Out-Null
                            }
                            else {
                                # Add the new dependency instead
                                $NewDependency = New-IntuneWin32AppDependency -ID $To.id -DependencyType $dependency.dependencyType
                                Add-IntuneWin32AppDependency -ID $dependency.targetId -Dependency $NewDependency | Out-Null
                            }
                        }

                    }
                }
            }
            elseif ($dependentApps) {
                while (!$successfullyAdded -and ($try++ -lt $maxRetries)) {
                    $NewDependency = New-IntuneWin32AppDependency -ID $To.id -DependencyType $dependency.dependencyType
                    Add-IntuneWin32AppDependency -ID $dependency.targetId -Dependency $NewDependency | Out-Null
                    # Check that it worked
                    $AssignedDependencies = Get-IntuneWin32AppDependency -ID $To.id
                    if($AssignedDependencies | Where-Object targetId -eq $dependency.targetId) {
                        $successfullyAdded = $true
                        Write-Log "Dependency $($dependency.dependencyId) added successfully to $($To.id)"
                    }
                }
            }
            else {
                Write-Log "No dependencies found for $($dependency.sourceId). Skipping dependency move."
            }
            
        }
    }
}


# Get-SameAppAllVersions
# Returns all versions of an app sorted from newest to oldest
# Accounts for edge cases where an application name might be similar to others, i.e. Mozilla Firefox vs. Mozilla Firefox ESR
# Param: [String] DisplayName
# Return: @(PSCustomObject) 
function Get-SameAppAllVersions($DisplayName) {
    $AllSimilarApps = Get-IntuneWin32App -DisplayName "$DisplayName"
    # Only return apps that have the same name (not ones that look the same) and apps that are pending approval
    $sortable = ($AllSimilarApps | Where-Object {($_.DisplayName -eq $DisplayName) -or ($_.DisplayName -like "$DisplayName (N-*")})
    # Pad out each version section to 8 digits of zeroes before sorting, remove any letters, and mash it all together so that versions sort correctly
    return $sortable | Sort-Object {$($($($_.displayVersion -replace "[A-Za-z]", "0") -replace "[\-\+]", ".").split(".") | ForEach-Object {'{0:d8}' -f [int]$_}) -join ''} -Descending
}


# Format-FileDetectionVersion
# Converts a version number to one padded with the appropriate number of zeroes for detection by Intune (4)
# Param: [String] Version
# Return: [String] FileDetectionVersion
function Format-FileDetectionVersion($Version) {
    $VersionComponents = $Version.split(".")
    Switch($VersionComponents.count) {
      1 {$FileVersion = "$($VersionComponents[0]).0.0.0"; Break}
      2 {$FileVersion = "$($VersionComponents[0]).$($VersionComponents[1]).0.0"; Break}
      3 {$FileVersion = "$($VersionComponents[0]).$($VersionComponents[1]).$($VersionComponents[2]).0"; Break}
      default {$FileVersion = "$($VersionComponents[0]).$($VersionComponents[1]).$($VersionComponents[2]).$($VersionComponents[3])"}
    }
    return $FileVersion
}


# Get-VersionLocked
# Returns $false if the version is allowed to update, $true if the versionlock parameter will not allow it.
# If the versionLock is null, it will always return $false.
# In the $versionLock parameter, values with x are ignored, i.e. 1.2.x is locked to 1.2
# Param: [String] $version, [String] $versionLock
# Return: [Boolean] $false if locked, otherwise $true
function Get-VersionLocked {
    param (
        [Parameter(Mandatory=$true)]
        [String]$version,
        [Parameter(Mandatory=$false)]
        [String]$versionLock
    )
    # If versionLock is null, return false
    if (-not $versionLock) {
        return $false
    }
    # Compare the version and versionLock
    $versionPattern = $versionLock -replace "[Xx]{1,}", "[0-9]{1,}"
    $versionPattern = $versionPattern -replace "\.", "\."
    return $version -notmatch "^$versionPattern"
}

